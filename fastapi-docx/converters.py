import logging
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path

from .config import CONVERSION_TIMEOUT

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# COM worker thread
#
# win32com / pywin32 require that every COM call happens on the same OS thread
# that called CoInitialize.  We pin one dedicated thread for the lifetime of
# the process and route every conversion through it via a simple queue.
#
# Each work item is a tuple:
#   (callable, result_event, result_box)
# The worker calls callable(), stores the return value or exception in
# result_box[0], then sets result_event so the caller can unblock.
# ---------------------------------------------------------------------------

_work_queue: queue.Queue = queue.Queue()
_worker_thread: threading.Thread | None = None
_word_pid: int | None = None  # PID of the Word.exe process, for forced kill on timeout


def _is_rpc_unavailable(exc: BaseException) -> bool:
    """Return whether pywin32 reports that the Word COM server disappeared."""
    # 0x800706BA: The RPC server is unavailable.
    return bool(getattr(exc, "args", ())) and getattr(exc, "args", ())[0] == -2147023174


def _com_worker() -> None:
    """Runs forever in its own thread; owns the Word COM object."""
    global _word_pid

    logger.info("COM worker thread started (platform=%s)", sys.platform)
    if sys.platform != "win32":
        logger.warning("COM worker unavailable: win32com is only available on Windows")
        # Drain the queue so callers get NotImplementedError instead of hanging.
        while True:
            item = _work_queue.get()
            if item is None:
                logger.debug("COM worker received shutdown signal on unsupported platform")
                break
            _, evt, box = item
            box[0] = NotImplementedError("win32com is only available on Windows.")
            evt.set()
            logger.debug("Rejected queued COM operation on unsupported platform")
        return

    # Windows-only path — imports are deferred so this module loads on Linux too.
    import pythoncom  # noqa: PLC0415
    import win32com.client  # noqa: PLC0415
    import win32process  # noqa: PLC0415

    pythoncom.CoInitialize()
    logger.debug("COM initialized")

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone — suppress all modal dialogs
        logger.info("Word COM application started")
    except Exception:
        logger.exception("Failed to start Word COM application")
        pythoncom.CoUninitialize()
        raise

    try:
        _, pid = win32process.GetWindowThreadProcessId(word.Hwnd)
        _word_pid = pid
        logger.debug("Word process identified (pid=%s)", pid)
    except Exception:
        _word_pid = None
        logger.warning("Could not identify the Word process ID", exc_info=True)

    try:
        while True:
            item = _work_queue.get()
            if item is None:
                break
            fn, evt, box = item
            word_lost = False
            try:
                logger.debug("Starting queued COM operation")
                box[0] = fn(word)
                logger.debug("Queued COM operation completed")
            except Exception as exc:
                box[0] = exc
                logger.exception("Queued COM operation failed")
                # A Word COM proxy cannot recover after Word.exe has exited.
                # Stop this worker so the next request creates a fresh proxy.
                word_lost = _is_rpc_unavailable(exc)
            finally:
                evt.set()
            if word_lost:
                logger.warning("Word COM connection was lost; stopping worker for restart")
                break
    finally:
        try:
            word.Quit()
            logger.info("Word COM application stopped")
        except Exception:
            logger.exception("Failed to stop Word COM application")
        pythoncom.CoUninitialize()
        logger.debug("COM uninitialized")


def _start_worker() -> None:
    global _worker_thread
    if _worker_thread is None or not _worker_thread.is_alive():
        logger.info("Starting COM worker thread")
        _worker_thread = threading.Thread(target=_com_worker, daemon=True, name="com-worker")
        _worker_thread.start()
    else:
        logger.debug("COM worker thread is already running")


def _run_on_com_thread(fn, timeout: float = CONVERSION_TIMEOUT):
    """
    Submit fn(word) to the COM worker and block until done or timeout.
    Raises TimeoutError on timeout (and attempts to kill Word).
    Re-raises any exception thrown inside fn.
    """
    global _worker_thread
    for attempt in range(2):
        _start_worker()
        evt = threading.Event()
        box: list = [None]
        logger.debug("Queueing COM operation (timeout=%ss, attempt=%s)", timeout, attempt + 1)
        _work_queue.put((fn, evt, box))

        finished = evt.wait(timeout=timeout)
        if not finished:
            logger.error("COM operation timed out after %ss; killing Word", timeout)
            _kill_word()
            raise TimeoutError(f"Conversion exceeded {timeout}s timeout — Word process killed.")

        result = box[0]
        if isinstance(result, BaseException):
            logger.error("COM operation raised %s", type(result).__name__)
            if _is_rpc_unavailable(result) and attempt == 0:
                # The worker exits after detecting a dead Word proxy. Wait for
                # its COM cleanup before starting a replacement thread.
                worker = _worker_thread
                if worker is not None:
                    worker.join(timeout=5)
                if worker is not None and worker.is_alive():
                    raise result
                _worker_thread = None
                logger.warning("Retrying operation with a new Word COM instance")
                continue
            raise result
        logger.debug("COM operation returned successfully")
        return result

    raise RuntimeError("COM operation could not be completed")


def _kill_word() -> None:
    """Force-kill the Word.exe process and restart the worker thread."""
    global _word_pid, _worker_thread
    pid = _word_pid
    if pid:
        logger.warning("Force-killing Word process (pid=%s)", pid)
        try:
            subprocess.call(["taskkill", "/F", "/PID", str(pid)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            logger.exception("Failed to force-kill Word process (pid=%s)", pid)
        _word_pid = None
    else:
        logger.warning("Requested Word termination, but no process ID is available")
    # Drain any pending item that will never complete
    try:
        _work_queue.get_nowait()
    except queue.Empty:
        logger.debug("No pending COM operation to discard")
    else:
        logger.debug("Discarded one pending COM operation after timeout")
    # Restart worker so subsequent requests are not permanently blocked
    _worker_thread = None
    _start_worker()
    logger.info("COM worker restart requested after Word termination")


# ---------------------------------------------------------------------------
# Office cache cleanup
# ---------------------------------------------------------------------------

_OFFICE_CACHE_GLOBS = [
    r"%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache\*",
    r"%LOCALAPPDATA%\Microsoft\Office\16.0\WebServiceCache\*",
    r"%APPDATA%\Microsoft\Office\Recent\*",
    r"%LOCALAPPDATA%\Temp\*.tmp",
]

# ---------------------------------------------------------------------------
# Cross-process "cleanup leader" lock
#
# When the app is run with multiple workers (e.g. multiple Uvicorn/Gunicorn
# processes), each worker runs its own lifespan and would otherwise try to
# clean up the same Office cache folders at the same time, racing each other.
# Only one worker (whichever gets there first) should actually do the work;
# the rest should skip it. We arbitrate this with an atomically-created lock
# file: os.O_CREAT | os.O_EXCL is atomic across processes on both Windows and
# POSIX, so exactly one process will win the race to create it.
# ---------------------------------------------------------------------------

_CLEANUP_LOCK_PATH = Path(tempfile.gettempdir()) / "fastapi_docx_office_cache_cleanup.lock"
_CLEANUP_LOCK_STALE_SECONDS = 60


def _try_become_cleanup_leader() -> bool:
    """
    Attempt to atomically claim the right to perform Office cache cleanup.
    Returns True if this process/worker won the race and should perform the
    cleanup, False if another worker already claimed it (or claimed it very
    recently).
    """
    for attempt in range(2):
        try:
            fd = os.open(_CLEANUP_LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, str(os.getpid()).encode())
            finally:
                os.close(fd)
            logger.debug("Became Office cache cleanup leader (pid=%s)", os.getpid())
            return True
        except FileExistsError:
            if attempt == 0:
                try:
                    age = time.time() - _CLEANUP_LOCK_PATH.stat().st_mtime
                except OSError:
                    age = None
                if age is not None and age > _CLEANUP_LOCK_STALE_SECONDS:
                    logger.warning(
                        "Cleanup lock file is stale (age=%.1fs); removing and retrying: %s",
                        age,
                        _CLEANUP_LOCK_PATH,
                    )
                    try:
                        _CLEANUP_LOCK_PATH.unlink(missing_ok=True)
                    except OSError:
                        logger.exception("Failed to remove stale cleanup lock file")
                        return False
                    continue
            logger.debug("Another worker already holds the Office cache cleanup lock")
            return False
        except OSError:
            logger.exception("Unexpected error while acquiring cleanup lock")
            return False
    return False


def _release_cleanup_leader() -> None:
    try:
        _CLEANUP_LOCK_PATH.unlink(missing_ok=True)
        logger.debug("Released Office cache cleanup lock")
    except OSError:
        logger.exception("Failed to release Office cache cleanup lock: %s", _CLEANUP_LOCK_PATH)


def _cleanup_office_cache() -> None:
    if sys.platform != "win32":
        logger.debug("Skipping Office cache cleanup on unsupported platform")
        return

    if not _try_become_cleanup_leader():
        logger.info("Skipping Office cache cleanup; another worker is already handling it")
        return

    try:
        logger.info("Cleaning Office cache files")
        removed = 0
        for pattern in _OFFICE_CACHE_GLOBS:
            expanded = Path(os.path.expandvars(pattern))
            parent = expanded.parent
            glob = expanded.name
            if not parent.exists():
                logger.debug("Office cache directory does not exist: %s", parent)
                continue
            try:
                items = list(parent.glob(glob))
            except OSError:
                logger.exception("Failed to list Office cache items in: %s", parent)
                continue
            for item in items:
                try:
                    if item.is_dir():
                        shutil.rmtree(item, ignore_errors=True)
                    else:
                        item.unlink(missing_ok=True)
                    removed += 1
                except Exception:
                    logger.exception("Failed to remove Office cache item: %s", item)
        logger.info("Office cache cleanup finished (items_removed=%s)", removed)
    finally:
        _release_cleanup_leader()


# ---------------------------------------------------------------------------
# Lifespan
# ---------------------------------------------------------------------------


@asynccontextmanager
async def lifespan(*_):
    logger.info("Starting converter application lifespan")
    _cleanup_office_cache()
    _start_worker()
    try:
        yield
    finally:
        logger.info("Stopping converter application lifespan")
        # Signal the COM worker to quit gracefully
        _work_queue.put(None)
        if _worker_thread:
            _worker_thread.join(timeout=10)
            if _worker_thread.is_alive():
                logger.warning("COM worker did not stop within 10 seconds")
            else:
                logger.info("COM worker stopped")
        _cleanup_office_cache()


# ---------------------------------------------------------------------------
# Conversion helpers
# ---------------------------------------------------------------------------


def _do_convert(input_path: Path, output_path: Path, wd_format: int):
    """Closure factory — returns a function that runs inside the COM thread."""

    def _inner(word):
        logger.info("Opening document for conversion: %s", input_path)
        doc = word.Documents.Open(str(input_path.resolve()))
        try:
            doc.WebOptions.Encoding = 65001
            doc.SaveAs(str(output_path.resolve()), FileFormat=wd_format)
            logger.info("Document conversion output created: %s", output_path)
        finally:
            doc.Close(0)
            logger.debug("Closed input document: %s", input_path)

    return _inner


def convert(input_path: Path, output_path: Path, wd_format: int) -> None:
    logger.info("Conversion requested: %s -> %s (Word format=%s)", input_path, output_path, wd_format)
    if sys.platform != "win32":
        logger.warning("Rejecting conversion on unsupported platform: %s", sys.platform)
        raise NotImplementedError("win32com is only available on Windows.")
    try:
        _run_on_com_thread(_do_convert(input_path, output_path, wd_format))
    except Exception:
        logger.exception("Conversion failed: %s -> %s", input_path, output_path)
        raise
    logger.info("Conversion completed: %s -> %s", input_path, output_path)


def doc_to_docx(input_path: Path, output_path: Path) -> None:
    logger.debug("Converting DOC to DOCX: %s -> %s", input_path, output_path)
    convert(input_path, output_path, 16)


def docx_to_pdf(input_path: Path, output_path: Path) -> None:
    logger.debug("Converting DOCX to PDF: %s -> %s", input_path, output_path)
    convert(input_path, output_path, 17)


def docx_to_html(input_path: Path, output_path: Path) -> None:
    logger.debug("Converting DOCX to HTML: %s -> %s", input_path, output_path)
    convert(input_path, output_path, 8)


def docx_to_html_zip(input_path: Path, output_zip_path: Path, base_name: str | None = None) -> Path:
    if base_name is None:
        base_name = output_zip_path.stem

    logger.info("Converting DOCX to HTML ZIP: %s -> %s", input_path, output_zip_path)

    temp_html_dir = output_zip_path.parent / f"{output_zip_path.stem}_html_temp"
    temp_html_dir.mkdir(exist_ok=True)
    html_output_path = temp_html_dir / f"{base_name}.html"

    try:
        convert(input_path, html_output_path, 8)

        with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(html_output_path, arcname=f"{base_name}.html")

            files_folder = temp_html_dir / f"{base_name}_files"
            if files_folder.exists():
                for file_path in files_folder.rglob("*"):
                    if file_path.is_file():
                        arcname = f"{base_name}_files/{file_path.relative_to(files_folder)}"
                        zipf.write(file_path, arcname=arcname)

        logger.info("Created HTML ZIP output: %s", output_zip_path)
        return output_zip_path
    except Exception:
        logger.exception("DOCX to HTML ZIP conversion failed: %s -> %s", input_path, output_zip_path)
        raise
    finally:
        if temp_html_dir.exists():
            shutil.rmtree(temp_html_dir, ignore_errors=True)
            logger.debug("Removed temporary HTML directory: %s", temp_html_dir)
        logger.debug("DOCX to HTML ZIP cleanup finished: %s", output_zip_path)
