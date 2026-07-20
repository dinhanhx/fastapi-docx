import logging
import os
import queue
import shutil
import subprocess
import sys
import threading
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
            try:
                logger.debug("Starting queued COM operation")
                box[0] = fn(word)
                logger.debug("Queued COM operation completed")
            except Exception as exc:
                box[0] = exc
                logger.exception("Queued COM operation failed")
            finally:
                evt.set()
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
    evt = threading.Event()
    box: list = [None]
    logger.debug("Queueing COM operation (timeout=%ss)", timeout)
    _work_queue.put((fn, evt, box))

    finished = evt.wait(timeout=timeout)
    if not finished:
        logger.error("COM operation timed out after %ss; killing Word", timeout)
        _kill_word()
        raise TimeoutError(f"Conversion exceeded {timeout}s timeout — Word process killed.")

    result = box[0]
    if isinstance(result, BaseException):
        logger.error("COM operation raised %s", type(result).__name__)
        raise result
    logger.debug("COM operation returned successfully")
    return result


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


def _cleanup_office_cache() -> None:
    if sys.platform != "win32":
        logger.debug("Skipping Office cache cleanup on unsupported platform")
        return
    logger.info("Cleaning Office cache files")
    removed = 0
    for pattern in _OFFICE_CACHE_GLOBS:
        expanded = Path(os.path.expandvars(pattern))
        parent = expanded.parent
        glob = expanded.name
        if not parent.exists():
            logger.debug("Office cache directory does not exist: %s", parent)
            continue
        for item in parent.glob(glob):
            try:
                if item.is_dir():
                    shutil.rmtree(item, ignore_errors=True)
                else:
                    item.unlink(missing_ok=True)
                removed += 1
            except Exception:
                logger.exception("Failed to remove Office cache item: %s", item)
    logger.info("Office cache cleanup finished (items_removed=%s)", removed)


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
