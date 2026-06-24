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

    if sys.platform != "win32":
        # Drain the queue so callers get NotImplementedError instead of hanging.
        while True:
            item = _work_queue.get()
            if item is None:
                break
            _, evt, box = item
            box[0] = NotImplementedError("win32com is only available on Windows.")
            evt.set()
        return

    # Windows-only path — imports are deferred so this module loads on Linux too.
    import pythoncom  # noqa: PLC0415
    import win32com.client  # noqa: PLC0415
    import win32process  # noqa: PLC0415

    pythoncom.CoInitialize()

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        _, pid = win32process.GetWindowThreadProcessId(word.Hwnd)
        _word_pid = pid
    except Exception:
        _word_pid = None

    try:
        while True:
            item = _work_queue.get()
            if item is None:
                break
            fn, evt, box = item
            try:
                box[0] = fn(word)
            except Exception as exc:
                box[0] = exc
            finally:
                evt.set()
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _start_worker() -> None:
    global _worker_thread
    if _worker_thread is None or not _worker_thread.is_alive():
        _worker_thread = threading.Thread(target=_com_worker, daemon=True, name="com-worker")
        _worker_thread.start()


def _run_on_com_thread(fn, timeout: float = CONVERSION_TIMEOUT):
    """
    Submit fn(word) to the COM worker and block until done or timeout.
    Raises TimeoutError on timeout (and attempts to kill Word).
    Re-raises any exception thrown inside fn.
    """
    evt = threading.Event()
    box: list = [None]
    _work_queue.put((fn, evt, box))

    finished = evt.wait(timeout=timeout)
    if not finished:
        _kill_word()
        raise TimeoutError(f"Conversion exceeded {timeout}s timeout — Word process killed.")

    result = box[0]
    if isinstance(result, BaseException):
        raise result
    return result


def _kill_word() -> None:
    """Force-kill the Word.exe process and restart the worker thread."""
    global _word_pid, _worker_thread
    pid = _word_pid
    if pid:
        try:
            subprocess.call(["taskkill", "/F", "/PID", str(pid)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
        _word_pid = None
    # Drain any pending item that will never complete
    try:
        _work_queue.get_nowait()
    except queue.Empty:
        pass
    # Restart worker so subsequent requests are not permanently blocked
    _worker_thread = None
    _start_worker()


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
        return
    for pattern in _OFFICE_CACHE_GLOBS:
        expanded = Path(os.path.expandvars(pattern))
        parent = expanded.parent
        glob = expanded.name
        if not parent.exists():
            continue
        for item in parent.glob(glob):
            try:
                if item.is_dir():
                    shutil.rmtree(item, ignore_errors=True)
                else:
                    item.unlink(missing_ok=True)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Lifespan
# ---------------------------------------------------------------------------


@asynccontextmanager
async def lifespan(*_):
    _cleanup_office_cache()
    _start_worker()
    yield
    # Signal the COM worker to quit gracefully
    _work_queue.put(None)
    if _worker_thread:
        _worker_thread.join(timeout=10)
    _cleanup_office_cache()


# ---------------------------------------------------------------------------
# Conversion helpers
# ---------------------------------------------------------------------------


def _do_convert(input_path: Path, output_path: Path, wd_format: int):
    """Closure factory — returns a function that runs inside the COM thread."""

    def _inner(word):
        doc = word.Documents.Open(str(input_path.resolve()))
        try:
            doc.WebOptions.Encoding = 65001
            doc.SaveAs(str(output_path.resolve()), FileFormat=wd_format)
        finally:
            doc.Close(0)

    return _inner


def convert(input_path: Path, output_path: Path, wd_format: int) -> None:
    if sys.platform != "win32":
        raise NotImplementedError("win32com is only available on Windows.")
    _run_on_com_thread(_do_convert(input_path, output_path, wd_format))


def doc_to_docx(input_path: Path, output_path: Path) -> None:
    convert(input_path, output_path, 16)


def docx_to_pdf(input_path: Path, output_path: Path) -> None:
    convert(input_path, output_path, 17)


def docx_to_html(input_path: Path, output_path: Path) -> None:
    convert(input_path, output_path, 8)


def docx_to_html_zip(input_path: Path, output_zip_path: Path, base_name: str | None = None) -> Path:
    if base_name is None:
        base_name = output_zip_path.stem

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

        return output_zip_path
    finally:
        if temp_html_dir.exists():
            shutil.rmtree(temp_html_dir, ignore_errors=True)
