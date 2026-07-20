import atexit
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import HTTPException, UploadFile
from starlette.background import BackgroundTask

from .config import MAX_UPLOAD_BYTES

# Each worker gets its own directory.  This module is imported independently by
# each worker process, and mkdtemp() guarantees that the directories do not
# collide or remove files belonging to another worker.
TEMP_DIR = Path(tempfile.mkdtemp(prefix="doc_converter-"))


def _remove_temp_dir() -> None:
    """Remove this worker's temporary directory when the process exits."""
    shutil.rmtree(TEMP_DIR, ignore_errors=True)


atexit.register(_remove_temp_dir)


def validate_file_extension(filename: str | None, allowed: list[str]) -> None:
    """
    Raise HTTP 422 if *filename* does not end with one of the *allowed*
    extensions (case-insensitive).
    """
    if not filename:
        raise HTTPException(status_code=422, detail="No filename provided.")

    ext = Path(filename).suffix.lower()
    if ext not in [a.lower() for a in allowed]:
        raise HTTPException(
            status_code=422,
            detail=f"Invalid file type '{ext}'. Expected one of: {allowed}.",
        )


async def save_upload(file: UploadFile) -> Path:
    """
    Stream *file* to a uniquely-named temp path and return that path.

    Raises:
        HTTPException 413: If the file exceeds MAX_UPLOAD_BYTES.
    """
    # Check Content-Length header if available
    if file.size and file.size > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Maximum size is {MAX_UPLOAD_BYTES // (1024 * 1024)} MB.",
        )

    suffix = Path(file.filename or "upload").suffix
    dest = TEMP_DIR / f"{uuid.uuid4().hex}{suffix}"

    total = 0
    chunk_size = 64 * 1024  # 64 KB

    try:
        with dest.open("wb") as fh:
            while chunk := await file.read(chunk_size):
                total += len(chunk)
                if total > MAX_UPLOAD_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail=f"File too large. Maximum size is {MAX_UPLOAD_BYTES // (1024 * 1024)} MB.",
                    )
                fh.write(chunk)
    except HTTPException:
        dest.unlink(missing_ok=True)
        raise
    except Exception:
        dest.unlink(missing_ok=True)
        raise

    return dest


def get_output_path(input_path: Path, new_suffix: str) -> Path:
    """Return a sibling temp path with *new_suffix* (e.g. '.pdf')."""
    return input_path.with_suffix(new_suffix)


def delete_file(path: Path) -> None:
    """Delete a single file immediately, best-effort."""
    try:
        path.unlink(missing_ok=True)
    except Exception:
        pass


def cleanup(*paths: Path) -> BackgroundTask:
    """
    Return a Starlette BackgroundTask that deletes all given paths after
    the response has been sent.
    """

    def _delete():
        for p in paths:
            try:
                Path(p).unlink(missing_ok=True)
            except Exception:
                pass  # best-effort

    return BackgroundTask(_delete)
