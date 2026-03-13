import sys
from contextlib import asynccontextmanager
from pathlib import Path


@asynccontextmanager
async def lifespan(app):
    yield
    if sys.platform == "win32":
        import win32com.client

        try:
            word = win32com.client.GetActiveObject("Word.Application")
            word.Quit()
        except Exception:
            pass  # Word was not running


def convert(input_path: Path, output_path: Path, wd_format: int, visible: bool = False, keep_active: bool = True):
    """
    Convert a Word document to another format using win32com (Windows only).

    Args:
        input_path (Path): Resolved path to the input file.
        output_path (Path): Resolved path to the output file.
        wd_format (int): Word save format constant, e.g.:
                            8  = wdFormatHTML
                            16 = wdFormatDocumentDefault (.docx)
                            17 = wdFormatPDF
        visible (bool): Whether to show the Word application window. Defaults to False.
        keep_active (bool): Prevent quitting Word after conversion. Defaults to True.
    """
    if sys.platform != "win32":
        raise NotImplementedError("win32com is only available on Windows.")

    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = visible
    try:
        doc = word.Documents.Open(str(input_path.resolve()))
        try:
            doc.WebOptions.Encoding = 65001
            doc.SaveAs(str(output_path.resolve()), FileFormat=wd_format)
        finally:
            doc.Close(0)
    finally:
        if not keep_active:
            word.Quit()


def doc_to_docx(input_path: Path, output_path: Path) -> None:
    convert(input_path, output_path, 16)


def docx_to_pdf(input_path: Path, output_path: Path) -> None:
    convert(input_path, output_path, 17)


def docx_to_html(input_path: Path, output_path: Path) -> None:
    convert(input_path, output_path, 8)
