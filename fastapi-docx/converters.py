import sys
import shutil
import zipfile
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


def docx_to_html_zip(input_path: Path, output_zip_path: Path, base_name: str | None = None) -> Path:
    """
    Convert a .docx file to HTML and return a ZIP file containing the HTML and its assets.

    Args:
        input_path (Path): Path to the input .docx file.
        output_zip_path (Path): Path where the output .zip file should be saved.
        base_name (str | None): Base name for files inside the ZIP (e.g., "document" for "document.html" and "document_files/").
                                If None, uses output_zip_path.stem.

    Returns:
        Path: Path to the created ZIP file.
    """
    if base_name is None:
        base_name = output_zip_path.stem

    # Create a temporary directory for the HTML conversion
    temp_html_dir = output_zip_path.parent / f"{output_zip_path.stem}_html_temp"
    temp_html_dir.mkdir(exist_ok=True)

    html_output_path = temp_html_dir / f"{base_name}.html"

    try:
        # Convert docx to HTML
        convert(input_path, html_output_path, 8)

        # Create ZIP file
        with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            # Add the HTML file with the correct name inside the ZIP
            zipf.write(html_output_path, arcname=f"{base_name}.html")

            # Add the _files folder if it exists
            files_folder = temp_html_dir / f"{base_name}_files"
            
            if files_folder.exists():
                for file_path in files_folder.rglob("*"):
                    if file_path.is_file():
                        arcname = f"{base_name}_files/{file_path.relative_to(files_folder)}"
                        zipf.write(file_path, arcname=arcname)

        return output_zip_path

    finally:
        # Cleanup temporary HTML conversion files
        if temp_html_dir.exists():
            shutil.rmtree(temp_html_dir, ignore_errors=True)
