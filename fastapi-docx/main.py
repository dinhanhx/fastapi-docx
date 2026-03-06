from pathlib import Path

import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from .converters import doc_to_docx, docx_to_html, docx_to_pdf, lifespan
from .utils import cleanup, get_output_path, save_upload, validate_file_extension

app = FastAPI(
    title="Document Conversion Service",
    description="Convert documents between doc, docx, pdf, and html formats.",
    version="0.0.1",
    docs_url="/docs",
    redoc_url="/redoc",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health", tags=["System"])
def health_check():
    return {"status": "ok", "version": "0.0.1"}


@app.post(
    "/convert/doc-to-docx",
    tags=["Conversions"],
    summary="Convert .doc → .docx",
    response_description="The converted .docx file as a binary download.",
)
async def convert_doc_to_docx(file: UploadFile = File(..., description="A .doc file to convert.")):

    validate_file_extension(file.filename, allowed=[".doc"])

    input_path = await save_upload(file)
    output_path = get_output_path(input_path, ".docx")
    download_name = Path(file.filename).stem + ".docx"

    try:
        doc_to_docx(input_path, output_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}") from exc
    finally:
        cleanup(input_path)

    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=download_name,
        background=cleanup(output_path),
    )


@app.post(
    "/convert/docx-to-pdf",
    tags=["Conversions"],
    summary="Convert .docx → .pdf",
    response_description="The converted .pdf file as a binary download.",
)
async def convert_docx_to_pdf(file: UploadFile = File(..., description="A .docx file to convert.")):

    validate_file_extension(file.filename, allowed=[".docx"])

    input_path = await save_upload(file)
    output_path = get_output_path(input_path, ".pdf")
    download_name = Path(file.filename).stem + ".pdf"

    try:
        docx_to_pdf(input_path, output_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}") from exc
    finally:
        cleanup(input_path)

    return FileResponse(
        path=output_path,
        media_type="application/pdf",
        filename=download_name,
        background=cleanup(output_path),
    )


@app.post(
    "/convert/docx-to-html",
    tags=["Conversions"],
    summary="Convert .docx → .html",
    response_description="The converted .html file as a binary download.",
)
async def convert_docx_to_html(file: UploadFile = File(..., description="A .docx file to convert.")):

    validate_file_extension(file.filename, allowed=[".docx"])

    input_path = await save_upload(file)
    output_path = get_output_path(input_path, ".html")
    download_name = Path(file.filename).stem + ".html"

    try:
        docx_to_html(input_path, output_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}") from exc
    finally:
        cleanup(input_path)

    return FileResponse(
        path=output_path,
        media_type="text/html",
        filename=download_name,
        background=cleanup(output_path),
    )


if __name__ == "__main__":
    uvicorn.run(
        "fastapi-docx.main:app",
        host="0.0.0.0",
        port=8000,
        workers=16,
    )
