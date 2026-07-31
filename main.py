import uuid
import zipfile
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from test import convert_pdf_to_excel


app = FastAPI()


UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
STATIC_DIR = Path("static")

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)


templates = Jinja2Templates(
    directory="templates"
)


app.mount(
    "/static",
    StaticFiles(directory=STATIC_DIR),
    name="static"
)


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse(
        request,
        "index.html",
    )


@app.post("/convert")
async def convert(
    file: UploadFile = File(...)
):
    if file.content_type != "application/pdf":
        raise HTTPException(
            status_code=400,
            detail="Upload must be a PDF file."
        )

    input_filename = Path(file.filename or f"{uuid.uuid4()}.pdf").name
    input_file_stem = Path(input_filename).stem

    pdf_path = (
        UPLOAD_DIR /
        input_filename
    )
    output_filename = f"{input_file_stem}.xlsx"
    html_output_dir = OUTPUT_DIR / input_file_stem
    zip_path = OUTPUT_DIR / f"{input_file_stem}.zip"
    download_filename = f"{input_file_stem}.zip"


    with open(pdf_path, "wb") as f:
        f.write(
            await file.read()
        )


    print("Starting conversion")
    print(pdf_path)


    try:
        conversion_result = convert_pdf_to_excel(
            str(pdf_path),
            str(OUTPUT_DIR),
            output_filename=output_filename,
            html_output_dir=str(html_output_dir),
        )
    except Exception as error:
        print("Conversion failed:", error)
        raise HTTPException(
            status_code=502,
            detail=f"Conversion failed: {error}"
        ) from error

    excel_file = conversion_result["excel_file"]
    html_files = conversion_result["html_files"]

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(
            excel_file,
            arcname=output_filename,
        )

        for html_file in dict.fromkeys(html_files):
            archive.write(
                html_file,
                arcname=f"{input_file_stem}/{Path(html_file).name}",
            )

    return FileResponse(
        zip_path,
        filename=download_filename,
        media_type="application/zip",
    )
