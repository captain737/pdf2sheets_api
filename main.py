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

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


templates = Jinja2Templates(
    directory="templates"
)


app.mount(
    "/static",
    StaticFiles(directory="static"),
    name="static"
)


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request
        }
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

    file_id = str(uuid.uuid4())

    pdf_path = (
        UPLOAD_DIR /
        f"{file_id}.pdf"
    )
    output_filename = f"{file_id}.xlsx"
    html_output_dir = OUTPUT_DIR / f"{file_id}-html"
    zip_path = OUTPUT_DIR / f"{file_id}.zip"
    download_filename = f"converted-{file_id}.zip"


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
        raise HTTPException(
            status_code=502,
            detail=str(error)
        ) from error

    excel_file = conversion_result["excel_file"]
    html_files = conversion_result["html_files"]

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(
            excel_file,
            arcname="converted.xlsx",
        )

        for html_file in dict.fromkeys(html_files):
            archive.write(
                html_file,
                arcname=f"intermediate_html/{Path(html_file).name}",
            )

    return FileResponse(
        zip_path,
        filename=download_filename,
        media_type="application/zip",
    )
