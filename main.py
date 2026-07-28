import uuid
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


    with open(pdf_path, "wb") as f:
        f.write(
            await file.read()
        )


    print("Starting conversion")
    print(pdf_path)


    try:
        excel_file = convert_pdf_to_excel(
            str(pdf_path),
            str(OUTPUT_DIR),
            output_filename=output_filename
        )
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=str(error)
        ) from error


    return FileResponse(
        excel_file,
        filename="converted.xlsx",
        media_type=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )
