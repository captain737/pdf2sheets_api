# pdf2sheets_api

FastAPI backend that accepts PDF uploads, runs a configured Datalab Chandra OCR pipeline, extracts table-like output, and returns a downloadable Excel file.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export DATALAB_API_KEY="your_datalab_api_key"
```

The default pipeline id is `pl_4AHLbwoxranz`. To override it:

```bash
export DATALAB_PIPELINE_ID="your_pipeline_id"
```

## Run

```bash
python3 -m uvicorn main:app --reload
```

Open:

```text
http://127.0.0.1:8000
```

## CLI Upload Test

```bash
curl -X POST \
  -F "file=@/path/to/document.pdf;type=application/pdf" \
  http://127.0.0.1:8000/convert \
  --output converted.xlsx
```
