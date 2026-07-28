import json
from pathlib import Path

from test import chandra_convert


def extract_html_from_pdf(pdf_path: str, output_dir: str):
    """
    Sends a PDF to Datalab Chandra OCR pipeline and saves returned JSON.

    Returns:
        Path to JSON file.
    """

    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    json_output = output_dir / f"{pdf_path.stem}.json"
    result = chandra_convert(pdf_path)

    json_output.write_text(
        json.dumps(result, indent=2),
        encoding="utf-8",
    )

    print(f"Saved pipeline JSON: {json_output}")

    return str(json_output)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print(
            "Usage:\n"
            "python api_extract.py file.pdf"
        )
        raise SystemExit(1)

    extract_html_from_pdf(
        sys.argv[1],
        "outputs/html",
    )
