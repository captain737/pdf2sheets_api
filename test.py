import json
import os
import re
import time
from io import StringIO
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl.styles import Alignment


load_dotenv()

DATALAB_API_BASE_URL = os.getenv(
    "DATALAB_API_BASE_URL",
    "https://www.datalab.to/api/v1",
).rstrip("/")
DATALAB_PIPELINE_ID = os.getenv("DATALAB_PIPELINE_ID", "pl_4AHLbwoxranz")
DATALAB_PIPELINE_VERSION = os.getenv("DATALAB_PIPELINE_VERSION")
DATALAB_POLL_INTERVAL_SECONDS = int(os.getenv("DATALAB_POLL_INTERVAL_SECONDS", "2"))
DATALAB_TIMEOUT_SECONDS = int(os.getenv("DATALAB_TIMEOUT_SECONDS", "900"))
SCAM_CENSUS_COLUMNS = [
    "Chamber",
    "Column",
    "Row",
    "Species",
    "Count",
    "Total_Height",
    "Green_Height",
    "Width_mm",
    "Flower",
]


def chandra_convert(pdf_path):
    """
    Run the configured Datalab Chandra OCR pipeline and return its step results.
    """

    return run_chandra_pipeline(pdf_path)


def run_chandra_pipeline(pdf_path, pipeline_id=None):
    api_key = os.getenv("DATALAB_API_KEY")

    if not api_key:
        raise RuntimeError(
            "DATALAB_API_KEY missing. Add it to .env or export it in terminal."
        )

    pipeline_id = pipeline_id or DATALAB_PIPELINE_ID
    pdf_path = Path(pdf_path)
    headers = {
        "X-API-Key": api_key,
    }

    data = {}

    if DATALAB_PIPELINE_VERSION:
        data["version"] = DATALAB_PIPELINE_VERSION

    print(f"Running Datalab Chandra OCR pipeline: {pipeline_id}")

    with open(pdf_path, "rb") as file_handle:
        response = requests.post(
            f"{DATALAB_API_BASE_URL}/pipelines/{pipeline_id}/run",
            headers=headers,
            files={
                "file": (
                    pdf_path.name,
                    file_handle,
                    "application/pdf",
                )
            },
            data=data,
            timeout=60,
        )

    print("Datalab pipeline submit status:", response.status_code)

    if not response.ok:
        raise RuntimeError(
            f"Datalab pipeline submit failed "
            f"({response.status_code}): {response.text}"
        )

    submission = response.json()
    execution_id = submission.get("execution_id")

    if not execution_id:
        raise RuntimeError(
            f"Datalab pipeline response did not include execution_id: {submission}"
        )

    execution = _poll_pipeline_execution(execution_id, headers)
    step_results = _fetch_pipeline_step_results(execution, headers)

    return {
        "execution": execution,
        "step_results": step_results,
    }


def _poll_pipeline_execution(execution_id, headers):
    deadline = time.monotonic() + DATALAB_TIMEOUT_SECONDS

    while time.monotonic() < deadline:
        response = requests.get(
            f"{DATALAB_API_BASE_URL}/pipelines/executions/{execution_id}",
            headers=headers,
            timeout=60,
        )

        if not response.ok:
            raise RuntimeError(
                f"Datalab pipeline poll failed "
                f"({response.status_code}): {response.text}"
            )

        execution = response.json()
        status = execution.get("status")

        print("Datalab pipeline status:", status)

        if status in ("completed", "completed_with_errors"):
            return execution

        if status == "failed":
            raise RuntimeError(
                f"Datalab pipeline failed: {execution}"
            )

        time.sleep(DATALAB_POLL_INTERVAL_SECONDS)

    raise TimeoutError(
        "Timed out waiting for Datalab pipeline execution to finish."
    )


def _fetch_pipeline_step_results(execution, headers):
    execution_id = execution.get("execution_id")
    steps = execution.get("steps") or []
    results = []

    for fallback_index, step in enumerate(steps):
        if step.get("status") != "completed":
            continue

        step_index = (
            step.get("step_index")
            or step.get("index")
            or fallback_index
        )

        response = requests.get(
            (
                f"{DATALAB_API_BASE_URL}/pipelines/executions/"
                f"{execution_id}/steps/{step_index}/result"
            ),
            headers=headers,
            timeout=60,
        )

        if not response.ok:
            print(
                f"Skipping step {step_index} result:",
                response.status_code,
                response.text,
            )
            continue

        results.append(
            {
                "step": step,
                "result": response.json(),
            }
        )

    if not results:
        raise RuntimeError(
            f"No completed pipeline step results found: {execution}"
        )

    return results


def pipeline_result_to_excel(pipeline_result, output_dir, output_filename="converted.xlsx"):
    dataframes = _extract_dataframes_from_pipeline_result(pipeline_result)
    return dataframes_to_excel(dataframes, output_dir, output_filename)


def _extract_dataframes_from_pipeline_result(pipeline_result):
    step_results = pipeline_result.get("step_results") or []
    all_dataframes = []

    for step_result in step_results:
        dataframes = _extract_dataframes_from_value(step_result.get("result"))

        all_dataframes.extend(dataframes)

    if all_dataframes:
        return all_dataframes

    all_dataframes = _extract_dataframes_from_value(pipeline_result)

    if not all_dataframes:
        raise RuntimeError(
            "Pipeline completed, but no HTML tables or structured table-like JSON "
            "were found in the step results."
        )

    return all_dataframes


def _extract_dataframes_from_value(value):
    rendered_texts = []
    _collect_rendered_texts(value, rendered_texts)

    if rendered_texts:
        dataframes = []

        for rendered_text in _dedupe_values(rendered_texts):
            try:
                dataframes.extend(html_to_dataframes(rendered_text))
            except RuntimeError as error:
                print("Skipping rendered text chunk:", error)

        if dataframes:
            return dataframes

    table_values = []
    _collect_table_like_values(value, table_values)

    dataframes = []

    for table_value in table_values:
        df = _value_to_dataframe(table_value)

        if df is not None and df.shape[0] >= 1 and df.shape[1] >= 1:
            dataframes.append(_clean_dataframe(df))

    return dataframes


def _collect_rendered_texts(value, rendered_texts):
    if isinstance(value, str):
        stripped = value.strip()

        if _looks_like_html_or_page_text(stripped):
            rendered_texts.append(stripped)
            return

        try:
            _collect_rendered_texts(json.loads(stripped), rendered_texts)
        except json.JSONDecodeError:
            return

    elif isinstance(value, dict):
        for key in ("html", "markdown", "content", "output", "result", "text"):
            if key in value:
                _collect_rendered_texts(value[key], rendered_texts)

        for key, nested_value in value.items():
            if key not in ("html", "markdown", "content", "output", "result", "text"):
                _collect_rendered_texts(nested_value, rendered_texts)

    elif isinstance(value, list):
        for item in value:
            _collect_rendered_texts(item, rendered_texts)


def _looks_like_html_or_page_text(value):
    lowered = value.lower()

    return (
        "<table" in lowered
        or "<html" in lowered
        or "<p" in lowered
        or re.search(r"\bpage\s*:\s*\d+\s+of\s+\d+", value, re.IGNORECASE)
    )


def _collect_table_like_values(value, table_values):
    if isinstance(value, str):
        try:
            _collect_table_like_values(json.loads(value), table_values)
        except json.JSONDecodeError:
            return

    elif isinstance(value, list):
        if value and all(isinstance(item, dict) for item in value):
            table_values.append(value)
            return

        for item in value:
            _collect_table_like_values(item, table_values)

    elif isinstance(value, dict):
        for nested_value in value.values():
            _collect_table_like_values(nested_value, table_values)


def _dedupe_values(values):
    seen = set()
    deduped = []

    for value in values:
        if value in seen:
            continue

        seen.add(value)
        deduped.append(value)

    return deduped


def _value_to_dataframe(value):
    if isinstance(value, list):
        if not value:
            return None

        if all(isinstance(item, dict) for item in value):
            return pd.json_normalize(value)

        if all(isinstance(item, list) for item in value):
            return pd.DataFrame(value)

    if isinstance(value, dict):
        if all(not isinstance(item, (dict, list)) for item in value.values()):
            return pd.DataFrame([value])

        return pd.json_normalize(value)

    return None


def html_to_excel(html, output_dir, output_filename="converted.xlsx"):
    dataframes = html_to_dataframes(html)
    return dataframes_to_excel(dataframes, output_dir, output_filename)


def html_to_dataframes(html):
    print("Reading botanical HTML tables...")

    soup = BeautifulSoup(
        html,
        "html.parser",
    )
    table_elements = soup.find_all("table")

    print("Tables found:", len(table_elements))

    cleaned = []

    for index, table_element in enumerate(table_elements, start=1):
        if not looks_like_botanical_data_table_element(table_element):
            continue

        print("Processing table", index)

        try:
            table = pd.read_html(
                StringIO(str(table_element)),
                header=[0, 1],
            )[0]
        except ValueError:
            try:
                table = pd.read_html(
                    StringIO(str(table_element)),
                    header=0,
                )[0]
            except ValueError as error:
                print("Skipping table:", error)
                continue

        table = flatten_headers(table)
        print("Headers:", table.columns.tolist())

        table = clean_headers(table)
        table = make_unique_columns(table)

        if not is_scam_census_table(table):
            continue

        cleaned.append(select_scam_census_columns(table))

    if not cleaned:
        raise RuntimeError(
            "No usable botanical tables found."
        )

    return cleaned


def looks_like_botanical_data_table_element(table_element):
    text = _normalize_header(table_element.get_text(" ", strip=True)).lower()
    rows = table_element.find_all("tr")

    if len(rows) < 4:
        return False

    signals = (
        "ch",
        "species",
        "count",
        "sub",
        "subplot",
        "code",
        "consec",
        "number",
        "plot",
        "quadrant",
        "quad",
        "height",
        "width",
        "flower",
    )
    matches = sum(
        1
        for signal in signals
        if re.search(rf"\b{re.escape(signal)}\b", text)
    )

    return matches >= 2


def flatten_headers(df):
    if not isinstance(df.columns, pd.MultiIndex):
        return df

    new_columns = []

    for column in df.columns:
        parts = []

        for item in column:
            item = str(item).strip()

            if item != "nan" and not item.startswith("Unnamed"):
                parts.append(item)

        new_columns.append(" ".join(_dedupe_adjacent(parts)))

    df.columns = new_columns

    return df


def clean_headers(df):
    df.columns = [
        _normalize_header(column)
        for column in df.columns
    ]

    rename = {
        "Subplot COL #": "Subplot",
        "Subplot Row Code": "Row_Code",
        "Consec Number 1, 2, 3... per quad": "Consec_Number",
        "Height (cm) (cm)": "Height_cm",
        "Width (mm) (only if > 200 cm)": "Width_mm",
        "Flower Y/N/C": "Flower",
        "Herbivory G/L/A/N": "Herbivory",
        "CEVO Flag #": "CEVO_Flag",
        "CH": "Chamber",
        "Col": "Column",
        "Col (see map)": "Column",
        "Row": "Row",
        "Species": "Species",
        "Count": "Count",
        "Height Total (cm)": "Total_Height",
        "Height Green (cm)": "Green_Height",
        "Total (cm)": "Total_Height",
        "Green (cm)": "Green_Height",
        "Width (mm)": "Width_mm",
        "Width": "Width_mm",
    }

    return df.rename(columns=rename)


def is_scam_census_table(df):
    return all(
        column in df.columns
        for column in SCAM_CENSUS_COLUMNS
    )


def select_scam_census_columns(df):
    df = df[SCAM_CENSUS_COLUMNS].copy()

    for column in ("Chamber", "Column", "Count", "Total_Height", "Green_Height", "Width_mm"):
        df[column] = df[column].map(_prefer_latter_ocr_candidate)

    return df


def _prefer_latter_ocr_candidate(value):
    if value is None:
        return value

    if isinstance(value, float) and pd.isna(value):
        return value

    text = str(value).strip()

    if not re.fullmatch(r"\d+(?:\.\d+)?\s+\d+(?:\.\d+)?", text):
        return value

    return text.split()[-1]


def _normalize_header(value):
    value = str(value).strip()
    value = re.sub(r"\s+", " ", value)
    value = value.replace(" (", " (")

    return value


def make_unique_columns(df):
    seen = {}
    columns = []

    for column in df.columns:
        column = str(column)

        if column not in seen:
            seen[column] = 0
            columns.append(column)
            continue

        seen[column] += 1
        columns.append(f"{column}_{seen[column]}")

    df.columns = columns

    return df


def is_data_table(df):
    columns = [
        str(column).lower()
        for column in df.columns
    ]

    keywords = [
        "height",
        "species",
        "subplot",
        "count",
        "flower",
        "width",
    ]

    matches = sum(
        any(keyword in column for column in columns)
        for keyword in keywords
    )

    return matches >= 2


def _dedupe_adjacent(values):
    deduped = []

    for value in values:
        if deduped and deduped[-1] == value:
            continue

        deduped.append(value)

    return deduped


def _page_sections_to_dataframes(text):
    page_rows = _parse_page_marker_rows(text)

    if not page_rows:
        return []

    return [
        _clean_dataframe(pd.DataFrame(rows))
        for rows in page_rows
    ]


def _parse_page_marker_rows(text):
    raw_lines = [
        line.strip()
        for line in _normalize_text(text).splitlines()
        if line.strip()
    ]
    lines = _join_split_field_value_lines(raw_lines)
    page_rows = []
    pending_header_rows = []
    current_rows = None
    current_page = None

    for line in lines:
        cleaned_line = _strip_markdown_emphasis(line)
        field_value = _parse_field_value_line(cleaned_line)

        if not field_value:
            continue

        field, value = field_value
        page_match = re.match(r"^(\d+)\s+of\s+\d+$", value, re.IGNORECASE)

        if field.lower() == "page" and page_match:
            if current_rows:
                page_rows.append(current_rows)

            current_page = int(page_match.group(1))
            current_rows = [
                _page_row(current_page, header_field, header_value)
                for header_field, header_value in pending_header_rows
            ]
            current_rows.append(_page_row(current_page, field, value))
            pending_header_rows = []
            continue

        if current_page is None:
            pending_header_rows.append((field, value))
            continue

        if field.lower() in ("date", "chamber") and _looks_like_new_page_header(field, value):
            pending_header_rows = [(field, value)]
            current_page = None
            continue

        current_rows.append(_page_row(current_page, field, value))

    if current_rows:
        page_rows.append(current_rows)

    return page_rows


def _join_split_field_value_lines(lines):
    joined_lines = []
    index = 0

    while index < len(lines):
        line = lines[index]

        if line.endswith(":") and index + 1 < len(lines):
            joined_lines.append(f"{line} {lines[index + 1]}")
            index += 2
            continue

        joined_lines.append(line)
        index += 1

    return joined_lines


def _parse_field_value_line(line):
    match = re.match(r"^([^:]{1,80})\s*:\s*(.+)$", line)

    if not match:
        return None

    return (
        match.group(1).strip(),
        match.group(2).strip(),
    )


def _looks_like_new_page_header(field, value):
    if field.lower() == "date":
        return bool(re.match(r"^\d{1,2}/\d{1,2}/\d{2,4}$", value))

    return True


def _page_row(page_number, field, value):
    return {
        "source_page": page_number,
        "field": field,
        "value": value,
    }


def _normalize_text(text):
    return (
        str(text)
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\xa0", " ")
    )


def _strip_markdown_emphasis(value):
    return re.sub(r"\*\*(.*?)\*\*", r"\1", value).strip()


def dataframes_to_excel(dataframes, output_dir, output_filename="converted.xlsx"):
    output_dir = Path(output_dir)
    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_file = output_dir / output_filename

    final = pd.concat(
        dataframes,
        ignore_index=True,
        sort=False,
    )
    final = remove_crossed_out_rows(final)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final.to_excel(
            writer,
            index=False,
            sheet_name="Data",
        )
        worksheet = writer.sheets["Data"]
        right_alignment = Alignment(horizontal="right")

        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = right_alignment

    print("Saved:", output_file)

    return str(output_file)


def remove_crossed_out_rows(df):
    df = df.copy()
    text_df = df.astype(str).apply(lambda column: column.str.strip())
    lowered_df = text_df.apply(lambda column: column.str.lower())

    crossed_out_text = lowered_df.apply(
        lambda row: row.str.contains("crossed out", na=False).any(),
        axis=1,
    )
    large_x_rows = _large_x_placeholder_rows(lowered_df)
    rows_to_drop = crossed_out_text | large_x_rows

    dropped = int(rows_to_drop.sum())

    if dropped:
        print(f"Dropped {dropped} crossed-out/X placeholder rows")

    return df.loc[~rows_to_drop].reset_index(drop=True)


def _large_x_placeholder_rows(lowered_df):
    measurement_columns = [
        column
        for column in ("Total_Height", "Green_Height", "Width_mm", "Flower")
        if column in lowered_df.columns
    ]

    if measurement_columns:
        return lowered_df[measurement_columns].eq("x").all(axis=1)

    non_identity_columns = [
        column
        for column in lowered_df.columns
        if column not in ("Chamber", "Column", "Row", "Species", "Count")
    ]

    if not non_identity_columns:
        return pd.Series(False, index=lowered_df.index)

    return lowered_df[non_identity_columns].eq("x").all(axis=1)


def _clean_dataframe(df):
    df = df.copy()

    df.columns = [
        _clean_cell(column)
        for column in df.columns
    ]

    df = df.map(_clean_cell)
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    return df


def _clean_cell(value):
    if pd.isna(value):
        return value

    return (
        str(value)
        .replace("\xa0", " ")
        .replace("\n", " ")
        .strip()
    )


def convert_pdf_to_excel(pdf_path, output_dir, output_filename="converted.xlsx"):
    pipeline_result = chandra_convert(pdf_path)

    return pipeline_result_to_excel(
        pipeline_result,
        output_dir,
        output_filename=output_filename,
    )
