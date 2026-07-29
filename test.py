import json
import os
import re
import time
from difflib import SequenceMatcher
from io import StringIO
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv


load_dotenv()

DATALAB_API_BASE_URL = os.getenv(
    "DATALAB_API_BASE_URL",
    "https://www.datalab.to/api/v1",
).rstrip("/")
DATALAB_CONVERT_MODE = os.getenv("DATALAB_CONVERT_MODE", "accurate")
DATALAB_OUTPUT_FORMAT = os.getenv("DATALAB_OUTPUT_FORMAT", "html")
DATALAB_EXTRAS = os.getenv("DATALAB_EXTRAS", "table_cell_bboxes")
DATALAB_POLL_INTERVAL_SECONDS = int(os.getenv("DATALAB_POLL_INTERVAL_SECONDS", "2"))
DATALAB_TIMEOUT_SECONDS = int(os.getenv("DATALAB_TIMEOUT_SECONDS", "900"))


def chandra_convert(pdf_path):
    """
    Run Datalab Convert with the same options as the playground setup.
    """

    return run_chandra_convert(pdf_path)


def run_chandra_convert(pdf_path):
    api_key = os.getenv("DATALAB_API_KEY")

    if not api_key:
        raise RuntimeError(
            "DATALAB_API_KEY missing. Add it to .env or export it in terminal."
        )

    pdf_path = Path(pdf_path)
    headers = {
        "X-API-Key": api_key,
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    }

    data = {
        "mode": DATALAB_CONVERT_MODE,
        "output_format": DATALAB_OUTPUT_FORMAT,
        "paginate": "true",
        "merge_cross_page": "false",
        "disable_image_captions": "true",
        "disable_image_extraction": "true",
        "skip_cache": "true",
        "additional_config": json.dumps(
            {
                "keep_pageheader_in_output": False,
                "keep_pagefooter_in_output": False,
            }
        ),
    }

    if DATALAB_EXTRAS:
        data["extras"] = DATALAB_EXTRAS

    print("Running Datalab Convert")
    print("Datalab Convert options:", data)

    with open(pdf_path, "rb") as file_handle:
        response = requests.post(
            f"{DATALAB_API_BASE_URL}/convert",
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

    print("Datalab Convert submit status:", response.status_code)

    if not response.ok:
        raise RuntimeError(
            f"Datalab Convert submit failed "
            f"({response.status_code}): {response.text}"
        )

    submission = response.json()
    request_id = submission.get("request_id")

    if not request_id:
        raise RuntimeError(
            f"Datalab Convert response did not include request_id: {submission}"
        )

    result = _poll_convert_request(request_id, headers)

    return {
        "execution": {
            "request_id": request_id,
            "status": result.get("status"),
            "mode": DATALAB_CONVERT_MODE,
            "output_format": DATALAB_OUTPUT_FORMAT,
            "extras": DATALAB_EXTRAS,
        },
        "step_results": [
            {
                "step": {
                    "status": result.get("status"),
                    "type": "convert",
                },
                "result": result,
            }
        ],
    }


def _poll_convert_request(request_id, headers):
    deadline = time.monotonic() + DATALAB_TIMEOUT_SECONDS

    while time.monotonic() < deadline:
        response = requests.get(
            f"{DATALAB_API_BASE_URL}/convert/{request_id}",
            headers=headers,
            timeout=60,
        )

        if not response.ok:
            raise RuntimeError(
                f"Datalab Convert poll failed "
                f"({response.status_code}): {response.text}"
            )

        result = response.json()
        status = result.get("status")

        print("Datalab Convert status:", status)

        if status in ("complete", "completed", "completed_with_errors"):
            result = _hydrate_convert_result(result)
            return result

        if status in ("failed", "error"):
            raise RuntimeError(
                f"Datalab Convert failed: {result}"
            )

        time.sleep(DATALAB_POLL_INTERVAL_SECONDS)

    raise TimeoutError(
        "Timed out waiting for Datalab Convert to finish."
    )


def _hydrate_convert_result(result):
    if result.get("html") or not result.get("result_url"):
        return result

    response = requests.get(
        result["result_url"],
        timeout=60,
    )

    if not response.ok:
        return result

    try:
        downloaded_result = response.json()
    except ValueError:
        return result

    if isinstance(downloaded_result, dict):
        merged_result = dict(result)
        merged_result.update(downloaded_result)
        return merged_result

    return result


def pipeline_result_to_excel(
    pipeline_result,
    output_dir,
    output_filename="converted.xlsx",
    run_metadata=None,
):
    dataframes = _extract_dataframes_from_pipeline_result(pipeline_result)
    return dataframes_to_excel(
        dataframes,
        output_dir,
        output_filename,
        run_metadata=run_metadata,
    )


def save_intermediate_html_outputs(
    pipeline_result,
    output_dir,
    filename_prefix="intermediate",
):
    output_dir = Path(output_dir)
    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    rendered_texts = _collect_pipeline_rendered_texts(pipeline_result)

    html_files = []
    html_documents = []

    for index, rendered_text in enumerate(rendered_texts, start=1):
        html_document = _as_html_document(rendered_text)
        html_file = output_dir / f"{filename_prefix}-{index:02d}.html"
        html_file.write_text(
            html_document,
            encoding="utf-8",
        )
        html_files.append(str(html_file))
        html_documents.append((html_file, html_document))

    default_html = _select_default_html_document(html_documents)

    if default_html:
        default_file = output_dir / "default.html"
        default_file.write_text(
            default_html,
            encoding="utf-8",
        )
        html_files.insert(0, str(default_file))

    return html_files


def _as_html_document(value):
    stripped = str(value).strip()

    if "<html" in stripped.lower():
        return stripped

    return (
        "<!DOCTYPE html>\n"
        "<html>\n"
        "<head><meta charset=\"utf-8\"></head>\n"
        "<body>\n"
        f"{stripped}\n"
        "</body>\n"
        "</html>\n"
    )


def _select_default_html_document(html_documents):
    if not html_documents:
        return None

    return max(
        html_documents,
        key=lambda item: _html_document_score(item[1]),
    )[1]


def _html_document_score(html_document):
    lowered = html_document.lower()
    page_markers = len(
        re.findall(
            r"\bpage\s*:?\s*\d+\s+of\s+\d+",
            html_document,
            re.IGNORECASE,
        )
    )
    real_tables = lowered.count("<table")
    markdown_table_lines = sum(
        1
        for line in html_document.splitlines()
        if line.strip().startswith("|") and line.strip().endswith("|")
    )

    return (
        page_markers * 10000
        + real_tables * 1000
        + markdown_table_lines * 50
        + len(html_document)
    )


def _collect_pipeline_rendered_texts(pipeline_result):
    rendered_texts = []

    for step_result in pipeline_result.get("step_results") or []:
        _collect_rendered_texts(step_result.get("result"), rendered_texts)

    if not rendered_texts:
        _collect_rendered_texts(pipeline_result, rendered_texts)

    return _dedupe_values(rendered_texts)


def _extract_dataframes_from_pipeline_result(pipeline_result):
    rendered_texts = _collect_pipeline_rendered_texts(pipeline_result)

    if rendered_texts:
        default_html = _select_default_html_document(
            [
                (None, _as_html_document(rendered_text))
                for rendered_text in rendered_texts
            ]
        )

        if default_html:
            dataframes = _extract_dataframes_from_value(default_html)

            if dataframes:
                print("Parsing selected default intermediate HTML only")
                return dataframes

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
    previous_columns = None

    for index, table_element in enumerate(table_elements, start=1):
        print("Processing table", index)
        table_element = _mark_crossed_out_cells(table_element)

        table = _best_table_dataframe(table_element, previous_columns)

        if table is not None and is_data_table(table):
            print("Headers:", table.columns.tolist())
            cleaned.append(table)
            previous_columns = table.columns.tolist()

    if not cleaned:
        cleaned = markdown_tables_to_dataframes(html)

    if not cleaned:
        raise RuntimeError("No usable botanical tables found.")

    return cleaned


def markdown_tables_to_dataframes(html):
    markdown_text = _html_to_markdownish_text(html)
    table_blocks = _markdown_table_blocks(markdown_text)
    cleaned = []

    for block in table_blocks:
        table = _markdown_block_to_dataframe(block)

        if table is None:
            continue

        table = _normalize_candidate_table(table)

        if is_data_table(table):
            print("Markdown headers:", table.columns.tolist())
            cleaned.append(table)

    return cleaned


def _html_to_markdownish_text(html):
    html = re.sub(
        r"<del>(.*?)</del>",
        r"\1 crossed out",
        str(html),
        flags=re.IGNORECASE | re.DOTALL,
    )
    html = re.sub(
        r"<br\s*/?>",
        " ",
        html,
        flags=re.IGNORECASE,
    )

    return BeautifulSoup(html, "html.parser").get_text("\n")


def _markdown_table_blocks(text):
    blocks = []
    current = []

    for line in text.splitlines():
        stripped = line.strip()

        if stripped.startswith("|") and stripped.endswith("|"):
            current.append(stripped)
            continue

        if current:
            blocks.append(current)
            current = []

    if current:
        blocks.append(current)

    return blocks


def _markdown_block_to_dataframe(block):
    rows = [
        _split_markdown_row(line)
        for line in block
    ]
    rows = [
        row
        for row in rows
        if not _is_markdown_separator_row(row)
    ]

    if len(rows) < 2:
        return None

    header_rows = []
    data_start = None

    for index, row in enumerate(rows):
        if _markdown_row_starts_data(row):
            data_start = index
            break

        header_rows.append(row)

    if data_start is None or not header_rows:
        return None

    headers = _headers_from_grid_rows(header_rows[:2])
    data_rows = [
        _fit_row_to_width(row, len(headers))
        for row in rows[data_start:]
        if _row_has_meaningful_value(row)
    ]

    if not data_rows:
        return None

    return pd.DataFrame(data_rows, columns=headers)


def _split_markdown_row(line):
    return [
        _clean_markdown_cell(cell)
        for cell in line.strip().strip("|").split("|")
    ]


def _clean_markdown_cell(value):
    value = re.sub(r"\*+", "", str(value))
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def _is_markdown_separator_row(row):
    return all(
        not cell
        or re.fullmatch(r":?-{2,}:?", cell.replace(" ", ""))
        for cell in row
    )


def _markdown_row_starts_data(row):
    if not row:
        return False

    return bool(re.fullmatch(r"\d+", str(row[0]).strip()))


def _mark_crossed_out_cells(table_element):
    table_element = BeautifulSoup(
        str(table_element),
        "html.parser",
    ).find("table")

    for element in table_element.find_all(True):
        style = (element.get("style") or "").lower()
        tag_is_strike = element.name in ("s", "strike", "del")
        style_is_strike = "line-through" in style

        if not tag_is_strike and not style_is_strike:
            continue

        text = element.get_text(" ", strip=True)

        if text:
            element.string = f"{text} crossed out"

    return table_element


def _best_table_dataframe(table_element, previous_columns=None):
    candidates = []

    manual = _manual_table_dataframe(table_element, previous_columns)

    if manual is not None:
        candidates.append(manual)

    for header in ([0, 1], 0, None):
        try:
            table = pd.read_html(
                StringIO(str(table_element)),
                header=header,
            )[0]
        except ValueError:
            continue

        if isinstance(table.columns, pd.MultiIndex):
            table = flatten_headers(table)

        candidates.append(table)

    scored = []

    for candidate in candidates:
        candidate = _normalize_candidate_table(candidate)
        score = _table_quality_score(candidate)

        if score > 0:
            scored.append((score, candidate))

    if not scored:
        return None

    scored.sort(
        key=lambda item: item[0],
        reverse=True,
    )

    return scored[0][1]


def _normalize_candidate_table(table):
    table = clean_headers(table)
    table = coalesce_duplicate_columns(table)
    table = make_unique_columns(table)
    table = normalize_botanical_table(table)

    return table


def _table_quality_score(table):
    if not is_data_table(table):
        return 0

    canonical_columns = sum(
        column in _identity_column_names() | _data_column_names()
        for column in table.columns
    )
    rows_with_measurements = _rows_with_measurements(table)
    artifact_rows = _artifact_row_count(table)
    invalid_values = _invalid_value_count(table)
    unnamed_columns = sum(str(column).startswith("Unnamed") or str(column) == "" for column in table.columns)

    return (
        canonical_columns * 20
        + rows_with_measurements * 5
        + len(table.index)
        - artifact_rows * 50
        - invalid_values * 20
        - unnamed_columns * 10
    )


def _manual_table_dataframe(table_element, previous_columns=None):
    grid = _html_table_grid(table_element)

    if not grid:
        return None

    header_index = _find_header_row_index(grid)

    if header_index is None:
        return _headerless_table_dataframe(grid, previous_columns)

    header_rows = [grid[header_index]]
    next_index = header_index + 1

    if next_index < len(grid) and _looks_like_subheader_row(grid[next_index]):
        header_rows.append(grid[next_index])
        next_index += 1

    headers = _headers_from_grid_rows(header_rows)

    while next_index < len(grid) and _looks_like_subheader_row(grid[next_index]):
        next_index += 1

    rows = [
        _fit_row_to_width(row, len(headers))
        for row in grid[next_index:]
        if _row_has_meaningful_value(row)
    ]

    if not rows:
        return None

    return pd.DataFrame(rows, columns=headers)


def _headerless_table_dataframe(grid, previous_columns):
    if not previous_columns:
        return None

    width = len(previous_columns)
    rows = [
        _fit_row_to_width(row, width)
        for row in grid
        if _row_has_meaningful_value(row)
    ]

    if not rows:
        return None

    if not _rows_match_previous_columns(rows, previous_columns):
        return None

    return pd.DataFrame(rows, columns=previous_columns)


def _rows_match_previous_columns(rows, previous_columns):
    sample_rows = rows[: min(len(rows), 5)]
    column_positions = {
        column: index
        for index, column in enumerate(previous_columns)
    }

    numeric_identity_columns = [
        column
        for column in ("Chamber", "Column", "Subplot", "Count")
        if column in column_positions
    ]

    if numeric_identity_columns:
        numeric_matches = 0
        possible_matches = 0

        for row in sample_rows:
            for column in numeric_identity_columns:
                possible_matches += 1
                value = row[column_positions[column]]

                if _looks_numeric(value):
                    numeric_matches += 1

        if possible_matches and numeric_matches / possible_matches < 0.6:
            return False

    if "Species" in column_positions:
        species_values = [
            row[column_positions["Species"]]
            for row in sample_rows
            if not _is_blank_cell(row[column_positions["Species"]])
        ]

        if species_values and not any(_looks_like_species_code(value) for value in species_values):
            return False

    return True


def _looks_numeric(value):
    if _is_blank_cell(value):
        return False

    return bool(re.match(r"^\d+(?:\.\d+)?$", str(value).strip()))


def _looks_like_species_code(value):
    if _is_blank_cell(value):
        return False

    return bool(re.match(r"^[A-Z]{3,6}$", str(value).strip()))


def _html_table_grid(table_element):
    grid = []
    rowspans = {}

    for row_index, row in enumerate(table_element.find_all("tr")):
        grid_row = []
        column_index = 0

        while (row_index, column_index) in rowspans:
            text, remaining = rowspans.pop((row_index, column_index))
            grid_row.append(text)

            if remaining > 1:
                rowspans[(row_index + 1, column_index)] = (text, remaining - 1)

            column_index += 1

        for cell in row.find_all(["th", "td"]):
            while (row_index, column_index) in rowspans:
                text, remaining = rowspans.pop((row_index, column_index))
                grid_row.append(text)

                if remaining > 1:
                    rowspans[(row_index + 1, column_index)] = (text, remaining - 1)

                column_index += 1

            text = _clean_cell(cell.get_text(" ", strip=True))
            colspan = int(cell.get("colspan") or 1)
            rowspan = int(cell.get("rowspan") or 1)

            for offset in range(colspan):
                grid_row.append(text)

                if rowspan > 1:
                    rowspans[(row_index + 1, column_index + offset)] = (
                        text,
                        rowspan - 1,
                    )

            column_index += colspan

        grid.append(grid_row)

    width = max((len(row) for row in grid), default=0)

    return [
        _fit_row_to_width(row, width)
        for row in grid
    ]


def _find_header_row_index(grid):
    best_index = None
    best_score = 0

    for index, row in enumerate(grid[:5]):
        score = sum(
            _canonical_header(value) in _identity_column_names() | _data_column_names()
            for value in row
        )

        if score > best_score:
            best_index = index
            best_score = score

    if best_score >= 3:
        return best_index

    return None


def _looks_like_subheader_row(row):
    values = [
        _header_match_key(value)
        for value in row
        if not _is_blank_cell(value)
    ]

    if not values:
        return True

    joined = " ".join(values)

    if any(token in joined for token in ("see map", "total", "green", "cm", "mm")):
        return True

    canonical_hits = sum(
        _canonical_header(value) in _identity_column_names() | _data_column_names()
        for value in row
    )

    return canonical_hits >= 2


def _headers_from_grid_rows(header_rows):
    if len(header_rows) == 2:
        aligned_headers = _aligned_split_header_rows(
            header_rows[0],
            header_rows[1],
        )

        if aligned_headers:
            return aligned_headers

    width = max(len(row) for row in header_rows)
    headers = []

    for column_index in range(width):
        parts = []

        for row in header_rows:
            value = row[column_index] if column_index < len(row) else ""

            if _is_blank_cell(value):
                continue

            if parts and parts[-1] == value:
                continue

            parts.append(value)

        headers.append(" ".join(parts).strip())

    return headers


def _aligned_split_header_rows(top_row, bottom_row):
    top_values = [
        _header_match_key(value)
        for value in top_row
    ]
    bottom_values = [
        _header_match_key(value)
        for value in bottom_row
    ]

    if (
        top_values
        and bottom_values
        and top_values[0] == "subplot"
        and len(bottom_values) > 2
        and bottom_values[0] in ("col number", "sub col number")
        and bottom_values[1] in ("row code", "sub code")
    ):
        headers = [
            f"{top_row[0]} {bottom_row[0]}",
            f"{top_row[0]} {bottom_row[1]}",
        ]

        for index in range(1, len(top_row)):
            bottom_index = index + 1
            bottom_value = bottom_row[bottom_index] if bottom_index < len(bottom_row) else ""

            if _is_blank_cell(bottom_value):
                headers.append(top_row[index])
            else:
                headers.append(f"{top_row[index]} {bottom_value}")

        return headers[: max(len(top_row), len(bottom_row))]

    return None


def _fit_row_to_width(row, width):
    row = list(row)

    if len(row) >= width:
        return row[:width]

    return row + [""] * (width - len(row))


def _row_has_meaningful_value(row):
    return any(not _is_blank_cell(value) for value in row)


def _rows_with_measurements(table):
    measurement_columns = [
        column
        for column in ("Count", "Height_cm", "Total_Height", "Green_Height", "Width_mm", "Width_cm", "Flower")
        if column in table.columns
    ]

    if not measurement_columns:
        return 0

    return int(
        table[measurement_columns].apply(
            lambda row: any(not _is_blank_cell(value) for value in row),
            axis=1,
        ).sum()
    )


def _artifact_row_count(table):
    return int(
        table.apply(
            lambda row: any(len(str(value)) > 200 for value in row if not _is_blank_cell(value)),
            axis=1,
        ).sum()
    )


def _invalid_value_count(table):
    invalid = 0

    validators = {
        "Chamber": _looks_numeric,
        "Column": _looks_numeric,
        "Subplot": _looks_numeric,
        "Consec_Number": _looks_numeric,
        "Height_cm": _looks_numeric_or_dash,
        "Total_Height": _looks_numeric_or_dash_or_none_found,
        "Green_Height": _looks_numeric_or_dash,
        "Width_mm": _looks_numeric_or_dash,
        "Width_cm": _looks_numeric_or_dash,
        "Flower": _looks_like_code_cell,
        "Herbivory": _looks_like_code_cell,
    }

    for column, validator in validators.items():
        if column not in table.columns:
            continue

        for value in table[column]:
            if _is_blank_cell(value):
                continue

            if not validator(value):
                invalid += 1

    return invalid


def _looks_numeric_or_dash(value):
    if _looks_numeric(value):
        return True

    return str(value).strip() in {"-", "—", "x", "X"}


def _looks_numeric_or_dash_or_none_found(value):
    if _looks_numeric_or_dash(value):
        return True

    return str(value).strip().lower() == "none found"


def _looks_like_code_cell(value):
    if _is_blank_cell(value):
        return True

    return bool(re.match(r"^[A-Za-z—xX/-]{1,12}$", str(value).strip()))


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
        _canonical_header(column)
        for column in df.columns
    ]

    return df


CANONICAL_HEADER_PATTERNS = {
    "Chamber": [
        "ch",
        "chamber",
        "chamber number",
    ],
    "Column": [
        "col",
        "column",
        "column see map",
        "subplot col number",
    ],
    "Row": [
        "row",
        "subplot row",
    ],
    "Subplot": [
        "subplot",
        "subplots",
        "sub",
        "sub number",
        "subplot number",
        "subplots number",
        "subplot col number",
        "sub col number",
    ],
    "Row_Code": [
        "code",
        "sub code",
        "row code",
        "subplot code",
        "subplots code",
        "subplot sub code",
        "subplots sub code",
        "subplot row code",
    ],
    "Consec_Number": [
        "consec number",
        "consecutive number",
        "consec number per quad",
    ],
    "Species": [
        "species",
        "species see list",
        "plant species",
    ],
    "Count": [
        "count",
        "stem count",
    ],
    "Height_cm": [
        "height",
        "height cm",
    ],
    "Total_Height": [
        "total",
        "total cm",
        "height total",
        "height total cm",
        "total height",
    ],
    "Green_Height": [
        "green",
        "green cm",
        "height green",
        "height green cm",
        "green height",
    ],
    "Width_mm": [
        "width",
        "width mm",
        "width only if greater than 200 cm",
    ],
    "Width_cm": [
        "width cm",
    ],
    "Flower": [
        "flower",
        "flower y n c",
        "flower yes no cut",
    ],
    "Herbivory": [
        "herbivory",
        "herbivory g l a n",
    ],
    "CEVO_Flag": [
        "cevo",
        "cevo flag",
        "cevo flag number",
    ],
}


def _normalize_header(value):
    value = str(value).strip()
    value = re.sub(r"\s+", " ", value)
    value = value.replace(" (", " (")

    return value


def _canonical_header(value):
    original = _normalize_header(value)
    normalized = _header_match_key(original)

    if "consec" in normalized:
        return "Consec_Number"

    if "cevo" in normalized:
        return "CEVO_Flag"

    if "code" in normalized:
        return "Row_Code"

    if "width" in normalized:
        if "cm" in normalized and "mm" not in normalized:
            return "Width_cm"

        return "Width_mm"

    best_name = original
    best_score = 0.0

    for canonical_name, patterns in CANONICAL_HEADER_PATTERNS.items():
        for pattern in patterns:
            score = _header_similarity(normalized, _header_match_key(pattern))

            if score > best_score:
                best_name = canonical_name
                best_score = score

    if best_score >= 0.78:
        return best_name

    return original


def _header_match_key(value):
    value = str(value).lower()
    value = value.replace("#", " number ")
    value = value.replace("?", " ")
    value = value.replace(">", " greater than ")
    value = re.sub(r"\bch\s+\d+\b", "ch", value)
    value = re.sub(r"\([^)]*see list[^)]*\)", " species ", value)
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\b(y|n|c|g|l|a)\b", " ", value)
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def _header_similarity(value, pattern):
    if not value or not pattern:
        return 0.0

    value_tokens = set(value.split())
    pattern_tokens = set(pattern.split())

    if value == pattern:
        return 1.0

    if pattern_tokens and pattern_tokens.issubset(value_tokens):
        return 0.96

    if value_tokens and value_tokens.issubset(pattern_tokens):
        return 0.9

    overlap = 0.0

    if value_tokens and pattern_tokens:
        overlap = len(value_tokens & pattern_tokens) / len(value_tokens | pattern_tokens)

    fuzzy = SequenceMatcher(None, value, pattern).ratio()

    return max(overlap, fuzzy)


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


def coalesce_duplicate_columns(df):
    duplicate_names = [
        column
        for column in dict.fromkeys(df.columns)
        if list(df.columns).count(column) > 1
    ]

    for column in duplicate_names:
        same_name = df.loc[:, df.columns == column]
        coalesced = same_name.bfill(axis=1).iloc[:, 0]
        df = df.drop(columns=[column])
        df[column] = coalesced

    return df


def is_data_table(df):
    column_set = set(df.columns)
    identity_columns = {
        "Chamber",
        "Column",
        "Row",
        "Subplot",
        "Row_Code",
        "Consec_Number",
    }
    data_columns = {
        "Species",
        "Count",
        "Height_cm",
        "Total_Height",
        "Green_Height",
        "Width_mm",
        "Flower",
        "Herbivory",
        "CEVO_Flag",
    }

    identity_matches = len(column_set & identity_columns)
    data_matches = len(column_set & data_columns)

    if "Species" in column_set and data_matches >= 2:
        return True

    return identity_matches >= 1 and data_matches >= 2


def normalize_botanical_table(df):
    df = _clean_dataframe(df)
    df = _drop_empty_named_columns(df)
    df = _drop_artifact_rows(df)
    df = _drop_full_x_rows(df)
    df = _fill_down_continuation_cells(df)
    df = _drop_empty_botanical_rows(df)

    return df.reset_index(drop=True)


def _drop_empty_named_columns(df):
    columns_to_keep = [
        column
        for column in df.columns
        if not _is_blank_cell(column)
    ]

    return df[columns_to_keep]


def _drop_artifact_rows(df):
    artifact_rows = df.apply(
        _looks_like_artifact_row,
        axis=1,
    )

    return df.loc[~artifact_rows]


def _drop_full_x_rows(df):
    x_rows = df.map(_normalize_cell_for_matching).apply(
        lambda row: (
            any(value == "x" for value in row)
            and all(value in ("", "x") for value in row)
        ),
        axis=1,
    )

    return df.loc[~x_rows]


def _looks_like_artifact_row(row):
    values = [
        str(value).strip()
        for value in row
        if not _is_blank_cell(value)
    ]

    if not values:
        return False

    if any(len(value) > 120 for value in values):
        return True

    if any("see map" in value.lower() for value in values):
        return True

    repeated_long_values = len(values) >= 3 and len(set(values)) <= 2
    many_tokens = any(len(value.split()) > 20 for value in values)

    return repeated_long_values and many_tokens


def _fill_down_continuation_cells(df):
    fill_columns = [
        column
        for column in (
            "Chamber",
            "Column",
            "Row",
            "Subplot",
            "Row_Code",
            "Species",
        )
        if column in df.columns
    ]

    for column in fill_columns:
        previous = None
        values = []

        for value in df[column]:
            if _is_continuation_marker(value):
                values.append(previous if previous is not None else value)
                continue

            if not _is_blank_cell(value):
                previous = value

            values.append(value)

        df[column] = values

    return df


def _drop_empty_botanical_rows(df):
    protected_columns = {
        "Chamber",
        "Column",
        "Row",
        "Subplot",
        "Row_Code",
        "Consec_Number",
        "Species",
    }
    value_columns = [
        column
        for column in df.columns
        if column not in protected_columns
    ]

    if not value_columns:
        return df.dropna(how="all")

    has_value = df[value_columns].apply(
        lambda row: any(not _is_blank_cell(value) for value in row),
        axis=1,
    )
    identity_columns = [
        column
        for column in df.columns
        if column in protected_columns and column != "Species"
    ]

    if identity_columns:
        has_identity = df[identity_columns].apply(
            lambda row: any(not _is_blank_cell(value) for value in row),
            axis=1,
        )
    else:
        has_identity = pd.Series(False, index=df.index)

    return df.loc[has_value | has_identity]


def _is_continuation_marker(value):
    return str(value).strip().lower() in {
        "↓",
        "v",
        '"',
        "''",
        "same",
        "ditto",
    }


def _is_blank_cell(value):
    if pd.isna(value):
        return True

    return str(value).strip().lower() in {
        "",
        "nan",
        "none",
        "null",
    }


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


def dataframes_to_excel(
    dataframes,
    output_dir,
    output_filename="converted.xlsx",
    run_metadata=None,
):
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

        if run_metadata:
            pd.DataFrame(
                [
                    {
                        "key": key,
                        "value": value,
                    }
                    for key, value in run_metadata.items()
                ]
            ).to_excel(
                writer,
                index=False,
                sheet_name="_RunInfo",
            )

    print("Saved:", output_file)

    return str(output_file)


def remove_crossed_out_rows(df):
    df = df.copy()
    normalized_df = df.map(_normalize_cell_for_matching)

    crossed_out_text = normalized_df.apply(
        lambda row: row.str.contains("crossed out", na=False).any(),
        axis=1,
    )
    x_placeholder_rows = _x_placeholder_rows(normalized_df)
    rows_to_drop = crossed_out_text | x_placeholder_rows

    dropped = int(rows_to_drop.sum())

    if dropped:
        print(f"Dropped {dropped} crossed-out/X placeholder rows")

    return df.loc[~rows_to_drop].reset_index(drop=True)


def _x_placeholder_rows(normalized_df):
    field_columns = [
        column
        for column in (
            "Count",
            "Height_cm",
            "Total_Height",
            "Green_Height",
            "Width_mm",
            "Width_cm",
            "Flower",
            "Herbivory",
            "CEVO_Flag",
        )
        if column in normalized_df.columns
    ]

    if not field_columns:
        field_columns = [
            column
            for column in normalized_df.columns
            if column not in _identity_column_names()
        ]

    if not field_columns:
        return pd.Series(False, index=normalized_df.index)

    field_df = normalized_df[field_columns]
    has_x = field_df.eq("x").any(axis=1)
    non_blank_values_are_x = field_df.apply(
        lambda row: all(value in ("", "x") for value in row),
        axis=1,
    )

    return has_x & non_blank_values_are_x


def _identity_column_names():
    return {
        "Chamber",
        "Column",
        "Row",
        "Subplot",
        "Row_Code",
        "Consec_Number",
        "Species",
    }


def _data_column_names():
    return {
        "Count",
        "Height_cm",
        "Total_Height",
        "Green_Height",
        "Width_mm",
        "Width_cm",
        "Flower",
        "Herbivory",
        "CEVO_Flag",
    }


def _normalize_cell_for_matching(value):
    if pd.isna(value):
        return ""

    value = str(value).strip().lower()
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    if value in ("nan", "none", "null"):
        return ""

    return value

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


def convert_pdf_to_excel(
    pdf_path,
    output_dir,
    output_filename="converted.xlsx",
    html_output_dir=None,
):
    pipeline_result = chandra_convert(pdf_path)
    pdf_path = Path(pdf_path)
    html_files = []

    if html_output_dir:
        html_files = save_intermediate_html_outputs(
            pipeline_result,
            html_output_dir,
            filename_prefix=pdf_path.stem,
        )

    excel_file = pipeline_result_to_excel(
        pipeline_result,
        output_dir,
        output_filename=output_filename,
        run_metadata={
            "source_pdf": pdf_path.name,
            "output_file": output_filename,
            "datalab_api": "convert",
            "datalab_mode": DATALAB_CONVERT_MODE,
            "datalab_output_format": DATALAB_OUTPUT_FORMAT,
            "datalab_extras": DATALAB_EXTRAS,
            "request_id": (
                pipeline_result.get("execution") or {}
            ).get("request_id"),
        },
    )

    if html_output_dir:
        return {
            "excel_file": excel_file,
            "html_files": html_files,
        }

    return excel_file
