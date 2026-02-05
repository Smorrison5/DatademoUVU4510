import json
import math
import zipfile
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET


DATA_FILE = Path("je_samples.xlsx")
OUTPUT_DIR = Path("outputs")


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for char in letters:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1


def parse_sheet(path: Path) -> list[list[str | None]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = []
        shared_path = "xl/sharedStrings.xml"
        if shared_path in archive.namelist():
            shared_root = ET.fromstring(archive.read(shared_path))
            ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            for item in shared_root.findall("main:si", ns):
                text_elem = item.find(".//main:t", ns)
                shared_strings.append(text_elem.text if text_elem is not None else "")

        sheet_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rows = []
        for row in sheet_root.findall("main:sheetData/main:row", ns):
            cells = {}
            for cell in row.findall("main:c", ns):
                cell_ref = cell.get("r")
                value_elem = cell.find("main:v", ns)
                if not cell_ref or value_elem is None:
                    continue
                value = value_elem.text
                if cell.get("t") == "s" and value is not None:
                    value = shared_strings[int(value)]
                cells[column_index(cell_ref)] = value
            max_index = max(cells.keys(), default=-1)
            row_values = [None] * (max_index + 1)
            for idx, value in cells.items():
                row_values[idx] = value
            rows.append(row_values)
    return rows


def parse_date(value: str) -> datetime | None:
    formats = [
        "%Y-%m-%d",
        "%Y-%m",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def numeric_stats(values: list[float]) -> dict:
    count = len(values)
    if count == 0:
        return {
            "count": 0,
            "mean": None,
            "std": None,
            "min": None,
            "max": None,
        }
    mean = sum(values) / count
    variance = sum((value - mean) ** 2 for value in values) / (count - 1) if count > 1 else 0.0
    std = math.sqrt(variance)
    return {
        "count": count,
        "mean": mean,
        "std": std,
        "min": min(values),
        "max": max(values),
    }


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = parse_sheet(DATA_FILE)
    if not rows:
        raise RuntimeError("No rows found in worksheet.")

    headers = rows[0]
    data_rows = rows[1:]
    column_count = len(headers)

    normalized_rows = []
    for row in data_rows:
        normalized = row + [None] * (column_count - len(row))
        normalized_rows.append(normalized[:column_count])

    columns = [str(header) if header is not None else f"Column{idx+1}" for idx, header in enumerate(headers)]
    column_values = {name: [] for name in columns}

    for row in normalized_rows:
        for idx, name in enumerate(columns):
            value = row[idx] if idx < len(row) else None
            column_values[name].append(value)

    missing_values = {name: sum(1 for value in values if value in (None, "")) for name, values in column_values.items()}

    date_columns = {}
    for name, values in column_values.items():
        parsed_dates = []
        for value in values:
            if value in (None, ""):
                continue
            parsed = parse_date(str(value))
            if parsed is not None:
                parsed_dates.append(parsed)
        if parsed_dates:
            ratio = len(parsed_dates) / max(1, len([v for v in values if v not in (None, "")]))
            if ratio >= 0.8:
                date_columns[name] = {
                    "min": min(parsed_dates).isoformat(),
                    "max": max(parsed_dates).isoformat(),
                    "non_null_ratio": round(ratio, 4),
                }

    numeric_summary = {}
    for name, values in column_values.items():
        numeric_values = []
        for value in values:
            if value in (None, ""):
                continue
            try:
                numeric_values.append(float(str(value)))
            except ValueError:
                continue
        stats = numeric_stats(numeric_values)
        if stats["count"] > 0:
            numeric_summary[name] = stats

    summary = {
        "file": str(DATA_FILE),
        "row_count": len(data_rows),
        "column_count": column_count,
        "columns": columns,
        "missing_values": missing_values,
        "date_columns": date_columns,
        "numeric_summary": numeric_summary,
    }

    (OUTPUT_DIR / "summary.json").write_text(json.dumps(summary, indent=2))

    numeric_csv_lines = ["column,count,mean,std,min,max"]
    for name, stats in numeric_summary.items():
        numeric_csv_lines.append(
            f"{name},{stats['count']},{stats['mean']},{stats['std']},{stats['min']},{stats['max']}"
        )
    (OUTPUT_DIR / "numeric_summary.csv").write_text("\n".join(numeric_csv_lines))

    summary_md_lines = [
        "# Journal Entry Sample Summary",
        "",
        f"**File:** `{DATA_FILE}`",
        f"**Row count:** {summary['row_count']}",
        f"**Column count:** {summary['column_count']}",
        "",
        "## Columns",
    ]
    for column in columns:
        summary_md_lines.append(f"- {column}")

    summary_md_lines.append("")
    summary_md_lines.append("## Missing Values (Top 10)")
    for column, count in sorted(missing_values.items(), key=lambda item: item[1], reverse=True)[:10]:
        summary_md_lines.append(f"- {column}: {count}")

    summary_md_lines.append("")
    summary_md_lines.append("## Date Ranges")
    if date_columns:
        for column, stats in date_columns.items():
            summary_md_lines.append(
                f"- {column}: {stats['min']} to {stats['max']} (non-null ratio {stats['non_null_ratio']})"
            )
    else:
        summary_md_lines.append("- No date columns detected with >= 80% non-null values.")

    summary_md_lines.append("")
    summary_md_lines.append("## Numeric Summary")
    summary_md_lines.append("See `outputs/numeric_summary.csv` for full descriptive statistics.")

    (OUTPUT_DIR / "summary.md").write_text("\n".join(summary_md_lines))


if __name__ == "__main__":
    main()
