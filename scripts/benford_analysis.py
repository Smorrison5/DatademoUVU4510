import argparse
import csv
import json
import math
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

DEFAULT_FILE_CANDIDATES = [Path("je_samples.xlsx"), Path("je_sample.xlsx")]
OUTPUT_DIR = Path("outputs")


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for char in letters:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1


def parse_sheet(path: Path, sheet_path: str = "xl/worksheets/sheet1.xml") -> list[list[str | None]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = []
        shared_path = "xl/sharedStrings.xml"
        if shared_path in archive.namelist():
            shared_root = ET.fromstring(archive.read(shared_path))
            ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            for item in shared_root.findall("main:si", ns):
                text_elem = item.find(".//main:t", ns)
                shared_strings.append(text_elem.text if text_elem is not None else "")

        if sheet_path not in archive.namelist():
            raise FileNotFoundError(f"Sheet XML not found: {sheet_path}")
        sheet_root = ET.fromstring(archive.read(sheet_path))
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


def resolve_default_file() -> Path:
    for candidate in DEFAULT_FILE_CANDIDATES:
        if candidate.exists():
            return candidate
    return DEFAULT_FILE_CANDIDATES[0]


def coerce_numeric(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(str(value).strip())
    except ValueError:
        return None


def leading_digit(value: float) -> int | None:
    if value == 0:
        return None
    magnitude = abs(value)
    while magnitude < 1:
        magnitude *= 10
    while magnitude >= 10:
        magnitude /= 10
    digit = int(magnitude)
    return digit if 1 <= digit <= 9 else None


def expected_benford_counts(total: int) -> dict[int, float]:
    return {digit: total * math.log10(1 + 1 / digit) for digit in range(1, 10)}


def write_svg_chart(path: Path, digits: list[int], observed: list[float], expected: list[float]) -> None:
    width = 900
    height = 500
    margin = 60
    chart_width = width - 2 * margin
    chart_height = height - 2 * margin
    max_value = max(max(observed), max(expected), 0.01)

    def x_pos(index: int) -> float:
        return margin + index * (chart_width / (len(digits) - 1))

    def y_pos(value: float) -> float:
        return height - margin - (value / max_value) * chart_height

    bar_width = chart_width / len(digits) * 0.6
    svg_lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        '<rect width="100%" height="100%" fill="#ffffff"/>',
        f'<text x="{width / 2}" y="{margin / 2}" text-anchor="middle" '
        'font-family="Arial" font-size="18">Benford\'s Law Analysis</text>',
    ]

    for i, digit in enumerate(digits):
        bar_x = margin + i * (chart_width / len(digits)) + (bar_width * 0.2)
        bar_height = (observed[i] / max_value) * chart_height
        bar_y = height - margin - bar_height
        svg_lines.append(
            f'<rect x="{bar_x:.2f}" y="{bar_y:.2f}" width="{bar_width:.2f}" height="{bar_height:.2f}" '
            'fill="#4C78A8" opacity="0.85"/>'
        )
        svg_lines.append(
            f'<text x="{bar_x + bar_width / 2:.2f}" y="{height - margin / 2:.2f}" '
            'text-anchor="middle" font-family="Arial" font-size="12">'
            f"{digit}</text>"
        )

    expected_points = " ".join(f"{x_pos(i):.2f},{y_pos(expected[i]):.2f}" for i in range(len(digits)))
    svg_lines.append(
        f'<polyline points="{expected_points}" fill="none" stroke="#F58518" '
        'stroke-width="2"/>'
    )
    for i in range(len(digits)):
        svg_lines.append(
            f'<circle cx="{x_pos(i):.2f}" cy="{y_pos(expected[i]):.2f}" r="4" fill="#F58518"/>'
        )

    for tick in range(0, 6):
        value = max_value * tick / 5
        y = y_pos(value)
        svg_lines.append(
            f'<line x1="{margin}" y1="{y:.2f}" x2="{width - margin}" y2="{y:.2f}" '
            'stroke="#E0E0E0" stroke-width="1"/>'
        )
        svg_lines.append(
            f'<text x="{margin - 10}" y="{y + 4:.2f}" text-anchor="end" '
            'font-family="Arial" font-size="12">'
            f"{value:.2f}</text>"
        )

    svg_lines.append(
        f'<text x="{width / 2}" y="{height - 10}" text-anchor="middle" '
        'font-family="Arial" font-size="14">Leading Digit</text>'
    )
    svg_lines.append(
        f'<text x="20" y="{height / 2}" text-anchor="middle" font-family="Arial" '
        'font-size="14" transform="rotate(-90 20,{height / 2})">Proportion</text>'
    )
    svg_lines.append("</svg>")

    path.write_text("\n".join(svg_lines))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run Benford's Law analysis on an Excel file.")
    parser.add_argument("--file", type=Path, default=resolve_default_file(), help="Path to .xlsx file")
    parser.add_argument("--column", type=str, default=None, help="Column header to analyze")
    parser.add_argument(
        "--sheet-xml",
        type=str,
        default="xl/worksheets/sheet1.xml",
        help="Sheet XML path inside the workbook (default: sheet1)",
    )
    parser.add_argument("--min-count", type=int, default=10, help="Minimum numeric values to accept a column")
    return parser


def pick_numeric_column(headers: list[str], column_values: dict[str, list[float]], min_count: int) -> str:
    for name in headers:
        if len(column_values.get(name, [])) >= min_count:
            return name
    raise ValueError("No numeric column meets the minimum count threshold.")


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    data_file = args.file
    if not data_file.exists():
        raise FileNotFoundError(f"Excel file not found: {data_file}")

    rows = parse_sheet(data_file, args.sheet_xml)
    if not rows:
        raise RuntimeError("No rows found in worksheet.")

    headers = [str(header) if header is not None else f"Column{idx+1}" for idx, header in enumerate(rows[0])]
    data_rows = rows[1:]
    column_count = len(headers)

    normalized_rows = []
    for row in data_rows:
        normalized = row + [None] * (column_count - len(row))
        normalized_rows.append(normalized[:column_count])

    column_values: dict[str, list[float]] = {name: [] for name in headers}
    for row in normalized_rows:
        for idx, name in enumerate(headers):
            numeric = coerce_numeric(row[idx] if idx < len(row) else None)
            if numeric is not None:
                column_values[name].append(numeric)

    selected_column = args.column
    if selected_column is None:
        selected_column = pick_numeric_column(headers, column_values, args.min_count)

    if selected_column not in column_values:
        raise ValueError(f"Column '{selected_column}' not found in headers: {headers}")

    values = column_values[selected_column]
    leading_digits = [digit for value in values if (digit := leading_digit(value)) is not None]

    if not leading_digits:
        raise ValueError("No valid leading digits found for Benford analysis.")

    total = len(leading_digits)
    observed_counts = {digit: leading_digits.count(digit) for digit in range(1, 10)}
    expected_counts = expected_benford_counts(total)
    observed_percent = {digit: observed_counts[digit] / total for digit in range(1, 10)}
    expected_percent = {digit: expected_counts[digit] / total for digit in range(1, 10)}

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    summary = {
        "file": str(data_file),
        "column": selected_column,
        "total_values": total,
        "observed_counts": observed_counts,
        "expected_counts": {digit: round(expected_counts[digit], 4) for digit in expected_counts},
        "observed_percent": {digit: round(observed_percent[digit], 4) for digit in observed_percent},
        "expected_percent": {digit: round(expected_percent[digit], 4) for digit in expected_percent},
    }

    (OUTPUT_DIR / "benford_summary.json").write_text(json.dumps(summary, indent=2))

    with (OUTPUT_DIR / "benford_summary.csv").open("w", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["digit", "observed_count", "expected_count", "observed_percent", "expected_percent"])
        for digit in range(1, 10):
            writer.writerow(
                [
                    digit,
                    observed_counts[digit],
                    round(expected_counts[digit], 4),
                    round(observed_percent[digit], 4),
                    round(expected_percent[digit], 4),
                ]
            )

    digits = list(range(1, 10))
    observed = [observed_percent[digit] for digit in digits]
    expected = [expected_percent[digit] for digit in digits]

    write_svg_chart(OUTPUT_DIR / "benford_chart.svg", digits, observed, expected)

    summary_md = [
        "# Benford's Law Analysis",
        "",
        f"**File:** `{data_file}`",
        f"**Column:** `{selected_column}`",
        f"**Total values analyzed:** {total}",
        "",
        "Outputs:",
        "- `outputs/benford_summary.json`",
        "- `outputs/benford_summary.csv`",
        "- `outputs/benford_chart.svg`",
    ]
    (OUTPUT_DIR / "benford_summary.md").write_text("\n".join(summary_md))


if __name__ == "__main__":
    main()
