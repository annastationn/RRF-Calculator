#!/usr/bin/env python3
"""
Automatic RRF calculation and dictionary-driven class summary.

Usage:
    python calculate_rrf.py input.xlsx classes_ready.txt

Optional:
    python calculate_rrf.py input.xlsx classes_ready.txt -o result.xlsx

Dependency:
    pip install openpyxl
"""

from __future__ import annotations

import argparse
import ast
import csv
import math
import re
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.formatting.rule import ColorScaleRule


C_AW = 12.011
H_AW = 1.008
NORMALIZATION_FACTOR = 0.83905


# Summary table configuration.
# The chemical classification itself is NOT hard-coded here.
# It is read from the external classes file via the "summary_group" field.
SUMMARY_SHEET_NAME = "Class summary"
SUMMARY_C_MIN = 7
SUMMARY_C_MAX = 25


class MissingCarbonNumber(Exception):
    """Raised when Group Name does not contain an explicit C number."""
    pass

# Examples accepted:
#   aromatic C10
#   aromatic C19+
#   tetralin C14 ?
#   indane
GROUP_RE = re.compile(
    r"^(?P<class>.+?)(?:\s+C(?P<C>\d+)(?P<suffix>\+|\s*\?)?)?$",
    re.IGNORECASE,
)


def normalize_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip())


def safe_formula_eval(expression: str, C: int) -> float:
    """Evaluate formulas such as 2*C-6 without using eval()."""
    tree = ast.parse(expression, mode="eval")

    def walk(node):
        if isinstance(node, ast.Expression):
            return walk(node.body)

        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return node.value

        if isinstance(node, ast.Name) and node.id == "C":
            return C

        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            value = walk(node.operand)
            return value if isinstance(node.op, ast.UAdd) else -value

        if isinstance(node, ast.BinOp):
            left = walk(node.left)
            right = walk(node.right)

            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            if isinstance(node.op, ast.Div):
                return left / right
            if isinstance(node.op, ast.Pow):
                return left ** right

        raise ValueError(f"Unsupported formula: {expression!r}")

    return float(walk(tree))


def load_classes(path: Path) -> dict[str, dict[str, object]]:
    """Load the external static class dictionary.

    Required columns:
        class;H_formula;summary_group

    Example:
        aromatic;2*C-6;monocyclic aromatics
    """
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")

        if not reader.fieldnames:
            raise ValueError("The classes file is empty.")

        required = {"class", "H_formula", "summary_group"}
        missing = required - set(reader.fieldnames)
        if missing:
            raise ValueError(
                "Missing columns in classes file: " + ", ".join(sorted(missing))
            )

        result = {}

        for line_no, row in enumerate(reader, start=2):
            class_name = normalize_text(row.get("class") or "")
            h_formula = normalize_text(row.get("H_formula") or "")
            summary_group = normalize_text(row.get("summary_group") or "")

            if not class_name or not h_formula or not summary_group:
                raise ValueError(
                    f"Invalid row {line_no} in classes file: "
                    "class, H_formula and summary_group must all be filled."
                )

            key = class_name.casefold()

            if key in result:
                raise ValueError(
                    f"Duplicate class in classes file: {class_name!r}"
                )

            result[key] = {
                "display_name": class_name,
                "H_formula": h_formula,
                "summary_group": summary_group,
            }

    return result


def parse_group_name(
    group_name: object,
    classes: dict[str, dict[str, object]],
) -> tuple[str, int, int, str | None, str]:
    """Parse Group Name and obtain chemistry rules from the static dictionary."""
    text = normalize_text(group_name)
    match = GROUP_RE.fullmatch(text)

    if not match:
        raise ValueError(f"Cannot parse Group Name {text!r}")

    class_name = normalize_text(match.group("class"))
    key = class_name.casefold()

    if key not in classes:
        raise ValueError(
            f"Unknown class {class_name!r}. Add it to the classes file."
        )

    rule = classes[key]

    c_text = match.group("C")
    if c_text:
        C = int(c_text)
    else:
        raise MissingCarbonNumber(
            f"Group Name {text!r} has no explicit C number."
        )

    H_value = safe_formula_eval(str(rule["H_formula"]), C)

    if not math.isclose(H_value, round(H_value), abs_tol=1e-9):
        raise ValueError(
            f"H is not an integer for {text!r}: {H_value}. "
            f"Check H_formula in the classes file."
        )

    H = int(round(H_value))
    suffix = match.group("suffix")
    summary_group = str(rule["summary_group"])

    return class_name, C, H, suffix, summary_group


def find_header_row_and_columns(ws):
    required = {"group name", "volume"}

    for row_idx in range(1, min(ws.max_row, 20) + 1):
        headers = {}
        for cell in ws[row_idx]:
            if cell.value is not None:
                headers[normalize_text(cell.value).casefold()] = cell.column

        if required.issubset(headers):
            return row_idx, headers

    raise ValueError(
        "Could not find a header row containing both 'Group Name' and 'Volume'."
    )


def get_or_create_column(ws, header_row: int, header_name: str, style_source_col: int):
    for cell in ws[header_row]:
        if cell.value is not None and normalize_text(cell.value).casefold() == header_name.casefold():
            return cell.column

    new_col = ws.max_column + 1
    target = ws.cell(header_row, new_col)
    source = ws.cell(header_row, style_source_col)

    target.value = header_name

    # Preserve the style of the original table header.
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.protection = copy(source.protection)
        target.number_format = source.number_format

    return new_col



def create_summary_sheet(
    wb,
    calculated_rows: list[dict],
    total_rrf_v: float,
    summary_columns: list[str],
) -> None:
    """Create a C-by-summary_group table.

    The mapping "chemical class -> summary_group" is read from classes.txt.
    The script does not contain chemical grouping rules.
    """
    if SUMMARY_SHEET_NAME in wb.sheetnames:
        del wb[SUMMARY_SHEET_NAME]

    ws_summary = wb.create_sheet(SUMMARY_SHEET_NAME)

    # Keep at least the historical C7..C25 range and extend it if necessary.
    observed_c = [int(item["C"]) for item in calculated_rows]
    c_min = min([SUMMARY_C_MIN] + observed_c) if observed_c else SUMMARY_C_MIN
    c_max = max([SUMMARY_C_MAX] + observed_c) if observed_c else SUMMARY_C_MAX
    carbon_numbers = list(range(c_min, c_max + 1))

    summary = {
        c: {group: 0.0 for group in summary_columns}
        for c in carbon_numbers
    }

    for item in calculated_rows:
        c = int(item["C"])
        group = str(item["summary_group"])
        w_percent = item["RRF*V"] / total_rrf_v * 100.0
        summary[c][group] += w_percent

    # Header.
    headers = ["C"] + summary_columns
    for col_idx, value in enumerate(headers, start=1):
        cell = ws_summary.cell(row=1, column=col_idx, value=value)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    # Main matrix.
    first_data_row = 2

    for offset, c in enumerate(carbon_numbers):
        row = first_data_row + offset
        ws_summary.cell(row=row, column=1, value=c)

        for col_idx, group in enumerate(summary_columns, start=2):
            value = summary[c][group]
            cell = ws_summary.cell(row=row, column=col_idx)

            if math.isclose(value, 0.0, abs_tol=1e-15):
                cell.value = "-"
            else:
                cell.value = value
                cell.number_format = "0.00"

    # Total by summary group.
    totals_row = first_data_row + len(carbon_numbers) + 1
    total_label = ws_summary.cell(row=totals_row, column=1, value="Total")
    total_label.font = Font(bold=True)

    grand_total = 0.0

    for col_idx, group in enumerate(summary_columns, start=2):
        total = sum(summary[c][group] for c in carbon_numbers)
        grand_total += total

        cell = ws_summary.cell(row=totals_row, column=col_idx, value=total)
        cell.number_format = "0.00"
        cell.font = Font(bold=True)

    grand_total_row = totals_row + 2
    ws_summary.cell(
        row=grand_total_row,
        column=1,
        value="Grand total",
    ).font = Font(bold=True)

    grand_cell = ws_summary.cell(
        row=grand_total_row,
        column=2,
        value=grand_total,
    )
    grand_cell.number_format = "0.00"
    grand_cell.font = Font(bold=True)

    # Heatmap for the main numeric matrix only (not totals).
    # Same logic as the old Excel table:
    # minimum -> green, 50th percentile -> yellow, maximum -> red.
    first_heatmap_col = 2
    last_heatmap_col = 1 + len(summary_columns)
    last_data_row = first_data_row + len(carbon_numbers) - 1

    first_cell = ws_summary.cell(
        row=first_data_row,
        column=first_heatmap_col,
    ).coordinate
    last_cell = ws_summary.cell(
        row=last_data_row,
        column=last_heatmap_col,
    ).coordinate

    heatmap_range = f"{first_cell}:{last_cell}"

    ws_summary.conditional_formatting.add(
        heatmap_range,
        ColorScaleRule(
            start_type="min",
            start_color="63BE7B",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFEB84",
            end_type="max",
            end_color="F8696B",
        ),
    )

    # Basic readability.
    ws_summary.freeze_panes = "B2"
    ws_summary.column_dimensions["A"].width = 14

    for col_idx in range(2, len(headers) + 1):
        column_letter = ws_summary.cell(
            row=1,
            column=col_idx,
        ).column_letter
        ws_summary.column_dimensions[column_letter].width = 28

    ws_summary.row_dimensions[1].height = 42


def calculate_file(input_path: Path, classes_path: Path, output_path: Path) -> None:
    classes = load_classes(classes_path)

    wb = load_workbook(input_path)
    ws = wb.active

    header_row, headers = find_header_row_and_columns(ws)
    group_col = headers["group name"]
    volume_col = headers["volume"]

    col_C = get_or_create_column(ws, header_row, "C", volume_col)
    col_H = get_or_create_column(ws, header_row, "H", volume_col)
    col_formula = get_or_create_column(ws, header_row, "Formula", volume_col)
    col_rrf = get_or_create_column(ws, header_row, "RRF", volume_col)
    col_rrfv = get_or_create_column(ws, header_row, "RRF*V", volume_col)
    col_w = get_or_create_column(ws, header_row, "w, %", volume_col)

    calculated_rows = []
    errors = []
    suffix_rows = []
    skipped_no_c = []

    for row_idx in range(header_row + 1, ws.max_row + 1):
        group_name = ws.cell(row_idx, group_col).value
        volume = ws.cell(row_idx, volume_col).value

        # Completely empty row: ignore it.
        if group_name is None and volume is None:
            continue

        if group_name is None or normalize_text(group_name) == "":
            errors.append(f"row {row_idx}: empty Group Name")
            continue

        try:
            volume = float(volume)
        except (TypeError, ValueError):
            errors.append(f"row {row_idx}: invalid Volume = {volume!r}")
            continue

        try:
            class_name, C, H, suffix, summary_group = parse_group_name(group_name, classes)
        except MissingCarbonNumber:
            skipped_no_c.append((row_idx, normalize_text(group_name)))
            continue
        except Exception as exc:
            errors.append(f"row {row_idx}: {exc}")
            continue

        if suffix:
            suffix_rows.append((row_idx, normalize_text(group_name)))

        rrf = (
            ((C_AW * C) + (H_AW * H))
            * NORMALIZATION_FACTOR
            / (C_AW * C)
        )
        rrf_v = rrf * volume

        calculated_rows.append(
            {
                "row": row_idx,
                "class": class_name,
                "summary_group": summary_group,
                "C": C,
                "H": H,
                "Formula": f"C{C}H{H}",
                "RRF": rrf,
                "RRF*V": rrf_v,
            }
        )

    if errors:
        preview = "\n".join(errors[:30])
        more = "" if len(errors) <= 30 else f"\n... and {len(errors) - 30} more error(s)"
        raise RuntimeError(
            f"Calculation stopped: {len(errors)} row(s) could not be processed.\n"
            f"{preview}{more}"
        )

    total_rrf_v = sum(item["RRF*V"] for item in calculated_rows)
    if total_rrf_v == 0:
        raise RuntimeError("Sum(RRF*V) is zero; w, % cannot be calculated.")

    for item in calculated_rows:
        row_idx = item["row"]

        ws.cell(row_idx, col_C).value = item["C"]
        ws.cell(row_idx, col_H).value = item["H"]
        ws.cell(row_idx, col_formula).value = item["Formula"]
        ws.cell(row_idx, col_rrf).value = item["RRF"]
        ws.cell(row_idx, col_rrfv).value = item["RRF*V"]
        ws.cell(row_idx, col_w).value = item["RRF*V"] / total_rrf_v * 100.0

        ws.cell(row_idx, col_C).number_format = "0"
        ws.cell(row_idx, col_H).number_format = "0"
        ws.cell(row_idx, col_rrf).number_format = "0.0000"
        ws.cell(row_idx, col_rrfv).number_format = "0.00E+00"
        ws.cell(row_idx, col_w).number_format = "0.00"

    # Preserve the order of summary groups as they first appear
    # in the external static dictionary.
    summary_columns = list(dict.fromkeys(
        str(rule["summary_group"]) for rule in classes.values()
    ))

    create_summary_sheet(
        wb,
        calculated_rows,
        total_rrf_v,
        summary_columns,
    )

    wb.save(output_path)

    print(f"Done: {output_path}")
    print(f"Calculated rows: {len(calculated_rows)}")
    print(f"Sum(RRF*V): {total_rrf_v:.6f}")
    print(f"Sum(w, %): {sum(item['RRF*V'] / total_rrf_v * 100 for item in calculated_rows):.10f}")
    print(f"Summary sheet: {SUMMARY_SHEET_NAME}")

    if suffix_rows:
        print(
            f"Warning: {len(suffix_rows)} row(s) contain '+' or '?'. "
            "The numeric value after C was used for the calculation."
        )
        for row_idx, name in suffix_rows[:10]:
            print(f"  row {row_idx}: {name}")
        if len(suffix_rows) > 10:
            print(f"  ... and {len(suffix_rows) - 10} more")


    if skipped_no_c:
        print(
            f"Skipped {len(skipped_no_c)} row(s) without an explicit C number "
            "in Group Name. Their calculated columns were left blank."
        )
        for row_idx, name in skipped_no_c[:10]:
            print(f"  row {row_idx}: {name}")
        if len(skipped_no_c) > 10:
            print(f"  ... and {len(skipped_no_c) - 10} more")


def main():
    parser = argparse.ArgumentParser(
        description="Calculate C, H, molecular formula, RRF, RRF*V, w, % and dictionary-driven class summary."
    )
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("classes", type=Path, help="Class rules text file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .xlsx file. Default: <input>_calculated.xlsx",
    )
    args = parser.parse_args()

    output = args.output
    if output is None:
        output = args.input.with_name(args.input.stem + "_calculated.xlsx")

    calculate_file(args.input, args.classes, output)


if __name__ == "__main__":
    main()
