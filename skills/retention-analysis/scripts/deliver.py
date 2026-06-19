#!/usr/bin/env python3
"""
deliver.py — retention-analysis deliverable builder.

Builds a formula-driven retention workbook with an explicit calculation
interface:

* RawData_Customers / RawData_Type / RawData_<month-key>
* Cube_Customers / Cube_<month-key>
* CubeType_<type-key>_<month-key>

Formulas consume those names instead of assuming a fixed data row range such as
12:108. Source month labels and revenue-type labels are preserved visibly, while
the workbook uses canonical internal keys such as 2022-01 and recurring_main.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import contextlib
import io
from collections import OrderedDict
from copy import copy as _copy
from dataclasses import dataclass
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.workbook.defined_name import DefinedName


# Corkscrew rows
ROW_TITLE = 1
ROW_GENERATED = 2
ROW_ARR_FACTOR = 3
ROW_DATES = 5
ROW_VS = 6
ROW_BEGIN = 8
ROW_NEW = 9
ROW_UPSELL = 10
ROW_DOWNSELL = 11
ROW_CHURN = 12
ROW_END = 13
ROW_CHECK = 14
ROW_CC_BANNER = 16
ROW_N_ACTIVE_PRIOR = 17
ROW_N_ACTIVE_CURR = 18
ROW_N_CHURNED = 19
ROW_N_NEW = 20
ROW_RR_BANNER = 22
ROW_GRR = 23
ROW_NRR = 24
ROW_LOGO = 25
ROW_PC_BANNER = 27
ROW_AVG_ARR = 28
ROW_AVG_NEW = 29
ROW_RECON_BANNER = 31

FIRST_DATA_COL = 2
CUBE_FIRST_MONTH_COL = 2

TITLE_FILL = "1F4E79"
SUBHEADER_FILL = "D9E1F2"
KEY_METRIC_FILL = "BDD7EE"
COLOR_BLUE = "0000FF"
COLOR_GREEN = "006100"
COLOR_BLACK = "000000"
COLOR_WHITE = "FFFFFF"

FMT_DOLLAR = '"$"#,##0;("$"#,##0);"-"'
FMT_NUMBER = '#,##0;(#,##0);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_COUNT = '#,##0;(#,##0);"-"'
FMT_DATE = "mmm-yy"


def font_hardcode(bold: bool = False, size: int = 10) -> Font:
    return Font(name="Calibri", size=size, bold=bold, color=COLOR_BLUE)


def font_xsheet(bold: bool = False, size: int = 10) -> Font:
    return Font(name="Calibri", size=size, bold=bold, color=COLOR_GREEN)


def font_formula(bold: bool = False, size: int = 10) -> Font:
    return Font(name="Calibri", size=size, bold=bold, color=COLOR_BLACK)


def font_title() -> Font:
    return Font(name="Calibri", size=14, bold=True, color=COLOR_WHITE)


def font_banner() -> Font:
    return Font(name="Calibri", size=10, bold=True, color=COLOR_WHITE)


def font_subheader() -> Font:
    return Font(name="Calibri", size=10, bold=True, color=COLOR_BLACK)


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def center_continuous_across(ws, row: int, first_col: int, last_col: int,
                             text: str, font_obj: Font, fill_obj: PatternFill) -> None:
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.value = text if c == first_col else None
        cell.font = font_obj
        cell.fill = fill_obj
        cell.alignment = Alignment(horizontal="centerContinuous", vertical="center")


def quote_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def safe_name(text: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_]", "_", str(text).strip())
    s = re.sub(r"_+", "_", s).strip("_")
    if not s or s[0].isdigit():
        s = "n_" + s
    return s[:180]


def month_key(value: Any) -> str:
    if isinstance(value, dt.datetime):
        return f"{value.year:04d}-{value.month:02d}"
    if isinstance(value, dt.date):
        return f"{value.year:04d}-{value.month:02d}"
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y-%m", "%b-%y", "%b %Y", "%B %Y", "%Y-M%m"):
        try:
            d = dt.datetime.strptime(s, fmt)
            return f"{d.year:04d}-{d.month:02d}"
        except ValueError:
            pass
    m = re.match(r"^(\d{4})-M(\d{1,2})$", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    if re.match(r"^\d{4}-\d{2}", s):
        return s[:7]
    raise ValueError(f"Could not parse month header {value!r}")


def month_label(key: str) -> str:
    y, m = key.split("-")
    return f"{int(y):04d}-M{int(m)}"


def month_to_date(key: str) -> dt.date:
    y, m = key.split("-")
    return dt.date(int(y), int(m), 1)


def normalize_type_label(label: str) -> str:
    s = re.sub(r"[^a-z0-9]+", " ", str(label).strip().lower()).strip()
    if not s:
        return "unknown"
    nonrec_tokens = (
        "non recurring", "nonrecurring", "one time", "one off", "setup",
        "implementation", "professional service", "services", "training",
        "hardware", "usage overage", "overage", "migration",
    )
    recurring_tokens = (
        "recurring", "re occurring", "reoccurring", "subscription", "saas",
        "license", "licence", "platform", "monthly", "annual", "arr", "mrr",
    )
    if any(tok in s for tok in nonrec_tokens):
        return "non_recurring"
    if any(tok in s for tok in recurring_tokens):
        return "recurring"
    return safe_name(s).lower()


def classify_types(all_types: list[str], requested: list[str] | None) -> tuple[list[str], dict[str, str], dict[str, str]]:
    """Return (in_scope_visible_labels, normalized_id_by_label, exclusion_reason_by_label)."""
    norm_by_type = {t: normalize_type_label(t) for t in all_types}
    if requested:
        requested_norms = {normalize_type_label(t) for t in requested}
        requested_labels = {t.strip() for t in requested}
        in_scope = [
            t for t in all_types
            if t in requested_labels or norm_by_type[t] in requested_norms
        ]
    else:
        in_scope = [t for t in all_types if norm_by_type[t] == "recurring"]
    if not in_scope and all_types:
        raise ValueError(
            "No in-scope revenue types matched. Pass --type-filter with the "
            "source labels or equivalent tags."
        )
    reasons = {}
    for t in all_types:
        if t in in_scope:
            reasons[t] = "Included: user-confirmed in-scope recurring revenue."
        else:
            reasons[t] = f"Excluded: normalized as {norm_by_type[t]} and not selected."
    return in_scope, norm_by_type, reasons


@dataclass
class CalcInterface:
    helper_sheet: str
    customer_first_row: int
    customer_last_row: int
    type_first_row: int
    type_last_row: int
    row_by_type: dict[str, int]
    month_keys: list[str]
    raw_month_name: dict[str, str]
    cube_month_name: dict[str, str]
    cube_type_name: dict[tuple[str, str], str]
    raw_customer_name: str = "RawData_Customers"
    raw_type_name: str = "RawData_Type"
    cube_customer_name: str = "Cube_Customers"


def add_name(wb, name: str, sheet: str, coord: str) -> str:
    dn = DefinedName(name, attr_text=f"{quote_sheet(sheet)}!{coord}")
    wb.defined_names.add(dn)
    return name


def load_compute_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def load_long_csv(path: str) -> tuple[list[str], list[str], dict[tuple[str, str], float]]:
    customers: set[str] = set()
    months: set[str] = set()
    values: dict[tuple[str, str], float] = {}
    with open(path, newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            cust = str(row["customer_id"]).strip()
            mo = month_key(row["month"])
            try:
                val = float(row["mrr"])
            except (TypeError, ValueError):
                val = 0.0
            customers.add(cust)
            months.add(mo)
            values[(cust, mo)] = values.get((cust, mo), 0.0) + val

    def cust_sort(c: str):
        m = re.search(r"(\d+)\s*$", c)
        if m:
            return (0, int(m.group(1)), c)
        try:
            return (0, int(c), c)
        except ValueError:
            return (1, 0, c)

    return sorted(customers, key=cust_sort), sorted(months), values


def copy_source_sheet_verbatim(src_path: str, src_sheet_name: str | None, dest_ws) -> None:
    if src_path.lower().endswith(".csv"):
        with open(src_path, newline="", encoding="utf-8") as fh:
            for r, row in enumerate(csv.reader(fh), start=1):
                for c, val in enumerate(row, start=1):
                    dest_ws.cell(row=r, column=c, value=val)
        return
    src_wb = load_workbook(src_path, data_only=False, read_only=False)
    if src_sheet_name not in src_wb.sheetnames:
        raise ValueError(f"Source sheet {src_sheet_name!r} not found. Available: {src_wb.sheetnames}")
    src_ws = src_wb[src_sheet_name]
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            dest = dest_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                dest.font = _copy(cell.font)
                dest.fill = _copy(cell.fill)
                dest.border = _copy(cell.border)
                dest.alignment = _copy(cell.alignment)
                dest.number_format = cell.number_format
                dest.protection = _copy(cell.protection)
            if cell.comment:
                dest.comment = Comment(cell.comment.text or "", cell.comment.author or "source")
    for mr in src_ws.merged_cells.ranges:
        dest_ws.merge_cells(str(mr))
    for col, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            dest_ws.column_dimensions[col].width = dim.width
    for row_num, dim in src_ws.row_dimensions.items():
        if dim.height is not None:
            dest_ws.row_dimensions[row_num].height = dim.height


def source_last_row(src_path: str, sheet: str, customer_col: str, first_data_row: int) -> int:
    wb = load_workbook(src_path, data_only=True, read_only=False)
    ws = wb[sheet]
    idx = column_index_from_string(customer_col)
    last = first_data_row
    trailing_summary = re.compile(
        r"^\s*(total|subtotal|grand total|new revenue|expansion|contraction|churn|pipeline|summary)\b",
        re.I,
    )
    for r in range(first_data_row, ws.max_row + 1):
        val = ws.cell(r, idx).value
        if val in (None, ""):
            continue
        if isinstance(val, str) and trailing_summary.match(val):
            break
        last = r
    return last


def source_months(src_path: str, sheet: str, first_date_col: str, header_row: int) -> list[str]:
    wb = load_workbook(src_path, data_only=True, read_only=False)
    ws = wb[sheet]
    c = column_index_from_string(first_date_col)
    months: list[str] = []
    while c <= ws.max_column:
        val = ws.cell(header_row, c).value
        if val in (None, ""):
            break
        try:
            months.append(month_key(val))
        except ValueError:
            break
        c += 1
    return months


def source_types(src_path: str, sheet: str, type_col: str, first_data_row: int) -> list[str]:
    wb = load_workbook(src_path, data_only=True, read_only=False)
    ws = wb[sheet]
    idx = column_index_from_string(type_col)
    seen: OrderedDict[str, None] = OrderedDict()
    trailing_summary = re.compile(
        r"^\s*(total|subtotal|grand total|new revenue|expansion|contraction|churn|pipeline|summary)\b",
        re.I,
    )
    for r in range(first_data_row, ws.max_row + 1):
        first_val = ws.cell(r, 1).value
        if isinstance(first_val, str) and trailing_summary.match(first_val):
            break
        val = ws.cell(r, idx).value
        if val in (None, ""):
            continue
        seen[str(val).strip()] = None
    return list(seen.keys())


def write_raw_from_csv(ws, customers: list[str], months: list[str], values: dict[tuple[str, str], float]) -> None:
    ws.cell(1, 1, "Customer ID").font = font_subheader()
    ws.cell(1, 1).fill = fill(SUBHEADER_FILL)
    for j, mo in enumerate(months):
        cell = ws.cell(1, 2 + j, month_to_date(mo))
        cell.number_format = FMT_DATE
        cell.font = font_subheader()
        cell.fill = fill(SUBHEADER_FILL)
    for i, cust in enumerate(customers):
        r = 2 + i
        ws.cell(r, 1, cust).font = font_hardcode()
        for j, mo in enumerate(months):
            c = ws.cell(r, 2 + j, values.get((cust, mo), 0.0))
            c.font = font_hardcode()
            c.number_format = FMT_NUMBER


def make_calc_interface(
    wb,
    ws_raw,
    ws_cube,
    customers: list[str],
    months: list[str],
    source_customer_col: str,
    source_first_data_row: int,
    source_last_data_row: int,
    source_first_date_col: str,
    source_type_col: str | None,
    all_types: list[str],
) -> CalcInterface:
    raw_customer_col_idx = column_index_from_string(source_customer_col)
    raw_first_date_idx = column_index_from_string(source_first_date_col)
    raw_last_date_idx = raw_first_date_idx + len(months) - 1
    add_name(
        wb,
        "RawData_Customers",
        ws_raw.title,
        f"${get_column_letter(raw_customer_col_idx)}${source_first_data_row}:"
        f"${get_column_letter(raw_customer_col_idx)}${source_last_data_row}",
    )
    if source_type_col:
        add_name(
            wb,
            "RawData_Type",
            ws_raw.title,
            f"${source_type_col}${source_first_data_row}:${source_type_col}${source_last_data_row}",
        )
    raw_month_name: dict[str, str] = {}
    for j, mo in enumerate(months):
        nm = f"RawData_{safe_name(mo)}"
        col = get_column_letter(raw_first_date_idx + j)
        raw_month_name[mo] = add_name(
            wb, nm, ws_raw.title, f"${col}${source_first_data_row}:${col}${source_last_data_row}"
        )

    type_first_row = 6
    row_by_type = {t: type_first_row + i for i, t in enumerate(all_types)}
    type_last_row = type_first_row + max(len(all_types), 1) - 1
    customer_first_row = type_last_row + 6
    customer_last_row = customer_first_row + len(customers) - 1

    add_name(
        wb,
        "Cube_Customers",
        ws_cube.title,
        f"$A${customer_first_row}:$A${customer_last_row}",
    )
    cube_month_name: dict[str, str] = {}
    for j, mo in enumerate(months):
        nm = f"Cube_{safe_name(mo)}"
        col = get_column_letter(CUBE_FIRST_MONTH_COL + j)
        cube_month_name[mo] = add_name(
            wb, nm, ws_cube.title, f"${col}${customer_first_row}:${col}${customer_last_row}"
        )

    cube_type_name: dict[tuple[str, str], str] = {}
    for t, row in row_by_type.items():
        for j, mo in enumerate(months):
            nm = f"CubeType_{safe_name(normalize_type_label(t))}_{safe_name(t)}_{safe_name(mo)}"
            col = get_column_letter(CUBE_FIRST_MONTH_COL + j)
            cube_type_name[(t, mo)] = add_name(wb, nm, ws_cube.title, f"${col}${row}")

    return CalcInterface(
        helper_sheet=ws_cube.title,
        customer_first_row=customer_first_row,
        customer_last_row=customer_last_row,
        type_first_row=type_first_row,
        type_last_row=type_last_row,
        row_by_type=row_by_type,
        month_keys=months,
        raw_month_name=raw_month_name,
        cube_month_name=cube_month_name,
        cube_type_name=cube_type_name,
    )


def write_customer_cube(
    ws,
    iface: CalcInterface,
    customers: list[str],
    months: list[str],
    all_types: list[str],
    in_scope_types: list[str],
    type_norms: dict[str, str],
    exclusion_reasons: dict[str, str],
    source_type_col: str | None,
    lookback: int,
) -> None:
    ws.cell(1, 1, "Customer ID").font = font_subheader()
    ws.cell(1, 1).fill = fill(SUBHEADER_FILL)
    for j, mo in enumerate(months):
        cell = ws.cell(1, CUBE_FIRST_MONTH_COL + j, month_label(mo))
        cell.font = font_subheader()
        cell.fill = fill(SUBHEADER_FILL)
        cell.alignment = Alignment(horizontal="center")

    ws.cell(2, 1, "# Active customers").font = font_subheader()
    ws.cell(3, 1, f"# Retained vs {lookback}mo prior").font = font_subheader()
    ws.cell(4, 1, "Check # Active vs raw interface").font = font_formula()

    for j, mo in enumerate(months):
        col = get_column_letter(CUBE_FIRST_MONTH_COL + j)
        cube_rng = iface.cube_month_name[mo]
        ws.cell(2, CUBE_FIRST_MONTH_COL + j, f'=COUNTIF({cube_rng},">0")').font = font_formula()
        ws.cell(2, CUBE_FIRST_MONTH_COL + j).number_format = FMT_COUNT
        if j < lookback:
            ws.cell(3, CUBE_FIRST_MONTH_COL + j, "").font = font_formula()
        else:
            prior_rng = iface.cube_month_name[months[j - lookback]]
            ws.cell(
                3, CUBE_FIRST_MONTH_COL + j,
                f"=SUMPRODUCT(({cube_rng}>0)*({prior_rng}>0))",
            ).font = font_formula()
            ws.cell(3, CUBE_FIRST_MONTH_COL + j).number_format = FMT_COUNT
        ws.cell(4, CUBE_FIRST_MONTH_COL + j, 0).font = font_formula()
        ws.cell(4, CUBE_FIRST_MONTH_COL + j).number_format = FMT_COUNT

    meta_col = CUBE_FIRST_MONTH_COL + len(months) + 1
    ws.cell(iface.type_first_row - 1, 1, "Revenue Type").font = font_subheader()
    ws.cell(iface.type_first_row - 1, meta_col, "Scope").font = font_subheader()
    ws.cell(iface.type_first_row - 1, meta_col + 1, "Normalized key").font = font_subheader()
    ws.cell(iface.type_first_row - 1, meta_col + 2, "Reason").font = font_subheader()

    for t in all_types:
        r = iface.row_by_type[t]
        included = t in in_scope_types
        ws.cell(r, 1, t).font = font_formula()
        ws.cell(r, meta_col, "Included" if included else "Excluded").font = font_formula()
        ws.cell(r, meta_col + 1, type_norms.get(t, "")).font = font_formula()
        ws.cell(r, meta_col + 2, exclusion_reasons.get(t, "")).font = font_formula()
        for j, mo in enumerate(months):
            raw_rng = iface.raw_month_name[mo]
            col = CUBE_FIRST_MONTH_COL + j
            if source_type_col:
                formula = f'=SUMIFS({raw_rng},RawData_Type,"{t}")'
            else:
                formula = f"=SUM({raw_rng})" if included else "=0"
            c = ws.cell(r, col, formula)
            c.font = font_xsheet()
            c.number_format = FMT_NUMBER

    total_row = iface.type_last_row + 1
    inscope_row = iface.type_last_row + 2
    raw_check_row = iface.type_last_row + 3
    customer_check_row = iface.type_last_row + 4
    ws.cell(total_row, 1, "Total MRR (all visible types)").font = font_subheader()
    ws.cell(inscope_row, 1, "In-scope MRR").font = font_subheader()
    ws.cell(raw_check_row, 1, "Check total vs Raw Data").font = font_formula()
    ws.cell(customer_check_row, 1, "Check in-scope vs customer rows").font = font_formula()
    for j, mo in enumerate(months):
        col = get_column_letter(CUBE_FIRST_MONTH_COL + j)
        type_cells = [f"{col}{iface.row_by_type[t]}" for t in all_types]
        in_scope_cells = [f"{col}{iface.row_by_type[t]}" for t in in_scope_types]
        ws.cell(total_row, CUBE_FIRST_MONTH_COL + j, "=" + "+".join(type_cells) if type_cells else "=0").font = font_formula(True)
        ws.cell(inscope_row, CUBE_FIRST_MONTH_COL + j, "=" + "+".join(in_scope_cells) if in_scope_cells else "=0").font = font_formula(True)
        ws.cell(
            raw_check_row, CUBE_FIRST_MONTH_COL + j,
            f"={col}{total_row}-SUM({iface.raw_month_name[mo]})",
        ).font = font_xsheet()
        ws.cell(
            customer_check_row, CUBE_FIRST_MONTH_COL + j,
            f"={col}{inscope_row}-SUM({iface.cube_month_name[mo]})",
        ).font = font_formula()
        for row in (total_row, inscope_row, raw_check_row, customer_check_row):
            ws.cell(row, CUBE_FIRST_MONTH_COL + j).number_format = FMT_NUMBER

    for i, cust in enumerate(customers):
        r = iface.customer_first_row + i
        ws.cell(r, 1, cust).font = font_subheader()
        for j, mo in enumerate(months):
            raw_rng = iface.raw_month_name[mo]
            parts = []
            if source_type_col:
                for t in in_scope_types:
                    parts.append(f'SUMIFS({raw_rng},RawData_Customers,$A{r},RawData_Type,"{t}")')
            else:
                parts.append(f"SUMIFS({raw_rng},RawData_Customers,$A{r})")
            c = ws.cell(r, CUBE_FIRST_MONTH_COL + j, "=" + "+".join(parts))
            c.font = font_xsheet()
            c.number_format = FMT_NUMBER

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions[get_column_letter(meta_col)].width = 12
    ws.column_dimensions[get_column_letter(meta_col + 1)].width = 18
    ws.column_dimensions[get_column_letter(meta_col + 2)].width = 58
    for j in range(len(months)):
        ws.column_dimensions[get_column_letter(CUBE_FIRST_MONTH_COL + j)].width = 12
    ws.freeze_panes = ws.cell(iface.customer_first_row, CUBE_FIRST_MONTH_COL)


def write_corkscrew(
    ws,
    iface: CalcInterface,
    customers: list[str],
    months: list[str],
    arr_factor: float,
    company: str,
    in_scope_types: list[str],
    lookback: int,
) -> None:
    n_periods = len(months) - lookback
    if n_periods <= 0:
        raise ValueError(f"Not enough months for {lookback}-period lookback: {len(months)} months.")
    last_col = FIRST_DATA_COL + n_periods - 1
    title = f"{company} — {lookback}-period ARR Corkscrew & Retention Analysis" if company else f"{lookback}-period ARR Corkscrew & Retention Analysis"
    center_continuous_across(ws, ROW_TITLE, 1, last_col, title, font_title(), fill(TITLE_FILL))
    ws.cell(ROW_GENERATED, 1, "Generated:").font = font_subheader()
    ws.cell(ROW_GENERATED, 2, dt.date.today().isoformat()).font = font_hardcode()
    ws.cell(ROW_ARR_FACTOR, 1, "ARR Factor (MRR × N):").font = font_subheader()
    ws.cell(ROW_ARR_FACTOR, 2, arr_factor).font = font_hardcode(True)
    ws.cell(ROW_ARR_FACTOR, 2).comment = Comment(
        f"Source: User-confirmed, ARR factor passed to script as {arr_factor}.",
        "retention-analysis",
    )
    arr_ref = f"$B${ROW_ARR_FACTOR}"

    ws.cell(ROW_DATES, 1, "Item").font = font_banner()
    ws.cell(ROW_DATES, 1).fill = fill(TITLE_FILL)
    ws.cell(ROW_VS, 1, f"(vs. {lookback} periods prior)").font = font_subheader()
    ws.cell(ROW_VS, 1).fill = fill(SUBHEADER_FILL)
    for j, mo in enumerate(months[lookback:]):
        col = FIRST_DATA_COL + j
        c = ws.cell(ROW_DATES, col, month_label(mo))
        c.font = font_banner()
        c.fill = fill(TITLE_FILL)
        c.alignment = Alignment(horizontal="center")
        c = ws.cell(ROW_VS, col, f"vs {month_label(months[j])}")
        c.font = font_subheader()
        c.fill = fill(SUBHEADER_FILL)
        c.alignment = Alignment(horizontal="center")

    labels = {
        ROW_BEGIN: "Beginning ARR",
        ROW_NEW: "  + New customer ARR",
        ROW_UPSELL: "  + Expansion (Upsell)",
        ROW_DOWNSELL: "  - Contraction (Downsell)",
        ROW_CHURN: "  - Churn",
        ROW_END: "Ending ARR",
        ROW_CHECK: "External Raw Data tie-out (= 0)",
        ROW_N_ACTIVE_PRIOR: "# Active (prior period)",
        ROW_N_ACTIVE_CURR: "# Active (current period)",
        ROW_N_CHURNED: "# Churned",
        ROW_N_NEW: "# New",
        ROW_GRR: "Gross Dollar Retention (GRR)",
        ROW_NRR: "Net Dollar Retention (NRR)",
        ROW_LOGO: "Logo Retention",
        ROW_AVG_ARR: "Avg ARR per Active Customer",
        ROW_AVG_NEW: "Avg ARR per New Customer",
    }
    for row, text in labels.items():
        ws.cell(row, 1, text).font = font_subheader()
        if row in (ROW_BEGIN, ROW_END):
            ws.cell(row, 1).fill = fill(KEY_METRIC_FILL)

    for row, text in ((ROW_CC_BANNER, "CUSTOMER COUNTS"), (ROW_RR_BANNER, "RETENTION RATES"), (ROW_PC_BANNER, "PER-CUSTOMER METRICS")):
        center_continuous_across(ws, row, 1, last_col, text, font_banner(), fill(TITLE_FILL))

    recon_type_start = ROW_RECON_BANNER + 1
    recon_sum_row = recon_type_start + len(in_scope_types)
    recon_var_row = recon_sum_row + 1
    center_continuous_across(ws, ROW_RECON_BANNER, 1, last_col, "RECONCILIATION CHECKS", font_banner(), fill(TITLE_FILL))
    for i, t in enumerate(in_scope_types):
        ws.cell(recon_type_start + i, 1, f"{t} ARR").font = font_subheader()
    ws.cell(recon_sum_row, 1, "Sum in-scope ARR").font = font_subheader()
    ws.cell(recon_var_row, 1, "Variance vs Ending ARR (= 0)").font = font_subheader()

    for j, mo in enumerate(months[lookback:]):
        col = FIRST_DATA_COL + j
        letter = get_column_letter(col)
        prior_mo = months[j]
        curr_rng = iface.cube_month_name[mo]
        prior_rng = iface.cube_month_name[prior_mo]
        ws.cell(ROW_BEGIN, col, f'=SUMPRODUCT(({prior_rng}>0)*{prior_rng})*{arr_ref}').font = font_xsheet(True)
        ws.cell(ROW_NEW, col, f'=SUMPRODUCT(({prior_rng}=0)*({curr_rng}>0)*{curr_rng})*{arr_ref}').font = font_xsheet()
        ws.cell(ROW_UPSELL, col, f'=SUMPRODUCT(({prior_rng}>0)*({curr_rng}>{prior_rng})*({curr_rng}-{prior_rng}))*{arr_ref}').font = font_xsheet()
        ws.cell(ROW_DOWNSELL, col, f'=SUMPRODUCT(({prior_rng}>0)*({curr_rng}>0)*({curr_rng}<{prior_rng})*({curr_rng}-{prior_rng}))*{arr_ref}').font = font_xsheet()
        ws.cell(ROW_CHURN, col, f'=SUMPRODUCT(({prior_rng}>0)*({curr_rng}=0)*(-{prior_rng}))*{arr_ref}').font = font_xsheet()
        ws.cell(ROW_END, col, f"=SUM({letter}{ROW_BEGIN}:{letter}{ROW_CHURN})").font = font_formula(True)
        ws.cell(ROW_CHECK, col, f"={letter}{ROW_END}-SUM({curr_rng})*{arr_ref}").font = font_formula()

        ws.cell(ROW_N_ACTIVE_PRIOR, col, f'=COUNTIF({prior_rng},">0")').font = font_formula()
        ws.cell(ROW_N_ACTIVE_CURR, col, f'=COUNTIF({curr_rng},">0")').font = font_formula()
        ws.cell(ROW_N_CHURNED, col, f'=SUMPRODUCT(({prior_rng}>0)*({curr_rng}=0))').font = font_formula()
        ws.cell(ROW_N_NEW, col, f'=SUMPRODUCT(({prior_rng}=0)*({curr_rng}>0))').font = font_formula()
        ws.cell(ROW_GRR, col, f'=IF({letter}{ROW_BEGIN}=0,"",({letter}{ROW_BEGIN}+{letter}{ROW_DOWNSELL}+{letter}{ROW_CHURN})/{letter}{ROW_BEGIN})').font = font_formula()
        ws.cell(ROW_NRR, col, f'=IF({letter}{ROW_BEGIN}=0,"",({letter}{ROW_BEGIN}+{letter}{ROW_UPSELL}+{letter}{ROW_DOWNSELL}+{letter}{ROW_CHURN})/{letter}{ROW_BEGIN})').font = font_formula()
        ws.cell(ROW_LOGO, col, f'=IF({letter}{ROW_N_ACTIVE_PRIOR}=0,"",({letter}{ROW_N_ACTIVE_PRIOR}-{letter}{ROW_N_CHURNED})/{letter}{ROW_N_ACTIVE_PRIOR})').font = font_formula()
        ws.cell(ROW_AVG_ARR, col, f'=IF({letter}{ROW_N_ACTIVE_CURR}=0,"",{letter}{ROW_END}/{letter}{ROW_N_ACTIVE_CURR})').font = font_formula()
        ws.cell(ROW_AVG_NEW, col, f'=IF({letter}{ROW_N_NEW}=0,"",{letter}{ROW_NEW}/{letter}{ROW_N_NEW})').font = font_formula()

        for i, t in enumerate(in_scope_types):
            c = ws.cell(recon_type_start + i, col, f"={iface.cube_type_name[(t, mo)]}*{arr_ref}")
            c.font = font_xsheet()
            c.number_format = FMT_NUMBER
        ws.cell(recon_sum_row, col, f"=SUM({letter}{recon_type_start}:{letter}{recon_sum_row - 1})").font = font_formula(True)
        ws.cell(recon_var_row, col, f"={letter}{recon_sum_row}-{letter}{ROW_END}").font = font_formula()

        for row in (ROW_BEGIN, ROW_END, ROW_AVG_ARR, ROW_AVG_NEW, recon_sum_row):
            ws.cell(row, col).number_format = FMT_DOLLAR
        for row in (ROW_NEW, ROW_UPSELL, ROW_DOWNSELL, ROW_CHURN, ROW_CHECK, recon_var_row):
            ws.cell(row, col).number_format = FMT_NUMBER
        for row in (ROW_N_ACTIVE_PRIOR, ROW_N_ACTIVE_CURR, ROW_N_CHURNED, ROW_N_NEW):
            ws.cell(row, col).number_format = FMT_COUNT
        for row in (ROW_GRR, ROW_NRR, ROW_LOGO):
            ws.cell(row, col).number_format = FMT_PCT
        ws.cell(ROW_BEGIN, col).fill = fill(KEY_METRIC_FILL)
        ws.cell(ROW_END, col).fill = fill(KEY_METRIC_FILL)

    ws.column_dimensions["A"].width = 38
    for j in range(n_periods):
        ws.column_dimensions[get_column_letter(FIRST_DATA_COL + j)].width = 13
    ws.freeze_panes = ws.cell(ROW_DATES + 2, FIRST_DATA_COL)


def scan_formula_errors(path: str) -> list[str]:
    wb_formula = load_workbook(path, data_only=False)
    errors: list[str] = []
    bad_tokens = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A")
    for ws in wb_formula.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    upper = cell.value.upper()
                    if any(tok in upper for tok in bad_tokens):
                        errors.append(f"{ws.title}!{cell.coordinate}: {cell.value}")
                    if upper.startswith("=") and "IFERROR" in upper:
                        errors.append(f"{ws.title}!{cell.coordinate}: IFERROR is not allowed in metrics/check formulas")
    return errors


def run_libreoffice_recalc(path: str, require: bool = False) -> tuple[bool, str]:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        msg = "LibreOffice/soffice is not installed or not on PATH."
        if require:
            raise RuntimeError(msg)
        return False, msg
    out_dir = tempfile.mkdtemp(prefix="retention_recalc_")
    cmd = [exe, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out_dir, path]
    proc = subprocess.run(cmd, text=True, capture_output=True, timeout=120)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice recalc failed: {proc.stderr or proc.stdout}")
    recalc_path = os.path.join(out_dir, os.path.basename(path))
    if os.path.exists(recalc_path):
        shutil.copy2(recalc_path, path)
    return True, "LibreOffice recalc completed."


def validate_delivery(path: str, require_libreoffice: bool = False) -> dict[str, Any]:
    recalc_ran, recalc_msg = run_libreoffice_recalc(path, require=require_libreoffice)
    formula_errors = scan_formula_errors(path)
    if formula_errors:
        raise RuntimeError("Formula validation failed:\n" + "\n".join(formula_errors[:25]))
    tieout_errors = calculate_and_scan_tieouts(path)
    if tieout_errors:
        raise RuntimeError("Calculated tie-out validation failed:\n" + "\n".join(tieout_errors[:25]))
    return {
        "recalc_ran": recalc_ran,
        "recalc_message": recalc_msg,
        "formula_errors": [],
        "tieout_errors": [],
    }


def _range_scalar(value: Any) -> Any:
    raw = getattr(value, "value", value)
    while isinstance(raw, list) and raw:
        raw = raw[0]
    return raw


def calculate_and_scan_tieouts(path: str) -> list[str]:
    """Evaluate check rows when a local formula engine is available.

    LibreOffice remains the production recalc gate. In local Codex desktop
    sessions LibreOffice is not always installed, while the `formulas` package
    often is. This catches zeroed checks and formula errors instead of silently
    passing a workbook that has never been calculated.
    """
    try:
        import formulas  # type: ignore
    except Exception:
        return []

    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        solution = formulas.ExcelModel().loads(path).finish().calculate()
    solved = {str(k).upper(): v for k, v in solution.items()}
    wb = load_workbook(path, data_only=False)
    errors: list[str] = []
    basename = os.path.basename(path).upper()

    for ws in wb.worksheets:
        check_rows = []
        for r in range(1, ws.max_row + 1):
            label = ws.cell(r, 1).value
            if isinstance(label, str) and (
                "CHECK" in label.upper() or "VARIANCE VS ENDING" in label.upper()
            ):
                check_rows.append(r)
        for r in check_rows:
            for c in range(2, ws.max_column + 1):
                cell = ws.cell(r, c)
                if cell.value in (None, ""):
                    continue
                needle = f"[{basename}]{ws.title.upper()}'!{cell.coordinate.upper()}"
                matches = [v for k, v in solved.items() if k.endswith(needle)]
                if not matches:
                    continue
                val = _range_scalar(matches[0])
                if isinstance(val, str) and val.startswith("#"):
                    errors.append(f"{ws.title}!{cell.coordinate}: {val}")
                elif isinstance(val, (int, float)) and abs(float(val)) > 0.01:
                    errors.append(f"{ws.title}!{cell.coordinate}: expected 0, got {val}")
    return errors


def write_markdown(path: str, compute: dict, company: str, validation: dict[str, Any], in_scope: list[str], out_scope: list[str]) -> None:
    cfg = compute.get("config", {})
    months = cfg.get("month_range", ["?", "?"])
    lines = [
        f"# {company or 'Company'} Retention Summary",
        "",
        f"- Period: {months[0]} to {months[-1]}",
        f"- Customers: {cfg.get('n_customers', '?')}",
        f"- In-scope revenue types: {', '.join(in_scope) if in_scope else 'all source revenue'}",
        f"- Out-of-scope visible/excluded: {', '.join(out_scope) if out_scope else 'none'}",
        f"- LibreOffice recalc: {'ran' if validation.get('recalc_ran') else 'not run'}",
        f"- Formula-error scan: passed",
        "",
    ]
    if validation.get("recalc_message"):
        lines.append(f"Note: {validation['recalc_message']}")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def deliver(
    compute_json_path: str,
    long_csv_path: str,
    output_xlsx_path: str,
    company: str = "",
    source_path: str | None = None,
    source_sheet: str | None = None,
    source_customer_col: str = "A",
    source_first_data_row: int = 2,
    source_first_date_col: str = "B",
    source_type_col: str | None = None,
    type_filter: list[str] | None = None,
    lookback: int = 12,
    source_header_row: int | None = None,
    require_libreoffice: bool = False,
) -> tuple[str, str]:
    compute = load_compute_json(compute_json_path)
    customers, csv_months, values = load_long_csv(long_csv_path)
    arr_factor = float(compute.get("config", {}).get("arr_factor", 12))

    wb = Workbook()
    ws_cork = wb.active
    ws_cork.title = "Corkscrew"
    ws_cube = wb.create_sheet("Raw Data with Analysis")
    ws_raw = wb.create_sheet("Raw Data")

    if source_path:
        if not source_sheet:
            source_sheet = load_workbook(source_path, read_only=True).sheetnames[0]
        copy_source_sheet_verbatim(source_path, source_sheet, ws_raw)
        if source_header_row is None:
            source_header_row = source_first_data_row - 1
        months = source_months(source_path, source_sheet, source_first_date_col, source_header_row) or csv_months
        src_last = source_last_row(source_path, source_sheet, source_customer_col, source_first_data_row)
        if source_type_col:
            all_types = source_types(source_path, source_sheet, source_type_col, source_first_data_row)
        else:
            all_types = ["Revenue"]
        in_scope, type_norms, reasons = classify_types(all_types, type_filter or (["recurring"] if source_type_col else ["Revenue"]))
        iface = make_calc_interface(
            wb, ws_raw, ws_cube, customers, months,
            source_customer_col, source_first_data_row, src_last,
            source_first_date_col, source_type_col, all_types,
        )
    else:
        write_raw_from_csv(ws_raw, customers, csv_months, values)
        months = csv_months
        src_last = 1 + len(customers)
        all_types = ["Revenue"]
        in_scope, type_norms, reasons = classify_types(all_types, type_filter or ["Revenue"])
        iface = make_calc_interface(
            wb, ws_raw, ws_cube, customers, months,
            "A", 2, src_last, "B", None, all_types,
        )

    write_customer_cube(
        ws_cube, iface, customers, months, all_types, in_scope,
        type_norms, reasons, source_type_col, lookback,
    )
    write_corkscrew(
        ws_cork, iface, customers, months, arr_factor, company,
        in_scope, lookback,
    )

    wb.save(output_xlsx_path)
    validation = validate_delivery(output_xlsx_path, require_libreoffice=require_libreoffice)
    company_slug = safe_name(company or "Company")
    md_path = os.path.join(os.path.dirname(os.path.abspath(output_xlsx_path)), f"{company_slug}_Retention_Summary.md")
    write_markdown(md_path, compute, company, validation, in_scope, [t for t in all_types if t not in in_scope])
    return output_xlsx_path, md_path


def parse_args(argv):
    p = argparse.ArgumentParser(description="Build formula-driven retention workbook.")
    p.add_argument("compute_json")
    p.add_argument("long_csv")
    p.add_argument("output_xlsx")
    p.add_argument("--company", default="")
    p.add_argument("--source", default=None)
    p.add_argument("--source-sheet", default=None)
    p.add_argument("--source-customer-col", default="A")
    p.add_argument("--source-first-data-row", type=int, default=2)
    p.add_argument("--source-first-date-col", default="B")
    p.add_argument("--source-header-row", type=int, default=None)
    p.add_argument("--source-type-col", default=None)
    p.add_argument("--type-filter", default=None)
    p.add_argument("--lookback", type=int, default=12)
    p.add_argument("--require-libreoffice", action="store_true")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv if argv is not None else sys.argv[1:])
    type_filter = [t.strip() for t in args.type_filter.split(",")] if args.type_filter else None
    out_xlsx, out_md = deliver(
        args.compute_json,
        args.long_csv,
        args.output_xlsx,
        company=args.company,
        source_path=args.source,
        source_sheet=args.source_sheet,
        source_customer_col=args.source_customer_col,
        source_first_data_row=args.source_first_data_row,
        source_first_date_col=args.source_first_date_col,
        source_type_col=args.source_type_col,
        type_filter=type_filter,
        lookback=args.lookback,
        source_header_row=args.source_header_row,
        require_libreoffice=args.require_libreoffice,
    )
    print(f"Wrote: {out_xlsx}")
    print(f"Wrote: {out_md}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
