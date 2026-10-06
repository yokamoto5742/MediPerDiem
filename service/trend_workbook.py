import shutil
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from service.revenue_reader import DoctorRevenue

PER_DIEM_SHEET_NAME = "外来日当円"
REVENUE_SHEET_NAME = "外来収益合計"
FIRST_MONTH_COLUMN = 4


def backup_workbook(trend_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = trend_path.with_name(f"{trend_path.stem}_backup_{timestamp}{trend_path.suffix}")
    shutil.copy2(trend_path, backup_path)
    return backup_path


def update_trend_workbook(trend_path: Path, month: int, revenues: list[DoctorRevenue]) -> list[DoctorRevenue]:
    """対象月の列を両シートに書き込み、変化表に行が無い医師を返す"""
    workbook = openpyxl.load_workbook(trend_path)
    per_diem_sheet = workbook[PER_DIEM_SHEET_NAME]
    registered_ids = _write_month_column(per_diem_sheet, month, {r.doctor_id: r.per_diem for r in revenues})
    _write_month_column(workbook[REVENUE_SHEET_NAME], month, {r.doctor_id: r.revenue for r in revenues})
    workbook.save(trend_path)
    return [r for r in revenues if r.doctor_id not in registered_ids]


def _write_month_column(sheet: Worksheet, month: int, value_by_id: dict[int, float | None]) -> set[int]:
    column = _find_month_column(sheet, month)
    letter = get_column_letter(column)
    if sheet.cell(1, column).value is None:
        _copy_column_format(sheet, column - 1, column)
        sheet[f"{letter}1"] = month

    registered_ids: set[int] = set()
    for row in range(2, sheet.max_row + 1):
        doctor_id = sheet.cell(row, 1).value
        if isinstance(doctor_id, int):
            registered_ids.add(doctor_id)
            # 入力に無い医師は空欄にする(同じ月の再実行で前回の値を残さない)
            sheet[f"{letter}{row}"] = value_by_id.get(doctor_id)
    return registered_ids


def _find_month_column(sheet: Worksheet, month: int) -> int:
    """対象月の列を返す。無ければ見出し行の次の空き列を返す"""
    column = FIRST_MONTH_COLUMN
    while sheet.cell(1, column).value not in (None, month):
        column += 1
    return column


def _copy_column_format(sheet: Worksheet, source_column: int, target_column: int) -> None:
    for row in range(1, sheet.max_row + 1):
        sheet.cell(row, target_column)._style = copy(sheet.cell(row, source_column)._style)
    source_width = sheet.column_dimensions[get_column_letter(source_column)].width
    sheet.column_dimensions[get_column_letter(target_column)].width = source_width
