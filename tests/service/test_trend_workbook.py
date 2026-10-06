from pathlib import Path

import openpyxl

from service.revenue_reader import DoctorRevenue
from service.trend_workbook import backup_workbook, update_trend_workbook

NUMBER_FORMAT = "#,##0_ "


def make_trend_workbook(path: Path) -> Path:
    workbook = openpyxl.Workbook()
    default_sheet = workbook.active
    assert default_sheet is not None
    workbook.remove(default_sheet)
    for title in ("外来日当円", "外来収益合計"):
        sheet = workbook.create_sheet(title)
        sheet.append(["ID", "医師名", "診療科", 202606, 202607])
        sheet.append([103, "内科 太郎", "内", 100.5, 200.5])
        sheet.append([107, "眼科 花子", "眼", 300, None])
        for row in (2, 3):
            sheet.cell(row, 5).number_format = NUMBER_FORMAT
        sheet.column_dimensions["E"].width = 11.5
    workbook.save(path)
    return path


def test_appends_month_column_to_both_sheets(tmp_path: Path) -> None:
    path = make_trend_workbook(tmp_path / "trend.xlsx")
    revenues = [DoctorRevenue(103, "内科 太郎", 4096162, 10639.38), DoctorRevenue(107, "眼科 花子", 21240, None)]

    assert update_trend_workbook(path, 202608, revenues) == []

    workbook = openpyxl.load_workbook(path)
    per_diem, revenue = workbook["外来日当円"], workbook["外来収益合計"]
    assert [per_diem.cell(row, 6).value for row in (1, 2, 3)] == [202608, 10639.38, None]
    assert [revenue.cell(row, 6).value for row in (1, 2, 3)] == [202608, 4096162, 21240]
    assert per_diem.cell(2, 5).value == 200.5
    assert per_diem.cell(2, 6).number_format == NUMBER_FORMAT
    assert per_diem.column_dimensions["F"].width == 11.5


def test_rerun_overwrites_same_month_column(tmp_path: Path) -> None:
    path = make_trend_workbook(tmp_path / "trend.xlsx")
    update_trend_workbook(path, 202608, [DoctorRevenue(103, "", 100, 10), DoctorRevenue(107, "", 200, 20)])
    update_trend_workbook(path, 202608, [DoctorRevenue(103, "", 300, 30)])

    sheet = openpyxl.load_workbook(path)["外来日当円"]
    assert [sheet.cell(row, 6).value for row in (1, 2, 3)] == [202608, 30, None]
    assert sheet.cell(1, 7).value is None


def test_returns_doctors_missing_from_workbook(tmp_path: Path) -> None:
    path = make_trend_workbook(tmp_path / "trend.xlsx")
    unknown = DoctorRevenue(901, "健診 次郎", 116821, 5310.05)

    assert update_trend_workbook(path, 202608, [unknown]) == [unknown]
    assert openpyxl.load_workbook(path)["外来日当円"].max_row == 3


def test_backup_keeps_original_content(tmp_path: Path) -> None:
    path = make_trend_workbook(tmp_path / "trend.xlsx")
    backup_path = backup_workbook(path)
    update_trend_workbook(path, 202608, [DoctorRevenue(103, "", 100, 10)])

    assert backup_path.parent == path.parent
    assert openpyxl.load_workbook(backup_path)["外来日当円"].cell(1, 6).value is None
