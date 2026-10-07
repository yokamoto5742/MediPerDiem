import csv
from pathlib import Path

import openpyxl
import pytest

from service.mail_merge_csv import fiscal_year, write_mail_merge_csv


def make_trend_workbook(path: Path, last_month: int) -> Path:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "外来日当円"
    sheet.append(["ID", "医師名", "診療科", 202504, 202603, 202604, last_month])
    sheet.append([103, "内科 太郎", "内", 9876.4, 12000, 123456, None])
    workbook.save(path)
    return path


def read_mail_merge_rows(tmp_path: Path, last_month: int) -> list[dict[str, str]]:
    trend_path = make_trend_workbook(tmp_path / "trend.xlsx", last_month)
    recipients_path = tmp_path / "recipients.csv"
    recipients_path.write_text("ID,メールアドレス\n103,太郎 <t@example.com>\n999,不明 <x@example.com>\n", encoding="utf-8-sig")
    output_path = tmp_path / "out.csv"
    write_mail_merge_csv(trend_path, recipients_path, output_path)
    with open(output_path, encoding="utf-8-sig", newline="") as csv_file:
        return list(csv.DictReader(csv_file))


@pytest.mark.parametrize(("year_month", "expected"), [(202604, 2026), (202612, 2026), (202703, 2026), (202704, 2027)])
def test_fiscal_year(year_month: int, expected: int) -> None:
    assert fiscal_year(year_month) == expected


def test_writes_previous_and_current_year_side_by_side(tmp_path: Path) -> None:
    rows = read_mail_merge_rows(tmp_path, 202610)

    assert len(rows) == 1
    row = rows[0]
    assert list(row)[:3] == ["医師名", "メールアドレス", "見出し"]
    assert list(row)[3:] == ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"]
    assert row["医師名"] == "内科 太郎"
    assert row["メールアドレス"] == "太郎 <t@example.com>"
    assert row["見出し"] == "月      2025年度    2026年度"
    assert row["4月"] == "４月     9,876円   123,456円"
    assert row["10月"] == "10月           -           -"
    assert row["3月"] == "３月    12,000円           -"


def test_switches_years_when_new_fiscal_year_starts(tmp_path: Path) -> None:
    row = read_mail_merge_rows(tmp_path, 202704)[0]

    assert row["見出し"] == "月      2026年度    2027年度"
    assert row["4月"] == "４月   123,456円           -"
