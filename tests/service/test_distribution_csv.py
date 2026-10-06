from pathlib import Path

import openpyxl
import pytest

from service.distribution_csv import format_per_diem, read_recipients, write_distribution_csv


def make_trend_workbook(path: Path) -> Path:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "外来日当円"
    sheet.append(["ID", "医師名", "診療科", 202607, 202608])
    sheet.append([103, "内科 太郎", "内", 11245.85, 10639.38])
    sheet.append([107, "眼科 花子", "眼", None, 0])
    sheet.append([144, "麻酔 三郎", "麻", 3304.5, 999.4])
    workbook.save(path)
    return path


@pytest.mark.parametrize(("value", "expected"), [
    (None, ""), (0, "0"), (999.4, "999"), (3304.5, "3,305"), (2.5, "3"), (10639.38, "10,639"),
])
def test_format_per_diem(value: float | None, expected: str) -> None:
    assert format_per_diem(value) == expected


@pytest.mark.parametrize("encoding", ["utf-8-sig", "cp932"])
def test_read_recipients_accepts_utf8_and_cp932(tmp_path: Path, encoding: str) -> None:
    path = tmp_path / "recipients.csv"
    path.write_text("ID,メールアドレス\n144,麻酔三郎 <m@example.com>\n103,内科太郎 <n@example.com>\n", encoding=encoding)
    assert read_recipients(path) == {144: "麻酔三郎 <m@example.com>", 103: "内科太郎 <n@example.com>"}


def test_read_recipients_requires_headers(tmp_path: Path) -> None:
    path = tmp_path / "recipients.csv"
    path.write_text("ID,mail\n144,m@example.com\n", encoding="utf-8")
    with pytest.raises(ValueError):
        read_recipients(path)


def test_writes_csv_in_recipient_order(tmp_path: Path) -> None:
    trend_path = make_trend_workbook(tmp_path / "trend.xlsx")
    recipients_path = tmp_path / "recipients.csv"
    recipients_path.write_text(
        "ID,メールアドレス\n107,花子 <h@example.com>\n999,不明 <x@example.com>\n103,太郎 <t@example.com>\n",
        encoding="utf-8-sig",
    )
    output_path = tmp_path / "out.csv"

    assert write_distribution_csv(trend_path, recipients_path, output_path) == [999]

    assert output_path.read_bytes() == (
        "﻿ID,医師名,メールアドレス,202607,202608\r\n"
        "107,眼科 花子,花子 <h@example.com>,,0\r\n"
        '103,内科 太郎,太郎 <t@example.com>,"11,246","10,639"\r\n'
    ).encode("utf-8")
