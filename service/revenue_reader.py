import re
from dataclasses import dataclass
from pathlib import Path

import openpyxl

LIST_SHEET_NAME = "一覧"
HEADER_ROW = 3
ADJUSTED_DEPARTMENT = "内"
# 一覧シートの列位置(0始まり)と、その列にあるべき見出し
COL_DEPARTMENT, COL_ID, COL_NAME, COL_MEDICINE, COL_INJECTION, COL_TOTAL, COL_PATIENTS = 0, 1, 2, 4, 5, 13, 14
EXPECTED_HEADERS = {
    COL_ID: "ID",
    COL_MEDICINE: "クスリ",
    COL_INJECTION: "注射",
    COL_TOTAL: "合計",
    COL_PATIENTS: "外来人数合計",
}


@dataclass(frozen=True)
class DoctorRevenue:
    doctor_id: int
    name: str
    revenue: float
    per_diem: float | None


def parse_target_month(revenue_path: Path) -> int:
    """ファイル名末尾のYYYYMMを対象月として返す"""
    match = re.search(r"(\d{4})(0[1-9]|1[0-2])$", revenue_path.stem)
    if match is None:
        raise ValueError(f"ファイル名から対象月(YYYYMM)を取得できません: {revenue_path.name}")
    return int(match.group(0))


def read_doctor_revenues(revenue_path: Path) -> list[DoctorRevenue]:
    workbook = openpyxl.load_workbook(revenue_path, read_only=True, data_only=True)
    try:
        if LIST_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(f"「{LIST_SHEET_NAME}」シートがありません: {revenue_path.name}")
        rows = list(workbook[LIST_SHEET_NAME].iter_rows(min_row=HEADER_ROW, max_col=COL_PATIENTS + 1, values_only=True))
    finally:
        workbook.close()

    for column, expected in EXPECTED_HEADERS.items():
        if not rows or rows[0][column] != expected:
            raise ValueError(f"「{LIST_SHEET_NAME}」シートの列構成が想定と異なります(見出し「{expected}」)")

    # 「◆マスタなし」や合計行はIDが無いので除外される
    return [_to_doctor_revenue(row) for row in rows[1:] if isinstance(row[COL_ID], int)]


def _to_doctor_revenue(row: tuple) -> DoctorRevenue:
    revenue = row[COL_TOTAL] or 0
    if row[COL_DEPARTMENT] == ADJUSTED_DEPARTMENT:
        # 内科はクスリと注射を1/10に調整する(「内科 (調整済み）」シートの合計と同じ)
        medicine_and_injection = (row[COL_MEDICINE] or 0) + (row[COL_INJECTION] or 0)
        revenue = revenue - medicine_and_injection + medicine_and_injection / 10
        if revenue == int(revenue):
            revenue = int(revenue)
    patients = row[COL_PATIENTS]
    return DoctorRevenue(
        doctor_id=row[COL_ID],
        name=str(row[COL_NAME] or ""),
        revenue=revenue,
        per_diem=revenue / patients if patients else None,
    )
