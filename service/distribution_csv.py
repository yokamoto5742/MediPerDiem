import csv
import io
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl

from service.trend_workbook import FIRST_MONTH_COLUMN, PER_DIEM_SHEET_NAME

RECIPIENT_ID_HEADER = "ID"
RECIPIENT_EMAIL_HEADER = "メールアドレス"
CSV_FIXED_HEADERS = ["ID", "医師名", "メールアドレス"]


def read_recipients(recipients_path: Path) -> dict[int, str]:
    """宛先マスタ(ID・メールアドレス)を記載順のまま返す"""
    raw = recipients_path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp932")
    reader = csv.DictReader(io.StringIO(text))
    if not {RECIPIENT_ID_HEADER, RECIPIENT_EMAIL_HEADER} <= set(reader.fieldnames or []):
        raise ValueError(f"宛先マスタに「{RECIPIENT_ID_HEADER}」「{RECIPIENT_EMAIL_HEADER}」列が必要です: {recipients_path.name}")
    return {int(row[RECIPIENT_ID_HEADER]): row[RECIPIENT_EMAIL_HEADER] for row in reader}


def read_distribution_rows(trend_path: Path, recipients_path: Path) -> tuple[list, list[tuple], list[int]]:
    """対象月の見出し、宛先マスタ順の(ID, 医師名, メールアドレス, 各月の外来日当円)、変化表に行が無い宛先IDを返す"""
    recipients = read_recipients(recipients_path)
    workbook = openpyxl.load_workbook(trend_path, read_only=True, data_only=True)
    try:
        sheet_rows = list(workbook[PER_DIEM_SHEET_NAME].iter_rows(values_only=True))
    finally:
        workbook.close()

    months = [month for month in sheet_rows[0][FIRST_MONTH_COLUMN - 1:] if month is not None]
    last_column = FIRST_MONTH_COLUMN - 1 + len(months)
    row_by_id = {row[0]: row for row in sheet_rows[1:]}

    missing_ids = [doctor_id for doctor_id in recipients if doctor_id not in row_by_id]
    doctor_rows = [
        (doctor_id, row_by_id[doctor_id][1], email, list(row_by_id[doctor_id][FIRST_MONTH_COLUMN - 1:last_column]))
        for doctor_id, email in recipients.items() if doctor_id in row_by_id
    ]
    return months, doctor_rows, missing_ids


def write_distribution_csv(trend_path: Path, recipients_path: Path, output_path: Path) -> list[int]:
    """配信用CSVを出力し、変化表に行が無い宛先IDを返す"""
    months, doctor_rows, missing_ids = read_distribution_rows(trend_path, recipients_path)
    with open(output_path, "w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(CSV_FIXED_HEADERS + months)
        for doctor_id, name, email, per_diems in doctor_rows:
            writer.writerow([doctor_id, name, email] + [format_per_diem(value) for value in per_diems])
    return missing_ids


def round_per_diem(value: object) -> int | None:
    """四捨五入した整数を返す(空欄はNone)"""
    if not isinstance(value, (int, float)):
        return None
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def format_per_diem(value: object) -> str:
    """四捨五入した整数を桁区切りで返す(空欄は空文字)"""
    rounded = round_per_diem(value)
    return "" if rounded is None else f"{rounded:,}"
