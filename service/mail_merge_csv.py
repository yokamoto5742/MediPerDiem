import csv
import unicodedata
from pathlib import Path

from service.distribution_csv import format_per_diem, read_distribution_rows, round_per_diem

FISCAL_MONTHS = (4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3)
FULLWIDTH_DIGITS = "０１２３４５６７８９"
MAIL_MERGE_FIXED_HEADERS = ["医師名", "メールアドレス", "見出し"]
KPI_FIXED_HEADERS = ["年月", "平均", "見出し"]
MONTH_HEADERS = [f"{month}月" for month in FISCAL_MONTHS]
HEADING_LABEL = "月  "
VALUE_WIDTH = 12


def fiscal_year(year_month: int) -> int:
    """年月(YYYYMM)が属する年度を返す"""
    year, month = divmod(year_month, 100)
    return year if month >= 4 else year - 1


def align_right(text: str, width: int) -> str:
    """全角を2桁と数えて半角空白で右寄せする"""
    display_width = sum(2 if unicodedata.east_asian_width(char) in "FW" else 1 for char in text)
    return " " * (width - display_width) + text


def month_label(month: int) -> str:
    """1桁の月は全角数字にして、2桁の月(半角)と表示幅をそろえる"""
    return f"{month}月" if month >= 10 else f"{FULLWIDTH_DIGITS[month]}月"


def write_mail_merge_csv(trend_path: Path, recipients_path: Path, output_path: Path) -> None:
    """前年度と今年度を横並びにした月ごとの行を持つ、Mail Merge 差し込み用CSVを出力する"""
    months, doctor_rows, _ = read_distribution_rows(trend_path, recipients_path)
    years = comparison_years(months)

    with open(output_path, "w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(MAIL_MERGE_FIXED_HEADERS + MONTH_HEADERS)
        for _, name, email, per_diems in doctor_rows:
            text_by_month = {int(month): format_per_diem(value) for month, value in zip(months, per_diems)}
            writer.writerow([name, email, heading_line(years)] + month_lines(text_by_month, years))


def write_kpi_mail_merge_csv(trend_path: Path, recipients_path: Path, output_path: Path) -> None:
    """全医師の外来日当円の平均を月ごとに並べた、1行だけの Mail Merge 差し込み用CSVを出力する"""
    months, doctor_rows, _ = read_distribution_rows(trend_path, recipients_path)
    years = comparison_years(months)
    text_by_month = {
        int(month): average_text([per_diems[index] for _, _, _, per_diems in doctor_rows])
        for index, month in enumerate(months)
    }
    latest_year, latest_month = divmod(int(months[-1]), 100)

    with open(output_path, "w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(KPI_FIXED_HEADERS + MONTH_HEADERS)
        writer.writerow([
            f"{latest_year}年{latest_month}月",
            per_diem_text(text_by_month, years[-1], latest_month),
            heading_line(years),
        ] + month_lines(text_by_month, years))


def average_text(per_diems: list) -> str:
    """四捨五入した各医師の値の平均を桁区切りで返す(空欄は除外。Excel の AVERAGE と同じ)"""
    values = [rounded for rounded in map(round_per_diem, per_diems) if rounded is not None]
    return format_per_diem(sum(values) / len(values)) if values else ""


def comparison_years(months: list) -> tuple[int, int]:
    """最新月の属する年度を今年度として、(前年度, 今年度)を返す"""
    current_year = fiscal_year(int(months[-1]))
    return current_year - 1, current_year


def heading_line(years: tuple[int, int]) -> str:
    return HEADING_LABEL + "".join(align_right(f"{year}年度", VALUE_WIDTH) for year in years)


def month_lines(text_by_month: dict[int, str], years: tuple[int, int]) -> list[str]:
    """4月から3月まで、前年度と今年度の値を横並びにした行を返す"""
    return [
        (month_label(month) + "".join(
            align_right(per_diem_text(text_by_month, year, month), VALUE_WIDTH) for year in years)).rstrip()
        for month in FISCAL_MONTHS
    ]


def per_diem_text(text_by_month: dict[int, str], year: int, month: int) -> str:
    """年度と月に対応する外来日当円を「円」付きで返す(値が無ければ空文字)"""
    calendar_year = year if month >= 4 else year + 1
    text = text_by_month.get(calendar_year * 100 + month, "")
    return f"{text}円" if text else ""
