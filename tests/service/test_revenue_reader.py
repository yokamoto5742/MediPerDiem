from pathlib import Path

import openpyxl
import pytest

from service.revenue_reader import parse_target_month, read_doctor_revenues

HEADERS = ["診療科目", "ID", "氏名", "診察", "クスリ", "注射", "処置", "手術", "検査", "画像",
           "その他", "入院料", "自費", "合計", "外来人数合計"]


def make_revenue_workbook(path: Path, rows: list[list[object]], headers: list[str] = HEADERS) -> Path:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "一覧"
    sheet.append([])
    sheet.append([None, None, "入外区分", "外来"])
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


def revenue_row(department: str | None, doctor_id: int | None, medicine: int, injection: int,
                total: int, patients: int | None, name: str = "医師") -> list[object]:
    exam = total - medicine - injection
    return [department, doctor_id, name, exam, medicine, injection, 0, 0, 0, 0, 0, 0, 0, total, patients]


def test_parse_target_month() -> None:
    assert parse_target_month(Path("data/医師別収益データ202608.xlsx")) == 202608


@pytest.mark.parametrize("name", ["医師別収益データ.xlsx", "医師別収益データ202613.xlsx"])
def test_parse_target_month_rejects_invalid_name(name: str) -> None:
    with pytest.raises(ValueError):
        parse_target_month(Path(name))


def test_per_diem_is_total_divided_by_patients(tmp_path: Path) -> None:
    path = make_revenue_workbook(tmp_path / "r.xlsx", [revenue_row("眼", 107, 500, 300, 1000, 3)])
    (doctor,) = read_doctor_revenues(path)
    assert (doctor.doctor_id, doctor.revenue, doctor.department) == (107, 1000, "眼")
    assert doctor.per_diem == pytest.approx(1000 / 3)


def test_internal_medicine_adjusts_medicine_and_injection(tmp_path: Path) -> None:
    path = make_revenue_workbook(tmp_path / "r.xlsx", [revenue_row("内", 103, 500, 300, 1000, 4)])
    (doctor,) = read_doctor_revenues(path)
    # 診察200 + クスリ50 + 注射30
    assert doctor.revenue == 280
    assert doctor.per_diem == 70


@pytest.mark.parametrize("patients", [0, None])
def test_per_diem_is_blank_without_patients(tmp_path: Path, patients: int | None) -> None:
    path = make_revenue_workbook(tmp_path / "r.xlsx", [revenue_row("整", 389, 0, 0, 21240, patients)])
    (doctor,) = read_doctor_revenues(path)
    assert doctor.revenue == 21240
    assert doctor.per_diem is None


def test_rows_without_id_are_skipped(tmp_path: Path) -> None:
    rows = [revenue_row("麻", 144, 0, 0, 100, 1), revenue_row(None, None, 0, 0, 64000, 12)]
    path = make_revenue_workbook(tmp_path / "r.xlsx", rows)
    assert [doctor.doctor_id for doctor in read_doctor_revenues(path)] == [144]


def test_same_name_is_merged_into_youngest_id(tmp_path: Path) -> None:
    rows = [
        revenue_row("耳", 612, 0, 0, 200, 1, name="佐藤　諒"),
        revenue_row("耳", 436, 0, 0, 1000, 3, name="佐藤　諒"),
        revenue_row("眼", 107, 0, 0, 500, 5, name="鈴木　一郎"),
    ]
    path = make_revenue_workbook(tmp_path / "r.xlsx", rows)
    merged, other = read_doctor_revenues(path)
    assert (merged.doctor_id, merged.revenue, merged.per_diem) == (436, 1200, 300)
    assert (other.doctor_id, other.revenue, other.per_diem) == (107, 500, 100)


def test_same_name_ignores_whitespace_and_adjusts_each_row(tmp_path: Path) -> None:
    rows = [
        revenue_row("内", 400, 500, 300, 1000, 4, name="田中 宏明"),
        revenue_row("外", 950, 500, 300, 1000, 1, name="田　中　宏　明"),
    ]
    path = make_revenue_workbook(tmp_path / "r.xlsx", rows)
    (doctor,) = read_doctor_revenues(path)
    # 内科の行だけ調整される: 280 + 1000
    assert (doctor.doctor_id, doctor.name, doctor.revenue) == (400, "田中 宏明", 1280)
    assert doctor.per_diem == 256


def test_health_checkup_rows_are_excluded(tmp_path: Path) -> None:
    rows = [
        revenue_row("内", 400, 0, 0, 1000, 4, name="田中 宏明"),
        revenue_row("健", 923, 0, 0, 470, 1, name="田　中　宏　明"),
        revenue_row("健", 901, 0, 0, 5000, 2, name="本多　正治"),
    ]
    path = make_revenue_workbook(tmp_path / "r.xlsx", rows)
    (doctor,) = read_doctor_revenues(path)
    assert (doctor.doctor_id, doctor.revenue, doctor.per_diem) == (400, 1000, 250)


def test_unexpected_layout_raises(tmp_path: Path) -> None:
    headers = HEADERS[:13] + ["外来人数合計", "合計"]
    path = make_revenue_workbook(tmp_path / "r.xlsx", [], headers)
    with pytest.raises(ValueError):
        read_doctor_revenues(path)
