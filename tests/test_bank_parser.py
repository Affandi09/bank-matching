from pathlib import Path
from app.bank_parser import parse_bank_workbook
from app.main import detect_workbook_type

ROOT_DIR = Path(__file__).resolve().parent.parent


def test_parse_bank_snj_multi_sheet() -> None:
    path = ROOT_DIR / "Laporan  Bank SNJ - OKTOBER 2026 (45).xlsx"
    data = path.read_bytes()
    assert detect_workbook_type(data) == "bank"

    res = parse_bank_workbook(data, path.name)
    assert len(res.sheets_processed) == 9
    assert len(res.transactions) > 0
    assert any(t.row_kind == "opening" for t in res.transactions)
    assert not any(t.keterangan.upper() == "TOTAL" for t in res.transactions)
    # Group header "HO Grosir:" should be carried as section, not as a transaction row
    assert not any(t.keterangan == "HO Grosir:" for t in res.transactions)
    assert any(t.section == "HO Grosir:" for t in res.transactions)


def test_parse_bca_snj_daily_sheets() -> None:
    path = ROOT_DIR / "Laporan BCA SNJ - OKTOBER 2026 (51).xlsx"
    data = path.read_bytes()
    assert detect_workbook_type(data) == "bank"

    res = parse_bank_workbook(data, path.name)
    assert len(res.sheets_processed) == 31
    assert len(res.transactions) > 0
    assert all(t.bank_name == "BCA" for t in res.transactions)
    assert all(t.account_no == "628-0910030" for t in res.transactions)
