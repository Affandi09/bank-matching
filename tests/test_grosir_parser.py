from pathlib import Path
from app.db import flatten_invoice_rows
from app.parser import parse_grosir_workbook

GROSIR_DIR = Path(__file__).resolve().parent.parent / "GROSIR"


def test_all_grosir_files_parse_cleanly() -> None:
    files = sorted(GROSIR_DIR.glob("*.xlsx"))
    assert len(files) == 37, f"Expected 37 xlsx files, found {len(files)}"

    total_invoices = 0
    for path in files:
        with path.open("rb") as f:
            res = parse_grosir_workbook(f, path.name)

        assert len(res.invoices) > 0, f"{path.name} produced 0 invoices"
        assert len(res.sheets_processed) > 0, f"{path.name} processed 0 sheets"
        for inv in res.invoices:
            assert inv.branch != "UNKNOWN", f"{path.name}: unknown branch for {inv.invoice_no}"
            assert inv.invoice_no
            assert inv.customer_name
            assert inv.invoice_date is not None
        total_invoices += len(res.invoices)

    assert total_invoices > 0


def test_template_branch_override_on_26_1() -> None:
    path = GROSIR_DIR / "26 (1).xlsx"
    with path.open("rb") as f:
        res = parse_grosir_workbook(f, path.name)

    assert res.invoices[0].branch == "KELAPA GADING"


def test_flatten_invoice_rows() -> None:
    path = GROSIR_DIR / "261026 LAP. GROSIR KRANGGAN.xlsx"
    with path.open("rb") as f:
        res = parse_grosir_workbook(f, path.name)

    flat = flatten_invoice_rows(res.invoices)
    assert len(flat) >= len(res.invoices)
    # First invoice SI.2026.10.00896 has split payment (KAS + Bank BCA) -> 2 rows
    first_inv_rows = [r for r in flat if r[3] == "SI.2026.10.00896"]
    assert len(first_inv_rows) == 2

