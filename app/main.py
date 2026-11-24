from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Any

import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse

from app.bank_parser import BankParseResult, parse_bank_workbook
from app.config import settings
from app.db import (
    check_db_status,
    flatten_invoice_rows,
    save_bank_transactions,
    save_invoices,
)
from app.parser import ParseResult, _clean_str, parse_grosir_workbook

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "index.html"

app = FastAPI(title="BankMatching - XLSX Uploader", version="1.1.0")


def detect_workbook_type(content: bytes) -> str:
    """Detect whether an .xlsx file is a Grosir sales report or a Bank statement."""
    wb = openpyxl.load_workbook(BytesIO(content), read_only=True, data_only=True)
    try:
        for sheet_name in wb.sheetnames[:3]:
            ws = wb[sheet_name]
            for idx, r in enumerate(ws.iter_rows(values_only=True)):
                if idx >= 15:
                    break
                upper_cells = [_clean_str(c).upper() for c in r]
                if any("NOMOR FAKTUR" in c for c in upper_cells):
                    return "grosir"
                if any("KETERANGAN" in c for c in upper_cells) and any(
                    k in upper_cells for k in ("DEBET", "KREDIT", "SALDO")
                ):
                    return "bank"
    finally:
        wb.close()
    return "unknown"


def _serialize_grosir(res: ParseResult, saved: dict[str, int] | None = None) -> dict[str, Any]:
    branches = sorted({inv.branch for inv in res.invoices})
    flat_rows = flatten_invoice_rows(res.invoices)
    sample_rows = [
        {
            "branch": r[0],
            "invoice_date": r[2].isoformat(),
            "invoice_no": r[3],
            "customer_name": r[4],
            "payment_group": r[5],
            "bank_channel": r[6],
            "payment_amount": float(r[7]),
        }
        for r in flat_rows[:20]
    ]
    payload: dict[str, Any] = {
        "file_type": "grosir",
        "filename": res.filename,
        "branches": branches,
        "invoices_count": len(res.invoices),
        "rows_count": len(flat_rows),
        "sheets_processed": res.sheets_processed,
        "sheets_skipped": res.sheets_skipped,
        "warnings": res.warnings[:20],
        "sample_rows": sample_rows,
    }
    if saved is not None:
        payload["db_saved"] = saved
    return payload


def _serialize_bank(res: BankParseResult, saved: dict[str, int] | None = None) -> dict[str, Any]:
    banks = sorted({t.bank_name for t in res.transactions})
    accounts = sorted({t.account_no for t in res.transactions})
    sample_rows = [
        {
            "bank_name": t.bank_name,
            "account_no": t.account_no,
            "source_sheet": t.source_sheet,
            "period": t.period.isoformat() if t.period else None,
            "section": t.section,
            "keterangan": t.keterangan,
            "debet": float(t.debet),
            "kredit": float(t.kredit),
            "row_kind": t.row_kind,
            "row_num": t.row_num,
        }
        for t in res.transactions[:20]
    ]
    payload: dict[str, Any] = {
        "file_type": "bank",
        "filename": res.filename,
        "banks": banks,
        "accounts": accounts,
        "rows_count": len(res.transactions),
        "total_debet": float(res.total_debet),
        "total_kredit": float(res.total_kredit),
        "sheets_processed": res.sheets_processed,
        "sheets_skipped": res.sheets_skipped,
        "warnings": res.warnings[:20],
        "sample_rows": sample_rows,
    }
    if saved is not None:
        payload["db_saved"] = saved
    return payload


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse(TEMPLATE_PATH.read_text(encoding="utf-8"))


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "target": f"{settings.db_host}:{settings.db_port}/{settings.db_name}.{settings.db_schema}",
        "db": check_db_status(),
    }


@app.post("/api/preview")
async def preview_files(files: list[UploadFile] = File(...)) -> dict[str, Any]:
    results = []
    for f in files:
        if not f.filename or not f.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"File '{f.filename}' bukan .xlsx")
        content = await f.read()
        ftype = detect_workbook_type(content)
        if ftype == "grosir":
            results.append(_serialize_grosir(parse_grosir_workbook(content, f.filename)))
        elif ftype == "bank":
            results.append(_serialize_bank(parse_bank_workbook(content, f.filename)))
        else:
            raise HTTPException(
                status_code=400,
                detail=f"Format '{f.filename}' tidak dikenali (bukan Laporan Grosir maupun Laporan Bank).",
            )
    return {"mode": "preview", "files": results}


@app.post("/api/upload")
async def upload_files(files: list[UploadFile] = File(...)) -> dict[str, Any]:
    results = []
    for f in files:
        if not f.filename or not f.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"File '{f.filename}' bukan .xlsx")
        content = await f.read()
        ftype = detect_workbook_type(content)
        try:
            if ftype == "grosir":
                parsed_g = parse_grosir_workbook(content, f.filename)
                saved = save_invoices(parsed_g.invoices)
                results.append(_serialize_grosir(parsed_g, saved=saved))
            elif ftype == "bank":
                parsed_b = parse_bank_workbook(content, f.filename)
                saved = save_bank_transactions(parsed_b.transactions)
                results.append(_serialize_bank(parsed_b, saved=saved))
            else:
                raise HTTPException(
                    status_code=400,
                    detail=f"Format '{f.filename}' tidak dikenali (bukan Laporan Grosir maupun Laporan Bank).",
                )
        except HTTPException:
            raise
        except Exception as exc:
            raise HTTPException(
                status_code=500,
                detail=f"Gagal simpan '{f.filename}' ke DB ({settings.db_host}): {exc}",
            ) from exc
    return {"mode": "upload", "files": results}
