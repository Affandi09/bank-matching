from __future__ import annotations

import re
from datetime import date
from decimal import Decimal
from io import BytesIO
from typing import Any, BinaryIO

import openpyxl
from pydantic import BaseModel, Field

from app.parser import _clean_str, _parse_date, _to_decimal


class BankTransactionRow(BaseModel):
    bank_name: str
    account_no: str
    source_file: str
    source_sheet: str
    period: date | None = None
    tgl: date | None = None
    section: str | None = None
    keterangan: str
    debet: Decimal = Decimal("0.00")
    kredit: Decimal = Decimal("0.00")
    saldo: Decimal | None = None
    row_kind: str = "transaction"
    row_num: int


class BankParseResult(BaseModel):
    filename: str
    transactions: list[BankTransactionRow] = Field(default_factory=list)
    sheets_processed: list[str] = Field(default_factory=list)
    sheets_skipped: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)

    @property
    def total_debet(self) -> Decimal:
        return sum(
            (t.debet for t in self.transactions if t.row_kind == "transaction"),
            Decimal("0.00"),
        )

    @property
    def total_kredit(self) -> Decimal:
        return sum(
            (t.kredit for t in self.transactions if t.row_kind == "transaction"),
            Decimal("0.00"),
        )


def _detect_bank_name(top_rows: list[tuple[Any, ...]], sheet_name: str) -> str:
    joined = " ".join(_clean_str(c) for r in top_rows[:3] for c in r if c).upper()
    for bank in ("MANDIRI", "BCA", "BRI", "DKI", "BNI"):
        if bank in joined:
            return bank
    sheet_upper = sheet_name.upper()
    for bank in ("MANDIRI", "BCA", "BRI", "DKI", "BNI"):
        if bank in sheet_upper:
            return bank
    return "UNKNOWN"


def _extract_account_no(top_rows: list[tuple[Any, ...]], header_idx: int) -> str:
    for r in top_rows[:header_idx]:
        for c in r:
            s = _clean_str(c)
            if not s:
                continue
            upper = s.upper()
            # Match "No. Rek. 129 00 77 33 9998", "No. 0122...", or "628-0910030"
            if "NO." in upper or "REK" in upper or re.match(r"^\d{3}-\d+$", s):
                digits = re.sub(r"[^\d-]", "", s)
                # Strip leading hyphens if any
                digits = digits.strip("-")
                if len(re.sub(r"\D", "", digits)) >= 6:
                    return digits
    return "UNKNOWN"


def _extract_period(top_rows: list[tuple[Any, ...]], header_idx: int) -> date | None:
    for r in top_rows[:header_idx]:
        row_has_per_tgl = any("PER TANGGAL" in _clean_str(c).upper() for c in r)
        for c in r:
            d = _parse_date(c)
            if d and (row_has_per_tgl or True):
                return d
    return None


def _has_numeric_value(val: Any) -> bool:
    if val is None or val == "":
        return False
    if isinstance(val, (int, float, Decimal)):
        return True
    s = str(val).strip().replace(",", "")
    try:
        Decimal(s)
        return True
    except Exception:
        return False


def parse_bank_workbook(file_obj: BinaryIO | bytes, filename: str) -> BankParseResult:
    """Parse a per-bank report Excel workbook into normalized BankTransactionRow items."""
    stream: BinaryIO = BytesIO(file_obj) if isinstance(file_obj, bytes) else file_obj
    wb = openpyxl.load_workbook(stream, read_only=True, data_only=True)
    result = BankParseResult(filename=filename)
    last_account_by_bank: dict[str, str] = {}

    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows_iter = ws.iter_rows(values_only=True)

            top_rows: list[tuple[Any, ...]] = []
            for _ in range(15):
                try:
                    top_rows.append(next(rows_iter))
                except StopIteration:
                    break

            # Find header row containing KETERANGAN and (DEBET or SALDO)
            header_idx: int | None = None
            for idx, r in enumerate(top_rows):
                upper_cells = [_clean_str(c).upper() for c in r]
                if any("KETERANGAN" in c for c in upper_cells):
                    header_idx = idx
                    break

            if header_idx is None:
                result.sheets_skipped.append(sheet_name)
                continue

            header_row = top_rows[header_idx]
            sub_row: tuple[Any, ...] | None = None
            data_start_idx = header_idx + 1
            if header_idx + 1 < len(top_rows):
                cand = top_rows[header_idx + 1]
                cand_upper = [_clean_str(c).upper() for c in cand]
                if any(k in cand_upper for k in ("DEBET", "KREDIT")):
                    sub_row = cand
                    data_start_idx = header_idx + 2

            # Build column indices
            col_tgl: int | None = None
            col_ket: int | None = None
            col_debet: int | None = None
            col_kredit: int | None = None
            col_saldo: int | None = None

            max_cols = max(len(header_row), len(sub_row) if sub_row else 0)
            for idx in range(max_cols):
                m_val = _clean_str(header_row[idx] if idx < len(header_row) else "").upper()
                s_val = _clean_str(sub_row[idx] if sub_row and idx < len(sub_row) else "").upper()
                label = s_val if s_val else m_val
                if label == "TGL":
                    col_tgl = idx
                elif label == "KETERANGAN":
                    col_ket = idx
                elif label == "DEBET":
                    col_debet = idx
                elif label == "KREDIT":
                    col_kredit = idx
                elif label == "SALDO":
                    col_saldo = idx

            if col_ket is None or col_debet is None or col_kredit is None:
                result.sheets_skipped.append(sheet_name)
                continue

            result.sheets_processed.append(sheet_name)
            bank_name = _detect_bank_name(top_rows, sheet_name)
            account_no = _extract_account_no(top_rows, header_idx)
            if account_no != "UNKNOWN":
                last_account_by_bank[bank_name] = account_no
            elif bank_name in last_account_by_bank:
                account_no = last_account_by_bank[bank_name]
            period = _extract_period(top_rows, header_idx)

            current_section: str | None = None
            remaining_rows = list(top_rows[data_start_idx:])
            remaining_rows.extend(rows_iter)

            for offset, r in enumerate(remaining_rows):
                excel_row_num = data_start_idx + offset + 1
                if not r or not any(r):
                    continue

                ket = _clean_str(r[col_ket] if col_ket < len(r) else "")
                if not ket:
                    continue

                ket_upper = ket.upper()
                if ket_upper in ("TOTAL", "JUMLAH", "GRAND TOTAL"):
                    continue

                raw_debet = r[col_debet] if col_debet < len(r) else None
                raw_kredit = r[col_kredit] if col_kredit < len(r) else None
                raw_saldo = r[col_saldo] if col_saldo is not None and col_saldo < len(r) else None

                has_deb = _has_numeric_value(raw_debet)
                has_kre = _has_numeric_value(raw_kredit)
                has_sal = _has_numeric_value(raw_saldo)

                is_opening = "SALDO AWAL" in ket_upper

                # Group/section header row: no debet/kredit/saldo (or ends with ":" with no debet/kredit)
                if not is_opening and not has_deb and not has_kre and not has_sal:
                    if ket.endswith(":") or "PEMBAYARAN" in ket_upper or "PENJUALAN GROSIR" in ket_upper or "HO GROSIR" in ket_upper:
                        current_section = ket
                    # Skip lines without any monetary values
                    continue

                # Also skip group summary rows like "HO Grosir:" or "Penjualan Grosir ...:" that have a subtotal outside DEBET/KREDIT
                if not is_opening and ket.endswith(":") and not has_deb and not has_kre:
                    current_section = ket
                    continue

                tgl_val = _parse_date(r[col_tgl] if col_tgl is not None and col_tgl < len(r) else None)
                debet_val = _to_decimal(raw_debet)
                kredit_val = _to_decimal(raw_kredit)
                saldo_val = _to_decimal(raw_saldo) if has_sal else None

                result.transactions.append(
                    BankTransactionRow(
                        bank_name=bank_name,
                        account_no=account_no,
                        source_file=filename,
                        source_sheet=sheet_name.strip(),
                        period=period,
                        tgl=tgl_val,
                        section=None if is_opening else current_section,
                        keterangan=ket,
                        debet=debet_val,
                        kredit=kredit_val,
                        saldo=saldo_val,
                        row_kind="opening" if is_opening else "transaction",
                        row_num=excel_row_num,
                    )
                )
    finally:
        wb.close()

    return result
