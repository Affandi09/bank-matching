from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO

import openpyxl
from pydantic import BaseModel, Field

QUANT_2 = Decimal("0.01")

INDONESIAN_MONTHS: dict[str, int] = {
    "JANUARI": 1,
    "JAN": 1,
    "FEBRUARI": 2,
    "FEB": 2,
    "MARET": 3,
    "MAR": 3,
    "APRIL": 4,
    "APR": 4,
    "MEI": 5,
    "JUNI": 6,
    "JUN": 6,
    "JULI": 7,
    "JUL": 7,
    "AGUSTUS": 8,
    "AGT": 8,
    "AGU": 8,
    "SEPTEMBER": 9,
    "SEP": 9,
    "OKTOBER": 10,
    "OKT": 10,
    "NOVEMBER": 11,
    "NOV": 11,
    "DESEMBER": 12,
    "DES": 12,
}

# Canonical branch mapping (longer/specific keywords first)
BRANCH_ALIASES: list[tuple[str, str]] = [
    ("CIKARANG UTARA", "CIKARANG UTARA"),
    ("CIKUT", "CIKARANG UTARA"),
    ("KELAPA GADING", "KELAPA GADING"),
    ("JATIRAWAMANGUN", "JATIRAWAMANGUN"),
    ("JATIRWAMANGUN", "JATIRAWAMANGUN"),
    ("JTRW", "JATIRAWAMANGUN"),
    ("SOEPOMO TEBET", "TEBET"),
    ("SOEPOMO", "TEBET"),
    ("TEBET", "TEBET"),
    ("YOGYAKARTA", "YOGYAKARTA"),
    ("JOGJA", "YOGYAKARTA"),
    ("PEKANBARU", "PEKANBARU"),
    ("MATRAMAN", "MATRAMAN"),
    ("MTM", "MATRAMAN"),
    ("SEMARANG", "SEMARANG"),
    ("SMG", "SEMARANG"),
    ("SURABAYA", "SURABAYA"),
    ("SBY", "SURABAYA"),
    ("KARAWACI", "KARAWACI"),
    ("KRANGGAN", "KRANGGAN"),
    ("SAWANGAN", "SAWANGAN"),
    ("KOSAMBI", "KOSAMBI"),
    ("KEMANG", "KEMANG"),
    ("KRANJI", "KRANJI"),
    ("BOGOR", "BOGOR"),
    ("DEPOK", "DEPOK"),
    ("MEDAN", "MEDAN"),
    ("BALI", "BALI"),
]

PAYMENT_COLUMN_MAP: dict[str, tuple[str, str]] = {
    "DEPOSIT": ("DEPOSIT", "DEPOSIT"),
    "KAS": ("KAS", "CASH"),
    "KAS/TUNAI": ("KAS", "CASH"),
    "BANK MANDIRI": ("TRANSFER", "MANDIRI"),
    "BANK BCA": ("TRANSFER", "BCA"),
    "BANK BRI": ("TRANSFER", "BRI"),
    "BANK DKI": ("TRANSFER", "DKI"),
    "DEBIT MANDIRI": ("EDC_DEBIT", "MANDIRI"),
    "CC MANDIRI": ("EDC_CC", "MANDIRI"),
    "QRIS MANDIRI": ("QRIS", "MANDIRI"),
    "DEBIT BCA": ("EDC_DEBIT", "BCA"),
    "CC BCA": ("EDC_CC", "BCA"),
    "QRIS BCA": ("QRIS", "BCA"),
    "DEBIT BRI": ("EDC_DEBIT", "BRI"),
    "CC BRI": ("EDC_CC", "BRI"),
    "QRIS BRI": ("QRIS", "BRI"),
}


class PaymentLine(BaseModel):
    payment_group: str
    bank_channel: str
    amount: Decimal


class GrosirInvoice(BaseModel):
    branch: str
    report_date: date | None = None
    invoice_date: date
    invoice_no: str
    customer_name: str
    transaction_amount: Decimal = Decimal("0.00")
    overpay: Decimal = Decimal("0.00")
    voucher: Decimal = Decimal("0.00")
    source_file: str
    sheet_name: str
    row_no: int | None = None
    payment_lines: list[PaymentLine] = Field(default_factory=list)


class ParseResult(BaseModel):
    filename: str
    invoices: list[GrosirInvoice] = Field(default_factory=list)
    sheets_processed: list[str] = Field(default_factory=list)
    sheets_skipped: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)

    @property
    def total_amount(self) -> Decimal:
        return sum((inv.transaction_amount for inv in self.invoices), Decimal("0.00"))


def _clean_str(val: Any) -> str:
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def _to_decimal(val: Any) -> Decimal:
    if val is None or val == "":
        return Decimal("0.00")
    if isinstance(val, Decimal):
        return val.quantize(QUANT_2, rounding=ROUND_HALF_UP)
    if isinstance(val, (int, float)):
        return Decimal(str(val)).quantize(QUANT_2, rounding=ROUND_HALF_UP)
    s = str(val).strip().replace(",", "")
    if not s or s == "-":
        return Decimal("0.00")
    try:
        return Decimal(s).quantize(QUANT_2, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return Decimal("0.00")


def _to_int(val: Any) -> int | None:
    if val is None or val == "":
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return int(val)
    s = str(val).strip()
    if s.isdigit():
        return int(s)
    return None


def _parse_date(val: Any) -> date | None:
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val

    text = _clean_str(val).upper()
    if not text:
        return None

    # Remove day name prefix like "KAMIS, "
    text = re.sub(r"^[A-Z]+,\s*", "", text)

    # Try Indonesian "01 OKTOBER 2026"
    m = re.match(r"^(\d{1,2})\s+([A-Z]+)\s+(\d{4})$", text)
    if m:
        day, month_name, year = int(m.group(1)), m.group(2), int(m.group(3))
        month = INDONESIAN_MONTHS.get(month_name)
        if month:
            try:
                return date(year, month, day)
            except ValueError:
                return None

    # Try ISO / numeric formats
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y"):
        try:
            return datetime.strptime(text.split()[0], fmt).date()
        except ValueError:
            continue
    return None


def _match_branch(text: str) -> str | None:
    upper = text.upper()
    for keyword, canonical in BRANCH_ALIASES:
        if re.search(rf"\b{re.escape(keyword)}\b", upper):
            return canonical
    return None


def _resolve_branch(filename: str, top_rows: list[tuple[Any, ...]], header_idx: int) -> str:
    # 1. Check filename first (most reliable when branch name is in filename)
    from_file = _match_branch(Path(filename).stem)
    if from_file:
        return from_file

    # 2. Check Row 2 (index 1) first line before newline (avoids copy-paste "GROSIR SEMARANG" below it)
    if len(top_rows) > 1 and top_rows[1] and top_rows[1][0]:
        first_line = str(top_rows[1][0]).split("\n")[0]
        from_r2 = _match_branch(first_line)
        if from_r2:
            return from_r2

    # 3. Check row right above header (usually R4 or R5)
    if header_idx > 0 and top_rows[header_idx - 1] and top_rows[header_idx - 1][0]:
        from_above = _match_branch(str(top_rows[header_idx - 1][0]))
        if from_above:
            return from_above

    return "UNKNOWN"


def _resolve_report_date(top_rows: list[tuple[Any, ...]], header_idx: int) -> date | None:
    for r in top_rows[:header_idx]:
        if r and r[0] is not None:
            parsed = _parse_date(r[0])
            if parsed:
                return parsed
    return None


def _build_column_map(
    header_row: tuple[Any, ...], sub_row: tuple[Any, ...] | None
) -> dict[int, str]:
    col_map: dict[int, str] = {}
    max_cols = max(len(header_row), len(sub_row) if sub_row else 0)
    for idx in range(max_cols):
        main_val = _clean_str(header_row[idx] if idx < len(header_row) else "").upper()
        sub_val = _clean_str(sub_row[idx] if sub_row and idx < len(sub_row) else "").upper()
        # Sub-header takes precedence for payment channels (e.g. BANK MANDIRI, DEBIT BCA)
        label = sub_val if sub_val else main_val
        if label:
            col_map[idx] = label
    return col_map


def parse_grosir_workbook(file_obj: BinaryIO | bytes, filename: str) -> ParseResult:
    """Parse a Grosir sales Excel workbook into normalized invoices and payment lines."""
    if isinstance(file_obj, bytes):
        stream: BinaryIO = BytesIO(file_obj)
    else:
        stream = file_obj

    wb = openpyxl.load_workbook(stream, read_only=True, data_only=True)
    result = ParseResult(filename=filename)

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

            # Find header row containing NOMOR FAKTUR and NAMA CUSTOMER
            header_idx: int | None = None
            for idx, r in enumerate(top_rows):
                upper_cells = [_clean_str(c).upper() for c in r]
                if any("NOMOR FAKTUR" in c for c in upper_cells) and any(
                    "NAMA CUSTOMER" in c for c in upper_cells
                ):
                    header_idx = idx
                    break

            if header_idx is None:
                result.sheets_skipped.append(sheet_name)
                continue

            result.sheets_processed.append(sheet_name)
            header_row = top_rows[header_idx]

            # Determine if next row is a sub-header row
            sub_row: tuple[Any, ...] | None = None
            data_start_idx = header_idx + 1
            if header_idx + 1 < len(top_rows):
                candidate = top_rows[header_idx + 1]
                cand_upper = [_clean_str(c).upper() for c in candidate]
                if any(k in cand_upper for k in ("BANK MANDIRI", "BANK BCA", "DEBIT MANDIRI", "KAS", "KAS/TUNAI")):
                    sub_row = candidate
                    data_start_idx = header_idx + 2

            col_map = _build_column_map(header_row, sub_row)
            idx_by_label = {label: idx for idx, label in col_map.items()}

            col_no = idx_by_label.get("NO")
            col_date = idx_by_label.get("TANGGAL FAKTUR")
            col_inv = idx_by_label.get("NOMOR FAKTUR")
            col_cust = idx_by_label.get("NAMA CUSTOMER")
            col_amount = idx_by_label.get("NILAI TRANSAKSI")
            col_overpay = idx_by_label.get("OVERPAY")
            col_voucher = idx_by_label.get("VOUCHER")

            if col_inv is None or col_cust is None:
                result.warnings.append(f"Sheet '{sheet_name}': kolom wajib tidak lengkap.")
                continue

            branch = _resolve_branch(filename, top_rows, header_idx)
            report_date = _resolve_report_date(top_rows, header_idx)

            # Combine remaining buffered top_rows + rest of sheet
            remaining_rows = list(top_rows[data_start_idx:])
            remaining_rows.extend(rows_iter)

            for r in remaining_rows:
                if not r or not any(r):
                    continue

                inv_no = _clean_str(r[col_inv] if col_inv < len(r) else "")
                cust_name = _clean_str(r[col_cust] if col_cust < len(r) else "")
                no_raw = _clean_str(r[col_no] if col_no is not None and col_no < len(r) else "")

                # Skip total/summary/blank rows
                if not inv_no or not cust_name:
                    continue
                upper_check = f"{no_raw} {inv_no} {cust_name}".upper()
                if any(word in upper_check for word in ("TOTAL", "JUMLAH", "GRAND TOTAL")):
                    continue

                inv_date = (
                    _parse_date(r[col_date] if col_date is not None and col_date < len(r) else None)
                    or report_date
                )
                if inv_date is None:
                    result.warnings.append(
                        f"Sheet '{sheet_name}' faktur '{inv_no}': tanggal tidak valid, dilewati."
                    )
                    continue

                tx_amount = _to_decimal(
                    r[col_amount] if col_amount is not None and col_amount < len(r) else None
                )
                overpay = _to_decimal(
                    r[col_overpay] if col_overpay is not None and col_overpay < len(r) else None
                )
                voucher = _to_decimal(
                    r[col_voucher] if col_voucher is not None and col_voucher < len(r) else None
                )
                row_no = _to_int(r[col_no] if col_no is not None and col_no < len(r) else None)

                payment_lines: list[PaymentLine] = []
                for idx, label in col_map.items():
                    if label in PAYMENT_COLUMN_MAP and idx < len(r):
                        amt = _to_decimal(r[idx])
                        if amt != Decimal("0.00"):
                            group, channel = PAYMENT_COLUMN_MAP[label]
                            payment_lines.append(
                                PaymentLine(
                                    payment_group=group,
                                    bank_channel=channel,
                                    amount=amt,
                                )
                            )

                result.invoices.append(
                    GrosirInvoice(
                        branch=branch,
                        report_date=report_date,
                        invoice_date=inv_date,
                        invoice_no=inv_no,
                        customer_name=cust_name,
                        transaction_amount=tx_amount,
                        overpay=overpay,
                        voucher=voucher,
                        source_file=filename,
                        sheet_name=sheet_name.strip(),
                        row_no=row_no,
                        payment_lines=payment_lines,
                    )
                )
    finally:
        wb.close()

    return result
