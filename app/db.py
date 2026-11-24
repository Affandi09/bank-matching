from __future__ import annotations

from decimal import Decimal
from typing import Any

import psycopg
from psycopg import sql

from app.bank_parser import BankTransactionRow
from app.config import settings
from app.parser import GrosirInvoice


def check_db_status() -> dict[str, Any]:
    """Check PostgreSQL connectivity and target table availability."""
    try:
        with psycopg.connect(settings.dsn, connect_timeout=5) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT
                        to_regclass(%s) IS NOT NULL,
                        to_regclass(%s) IS NOT NULL
                    """,
                    (
                        f"{settings.db_schema}.grosir_invoices",
                        f"{settings.db_schema}.bank_transactions",
                    ),
                )
                row = cur.fetchone()
                return {
                    "connected": True,
                    "host": settings.db_host,
                    "database": settings.db_name,
                    "schema": settings.db_schema,
                    "has_invoices_table": bool(row and row[0]),
                    "has_bank_table": bool(row and row[1]),
                }
    except Exception as exc:
        return {
            "connected": False,
            "host": settings.db_host,
            "database": settings.db_name,
            "schema": settings.db_schema,
            "error": str(exc),
        }


def flatten_invoice_rows(invoices: list[GrosirInvoice]) -> list[tuple[Any, ...]]:
    """Flatten invoices so each payment channel becomes a row in grosir_invoices."""
    rows: list[tuple[Any, ...]] = []
    for inv in invoices:
        lines = inv.payment_lines or []
        if not lines:
            rows.append(
                (
                    inv.branch,
                    inv.report_date,
                    inv.invoice_date,
                    inv.invoice_no,
                    inv.customer_name,
                    "UNPAID",
                    "NONE",
                    Decimal("0.00"),
                    inv.overpay,
                    inv.voucher,
                    inv.source_file,
                    inv.sheet_name,
                    inv.row_no,
                )
            )
            continue

        for pl in lines:
            rows.append(
                (
                    inv.branch,
                    inv.report_date,
                    inv.invoice_date,
                    inv.invoice_no,
                    inv.customer_name,
                    pl.payment_group,
                    pl.bank_channel,
                    pl.amount,
                    inv.overpay,
                    inv.voucher,
                    inv.source_file,
                    inv.sheet_name,
                    inv.row_no,
                )
            )
    return rows


def save_invoices(invoices: list[GrosirInvoice]) -> dict[str, int]:
    """Upsert flattened payment rows into schm_snj.grosir_invoices in one transaction."""
    rows = flatten_invoice_rows(invoices)
    if not rows:
        return {"rows_saved": 0}

    schema_id = sql.Identifier(settings.db_schema)
    upsert_sql = sql.SQL(
        """
        INSERT INTO {schema}.grosir_invoices (
            branch, report_date, invoice_date, invoice_no, customer_name,
            payment_group, bank_channel, payment_amount, overpay, voucher,
            source_file, sheet_name, row_no
        ) VALUES (
            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
        )
        ON CONFLICT (branch, invoice_date, invoice_no, payment_group, bank_channel)
        DO UPDATE SET
            report_date    = EXCLUDED.report_date,
            customer_name  = EXCLUDED.customer_name,
            payment_amount = EXCLUDED.payment_amount,
            overpay        = EXCLUDED.overpay,
            voucher        = EXCLUDED.voucher,
            source_file    = EXCLUDED.source_file,
            sheet_name     = EXCLUDED.sheet_name,
            row_no         = EXCLUDED.row_no
        """
    ).format(schema=schema_id)

    with psycopg.connect(settings.dsn, connect_timeout=10) as conn:
        with conn.transaction():
            with conn.cursor() as cur:
                cur.executemany(upsert_sql, rows)

    return {"rows_saved": len(rows)}


def save_bank_transactions(transactions: list[BankTransactionRow]) -> dict[str, int]:
    """Upsert bank statement rows into schm_snj.bank_transactions in one transaction."""
    if not transactions:
        return {"rows_saved": 0}

    schema_id = sql.Identifier(settings.db_schema)
    upsert_sql = sql.SQL(
        """
        INSERT INTO {schema}.bank_transactions (
            bank_name, account_no, source_file, source_sheet, period, tgl,
            section, keterangan, debet, kredit, saldo, row_kind, row_num
        ) VALUES (
            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
        )
        ON CONFLICT (account_no, source_sheet, period, row_num)
        DO UPDATE SET
            bank_name   = EXCLUDED.bank_name,
            source_file = EXCLUDED.source_file,
            tgl         = EXCLUDED.tgl,
            section     = EXCLUDED.section,
            keterangan  = EXCLUDED.keterangan,
            debet       = EXCLUDED.debet,
            kredit      = EXCLUDED.kredit,
            saldo       = EXCLUDED.saldo,
            row_kind    = EXCLUDED.row_kind
        """
    ).format(schema=schema_id)

    rows = [
        (
            t.bank_name,
            t.account_no,
            t.source_file,
            t.source_sheet,
            t.period,
            t.tgl,
            t.section,
            t.keterangan,
            t.debet,
            t.kredit,
            t.saldo,
            t.row_kind,
            t.row_num,
        )
        for t in transactions
    ]

    with psycopg.connect(settings.dsn, connect_timeout=10) as conn:
        with conn.transaction():
            with conn.cursor() as cur:
                cur.executemany(upsert_sql, rows)

    return {"rows_saved": len(rows)}
