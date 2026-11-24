-- Target: 10.50.100.28 | Database: db_snj_group | Schema: schm_snj

CREATE SCHEMA IF NOT EXISTS schm_snj;

-- 1. Tabel Grosir (1 tabel flat: setiap kanal pembayaran per faktur menjadi 1 baris)
CREATE TABLE IF NOT EXISTS schm_snj.grosir_invoices (
    id             BIGSERIAL PRIMARY KEY,
    branch         VARCHAR(100) NOT NULL,
    report_date    DATE,
    invoice_date   DATE NOT NULL,
    invoice_no     VARCHAR(100) NOT NULL,
    customer_name  VARCHAR(255) NOT NULL,
    payment_group  VARCHAR(50) NOT NULL,   -- DEPOSIT, KAS, TRANSFER, EDC_DEBIT, EDC_CC, QRIS, UNPAID
    bank_channel   VARCHAR(50) NOT NULL,   -- MANDIRI, BCA, BRI, DKI, CASH, DEPOSIT, NONE
    payment_amount NUMERIC(18, 2) NOT NULL DEFAULT 0,
    overpay        NUMERIC(18, 2) NOT NULL DEFAULT 0,
    voucher        NUMERIC(18, 2) NOT NULL DEFAULT 0,
    source_file    VARCHAR(255) NOT NULL,
    sheet_name     VARCHAR(100) NOT NULL,
    row_no         INTEGER,
    created_at     TIMESTAMP WITHOUT TIME ZONE NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_grosir_invoice_payment UNIQUE (
        branch, invoice_date, invoice_no, payment_group, bank_channel
    )
);

CREATE INDEX IF NOT EXISTS idx_grosir_invoices_date_branch
    ON schm_snj.grosir_invoices (invoice_date, branch);

CREATE INDEX IF NOT EXISTS idx_grosir_invoices_no
    ON schm_snj.grosir_invoices (invoice_no);

CREATE INDEX IF NOT EXISTS idx_grosir_invoices_payment_channel
    ON schm_snj.grosir_invoices (payment_group, bank_channel);

-- 2. Tabel Mutasi Bank (referensi: gold_bank_transactions)
CREATE TABLE IF NOT EXISTS schm_snj.bank_transactions (
    id           BIGSERIAL PRIMARY KEY,
    bank_name    VARCHAR(50) NOT NULL,     -- BCA, MANDIRI, BRI, DKI, BNI
    account_no   VARCHAR(50) NOT NULL,
    source_file  VARCHAR(255) NOT NULL,
    source_sheet VARCHAR(100) NOT NULL,
    period       DATE,
    tgl          DATE,
    section      TEXT,
    keterangan   TEXT NOT NULL,
    debet        NUMERIC(18, 2) NOT NULL DEFAULT 0,
    kredit       NUMERIC(18, 2) NOT NULL DEFAULT 0,
    saldo        NUMERIC(18, 2),
    row_kind     VARCHAR(30) NOT NULL DEFAULT 'transaction', -- opening | transaction
    row_num      INTEGER NOT NULL,
    created_at   TIMESTAMP WITHOUT TIME ZONE NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_bank_tx_row UNIQUE (account_no, source_sheet, period, row_num)
);

CREATE INDEX IF NOT EXISTS idx_bank_tx_period_bank
    ON schm_snj.bank_transactions (period, bank_name);

CREATE INDEX IF NOT EXISTS idx_bank_tx_account
    ON schm_snj.bank_transactions (account_no, tgl);
