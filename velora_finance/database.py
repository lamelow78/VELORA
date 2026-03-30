from __future__ import annotations

import calendar as month_calendar
import json
import sqlite3
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from .config import asset_path, storage_root


DEFAULT_COMPANY_PROFILE = {
    "company_name": "Votre entreprise",
    "legal_name": "Votre entreprise SAS",
    "siret": "",
    "vat_number": "",
    "email": "",
    "phone": "",
    "address": "",
    "footer": "Merci pour votre confiance.",
    "logo_path": "",
}

DEFAULT_DOCUMENT_PREFERENCES = {
    "default_tax_rate": 20.0,
    "invoice_due_days": 30,
    "quote_validity_days": 30,
    "default_client_name": "",
    "default_client_email": "",
    "default_client_address": "",
    "default_invoice_notes": "",
    "default_quote_notes": "",
    "default_invoice_status": "Brouillon",
    "default_quote_status": "Brouillon",
}


class Database:
    def __init__(self, db_path: Path | None = None) -> None:
        self.root = storage_root()
        self.db_path = Path(db_path) if db_path else self.root / "velora_finance.db"
        self.connection = sqlite3.connect(self.db_path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.execute("PRAGMA journal_mode = WAL")
        self._create_tables()
        self._migrate_tables()
        self._ensure_defaults()
        self._ensure_schema_version()

    def _create_tables(self) -> None:
        self.connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_date TEXT NOT NULL CHECK(length(sale_date) = 10),
                client TEXT NOT NULL CHECK(TRIM(client) <> ''),
                company_name TEXT NOT NULL DEFAULT '' CHECK(TRIM(company_name) <> ''),
                category TEXT NOT NULL CHECK(TRIM(category) <> ''),
                amount REAL NOT NULL DEFAULT 0 CHECK(amount >= 0),
                amount_ht REAL NOT NULL DEFAULT 0 CHECK(amount_ht >= 0),
                amount_ttc REAL NOT NULL DEFAULT 0 CHECK(amount_ttc >= amount_ht),
                source_type TEXT NOT NULL DEFAULT 'manual' CHECK(source_type IN ('manual', 'invoice')),
                source_invoice_id INTEGER,
                notes TEXT DEFAULT '',
                created_at TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT '',
                CHECK(
                    (source_type = 'manual' AND source_invoice_id IS NULL)
                    OR (source_type = 'invoice' AND source_invoice_id IS NOT NULL)
                ),
                FOREIGN KEY(source_invoice_id) REFERENCES invoices(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                expense_date TEXT NOT NULL CHECK(length(expense_date) = 10),
                vendor TEXT NOT NULL CHECK(TRIM(vendor) <> ''),
                company_name TEXT NOT NULL DEFAULT '' CHECK(TRIM(company_name) <> ''),
                category TEXT NOT NULL CHECK(TRIM(category) <> ''),
                amount REAL NOT NULL DEFAULT 0 CHECK(amount >= 0),
                amount_ht REAL NOT NULL DEFAULT 0 CHECK(amount_ht >= 0),
                amount_ttc REAL NOT NULL DEFAULT 0 CHECK(amount_ttc >= amount_ht),
                source_type TEXT NOT NULL DEFAULT 'manual' CHECK(source_type IN ('manual', 'invoice')),
                source_invoice_id INTEGER,
                notes TEXT DEFAULT '',
                created_at TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT '',
                CHECK(
                    (source_type = 'manual' AND source_invoice_id IS NULL)
                    OR (source_type = 'invoice' AND source_invoice_id IS NOT NULL)
                ),
                FOREIGN KEY(source_invoice_id) REFERENCES invoices(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT NOT NULL UNIQUE,
                issue_date TEXT NOT NULL CHECK(length(issue_date) = 10),
                due_date TEXT NOT NULL CHECK(length(due_date) = 10 AND due_date >= issue_date),
                status TEXT NOT NULL CHECK(status IN ('Brouillon', 'Envoyee', 'Payee', 'En retard')),
                client_name TEXT NOT NULL CHECK(TRIM(client_name) <> ''),
                client_email TEXT DEFAULT '',
                client_address TEXT DEFAULT '',
                items_json TEXT NOT NULL,
                notes TEXT DEFAULT '',
                tax_rate REAL NOT NULL CHECK(tax_rate >= 0 AND tax_rate <= 100),
                subtotal REAL NOT NULL CHECK(subtotal >= 0),
                tax_amount REAL NOT NULL CHECK(tax_amount >= 0),
                total REAL NOT NULL CHECK(total >= 0 AND total >= subtotal),
                company_json TEXT NOT NULL,
                html_path TEXT NOT NULL CHECK(TRIM(html_path) <> ''),
                created_at TEXT NOT NULL CHECK(TRIM(created_at) <> ''),
                updated_at TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS quotes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT NOT NULL UNIQUE,
                issue_date TEXT NOT NULL CHECK(length(issue_date) = 10),
                valid_until TEXT NOT NULL CHECK(length(valid_until) = 10 AND valid_until >= issue_date),
                status TEXT NOT NULL CHECK(status IN ('Brouillon', 'Envoye', 'Accepte', 'Refuse', 'Expire')),
                client_name TEXT NOT NULL CHECK(TRIM(client_name) <> ''),
                client_email TEXT DEFAULT '',
                client_address TEXT DEFAULT '',
                items_json TEXT NOT NULL,
                notes TEXT DEFAULT '',
                tax_rate REAL NOT NULL CHECK(tax_rate >= 0 AND tax_rate <= 100),
                subtotal REAL NOT NULL CHECK(subtotal >= 0),
                tax_amount REAL NOT NULL CHECK(tax_amount >= 0),
                total REAL NOT NULL CHECK(total >= 0 AND total >= subtotal),
                company_json TEXT NOT NULL,
                html_path TEXT NOT NULL CHECK(TRIM(html_path) <> ''),
                created_at TEXT NOT NULL CHECK(TRIM(created_at) <> ''),
                updated_at TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS todos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL CHECK(TRIM(title) <> ''),
                details TEXT DEFAULT '',
                task_date TEXT NOT NULL CHECK(length(task_date) = 10),
                task_time TEXT DEFAULT '' CHECK(
                    task_time = ''
                    OR (
                        length(task_time) = 5
                        AND substr(task_time, 3, 1) = ':'
                        AND CAST(substr(task_time, 1, 2) AS INTEGER) BETWEEN 0 AND 23
                        AND CAST(substr(task_time, 4, 2) AS INTEGER) BETWEEN 0 AND 59
                    )
                ),
                status TEXT NOT NULL DEFAULT 'A faire' CHECK(status IN ('A faire', 'En cours', 'Terminee')),
                created_at TEXT NOT NULL CHECK(TRIM(created_at) <> ''),
                updated_at TEXT NOT NULL CHECK(TRIM(updated_at) <> '')
            );

            CREATE INDEX IF NOT EXISTS idx_sales_date ON sales(sale_date);
            CREATE INDEX IF NOT EXISTS idx_expenses_date ON expenses(expense_date);
            CREATE INDEX IF NOT EXISTS idx_invoices_issue_date ON invoices(issue_date);
            CREATE INDEX IF NOT EXISTS idx_quotes_issue_date ON quotes(issue_date);
            CREATE INDEX IF NOT EXISTS idx_todos_task_date ON todos(task_date);
            """
        )
        self.connection.commit()

    def _migrate_tables(self) -> None:
        self._ensure_column("sales", "company_name", "TEXT DEFAULT ''")
        self._ensure_column("sales", "amount_ht", "REAL NOT NULL DEFAULT 0")
        self._ensure_column("sales", "amount_ttc", "REAL NOT NULL DEFAULT 0")
        self._ensure_column("sales", "source_type", "TEXT NOT NULL DEFAULT 'manual'")
        self._ensure_column("sales", "source_invoice_id", "INTEGER")
        self._ensure_column("sales", "created_at", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column("sales", "updated_at", "TEXT NOT NULL DEFAULT ''")

        self._ensure_column("expenses", "company_name", "TEXT DEFAULT ''")
        self._ensure_column("expenses", "amount_ht", "REAL NOT NULL DEFAULT 0")
        self._ensure_column("expenses", "amount_ttc", "REAL NOT NULL DEFAULT 0")
        self._ensure_column("expenses", "source_type", "TEXT NOT NULL DEFAULT 'manual'")
        self._ensure_column("expenses", "source_invoice_id", "INTEGER")
        self._ensure_column("expenses", "created_at", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column("expenses", "updated_at", "TEXT NOT NULL DEFAULT ''")

        self._ensure_column("invoices", "updated_at", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column("quotes", "updated_at", "TEXT NOT NULL DEFAULT ''")

        self.connection.executescript(
            """
            UPDATE sales
            SET company_name = CASE
                WHEN TRIM(COALESCE(company_name, '')) = '' THEN client
                ELSE company_name
            END;

            UPDATE sales
            SET amount_ttc = CASE
                WHEN COALESCE(amount_ttc, 0) = 0 THEN amount
                ELSE amount_ttc
            END;

            UPDATE sales
            SET amount_ht = CASE
                WHEN COALESCE(amount_ht, 0) = 0 THEN amount_ttc
                ELSE amount_ht
            END;

            UPDATE sales
            SET created_at = CASE
                WHEN TRIM(COALESCE(created_at, '')) = '' THEN sale_date || 'T00:00:00'
                ELSE created_at
            END;

            UPDATE sales
            SET updated_at = CASE
                WHEN TRIM(COALESCE(updated_at, '')) = '' THEN created_at
                ELSE updated_at
            END;

            UPDATE expenses
            SET company_name = CASE
                WHEN TRIM(COALESCE(company_name, '')) = '' THEN vendor
                ELSE company_name
            END;

            UPDATE expenses
            SET amount_ttc = CASE
                WHEN COALESCE(amount_ttc, 0) = 0 THEN amount
                ELSE amount_ttc
            END;

            UPDATE expenses
            SET amount_ht = CASE
                WHEN COALESCE(amount_ht, 0) = 0 THEN amount_ttc
                ELSE amount_ht
            END;

            UPDATE expenses
            SET created_at = CASE
                WHEN TRIM(COALESCE(created_at, '')) = '' THEN expense_date || 'T00:00:00'
                ELSE created_at
            END;

            UPDATE expenses
            SET updated_at = CASE
                WHEN TRIM(COALESCE(updated_at, '')) = '' THEN created_at
                ELSE updated_at
            END;

            UPDATE invoices
            SET updated_at = CASE
                WHEN TRIM(COALESCE(updated_at, '')) = '' THEN created_at
                ELSE updated_at
            END;

            UPDATE quotes
            SET updated_at = CASE
                WHEN TRIM(COALESCE(updated_at, '')) = '' THEN created_at
                ELSE updated_at
            END;
            """
        )
        self.connection.commit()

    def _ensure_column(self, table: str, column: str, definition: str) -> None:
        existing = {
            row["name"]
            for row in self.connection.execute(f"PRAGMA table_info({table})").fetchall()
        }
        if column not in existing:
            self.connection.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
            self.connection.commit()

    def _ensure_defaults(self) -> None:
        if self.get_setting("company_profile") is None:
            self.set_setting("company_profile", DEFAULT_COMPANY_PROFILE)
        if self.get_setting("document_preferences") is None:
            self.set_setting("document_preferences", DEFAULT_DOCUMENT_PREFERENCES)

    def _ensure_schema_version(self) -> None:
        current_version = int(
            self.connection.execute("PRAGMA user_version").fetchone()[0]
        )
        target_version = 2
        if current_version >= target_version:
            return
        self._rebuild_tables_with_constraints()
        self.connection.execute(f"PRAGMA user_version = {target_version}")
        self.connection.commit()

    def _rebuild_tables_with_constraints(self) -> None:
        self.connection.execute("PRAGMA foreign_keys = OFF")
        try:
            self._rebuild_invoices_table()
            self._rebuild_quotes_table()
            self._rebuild_todos_table()
            self._rebuild_sales_table()
            self._rebuild_expenses_table()
        finally:
            self.connection.execute("PRAGMA foreign_keys = ON")
        violations = self.connection.execute("PRAGMA foreign_key_check").fetchall()
        if violations:
            raise ValueError("La migration de schema a laisse des references invalides.")

    def _rebuild_invoices_table(self) -> None:
        self.connection.executescript(
            """
            CREATE TABLE invoices__new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT NOT NULL UNIQUE,
                issue_date TEXT NOT NULL CHECK(length(issue_date) = 10),
                due_date TEXT NOT NULL CHECK(length(due_date) = 10 AND due_date >= issue_date),
                status TEXT NOT NULL CHECK(status IN ('Brouillon', 'Envoyee', 'Payee', 'En retard')),
                client_name TEXT NOT NULL CHECK(TRIM(client_name) <> ''),
                client_email TEXT DEFAULT '',
                client_address TEXT DEFAULT '',
                items_json TEXT NOT NULL,
                notes TEXT DEFAULT '',
                tax_rate REAL NOT NULL CHECK(tax_rate >= 0 AND tax_rate <= 100),
                subtotal REAL NOT NULL CHECK(subtotal >= 0),
                tax_amount REAL NOT NULL CHECK(tax_amount >= 0),
                total REAL NOT NULL CHECK(total >= 0 AND total >= subtotal),
                company_json TEXT NOT NULL,
                html_path TEXT NOT NULL CHECK(TRIM(html_path) <> ''),
                created_at TEXT NOT NULL CHECK(TRIM(created_at) <> ''),
                updated_at TEXT NOT NULL DEFAULT ''
            );

            INSERT INTO invoices__new (
                id, number, issue_date, due_date, status, client_name, client_email,
                client_address, items_json, notes, tax_rate, subtotal, tax_amount,
                total, company_json, html_path, created_at, updated_at
            )
            SELECT
                id,
                number,
                issue_date,
                CASE WHEN due_date < issue_date THEN issue_date ELSE due_date END,
                CASE
                    WHEN status IN ('Brouillon', 'Envoyee', 'Payee', 'En retard') THEN status
                    ELSE 'Brouillon'
                END,
                CASE
                    WHEN TRIM(COALESCE(client_name, '')) = '' THEN 'Client'
                    ELSE client_name
                END,
                COALESCE(client_email, ''),
                COALESCE(client_address, ''),
                items_json,
                COALESCE(notes, ''),
                CASE
                    WHEN tax_rate < 0 THEN 0
                    WHEN tax_rate > 100 THEN 100
                    ELSE tax_rate
                END,
                CASE WHEN subtotal < 0 THEN 0 ELSE subtotal END,
                CASE WHEN tax_amount < 0 THEN 0 ELSE tax_amount END,
                CASE WHEN total < 0 THEN 0 ELSE total END,
                company_json,
                html_path,
                CASE
                    WHEN TRIM(COALESCE(created_at, '')) = '' THEN issue_date || 'T00:00:00'
                    ELSE created_at
                END,
                CASE
                    WHEN TRIM(COALESCE(updated_at, '')) = '' THEN
                        CASE
                            WHEN TRIM(COALESCE(created_at, '')) = '' THEN issue_date || 'T00:00:00'
                            ELSE created_at
                        END
                    ELSE updated_at
                END
            FROM invoices;

            DROP TABLE invoices;
            ALTER TABLE invoices__new RENAME TO invoices;
            CREATE INDEX IF NOT EXISTS idx_invoices_issue_date ON invoices(issue_date);
            """
        )
        self.connection.commit()

    def _rebuild_quotes_table(self) -> None:
        self.connection.executescript(
            """
            CREATE TABLE quotes__new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT NOT NULL UNIQUE,
                issue_date TEXT NOT NULL CHECK(length(issue_date) = 10),
                valid_until TEXT NOT NULL CHECK(length(valid_until) = 10 AND valid_until >= issue_date),
                status TEXT NOT NULL CHECK(status IN ('Brouillon', 'Envoye', 'Accepte', 'Refuse', 'Expire')),
                client_name TEXT NOT NULL CHECK(TRIM(client_name) <> ''),
                client_email TEXT DEFAULT '',
                client_address TEXT DEFAULT '',
                items_json TEXT NOT NULL,
                notes TEXT DEFAULT '',
                tax_rate REAL NOT NULL CHECK(tax_rate >= 0 AND tax_rate <= 100),
                subtotal REAL NOT NULL CHECK(subtotal >= 0),
                tax_amount REAL NOT NULL CHECK(tax_amount >= 0),
                total REAL NOT NULL CHECK(total >= 0 AND total >= subtotal),
                company_json TEXT NOT NULL,
                html_path TEXT NOT NULL CHECK(TRIM(html_path) <> ''),
                created_at TEXT NOT NULL CHECK(TRIM(created_at) <> ''),
                updated_at TEXT NOT NULL DEFAULT ''
            );

            INSERT INTO quotes__new (
                id, number, issue_date, valid_until, status, client_name, client_email,
                client_address, items_json, notes, tax_rate, subtotal, tax_amount,
                total, company_json, html_path, created_at, updated_at
            )
            SELECT
                id,
                number,
                issue_date,
                CASE WHEN valid_until < issue_date THEN issue_date ELSE valid_until END,
                CASE
                    WHEN status IN ('Brouillon', 'Envoye', 'Accepte', 'Refuse', 'Expire') THEN status
                    ELSE 'Brouillon'
                END,
                CASE
                    WHEN TRIM(COALESCE(client_name, '')) = '' THEN 'Client'
                    ELSE client_name
                END,
                COALESCE(client_email, ''),
                COALESCE(client_address, ''),
                items_json,
                COALESCE(notes, ''),
                CASE
                    WHEN tax_rate < 0 THEN 0
                    WHEN tax_rate > 100 THEN 100
                    ELSE tax_rate
                END,
                CASE WHEN subtotal < 0 THEN 0 ELSE subtotal END,
                CASE WHEN tax_amount < 0 THEN 0 ELSE tax_amount END,
                CASE WHEN total < 0 THEN 0 ELSE total END,
                company_json,
                html_path,
                CASE
                    WHEN TRIM(COALESCE(created_at, '')) = '' THEN issue_date || 'T00:00:00'
                    ELSE created_at
                END,
                CASE
                    WHEN TRIM(COALESCE(updated_at, '')) = '' THEN
                        CASE
                            WHEN TRIM(COALESCE(created_at, '')) = '' THEN issue_date || 'T00:00:00'
                            ELSE created_at
                        END
                    ELSE updated_at
                END
            FROM quotes;

            DROP TABLE quotes;
            ALTER TABLE quotes__new RENAME TO quotes;
            CREATE INDEX IF NOT EXISTS idx_quotes_issue_date ON quotes(issue_date);
            """
        )
        self.connection.commit()

    def _rebuild_todos_table(self) -> None:
        self.connection.executescript(
            """
            CREATE TABLE todos__new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL CHECK(TRIM(title) <> ''),
                details TEXT DEFAULT '',
                task_date TEXT NOT NULL CHECK(length(task_date) = 10),
                task_time TEXT DEFAULT '' CHECK(
                    task_time = ''
                    OR (
                        length(task_time) = 5
                        AND substr(task_time, 3, 1) = ':'
                        AND CAST(substr(task_time, 1, 2) AS INTEGER) BETWEEN 0 AND 23
                        AND CAST(substr(task_time, 4, 2) AS INTEGER) BETWEEN 0 AND 59
                    )
                ),
                status TEXT NOT NULL DEFAULT 'A faire' CHECK(status IN ('A faire', 'En cours', 'Terminee')),
                created_at TEXT NOT NULL CHECK(TRIM(created_at) <> ''),
                updated_at TEXT NOT NULL CHECK(TRIM(updated_at) <> '')
            );

            INSERT INTO todos__new (
                id, title, details, task_date, task_time, status, created_at, updated_at
            )
            SELECT
                id,
                CASE
                    WHEN TRIM(COALESCE(title, '')) = '' THEN 'Tache'
                    ELSE title
                END,
                COALESCE(details, ''),
                task_date,
                CASE
                    WHEN TRIM(COALESCE(task_time, '')) = '' THEN ''
                    WHEN length(task_time) = 5
                        AND substr(task_time, 3, 1) = ':'
                        AND CAST(substr(task_time, 1, 2) AS INTEGER) BETWEEN 0 AND 23
                        AND CAST(substr(task_time, 4, 2) AS INTEGER) BETWEEN 0 AND 59
                    THEN task_time
                    ELSE ''
                END,
                CASE
                    WHEN status IN ('A faire', 'En cours', 'Terminee') THEN status
                    ELSE 'A faire'
                END,
                CASE
                    WHEN TRIM(COALESCE(created_at, '')) = '' THEN task_date || 'T00:00:00'
                    ELSE created_at
                END,
                CASE
                    WHEN TRIM(COALESCE(updated_at, '')) = '' THEN
                        CASE
                            WHEN TRIM(COALESCE(created_at, '')) = '' THEN task_date || 'T00:00:00'
                            ELSE created_at
                        END
                    ELSE updated_at
                END
            FROM todos;

            DROP TABLE todos;
            ALTER TABLE todos__new RENAME TO todos;
            CREATE INDEX IF NOT EXISTS idx_todos_task_date ON todos(task_date);
            """
        )
        self.connection.commit()

    def _rebuild_sales_table(self) -> None:
        self.connection.executescript(
            """
            CREATE TABLE sales__new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_date TEXT NOT NULL CHECK(length(sale_date) = 10),
                client TEXT NOT NULL CHECK(TRIM(client) <> ''),
                company_name TEXT NOT NULL DEFAULT '' CHECK(TRIM(company_name) <> ''),
                category TEXT NOT NULL CHECK(TRIM(category) <> ''),
                amount REAL NOT NULL DEFAULT 0 CHECK(amount >= 0),
                amount_ht REAL NOT NULL DEFAULT 0 CHECK(amount_ht >= 0),
                amount_ttc REAL NOT NULL DEFAULT 0 CHECK(amount_ttc >= amount_ht),
                source_type TEXT NOT NULL DEFAULT 'manual' CHECK(source_type IN ('manual', 'invoice')),
                source_invoice_id INTEGER,
                notes TEXT DEFAULT '',
                created_at TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT '',
                CHECK(
                    (source_type = 'manual' AND source_invoice_id IS NULL)
                    OR (source_type = 'invoice' AND source_invoice_id IS NOT NULL)
                ),
                FOREIGN KEY(source_invoice_id) REFERENCES invoices(id) ON DELETE SET NULL
            );

            INSERT INTO sales__new (
                id, sale_date, client, company_name, category, amount, amount_ht, amount_ttc,
                source_type, source_invoice_id, notes, created_at, updated_at
            )
            SELECT
                id,
                sale_date,
                CASE
                    WHEN TRIM(COALESCE(company_name, '')) = '' THEN client
                    ELSE company_name
                END,
                CASE
                    WHEN TRIM(COALESCE(company_name, '')) = '' THEN client
                    ELSE company_name
                END,
                CASE
                    WHEN TRIM(COALESCE(category, '')) = '' THEN 'Autre'
                    ELSE category
                END,
                CASE
                    WHEN COALESCE(NULLIF(amount_ttc, 0), amount) < 0 THEN 0
                    ELSE COALESCE(NULLIF(amount_ttc, 0), amount)
                END,
                CASE
                    WHEN amount_ht < 0 THEN 0
                    ELSE amount_ht
                END,
                CASE
                    WHEN COALESCE(NULLIF(amount_ttc, 0), amount) < amount_ht THEN amount_ht
                    ELSE COALESCE(NULLIF(amount_ttc, 0), amount)
                END,
                CASE
                    WHEN source_type = 'invoice'
                        AND source_invoice_id IS NOT NULL
                        AND EXISTS(SELECT 1 FROM invoices WHERE id = source_invoice_id)
                    THEN 'invoice'
                    ELSE 'manual'
                END,
                CASE
                    WHEN source_type = 'invoice'
                        AND source_invoice_id IS NOT NULL
                        AND EXISTS(SELECT 1 FROM invoices WHERE id = source_invoice_id)
                    THEN source_invoice_id
                    ELSE NULL
                END,
                COALESCE(notes, ''),
                CASE
                    WHEN TRIM(COALESCE(created_at, '')) = '' THEN sale_date || 'T00:00:00'
                    ELSE created_at
                END,
                CASE
                    WHEN TRIM(COALESCE(updated_at, '')) = '' THEN
                        CASE
                            WHEN TRIM(COALESCE(created_at, '')) = '' THEN sale_date || 'T00:00:00'
                            ELSE created_at
                        END
                    ELSE updated_at
                END
            FROM sales;

            DROP TABLE sales;
            ALTER TABLE sales__new RENAME TO sales;
            CREATE INDEX IF NOT EXISTS idx_sales_date ON sales(sale_date);
            """
        )
        self.connection.commit()

    def _rebuild_expenses_table(self) -> None:
        self.connection.executescript(
            """
            CREATE TABLE expenses__new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                expense_date TEXT NOT NULL CHECK(length(expense_date) = 10),
                vendor TEXT NOT NULL CHECK(TRIM(vendor) <> ''),
                company_name TEXT NOT NULL DEFAULT '' CHECK(TRIM(company_name) <> ''),
                category TEXT NOT NULL CHECK(TRIM(category) <> ''),
                amount REAL NOT NULL DEFAULT 0 CHECK(amount >= 0),
                amount_ht REAL NOT NULL DEFAULT 0 CHECK(amount_ht >= 0),
                amount_ttc REAL NOT NULL DEFAULT 0 CHECK(amount_ttc >= amount_ht),
                source_type TEXT NOT NULL DEFAULT 'manual' CHECK(source_type IN ('manual', 'invoice')),
                source_invoice_id INTEGER,
                notes TEXT DEFAULT '',
                created_at TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT '',
                CHECK(
                    (source_type = 'manual' AND source_invoice_id IS NULL)
                    OR (source_type = 'invoice' AND source_invoice_id IS NOT NULL)
                ),
                FOREIGN KEY(source_invoice_id) REFERENCES invoices(id) ON DELETE SET NULL
            );

            INSERT INTO expenses__new (
                id, expense_date, vendor, company_name, category, amount, amount_ht, amount_ttc,
                source_type, source_invoice_id, notes, created_at, updated_at
            )
            SELECT
                id,
                expense_date,
                CASE
                    WHEN TRIM(COALESCE(company_name, '')) = '' THEN vendor
                    ELSE company_name
                END,
                CASE
                    WHEN TRIM(COALESCE(company_name, '')) = '' THEN vendor
                    ELSE company_name
                END,
                CASE
                    WHEN TRIM(COALESCE(category, '')) = '' THEN 'Autre'
                    ELSE category
                END,
                CASE
                    WHEN COALESCE(NULLIF(amount_ttc, 0), amount) < 0 THEN 0
                    ELSE COALESCE(NULLIF(amount_ttc, 0), amount)
                END,
                CASE
                    WHEN amount_ht < 0 THEN 0
                    ELSE amount_ht
                END,
                CASE
                    WHEN COALESCE(NULLIF(amount_ttc, 0), amount) < amount_ht THEN amount_ht
                    ELSE COALESCE(NULLIF(amount_ttc, 0), amount)
                END,
                CASE
                    WHEN source_type = 'invoice'
                        AND source_invoice_id IS NOT NULL
                        AND EXISTS(SELECT 1 FROM invoices WHERE id = source_invoice_id)
                    THEN 'invoice'
                    ELSE 'manual'
                END,
                CASE
                    WHEN source_type = 'invoice'
                        AND source_invoice_id IS NOT NULL
                        AND EXISTS(SELECT 1 FROM invoices WHERE id = source_invoice_id)
                    THEN source_invoice_id
                    ELSE NULL
                END,
                COALESCE(notes, ''),
                CASE
                    WHEN TRIM(COALESCE(created_at, '')) = '' THEN expense_date || 'T00:00:00'
                    ELSE created_at
                END,
                CASE
                    WHEN TRIM(COALESCE(updated_at, '')) = '' THEN
                        CASE
                            WHEN TRIM(COALESCE(created_at, '')) = '' THEN expense_date || 'T00:00:00'
                            ELSE created_at
                        END
                    ELSE updated_at
                END
            FROM expenses;

            DROP TABLE expenses;
            ALTER TABLE expenses__new RENAME TO expenses;
            CREATE INDEX IF NOT EXISTS idx_expenses_date ON expenses(expense_date);
            """
        )
        self.connection.commit()

    def close(self) -> None:
        self.connection.close()

    def _now(self) -> str:
        return datetime.now().isoformat(timespec="seconds")

    def get_setting(self, key: str):
        row = self.connection.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        if row is None:
            return None
        return json.loads(row["value"])

    def set_setting(self, key: str, value) -> None:
        payload = json.dumps(value, ensure_ascii=False)
        self.connection.execute(
            """
            INSERT INTO settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, payload),
        )
        self.connection.commit()

    def get_company_profile(self) -> dict:
        profile = DEFAULT_COMPANY_PROFILE.copy()
        profile.update(self.get_setting("company_profile") or {})
        if not profile.get("logo_path"):
            profile["logo_path"] = str(asset_path("velora_logo.png"))
        return profile

    def save_company_profile(self, profile: dict) -> None:
        clean_profile = DEFAULT_COMPANY_PROFILE.copy()
        clean_profile.update(profile)
        self.set_setting("company_profile", clean_profile)

    def get_document_preferences(self) -> dict:
        preferences = DEFAULT_DOCUMENT_PREFERENCES.copy()
        preferences.update(self.get_setting("document_preferences") or {})
        return preferences

    def save_document_preferences(self, preferences: dict) -> None:
        clean_preferences = DEFAULT_DOCUMENT_PREFERENCES.copy()
        clean_preferences.update(preferences)
        self.set_setting("document_preferences", clean_preferences)

    def _parse_iso_date(self, value: str, label: str) -> date:
        try:
            return date.fromisoformat(str(value).strip())
        except Exception as exc:
            raise ValueError(f"{label} invalide.") from exc

    def _parse_task_time(self, value: str) -> str:
        cleaned = str(value or "").strip()
        if not cleaned:
            return ""
        try:
            return datetime.strptime(cleaned, "%H:%M").strftime("%H:%M")
        except ValueError as exc:
            raise ValueError("L'heure doit etre au format HH:MM.") from exc

    def _validate_money_pair(self, amount_ht: float, amount_ttc: float) -> None:
        if amount_ht < 0 or amount_ttc < 0:
            raise ValueError("Les montants HT et TTC doivent etre positifs.")
        if amount_ttc < amount_ht:
            raise ValueError("Le montant TTC doit etre superieur ou egal au montant HT.")

    def _validate_tax_rate(self, tax_rate: float) -> None:
        if tax_rate < 0 or tax_rate > 100:
            raise ValueError("Le taux de TVA doit etre compris entre 0 et 100.")

    def _validate_document_status(self, kind: str, status: str) -> str:
        allowed = {
            "invoice": {"Brouillon", "Envoyee", "Payee", "En retard"},
            "quote": {"Brouillon", "Envoye", "Accepte", "Refuse", "Expire"},
        }
        cleaned = str(status).strip()
        if cleaned not in allowed[kind]:
            raise ValueError("Statut de document invalide.")
        return cleaned

    def _validate_financial_source(self, payload: dict) -> tuple[str, int | None]:
        source_type = str(payload.get("source_type", "manual")).strip() or "manual"
        if source_type not in {"manual", "invoice"}:
            raise ValueError("Source financiere invalide.")
        if source_type == "invoice":
            source_invoice_id = payload.get("source_invoice_id")
            if source_invoice_id is None:
                raise ValueError("La facture source est obligatoire.")
            row = self.connection.execute(
                "SELECT id FROM invoices WHERE id = ?",
                (int(source_invoice_id),),
            ).fetchone()
            if row is None:
                raise ValueError("La facture source n'existe pas.")
            return source_type, int(source_invoice_id)
        return "manual", None

    def _validate_financial_payload(self, payload: dict, date_field: str) -> dict:
        date_value = self._parse_iso_date(payload[date_field], "La date").isoformat()
        company_name = str(payload.get("company_name", "")).strip()
        if not company_name:
            raise ValueError("Le nom de l'entreprise est obligatoire.")
        category = str(payload.get("category", "")).strip()
        if not category:
            raise ValueError("La categorie est obligatoire.")
        amount_ht = float(payload["amount_ht"])
        amount_ttc = float(payload["amount_ttc"])
        self._validate_money_pair(amount_ht, amount_ttc)
        source_type, source_invoice_id = self._validate_financial_source(payload)
        return {
            "date_value": date_value,
            "company_name": company_name,
            "category": category,
            "amount_ht": amount_ht,
            "amount_ttc": amount_ttc,
            "source_type": source_type,
            "source_invoice_id": source_invoice_id,
            "notes": str(payload.get("notes", "")).strip(),
        }

    def _validate_document_payload(self, kind: str, payload: dict) -> dict:
        issue_date = self._parse_iso_date(payload["issue_date"], "La date du document")
        period_key = "due_date" if kind == "invoice" else "valid_until"
        period_label = "L'echeance" if kind == "invoice" else "La date de validite"
        period_date = self._parse_iso_date(payload[period_key], period_label)
        if period_date < issue_date:
            if kind == "invoice":
                raise ValueError("L'echeance ne peut pas etre anterieure a la date de facture.")
            raise ValueError("La validite du devis ne peut pas etre anterieure a la date du devis.")

        number = str(payload.get("number", "")).strip()
        if not number:
            raise ValueError("Le numero du document est obligatoire.")
        status = self._validate_document_status(kind, payload.get("status", ""))
        client_name = str(payload.get("client_name", "")).strip()
        if not client_name:
            raise ValueError("Le client est obligatoire.")

        items = payload.get("items") or []
        if not items:
            raise ValueError("Ajoutez au moins une ligne au document.")

        subtotal = 0.0
        clean_items = []
        for item in items:
            description = str(item.get("description", "")).strip()
            if not description:
                raise ValueError("Chaque ligne doit contenir une description.")
            quantity = float(item["quantity"])
            unit_price = float(item["unit_price"])
            if quantity <= 0:
                raise ValueError("La quantite doit etre superieure a 0.")
            if unit_price < 0:
                raise ValueError("Le prix unitaire doit etre positif.")
            clean_items.append(
                {
                    "description": description,
                    "quantity": quantity,
                    "unit_price": unit_price,
                }
            )
            subtotal += quantity * unit_price

        tax_rate = float(payload["tax_rate"])
        self._validate_tax_rate(tax_rate)
        declared_subtotal = float(payload["subtotal"])
        declared_tax_amount = float(payload["tax_amount"])
        declared_total = float(payload["total"])
        if declared_subtotal < 0 or declared_tax_amount < 0 or declared_total < 0:
            raise ValueError("Les montants du document doivent etre positifs.")

        expected_tax_amount = subtotal * (tax_rate / 100)
        expected_total = subtotal + expected_tax_amount
        if abs(declared_subtotal - subtotal) > 0.01:
            raise ValueError("Le sous-total du document est incoherent.")
        if abs(declared_tax_amount - expected_tax_amount) > 0.01:
            raise ValueError("Le montant de TVA du document est incoherent.")
        if abs(declared_total - expected_total) > 0.01:
            raise ValueError("Le total TTC du document est incoherent.")

        html_path = str(payload.get("html_path", "")).strip()
        if not html_path:
            raise ValueError("Le chemin local du document est obligatoire.")

        company_profile = payload.get("company_profile")
        if not isinstance(company_profile, dict):
            raise ValueError("Le profil entreprise du document est invalide.")

        validated = {
            "number": number,
            "issue_date": issue_date.isoformat(),
            period_key: period_date.isoformat(),
            "status": status,
            "client_name": client_name,
            "client_email": str(payload.get("client_email", "")).strip(),
            "client_address": str(payload.get("client_address", "")).strip(),
            "items": clean_items,
            "notes": str(payload.get("notes", "")).strip(),
            "tax_rate": tax_rate,
            "subtotal": declared_subtotal,
            "tax_amount": declared_tax_amount,
            "total": declared_total,
            "company_profile": company_profile,
            "html_path": html_path,
        }
        if "generated_at" in payload:
            validated["generated_at"] = str(payload["generated_at"]).strip()
        return validated

    def _validate_todo_payload(self, payload: dict) -> dict:
        title = str(payload.get("title", "")).strip()
        if not title:
            raise ValueError("Le titre est obligatoire.")
        task_date = self._parse_iso_date(payload["task_date"], "La date de tache").isoformat()
        task_time = self._parse_task_time(payload.get("task_time", ""))
        status = str(payload.get("status", "A faire")).strip() or "A faire"
        if status not in {"A faire", "En cours", "Terminee"}:
            raise ValueError("Statut de tache invalide.")
        return {
            "title": title,
            "task_date": task_date,
            "task_time": task_time,
            "status": status,
            "details": str(payload.get("details", "")).strip(),
        }

    def _delete_local_document_file(self, html_path: str | None) -> None:
        if not html_path:
            return
        try:
            target = Path(html_path).resolve()
        except Exception:
            return
        documents_root = (self.root / "documents").resolve()
        if documents_root not in target.parents:
            return
        try:
            target.unlink(missing_ok=True)
        except OSError:
            return

    def add_sale(self, payload: dict) -> None:
        validated = self._validate_financial_payload(payload, "sale_date")
        now = self._now()
        self.connection.execute(
            """
            INSERT INTO sales (
                sale_date, client, company_name, category, amount, amount_ht, amount_ttc,
                source_type, source_invoice_id, notes, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                validated["date_value"],
                validated["company_name"],
                validated["company_name"],
                validated["category"],
                validated["amount_ttc"],
                validated["amount_ht"],
                validated["amount_ttc"],
                validated["source_type"],
                validated["source_invoice_id"],
                validated["notes"],
                now,
                now,
            ),
        )
        self.connection.commit()

    def list_sales(self) -> list[sqlite3.Row]:
        return self.connection.execute(
            """
            SELECT
                id,
                sale_date,
                client,
                company_name,
                category,
                amount,
                amount_ht,
                amount_ttc,
                source_type,
                source_invoice_id,
                notes,
                created_at,
                updated_at
            FROM sales
            ORDER BY sale_date DESC, id DESC
            """
        ).fetchall()

    def delete_sale(self, sale_id: int) -> None:
        self.connection.execute("DELETE FROM sales WHERE id = ?", (sale_id,))
        self.connection.commit()

    def add_expense(self, payload: dict) -> None:
        validated = self._validate_financial_payload(payload, "expense_date")
        now = self._now()
        self.connection.execute(
            """
            INSERT INTO expenses (
                expense_date, vendor, company_name, category, amount, amount_ht, amount_ttc,
                source_type, source_invoice_id, notes, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                validated["date_value"],
                validated["company_name"],
                validated["company_name"],
                validated["category"],
                validated["amount_ttc"],
                validated["amount_ht"],
                validated["amount_ttc"],
                validated["source_type"],
                validated["source_invoice_id"],
                validated["notes"],
                now,
                now,
            ),
        )
        self.connection.commit()

    def list_expenses(self) -> list[sqlite3.Row]:
        return self.connection.execute(
            """
            SELECT
                id,
                expense_date,
                vendor,
                company_name,
                category,
                amount,
                amount_ht,
                amount_ttc,
                source_type,
                source_invoice_id,
                notes,
                created_at,
                updated_at
            FROM expenses
            ORDER BY expense_date DESC, id DESC
            """
        ).fetchall()

    def delete_expense(self, expense_id: int) -> None:
        self.connection.execute("DELETE FROM expenses WHERE id = ?", (expense_id,))
        self.connection.commit()

    def create_invoice(self, payload: dict) -> int:
        validated = self._validate_document_payload("invoice", payload)
        now = self._now()
        cursor = self.connection.execute(
            """
            INSERT INTO invoices (
                number, issue_date, due_date, status, client_name, client_email,
                client_address, items_json, notes, tax_rate, subtotal, tax_amount,
                total, company_json, html_path, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                validated["number"],
                validated["issue_date"],
                validated["due_date"],
                validated["status"],
                validated["client_name"],
                validated["client_email"],
                validated["client_address"],
                json.dumps(validated["items"], ensure_ascii=False),
                validated["notes"],
                validated["tax_rate"],
                validated["subtotal"],
                validated["tax_amount"],
                validated["total"],
                json.dumps(validated["company_profile"], ensure_ascii=False),
                validated["html_path"],
                now,
                now,
            ),
        )
        self.connection.commit()
        return int(cursor.lastrowid)

    def list_invoices(self) -> list[dict]:
        rows = self.connection.execute(
            "SELECT * FROM invoices ORDER BY issue_date DESC, id DESC"
        ).fetchall()
        return [self._deserialize_document(row) for row in rows]

    def update_invoice_status(self, invoice_id: int, status: str) -> None:
        clean_status = self._validate_document_status("invoice", status)
        self.connection.execute(
            "UPDATE invoices SET status = ?, updated_at = ? WHERE id = ?",
            (clean_status, self._now(), invoice_id),
        )
        self.connection.commit()

    def delete_invoice(self, invoice_id: int) -> None:
        row = self.connection.execute(
            "SELECT html_path FROM invoices WHERE id = ?",
            (invoice_id,),
        ).fetchone()
        self.connection.execute(
            """
            UPDATE sales
            SET source_type = 'manual',
                source_invoice_id = NULL,
                updated_at = ?,
                notes = TRIM(
                    CASE
                        WHEN notes IS NULL OR notes = '' THEN 'Lien facture supprime'
                        ELSE notes || ' | Lien facture supprime'
                    END
                )
            WHERE source_type = 'invoice' AND source_invoice_id = ?
            """,
            (self._now(), invoice_id),
        )
        self.connection.execute(
            """
            UPDATE expenses
            SET source_type = 'manual',
                source_invoice_id = NULL,
                updated_at = ?,
                notes = TRIM(
                    CASE
                        WHEN notes IS NULL OR notes = '' THEN 'Lien facture supprime'
                        ELSE notes || ' | Lien facture supprime'
                    END
                )
            WHERE source_type = 'invoice' AND source_invoice_id = ?
            """,
            (self._now(), invoice_id),
        )
        self.connection.execute("DELETE FROM invoices WHERE id = ?", (invoice_id,))
        self.connection.commit()
        self._delete_local_document_file(row["html_path"] if row else None)

    def create_quote(self, payload: dict) -> int:
        validated = self._validate_document_payload("quote", payload)
        now = self._now()
        cursor = self.connection.execute(
            """
            INSERT INTO quotes (
                number, issue_date, valid_until, status, client_name, client_email,
                client_address, items_json, notes, tax_rate, subtotal, tax_amount,
                total, company_json, html_path, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                validated["number"],
                validated["issue_date"],
                validated["valid_until"],
                validated["status"],
                validated["client_name"],
                validated["client_email"],
                validated["client_address"],
                json.dumps(validated["items"], ensure_ascii=False),
                validated["notes"],
                validated["tax_rate"],
                validated["subtotal"],
                validated["tax_amount"],
                validated["total"],
                json.dumps(validated["company_profile"], ensure_ascii=False),
                validated["html_path"],
                now,
                now,
            ),
        )
        self.connection.commit()
        return int(cursor.lastrowid)

    def list_quotes(self) -> list[dict]:
        rows = self.connection.execute(
            "SELECT * FROM quotes ORDER BY issue_date DESC, id DESC"
        ).fetchall()
        return [self._deserialize_document(row) for row in rows]

    def update_quote_status(self, quote_id: int, status: str) -> None:
        clean_status = self._validate_document_status("quote", status)
        self.connection.execute(
            "UPDATE quotes SET status = ?, updated_at = ? WHERE id = ?",
            (clean_status, self._now(), quote_id),
        )
        self.connection.commit()

    def delete_quote(self, quote_id: int) -> None:
        row = self.connection.execute(
            "SELECT html_path FROM quotes WHERE id = ?",
            (quote_id,),
        ).fetchone()
        self.connection.execute("DELETE FROM quotes WHERE id = ?", (quote_id,))
        self.connection.commit()
        self._delete_local_document_file(row["html_path"] if row else None)

    def add_todo(self, payload: dict) -> None:
        validated = self._validate_todo_payload(payload)
        now = self._now()
        self.connection.execute(
            """
            INSERT INTO todos (title, details, task_date, task_time, status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                validated["title"],
                validated["details"],
                validated["task_date"],
                validated["task_time"],
                validated["status"],
                now,
                now,
            ),
        )
        self.connection.commit()

    def update_todo(self, todo_id: int, payload: dict) -> None:
        validated = self._validate_todo_payload(payload)
        self.connection.execute(
            """
            UPDATE todos
            SET title = ?, details = ?, task_date = ?, task_time = ?, status = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                validated["title"],
                validated["details"],
                validated["task_date"],
                validated["task_time"],
                validated["status"],
                self._now(),
                todo_id,
            ),
        )
        self.connection.commit()

    def delete_todo(self, todo_id: int) -> None:
        self.connection.execute("DELETE FROM todos WHERE id = ?", (todo_id,))
        self.connection.commit()

    def list_todos(self) -> list[dict]:
        rows = self.connection.execute(
            """
            SELECT *
            FROM todos
            ORDER BY
                CASE WHEN status = 'Terminee' THEN 1 ELSE 0 END,
                task_date ASC,
                CASE WHEN TRIM(COALESCE(task_time, '')) = '' THEN '99:99' ELSE task_time END ASC,
                id DESC
            """
        ).fetchall()
        return [dict(row) for row in rows]

    def list_invoice_options(self) -> list[dict]:
        rows = self.connection.execute(
            """
            SELECT id, number, client_name, issue_date, due_date, subtotal, total, status
            FROM invoices
            ORDER BY issue_date DESC, id DESC
            """
        ).fetchall()
        return [dict(row) for row in rows]

    def next_document_number(self, kind: str) -> str:
        prefix = "FACT" if kind == "invoice" else "DEV"
        table = "invoices" if kind == "invoice" else "quotes"
        period = date.today().strftime("%Y%m")
        like_pattern = f"{prefix}-{period}-%"
        rows = self.connection.execute(
            f"SELECT number FROM {table} WHERE number LIKE ?",
            (like_pattern,),
        ).fetchall()
        sequence = 0
        for row in rows:
            try:
                sequence = max(sequence, int(row["number"].split("-")[-1]))
            except ValueError:
                continue
        return f"{prefix}-{period}-{sequence + 1:03d}"

    def dashboard_snapshot(self) -> dict:
        invoice_total = self._scalar("SELECT COALESCE(SUM(total), 0) FROM invoices")
        manual_sales_total = self._scalar(
            """
            SELECT COALESCE(SUM(CASE
                WHEN source_type = 'invoice' AND source_invoice_id IS NOT NULL THEN 0
                ELSE COALESCE(NULLIF(amount_ttc, 0), amount)
            END), 0)
            FROM sales
            """
        )
        expenses_total = self._scalar(
            "SELECT COALESCE(SUM(COALESCE(NULLIF(amount_ttc, 0), amount)), 0) FROM expenses"
        )
        revenue_total = float(invoice_total) + float(manual_sales_total)
        profit_total = revenue_total - float(expenses_total)

        open_invoices = self._scalar(
            "SELECT COUNT(*) FROM invoices WHERE status IN ('Brouillon', 'Envoyee', 'En retard')"
        )
        pending_quotes = self._scalar(
            "SELECT COUNT(*) FROM quotes WHERE status IN ('Brouillon', 'Envoye')"
        )
        pending_todos = self._scalar("SELECT COUNT(*) FROM todos WHERE status <> 'Terminee'")

        months = self._last_month_keys(6)
        revenue_map = {item["key"]: 0.0 for item in months}
        expense_map = {item["key"]: 0.0 for item in months}

        for row in self.connection.execute("SELECT issue_date, total FROM invoices").fetchall():
            key = row["issue_date"][:7]
            if key in revenue_map:
                revenue_map[key] += float(row["total"])

        for row in self.connection.execute(
            """
            SELECT sale_date, amount, amount_ttc, source_type, source_invoice_id
            FROM sales
            """
        ).fetchall():
            if row["source_type"] == "invoice" and row["source_invoice_id"] is not None:
                continue
            key = row["sale_date"][:7]
            if key in revenue_map:
                revenue_map[key] += float(row["amount_ttc"] or row["amount"])

        for row in self.connection.execute(
            "SELECT expense_date, amount, amount_ttc FROM expenses"
        ).fetchall():
            key = row["expense_date"][:7]
            if key in expense_map:
                expense_map[key] += float(row["amount_ttc"] or row["amount"])

        revenue_series = []
        profit_series = []
        for month in months:
            revenue = revenue_map[month["key"]]
            expenses = expense_map[month["key"]]
            revenue_series.append((month["label"], revenue))
            profit_series.append((month["label"], revenue - expenses))

        sales_categories = defaultdict(float)
        invoice_categories_total = self._scalar("SELECT COALESCE(SUM(total), 0) FROM invoices")
        if float(invoice_categories_total):
            sales_categories["Factures"] = float(invoice_categories_total)

        for row in self.connection.execute(
            """
            SELECT category, amount, amount_ttc, source_type, source_invoice_id
            FROM sales
            """
        ).fetchall():
            if row["source_type"] == "invoice" and row["source_invoice_id"] is not None:
                continue
            sales_categories[row["category"]] += float(row["amount_ttc"] or row["amount"])

        expense_categories = defaultdict(float)
        for row in self.connection.execute(
            "SELECT category, amount, amount_ttc FROM expenses"
        ).fetchall():
            expense_categories[row["category"]] += float(row["amount_ttc"] or row["amount"])

        latest_invoices = self.connection.execute(
            """
            SELECT number, client_name, total, status, created_at
            FROM invoices
            ORDER BY id DESC
            LIMIT 5
            """
        ).fetchall()
        latest_quotes = self.connection.execute(
            """
            SELECT number, client_name, total, status, created_at
            FROM quotes
            ORDER BY id DESC
            LIMIT 5
            """
        ).fetchall()
        latest_todos = self.connection.execute(
            """
            SELECT title, task_date, task_time, status
            FROM todos
            ORDER BY
                CASE WHEN status = 'Terminee' THEN 1 ELSE 0 END,
                task_date ASC,
                CASE WHEN TRIM(COALESCE(task_time, '')) = '' THEN '99:99' ELSE task_time END ASC
            LIMIT 5
            """
        ).fetchall()

        return {
            "revenue_total": revenue_total,
            "expenses_total": float(expenses_total),
            "profit_total": profit_total,
            "open_invoices": int(open_invoices),
            "pending_quotes": int(pending_quotes),
            "pending_todos": int(pending_todos),
            "revenue_series": revenue_series,
            "profit_series": profit_series,
            "sales_categories": self._top_categories(sales_categories),
            "expense_categories": self._top_categories(expense_categories),
            "latest_invoices": [dict(row) for row in latest_invoices],
            "latest_quotes": [dict(row) for row in latest_quotes],
            "latest_todos": [dict(row) for row in latest_todos],
        }

    def calendar_snapshot(self, year: int, month: int) -> dict:
        month_key = f"{year:04d}-{month:02d}"
        day_count = month_calendar.monthrange(year, month)[1]
        days = {
            f"{month_key}-{day:02d}": {
                "date": f"{month_key}-{day:02d}",
                "day": day,
                "revenue": 0.0,
                "expenses": 0.0,
                "balance": 0.0,
                "invoice_count": 0,
                "expense_count": 0,
                "todo_count": 0,
                "items": [],
            }
            for day in range(1, day_count + 1)
        }

        month_revenue = 0.0
        month_expenses = 0.0
        invoice_revenue = 0.0
        manual_revenue = 0.0

        invoice_rows = self.connection.execute(
            """
            SELECT id, number, client_name, issue_date, due_date, total, subtotal, status
            FROM invoices
            WHERE substr(issue_date, 1, 7) = ? OR substr(due_date, 1, 7) = ?
            ORDER BY issue_date ASC, id ASC
            """,
            (month_key, month_key),
        ).fetchall()
        for row in invoice_rows:
            if row["issue_date"].startswith(month_key):
                entry = days[row["issue_date"]]
                amount = float(row["total"])
                entry["revenue"] += amount
                entry["invoice_count"] += 1
                entry["items"].append(
                    {
                        "kind": "invoice",
                        "label": f"Facture {row['number']} - {row['client_name']}",
                        "amount": amount,
                        "status": row["status"],
                    }
                )
                month_revenue += amount
                invoice_revenue += amount
            if row["due_date"].startswith(month_key) and row["due_date"] != row["issue_date"]:
                days[row["due_date"]]["items"].append(
                    {
                        "kind": "invoice_due",
                        "label": f"Echeance {row['number']} - {row['client_name']}",
                        "amount": float(row["total"]),
                        "status": row["status"],
                    }
                )

        sale_rows = self.connection.execute(
            """
            SELECT id, sale_date, company_name, category, amount, amount_ht, amount_ttc, source_type, source_invoice_id
            FROM sales
            WHERE substr(sale_date, 1, 7) = ?
            ORDER BY sale_date ASC, id ASC
            """,
            (month_key,),
        ).fetchall()
        for row in sale_rows:
            if row["source_type"] == "invoice" and row["source_invoice_id"] is not None:
                continue
            amount = float(row["amount_ttc"] or row["amount"])
            entry = days[row["sale_date"]]
            entry["revenue"] += amount
            entry["items"].append(
                {
                    "kind": "sale",
                    "label": f"Recette manuelle - {row['company_name']}",
                    "amount": amount,
                    "status": row["category"],
                }
            )
            month_revenue += amount
            manual_revenue += amount

        expense_rows = self.connection.execute(
            """
            SELECT id, expense_date, company_name, category, amount, amount_ht, amount_ttc, source_type
            FROM expenses
            WHERE substr(expense_date, 1, 7) = ?
            ORDER BY expense_date ASC, id ASC
            """,
            (month_key,),
        ).fetchall()
        for row in expense_rows:
            amount = float(row["amount_ttc"] or row["amount"])
            entry = days[row["expense_date"]]
            entry["expenses"] += amount
            entry["expense_count"] += 1
            entry["items"].append(
                {
                    "kind": "expense",
                    "label": f"Depense - {row['company_name']}",
                    "amount": amount,
                    "status": row["category"],
                }
            )
            month_expenses += amount

        todo_rows = self.connection.execute(
            """
            SELECT id, title, task_date, task_time, status
            FROM todos
            WHERE substr(task_date, 1, 7) = ?
            ORDER BY task_date ASC, task_time ASC, id ASC
            """,
            (month_key,),
        ).fetchall()
        for row in todo_rows:
            entry = days[row["task_date"]]
            entry["todo_count"] += 1
            entry["items"].append(
                {
                    "kind": "todo",
                    "label": row["title"],
                    "amount": 0.0,
                    "status": f"{row['status']} {row['task_time']}".strip(),
                }
            )

        quote_rows = self.connection.execute(
            """
            SELECT number, client_name, valid_until, status, total
            FROM quotes
            WHERE substr(valid_until, 1, 7) = ?
            ORDER BY valid_until ASC, id ASC
            """,
            (month_key,),
        ).fetchall()
        for row in quote_rows:
            days[row["valid_until"]]["items"].append(
                {
                    "kind": "quote",
                    "label": f"Fin de validite devis {row['number']} - {row['client_name']}",
                    "amount": float(row["total"]),
                    "status": row["status"],
                }
            )

        for payload in days.values():
            payload["balance"] = payload["revenue"] - payload["expenses"]

        return {
            "year": year,
            "month": month,
            "days": days,
            "month_revenue": month_revenue,
            "month_expenses": month_expenses,
            "month_balance": month_revenue - month_expenses,
            "invoice_revenue": invoice_revenue,
            "manual_revenue": manual_revenue,
            "pending_todos": sum(1 for row in todo_rows if row["status"] != "Terminee"),
        }

    def _deserialize_document(self, row: sqlite3.Row) -> dict:
        payload = dict(row)
        payload["items"] = json.loads(payload.pop("items_json"))
        payload["company_profile"] = json.loads(payload.pop("company_json"))
        return payload

    def _scalar(self, query: str):
        return self.connection.execute(query).fetchone()[0]

    def _last_month_keys(self, length: int) -> list[dict]:
        cursor = date.today().replace(day=1)
        months = []
        for _ in range(length):
            months.append({"key": cursor.strftime("%Y-%m"), "label": cursor.strftime("%b")})
            if cursor.month == 1:
                cursor = cursor.replace(year=cursor.year - 1, month=12)
            else:
                cursor = cursor.replace(month=cursor.month - 1)
        months.reverse()
        return months

    def _top_categories(self, bucket: dict[str, float]) -> list[tuple[str, float]]:
        items = sorted(bucket.items(), key=lambda item: item[1], reverse=True)
        return items[:5]
