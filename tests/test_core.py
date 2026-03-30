from __future__ import annotations

import shutil
import sqlite3
import unittest
from pathlib import Path

from velora_finance.config import storage_root
from velora_finance.database import Database
from velora_finance.documents import build_document_html, document_target_path, save_document_html
from velora_finance.exports import export_month_bundle


class VeloraCoreTests(unittest.TestCase):
    def setUp(self) -> None:
        self.db_path = Path("test_suite.db")
        self.export_root = Path("test_exports")
        if self.db_path.exists():
            self.db_path.unlink()
        if self.export_root.exists():
            shutil.rmtree(self.export_root)
        self.db = Database(self.db_path)
        self.company = self.db.get_company_profile()

    def tearDown(self) -> None:
        self.db.close()
        self.db_path.unlink(missing_ok=True)
        if self.export_root.exists():
            shutil.rmtree(self.export_root)
        (storage_root() / "documents" / "depenses" / "test-depense-piece.pdf").unlink(missing_ok=True)

    def _invoice_payload(self, number: str, client_name: str = "Client") -> dict:
        return {
            "number": number,
            "issue_date": "2026-03-10",
            "due_date": "2026-03-20",
            "status": "Brouillon",
            "client_name": client_name,
            "client_email": "",
            "client_address": "",
            "items": [{"description": "Mission", "quantity": 1, "unit_price": 100.0}],
            "notes": "",
            "tax_rate": 20.0,
            "subtotal": 100.0,
            "tax_amount": 20.0,
            "total": 120.0,
            "company_profile": self.company,
            "generated_at": "2026-03-30T12:00:00",
        }

    def test_document_target_path_sanitizes_number(self) -> None:
        target = document_target_path("invoice", "..\\evil/facture:test")
        self.assertEqual(target.name, "evil-facture-test.html")
        self.assertIn("documents", str(target))
        self.assertIn("factures", str(target))

    def test_generated_document_uses_simple_header_without_color_banner(self) -> None:
        payload = self._invoice_payload("FACT-HEADER-001", "Client Test")
        rendered = build_document_html("invoice", payload)
        self.assertIn('class="header"', rendered)
        self.assertNotIn('class="banner"', rendered)

    def test_duplicate_invoice_does_not_overwrite_existing_html(self) -> None:
        payload = self._invoice_payload("FACT-TEST-001", "Premier")
        html_path = document_target_path("invoice", payload["number"])
        payload["html_path"] = html_path
        self.db.create_invoice(payload)
        save_document_html("invoice", payload, html_path)

        duplicate = self._invoice_payload("FACT-TEST-001", "Second")
        original_content = html_path.read_text(encoding="utf-8")
        with self.assertRaises(sqlite3.IntegrityError):
            self.db.create_invoice({**duplicate, "html_path": html_path})
        self.assertEqual(html_path.read_text(encoding="utf-8"), original_content)
        html_path.unlink(missing_ok=True)

    def test_delete_invoice_detaches_linked_financial_rows(self) -> None:
        payload = self._invoice_payload("FACT-LINK-001", "Acme")
        html_path = document_target_path("invoice", payload["number"])
        invoice_id = self.db.create_invoice({**payload, "html_path": html_path})

        self.db.add_sale(
            {
                "sale_date": "2026-03-11",
                "company_name": "Acme",
                "category": "Facture",
                "amount_ht": 100.0,
                "amount_ttc": 120.0,
                "source_type": "invoice",
                "source_invoice_id": invoice_id,
                "notes": "",
            }
        )
        self.db.add_expense(
            {
                "expense_date": "2026-03-12",
                "company_name": "Acme",
                "category": "Achats",
                "amount_ht": 50.0,
                "amount_ttc": 60.0,
                "source_type": "invoice",
                "source_invoice_id": invoice_id,
                "notes": "",
            }
        )

        self.db.delete_invoice(invoice_id)
        sale = dict(self.db.list_sales()[0])
        expense = dict(self.db.list_expenses()[0])
        self.assertEqual(sale["source_type"], "manual")
        self.assertIsNone(sale["source_invoice_id"])
        self.assertEqual(expense["source_type"], "manual")
        self.assertIsNone(expense["source_invoice_id"])

    def test_delete_invoice_removes_local_html(self) -> None:
        payload = self._invoice_payload("FACT-DELETE-FILE-001", "Acme")
        html_path = document_target_path("invoice", payload["number"])
        invoice_id = self.db.create_invoice({**payload, "html_path": html_path})
        save_document_html("invoice", payload, html_path)

        self.assertTrue(html_path.exists())
        self.db.delete_invoice(invoice_id)
        self.assertFalse(html_path.exists())

    def test_delete_quote_removes_local_html(self) -> None:
        payload = {
            **self._invoice_payload("DEV-DELETE-FILE-001", "Acme"),
            "valid_until": "2026-03-25",
        }
        payload.pop("due_date")
        html_path = document_target_path("quote", payload["number"])
        quote_id = self.db.create_quote({**payload, "html_path": html_path})
        save_document_html("quote", payload, html_path)

        self.assertTrue(html_path.exists())
        self.db.delete_quote(quote_id)
        self.assertFalse(html_path.exists())

    def test_delete_expense_removes_local_attachment(self) -> None:
        attachment = storage_root() / "documents" / "depenses" / "test-depense-piece.pdf"
        attachment.write_text("piece", encoding="utf-8")
        self.db.add_expense(
            {
                "expense_date": "2026-03-12",
                "company_name": "Fournisseur",
                "category": "Achats",
                "amount_ht": 40.0,
                "amount_ttc": 48.0,
                "source_type": "manual",
                "source_invoice_id": None,
                "attachment_path": str(attachment),
                "notes": "",
            }
        )

        expense_id = int(self.db.list_expenses()[0]["id"])
        self.assertTrue(attachment.exists())
        self.db.delete_expense(expense_id)
        self.assertFalse(attachment.exists())

    def test_dashboard_avoids_double_counting_invoice_linked_sale(self) -> None:
        first_invoice = self._invoice_payload("FACT-DASH-001", "Acme")
        second_invoice = self._invoice_payload("FACT-DASH-002", "Beta")
        first_invoice["status"] = "Envoyee"
        second_invoice["status"] = "Payee"
        first_id = self.db.create_invoice({**first_invoice, "html_path": document_target_path("invoice", first_invoice["number"])})
        self.db.create_invoice({**second_invoice, "html_path": document_target_path("invoice", second_invoice["number"])})

        self.db.add_sale(
            {
                "sale_date": "2026-03-11",
                "company_name": "Acme",
                "category": "Facture",
                "amount_ht": 100.0,
                "amount_ttc": 120.0,
                "source_type": "invoice",
                "source_invoice_id": first_id,
                "notes": "",
            }
        )
        self.db.add_sale(
            {
                "sale_date": "2026-03-11",
                "company_name": "Conseil",
                "category": "Conseil",
                "amount_ht": 50.0,
                "amount_ttc": 60.0,
                "source_type": "manual",
                "source_invoice_id": None,
                "notes": "",
            }
        )
        snapshot = self.db.dashboard_snapshot()
        self.assertEqual(snapshot["revenue_total"], 300.0)
        self.assertEqual(snapshot["draft_revenue_total"], 0.0)

    def test_document_preferences_are_persisted(self) -> None:
        self.db.save_document_preferences(
            {
                "default_tax_rate": 5.5,
                "invoice_due_days": 45,
                "quote_validity_days": 15,
                "default_client_name": "Client modele",
                "default_client_email": "modele@test.fr",
                "default_client_address": "1 rue modele",
                "default_invoice_notes": "Note facture",
                "default_quote_notes": "Note devis",
                "default_invoice_status": "Envoyee",
                "default_quote_status": "Envoye",
                "ui_theme": "Sombre",
            }
        )
        preferences = self.db.get_document_preferences()
        self.assertEqual(preferences["default_tax_rate"], 5.5)
        self.assertEqual(preferences["invoice_due_days"], 45)
        self.assertEqual(preferences["quote_validity_days"], 15)
        self.assertEqual(preferences["default_client_name"], "Client modele")
        self.assertEqual(preferences["default_quote_status"], "Envoye")
        self.assertEqual(preferences["ui_theme"], "Sombre")

    def test_create_invoice_rejects_due_date_before_issue_date(self) -> None:
        payload = self._invoice_payload("FACT-BAD-DATE-001", "Acme")
        payload["due_date"] = "2026-03-01"
        with self.assertRaises(ValueError):
            self.db.create_invoice({**payload, "html_path": document_target_path("invoice", payload["number"])})

    def test_create_quote_rejects_valid_until_before_issue_date(self) -> None:
        payload = {
            **self._invoice_payload("DEV-BAD-DATE-001", "Acme"),
            "valid_until": "2026-03-01",
        }
        payload.pop("due_date")
        with self.assertRaises(ValueError):
            self.db.create_quote({**payload, "html_path": document_target_path("quote", payload["number"])})

    def test_save_company_profile_rejects_invalid_siret(self) -> None:
        with self.assertRaises(ValueError):
            self.db.save_company_profile(
                {
                    "company_name": "Velora",
                    "legal_name": "Velora SAS",
                    "siret": "1234",
                    "vat_number": "FR12345678901",
                    "email": "contact@velora.fr",
                    "phone": "0102030405",
                    "address": "1 rue test",
                    "footer": "Merci",
                    "logo_path": "",
                }
            )

    def test_save_document_preferences_rejects_invalid_default_client_email(self) -> None:
        with self.assertRaises(ValueError):
            self.db.save_document_preferences(
                {
                    "default_client_email": "email-invalide",
                }
            )

    def test_add_sale_rejects_invalid_data_at_database_level(self) -> None:
        with self.assertRaises(ValueError):
            self.db.add_sale(
                {
                    "sale_date": "2026-03-30",
                    "company_name": "Client Test",
                    "category": "Conseil",
                    "amount_ht": -50.0,
                    "amount_ttc": -60.0,
                    "source_type": "invoice",
                    "source_invoice_id": 999999,
                    "notes": "test",
                }
            )

    def test_add_todo_rejects_invalid_time_at_database_level(self) -> None:
        with self.assertRaises(ValueError):
            self.db.add_todo(
                {
                    "title": "Relance client",
                    "task_date": "2026-03-30",
                    "task_time": "25:61",
                    "status": "A faire",
                    "details": "",
                }
            )

    def test_add_employee_expense_persists_employee_fields(self) -> None:
        self.db.add_expense(
            {
                "expense_date": "2026-03-18",
                "company_name": "Alice Martin",
                "category": "Salaires",
                "expense_kind": "employee",
                "expense_label": "Prime",
                "employee_name": "Alice Martin",
                "payroll_month": "2026-03",
                "amount_ht": 1000.0,
                "amount_ttc": 1000.0,
                "source_type": "manual",
                "source_invoice_id": None,
                "notes": "Prime trimestrielle",
            }
        )

        row = dict(self.db.list_expenses("employee")[0])
        self.assertEqual(row["expense_kind"], "employee")
        self.assertEqual(row["employee_name"], "Alice Martin")
        self.assertEqual(row["expense_label"], "Prime")
        self.assertEqual(row["payroll_month"], "2026-03")

    def test_dashboard_snapshot_excludes_draft_invoices_from_health_revenue(self) -> None:
        draft_invoice = self._invoice_payload("FACT-DRAFT-001", "Draft Client")
        confirmed_invoice = self._invoice_payload("FACT-CONF-001", "Confirmed Client")
        confirmed_invoice["status"] = "Envoyee"
        self.db.create_invoice({**draft_invoice, "html_path": document_target_path("invoice", draft_invoice["number"])})
        self.db.create_invoice({**confirmed_invoice, "html_path": document_target_path("invoice", confirmed_invoice["number"])})

        snapshot = self.db.dashboard_snapshot()

        self.assertEqual(snapshot["revenue_total"], 120.0)
        self.assertEqual(snapshot["draft_revenue_total"], 120.0)
        self.assertEqual(snapshot["invoice_revenue_total"], 120.0)
        self.assertEqual(len(snapshot["activity_points"]), 1)

    def test_sqlite_schema_rejects_invalid_invoice_dates(self) -> None:
        with self.assertRaises(sqlite3.IntegrityError):
            self.db.connection.execute(
                """
                INSERT INTO invoices (
                    number, issue_date, due_date, status, client_name, client_email,
                    client_address, items_json, notes, tax_rate, subtotal, tax_amount,
                    total, company_json, html_path, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    "FACT-SCHEMA-001",
                    "2026-03-20",
                    "2026-03-01",
                    "Brouillon",
                    "Client Schema",
                    "",
                    "",
                    "[{\"description\":\"Test\",\"quantity\":1,\"unit_price\":100}]",
                    "",
                    20.0,
                    100.0,
                    20.0,
                    120.0,
                    "{}",
                    str(document_target_path("invoice", "FACT-SCHEMA-001")),
                    "2026-03-20T10:00:00",
                    "2026-03-20T10:00:00",
                ),
            )

    def test_sqlite_schema_rejects_missing_invoice_reference(self) -> None:
        with self.assertRaises(sqlite3.IntegrityError):
            self.db.connection.execute(
                """
                INSERT INTO sales (
                    sale_date, client, company_name, category, amount, amount_ht, amount_ttc,
                    source_type, source_invoice_id, notes, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    "2026-03-20",
                    "Client Schema",
                    "Client Schema",
                    "Facture",
                    120.0,
                    100.0,
                    120.0,
                    "invoice",
                    999999,
                    "",
                    "2026-03-20T10:00:00",
                    "2026-03-20T10:00:00",
                ),
            )

    def test_export_invoice_month_bundle_copies_html_and_csv(self) -> None:
        payload = self._invoice_payload("FACT-EXPORT-001", "Client Export")
        html_path = self.export_root / "source" / "facture-export.html"
        html_path.parent.mkdir(parents=True, exist_ok=True)
        payload["html_path"] = html_path
        self.db.create_invoice(payload)
        save_document_html("invoice", payload, html_path)

        export_result = export_month_bundle("invoice", self.db.list_invoices(), "2026-03", self.export_root)

        copied_html = export_result.export_root / html_path.name
        csv_file = export_result.export_root / "factures-2026-03.csv"
        self.assertTrue(copied_html.exists())
        self.assertTrue(csv_file.exists())
        self.assertIn("FACT-EXPORT-001", csv_file.read_text(encoding="utf-8"))
        self.assertEqual(export_result.missing_documents, [])

    def test_export_expense_month_bundle_creates_separate_csv(self) -> None:
        self.db.add_expense(
            {
                "expense_date": "2026-03-18",
                "company_name": "Fournisseur Export",
                "category": "Logiciels",
                "amount_ht": 80.0,
                "amount_ttc": 96.0,
                "source_type": "manual",
                "source_invoice_id": None,
                "notes": "Licence mensuelle",
            }
        )

        export_result = export_month_bundle(
            "expense",
            [dict(row) for row in self.db.list_expenses()],
            "2026-03",
            self.export_root,
        )

        csv_file = export_result.export_root / "depenses-2026-03.csv"
        self.assertTrue(csv_file.exists())
        content = csv_file.read_text(encoding="utf-8")
        self.assertIn("Fournisseur Export", content)
        self.assertIn("Licence mensuelle", content)

    def test_export_employee_expense_month_bundle_reports_missing_attachment(self) -> None:
        attachment = self.export_root / "missing" / "fiche-paie.pdf"
        self.db.add_expense(
            {
                "expense_date": "2026-03-18",
                "company_name": "Alice Martin",
                "category": "Salaires",
                "expense_kind": "employee",
                "expense_label": "Salaire",
                "employee_name": "Alice Martin",
                "payroll_month": "2026-03",
                "amount_ht": 1500.0,
                "amount_ttc": 1500.0,
                "source_type": "manual",
                "source_invoice_id": None,
                "attachment_path": str(attachment),
                "notes": "",
            }
        )

        export_result = export_month_bundle(
            "employee_expense",
            [dict(row) for row in self.db.list_expenses("employee")],
            "2026-03",
            self.export_root,
        )

        self.assertEqual(len(export_result.missing_documents), 1)
        self.assertIn("Alice Martin", export_result.missing_documents[0])
        missing_csv = export_result.export_root / "documents-manquants-2026-03.csv"
        self.assertTrue(missing_csv.exists())

    def test_reset_all_data_clears_database_and_local_documents(self) -> None:
        invoice_payload = self._invoice_payload("FACT-RESET-001", "Client Reset")
        invoice_payload["status"] = "Envoyee"
        invoice_html = document_target_path("invoice", invoice_payload["number"])
        self.db.create_invoice({**invoice_payload, "html_path": invoice_html})
        save_document_html("invoice", invoice_payload, invoice_html)

        expense_attachment = storage_root() / "documents" / "depenses" / "test-reset-piece.pdf"
        expense_attachment.write_text("piece", encoding="utf-8")
        self.db.add_expense(
            {
                "expense_date": "2026-03-20",
                "company_name": "Fournisseur Reset",
                "category": "Logiciels",
                "amount_ht": 10.0,
                "amount_ttc": 12.0,
                "source_type": "manual",
                "source_invoice_id": None,
                "attachment_path": str(expense_attachment),
                "notes": "",
            }
        )

        self.assertTrue(invoice_html.exists())
        self.assertTrue(expense_attachment.exists())

        self.db.reset_all_data()

        self.assertEqual(self.db.list_invoices(), [])
        self.assertEqual(self.db.list_expenses(), [])
        self.assertEqual(self.db.list_sales(), [])
        self.assertEqual(self.db.list_quotes(), [])
        self.assertEqual(self.db.list_todos(), [])
        self.assertFalse(invoice_html.exists())
        self.assertFalse(expense_attachment.exists())

    def test_export_invoice_month_bundle_reports_missing_html(self) -> None:
        payload = self._invoice_payload("FACT-EXPORT-MISSING-001", "Client Export")
        html_path = self.export_root / "source" / "facture-manquante.html"
        html_path.parent.mkdir(parents=True, exist_ok=True)
        payload["html_path"] = html_path
        self.db.create_invoice(payload)

        export_result = export_month_bundle("invoice", self.db.list_invoices(), "2026-03", self.export_root)

        self.assertEqual(export_result.missing_documents, ["FACT-EXPORT-MISSING-001"])
        missing_csv = export_result.export_root / "documents-manquants-2026-03.csv"
        self.assertTrue(missing_csv.exists())
        self.assertIn("FACT-EXPORT-MISSING-001", missing_csv.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
