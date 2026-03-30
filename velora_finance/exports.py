from __future__ import annotations

import csv
import shutil
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ExportResult:
    export_root: Path
    copied_files: int
    missing_documents: list[str]
    report_path: Path


def month_key_from_date(value: str) -> str:
    return value[:7]


def collect_month_keys(rows: list[dict], date_field: str) -> list[str]:
    months = sorted({month_key_from_date(row[date_field]) for row in rows if row.get(date_field)}, reverse=True)
    return months


def export_month_bundle(kind: str, rows: list[dict], month_key: str, destination_root: Path) -> ExportResult:
    folder_names = {
        "invoice": "factures",
        "quote": "devis",
        "expense": "depenses",
    }
    if kind not in folder_names:
        raise ValueError("Type d'export inconnu.")

    export_root = Path(destination_root) / f"velora-export-{month_key}" / folder_names[kind]
    export_root.mkdir(parents=True, exist_ok=True)
    copied_files = 0
    missing_documents: list[str] = []

    if kind in {"invoice", "quote"}:
        copied_files, missing_documents = _export_documents(rows, month_key, export_root, kind)
    else:
        _export_expenses(rows, month_key, export_root)
    report_path = _write_export_report(export_root, kind, month_key, copied_files, missing_documents)
    return ExportResult(export_root=export_root, copied_files=copied_files, missing_documents=missing_documents, report_path=report_path)


def _export_documents(rows: list[dict], month_key: str, export_root: Path, kind: str) -> tuple[int, list[str]]:
    selected = [row for row in rows if row["issue_date"].startswith(month_key)]
    csv_path = export_root / f"{'factures' if kind == 'invoice' else 'devis'}-{month_key}.csv"
    copied_files = 0
    missing_documents: list[str] = []
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(["numero", "client", "date", "statut", "ht", "ttc", "html", "html_present"])
        for row in selected:
            source = Path(row["html_path"])
            html_present = source.exists()
            if source.exists():
                shutil.copy2(source, export_root / source.name)
                copied_files += 1
            else:
                missing_documents.append(row["number"])
            writer.writerow(
                [
                    row["number"],
                    row["client_name"],
                    row["issue_date"],
                    row["status"],
                    row["subtotal"],
                    row["total"],
                    source.name,
                    "oui" if html_present else "non",
                ]
            )
    if missing_documents:
        missing_path = export_root / f"documents-manquants-{month_key}.csv"
        with missing_path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(["numero"])
            for number in missing_documents:
                writer.writerow([number])
    return copied_files, missing_documents


def _export_expenses(rows: list[dict], month_key: str, export_root: Path) -> None:
    selected = [row for row in rows if row["expense_date"].startswith(month_key)]
    csv_path = export_root / f"depenses-{month_key}.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(["date", "entreprise", "categorie", "source", "ht", "ttc", "notes", "ajoute_le"])
        for row in selected:
            writer.writerow(
                [
                    row["expense_date"],
                    row["company_name"],
                    row["category"],
                    row["source_type"],
                    row["amount_ht"],
                    row["amount_ttc"],
                    row["notes"],
                    row["created_at"],
                ]
            )


def _write_export_report(export_root: Path, kind: str, month_key: str, copied_files: int, missing_documents: list[str]) -> Path:
    labels = {
        "invoice": "Factures",
        "quote": "Devis",
        "expense": "Depenses",
    }
    report_path = export_root / f"rapport-export-{month_key}.txt"
    lines = [
        f"Type: {labels[kind]}",
        f"Mois: {month_key}",
        f"Documents copies: {copied_files}",
        f"Documents manquants: {len(missing_documents)}",
    ]
    if missing_documents:
        lines.append("")
        lines.append("Numeros manquants:")
        lines.extend(missing_documents)
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return report_path
