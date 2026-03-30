from __future__ import annotations

import base64
import html
import re
from datetime import date
from pathlib import Path

from .config import document_directory


DOCUMENT_PREFIXES = {
    "invoice": "facture",
    "quote": "devis",
    "expense": "depense",
    "employee_expense": "paie",
    "sale": "recette",
    "other": "piece",
}


def money(value: float) -> str:
    rendered = f"{value:,.2f}".replace(",", " ").replace(".", ",")
    return f"{rendered} EUR"


def image_to_data_uri(path: str | None) -> str:
    if not path:
        return ""
    file_path = Path(path)
    if not file_path.exists():
        return ""
    mime = "image/png"
    if file_path.suffix.lower() == ".svg":
        mime = "image/svg+xml"
    if file_path.suffix.lower() in {".jpg", ".jpeg"}:
        mime = "image/jpeg"
    encoded = base64.b64encode(file_path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def safe_document_stem(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", str(value or "").strip())
    cleaned = re.sub(r"-{2,}", "-", cleaned)
    cleaned = cleaned.strip(".-_")
    return cleaned or "document"


def normalize_document_date(value: str | date | None) -> str:
    if isinstance(value, date):
        return value.isoformat()
    cleaned = str(value or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", cleaned):
        return cleaned
    return date.today().isoformat()


def build_document_filename(kind: str, subject: str, document_date: str | date | None = None, extension: str = ".html") -> str:
    prefix = DOCUMENT_PREFIXES.get(kind, "document")
    subject_stem = safe_document_stem(subject)
    date_stem = normalize_document_date(document_date)
    clean_extension = extension if extension.startswith(".") else f".{extension}"
    return f"{prefix}_{subject_stem}_{date_stem}{clean_extension.lower()}"


def _unique_path(target: Path) -> Path:
    if not target.exists():
        return target
    suffix = target.suffix
    stem = target.stem
    counter = 2
    while True:
        candidate = target.with_name(f"{stem}-{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def document_target_path(kind: str, subject: str, document_date: str | date | None = None, extension: str = ".html") -> Path:
    base = document_directory(kind).resolve()
    target = _unique_path((base / build_document_filename(kind, subject, document_date, extension)).resolve())
    if target.parent != base:
        raise ValueError("Chemin de document invalide.")
    return target


def save_document_html(kind: str, payload: dict, target: Path | None = None) -> Path:
    target = target or document_target_path(kind, payload.get("client_name", "document"), payload.get("issue_date"), ".html")
    target.parent.mkdir(parents=True, exist_ok=True)
    temp_target = target.with_suffix(f"{target.suffix}.tmp")
    temp_target.write_text(build_document_html(kind, payload), encoding="utf-8")
    temp_target.replace(target)
    return target


def build_document_html(kind: str, payload: dict) -> str:
    company = payload["company_profile"]
    logo_data = image_to_data_uri(company.get("logo_path"))
    doc_label = "Facture" if kind == "invoice" else "Devis"
    period_label = "Echeance" if kind == "invoice" else "Valide jusqu'au"
    period_value = payload["due_date"] if kind == "invoice" else payload["valid_until"]
    generated_at = payload.get("generated_at", "")

    items_rows = []
    for item in payload["items"]:
        total = float(item["quantity"]) * float(item["unit_price"])
        items_rows.append(
            f"""
            <tr>
                <td>{html.escape(item['description'])}</td>
                <td>{item['quantity']}</td>
                <td>{money(float(item['unit_price']))}</td>
                <td>{money(total)}</td>
            </tr>
            """
        )

    logo_block = (
        f'<img class="logo" src="{logo_data}" alt="Logo entreprise" />'
        if logo_data
        else f'<div class="wordmark">{html.escape(company.get("company_name", ""))}</div>'
    )

    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8" />
    <title>{doc_label} {payload['number']}</title>
    <style>
        :root {{
            color-scheme: light only;
        }}
        body {{
            background:
                radial-gradient(circle at top left, rgba(36, 107, 255, 0.08), transparent 34%),
                linear-gradient(180deg, #f3f7fb 0%, #eef4fb 100%);
            color: #132238;
            font-family: "Segoe UI", "Inter", sans-serif;
            margin: 0;
            padding: 32px;
        }}
        .sheet {{
            background: #ffffff;
            border: 1px solid #dce6f0;
            border-radius: 24px;
            box-shadow: 0 18px 48px rgba(16, 35, 61, 0.12);
            margin: 0 auto;
            max-width: 980px;
            overflow: hidden;
        }}
        .header {{
            align-items: flex-start;
            background: linear-gradient(135deg, #10233d 0%, #133154 58%, #0ea5a3 140%);
            color: #f7fbff;
            display: flex;
            gap: 24px;
            justify-content: space-between;
            padding: 32px 36px 28px;
        }}
        .brand {{
            align-items: center;
            display: flex;
            gap: 16px;
        }}
        .logo {{
            background: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.16);
            border-radius: 16px;
            height: 58px;
            object-fit: contain;
            padding: 6px;
            width: 58px;
        }}
        .wordmark {{
            align-items: center;
            color: #ffffff;
            display: flex;
            font-size: 24px;
            font-weight: 700;
            height: 58px;
            letter-spacing: -0.02em;
        }}
        .brand-copy p,
        .doc-meta p {{
            color: rgba(247, 251, 255, 0.82);
            margin: 6px 0 0 0;
        }}
        .doc-title {{
            font-size: 34px;
            font-weight: 700;
            letter-spacing: -0.03em;
            margin: 0 0 10px 0;
        }}
        .doc-meta {{
            min-width: 280px;
            text-align: right;
        }}
        .content {{
            padding: 30px 36px 36px;
        }}
        .grid {{
            display: grid;
            gap: 18px;
            grid-template-columns: 1.1fr 0.9fr;
            margin-bottom: 24px;
        }}
        .panel {{
            background: #f5f9fd;
            border: 1px solid #dce6f0;
            border-radius: 18px;
            padding: 18px 20px;
        }}
        .panel h3 {{
            color: #246bff;
            font-size: 12px;
            letter-spacing: 0.12em;
            margin: 0 0 12px 0;
            text-transform: uppercase;
        }}
        .panel p {{
            margin: 4px 0;
            white-space: pre-line;
        }}
        table {{
            border-collapse: collapse;
            margin-top: 8px;
            width: 100%;
        }}
        th, td {{
            border-bottom: 1px solid #e5edf5;
            padding: 14px 10px;
            text-align: left;
        }}
        th {{
            color: #5b6b82;
            font-size: 12px;
            letter-spacing: 0.06em;
            text-transform: uppercase;
        }}
        .totals {{
            display: grid;
            gap: 10px;
            justify-content: end;
            margin-top: 22px;
        }}
        .total-line {{
            align-items: center;
            display: flex;
            gap: 28px;
            justify-content: space-between;
            min-width: 280px;
        }}
        .grand-total {{
            background: linear-gradient(135deg, #eef5ff 0%, #e4fbf6 100%);
            border: 1px solid #d9ebff;
            border-radius: 16px;
            font-size: 20px;
            font-weight: 700;
            padding: 14px 18px;
        }}
        .notes {{
            margin-top: 26px;
        }}
        .footer {{
            border-top: 1px solid #e5edf5;
            color: #5b6b82;
            margin-top: 28px;
            padding-top: 18px;
        }}
    </style>
</head>
<body>
    <div class="sheet">
        <div class="header">
            <div class="brand">
                {logo_block}
                <div class="brand-copy">
                    <div class="doc-title">{doc_label}</div>
                    <p>{html.escape(company.get('company_name', ''))}</p>
                </div>
            </div>
            <div class="doc-meta">
                <p><strong>Numero:</strong> {html.escape(payload['number'])}</p>
                <p><strong>Date:</strong> {html.escape(payload['issue_date'])}</p>
                <p><strong>{period_label}:</strong> {html.escape(period_value)}</p>
                <p><strong>Statut:</strong> {html.escape(payload['status'])}</p>
                <p><strong>Ajoute le:</strong> {html.escape(generated_at)}</p>
            </div>
        </div>
        <div class="content">
            <div class="grid">
                <div class="panel">
                    <h3>Entreprise</h3>
                    <p><strong>{html.escape(company.get('company_name', ''))}</strong></p>
                    <p>{html.escape(company.get('legal_name', ''))}</p>
                    <p>{html.escape(company.get('address', ''))}</p>
                    <p>SIRET: {html.escape(company.get('siret', ''))}</p>
                    <p>TVA: {html.escape(company.get('vat_number', ''))}</p>
                    <p>{html.escape(company.get('email', ''))} - {html.escape(company.get('phone', ''))}</p>
                </div>
                <div class="panel">
                    <h3>Client</h3>
                    <p><strong>{html.escape(payload['client_name'])}</strong></p>
                    <p>{html.escape(payload.get('client_address', ''))}</p>
                    <p>{html.escape(payload.get('client_email', ''))}</p>
                </div>
            </div>

            <table>
                <thead>
                    <tr>
                        <th>Description</th>
                        <th>Quantite</th>
                        <th>Prix unitaire</th>
                        <th>Total</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(items_rows)}
                </tbody>
            </table>

            <div class="totals">
                <div class="total-line"><span>Sous-total</span><strong>{money(payload['subtotal'])}</strong></div>
                <div class="total-line"><span>TVA ({payload['tax_rate']} %)</span><strong>{money(payload['tax_amount'])}</strong></div>
                <div class="total-line grand-total"><span>Total</span><span>{money(payload['total'])}</span></div>
            </div>

            <div class="notes">
                <h3 style="margin-bottom:8px;">Notes</h3>
                <p>{html.escape(payload.get('notes', '') or 'Aucune note.')}</p>
            </div>

            <div class="footer">
                {html.escape(company.get('footer', ''))}
            </div>
        </div>
    </div>
</body>
</html>
"""
