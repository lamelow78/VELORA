from __future__ import annotations

import base64
import html
import re
from pathlib import Path

from .config import COLOR_CORAL, COLOR_NAVY, COLOR_TEAL, storage_root


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
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip())
    cleaned = cleaned.strip(".-")
    return cleaned or "document"


def document_target_path(kind: str, number: str) -> Path:
    folder = "factures" if kind == "invoice" else "devis"
    base = (storage_root() / "documents" / folder).resolve()
    target = (base / f"{safe_document_stem(number)}.html").resolve()
    if target.parent != base:
        raise ValueError("Numero de document invalide.")
    return target


def save_document_html(kind: str, payload: dict, target: Path | None = None) -> Path:
    target = target or document_target_path(kind, payload["number"])
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
        body {{
            background: linear-gradient(180deg, #f5f7fb 0%, #ffffff 100%);
            color: #17324d;
            font-family: "Segoe UI", sans-serif;
            margin: 0;
            padding: 40px;
        }}
        .sheet {{
            background: #ffffff;
            border-radius: 28px;
            box-shadow: 0 25px 80px rgba(21, 48, 75, 0.12);
            margin: 0 auto;
            max-width: 980px;
            overflow: hidden;
        }}
        .banner {{
            background: linear-gradient(120deg, {COLOR_NAVY} 0%, {COLOR_TEAL} 70%, {COLOR_CORAL} 100%);
            color: #ffffff;
            display: flex;
            justify-content: space-between;
            padding: 36px 44px;
        }}
        .logo {{
            background: rgba(255, 255, 255, 0.12);
            border-radius: 20px;
            height: 68px;
            object-fit: contain;
            padding: 10px;
            width: 68px;
        }}
        .wordmark {{
            align-items: center;
            background: rgba(255, 255, 255, 0.12);
            border-radius: 20px;
            display: inline-flex;
            font-size: 24px;
            font-weight: 700;
            height: 68px;
            justify-content: center;
            padding: 0 22px;
        }}
        .banner h1 {{
            font-size: 42px;
            margin: 0 0 6px 0;
        }}
        .banner p {{
            margin: 4px 0;
            opacity: 0.88;
        }}
        .content {{
            padding: 36px 44px 48px;
        }}
        .grid {{
            display: grid;
            gap: 20px;
            grid-template-columns: 1.1fr 0.9fr;
            margin-bottom: 28px;
        }}
        .panel {{
            background: #f8fafc;
            border: 1px solid #dce6f1;
            border-radius: 20px;
            padding: 20px 22px;
        }}
        .panel h3 {{
            font-size: 12px;
            letter-spacing: 0.12em;
            margin: 0 0 10px 0;
            text-transform: uppercase;
        }}
        .panel p {{
            margin: 4px 0;
            white-space: pre-line;
        }}
        table {{
            border-collapse: collapse;
            margin-top: 24px;
            width: 100%;
        }}
        th, td {{
            border-bottom: 1px solid #e4ebf4;
            padding: 14px 10px;
            text-align: left;
        }}
        th {{
            color: #607287;
            font-size: 12px;
            letter-spacing: 0.05em;
            text-transform: uppercase;
        }}
        .totals {{
            display: grid;
            gap: 8px;
            justify-content: end;
            margin-top: 22px;
        }}
        .total-line {{
            display: flex;
            gap: 26px;
            justify-content: space-between;
            min-width: 260px;
        }}
        .grand-total {{
            background: #f3f7fb;
            border-radius: 16px;
            font-size: 20px;
            font-weight: 700;
            padding: 14px 18px;
        }}
        .notes {{
            margin-top: 24px;
        }}
        .footer {{
            color: #607287;
            margin-top: 28px;
        }}
    </style>
</head>
<body>
    <div class="sheet">
        <div class="banner">
            <div>
                {logo_block}
            </div>
            <div style="text-align:right;">
                <h1>{doc_label}</h1>
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
