from __future__ import annotations

import calendar as pycalendar
import tkinter as tk
import webbrowser
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import (
    APP_NAME,
    APP_TAGLINE,
    COLOR_BG,
    COLOR_BORDER,
    COLOR_CORAL,
    COLOR_DANGER,
    COLOR_GOLD,
    COLOR_NAVY,
    COLOR_NAVY_DEEP,
    COLOR_SUCCESS,
    COLOR_SURFACE,
    COLOR_SURFACE_SOFT,
    COLOR_TEAL,
    COLOR_TEXT,
    COLOR_TEXT_MUTED,
    asset_path,
    storage_root,
)
from .database import Database
from .documents import document_target_path, money, save_document_html
from .exports import collect_month_keys, export_month_bundle


FRENCH_MONTHS = [
    "janvier",
    "fevrier",
    "mars",
    "avril",
    "mai",
    "juin",
    "juillet",
    "aout",
    "septembre",
    "octobre",
    "novembre",
    "decembre",
]


def main() -> None:
    app = VeloraApp()
    app.mainloop()


def parse_amount(value: str) -> float:
    cleaned = value.replace(" ", "").replace(",", ".").strip()
    if not cleaned:
        raise ValueError("Montant manquant.")
    return float(cleaned)


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def parse_time_text(value: str) -> str:
    cleaned = value.strip()
    if not cleaned:
        return ""
    return datetime.strptime(cleaned, "%H:%M").strftime("%H:%M")


def validate_money_pair(amount_ht: float, amount_ttc: float) -> None:
    if amount_ht < 0 or amount_ttc < 0:
        raise ValueError("Les montants HT et TTC doivent etre positifs.")
    if amount_ttc < amount_ht:
        raise ValueError("Le montant TTC doit etre superieur ou egal au montant HT.")


def validate_tax_rate(tax_rate: float) -> None:
    if tax_rate < 0 or tax_rate > 100:
        raise ValueError("Le taux de TVA doit etre compris entre 0 et 100.")


def format_timestamp(value: str) -> str:
    if not value:
        return "-"
    try:
        return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return value


def format_day(value: str) -> str:
    try:
        return parse_iso_date(value).strftime("%d/%m/%Y")
    except ValueError:
        return value


def short_money(value: float) -> str:
    absolute = abs(value)
    if absolute >= 1000:
        return f"{value / 1000:.1f}k EUR".replace(".", ",")
    return f"{int(round(value))} EUR"


def month_label(year: int, month: int) -> str:
    return f"{FRENCH_MONTHS[month - 1].capitalize()} {year}"


def health_message(balance: float, revenue: float, expenses: float) -> str:
    if revenue == 0 and expenses == 0:
        return "Calme"
    if balance > 0:
        return "Sain"
    if balance < 0:
        return "Sous tension"
    return "Equilibre"


def health_cell_color(balance: float, revenue: float, expenses: float, todo_count: int) -> str:
    if balance > 0 and revenue > 0:
        return "#daf2e7"
    if balance < 0 or expenses > revenue:
        return "#fde5df"
    if todo_count:
        return "#fbf0da"
    return COLOR_SURFACE_SOFT


def source_label(source_type: str) -> str:
    return "Facture" if source_type == "invoice" else "Manuel"


class MetricCard(ttk.Frame):
    def __init__(self, master, title: str, accent: str) -> None:
        super().__init__(master, style="Surface.TFrame", padding=18)
        self.value_var = tk.StringVar(value="0")
        self.subtitle_var = tk.StringVar(value="")

        dot = tk.Canvas(self, width=12, height=12, bg=COLOR_SURFACE, highlightthickness=0)
        dot.create_oval(1, 1, 11, 11, fill=accent, outline=accent)
        dot.grid(row=0, column=0, sticky="w")

        ttk.Label(self, text=title, style="CardTitle.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(self, textvariable=self.value_var, style="MetricValue.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(12, 4))
        ttk.Label(self, textvariable=self.subtitle_var, style="MutedSurface.TLabel").grid(row=2, column=0, columnspan=2, sticky="w")
        self.columnconfigure(1, weight=1)

    def set(self, value: str, subtitle: str) -> None:
        self.value_var.set(value)
        self.subtitle_var.set(subtitle)


class LineChart(tk.Canvas):
    def __init__(self, master, **kwargs) -> None:
        super().__init__(master, bg=COLOR_SURFACE, highlightthickness=0, height=250, **kwargs)
        self.labels: list[str] = []
        self.series: list[tuple[str, list[float], str]] = []
        self.bind("<Configure>", lambda _event: self.redraw())

    def update_data(self, labels: list[str], series: list[tuple[str, list[float], str]]) -> None:
        self.labels = labels
        self.series = series
        self.redraw()

    def redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 420)
        height = max(self.winfo_height(), 240)
        if not self.labels or not self.series:
            self.create_text(width / 2, height / 2, text="Aucune donnee pour le moment", fill=COLOR_TEXT_MUTED, font=("Segoe UI", 12))
            return

        max_value = max([max(values or [0]) for _, values, _ in self.series] + [1])
        left, top, right, bottom = 56, 28, width - 26, height - 40
        plot_width = max(right - left, 1)
        plot_height = max(bottom - top, 1)

        for step in range(5):
            y = top + plot_height * step / 4
            self.create_line(left, y, right, y, fill="#e8edf4", width=1)
            value = max_value - (max_value * step / 4)
            self.create_text(left - 10, y, text=f"{int(value)}", fill=COLOR_TEXT_MUTED, anchor="e")

        count = max(len(self.labels) - 1, 1)
        for index, label in enumerate(self.labels):
            x = left + plot_width * index / count
            self.create_text(x, bottom + 18, text=label, fill=COLOR_TEXT_MUTED)

        for series_index, (name, values, color) in enumerate(self.series):
            points = []
            for index, value in enumerate(values):
                x = left + plot_width * index / count
                y = bottom - (value / max_value) * plot_height if max_value else bottom
                points.extend([x, y])
            if len(points) >= 4:
                self.create_line(*points, fill=color, smooth=True, width=3)
            for point_index in range(0, len(points), 2):
                x, y = points[point_index], points[point_index + 1]
                self.create_oval(x - 4, y - 4, x + 4, y + 4, fill=color, outline=COLOR_SURFACE)
            legend_x = left + (140 * series_index)
            self.create_rectangle(legend_x, 8, legend_x + 10, 18, fill=color, outline=color)
            self.create_text(legend_x + 16, 13, text=name, anchor="w", fill=COLOR_TEXT, font=("Segoe UI", 10, "bold"))


class HorizontalBarChart(tk.Canvas):
    def __init__(self, master, **kwargs) -> None:
        super().__init__(master, bg=COLOR_SURFACE, highlightthickness=0, height=250, **kwargs)
        self.data: list[tuple[str, float]] = []
        self.bar_color = COLOR_TEAL
        self.bind("<Configure>", lambda _event: self.redraw())

    def update_data(self, data: list[tuple[str, float]], bar_color: str) -> None:
        self.data = data
        self.bar_color = bar_color
        self.redraw()

    def redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 340)
        height = max(self.winfo_height(), 240)
        if not self.data:
            self.create_text(width / 2, height / 2, text="Aucune categorie a afficher", fill=COLOR_TEXT_MUTED, font=("Segoe UI", 12))
            return

        left = 18
        top = 18
        row_height = 38
        max_value = max(value for _, value in self.data) or 1
        for index, (label, value) in enumerate(self.data):
            y = top + index * row_height
            self.create_text(left, y + 12, text=label, fill=COLOR_TEXT, anchor="w", font=("Segoe UI", 10, "bold"))
            self.create_rectangle(left, y + 20, width - 92, y + 30, fill="#edf2f8", outline="#edf2f8")
            ratio = value / max_value
            self.create_rectangle(left, y + 20, left + (width - 110) * ratio, y + 30, fill=self.bar_color, outline=self.bar_color)
            self.create_text(width - 16, y + 25, text=money(value), fill=COLOR_TEXT_MUTED, anchor="e", font=("Segoe UI", 10))


class DashboardPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app

        header = ttk.Frame(self, style="App.TFrame")
        header.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(header, text="Tableau de bord", style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="Vue globale du chiffre d'affaires, des benefices, des factures et des taches.", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        cards = ttk.Frame(self, style="App.TFrame")
        cards.pack(fill="x", padx=24)
        cards.columnconfigure((0, 1, 2, 3, 4), weight=1)
        self.revenue_card = MetricCard(cards, "Chiffre d'affaires", COLOR_TEAL)
        self.profit_card = MetricCard(cards, "Benefice estime", COLOR_CORAL)
        self.expense_card = MetricCard(cards, "Depenses", COLOR_GOLD)
        self.invoice_card = MetricCard(cards, "Factures / devis", COLOR_NAVY)
        self.todo_card = MetricCard(cards, "Taches ouvertes", COLOR_SUCCESS)
        for column, widget in enumerate(
            [self.revenue_card, self.profit_card, self.expense_card, self.invoice_card, self.todo_card]
        ):
            widget.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 12, 0))

        charts = ttk.Frame(self, style="App.TFrame")
        charts.pack(fill="both", expand=True, padx=24, pady=20)
        charts.columnconfigure(0, weight=3)
        charts.columnconfigure(1, weight=2)
        charts.rowconfigure(0, weight=1)
        charts.rowconfigure(1, weight=1)

        trend_card = self._chart_card(charts, "Evolution sur 6 mois", "Toutes les factures comptent automatiquement dans le CA")
        trend_card.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        self.trend_chart = LineChart(trend_card)
        self.trend_chart.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        sales_card = self._chart_card(charts, "Recettes par categorie", "Factures + recettes manuelles")
        sales_card.grid(row=0, column=1, sticky="nsew")
        self.sales_chart = HorizontalBarChart(sales_card)
        self.sales_chart.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        expense_card = self._chart_card(charts, "Depenses par categorie", "Pour surveiller la sante de l'entreprise")
        expense_card.grid(row=1, column=1, sticky="nsew", pady=(12, 0))
        self.expense_chart = HorizontalBarChart(expense_card)
        self.expense_chart.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        foot = ttk.Frame(self, style="App.TFrame")
        foot.pack(fill="x", padx=24, pady=(0, 24))
        foot.columnconfigure((0, 1, 2), weight=1)
        self.invoices_panel = self._latest_panel(foot, "Dernieres factures")
        self.invoices_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.quotes_panel = self._latest_panel(foot, "Derniers devis")
        self.quotes_panel.grid(row=0, column=1, sticky="nsew", padx=(0, 12))
        self.todos_panel = self._latest_panel(foot, "Todo liste")
        self.todos_panel.grid(row=0, column=2, sticky="nsew")

    def _chart_card(self, parent, title: str, subtitle: str) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame", padding=18)
        ttk.Label(frame, text=title, style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(frame, text=subtitle, style="MutedSurface.TLabel").pack(anchor="w", pady=(2, 14))
        return frame

    def _latest_panel(self, parent, title: str) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame", padding=18)
        ttk.Label(frame, text=title, style="SectionTitle.TLabel").pack(anchor="w")
        body = ttk.Frame(frame, style="Surface.TFrame")
        body.pack(fill="x", pady=(12, 0))
        frame.body = body
        return frame

    def _fill_latest_panel(self, container: ttk.Frame, rows: list[dict], kind: str) -> None:
        for child in container.body.winfo_children():
            child.destroy()
        if not rows:
            ttk.Label(container.body, text="Aucun element pour le moment.", style="MutedSurface.TLabel").pack(anchor="w")
            return
        for row in rows:
            line = ttk.Frame(container.body, style="Surface.TFrame")
            line.pack(fill="x", pady=4)
            if kind == "todo":
                left = f"{row['task_date']} {row['task_time']}".strip()
                ttk.Label(line, text=row["title"], style="TableStrong.TLabel").pack(side="left")
                ttk.Label(line, text=left, style="MutedSurface.TLabel").pack(side="left", padx=10)
                ttk.Label(line, text=row["status"], style="BadgeSurface.TLabel").pack(side="right")
            else:
                ttk.Label(line, text=row["number"], style="TableStrong.TLabel").pack(side="left")
                ttk.Label(line, text=f"{row['client_name']} - {money(row['total'])}", style="MutedSurface.TLabel").pack(side="left", padx=10)
                ttk.Label(line, text=row["status"], style="BadgeSurface.TLabel").pack(side="right")

    def refresh(self) -> None:
        snapshot = self.app.db.dashboard_snapshot()
        self.revenue_card.set(money(snapshot["revenue_total"]), "Factures creees + recettes manuelles")
        self.profit_card.set(money(snapshot["profit_total"]), "CA - depenses TTC")
        self.expense_card.set(money(snapshot["expenses_total"]), "Toutes les depenses enregistrees")
        self.invoice_card.set(f"{snapshot['open_invoices']} / {snapshot['pending_quotes']}", "Factures ouvertes / devis en attente")
        self.todo_card.set(str(snapshot["pending_todos"]), "Taches encore a traiter")

        labels = [label for label, _ in snapshot["revenue_series"]]
        revenue_values = [value for _, value in snapshot["revenue_series"]]
        profit_values = [value for _, value in snapshot["profit_series"]]
        self.trend_chart.update_data(labels, [("CA", revenue_values, COLOR_TEAL), ("Benefice", profit_values, COLOR_CORAL)])
        self.sales_chart.update_data(snapshot["sales_categories"], COLOR_TEAL)
        self.expense_chart.update_data(snapshot["expense_categories"], COLOR_CORAL)
        self._fill_latest_panel(self.invoices_panel, snapshot["latest_invoices"], "doc")
        self._fill_latest_panel(self.quotes_panel, snapshot["latest_quotes"], "doc")
        self._fill_latest_panel(self.todos_panel, snapshot["latest_todos"], "todo")


class BaseEntryPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp", title: str, subtitle: str) -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        head = ttk.Frame(self, style="App.TFrame")
        head.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(head, text=title, style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(head, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        self.content = ttk.Frame(self, style="App.TFrame")
        self.content.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        self.content.columnconfigure(0, weight=2)
        self.content.columnconfigure(1, weight=3)
        self.form_card = ttk.Frame(self.content, style="Surface.TFrame", padding=18)
        self.form_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.table_card = ttk.Frame(self.content, style="Surface.TFrame", padding=18)
        self.table_card.grid(row=0, column=1, sticky="nsew")

    def _create_label(self, master, text: str, row: int, column: int = 0) -> None:
        ttk.Label(master, text=text, style="FieldLabel.TLabel").grid(row=row, column=column, sticky="w", pady=(0, 6))


class FinancialEntryPage(BaseEntryPage):
    def __init__(
        self,
        master,
        app: "VeloraApp",
        title: str,
        subtitle: str,
        record_kind: str,
        default_category: str,
        categories: list[str],
        action_label: str,
        button_style: str,
    ) -> None:
        super().__init__(master, app, title, subtitle)
        self.record_kind = record_kind
        self.categories = categories
        self.action_label = action_label
        self.button_style = button_style
        self.rows: dict[str, dict] = {}
        self.invoice_map: dict[str, dict] = {}

        self.date_var = tk.StringVar(value=date.today().isoformat())
        self.source_mode_var = tk.StringVar(value="Manuel")
        self.invoice_choice_var = tk.StringVar()
        self.company_var = tk.StringVar()
        self.category_var = tk.StringVar(value=default_category)
        self.amount_ht_var = tk.StringVar()
        self.amount_ttc_var = tk.StringVar()
        self.notes_var = tk.StringVar()

        self._build_form()
        self._build_table()
        self.refresh_invoice_options()

    def _build_form(self) -> None:
        form_title = "Nouvel enregistrement" if self.record_kind == "sale" else "Nouvelle depense"
        ttk.Label(self.form_card, text=form_title, style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")

        fields = [
            ("Date", self.date_var),
            ("Mode", self.source_mode_var),
            ("Facture source", self.invoice_choice_var),
            ("Nom de l'entreprise", self.company_var),
            ("Categorie", self.category_var),
            ("Montant HT", self.amount_ht_var),
            ("Montant TTC", self.amount_ttc_var),
            ("Notes", self.notes_var),
        ]
        for index, (label, variable) in enumerate(fields, start=1):
            self._create_label(self.form_card, label, index * 2 - 1)
            if label == "Mode":
                widget = ttk.Combobox(self.form_card, textvariable=variable, values=["Manuel", "Depuis facture"], state="readonly")
                widget.bind("<<ComboboxSelected>>", lambda _event: self.on_mode_change())
            elif label == "Facture source":
                widget = ttk.Combobox(self.form_card, textvariable=variable, state="disabled")
                widget.bind("<<ComboboxSelected>>", lambda _event: self.apply_invoice_source())
                self.invoice_widget = widget
            elif label == "Categorie":
                widget = ttk.Combobox(self.form_card, textvariable=variable, values=self.categories, state="readonly")
            else:
                widget = ttk.Entry(self.form_card, textvariable=variable)
            widget.grid(row=index * 2, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        ttk.Button(self.form_card, text=self.action_label, style=self.button_style, command=self.add_record).grid(row=18, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(
            self.form_card,
            text="Mode facture: les informations HT/TTC et le nom client sont recuperes automatiquement. Mode manuel: vous saisissez vos propres montants.",
            style="MutedSurface.TLabel",
            wraplength=260,
        ).grid(row=19, column=0, sticky="w", pady=(12, 0))
        self.form_card.columnconfigure(0, weight=1)

    def _build_table(self) -> None:
        top = ttk.Frame(self.table_card, style="Surface.TFrame")
        top.pack(fill="x")
        table_title = "Historique des recettes" if self.record_kind == "sale" else "Historique des depenses"
        ttk.Label(top, text=table_title, style="SectionTitle.TLabel").pack(side="left")
        ttk.Button(top, text="Supprimer la selection", style="Ghost.TButton", command=self.delete_selected).pack(side="right")

        columns = ("date", "company", "source", "category", "ht", "ttc", "added")
        self.tree = ttk.Treeview(self.table_card, columns=columns, show="headings", height=18)
        headings = {
            "date": "Date",
            "company": "Entreprise",
            "source": "Source",
            "category": "Categorie",
            "ht": "HT",
            "ttc": "TTC",
            "added": "Ajout",
        }
        widths = {"date": 95, "company": 165, "source": 90, "category": 110, "ht": 90, "ttc": 90, "added": 130}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.pack(fill="both", expand=True, pady=(12, 0))

    def refresh_invoice_options(self) -> None:
        self.invoice_map.clear()
        values = []
        for row in self.app.db.list_invoice_options():
            label = f"{row['number']} | {row['client_name']} | {money(row['total'])}"
            values.append(label)
            self.invoice_map[label] = row
        self.invoice_widget.configure(values=values)

    def on_mode_change(self) -> None:
        if self.source_mode_var.get() == "Depuis facture":
            self.invoice_widget.configure(state="readonly")
            values = list(self.invoice_widget.cget("values"))
            if values:
                self.invoice_choice_var.set(values[0])
                self.apply_invoice_source()
        else:
            self.invoice_widget.configure(state="disabled")
            self.invoice_choice_var.set("")

    def apply_invoice_source(self) -> None:
        invoice = self.invoice_map.get(self.invoice_choice_var.get())
        if not invoice:
            return
        self.company_var.set(invoice["client_name"])
        self.amount_ht_var.set(str(invoice["subtotal"]))
        self.amount_ttc_var.set(str(invoice["total"]))
        if self.record_kind == "sale":
            self.category_var.set("Facture")
        self.notes_var.set(f"Rattache a la facture {invoice['number']}")

    def add_record(self) -> None:
        try:
            parse_iso_date(self.date_var.get())
            source_type = "invoice" if self.source_mode_var.get() == "Depuis facture" else "manual"
            source_invoice_id = None
            if source_type == "invoice":
                invoice = self.invoice_map.get(self.invoice_choice_var.get())
                if not invoice:
                    raise ValueError("Choisissez une facture source.")
                source_invoice_id = invoice["id"]
            company_name = self.company_var.get().strip() or "Entreprise"
            amount_ht = parse_amount(self.amount_ht_var.get())
            amount_ttc = parse_amount(self.amount_ttc_var.get())
            validate_money_pair(amount_ht, amount_ttc)
            payload = {
                self.date_key: self.date_var.get(),
                "company_name": company_name,
                "category": self.category_var.get().strip() or self.categories[0],
                "amount_ht": amount_ht,
                "amount_ttc": amount_ttc,
                "source_type": source_type,
                "source_invoice_id": source_invoice_id,
                "notes": self.notes_var.get().strip(),
            }
        except Exception as exc:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer.\n\n{exc}")
            return

        if self.record_kind == "sale":
            self.app.db.add_sale(payload)
        else:
            self.app.db.add_expense(payload)

        self.company_var.set("")
        self.amount_ht_var.set("")
        self.amount_ttc_var.set("")
        self.notes_var.set("")
        if self.source_mode_var.get() == "Depuis facture":
            self.apply_invoice_source()
        self.app.refresh_all_pages()

    def refresh(self) -> None:
        self.refresh_invoice_options()
        self.rows.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = self.app.db.list_sales() if self.record_kind == "sale" else self.app.db.list_expenses()
        date_field = self.date_key
        for row in rows:
            record = dict(row)
            iid = str(record["id"])
            self.rows[iid] = record
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    record[date_field],
                    record["company_name"],
                    source_label(record["source_type"]),
                    record["category"],
                    money(record["amount_ht"]),
                    money(record["amount_ttc"]),
                    format_timestamp(record["created_at"]),
                ),
            )

    def delete_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        if not messagebox.askyesno("Supprimer", "Supprimer la ligne selectionnee ?"):
            return
        record_id = int(selection[0])
        if self.record_kind == "sale":
            self.app.db.delete_sale(record_id)
        else:
            self.app.db.delete_expense(record_id)
        self.app.refresh_all_pages()


class SalesPage(FinancialEntryPage):
    date_key = "sale_date"

    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(
            master,
            app,
            "Recettes",
            "Les factures creees generent automatiquement du chiffre d'affaires. Cette page permet d'ajouter des recettes manuelles ou des lignes liees a une facture.",
            "sale",
            "Conseil",
            ["Facture", "Conseil", "Abonnement", "Projet", "Formation", "Autre"],
            "Ajouter la recette",
            "Accent.TButton",
        )


class ExpensesPage(FinancialEntryPage):
    date_key = "expense_date"

    def __init__(self, master, app: "VeloraApp") -> None:
        self.export_month_var = tk.StringVar(master=app)
        super().__init__(
            master,
            app,
            "Depenses",
            "Enregistrez les depenses en HT et TTC. Vous pouvez recuperer automatiquement les informations d'une facture ou saisir les montants manuellement.",
            "expense",
            "Achats",
            ["Achats", "Marketing", "Logiciels", "Salaires", "Transport", "Sous-traitance", "Autre"],
            "Ajouter la depense",
            "AccentAlt.TButton",
        )

    def _build_table(self) -> None:
        top = ttk.Frame(self.table_card, style="Surface.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="Historique des depenses", style="SectionTitle.TLabel").pack(side="left")
        ttk.Button(top, text="Supprimer la selection", style="Ghost.TButton", command=self.delete_selected).pack(side="right")

        export_bar = ttk.Frame(self.table_card, style="Surface.TFrame")
        export_bar.pack(fill="x", pady=(12, 12))
        ttk.Label(export_bar, text="Export local du mois", style="FieldLabel.TLabel").pack(side="left")
        self.export_month_widget = ttk.Combobox(export_bar, textvariable=self.export_month_var, state="readonly", width=12)
        self.export_month_widget.pack(side="left", padx=(10, 8))
        ttk.Button(export_bar, text="Exporter le dossier", style="Ghost.TButton", command=self.export_selected_month).pack(side="left")

        columns = ("date", "company", "source", "category", "ht", "ttc", "added")
        self.tree = ttk.Treeview(self.table_card, columns=columns, show="headings", height=18)
        headings = {
            "date": "Date",
            "company": "Entreprise",
            "source": "Source",
            "category": "Categorie",
            "ht": "HT",
            "ttc": "TTC",
            "added": "Ajout",
        }
        widths = {"date": 95, "company": 165, "source": 90, "category": 110, "ht": 90, "ttc": 90, "added": 130}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.pack(fill="both", expand=True, pady=(0, 0))

    def refresh(self) -> None:
        super().refresh()
        months = collect_month_keys([dict(row) for row in self.app.db.list_expenses()], "expense_date")
        self.export_month_widget.configure(values=months)
        if months:
            if self.export_month_var.get() not in months:
                self.export_month_var.set(months[0])
        else:
            self.export_month_var.set("")

    def export_selected_month(self) -> None:
        month_key = self.export_month_var.get().strip()
        if not month_key:
            messagebox.showwarning("Export", "Choisissez un mois a exporter.")
            return
        destination = filedialog.askdirectory(title="Choisir un dossier pour l'export des depenses")
        if not destination:
            return
        export_result = export_month_bundle(
            "expense",
            [dict(row) for row in self.app.db.list_expenses()],
            month_key,
            Path(destination),
        )
        messagebox.showinfo(
            "Export depenses",
            "Les depenses de "
            f"{month_key} ont ete exportees localement dans:\n\n{export_result.export_root}\n\n"
            f"Rapport genere: {export_result.report_path.name}",
        )


class DocumentPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp", title: str, subtitle: str, kind: str) -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.kind = kind
        self.records: dict[str, dict] = {}
        self.export_month_var = tk.StringVar()

        head = ttk.Frame(self, style="App.TFrame")
        head.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(head, text=title, style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(head, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(self, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        self.table_card = ttk.Frame(body, style="Surface.TFrame", padding=18)
        self.table_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.detail_card = ttk.Frame(body, style="Surface.TFrame", padding=18)
        self.detail_card.grid(row=0, column=1, sticky="nsew")

        self._build_table()
        self._build_detail()

    def _build_table(self) -> None:
        ttk.Label(self.table_card, text="Documents", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(self.table_card, text="HT, TTC, date du document, apercu et export mensuel local.", style="MutedSurface.TLabel").pack(anchor="w", pady=(2, 12))

        actions = ttk.Frame(self.table_card, style="Surface.TFrame")
        actions.pack(fill="x", pady=(0, 12))
        new_label = "Nouvelle facture" if self.kind == "invoice" else "Nouveau devis"
        ttk.Button(actions, text=new_label, style="Accent.TButton", command=self.open_editor).pack(side="left")
        ttk.Button(actions, text="Ouvrir HTML", style="Ghost.TButton", command=self.open_selected_html).pack(side="left", padx=8)
        ttk.Button(actions, text="Supprimer", style="Danger.TButton", command=self.delete_selected).pack(side="right")

        export_actions = ttk.Frame(self.table_card, style="Surface.TFrame")
        export_actions.pack(fill="x", pady=(0, 12))
        ttk.Label(export_actions, text="Export local du mois", style="FieldLabel.TLabel").pack(side="left")
        self.export_month_widget = ttk.Combobox(export_actions, textvariable=self.export_month_var, state="readonly", width=12)
        self.export_month_widget.pack(side="left", padx=(10, 8))
        ttk.Button(export_actions, text="Exporter le dossier", style="Ghost.TButton", command=self.export_selected_month).pack(side="left")

        status_actions = ttk.Frame(self.table_card, style="Surface.TFrame")
        status_actions.pack(fill="x", pady=(0, 12))
        for label in self.available_statuses():
            ttk.Button(status_actions, text=label, style="Ghost.TButton", command=lambda value=label: self.change_status(value)).pack(side="left", padx=(0, 8))

        columns = ("number", "client", "date", "deadline", "added", "status", "ttc")
        self.tree = ttk.Treeview(self.table_card, columns=columns, show="headings", height=18)
        headings = {
            "number": "Numero",
            "client": "Client",
            "date": "Date",
            "deadline": "Echeance" if self.kind == "invoice" else "Validite",
            "added": "Ajout",
            "status": "Statut",
            "ttc": "TTC",
        }
        widths = {"number": 140, "client": 170, "date": 95, "deadline": 95, "added": 130, "status": 100, "ttc": 95}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", lambda _event: self.show_detail())

    def _build_detail(self) -> None:
        ttk.Label(self.detail_card, text="Apercu visuel", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(self.detail_card, text="Le document reste stocke localement sur la machine et peut etre ouvert en HTML.", style="MutedSurface.TLabel").pack(anchor="w", pady=(2, 12))

        preview = ttk.Frame(self.detail_card, style="SurfaceSoft.TFrame", padding=16)
        preview.pack(fill="both", expand=True)
        self.preview_title_var = tk.StringVar(value="Aucun document selectionne")
        self.preview_meta_var = tk.StringVar(value="")
        self.preview_company_var = tk.StringVar(value="")
        self.preview_amounts_var = tk.StringVar(value="")
        self.preview_path_var = tk.StringVar(value="")

        ttk.Label(preview, textvariable=self.preview_title_var, style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(preview, textvariable=self.preview_meta_var, style="MutedSoft.TLabel").pack(anchor="w", pady=(4, 6))
        ttk.Label(preview, textvariable=self.preview_company_var, style="TableStrongSurface.TLabel").pack(anchor="w", pady=(0, 12))
        ttk.Label(preview, textvariable=self.preview_amounts_var, style="MetricMiniSoft.TLabel").pack(anchor="w", pady=(0, 12))

        self.preview_items = ttk.Treeview(preview, columns=("description", "qty", "price", "total"), show="headings", height=8)
        for column, heading, width in [("description", "Description", 220), ("qty", "Qt", 50), ("price", "Prix HT", 90), ("total", "Total HT", 90)]:
            self.preview_items.heading(column, text=heading)
            self.preview_items.column(column, width=width, anchor="w")
        self.preview_items.pack(fill="x", pady=(0, 12))

        ttk.Label(preview, text="Notes", style="SectionTitle.TLabel").pack(anchor="w")
        self.preview_notes = tk.Text(preview, bg=COLOR_SURFACE_SOFT, bd=0, fg=COLOR_TEXT, height=5, font=("Segoe UI", 10), wrap="word", highlightthickness=0)
        self.preview_notes.pack(fill="x", pady=(6, 12))
        self.preview_notes.configure(state="disabled")

        ttk.Label(preview, text="Fichier local", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(preview, textvariable=self.preview_path_var, style="MutedSoft.TLabel", wraplength=360).pack(anchor="w", pady=(6, 0))

    def available_statuses(self) -> list[str]:
        if self.kind == "invoice":
            return ["Brouillon", "Envoyee", "Payee", "En retard"]
        return ["Brouillon", "Envoye", "Accepte", "Refuse", "Expire"]

    def open_editor(self) -> None:
        DocumentEditor(self.app, self.kind)

    def selected_record(self) -> dict | None:
        selection = self.tree.selection()
        if not selection:
            return None
        return self.records.get(selection[0])

    def show_detail(self) -> None:
        record = self.selected_record()
        if not record:
            self.preview_title_var.set("Aucun document selectionne")
            self.preview_meta_var.set("")
            self.preview_company_var.set("")
            self.preview_amounts_var.set("")
            self.preview_path_var.set("")
            self.preview_notes.configure(state="normal")
            self.preview_notes.delete("1.0", "end")
            self.preview_notes.configure(state="disabled")
            for item in self.preview_items.get_children():
                self.preview_items.delete(item)
            return
        deadline_label = "Echeance" if self.kind == "invoice" else "Valide jusqu'au"
        deadline_value = record["due_date"] if self.kind == "invoice" else record["valid_until"]
        self.preview_title_var.set(f"{'Facture' if self.kind == 'invoice' else 'Devis'} {record['number']}")
        self.preview_meta_var.set(
            f"Date {record['issue_date']} | {deadline_label} {deadline_value} | Statut {record['status']} | Ajoute le {format_timestamp(record['created_at'])}"
        )
        self.preview_company_var.set(
            f"Client: {record['client_name']} | {record.get('client_email', '') or 'Sans email'}"
        )
        self.preview_amounts_var.set(
            f"HT {money(record['subtotal'])} | TVA {money(record['tax_amount'])} | TTC {money(record['total'])}"
        )
        self.preview_path_var.set(record["html_path"])
        for item in self.preview_items.get_children():
            self.preview_items.delete(item)
        for index, item in enumerate(record["items"]):
            total = item["quantity"] * item["unit_price"]
            self.preview_items.insert("", "end", iid=str(index), values=(item["description"], item["quantity"], money(item["unit_price"]), money(total)))
        self.preview_notes.configure(state="normal")
        self.preview_notes.delete("1.0", "end")
        self.preview_notes.insert("1.0", record.get("notes", "") or "Aucune note.")
        self.preview_notes.configure(state="disabled")

    def refresh(self) -> None:
        self.records.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        dataset = self.app.db.list_invoices() if self.kind == "invoice" else self.app.db.list_quotes()
        months = collect_month_keys(dataset, "issue_date")
        self.export_month_widget.configure(values=months)
        if months:
            if self.export_month_var.get() not in months:
                self.export_month_var.set(months[0])
        else:
            self.export_month_var.set("")
        for row in dataset:
            iid = str(row["id"])
            self.records[iid] = row
            deadline = row["due_date"] if self.kind == "invoice" else row["valid_until"]
            self.tree.insert("", "end", iid=iid, values=(row["number"], row["client_name"], row["issue_date"], deadline, format_timestamp(row["created_at"]), row["status"], money(row["total"])))
        self.show_detail()

    def export_selected_month(self) -> None:
        month_key = self.export_month_var.get().strip()
        if not month_key:
            messagebox.showwarning("Export", "Choisissez un mois a exporter.")
            return
        destination = filedialog.askdirectory(title=f"Choisir un dossier pour l'export des {'factures' if self.kind == 'invoice' else 'devis'}")
        if not destination:
            return
        kind = "invoice" if self.kind == "invoice" else "quote"
        export_result = export_month_bundle(kind, list(self.records.values()), month_key, Path(destination))
        if export_result.missing_documents:
            messagebox.showwarning(
                "Export",
                "L'export est termine, mais certains documents HTML etaient manquants.\n\n"
                f"Dossier: {export_result.export_root}\n"
                f"Rapport: {export_result.report_path.name}\n"
                f"Documents manquants: {len(export_result.missing_documents)}",
            )
            return
        messagebox.showinfo(
            "Export",
            f"Les {'factures' if self.kind == 'invoice' else 'devis'} de {month_key} ont ete exportes localement dans:\n\n"
            f"{export_result.export_root}\n\nRapport genere: {export_result.report_path.name}",
        )

    def open_selected_html(self) -> None:
        record = self.selected_record()
        if not record:
            return
        webbrowser.open(Path(record["html_path"]).as_uri())

    def change_status(self, status: str) -> None:
        record = self.selected_record()
        if not record:
            return
        if self.kind == "invoice":
            self.app.db.update_invoice_status(record["id"], status)
        else:
            self.app.db.update_quote_status(record["id"], status)
        self.app.refresh_all_pages()

    def delete_selected(self) -> None:
        record = self.selected_record()
        if not record:
            return
        if not messagebox.askyesno("Supprimer", "Supprimer ce document ?"):
            return
        if self.kind == "invoice":
            self.app.db.delete_invoice(record["id"])
        else:
            self.app.db.delete_quote(record["id"])
        self.app.refresh_all_pages()


class CompanyPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.logo_var = tk.StringVar()

        head = ttk.Frame(self, style="App.TFrame")
        head.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(head, text="Entreprise", style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(head, text="Personnalisez le nom, le logo, le SIRET, la TVA et les coordonnees.", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(self, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        body.columnconfigure(0, weight=2)
        body.columnconfigure(1, weight=1)
        self.form_card = ttk.Frame(body, style="Surface.TFrame", padding=18)
        self.form_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.info_card = ttk.Frame(body, style="Surface.TFrame", padding=18)
        self.info_card.grid(row=0, column=1, sticky="nsew")

        self.fields = {
            "company_name": tk.StringVar(),
            "legal_name": tk.StringVar(),
            "siret": tk.StringVar(),
            "vat_number": tk.StringVar(),
            "email": tk.StringVar(),
            "phone": tk.StringVar(),
        }
        self.document_fields = {
            "default_tax_rate": tk.StringVar(),
            "invoice_due_days": tk.StringVar(),
            "quote_validity_days": tk.StringVar(),
            "default_client_name": tk.StringVar(),
            "default_client_email": tk.StringVar(),
            "default_invoice_status": tk.StringVar(),
            "default_quote_status": tk.StringVar(),
        }

        self._build_form()
        self._build_info_card()

    def _build_form(self) -> None:
        ttk.Label(self.form_card, text="Identite entreprise", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        labels = [
            ("Nom commercial", "company_name"),
            ("Raison sociale", "legal_name"),
            ("SIRET", "siret"),
            ("Numero TVA", "vat_number"),
            ("Email", "email"),
            ("Telephone", "phone"),
        ]
        row_cursor = 1
        for label, key in labels:
            ttk.Label(self.form_card, text=label, style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
            ttk.Entry(self.form_card, textvariable=self.fields[key]).grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
            row_cursor += 2

        ttk.Label(self.form_card, text="Adresse", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
        self.address_text = tk.Text(self.form_card, height=4, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.address_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        row_cursor += 2

        ttk.Label(self.form_card, text="Message bas de page", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
        self.footer_text = tk.Text(self.form_card, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.footer_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        row_cursor += 2

        ttk.Label(self.form_card, text="Logo", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
        ttk.Entry(self.form_card, textvariable=self.logo_var).grid(row=row_cursor + 1, column=0, sticky="ew")
        actions = ttk.Frame(self.form_card, style="Surface.TFrame")
        actions.grid(row=row_cursor + 1, column=1, sticky="e", padx=(12, 0))
        ttk.Button(actions, text="Parcourir", style="Ghost.TButton", command=self.pick_logo).pack(side="left")
        ttk.Button(actions, text="Logo Velora", style="Ghost.TButton", command=self.use_default_logo).pack(side="left", padx=(8, 0))
        row_cursor += 2

        ttk.Separator(self.form_card, orient="horizontal").grid(row=row_cursor, column=0, columnspan=2, sticky="ew", pady=(18, 12))
        row_cursor += 1

        ttk.Label(self.form_card, text="Parametres facture / devis", style="SectionTitle.TLabel").grid(row=row_cursor, column=0, columnspan=2, sticky="w")
        row_cursor += 1

        document_labels = [
            ("TVA par defaut %", "default_tax_rate"),
            ("Delai facture en jours", "invoice_due_days"),
            ("Validite devis en jours", "quote_validity_days"),
            ("Client par defaut", "default_client_name"),
            ("Email client par defaut", "default_client_email"),
            ("Statut facture par defaut", "default_invoice_status"),
            ("Statut devis par defaut", "default_quote_status"),
        ]
        for label, key in document_labels:
            ttk.Label(self.form_card, text=label, style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
            if key == "default_invoice_status":
                widget = ttk.Combobox(self.form_card, textvariable=self.document_fields[key], values=["Brouillon", "Envoyee", "Payee", "En retard"], state="readonly")
            elif key == "default_quote_status":
                widget = ttk.Combobox(self.form_card, textvariable=self.document_fields[key], values=["Brouillon", "Envoye", "Accepte", "Refuse", "Expire"], state="readonly")
            else:
                widget = ttk.Entry(self.form_card, textvariable=self.document_fields[key])
            widget.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
            row_cursor += 2

        ttk.Label(self.form_card, text="Adresse client par defaut", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
        self.default_client_address_text = tk.Text(self.form_card, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.default_client_address_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        row_cursor += 2

        ttk.Label(self.form_card, text="Notes facture par defaut", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
        self.default_invoice_notes_text = tk.Text(self.form_card, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.default_invoice_notes_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        row_cursor += 2

        ttk.Label(self.form_card, text="Notes devis par defaut", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(10, 6))
        self.default_quote_notes_text = tk.Text(self.form_card, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.default_quote_notes_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        row_cursor += 2

        ttk.Button(self.form_card, text="Enregistrer les informations", style="Accent.TButton", command=self.save).grid(row=row_cursor + 1, column=0, sticky="ew", pady=(18, 0))
        self.form_card.columnconfigure(0, weight=1)

    def _build_info_card(self) -> None:
        ttk.Label(self.info_card, text="Stockage local", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(self.info_card, text="Toutes les factures, depenses, taches et la base sont enregistrees localement sur cette machine.", style="MutedSurface.TLabel", wraplength=280).pack(anchor="w", pady=(4, 12))

        storage = storage_root()
        info_lines = [
            ("Base SQLite", str(storage / "velora_finance.db")),
            ("Factures HTML", str(storage / "documents" / "factures")),
            ("Devis HTML", str(storage / "documents" / "devis")),
        ]
        for title, path in info_lines:
            block = ttk.Frame(self.info_card, style="SurfaceSoft.TFrame", padding=14)
            block.pack(fill="x", pady=6)
            ttk.Label(block, text=title, style="TableStrongSurface.TLabel").pack(anchor="w")
            ttk.Label(block, text=path, style="MutedSoft.TLabel", wraplength=260).pack(anchor="w", pady=(4, 0))

        ttk.Label(self.info_card, text="Rappel finance", style="SectionTitle.TLabel").pack(anchor="w", pady=(18, 6))
        ttk.Label(self.info_card, text="Chaque facture creee genere automatiquement du chiffre d'affaires dans le tableau de bord et le calendrier. Les depenses TTC sont deduites pour suivre la sante de l'entreprise.", style="MutedSurface.TLabel", wraplength=280).pack(anchor="w")

    def pick_logo(self) -> None:
        filename = filedialog.askopenfilename(title="Choisir un logo", filetypes=[("Images", "*.png *.gif *.ppm *.jpg *.jpeg *.svg"), ("Tous les fichiers", "*.*")])
        if filename:
            self.logo_var.set(filename)

    def use_default_logo(self) -> None:
        self.logo_var.set(str(asset_path("velora_logo.png")))

    def save(self) -> None:
        try:
            profile = {key: variable.get().strip() for key, variable in self.fields.items()}
            profile["address"] = self.address_text.get("1.0", "end").strip()
            profile["footer"] = self.footer_text.get("1.0", "end").strip()
            profile["logo_path"] = self.logo_var.get().strip()
            preferences = {
                "default_tax_rate": parse_amount(self.document_fields["default_tax_rate"].get()),
                "invoice_due_days": int(self.document_fields["invoice_due_days"].get().strip()),
                "quote_validity_days": int(self.document_fields["quote_validity_days"].get().strip()),
                "default_client_name": self.document_fields["default_client_name"].get().strip(),
                "default_client_email": self.document_fields["default_client_email"].get().strip(),
                "default_client_address": self.default_client_address_text.get("1.0", "end").strip(),
                "default_invoice_notes": self.default_invoice_notes_text.get("1.0", "end").strip(),
                "default_quote_notes": self.default_quote_notes_text.get("1.0", "end").strip(),
                "default_invoice_status": self.document_fields["default_invoice_status"].get().strip() or "Brouillon",
                "default_quote_status": self.document_fields["default_quote_status"].get().strip() or "Brouillon",
            }
            validate_tax_rate(float(preferences["default_tax_rate"]))
            if preferences["invoice_due_days"] < 0 or preferences["quote_validity_days"] < 0:
                raise ValueError("Les delais doivent etre positifs.")
        except Exception as exc:
            messagebox.showerror("Entreprise", f"Impossible d'enregistrer les parametres.\n\n{exc}")
            return

        self.app.db.save_company_profile(profile)
        self.app.db.save_document_preferences(preferences)
        messagebox.showinfo("Entreprise", "Les informations entreprise et les parametres de documents ont bien ete enregistres.")
        self.app.refresh_all_pages()

    def refresh(self) -> None:
        profile = self.app.db.get_company_profile()
        preferences = self.app.db.get_document_preferences()
        for key, variable in self.fields.items():
            variable.set(profile.get(key, ""))
        for key, variable in self.document_fields.items():
            variable.set(str(preferences.get(key, "")))
        self.address_text.delete("1.0", "end")
        self.address_text.insert("1.0", profile.get("address", ""))
        self.footer_text.delete("1.0", "end")
        self.footer_text.insert("1.0", profile.get("footer", ""))
        self.default_client_address_text.delete("1.0", "end")
        self.default_client_address_text.insert("1.0", preferences.get("default_client_address", ""))
        self.default_invoice_notes_text.delete("1.0", "end")
        self.default_invoice_notes_text.insert("1.0", preferences.get("default_invoice_notes", ""))
        self.default_quote_notes_text.delete("1.0", "end")
        self.default_quote_notes_text.insert("1.0", preferences.get("default_quote_notes", ""))
        self.logo_var.set(profile.get("logo_path", ""))


class TodoPage(BaseEntryPage):
    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(master, app, "Todo liste", "Ajoutez, modifiez, supprimez et datez vos taches avec heure.")
        self.selected_todo_id: int | None = None
        self.title_var = tk.StringVar()
        self.date_var = tk.StringVar(value=date.today().isoformat())
        self.time_var = tk.StringVar(value="09:00")
        self.status_var = tk.StringVar(value="A faire")

        self._build_form()
        self._build_table()

    def _build_form(self) -> None:
        ttk.Label(self.form_card, text="Nouvelle tache", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        fields = [
            ("Titre", self.title_var),
            ("Date", self.date_var),
            ("Heure", self.time_var),
            ("Statut", self.status_var),
        ]
        for index, (label, variable) in enumerate(fields, start=1):
            self._create_label(self.form_card, label, index * 2 - 1)
            if label == "Statut":
                widget = ttk.Combobox(self.form_card, textvariable=variable, values=["A faire", "En cours", "Terminee"], state="readonly")
            else:
                widget = ttk.Entry(self.form_card, textvariable=variable)
            widget.grid(row=index * 2, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        self._create_label(self.form_card, "Details", 9)
        self.details_text = tk.Text(self.form_card, height=7, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.details_text.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        actions = ttk.Frame(self.form_card, style="Surface.TFrame")
        actions.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Ajouter", style="Accent.TButton", command=self.save_todo).pack(side="left")
        ttk.Button(actions, text="Mettre a jour", style="Ghost.TButton", command=self.update_todo).pack(side="left", padx=8)
        ttk.Button(actions, text="Vider", style="Ghost.TButton", command=self.clear_form).pack(side="left")
        self.todo_info_label = ttk.Label(self.form_card, text="", style="MutedSurface.TLabel", wraplength=260)
        self.todo_info_label.grid(row=12, column=0, sticky="w", pady=(12, 0))
        self.form_card.columnconfigure(0, weight=1)

    def _build_table(self) -> None:
        top = ttk.Frame(self.table_card, style="Surface.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="Taches", style="SectionTitle.TLabel").pack(side="left")
        ttk.Button(top, text="Supprimer", style="Danger.TButton", command=self.delete_selected).pack(side="right")

        columns = ("date", "time", "title", "status", "added")
        self.tree = ttk.Treeview(self.table_card, columns=columns, show="headings", height=16)
        headings = {"date": "Date", "time": "Heure", "title": "Titre", "status": "Statut", "added": "Ajout"}
        widths = {"date": 95, "time": 70, "title": 220, "status": 100, "added": 130}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.pack(fill="both", expand=True, pady=(12, 0))
        self.tree.bind("<<TreeviewSelect>>", lambda _event: self.load_selected())

    def _payload(self) -> dict:
        parse_iso_date(self.date_var.get())
        task_time = parse_time_text(self.time_var.get())
        title = self.title_var.get().strip()
        if not title:
            raise ValueError("Le titre est obligatoire.")
        return {
            "title": title,
            "task_date": self.date_var.get(),
            "task_time": task_time,
            "status": self.status_var.get(),
            "details": self.details_text.get("1.0", "end").strip(),
        }

    def save_todo(self) -> None:
        try:
            self.app.db.add_todo(self._payload())
        except Exception as exc:
            messagebox.showerror("Todo", f"Impossible d'ajouter la tache.\n\n{exc}")
            return
        self.clear_form()
        self.app.refresh_all_pages()

    def update_todo(self) -> None:
        if self.selected_todo_id is None:
            return
        try:
            self.app.db.update_todo(self.selected_todo_id, self._payload())
        except Exception as exc:
            messagebox.showerror("Todo", f"Impossible de modifier la tache.\n\n{exc}")
            return
        self.clear_form()
        self.app.refresh_all_pages()

    def delete_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        if not messagebox.askyesno("Todo", "Supprimer cette tache ?"):
            return
        self.app.db.delete_todo(int(selection[0]))
        self.clear_form()
        self.app.refresh_all_pages()

    def load_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        record = self.rows[selection[0]]
        self.selected_todo_id = record["id"]
        self.title_var.set(record["title"])
        self.date_var.set(record["task_date"])
        self.time_var.set(record["task_time"])
        self.status_var.set(record["status"])
        self.details_text.delete("1.0", "end")
        self.details_text.insert("1.0", record["details"])
        self.todo_info_label.configure(text=f"Ajoute le {format_timestamp(record['created_at'])}\nMis a jour le {format_timestamp(record['updated_at'])}")

    def clear_form(self) -> None:
        self.selected_todo_id = None
        self.title_var.set("")
        self.date_var.set(date.today().isoformat())
        self.time_var.set("09:00")
        self.status_var.set("A faire")
        self.details_text.delete("1.0", "end")
        self.todo_info_label.configure(text="")

    def refresh(self) -> None:
        self.rows = {}
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.app.db.list_todos():
            iid = str(row["id"])
            self.rows[iid] = row
            self.tree.insert("", "end", iid=iid, values=(row["task_date"], row["task_time"], row["title"], row["status"], format_timestamp(row["created_at"])))


class CalendarPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        today = date.today()
        self.current_year = today.year
        self.current_month = today.month
        self.selected_date = today.isoformat()
        self.snapshot: dict = {}

        header = ttk.Frame(self, style="App.TFrame")
        header.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(header, text="Calendrier sante", style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="Visualisez chaque jour l'equilibre entre recettes, depenses, echeances et taches.", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        controls = ttk.Frame(self, style="App.TFrame")
        controls.pack(fill="x", padx=24)
        ttk.Button(controls, text="Mois precedent", style="Ghost.TButton", command=self.previous_month).pack(side="left")
        ttk.Button(controls, text="Aujourd'hui", style="Ghost.TButton", command=self.current_month_view).pack(side="left", padx=8)
        ttk.Button(controls, text="Mois suivant", style="Ghost.TButton", command=self.next_month).pack(side="left")
        self.month_var = tk.StringVar()
        ttk.Label(controls, textvariable=self.month_var, style="PageTitle.TLabel").pack(side="right")

        self.summary = ttk.Frame(self, style="App.TFrame")
        self.summary.pack(fill="x", padx=24, pady=18)
        self.summary.columnconfigure((0, 1, 2, 3), weight=1)
        self.month_revenue_card = MetricCard(self.summary, "Recettes du mois", COLOR_TEAL)
        self.month_expense_card = MetricCard(self.summary, "Depenses du mois", COLOR_CORAL)
        self.month_balance_card = MetricCard(self.summary, "Solde du mois", COLOR_GOLD)
        self.month_todo_card = MetricCard(self.summary, "Todos du mois", COLOR_NAVY)
        for index, widget in enumerate([self.month_revenue_card, self.month_expense_card, self.month_balance_card, self.month_todo_card]):
            widget.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 12, 0))

        body = ttk.Frame(self, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        self.calendar_card = ttk.Frame(body, style="Surface.TFrame", padding=18)
        self.calendar_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        ttk.Label(self.calendar_card, text="Vue mensuelle", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(self.calendar_card, text="Vert: positif, corail: depenses dominantes, sable: jour avec taches.", style="MutedSurface.TLabel").pack(anchor="w", pady=(2, 14))
        self.calendar_grid = tk.Frame(self.calendar_card, bg=COLOR_SURFACE)
        self.calendar_grid.pack(fill="both", expand=True)

        self.detail_card = ttk.Frame(body, style="Surface.TFrame", padding=18)
        self.detail_card.grid(row=0, column=1, sticky="nsew")
        ttk.Label(self.detail_card, text="Detail du jour", style="SectionTitle.TLabel").pack(anchor="w")
        self.selected_day_var = tk.StringVar()
        ttk.Label(self.detail_card, textvariable=self.selected_day_var, style="PageTitle.TLabel").pack(anchor="w", pady=(6, 6))
        self.day_state_var = tk.StringVar()
        ttk.Label(self.detail_card, textvariable=self.day_state_var, style="MutedSurface.TLabel").pack(anchor="w", pady=(0, 12))
        self.detail_text = tk.Text(self.detail_card, bg=COLOR_SURFACE, bd=0, fg=COLOR_TEXT, height=30, font=("Segoe UI", 10), wrap="word", highlightthickness=0)
        self.detail_text.pack(fill="both", expand=True)
        self.detail_text.configure(state="disabled")

    def previous_month(self) -> None:
        if self.current_month == 1:
            self.current_year -= 1
            self.current_month = 12
        else:
            self.current_month -= 1
        self.refresh()

    def next_month(self) -> None:
        if self.current_month == 12:
            self.current_year += 1
            self.current_month = 1
        else:
            self.current_month += 1
        self.refresh()

    def current_month_view(self) -> None:
        today = date.today()
        self.current_year = today.year
        self.current_month = today.month
        self.selected_date = today.isoformat()
        self.refresh()

    def refresh(self) -> None:
        self.snapshot = self.app.db.calendar_snapshot(self.current_year, self.current_month)
        self.month_var.set(month_label(self.current_year, self.current_month))
        self.month_revenue_card.set(money(self.snapshot["month_revenue"]), f"Factures: {money(self.snapshot['invoice_revenue'])}")
        self.month_expense_card.set(money(self.snapshot["month_expenses"]), "Depenses TTC du mois")
        self.month_balance_card.set(money(self.snapshot["month_balance"]), "Solde mensuel")
        self.month_todo_card.set(str(self.snapshot["pending_todos"]), "Taches du mois non terminees")
        self.render_calendar()
        if not self.selected_date.startswith(f"{self.current_year:04d}-{self.current_month:02d}"):
            self.selected_date = f"{self.current_year:04d}-{self.current_month:02d}-01"
        self.show_day(self.selected_date)

    def render_calendar(self) -> None:
        for child in self.calendar_grid.winfo_children():
            child.destroy()

        weekday_labels = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]
        for column, label in enumerate(weekday_labels):
            header = tk.Label(self.calendar_grid, text=label, bg=COLOR_SURFACE, fg=COLOR_TEXT_MUTED, font=("Segoe UI Semibold", 10))
            header.grid(row=0, column=column, sticky="nsew", padx=4, pady=(0, 6))

        calendar_builder = pycalendar.Calendar(firstweekday=0)
        weeks = calendar_builder.monthdayscalendar(self.current_year, self.current_month)
        for week_row, week in enumerate(weeks, start=1):
            self.calendar_grid.rowconfigure(week_row, weight=1)
            for column, day_number in enumerate(week):
                self.calendar_grid.columnconfigure(column, weight=1)
                if day_number == 0:
                    blank = tk.Frame(self.calendar_grid, bg=COLOR_SURFACE)
                    blank.grid(row=week_row, column=column, sticky="nsew", padx=4, pady=4)
                    continue
                day_key = f"{self.current_year:04d}-{self.current_month:02d}-{day_number:02d}"
                payload = self.snapshot["days"][day_key]
                lines = [str(day_number)]
                if payload["revenue"]:
                    lines.append(f"CA {short_money(payload['revenue'])}")
                if payload["expenses"]:
                    lines.append(f"Dep {short_money(payload['expenses'])}")
                if payload["todo_count"]:
                    lines.append(f"Todo {payload['todo_count']}")
                background = health_cell_color(payload["balance"], payload["revenue"], payload["expenses"], payload["todo_count"])
                relief = "solid" if day_key == self.selected_date else "flat"
                button = tk.Button(
                    self.calendar_grid,
                    text="\n".join(lines),
                    bg=background,
                    fg=COLOR_TEXT,
                    activebackground=background,
                    activeforeground=COLOR_TEXT,
                    relief=relief,
                    bd=2 if day_key == self.selected_date else 0,
                    justify="left",
                    anchor="nw",
                    wraplength=110,
                    font=("Segoe UI", 9),
                    command=lambda value=day_key: self.show_day(value),
                    padx=10,
                    pady=8,
                )
                button.grid(row=week_row, column=column, sticky="nsew", padx=4, pady=4)

    def show_day(self, day_key: str) -> None:
        self.selected_date = day_key
        if not self.snapshot:
            return
        payload = self.snapshot["days"].get(day_key)
        if not payload:
            return
        self.selected_day_var.set(format_day(day_key))
        self.day_state_var.set(
            f"{health_message(payload['balance'], payload['revenue'], payload['expenses'])} | CA {money(payload['revenue'])} | Depenses {money(payload['expenses'])} | Solde {money(payload['balance'])}"
        )
        lines = []
        if not payload["items"]:
            lines.append("Aucun evenement sur cette journee.")
        else:
            for item in payload["items"]:
                if item["kind"] in {"invoice", "sale"}:
                    lines.append(f"[Recette] {item['label']} - {money(item['amount'])} - {item['status']}")
                elif item["kind"] == "expense":
                    lines.append(f"[Depense] {item['label']} - {money(item['amount'])} - {item['status']}")
                elif item["kind"] == "todo":
                    lines.append(f"[Todo] {item['label']} - {item['status']}")
                elif item["kind"] == "invoice_due":
                    lines.append(f"[Echeance] {item['label']} - {money(item['amount'])} - {item['status']}")
                else:
                    lines.append(f"[Devis] {item['label']} - {money(item['amount'])} - {item['status']}")
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", "\n".join(lines))
        self.detail_text.configure(state="disabled")
        self.render_calendar()


class DocumentEditor(tk.Toplevel):
    def __init__(self, app: "VeloraApp", kind: str) -> None:
        super().__init__(app)
        self.app = app
        self.kind = kind
        self.items: list[dict] = []
        self.company_profile = self.app.db.get_company_profile()
        self.document_preferences = self.app.db.get_document_preferences()

        self.title("Nouvelle facture" if kind == "invoice" else "Nouveau devis")
        self.geometry("1080x760")
        self.minsize(980, 700)
        self.configure(bg=COLOR_BG)
        self.transient(app)
        self.grab_set()

        self.number_var = tk.StringVar(value=self.app.db.next_document_number(kind))
        self.issue_date_var = tk.StringVar(value=date.today().isoformat())
        default_status = self.document_preferences["default_invoice_status"] if kind == "invoice" else self.document_preferences["default_quote_status"]
        self.status_var = tk.StringVar(value=default_status)
        self.tax_rate_var = tk.StringVar(value=str(self.document_preferences["default_tax_rate"]))
        self.client_name_var = tk.StringVar(value=self.document_preferences["default_client_name"])
        self.client_email_var = tk.StringVar(value=self.document_preferences["default_client_email"])
        self.client_address_var = tk.StringVar(value=self.document_preferences["default_client_address"])
        self.description_var = tk.StringVar()
        self.quantity_var = tk.StringVar(value="1")
        self.unit_price_var = tk.StringVar(value="0")
        self.total_var = tk.StringVar(value=money(0))
        self.subtotal_var = tk.StringVar(value=money(0))
        self.tax_amount_var = tk.StringVar(value=money(0))
        self.tax_rate_var.trace_add("write", lambda *_args: self.update_totals())
        if kind == "invoice":
            self.period_var = tk.StringVar(value=(date.today() + timedelta(days=int(self.document_preferences["invoice_due_days"]))).isoformat())
        else:
            self.validity_days_var = tk.StringVar(value=str(self.document_preferences["quote_validity_days"]))
            self.period_var = tk.StringVar(value=(date.today() + timedelta(days=int(self.document_preferences["quote_validity_days"]))).isoformat())
            self.validity_days_var.trace_add("write", lambda *_args: self.update_valid_until())
            self.issue_date_var.trace_add("write", lambda *_args: self.update_valid_until())

        shell = ttk.Frame(self, style="App.TFrame", padding=20)
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=3)
        shell.columnconfigure(1, weight=2)
        shell.rowconfigure(1, weight=1)

        left = ttk.Frame(shell, style="Surface.TFrame", padding=18)
        left.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        right = ttk.Frame(shell, style="Surface.TFrame", padding=18)
        right.grid(row=0, column=1, sticky="nsew")
        bottom = ttk.Frame(shell, style="Surface.TFrame", padding=18)
        bottom.grid(row=1, column=1, sticky="nsew", pady=(12, 0))

        self._build_left(left)
        self._build_right(right)
        self._build_bottom(bottom)
        if self.kind == "invoice":
            self.notes_text.insert("1.0", self.document_preferences.get("default_invoice_notes", ""))
        else:
            self.notes_text.insert("1.0", self.document_preferences.get("default_quote_notes", ""))
        self.update_totals()

    def _labeled_entry(self, master, label: str, variable: tk.StringVar, row: int, column: int = 0, readonly: bool = False) -> ttk.Entry:
        ttk.Label(master, text=label, style="FieldLabel.TLabel").grid(row=row, column=column, sticky="w", pady=(8, 6))
        entry = ttk.Entry(master, textvariable=variable)
        if readonly:
            entry.state(["readonly"])
        entry.grid(row=row + 1, column=column, sticky="ew", pady=(0, 6), padx=(0, 10))
        return entry

    def _build_left(self, master) -> None:
        ttk.Label(master, text="Informations du document", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        self._labeled_entry(master, "Numero", self.number_var, 1, 0)
        self._labeled_entry(master, "Date", self.issue_date_var, 1, 1)

        if self.kind == "invoice":
            self._labeled_entry(master, "Echeance", self.period_var, 3, 0)
            ttk.Label(master, text="Statut", style="FieldLabel.TLabel").grid(row=3, column=1, sticky="w", pady=(8, 6))
            ttk.Combobox(master, textvariable=self.status_var, values=["Brouillon", "Envoyee", "Payee", "En retard"], state="readonly").grid(row=4, column=1, sticky="ew", pady=(0, 6))
        else:
            self._labeled_entry(master, "Validite en jours", self.validity_days_var, 3, 0)
            self._labeled_entry(master, "Valide jusqu'au", self.period_var, 3, 1, readonly=True)
            ttk.Label(master, text="Statut", style="FieldLabel.TLabel").grid(row=5, column=0, sticky="w", pady=(8, 6))
            ttk.Combobox(master, textvariable=self.status_var, values=["Brouillon", "Envoye", "Accepte", "Refuse", "Expire"], state="readonly").grid(row=6, column=0, sticky="ew", pady=(0, 6))

        start_row = 5 if self.kind == "invoice" else 7
        self._labeled_entry(master, "Client", self.client_name_var, start_row, 0)
        self._labeled_entry(master, "Email client", self.client_email_var, start_row, 1)
        self._labeled_entry(master, "Adresse client", self.client_address_var, start_row + 2, 0)
        self._labeled_entry(master, "TVA %", self.tax_rate_var, start_row + 2, 1)

        ttk.Label(master, text="Notes", style="FieldLabel.TLabel").grid(row=start_row + 4, column=0, sticky="w", pady=(8, 6))
        self.notes_text = tk.Text(master, height=4, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.notes_text.grid(row=start_row + 5, column=0, columnspan=2, sticky="nsew")

        ttk.Label(master, text="Lignes", style="SectionTitle.TLabel").grid(row=start_row + 6, column=0, columnspan=2, sticky="w", pady=(18, 8))
        self._labeled_entry(master, "Description", self.description_var, start_row + 7, 0)
        self._labeled_entry(master, "Quantite", self.quantity_var, start_row + 7, 1)
        self._labeled_entry(master, "Prix unitaire HT", self.unit_price_var, start_row + 9, 0)
        ttk.Button(master, text="Ajouter la ligne", style="Accent.TButton", command=self.add_item).grid(row=start_row + 10, column=1, sticky="ew", pady=(24, 0))

        columns = ("description", "quantity", "price", "total")
        self.items_tree = ttk.Treeview(master, columns=columns, show="headings", height=8)
        for column, heading, width in [("description", "Description", 260), ("quantity", "Qt", 60), ("price", "Prix HT", 100), ("total", "Total HT", 100)]:
            self.items_tree.heading(column, text=heading)
            self.items_tree.column(column, width=width, anchor="w")
        self.items_tree.grid(row=start_row + 11, column=0, columnspan=2, sticky="nsew", pady=(12, 8))
        ttk.Button(master, text="Retirer la ligne", style="Ghost.TButton", command=self.remove_item).grid(row=start_row + 12, column=0, sticky="w")
        master.columnconfigure((0, 1), weight=1)

    def _build_right(self, master) -> None:
        ttk.Label(master, text="Entreprise emettrice", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(master, text=self.company_profile.get("company_name", ""), style="PageTitle.TLabel").pack(anchor="w", pady=(6, 0))
        info = [
            self.company_profile.get("legal_name", ""),
            self.company_profile.get("address", ""),
            f"SIRET: {self.company_profile.get('siret', '')}",
            f"TVA: {self.company_profile.get('vat_number', '')}",
            f"{self.company_profile.get('email', '')} - {self.company_profile.get('phone', '')}",
        ]
        ttk.Label(master, text="\n".join([line for line in info if line]), style="MutedSurface.TLabel", wraplength=300).pack(anchor="w", pady=(8, 0))
        extra_text = "Chaque facture creee genere automatiquement du chiffre d'affaires dans le tableau de bord et le calendrier." if self.kind == "invoice" else "Le devis genere un document HTML local, avec une validite modifiable."
        ttk.Label(master, text=extra_text, style="MutedSurface.TLabel", wraplength=300).pack(anchor="w", pady=(18, 0))

    def _build_bottom(self, master) -> None:
        ttk.Label(master, text="Synthese", style="SectionTitle.TLabel").pack(anchor="w")
        for label, variable in [("Sous-total HT", self.subtotal_var), ("TVA", self.tax_amount_var), ("Total TTC", self.total_var)]:
            line = ttk.Frame(master, style="SurfaceSoft.TFrame", padding=12)
            line.pack(fill="x", pady=6)
            ttk.Label(line, text=label, style="TableStrongSurface.TLabel").pack(side="left")
            ttk.Label(line, textvariable=variable, style="MetricMiniSoft.TLabel").pack(side="right")

        actions = ttk.Frame(master, style="Surface.TFrame")
        actions.pack(fill="x", pady=(16, 0))
        ttk.Button(actions, text="Fermer", style="Ghost.TButton", command=self.destroy).pack(side="left")
        ttk.Button(actions, text="Generer et enregistrer", style="Accent.TButton", command=self.save).pack(side="right")

    def update_valid_until(self) -> None:
        try:
            start = parse_iso_date(self.issue_date_var.get())
            days = int(self.validity_days_var.get())
            self.period_var.set((start + timedelta(days=days)).isoformat())
        except Exception:
            return

    def add_item(self) -> None:
        description = self.description_var.get().strip()
        if not description:
            messagebox.showwarning("Ligne", "Ajoutez une description de ligne.")
            return
        try:
            quantity = parse_amount(self.quantity_var.get())
            unit_price = parse_amount(self.unit_price_var.get())
            if quantity <= 0:
                raise ValueError("La quantite doit etre superieure a 0.")
            if unit_price < 0:
                raise ValueError("Le prix unitaire doit etre positif.")
        except Exception as exc:
            messagebox.showerror("Ligne", f"Impossible d'ajouter la ligne.\n\n{exc}")
            return
        self.items.append({"description": description, "quantity": quantity, "unit_price": unit_price})
        self.description_var.set("")
        self.quantity_var.set("1")
        self.unit_price_var.set("0")
        self.refresh_items()
        self.update_totals()

    def refresh_items(self) -> None:
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)
        for index, item in enumerate(self.items, start=1):
            total = item["quantity"] * item["unit_price"]
            self.items_tree.insert("", "end", iid=str(index - 1), values=(item["description"], item["quantity"], money(item["unit_price"]), money(total)))

    def remove_item(self) -> None:
        selection = self.items_tree.selection()
        if not selection:
            return
        self.items.pop(int(selection[0]))
        self.refresh_items()
        self.update_totals()

    def update_totals(self) -> None:
        subtotal = sum(item["quantity"] * item["unit_price"] for item in self.items)
        try:
            tax_rate = parse_amount(self.tax_rate_var.get())
            validate_tax_rate(tax_rate)
        except Exception:
            tax_rate = 0.0
        tax_amount = subtotal * (tax_rate / 100)
        total = subtotal + tax_amount
        self.subtotal_var.set(money(subtotal))
        self.tax_amount_var.set(money(tax_amount))
        self.total_var.set(money(total))

    def save(self) -> None:
        if not self.items:
            messagebox.showwarning("Document", "Ajoutez au moins une ligne avant de generer le document.")
            return
        created_id: int | None = None
        try:
            issue_date = parse_iso_date(self.issue_date_var.get())
            deadline = self.period_var.get()
            deadline_date = parse_iso_date(deadline)
            number = self.number_var.get().strip()
            if not number:
                raise ValueError("Le numero du document est obligatoire.")
            if deadline_date < issue_date:
                if self.kind == "invoice":
                    raise ValueError("L'echeance ne peut pas etre anterieure a la date de facture.")
                raise ValueError("La validite du devis ne peut pas etre anterieure a la date du devis.")
            subtotal = sum(item["quantity"] * item["unit_price"] for item in self.items)
            tax_rate = parse_amount(self.tax_rate_var.get())
            validate_tax_rate(tax_rate)
            tax_amount = subtotal * (tax_rate / 100)
            total = subtotal + tax_amount
            html_path = document_target_path(self.kind, number)
            payload = {
                "number": number,
                "issue_date": self.issue_date_var.get().strip(),
                "status": self.status_var.get().strip(),
                "client_name": self.client_name_var.get().strip() or "Client",
                "client_email": self.client_email_var.get().strip(),
                "client_address": self.client_address_var.get().strip(),
                "items": self.items,
                "notes": self.notes_text.get("1.0", "end").strip(),
                "tax_rate": tax_rate,
                "subtotal": subtotal,
                "tax_amount": tax_amount,
                "total": total,
                "company_profile": self.company_profile,
                "generated_at": datetime.now().isoformat(timespec="seconds"),
                "html_path": html_path,
            }
            if self.kind == "invoice":
                payload["due_date"] = deadline
                created_id = self.app.db.create_invoice(payload)
                save_document_html("invoice", payload, html_path)
            else:
                payload["valid_until"] = deadline
                created_id = self.app.db.create_quote(payload)
                save_document_html("quote", payload, html_path)
        except Exception as exc:
            if created_id is not None:
                try:
                    if self.kind == "invoice":
                        self.app.db.delete_invoice(created_id)
                    else:
                        self.app.db.delete_quote(created_id)
                except Exception:
                    pass
            messagebox.showerror("Document", f"Impossible de generer le document.\n\n{exc}")
            return

        self.app.refresh_all_pages()
        webbrowser.open(Path(html_path).as_uri())
        messagebox.showinfo("Document", "Le document a ete genere localement et ouvert dans votre navigateur.")
        self.destroy()


class VeloraApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.db = Database()
        self.title(APP_NAME)
        self.geometry("1500x940")
        self.minsize(1320, 840)
        self.configure(bg=COLOR_BG)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._load_logo()
        self._configure_styles()
        self._build_shell()
        self.refresh_all_pages()

    def _load_logo(self) -> None:
        self.logo_image = None
        png_path = asset_path("velora_logo.png")
        if png_path.exists():
            self.logo_image = tk.PhotoImage(file=str(png_path))
            self.iconphoto(True, self.logo_image)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("App.TFrame", background=COLOR_BG)
        style.configure("Surface.TFrame", background=COLOR_SURFACE)
        style.configure("SurfaceSoft.TFrame", background=COLOR_SURFACE_SOFT)
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT_MUTED, font=("Segoe UI", 10, "bold"))
        style.configure("MetricValue.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=("Bahnschrift SemiBold", 23))
        style.configure("MetricMiniSoft.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, font=("Bahnschrift SemiBold", 15))
        style.configure("PageTitle.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Bahnschrift SemiBold", 24))
        style.configure("SectionTitle.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=("Bahnschrift SemiBold", 13))
        style.configure("TableStrong.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=("Segoe UI Semibold", 10))
        style.configure("TableStrongSurface.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, font=("Segoe UI Semibold", 10))
        style.configure("Muted.TLabel", background=COLOR_BG, foreground=COLOR_TEXT_MUTED, font=("Segoe UI", 10))
        style.configure("MutedSurface.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT_MUTED, font=("Segoe UI", 10))
        style.configure("MutedSoft.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT_MUTED, font=("Segoe UI", 10))
        style.configure("FieldLabel.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT_MUTED, font=("Segoe UI Semibold", 9))
        style.configure("BadgeSurface.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEAL, font=("Segoe UI", 9, "bold"))

        for name, color in [("Accent.TButton", COLOR_TEAL), ("AccentAlt.TButton", COLOR_CORAL)]:
            style.configure(name, background=color, foreground="#ffffff", padding=(16, 10), borderwidth=0)
            style.map(name, background=[("active", COLOR_NAVY_DEEP)])
        style.configure("Ghost.TButton", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, padding=(14, 10), borderwidth=0)
        style.map("Ghost.TButton", background=[("active", "#dfe7f2")])
        style.configure("Danger.TButton", background=COLOR_DANGER, foreground="#ffffff", padding=(14, 10), borderwidth=0)
        style.map("Danger.TButton", background=[("active", "#a94141")])

        style.configure("Treeview", background=COLOR_SURFACE, foreground=COLOR_TEXT, fieldbackground=COLOR_SURFACE, bordercolor=COLOR_BORDER, rowheight=34, font=("Segoe UI", 10))
        style.map("Treeview", background=[("selected", "#dbece6")], foreground=[("selected", COLOR_TEXT)])
        style.configure("Treeview.Heading", background="#eef2f7", foreground=COLOR_TEXT_MUTED, font=("Segoe UI Semibold", 9), relief="flat")
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER, padding=8)
        style.configure("TCombobox", fieldbackground="#ffffff", bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER, padding=8)

    def _build_shell(self) -> None:
        self.sidebar = tk.Frame(self, bg=COLOR_NAVY, width=250)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = tk.Frame(self.sidebar, bg=COLOR_NAVY)
        brand.pack(fill="x", padx=24, pady=(26, 26))
        if self.logo_image is not None:
            icon = self.logo_image.subsample(5, 5)
            self.sidebar_logo = icon
            tk.Label(brand, image=icon, bg=COLOR_NAVY).pack(anchor="w")
        tk.Label(brand, text=APP_NAME, bg=COLOR_NAVY, fg="#ffffff", font=("Bahnschrift SemiBold", 18)).pack(anchor="w", pady=(12, 2))
        tk.Label(brand, text=APP_TAGLINE, bg=COLOR_NAVY, fg="#c7d7e8", wraplength=190, justify="left", font=("Segoe UI", 9)).pack(anchor="w")

        self.nav_buttons: dict[str, tk.Button] = {}
        for key, label in [
            ("dashboard", "Tableau de bord"),
            ("calendar", "Calendrier"),
            ("invoices", "Factures"),
            ("quotes", "Devis"),
            ("sales", "Recettes"),
            ("expenses", "Depenses"),
            ("todos", "Todo liste"),
            ("company", "Entreprise"),
        ]:
            button = tk.Button(
                self.sidebar,
                text=label,
                bg=COLOR_NAVY,
                fg="#ecf3fb",
                activebackground=COLOR_TEAL,
                activeforeground="#ffffff",
                borderwidth=0,
                anchor="w",
                padx=24,
                pady=12,
                font=("Segoe UI Semibold", 10),
                command=lambda page=key: self.show_page(page),
            )
            button.pack(fill="x", padx=12, pady=4)
            self.nav_buttons[key] = button

        footer = tk.Frame(self.sidebar, bg=COLOR_NAVY)
        footer.pack(side="bottom", fill="x", padx=24, pady=24)
        tk.Label(footer, text="Stockage local", bg=COLOR_NAVY, fg="#ffffff", font=("Segoe UI Semibold", 10)).pack(anchor="w")
        tk.Label(footer, text=str(storage_root()), bg=COLOR_NAVY, fg="#c7d7e8", wraplength=190, justify="left", font=("Segoe UI", 8)).pack(anchor="w", pady=(6, 0))

        self.main_area = ttk.Frame(self, style="App.TFrame")
        self.main_area.pack(side="left", fill="both", expand=True)

        self.pages = {
            "dashboard": DashboardPage(self.main_area, self),
            "calendar": CalendarPage(self.main_area, self),
            "invoices": DocumentPage(self.main_area, self, "Factures", "Creation, suivi et ouverture locale de vos factures clients.", "invoice"),
            "quotes": DocumentPage(self.main_area, self, "Devis", "Generateur de devis avec validite modifiable et suivi d'acceptation.", "quote"),
            "sales": SalesPage(self.main_area, self),
            "expenses": ExpensesPage(self.main_area, self),
            "todos": TodoPage(self.main_area, self),
            "company": CompanyPage(self.main_area, self),
        }
        for page in self.pages.values():
            page.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.show_page("dashboard")

    def show_page(self, name: str) -> None:
        for key, page in self.pages.items():
            if key == name:
                page.lift()
                self.nav_buttons[key].configure(bg=COLOR_TEAL)
            else:
                self.nav_buttons[key].configure(bg=COLOR_NAVY)

    def refresh_all_pages(self) -> None:
        for page in self.pages.values():
            if hasattr(page, "refresh"):
                page.refresh()

    def on_close(self) -> None:
        self.db.close()
        self.destroy()
