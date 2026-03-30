from __future__ import annotations

import calendar as pycalendar
import shutil
import tkinter as tk
import webbrowser
import tkinter.font as tkfont
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import (
    APP_NAME,
    APP_TAGLINE,
    APP_THEMES,
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
    COLOR_WARNING,
    asset_path,
    get_theme_palette,
    storage_root,
)
from .database import Database
from .documents import document_target_path, money, safe_document_stem, save_document_html
from .exports import collect_month_keys, export_month_bundle


COLOR_HEALTH_GOOD = "#dbeee4"
COLOR_HEALTH_BAD = "#f7e0d9"
COLOR_HEALTH_TODO = "#f5e8cf"
COLOR_INPUT_BG = "#ffffff"
COLOR_TEXT_SELECT_BG = "#dce8e1"
COLOR_CHART_GRID = "#e8dece"
COLOR_BAR_TRACK = "#f1eadf"
COLOR_GHOST_HOVER = "#e7e0d2"
COLOR_GHOST_PRESSED = "#dfd7c6"
COLOR_DANGER_HOVER = "#a94141"
COLOR_DANGER_PRESSED = "#963838"
COLOR_TREE_SELECTED = "#dce8e1"
COLOR_HEADING_BG = "#f1eadf"
COLOR_HEADING_ACTIVE = "#ebe2d4"
COLOR_SIDEBAR_OUTLINE = "#294563"
COLOR_SIDEBAR_TEXT = "#edf4fb"
COLOR_SIDEBAR_MUTED = "#d8e3ef"
COLOR_SIDEBAR_CAPTION = "#b7c7d7"
COLOR_SIDEBAR_ACCENT = "#8fd1c7"

FONT_BODY = "Aptos"
FONT_BODY_SEMIBOLD = "Aptos Semibold"
FONT_HEADLINE = "Aptos Display"
FONT_NUMERIC = "Bahnschrift SemiBold"

VALUE_ANIMATION_STEPS = 8
VALUE_ANIMATION_DELAY_MS = 20
CHART_ANIMATION_STEPS = 8
CHART_ANIMATION_DELAY_MS = 22
CARD_PULSE_HEIGHTS = [4, 5, 6, 5, 4]


def apply_runtime_palette(theme_name: str) -> str:
    palette = get_theme_palette(theme_name)
    global COLOR_NAVY, COLOR_NAVY_DEEP, COLOR_TEAL, COLOR_CORAL, COLOR_GOLD
    global COLOR_BG, COLOR_SURFACE, COLOR_SURFACE_SOFT, COLOR_TEXT, COLOR_TEXT_MUTED
    global COLOR_BORDER, COLOR_SUCCESS, COLOR_WARNING, COLOR_DANGER
    global COLOR_HEALTH_GOOD, COLOR_HEALTH_BAD, COLOR_HEALTH_TODO, COLOR_INPUT_BG
    global COLOR_TEXT_SELECT_BG, COLOR_CHART_GRID, COLOR_BAR_TRACK, COLOR_GHOST_HOVER
    global COLOR_GHOST_PRESSED, COLOR_DANGER_HOVER, COLOR_DANGER_PRESSED
    global COLOR_TREE_SELECTED, COLOR_HEADING_BG, COLOR_HEADING_ACTIVE
    global COLOR_SIDEBAR_OUTLINE, COLOR_SIDEBAR_TEXT, COLOR_SIDEBAR_MUTED
    global COLOR_SIDEBAR_CAPTION, COLOR_SIDEBAR_ACCENT

    COLOR_NAVY = palette["navy"]
    COLOR_NAVY_DEEP = palette["navy_deep"]
    COLOR_TEAL = palette["teal"]
    COLOR_CORAL = palette["coral"]
    COLOR_GOLD = palette["gold"]
    COLOR_BG = palette["bg"]
    COLOR_SURFACE = palette["surface"]
    COLOR_SURFACE_SOFT = palette["surface_soft"]
    COLOR_TEXT = palette["text"]
    COLOR_TEXT_MUTED = palette["text_muted"]
    COLOR_BORDER = palette["border"]
    COLOR_SUCCESS = palette["success"]
    COLOR_WARNING = palette["warning"]
    COLOR_DANGER = palette["danger"]
    COLOR_HEALTH_GOOD = palette["health_good"]
    COLOR_HEALTH_BAD = palette["health_bad"]
    COLOR_HEALTH_TODO = palette["health_todo"]
    COLOR_INPUT_BG = palette["input_bg"]
    COLOR_TEXT_SELECT_BG = palette["text_select_bg"]
    COLOR_CHART_GRID = palette["chart_grid"]
    COLOR_BAR_TRACK = palette["bar_track"]
    COLOR_GHOST_HOVER = palette["ghost_hover"]
    COLOR_GHOST_PRESSED = palette["ghost_pressed"]
    COLOR_DANGER_HOVER = palette["danger_hover"]
    COLOR_DANGER_PRESSED = palette["danger_pressed"]
    COLOR_TREE_SELECTED = palette["tree_selected"]
    COLOR_HEADING_BG = palette["heading_bg"]
    COLOR_HEADING_ACTIVE = palette["heading_active"]
    COLOR_SIDEBAR_OUTLINE = palette["sidebar_outline"]
    COLOR_SIDEBAR_TEXT = palette["sidebar_text"]
    COLOR_SIDEBAR_MUTED = palette["sidebar_muted"]
    COLOR_SIDEBAR_CAPTION = palette["sidebar_caption"]
    COLOR_SIDEBAR_ACCENT = palette["sidebar_accent"]
    return str(theme_name or "Clair").strip().capitalize() if str(theme_name or "Clair").strip().capitalize() in APP_THEMES else "Clair"


apply_runtime_palette("Clair")


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
        return COLOR_HEALTH_GOOD
    if balance < 0 or expenses > revenue:
        return COLOR_HEALTH_BAD
    if todo_count:
        return COLOR_HEALTH_TODO
    return COLOR_SURFACE_SOFT


def source_label(source_type: str) -> str:
    return "Facture" if source_type == "invoice" else "Manuel"


def style_text_widget(widget: tk.Text, background: str | None = None) -> None:
    background = background or COLOR_INPUT_BG
    widget.configure(
        bg=background,
        fg=COLOR_TEXT,
        insertbackground=COLOR_TEXT,
        selectbackground=COLOR_TEXT_SELECT_BG,
        selectforeground=COLOR_TEXT,
        highlightbackground=COLOR_BORDER,
        highlightcolor=COLOR_TEAL,
        padx=12,
        pady=10,
        font=(FONT_BODY, 10),
    )


def copy_expense_attachment(source_path: str) -> Path:
    source = Path(source_path).expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise ValueError("Le fichier de depense selectionne est introuvable.")
    target_root = storage_root() / "documents" / "depenses"
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe_name = safe_document_stem(source.stem)
    target = target_root / f"{timestamp}-{safe_name}{source.suffix.lower()}"
    shutil.copy2(source, target)
    return target


def pulse_panel(panel: ttk.Frame, base_padding: int = 18) -> None:
    try:
        panel.configure(padding=base_padding)
    except tk.TclError:
        return


def ease_out_cubic(progress: float) -> float:
    progress = min(max(progress, 0.0), 1.0)
    return 1 - ((1 - progress) ** 3)


def widget_background(widget, fallback: str) -> str:
    try:
        bg = str(widget.cget("bg")).strip()
        if bg:
            return bg
    except Exception:
        pass

    try:
        style_name = str(widget.cget("style")).strip()
    except Exception:
        style_name = ""

    style = ttk.Style(widget)
    for candidate in [style_name, widget.winfo_class(), "TFrame", "TLabel"]:
        if not candidate:
            continue
        bg = str(style.lookup(candidate, "background")).strip()
        if bg:
            return bg
    return fallback


def rounded_parts(canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int, radius: int, fill: str) -> None:
    radius = max(0, min(radius, int((x2 - x1) / 2), int((y2 - y1) / 2)))
    if radius <= 1:
        canvas.create_rectangle(x1, y1, x2, y2, fill=fill, outline=fill)
        return
    canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2, fill=fill, outline=fill)
    canvas.create_rectangle(x1, y1 + radius, x2, y2 - radius, fill=fill, outline=fill)
    canvas.create_arc(x1, y1, x1 + (radius * 2), y1 + (radius * 2), start=90, extent=90, style="pieslice", fill=fill, outline=fill)
    canvas.create_arc(x2 - (radius * 2), y1, x2, y1 + (radius * 2), start=0, extent=90, style="pieslice", fill=fill, outline=fill)
    canvas.create_arc(x1, y2 - (radius * 2), x1 + (radius * 2), y2, start=180, extent=90, style="pieslice", fill=fill, outline=fill)
    canvas.create_arc(x2 - (radius * 2), y2 - (radius * 2), x2, y2, start=270, extent=90, style="pieslice", fill=fill, outline=fill)


def capsule_parts(canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int, fill: str, border_fill: str, border_width: int = 1) -> None:
    height = max(int(y2 - y1), 2)
    center_y = (y1 + y2) / 2
    line_start = x1 + (height / 2)
    line_end = x2 - (height / 2)
    if line_end <= line_start:
        rounded_parts(canvas, x1, y1, x2, y2, max(height // 2, 1), border_fill)
        if border_width > 0 and (x2 - x1) > (border_width * 2) and (y2 - y1) > (border_width * 2):
            rounded_parts(canvas, x1 + border_width, y1 + border_width, x2 - border_width, y2 - border_width, max((height // 2) - border_width, 1), fill)
        return

    canvas.create_line(line_start, center_y, line_end, center_y, fill=border_fill, width=height, capstyle=tk.ROUND)
    inner_width = max(height - (border_width * 2), 2)
    canvas.create_line(line_start, center_y, line_end, center_y, fill=fill, width=inner_width, capstyle=tk.ROUND)


class RoundButton(tk.Canvas):
    def __init__(
        self,
        master,
        text: str,
        command,
        style_name: str = "Ghost.TButton",
        anchor: str = "center",
        radius: int = 18,
        min_height: int = 46,
        min_width: int = 0,
        canvas_bg: str | None = None,
        fill_color: str | None = None,
        hover_fill: str | None = None,
        pressed_fill: str | None = None,
        text_color: str | None = None,
        border_color: str | None = None,
        font_spec: tuple[str, int] | tuple[str, int, str] | None = None,
    ) -> None:
        self._canvas_bg = canvas_bg or widget_background(master, COLOR_BG)
        super().__init__(master, bg=self._canvas_bg, highlightthickness=0, bd=0, relief="flat", cursor="hand2", height=min_height)
        self.text = text
        self.command = command
        self.style_name = style_name
        self.anchor_mode = anchor
        self.radius = radius
        self.min_height = min_height
        self.min_width = min_width
        self.fill_color = fill_color
        self.hover_fill = hover_fill
        self.pressed_fill = pressed_fill
        self.text_color = text_color
        self.border_color = border_color
        self._hovered = False
        self._pressed = False
        self._disabled = False
        self._font = tkfont.Font(font=font_spec) if font_spec else tkfont.Font(family=FONT_BODY_SEMIBOLD, size=11)

        target_width = max(min_width, self._font.measure(text) + 44)
        self.configure(width=target_width)

        self.bind("<Configure>", lambda _event: self.redraw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.redraw()

    def redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), self.min_width or 80)
        height = max(self.winfo_height(), self.min_height)
        border_fill, inner_fill, text_fill, gloss_fill = self._palette()
        is_capsule = height <= 60

        if is_capsule:
            capsule_parts(self, 0, 1, width, height - 1, inner_fill, border_fill, 1)
            if gloss_fill:
                highlight_y = max(10, int(height * 0.34))
                highlight_start = max(16, int(height * 0.42))
                highlight_end = width - highlight_start
                if highlight_end > highlight_start:
                    self.create_line(
                        highlight_start,
                        highlight_y,
                        highlight_end,
                        highlight_y,
                        fill=gloss_fill,
                        width=max(2, int(height * 0.12)),
                        capstyle=tk.ROUND,
                    )
        else:
            rounded_parts(self, 0, 0, width, height, self.radius, border_fill)
            rounded_parts(self, 1, 1, width - 1, height - 1, max(self.radius - 1, 1), inner_fill)
            if gloss_fill:
                self.create_line(18, 14, width - 18, 14, fill=gloss_fill, width=3, capstyle=tk.ROUND)

        if self.anchor_mode == "w":
            text_x = 20
            anchor = "w"
        elif self.anchor_mode == "e":
            text_x = width - 20
            anchor = "e"
        else:
            text_x = width / 2
            anchor = "center"

        offset_y = 1 if self._pressed else 0
        self.create_text(
            text_x,
            (height / 2) + offset_y,
            text=self.text,
            fill=text_fill,
            font=self._font,
            anchor=anchor,
            width=max(width - 30, 20) if anchor != "center" else 0,
            justify="left" if anchor != "center" else "center",
        )

    def _palette(self) -> tuple[str, str, str, str]:
        if self.fill_color:
            border_fill = self.border_color or self._shade(self.fill_color, 0.88)
            base_fill = self.fill_color
            text_fill = self.text_color or COLOR_TEXT
            hover_fill = self.hover_fill or self._shade(base_fill, 1.04)
            pressed_fill = self.pressed_fill or self._shade(base_fill, 0.9)
            gloss_fill = self._shade(base_fill, 1.11)
            if self._disabled:
                return border_fill, self._shade(base_fill, 1.02), COLOR_TEXT_MUTED, ""
            if self._pressed:
                return border_fill, pressed_fill, text_fill, ""
            if self._hovered:
                return border_fill, hover_fill, text_fill, gloss_fill
            return border_fill, base_fill, text_fill, gloss_fill

        palettes = {
            "Accent.TButton": (self._shade(COLOR_TEAL, 0.82), COLOR_TEAL, "#ffffff", self._shade(COLOR_TEAL, 1.04), self._shade(COLOR_TEAL, 0.92), self._shade(COLOR_TEAL, 1.16)),
            "AccentAlt.TButton": (self._shade(COLOR_CORAL, 0.82), COLOR_CORAL, "#ffffff", self._shade(COLOR_CORAL, 1.04), self._shade(COLOR_CORAL, 0.92), self._shade(COLOR_CORAL, 1.16)),
            "Danger.TButton": (self._shade(COLOR_DANGER, 0.82), COLOR_DANGER, "#ffffff", self._shade(COLOR_DANGER, 1.04), self._shade(COLOR_DANGER, 0.92), self._shade(COLOR_DANGER, 1.14)),
            "Ghost.TButton": ("#d7e0ec", "#ffffff", COLOR_TEXT, "#f8fbff", "#eef4fb", "#ffffff"),
        }
        border_fill, base_fill, text_fill, hover_fill, pressed_fill, gloss_fill = palettes.get(self.style_name, palettes["Ghost.TButton"])

        if self._disabled:
            return border_fill, self._shade(base_fill, 1.02), COLOR_TEXT_MUTED, ""
        if self._pressed:
            return border_fill, pressed_fill, "#ffffff" if self.style_name != "Ghost.TButton" else COLOR_TEXT, ""
        if self._hovered:
            return border_fill, hover_fill, "#ffffff" if self.style_name != "Ghost.TButton" else COLOR_TEXT, gloss_fill
        return border_fill, base_fill, text_fill, gloss_fill

    def _on_enter(self, _event) -> None:
        if not self._disabled:
            self._hovered = True
            self.redraw()

    def _on_leave(self, _event) -> None:
        self._hovered = False
        self._pressed = False
        self.redraw()

    def _on_press(self, _event) -> None:
        if not self._disabled:
            self._pressed = True
            self.redraw()

    def _on_release(self, event) -> None:
        if self._disabled:
            return
        was_pressed = self._pressed
        self._pressed = False
        self.redraw()
        if not was_pressed:
            return
        if 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self.command()

    def set_disabled(self, disabled: bool) -> None:
        self._disabled = disabled
        self.configure(cursor="arrow" if disabled else "hand2")
        self.redraw()

    def _shade(self, color: str, factor: float) -> str:
        color = color.lstrip("#")
        red = max(0, min(255, int(int(color[0:2], 16) * factor)))
        green = max(0, min(255, int(int(color[2:4], 16) * factor)))
        blue = max(0, min(255, int(int(color[4:6], 16) * factor)))
        return f"#{red:02x}{green:02x}{blue:02x}"


class SidebarScrollArea(tk.Frame):
    def __init__(self, master, width: int) -> None:
        super().__init__(master, bg=COLOR_NAVY_DEEP)
        self.canvas = tk.Canvas(self, bg=COLOR_NAVY_DEEP, highlightthickness=0, bd=0, width=width)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.content = tk.Frame(self.canvas, bg=COLOR_NAVY_DEEP)
        self._content_window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.content.bind("<Configure>", self._sync_scrollregion)
        self.canvas.bind("<Configure>", self._sync_width)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.content.bind("<MouseWheel>", self._on_mousewheel)

    def _sync_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_width(self, event) -> None:
        self.canvas.itemconfigure(self._content_window, width=event.width)

    def _on_mousewheel(self, event) -> str:
        delta = int(getattr(event, "delta", 0))
        units = -int(delta / 120) if delta else 0
        if units == 0:
            units = -1 if delta > 0 else 1
        self.scroll_units(units * 3)
        return "break"

    def scroll_units(self, units: int) -> None:
        self.canvas.yview_scroll(units, "units")


class SidebarNavButton(tk.Canvas):
    def __init__(self, master, text: str, command, kind: str = "item") -> None:
        super().__init__(master, bg=COLOR_NAVY_DEEP, highlightthickness=0, bd=0, relief="flat", cursor="hand2", height=48)
        self.base_text = text
        self.command = command
        self.kind = kind
        self.active = False
        self.expanded = False
        self.hovered = False
        self.pressed = False
        self._font = tkfont.Font(family=FONT_HEADLINE if kind == "group" else FONT_BODY_SEMIBOLD, size=12 if kind == "group" else 10, weight="bold" if kind == "group" else "normal")

        self.bind("<Configure>", lambda _event: self.redraw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.redraw()

    def set_state(self, active: bool = False, expanded: bool = False) -> None:
        self.active = active
        self.expanded = expanded
        self.redraw()

    def redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 230)
        height = max(self.winfo_height(), 48)
        outer_fill, inner_fill, text_fill = self._palette()
        rounded_parts(self, 0, 0, width, height, 20 if self.kind == "group" else 18, outer_fill)
        rounded_parts(self, 1, 1, width - 1, height - 1, 19 if self.kind == "group" else 17, inner_fill)
        if self.hovered or self.active or self.expanded:
            self.create_line(18, 13, width - 18, 13, fill=self._highlight_color(), width=3, capstyle=tk.ROUND)

        label = self.base_text
        if self.kind == "group":
            arrow = "v" if self.expanded else ">"
            label = f"{self.base_text}   {arrow}"
        self.create_text(18, height / 2 + (1 if self.pressed else 0), text=label, anchor="w", fill=text_fill, font=self._font)

    def _highlight_color(self) -> str:
        if self.kind == "group":
            return "#4a6a90"
        return self._shade(COLOR_TEAL if self.active else COLOR_NAVY, 1.28)

    def _palette(self) -> tuple[str, str, str]:
        if self.kind == "group":
            border = COLOR_SIDEBAR_OUTLINE
            if self.hovered or self.expanded or self.active:
                return border, COLOR_NAVY, "#ffffff"
            return border, COLOR_NAVY_DEEP, COLOR_SIDEBAR_CAPTION
        border = COLOR_SIDEBAR_OUTLINE
        if self.active:
            return self._shade(COLOR_TEAL, 0.82), COLOR_TEAL, "#ffffff"
        if self.pressed:
            return border, COLOR_NAVY, "#ffffff"
        if self.hovered:
            return border, COLOR_NAVY, "#ffffff"
        return border, COLOR_NAVY_DEEP, COLOR_SIDEBAR_TEXT

    def _on_enter(self, _event) -> None:
        self.hovered = True
        self.redraw()

    def _on_leave(self, _event) -> None:
        self.hovered = False
        self.pressed = False
        self.redraw()

    def _on_press(self, _event) -> None:
        self.pressed = True
        self.redraw()

    def _on_release(self, event) -> None:
        was_pressed = self.pressed
        self.pressed = False
        self.redraw()
        if was_pressed and 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self.command()

    def _shade(self, color: str, factor: float) -> str:
        color = color.lstrip("#")
        red = max(0, min(255, int(int(color[0:2], 16) * factor)))
        green = max(0, min(255, int(int(color[2:4], 16) * factor)))
        blue = max(0, min(255, int(int(color[4:6], 16) * factor)))
        return f"#{red:02x}{green:02x}{blue:02x}"


class ScrollableFrame(ttk.Frame):
    def __init__(self, master, canvas_bg: str) -> None:
        super().__init__(master, style="App.TFrame")
        self.canvas = tk.Canvas(self, bg=canvas_bg, highlightthickness=0, bd=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.content = ttk.Frame(self.canvas, style="App.TFrame")
        self._content_window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.content.bind("<Configure>", self._sync_scrollregion)
        self.canvas.bind("<Configure>", self._sync_width)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.content.bind("<MouseWheel>", self._on_mousewheel)

    def _sync_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_width(self, event) -> None:
        self.canvas.itemconfigure(self._content_window, width=event.width)

    def _on_mousewheel(self, event) -> str:
        delta = -1 if event.delta > 0 else 1
        self.scroll_units(delta * 3)
        return "break"

    def scroll_units(self, units: int) -> None:
        self.canvas.yview_scroll(units, "units")

    def scroll_pages(self, pages: int) -> None:
        self.canvas.yview_scroll(pages, "pages")

    def scroll_home(self) -> None:
        self.canvas.yview_moveto(0)


class MetricCard(ttk.Frame):
    def __init__(self, master, title: str, accent: str) -> None:
        super().__init__(master, style="CardPanel.TFrame", padding=20)
        self.value_var = tk.StringVar(value="0")
        self.subtitle_var = tk.StringVar(value="")
        self.accent_color = accent
        self._value_after_id: str | None = None
        self._pulse_after_ids: list[str] = []

        self.accent_strip = tk.Frame(self, bg=accent, height=4)
        self.accent_strip.grid(row=0, column=0, sticky="ew")

        body = ttk.Frame(self, style="CardPanel.TFrame")
        body.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        body.columnconfigure(1, weight=1)

        icon_box = tk.Frame(body, bg=COLOR_SURFACE_SOFT, width=46, height=46)
        icon_box.grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 14))
        icon_box.grid_propagate(False)
        tk.Label(
            icon_box,
            text=self._icon_text(title),
            bg=COLOR_SURFACE_SOFT,
            fg=accent,
            font=(FONT_BODY_SEMIBOLD, 10),
        ).pack(expand=True)

        header = ttk.Frame(body, style="CardPanel.TFrame")
        header.grid(row=0, column=1, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=title, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Live", style="CardBadge.TLabel").grid(row=0, column=1, sticky="e")

        ttk.Label(body, textvariable=self.value_var, style="MetricValue.TLabel").grid(row=1, column=1, sticky="w", pady=(12, 2))
        ttk.Label(body, textvariable=self.subtitle_var, style="MutedSurface.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

    def set(self, value: str, subtitle: str) -> None:
        self.subtitle_var.set(subtitle)
        target_number = self._extract_numeric(value)
        current_number = self._extract_numeric(self.value_var.get())
        if target_number is not None and current_number is not None:
            self._animate_value(current_number, target_number, value)
        else:
            if self._value_after_id is not None:
                try:
                    self.after_cancel(self._value_after_id)
                except tk.TclError:
                    pass
                self._value_after_id = None
            self.value_var.set(value)
        self.pulse()

    def pulse(self) -> None:
        for after_id in self._pulse_after_ids:
            try:
                self.after_cancel(after_id)
            except tk.TclError:
                pass
        self._pulse_after_ids.clear()
        for index, height in enumerate(CARD_PULSE_HEIGHTS):
            after_id = self.after(index * 36, lambda current_height=height: self.accent_strip.configure(height=current_height, bg=self.accent_color))
            self._pulse_after_ids.append(after_id)

    def _icon_text(self, title: str) -> str:
        words = [part[0] for part in title.split() if part]
        return "".join(words[:2]).upper() or "VF"

    def _extract_numeric(self, value: str) -> float | None:
        rendered = str(value).strip()
        if not rendered or "/" in rendered:
            return None
        cleaned = rendered.replace("EUR", "").replace(" ", "").replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return None

    def _animate_value(self, start: float, end: float, template: str) -> None:
        if self._value_after_id is not None:
            try:
                self.after_cancel(self._value_after_id)
            except tk.TclError:
                pass
            self._value_after_id = None

        if abs(end - start) < 0.01:
            self.value_var.set(template)
            return

        steps = VALUE_ANIMATION_STEPS
        delay_ms = VALUE_ANIMATION_DELAY_MS

        def render(step: int) -> None:
            progress = step / steps
            current = start + ((end - start) * progress)
            self.value_var.set(self._format_numeric(current, template))
            if step < steps:
                self._value_after_id = self.after(delay_ms, lambda: render(step + 1))
            else:
                self.value_var.set(template)
                self._value_after_id = None

        render(0)

    def _format_numeric(self, value: float, template: str) -> str:
        if "EUR" in template:
            return money(value)
        return str(int(round(value)))


class LineChart(tk.Canvas):
    def __init__(self, master, **kwargs) -> None:
        kwargs.setdefault("height", 250)
        super().__init__(master, bg=COLOR_SURFACE, highlightthickness=0, **kwargs)
        self.labels: list[str] = []
        self.series: list[tuple[str, list[float], str]] = []
        self._animation_after_id: str | None = None
        self.bind("<Configure>", lambda _event: self.redraw())

    def update_data(self, labels: list[str], series: list[tuple[str, list[float], str]]) -> None:
        self.labels = labels
        target_series = [(name, [float(value) for value in values], color) for name, values, color in series]
        start_map = {name: values for name, values, _color in self.series}
        start_series = []
        for name, values, color in target_series:
            previous = start_map.get(name, [])
            normalized = [float(previous[index]) if index < len(previous) else 0.0 for index in range(len(values))]
            start_series.append((name, normalized, color))
        self._animate_series(start_series, target_series)

    def _animate_series(self, start_series, target_series) -> None:
        if self._animation_after_id is not None:
            try:
                self.after_cancel(self._animation_after_id)
            except tk.TclError:
                pass
            self._animation_after_id = None
        steps = CHART_ANIMATION_STEPS

        def render(step: int) -> None:
            progress = ease_out_cubic(step / steps)
            current = []
            for (name, start_values, color), (_, target_values, _target_color) in zip(start_series, target_series):
                interpolated = [
                    start_value + ((target_value - start_value) * progress)
                    for start_value, target_value in zip(start_values, target_values)
                ]
                current.append((name, interpolated, color))
            self.series = current
            self.redraw()
            if step < steps:
                self._animation_after_id = self.after(CHART_ANIMATION_DELAY_MS, lambda: render(step + 1))
            else:
                self.series = target_series
                self.redraw()
                self._animation_after_id = None

        render(0)

    def redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 420)
        height = max(self.winfo_height(), 240)
        if not self.labels or not self.series:
            self.create_text(width / 2, height / 2, text="Aucune donnee pour le moment", fill=COLOR_TEXT_MUTED, font=(FONT_BODY, 12))
            return

        max_value = max([max(values or [0]) for _, values, _ in self.series] + [1])
        left, top, right, bottom = 56, 28, width - 26, height - 40
        plot_width = max(right - left, 1)
        plot_height = max(bottom - top, 1)

        for step in range(5):
            y = top + plot_height * step / 4
            self.create_line(left, y, right, y, fill=COLOR_CHART_GRID, width=1)
            value = max_value - (max_value * step / 4)
            self.create_text(left - 10, y, text=f"{int(value)}", fill=COLOR_TEXT_MUTED, anchor="e", font=(FONT_BODY, 9))

        count = max(len(self.labels) - 1, 1)
        for index, label in enumerate(self.labels):
            x = left + plot_width * index / count
            self.create_text(x, bottom + 18, text=label, fill=COLOR_TEXT_MUTED, font=(FONT_BODY, 9))

        for series_index, (name, values, color) in enumerate(self.series):
            points = []
            for index, value in enumerate(values):
                x = left + plot_width * index / count
                y = bottom - (value / max_value) * plot_height if max_value else bottom
                points.extend([x, y])
            if len(points) >= 4:
                self.create_line(*points, fill=color, smooth=True, splinesteps=18, width=3)
            for point_index in range(0, len(points), 2):
                x, y = points[point_index], points[point_index + 1]
                self.create_oval(x - 4, y - 4, x + 4, y + 4, fill=color, outline=COLOR_SURFACE)
            legend_x = left + (140 * series_index)
            self.create_rectangle(legend_x, 8, legend_x + 10, 18, fill=color, outline=color)
            self.create_text(legend_x + 16, 13, text=name, anchor="w", fill=COLOR_TEXT, font=(FONT_BODY_SEMIBOLD, 10))


class HorizontalBarChart(tk.Canvas):
    def __init__(self, master, **kwargs) -> None:
        kwargs.setdefault("height", 250)
        super().__init__(master, bg=COLOR_SURFACE, highlightthickness=0, **kwargs)
        self.data: list[tuple[str, float]] = []
        self.bar_color = COLOR_TEAL
        self._animation_after_id: str | None = None
        self.bind("<Configure>", lambda _event: self.redraw())

    def update_data(self, data: list[tuple[str, float]], bar_color: str) -> None:
        self.bar_color = bar_color
        target_data = [(label, float(value)) for label, value in data]
        start_map = {label: value for label, value in self.data}
        start_data = [(label, float(start_map.get(label, 0.0))) for label, _value in target_data]
        self._animate_bars(start_data, target_data)

    def _animate_bars(self, start_data, target_data) -> None:
        if self._animation_after_id is not None:
            try:
                self.after_cancel(self._animation_after_id)
            except tk.TclError:
                pass
            self._animation_after_id = None
        steps = CHART_ANIMATION_STEPS

        def render(step: int) -> None:
            progress = ease_out_cubic(step / steps)
            self.data = [
                (label, start_value + ((target_value - start_value) * progress))
                for (label, start_value), (_target_label, target_value) in zip(start_data, target_data)
            ]
            self.redraw()
            if step < steps:
                self._animation_after_id = self.after(CHART_ANIMATION_DELAY_MS, lambda: render(step + 1))
            else:
                self.data = target_data
                self.redraw()
                self._animation_after_id = None

        render(0)

    def redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 340)
        height = max(self.winfo_height(), 240)
        if not self.data:
            self.create_text(width / 2, height / 2, text="Aucune categorie a afficher", fill=COLOR_TEXT_MUTED, font=(FONT_BODY, 12))
            return

        left = 18
        top = 18
        row_height = 42
        max_value = max(value for _, value in self.data) or 1
        for index, (label, value) in enumerate(self.data):
            y = top + index * row_height
            self.create_text(left, y + 12, text=label, fill=COLOR_TEXT, anchor="w", font=(FONT_BODY_SEMIBOLD, 10))
            rounded_parts(self, left, y + 20, width - 92, y + 32, 6, COLOR_BAR_TRACK)
            ratio = value / max_value
            bar_end = left + (width - 110) * ratio
            rounded_parts(self, left, y + 20, max(int(bar_end), left + 12), y + 32, 6, self.bar_color)
            self.create_text(width - 16, y + 26, text=money(value), fill=COLOR_TEXT_MUTED, anchor="e", font=(FONT_BODY, 10))


class DonutChart(tk.Canvas):
    def __init__(self, master, **kwargs) -> None:
        kwargs.setdefault("height", 280)
        super().__init__(master, bg=COLOR_SURFACE, highlightthickness=0, **kwargs)
        self.data: list[tuple[str, float, str]] = []
        self.center_title = "Flux"
        self.center_value = ""
        self._animation_after_id: str | None = None
        self.bind("<Configure>", lambda _event: self.redraw())

    def update_data(self, data: list[tuple[str, float, str]], center_title: str, center_value: str) -> None:
        self.center_title = center_title
        self.center_value = center_value
        target_data = [(label, float(value), color) for label, value, color in data]
        start_map = {label: value for label, value, _color in self.data}
        start_data = [(label, float(start_map.get(label, 0.0)), color) for label, _value, color in target_data]
        self._animate_donut(start_data, target_data)

    def _animate_donut(self, start_data, target_data) -> None:
        if self._animation_after_id is not None:
            try:
                self.after_cancel(self._animation_after_id)
            except tk.TclError:
                pass
            self._animation_after_id = None
        steps = CHART_ANIMATION_STEPS

        def render(step: int) -> None:
            progress = ease_out_cubic(step / steps)
            self.data = [
                (label, start_value + ((target_value - start_value) * progress), color)
                for (label, start_value, color), (_label, target_value, _target_color) in zip(start_data, target_data)
            ]
            self.redraw()
            if step < steps:
                self._animation_after_id = self.after(CHART_ANIMATION_DELAY_MS, lambda: render(step + 1))
            else:
                self.data = target_data
                self.redraw()
                self._animation_after_id = None

        render(0)

    def redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 320)
        height = max(self.winfo_height(), 280)
        positive = [(label, value, color) for label, value, color in self.data if value > 0]
        if not positive:
            self.create_text(width / 2, height / 2, text="Aucune repartition a afficher", fill=COLOR_TEXT_MUTED, font=(FONT_BODY, 12))
            return

        legend_top = height - 86
        chart_size = min(width - 80, height - 120, 220)
        left = (width - chart_size) / 2
        top = 12
        right = left + chart_size
        bottom = top + chart_size
        total = sum(value for _, value, _ in positive) or 1
        start_angle = 90
        ring_width = max(int(chart_size * 0.22), 24)

        for label, value, color in positive:
            extent = -(value / total) * 360
            self.create_arc(
                left,
                top,
                right,
                bottom,
                start=start_angle,
                extent=extent,
                style="arc",
                outline=color,
                width=ring_width,
            )
            start_angle += extent

        self.create_oval(
            left + ring_width,
            top + ring_width,
            right - ring_width,
            bottom - ring_width,
            fill=COLOR_SURFACE,
            outline=COLOR_SURFACE,
        )
        self.create_text(width / 2, top + (chart_size / 2) - 10, text=self.center_title, fill=COLOR_TEXT_MUTED, font=(FONT_BODY_SEMIBOLD, 10))
        self.create_text(width / 2, top + (chart_size / 2) + 18, text=self.center_value, fill=COLOR_TEXT, font=(FONT_NUMERIC, 15))

        columns = 2
        column_width = (width - 36) / columns
        for index, (label, value, color) in enumerate(positive[:4]):
            row = index // columns
            column = index % columns
            x = 18 + (column * column_width)
            y = legend_top + (row * 28)
            self.create_rectangle(x, y + 4, x + 12, y + 16, fill=color, outline=color)
            self.create_text(x + 18, y + 10, text=label, anchor="w", fill=COLOR_TEXT, font=(FONT_BODY_SEMIBOLD, 9))
            self.create_text(x + column_width - 8, y + 10, text=money(value), anchor="e", fill=COLOR_TEXT_MUTED, font=(FONT_BODY, 9))


class ActivityTimelineChart(tk.Canvas):
    def __init__(self, master, **kwargs) -> None:
        kwargs.setdefault("height", 260)
        super().__init__(master, bg=COLOR_SURFACE, highlightthickness=0, **kwargs)
        self.data: list[dict] = []
        self._hotspots: list[dict] = []
        self._hover_callback = None
        self.bind("<Configure>", lambda _event: self.redraw())
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", self._on_leave)

    def set_hover_callback(self, callback) -> None:
        self._hover_callback = callback

    def update_data(self, data: list[dict]) -> None:
        self.data = list(data)
        self.redraw()

    def redraw(self) -> None:
        self.delete("all")
        self._hotspots.clear()
        width = max(self.winfo_width(), 420)
        height = max(self.winfo_height(), 260)
        if not self.data:
            self.create_text(width / 2, height / 2, text="Aucune activite recente", fill=COLOR_TEXT_MUTED, font=(FONT_BODY, 12))
            return

        parsed = []
        for item in self.data:
            parsed.append({**item, "_date": parse_iso_date(item["date"])})
        min_date = min(item["_date"] for item in parsed)
        max_date = max(item["_date"] for item in parsed)
        span_days = max((max_date - min_date).days, 1)
        max_amount = max(item["amount"] for item in parsed) or 1
        left, top, right, bottom = 58, 20, width - 18, height - 44
        plot_width = max(right - left, 1)
        plot_height = max(bottom - top, 1)

        for step in range(4):
            y = top + plot_height * step / 3
            self.create_line(left, y, right, y, fill=COLOR_CHART_GRID, width=1)
            value = max_amount - (max_amount * step / 3)
            self.create_text(left - 10, y, text=short_money(value), fill=COLOR_TEXT_MUTED, anchor="e", font=(FONT_BODY, 9))

        for offset in range(5):
            progress = offset / 4
            current_day = min_date + timedelta(days=int(span_days * progress))
            x = left + (plot_width * progress)
            self.create_text(x, bottom + 18, text=current_day.strftime("%d/%m"), fill=COLOR_TEXT_MUTED, font=(FONT_BODY, 9))

        self.create_line(left, bottom, right, bottom, fill=COLOR_BORDER, width=1)

        color_map = {
            "invoice": COLOR_SUCCESS,
            "sale": COLOR_TEAL,
            "expense": COLOR_GOLD,
            "employee_expense": COLOR_CORAL,
        }
        for item in parsed:
            x = left + (((item["_date"] - min_date).days / span_days) * plot_width)
            y = bottom - ((item["amount"] / max_amount) * plot_height)
            color = color_map.get(item["kind"], COLOR_TEAL)
            radius = 5 if item["direction"] == "income" else 6
            self.create_oval(x - radius, y - radius, x + radius, y + radius, fill=color, outline=COLOR_SURFACE, width=2)
            self._hotspots.append(
                {
                    "x": x,
                    "y": y,
                    "radius": radius + 4,
                    "text": f"{format_day(item['date'])} | {money(item['amount'])} | {item['label']}",
                }
            )

        legend = [
            ("Facture", COLOR_SUCCESS),
            ("Recette", COLOR_TEAL),
            ("Depense", COLOR_GOLD),
            ("Paie", COLOR_CORAL),
        ]
        for index, (label, color) in enumerate(legend):
            x = left + (index * 102)
            y = 4
            self.create_rectangle(x, y + 5, x + 12, y + 17, fill=color, outline=color)
            self.create_text(x + 18, y + 11, text=label, anchor="w", fill=COLOR_TEXT_MUTED, font=(FONT_BODY_SEMIBOLD, 9))

    def _on_motion(self, event) -> None:
        for item in self._hotspots:
            if abs(event.x - item["x"]) <= item["radius"] and abs(event.y - item["y"]) <= item["radius"]:
                if self._hover_callback:
                    self._hover_callback(item["text"])
                return
        if self._hover_callback:
            self._hover_callback("Survolez un point pour voir la date, le montant et l'origine.")

    def _on_leave(self, _event) -> None:
        if self._hover_callback:
            self._hover_callback("Survolez un point pour voir la date, le montant et l'origine.")


class DashboardPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.scroll_shell = ScrollableFrame(self, COLOR_BG)
        self.scroll_shell.pack(fill="both", expand=True)
        shell = self.scroll_shell.content

        header = ttk.Frame(shell, style="App.TFrame")
        header.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(header, text="Tableau de bord", style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="Vue d'ensemble des flux, du chiffre d'affaires, des dépenses et des priorités en cours.", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        cards = ttk.Frame(shell, style="App.TFrame")
        cards.pack(fill="x", padx=24)
        cards.columnconfigure((0, 1, 2, 3, 4), weight=1)
        self.revenue_card = MetricCard(cards, "Chiffre d'affaires", COLOR_TEAL)
        self.profit_card = MetricCard(cards, "Benefice estime", COLOR_CORAL)
        self.expense_card = MetricCard(cards, "Depenses", COLOR_GOLD)
        self.invoice_card = MetricCard(cards, "Factures / devis", COLOR_NAVY)
        self.todo_card = MetricCard(cards, "Taches ouvertes", COLOR_SUCCESS)
        self.metric_cards = [self.revenue_card, self.profit_card, self.expense_card, self.invoice_card, self.todo_card]
        for column, widget in enumerate(
            self.metric_cards
        ):
            widget.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 12, 0))

        charts = ttk.Frame(shell, style="App.TFrame")
        charts.pack(fill="both", expand=True, padx=24, pady=20)
        charts.columnconfigure(0, weight=3)
        charts.columnconfigure(1, weight=2)
        charts.rowconfigure(0, weight=1)
        charts.rowconfigure(1, weight=1)

        self.trend_card = self._chart_card(charts, "Evolution sur 6 mois", "Le chiffre d'affaires confirme reste au centre du pilotage")
        self.trend_card.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        self.trend_chart = LineChart(self.trend_card)
        self.trend_chart.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        self.donut_card = self._chart_card(charts, "Repartition des flux", "Factures, recettes manuelles, depenses et paie en un coup d'oeil")
        self.donut_card.grid(row=0, column=1, sticky="nsew")
        self.flow_donut = DonutChart(self.donut_card)
        self.flow_donut.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        self.right_stack = ttk.Frame(charts, style="App.TFrame")
        self.right_stack.grid(row=1, column=1, sticky="nsew", pady=(12, 0))
        self.right_stack.columnconfigure(0, weight=1)
        self.right_stack.rowconfigure((0, 1), weight=1)

        self.sales_card = self._chart_card(self.right_stack, "Recettes par categorie", "Factures confirmees + recettes manuelles")
        self.sales_card.grid(row=0, column=0, sticky="nsew")
        self.sales_chart = HorizontalBarChart(self.sales_card, height=170)
        self.sales_chart.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        self.expense_category_card = self._chart_card(self.right_stack, "Depenses par categorie", "Depenses generales + paie employe")
        self.expense_category_card.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.expense_chart = HorizontalBarChart(self.expense_category_card, height=170)
        self.expense_chart.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        self.activity_card = self._chart_card(shell, "Activite quotidienne", "Chaque point represente une facture, une recette, une depense ou une paie avec sa date")
        self.activity_card.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.activity_chart = ActivityTimelineChart(self.activity_card)
        self.activity_chart.pack(fill="both", expand=True, padx=4, pady=(0, 8))
        self.activity_hover_var = tk.StringVar(value="Survolez un point pour voir la date, le montant et l'origine.")
        ttk.Label(self.activity_card, textvariable=self.activity_hover_var, style="MutedSurface.TLabel").pack(anchor="w", padx=4)
        self.activity_chart.set_hover_callback(self.activity_hover_var.set)

        foot = ttk.Frame(shell, style="App.TFrame")
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
        self.revenue_card.set(money(snapshot["revenue_total"]), f"CA confirme | brouillons {money(snapshot['draft_revenue_total'])}")
        self.profit_card.set(money(snapshot["profit_total"]), "CA confirme - depenses TTC")
        self.expense_card.set(money(snapshot["expenses_total"]), f"Dont paie {money(snapshot['employee_expenses_total'])}")
        self.invoice_card.set(f"{snapshot['open_invoices']} / {snapshot['pending_quotes']}", f"Factures a relancer | brouillons {snapshot['draft_invoices']}")
        self.todo_card.set(str(snapshot["pending_todos"]), "Taches encore a traiter")

        labels = [label for label, _ in snapshot["revenue_series"]]
        revenue_values = [value for _, value in snapshot["revenue_series"]]
        profit_values = [value for _, value in snapshot["profit_series"]]
        self.trend_chart.update_data(labels, [("CA", revenue_values, COLOR_TEAL), ("Benefice", profit_values, COLOR_CORAL)])
        self.flow_donut.update_data(
            [
                ("Factures", snapshot["invoice_revenue_total"], COLOR_TEAL),
                ("Recettes", snapshot["manual_revenue_total"], COLOR_SUCCESS),
                ("Depenses", snapshot["general_expenses_total"], COLOR_GOLD),
                ("Paie", snapshot["employee_expenses_total"], COLOR_CORAL),
            ],
            "Flux du mois",
            money(snapshot["revenue_total"]),
        )
        self.sales_chart.update_data(snapshot["sales_categories"], COLOR_TEAL)
        self.expense_chart.update_data(snapshot["expense_categories"], COLOR_CORAL)
        self.activity_chart.update_data(snapshot["activity_points"])
        self.activity_hover_var.set("Survolez un point pour voir la date, le montant et l'origine.")
        self._fill_latest_panel(self.invoices_panel, snapshot["latest_invoices"], "doc")
        self._fill_latest_panel(self.quotes_panel, snapshot["latest_quotes"], "doc")
        self._fill_latest_panel(self.todos_panel, snapshot["latest_todos"], "todo")

    def animate_in(self) -> None:
        for index, card in enumerate(self.metric_cards):
            self.after(index * 60, card.pulse)
        for index, panel in enumerate([self.trend_card, self.donut_card, self.sales_card, self.expense_category_card, self.activity_card], start=1):
            self.after(index * 65, lambda current_panel=panel: pulse_panel(current_panel))
        for index, panel in enumerate([self.invoices_panel, self.quotes_panel, self.todos_panel], start=1):
            self.after((index + 5) * 65, lambda current_panel=panel: pulse_panel(current_panel))

    def scroll_units(self, units: int) -> None:
        self.scroll_shell.scroll_units(units)

    def scroll_pages(self, pages: int) -> None:
        self.scroll_shell.scroll_pages(pages)


class BaseEntryPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp", title: str, subtitle: str) -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.scroll_shell = ScrollableFrame(self, COLOR_BG)
        self.scroll_shell.pack(fill="both", expand=True)

        head = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        head.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(head, text=title, style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(head, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        self.content = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        self.content.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        self.content.columnconfigure(0, weight=2)
        self.content.columnconfigure(1, weight=3)
        self.form_card = ttk.Frame(self.content, style="Surface.TFrame", padding=20)
        self.form_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.table_card = ttk.Frame(self.content, style="Surface.TFrame", padding=20)
        self.table_card.grid(row=0, column=1, sticky="nsew")

    def _create_label(self, master, text: str, row: int, column: int = 0) -> None:
        ttk.Label(master, text=text, style="FieldLabel.TLabel").grid(row=row, column=column, sticky="w", pady=(0, 6))

    def animate_in(self) -> None:
        pulse_panel(self.form_card)
        self.after(70, lambda: pulse_panel(self.table_card))

    def scroll_units(self, units: int) -> None:
        self.scroll_shell.scroll_units(units)

    def scroll_pages(self, pages: int) -> None:
        self.scroll_shell.scroll_pages(pages)


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
        self.attachment_label_var = tk.StringVar(value="Aucune facture de depense selectionnee")
        self.attachment_source_path = ""

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

        action_row = 18
        help_row = 19
        if self.record_kind == "expense":
            self._create_label(self.form_card, "Facture de depense locale", 17)
            file_actions = ttk.Frame(self.form_card, style="Surface.TFrame")
            file_actions.grid(row=18, column=0, columnspan=2, sticky="ew", pady=(0, 8))
            RoundButton(file_actions, text="Choisir un fichier", style_name="Ghost.TButton", command=self.pick_attachment).pack(side="left")
            RoundButton(file_actions, text="Ouvrir", style_name="Ghost.TButton", command=self.open_attachment).pack(side="left", padx=8)
            RoundButton(file_actions, text="Retirer", style_name="Ghost.TButton", command=self.clear_attachment).pack(side="left")
            ttk.Label(self.form_card, textvariable=self.attachment_label_var, style="MutedSurface.TLabel", wraplength=260).grid(row=19, column=0, columnspan=2, sticky="w", pady=(0, 10))
            action_row = 20
            help_row = 21

        RoundButton(self.form_card, text=self.action_label, style_name=self.button_style, command=self.add_record).grid(row=action_row, column=0, sticky="ew", pady=(8, 0))
        help_text = (
            "Mode facture: les informations HT/TTC et le nom client sont recuperes automatiquement. Mode manuel: vous saisissez vos propres montants."
            if self.record_kind == "sale"
            else "Mode facture: recupere une facture client existante. Mode manuel: ajoutez vos montants. Facture de depense locale: joignez un fichier de votre machine qui sera copie localement dans le logiciel."
        )
        ttk.Label(
            self.form_card,
            text=help_text,
            style="MutedSurface.TLabel",
            wraplength=260,
        ).grid(row=help_row, column=0, sticky="w", pady=(12, 0))
        self.form_card.columnconfigure(0, weight=1)

    def pick_attachment(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choisir une facture de depense",
            filetypes=[
                ("Documents", "*.pdf *.html *.htm *.png *.jpg *.jpeg *.webp"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        if filename:
            self.attachment_source_path = filename
            self.attachment_label_var.set(Path(filename).name)

    def open_attachment(self) -> None:
        source = self.attachment_source_path.strip()
        if not source:
            return
        try:
            webbrowser.open(Path(source).resolve().as_uri())
        except Exception:
            messagebox.showerror("Depense", "Impossible d'ouvrir la facture locale selectionnee.")

    def clear_attachment(self) -> None:
        self.attachment_source_path = ""
        self.attachment_label_var.set("Aucune facture de depense selectionnee")

    def _build_table(self) -> None:
        top = ttk.Frame(self.table_card, style="Surface.TFrame")
        top.pack(fill="x")
        table_title = "Historique des recettes" if self.record_kind == "sale" else "Historique des depenses"
        ttk.Label(top, text=table_title, style="SectionTitle.TLabel").pack(side="left")
        RoundButton(top, text="Supprimer la selection", style_name="Ghost.TButton", command=self.delete_selected).pack(side="right")

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
        copied_attachment = ""
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
            if self.record_kind == "expense" and self.attachment_source_path.strip():
                copied_attachment = str(copy_expense_attachment(self.attachment_source_path))
                payload["attachment_path"] = copied_attachment
        except Exception as exc:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer.\n\n{exc}")
            return

        try:
            if self.record_kind == "sale":
                self.app.db.add_sale(payload)
            else:
                self.app.db.add_expense(payload)
        except Exception as exc:
            if copied_attachment:
                Path(copied_attachment).unlink(missing_ok=True)
            messagebox.showerror("Erreur", f"Impossible d'enregistrer.\n\n{exc}")
            return

        self.company_var.set("")
        self.amount_ht_var.set("")
        self.amount_ttc_var.set("")
        self.notes_var.set("")
        if self.record_kind == "expense":
            self.clear_attachment()
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
            "Depenses generales",
            "Achats, logiciels, marketing et autres frais hors paie. Vous pouvez lier une facture client existante, saisir manuellement ou joindre une vraie facture locale.",
            "expense",
            "Achats",
            ["Achats", "Marketing", "Logiciels", "Transport", "Sous-traitance", "Autre"],
            "Ajouter la depense",
            "AccentAlt.TButton",
        )

    def _build_table(self) -> None:
        top = ttk.Frame(self.table_card, style="Surface.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="Historique des depenses", style="SectionTitle.TLabel").pack(side="left")
        RoundButton(top, text="Ouvrir la facture", style_name="Ghost.TButton", command=self.open_selected_attachment).pack(side="right", padx=(0, 8))
        RoundButton(top, text="Supprimer la selection", style_name="Ghost.TButton", command=self.delete_selected).pack(side="right")

        export_bar = ttk.Frame(self.table_card, style="Surface.TFrame")
        export_bar.pack(fill="x", pady=(12, 12))
        ttk.Label(export_bar, text="Export depenses seulement", style="FieldLabel.TLabel").pack(side="left")
        self.export_month_widget = ttk.Combobox(export_bar, textvariable=self.export_month_var, state="readonly", width=12)
        self.export_month_widget.pack(side="left", padx=(10, 8))
        RoundButton(export_bar, text="Exporter les depenses", style_name="Ghost.TButton", command=self.export_selected_month).pack(side="left")

        columns = ("date", "company", "source", "category", "ht", "ttc", "file", "added")
        self.tree = ttk.Treeview(self.table_card, columns=columns, show="headings", height=18)
        headings = {
            "date": "Date",
            "company": "Entreprise",
            "source": "Source",
            "category": "Categorie",
            "ht": "HT",
            "ttc": "TTC",
            "file": "Facture",
            "added": "Ajout",
        }
        widths = {"date": 95, "company": 160, "source": 90, "category": 105, "ht": 90, "ttc": 90, "file": 110, "added": 130}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.pack(fill="both", expand=True, pady=(0, 0))

    def refresh(self) -> None:
        self.refresh_invoice_options()
        self.rows.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.app.db.list_expenses("general"):
            record = dict(row)
            iid = str(record["id"])
            self.rows[iid] = record
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    record["expense_date"],
                    record["company_name"],
                    source_label(record["source_type"]),
                    record["category"],
                    money(record["amount_ht"]),
                    money(record["amount_ttc"]),
                    Path(record["attachment_path"]).name if record.get("attachment_path") else "-",
                    format_timestamp(record["created_at"]),
                ),
            )
        months = collect_month_keys([dict(row) for row in self.app.db.list_expenses("general")], "expense_date")
        self.export_month_widget.configure(values=months)
        if months:
            if self.export_month_var.get() not in months:
                self.export_month_var.set(months[0])
        else:
            self.export_month_var.set("")

    def open_selected_attachment(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        record = self.rows.get(selection[0], {})
        attachment = str(record.get("attachment_path", "")).strip()
        if not attachment:
            messagebox.showwarning("Depense", "Aucune facture locale rattachee a cette depense.")
            return
        try:
            webbrowser.open(Path(attachment).resolve().as_uri())
        except Exception:
            messagebox.showerror("Depense", "Impossible d'ouvrir la facture de depense.")

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
            [dict(row) for row in self.app.db.list_expenses("general")],
            month_key,
            Path(destination),
        )
        if export_result.missing_documents:
            messagebox.showwarning(
                "Export depenses",
                "L'export est termine, mais certaines factures locales etaient manquantes.\n\n"
                f"Dossier: {export_result.export_root}\n"
                f"Rapport: {export_result.report_path.name}\n"
                f"Documents manquants: {len(export_result.missing_documents)}",
            )
            return
        messagebox.showinfo(
            "Export depenses",
            "Les depenses de "
            f"{month_key} ont ete exportees dans un dossier separe:\n\n{export_result.export_root}\n\n"
            f"Rapport genere: {export_result.report_path.name}",
        )


class EmployeeExpensesPage(BaseEntryPage):
    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(
            master,
            app,
            "Paie employe",
            "Centralisez salaires, primes, charges et remboursements pour suivre la vraie sante de l'entreprise.",
        )
        self.rows: dict[str, dict] = {}
        self.export_month_var = tk.StringVar(master=app)
        self.date_var = tk.StringVar(value=date.today().isoformat())
        self.payroll_month_var = tk.StringVar(value=date.today().strftime("%Y-%m"))
        self.employee_var = tk.StringVar()
        self.label_var = tk.StringVar(value="Salaire")
        self.amount_ht_var = tk.StringVar()
        self.amount_ttc_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        self.attachment_label_var = tk.StringVar(value="Aucun justificatif de paie selectionne")
        self.attachment_source_path = ""

        self._build_form()
        self._build_table()

    def _build_form(self) -> None:
        ttk.Label(self.form_card, text="Nouvelle depense employe", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(
            self.form_card,
            text="Une ligne = un employe, un mois, un montant et un justificatif local si besoin.",
            style="MutedSurface.TLabel",
            wraplength=280,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 12))

        fields = [
            ("Date", self.date_var),
            ("Mois de paie", self.payroll_month_var),
            ("Employe", self.employee_var),
            ("Type", self.label_var),
            ("Montant HT", self.amount_ht_var),
            ("Montant TTC", self.amount_ttc_var),
            ("Notes", self.notes_var),
        ]
        for index, (label, variable) in enumerate(fields, start=1):
            self._create_label(self.form_card, label, index * 2)
            if label == "Type":
                widget = ttk.Combobox(
                    self.form_card,
                    textvariable=variable,
                    values=["Salaire", "Prime", "Charge patronale", "Remboursement", "Avantage", "Autre"],
                    state="readonly",
                )
            else:
                widget = ttk.Entry(self.form_card, textvariable=variable)
            widget.grid(row=index * 2 + 1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        self._create_label(self.form_card, "Justificatif local", 16)
        file_actions = ttk.Frame(self.form_card, style="Surface.TFrame")
        file_actions.grid(row=17, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        RoundButton(file_actions, text="Choisir un fichier", style_name="Ghost.TButton", command=self.pick_attachment).pack(side="left")
        RoundButton(file_actions, text="Ouvrir", style_name="Ghost.TButton", command=self.open_attachment).pack(side="left", padx=8)
        RoundButton(file_actions, text="Retirer", style_name="Ghost.TButton", command=self.clear_attachment).pack(side="left")
        ttk.Label(self.form_card, textvariable=self.attachment_label_var, style="MutedSurface.TLabel", wraplength=280).grid(row=18, column=0, columnspan=2, sticky="w", pady=(0, 10))

        actions = ttk.Frame(self.form_card, style="Surface.TFrame")
        actions.grid(row=19, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        RoundButton(actions, text="Ajouter la depense employe", style_name="AccentAlt.TButton", command=self.add_record).pack(side="left")
        RoundButton(actions, text="Vider", style_name="Ghost.TButton", command=self.clear_form).pack(side="left", padx=8)
        ttk.Label(
            self.form_card,
            text="La paie est ajoutee dans la sante d'entreprise, le calendrier et les exports.",
            style="MutedSurface.TLabel",
            wraplength=280,
        ).grid(row=20, column=0, columnspan=2, sticky="w", pady=(12, 0))
        self.form_card.columnconfigure(0, weight=1)

    def _build_table(self) -> None:
        top = ttk.Frame(self.table_card, style="Surface.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="Historique paie", style="SectionTitle.TLabel").pack(side="left")
        RoundButton(top, text="Ouvrir le justificatif", style_name="Ghost.TButton", command=self.open_selected_attachment).pack(side="right", padx=(0, 8))
        RoundButton(top, text="Supprimer la selection", style_name="Ghost.TButton", command=self.delete_selected).pack(side="right")

        export_bar = ttk.Frame(self.table_card, style="Surface.TFrame")
        export_bar.pack(fill="x", pady=(12, 12))
        ttk.Label(export_bar, text="Export paie seulement", style="FieldLabel.TLabel").pack(side="left")
        self.export_month_widget = ttk.Combobox(export_bar, textvariable=self.export_month_var, state="readonly", width=12)
        self.export_month_widget.pack(side="left", padx=(10, 8))
        RoundButton(export_bar, text="Exporter la paie", style_name="Ghost.TButton", command=self.export_selected_month).pack(side="left")

        columns = ("date", "employee", "label", "period", "ttc", "file", "added")
        self.tree = ttk.Treeview(self.table_card, columns=columns, show="headings", height=18)
        headings = {
            "date": "Date",
            "employee": "Employe",
            "label": "Type",
            "period": "Mois",
            "ttc": "TTC",
            "file": "Piece",
            "added": "Ajout",
        }
        widths = {"date": 95, "employee": 160, "label": 125, "period": 90, "ttc": 90, "file": 110, "added": 130}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.pack(fill="both", expand=True)

    def pick_attachment(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choisir un justificatif de paie",
            filetypes=[("Documents", "*.pdf *.html *.htm *.png *.jpg *.jpeg *.webp"), ("Tous les fichiers", "*.*")],
        )
        if filename:
            self.attachment_source_path = filename
            self.attachment_label_var.set(Path(filename).name)

    def open_attachment(self) -> None:
        source = self.attachment_source_path.strip()
        if not source:
            return
        try:
            webbrowser.open(Path(source).resolve().as_uri())
        except Exception:
            messagebox.showerror("Paie", "Impossible d'ouvrir le justificatif selectionne.")

    def clear_attachment(self) -> None:
        self.attachment_source_path = ""
        self.attachment_label_var.set("Aucun justificatif de paie selectionne")

    def clear_form(self) -> None:
        self.date_var.set(date.today().isoformat())
        self.payroll_month_var.set(date.today().strftime("%Y-%m"))
        self.employee_var.set("")
        self.label_var.set("Salaire")
        self.amount_ht_var.set("")
        self.amount_ttc_var.set("")
        self.notes_var.set("")
        self.clear_attachment()

    def add_record(self) -> None:
        copied_attachment = ""
        try:
            parse_iso_date(self.date_var.get())
            if len(self.payroll_month_var.get().strip()) != 7:
                raise ValueError("Le mois de paie doit etre au format AAAA-MM.")
            employee_name = self.employee_var.get().strip()
            if not employee_name:
                raise ValueError("Le nom de l'employe est obligatoire.")
            amount_ht = parse_amount(self.amount_ht_var.get())
            amount_ttc = parse_amount(self.amount_ttc_var.get())
            validate_money_pair(amount_ht, amount_ttc)
            payload = {
                "expense_date": self.date_var.get(),
                "company_name": employee_name,
                "category": "Salaires",
                "expense_kind": "employee",
                "expense_label": self.label_var.get().strip() or "Salaire",
                "employee_name": employee_name,
                "payroll_month": self.payroll_month_var.get().strip(),
                "amount_ht": amount_ht,
                "amount_ttc": amount_ttc,
                "source_type": "manual",
                "source_invoice_id": None,
                "notes": self.notes_var.get().strip(),
            }
            if self.attachment_source_path.strip():
                copied_attachment = str(copy_expense_attachment(self.attachment_source_path))
                payload["attachment_path"] = copied_attachment
        except Exception as exc:
            messagebox.showerror("Paie", f"Impossible d'enregistrer.\n\n{exc}")
            return

        try:
            self.app.db.add_expense(payload)
        except Exception as exc:
            if copied_attachment:
                Path(copied_attachment).unlink(missing_ok=True)
            messagebox.showerror("Paie", f"Impossible d'enregistrer.\n\n{exc}")
            return

        self.clear_form()
        self.app.refresh_all_pages()

    def refresh(self) -> None:
        self.rows.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.app.db.list_expenses("employee"):
            record = dict(row)
            iid = str(record["id"])
            self.rows[iid] = record
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    record["expense_date"],
                    record.get("employee_name") or record["company_name"],
                    record.get("expense_label") or "Salaire",
                    record.get("payroll_month") or "-",
                    money(record["amount_ttc"]),
                    Path(record["attachment_path"]).name if record.get("attachment_path") else "-",
                    format_timestamp(record["created_at"]),
                ),
            )
        months = collect_month_keys([dict(row) for row in self.app.db.list_expenses("employee")], "expense_date")
        self.export_month_widget.configure(values=months)
        if months:
            if self.export_month_var.get() not in months:
                self.export_month_var.set(months[0])
        else:
            self.export_month_var.set("")

    def open_selected_attachment(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        record = self.rows.get(selection[0], {})
        attachment = str(record.get("attachment_path", "")).strip()
        if not attachment:
            messagebox.showwarning("Paie", "Aucun justificatif local rattache a cette ligne.")
            return
        try:
            webbrowser.open(Path(attachment).resolve().as_uri())
        except Exception:
            messagebox.showerror("Paie", "Impossible d'ouvrir le justificatif de paie.")

    def delete_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        if not messagebox.askyesno("Paie", "Supprimer la ligne selectionnee ?"):
            return
        self.app.db.delete_expense(int(selection[0]))
        self.app.refresh_all_pages()

    def export_selected_month(self) -> None:
        month_key = self.export_month_var.get().strip()
        if not month_key:
            messagebox.showwarning("Export", "Choisissez un mois a exporter.")
            return
        destination = filedialog.askdirectory(title="Choisir un dossier pour l'export de la paie")
        if not destination:
            return
        export_result = export_month_bundle(
            "employee_expense",
            [dict(row) for row in self.app.db.list_expenses("employee")],
            month_key,
            Path(destination),
        )
        if export_result.missing_documents:
            messagebox.showwarning(
                "Export paie",
                "L'export est termine, mais certains justificatifs locaux etaient manquants.\n\n"
                f"Dossier: {export_result.export_root}\n"
                f"Rapport: {export_result.report_path.name}\n"
                f"Documents manquants: {len(export_result.missing_documents)}",
            )
            return
        messagebox.showinfo(
            "Export paie",
            f"La paie de {month_key} a ete exportee dans un dossier separe:\n\n{export_result.export_root}\n\n"
            f"Rapport genere: {export_result.report_path.name}",
        )


class DocumentPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp", title: str, subtitle: str, kind: str) -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.kind = kind
        self.records: dict[str, dict] = {}
        self.export_month_var = tk.StringVar()
        self.scroll_shell = ScrollableFrame(self, COLOR_BG)
        self.scroll_shell.pack(fill="both", expand=True)

        head = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        head.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(head, text=title, style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(head, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        self.table_card = ttk.Frame(body, style="Surface.TFrame", padding=20)
        self.table_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.detail_card = ttk.Frame(body, style="Surface.TFrame", padding=20)
        self.detail_card.grid(row=0, column=1, sticky="nsew")

        self._build_table()
        self._build_detail()

    def _build_table(self) -> None:
        ttk.Label(self.table_card, text="Documents", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(self.table_card, text="HT, TTC, date du document, apercu et export mensuel local.", style="MutedSurface.TLabel").pack(anchor="w", pady=(2, 12))

        actions = ttk.Frame(self.table_card, style="Surface.TFrame")
        actions.pack(fill="x", pady=(0, 12))
        new_label = "Nouvelle facture" if self.kind == "invoice" else "Nouveau devis"
        RoundButton(actions, text=new_label, style_name="Accent.TButton", command=self.open_editor).pack(side="left")
        RoundButton(actions, text="Ouvrir HTML", style_name="Ghost.TButton", command=self.open_selected_html).pack(side="left", padx=8)
        RoundButton(actions, text="Supprimer", style_name="Danger.TButton", command=self.delete_selected).pack(side="right")

        export_actions = ttk.Frame(self.table_card, style="Surface.TFrame")
        export_actions.pack(fill="x", pady=(0, 12))
        ttk.Label(
            export_actions,
            text=f"Export {'factures' if self.kind == 'invoice' else 'devis'} seulement",
            style="FieldLabel.TLabel",
        ).pack(side="left")
        self.export_month_widget = ttk.Combobox(export_actions, textvariable=self.export_month_var, state="readonly", width=12)
        self.export_month_widget.pack(side="left", padx=(10, 8))
        RoundButton(
            export_actions,
            text=f"Exporter les {'factures' if self.kind == 'invoice' else 'devis'}",
            style_name="Ghost.TButton",
            command=self.export_selected_month,
        ).pack(side="left")

        status_actions = ttk.Frame(self.table_card, style="Surface.TFrame")
        status_actions.pack(fill="x", pady=(0, 12))
        for label in self.available_statuses():
            RoundButton(status_actions, text=label, style_name="Ghost.TButton", command=lambda value=label: self.change_status(value)).pack(side="left", padx=(0, 8))

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

        ttk.Label(preview, textvariable=self.preview_title_var, style="PageTitleSoft.TLabel").pack(anchor="w")
        ttk.Label(preview, textvariable=self.preview_meta_var, style="MutedSoft.TLabel").pack(anchor="w", pady=(4, 6))
        ttk.Label(preview, textvariable=self.preview_company_var, style="TableStrongSurface.TLabel").pack(anchor="w", pady=(0, 12))
        ttk.Label(preview, textvariable=self.preview_amounts_var, style="MetricMiniSoft.TLabel").pack(anchor="w", pady=(0, 12))

        self.preview_items = ttk.Treeview(preview, columns=("description", "qty", "price", "total"), show="headings", height=8)
        for column, heading, width in [("description", "Description", 220), ("qty", "Qt", 50), ("price", "Prix HT", 90), ("total", "Total HT", 90)]:
            self.preview_items.heading(column, text=heading)
            self.preview_items.column(column, width=width, anchor="w")
        self.preview_items.pack(fill="x", pady=(0, 12))

        ttk.Label(preview, text="Notes", style="SectionTitleSoft.TLabel").pack(anchor="w")
        self.preview_notes = tk.Text(preview, bg=COLOR_SURFACE_SOFT, bd=0, fg=COLOR_TEXT, height=5, font=(FONT_BODY, 10), wrap="word", highlightthickness=0)
        self.preview_notes.pack(fill="x", pady=(6, 12))
        style_text_widget(self.preview_notes, COLOR_SURFACE_SOFT)
        self.preview_notes.configure(state="disabled")

        ttk.Label(preview, text="Fichier local", style="SectionTitleSoft.TLabel").pack(anchor="w")
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
            f"Les {'factures' if self.kind == 'invoice' else 'devis'} de {month_key} ont ete exportes dans un dossier separe:\n\n"
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

    def animate_in(self) -> None:
        pulse_panel(self.table_card)
        self.after(70, lambda: pulse_panel(self.detail_card))

    def scroll_units(self, units: int) -> None:
        self.scroll_shell.scroll_units(units)

    def scroll_pages(self, pages: int) -> None:
        self.scroll_shell.scroll_pages(pages)


class CompanyPage(ttk.Frame):
    def __init__(self, master, app: "VeloraApp") -> None:
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.logo_var = tk.StringVar()

        self.scroll_shell = ScrollableFrame(self, COLOR_BG)
        self.scroll_shell.pack(fill="both", expand=True)

        head = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        head.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(head, text="Parametres", style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(head, text="Identite, documents et theme dans une zone plus simple et plus propre.", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
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
            "ui_theme": tk.StringVar(value="Clair"),
        }

        self._build_form()
        self._build_info_card()

    def _build_form(self) -> None:
        ttk.Label(self.form_card, text="Parametres", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(self.form_card, text="Trois onglets, uniquement l'essentiel.", style="MutedSurface.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 12))

        notebook = ttk.Notebook(self.form_card)
        notebook.grid(row=2, column=0, sticky="nsew")
        identity_tab = ttk.Frame(notebook, style="Surface.TFrame", padding=14)
        documents_tab = ttk.Frame(notebook, style="Surface.TFrame", padding=14)
        appearance_tab = ttk.Frame(notebook, style="Surface.TFrame", padding=14)
        notebook.add(identity_tab, text="Societe")
        notebook.add(documents_tab, text="Documents")
        notebook.add(appearance_tab, text="Theme")

        self._build_identity_tab(identity_tab)
        self._build_documents_tab(documents_tab)
        self._build_appearance_tab(appearance_tab)

        actions = ttk.Frame(self.form_card, style="Surface.TFrame")
        actions.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        RoundButton(actions, text="Enregistrer", style_name="Accent.TButton", command=self.save).pack(side="left")
        RoundButton(actions, text="Logo Velora", style_name="Ghost.TButton", command=self.use_default_logo).pack(side="left", padx=8)
        self.form_card.columnconfigure(0, weight=1)
        self.form_card.rowconfigure(2, weight=1)

    def _build_identity_tab(self, master) -> None:
        ttk.Label(master, text="SIRET: 14 chiffres. TVA: ex FR12345678901. Les formats sont verifies a l'enregistrement.", style="MutedSurface.TLabel", wraplength=420).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
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
            ttk.Label(master, text=label, style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
            ttk.Entry(master, textvariable=self.fields[key]).grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
            row_cursor += 2

        ttk.Label(master, text="Adresse", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
        self.address_text = tk.Text(master, height=4, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.address_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        style_text_widget(self.address_text)
        row_cursor += 2

        ttk.Label(master, text="Message bas de page", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
        self.footer_text = tk.Text(master, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.footer_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        style_text_widget(self.footer_text)
        row_cursor += 2

        ttk.Label(master, text="Logo", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
        ttk.Entry(master, textvariable=self.logo_var).grid(row=row_cursor + 1, column=0, sticky="ew")
        logo_actions = ttk.Frame(master, style="Surface.TFrame")
        logo_actions.grid(row=row_cursor + 1, column=1, sticky="e", padx=(12, 0))
        RoundButton(logo_actions, text="Parcourir", style_name="Ghost.TButton", command=self.pick_logo).pack(side="left")
        master.columnconfigure(0, weight=1)
        master.columnconfigure(1, weight=1)

    def _build_documents_tab(self, master) -> None:
        document_labels = [
            ("TVA par defaut %", "default_tax_rate"),
            ("Delai facture en jours", "invoice_due_days"),
            ("Validite devis en jours", "quote_validity_days"),
            ("Client par defaut", "default_client_name"),
            ("Email client par defaut", "default_client_email"),
            ("Statut facture par defaut", "default_invoice_status"),
            ("Statut devis par defaut", "default_quote_status"),
        ]
        row_cursor = 0
        for label, key in document_labels:
            ttk.Label(master, text=label, style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
            if key == "default_invoice_status":
                widget = ttk.Combobox(master, textvariable=self.document_fields[key], values=["Brouillon", "Envoyee", "Payee", "En retard"], state="readonly")
            elif key == "default_quote_status":
                widget = ttk.Combobox(master, textvariable=self.document_fields[key], values=["Brouillon", "Envoye", "Accepte", "Refuse", "Expire"], state="readonly")
            else:
                widget = ttk.Entry(master, textvariable=self.document_fields[key])
            widget.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
            row_cursor += 2

        ttk.Label(master, text="Adresse client par defaut", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
        self.default_client_address_text = tk.Text(master, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.default_client_address_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        style_text_widget(self.default_client_address_text)
        row_cursor += 2

        ttk.Label(master, text="Notes facture par defaut", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
        self.default_invoice_notes_text = tk.Text(master, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.default_invoice_notes_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        style_text_widget(self.default_invoice_notes_text)
        row_cursor += 2

        ttk.Label(master, text="Notes devis par defaut", style="FieldLabel.TLabel").grid(row=row_cursor, column=0, sticky="w", pady=(8, 6))
        self.default_quote_notes_text = tk.Text(master, height=3, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.default_quote_notes_text.grid(row=row_cursor + 1, column=0, columnspan=2, sticky="ew")
        style_text_widget(self.default_quote_notes_text)
        master.columnconfigure(0, weight=1)
        master.columnconfigure(1, weight=1)

    def _build_appearance_tab(self, master) -> None:
        ttk.Label(master, text="Theme de l'interface", style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w", pady=(8, 6))
        ttk.Combobox(master, textvariable=self.document_fields["ui_theme"], values=list(APP_THEMES), state="readonly").grid(row=1, column=0, sticky="ew")

        quick = ttk.Frame(master, style="SurfaceSoft.TFrame", padding=14)
        quick.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        ttk.Label(quick, text="Actions rapides", style="SectionTitleSoft.TLabel").pack(anchor="w")
        ttk.Label(quick, text="Un clic pour passer du clair au sombre ou remettre le logo d'origine.", style="MutedSoft.TLabel", wraplength=320).pack(anchor="w", pady=(4, 10))
        buttons = ttk.Frame(quick, style="SurfaceSoft.TFrame")
        buttons.pack(fill="x")
        RoundButton(buttons, text="Clair", style_name="Ghost.TButton", command=lambda: self.document_fields["ui_theme"].set("Clair")).pack(side="left")
        RoundButton(buttons, text="Sombre", style_name="Ghost.TButton", command=lambda: self.document_fields["ui_theme"].set("Sombre")).pack(side="left", padx=8)
        RoundButton(buttons, text="Logo Velora", style_name="Ghost.TButton", command=self.use_default_logo).pack(side="left")
        master.columnconfigure(0, weight=1)

    def _build_info_card(self) -> None:
        ttk.Label(self.info_card, text="Resume", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(self.info_card, text="Un panneau court avec les actions utiles et le stockage local.", style="MutedSurface.TLabel", wraplength=280).pack(anchor="w", pady=(4, 12))

        quick_actions = ttk.Frame(self.info_card, style="SurfaceSoft.TFrame", padding=14)
        quick_actions.pack(fill="x", pady=(0, 12))
        ttk.Label(quick_actions, text="Theme rapide", style="SectionTitleSoft.TLabel").pack(anchor="w")
        action_buttons = ttk.Frame(quick_actions, style="SurfaceSoft.TFrame")
        action_buttons.pack(fill="x", pady=(10, 0))
        RoundButton(action_buttons, text="Theme clair", style_name="Ghost.TButton", command=lambda: self.document_fields["ui_theme"].set("Clair")).pack(side="left")
        RoundButton(action_buttons, text="Theme sombre", style_name="Ghost.TButton", command=lambda: self.document_fields["ui_theme"].set("Sombre")).pack(side="left", padx=8)
        RoundButton(action_buttons, text="Logo Velora", style_name="Ghost.TButton", command=self.use_default_logo).pack(side="left")

        storage = storage_root()
        local_block = ttk.Frame(self.info_card, style="SurfaceSoft.TFrame", padding=14)
        local_block.pack(fill="x", pady=6)
        ttk.Label(local_block, text="Stockage local", style="TableStrongSurface.TLabel").pack(anchor="w")
        ttk.Label(local_block, text=str(storage), style="MutedSoft.TLabel", wraplength=260).pack(anchor="w", pady=(4, 0))

        docs_block = ttk.Frame(self.info_card, style="SurfaceSoft.TFrame", padding=14)
        docs_block.pack(fill="x", pady=6)
        ttk.Label(docs_block, text="Dossiers utiles", style="TableStrongSurface.TLabel").pack(anchor="w")
        ttk.Label(
            docs_block,
            text="factures, devis, depenses et paie restent sur cette machine.",
            style="MutedSoft.TLabel",
            wraplength=260,
        ).pack(anchor="w", pady=(4, 0))

        reset_block = ttk.Frame(self.info_card, style="SurfaceSoft.TFrame", padding=14)
        reset_block.pack(fill="x", pady=6)
        ttk.Label(reset_block, text="Remise a zero", style="TableStrongSurface.TLabel").pack(anchor="w")
        ttk.Label(
            reset_block,
            text="Supprime toutes les donnees locales, les graphiques, les factures, les devis et les depenses sans reinstaller le logiciel.",
            style="MutedSoft.TLabel",
            wraplength=260,
        ).pack(anchor="w", pady=(4, 10))
        RoundButton(reset_block, text="Supprimer toutes les donnees", style_name="Danger.TButton", command=self.reset_software).pack(anchor="w")

    def pick_logo(self) -> None:
        filename = filedialog.askopenfilename(title="Choisir un logo", filetypes=[("Images", "*.png *.gif *.ppm *.jpg *.jpeg *.svg"), ("Tous les fichiers", "*.*")])
        if filename:
            self.logo_var.set(filename)

    def use_default_logo(self) -> None:
        self.logo_var.set(str(asset_path("velora_logo.png")))

    def reset_software(self) -> None:
        confirmed = messagebox.askokcancel(
            "Remise a zero",
            "Attention.\n\nCette action va supprimer toutes les donnees du logiciel sur cette machine:\n"
            "- factures\n- devis\n- recettes\n- depenses\n- paie\n- todo liste\n- parametres enregistres\n\n"
            "Confirmer pour repartir de zero, ou Annuler pour garder les donnees.",
            icon="warning",
        )
        if not confirmed:
            return
        self.app.db.reset_all_data()
        messagebox.showinfo("Remise a zero", "Toutes les donnees locales du logiciel ont ete supprimees. Velora repart maintenant de zero.")
        self.app.apply_theme("Clair", keep_page="company")

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
                "ui_theme": self.document_fields["ui_theme"].get().strip() or "Clair",
            }
            validate_tax_rate(float(preferences["default_tax_rate"]))
            if preferences["invoice_due_days"] < 0 or preferences["quote_validity_days"] < 0:
                raise ValueError("Les delais doivent etre positifs.")
            if preferences["ui_theme"] not in APP_THEMES:
                raise ValueError("Le theme choisi est invalide.")
        except Exception as exc:
            messagebox.showerror("Parametres", f"Impossible d'enregistrer les parametres.\n\n{exc}")
            return

        theme_changed = self.app.current_theme != preferences["ui_theme"]
        self.app.db.save_company_profile(profile)
        self.app.db.save_document_preferences(preferences)
        messagebox.showinfo("Parametres", "Les informations entreprise et les reglages de documents ont bien ete enregistres.")
        if theme_changed:
            self.app.apply_theme(preferences["ui_theme"], keep_page="company")
        else:
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

    def animate_in(self) -> None:
        pulse_panel(self.form_card)
        self.after(70, lambda: pulse_panel(self.info_card))

    def scroll_units(self, units: int) -> None:
        self.scroll_shell.scroll_units(units)

    def scroll_pages(self, pages: int) -> None:
        self.scroll_shell.scroll_pages(pages)


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
        style_text_widget(self.details_text)

        actions = ttk.Frame(self.form_card, style="Surface.TFrame")
        actions.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        RoundButton(actions, text="Ajouter", style_name="Accent.TButton", command=self.save_todo).pack(side="left")
        RoundButton(actions, text="Mettre a jour", style_name="Ghost.TButton", command=self.update_todo).pack(side="left", padx=8)
        RoundButton(actions, text="Vider", style_name="Ghost.TButton", command=self.clear_form).pack(side="left")
        self.todo_info_label = ttk.Label(self.form_card, text="", style="MutedSurface.TLabel", wraplength=260)
        self.todo_info_label.grid(row=12, column=0, sticky="w", pady=(12, 0))
        self.form_card.columnconfigure(0, weight=1)

    def _build_table(self) -> None:
        top = ttk.Frame(self.table_card, style="Surface.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="Taches", style="SectionTitle.TLabel").pack(side="left")
        RoundButton(top, text="Supprimer", style_name="Danger.TButton", command=self.delete_selected).pack(side="right")

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
        self.scroll_shell = ScrollableFrame(self, COLOR_BG)
        self.scroll_shell.pack(fill="both", expand=True)

        header = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        header.pack(fill="x", padx=24, pady=(24, 12))
        ttk.Label(header, text="Calendrier sante", style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="Visualisez chaque jour le CA confirme, la paie, les depenses, les echeances et les taches.", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        controls = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        controls.pack(fill="x", padx=24)
        RoundButton(controls, text="Mois precedent", style_name="Ghost.TButton", command=self.previous_month).pack(side="left")
        RoundButton(controls, text="Aujourd'hui", style_name="Ghost.TButton", command=self.current_month_view).pack(side="left", padx=8)
        RoundButton(controls, text="Mois suivant", style_name="Ghost.TButton", command=self.next_month).pack(side="left")
        self.month_var = tk.StringVar()
        ttk.Label(controls, textvariable=self.month_var, style="PageTitle.TLabel").pack(side="right")

        self.summary = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        self.summary.pack(fill="x", padx=24, pady=18)
        self.summary.columnconfigure((0, 1, 2, 3), weight=1)
        self.month_revenue_card = MetricCard(self.summary, "Recettes du mois", COLOR_TEAL)
        self.month_expense_card = MetricCard(self.summary, "Depenses du mois", COLOR_CORAL)
        self.month_balance_card = MetricCard(self.summary, "Solde du mois", COLOR_GOLD)
        self.month_todo_card = MetricCard(self.summary, "Todos du mois", COLOR_NAVY)
        self.summary_cards = [self.month_revenue_card, self.month_expense_card, self.month_balance_card, self.month_todo_card]
        for index, widget in enumerate(self.summary_cards):
            widget.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 12, 0))

        body = ttk.Frame(self.scroll_shell.content, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        self.calendar_card = ttk.Frame(body, style="Surface.TFrame", padding=20)
        self.calendar_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        ttk.Label(self.calendar_card, text="Vue mensuelle", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(self.calendar_card, text="Vert: positif, corail: depenses dominantes, sable: jour avec taches.", style="MutedSurface.TLabel").pack(anchor="w", pady=(2, 14))
        self.calendar_grid = tk.Frame(self.calendar_card, bg=COLOR_SURFACE)
        self.calendar_grid.pack(fill="both", expand=True)

        self.detail_card = ttk.Frame(body, style="Surface.TFrame", padding=20)
        self.detail_card.grid(row=0, column=1, sticky="nsew")
        ttk.Label(self.detail_card, text="Detail du jour", style="SectionTitle.TLabel").pack(anchor="w")
        self.selected_day_var = tk.StringVar()
        ttk.Label(self.detail_card, textvariable=self.selected_day_var, style="PageTitleSurface.TLabel").pack(anchor="w", pady=(6, 6))
        self.day_state_var = tk.StringVar()
        ttk.Label(self.detail_card, textvariable=self.day_state_var, style="MutedSurface.TLabel").pack(anchor="w", pady=(0, 12))
        self.detail_text = tk.Text(self.detail_card, bg=COLOR_SURFACE, bd=0, fg=COLOR_TEXT, height=30, font=(FONT_BODY, 10), wrap="word", highlightthickness=0)
        self.detail_text.pack(fill="both", expand=True)
        style_text_widget(self.detail_text, COLOR_SURFACE)
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
        self.month_revenue_card.set(money(self.snapshot["month_revenue"]), f"Factures confirmees: {money(self.snapshot['invoice_revenue'])}")
        self.month_expense_card.set(money(self.snapshot["month_expenses"]), f"Dont paie {money(self.snapshot['employee_expenses'])}")
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
            header = tk.Label(self.calendar_grid, text=label, bg=COLOR_SURFACE, fg=COLOR_TEXT_MUTED, font=(FONT_BODY_SEMIBOLD, 10))
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
                button = RoundButton(
                    self.calendar_grid,
                    text="\n".join(lines),
                    command=lambda value=day_key: self.show_day(value),
                    anchor="w",
                    radius=18,
                    min_height=94,
                    canvas_bg=COLOR_BG,
                    fill_color=background,
                    hover_fill=background,
                    pressed_fill=background,
                    text_color=COLOR_TEXT,
                    border_color=COLOR_TEAL if day_key == self.selected_date else COLOR_BORDER,
                    font_spec=(FONT_BODY, 9),
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
                elif item["kind"] == "invoice_draft":
                    lines.append(f"[Brouillon] {item['label']} - {money(item['amount'])} - {item['status']}")
                elif item["kind"] == "employee_expense":
                    lines.append(f"[Paie] {item['label']} - {money(item['amount'])} - {item['status']}")
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

    def animate_in(self) -> None:
        for index, card in enumerate(self.summary_cards):
            self.after(index * 60, card.pulse)
        self.after(70, lambda: pulse_panel(self.calendar_card))
        self.after(140, lambda: pulse_panel(self.detail_card))
        self.render_calendar()

    def scroll_units(self, units: int) -> None:
        self.scroll_shell.scroll_units(units)

    def scroll_pages(self, pages: int) -> None:
        self.scroll_shell.scroll_pages(pages)


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

        self.scroll_shell = ScrollableFrame(self, COLOR_BG)
        self.scroll_shell.pack(fill="both", expand=True)
        shell = ttk.Frame(self.scroll_shell.content, style="App.TFrame", padding=20)
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=3)
        shell.columnconfigure(1, weight=2)
        shell.rowconfigure(1, weight=1)

        self.left_panel = ttk.Frame(shell, style="Surface.TFrame", padding=20)
        self.left_panel.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        self.right_panel = ttk.Frame(shell, style="Surface.TFrame", padding=20)
        self.right_panel.grid(row=0, column=1, sticky="nsew")
        self.bottom_panel = ttk.Frame(shell, style="Surface.TFrame", padding=20)
        self.bottom_panel.grid(row=1, column=1, sticky="nsew", pady=(12, 0))

        self._build_left(self.left_panel)
        self._build_right(self.right_panel)
        self._build_bottom(self.bottom_panel)
        if self.kind == "invoice":
            self.notes_text.insert("1.0", self.document_preferences.get("default_invoice_notes", ""))
        else:
            self.notes_text.insert("1.0", self.document_preferences.get("default_quote_notes", ""))
        self.update_totals()
        self._bind_shortcuts()
        self.after(40, lambda: pulse_panel(self.left_panel, 20))
        self.after(110, lambda: pulse_panel(self.right_panel, 20))
        self.after(180, lambda: pulse_panel(self.bottom_panel, 20))

    def _bind_shortcuts(self) -> None:
        for sequence, step in [("<Next>", 1), ("<KP_Next>", 1), ("<Prior>", -1), ("<KP_Prior>", -1)]:
            self.bind(sequence, lambda _event, pages=step: self._scroll_pages(pages))
        for sequence, step in [("<KP_Down>", 6), ("<KP_Up>", -6)]:
            self.bind(sequence, lambda _event, units=step: self._scroll_units(units))

    def _scroll_pages(self, pages: int) -> str:
        self.scroll_shell.scroll_pages(pages)
        return "break"

    def _scroll_units(self, units: int) -> str:
        self.scroll_shell.scroll_units(units)
        return "break"

    def _labeled_entry(self, master, label: str, variable: tk.StringVar, row: int, column: int = 0, readonly: bool = False) -> ttk.Entry:
        ttk.Label(master, text=label, style="FieldLabel.TLabel").grid(row=row, column=column, sticky="w", pady=(8, 6))
        entry = ttk.Entry(master, textvariable=variable)
        if readonly:
            entry.state(["readonly"])
        entry.grid(row=row + 1, column=column, sticky="ew", pady=(0, 6), padx=(0, 10))
        return entry

    def _build_left(self, master) -> None:
        ttk.Label(master, text="Informations du document", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(
            master,
            text="Un formulaire plus direct, pense pour aller vite sans perdre le rendu final.",
            style="MutedSurface.TLabel",
            wraplength=460,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 10))
        self._labeled_entry(master, "Numero", self.number_var, 2, 0)
        self._labeled_entry(master, "Date", self.issue_date_var, 2, 1)

        if self.kind == "invoice":
            self._labeled_entry(master, "Echeance", self.period_var, 4, 0)
            ttk.Label(master, text="Statut", style="FieldLabel.TLabel").grid(row=4, column=1, sticky="w", pady=(8, 6))
            ttk.Combobox(master, textvariable=self.status_var, values=["Brouillon", "Envoyee", "Payee", "En retard"], state="readonly").grid(row=5, column=1, sticky="ew", pady=(0, 6))
        else:
            self._labeled_entry(master, "Validite en jours", self.validity_days_var, 4, 0)
            self._labeled_entry(master, "Valide jusqu'au", self.period_var, 4, 1, readonly=True)
            ttk.Label(master, text="Statut", style="FieldLabel.TLabel").grid(row=6, column=0, sticky="w", pady=(8, 6))
            ttk.Combobox(master, textvariable=self.status_var, values=["Brouillon", "Envoye", "Accepte", "Refuse", "Expire"], state="readonly").grid(row=7, column=0, sticky="ew", pady=(0, 6))

        start_row = 6 if self.kind == "invoice" else 8
        self._labeled_entry(master, "Client", self.client_name_var, start_row, 0)
        self._labeled_entry(master, "Email client", self.client_email_var, start_row, 1)
        self._labeled_entry(master, "Adresse client", self.client_address_var, start_row + 2, 0)
        self._labeled_entry(master, "TVA %", self.tax_rate_var, start_row + 2, 1)

        ttk.Label(master, text="Notes", style="FieldLabel.TLabel").grid(row=start_row + 4, column=0, sticky="w", pady=(8, 6))
        self.notes_text = tk.Text(master, height=4, wrap="word", bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER, relief="flat")
        self.notes_text.grid(row=start_row + 5, column=0, columnspan=2, sticky="nsew")
        style_text_widget(self.notes_text)

        ttk.Label(master, text="Lignes", style="SectionTitle.TLabel").grid(row=start_row + 6, column=0, columnspan=2, sticky="w", pady=(18, 8))
        self._labeled_entry(master, "Description", self.description_var, start_row + 7, 0)
        self._labeled_entry(master, "Quantite", self.quantity_var, start_row + 7, 1)
        self._labeled_entry(master, "Prix unitaire HT", self.unit_price_var, start_row + 9, 0)
        RoundButton(master, text="Ajouter la ligne", style_name="Accent.TButton", command=self.add_item).grid(row=start_row + 10, column=1, sticky="ew", pady=(24, 0))

        columns = ("description", "quantity", "price", "total")
        self.items_tree = ttk.Treeview(master, columns=columns, show="headings", height=8)
        for column, heading, width in [("description", "Description", 260), ("quantity", "Qt", 60), ("price", "Prix HT", 100), ("total", "Total HT", 100)]:
            self.items_tree.heading(column, text=heading)
            self.items_tree.column(column, width=width, anchor="w")
        self.items_tree.grid(row=start_row + 11, column=0, columnspan=2, sticky="nsew", pady=(12, 8))
        RoundButton(master, text="Retirer la ligne", style_name="Ghost.TButton", command=self.remove_item).grid(row=start_row + 12, column=0, sticky="w")
        master.columnconfigure((0, 1), weight=1)

    def _build_right(self, master) -> None:
        ttk.Label(master, text="Entreprise emettrice", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(master, text=self.company_profile.get("company_name", ""), style="PageTitleSurface.TLabel").pack(anchor="w", pady=(6, 0))
        identity = ttk.Frame(master, style="SurfaceSoft.TFrame", padding=14)
        identity.pack(fill="x", pady=(12, 12))
        ttk.Label(identity, text="Coordonnees", style="SectionTitleSoft.TLabel").pack(anchor="w")
        info = [
            self.company_profile.get("legal_name", ""),
            self.company_profile.get("address", ""),
            f"SIRET: {self.company_profile.get('siret', '')}",
            f"TVA: {self.company_profile.get('vat_number', '')}",
            f"{self.company_profile.get('email', '')} - {self.company_profile.get('phone', '')}",
        ]
        ttk.Label(identity, text="\n".join([line for line in info if line]), style="MutedSoft.TLabel", wraplength=300).pack(anchor="w", pady=(8, 0))

        impact = ttk.Frame(master, style="SurfaceSoft.TFrame", padding=14)
        impact.pack(fill="x")
        ttk.Label(impact, text="Impact metier", style="SectionTitleSoft.TLabel").pack(anchor="w")
        extra_text = (
            "Cette facture entre ensuite dans le chiffre d'affaires confirme et dans le suivi du calendrier."
            if self.kind == "invoice"
            else "Ce devis reste stocke en local et garde une date de validite modifiable."
        )
        ttk.Label(impact, text=extra_text, style="MutedSoft.TLabel", wraplength=300).pack(anchor="w", pady=(8, 0))

        local = ttk.Frame(master, style="SurfaceSoft.TFrame", padding=14)
        local.pack(fill="x", pady=(12, 0))
        ttk.Label(local, text="Stockage", style="SectionTitleSoft.TLabel").pack(anchor="w")
        ttk.Label(local, text="Le document HTML reste sur cette machine.", style="MutedSoft.TLabel", wraplength=300).pack(anchor="w", pady=(8, 0))

    def _build_bottom(self, master) -> None:
        ttk.Label(master, text="Synthese", style="SectionTitle.TLabel").pack(anchor="w")
        for label, variable in [("Sous-total HT", self.subtotal_var), ("TVA", self.tax_amount_var), ("Total TTC", self.total_var)]:
            line = ttk.Frame(master, style="SurfaceSoft.TFrame", padding=12)
            line.pack(fill="x", pady=6)
            ttk.Label(line, text=label, style="TableStrongSurface.TLabel").pack(side="left")
            ttk.Label(line, textvariable=variable, style="MetricMiniSoft.TLabel").pack(side="right")

        actions = ttk.Frame(master, style="Surface.TFrame")
        actions.pack(fill="x", pady=(16, 0))
        RoundButton(actions, text="Fermer", style_name="Ghost.TButton", command=self.destroy).pack(side="left")
        RoundButton(actions, text="Generer et enregistrer", style_name="Accent.TButton", command=self.save).pack(side="right")

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
        self.current_theme = apply_runtime_palette(self.db.get_document_preferences().get("ui_theme", "Clair"))
        self.title(APP_NAME)
        self.geometry("1500x940")
        self.minsize(1320, 840)
        self.configure(bg=COLOR_BG)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._load_logo()
        self._configure_styles()
        self._build_shell("dashboard")
        self._bind_shortcuts()
        self.refresh_all_pages()

    def _load_logo(self) -> None:
        self.logo_image = None
        png_path = asset_path("velora_logo.png")
        if png_path.exists():
            self.logo_image = tk.PhotoImage(file=str(png_path))
            self.iconphoto(True, self.logo_image)

    def _bind_shortcuts(self) -> None:
        navigation = [
            ("<Control-Key-1>", "dashboard"),
            ("<Control-Key-2>", "calendar"),
            ("<Control-Key-3>", "invoices"),
            ("<Control-Key-4>", "quotes"),
            ("<Control-Key-5>", "sales"),
            ("<Control-Key-6>", "expenses"),
            ("<Control-Key-7>", "payroll"),
            ("<Control-Key-8>", "todos"),
            ("<Control-Key-9>", "company"),
        ]
        for sequence, page_name in navigation:
            self.bind_all(sequence, lambda _event, target=page_name: self._shortcut_show_page(target))

        self.bind_all("<F5>", lambda _event: self._shortcut_refresh())
        self.bind_all("<MouseWheel>", self._shortcut_mousewheel)
        for sequence, step in [("<Next>", 1), ("<KP_Next>", 1), ("<Prior>", -1), ("<KP_Prior>", -1)]:
            self.bind_all(sequence, lambda _event, pages=step: self._shortcut_scroll_pages(pages))
        for sequence, step in [("<KP_Down>", 6), ("<KP_Up>", -6)]:
            self.bind_all(sequence, lambda _event, units=step: self._shortcut_scroll_units(units))

    def _shortcut_show_page(self, page_name: str) -> str:
        self.show_page(page_name)
        return "break"

    def _shortcut_refresh(self) -> str:
        self.refresh_all_pages()
        return "break"

    def _is_descendant(self, widget, ancestor) -> bool:
        current = widget
        while current is not None:
            if current == ancestor:
                return True
            current = getattr(current, "master", None)
        return False

    def _shortcut_mousewheel(self, event) -> str:
        try:
            if event.widget.winfo_toplevel() is not self:
                return ""
        except Exception:
            return ""
        delta = int(getattr(event, "delta", 0))
        units = -int(delta / 120) if delta else 0
        if units == 0:
            units = -1 if delta > 0 else 1
        if hasattr(self, "sidebar") and self._is_descendant(event.widget, self.sidebar):
            if hasattr(self, "sidebar_scroll"):
                self.sidebar_scroll.scroll_units(units * 3)
                return "break"
        return self._shortcut_scroll_units(units * 3)

    def _shortcut_scroll_units(self, units: int) -> str:
        active_page = self.pages.get(self.active_page)
        if active_page and hasattr(active_page, "scroll_units"):
            active_page.scroll_units(units)
            return "break"
        return ""

    def _shortcut_scroll_pages(self, pages: int) -> str:
        active_page = self.pages.get(self.active_page)
        if active_page and hasattr(active_page, "scroll_pages"):
            active_page.scroll_pages(pages)
            return "break"
        return ""

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("App.TFrame", background=COLOR_BG)
        style.configure("Surface.TFrame", background=COLOR_SURFACE, borderwidth=1, relief="flat")
        style.configure("SurfaceSoft.TFrame", background=COLOR_SURFACE_SOFT, borderwidth=0, relief="flat")
        style.configure("CardPanel.TFrame", background=COLOR_SURFACE, relief="flat", borderwidth=1)
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=(FONT_BODY, 10))
        style.configure("CardTitle.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT_MUTED, font=(FONT_BODY_SEMIBOLD, 10))
        style.configure("CardBadge.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEAL, font=(FONT_BODY_SEMIBOLD, 8), padding=(9, 4))
        style.configure("MetricValue.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=(FONT_NUMERIC, 25))
        style.configure("MetricMiniSoft.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, font=(FONT_NUMERIC, 16))
        style.configure("PageTitle.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=(FONT_HEADLINE, 28, "bold"))
        style.configure("PageTitleSurface.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=(FONT_HEADLINE, 24, "bold"))
        style.configure("PageTitleSoft.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, font=(FONT_HEADLINE, 24, "bold"))
        style.configure("SectionTitle.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=(FONT_HEADLINE, 14, "bold"))
        style.configure("SectionTitleSoft.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, font=(FONT_HEADLINE, 13, "bold"))
        style.configure("TableStrong.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=(FONT_BODY_SEMIBOLD, 10))
        style.configure("TableStrongSurface.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, font=(FONT_BODY_SEMIBOLD, 10))
        style.configure("Muted.TLabel", background=COLOR_BG, foreground=COLOR_TEXT_MUTED, font=(FONT_BODY, 10))
        style.configure("MutedSurface.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT_MUTED, font=(FONT_BODY, 10))
        style.configure("MutedSoft.TLabel", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT_MUTED, font=(FONT_BODY, 10))
        style.configure("FieldLabel.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT_MUTED, font=(FONT_BODY_SEMIBOLD, 9))
        style.configure("BadgeSurface.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEAL, font=(FONT_BODY_SEMIBOLD, 9), padding=(8, 3))
        style.configure("TNotebook", background=COLOR_SURFACE, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT_MUTED, padding=(18, 11), borderwidth=0, font=(FONT_BODY_SEMIBOLD, 9))
        style.map("TNotebook.Tab", background=[("selected", COLOR_SURFACE), ("active", COLOR_SURFACE)], foreground=[("selected", COLOR_TEXT), ("active", COLOR_TEXT)])

        for name, color in [("Accent.TButton", COLOR_TEAL), ("AccentAlt.TButton", COLOR_CORAL)]:
            style.configure(name, background=color, foreground="#ffffff", padding=(18, 12), borderwidth=0, focusthickness=0, font=(FONT_BODY_SEMIBOLD, 10))
            style.map(name, background=[("active", COLOR_NAVY), ("pressed", COLOR_NAVY_DEEP)])
        style.configure("Ghost.TButton", background=COLOR_SURFACE_SOFT, foreground=COLOR_TEXT, padding=(16, 10), borderwidth=0, focusthickness=0, font=(FONT_BODY_SEMIBOLD, 10))
        style.map("Ghost.TButton", background=[("active", COLOR_GHOST_HOVER), ("pressed", COLOR_GHOST_PRESSED)])
        style.configure("Danger.TButton", background=COLOR_DANGER, foreground="#ffffff", padding=(14, 10), borderwidth=0, font=(FONT_BODY_SEMIBOLD, 10))
        style.map("Danger.TButton", background=[("active", COLOR_DANGER_HOVER), ("pressed", COLOR_DANGER_PRESSED)])

        style.configure(
            "Treeview",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            fieldbackground=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            borderwidth=0,
            rowheight=40,
            font=(FONT_BODY, 10),
        )
        style.map("Treeview", background=[("selected", COLOR_TREE_SELECTED)], foreground=[("selected", COLOR_TEXT)])
        style.configure("Treeview.Heading", background=COLOR_HEADING_BG, foreground=COLOR_TEXT_MUTED, font=(FONT_BODY_SEMIBOLD, 9), relief="flat", padding=(10, 10))
        style.map("Treeview.Heading", background=[("active", COLOR_HEADING_ACTIVE)])
        style.configure("TEntry", fieldbackground=COLOR_INPUT_BG, bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER, foreground=COLOR_TEXT, padding=(12, 10), relief="flat")
        style.map("TEntry", bordercolor=[("focus", COLOR_TEAL)], lightcolor=[("focus", COLOR_TEAL)], darkcolor=[("focus", COLOR_TEAL)])
        style.configure(
            "TCombobox",
            fieldbackground=COLOR_SURFACE,
            background=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            arrowcolor=COLOR_TEXT_MUTED,
            foreground=COLOR_TEXT,
            padding=(12, 10),
            relief="flat",
            arrowsize=16,
        )
        style.map(
            "TCombobox",
            bordercolor=[("focus", COLOR_TEAL), ("readonly", COLOR_BORDER)],
            lightcolor=[("focus", COLOR_TEAL), ("readonly", COLOR_BORDER)],
            darkcolor=[("focus", COLOR_TEAL), ("readonly", COLOR_BORDER)],
            fieldbackground=[("readonly", COLOR_SURFACE)],
            background=[("readonly", COLOR_SURFACE)],
            foreground=[("readonly", COLOR_TEXT)],
            arrowcolor=[("focus", COLOR_TEAL), ("readonly", COLOR_TEXT_MUTED)],
        )

    def _build_shell(self, initial_page: str = "dashboard") -> None:
        self.active_page = ""
        self.sidebar = tk.Frame(self, bg=COLOR_NAVY_DEEP, width=286)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = tk.Frame(self.sidebar, bg=COLOR_NAVY, bd=0, highlightthickness=1, highlightbackground=COLOR_SIDEBAR_OUTLINE)
        brand.pack(fill="x", padx=20, pady=(24, 18), ipadx=18, ipady=18)
        if self.logo_image is not None:
            icon = self.logo_image.subsample(5, 5)
            self.sidebar_logo = icon
            tk.Label(brand, image=icon, bg=COLOR_NAVY).pack(anchor="w")
        tk.Label(brand, text=APP_NAME, bg=COLOR_NAVY, fg="#ffffff", font=(FONT_HEADLINE, 20, "bold")).pack(anchor="w", pady=(14, 2))
        tk.Label(brand, text=APP_TAGLINE, bg=COLOR_NAVY, fg=COLOR_SIDEBAR_MUTED, wraplength=218, justify="left", font=(FONT_BODY, 9)).pack(anchor="w", pady=(0, 6))
        tk.Label(brand, text="Simple, local, moderne", bg=COLOR_NAVY, fg=COLOR_SIDEBAR_ACCENT, font=(FONT_BODY_SEMIBOLD, 9)).pack(anchor="w")

        nav_shell = tk.Frame(self.sidebar, bg=COLOR_NAVY_DEEP)
        nav_shell.pack(fill="both", expand=True, padx=14)
        self.sidebar_scroll = SidebarScrollArea(nav_shell, width=246)
        self.sidebar_scroll.pack(fill="both", expand=True)
        nav_zone = self.sidebar_scroll.content

        self.nav_buttons: dict[str, SidebarNavButton] = {}
        self.nav_groups: dict[str, dict] = {}
        self.page_to_group: dict[str, str] = {}
        sections = [
            ("pilotage", "Pilotage", [("dashboard", "Tableau de bord"), ("calendar", "Calendrier")]),
            ("documents", "Documents", [("invoices", "Factures"), ("quotes", "Devis")]),
            ("flux", "Flux", [("sales", "Recettes"), ("expenses", "Depenses"), ("payroll", "Paie")]),
            ("organisation", "Organisation", [("todos", "Todo liste"), ("company", "Parametres")]),
        ]

        for group_key, section_title, items in sections:
            group_frame = tk.Frame(nav_zone, bg=COLOR_NAVY_DEEP)
            group_frame.pack(fill="x", pady=(0, 8))
            header_button = SidebarNavButton(
                group_frame,
                text=section_title,
                command=lambda current_group=group_key: self._toggle_nav_group(current_group),
                kind="group",
            )
            header_button.pack(fill="x")
            body_frame = tk.Frame(group_frame, bg=COLOR_NAVY_DEEP)
            self.nav_groups[group_key] = {
                "button": header_button,
                "body": body_frame,
                "label": section_title,
                "expanded": False,
            }

            for key, label in items:
                button = SidebarNavButton(
                    body_frame,
                    text=label,
                    command=lambda page=key: self.show_page(page),
                    kind="item",
                )
                button.pack(fill="x", padx=0, pady=4)
                self.nav_buttons[key] = button
                self.page_to_group[key] = group_key

        footer = tk.Frame(self.sidebar, bg=COLOR_NAVY, bd=0, highlightthickness=1, highlightbackground=COLOR_SIDEBAR_OUTLINE)
        footer.pack(side="bottom", fill="x", padx=20, pady=20, ipadx=18, ipady=16)
        tk.Label(footer, text="Local", bg=COLOR_NAVY, fg="#ffffff", font=(FONT_BODY_SEMIBOLD, 10)).pack(anchor="w")
        tk.Label(footer, text="100% sur cette machine", bg=COLOR_NAVY, fg=COLOR_SIDEBAR_ACCENT, font=(FONT_BODY_SEMIBOLD, 9)).pack(anchor="w", pady=(4, 8))
        tk.Label(footer, text=str(storage_root()), bg=COLOR_NAVY, fg=COLOR_SIDEBAR_MUTED, wraplength=220, justify="left", font=(FONT_BODY, 8)).pack(anchor="w")

        self.main_area = ttk.Frame(self, style="App.TFrame")
        self.main_area.pack(side="left", fill="both", expand=True)

        self.pages = {
            "dashboard": DashboardPage(self.main_area, self),
            "calendar": CalendarPage(self.main_area, self),
            "invoices": DocumentPage(self.main_area, self, "Factures", "Creation, suivi et ouverture locale de vos factures clients.", "invoice"),
            "quotes": DocumentPage(self.main_area, self, "Devis", "Generateur de devis avec validite modifiable et suivi d'acceptation.", "quote"),
            "sales": SalesPage(self.main_area, self),
            "expenses": ExpensesPage(self.main_area, self),
            "payroll": EmployeeExpensesPage(self.main_area, self),
            "todos": TodoPage(self.main_area, self),
            "company": CompanyPage(self.main_area, self),
        }
        for page in self.pages.values():
            page.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.show_page(initial_page)

    def show_page(self, name: str) -> None:
        self.active_page = name
        active_group = self.page_to_group.get(name)
        if active_group:
            self._expand_only_group(active_group)
        for key, page in self.pages.items():
            if key == name:
                page.lift()
            self._set_nav_button_state(key, "default")
        active_page = self.pages.get(name)
        if active_page and hasattr(active_page, "animate_in"):
            active_page.animate_in()

    def _set_nav_button_state(self, key: str, mode: str) -> None:
        button = self.nav_buttons[key]
        button.set_state(active=(key == self.active_page), expanded=False)

    def _toggle_nav_group(self, group_key: str) -> None:
        group = self.nav_groups[group_key]
        if group["expanded"] and self.page_to_group.get(self.active_page) != group_key:
            group["expanded"] = False
            self._apply_nav_group_layout()
            return
        self._expand_only_group(group_key)

    def _expand_only_group(self, group_key: str) -> None:
        for current_key, group in self.nav_groups.items():
            group["expanded"] = current_key == group_key
        self._apply_nav_group_layout()

    def _apply_nav_group_layout(self) -> None:
        for group_key, group in self.nav_groups.items():
            if group["expanded"]:
                if not group["body"].winfo_manager():
                    group["body"].pack(fill="x", pady=(8, 0))
            else:
                if group["body"].winfo_manager():
                    group["body"].pack_forget()
            self._set_nav_group_state(group_key)
        if hasattr(self, "sidebar_scroll"):
            self.sidebar_scroll.canvas.configure(scrollregion=self.sidebar_scroll.canvas.bbox("all"))

    def _set_nav_group_state(self, group_key: str) -> None:
        group = self.nav_groups[group_key]
        active = self.page_to_group.get(self.active_page) == group_key
        group["button"].set_state(active=active, expanded=group["expanded"])

    def apply_theme(self, theme_name: str, keep_page: str | None = None) -> None:
        self.current_theme = apply_runtime_palette(theme_name)
        current_page = keep_page or self.active_page or "dashboard"
        self.configure(bg=COLOR_BG)
        if hasattr(self, "sidebar"):
            self.sidebar.destroy()
        if hasattr(self, "main_area"):
            self.main_area.destroy()
        self._configure_styles()
        self._build_shell(current_page)
        self.refresh_all_pages()

    def refresh_all_pages(self) -> None:
        for page in self.pages.values():
            if hasattr(page, "refresh"):
                page.refresh()

    def on_close(self) -> None:
        self.db.close()
        self.destroy()
