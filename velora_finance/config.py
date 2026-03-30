from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "Velora Finance"
APP_SLUG = "velora-finance"
APP_TAGLINE = "Pilotage local des finances pour les entreprises"

APP_THEMES = ("Clair", "Sombre")

THEME_PALETTES = {
    "Clair": {
        "navy": "#0f2340",
        "navy_deep": "#0a1628",
        "teal": "#2563eb",
        "coral": "#ef4444",
        "gold": "#f97316",
        "bg": "#f8fafc",
        "surface": "#ffffff",
        "surface_soft": "#f1f5f9",
        "text": "#1e293b",
        "text_muted": "#64748b",
        "border": "#e2e8f0",
        "success": "#10b981",
        "warning": "#ea580c",
        "danger": "#dc2626",
        "health_good": "#dcfce7",
        "health_bad": "#fee2e2",
        "health_todo": "#ffedd5",
        "input_bg": "#ffffff",
        "text_select_bg": "#dbeafe",
        "chart_grid": "#e2e8f0",
        "bar_track": "#e2e8f0",
        "ghost_hover": "#e2e8f0",
        "ghost_pressed": "#cbd5e1",
        "danger_hover": "#b91c1c",
        "danger_pressed": "#991b1b",
        "tree_selected": "#dbeafe",
        "heading_bg": "#f8fafc",
        "heading_active": "#eff6ff",
        "sidebar_outline": "#1a365d",
        "sidebar_text": "#e2e8f0",
        "sidebar_muted": "#94a3b8",
        "sidebar_caption": "#64748b",
        "sidebar_accent": "#3b82f6",
    },
    "Sombre": {
        "navy": "#0a1628",
        "navy_deep": "#081120",
        "teal": "#3b82f6",
        "coral": "#ef4444",
        "gold": "#f97316",
        "bg": "#0f172a",
        "surface": "#111827",
        "surface_soft": "#1e293b",
        "text": "#f8fafc",
        "text_muted": "#94a3b8",
        "border": "#334155",
        "success": "#10b981",
        "warning": "#f97316",
        "danger": "#ef4444",
        "health_good": "#052e16",
        "health_bad": "#450a0a",
        "health_todo": "#422006",
        "input_bg": "#0f172a",
        "text_select_bg": "#1d4ed8",
        "chart_grid": "#334155",
        "bar_track": "#1f2937",
        "ghost_hover": "#243041",
        "ghost_pressed": "#1b2432",
        "danger_hover": "#dc2626",
        "danger_pressed": "#b91c1c",
        "tree_selected": "#1e3a8a",
        "heading_bg": "#0f172a",
        "heading_active": "#1e293b",
        "sidebar_outline": "#1a365d",
        "sidebar_text": "#e2e8f0",
        "sidebar_muted": "#94a3b8",
        "sidebar_caption": "#64748b",
        "sidebar_accent": "#3b82f6",
    },
}


def get_theme_palette(theme_name: str) -> dict[str, str]:
    cleaned = str(theme_name or "Clair").strip().capitalize()
    if cleaned not in THEME_PALETTES:
        cleaned = "Clair"
    return THEME_PALETTES[cleaned].copy()


_default_palette = get_theme_palette("Clair")

COLOR_NAVY = _default_palette["navy"]
COLOR_NAVY_DEEP = _default_palette["navy_deep"]
COLOR_TEAL = _default_palette["teal"]
COLOR_CORAL = _default_palette["coral"]
COLOR_GOLD = _default_palette["gold"]
COLOR_BG = _default_palette["bg"]
COLOR_SURFACE = _default_palette["surface"]
COLOR_SURFACE_SOFT = _default_palette["surface_soft"]
COLOR_TEXT = _default_palette["text"]
COLOR_TEXT_MUTED = _default_palette["text_muted"]
COLOR_BORDER = _default_palette["border"]
COLOR_SUCCESS = _default_palette["success"]
COLOR_WARNING = _default_palette["warning"]
COLOR_DANGER = _default_palette["danger"]


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent.parent


def asset_path(name: str) -> Path:
    return project_root() / "assets" / name


def storage_root() -> Path:
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path.home() / ".local" / "share"
    root = base / APP_NAME
    (root / "documents" / "factures").mkdir(parents=True, exist_ok=True)
    (root / "documents" / "devis").mkdir(parents=True, exist_ok=True)
    (root / "documents" / "depenses").mkdir(parents=True, exist_ok=True)
    return root
