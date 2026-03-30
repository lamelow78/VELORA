from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

APP_NAME = "Noryven"
APP_SLUG = "noryven"
APP_TAGLINE = "Gestion financiere locale"
LEGACY_APP_NAMES = ("Velora Finance",)

APP_THEMES = ("Clair", "Sombre")

DOCUMENT_FOLDERS = {
    "invoice": "factures",
    "quote": "devis",
    "expense": "depenses",
    "employee_expense": "paie",
    "sale": "recettes",
    "other": "pieces-diverses",
}

THEME_PALETTES = {
    "Clair": {
        "navy": "#10233d",
        "navy_deep": "#0a1728",
        "teal": "#246bff",
        "coral": "#0ea5a3",
        "gold": "#14b8a6",
        "bg": "#f4f7fb",
        "surface": "#ffffff",
        "surface_soft": "#eef4fb",
        "text": "#132238",
        "text_muted": "#5b6b82",
        "border": "#d9e3ef",
        "success": "#0f9f6e",
        "warning": "#d97706",
        "danger": "#dc2626",
        "health_good": "#ddfaee",
        "health_bad": "#fee7e7",
        "health_todo": "#e7f4ff",
        "input_bg": "#ffffff",
        "text_select_bg": "#d9ebff",
        "chart_grid": "#dce8f5",
        "bar_track": "#dce7f3",
        "ghost_hover": "#e5eef8",
        "ghost_pressed": "#d8e3f0",
        "danger_hover": "#c52121",
        "danger_pressed": "#a91a1a",
        "tree_selected": "#dcebff",
        "heading_bg": "#f0f5fb",
        "heading_active": "#e7f1fb",
        "sidebar_outline": "#1d3a5c",
        "sidebar_text": "#eef5ff",
        "sidebar_muted": "#b1c6e0",
        "sidebar_caption": "#8aa3c0",
        "sidebar_accent": "#5dc4ff",
    },
    "Sombre": {
        "navy": "#0b1626",
        "navy_deep": "#08111e",
        "teal": "#5b8cff",
        "coral": "#22c7be",
        "gold": "#2dd4bf",
        "bg": "#08111d",
        "surface": "#0f1b2c",
        "surface_soft": "#122338",
        "text": "#ecf4ff",
        "text_muted": "#9fb2c9",
        "border": "#21364f",
        "success": "#22c55e",
        "warning": "#f59e0b",
        "danger": "#f87171",
        "health_good": "#0c2d23",
        "health_bad": "#3c171c",
        "health_todo": "#14293f",
        "input_bg": "#091425",
        "text_select_bg": "#214b83",
        "chart_grid": "#22344b",
        "bar_track": "#17283b",
        "ghost_hover": "#17304a",
        "ghost_pressed": "#15263b",
        "danger_hover": "#ef5f5f",
        "danger_pressed": "#db4e4e",
        "tree_selected": "#173d6f",
        "heading_bg": "#0d1726",
        "heading_active": "#17304a",
        "sidebar_outline": "#1d3a5c",
        "sidebar_text": "#e6f1ff",
        "sidebar_muted": "#98afd0",
        "sidebar_caption": "#7393bd",
        "sidebar_accent": "#5dc4ff",
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


def _base_storage_dir() -> Path:
    if sys.platform.startswith("win"):
        return Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support"
    return Path.home() / ".local" / "share"


def _migrate_legacy_storage(target: Path) -> None:
    if target.exists():
        return
    base = _base_storage_dir()
    for legacy_name in LEGACY_APP_NAMES:
        legacy_root = base / legacy_name
        if legacy_root.exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copytree(legacy_root, target, dirs_exist_ok=True)
            return


def storage_root() -> Path:
    base = _base_storage_dir()
    root = base / APP_NAME
    _migrate_legacy_storage(root)
    root.mkdir(parents=True, exist_ok=True)
    return root


def app_state_path() -> Path:
    return storage_root() / "app_state.json"


def _load_app_state() -> dict:
    target = app_state_path()
    if not target.exists():
        return {}
    try:
        payload = json.loads(target.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def _save_app_state(payload: dict) -> None:
    app_state_path().write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def legacy_documents_root() -> Path:
    return storage_root() / "documents"


def default_documents_root() -> Path:
    return legacy_documents_root()


def preferred_documents_root() -> Path:
    configured = _load_app_state().get("documents_root")
    if configured:
        return Path(configured).expanduser()
    return default_documents_root()


def ensure_documents_root(root: str | Path | None = None) -> Path:
    target = Path(root).expanduser() if root else preferred_documents_root()
    resolved = target.resolve()
    resolved.mkdir(parents=True, exist_ok=True)
    for folder in DOCUMENT_FOLDERS.values():
        (resolved / folder).mkdir(parents=True, exist_ok=True)
    return resolved


def documents_root() -> Path:
    return ensure_documents_root(preferred_documents_root())


def set_documents_root(root: str | Path) -> Path:
    resolved = ensure_documents_root(root)
    payload = _load_app_state()
    payload["documents_root"] = str(resolved)
    _save_app_state(payload)
    return resolved


def document_directory(kind: str, root: str | Path | None = None) -> Path:
    folder = DOCUMENT_FOLDERS.get(kind)
    if folder is None:
        raise ValueError("Type de document inconnu.")
    return ensure_documents_root(root) / folder


def available_document_directories(root: str | Path | None = None) -> dict[str, Path]:
    base = ensure_documents_root(root)
    return {kind: base / folder for kind, folder in DOCUMENT_FOLDERS.items()}
