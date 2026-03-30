from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "Velora Finance"
APP_SLUG = "velora-finance"
APP_TAGLINE = "Pilotage local des finances pour les entreprises"

COLOR_NAVY = "#15304b"
COLOR_NAVY_DEEP = "#10263c"
COLOR_TEAL = "#1f8a70"
COLOR_CORAL = "#ff8b5e"
COLOR_GOLD = "#f2c572"
COLOR_BG = "#f3f5f9"
COLOR_SURFACE = "#ffffff"
COLOR_SURFACE_SOFT = "#eef2f7"
COLOR_TEXT = "#17324d"
COLOR_TEXT_MUTED = "#6b7a8c"
COLOR_BORDER = "#d7dfeb"
COLOR_SUCCESS = "#1d9a6c"
COLOR_WARNING = "#d98c2f"
COLOR_DANGER = "#d25555"


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
    return root
