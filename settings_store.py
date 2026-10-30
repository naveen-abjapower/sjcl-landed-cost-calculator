"""
Tiny local persistence layer for the manually-entered settings (Andhra
Pradesh Energy Cost, CSS rate, trading margin) so they don't reset to
zero every time the app is restarted.

Stores a plain JSON file next to app.py. No database, no external
service - just a file on your own machine.
"""

from __future__ import annotations

import json
import os

SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

DEFAULTS = {
    "margin_paisa": 2.0,
    "ap_peak": 0.0,
    "ap_normal": 0.0,
    "ap_offpeak": 0.0,
    "css_rate": 0.0,
}


def is_streamlit_cloud() -> bool:
    return os.environ.get("IS_STREAMLIT_CLOUD") == "true"


def load_settings() -> dict:
    if is_streamlit_cloud():
        return dict(DEFAULTS)
    if not os.path.exists(SETTINGS_PATH):
        return dict(DEFAULTS)
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        merged = dict(DEFAULTS)
        merged.update({k: v for k, v in data.items() if k in DEFAULTS})
        return merged
    except Exception:
        # Corrupt/unreadable settings file - fall back to defaults rather
        # than crashing the app.
        return dict(DEFAULTS)


def save_settings(values: dict) -> None:
    if is_streamlit_cloud():
        return
    current = load_settings()
    current.update({k: v for k, v in values.items() if k in DEFAULTS})
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(current, f, indent=2)
    except Exception:
        pass  # best-effort; never crash the app over a failed save
