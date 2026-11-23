#!/usr/bin/env python3
"""
i18n_ar.py — Single-source Arabic UI string loader for Smart_Local_Agent_Free.

Reads all user-facing strings from `strings_ar.json` and supports two styles:
  - `msa` (default): Simplified Modern Standard Arabic (فصحى مبسطة)
  - `eg`:            Egyptian Colloquial Arabic (عامية مصرية)
"""
from __future__ import annotations

import json
import os
import re
import threading
from pathlib import Path
from typing import Any

STRINGS_PATH = Path(__file__).resolve().parent / "strings_ar.json"
DEFAULT_STYLE = "msa"
VALID_STYLES = ("msa", "eg")

_PLACEHOLDER_RE = re.compile(r"\{([a-zA-Z0-9_]+)\}")
_LOCK = threading.RLock()
_CACHE_MTIME: float = -1.0
_CACHE_RAW: dict[str, Any] = {}
_CACHE_CLIENT: dict[str, Any] = {}
_ACTIVE_STYLE: str = DEFAULT_STYLE


def normalize_style(style: str | None) -> str:
    s = str(style or "").strip().lower()
    return s if s in VALID_STYLES else DEFAULT_STYLE


def get_active_style() -> str:
    env_style = os.environ.get("SMART_AGENT_UI_STYLE", "").strip().lower()
    if env_style in VALID_STYLES:
        return env_style
    with _LOCK:
        return _ACTIVE_STYLE


def set_active_style(style: str | None) -> str:
    global _ACTIVE_STYLE
    norm = normalize_style(style)
    with _LOCK:
        _ACTIVE_STYLE = norm
    return norm


def _load_locked() -> dict[str, Any]:
    global _CACHE_MTIME, _CACHE_RAW, _CACHE_CLIENT
    try:
        st = STRINGS_PATH.stat()
        mtime = st.st_mtime
    except OSError:
        return _CACHE_RAW
    if _CACHE_RAW and mtime == _CACHE_MTIME:
        return _CACHE_RAW
    raw = json.loads(STRINGS_PATH.read_text(encoding="utf-8"))
    strings = raw.get("strings") if isinstance(raw.get("strings"), dict) else {}
    compact: dict[str, dict[str, str]] = {}
    for k, v in strings.items():
        if k.startswith("_") or not isinstance(v, dict):
            continue
        compact[k] = {
            "msa": str(v.get("msa") or ""),
            "eg": str(v.get("eg") or v.get("msa") or ""),
        }
    meta = raw.get("_meta") if isinstance(raw.get("_meta"), dict) else {}
    _CACHE_RAW = {"_meta": meta, "strings": strings}
    _CACHE_CLIENT = {
        "default_style": DEFAULT_STYLE,
        "styles": list(VALID_STYLES),
        "strings": compact,
    }
    _CACHE_MTIME = mtime
    return _CACHE_RAW


def tr(key: str, style: str | None = None, **kwargs: Any) -> str:
    """Return localized Arabic string for `key` in `style` (default: active style `msa`)."""
    with _LOCK:
        raw = _load_locked()
        strings = raw.get("strings") or {}
        entry = strings.get(key)
    st = normalize_style(style) if style else get_active_style()
    if not isinstance(entry, dict):
        return key
    template = str(entry.get(st) or entry.get(DEFAULT_STYLE) or key)
    if not kwargs:
        return template

    def _sub(m: re.Match[str]) -> str:
        name = m.group(1)
        if name in kwargs and kwargs[name] is not None:
            return str(kwargs[name])
        return m.group(0)

    return _PLACEHOLDER_RE.sub(_sub, template)


def get_client_bundle(active_style: str | None = None) -> dict[str, Any]:
    """Return compact dictionary payload for frontend injection and `/api/i18n`."""
    with _LOCK:
        _load_locked()
        base = dict(_CACHE_CLIENT)
    base["active_style"] = normalize_style(active_style) if active_style else get_active_style()
    return base


def section_label(section_id: str, style: str | None = None) -> str:
    key = f"sec.{section_id}"
    val = tr(key, style=style)
    return val if val != key else tr("sec.general", style=style)


_AUTO_SRC_MAP = {
    "project.name": "autosrc.project_name",
    "project.code": "autosrc.project_code",
    "project.contract_no": "autosrc.project_contract_no",
    "project.owner": "autosrc.project_owner",
    "project.consultant": "autosrc.project_consultant",
    "project.contractor": "autosrc.project_contractor",
    "project.location": "autosrc.project_location",
    "project.engineer": "autosrc.project_engineer",
    "system.next_ref": "autosrc.system_next_ref",
    "system.today_iso": "autosrc.system_today_iso",
    "system.today_dmy": "autosrc.system_today_dmy",
    "system.today_dmy_dash": "autosrc.system_today_dmy_dash",
    "system.month_year": "autosrc.system_month_year",
    "system.day_name_ar": "autosrc.system_day_name_ar",
    "system.revision": "autosrc.system_revision",
    "system.year": "autosrc.system_year",
    "date.today_dmy": "autosrc.system_today_dmy",
    "date.today_iso": "autosrc.system_today_iso",
    "date.today_dmy_dash": "autosrc.system_today_dmy_dash",
    "date.month_year": "autosrc.system_month_year",
    "date.day_name_ar": "autosrc.system_day_name_ar",
    "date.year": "autosrc.system_year",
}


def auto_source_label(source_key: str, style: str | None = None) -> str:
    i18n_key = _AUTO_SRC_MAP.get(source_key)
    return tr(i18n_key, style=style) if i18n_key else source_key
