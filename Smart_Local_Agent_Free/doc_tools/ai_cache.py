#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ai_cache.py — Persistent content-addressed AI result cache for Smart_Local_Agent_Free.

Stores successful AI Vision / LLM results on disk (keyed by file SHA-256 + context
signature) so reopening modals, re-scanning folders, or switching between Scanner /
Inbox / Organizer / WhatsApp / Template Studio reuses prior AI work instantly without
burning API quota or time.
"""

from __future__ import annotations

import copy
import datetime as dt
import hashlib
import json
import re
import threading
import time
from pathlib import Path
from typing import Any

from security_utils import atomic_write_json, read_json_safe

CACHE_SCHEMA_VERSION = 1
DEFAULT_MAX_ENTRIES = 600

_LOCKS_GUARD = threading.Lock()
_FILE_LOCKS: dict[str, threading.RLock] = {}


def _lock_for(path: Path) -> threading.RLock:
    key = str(path.resolve()) if path else str(path)
    with _LOCKS_GUARD:
        lk = _FILE_LOCKS.get(key)
        if lk is None:
            lk = threading.RLock()
            _FILE_LOCKS[key] = lk
        return lk


def _clean_namespace(namespace: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", str(namespace or "default")).strip("._-")
    return cleaned or "default"


def project_ai_cache_dir(proj_root: Path | str) -> Path:
    return Path(proj_root) / "_System_Agent" / "_ai_cache"


def cache_file_path(cache_dir: Path | str, namespace: str) -> Path:
    return Path(cache_dir) / f"{_clean_namespace(namespace)}.json"


def stable_sig(*parts: Any) -> str:
    """Return a deterministic short SHA-256 hex signature for arbitrary JSON-able context."""
    try:
        blob = json.dumps(parts, ensure_ascii=False, sort_keys=True, default=str)
    except Exception:
        blob = "|".join(str(p) for p in parts)
    return hashlib.sha256(blob.encode("utf-8", errors="replace")).hexdigest()[:20]


def _now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def _load_store(path: Path) -> dict[str, Any]:
    raw = read_json_safe(path, {})
    if not isinstance(raw, dict) or not isinstance(raw.get("entries"), dict):
        return {"version": CACHE_SCHEMA_VERSION, "entries": {}}
    return {"version": int(raw.get("version") or CACHE_SCHEMA_VERSION), "entries": dict(raw["entries"])}


def _prune_entries(entries: dict[str, Any], max_entries: int) -> None:
    lim = max(10, int(max_entries or DEFAULT_MAX_ENTRIES))
    if len(entries) <= lim:
        return
    ordered = sorted(
        entries.items(),
        key=lambda kv: float((kv[1] or {}).get("ts") or 0.0),
    )
    excess = len(entries) - lim
    for k, _ in ordered[:excess]:
        entries.pop(k, None)


def get_cached(
    cache_dir: Path | str,
    namespace: str,
    key: str,
    *,
    ttl_seconds: float | None = None,
) -> dict[str, Any] | None:
    """Return a deep copy of a cached payload dict if present and non-expired, else None."""
    if not key:
        return None
    cpath = cache_file_path(cache_dir, namespace)
    if not cpath.exists():
        return None
    with _lock_for(cpath):
        store = _load_store(cpath)
        rec = store["entries"].get(str(key))
        if not isinstance(rec, dict):
            return None
        val = rec.get("value")
        if not isinstance(val, dict) or not val:
            return None
        if ttl_seconds is not None and ttl_seconds > 0:
            ts = float(rec.get("ts") or 0.0)
            if ts <= 0 or (time.time() - ts) > float(ttl_seconds):
                return None
        out = copy.deepcopy(val)
        out["_cached_at"] = str(rec.get("cached_at") or "")
        return out


def set_cached(
    cache_dir: Path | str,
    namespace: str,
    key: str,
    value: dict[str, Any],
    *,
    max_entries: int = DEFAULT_MAX_ENTRIES,
) -> bool:
    """Persist a non-empty valid AI result dict atomically under `namespace` + `key`."""
    if not key or not isinstance(value, dict) or not value:
        return False
    # Never cache errors, empty results, or unreadable scan fallbacks
    if value.get("error") or value.get("is_blank_or_unreadable"):
        return False

    cpath = cache_file_path(cache_dir, namespace)
    clean_val = {k: v for k, v in copy.deepcopy(value).items() if not str(k).startswith("_cached_")}
    if not clean_val:
        return False

    with _lock_for(cpath):
        cpath.parent.mkdir(parents=True, exist_ok=True)
        store = _load_store(cpath)
        store["entries"][str(key)] = {
            "ts": time.time(),
            "cached_at": _now_iso(),
            "value": clean_val,
        }
        _prune_entries(store["entries"], max_entries)
        atomic_write_json(cpath, store)
    return True


def delete_cached(cache_dir: Path | str, namespace: str, key: str) -> bool:
    cpath = cache_file_path(cache_dir, namespace)
    if not cpath.exists():
        return False
    with _lock_for(cpath):
        store = _load_store(cpath)
        if str(key) not in store["entries"]:
            return False
        store["entries"].pop(str(key), None)
        atomic_write_json(cpath, store)
        return True


def clear_cache(cache_dir: Path | str, namespace: str | None = None) -> int:
    cdir = Path(cache_dir)
    if not cdir.exists() or not cdir.is_dir():
        return 0
    removed = 0
    if namespace:
        targets = [cache_file_path(cdir, namespace)]
    else:
        targets = [p for p in cdir.glob("*.json") if p.is_file()]
    for tp in targets:
        if not tp.exists():
            continue
        with _lock_for(tp):
            store = _load_store(tp)
            removed += len(store.get("entries") or {})
            try:
                tp.unlink(missing_ok=True)
            except OSError:
                atomic_write_json(tp, {"version": CACHE_SCHEMA_VERSION, "entries": {}})
    return removed


def cache_stats(cache_dir: Path | str) -> dict[str, Any]:
    cdir = Path(cache_dir)
    if not cdir.exists() or not cdir.is_dir():
        return {"total_entries": 0, "namespaces": {}}
    by_ns: dict[str, int] = {}
    total = 0
    for tp in sorted(cdir.glob("*.json")):
        with _lock_for(tp):
            store = _load_store(tp)
            cnt = len(store.get("entries") or {})
            by_ns[tp.stem] = cnt
            total += cnt
    return {"total_entries": total, "namespaces": by_ns}


def cached_sizes(cache_dir: Path | str, namespace: str) -> set[int]:
    """Return all recorded `size_bytes` values in `namespace` for O(1) pre-filtering before hashing."""
    cpath = cache_file_path(cache_dir, namespace)
    if not cpath.exists():
        return set()
    sizes: set[int] = set()
    with _lock_for(cpath):
        store = _load_store(cpath)
        for rec in (store.get("entries") or {}).values():
            if not isinstance(rec, dict):
                continue
            val = rec.get("value")
            if isinstance(val, dict) and val.get("size_bytes") is not None:
                try:
                    sizes.add(int(val["size_bytes"]))
                except (TypeError, ValueError):
                    pass
    return sizes
