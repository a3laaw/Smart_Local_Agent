#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
security_utils.py — Core local security, path-traversal protection, zero-delete archival,
secret scrubbing, atomic JSON I/O, and per-device identity (Smart_Local_Agent_Free).
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import secrets
import shutil
import socket
import tempfile
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from i18n_ar import tr

DANGEROUS_EXTENSIONS: frozenset[str] = frozenset({
    ".exe", ".dll", ".bat", ".cmd", ".ps1", ".vbs", ".vbe", ".js", ".jse",
    ".wsf", ".wsh", ".msi", ".msp", ".com", ".scr", ".pif", ".hta", ".cpl",
    ".reg", ".inf", ".sys", ".drv", ".lnk", ".url", ".sh", ".command", ".app",
})

SYSTEM_TEMP_FILENAMES: frozenset[str] = frozenset({
    "thumbs.db", "desktop.ini", ".ds_store", "ehthumbs.db", "icon\r",
    "_rename_journal.json", "_archive_journal.json", "pending_log_ops.json",
    "pending_excel_rows.json", "watcher_state.json", "watcher_config.json",
})

SECRET_KEY_PATTERNS: tuple[str, ...] = (
    "api_key", "apikey", "gemini_api_key", "secret", "password", "passwd",
    "access_token", "refresh_token", "auth_token", "private_key", "authorization",
    "verify_token", "app_secret", "wa_token",
)

SECRET_VALUE_RE = re.compile(
    r"(?:AIza[0-9A-Za-z_\-]{25,}|sk-[0-9A-Za-z_\-]{20,}|ghp_[0-9A-Za-z]{25,}|EAAG[0-9A-Za-z]{20,})"
)


class PathTraversalError(ValueError):
    """Raised when a path escapes its permitted root directory."""


class UnsafeFileError(ValueError):
    """Raised when an upload or file target violates extension or name safety rules."""


def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def sha256_of_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def scrub_secrets(obj: Any) -> Any:
    """Return a deep copy with any secret keys or raw API key strings stripped."""
    if isinstance(obj, dict):
        out: dict[str, Any] = {}
        for k, v in obj.items():
            kl = str(k).strip().lower()
            if any(pat in kl for pat in SECRET_KEY_PATTERNS):
                continue
            out[k] = scrub_secrets(v)
        return out
    if isinstance(obj, list):
        return [scrub_secrets(x) for x in obj]
    if isinstance(obj, tuple):
        return [scrub_secrets(x) for x in obj]
    if isinstance(obj, str):
        if SECRET_VALUE_RE.search(obj):
            return SECRET_VALUE_RE.sub("[REDACTED]", obj)
        return obj
    return obj


def contains_secret(text: str) -> bool:
    if not text:
        return False
    if SECRET_VALUE_RE.search(text):
        return True
    low = text.lower()
    for pat in ("gemini_api_key", "access_token", "app_secret", "private_key"):
        m = re.search(rf'"{re.escape(pat)}"\s*:\s*"([^"]+)"', low)
        if m and m.group(1).strip() and m.group(1).strip() != "[redacted]":
            return True
    return False


def atomic_write_json(path: Path, data: Any, *, scrub: bool = True) -> None:
    """Atomically write UTF-8 JSON to disk, optionally scrubbing secret fields."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = scrub_secrets(data) if scrub else data
    encoded = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    fd, tmp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent))
    tmp_path = Path(tmp_name)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(encoded)
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
        os.replace(tmp_path, path)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def read_json_safe(path: Path, default: Any = None) -> Any:
    p = Path(path)
    if not p.exists():
        return default
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return default


def get_or_create_device_id(state_dir: Path) -> str:
    """
    Return a stable, non-secret local device identifier for audit trails across
    multi-PC setups (e.g. PC-A with 1 project, PC-B with 2 projects).
    """
    state_dir = Path(state_dir)
    state_dir.mkdir(parents=True, exist_ok=True)
    id_file = state_dir / "device_identity.json"
    existing = read_json_safe(id_file, {})
    if isinstance(existing, dict):
        dev = str(existing.get("device_id") or "").strip()
        if dev and re.fullmatch(r"[A-Za-z0-9_\-]{4,64}", dev):
            return dev
    raw_host = ""
    try:
        raw_host = socket.gethostname() or ""
    except Exception:
        raw_host = "PC"
    clean_host = re.sub(r"[^A-Za-z0-9]", "", raw_host)[:14].upper() or "PC"
    suffix = secrets.token_hex(3).upper()
    device_id = f"DEV-{clean_host}-{suffix}"
    atomic_write_json(id_file, {
        "device_id": device_id,
        "hostname_label": clean_host,
        "created_at": now_iso(),
    })
    return device_id


def normalize_win_path(path: Path | str) -> Path:
    """
    Normalize Windows extended-length UNC/local paths (`\\\\?\\UNC\\server\\share` ->
    `\\\\server\\share`, `\\\\?\\C:\\...` -> `C:\\...`) so `relative_to()` and path
    comparisons work consistently across Windows Server SMB shares and mapped drives.
    """
    s = str(path)
    if s.startswith("\\\\?\\UNC\\") or s.startswith("//?/UNC/"):
        s = "\\\\" + s[8:]
    elif s.startswith("\\\\?\\") or s.startswith("//?/"):
        s = s[4:]
    return Path(s)


def resolve_local_or_unc(path: Path | str) -> Path:
    """Resolve a local or Windows Server UNC path safely without throwing on network shares."""
    p = normalize_win_path(path)
    try:
        return normalize_win_path(p.resolve())
    except OSError:
        return normalize_win_path(os.path.abspath(str(p)))


def safe_resolve_within(
    base: Path,
    candidate: Path | str,
    *,
    allow_base: bool = False,
    reject_symlinks: bool = True,
) -> Path:
    """
    Safely resolve `candidate` (relative or absolute) and guarantee it sits strictly
    inside `base` (or equals `base` when `allow_base=True`).
    Blocks `../`, `..\\`, null bytes, drive jumps, and symlink escapes, while fully
    supporting Windows Server UNC bases (`\\\\Server\\Share\\...`).
    """
    if candidate is None:
        raise PathTraversalError(tr("server.path_blocked"))
    raw = str(candidate)
    if "\x00" in raw or not raw.strip():
        raise PathTraversalError(tr("server.path_blocked"))

    base_resolved = resolve_local_or_unc(base)

    # Reject Windows-style backslash traversal or unexpected UNC jumps
    norm_slashes = raw.replace("\\", "/")
    if re.match(r"^[A-Za-z]:", norm_slashes) and os.name != "nt":
        raise PathTraversalError(tr("server.path_blocked"))
    cand_norm_win = normalize_win_path(raw)
    if norm_slashes.startswith("//"):
        # Only allow an absolute UNC candidate if its normalized root actually lives within `base_resolved`
        base_slashes = str(base_resolved).replace("\\", "/").rstrip("/")
        cand_slashes = str(cand_norm_win).replace("\\", "/")
        cand_lc = cand_slashes.lower()
        base_lc = base_slashes.lower()
        within_unc = (cand_lc == base_lc) or cand_lc.startswith(base_lc + "/")
        if os.name != "nt" or not within_unc:
            raise PathTraversalError(tr("server.path_blocked"))

    cand_path = Path(norm_slashes) if os.name != "nt" else cand_norm_win
    joined = cand_path if cand_path.is_absolute() else (base_resolved / cand_path)

    if reject_symlinks:
        # Walk existing components from base down to joined to reject symlinks escaping base
        try:
            probe = joined
            checked: list[Path] = []
            while probe != probe.parent and probe != base_resolved:
                checked.append(probe)
                probe = probe.parent
            for item in reversed(checked):
                if item.is_symlink():
                    link_target = resolve_local_or_unc(item)
                    if not link_target.is_relative_to(base_resolved):
                        raise PathTraversalError(tr("server.path_blocked"))
        except PathTraversalError:
            raise
        except OSError:
            pass

    resolved = resolve_local_or_unc(joined)
    if resolved == base_resolved:
        if allow_base:
            return resolved
        raise PathTraversalError(tr("server.path_blocked"))
    if not resolved.is_relative_to(base_resolved):
        raise PathTraversalError(tr("server.path_blocked"))
    return resolved


def sanitize_upload_filename(
    raw_name: str,
    allowed_exts: set[str] | frozenset[str] | None = None,
    max_stem_len: int = 90,
) -> str:
    """
    Validate and sanitize an uploaded filename:
    - strips any directory traversal (`../`, `..\\`, `/`, `\\`)
    - blocks dangerous executable extensions and Office lock files (`~$`)
    - enforces `allowed_exts` when provided
    """
    if not raw_name or "\x00" in str(raw_name):
        raise UnsafeFileError(tr("server.invalid_filename"))
    cleaned = str(raw_name).replace("\\", "/").strip()
    # Explicit rejection if raw name looks purely like traversal
    base_name = cleaned.split("/")[-1].strip()
    if not base_name or base_name in (".", "..") or base_name.startswith("~$"):
        raise UnsafeFileError(tr("server.invalid_filename"))
    if base_name.startswith("."):
        raise UnsafeFileError(tr("server.invalid_filename"))

    p = Path(base_name)
    ext = p.suffix.lower()
    if not ext:
        raise UnsafeFileError(tr("server.invalid_filename"))
    if ext in DANGEROUS_EXTENSIONS:
        raise UnsafeFileError(tr("server.file_type_blocked", ext=ext))
    if allowed_exts is not None:
        norm_allowed = {e.lower() if e.startswith(".") else f".{e.lower()}" for e in allowed_exts}
        if ext not in norm_allowed:
            raise UnsafeFileError(tr("server.file_type_blocked", ext=ext))

    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", p.stem).strip(" ._")
    if not stem or stem.upper() in {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    }:
        stem = f"file_{stem}" if stem else "file"
    stem = stem[:max_stem_len].rstrip(" ._")
    return f"{stem}{ext}"


def safe_template_name(raw_name: str) -> str:
    """Validate a template or log filename coming from a URL path segment."""
    if not raw_name or "\x00" in raw_name:
        raise PathTraversalError(tr("server.invalid_filename"))
    if "/" in raw_name or "\\" in raw_name or ".." in raw_name:
        raise PathTraversalError(tr("server.path_blocked"))
    if raw_name.startswith(".") or raw_name.startswith("~$"):
        raise PathTraversalError(tr("server.invalid_filename"))
    return raw_name.strip()


def archive_instead_of_delete(
    target: Path,
    archive_dir: Path,
    *,
    reason: str = "archived",
    device_id: str = "",
) -> Path:
    """
    Zero-deletion guarantee: moves `target` to `archive_dir` with a timestamp prefix
    and records a reversible entry in `_archive_journal.json`.
    """
    target = Path(target)
    if not target.exists():
        raise FileNotFoundError(tr("server.file_not_found"))
    archive_dir = Path(archive_dir)
    archive_dir.mkdir(parents=True, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = archive_dir / f"{ts}_{target.name}"
    idx = 2
    while dest.exists():
        dest = archive_dir / f"{ts}_{idx}_{target.name}"
        idx += 1
    file_hash = ""
    if target.is_file():
        try:
            file_hash = sha256_of_file(target)
        except Exception:
            file_hash = ""
    shutil.move(str(target), str(dest))
    journal_path = archive_dir / "_archive_journal.json"
    journal = read_json_safe(journal_path, [])
    if not isinstance(journal, list):
        journal = []
    journal.append({
        "id": f"ARC-{ts}-{secrets.token_hex(2).upper()}",
        "original_path": str(target),
        "archived_path": str(dest),
        "original_name": target.name,
        "sha256": file_hash,
        "reason": reason,
        "device_id": device_id,
        "timestamp": now_iso(),
    })
    atomic_write_json(journal_path, journal)
    return dest


def is_loopback_host(host: str) -> bool:
    h = (host or "").strip().lower()
    return h in ("127.0.0.1", "localhost", "::1")


def validate_bind_host(host: str, *, allow_remote_preview: bool = False) -> str:
    """
    Enforce local-only binding (`127.0.0.1`) by default.
    Non-loopback binding requires explicit `allow_remote_preview=True`.
    """
    h = (host or "127.0.0.1").strip()
    if is_loopback_host(h):
        return h
    if allow_remote_preview:
        return h
    raise PermissionError(tr("server.request_blocked"))


def validate_request_origin(headers: Any, *, allow_remote_preview: bool = False) -> bool:
    """
    Block cross-site CSRF / DNS-rebinding on mutating API calls when bound to localhost.
    Same-origin browser calls and local CLI/test clients without Origin pass cleanly.
    """
    if allow_remote_preview:
        return True
    host_hdr = str(headers.get("Host") or "").strip().lower()
    if host_hdr:
        try:
            parsed_host = urlparse(f"http://{host_hdr}")
            host_only = (parsed_host.hostname or "").lower()
            _ = parsed_host.port
        except ValueError:
            return False
        if host_only not in ("127.0.0.1", "localhost", "::1"):
            return False
    origin = str(headers.get("Origin") or "").strip()
    if origin:
        if origin.lower() == "null":
            return False
        try:
            parsed = urlparse(origin)
            ohost = (parsed.hostname or "").lower()
        except ValueError:
            return False
        if ohost not in ("127.0.0.1", "localhost", "::1"):
            return False
    return True
