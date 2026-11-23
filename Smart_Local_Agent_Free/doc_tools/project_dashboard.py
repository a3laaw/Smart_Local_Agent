#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
project_dashboard.py — Smart Local Agent Free (v10.0)
Local-first Multi-Project Document Control Server & REST API.

Core guarantees:
  1. Per-project UPLOADED Excel logs only (.xlsx / .xlsm with VBA intact) — no forced generic log.
  2. Per-project UPLOADED templates only (.xlsx / .xlsm / .docx) — zero system-suggested templates.
  3. Unrelated-file protection (rename_in_place + Undo, new_project_by_name fuzzy dedup, ask + learned_rules, dry_run, ignore_patterns).
  4. Active Project + Start/Sync (counter recalculation from uploaded log) + End Session (dated _Backups + external backup + >7d alert).
  5. Visual Template Studio (.tplmap.json sidecar, 6 roles incl. repeating rows & formula protection, anchor auto-rebind, orphan blocker, live preview).
  6. Localhost 127.0.0.1 binding, Host/Origin checks on all HTTP verbs, strict path-traversal & symlink guards, zero-deletion auto-archiving, and zero secrets in project/log/state files.
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import mimetypes
import os
from pathlib import Path
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from typing import Any, Callable
from urllib.parse import parse_qs, quote, unquote, urlparse
import uuid
import webbrowser

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from active_project import (
    DEFAULT_BACKUP_WARN_DAYS,
    PROCESS_MODES,
    check_backup_status,
    create_session_backup,
    detect_cloud_backup_folders,
    detect_shared_inboxes,
    effective_process_mode,
    get_projects_to_process,
    run_end_session,
    run_start_sync,
)
from decisions import DecisionStore
from inbox_pipeline import (
    analyze_multipage_pdf_split,
    apply_multipage_pdf_split,
    archive_dir_for_project,
    clear_scanner_resolved_items,
    get_scanner_triage_state,
    process_project_inbox_once,
    rename_journal_path_for,
    render_document_preview_page,
    render_triage_pdf_page,
    resolve_pending_decision,
    resolve_triage_bulk,
    resolve_triage_item,
    scan_inbox_to_triage,
    sync_decision_to_triage,
)
import letter_links as llinks
import project_log as plog
from relevance import (
    DEFAULT_IGNORE_PATTERNS,
    DEFAULT_RELEVANCE_THRESHOLD,
    UNRELATED_POLICIES,
    load_rename_journal,
    parse_ignore_patterns,
    undo_rename_entry,
)
import scan_watcher as sw
from security_utils import (
    DANGEROUS_EXTENSIONS,
    PathTraversalError,
    UnsafeFileError,
    atomic_write_json,
    get_or_create_device_id,
    normalize_win_path,
    now_iso,
    read_json_safe,
    resolve_local_or_unc,
    safe_resolve_within,
    safe_template_name,
    sanitize_upload_filename,
    scrub_secrets,
    sha256_of_file,
    validate_bind_host,
    validate_request_origin,
)
import ai_cache
import project_organizer as porg
import whatsapp_archiver as waa
import content_duplicates as dups
import task_bar as tbar
import template_profile as tprof
import template_render as trender
from i18n_ar import (
    DEFAULT_STYLE,
    VALID_STYLES,
    get_client_bundle,
    normalize_style,
    set_active_style,
    tr as _tr,
)

try:
    import whatsapp_bridge as wab
except Exception:
    wab = None  # type: ignore

APP_VERSION = "10.13.2"
MAX_BODY_BYTES = 64 * 1024 * 1024  # 64 MB (supports large .xlsm project logs)
MAX_WA_UPLOAD_BYTES = 1024 * 1024 * 1024  # 1 GB (streaming WhatsApp zip/video uploads)

DEFAULT_GLOBAL_SETTINGS: dict[str, Any] = {
    "ui_style": DEFAULT_STYLE,
    "process_mode": "active_only",
    "process_mode_explicit": False,
    "default_unrelated_policy": "rename_in_place",
    "relevance_threshold": DEFAULT_RELEVANCE_THRESHOLD,
    "dry_run": False,
    "ignore_patterns": [],
    "external_backup_dir": "",
    "auto_backup_on_end": True,
    "auto_projects_dir": "",
    "backup_warn_days": DEFAULT_BACKUP_WARN_DAYS,
    "poll_seconds": 6,
    "confirm_before_close": True,
    "engine": "gemini",
    "ollama_url": "http://127.0.0.1:11434",
    "ollama_model": "qwen2.5vl:7b",
}

ARCHIVE_SCAFFOLD_DIRS: tuple[str, ...] = (
    "01_Incoming - الوارد",
    "02_Outgoing - الصادر",
    "03_Submittals - الاعتمادات الفنية",
    "04_Contracts_and_IPCs - العقود والمستخلصات",
    "05_Petty_Cash_and_Invoices - العهد والفواتير",
    "06_Site_Reports_and_MOM - التقارير والمحاضر",
    "09_General - مستندات عامة",
)


def _clean_folder_seg(text: str, max_len: int = 60) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", str(text or "")).strip(" ._-")
    return re.sub(r"\s+", " ", s)[:max_len].strip() or "Project"


def scaffold_project_dirs(
    proj_root: Path,
    inbox_dir: Path | None = None,
    archive_dir: Path | None = None,
) -> None:
    """
    Create clean project folder structure WITHOUT generating any generic system log
    and WITHOUT copying any built-in/system templates!
    """
    proj_root.mkdir(parents=True, exist_ok=True)
    for sub in ("1_Inbox", "2_Archive", "3_Project_Logs", "4_Project_Templates", "_Backups", "_System_Agent"):
        (proj_root / sub).mkdir(parents=True, exist_ok=True)
    eff_archive = resolve_local_or_unc(archive_dir) if archive_dir is not None else (proj_root / "2_Archive")
    try:
        eff_archive.mkdir(parents=True, exist_ok=True)
        for cat in ARCHIVE_SCAFFOLD_DIRS:
            (eff_archive / cat).mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    if inbox_dir is not None:
        try:
            inbox_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass


def _has_invalid_win_question_marks(raw_path: str) -> bool:
    """Detect corrupted '?' characters (e.g. from codepage loss) excluding legal `\\\\?\\` prefix."""
    s = str(raw_path or "").strip()
    if not s:
        return False
    for prefix in ("\\\\?\\UNC\\", "//?/UNC/", "\\\\?\\", "//?/"):
        if s.startswith(prefix):
            s = s[len(prefix):]
            break
    return "?" in s


def _decode_ps_path_output(raw: str) -> str:
    """Decode Base64 UTF-8 path returned from PowerShell (`B64:...`) safely preserving Arabic names."""
    lines = [ln.strip() for ln in str(raw or "").splitlines() if ln.strip()]
    if not lines:
        return ""
    for line in reversed(lines):
        if line.startswith("B64:"):
            try:
                return base64.b64decode(line[4:].strip()).decode("utf-8", errors="strict").strip()
            except Exception:
                return ""
    return lines[-1]


def _resolve_picker_initial_dir(initial_dir: str) -> str:
    """Resolve nearest existing directory for picker initialization, ignoring corrupted `?` segments."""
    init_raw = str(initial_dir or "").strip()
    if not init_raw:
        return str(Path.home())
    if _has_invalid_win_question_marks(init_raw):
        # Walk up to nearest existing parent (e.g. `E:\\ICBS SHARE` from `E:\\ICBS SHARE\\?????`)
        clean_parts = re.split(r"[\\/]+", init_raw)
        valid_parts: list[str] = []
        for part in clean_parts:
            if "?" in part:
                break
            valid_parts.append(part)
        init_raw = "\\".join(valid_parts) if "\\" in init_raw else "/".join(valid_parts)
    init_p = str(normalize_win_path(init_raw)) if init_raw else ""
    cur = Path(init_p) if init_p else None
    while cur is not None:
        try:
            if cur.exists() and cur.is_dir():
                return str(cur)
        except OSError:
            pass
        if cur.parent == cur:
            break
        cur = cur.parent
    return str(Path.home())


# ══════════════════════════════════════════════════════════════════════════════
# HubState — Persistent Multi-Project State (Zero Secrets, Active-Project Aware)
# ══════════════════════════════════════════════════════════════════════════════
class HubState:
    def __init__(
        self,
        registry_path: Path,
        projects_root: Path,
        *,
        analyzer_override: Any = None,
    ):
        self.registry_path = Path(registry_path).resolve()
        self.state_dir = self.registry_path.parent
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.projects_root = Path(projects_root).resolve()
        self.projects_root.mkdir(parents=True, exist_ok=True)

        self.lock = threading.RLock()
        self.device_id = get_or_create_device_id(self.state_dir)
        self.decisions = DecisionStore(self.state_dir, device_id=self.device_id)
        self.analyzer_override = analyzer_override

        self.data = self._load()
        self.jobs: dict[str, dict[str, Any]] = {}
        self.legacy_jobs: dict[str, dict[str, Any]] = {}
        self._legacy_stops: dict[str, threading.Event] = {}
        self.organizer_jobs: dict[str, dict[str, Any]] = {}
        self._organizer_stops: dict[str, threading.Event] = {}
        self.dup_jobs: dict[str, dict[str, Any]] = {}
        self._dup_stops: dict[str, threading.Event] = {}
        self.wa_jobs: dict[str, dict[str, Any]] = {}
        self._wa_stops: dict[str, threading.Event] = {}
        self._watcher_stop = threading.Event()
        self._watcher_thread: threading.Thread | None = None
        self.watcher_running = False
        self.last_watcher_error: dict[str, str] | None = None

    def _load(self) -> dict[str, Any]:
        raw = read_json_safe(self.registry_path, {})
        if not isinstance(raw, dict):
            raw = {}
        settings = dict(DEFAULT_GLOBAL_SETTINGS)
        if isinstance(raw.get("settings"), dict):
            settings.update(scrub_secrets(raw["settings"]))
        # Migrate legacy top-level engine/poll keys cleanly
        for k in ("engine", "ollama_url", "ollama_model", "poll_seconds"):
            if k in raw and k not in (raw.get("settings") or {}):
                settings[k] = raw[k]

        projects: list[dict[str, Any]] = []
        for p in (raw.get("projects") or []):
            if not isinstance(p, dict) or not p.get("id") or not p.get("root"):
                continue
            cp = scrub_secrets(dict(p))
            cp.setdefault("enabled", True)
            cp.setdefault("domain", "contracting")
            cp.setdefault("unrelated_policy", "")  # empty => inherit global default (`rename_in_place`)
            cp.setdefault("relevance_threshold", None)
            cp.setdefault("dry_run", False)
            cp.setdefault("ignore_patterns", [])
            cp.setdefault("log_unrelated_sheet", False)
            cp.setdefault("global_submittal_numbering", None)
            cp.setdefault("overdue_days", None)
            cp.setdefault("aliases", [])
            cp.setdefault("keywords", [])
            cp.setdefault("subcontractors", [])
            cp.setdefault("external_backup_dir", "")
            cp.setdefault("baseline_counters", cp.get("starting_counters") or {})
            cp.setdefault("baseline_disc_counters", cp.get("starting_disc_counters") or {})
            cp.setdefault("synced_counters", {})
            cp.setdefault("contacts", [])
            cp.setdefault("archive_dir", str(cp.get("archive_dir") or ""))
            cp.setdefault("custom_sections", list(cp.get("custom_sections") or []))
            projects.append(cp)

        active_id = raw.get("active_project_id")
        if active_id and not any(p["id"] == active_id for p in projects):
            active_id = None
        # Auto-enforce `active_only` when >1 project and not explicitly changed
        settings["process_mode"] = effective_process_mode(settings, len(projects))
        settings["ui_style"] = normalize_style(settings.get("ui_style"))
        set_active_style(settings["ui_style"])

        out = {
            "version": APP_VERSION,
            "device_id": self.device_id,
            "active_project_id": active_id,
            "session_active": bool(raw.get("session_active", False) and active_id),
            "session_started_at": raw.get("session_started_at") if active_id else None,
            "settings": settings,
            "projects": projects,
            "updated_at": now_iso(),
        }
        atomic_write_json(self.registry_path, out, scrub=True)
        return out

    def save(self) -> None:
        with self.lock:
            self.data["device_id"] = self.device_id
            self.data["settings"]["process_mode"] = effective_process_mode(
                self.data["settings"], len(self.data["projects"])
            )
            self.data["updated_at"] = now_iso()
            atomic_write_json(self.registry_path, self.data, scrub=True)

    def list_projects(self) -> list[dict[str, Any]]:
        with self.lock:
            return [dict(p) for p in self.data["projects"]]

    def get_project(self, pid: str) -> dict[str, Any] | None:
        with self.lock:
            for p in self.data["projects"]:
                if str(p.get("id")) == str(pid):
                    return dict(p)
        return None

    def _get_project_ref_unlocked(self, pid: str) -> dict[str, Any] | None:
        for p in self.data["projects"]:
            if str(p.get("id")) == str(pid):
                return p
        return None

    def get_settings(self) -> dict[str, Any]:
        with self.lock:
            s = dict(self.data["settings"])
            s["ui_style"] = normalize_style(s.get("ui_style"))
            set_active_style(s["ui_style"])
            s["process_mode"] = effective_process_mode(s, len(self.data["projects"]))
            s["device_id"] = self.device_id
            s["projects_root"] = str(self.projects_root)
            s["active_project_id"] = self.data.get("active_project_id")
            s["session_active"] = bool(self.data.get("session_active"))
            s["default_system_ignores"] = list(DEFAULT_IGNORE_PATTERNS)
            return s

    def update_settings(self, patch: dict[str, Any]) -> dict[str, Any]:
        with self.lock:
            s = self.data["settings"]
            if "ui_style" in patch:
                us = str(patch["ui_style"] or "").strip().lower()
                if us not in VALID_STYLES:
                    raise ValueError(_tr("server.invalid_setting"))
                s["ui_style"] = us
                set_active_style(us)
            if "process_mode" in patch:
                pm = str(patch["process_mode"] or "").strip().lower()
                if pm not in PROCESS_MODES:
                    raise ValueError(_tr("server.invalid_setting"))
                s["process_mode"] = pm
                s["process_mode_explicit"] = True
            if "default_unrelated_policy" in patch:
                pol = str(patch["default_unrelated_policy"] or "").strip()
                if pol not in UNRELATED_POLICIES:
                    raise ValueError(_tr("server.invalid_setting"))
                s["default_unrelated_policy"] = pol
            if "relevance_threshold" in patch and patch["relevance_threshold"] not in (None, ""):
                thr = float(patch["relevance_threshold"])
                if not (0.1 <= thr <= 0.95):
                    raise ValueError(_tr("server.sensitivity_range"))
                s["relevance_threshold"] = round(thr, 2)
            if "dry_run" in patch:
                s["dry_run"] = bool(patch["dry_run"])
            if "ignore_patterns" in patch:
                s["ignore_patterns"] = parse_ignore_patterns(patch["ignore_patterns"])
            if "external_backup_dir" in patch:
                s["external_backup_dir"] = str(patch["external_backup_dir"] or "").strip()
            if "auto_backup_on_end" in patch:
                s["auto_backup_on_end"] = bool(patch["auto_backup_on_end"])
            if "auto_projects_dir" in patch:
                s["auto_projects_dir"] = str(patch["auto_projects_dir"] or "").strip()
            if "backup_warn_days" in patch and patch["backup_warn_days"] not in (None, ""):
                s["backup_warn_days"] = max(1, min(90, int(patch["backup_warn_days"])))
            if "poll_seconds" in patch and patch["poll_seconds"] not in (None, ""):
                s["poll_seconds"] = max(3, min(300, int(patch["poll_seconds"])))
            if "confirm_before_close" in patch:
                s["confirm_before_close"] = bool(patch["confirm_before_close"])
            if "engine" in patch and str(patch["engine"]).lower() in ("gemini", "ollama"):
                s["engine"] = str(patch["engine"]).lower()
            if "ollama_url" in patch and str(patch["ollama_url"]).strip():
                s["ollama_url"] = str(patch["ollama_url"]).strip()
            if "ollama_model" in patch and str(patch["ollama_model"]).strip():
                s["ollama_model"] = str(patch["ollama_model"]).strip()
            self.save()
        return self.get_settings()

    # ── Project CRUD (Zero System Templates, Per-Project Uploaded Logs Only!) ─
    def create_project(self, payload: dict[str, Any]) -> dict[str, Any]:
        name = str(payload.get("name") or "").strip()
        if not name:
            raise ValueError(_tr("new_project.name_required"))
        raw_code = str(payload.get("code") or "").strip().upper()
        code = re.sub(r"[^A-Z0-9\u0600-\u06FF\-_]", "", raw_code)[:16] or "PRJ"

        custom_root = str(payload.get("root") or payload.get("project_dir") or "").strip()
        custom_inbox = str(payload.get("inbox_dir") or "").strip()
        custom_archive = str(payload.get("archive_dir") or "").strip()
        linked_log_path_chk = str(payload.get("linked_log_path") or "").strip()
        for cand_p in (custom_root, custom_inbox, custom_archive, linked_log_path_chk):
            if _has_invalid_win_question_marks(cand_p):
                raise ValueError(_tr("server.invalid_path_question_marks"))

        if custom_root:
            proj_root = resolve_local_or_unc(custom_root)
        else:
            folder_nm = _clean_folder_seg(f"{code} - {name}", 75)
            proj_root = resolve_local_or_unc(self.projects_root / folder_nm)

        inbox_dir = resolve_local_or_unc(custom_inbox) if custom_inbox else (proj_root / "1_Inbox")

        archive_dir = resolve_local_or_unc(custom_archive) if custom_archive else (proj_root / "2_Archive")

        scaffold_project_dirs(proj_root, inbox_dir, archive_dir)

        pid = f"prj_{dt.datetime.now().strftime('%Y%m%d%H%M%S')}_{secrets.token_hex(2)}"
        pol = str(payload.get("unrelated_policy") or "").strip()
        if pol and pol not in UNRELATED_POLICIES:
            pol = ""

        def _to_list(v: Any) -> list[str]:
            if isinstance(v, list):
                return [str(x).strip() for x in v if str(x).strip()]
            if isinstance(v, str):
                return [x.strip() for x in re.split(r"[,،;\n]+", v) if x.strip()]
            return []

        proj: dict[str, Any] = {
            "id": pid,
            "name": name,
            "code": code,
            "contract_no": str(payload.get("contract_no") or "").strip(),
            "owner": str(payload.get("owner") or "").strip(),
            "consultant": str(payload.get("consultant") or "").strip(),
            "contractor": str(payload.get("contractor") or "").strip(),
            "location": str(payload.get("location") or "").strip(),
            "engineer": str(payload.get("engineer") or "").strip(),
            "domain": str(payload.get("domain") or "contracting").strip(),
            "root": str(proj_root),
            "inbox_dir": str(inbox_dir),
            "archive_dir": str(archive_dir),
            "custom_sections": list(payload.get("custom_sections") or []),
            "enabled": True,
            "unrelated_policy": pol,
            "relevance_threshold": float(payload["relevance_threshold"]) if payload.get("relevance_threshold") not in (None, "") else None,
            "dry_run": bool(payload.get("dry_run", False)),
            "ignore_patterns": parse_ignore_patterns(payload.get("ignore_patterns")),
            "log_unrelated_sheet": bool(payload.get("log_unrelated_sheet", False)),
            "global_submittal_numbering": (bool(payload["global_submittal_numbering"]) if payload.get("global_submittal_numbering") is not None else None),
            "overdue_days": (max(1, min(90, int(payload["overdue_days"]))) if payload.get("overdue_days") not in (None, "") else None),
            "aliases": _to_list(payload.get("aliases")),
            "keywords": _to_list(payload.get("keywords")),
            "subcontractors": _to_list(payload.get("subcontractors")),
            "external_backup_dir": str(payload.get("external_backup_dir") or "").strip(),
            "baseline_counters": {},
            "baseline_disc_counters": {},
            "synced_counters": {},
            "contacts": [],
            "created_at": now_iso(),
            "created_by_device": self.device_id,
            "last_started_at": None,
            "last_backup_at": None,
            "last_external_backup_at": None,
        }

        # Optional: if user uploaded OR directly linked an external project log in "New Project" modal!
        init_log = payload.get("initial_log")
        linked_log_path = str(payload.get("linked_log_path") or "").strip()
        if isinstance(init_log, dict) and init_log.get("name") and init_log.get("data_b64"):
            raw_b = base64.b64decode(str(init_log["data_b64"]))
            plog.upload_project_log(proj_root, str(init_log["name"]), raw_b, device_id=self.device_id)
            sync = plog.recalculate_counters_from_logs(proj, device_id=self.device_id)
            proj["synced_counters"] = {
                "effective_counters": sync["effective_counters"],
                "effective_disc": sync["effective_disc"],
                "arabic_series": sync["arabic_series"],
                "synced_at": sync["synced_at"],
            }
        elif linked_log_path:
            plog.link_external_project_log(proj_root, linked_log_path, device_id=self.device_id)
            sync = plog.recalculate_counters_from_logs(proj, device_id=self.device_id)
            proj["synced_counters"] = {
                "effective_counters": sync["effective_counters"],
                "effective_disc": sync["effective_disc"],
                "arabic_series": sync["arabic_series"],
                "synced_at": sync["synced_at"],
            }

        with self.lock:
            self.data["projects"].append(proj)
            # If this is the first project or caller requested activation, set active
            if len(self.data["projects"]) == 1 or payload.get("set_active", True):
                self.data["active_project_id"] = pid
            # Automatically switch to `active_only` when adding a 2nd project unless user explicitly overrode
            self.data["settings"]["process_mode"] = effective_process_mode(
                self.data["settings"], len(self.data["projects"])
            )
            self.save()
        return dict(proj)

    def update_project(self, pid: str, patch: dict[str, Any]) -> dict[str, Any]:
        if "enabled" in patch and not bool(patch["enabled"]):
            self.stop_legacy_import(pid)
            self.stop_organizer_job(pid)
        with self.lock:
            proj = self._get_project_ref_unlocked(pid)
            if not proj:
                raise KeyError(_tr("server.project_not_found"))
            for k in ("name", "contract_no", "owner", "consultant", "contractor", "location", "engineer", "domain", "external_backup_dir"):
                if k in patch:
                    val = str(patch[k] or "").strip()
                    if k == "name" and not val:
                        continue
                    proj[k] = val
            if "code" in patch and str(patch["code"]).strip():
                proj["code"] = re.sub(r"[^A-Z0-9\u0600-\u06FF\-_]", "", str(patch["code"]).strip().upper())[:16] or proj["code"]
            if "inbox_dir" in patch and str(patch["inbox_dir"]).strip():
                raw_ibx = str(patch["inbox_dir"]).strip()
                if _has_invalid_win_question_marks(raw_ibx):
                    raise ValueError(_tr("server.invalid_path_question_marks"))
                ibx = resolve_local_or_unc(raw_ibx)
                proj["inbox_dir"] = str(ibx)
                try:
                    ibx.mkdir(parents=True, exist_ok=True)
                except OSError:
                    pass
            if "archive_dir" in patch:
                raw_arc = str(patch["archive_dir"] or "").strip()
                if raw_arc:
                    if _has_invalid_win_question_marks(raw_arc):
                        raise ValueError(_tr("server.invalid_path_question_marks"))
                    arc_p = resolve_local_or_unc(raw_arc)
                    proj["archive_dir"] = str(arc_p)
                    try:
                        arc_p.mkdir(parents=True, exist_ok=True)
                        for cat in ARCHIVE_SCAFFOLD_DIRS:
                            (arc_p / cat).mkdir(parents=True, exist_ok=True)
                    except OSError:
                        pass
                else:
                    proj["archive_dir"] = str(resolve_local_or_unc(str(proj["root"])) / "2_Archive")
            if "custom_sections" in patch and isinstance(patch["custom_sections"], list):
                proj["custom_sections"] = self._normalize_custom_sections(proj, patch["custom_sections"])
            if "enabled" in patch:
                en = bool(patch["enabled"])
                proj["enabled"] = en
                if not en and self.data.get("active_project_id") == pid:
                    self.data["active_project_id"] = None
                    self.data["session_active"] = False
                    self.data["session_started_at"] = None
            if "unrelated_policy" in patch:
                pol = str(patch["unrelated_policy"] or "").strip()
                if pol and pol not in UNRELATED_POLICIES:
                    raise ValueError(_tr("server.invalid_setting"))
                proj["unrelated_policy"] = pol
            if "relevance_threshold" in patch:
                rt = patch["relevance_threshold"]
                if rt in (None, ""):
                    proj["relevance_threshold"] = None
                else:
                    fv = float(rt)
                    if not (0.1 <= fv <= 0.95):
                        raise ValueError(_tr("server.sensitivity_range"))
                    proj["relevance_threshold"] = round(fv, 2)
            if "dry_run" in patch:
                proj["dry_run"] = bool(patch["dry_run"])
            if "ignore_patterns" in patch:
                proj["ignore_patterns"] = parse_ignore_patterns(patch["ignore_patterns"])
            if "log_unrelated_sheet" in patch:
                proj["log_unrelated_sheet"] = bool(patch["log_unrelated_sheet"])
            if "global_submittal_numbering" in patch:
                gsn = patch["global_submittal_numbering"]
                proj["global_submittal_numbering"] = None if gsn in (None, "") else bool(gsn)
            if "overdue_days" in patch:
                od = patch["overdue_days"]
                proj["overdue_days"] = None if od in (None, "") else max(1, min(90, int(od)))
            for lk in ("aliases", "keywords", "subcontractors"):
                if lk in patch:
                    rv = patch[lk]
                    if isinstance(rv, list):
                        proj[lk] = [str(x).strip() for x in rv if str(x).strip()]
                    else:
                        proj[lk] = [x.strip() for x in re.split(r"[,،;\n]+", str(rv or "")) if x.strip()]
            if "baseline_counters" in patch and isinstance(patch["baseline_counters"], dict):
                proj["baseline_counters"] = {
                    str(k).strip().upper(): max(0, int(v or 0))
                    for k, v in patch["baseline_counters"].items()
                    if str(k).strip()
                }
            if "baseline_disc_counters" in patch and isinstance(patch["baseline_disc_counters"], dict):
                proj["baseline_disc_counters"] = {
                    plog.canonical_discipline(str(k)): max(0, int(v or 0))
                    for k, v in patch["baseline_disc_counters"].items()
                    if str(k).strip()
                }
            proj["updated_at"] = now_iso()
            self.save()
            return dict(proj)

    def unlink_project(self, pid: str) -> bool:
        """Unlink project from dashboard registry — ZERO files deleted on disk!"""
        self.stop_legacy_import(pid)
        self.stop_organizer_job(pid)
        with self.lock:
            before = len(self.data["projects"])
            self.data["projects"] = [p for p in self.data["projects"] if str(p.get("id")) != str(pid)]
            if len(self.data["projects"]) == before:
                return False
            if self.data.get("active_project_id") == pid:
                self.data["active_project_id"] = self.data["projects"][0]["id"] if self.data["projects"] else None
                self.data["session_active"] = False
                self.data["session_started_at"] = None
            self.save()
            return True

    def delete_project(self, pid: str, *, delete_files: bool = False) -> dict[str, Any]:
        """
        Remove a project from the hub.
        When `delete_files=False` (default), behaves identically to `unlink_project` (zero files touched).
        When `delete_files=True`, deletes `proj['root']` strictly after validating it is a dedicated
        project folder (never a drive root, user home, Desktop, hub state dir, or shared with another project).
        """
        if not delete_files:
            ok = self.unlink_project(pid)
            return {"ok": ok, "deleted_files": False}

        self.stop_legacy_import(pid)
        self.stop_organizer_job(pid)
        with self.lock:
            target = self._get_project_ref_unlocked(pid)
            if not target:
                return {"ok": False, "deleted_files": False}
            proj_root = Path(str(target.get("root") or "")).resolve()
            forbidden_paths = {
                Path(proj_root.anchor).resolve(),
                Path.home().resolve(),
                (Path.home() / "Desktop").resolve(),
                (Path.home() / "Documents").resolve(),
                (Path.home() / "Downloads").resolve(),
                self.projects_root.resolve(),
                self.state_dir.resolve(),
                self.state_dir.parent.resolve(),
                SCRIPT_DIR.resolve(),
                SCRIPT_DIR.parent.resolve(),
                Path("/").resolve(),
            }
            if len(proj_root.parts) <= 2 or proj_root in forbidden_paths:
                raise ValueError(_tr("hub.delete_unsafe_err"))
            for other in self.data["projects"]:
                if str(other.get("id")) == str(pid):
                    continue
                other_root = Path(str(other.get("root") or "")).resolve()
                if other_root == proj_root or other_root in proj_root.parents or proj_root in other_root.parents:
                    raise ValueError(_tr("hub.delete_unsafe_err"))

            self.data["projects"] = [p for p in self.data["projects"] if str(p.get("id")) != str(pid)]
            if self.data.get("active_project_id") == pid:
                self.data["active_project_id"] = self.data["projects"][0]["id"] if self.data["projects"] else None
                self.data["session_active"] = False
                self.data["session_started_at"] = None
            self.save()

        removed_disk = False
        disk_warning = ""
        if proj_root.exists() and proj_root.is_dir():
            try:
                shutil.rmtree(proj_root)
                removed_disk = True
            except Exception as exc:
                disk_warning = str(exc)
        return {"ok": True, "deleted_files": removed_disk, "warning": disk_warning}

    def stop_all_background(self) -> None:
        self.toggle_watcher(False)
        for p in self.list_projects():
            self.stop_legacy_import(p["id"])
            self.stop_organizer_job(p["id"])
            self.stop_scoped_job("dup", p["id"])
            self.stop_scoped_job("wa", p["id"])

    def reset_system(
        self,
        *,
        delete_project_folders: bool = False,
        keep_ai_key: bool = True,
        keep_whatsapp: bool = True,
        reset_settings: bool = True,
    ) -> dict[str, Any]:
        """
        Full factory reset for clearing trial/test projects, pending decisions,
        learned rules, caches, and optionally resetting settings & dedicated project folders.
        External scanner/archive folders outside dedicated project roots are never deleted.
        """
        self.stop_all_background()
        folders_deleted = 0
        folder_warnings: list[str] = []

        with self.lock:
            existing_projects = list(self.data.get("projects") or [])
            projects_removed = len(existing_projects)
            cur_style = normalize_style((self.data.get("settings") or {}).get("ui_style"))

            roots_to_delete: list[Path] = []
            if delete_project_folders:
                forbidden_paths = {
                    Path(self.projects_root.anchor).resolve(),
                    Path.home().resolve(),
                    (Path.home() / "Desktop").resolve(),
                    (Path.home() / "Documents").resolve(),
                    (Path.home() / "Downloads").resolve(),
                    self.projects_root.resolve(),
                    self.state_dir.resolve(),
                    self.state_dir.parent.resolve(),
                    SCRIPT_DIR.resolve(),
                    SCRIPT_DIR.parent.resolve(),
                    Path("/").resolve(),
                }
                for p in existing_projects:
                    raw_r = str(p.get("root") or "").strip()
                    if not raw_r:
                        continue
                    pr = Path(raw_r).resolve()
                    if len(pr.parts) <= 2 or pr in forbidden_paths:
                        continue
                    if any(fp == pr or pr in fp.parents for fp in forbidden_paths):
                        continue
                    roots_to_delete.append(pr)

            self.data["projects"] = []
            self.data["active_project_id"] = None
            self.data["session_active"] = False
            self.data["session_started_at"] = None
            self.data["version"] = APP_VERSION
            if reset_settings:
                fresh_settings = dict(DEFAULT_GLOBAL_SETTINGS)
                fresh_settings["ui_style"] = cur_style
                self.data["settings"] = fresh_settings
            self.jobs.clear()
            self.legacy_jobs.clear()
            self.organizer_jobs.clear()
            self.dup_jobs.clear()
            self.wa_jobs.clear()
            self.last_watcher_error = None
            self.save()

        dec_res = self.decisions.clear_all()
        try:
            plog._SUMMARY_CACHE.clear()
        except Exception:
            pass
        prev_cache = self.state_dir / "_preview_cache"
        if prev_cache.exists():
            shutil.rmtree(prev_cache, ignore_errors=True)

        if delete_project_folders:
            seen_roots: set[Path] = set()
            for pr in roots_to_delete:
                if pr in seen_roots:
                    continue
                seen_roots.add(pr)
                if pr.exists() and pr.is_dir():
                    try:
                        shutil.rmtree(pr)
                        folders_deleted += 1
                    except Exception as exc:
                        folder_warnings.append(f"{pr.name}: {exc}")

        if not keep_ai_key:
            try:
                sw.save_gemini_key("", "", backup_api_key="")
            except Exception:
                pass

        if not keep_whatsapp:
            wa_cfg_p = self.state_dir / "whatsapp_config.json"
            if wa_cfg_p.exists():
                try:
                    atomic_write_json(wa_cfg_p, {})
                except Exception:
                    pass

        return {
            "ok": True,
            "projects_removed": projects_removed,
            "folders_deleted": folders_deleted,
            "folder_warnings": folder_warnings,
            "kept_ai_key": bool(keep_ai_key),
            "kept_whatsapp": bool(keep_whatsapp),
            "settings_reset": bool(reset_settings),
            **dec_res,
        }

    # ── Active Project + Start/Sync + End Session ────────────────────────────
    def set_active_project(self, pid: str | None) -> dict[str, Any]:
        with self.lock:
            if pid is None or str(pid).strip() == "":
                self.data["active_project_id"] = None
                self.data["session_active"] = False
                self.data["session_started_at"] = None
                self.save()
                return {"ok": True, "active_project_id": None}
            proj = self._get_project_ref_unlocked(str(pid))
            if not proj:
                raise KeyError(_tr("server.project_not_found"))
            # Stop legacy imports on other projects if in active_only mode
            pm = effective_process_mode(self.data["settings"], len(self.data["projects"]))
            if pm == "active_only":
                for op in self.data["projects"]:
                    if op["id"] != proj["id"]:
                        ev = self._legacy_stops.get(op["id"])
                        if ev:
                            ev.set()
            self.data["active_project_id"] = proj["id"]
            self.save()
            return {
                "ok": True,
                "active_project_id": proj["id"],
                "project_name": proj["name"],
                "process_mode": pm,
                "device_id": self.device_id,
            }

    def start_project(self, pid: str) -> dict[str, Any]:
        """Full Start Work / Sync routine."""
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        all_p = self.list_projects()
        gcfg = self.get_settings()
        # In active_only mode, halt background imports on all other projects
        if gcfg["process_mode"] == "active_only":
            for op in all_p:
                if op["id"] != pid:
                    self.stop_legacy_import(op["id"])

        pcount = self.decisions.pending_count(project_id=pid)
        report = run_start_sync(
            proj,
            all_p,
            gcfg,
            device_id=self.device_id,
            pending_decisions_count=pcount,
        )
        with self.lock:
            pref = self._get_project_ref_unlocked(pid)
            if pref is not None:
                # Overwrite any stale cached counters cleanly with actual log counters!
                pref["synced_counters"] = {
                    "effective_counters": report["counters"]["effective_counters"],
                    "effective_disc": report["counters"]["effective_disc"],
                    "arabic_series": report["counters"]["arabic_series"],
                    "rows_scanned": report["counters"]["rows_scanned"],
                    "synced_at": report["started_at"],
                }
                # Purge any legacy v8/v9 inflated `counters` cache so nothing can ever read stale numbers
                pref.pop("counters", None)
                pref.pop("disc_counters", None)
                pref["last_started_at"] = report["started_at"]
                pref["last_device_id"] = self.device_id
            self.data["active_project_id"] = pid
            self.data["session_active"] = True
            self.data["session_started_at"] = report["started_at"]
            self.save()
        return report

    def end_project(self, pid: str, *, copy_external: bool = True, include_archive_external: bool = False) -> dict[str, Any]:
        """Full End Session routine with clean shutdown & dated backup."""
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))

        # 1. Stop watcher (if active_only or targeting this project) & stop legacy import
        self.stop_legacy_import(pid)
        gcfg = self.get_settings()
        if self.watcher_running and (gcfg["process_mode"] == "active_only" or self.data.get("active_project_id") == pid):
            self.toggle_watcher(False)

        # Wait briefly if an organize job on this project is finishing
        deadline = time.time() + 8.0
        while time.time() < deadline:
            with self.lock:
                j = self.jobs.get(pid) or {}
                if not j.get("running"):
                    break
            time.sleep(0.15)

        report = run_end_session(
            proj,
            gcfg,
            device_id=self.device_id,
            copy_external=copy_external,
            include_archive_external=include_archive_external,
        )
        with self.lock:
            pref = self._get_project_ref_unlocked(pid)
            if pref is not None:
                pref["last_backup_at"] = report["backup"]["created_at"]
                pref["last_ended_at"] = report["ended_at"]
                pref["last_device_id"] = self.device_id
                if report["backup"]["external"].get("copied"):
                    pref["last_external_backup_at"] = report["backup"]["created_at"]
            if self.data.get("active_project_id") == pid:
                self.data["active_project_id"] = None
                self.data["session_active"] = False
                self.data["session_started_at"] = None
            self.save()
        return report

    def quick_backup_project(
        self,
        pid: str,
        *,
        copy_external: bool = True,
        include_archive_external: bool = False,
        external_dir_override: str = "",
    ) -> dict[str, Any]:
        """Create an immediate dated backup snapshot without ending the active work session."""
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        gcfg = self.get_settings()
        if external_dir_override.strip():
            gcfg = {**gcfg, "external_backup_dir": external_dir_override.strip()}
        bkp = create_session_backup(
            proj,
            device_id=self.device_id,
            global_settings=gcfg,
            copy_external=copy_external,
            include_archive_external=include_archive_external,
            reason="quick_backup",
        )
        with self.lock:
            pref = self._get_project_ref_unlocked(pid)
            if pref is not None:
                pref["last_backup_at"] = bkp["created_at"]
                pref["last_device_id"] = self.device_id
                if bkp.get("external", {}).get("copied"):
                    pref["last_external_backup_at"] = bkp["created_at"]
            self.save()
        return {"ok": True, "project_id": pid, "project_name": proj.get("name", ""), "backup": bkp}

    def refresh_project_counters(self, pid: str) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        sync = plog.recalculate_counters_from_logs(proj, device_id=self.device_id)
        with self.lock:
            pref = self._get_project_ref_unlocked(pid)
            if pref is not None:
                pref["synced_counters"] = {
                    "effective_counters": sync["effective_counters"],
                    "effective_disc": sync["effective_disc"],
                    "arabic_series": sync["arabic_series"],
                    "rows_scanned": sync["rows_scanned"],
                    "synced_at": sync["synced_at"],
                }
                pref.pop("counters", None)
                pref.pop("disc_counters", None)
                self.save()
        return sync

    @staticmethod
    def _normalize_custom_sections(proj: dict[str, Any], raw_sections: list[dict[str, Any]]) -> list[dict[str, Any]]:
        clean: list[dict[str, Any]] = []
        seen: set[str] = set()
        arc_root = archive_dir_for_project(proj)
        for item in raw_sections:
            if not isinstance(item, dict):
                continue
            lbl = str(item.get("label") or "").strip()
            sid = str(item.get("id") or "").strip()
            if not sid:
                if not lbl:
                    continue
                slug = re.sub(r"[^a-zA-Z0-9_]+", "_", lbl).strip("_").lower()[:24]
                sid = f"custom_{slug or secrets.token_hex(3)}"
            if sid in seen:
                continue
            seen.add(sid)
            pfx = re.sub(r"[^A-Z0-9_-]", "", str(item.get("ref_prefix") or "DOC").upper())[:12] or "DOC"
            short = str(item.get("short") or lbl or pfx).strip()[:18]
            af = re.sub(r'[<>:"|?*]', "_", str(item.get("archive_folder") or item.get("folder") or "")).strip("/\\. ")
            is_hidden = bool(item.get("hidden", False)) or (item.get("enabled") is False)
            entry: dict[str, Any] = {
                "id": sid,
                "label": lbl or sid,
                "short": short,
                "ref_prefix": pfx,
                "hidden": is_hidden,
            }
            if af:
                entry["archive_folder"] = af
                try:
                    (arc_root / af).mkdir(parents=True, exist_ok=True)
                except OSError:
                    pass
            elif sid not in tprof.SECTIONS_BY_ID:
                safe_af = f"10_{_clean_folder_seg(lbl or sid, 40)}"
                entry["archive_folder"] = safe_af
                try:
                    (arc_root / safe_af).mkdir(parents=True, exist_ok=True)
                except OSError:
                    pass
            if item.get("log_section"):
                entry["log_section"] = str(item["log_section"])
            if item.get("subtype"):
                entry["subtype"] = str(item["subtype"])
            clean.append(entry)
        return clean

    def update_project_sections(self, pid: str, raw_sections: list[dict[str, Any]]) -> list[dict[str, Any]]:
        with self.lock:
            proj = self._get_project_ref_unlocked(pid)
            if not proj:
                raise KeyError(_tr("server.project_not_found"))
            proj["custom_sections"] = self._normalize_custom_sections(proj, raw_sections)
            proj["updated_at"] = now_iso()
            self.save()
            style = self.get_settings().get("ui_style")
            return tprof.get_project_sections(proj, style=style, include_hidden=True)

    # ── Scanner Triage (Read-only Scan -> User Approve/Reject) ────────────────
    def get_scanner_job_status(self, pid: str) -> dict[str, Any]:
        with self.lock:
            job = dict(self.jobs.get(pid) or {"running": False})
        if job.get("running") and job.get("started_ts"):
            job["elapsed_seconds"] = max(0, int(time.time() - float(job["started_ts"])))
        return job

    def dismiss_scanner_job(self, pid: str) -> None:
        with self.lock:
            cur = self.jobs.get(pid)
            if cur and not cur.get("running"):
                self.jobs.pop(pid, None)

    def run_scanner_triage_scan(
        self,
        pid: str,
        *,
        blocking: bool = False,
        analyzer_override: Any = None,
    ) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        now_ts = time.time()
        with self.lock:
            if (self.jobs.get(pid) or {}).get("running"):
                return {"ok": False, "error": _tr("server.scan_already_running")}
            self.jobs[pid] = {
                "running": True,
                "kind": "scanner_scan",
                "phase": "scanning",
                "started_at": now_iso(),
                "started_ts": now_ts,
                "elapsed_seconds": 0,
                "last_summary": None,
                "error": "",
            }

        eff_analyzer = analyzer_override if analyzer_override is not None else self.analyzer_override

        def _exec() -> dict[str, Any]:
            try:
                res = scan_inbox_to_triage(
                    self.get_project(pid) or proj,
                    all_projects=self.list_projects(),
                    global_settings=self.get_settings(),
                    decision_store=self.decisions,
                    analyzer_override=eff_analyzer,
                    device_id=self.device_id,
                )
                errs = list(res.get("errors") or []) if isinstance(res, dict) else []
                with self.lock:
                    prev_st = (self.jobs.get(pid) or {}).get("started_ts") or now_ts
                    self.jobs[pid] = {
                        "running": False,
                        "kind": "scanner_scan",
                        "phase": "done",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "last_summary": res,
                        "summary": res,
                        "error": errs[0] if errs and not (res.get("added") or res.get("archived")) else "",
                    }
                return res
            except Exception as exc:
                err = {"ok": False, "error": str(exc)}
                with self.lock:
                    prev_st = (self.jobs.get(pid) or {}).get("started_ts") or now_ts
                    self.jobs[pid] = {
                        "running": False,
                        "kind": "scanner_scan",
                        "phase": "error",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "last_summary": err,
                        "error": str(exc),
                    }
                return err

        if blocking:
            return _exec()
        t = threading.Thread(target=_exec, daemon=True)
        t.start()
        return {"ok": True, "started": True, "message": _tr("server.scan_started"), "job": self.get_scanner_job_status(pid)}

    # ── Inbox Processing & Continuous Watcher (Respects `active_only`!) ──────
    def run_inbox_sync_now(self, pid: str, *, blocking: bool = False, analyzer_override: Any = None) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        now_ts = time.time()
        with self.lock:
            if (self.jobs.get(pid) or {}).get("running"):
                return {"ok": False, "error": _tr("server.scan_already_running")}
            self.jobs[pid] = {
                "running": True,
                "kind": "inbox_sync",
                "phase": "scanning",
                "started_at": now_iso(),
                "started_ts": now_ts,
                "elapsed_seconds": 0,
                "last_summary": None,
                "error": "",
            }

        eff_analyzer = analyzer_override if analyzer_override is not None else self.analyzer_override

        def _exec() -> dict[str, Any]:
            try:
                res = process_project_inbox_once(
                    self.get_project(pid) or proj,
                    all_projects=self.list_projects(),
                    global_settings=self.get_settings(),
                    decision_store=self.decisions,
                    device_id=self.device_id,
                    analyzer_override=eff_analyzer,
                )
                self.refresh_project_counters(pid)
                with self.lock:
                    prev_st = (self.jobs.get(pid) or {}).get("started_ts") or now_ts
                    self.jobs[pid] = {
                        "running": False,
                        "kind": "inbox_sync",
                        "phase": "done",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "last_summary": res,
                        "summary": res,
                        "error": "",
                    }
                return res
            except Exception as exc:
                err = {"ok": False, "error": str(exc)}
                with self.lock:
                    prev_st = (self.jobs.get(pid) or {}).get("started_ts") or now_ts
                    self.jobs[pid] = {
                        "running": False,
                        "kind": "inbox_sync",
                        "phase": "error",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "last_summary": err,
                        "error": str(exc),
                    }
                return err

        if blocking:
            return _exec()
        t = threading.Thread(target=_exec, daemon=True)
        t.start()
        return {"ok": True, "started": True, "message": _tr("server.scan_started"), "job": self.get_scanner_job_status(pid)}

    def run_watcher_cycle_once(self, *, analyzer_override: Any = None) -> dict[str, Any]:
        """
        Run one Auto-Scanner tick strictly respecting `process_mode` (`active_only` vs `all_projects`):
        scans inbox folders read-only and populates the Scanner Triage screen ready for user Approve/Reject
        (never moves or auto-archives files without explicit user approval!).
        """
        all_p = self.list_projects()
        gcfg = self.get_settings()
        targets = get_projects_to_process(all_p, gcfg.get("active_project_id"), gcfg["process_mode"])
        results: dict[str, Any] = {}
        for p in targets:
            results[p["id"]] = self.run_scanner_triage_scan(p["id"], blocking=True, analyzer_override=analyzer_override)
        return {
            "process_mode": gcfg["process_mode"],
            "active_project_id": gcfg.get("active_project_id"),
            "processed_project_ids": list(results.keys()),
            "results": results,
        }

    def toggle_watcher(self, enable: bool | None = None) -> bool:
        with self.lock:
            want = (not self.watcher_running) if enable is None else bool(enable)
            if want and not self.watcher_running:
                self._watcher_stop.clear()
                self.watcher_running = True
                self._watcher_thread = threading.Thread(target=self._bg_watch_loop, daemon=True)
                self._watcher_thread.start()
            elif not want and self.watcher_running:
                self._watcher_stop.set()
                self.watcher_running = False
            th = self._watcher_thread
        if not want and th is not None and th.is_alive() and th is not threading.current_thread():
            th.join(timeout=3.0)
        return self.watcher_running

    def _bg_watch_loop(self) -> None:
        while not self._watcher_stop.is_set():
            try:
                res = self.run_watcher_cycle_once()
                errs: list[str] = []
                for p_res in (res.get("results") or {}).values():
                    if isinstance(p_res, dict):
                        if p_res.get("error"):
                            errs.append(str(p_res["error"]))
                        for e in (p_res.get("errors") or []):
                            errs.append(str(e))
                if errs:
                    self.last_watcher_error = {"at": now_iso(), "message": errs[0]}
                else:
                    self.last_watcher_error = None
            except Exception as exc:
                self.last_watcher_error = {"at": now_iso(), "message": str(exc)}
            poll = int(self.get_settings().get("poll_seconds") or 6)
            self._watcher_stop.wait(timeout=max(2, poll))

    # ── Batch Copy-Only Legacy Import ────────────────────────────────────────
    def scan_legacy_folder(self, pid: str, src_folder: str) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        sp = Path(str(src_folder or "").strip()).resolve()
        if not sp.exists() or not sp.is_dir():
            raise ValueError(_tr("server.folder_not_found", path=src_folder))
        state = read_json_safe(Path(proj["root"]) / "_System_Agent" / "watcher_state.json", {}) or {}
        done_hashes = set((state.get("hashes") or {}).keys()) | set((state.get("native_hashes") or {}).keys())
        supported = sw.SCANNED_EXTS | sw.NATIVE_DRAFT_EXTS | {".xlsm", ".txt"}
        total, pending = 0, 0
        for f in sp.rglob("*"):
            if f.is_file() and not f.name.startswith(("~$", ".")) and f.suffix.lower() in supported:
                total += 1
                try:
                    from security_utils import sha256_of_file
                    if sha256_of_file(f) not in done_hashes:
                        pending += 1
                except Exception:
                    pending += 1
        return {"ok": True, "folder": str(sp), "total_files": total, "already_imported": total - pending, "remaining": pending}

    def start_legacy_import(self, pid: str, src_folder: str, max_files: int = 30, pace_seconds: float = 4.0) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        sp = Path(str(src_folder or "").strip()).resolve()
        if not sp.exists() or not sp.is_dir():
            raise ValueError(_tr("server.folder_not_found", path=src_folder))
        stop_ev = threading.Event()
        with self.lock:
            if (self.legacy_jobs.get(pid) or {}).get("running"):
                return {"ok": False, "error": _tr("server.import_running")}
            self._legacy_stops[pid] = stop_ev
            self.legacy_jobs[pid] = {"running": True, "folder": str(sp), "started_at": now_iso()}

        def _worker() -> None:
            res = process_project_inbox_once(
                self.get_project(pid) or proj,
                all_projects=self.list_projects(),
                global_settings=self.get_settings(),
                decision_store=self.decisions,
                device_id=self.device_id,
                inbox_override=sp,
                copy_only=True,
                max_files=max(1, min(500, int(max_files))),
                pace_seconds=max(0.0, float(pace_seconds)),
                stop_event=stop_ev,
                analyzer_override=self.analyzer_override,
            )
            self.refresh_project_counters(pid)
            with self.lock:
                self.legacy_jobs[pid] = {"running": False, "folder": str(sp), "finished_at": now_iso(), "summary": res}

        threading.Thread(target=_worker, daemon=True).start()
        return {"ok": True, "started": True, "message": _tr("server.import_started", count=max_files)}

    def stop_legacy_import(self, pid: str) -> bool:
        with self.lock:
            ev = self._legacy_stops.get(pid)
            if ev:
                ev.set()
                return True
        return False

    # ── Large-Scale Project Organizer (1,000 - 10,000+ files) ────────────────
    def start_organizer_scan(
        self,
        pid: str,
        src_folder: str,
        *,
        extract_archives: bool = True,
        read_text: bool = True,
        blocking: bool = False,
    ) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        sp = resolve_local_or_unc(str(src_folder or "").strip())
        if not str(src_folder or "").strip() or not sp.exists() or not sp.is_dir():
            raise ValueError(_tr("server.folder_not_found", path=src_folder))
        stop_ev = threading.Event()
        with self.lock:
            if (self.organizer_jobs.get(pid) or {}).get("running"):
                return {"ok": False, "error": _tr("server.import_running")}
            self._organizer_stops[pid] = stop_ev
            now_ts = time.time()
            self.organizer_jobs[pid] = {
                "running": True,
                "kind": "scan",
                "phase": "walking",
                "scanned": 0,
                "discovered": 0,
                "archives_found": 0,
                "folders_scanned": 0,
                "total": 0,
                "current_file": sp.name,
                "started_at": now_iso(),
                "started_ts": now_ts,
                "updated_ts": now_ts,
            }

        def _on_prog(info: dict[str, Any]) -> None:
            with self.lock:
                cur = dict(self.organizer_jobs.get(pid) or {})
                cur.update(info)
                cur["updated_ts"] = time.time()
                self.organizer_jobs[pid] = cur

        def _work() -> dict[str, Any]:
            try:
                res = porg.build_organization_plan(
                    self.get_project(pid) or proj,
                    str(sp),
                    extract_archives=extract_archives,
                    read_text=read_text,
                    device_id=self.device_id,
                    stop_event=stop_ev,
                    progress_cb=_on_prog,
                )
                with self.lock:
                    prev_st = (self.organizer_jobs.get(pid) or {}).get("started_ts") or time.time()
                    was_stopped = bool(stop_ev.is_set() or (isinstance(res, dict) and res.get("stopped")))
                    self.organizer_jobs[pid] = {
                        "running": False,
                        "kind": "scan",
                        "phase": "stopped" if was_stopped else "done",
                        "stopped": was_stopped,
                        "started_ts": prev_st,
                        "finished_at": now_iso(),
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "summary": res,
                    }
                return res
            except Exception as exc:
                with self.lock:
                    prev_st = (self.organizer_jobs.get(pid) or {}).get("started_ts") or time.time()
                    self.organizer_jobs[pid] = {
                        "running": False,
                        "kind": "scan",
                        "phase": "error",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "error": str(exc),
                        "finished_at": now_iso(),
                    }
                if blocking:
                    raise
                return {}

        if blocking:
            res = _work()
            return {"ok": True, "summary": res}
        threading.Thread(target=_work, daemon=True).start()
        return {"ok": True, "started": True}

    def start_organizer_execute(
        self,
        pid: str,
        *,
        mode: str = "copy",
        update_excel_log: bool = True,
        blocking: bool = False,
    ) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        stop_ev = threading.Event()
        with self.lock:
            if (self.organizer_jobs.get(pid) or {}).get("running"):
                return {"ok": False, "error": _tr("server.import_running")}
            self._organizer_stops[pid] = stop_ev
            now_ts = time.time()
            self.organizer_jobs[pid] = {
                "running": True,
                "kind": "execute",
                "phase": "executing",
                "processed": 0,
                "total": 0,
                "copied": 0,
                "logged": 0,
                "current_file": "",
                "started_at": now_iso(),
                "started_ts": now_ts,
                "updated_ts": now_ts,
            }

        def _on_prog(info: dict[str, Any]) -> None:
            with self.lock:
                cur = dict(self.organizer_jobs.get(pid) or {})
                cur.update(info)
                cur["updated_ts"] = time.time()
                self.organizer_jobs[pid] = cur

        def _work() -> dict[str, Any]:
            try:
                res = porg.execute_organizer_plan(
                    self.get_project(pid) or proj,
                    mode=mode,
                    update_excel_log=update_excel_log,
                    device_id=self.device_id,
                    stop_event=stop_ev,
                    progress_cb=_on_prog,
                )
                self.refresh_project_counters(pid)
                with self.lock:
                    prev_st = (self.organizer_jobs.get(pid) or {}).get("started_ts") or time.time()
                    was_stopped = bool(stop_ev.is_set() or (isinstance(res, dict) and res.get("stopped")))
                    self.organizer_jobs[pid] = {
                        "running": False,
                        "kind": "execute",
                        "phase": "stopped" if was_stopped else "done",
                        "stopped": was_stopped,
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "summary": None if was_stopped else res,
                    }
                return res
            except Exception as exc:
                with self.lock:
                    prev_st = (self.organizer_jobs.get(pid) or {}).get("started_ts") or time.time()
                    self.organizer_jobs[pid] = {
                        "running": False,
                        "kind": "execute",
                        "phase": "error",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "error": str(exc),
                        "finished_at": now_iso(),
                    }
                if blocking:
                    raise
                return {}

        if blocking:
            res = _work()
            return {"ok": True, "execution": res}
        threading.Thread(target=_work, daemon=True).start()
        return {"ok": True, "started": True}

    def start_organizer_ai_rename(
        self,
        pid: str,
        *,
        item_ids: list[str] | None = None,
        scope: str = "unclassified",
        limit: int = 25,
        force_refresh: bool = False,
        blocking: bool = False,
    ) -> dict[str, Any]:
        proj = self.get_project(pid)
        if not proj:
            raise KeyError(_tr("server.project_not_found"))
        settings = self.get_settings()
        engine = str(settings.get("engine") or "gemini").strip().lower()
        ollama_url = str(settings.get("ollama_url") or "http://127.0.0.1:11434").strip()
        if ollama_url and not ollama_url.rstrip("/").endswith("/api/chat"):
            ollama_url = ollama_url.rstrip("/") + "/api/chat"
        ollama_model = str(settings.get("ollama_model") or "qwen2.5vl:7b").strip()

        stop_ev = threading.Event()
        with self.lock:
            if (self.organizer_jobs.get(pid) or {}).get("running"):
                return {"ok": False, "error": _tr("server.import_running")}
            self._organizer_stops[pid] = stop_ev
            now_ts = time.time()
            self.organizer_jobs[pid] = {
                "running": True,
                "kind": "ai_rename",
                "phase": "ai_renaming",
                "processed": 0,
                "total": len(item_ids) if item_ids else max(1, int(limit)),
                "updated": 0,
                "current_file": "",
                "started_at": now_iso(),
                "started_ts": now_ts,
                "updated_ts": now_ts,
            }

        def _on_prog(info: dict[str, Any]) -> None:
            with self.lock:
                cur = dict(self.organizer_jobs.get(pid) or {})
                cur.update(info)
                cur["updated_ts"] = time.time()
                self.organizer_jobs[pid] = cur

        def _work() -> dict[str, Any]:
            try:
                res = porg.ai_rename_organizer_items(
                    self.get_project(pid) or proj,
                    item_ids=item_ids,
                    scope=scope,
                    limit=limit,
                    engine=engine,
                    ollama_url=ollama_url,
                    ollama_model=ollama_model,
                    device_id=self.device_id,
                    force_refresh=force_refresh,
                    stop_event=stop_ev,
                    progress_cb=_on_prog,
                )
                with self.lock:
                    prev_st = (self.organizer_jobs.get(pid) or {}).get("started_ts") or time.time()
                    was_stopped = bool(stop_ev.is_set() or (isinstance(res, dict) and res.get("stopped")))
                    self.organizer_jobs[pid] = {
                        "running": False,
                        "kind": "ai_rename",
                        "phase": "stopped" if was_stopped else "done",
                        "stopped": was_stopped,
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "summary": res,
                    }
                return res
            except Exception as exc:
                with self.lock:
                    prev_st = (self.organizer_jobs.get(pid) or {}).get("started_ts") or time.time()
                    self.organizer_jobs[pid] = {
                        "running": False,
                        "kind": "ai_rename",
                        "phase": "error",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "error": str(exc),
                        "finished_at": now_iso(),
                    }
                if blocking:
                    raise
                return {}

        if blocking:
            res = _work()
            return res
        threading.Thread(target=_work, daemon=True).start()
        return {"ok": True, "started": True}

    def get_organizer_job_status(self, pid: str) -> dict[str, Any]:
        with self.lock:
            job = dict(self.organizer_jobs.get(pid) or {"running": False})
        if job.get("running") and job.get("started_ts"):
            job["elapsed_seconds"] = max(0, int(time.time() - float(job["started_ts"])))
            if job.get("updated_ts"):
                job["last_update_ago"] = max(0, int(time.time() - float(job["updated_ts"])))
        return job

    def stop_organizer_job(self, pid: str) -> bool:
        with self.lock:
            ev = self._organizer_stops.get(pid)
            if ev:
                ev.set()
                return True
        return False

    def dismiss_organizer_job(self, pid: str) -> None:
        with self.lock:
            cur = self.organizer_jobs.get(pid)
            if cur and not cur.get("running"):
                self.organizer_jobs.pop(pid, None)

    # ── Scoped Background Jobs (Duplicates & WhatsApp Archiver) ──────────────
    def _scope_maps(self, scope: str) -> tuple[dict[str, dict[str, Any]], dict[str, threading.Event]]:
        if scope == "dup":
            return self.dup_jobs, self._dup_stops
        if scope == "wa":
            return self.wa_jobs, self._wa_stops
        return self.organizer_jobs, self._organizer_stops

    def get_scoped_job_status(self, scope: str, pid: str) -> dict[str, Any]:
        jobs_map, _ = self._scope_maps(scope)
        with self.lock:
            job = dict(jobs_map.get(pid) or {"running": False})
        if job.get("running") and job.get("started_ts"):
            job["elapsed_seconds"] = max(0, int(time.time() - float(job["started_ts"])))
            if job.get("updated_ts"):
                job["last_update_ago"] = max(0, int(time.time() - float(job["updated_ts"])))
        return job

    def stop_scoped_job(self, scope: str, pid: str) -> bool:
        _, stops_map = self._scope_maps(scope)
        with self.lock:
            ev = stops_map.get(pid)
            if ev:
                ev.set()
                return True
        return False

    def dismiss_scoped_job(self, scope: str, pid: str) -> None:
        jobs_map, _ = self._scope_maps(scope)
        with self.lock:
            cur = jobs_map.get(pid)
            if cur and not cur.get("running"):
                jobs_map.pop(pid, None)

    def _run_scoped_job(
        self,
        scope: str,
        pid: str,
        kind: str,
        initial_phase: str,
        fn: Callable[[threading.Event, Callable[[dict[str, Any]], None]], dict[str, Any]],
        *,
        blocking: bool = True,
        refresh_counters: bool = False,
        total_hint: int = 0,
    ) -> dict[str, Any]:
        if not self.get_project(pid):
            raise KeyError(_tr("server.project_not_found"))
        jobs_map, stops_map = self._scope_maps(scope)
        stop_ev = threading.Event()
        now_ts = time.time()
        with self.lock:
            if (jobs_map.get(pid) or {}).get("running"):
                return {"ok": False, "error": _tr("server.import_running")}
            stops_map[pid] = stop_ev
            jobs_map[pid] = {
                "running": True,
                "scope": scope,
                "kind": kind,
                "phase": initial_phase,
                "processed": 0,
                "total": int(total_hint or 0),
                "discovered": 0,
                "updated": 0,
                "current_file": "",
                "started_at": now_iso(),
                "started_ts": now_ts,
                "updated_ts": now_ts,
                "elapsed_seconds": 0,
                "error": "",
            }

        def _on_prog(info: dict[str, Any]) -> None:
            with self.lock:
                cur = dict(jobs_map.get(pid) or {})
                cur.update(info)
                cur["updated_ts"] = time.time()
                jobs_map[pid] = cur

        def _work() -> dict[str, Any]:
            try:
                res = fn(stop_ev, _on_prog)
                if refresh_counters:
                    self.refresh_project_counters(pid)
                was_stopped = bool(stop_ev.is_set() or (isinstance(res, dict) and res.get("stopped")))
                with self.lock:
                    prev_st = (jobs_map.get(pid) or {}).get("started_ts") or now_ts
                    jobs_map[pid] = {
                        "running": False,
                        "scope": scope,
                        "kind": kind,
                        "phase": "stopped" if was_stopped else "done",
                        "stopped": was_stopped,
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "summary": res,
                        "error": "",
                    }
                return res
            except (dups.DuplicatesStoppedError, waa.WhatsappStoppedError):
                with self.lock:
                    prev = dict(jobs_map.get(pid) or {})
                    prev_st = prev.get("started_ts") or now_ts
                    jobs_map[pid] = {
                        "running": False,
                        "scope": scope,
                        "kind": kind,
                        "phase": "stopped",
                        "stopped": True,
                        "processed": prev.get("processed", 0),
                        "total": prev.get("total", 0),
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "error": "",
                    }
                return {"ok": True, "stopped": True}
            except Exception as exc:
                with self.lock:
                    prev_st = (jobs_map.get(pid) or {}).get("started_ts") or now_ts
                    jobs_map[pid] = {
                        "running": False,
                        "scope": scope,
                        "kind": kind,
                        "phase": "error",
                        "started_ts": prev_st,
                        "elapsed_seconds": max(0, int(time.time() - float(prev_st))),
                        "finished_at": now_iso(),
                        "error": str(exc),
                    }
                if blocking:
                    raise
                return {"ok": False, "error": str(exc)}

        if blocking:
            return _work()
        threading.Thread(target=_work, daemon=True).start()
        return {"ok": True, "started": True, "job": self.get_scoped_job_status(scope, pid)}

    # ── Contacts (for WhatsApp quick-send) ───────────────────────────────────
    def save_contact(self, pid: str, name: str, phone: str, role: str = "") -> list[dict[str, Any]]:
        clean_ph = re.sub(r"[^\d+]", "", str(phone or ""))
        if not name.strip() or not clean_ph:
            raise ValueError(_tr("server.contact_required"))
        with self.lock:
            proj = self._get_project_ref_unlocked(pid)
            if not proj:
                raise KeyError(_tr("server.project_not_found"))
            contacts = [c for c in (proj.get("contacts") or []) if c.get("phone") != clean_ph]
            contacts.append({"name": name.strip(), "phone": clean_ph, "role": role.strip()})
            proj["contacts"] = contacts
            self.save()
            return list(contacts)

    def delete_contact(self, pid: str, phone: str) -> list[dict[str, Any]]:
        clean_ph = re.sub(r"[^\d+]", "", str(phone or ""))
        with self.lock:
            proj = self._get_project_ref_unlocked(pid)
            if not proj:
                raise KeyError(_tr("server.project_not_found"))
            proj["contacts"] = [c for c in (proj.get("contacts") or []) if c.get("phone") != clean_ph]
            self.save()
            return list(proj["contacts"])


# ══════════════════════════════════════════════════════════════════════════════
# Template Issuance + Optional AI Field Suggestion
# ══════════════════════════════════════════════════════════════════════════════
def issue_template_document(
    hub: HubState,
    proj: dict[str, Any],
    tpl_name: str,
    payload: dict[str, Any],
) -> dict[str, Any]:
    """
    Deterministic end-to-end issuance from an UPLOADED project template:
      1) Load & SHA-256-verify `<tpl_name>.tplmap.json` (auto-rebind anchors; block on orphan fields!)
      2) Compute exact next reference directly from the project's uploaded Excel log (`max(log)+1`)
      3) Validate all inputs & run `fill_by_profile` (zero guessing, formulas protected, styles & VBA kept)
      4) Save filled editable file into `2_Archive/<section_folder>/00_Editable_Originals/`
      5) Immediately update the project's OWN uploaded Excel log with status `Under Review` + `device_id`!
    """
    proj_root = Path(str(proj["root"])).resolve()
    tpl_path = tprof.resolve_project_template(proj_root, tpl_name)
    profile = tprof.get_template_profile(proj_root, tpl_path.name, device_id=hub.device_id)

    if not profile.get("has_saved_profile") or not profile.get("section_confirmed"):
        raise tprof.ProfileValidationError(
            _tr("server.tpl_not_ready"),
            errors=[{"key": "profile", "label": _tr("templates.title"), "error": _tr("server.tpl_not_ready")}],
        )
    if profile.get("orphan_fields"):
        raise tprof.ProfileOrphanError(
            _tr("server.tpl_moved_fields", count=len(profile["orphan_fields"]), name=tpl_path.name),
            orphan_fields=profile["orphan_fields"],
        )

    sec_id = str(profile["section"])
    proj_sections = tprof.get_project_sections(proj, include_hidden=True)
    sec_map = {s["id"]: s for s in proj_sections}
    sec_meta = sec_map.get(sec_id) or tprof.SECTIONS_BY_ID[sec_id]
    values_in = payload.get("values") if isinstance(payload.get("values"), dict) else payload

    disc = str(
        payload.get("discipline")
        or values_in.get("discipline")
        or profile.get("default_discipline")
        or ""
    ).strip().upper()
    rev = str(payload.get("revision") or values_in.get("revision") or "R00").strip()
    pfx = str(profile.get("ref_prefix") or sec_meta["ref_prefix"] or "DOC").strip().upper()

    explicit_ref = str(payload.get("ref_no") or values_in.get("ref_no") or "").strip()
    if explicit_ref:
        ref_no = explicit_ref
    else:
        counter_fld = next((f for f in (profile.get("fields") or []) if f.get("role") == "auto_counter" and not f.get("counter_sync_log")), None)
        if counter_fld is not None:
            ref_no = tprof.format_counter_value(counter_fld)
        else:
            ref_no, _, _ = plog.next_reference_from_logs(
                proj,
                prefix=pfx,
                discipline=disc,
                use_discipline_series=bool(profile.get("use_discipline_series", False)),
                device_id=hub.device_id,
            )

    today_iso = str(payload.get("doc_date") or values_in.get("doc_date") or dt.date.today().isoformat()).strip()
    sys_ctx = {"next_ref": ref_no, "revision": rev, "today_iso": today_iso}

    # Pre-validate before creating target folder
    resolved = tprof.validate_and_resolve_payload(profile, payload, proj, system_ctx=sys_ctx)
    fvals = resolved["fields"]

    subject = str(
        payload.get("subject")
        or fvals.get("subject")
        or next((v for k, v in fvals.items() if isinstance(v, str) and v.strip() and k not in ("project_name", "contract_no", "owner", "consultant", "contractor", "ref_no", "doc_date", "revision", "approval_code", "discipline")), "")
        or profile.get("display_name")
        or tpl_path.stem
    ).strip()

    safe_ref = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", ref_no).strip(" ._-") or "DOC"
    safe_subj = _clean_folder_seg(subject, 50).replace(" ", "_")
    archive_root = archive_dir_for_project(proj)
    out_dir = archive_root / sec_meta["archive_folder"] / "00_Editable_Originals"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_name = f"{safe_ref}_{today_iso}_{safe_subj}{tpl_path.suffix.lower()}"
    out_path = out_dir / out_name
    idx = 2
    while out_path.exists():
        out_path = out_dir / f"{safe_ref}_{today_iso}_{safe_subj}_{idx}{tpl_path.suffix.lower()}"
        idx += 1

    try:
        fill_res = tprof.fill_by_profile(
            tpl_path,
            profile,
            payload,
            proj,
            out_path,
            system_ctx=sys_ctx,
            verify_drift=True,
            project_sections=proj_sections,
        )
        if not out_path.exists() or not out_path.is_file() or out_path.stat().st_size == 0:
            raise RuntimeError(_tr("server.file_not_found"))
    except Exception:
        out_path.unlink(missing_ok=True)
        raise
    # Increment any custom auto_counter fields stored inside this template's .tplmap.json
    advanced_counters = tprof.advance_template_counters(proj_root, tpl_path.name, device_id=hub.device_id)

    try:
        rel_from_root = out_path.relative_to(proj_root).as_posix()
        out_area = "root"
        dl_url = f"/api/projects/{proj['id']}/download?rel={quote(rel_from_root, safe='/')}"
    except ValueError:
        rel_from_root = out_path.relative_to(archive_root).as_posix()
        out_area = "archive"
        dl_url = f"/api/projects/{proj['id']}/download?area=archive&rel={quote(rel_from_root, safe='/')}"

    # Optional ready-to-print/send PDF export + attachment PDF/image merging (Feature 2)
    pdf_info: dict[str, Any] = {
        "pdf_name": "",
        "pdf_rel": "",
        "pdf_area": out_area,
        "pdf_download_url": "",
        "pdf_preview_url": "",
        "pdf_pages": 0,
        "pdf_engine": "",
    }
    if bool(payload.get("export_pdf") or payload.get("also_pdf")):
        out_pdf = out_path.with_suffix(".pdf")
        p_idx = 2
        while out_pdf.exists():
            out_pdf = out_dir / f"{out_path.stem}_{p_idx}.pdf"
            p_idx += 1
        summary_lines = [
            f"{sec_meta['label']} | {ref_no} ({rev}) | {today_iso}",
            f"{proj.get('name') or ''} — {subject}",
        ]
        for fk, fv in fvals.items():
            if fv not in (None, ""):
                summary_lines.append(f"{fk}: {fv}")

        raw_atts = payload.get("attachments") if isinstance(payload.get("attachments"), list) else []
        valid_ext_atts = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
        with tempfile.TemporaryDirectory(prefix="sla_issue_pdf_") as td_pdf:
            td_p = Path(td_pdf)
            base_pdf = td_p / "base_issued.pdf"
            conv = trender.convert_office_to_pdf(
                out_path,
                base_pdf,
                title_hint=f"{ref_no} — {subject}",
                summary_lines=summary_lines,
            )
            att_paths: list[Path] = []
            for ai, att in enumerate(raw_atts[:12]):
                if not isinstance(att, dict) or not att.get("data_b64"):
                    continue
                att_name = sanitize_upload_filename(str(att.get("name") or f"attach_{ai+1}.pdf"))
                att_ext = Path(att_name).suffix.lower()
                if att_ext not in valid_ext_atts:
                    continue
                raw_b = base64.b64decode(str(att["data_b64"]), validate=False)
                if len(raw_b) > 20 * 1024 * 1024:
                    continue
                att_p = td_p / f"att_{ai:02d}_{att_name}"
                att_p.write_bytes(raw_b)
                att_paths.append(att_p)

            arc_rels = payload.get("archive_attachment_rels") if isinstance(payload.get("archive_attachment_rels"), list) else []
            for rel_it in arc_rels[:8]:
                if isinstance(rel_it, dict):
                    r_str = str(rel_it.get("rel") or "").strip()
                    r_area = str(rel_it.get("area") or "root").strip()
                else:
                    r_str = str(rel_it or "").strip()
                    r_area = "root"
                if not r_str:
                    continue
                base_b = archive_root if r_area == "archive" else proj_root
                try:
                    cand_a = safe_resolve_within(base_b, r_str, allow_base=False, reject_symlinks=True)
                    if cand_a.exists() and cand_a.is_file() and cand_a.suffix.lower() in valid_ext_atts:
                        att_paths.append(cand_a)
                except PathTraversalError:
                    continue

            if att_paths:
                m_res = trender.merge_pdf_and_images(base_pdf, att_paths, out_pdf)
                tot_pages = int(m_res.get("pages") or 1)
            else:
                shutil.copy2(base_pdf, out_pdf)
                tot_pages = int(conv.get("pages") or trender.count_pdf_pages(out_pdf))

        if out_pdf.exists() and out_pdf.is_file() and out_pdf.stat().st_size > 0:
            try:
                pdf_rel = out_pdf.relative_to(proj_root).as_posix()
                pdf_area = "root"
                pdf_dl = f"/api/projects/{proj['id']}/download?rel={quote(pdf_rel, safe='/')}"
                pdf_prev = f"/api/projects/{proj['id']}/preview-info?rel={quote(pdf_rel, safe='/')}"
            except ValueError:
                pdf_rel = out_pdf.relative_to(archive_root).as_posix()
                pdf_area = "archive"
                pdf_dl = f"/api/projects/{proj['id']}/download?area=archive&rel={quote(pdf_rel, safe='/')}"
                pdf_prev = f"/api/projects/{proj['id']}/preview-info?area=archive&rel={quote(pdf_rel, safe='/')}"
            pdf_info = {
                "pdf_name": out_pdf.name,
                "pdf_rel": pdf_rel,
                "pdf_area": pdf_area,
                "pdf_download_url": pdf_dl,
                "pdf_preview_url": pdf_prev,
                "pdf_pages": tot_pages,
                "pdf_engine": conv.get("engine") or "",
            }
    try:
        rel_from_logs = Path(os.path.relpath(out_path, plog.logs_dir_for_project(proj_root))).as_posix()
    except ValueError:
        rel_from_logs = str(out_path)

    # Update the project's OWN uploaded Excel log immediately!
    log_res: dict[str, Any] = {"skipped": True}
    if profile.get("log_on_issue", True):
        status_val = str(
            payload.get("status")
            or fvals.get("approval_code")
            or profile.get("default_status")
            or "pending"
        ).strip()
        amount_val = payload.get("amount") if payload.get("amount") not in (None, "") else fvals.get("amount")
        # If repeating table has numeric amounts and top-level amount empty, sum them cleanly
        if amount_val in (None, ""):
            total_sum = 0.0
            found_money = False
            for trows in (fill_res.get("resolved_tables") or {}).values():
                for tr in trows:
                    if isinstance(tr.get("amount"), (int, float)):
                        total_sum += float(tr["amount"])
                        found_money = True
                    elif isinstance(tr.get("qty"), (int, float)) and isinstance(tr.get("unit_price"), (int, float)):
                        total_sum += float(tr["qty"]) * float(tr["unit_price"])
                        found_money = True
            if found_money:
                amount_val = round(total_sum, 3)

        log_entry = {
            "section": sec_id,
            "subtype": str(payload.get("subtype") or sec_meta.get("subtype") or pfx),
            "discipline": disc,
            "ref_no": ref_no,
            "revision": rev,
            "doc_date": today_iso,
            "subject_ar": subject,
            "party_name": str(payload.get("party_name") or fvals.get("party_name") or ""),
            "party_role": str(payload.get("party_role") or ""),
            "status": status_val,
            "amount": amount_val,
            "notes": str(payload.get("notes") or fvals.get("notes") or ""),
            "file_rel_path": rel_from_logs,
            "target_log": str(payload.get("target_log") or profile.get("target_log") or ""),
            "target_sheet": str(payload.get("target_sheet") or profile.get("target_sheet") or ""),
        }
        log_res = plog.record_entry_in_project_log(proj, log_entry, device_id=hub.device_id, queue_if_locked=True)
        hub.refresh_project_counters(str(proj["id"]))

    link_res: dict[str, Any] | None = None
    link_targets = payload.get("link_to_refs")
    if isinstance(link_targets, list) and link_targets and sec_id in ("incoming", "outgoing"):
        try:
            lr_d = log_res if isinstance(log_res, dict) else {}
            new_letter_item = {
                "ref_no": ref_no,
                "sheet": str(lr_d.get("sheet") or payload.get("target_sheet") or profile.get("target_sheet") or ""),
                "row": lr_d.get("row"),
                "log_name": str(lr_d.get("log_name") or ""),
                "section": sec_id,
                "party_role": str(payload.get("party_role") or ""),
                "party_name": str(payload.get("party_name") or fvals.get("party_name") or ""),
                "subject": subject,
                "doc_date": today_iso,
            }
            link_res = llinks.link_letters(
                proj,
                [new_letter_item, *link_targets],
                sync_excel=True,
                device_id=hub.device_id,
            )
        except Exception:
            link_res = None

    try:
        hub.refresh_project_counters(str(proj.get("id") or ""))
    except Exception:
        pass

    return {
        "ok": True,
        "ref_no": ref_no,
        "revision": rev,
        "section": sec_id,
        "section_label": sec_meta["label"],
        "output_name": out_path.name,
        "output_rel": rel_from_root,
        "output_area": out_area,
        "output_path": str(out_path),
        "download_url": dl_url,
        **pdf_info,
        "resolved_fields": fill_res["resolved_fields"],
        "touched_targets": fill_res["touched_targets"],
        "rebound_fields": fill_res["rebound_fields"],
        "advanced_counters": advanced_counters,
        "log_result": log_res,
        "link_result": link_res,
        "device_id": hub.device_id,
    }


def suggest_fields_with_ai(
    proj_root: Path,
    tpl_name: str,
    engine: str = "gemini",
    ollama_url: str = "",
    ollama_model: str = "",
    *,
    force_refresh: bool = False,
) -> dict[str, Any]:
    """Optional AI suggestion layer — always tagged `[مقترح]` and NEVER applied without user Save!"""
    tpl_path = tprof.resolve_project_template(proj_root, tpl_name)
    det = tprof.suggest_fields_deterministic(tpl_path)
    cache_dir = ai_cache.project_ai_cache_dir(proj_root)
    tpl_sha = sha256_of_file(tpl_path)
    if not force_refresh and tpl_sha:
        cached = ai_cache.get_cached(cache_dir, "template_ai", tpl_sha)
        if cached and isinstance(cached.get("ai_hints"), list) and cached["ai_hints"]:
            return {
                "ok": True,
                "ai_used": True,
                "from_cache": True,
                "cached_at": str(cached.get("_cached_at") or ""),
                "suggested_fields": det["suggested_fields"],
                "suggested_tables": det["suggested_tables"],
                "protected_cells": det["protected_cells"],
                "ai_hints": list(cached["ai_hints"]),
            }

    txt = sw.extract_native_office_text(tpl_path)[:3500]
    ai_extra: list[dict[str, Any]] = []
    ai_used = False
    prompt = (
        f"افحص هيكل النموذج الهندسي التالي ({tpl_path.name}) واقترح الحقول القابلة للتعبئة بصيغة JSON فقط:\n"
        f"{txt}\n"
        'أرجع: {"suggested_section": "submittal_mar|submittal_mir|submittal_ir|submittal_sd|submittal_ms|outgoing|contract_ipc|variation_vo|purchase_po|petty_cash|site_report|general", '
        '"fields": [{"label": "...", "key": "...", "role": "manual|auto_project|auto_system", "input_type": "text|textarea|date|number|kwd|dropdown|approval_code", "target_hint": "C4 أو T0:R1C1"}]}'
    )
    try:
        if engine == "ollama" and ollama_url:
            res = sw.analyze_with_ollama(prompt, [], ollama_url, ollama_model or "qwen2.5vl:7b")
            ai_used = True
        elif sw.load_saved_gemini_key():
            res = sw.analyze_with_gemini(prompt, tpl_path.parent, sw.load_saved_gemini_model(), images=[])
            ai_used = True
        else:
            res = {}
        for item in (res.get("fields") or []):
            if isinstance(item, dict) and item.get("label"):
                item["suggested"] = True
                item["ai_suggested"] = True
                ai_extra.append(item)
    except Exception:
        ai_used = False

    if ai_used and ai_extra and tpl_sha:
        ai_cache.set_cached(cache_dir, "template_ai", tpl_sha, {"ai_hints": ai_extra})

    return {
        "ok": True,
        "ai_used": ai_used,
        "from_cache": False,
        "suggested_fields": det["suggested_fields"],
        "suggested_tables": det["suggested_tables"],
        "protected_cells": det["protected_cells"],
        "ai_hints": ai_extra,
    }


# ══════════════════════════════════════════════════════════════════════════════
# Desktop Helpers & State Builders
# ══════════════════════════════════════════════════════════════════════════════
def open_on_desktop(target: Path) -> bool:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(target))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(target)])
        else:
            subprocess.Popen(["xdg-open", str(target)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


def browse_local_folder(initial_dir: str = "", title: str = "") -> tuple[bool, str, str]:
    """Open native OS folder picker in foreground/topmost (supports Windows Server UNC shares & Arabic/Unicode paths)."""
    init_p = _resolve_picker_initial_dir(initial_dir)
    if sys.platform.startswith("win"):
        try:
            env = os.environ.copy()
            env["SLA_PICK_INIT"] = init_p
            env["SLA_PICK_TITLE"] = str(title or "Select Folder")
            # Modern OpenFileDialog folder trick shows Network / Windows Server UNC shares and address bar.
            # Attach an invisible TopMost centered owner Form so the dialog always opens IN FRONT of the browser/app window.
            # Return UTF-8 bytes encoded as Base64 (`B64:...`) so Windows OEM/ANSI console codepages never mangle Arabic names to `????`.
            ps_script = (
                "[Console]::InputEncoding = [System.Text.Encoding]::UTF8;"
                "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8;"
                "Add-Type -AssemblyName System.Windows.Forms;"
                "$owner = New-Object System.Windows.Forms.Form;"
                "$owner.TopMost = $true;"
                "$owner.ShowInTaskbar = $false;"
                "$owner.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None;"
                "$owner.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;"
                "$owner.Width = 1; $owner.Height = 1; $owner.Opacity = 0;"
                "$owner.Add_Shown({ $this.Activate(); $this.BringToFront() });"
                "$owner.Show();"
                "[System.Windows.Forms.Application]::DoEvents();"
                "$owner.Activate();"
                "try {"
                "  $d = New-Object System.Windows.Forms.OpenFileDialog;"
                "  $d.ValidateNames = $false;"
                "  $d.CheckFileExists = $false;"
                "  $d.CheckPathExists = $true;"
                "  $d.FileName = 'Folder Selection';"
                "  if ($env:SLA_PICK_TITLE) { $d.Title = $env:SLA_PICK_TITLE };"
                "  if ($env:SLA_PICK_INIT -and (Test-Path -LiteralPath $env:SLA_PICK_INIT)) { $d.InitialDirectory = $env:SLA_PICK_INIT };"
                "  $res = $d.ShowDialog($owner);"
                "  if ($res -eq [System.Windows.Forms.DialogResult]::OK) {"
                "    $p = [System.IO.Path]::GetDirectoryName($d.FileName);"
                "    if ($p) { Write-Output ('B64:' + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($p))) }"
                "  }"
                "} finally { $owner.Close(); $owner.Dispose() }"
            )
            try:
                import ctypes
                ctypes.windll.user32.AllowSetForegroundWindow(-1)  # type: ignore[attr-defined]
            except Exception:
                pass
            kw: dict[str, Any] = {}
            if hasattr(subprocess, "CREATE_NO_WINDOW"):
                kw["creationflags"] = subprocess.CREATE_NO_WINDOW
            cp = subprocess.run(
                ["powershell", "-NoProfile", "-STA", "-Command", ps_script],
                env=env, capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=90, check=False, **kw,
            )
            chosen = _decode_ps_path_output(cp.stdout)
            if chosen:
                return True, chosen, ""
            return False, "", ""
        except Exception:
            pass
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
            root.update()
            root.lift()
            root.focus_force()
        except Exception:
            pass
        try:
            chosen = filedialog.askdirectory(parent=root, initialdir=init_p, title=title or "Select Folder")
        finally:
            root.destroy()
        return (True, str(chosen), "") if chosen else (False, "", "")
    except Exception:
        return False, "", _tr("new_project.browse_unavailable")


def browse_local_file(
    initial_dir: str = "",
    title: str = "",
    file_filter: str = "Excel Logs (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*",
) -> tuple[bool, str, str]:
    """Open native OS file picker in foreground/topmost (supports local drives, UNC shares & Arabic/Unicode paths)."""
    init_p = _resolve_picker_initial_dir(initial_dir)
    if sys.platform.startswith("win"):
        try:
            env = os.environ.copy()
            env["SLA_PICK_INIT"] = init_p
            env["SLA_PICK_TITLE"] = str(title or "Select Excel Log")
            env["SLA_PICK_FILTER"] = str(file_filter)
            ps_script = (
                "[Console]::InputEncoding = [System.Text.Encoding]::UTF8;"
                "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8;"
                "Add-Type -AssemblyName System.Windows.Forms;"
                "$owner = New-Object System.Windows.Forms.Form;"
                "$owner.TopMost = $true;"
                "$owner.ShowInTaskbar = $false;"
                "$owner.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None;"
                "$owner.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;"
                "$owner.Width = 1; $owner.Height = 1; $owner.Opacity = 0;"
                "$owner.Add_Shown({ $this.Activate(); $this.BringToFront() });"
                "$owner.Show();"
                "[System.Windows.Forms.Application]::DoEvents();"
                "$owner.Activate();"
                "try {"
                "  $d = New-Object System.Windows.Forms.OpenFileDialog;"
                "  if ($env:SLA_PICK_TITLE) { $d.Title = $env:SLA_PICK_TITLE };"
                "  if ($env:SLA_PICK_FILTER) { $d.Filter = $env:SLA_PICK_FILTER };"
                "  if ($env:SLA_PICK_INIT -and (Test-Path -LiteralPath $env:SLA_PICK_INIT)) { $d.InitialDirectory = $env:SLA_PICK_INIT };"
                "  if ($d.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK -and $d.FileName) {"
                "    Write-Output ('B64:' + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($d.FileName)))"
                "  }"
                "} finally { $owner.Close(); $owner.Dispose() }"
            )
            try:
                import ctypes
                ctypes.windll.user32.AllowSetForegroundWindow(-1)  # type: ignore[attr-defined]
            except Exception:
                pass
            kw: dict[str, Any] = {}
            if hasattr(subprocess, "CREATE_NO_WINDOW"):
                kw["creationflags"] = subprocess.CREATE_NO_WINDOW
            cp = subprocess.run(
                ["powershell", "-NoProfile", "-STA", "-Command", ps_script],
                env=env, capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=90, check=False, **kw,
            )
            chosen = _decode_ps_path_output(cp.stdout)
            if chosen:
                return True, chosen, ""
            return False, "", ""
        except Exception:
            pass
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
            root.update()
            root.lift()
            root.focus_force()
        except Exception:
            pass
        tk_types = [("PDF Files", "*.pdf"), ("All files", "*.*")] if "*.pdf" in file_filter else [("Excel Log", "*.xlsx *.xlsm"), ("All files", "*.*")]
        try:
            chosen = filedialog.askopenfilename(
                parent=root,
                initialdir=init_p,
                title=title or "Select File",
                filetypes=tk_types,
            )
        finally:
            root.destroy()
        return (True, str(chosen), "") if chosen else (False, "", "")
    except Exception:
        return False, "", _tr("new_project.browse_unavailable")


def launch_gemini_in_project(proj: dict[str, Any], initial_prompt: str = "") -> tuple[bool, str]:
    root = resolve_local_or_unc(str(proj.get("root") or ""))
    if not root.exists():
        return False, _tr("server.folder_not_found", path=str(root))
    try:
        if sys.platform.startswith("win"):
            gemini_cmd = "gemini"
            if initial_prompt.strip():
                safe_p = re.sub(r'[&|<>^%\r\n;"\x00`]', " ", initial_prompt).strip()[:300]
                if safe_p:
                    gemini_cmd += f' -i "{safe_p}"'
            subprocess.Popen(["cmd", "/c", "start", "", "/D", str(root), "cmd", "/k", gemini_cmd])
            return True, _tr("server.helper_opened")
        return False, _tr("server.windows_only_helper")
    except Exception:
        return False, _tr("common.error_generic")


def _find_windows_app_browser() -> str:
    if not sys.platform.startswith("win"):
        return ""
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    la = os.environ.get("LOCALAPPDATA", "")
    candidates = [
        Path(pf86) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(pf) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(pf) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(pf86) / "Google" / "Chrome" / "Application" / "chrome.exe",
    ]
    if la:
        candidates.extend([
            Path(la) / "Google" / "Chrome" / "Application" / "chrome.exe",
            Path(la) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        ])
    for c in candidates:
        if c.exists():
            return str(c)
    return ""


def open_standalone_app_window(url: str) -> dict[str, Any]:
    """Open `url` in a standalone Chromium/Edge `--app` window (fallback to default browser)."""
    browser_exe = _find_windows_app_browser()
    if browser_exe:
        try:
            subprocess.Popen([browser_exe, f"--app={url}"])
            return {"ok": True, "mode": "standalone_app", "browser": Path(browser_exe).name}
        except Exception:
            pass
    try:
        webbrowser.open(url)
        return {"ok": True, "mode": "browser_tab", "browser": "default"}
    except Exception as exc:
        return {"ok": False, "mode": "none", "error": str(exc)}


def create_desktop_app_shortcut() -> dict[str, Any]:
    """Create a silent Desktop launcher icon (`Smart Projects Hub.lnk`) on Windows."""
    if not sys.platform.startswith("win"):
        return {
            "ok": True,
            "platform": sys.platform,
            "mode": "browser_pwa",
            "shortcut_path": "",
        }
    gemini_vbs = Path.home() / ".gemini" / "Open_Projects_Silent.vbs"
    bundled_vbs = SCRIPT_DIR.parent / "0_OPEN_SILENT.vbs"
    if bundled_vbs.exists() and (Path.home() / ".gemini").exists():
        try:
            shutil.copy2(bundled_vbs, gemini_vbs)
        except Exception:
            pass
    target_vbs = gemini_vbs if gemini_vbs.exists() else bundled_vbs
    icon_file = SCRIPT_DIR / "app_icon.ico"
    gemini_icon = Path.home() / ".gemini" / "doc_tools" / "app_icon.ico"
    if icon_file.exists() and (Path.home() / ".gemini" / "doc_tools").exists():
        try:
            shutil.copy2(icon_file, gemini_icon)
        except Exception:
            pass
    use_icon = gemini_icon if gemini_icon.exists() else icon_file

    if not target_vbs.exists():
        return {"ok": False, "error": _tr("server.not_found")}

    env = os.environ.copy()
    env["SPH_VBS"] = str(target_vbs)
    env["SPH_WDIR"] = str(target_vbs.parent)
    env["SPH_ICON"] = str(use_icon) if use_icon.exists() else ""

    ps_script = (
        "$dt=[Environment]::GetFolderPath('Desktop'); "
        "$ws=New-Object -ComObject WScript.Shell; "
        "$sc=$ws.CreateShortcut((Join-Path $dt 'Smart Projects Hub.lnk')); "
        "$sc.TargetPath='wscript.exe'; "
        "$sc.Arguments='//nologo \"' + $env:SPH_VBS + '\"'; "
        "$sc.WorkingDirectory=$env:SPH_WDIR; "
        "if($env:SPH_ICON -and (Test-Path -LiteralPath $env:SPH_ICON)){$sc.IconLocation=$env:SPH_ICON}; "
        "$sc.Description='Smart Projects Hub v10.13.2'; "
        "$sc.Save(); "
        "Write-Output ('B64:' + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($sc.FullName)))"
    )
    kw: dict[str, Any] = {}
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        kw["creationflags"] = subprocess.CREATE_NO_WINDOW
    res = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps_script],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=15, env=env, **kw,
    )
    if res.returncode != 0:
        return {"ok": False, "error": res.stderr.strip() or "PowerShell shortcut failed"}
    return {
        "ok": True,
        "platform": "win32",
        "mode": "desktop_shortcut",
        "shortcut_path": _decode_ps_path_output(res.stdout),
    }


def build_project_card(hub: HubState, proj: dict[str, Any], gcfg: dict[str, Any]) -> dict[str, Any]:
    proj_root = Path(str(proj["root"])).resolve()
    proj_sections = tprof.get_project_sections(proj, style=gcfg.get("ui_style"), include_hidden=True)
    log_sum = plog.read_project_log_summary(proj, limit=0, device_id=hub.device_id)
    tpls = tprof.list_project_templates(proj_root, device_id=hub.device_id, project_sections=proj_sections) if proj_root.exists() else []
    bkp = check_backup_status(proj, gcfg, warn_days=int(gcfg.get("backup_warn_days") or DEFAULT_BACKUP_WARN_DAYS))
    job = hub.jobs.get(proj["id"]) or {}
    wstate = read_json_safe(proj_root / "_System_Agent" / "watcher_state.json", {}) if proj_root.exists() else {}
    tstate = read_json_safe(proj_root / "_System_Agent" / "scanner_triage.json", {}) if proj_root.exists() else {}
    wa_state = read_json_safe(proj_root / "_System_Agent" / "whatsapp_archive_plan.json", {}) if proj_root.exists() else {}
    dup_state = read_json_safe(proj_root / "_System_Agent" / "duplicates_plan.json", {}) if proj_root.exists() else {}
    t_rej = set((tstate or {}).get("rejected_hashes") or [])
    t_app = set((tstate or {}).get("approved_hashes") or [])
    triage_pending = sum(
        1 for it in ((tstate or {}).get("items") or [])
        if it.get("status", "pending") == "pending"
        and str(it.get("file_sha256") or "") not in t_rej
        and str(it.get("file_sha256") or "") not in t_app
        and not (it.get("file_sha256") and hub.decisions.is_ignored_hash(str(it.get("file_sha256"))))
        and (not it.get("file_path") or Path(str(it.get("file_path"))).exists())
    )
    wa_plan_count = len((wa_state or {}).get("items") or [])
    dup_groups_count = sum(1 for g in ((dup_state or {}).get("groups") or []) if not g.get("reviewed"))
    w_errs = list((wstate or {}).get("last_errors") or [])
    w_warns = list((wstate or {}).get("last_warnings") or [])
    return {
        "id": proj["id"],
        "name": proj["name"],
        "code": proj.get("code", "PRJ"),
        "contract_no": proj.get("contract_no", ""),
        "owner": proj.get("owner", ""),
        "consultant": proj.get("consultant", ""),
        "contractor": proj.get("contractor", ""),
        "location": proj.get("location", ""),
        "engineer": proj.get("engineer", ""),
        "root": str(proj_root),
        "inbox_dir": proj.get("inbox_dir", ""),
        "archive_dir": str(archive_dir_for_project(proj)),
        "enabled": proj.get("enabled", True),
        "is_active": gcfg.get("active_project_id") == proj["id"],
        "unrelated_policy": proj.get("unrelated_policy") or gcfg.get("default_unrelated_policy", "rename_in_place"),
        "unrelated_policy_override": proj.get("unrelated_policy", ""),
        "relevance_threshold": proj.get("relevance_threshold") if proj.get("relevance_threshold") is not None else gcfg.get("relevance_threshold", DEFAULT_RELEVANCE_THRESHOLD),
        "dry_run": bool(proj.get("dry_run") or gcfg.get("dry_run")),
        "log_unrelated_sheet": bool(proj.get("log_unrelated_sheet", False)),
        "global_submittal_numbering": proj.get("global_submittal_numbering"),
        "overdue_days": proj.get("overdue_days"),
        "stats": log_sum["stats"],
        "logs": [
            {"name": lg["name"], "ext": lg["ext"], "has_vba": lg.get("has_vba", False), "sheets_enabled": lg.get("sheets_enabled", 0)}
            for lg in log_sum["logs"]
        ],
        "templates_count": len(tpls),
        "templates_ready": sum(1 for t in tpls if t.get("ready_to_issue")),
        "orphan_templates_count": sum(1 for t in tpls if t.get("orphan_fields")),
        "pending_decisions": hub.decisions.pending_count(project_id=proj["id"]),
        "scanner_triage_pending": triage_pending,
        "wa_plan_count": wa_plan_count,
        "dup_groups_count": dup_groups_count,
        "backup": bkp,
        "job_running": bool(job.get("running")),
        "last_scan_errors": w_errs[-5:],
        "last_scan_warnings": w_warns[-5:],
        "last_started_at": proj.get("last_started_at"),
        "last_backup_at": proj.get("last_backup_at"),
    }


def build_project_detail(hub: HubState, proj: dict[str, Any], *, query: str = "") -> dict[str, Any]:
    gcfg = hub.get_settings()
    proj_root = Path(str(proj["root"])).resolve()
    style = gcfg.get("ui_style")
    proj_sections_all = tprof.get_project_sections(proj, style=style, include_hidden=True)
    proj_sections_vis = [s for s in proj_sections_all if not s.get("hidden")]
    log_sum = plog.read_project_log_summary(proj, query=query, limit=300, recent_first=True, device_id=hub.device_id)
    counters = plog.recalculate_counters_from_logs(proj, device_id=hub.device_id)
    tpls = tprof.list_project_templates(proj_root, device_id=hub.device_id, project_sections=proj_sections_all)
    bkp = check_backup_status(proj, gcfg, warn_days=int(gcfg.get("backup_warn_days") or DEFAULT_BACKUP_WARN_DAYS))
    ibx = Path(str(proj.get("inbox_dir") or (proj_root / "1_Inbox")))
    rjournal = load_rename_journal(rename_journal_path_for(ibx)) if ibx.exists() else []
    wstate = read_json_safe(proj_root / "_System_Agent" / "watcher_state.json", {}) if proj_root.exists() else {}
    triage = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)

    safe_proj = scrub_secrets(dict(proj))
    safe_proj["archive_dir"] = str(archive_dir_for_project(proj))
    return {
        "ok": True,
        "project": safe_proj,
        "is_active": gcfg.get("active_project_id") == proj["id"],
        "device_id": hub.device_id,
        "stats": log_sum["stats"],
        "logs": log_sum["logs"],
        "recent_rows": log_sum["rows"],
        "counters": counters,
        "templates": tpls,  # STRICTLY user-uploaded templates only!
        "sections": proj_sections_vis,
        "all_sections": proj_sections_all,
        "scanner_triage": triage,
        "backup": bkp,
        "pending_decisions": hub.decisions.list_decisions(status="pending", project_id=proj["id"], limit=100),
        "rename_journal": list(reversed(rjournal[-80:])),
        "job": hub.jobs.get(proj["id"]) or {"running": False},
        "legacy_job": hub.legacy_jobs.get(proj["id"]) or {"running": False},
        "last_scan_at": (wstate or {}).get("last_scan_at"),
        "last_scan_errors": list((wstate or {}).get("last_errors") or []),
        "last_scan_warnings": list((wstate or {}).get("last_warnings") or []),
    }


def global_search_hub(
    hub: HubState,
    query: str,
    *,
    project_id: str = "",
    kind: str = "all",
    limit: int = 90,
) -> dict[str, Any]:
    """
    Unified search across all projects (or one selected project):
      1) Excel Log rows (incoming/outgoing letters, submittals, invoices, ...)
      2) Archived files in `archive_dir`
      3) WhatsApp Archive items (`whatsapp_archive_plan.json`)
    """
    q = str(query or "").strip()
    if not q:
        return {
            "ok": True,
            "query": "",
            "total": 0,
            "counts": {"all": 0, "log_row": 0, "archive_file": 0, "whatsapp": 0},
            "results": [],
        }

    q_norm = plog.normalize_ar(q)
    lim = max(1, min(int(limit or 90), 250))
    kind_f = str(kind or "all").strip().lower()
    if kind_f not in ("all", "log_row", "archive_file", "whatsapp"):
        kind_f = "all"

    projects = [
        p for p in hub.list_projects()
        if (project_id and p["id"] == project_id) or (not project_id and p.get("enabled", True))
    ]
    previewable_exts = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tif", ".tiff", ".docx", ".xlsx", ".xlsm"}

    log_hits: list[dict[str, Any]] = []
    arc_hits: list[dict[str, Any]] = []
    wa_hits: list[dict[str, Any]] = []

    for proj in projects:
        pid = proj["id"]
        pcode = proj.get("code", "")
        pname = proj.get("name", "")
        proj_root = Path(str(proj["root"])).resolve()
        arc_root = archive_dir_for_project(proj)
        try:
            arc_root.relative_to(proj_root)
            arc_area = "root"
            arc_rel_base = proj_root
        except ValueError:
            arc_area = "archive"
            arc_rel_base = arc_root

        # 1) Log rows (includes incoming/outgoing letters, submittals, IPC, petty cash, etc.)
        try:
            lsum = plog.read_project_log_summary(
                proj, query=q, limit=40, recent_first=True, device_id=hub.device_id
            )
            for r in lsum.get("rows") or []:
                f_rel = str(r.get("file_rel") or "")
                f_area = str(r.get("file_area") or "root")
                f_ext = str(r.get("file_ext") or "")
                dl_url = ""
                if f_rel:
                    dl_url = f"/api/projects/{pid}/download?rel={quote(f_rel, safe='/')}" + ("&area=archive" if f_area == "archive" else "")
                log_hits.append({
                    "kind": "log_row",
                    "project_id": pid,
                    "project_code": pcode,
                    "project_name": pname,
                    "ref_no": r.get("ref_no") or "",
                    "title": r.get("subject") or r.get("ref_no") or r.get("sheet") or "",
                    "subtitle": " · ".join([x for x in (r.get("sheet"), r.get("party_label") or r.get("party")) if x]),
                    "section": r.get("section") or "",
                    "date": r.get("date") or "",
                    "status": r.get("status") or "",
                    "amount_kwd": r.get("amount_kwd"),
                    "file_rel": f_rel,
                    "file_area": f_area,
                    "file_ext": f_ext,
                    "is_previewable": bool(r.get("is_previewable")),
                    "download_url": dl_url,
                })
        except Exception:
            pass

        # 2) Archive files
        if arc_root.exists() and arc_root.is_dir():
            walked = 0
            for fp in arc_root.rglob("*"):
                if walked >= 2500 or len(arc_hits) >= 120:
                    break
                if not fp.is_file() or fp.name.startswith(("~$", ".")):
                    continue
                if any(part.startswith("_System_Agent") or part == "_Backups" for part in fp.parts):
                    continue
                walked += 1
                try:
                    rel_arc = fp.relative_to(arc_root).as_posix()
                    rel_dl = fp.relative_to(arc_rel_base).as_posix()
                except ValueError:
                    continue
                if q_norm not in plog.normalize_ar(f"{fp.name} {rel_arc}"):
                    continue
                ext = fp.suffix.lower()
                if ext.endswith(".json"):
                    continue
                try:
                    st = fp.stat()
                    m_date = dt.date.fromtimestamp(st.st_mtime).isoformat()
                except OSError:
                    m_date = ""
                head_ref = fp.stem.split("_")[0].strip()
                ref_guess = head_ref if re.search(r"[A-Za-z\u0600-\u06FF]+\s*[-/]\s*\d+", head_ref) else ""
                dl_url = f"/api/projects/{pid}/download?rel={quote(rel_dl, safe='/')}" + ("&area=archive" if arc_area == "archive" else "")
                arc_hits.append({
                    "kind": "archive_file",
                    "project_id": pid,
                    "project_code": pcode,
                    "project_name": pname,
                    "ref_no": ref_guess,
                    "title": fp.name,
                    "subtitle": rel_arc,
                    "section": "",
                    "date": m_date,
                    "status": ext.lstrip(".").upper(),
                    "amount_kwd": None,
                    "file_rel": rel_dl,
                    "file_area": arc_area,
                    "file_ext": ext,
                    "is_previewable": ext in previewable_exts,
                    "download_url": dl_url,
                })

        # 3) WhatsApp Archive items
        try:
            wa_p = waa.whatsapp_plan_path(proj_root)
            if wa_p.exists():
                wa_plan = read_json_safe(wa_p, {}) or {}
                for it in wa_plan.get("items") or []:
                    hay_wa = plog.normalize_ar(
                        f"{it.get('final_name') or ''} {it.get('orig_name') or ''} {it.get('context_text') or ''} {it.get('sender') or ''} {it.get('ref_no') or ''} {it.get('date_iso') or ''}"
                    )
                    if q_norm not in hay_wa:
                        continue
                    fp_raw = str(it.get("final_path") or it.get("src_path") or "")
                    ext = Path(fp_raw).suffix.lower() if fp_raw else ""
                    wa_id = str(it.get("id") or "")
                    dl_url = f"/api/projects/{pid}/whatsapp-archive/file?id={quote(wa_id)}" if wa_id else ""
                    wa_rel = ""
                    wa_area = "root"
                    if fp_raw:
                        fp_p = resolve_local_or_unc(fp_raw)
                        try:
                            wa_rel = fp_p.relative_to(proj_root).as_posix()
                            wa_area = "root"
                        except ValueError:
                            try:
                                wa_rel = fp_p.relative_to(arc_root).as_posix()
                                wa_area = "archive"
                            except ValueError:
                                wa_rel = ""
                    wa_hits.append({
                        "kind": "whatsapp",
                        "project_id": pid,
                        "project_code": pcode,
                        "project_name": pname,
                        "ref_no": it.get("ref_no") or "",
                        "title": it.get("final_name") or it.get("orig_name") or "",
                        "subtitle": (it.get("context_text") or it.get("sender") or "")[:140],
                        "section": "whatsapp",
                        "date": it.get("date_iso") or "",
                        "status": it.get("sender") or it.get("media_kind") or "",
                        "amount_kwd": None,
                        "file_rel": wa_rel,
                        "file_area": wa_area,
                        "file_ext": ext,
                        "wa_item_id": wa_id,
                        "is_previewable": bool(wa_rel and ext in previewable_exts),
                        "download_url": dl_url,
                    })
                    if len(wa_hits) >= 80:
                        break
        except Exception:
            pass

    counts = {
        "all": len(log_hits) + len(arc_hits) + len(wa_hits),
        "log_row": len(log_hits),
        "archive_file": len(arc_hits),
        "whatsapp": len(wa_hits),
    }
    if kind_f == "log_row":
        combined = log_hits
    elif kind_f == "archive_file":
        combined = arc_hits
    elif kind_f == "whatsapp":
        combined = wa_hits
    else:
        combined = log_hits + arc_hits + wa_hits

    combined.sort(key=lambda x: str(x.get("date") or ""), reverse=True)
    return {
        "ok": True,
        "query": q,
        "total": counts[kind_f],
        "counts": counts,
        "results": combined[:lim],
    }


def build_hub_state(hub: HubState) -> dict[str, Any]:
    gcfg = hub.get_settings()
    all_p = hub.list_projects()
    cards = [build_project_card(hub, p, gcfg) for p in all_p]
    shared_warns = detect_shared_inboxes(all_p)
    stale_warns = [
        w
        for c in cards
        for w in ((c.get("backup") or {}).get("warnings") or [])
    ]
    ai_st = sw.get_gemini_keys_status()

    style = gcfg.get("ui_style") or DEFAULT_STYLE
    active_jobs: list[dict[str, Any]] = []
    for p_item in all_p:
        p_id = p_item["id"]
        p_code = p_item.get("code", "")
        for sc_name, st_job in (
            ("organizer", hub.get_organizer_job_status(p_id)),
            ("duplicates", hub.get_scoped_job_status("dup", p_id)),
            ("whatsapp", hub.get_scoped_job_status("wa", p_id)),
            ("scanner", hub.get_scanner_job_status(p_id)),
            ("legacy", dict(hub.legacy_jobs.get(p_id) or {"running": False})),
        ):
            if st_job.get("running"):
                active_jobs.append({
                    **st_job,
                    "scope": sc_name,
                    "project_id": p_id,
                    "project_code": p_code,
                })

    t_store = tbar._load_tasks_store(hub.state_dir)
    manual_open_list = [t for t in (t_store.get("tasks") or []) if not t.get("done")]
    manual_urgent = sum(1 for t in manual_open_list if t.get("priority") == "urgent")
    snoozed_cnt = len(t_store.get("dismissed_auto") or {})
    auto_est = sum(
        int((c.get("stats") or {}).get("overdue") or 0)
        + int((c.get("stats") or {}).get("unreplied_incoming") or 0)
        + int((c.get("stats") or {}).get("rejected") or 0)
        + (1 if int(c.get("scanner_triage_pending") or 0) > 0 else 0)
        + (1 if (c.get("backup") or {}).get("stale") else 0)
        for c in cards
        if c.get("enabled", True)
    )
    auto_est_eff = max(0, auto_est - snoozed_cnt)
    tasks_summary = {
        "manual_open": len(manual_open_list),
        "manual_urgent": manual_urgent,
        "auto_open": auto_est_eff,
        "total_open": len(manual_open_list) + auto_est_eff,
    }

    return {
        "ok": True,
        "version": APP_VERSION,
        "build_tag": "2026-11-18-v10.13.2",
        "app_dir": str(SCRIPT_DIR),
        "device_id": hub.device_id,
        "projects_root": str(hub.projects_root),
        "ui_style": style,
        "active_project_id": gcfg.get("active_project_id"),
        "session_active": bool(gcfg.get("session_active")),
        "session_started_at": hub.data.get("session_started_at"),
        "watcher_running": hub.watcher_running,
        "watcher_last_error": hub.last_watcher_error,
        "active_jobs": active_jobs,
        "tasks_summary": tasks_summary,
        "pymupdf_available": getattr(sw, "fitz", None) is not None,
        "cloud_backup_folders": detect_cloud_backup_folders(),
        "settings": gcfg,
        "projects": cards,
        "pending_decisions_count": hub.decisions.pending_count(),
        "learned_rules_count": len(hub.decisions.list_rules(include_disabled=False)),
        "shared_inbox_warnings": shared_warns,
        "stale_backup_warnings": stale_warns,
        "sections": tprof.get_localized_sections(style),
        "auto_date_sources": tprof.get_localized_auto_date_sources(style),
        "auto_project_sources": tprof.get_localized_auto_project_sources(style),
        "auto_system_sources": tprof.get_localized_auto_system_sources(style),
        "ai_status": {
            **ai_st,
            "engine": gcfg.get("engine", "gemini"),
        },
    }


# ══════════════════════════════════════════════════════════════════════════════
# HTTP Server & Security-Hardened REST Router
# ══════════════════════════════════════════════════════════════════════════════
def make_handler(hub: HubState, *, allow_remote_preview: bool = False):
    ui_path = SCRIPT_DIR / "dashboard_ui.html"

    ACTIVE_WEB_EXTS = {
        ".html", ".htm", ".xhtml", ".svg", ".xml", ".js", ".mjs",
        ".css", ".swf", ".mht", ".mhtml", ".hta",
    }
    RESERVED_WIN_NAMES = {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    }

    def _resolve_decision_file_safely(dec: dict[str, Any]) -> Path:
        raw_fp = str(dec.get("file_path") or "").strip()
        if not raw_fp or "\x00" in raw_fp:
            raise FileNotFoundError(_tr("scanner.file_missing_warn"))
        if os.path.islink(raw_fp):
            raise PathTraversalError(_tr("server.path_blocked"))
        fp = resolve_local_or_unc(raw_fp)
        if fp.is_symlink():
            raise PathTraversalError(_tr("server.path_blocked"))
        allowed_bases: list[Path] = [resolve_local_or_unc(hub.state_dir), resolve_local_or_unc(hub.projects_root)]
        trusted_dir = str(dec.get("trusted_source_dir") or "").strip()
        if trusted_dir:
            allowed_bases.append(resolve_local_or_unc(trusted_dir))
        try:
            cfg = load_watcher_config()
            for k in ("inbox", "drafts", "archive"):
                v = str(cfg.get(k) or "").strip()
                if v:
                    allowed_bases.append(resolve_local_or_unc(v))
        except Exception:
            pass
        for p in hub.list_projects():
            for k in ("inbox_dir", "archive_dir", "root"):
                val = str(p.get(k) or "").strip()
                if val:
                    allowed_bases.append(resolve_local_or_unc(val))
        for base in allowed_bases:
            try:
                rel_p = fp.relative_to(base)
                return safe_resolve_within(base, rel_p.as_posix(), allow_base=False, reject_symlinks=True)
            except PathTraversalError:
                raise
            except Exception:
                continue
        raise PathTraversalError(_tr("server.path_blocked"))

    def _sanitize_upload_rel_segments(raw_rel: str) -> str:
        norm_rel = str(raw_rel or "").replace("\\", "/")
        if "\x00" in norm_rel:
            raise PathTraversalError(_tr("server.path_blocked"))
        rel_segs = [s for s in norm_rel.split("/") if s not in ("", ".")]
        if not rel_segs or any(s == ".." for s in rel_segs):
            raise PathTraversalError(_tr("server.path_blocked"))
        dir_segs: list[str] = []
        for s in rel_segs[:-1]:
            cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", s).strip(" ._")
            if not cleaned or cleaned.upper() in RESERVED_WIN_NAMES:
                cleaned = f"dir_{cleaned}" if cleaned else "dir"
            dir_segs.append(cleaned[:70])
        leaf_safe = sanitize_upload_filename(rel_segs[-1], allowed_exts=None)
        return "/".join(dir_segs + [leaf_safe])

    class DashboardHandler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def log_message(self, fmt: str, *args: Any) -> None:
            return

        def _add_security_headers(
            self,
            is_api: bool = True,
            *,
            is_html: bool = False,
            sandbox_csp: str = "",
        ) -> None:
            self.send_header("X-Content-Type-Options", "nosniff")
            if not allow_remote_preview:
                self.send_header("X-Frame-Options", "SAMEORIGIN")
            if sandbox_csp:
                self.send_header("Content-Security-Policy", sandbox_csp)
            elif is_html:
                fa = "" if allow_remote_preview else "; frame-ancestors 'self'"
                self.send_header(
                    "Content-Security-Policy",
                    "default-src 'self' data: blob:; "
                    "script-src 'self' 'unsafe-inline'; "
                    "style-src 'self' 'unsafe-inline'; "
                    "img-src 'self' data: blob:; "
                    "media-src 'self' blob:; "
                    "connect-src 'self'; "
                    "object-src 'none'; "
                    f"base-uri 'self'{fa}",
                )
            if is_api:
                self.send_header("Cache-Control", "no-store")

        def _send_json(self, obj: Any, status: int = 200) -> None:
            raw = json.dumps(obj, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self._add_security_headers(is_api=True)
            self.end_headers()
            self.wfile.write(raw)

        def _send_bytes(
            self,
            raw: bytes,
            content_type: str,
            *,
            filename: str = "",
            status: int = 200,
            is_html: bool = False,
            sandbox_csp: str = "",
        ) -> None:
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(raw)))
            if filename:
                ascii_fb = re.sub(r"[^\x20-\x7E]", "_", filename).replace('"', "")
                self.send_header(
                    "Content-Disposition",
                    f'attachment; filename="{ascii_fb}"; filename*=UTF-8\'\'{quote(filename)}',
                )
            self._add_security_headers(is_api=bool(filename), is_html=is_html, sandbox_csp=sandbox_csp)
            self.end_headers()
            self.wfile.write(raw)

        def _send_user_file(
            self,
            safe_target: Path,
            *,
            force_inline: bool = True,
            allow_range: bool = False,
        ) -> None:
            ext = safe_target.suffix.lower()
            is_active = ext in ACTIVE_WEB_EXTS
            ctype = "application/octet-stream" if is_active else (mimetypes.guess_type(safe_target.name)[0] or "application/octet-stream")
            sandbox_csp = "sandbox; default-src 'none'" if is_active else "default-src 'none'; img-src 'self'; media-src 'self'; object-src 'self'; script-src 'none'"
            fsize = safe_target.stat().st_size
            ascii_fb = re.sub(r"[^\x20-\x7E]", "_", safe_target.name).replace('"', "")
            disp_kind = "attachment" if (is_active or not force_inline) else "inline"
            disp_hdr = f'{disp_kind}; filename="{ascii_fb}"; filename*=UTF-8\'\'{quote(safe_target.name)}'

            rng_hdr = self.headers.get("Range") or ""
            if allow_range and not is_active and rng_hdr.startswith("bytes=") and fsize > 0:
                spec = rng_hdr[6:].split(",")[0].strip()
                start_s, _, end_s = spec.partition("-")
                try:
                    start_b = int(start_s) if start_s else 0
                    end_b = int(end_s) if end_s else (fsize - 1)
                    start_b = max(0, min(start_b, fsize - 1))
                    end_b = max(start_b, min(end_b, fsize - 1))
                    chunk_len = end_b - start_b + 1
                    self.send_response(206)
                    self.send_header("Content-Type", ctype)
                    self.send_header("Accept-Ranges", "bytes")
                    self.send_header("Content-Range", f"bytes {start_b}-{end_b}/{fsize}")
                    self.send_header("Content-Length", str(chunk_len))
                    self.send_header("Content-Disposition", disp_hdr)
                    self._add_security_headers(is_api=False, sandbox_csp=sandbox_csp)
                    self.end_headers()
                    with open(safe_target, "rb") as rf:
                        rf.seek(start_b)
                        rem = chunk_len
                        while rem > 0:
                            buf = rf.read(min(64 * 1024, rem))
                            if not buf:
                                break
                            self.wfile.write(buf)
                            rem -= len(buf)
                    return
                except ValueError:
                    pass

            self.send_response(200)
            self.send_header("Content-Type", ctype)
            if allow_range and not is_active:
                self.send_header("Accept-Ranges", "bytes")
            self.send_header("Content-Length", str(fsize))
            self.send_header("Content-Disposition", disp_hdr)
            self._add_security_headers(is_api=is_active or not force_inline, sandbox_csp=sandbox_csp)
            self.end_headers()
            with open(safe_target, "rb") as rf:
                while True:
                    buf = rf.read(64 * 1024)
                    if not buf:
                        break
                    self.wfile.write(buf)

        def _read_json(self) -> dict[str, Any]:
            try:
                length = int(self.headers.get("Content-Length") or 0)
            except ValueError as exc:
                raise ValueError(_tr("server.invalid_setting")) from exc
            if length < 0:
                raise ValueError(_tr("server.invalid_setting"))
            if length == 0:
                return {}
            if length > MAX_BODY_BYTES:
                raise OverflowError(_tr("server.file_too_large"))
            raw = self.rfile.read(length)
            try:
                data = json.loads(raw.decode("utf-8"))
            except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                raise ValueError(_tr("server.invalid_setting")) from exc
            return data if isinstance(data, dict) else {}

        def _check_origin(self) -> bool:
            if not validate_request_origin(self.headers, allow_remote_preview=allow_remote_preview):
                self._send_json({"ok": False, "error": _tr("server.request_blocked")}, 403)
                return False
            return True

        def _split_api_parts(self, raw_path: str) -> list[str]:
            # Check for encoded traversal BEFORE/AFTER unquoting
            decoded = unquote(raw_path)
            if "\x00" in decoded:
                raise PathTraversalError(_tr("server.path_blocked"))
            raw_segs = [unquote(seg) for seg in raw_path.strip("/").split("/") if seg != ""]
            if any(seg in (".", "..") for seg in raw_segs):
                raise PathTraversalError(_tr("server.path_blocked"))
            return raw_segs

        # ── GET ──────────────────────────────────────────────────────────────
        def do_GET(self) -> None:
            if not self._check_origin():
                return
            parsed = urlparse(self.path)
            qs = parse_qs(parsed.query)
            try:
                cur_style = hub.get_settings().get("ui_style") or DEFAULT_STYLE
                if parsed.path in ("/", "/index.html"):
                    if not ui_path.exists():
                        return self._send_bytes(b"<h1>UI missing</h1>", "text/html; charset=utf-8")
                    html_txt = ui_path.read_text(encoding="utf-8")
                    bundle_json = json.dumps(
                        get_client_bundle(cur_style), ensure_ascii=False, separators=(",", ":")
                    ).replace("</", "<\\/")
                    html_txt = html_txt.replace("/*__I18N_BUNDLE__*/null", bundle_json, 1)
                    return self._send_bytes(html_txt.encode("utf-8"), "text/html; charset=utf-8", is_html=True)
                if parsed.path == "/favicon.ico":
                    ico_p = SCRIPT_DIR / "app_icon.ico"
                    if ico_p.exists():
                        return self._send_bytes(ico_p.read_bytes(), "image/x-icon", status=200)
                    return self._send_bytes(b"", "image/x-icon", status=204)
                if parsed.path == "/app-icon-192.png":
                    p192 = SCRIPT_DIR / "app_icon_192.png"
                    if p192.exists():
                        return self._send_bytes(p192.read_bytes(), "image/png", status=200)
                    return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                if parsed.path == "/app-icon-512.png":
                    p512 = SCRIPT_DIR / "app_icon_512.png"
                    if p512.exists():
                        return self._send_bytes(p512.read_bytes(), "image/png", status=200)
                    return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                if parsed.path == "/manifest.webmanifest":
                    manifest = {
                        "name": "Smart Projects Hub",
                        "short_name": "Projects Hub",
                        "description": "Smart Local Document & Correspondence Control Hub",
                        "start_url": "/",
                        "scope": "/",
                        "display": "standalone",
                        "background_color": "#f6f8fb",
                        "theme_color": "#059BE5",
                        "icons": [
                            {"src": "/app-icon-192.png", "sizes": "192x192", "type": "image/png"},
                            {"src": "/app-icon-512.png", "sizes": "512x512", "type": "image/png"},
                        ],
                    }
                    return self._send_bytes(
                        json.dumps(manifest, ensure_ascii=False).encode("utf-8"),
                        "application/manifest+json; charset=utf-8",
                        status=200,
                    )

                parts = self._split_api_parts(parsed.path)
                if not parts or parts[0] != "api":
                    return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)

                # GET /api/i18n
                if parts == ["api", "i18n"]:
                    req_style = (qs.get("style") or [cur_style])[0]
                    return self._send_json({"ok": True, **get_client_bundle(req_style)})

                # GET /api/state
                if parts == ["api", "state"]:
                    return self._send_json(build_hub_state(hub))

                # GET /api/settings
                if parts == ["api", "settings"]:
                    return self._send_json({
                        "ok": True,
                        "settings": hub.get_settings(),
                        "cloud_backup_folders": detect_cloud_backup_folders(),
                    })

                # GET /api/search?q=...&project_id=...&kind=...
                if parts == ["api", "search"]:
                    q_val = (qs.get("q") or [""])[0]
                    pid_val = (qs.get("project_id") or [""])[0]
                    kind_val = (qs.get("kind") or ["all"])[0]
                    lim_val = int((qs.get("limit") or ["90"])[0] or 90)
                    return self._send_json(
                        global_search_hub(hub, q_val, project_id=pid_val, kind=kind_val, limit=lim_val)
                    )

                # GET /api/tasks?project_id=...&include_done=1&limit=80
                if parts == ["api", "tasks"]:
                    pid_val = (qs.get("project_id") or [""])[0]
                    inc_done = str((qs.get("include_done") or ["1"])[0]).strip().lower() not in ("0", "false", "no")
                    lim_val = int((qs.get("limit") or ["80"])[0] or 80)
                    return self._send_json(
                        tbar.build_tasks_payload(hub, project_id=pid_val, include_done=inc_done, limit=lim_val)
                    )

                # GET /api/decisions
                if parts == ["api", "decisions"]:
                    status_f = (qs.get("status") or [None])[0]
                    pid_f = (qs.get("project_id") or [None])[0]
                    items = hub.decisions.list_decisions(status=status_f, project_id=pid_f)
                    return self._send_json({
                        "ok": True,
                        "device_id": hub.device_id,
                        "pending_count": hub.decisions.pending_count(project_id=pid_f),
                        "decisions": items,
                    })

                # GET /api/decisions/<id>/preview?page=1&dpi=120
                if len(parts) == 4 and parts[:2] == ["api", "decisions"] and parts[3] == "preview":
                    dec = hub.decisions.get_decision(parts[2])
                    if not dec:
                        return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                    src = _resolve_decision_file_safely(dec)
                    page_no = int((qs.get("page") or ["1"])[0] or 1)
                    dpi_val = int((qs.get("dpi") or ["120"])[0] or 120)
                    png_bytes, total_pages = render_document_preview_page(
                        src,
                        page_no,
                        dpi=dpi_val,
                        cache_dir=hub.state_dir / "_preview_cache",
                        cache_key=str(dec["id"]),
                    )
                    self.send_response(200)
                    self.send_header("Content-Type", "image/png")
                    self.send_header("Content-Length", str(len(png_bytes)))
                    self.send_header("X-Total-Pages", str(total_pages))
                    self.send_header("X-Page-Number", str(page_no))
                    self.send_header("Cache-Control", "private, max-age=120")
                    self._add_security_headers(is_api=False)
                    self.end_headers()
                    self.wfile.write(png_bytes)
                    return

                # GET /api/decisions/<id>/file
                if len(parts) == 4 and parts[:2] == ["api", "decisions"] and parts[3] == "file":
                    dec = hub.decisions.get_decision(parts[2])
                    if not dec:
                        return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                    src = _resolve_decision_file_safely(dec)
                    if not src.exists() or not src.is_file():
                        return self._send_json({"ok": False, "error": _tr("scanner.file_missing_warn")}, 404)
                    return self._send_user_file(src, force_inline=True, allow_range=False)

                # GET /api/rename-journal
                if parts == ["api", "rename-journal"]:
                    pid_f = (qs.get("project_id") or [""])[0]
                    all_entries: list[dict[str, Any]] = []
                    seen_j: set[str] = set()
                    for p in hub.list_projects():
                        if pid_f and p["id"] != pid_f:
                            continue
                        jp = rename_journal_path_for(Path(str(p.get("inbox_dir") or (Path(p["root"]) / "1_Inbox"))))
                        jkey = str(jp)
                        if jkey in seen_j or not jp.exists():
                            continue
                        seen_j.add(jkey)
                        for e in load_rename_journal(jp):
                            all_entries.append({**e, "journal_project_id": p["id"]})
                    all_entries.sort(key=lambda x: str(x.get("timestamp") or ""), reverse=True)
                    return self._send_json({"ok": True, "entries": all_entries[:200]})

                # GET /api/learned-rules
                if parts == ["api", "learned-rules"]:
                    return self._send_json({"ok": True, "rules": hub.decisions.list_rules(include_disabled=True)})

                # GET /api/whatsapp/config & webhook
                if parts == ["api", "whatsapp", "config"] and wab:
                    raw_wa = wab.load_wa_config(hub.state_dir / "whatsapp_config.json")
                    return self._send_json({"ok": True, "config": wab.public_wa_config(raw_wa)})

                # Project-scoped GET routes: /api/projects/<id>/...
                if len(parts) >= 3 and parts[1] == "projects":
                    pid = parts[2]
                    proj = hub.get_project(pid)
                    if not proj:
                        return self._send_json({"ok": False, "error": _tr("server.project_not_found")}, 404)
                    proj_root = Path(str(proj["root"])).resolve()

                    # GET /api/projects/<id>
                    if len(parts) == 3:
                        q = (qs.get("q") or [""])[0]
                        return self._send_json(build_project_detail(hub, proj, query=q))

                    # GET /api/projects/<id>/download?rel=...&area=...
                    if len(parts) == 4 and parts[3] == "download":
                        rel = (qs.get("rel") or [""])[0]
                        area = (qs.get("area") or ["root"])[0]
                        base_dir = archive_dir_for_project(proj) if area == "archive" else proj_root
                        target = safe_resolve_within(base_dir, rel, allow_base=False, reject_symlinks=True)
                        if (not target.exists() or not target.is_file()) and area != "archive":
                            alt = safe_resolve_within(archive_dir_for_project(proj), rel, allow_base=False, reject_symlinks=True)
                            if alt.exists() and alt.is_file():
                                target = alt
                        if not target.exists() or not target.is_file():
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        ctype = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
                        return self._send_bytes(target.read_bytes(), ctype, filename=target.name)

                    # GET /api/projects/<id>/preview-info?rel=...&area=...
                    if len(parts) == 4 and parts[3] == "preview-info":
                        rel = (qs.get("rel") or [""])[0]
                        area = (qs.get("area") or ["root"])[0]
                        base_dir = archive_dir_for_project(proj) if area == "archive" else proj_root
                        target = safe_resolve_within(base_dir, rel, allow_base=False, reject_symlinks=True)
                        if (not target.exists() or not target.is_file()) and area != "archive":
                            alt = safe_resolve_within(archive_dir_for_project(proj), rel, allow_base=False, reject_symlinks=True)
                            if alt.exists() and alt.is_file():
                                target = alt
                                area = "archive"
                        if not target.exists() or not target.is_file():
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        ext = target.suffix.lower()
                        dl_url = f"/api/projects/{pid}/download?rel={quote(rel, safe='/')}" + ("&area=archive" if area == "archive" else "")
                        if ext in (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tif", ".tiff"):
                            return self._send_json({
                                "ok": True, "kind": "image", "total_pages": 1,
                                "name": target.name, "ext": ext, "download_url": dl_url,
                            })
                        if ext == ".pdf":
                            tot_p = trender.count_pdf_pages(target)
                            has_fitz = getattr(sw, "fitz", None) is not None
                            return self._send_json({
                                "ok": True,
                                "kind": "pdf" if has_fitz else "pdf_no_fitz",
                                "total_pages": tot_p,
                                "name": target.name,
                                "ext": ext,
                                "download_url": dl_url,
                            })
                        if ext in (".docx", ".xlsx", ".xlsm"):
                            vis = (
                                trender.render_xlsx_file(target)
                                if ext in (".xlsx", ".xlsm")
                                else trender.render_docx_file(target, try_libreoffice=True)
                            )
                            return self._send_json({
                                "ok": True, "kind": "office", "total_pages": 1,
                                "name": target.name, "ext": ext, "download_url": dl_url, "render": vis,
                            })
                        return self._send_json({
                            "ok": True, "kind": "unsupported", "total_pages": 0,
                            "name": target.name, "ext": ext, "download_url": dl_url,
                        })

                    # GET /api/projects/<id>/preview-page?rel=...&area=...&page=1&dpi=135
                    if len(parts) == 4 and parts[3] == "preview-page":
                        rel = (qs.get("rel") or [""])[0]
                        area = (qs.get("area") or ["root"])[0]
                        page_no = int((qs.get("page") or ["1"])[0] or 1)
                        dpi_val = int((qs.get("dpi") or ["135"])[0] or 135)
                        base_dir = archive_dir_for_project(proj) if area == "archive" else proj_root
                        target = safe_resolve_within(base_dir, rel, allow_base=False, reject_symlinks=True)
                        if (not target.exists() or not target.is_file()) and area != "archive":
                            alt = safe_resolve_within(archive_dir_for_project(proj), rel, allow_base=False, reject_symlinks=True)
                            if alt.exists() and alt.is_file():
                                target = alt
                        if not target.exists() or not target.is_file():
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        cache_dir = proj_root / "_System_Agent" / "_preview_cache"
                        png_bytes, total_pages = render_document_preview_page(
                            target,
                            page_no,
                            dpi=dpi_val,
                            cache_dir=cache_dir,
                            cache_key=sha256_of_file(target)[:24],
                        )
                        self.send_response(200)
                        self.send_header("Content-Type", "image/png")
                        self.send_header("Content-Length", str(len(png_bytes)))
                        self.send_header("X-Total-Pages", str(total_pages))
                        self.send_header("X-Page-Number", str(page_no))
                        self.send_header("Cache-Control", "private, max-age=120")
                        self._add_security_headers(is_api=False)
                        self.end_headers()
                        self.wfile.write(png_bytes)
                        return

                    # GET /api/projects/<id>/scanner
                    if len(parts) == 4 and parts[3] == "scanner":
                        st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                        return self._send_json({**st, "job": hub.get_scanner_job_status(pid)})

                    # GET /api/projects/<id>/scanner/page-thumb?id=...&page=1&dpi=110
                    if len(parts) == 5 and parts[3] == "scanner" and parts[4] == "page-thumb":
                        item_id = (qs.get("id") or [""])[0]
                        page_no = int((qs.get("page") or ["1"])[0] or 1)
                        dpi_val = int((qs.get("dpi") or ["110"])[0] or 110)
                        png_bytes = render_triage_pdf_page(proj, item_id, page_no, dpi=dpi_val)
                        self.send_response(200)
                        self.send_header("Content-Type", "image/png")
                        self.send_header("Content-Length", str(len(png_bytes)))
                        self.send_header("Cache-Control", "private, max-age=60")
                        self._add_security_headers(is_api=False)
                        self.end_headers()
                        self.wfile.write(png_bytes)
                        return

                    # GET /api/projects/<id>/scanner/file?id=...
                    if len(parts) == 5 and parts[3] == "scanner" and parts[4] == "file":
                        item_id = (qs.get("id") or [""])[0]
                        st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                        found = next(
                            (it for bucket in ("pending", "approved", "rejected") for it in st[bucket] if str(it.get("id")) == str(item_id)),
                            None,
                        )
                        if not found:
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        ibx_root = resolve_local_or_unc(str(proj.get("inbox_dir") or (proj_root / "1_Inbox")))
                        arc_root = archive_dir_for_project(proj)
                        proj_root_norm = resolve_local_or_unc(proj_root)
                        fp = resolve_local_or_unc(str(found.get("archive_path") or found.get("file_path") or ""))
                        # Strictly verify file lies within inbox or archive or proj_root (UNC-safe)!
                        allowed_bases = [ibx_root, arc_root, proj_root_norm]
                        safe_target: Path | None = None
                        for b in allowed_bases:
                            try:
                                rel_p = fp.relative_to(b)
                                safe_target = safe_resolve_within(b, rel_p.as_posix(), allow_base=False, reject_symlinks=True)
                                break
                            except Exception:
                                continue
                        if not safe_target or not safe_target.exists() or not safe_target.is_file():
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        inline = (qs.get("inline") or ["1"])[0] == "1"
                        return self._send_user_file(safe_target, force_inline=inline, allow_range=False)

                    # GET /api/projects/<id>/organizer
                    if len(parts) == 4 and parts[3] == "organizer":
                        page = int((qs.get("page") or ["1"])[0] or 1)
                        per_page = int((qs.get("per_page") or ["50"])[0] or 50)
                        flt = (qs.get("filter") or ["all"])[0]
                        sec_f = (qs.get("section") or [""])[0]
                        q_s = (qs.get("q") or [""])[0]
                        show_res = str((qs.get("show_resolved") or ["0"])[0]).strip().lower() in ("1", "true", "yes")
                        plan_pg = porg.get_organizer_plan_page(
                            proj_root,
                            page=page,
                            per_page=per_page,
                            status_filter=flt,
                            section_filter=sec_f,
                            query=q_s,
                            show_resolved=show_res,
                        )
                        return self._send_json({
                            "ok": True,
                            "job": hub.get_organizer_job_status(pid),
                            **plan_pg,
                        })

                    # GET /api/projects/<id>/whatsapp-archive
                    if len(parts) == 4 and parts[3] == "whatsapp-archive":
                        page = int((qs.get("page") or ["1"])[0] or 1)
                        per_page = int((qs.get("per_page") or ["50"])[0] or 50)
                        flt = (qs.get("filter") or ["all"])[0]
                        m_flt = (qs.get("media") or [""])[0]
                        q_s = (qs.get("q") or [""])[0]
                        hide_exc = str((qs.get("hide_excluded") or ["0"])[0]).strip().lower() in ("1", "true", "yes")
                        wa_pg = waa.get_whatsapp_plan_page(
                            proj_root,
                            page=page,
                            per_page=per_page,
                            status_filter=flt,
                            media_filter=m_flt,
                            query=q_s,
                            hide_excluded=hide_exc,
                        )
                        return self._send_json({"ok": True, "job": hub.get_scoped_job_status("wa", pid), **wa_pg})

                    # GET /api/projects/<id>/whatsapp-archive/download-zip
                    if len(parts) == 5 and parts[3] == "whatsapp-archive" and parts[4] == "download-zip":
                        zip_p = waa.build_whatsapp_export_zip(proj)
                        return self._send_bytes(zip_p.read_bytes(), "application/zip", filename=zip_p.name)

                    # GET /api/projects/<id>/whatsapp-archive/file?id=...
                    if len(parts) == 5 and parts[3] == "whatsapp-archive" and parts[4] == "file":
                        item_id = (qs.get("id") or [""])[0]
                        wa_plan = read_json_safe(waa.whatsapp_plan_path(proj_root), {}) or {}
                        found_wa = next(
                            (it for it in (wa_plan.get("items") or []) if str(it.get("id")) == str(item_id)),
                            None,
                        )
                        if not found_wa:
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        fp = resolve_local_or_unc(str(found_wa.get("final_path") or found_wa.get("src_path") or ""))
                        allowed_wa_bases = [resolve_local_or_unc(proj_root), archive_dir_for_project(proj)]
                        if wa_plan.get("source_dir"):
                            allowed_wa_bases.insert(0, resolve_local_or_unc(str(wa_plan["source_dir"])))
                        safe_wa_target: Path | None = None
                        for b in allowed_wa_bases:
                            try:
                                rel_p = fp.relative_to(b)
                                safe_wa_target = safe_resolve_within(b, rel_p.as_posix(), allow_base=False, reject_symlinks=True)
                                break
                            except Exception:
                                continue
                        if not safe_wa_target or not safe_wa_target.exists() or not safe_wa_target.is_file():
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        return self._send_user_file(safe_wa_target, force_inline=True, allow_range=True)

                    # GET /api/projects/<id>/duplicates
                    if len(parts) == 4 and parts[3] == "duplicates":
                        page = int((qs.get("page") or ["1"])[0] or 1)
                        per_page = int((qs.get("per_page") or ["20"])[0] or 20)
                        flt = (qs.get("filter") or ["all"])[0]
                        q_s = (qs.get("q") or [""])[0]
                        show_rev = str((qs.get("show_reviewed") or ["0"])[0]).strip().lower() in ("1", "true", "yes")
                        dup_pg = dups.get_duplicates_plan_page(
                            proj_root,
                            page=page,
                            per_page=per_page,
                            match_filter=flt,
                            query=q_s,
                            show_reviewed=show_rev,
                        )
                        return self._send_json({"ok": True, "job": hub.get_scoped_job_status("dup", pid), **dup_pg})

                    # GET /api/projects/<id>/duplicates/download-zip
                    if len(parts) == 5 and parts[3] == "duplicates" and parts[4] == "download-zip":
                        zip_p = dups.build_duplicates_unique_zip(proj)
                        return self._send_bytes(zip_p.read_bytes(), "application/zip", filename=zip_p.name)

                    # GET /api/projects/<id>/duplicates/file?id=...
                    if len(parts) == 5 and parts[3] == "duplicates" and parts[4] == "file":
                        item_id = (qs.get("id") or [""])[0]
                        d_plan = read_json_safe(dups.dups_plan_path(proj_root), {}) or {}
                        found_d = None
                        for grp in (d_plan.get("groups") or []):
                            for it in (grp.get("items") or []):
                                if str(it.get("id")) == str(item_id):
                                    found_d = it
                                    break
                            if found_d:
                                break
                        if not found_d:
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        fp = resolve_local_or_unc(str(found_d.get("abs_path") or ""))
                        allowed_d_bases = [resolve_local_or_unc(proj_root), archive_dir_for_project(proj)]
                        if d_plan.get("source_dir"):
                            allowed_d_bases.insert(0, resolve_local_or_unc(str(d_plan["source_dir"])))
                        safe_d_target: Path | None = None
                        for b in allowed_d_bases:
                            try:
                                rel_p = fp.relative_to(b)
                                safe_d_target = safe_resolve_within(b, rel_p.as_posix(), allow_base=False, reject_symlinks=True)
                                break
                            except Exception:
                                continue
                        if not safe_d_target or not safe_d_target.exists() or not safe_d_target.is_file():
                            return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        return self._send_user_file(safe_d_target, force_inline=True, allow_range=True)

                    # GET /api/projects/<id>/sections
                    if len(parts) == 4 and parts[3] == "sections":
                        all_secs = tprof.get_project_sections(proj, style=cur_style, include_hidden=True)
                        return self._send_json({"ok": True, "sections": [s for s in all_secs if not s.get("hidden")], "all_sections": all_secs})

                    # GET /api/projects/<id>/sections/<section>
                    if len(parts) == 5 and parts[3] == "sections":
                        sec_id = parts[4]
                        proj_sections = tprof.get_project_sections(proj, style=cur_style, include_hidden=True)
                        sec_map = {s["id"]: s for s in proj_sections}
                        if sec_id not in sec_map:
                            return self._send_json({"ok": False, "error": _tr("server.section_unknown")}, 404)
                        sec_meta = dict(sec_map[sec_id])
                        tpls = tprof.list_project_templates(proj_root, section=sec_id, device_id=hub.device_id, project_sections=proj_sections)
                        log_data = plog.read_project_log_summary(
                            proj, section_filter=sec_id, query=(qs.get("q") or [""])[0], limit=250, recent_first=True, device_id=hub.device_id
                        )
                        next_ref, _, sync = plog.next_reference_from_logs(
                            proj, prefix=sec_meta["ref_prefix"], device_id=hub.device_id
                        )
                        return self._send_json({
                            "ok": True,
                            "section": sec_meta,
                            "next_ref": next_ref,
                            "next_disc_preview": sync["next_disc_preview"],
                            "next_global_disc_preview": sync.get("next_global_disc_preview", {}),
                            "is_global_disc_series": sync.get("is_global_disc_series", False),
                            "next_arabic_preview": sync["next_arabic_preview"],
                            "templates": tpls,  # ONLY uploaded templates for this section!
                            "rows": log_data["rows"],
                            "has_uploaded_log": log_data["stats"]["has_uploaded_log"],
                        })

                    # GET /api/projects/<id>/letters/links
                    if len(parts) == 5 and parts[3] == "letters" and parts[4] == "links":
                        ref_q = (qs.get("ref_no") or [""])[0]
                        sh_q = (qs.get("sheet") or [""])[0]
                        subj_q = (qs.get("subject") or [""])[0]
                        sec_q = (qs.get("section") or [""])[0]
                        date_q = (qs.get("doc_date") or [""])[0]
                        txt_q = (qs.get("q") or [""])[0]
                        all_letters = llinks.collect_project_letters(proj, device_id=hub.device_id)
                        store = llinks._load_links_store(proj_root)
                        suggestions: list[dict[str, Any]] = []
                        unlinked: list[dict[str, Any]] = []
                        if ref_q or subj_q:
                            suggestions = llinks.suggest_letter_links(
                                proj,
                                ref_no=ref_q,
                                sheet=sh_q,
                                section=sec_q,
                                subject=subj_q,
                                doc_date=date_q,
                                limit=15,
                                min_score=0.45,
                                device_id=hub.device_id,
                            )
                        else:
                            unlinked = llinks.discover_unlinked_matches(proj, min_score=0.58, device_id=hub.device_id)
                        if txt_q:
                            qn = plog.normalize_ar(txt_q)
                            filtered_letters = [
                                x for x in all_letters
                                if qn in plog.normalize_ar(f"{x.get('ref_no','')} {x.get('subject','')} {x.get('party_label','')} {x.get('sheet','')}")
                            ]
                        else:
                            filtered_letters = all_letters
                        return self._send_json({
                            "ok": True,
                            "letters": filtered_letters,
                            "threads": store.get("threads") or [],
                            "suggestions": suggestions,
                            "unlinked_matches": unlinked,
                        })

                    # GET /api/projects/<id>/logs
                    if len(parts) == 4 and parts[3] == "logs":
                        logs = plog.list_project_logs(proj_root, device_id=hub.device_id)
                        cnt = plog.recalculate_counters_from_logs(proj, device_id=hub.device_id)
                        sum_d = plog.read_project_log_summary(proj, limit=0, device_id=hub.device_id)
                        return self._send_json({
                            "ok": True,
                            "logs": logs,
                            "pending_queue": plog.pending_queue_count(proj_root),
                            "counters": cnt,
                            "missing_reply_dates": sum_d["stats"].get("missing_reply_dates", 0),
                            "missing_reply_rows": sum_d.get("missing_reply_rows") or [],
                        })

                    # GET /api/projects/<id>/logs/<name>/map | download
                    if len(parts) == 6 and parts[3] == "logs":
                        lname = safe_template_name(parts[4])
                        act = parts[5]
                        lp, _is_linked = plog.resolve_project_log_path(proj_root, lname)
                        if not lp.exists():
                            return self._send_json({"ok": False, "error": _tr("server.log_not_found")}, 404)
                        if act == "map":
                            sh = (qs.get("sheet") or [""])[0]
                            prev = plog.preview_log_sheet(proj_root, lname, sheet_name=sh)
                            return self._send_json({"ok": True, **prev})
                        if act == "download":
                            ctype = mimetypes.guess_type(lp.name)[0] or "application/octet-stream"
                            return self._send_bytes(lp.read_bytes(), ctype, filename=lp.name)

                    # GET /api/projects/<id>/templates
                    if len(parts) == 4 and parts[3] == "templates":
                        sec_f = (qs.get("section") or [None])[0]
                        proj_sections = tprof.get_project_sections(proj, style=cur_style, include_hidden=True)
                        tpls = tprof.list_project_templates(proj_root, section=sec_f, device_id=hub.device_id, project_sections=proj_sections)
                        return self._send_json({"ok": True, "templates": tpls, "sections": [s for s in proj_sections if not s.get("hidden")]})

                    # GET /api/projects/<id>/templates/<name>/profile | render
                    if len(parts) == 6 and parts[3] == "templates":
                        tname = safe_template_name(parts[4])
                        sub_act = parts[5]
                        if sub_act == "profile":
                            prof = tprof.get_template_profile(proj_root, tname, device_id=hub.device_id)
                            proj_sections = tprof.get_project_sections(proj, style=cur_style, include_hidden=True)
                            sec_map = {s["id"]: s for s in proj_sections}
                            sec_id = str(prof.get("section") or "")
                            pfx = str(prof.get("ref_prefix") or (sec_map.get(sec_id, {}).get("ref_prefix") or "DOC"))
                            next_ref, _, sync = plog.next_reference_from_logs(
                                proj,
                                prefix=pfx,
                                discipline=str(prof.get("default_discipline") or ""),
                                use_discipline_series=bool(prof.get("use_discipline_series", False)),
                                device_id=hub.device_id,
                            )
                            return self._send_json({
                                "ok": True,
                                "profile": prof,
                                "next_ref": next_ref,
                                "next_disc_preview": sync["next_disc_preview"],
                                "next_global_disc_preview": sync.get("next_global_disc_preview", {}),
                                "is_global_disc_series": sync.get("is_global_disc_series", False),
                                "sections": [s for s in proj_sections if not s.get("hidden")],
                                "auto_date_sources": tprof.get_localized_auto_date_sources(cur_style),
                                "auto_project_sources": tprof.get_localized_auto_project_sources(cur_style),
                                "auto_system_sources": tprof.get_localized_auto_system_sources(cur_style),
                            })
                        if sub_act == "render":
                            sheet = (qs.get("sheet") or [""])[0]
                            use_lo = (qs.get("lo") or ["1"])[0] != "0"
                            vis = trender.render_template(proj_root, tname, sheet_name=sheet, try_libreoffice=use_lo)
                            return self._send_json({"ok": True, "render": vis})

                return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)

            except PathTraversalError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 403)
            except PermissionError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 403)
            except KeyError as exc:
                msg = str(exc.args[0]) if exc.args else str(exc)
                return self._send_json({"ok": False, "error": msg}, 404)
            except FileNotFoundError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 404)
            except (ValueError, UnsafeFileError) as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 400)
            except Exception as exc:
                print(f"[Dashboard GET Error] {type(exc).__name__}: {exc}", file=sys.stderr)
                return self._send_json({"ok": False, "error": _tr("common.error_generic")}, 500)

        # ── POST ─────────────────────────────────────────────────────────────
        def do_POST(self) -> None:
            if not self._check_origin():
                return
            parsed = urlparse(self.path)
            try:
                parts = self._split_api_parts(parsed.path)
                ctype_hdr = str(self.headers.get("Content-Type") or "").lower()
                if (
                    len(parts) == 5
                    and parts[0] == "api"
                    and parts[1] == "projects"
                    and parts[3] in ("whatsapp-archive", "duplicates")
                    and parts[4] == "upload-file"
                    and "application/json" not in ctype_hdr
                ):
                    pid_up = parts[2]
                    proj_up = hub.get_project(pid_up)
                    if not proj_up:
                        return self._send_json({"ok": False, "error": _tr("server.project_not_found")}, 404)
                    try:
                        clen = int(self.headers.get("Content-Length") or 0)
                    except ValueError as exc:
                        raise ValueError(_tr("server.invalid_setting")) from exc
                    if clen < 0 or clen > MAX_WA_UPLOAD_BYTES:
                        raise OverflowError(_tr("server.file_too_large"))
                    raw_batch = re.sub(r"[^A-Za-z0-9_-]", "", str(self.headers.get("X-Upload-Batch") or ""))[:40]
                    batch_id = raw_batch or uuid.uuid4().hex[:10]
                    raw_fn = unquote(str(self.headers.get("X-Upload-Filename") or "upload.bin"))
                    raw_rel = unquote(str(self.headers.get("X-Upload-Relpath") or raw_fn)).replace("\\", "/")
                    clean_rel = _sanitize_upload_rel_segments(raw_rel)
                    proj_root_up = Path(str(proj_up["root"])).resolve()
                    up_root = waa.whatsapp_chats_root(proj_root_up) if parts[3] == "whatsapp-archive" else dups.dups_uploads_root(proj_root_up)
                    up_root.mkdir(parents=True, exist_ok=True)
                    batch_dir = safe_resolve_within(up_root, batch_id, allow_base=False, reject_symlinks=True)
                    batch_dir.mkdir(parents=True, exist_ok=True)
                    dest_f = safe_resolve_within(batch_dir, clean_rel, allow_base=False, reject_symlinks=True)
                    dest_f.parent.mkdir(parents=True, exist_ok=True)
                    written = 0
                    with open(dest_f, "wb") as out_fh:
                        rem = clen
                        while rem > 0:
                            chunk = self.rfile.read(min(64 * 1024, rem))
                            if not chunk:
                                break
                            out_fh.write(chunk)
                            written += len(chunk)
                            rem -= len(chunk)
                    return self._send_json({
                        "ok": True,
                        "batch_id": batch_id,
                        "rel_path": clean_rel,
                        "bytes": written,
                        "is_zip": dest_f.suffix.lower() == ".zip",
                    })
                body = self._read_json()

                # POST /api/settings
                if parts == ["api", "settings"]:
                    updated = hub.update_settings(body)
                    return self._send_json({"ok": True, "settings": updated})

                # POST /api/browse-folder
                if parts == ["api", "browse-folder"]:
                    ok_b, path_b, msg_b = browse_local_folder(
                        initial_dir=str(body.get("initial_dir") or ""),
                        title=str(body.get("title") or ""),
                    )
                    if not ok_b and msg_b:
                        return self._send_json({"ok": False, "path": "", "error": msg_b})
                    return self._send_json({"ok": True, "cancelled": not ok_b, "path": path_b})

                # POST /api/browse-file
                if parts == ["api", "browse-file"]:
                    kind = str(body.get("kind") or "excel").strip().lower()
                    if kind == "pdf":
                        flt = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
                    else:
                        flt = "Excel Logs (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*"
                    ok_f, path_f, msg_f = browse_local_file(
                        initial_dir=str(body.get("initial_dir") or ""),
                        title=str(body.get("title") or ""),
                        file_filter=flt,
                    )
                    if not ok_f and msg_f:
                        return self._send_json({"ok": False, "path": "", "error": msg_f})
                    return self._send_json({"ok": True, "cancelled": not ok_f, "path": path_f})

                # POST /api/active-project
                if parts == ["api", "active-project"]:
                    pid = body.get("project_id")
                    if body.get("sync") and pid:
                        return self._send_json(hub.start_project(str(pid)))
                    return self._send_json(hub.set_active_project(str(pid) if pid else None))

                # POST /api/watcher/toggle
                if parts == ["api", "watcher", "toggle"]:
                    en = body.get("enable") if "enable" in body else None
                    state = hub.toggle_watcher(en)
                    return self._send_json({
                        "ok": True,
                        "watcher_running": state,
                        "process_mode": hub.get_settings()["process_mode"],
                        "active_project_id": hub.get_settings()["active_project_id"],
                    })

                # POST /api/app/install-shortcut
                if parts == ["api", "app", "install-shortcut"]:
                    res_sc = create_desktop_app_shortcut()
                    if res_sc.get("ok") and body.get("launch_app_window"):
                        host_hdr = self.headers.get("Host") or "127.0.0.1:8765"
                        res_sc["window"] = open_standalone_app_window(f"http://{host_hdr}")
                    return self._send_json(res_sc, 200 if res_sc.get("ok") else 400)

                # POST /api/app/open-window
                if parts == ["api", "app", "open-window"]:
                    host_hdr = self.headers.get("Host") or "127.0.0.1:8765"
                    res_win = open_standalone_app_window(f"http://{host_hdr}")
                    return self._send_json(res_win)

                # POST /api/app/shutdown
                if parts == ["api", "app", "shutdown"]:
                    hub.stop_all_background()
                    sim = bool(body.get("simulate", False)) or not getattr(hub, "allow_shutdown", True)
                    if not sim:
                        srv = self.server
                        threading.Timer(0.25, srv.shutdown).start()
                    return self._send_json({"ok": True, "shutting_down": True, "simulated": sim})

                # POST /api/system/reset  (factory reset / wipe test data cleanly)
                if parts == ["api", "system", "reset"]:
                    if not bool(body.get("confirm", False)):
                        return self._send_json({"ok": False, "error": _tr("settings.reset_ack_required")}, 400)
                    res_rst = hub.reset_system(
                        delete_project_folders=bool(body.get("delete_project_folders", False)),
                        keep_ai_key=bool(body.get("keep_ai_key", True)),
                        keep_whatsapp=bool(body.get("keep_whatsapp", True)),
                        reset_settings=bool(body.get("reset_settings", True)),
                    )
                    return self._send_json(res_rst)

                # POST /api/ai-key (stored strictly in ~/.gemini/ outside all project & hub files!)
                if parts == ["api", "ai-key"]:
                    if body.get("clear_all"):
                        sw.save_gemini_key("", "", backup_api_key="")
                        return self._send_json({"ok": True, **sw.get_gemini_keys_status()})
                    if body.get("clear_backup"):
                        sw.save_gemini_backup_key("")
                        return self._send_json({"ok": True, **sw.get_gemini_keys_status()})
                    if body.get("clear"):
                        sw.save_gemini_key("", "")
                        return self._send_json({"ok": True, **sw.get_gemini_keys_status()})
                    if "remove_backup_index" in body and body.get("remove_backup_index") is not None:
                        sw.remove_gemini_backup_key(int(body.get("remove_backup_index")))
                        return self._send_json({"ok": True, **sw.get_gemini_keys_status()})
                    if "promote_backup_index" in body and body.get("promote_backup_index") is not None:
                        sw.promote_gemini_backup_key(int(body.get("promote_backup_index")))
                        return self._send_json({"ok": True, **sw.get_gemini_keys_status()})

                    add_raw = body.get("add_backup_keys")
                    if add_raw is not None:
                        tokens = sw._parse_key_tokens(add_raw)
                        if not tokens or any(len(tok) < 20 for tok in tokens):
                            return self._send_json({"ok": False, "error": _tr("server.ai_key_short")}, 400)
                        top_model = ""
                        if body.get("verify", False):
                            models = sw.discover_free_gemini_models(tokens[0], force_refresh=True)
                            top_model = models[0] if models else ""
                        sw.add_gemini_backup_keys(tokens, top_model=top_model)
                        return self._send_json({"ok": True, **sw.get_gemini_keys_status()})

                    has_prim_arg = "api_key" in body and str(body.get("api_key") or "").strip() != ""
                    has_bak_arg = "backup_api_key" in body and str(body.get("backup_api_key") or "").strip() != ""
                    has_baks_arg = "backup_api_keys" in body and isinstance(body.get("backup_api_keys"), list)
                    raw_key = str(body.get("api_key") or "").strip() if has_prim_arg else None
                    bak_key = str(body.get("backup_api_key") or "").strip() if has_bak_arg else None
                    bak_list = sw._parse_key_tokens(body.get("backup_api_keys")) if has_baks_arg else None

                    if raw_key is None and bak_key is None and bak_list is None:
                        return self._send_json({"ok": False, "error": _tr("server.ai_key_short")}, 400)
                    if raw_key is not None and len(raw_key) < 20:
                        return self._send_json({"ok": False, "error": _tr("server.ai_key_short")}, 400)
                    if bak_key is not None and any(len(tok) < 20 for tok in sw._parse_key_tokens(bak_key)):
                        return self._send_json({"ok": False, "error": _tr("server.ai_key_short")}, 400)
                    if bak_list is not None and any(len(tok) < 20 for tok in bak_list):
                        return self._send_json({"ok": False, "error": _tr("server.ai_key_short")}, 400)

                    top_model = ""
                    probe_k = raw_key or (sw._parse_key_tokens(bak_key)[0] if bak_key else (bak_list[0] if bak_list else ""))
                    if body.get("verify", False) and probe_k:
                        models = sw.discover_free_gemini_models(probe_k, force_refresh=True)
                        top_model = models[0] if models else ""
                    sw.save_gemini_key(raw_key, top_model, backup_api_key=bak_key, backup_api_keys=bak_list)
                    return self._send_json({"ok": True, **sw.get_gemini_keys_status()})

                # POST /api/decisions/ignore-all
                if parts == ["api", "decisions", "ignore-all"]:
                    cleared = hub.decisions.ignore_all_pending()
                    closed_items = list(getattr(hub.decisions, "_last_bulk_closed", []) or [])
                    if closed_items:
                        all_p = hub.list_projects()
                        for ci in closed_items:
                            for p_item in all_p:
                                if not ci.get("project_id") or str(p_item.get("id")) == str(ci.get("project_id")):
                                    sync_decision_to_triage(
                                        p_item,
                                        file_sha256=str(ci.get("file_sha256") or ""),
                                        file_path=str(ci.get("file_path") or ""),
                                        action="ignore_once",
                                    )
                    return self._send_json({"ok": True, "cleared": cleared})

                # POST /api/decisions/<id>/resolve
                if len(parts) == 4 and parts[:2] == ["api", "decisions"] and parts[3] == "resolve":
                    dec_id = parts[2]
                    act = str(body.get("action") or "").strip()
                    res = resolve_pending_decision(
                        dec_id,
                        action=act,
                        params=body,
                        decision_store=hub.decisions,
                        all_projects=hub.list_projects(),
                        global_settings=hub.get_settings(),
                        device_id=hub.device_id,
                    )
                    dec_obj = res.get("decision") or {}
                    src_pid = str(dec_obj.get("source_project_id") or "")
                    res_pid = str(dec_obj.get("resolved_project_id") or "")
                    for p_item in hub.list_projects():
                        if str(p_item.get("id")) in (src_pid, res_pid) or not src_pid:
                            sync_decision_to_triage(
                                p_item,
                                file_sha256=str(dec_obj.get("file_sha256") or ""),
                                file_path=str(dec_obj.get("file_path") or ""),
                                action=act,
                            )
                    if act == "ignore_pattern" and res["result"].get("pattern"):
                        cur = list(hub.get_settings().get("ignore_patterns") or [])
                        pat = str(res["result"]["pattern"])
                        if pat not in cur:
                            cur.append(pat)
                            hub.update_settings({"ignore_patterns": cur})
                    if res_pid:
                        hub.refresh_project_counters(res_pid)
                    return self._send_json(res)

                # POST /api/rename-journal/<id>/undo
                if len(parts) == 4 and parts[:2] == ["api", "rename-journal"] and parts[3] == "undo":
                    entry_id = parts[2]
                    pid_hint = str(body.get("project_id") or "")
                    for p in hub.list_projects():
                        if pid_hint and p["id"] != pid_hint:
                            continue
                        jp = rename_journal_path_for(Path(str(p.get("inbox_dir") or (Path(p["root"]) / "1_Inbox"))))
                        if not jp.exists():
                            continue
                        try:
                            undone = undo_rename_entry(jp, entry_id, device_id=hub.device_id)
                            return self._send_json({"ok": True, "entry": undone})
                        except KeyError:
                            continue
                    return self._send_json({"ok": False, "error": _tr("server.journal_not_found")}, 404)

                # POST /api/learned-rules  |  POST /api/learned-rules/<id>
                if parts == ["api", "learned-rules"]:
                    rule = hub.decisions.add_rule(
                        match_type=str(body.get("match_type") or "entity"),
                        pattern=str(body.get("pattern") or ""),
                        action=str(body.get("action") or "rename_in_place"),
                        target_project_id=str(body.get("target_project_id") or ""),
                        entity_name=str(body.get("entity_name") or ""),
                        note=str(body.get("note") or ""),
                        enabled=bool(body.get("enabled", True)),
                    )
                    return self._send_json({"ok": True, "rule": rule})

                if len(parts) == 3 and parts[:2] == ["api", "learned-rules"]:
                    rid = parts[2]
                    if body.get("delete"):
                        ok = hub.decisions.delete_rule(rid)
                        return self._send_json({"ok": ok})
                    updated_r = hub.decisions.update_rule(rid, body)
                    return self._send_json({"ok": True, "rule": updated_r})

                # POST /api/tasks  |  POST /api/tasks/auto-dismiss  |  POST /api/tasks/<id>
                if parts == ["api", "tasks"]:
                    proj_lookup = {p["id"]: p for p in hub.list_projects()}
                    task = tbar.create_manual_task(hub.state_dir, body, project_lookup=proj_lookup)
                    return self._send_json({"ok": True, "task": task}, 201)

                if parts == ["api", "tasks", "auto-dismiss"]:
                    res_dis = tbar.dismiss_auto_task(
                        hub.state_dir,
                        str(body.get("auto_key") or ""),
                        restore=bool(body.get("restore", False)),
                        meta=body.get("meta") if isinstance(body.get("meta"), dict) else None,
                    )
                    return self._send_json({"ok": True, **res_dis})

                if parts == ["api", "tasks", "clear-done"]:
                    res_cl = tbar.clear_completed_manual_tasks(
                        hub.state_dir,
                        project_id=str(body.get("project_id") or ""),
                    )
                    return self._send_json({"ok": True, **res_cl})

                if len(parts) == 3 and parts[:2] == ["api", "tasks"]:
                    tid = parts[2]
                    if body.get("delete"):
                        deleted = tbar.delete_manual_task(hub.state_dir, tid)
                        if not deleted:
                            return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                        return self._send_json({"ok": True, "deleted": True})
                    proj_lookup = {p["id"]: p for p in hub.list_projects()}
                    try:
                        updated_t = tbar.update_manual_task(hub.state_dir, tid, body, project_lookup=proj_lookup)
                    except KeyError:
                        return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                    return self._send_json({"ok": True, "task": updated_t})

                # POST /api/whatsapp/config
                if parts == ["api", "whatsapp", "config"]:
                    if not wab:
                        return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                    wa_path = hub.state_dir / "whatsapp_config.json"
                    cur_wa = wab.load_wa_config(wa_path)
                    saved_wa = wab.save_wa_config(wa_path, cur_wa, body)
                    return self._send_json({"ok": True, "config": wab.public_wa_config(saved_wa)})

                # POST /api/projects (Create new project)
                if parts == ["api", "projects"]:
                    proj = hub.create_project(body)
                    return self._send_json({"ok": True, "project": proj}, 201)

                # Project-scoped POST routes: /api/projects/<id>/...
                if len(parts) >= 4 and parts[1] == "projects":
                    pid = parts[2]
                    proj = hub.get_project(pid)
                    if not proj:
                        return self._send_json({"ok": False, "error": _tr("server.project_not_found")}, 404)
                    proj_root = Path(str(proj["root"])).resolve()
                    action = parts[3]

                    # POST /api/projects/<id>/start  (بدء العمل / مزامنة)
                    if len(parts) == 4 and action == "start":
                        return self._send_json(hub.start_project(pid))

                    # POST /api/projects/<id>/end  (إنهاء الجلسة + نسخ احتياطي)
                    if len(parts) == 4 and action == "end":
                        return self._send_json(
                            hub.end_project(
                                pid,
                                copy_external=bool(body.get("copy_external", True)),
                                include_archive_external=bool(body.get("include_archive_external", False)),
                            )
                        )

                    # POST /api/projects/<id>/backup (نسخ احتياطي فوري من البطاقة)
                    if len(parts) == 4 and action == "backup":
                        ext_override = str(body.get("external_backup_dir") or "").strip()
                        if ext_override and bool(body.get("save_external_dir", False)):
                            hub.update_settings({"external_backup_dir": ext_override})
                        return self._send_json(
                            hub.quick_backup_project(
                                pid,
                                copy_external=bool(body.get("copy_external", True)),
                                include_archive_external=bool(body.get("include_archive_external", False)),
                                external_dir_override=ext_override,
                            )
                        )

                    # POST /api/projects/<id>/reports/weekly (تقرير المتابعة الأسبوعي)
                    if len(parts) == 5 and action == "reports" and parts[4] == "weekly":
                        rep = porg.generate_weekly_followup_report(
                            proj,
                            device_id=hub.device_id,
                            export_pdf=bool(body.get("export_pdf", True)),
                            days_window=max(1, min(30, int(body.get("days_window") or 7))),
                        )
                        rep["xlsx_download_url"] = (
                            f"/api/projects/{pid}/download?rel={quote(rep['xlsx_rel'], safe='/')}"
                            if rep.get("xlsx_rel") else ""
                        )
                        rep["pdf_download_url"] = (
                            f"/api/projects/{pid}/download?rel={quote(rep['pdf_rel'], safe='/')}"
                            if rep.get("pdf_rel") else ""
                        )
                        rep["pdf_preview_url"] = (
                            f"/api/projects/{pid}/preview-info?rel={quote(rep['pdf_rel'], safe='/')}"
                            if rep.get("pdf_rel") else ""
                        )
                        return self._send_json(rep)

                    # POST /api/projects/<id>/update | settings
                    if len(parts) == 4 and action in ("update", "settings"):
                        updated_p = hub.update_project(pid, body)
                        return self._send_json({"ok": True, "project": scrub_secrets(updated_p)})

                    # POST /api/projects/<id>/set-counters
                    if len(parts) == 4 and action == "set-counters":
                        hub.update_project(pid, {
                            "baseline_counters": body.get("counters") or body.get("baseline_counters") or {},
                            "baseline_disc_counters": body.get("disc_counters") or body.get("baseline_disc_counters") or {},
                        })
                        cnt = hub.refresh_project_counters(pid)
                        return self._send_json({"ok": True, "counters": cnt})

                    # POST /api/projects/<id>/unlink (zero deletion!)
                    if len(parts) == 4 and action == "unlink":
                        return self._send_json({"ok": hub.unlink_project(pid)})

                    # POST /api/projects/<id>/delete (unlink or delete project folder safely)
                    if len(parts) == 4 and action == "delete":
                        del_files = bool(body.get("delete_files", False))
                        if del_files:
                            ack = bool(body.get("confirm", False))
                            typed = str(body.get("confirm_name") or "").strip().lower()
                            allowed_names = {
                                str(proj.get("code") or "").strip().lower(),
                                str(proj.get("name") or "").strip().lower(),
                            } - {""}
                            if not ack or not typed or typed not in allowed_names:
                                return self._send_json({"ok": False, "error": _tr("hub.delete_confirm_name_err")}, 400)
                        return self._send_json(
                            hub.delete_project(pid, delete_files=del_files)
                        )

                    # POST /api/projects/<id>/toggle-enabled (pause/disable or re-enable project)
                    if len(parts) == 4 and action == "toggle-enabled":
                        en_val = bool(body.get("enabled", not bool(proj.get("enabled", True))))
                        updated_p = hub.update_project(pid, {"enabled": en_val})
                        return self._send_json({"ok": True, "enabled": en_val, "project": scrub_secrets(updated_p)})

                    # POST /api/projects/<id>/organize (scan inbox now)
                    if len(parts) == 4 and action == "organize":
                        blocking = bool(body.get("blocking", False))
                        return self._send_json(hub.run_inbox_sync_now(pid, blocking=blocking))

                    # POST /api/projects/<id>/import-legacy
                    if len(parts) == 4 and action == "import-legacy":
                        sub = str(body.get("mode") or "scan")
                        if sub == "stop":
                            return self._send_json({"ok": hub.stop_legacy_import(pid)})
                        folder = str(body.get("folder") or "").strip()
                        if sub == "start":
                            return self._send_json(hub.start_legacy_import(
                                pid, folder,
                                max_files=int(body.get("max_files") or 30),
                                pace_seconds=float(body.get("pace_seconds") or 4.0),
                            ))
                        return self._send_json(hub.scan_legacy_folder(pid, folder))

                    # POST /api/projects/<id>/upload (drop files into inbox / templates / logs)
                    if len(parts) == 4 and action == "upload":
                        target_kind = str(body.get("target") or "inbox").strip()
                        files_in = body.get("files") or []
                        if not isinstance(files_in, list) or not files_in:
                            return self._send_json({"ok": False, "error": _tr("server.no_files_uploaded")}, 400)
                        saved_items: list[dict[str, Any]] = []
                        for fobj in files_in:
                            raw_nm = str(fobj.get("name") or "")
                            raw_b = base64.b64decode(str(fobj.get("data_b64") or fobj.get("data") or ""))
                            if target_kind == "logs":
                                lm = plog.upload_project_log(proj_root, raw_nm, raw_b, device_id=hub.device_id)
                                saved_items.append({"name": lm["log_name"], "sheets": len(lm["sheets"]), "has_vba": lm["has_vba"]})
                            elif target_kind == "templates":
                                proj_sections = tprof.get_project_sections(proj, style=hub.get_settings().get("ui_style"), include_hidden=True)
                                tr_up = tprof.upload_project_template(
                                    proj_root, raw_nm, raw_b,
                                    section=str(fobj.get("section") or body.get("section") or ""),
                                    device_id=hub.device_id,
                                    project_sections=proj_sections,
                                )
                                saved_items.append(tr_up)
                            else:
                                safe_nm = sanitize_upload_filename(raw_nm)
                                ibx_dir = resolve_local_or_unc(str(proj.get("inbox_dir") or (proj_root / "1_Inbox")))
                                ibx_dir.mkdir(parents=True, exist_ok=True)
                                dst = safe_resolve_within(ibx_dir, safe_nm)
                                dst.write_bytes(raw_b)
                                saved_items.append({"name": dst.name})
                        if target_kind == "logs":
                            plog.flush_project_log_queue(proj, device_id=hub.device_id)
                            hub.refresh_project_counters(pid)
                        return self._send_json({"ok": True, "saved": saved_items})

                    # ── Correspondence Cross-Linking Endpoints ────────────────
                    if action == "letters":
                        if len(parts) == 5 and parts[4] == "link":
                            res_lk = llinks.link_letters(
                                proj,
                                body.get("items") or [],
                                topic=str(body.get("topic") or ""),
                                sync_excel=bool(body.get("sync_excel", True)),
                                device_id=hub.device_id,
                            )
                            return self._send_json(res_lk)
                        if len(parts) == 5 and parts[4] == "unlink":
                            res_ul = llinks.unlink_letter(
                                proj,
                                ref_no=str(body.get("ref_no") or ""),
                                sheet=str(body.get("sheet") or ""),
                                target_ref_no=str(body.get("target_ref_no") or ""),
                                target_sheet=str(body.get("target_sheet") or ""),
                                sync_excel=bool(body.get("sync_excel", True)),
                                device_id=hub.device_id,
                            )
                            return self._send_json(res_ul)
                        if len(parts) == 5 and parts[4] == "auto-link":
                            res_au = llinks.auto_link_project_letters(
                                proj,
                                min_score=float(body.get("min_score") or 0.58),
                                apply=bool(body.get("apply", True)),
                                sync_excel=bool(body.get("sync_excel", True)),
                                pairs=body.get("pairs") if isinstance(body.get("pairs"), list) else None,
                                device_id=hub.device_id,
                            )
                            return self._send_json(res_au)

                    # ── Project Uploaded Logs Endpoints ──────────────────────
                    if action == "logs":
                        # POST /api/projects/<id>/logs/upload
                        if len(parts) == 5 and parts[4] == "upload":
                            raw_nm = str(body.get("name") or "")
                            raw_b = base64.b64decode(str(body.get("data_b64") or ""))
                            lmap = plog.upload_project_log(proj_root, raw_nm, raw_b, device_id=hub.device_id)
                            flushed = plog.flush_project_log_queue(proj, device_id=hub.device_id)
                            cnt = hub.refresh_project_counters(pid)
                            return self._send_json({
                                "ok": True,
                                "log_name": lmap["log_name"],
                                "logmap": lmap,
                                "queue_flush": flushed,
                                "counters": cnt,
                            })
                        # POST /api/projects/<id>/logs/link (direct link to local/server Excel without copying!)
                        if len(parts) == 5 and parts[4] == "link":
                            raw_p = str(body.get("path") or "").strip()
                            if _has_invalid_win_question_marks(raw_p):
                                return self._send_json({"ok": False, "error": _tr("server.invalid_path_question_marks")}, 400)
                            lmap = plog.link_external_project_log(proj_root, raw_p, device_id=hub.device_id)
                            flushed = plog.flush_project_log_queue(proj, device_id=hub.device_id)
                            cnt = hub.refresh_project_counters(pid)
                            return self._send_json({
                                "ok": True,
                                "log_name": lmap["log_name"],
                                "logmap": lmap,
                                "queue_flush": flushed,
                                "counters": cnt,
                            })
                        # POST /api/projects/<id>/logs/flush
                        if len(parts) == 5 and parts[4] == "flush":
                            flushed = plog.flush_project_log_queue(proj, device_id=hub.device_id)
                            cnt = hub.refresh_project_counters(pid)
                            return self._send_json({"ok": True, "queue_flush": flushed, "counters": cnt})
                        # POST /api/projects/<id>/logs/set-reply-date
                        if len(parts) == 5 and parts[4] == "set-reply-date":
                            res_d = plog.update_log_row_reply_date(
                                proj,
                                log_name=str(body.get("log_name") or ""),
                                sheet=str(body.get("sheet") or ""),
                                row=int(body.get("row") or 0),
                                consultant_reply_date=str(body.get("consultant_reply_date") or ""),
                                moh_reply_date=str(body.get("moh_reply_date") or ""),
                                revision=str(body.get("revision") or ""),
                                device_id=hub.device_id,
                            )
                            sum_d = plog.read_project_log_summary(proj, limit=0, device_id=hub.device_id)
                            return self._send_json({
                                **res_d,
                                "missing_reply_dates": sum_d["stats"].get("missing_reply_dates", 0),
                                "missing_reply_rows": sum_d.get("missing_reply_rows") or [],
                            })
                        # POST /api/projects/<id>/logs/<name>/map | archive
                        if len(parts) == 6:
                            lname = safe_template_name(parts[4])
                            sub_act = parts[5]
                            lp, _is_linked = plog.resolve_project_log_path(proj_root, lname)
                            if sub_act == "map":
                                saved_m = plog.save_logmap(lp, body, device_id=hub.device_id, proj_root=proj_root)
                                flushed = plog.flush_project_log_queue(proj, device_id=hub.device_id)
                                cnt = hub.refresh_project_counters(pid)
                                return self._send_json({"ok": True, "logmap": saved_m, "queue_flush": flushed, "counters": cnt})
                            if sub_act == "archive":
                                arc = plog.archive_project_log(proj_root, lname, device_id=hub.device_id)
                                cnt = hub.refresh_project_counters(pid)
                                return self._send_json({"ok": True, "archived_to": arc.name, "counters": cnt})

                    # ── Large-Scale Project Organizer & Audit Endpoints ─────
                    if action == "organizer":
                        if len(parts) == 5 and parts[4] == "scan":
                            res = hub.start_organizer_scan(
                                pid,
                                str(body.get("folder") or ""),
                                extract_archives=bool(body.get("extract_archives", True)),
                                read_text=bool(body.get("read_text", True)),
                                blocking=bool(body.get("blocking", False)),
                            )
                            return self._send_json(res)
                        if len(parts) == 5 and parts[4] == "update":
                            res = porg.update_organizer_plan_items(
                                proj,
                                body.get("updates"),
                                bulk_action=str(body.get("bulk_action") or ""),
                            )
                            return self._send_json({"ok": True, **res})
                        if len(parts) == 5 and parts[4] == "execute":
                            res = hub.start_organizer_execute(
                                pid,
                                mode=str(body.get("mode") or "copy"),
                                update_excel_log=bool(body.get("update_excel_log", True)),
                                blocking=bool(body.get("blocking", False)),
                            )
                            return self._send_json(res)
                        if len(parts) == 5 and parts[4] == "stop":
                            return self._send_json({"ok": hub.stop_organizer_job(pid)})
                        if len(parts) == 5 and parts[4] == "job-dismiss":
                            hub.dismiss_organizer_job(pid)
                            return self._send_json({"ok": True})
                        if len(parts) == 5 and parts[4] == "ai-rename":
                            raw_ids = body.get("item_ids")
                            ids_list = [str(x) for x in raw_ids if str(x).strip()] if isinstance(raw_ids, list) else None
                            res = hub.start_organizer_ai_rename(
                                pid,
                                item_ids=ids_list,
                                scope=str(body.get("scope") or "unclassified"),
                                limit=max(1, min(100, int(body.get("limit") or 25))),
                                force_refresh=bool(body.get("force_refresh", False)),
                                blocking=bool(body.get("blocking", False)),
                            )
                            return self._send_json(res)
                        if len(parts) == 5 and parts[4] == "audit":
                            res = porg.audit_archive_and_generate_weekly_report(proj, device_id=hub.device_id)
                            return self._send_json(res)
                        if len(parts) == 5 and parts[4] == "register-unlogged":
                            res = porg.register_unlogged_archive_files(proj, device_id=hub.device_id)
                            hub.refresh_project_counters(pid)
                            return self._send_json(res)
                        if len(parts) == 5 and parts[4] == "pdf-split":
                            src_pdf = str(body.get("src_pdf") or "").strip()
                            dest_dir = str(body.get("dest_dir") or "").strip() or str(
                                resolve_local_or_unc(str(proj.get("inbox_dir") or (proj_root / "1_Inbox")))
                            )
                            res = porg.split_pdf_document(
                                src_pdf,
                                dest_dir,
                                every_n_pages=int(body.get("every_n_pages") or 1),
                                ranges_spec=str(body.get("ranges_spec") or ""),
                            )
                            file_links: list[dict[str, str]] = []
                            ddir_p = resolve_local_or_unc(str(res.get("dest_dir") or dest_dir))
                            arc_d_root = archive_dir_for_project(proj)
                            for fname in (res.get("files") or []):
                                fp = ddir_p / fname
                                if not fp.exists() or not fp.is_file():
                                    continue
                                dl_u = ""
                                try:
                                    rel_r = fp.relative_to(proj_root).as_posix()
                                    dl_u = f"/api/projects/{pid}/download?rel={quote(rel_r, safe='/')}&area=root"
                                except ValueError:
                                    try:
                                        rel_a = fp.relative_to(arc_d_root).as_posix()
                                        dl_u = f"/api/projects/{pid}/download?rel={quote(rel_a, safe='/')}&area=archive"
                                    except ValueError:
                                        dl_u = ""
                                file_links.append({"name": fname, "download_url": dl_u})
                            res["file_links"] = file_links
                            return self._send_json(res)
                        if len(parts) == 5 and parts[4] == "pdf-merge":
                            srcs = [str(x).strip() for x in (body.get("src_pdfs") or []) if str(x).strip()]
                            out_raw = str(body.get("out_pdf") or "").strip()
                            if not out_raw:
                                arc_d = archive_dir_for_project(proj) / "_Merged_PDFs"
                                out_raw = str(arc_d / f"Merged_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
                            res = porg.merge_pdf_documents(srcs, out_raw)
                            out_p = resolve_local_or_unc(str(res.get("output_path") or out_raw))
                            arc_d_root = archive_dir_for_project(proj)
                            dl_u = ""
                            out_rel = ""
                            if out_p.exists() and out_p.is_file():
                                try:
                                    out_rel = out_p.relative_to(proj_root).as_posix()
                                    dl_u = f"/api/projects/{pid}/download?rel={quote(out_rel, safe='/')}&area=root"
                                except ValueError:
                                    try:
                                        out_rel = out_p.relative_to(arc_d_root).as_posix()
                                        dl_u = f"/api/projects/{pid}/download?rel={quote(out_rel, safe='/')}&area=archive"
                                    except ValueError:
                                        dl_u = ""
                            res["output_rel"] = out_rel
                            res["download_url"] = dl_u
                            return self._send_json(res)

                    # ── WhatsApp Media & Chat Archiver ─────────────────────
                    if action == "whatsapp-archive":
                        if len(parts) == 5 and parts[4] == "upload-file":
                            raw_batch = re.sub(r"[^A-Za-z0-9_-]", "", str(body.get("batch_id") or ""))[:40]
                            batch_id = raw_batch or uuid.uuid4().hex[:10]
                            wa_root_up = waa.whatsapp_chats_root(proj_root)
                            wa_root_up.mkdir(parents=True, exist_ok=True)
                            batch_dir = safe_resolve_within(wa_root_up, batch_id, allow_base=False, reject_symlinks=True)
                            batch_dir.mkdir(parents=True, exist_ok=True)
                            files_list = body.get("files") if isinstance(body.get("files"), list) else [body]
                            saved_count = 0
                            for ent in files_list:
                                if not isinstance(ent, dict):
                                    continue
                                raw_fn = str(ent.get("rel_path") or ent.get("name") or "").replace("\\", "/")
                                clean_rel = _sanitize_upload_rel_segments(raw_fn)
                                dest_f = safe_resolve_within(batch_dir, clean_rel, allow_base=False, reject_symlinks=True)
                                dest_f.parent.mkdir(parents=True, exist_ok=True)
                                try:
                                    raw_bytes = base64.b64decode(str(ent.get("data_b64") or ""), validate=True)
                                except Exception as exc:
                                    raise ValueError(_tr("server.invalid_setting")) from exc
                                if len(raw_bytes) > MAX_WA_UPLOAD_BYTES:
                                    raise OverflowError(_tr("server.file_too_large"))
                                dest_f.write_bytes(raw_bytes)
                                saved_count += 1
                            return self._send_json({
                                "ok": True,
                                "batch_id": batch_id,
                                "saved_count": saved_count,
                                "batch_dir": str(batch_dir),
                            })

                        if len(parts) == 5 and parts[4] == "scan":
                            src_in = str(body.get("source_path") or "").strip()
                            batch_id = re.sub(r"[^A-Za-z0-9_-]", "", str(body.get("batch_id") or ""))[:40]
                            batch_lbl = str(body.get("batch_label") or "").strip()
                            bdir: Path | None = None
                            if not src_in and batch_id:
                                wa_root_up = waa.whatsapp_chats_root(proj_root)
                                bdir = safe_resolve_within(wa_root_up, batch_id, allow_base=False, reject_symlinks=True)
                                if not bdir.exists():
                                    return self._send_json({"ok": False, "error": _tr("wa.no_files_selected")}, 400)
                            elif not src_in:
                                return self._send_json({"ok": False, "error": _tr("wa.no_files_selected")}, 400)

                            win_min = int(body["window_minutes"]) if body.get("window_minutes") is not None else 0
                            s_mode = str(body.get("sender_mode") or "same_sender")
                            t_prio = str(body.get("text_priority") or "before_then_after")
                            n_fmt = str(body.get("name_format") or "date_name")
                            d_from = str(body.get("date_from") or "")
                            d_to = str(body.get("date_to") or "")
                            inc_docs = bool(body.get("include_docs", True))
                            blocking = bool(body.get("blocking", True))

                            def _do_wa_scan(stop_ev, on_prog):
                                eff_src = src_in
                                eff_lbl = batch_lbl
                                if not eff_src and bdir is not None:
                                    all_up = [p for p in bdir.rglob("*") if p.is_file()]
                                    zips = [p for p in all_up if p.suffix.lower() == ".zip"]
                                    non_zips = [p for p in all_up if p.suffix.lower() != ".zip"]
                                    if len(zips) == 1 and not non_zips:
                                        eff_src = str(zips[0])
                                        eff_lbl = eff_lbl or zips[0].name
                                    else:
                                        for zp in zips:
                                            on_prog({"phase": "extracting", "current_file": zp.name})
                                            waa.extract_zip_safely(zp, bdir / f"_extracted_{zp.stem[:24]}")
                                        eff_src = str(bdir)
                                return waa.scan_and_match_whatsapp_source(
                                    hub.get_project(pid) or proj,
                                    eff_src,
                                    window_minutes=win_min,
                                    sender_mode=s_mode,
                                    text_priority=t_prio,
                                    name_format=n_fmt,
                                    date_from=d_from,
                                    date_to=d_to,
                                    include_docs=inc_docs,
                                    batch_label=eff_lbl,
                                    stop_event=stop_ev,
                                    progress_cb=on_prog,
                                )

                            res = hub._run_scoped_job("wa", pid, "scan", "walking", _do_wa_scan, blocking=blocking)
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "recompute":
                            res = waa.recompute_whatsapp_plan(
                                proj,
                                window_minutes=body.get("window_minutes"),
                                sender_mode=body.get("sender_mode"),
                                text_priority=body.get("text_priority"),
                                name_format=body.get("name_format"),
                                date_from=body.get("date_from"),
                                date_to=body.get("date_to"),
                                include_docs=body.get("include_docs"),
                            )
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "update":
                            res = waa.update_whatsapp_plan_items(
                                proj,
                                body.get("updates"),
                                bulk_action=str(body.get("bulk_action") or ""),
                            )
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "ai-rename":
                            g = hub.get_settings()
                            raw_ids = body.get("item_ids")
                            ids_list = [str(x) for x in raw_ids if str(x).strip()] if isinstance(raw_ids, list) else None
                            ollama_url = str(g.get("ollama_url") or "http://127.0.0.1:11434").strip()
                            if ollama_url and not ollama_url.rstrip("/").endswith("/api/chat"):
                                ollama_url = ollama_url.rstrip("/") + "/api/chat"
                            scope_v = str(body.get("scope") or "date_only")
                            limit_v = max(1, min(100, int(body.get("limit") or 25)))
                            engine_v = str(g.get("engine") or "gemini")
                            omodel_v = str(g.get("ollama_model") or "qwen2.5vl:7b")
                            blocking = bool(body.get("blocking", True))
                            res = hub._run_scoped_job(
                                "wa",
                                pid,
                                "ai_rename",
                                "ai_renaming",
                                lambda stop_ev, on_prog: waa.ai_rename_whatsapp_items(
                                    hub.get_project(pid) or proj,
                                    item_ids=ids_list,
                                    scope=scope_v,
                                    limit=limit_v,
                                    engine=engine_v,
                                    ollama_url=ollama_url,
                                    ollama_model=omodel_v,
                                    force_refresh=bool(body.get("force_refresh", False)),
                                    stop_event=stop_ev,
                                    progress_cb=on_prog,
                                ),
                                blocking=blocking,
                                total_hint=len(ids_list) if ids_list else limit_v,
                            )
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "execute":
                            mode_v = str(body.get("mode") or "rename_in_place")
                            dest_v = str(body.get("dest_folder") or "")
                            blocking = bool(body.get("blocking", True))
                            res = hub._run_scoped_job(
                                "wa",
                                pid,
                                "execute",
                                "executing",
                                lambda stop_ev, on_prog: waa.execute_whatsapp_archive_plan(
                                    hub.get_project(pid) or proj,
                                    mode=mode_v,
                                    dest_folder=dest_v,
                                    stop_event=stop_ev,
                                    progress_cb=on_prog,
                                ),
                                blocking=blocking,
                                refresh_counters=True,
                            )
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "stop":
                            return self._send_json({"ok": hub.stop_scoped_job("wa", pid), "job": hub.get_scoped_job_status("wa", pid)})

                        if len(parts) == 5 and parts[4] == "job-dismiss":
                            hub.dismiss_scoped_job("wa", pid)
                            return self._send_json({"ok": True})

                        if len(parts) == 5 and parts[4] == "undo":
                            res = waa.undo_last_whatsapp_rename(proj)
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "clear":
                            res = waa.clear_whatsapp_plan(proj)
                            return self._send_json(res)

                    # ── Content-Based Duplicates & Visual Similarity Reviewer ──
                    if action == "duplicates":
                        if len(parts) == 5 and parts[4] == "scan":
                            raw_sim = body.get("min_similarity")
                            sim_val = 85 if raw_sim is None else int(raw_sim)
                            s_kind = str(body.get("source_kind") or "folder")
                            s_path = str(body.get("source_path") or "")
                            b_id = str(body.get("batch_id") or "")
                            sc_v = str(body.get("scope") or "all")
                            inc_sub = bool(body.get("include_subfolders", True))
                            blocking = bool(body.get("blocking", True))
                            res = hub._run_scoped_job(
                                "dup",
                                pid,
                                "scan",
                                "walking",
                                lambda stop_ev, on_prog: dups.scan_content_duplicates(
                                    hub.get_project(pid) or proj,
                                    source_kind=s_kind,
                                    source_path=s_path,
                                    batch_id=b_id,
                                    scope=sc_v,
                                    min_similarity=sim_val,
                                    include_subfolders=inc_sub,
                                    stop_event=stop_ev,
                                    progress_cb=on_prog,
                                ),
                                blocking=blocking,
                            )
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "update":
                            res = dups.update_duplicates_plan(
                                proj,
                                item_id=str(body.get("item_id") or ""),
                                group_id=str(body.get("group_id") or ""),
                                action=str(body.get("action") or ""),
                                bulk_action=str(body.get("bulk_action") or ""),
                            )
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "execute":
                            mode_v = str(body.get("mode") or "quarantine")
                            conf_del = bool(body.get("confirm_delete", False))
                            blocking = bool(body.get("blocking", True))
                            res = hub._run_scoped_job(
                                "dup",
                                pid,
                                "execute",
                                "executing",
                                lambda stop_ev, on_prog: dups.execute_duplicates_plan(
                                    hub.get_project(pid) or proj,
                                    mode=mode_v,
                                    confirm_delete=conf_del,
                                    stop_event=stop_ev,
                                    progress_cb=on_prog,
                                ),
                                blocking=blocking,
                                refresh_counters=True,
                            )
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "stop":
                            return self._send_json({"ok": hub.stop_scoped_job("dup", pid), "job": hub.get_scoped_job_status("dup", pid)})

                        if len(parts) == 5 and parts[4] == "job-dismiss":
                            hub.dismiss_scoped_job("dup", pid)
                            return self._send_json({"ok": True})

                        if len(parts) == 5 and parts[4] == "undo":
                            res = dups.undo_last_duplicates_quarantine(proj)
                            return self._send_json(res)

                        if len(parts) == 5 and parts[4] == "clear":
                            res = dups.clear_duplicates_plan(proj)
                            return self._send_json(res)

                    # ── Custom Sections Management ──────────────────────────
                    if len(parts) == 4 and action == "sections":
                        secs_in = body.get("sections") or []
                        if not isinstance(secs_in, list):
                            return self._send_json({"ok": False, "error": _tr("server.invalid_setting")}, 400)
                        updated_secs = hub.update_project_sections(pid, secs_in)
                        return self._send_json({
                            "ok": True,
                            "sections": [s for s in updated_secs if not s.get("hidden")],
                            "all_sections": updated_secs,
                        })

                    # ── Scanner Triage Endpoints (Scan -> Approve/Reject) ────
                    if action == "scanner":
                        if len(parts) == 5 and parts[4] == "scan":
                            blocking = bool(body.get("blocking", True))
                            res = hub.run_scanner_triage_scan(pid, blocking=blocking)
                            if not blocking:
                                return self._send_json(res)
                            st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                            return self._send_json({**st, "scan_summary": res, "job": hub.get_scanner_job_status(pid)})
                        if len(parts) == 5 and parts[4] == "job-dismiss":
                            hub.dismiss_scanner_job(pid)
                            return self._send_json({"ok": True})
                        if len(parts) == 5 and parts[4] == "approve":
                            if body.get("all") or isinstance(body.get("ids"), list):
                                res = resolve_triage_bulk(
                                    proj, action="approve",
                                    item_ids=body.get("ids"),
                                    device_id=hub.device_id,
                                    decision_store=hub.decisions,
                                )
                            else:
                                res = resolve_triage_item(
                                    proj, str(body.get("id") or ""),
                                    action="approve",
                                    overrides=body.get("overrides"),
                                    device_id=hub.device_id,
                                    decision_store=hub.decisions,
                                )
                            hub.refresh_project_counters(pid)
                            st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                            return self._send_json({**st, "resolution": res})
                        if len(parts) == 5 and parts[4] == "reject":
                            if body.get("all") or isinstance(body.get("ids"), list):
                                res = resolve_triage_bulk(
                                    proj, action="reject",
                                    item_ids=body.get("ids"),
                                    device_id=hub.device_id,
                                    decision_store=hub.decisions,
                                )
                            else:
                                res = resolve_triage_item(
                                    proj, str(body.get("id") or ""),
                                    action="reject",
                                    device_id=hub.device_id,
                                    decision_store=hub.decisions,
                                )
                            st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                            return self._send_json({**st, "resolution": res})
                        if len(parts) == 5 and parts[4] == "requeue":
                            res = resolve_triage_item(
                                proj, str(body.get("id") or ""),
                                action="requeue",
                                device_id=hub.device_id,
                                decision_store=hub.decisions,
                            )
                            st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                            return self._send_json({**st, "resolution": res})
                        if len(parts) == 5 and parts[4] == "clear-resolved":
                            res = clear_scanner_resolved_items(
                                proj,
                                bucket=str(body.get("bucket") or "all"),
                            )
                            st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                            return self._send_json({**st, "clear_result": res})
                        if len(parts) == 5 and parts[4] == "analyze-split":
                            res = analyze_multipage_pdf_split(
                                proj,
                                str(body.get("id") or ""),
                                global_settings=hub.get_settings(),
                                analyzer_override=hub.analyzer_override,
                                force_refresh=bool(body.get("force_refresh", False)),
                            )
                            return self._send_json(res)
                        if len(parts) == 5 and parts[4] == "apply-split":
                            res = apply_multipage_pdf_split(
                                proj,
                                str(body.get("id") or ""),
                                list(body.get("ranges") or []),
                                auto_approve=bool(body.get("auto_approve")),
                                device_id=hub.device_id,
                            )
                            hub.refresh_project_counters(pid)
                            st = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
                            return self._send_json({**st, "split_result": res})

                    # ── Template Studio Endpoints (Uploaded Templates ONLY!) ─
                    if action == "templates":
                        proj_sections = tprof.get_project_sections(proj, style=hub.get_settings().get("ui_style"), include_hidden=True)
                        sec_map = {s["id"]: s for s in proj_sections}
                        # POST /api/projects/<id>/templates/upload
                        if len(parts) == 5 and parts[4] == "upload":
                            raw_nm = str(body.get("name") or "")
                            raw_b = base64.b64decode(str(body.get("data_b64") or ""))
                            res = tprof.upload_project_template(
                                proj_root,
                                raw_nm,
                                raw_b,
                                section=str(body.get("section") or ""),
                                display_name=str(body.get("display_name") or ""),
                                device_id=hub.device_id,
                                project_sections=proj_sections,
                            )
                            return self._send_json(res, 201)

                        if len(parts) == 6:
                            tname = safe_template_name(parts[4])
                            sub_act = parts[5]

                            # POST /api/projects/<id>/templates/<name>/profile
                            if sub_act == "profile":
                                saved_p = tprof.save_template_profile(
                                    proj_root, tname, body,
                                    device_id=hub.device_id,
                                    project_sections=proj_sections,
                                )
                                return self._send_json({"ok": True, "profile": saved_p})

                            # POST /api/projects/<id>/templates/<name>/suggest-ai
                            if sub_act == "suggest-ai":
                                g = hub.get_settings()
                                res = suggest_fields_with_ai(
                                    proj_root, tname,
                                    engine=str(g.get("engine") or "gemini"),
                                    ollama_url=str(g.get("ollama_url") or ""),
                                    ollama_model=str(g.get("ollama_model") or ""),
                                    force_refresh=bool(body.get("force_refresh", False)),
                                )
                                return self._send_json(res)

                            # POST /api/projects/<id>/templates/<name>/preview
                            if sub_act == "preview":
                                prof = body.get("profile") if isinstance(body.get("profile"), dict) else tprof.get_template_profile(proj_root, tname, device_id=hub.device_id)
                                sec_id = str(prof.get("section") or "")
                                pfx = str(prof.get("ref_prefix") or (sec_map.get(sec_id, {}).get("ref_prefix") or "DOC"))
                                vals = body.get("values") if isinstance(body.get("values"), dict) else body
                                disc = str(body.get("discipline") or vals.get("discipline") or prof.get("default_discipline") or "")
                                next_ref, _, _ = plog.next_reference_from_logs(
                                    proj, prefix=pfx, discipline=disc,
                                    use_discipline_series=bool(prof.get("use_discipline_series", False)),
                                    device_id=hub.device_id,
                                )
                                sys_ctx = {
                                    "next_ref": str(vals.get("ref_no") or next_ref),
                                    "revision": str(vals.get("revision") or "R00"),
                                    "today_iso": str(vals.get("doc_date") or dt.date.today().isoformat()),
                                }
                                prev = trender.render_filled_preview(
                                    proj_root, tname, prof, body, proj,
                                    system_ctx=sys_ctx,
                                    sheet_name=str(body.get("sheet") or ""),
                                    try_libreoffice=bool(body.get("try_libreoffice", True)),
                                )
                                prev["next_ref"] = next_ref
                                return self._send_json(prev)

                            # POST /api/projects/<id>/templates/<name>/issue
                            if sub_act == "issue":
                                try:
                                    issued = issue_template_document(hub, proj, tname, body)
                                    return self._send_json(issued)
                                except RuntimeError as exc:
                                    return self._send_json({"ok": False, "error": str(exc)}, 400)

                            # POST /api/projects/<id>/templates/<name>/archive (zero deletion!)
                            if sub_act == "archive":
                                arc = tprof.archive_project_template(proj_root, tname, device_id=hub.device_id)
                                return self._send_json({"ok": True, "archived_to": arc.name})

                    # POST /api/projects/<id>/open (strict path-traversal guarded!)
                    if len(parts) == 4 and action == "open":
                        what = str(body.get("what") or "root")
                        if what == "root":
                            target = resolve_local_or_unc(proj_root)
                        elif what == "inbox":
                            target = resolve_local_or_unc(str(proj.get("inbox_dir") or (proj_root / "1_Inbox")))
                        elif what == "archive":
                            target = archive_dir_for_project(proj)
                        elif what == "logs":
                            target = safe_resolve_within(proj_root, "3_Project_Logs")
                        elif what == "templates":
                            target = safe_resolve_within(proj_root, "4_Project_Templates")
                        elif what == "backups":
                            target = safe_resolve_within(proj_root, "_Backups")
                        elif what == "reports":
                            target = safe_resolve_within(proj_root, "_Reports")
                            target.mkdir(parents=True, exist_ok=True)
                        elif what == "wa_source":
                            wa_pl = read_json_safe(waa.whatsapp_plan_path(proj_root), {}) or {}
                            if wa_pl.get("source_dir"):
                                target = resolve_local_or_unc(str(wa_pl["source_dir"]))
                            else:
                                target = waa.whatsapp_chats_root(proj_root)
                                target.mkdir(parents=True, exist_ok=True)
                        elif what == "dups_quarantine":
                            target = dups.dups_quarantine_root(proj_root)
                            target.mkdir(parents=True, exist_ok=True)
                        elif what == "dups_source":
                            d_pl = read_json_safe(dups.dups_plan_path(proj_root), {}) or {}
                            if d_pl.get("source_dir"):
                                target = resolve_local_or_unc(str(d_pl["source_dir"]))
                            else:
                                target = dups.dups_quarantine_root(proj_root)
                                target.mkdir(parents=True, exist_ok=True)
                        elif what == "log_file":
                            lname = safe_template_name(str(body.get("log_name") or ""))
                            target, _is_linked = plog.resolve_project_log_path(proj_root, lname)
                        elif what == "rel":
                            rel_s = str(body.get("rel") or "")
                            try:
                                target = safe_resolve_within(proj_root, rel_s, allow_base=False, reject_symlinks=True)
                            except PathTraversalError:
                                raise
                            if not target.exists():
                                alt = safe_resolve_within(archive_dir_for_project(proj), rel_s, allow_base=False, reject_symlinks=True)
                                if alt.exists():
                                    target = alt
                        else:
                            raise PathTraversalError(_tr("server.path_blocked"))
                        opened = open_on_desktop(target)
                        return self._send_json({"ok": True, "opened": opened, "path": str(target)})

                    # POST /api/projects/<id>/contacts | launch-agent
                    if len(parts) == 4 and action == "contacts":
                        if body.get("delete"):
                            cs = hub.delete_contact(pid, str(body.get("phone") or ""))
                        else:
                            cs = hub.save_contact(pid, str(body.get("name") or ""), str(body.get("phone") or ""), str(body.get("role") or ""))
                        return self._send_json({"ok": True, "contacts": cs})

                    if len(parts) == 4 and action == "launch-agent":
                        ok, msg = launch_gemini_in_project(proj, str(body.get("prompt") or ""))
                        return self._send_json({"ok": ok, "message": msg})

                    # POST /api/projects/<id>/whatsapp/send
                    if len(parts) == 5 and action == "whatsapp" and parts[4] == "send":
                        if not wab:
                            return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)
                        wa_cfg = wab.load_wa_config(hub.state_dir / "whatsapp_config.json")
                        mode = str(body.get("mode") or "web").strip().lower()
                        phone = str(body.get("phone") or "").strip()
                        message = str(body.get("message") or "").strip()
                        rel = str(body.get("rel") or "").strip()
                        area = str(body.get("area") or "root").strip()
                        target_f: Path | None = None
                        if rel:
                            base_dir = archive_dir_for_project(proj) if area == "archive" else proj_root
                            target_f = safe_resolve_within(base_dir, rel, allow_base=False, reject_symlinks=True)
                            if (not target_f.exists() or not target_f.is_file()) and area != "archive":
                                alt_f = safe_resolve_within(archive_dir_for_project(proj), rel, allow_base=False, reject_symlinks=True)
                                if alt_f.exists() and alt_f.is_file():
                                    target_f = alt_f
                            if not target_f.exists() or not target_f.is_file():
                                return self._send_json({"ok": False, "error": _tr("server.file_not_found")}, 404)
                        if mode == "cloud":
                            if target_f is not None:
                                res_wa = wab.cloud_send_file(
                                    target_f, phone, message, wa_cfg,
                                    use_template=bool(body.get("use_template", False)),
                                )
                            else:
                                res_wa = wab.cloud_send_text(phone, message, wa_cfg)
                            return self._send_json(res_wa)
                        clip_ok = False
                        clip_msg = ""
                        if target_f is not None and body.get("copy_to_clipboard", True):
                            clip_ok, clip_msg = wab.copy_file_to_os_clipboard(target_f)
                        wa_url = wab.build_whatsapp_web_url(phone, message, str(wa_cfg.get("default_country_code") or "965"))
                        return self._send_json({
                            "ok": True,
                            "mode": "web",
                            "url": wa_url,
                            "clipboard_copied": clip_ok,
                            "clipboard_message": clip_msg,
                        })

                return self._send_json({"ok": False, "error": _tr("server.not_found")}, 404)

            except PathTraversalError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 403)
            except UnsafeFileError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 400)
            except tprof.ProfileOrphanError as exc:
                return self._send_json({
                    "ok": False,
                    "error": str(exc),
                    "code": "ORPHAN_FIELDS",
                    "orphan_fields": exc.orphan_fields,
                }, 409)
            except tprof.FormulaOverwriteError as exc:
                return self._send_json({
                    "ok": False,
                    "error": str(exc),
                    "code": "FORMULA_PROTECTED",
                }, 400)
            except tprof.ProfileValidationError as exc:
                return self._send_json({
                    "ok": False,
                    "error": str(exc),
                    "code": "VALIDATION_ERROR",
                    "errors": exc.errors,
                }, 422)
            except OverflowError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 413)
            except PermissionError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 403)
            except KeyError as exc:
                msg = str(exc.args[0]) if exc.args else str(exc)
                return self._send_json({"ok": False, "error": msg}, 404)
            except FileNotFoundError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 404)
            except ValueError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 400)
            except Exception as exc:
                print(f"[Dashboard POST Error] {type(exc).__name__}: {exc}", file=sys.stderr)
                return self._send_json({"ok": False, "error": _tr("common.error_generic")}, 500)

        # ── DELETE ───────────────────────────────────────────────────────────
        def do_DELETE(self) -> None:
            if not self._check_origin():
                return
            parsed = urlparse(self.path)
            try:
                parts = self._split_api_parts(parsed.path)
                if len(parts) == 3 and parts[:2] == ["api", "learned-rules"]:
                    ok = hub.decisions.delete_rule(parts[2])
                    return self._send_json({"ok": ok}, 200 if ok else 404)
                return self._send_json({"ok": False, "error": _tr("server.delete_disabled")}, 405)
            except PathTraversalError as exc:
                return self._send_json({"ok": False, "error": str(exc)}, 403)

    return DashboardHandler


def create_server(
    hub: HubState,
    host: str = "127.0.0.1",
    port: int = 8765,
    *,
    allow_remote_preview: bool = False,
) -> ThreadingHTTPServer:
    safe_host = validate_bind_host(host, allow_remote_preview=allow_remote_preview)
    handler = make_handler(hub, allow_remote_preview=allow_remote_preview)
    return ThreadingHTTPServer((safe_host, int(port)), handler)


def main() -> None:
    parser = argparse.ArgumentParser(description="Smart Local Agent Free")
    parser.add_argument("--host", default="127.0.0.1", help="Local bind host (default 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8765, help="Port number (default 8765)")
    parser.add_argument("--allow-remote-preview", action="store_true", help="Allow isolated preview bind")
    parser.add_argument("--no-browser", action="store_true", help="Do not open browser automatically")
    parser.add_argument("--projects-root", default="", help="Projects root folder")
    parser.add_argument("--registry", default="", help="Registry file path")
    args = parser.parse_args()

    pkg_root = SCRIPT_DIR.parent
    if args.projects_root:
        default_projects_root = Path(args.projects_root).resolve()
    elif pkg_root.name == ".gemini":
        desk = Path.home() / "Desktop"
        default_projects_root = (desk if desk.exists() else Path.home()) / "Smart_Projects_Hub"
    else:
        default_projects_root = (pkg_root.parent / "Smart_Projects_مشاريعي").resolve()

    if args.registry:
        default_registry = Path(args.registry).resolve()
    elif pkg_root.name == ".gemini":
        default_registry = pkg_root / "projects_hub" / "registry.json"
    else:
        gemini_reg = Path.home() / ".gemini" / "projects_hub" / "registry.json"
        local_reg = pkg_root / "_hub_state" / "registry.json"
        legacy_reg = default_projects_root / "_hub_registry.json"
        if gemini_reg.exists() or not (local_reg.exists() or legacy_reg.exists()):
            default_registry = gemini_reg
        elif local_reg.exists():
            default_registry = local_reg
        else:
            default_registry = legacy_reg

    hub = HubState(default_registry, default_projects_root)
    server = create_server(hub, host=args.host, port=args.port, allow_remote_preview=args.allow_remote_preview)
    actual_port = server.server_address[1]
    url = f"http://127.0.0.1:{actual_port}"
    print(f"{_tr('server.app_ready_console')} — {url}", flush=True)

    if not args.no_browser and not args.allow_remote_preview:
        threading.Timer(0.7, lambda: open_standalone_app_window(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print(f"\n{_tr('server.app_stopped_console')}")
    finally:
        hub.toggle_watcher(False)
        server.server_close()


if __name__ == "__main__":
    main()
