#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
active_project.py — Active-project context, `active_only` vs `all_projects` device mode,
shared-inbox detection, Start-Work / Sync routine, End-Session clean shutdown, dated
`_Backups/` snapshots, external backup mirror, and >7-day stale-backup warnings.
"""

from __future__ import annotations

import datetime as dt
import os
import re
import shutil
from pathlib import Path
from typing import Any

import project_log as plog
from i18n_ar import tr
from relevance import is_file_locked
from security_utils import atomic_write_json, now_iso, scrub_secrets, sha256_of_file

PROCESS_MODES: tuple[str, ...] = ("active_only", "all_projects")
DEFAULT_BACKUP_WARN_DAYS: int = 7
_GD_AR_DRIVE_FOLDER = "".join(chr(c) for c in (0x0645, 0x0644, 0x0641, 0x0627, 0x062A, 0x064A))


def effective_process_mode(settings: dict[str, Any], projects_count: int) -> str:
    """
    Determine effective device processing mode:
    - If user explicitly chose a mode (`process_mode_explicit=True`), honor it.
    - Otherwise default automatically to `active_only` whenever >1 project exists on the PC!
    """
    explicit = bool(settings.get("process_mode_explicit", False))
    saved = str(settings.get("process_mode") or "").strip().lower()
    if explicit and saved in PROCESS_MODES:
        return saved
    if projects_count > 1:
        return "active_only"
    if saved in PROCESS_MODES:
        return saved
    return "active_only"


def get_projects_to_process(
    projects: list[dict[str, Any]],
    active_project_id: str | None,
    process_mode: str,
) -> list[dict[str, Any]]:
    """
    Filter which projects background watchers / bulk cycles are allowed to touch.
    In `active_only` mode, ONLY the currently active project is ever processed!
    """
    enabled = [p for p in projects if p.get("enabled", True)]
    if process_mode == "all_projects":
        return enabled
    if not active_project_id:
        return []
    return [p for p in enabled if str(p.get("id")) == str(active_project_id)]


def _norm_path_key(raw_path: str) -> str:
    if not raw_path:
        return ""
    s = str(raw_path).strip()
    if not s:
        return ""
    try:
        resolved = str(Path(s).resolve())
    except Exception:
        resolved = s
    resolved = resolved.replace("\\", "/").rstrip("/")
    # Case-fold Windows paths too
    if os.name == "nt" or re.match(r"^[A-Za-z]:/", resolved):
        resolved = resolved.lower()
    return resolved


def detect_shared_inboxes(projects: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Detect any pair of registered projects that share the exact same inbox folder
    or have nested (parent/child) inbox folders.
    """
    warnings: list[dict[str, Any]] = []
    entries: list[tuple[dict[str, Any], str, str]] = []
    for p in projects:
        inbox = str(p.get("inbox_dir") or "").strip()
        key = _norm_path_key(inbox)
        if key:
            entries.append((p, inbox, key))

    for i in range(len(entries)):
        p1, raw1, k1 = entries[i]
        for j in range(i + 1, len(entries)):
            p2, _raw2, k2 = entries[j]
            relation = ""
            if k1 == k2:
                relation = "same_folder"
            elif k2.startswith(k1 + "/") or k1.startswith(k2 + "/"):
                relation = "nested_folder"
            if relation:
                msg = (
                    tr("server.shared_scanner_warn", a=p1.get("name"), b=p2.get("name"))
                    if relation == "same_folder"
                    else tr("server.nested_scanner_warn", a=p1.get("name"), b=p2.get("name"))
                )
                warnings.append({
                    "kind": relation,
                    "project_ids": [p1.get("id"), p2.get("id")],
                    "project_names": [p1.get("name"), p2.get("name")],
                    "inbox_path": raw1,
                    "message": msg,
                })
    return warnings


def is_project_inbox_shared(projects: list[dict[str, Any]], project_id: str) -> bool:
    for w in detect_shared_inboxes(projects):
        if project_id in (w.get("project_ids") or []):
            return True
    return False


def scan_open_or_locked_files(proj_root: Path) -> list[dict[str, Any]]:
    """Inspect `3_Project_Logs` and `4_Project_Templates` for active Office locks / locked files."""
    locked: list[dict[str, Any]] = []
    for sub in ("3_Project_Logs", "4_Project_Templates"):
        d = proj_root / sub
        if not d.exists():
            continue
        for f in sorted(d.iterdir()):
            if not f.is_file():
                continue
            if f.name.startswith("~$"):
                real_name = f.name[2:] or f.name
                locked.append({
                    "folder": sub,
                    "file": real_name,
                    "lock_file": f.name,
                    "reason": tr("server.file_in_use", name=real_name),
                })
            elif f.suffix.lower() in (".xlsx", ".xlsm", ".docx") and not f.name.startswith((".", "_")):
                if is_file_locked(f):
                    locked.append({
                        "folder": sub,
                        "file": f.name,
                        "reason": tr("server.file_in_use", name=f.name),
                    })
    # Deduplicate by folder+file
    seen: set[tuple[str, str]] = set()
    unique: list[dict[str, Any]] = []
    for item in locked:
        key = (item["folder"], item["file"])
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def _parse_iso_dt(val: Any) -> dt.datetime | None:
    if not val or not isinstance(val, str):
        return None
    try:
        return dt.datetime.fromisoformat(val)
    except ValueError:
        return None


def check_backup_status(
    proj: dict[str, Any],
    global_settings: dict[str, Any] | None = None,
    *,
    warn_days: int = DEFAULT_BACKUP_WARN_DAYS,
    now_override: dt.datetime | None = None,
) -> dict[str, Any]:
    """
    Evaluate local & external backup freshness and emit a warning when > `warn_days` (default 7)
    have elapsed since the last backup.
    """
    gcfg = global_settings or {}
    now_dt = now_override or dt.datetime.now()
    proj_root = Path(str(proj.get("root") or ""))
    backups_dir = proj_root / "_Backups"

    last_local = _parse_iso_dt(proj.get("last_backup_at"))
    # Also inspect session folders on disk if metadata was cleared
    session_dirs: list[Path] = []
    if backups_dir.exists():
        session_dirs = sorted(
            [d for d in backups_dir.iterdir() if d.is_dir() and d.name.startswith("Session_")],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if session_dirs and last_local is None:
            last_local = dt.datetime.fromtimestamp(session_dirs[0].stat().st_mtime)

    ext_dir = str(proj.get("external_backup_dir") or gcfg.get("external_backup_dir") or "").strip()
    last_ext = _parse_iso_dt(proj.get("last_external_backup_at"))
    created_dt = _parse_iso_dt(proj.get("created_at"))

    days_since_local: int | None = None
    if last_local is not None:
        days_since_local = max(0, int((now_dt - last_local).total_seconds() // 86400))

    days_since_ext: int | None = None
    if last_ext is not None:
        days_since_ext = max(0, int((now_dt - last_ext).total_seconds() // 86400))

    stale = False
    warnings: list[str] = []
    if last_local is None:
        if created_dt is not None and (now_dt - created_dt).total_seconds() > warn_days * 86400:
            stale = True
            warnings.append(
                tr("server.backup_overdue_warn", days=warn_days, project=proj.get("name"))
            )
    elif days_since_local is not None and days_since_local > warn_days:
        stale = True
        warnings.append(
            tr("server.backup_overdue_warn", days=days_since_local, project=proj.get("name"))
        )

    if ext_dir:
        if last_ext is None:
            if created_dt is not None and (now_dt - created_dt).total_seconds() > warn_days * 86400:
                stale = True
                warnings.append(tr("server.ext_backup_missing"))
        elif days_since_ext is not None and days_since_ext > warn_days:
            stale = True
            warnings.append(
                tr("server.backup_overdue_warn", days=days_since_ext, project=proj.get("name"))
            )

    ext_low = ext_dir.lower()
    if not ext_dir:
        ext_provider = ""
    elif any(k in ext_low for k in ("google", "my drive", "googledrive", _GD_AR_DRIVE_FOLDER)) or ext_low.startswith("g:\\"):
        ext_provider = "google_drive"
    elif "onedrive" in ext_low:
        ext_provider = "onedrive"
    elif "dropbox" in ext_low:
        ext_provider = "dropbox"
    else:
        ext_provider = "folder"

    bkp_status = "overdue" if stale else ("never" if last_local is None else "ok")

    return {
        "stale": stale,
        "status": bkp_status,
        "warn_days": warn_days,
        "last_backup_at": last_local.isoformat(timespec="seconds") if last_local else None,
        "days_since_backup": days_since_local,
        "sessions_count": len(session_dirs),
        "latest_session_folder": str(session_dirs[0]) if session_dirs else None,
        "external_backup_dir": ext_dir,
        "external_provider": ext_provider,
        "external_reachable": bool(ext_dir and Path(ext_dir).exists()),
        "last_external_backup_at": last_ext.isoformat(timespec="seconds") if last_ext else None,
        "days_since_external": days_since_ext,
        "warnings": warnings,
    }


def detect_cloud_backup_folders(extra_candidates: list[Path] | None = None) -> list[dict[str, str]]:
    """
    Detect local synced cloud folders (Google Drive Desktop, OneDrive, Dropbox)
    so users can link auto-backup in 1 click without hunting through drives.
    Never creates directories during detection; suggests a clean `SmartProjects_Backups` subfolder.
    """
    home = Path.home()
    candidates: list[tuple[str, str, Path]] = []

    # Windows Google Drive Desktop streamed mount letters (commonly G:\ or H:\ ..)
    if os.name == "nt":
        for drv in ("G", "H", "I", "D", "E"):
            for sub in ("My Drive", _GD_AR_DRIVE_FOLDER):
                candidates.append(("google_drive", "Google Drive", Path(f"{drv}:/{sub}")))

    # Cross-platform user-home cloud folders
    candidates.extend([
        ("google_drive", "Google Drive", home / "Google Drive" / "My Drive"),
        ("google_drive", "Google Drive", home / "Google Drive"),
        ("google_drive", "Google Drive", home / "My Drive"),
        ("onedrive", "OneDrive", home / "OneDrive"),
        ("dropbox", "Dropbox", home / "Dropbox"),
    ])

    cloud_storage = home / "Library" / "CloudStorage"
    if cloud_storage.exists() and cloud_storage.is_dir():
        try:
            for d in sorted(cloud_storage.iterdir()):
                if not d.is_dir():
                    continue
                if d.name.startswith("GoogleDrive"):
                    sub = d / "My Drive"
                    candidates.append(("google_drive", "Google Drive", sub if sub.exists() else d))
                elif d.name.startswith("OneDrive"):
                    candidates.append(("onedrive", "OneDrive", d))
        except OSError:
            pass

    if extra_candidates:
        for ec in extra_candidates:
            candidates.append(("google_drive", "Google Drive", Path(ec)))

    found: list[dict[str, str]] = []
    seen_paths: set[str] = set()
    for prov, label, base_p in candidates:
        try:
            if not base_p.exists() or not base_p.is_dir():
                continue
            target = base_p / "SmartProjects_Backups"
            norm_k = str(target).lower()
            if norm_k in seen_paths:
                continue
            seen_paths.add(norm_k)
            found.append({
                "provider": prov,
                "label": f"{label} ({base_p})",
                "path": str(target),
            })
        except OSError:
            continue
    return found


def create_session_backup(
    proj: dict[str, Any],
    *,
    device_id: str,
    global_settings: dict[str, Any] | None = None,
    copy_external: bool = True,
    include_archive_external: bool = False,
    reason: str = "end_session",
) -> dict[str, Any]:
    """
    Create a clean, timestamped snapshot in `<project_root>/_Backups/Session_<ts>_<device>/`
    and optionally mirror it to the configured external backup path (USB/Drive/NAS).
    Zero deletion — previous snapshots stay untouched.
    """
    gcfg = global_settings or {}
    proj_root = Path(str(proj.get("root") or "")).resolve()
    backups_root = proj_root / "_Backups"
    backups_root.mkdir(parents=True, exist_ok=True)

    ts = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    short_dev = re.sub(r"[^A-Za-z0-9\-]", "", device_id or "LOCAL")[-10:]
    snap_name = f"Session_{ts}_{short_dev}"
    snap_dir = backups_root / snap_name
    idx = 2
    while snap_dir.exists():
        snap_dir = backups_root / f"{snap_name}_{idx}"
        idx += 1
    snap_dir.mkdir(parents=True, exist_ok=True)

    copied_files: list[dict[str, Any]] = []

    # 1. Copy all uploaded project logs + `.logmap.json` + pending queue + audit log
    logs_src = proj_root / "3_Project_Logs"
    if logs_src.exists():
        logs_dst = snap_dir / "3_Project_Logs"
        logs_dst.mkdir(parents=True, exist_ok=True)
        for f in sorted(logs_src.iterdir()):
            if f.is_file() and not f.name.startswith("~$"):
                dst_f = logs_dst / f.name
                shutil.copy2(f, dst_f)
                copied_files.append({
                    "rel": f"3_Project_Logs/{f.name}",
                    "size_bytes": dst_f.stat().st_size,
                    "sha256": sha256_of_file(dst_f),
                })

    # 2. Copy all uploaded templates + their `.tplmap.json` profiles
    tpl_src = proj_root / "4_Project_Templates"
    if tpl_src.exists():
        tpl_dst = snap_dir / "4_Project_Templates"
        tpl_dst.mkdir(parents=True, exist_ok=True)
        for f in sorted(tpl_src.iterdir()):
            if f.is_file() and not f.name.startswith("~$"):
                dst_f = tpl_dst / f.name
                shutil.copy2(f, dst_f)
                copied_files.append({
                    "rel": f"4_Project_Templates/{f.name}",
                    "size_bytes": dst_f.stat().st_size,
                    "sha256": sha256_of_file(dst_f),
                })

    # 3. Write scrubbed manifest (ZERO secrets!)
    safe_proj_meta = scrub_secrets({
        "id": proj.get("id"),
        "name": proj.get("name"),
        "code": proj.get("code"),
        "contract_no": proj.get("contract_no"),
        "owner": proj.get("owner"),
        "consultant": proj.get("consultant"),
        "contractor": proj.get("contractor"),
        "synced_counters": proj.get("synced_counters", {}),
        "baseline_counters": proj.get("baseline_counters", {}),
    })
    manifest = {
        "snapshot_name": snap_dir.name,
        "reason": reason,
        "created_at": now_iso(),
        "device_id": device_id,
        "project": safe_proj_meta,
        "files_count": len(copied_files),
        "files": copied_files,
    }
    atomic_write_json(snap_dir / "session_manifest.json", manifest)

    # 4. Optional External Backup copy (USB / OneDrive / Google Drive / NAS)
    ext_dir_raw = str(proj.get("external_backup_dir") or gcfg.get("external_backup_dir") or "").strip()
    ext_result: dict[str, Any] = {"configured": bool(ext_dir_raw), "copied": False, "path": None, "error": None}
    if copy_external and ext_dir_raw:
        try:
            ext_base = Path(ext_dir_raw).resolve()
            ext_base.mkdir(parents=True, exist_ok=True)
            safe_proj_folder = re.sub(
                r'[<>:"/\\|?*\x00-\x1f]+',
                "_",
                f"{proj.get('code') or 'PRJ'}_{proj.get('name') or proj.get('id')}",
            ).strip(" ._")[:70]
            ext_proj_dir = ext_base / safe_proj_folder
            ext_snap_dir = ext_proj_dir / "_Session_Backups" / snap_dir.name
            shutil.copytree(snap_dir, ext_snap_dir, dirs_exist_ok=True)

            # Optionally mirror new/modified archive files incrementally (zero deletion!)
            arc_copied = 0
            arc_raw = str(proj.get("archive_dir") or "").strip()
            arc_src = Path(arc_raw).resolve() if arc_raw else (proj_root / "2_Archive")
            if include_archive_external and arc_src.exists():
                arc_dst = ext_proj_dir / "2_Archive"
                for src_f in arc_src.rglob("*"):
                    if not src_f.is_file() or src_f.name.startswith("~$"):
                        continue
                    rel = src_f.relative_to(arc_src)
                    dst_f = arc_dst / rel
                    if not dst_f.exists() or dst_f.stat().st_size != src_f.stat().st_size:
                        dst_f.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(src_f, dst_f)
                        arc_copied += 1

            ext_result = {
                "configured": True,
                "copied": True,
                "path": str(ext_snap_dir),
                "archive_files_synced": arc_copied,
                "error": None,
            }
        except Exception:
            ext_result = {
                "configured": True,
                "copied": False,
                "path": ext_dir_raw,
                "error": tr("server.ext_backup_failed"),
            }

    return {
        "ok": True,
        "snapshot_name": snap_dir.name,
        "snapshot_path": str(snap_dir),
        "created_at": manifest["created_at"],
        "files_count": len(copied_files),
        "device_id": device_id,
        "external": ext_result,
    }


def run_start_sync(
    proj: dict[str, Any],
    all_projects: list[dict[str, Any]],
    global_settings: dict[str, Any],
    *,
    device_id: str,
    pending_decisions_count: int = 0,
) -> dict[str, Any]:
    """
    Execute the full «بدء العمل / مزامنة» (Start Work / Sync) routine for `proj`:
      1) Verify folder structure, inbox, uploaded project log(s), template profiles, locks
      2) Flush all pending log queue operations into the uploaded project log(s)
      3) Recalculate ALL counters directly from the uploaded project log(s) (`max(log)+1`)
      4) Check shared-inbox conflicts & >7-day stale backup alerts
      5) Return complete readiness & counter summary
    """
    proj_root = Path(str(proj.get("root") or "")).resolve()
    checks: list[dict[str, Any]] = []
    warnings: list[str] = []

    # 1. Core folder checks (creates missing scaffolding safely, NEVER creates generic log or system templates!)
    required_dirs = ("1_Inbox", "2_Archive", "3_Project_Logs", "4_Project_Templates", "_Backups")
    proj_root.mkdir(parents=True, exist_ok=True)
    for dn in required_dirs:
        (proj_root / dn).mkdir(parents=True, exist_ok=True)
    arc_raw = str(proj.get("archive_dir") or "").strip()
    if arc_raw:
        Path(arc_raw).resolve().mkdir(parents=True, exist_ok=True)
    checks.append({"item": "folders", "ok": True, "label": tr("sync.folders_ok")})

    inbox_raw = str(proj.get("inbox_dir") or "").strip()
    inbox_ok = bool(inbox_raw and Path(inbox_raw).exists())
    if inbox_ok:
        checks.append({"item": "inbox", "ok": True, "label": tr("sync.scanner_ok")})
    else:
        checks.append({"item": "inbox", "ok": False, "label": tr("sync.scanner_missing")})
        warnings.append(tr("sync.scanner_missing"))

    # 2. Check locks & flush pending log queue
    locked_files = scan_open_or_locked_files(proj_root)
    if locked_files:
        names = ", ".join(x["file"] for x in locked_files)
        warnings.append(tr("server.open_files_warn", files=names))

    flush_res = plog.flush_project_log_queue(proj, device_id=device_id)
    if flush_res["flushed"] > 0:
        checks.append({
            "item": "queue_flush",
            "ok": True,
            "label": tr("sync.flushed_msg", count=flush_res["flushed"]),
        })
    if flush_res["remaining"] > 0:
        warnings.append(tr("sync.waiting_count", count=flush_res["remaining"]))

    # 3. Inspect uploaded project log(s)
    uploaded_logs = plog.list_project_logs(proj_root, device_id=device_id)
    if uploaded_logs:
        log_names = ", ".join(lg["name"] for lg in uploaded_logs)
        checks.append({
            "item": "project_log",
            "ok": True,
            "label": tr("sync.log_ready", names=log_names),
        })
        for lg in uploaded_logs:
            for orph in (lg.get("orphan_columns") or []):
                if orph.get("reason"):
                    warnings.append(str(orph.get("reason")))
    else:
        checks.append({
            "item": "project_log",
            "ok": False,
            "label": tr("overview.no_log_warn"),
        })
        warnings.append(tr("overview.no_log_warn"))

    # 4. Recalculate ALL counters directly from uploaded project log(s)
    counters_sync = plog.recalculate_counters_from_logs(proj, device_id=device_id)
    checks.append({
        "item": "counters",
        "ok": True,
        "label": tr("sync.counters_ready", count=counters_sync["rows_scanned"]),
    })

    # 5. Check uploaded templates & orphan profile fields
    tpl_summary = {"uploaded_count": 0, "mapped_count": 0, "orphan_templates": []}
    try:
        import template_profile as tprof
        tpls = tprof.list_project_templates(proj_root, device_id=device_id)
        tpl_summary["uploaded_count"] = len(tpls)
        tpl_summary["mapped_count"] = sum(1 for t in tpls if t.get("has_profile") and t.get("section_confirmed"))
        orph_tpls = [t["name"] for t in tpls if t.get("orphan_fields")]
        tpl_summary["orphan_templates"] = orph_tpls
        if orph_tpls:
            warnings.append(tr("sync.moved_templates", names=", ".join(orph_tpls)))
    except ImportError:
        pass

    # 6. Shared-inbox warnings & backup staleness
    shared_warns = [
        w["message"] for w in detect_shared_inboxes(all_projects)
        if proj.get("id") in (w.get("project_ids") or [])
    ]
    warnings.extend(shared_warns)

    backup_info = check_backup_status(
        proj,
        global_settings,
        warn_days=int(global_settings.get("backup_warn_days") or DEFAULT_BACKUP_WARN_DAYS),
    )
    warnings.extend(backup_info.get("warnings") or [])

    return {
        "ok": True,
        "project_id": proj.get("id"),
        "project_name": proj.get("name"),
        "project_code": proj.get("code"),
        "device_id": device_id,
        "started_at": now_iso(),
        "process_mode": effective_process_mode(global_settings, len(all_projects)),
        "checks": checks,
        "warnings": warnings,
        "locked_files": locked_files,
        "queue_flush": flush_res,
        "uploaded_logs": [
            {
                "name": lg["name"],
                "sheets_enabled": lg.get("sheets_enabled", 0),
                "has_vba": lg.get("has_vba", False),
                "sections_covered": lg.get("sections_covered", []),
            }
            for lg in uploaded_logs
        ],
        "templates": tpl_summary,
        "counters": {
            "rows_scanned": counters_sync["rows_scanned"],
            "next_preview": counters_sync["next_preview"],
            "next_disc_preview": counters_sync["next_disc_preview"],
            "next_global_disc_preview": counters_sync.get("next_global_disc_preview", {}),
            "is_global_disc_series": counters_sync.get("is_global_disc_series", False),
            "next_arabic_preview": counters_sync["next_arabic_preview"],
            "effective_counters": counters_sync["effective_counters"],
            "effective_disc": counters_sync["effective_disc"],
            "arabic_series": counters_sync["arabic_series"],
            "global_series": counters_sync.get("global_series", {}),
        },
        "pending_decisions_count": pending_decisions_count,
        "backup": backup_info,
    }


def run_end_session(
    proj: dict[str, Any],
    global_settings: dict[str, Any],
    *,
    device_id: str,
    copy_external: bool = True,
    include_archive_external: bool = False,
) -> dict[str, Any]:
    """
    Execute the full «إنهاء الجلسة» (End Session) routine for `proj`:
      1) Flush pending log queue
      2) Verify no open/locked files remain
      3) Save timestamped snapshot to `<project_root>/_Backups/` (+ external path if configured)
      4) Return complete closure report (caller clears `active_project_id`)
    """
    proj_root = Path(str(proj.get("root") or "")).resolve()
    warnings: list[str] = []

    flush_res = plog.flush_project_log_queue(proj, device_id=device_id)
    locked_files = scan_open_or_locked_files(proj_root)
    if locked_files:
        warnings.append(tr("end_session.open_files_warn"))
    if flush_res["remaining"] > 0:
        warnings.append(tr("sync.waiting_count", count=flush_res["remaining"]))

    backup_res = create_session_backup(
        proj,
        device_id=device_id,
        global_settings=global_settings,
        copy_external=copy_external,
        include_archive_external=include_archive_external,
        reason="end_session",
    )
    if backup_res["external"].get("error"):
        warnings.append(backup_res["external"]["error"])

    return {
        "ok": True,
        "project_id": proj.get("id"),
        "project_name": proj.get("name"),
        "ended_at": now_iso(),
        "device_id": device_id,
        "active_cleared": True,
        "queue_flush": flush_res,
        "locked_files": locked_files,
        "all_writes_closed": len(locked_files) == 0 and flush_res["remaining"] == 0,
        "backup": backup_res,
        "warnings": warnings,
    }
