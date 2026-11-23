#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
task_bar.py — Smart Taskbar Backend for Smart_Local_Agent_Free (v10.13.0)

Combines:
  1) Local personal & quick-capture tasks (`_hub_state/tasks.json`) — zero writes to project folders/logs.
  2) Auto-extracted actionable items across enabled projects:
     - Overdue submittals waiting beyond `overdue_days`
     - Unreplied incoming correspondence
     - Rejected submittals (`Code C / D`) needing resubmission
     - Pending scanner items / decisions
     - Overdue project backups
  3) Today-only snooze (`dismissed_auto`) so resolved/deferred auto-reminders stay out of the way cleanly.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path
import re
from typing import Any
import uuid

from active_project import DEFAULT_BACKUP_WARN_DAYS, check_backup_status
from inbox_pipeline import get_scanner_triage_state
import project_log as plog
from security_utils import atomic_write_json, now_iso, read_json_safe, scrub_secrets

VALID_PRIORITIES = ("urgent", "normal", "low")


def tasks_file_path(state_dir: Path) -> Path:
    return Path(state_dir).resolve() / "tasks.json"


def _load_tasks_store(state_dir: Path) -> dict[str, Any]:
    fp = tasks_file_path(state_dir)
    raw = read_json_safe(fp, {})
    if not isinstance(raw, dict):
        raw = {}
    items = [t for t in (raw.get("tasks") or []) if isinstance(t, dict) and t.get("id") and t.get("title")]
    dismissed = raw.get("dismissed_auto") if isinstance(raw.get("dismissed_auto"), dict) else {}
    raw_meta = raw.get("dismissed_auto_meta") if isinstance(raw.get("dismissed_auto_meta"), dict) else {}
    # Prune expired day-snoozes (keep today's or future dates)
    today_iso = dt.date.today().isoformat()
    clean_dismissed = {
        str(k): str(v)
        for k, v in dismissed.items()
        if str(k).strip() and str(v or "")[:10] >= today_iso
    }
    clean_meta = {
        str(k): dict(v)
        for k, v in raw_meta.items()
        if str(k) in clean_dismissed and isinstance(v, dict)
    }
    return {
        "tasks": items,
        "dismissed_auto": clean_dismissed,
        "dismissed_auto_meta": clean_meta,
        "updated_at": raw.get("updated_at") or "",
    }


def _save_tasks_store(state_dir: Path, store: dict[str, Any]) -> None:
    store["updated_at"] = now_iso()
    atomic_write_json(tasks_file_path(state_dir), scrub_secrets(store))


def _parse_quick_hints(
    title: str,
    *,
    priority: str = "",
    due_date: str = "",
) -> tuple[str, str, str]:
    """Extract optional `#عاجل` / `#غدا` / `#اليوم` quick tokens cleanly without breaking plain titles."""
    t_clean = str(title or "").strip()
    prio = str(priority or "").strip().lower()
    due = str(due_date or "").strip()[:10]
    today = dt.date.today()

    if prio not in VALID_PRIORITIES:
        if re.search(r"(?:^|\s)[#!](?:عاجل|urgent|p1)\b", t_clean, re.I):
            prio = "urgent"
            t_clean = re.sub(r"(?:^|\s)[#!](?:عاجل|urgent|p1)\b", " ", t_clean, flags=re.I)
        elif re.search(r"(?:^|\s)[#!](?:منخفض|low|p3)\b", t_clean, re.I):
            prio = "low"
            t_clean = re.sub(r"(?:^|\s)[#!](?:منخفض|low|p3)\b", " ", t_clean, flags=re.I)
        else:
            prio = "normal"

    if not due:
        if re.search(r"(?:^|\s)#(?:اليوم|today)\b", t_clean, re.I):
            due = today.isoformat()
            t_clean = re.sub(r"(?:^|\s)#(?:اليوم|today)\b", " ", t_clean, flags=re.I)
        elif re.search(r"(?:^|\s)#(?:غدا|غداً|بكرة|tomorrow)\b", t_clean, re.I):
            due = (today + dt.timedelta(days=1)).isoformat()
            t_clean = re.sub(r"(?:^|\s)#(?:غدا|غداً|بكرة|tomorrow)\b", " ", t_clean, flags=re.I)

    t_clean = re.sub(r"\s+", " ", t_clean).strip() or str(title or "").strip()
    if due and not re.match(r"^\d{4}-\d{2}-\d{2}$", due):
        due = ""
    return t_clean, prio, due


def list_manual_tasks(
    state_dir: Path,
    *,
    project_id: str = "",
    include_done: bool = True,
) -> list[dict[str, Any]]:
    store = _load_tasks_store(state_dir)
    out: list[dict[str, Any]] = []
    for t in store["tasks"]:
        if project_id and str(t.get("project_id") or "") != project_id:
            continue
        if not include_done and t.get("done"):
            continue
        out.append(dict(t))
    prio_rank = {"urgent": 0, "normal": 1, "low": 2}
    out.sort(
        key=lambda x: (
            1 if x.get("done") else 0,
            prio_rank.get(str(x.get("priority") or "normal"), 1),
            str(x.get("due_date") or "9999-99-99"),
            str(x.get("created_at") or ""),
        )
    )
    return out


def create_manual_task(
    state_dir: Path,
    payload: dict[str, Any],
    *,
    project_lookup: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    raw_title = str(payload.get("title") or "").strip()
    if not raw_title:
        raise ValueError("Task title is required")
    title, prio, due = _parse_quick_hints(
        raw_title,
        priority=str(payload.get("priority") or ""),
        due_date=str(payload.get("due_date") or ""),
    )
    pid = str(payload.get("project_id") or "").strip()
    p_meta = (project_lookup or {}).get(pid) or {}
    now_s = now_iso()
    task = scrub_secrets({
        "id": f"tsk-{uuid.uuid4().hex[:8]}",
        "title": title[:240],
        "project_id": pid,
        "project_code": str(p_meta.get("code") or payload.get("project_code") or ""),
        "project_name": str(p_meta.get("name") or payload.get("project_name") or ""),
        "priority": prio,
        "due_date": due,
        "ref_no": str(payload.get("ref_no") or "")[:60],
        "notes": str(payload.get("notes") or "")[:600],
        "done": False,
        "done_at": "",
        "created_at": now_s,
        "updated_at": now_s,
    })
    store = _load_tasks_store(state_dir)
    store["tasks"].insert(0, task)
    _save_tasks_store(state_dir, store)
    return task


def update_manual_task(
    state_dir: Path,
    task_id: str,
    patch: dict[str, Any],
    *,
    project_lookup: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    store = _load_tasks_store(state_dir)
    idx = next((i for i, t in enumerate(store["tasks"]) if t.get("id") == task_id), -1)
    if idx < 0:
        raise KeyError("Task not found")
    cur = dict(store["tasks"][idx])
    if "title" in patch:
        nt = str(patch.get("title") or "").strip()
        if not nt:
            raise ValueError("Task title cannot be empty")
        cur["title"] = nt[:240]
    if "priority" in patch:
        pv = str(patch.get("priority") or "normal").strip().lower()
        cur["priority"] = pv if pv in VALID_PRIORITIES else "normal"
    if "due_date" in patch:
        dv = str(patch.get("due_date") or "").strip()[:10]
        cur["due_date"] = dv if re.match(r"^\d{4}-\d{2}-\d{2}$", dv) else ""
    if "project_id" in patch:
        pid = str(patch.get("project_id") or "").strip()
        p_meta = (project_lookup or {}).get(pid) or {}
        cur["project_id"] = pid
        cur["project_code"] = str(p_meta.get("code") or "")
        cur["project_name"] = str(p_meta.get("name") or "")
    if "ref_no" in patch:
        cur["ref_no"] = str(patch.get("ref_no") or "")[:60]
    if "notes" in patch:
        cur["notes"] = str(patch.get("notes") or "")[:600]
    if "done" in patch:
        done_b = bool(patch.get("done"))
        cur["done"] = done_b
        cur["done_at"] = now_iso() if done_b else ""
    cur["updated_at"] = now_iso()
    cur = scrub_secrets(cur)
    store["tasks"][idx] = cur
    _save_tasks_store(state_dir, store)
    return cur


def delete_manual_task(state_dir: Path, task_id: str) -> bool:
    store = _load_tasks_store(state_dir)
    before = len(store["tasks"])
    store["tasks"] = [t for t in store["tasks"] if t.get("id") != task_id]
    if len(store["tasks"]) == before:
        return False
    _save_tasks_store(state_dir, store)
    return True


def dismiss_auto_task(
    state_dir: Path,
    auto_key: str,
    *,
    restore: bool = False,
    meta: dict[str, Any] | None = None,
) -> dict[str, Any]:
    store = _load_tasks_store(state_dir)
    key = str(auto_key or "").strip()[:180]
    if not key:
        raise ValueError("Auto task key is required")
    if restore:
        store["dismissed_auto"].pop(key, None)
        store["dismissed_auto_meta"].pop(key, None)
    else:
        store["dismissed_auto"][key] = dt.date.today().isoformat()
        m_clean: dict[str, Any] = {
            "auto_key": key,
            "done_at": now_iso(),
        }
        if isinstance(meta, dict):
            for f in ("kind", "priority", "project_id", "project_code", "project_name", "ref_no", "subject", "section", "file_rel", "file_area", "target_view"):
                if f in meta:
                    m_clean[f] = str(meta.get(f) or "")[:200]
        store["dismissed_auto_meta"][key] = scrub_secrets(m_clean)
    _save_tasks_store(state_dir, store)
    return {"dismissed_count": len(store["dismissed_auto"])}


def clear_completed_manual_tasks(state_dir: Path, *, project_id: str = "") -> dict[str, int]:
    store = _load_tasks_store(state_dir)
    kept: list[dict[str, Any]] = []
    removed = 0
    for t in store["tasks"]:
        is_done = bool(t.get("done"))
        matches_proj = (not project_id) or str(t.get("project_id") or "") == project_id
        if is_done and matches_proj:
            removed += 1
        else:
            kept.append(t)
    if removed:
        store["tasks"] = kept
        _save_tasks_store(state_dir, store)
    return {"removed": removed}


def _row_age_days(date_str: str, today: dt.date) -> int:
    s = str(date_str or "").strip()[:10]
    if not s:
        return 0
    try:
        return max(0, (today - dt.date.fromisoformat(s)).days)
    except ValueError:
        return 0


def collect_auto_tasks(
    hub: Any,
    *,
    project_id: str = "",
    include_dismissed: bool = False,
) -> list[dict[str, Any]]:
    """
    Collect actionable follow-up items across enabled projects (or one project) without
    mutating any project files or logs.
    """
    gcfg = hub.get_settings()
    store = _load_tasks_store(hub.state_dir)
    dismissed = store.get("dismissed_auto") or {}
    today = dt.date.today()
    today_iso = today.isoformat()

    projects = [
        p for p in hub.list_projects()
        if (project_id and p["id"] == project_id) or (not project_id and p.get("enabled", True))
    ]
    out: list[dict[str, Any]] = []
    seen_keys: set[str] = set()

    for proj in projects:
        pid = proj["id"]
        pcode = str(proj.get("code") or "")
        pname = str(proj.get("name") or "")
        overdue_days = int(proj.get("overdue_days") or 14)

        try:
            log_sum = plog.read_project_log_summary(
                proj,
                limit=350,
                recent_first=True,
                device_id=hub.device_id,
            )
            rows = log_sum.get("rows") or []
        except Exception:
            rows = []

        for r in rows:
            ref_s = str(r.get("ref_no") or "").strip()
            sec = str(r.get("section") or "")
            ds = str(r.get("date") or "")[:10]
            subj = str(r.get("subject") or "").strip()
            sl = str(r.get("status") or "").lower()
            ml = str(r.get("moh_status") or "").lower()
            reply_s = str(r.get("reply_date") or "").strip()

            # 1) Rejected / Revise & Resubmit (Code C / D)
            is_rej = any(k in sl for k in ("reject", "مرفوض", "revise")) or sl in ("c", "d", "code c", "code d")
            is_rej_m = any(k in ml for k in ("reject", "مرفوض", "revise")) or ml in ("c", "d", "code c", "code d")
            if (is_rej or is_rej_m) and not r.get("is_superseded"):
                akey = f"rejected::{pid}::{ref_s or r.get('row')}"
                if akey not in seen_keys and (include_dismissed or dismissed.get(akey, "") < today_iso):
                    seen_keys.add(akey)
                    out.append({
                        "auto_key": akey,
                        "kind": "rejected",
                        "priority": "urgent",
                        "project_id": pid,
                        "project_code": pcode,
                        "project_name": pname,
                        "ref_no": ref_s,
                        "subject": subj or ref_s,
                        "section": sec,
                        "sheet": str(r.get("sheet") or ""),
                        "date": ds,
                        "days": _row_age_days(ds, today),
                        "file_rel": str(r.get("file_rel") or ""),
                        "file_area": str(r.get("file_area") or "root"),
                        "target_view": "section" if sec else "overview",
                    })
                continue

            # 2) Unreplied incoming correspondence
            if sec == "incoming" and r.get("unreplied"):
                wait_d = int(r.get("days_waiting") if r.get("days_waiting") is not None else _row_age_days(ds, today))
                akey = f"unreplied::{pid}::{ref_s or r.get('row')}"
                if akey not in seen_keys and (include_dismissed or dismissed.get(akey, "") < today_iso):
                    seen_keys.add(akey)
                    out.append({
                        "auto_key": akey,
                        "kind": "unreplied",
                        "priority": "urgent" if wait_d >= 5 else "normal",
                        "project_id": pid,
                        "project_code": pcode,
                        "project_name": pname,
                        "ref_no": ref_s,
                        "subject": subj or ref_s,
                        "section": sec,
                        "sheet": str(r.get("sheet") or ""),
                        "date": ds,
                        "days": wait_d,
                        "file_rel": str(r.get("file_rel") or ""),
                        "file_area": str(r.get("file_area") or "root"),
                        "target_view": "section",
                    })
                continue

            # 3) Overdue submittals waiting beyond overdue_days
            is_final_ok = any(k in sl for k in ("approv", "معتمد", "withdraw", "cancel", "ملغي"))
            if sec not in ("incoming", "outgoing", "") and not r.get("is_superseded") and not reply_s and not is_final_ok and ds:
                age = _row_age_days(ds, today)
                if age > overdue_days:
                    akey = f"overdue::{pid}::{ref_s or r.get('row')}"
                    if akey not in seen_keys and (include_dismissed or dismissed.get(akey, "") < today_iso):
                        seen_keys.add(akey)
                        out.append({
                            "auto_key": akey,
                            "kind": "overdue",
                            "priority": "urgent" if age >= overdue_days + 7 else "normal",
                            "project_id": pid,
                            "project_code": pcode,
                            "project_name": pname,
                            "ref_no": ref_s,
                            "subject": subj or ref_s,
                            "section": sec,
                            "sheet": str(r.get("sheet") or ""),
                            "date": ds,
                            "days": age,
                            "file_rel": str(r.get("file_rel") or ""),
                            "file_area": str(r.get("file_area") or "root"),
                            "target_view": "section" if sec else "overview",
                        })

        # 4) Pending scanner items / decisions for this project
        try:
            triage = get_scanner_triage_state(proj, watcher_running=hub.watcher_running, decision_store=hub.decisions)
            sc_pending = len(triage.get("pending") or [])
        except Exception:
            sc_pending = 0
        if sc_pending > 0:
            akey = f"scanner::{pid}"
            if include_dismissed or dismissed.get(akey, "") < today_iso:
                out.append({
                    "auto_key": akey,
                    "kind": "scanner_pending",
                    "priority": "normal",
                    "project_id": pid,
                    "project_code": pcode,
                    "project_name": pname,
                    "ref_no": "",
                    "subject": pname,
                    "section": "",
                    "sheet": "",
                    "date": today_iso,
                    "days": sc_pending,
                    "count": sc_pending,
                    "file_rel": "",
                    "file_area": "root",
                    "target_view": "scanner",
                })

        # 5) Overdue backup alert
        try:
            bkp = check_backup_status(
                proj,
                gcfg,
                warn_days=int(gcfg.get("backup_warn_days") or DEFAULT_BACKUP_WARN_DAYS),
            )
            if bkp.get("stale"):
                akey = f"backup::{pid}"
                if include_dismissed or dismissed.get(akey, "") < today_iso:
                    out.append({
                        "auto_key": akey,
                        "kind": "backup_overdue",
                        "priority": "normal",
                        "project_id": pid,
                        "project_code": pcode,
                        "project_name": pname,
                        "ref_no": "",
                        "subject": pname,
                        "section": "",
                        "sheet": "",
                        "date": str(bkp.get("last_backup_at") or "")[:10],
                        "days": int(bkp.get("days_since") or 0),
                        "file_rel": "",
                        "file_area": "root",
                        "target_view": "backup",
                    })
        except Exception:
            pass

    prio_rank = {"urgent": 0, "normal": 1, "low": 2}
    kind_rank = {"rejected": 0, "unreplied": 1, "overdue": 2, "scanner_pending": 3, "backup_overdue": 4}
    out.sort(
        key=lambda x: (
            prio_rank.get(str(x.get("priority") or "normal"), 1),
            kind_rank.get(str(x.get("kind") or ""), 9),
            -int(x.get("days") or 0),
        )
    )
    return out


def build_tasks_payload(
    hub: Any,
    *,
    project_id: str = "",
    include_done: bool = True,
    limit: int = 80,
) -> dict[str, Any]:
    manual_all = list_manual_tasks(hub.state_dir, project_id=project_id, include_done=True)
    manual_open = [t for t in manual_all if not t.get("done")]
    manual_done = [t for t in manual_all if t.get("done")]
    manual_ret = manual_all if include_done else manual_open

    auto_all = collect_auto_tasks(hub, project_id=project_id, include_dismissed=False)
    lim = max(1, min(int(limit or 80), 250))

    urgent_count = sum(1 for t in manual_open if t.get("priority") == "urgent") + sum(
        1 for a in auto_all if a.get("priority") == "urgent"
    )
    today_iso = dt.date.today().isoformat()
    today_manual = [t for t in manual_open if (t.get("due_date") or "") == today_iso or t.get("priority") == "urgent"]

    # Today's Top 5 Focus mix (urgent manual first, then top urgent auto, then normal manual/auto)
    focus_items: list[dict[str, Any]] = []
    for t in today_manual[:3]:
        focus_items.append({"source": "manual", **t})
    for a in auto_all:
        if len(focus_items) >= 5:
            break
        focus_items.append({"source": "auto", **a})
    for t in manual_open:
        if len(focus_items) >= 5:
            break
        if not any(f.get("id") == t["id"] for f in focus_items if f.get("source") == "manual"):
            focus_items.append({"source": "manual", **t})

    # Completed Archive ("اللي تم"): done manual tasks + snoozed/completed auto follow-ups (newest first)
    store = _load_tasks_store(hub.state_dir)
    done_archive: list[dict[str, Any]] = []
    for t in manual_done:
        done_archive.append({
            "source": "manual",
            **t,
            "done_at": str(t.get("done_at") or t.get("updated_at") or t.get("created_at") or ""),
        })
    for akey, d_day in (store.get("dismissed_auto") or {}).items():
        meta = dict((store.get("dismissed_auto_meta") or {}).get(akey) or {})
        if project_id and str(meta.get("project_id") or "") != project_id:
            continue
        if not meta:
            parts_k = akey.split("::", 2)
            meta = {
                "auto_key": akey,
                "kind": parts_k[0] if len(parts_k) > 0 else "",
                "project_id": parts_k[1] if len(parts_k) > 1 else "",
                "ref_no": parts_k[2] if len(parts_k) > 2 else "",
                "subject": parts_k[2] if len(parts_k) > 2 else akey,
                "done_at": d_day,
            }
        done_archive.append({
            "source": "auto",
            "auto_key": akey,
            "snoozed_until": d_day,
            **meta,
            "done_at": str(meta.get("done_at") or d_day),
        })
    done_archive.sort(key=lambda x: str(x.get("done_at") or ""), reverse=True)

    return {
        "ok": True,
        "manual_tasks": manual_ret[:lim],
        "auto_tasks": auto_all[:lim],
        "today_focus": focus_items[:5],
        "done_tasks": done_archive[:lim],
        "counts": {
            "total_open": len(manual_open) + len(auto_all),
            "urgent": urgent_count,
            "manual_open": len(manual_open),
            "manual_done": len(manual_done),
            "auto_open": len(auto_all),
            "focus_count": len(focus_items[:5]),
            "done_total": len(done_archive),
        },
    }


def lightweight_tasks_counts(hub: Any) -> dict[str, int]:
    """Fast badge counts for /api/state without walking logs every 12s poll."""
    manual_open = list_manual_tasks(hub.state_dir, include_done=False)
    urgent_m = sum(1 for t in manual_open if t.get("priority") == "urgent")
    overdue_p = 0
    scanner_p = 0
    gcfg = hub.get_settings()
    warn_days = int(gcfg.get("backup_warn_days") or DEFAULT_BACKUP_WARN_DAYS)
    for proj in hub.list_projects():
        if not proj.get("enabled", True):
            continue
        try:
            bkp = check_backup_status(proj, gcfg, warn_days=warn_days)
            if bkp.get("stale"):
                overdue_p += 1
        except Exception:
            pass
    pending_dec = hub.decisions.pending_count()
    return {
        "manual_open": len(manual_open),
        "manual_urgent": urgent_m,
        "backup_overdue": overdue_p,
        "pending_decisions": pending_dec,
        "badge_total": len(manual_open) + overdue_p + pending_dec + scanner_p,
    }
