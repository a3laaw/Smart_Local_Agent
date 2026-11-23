#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inbox_pipeline.py — Clean glue layer connecting inbox file ingestion with:
  - `relevance.py` (ignore patterns, locked-file guard, relevance score, conflict detection,
    `rename_in_place` + Undo journal, `new_project_by_name` fuzzy deduplication, `dry_run`)
  - `decisions.py` (Pending Decisions queue + user-editable `learned_rules.json`)
  - `project_log.py` (direct updates into each project's own uploaded `.xlsx`/`.xlsm` log —
    never creating a forced generic system log!)
"""

from __future__ import annotations

import datetime as dt
import os
import re
import secrets
import shutil
from pathlib import Path
from typing import Any

from active_project import is_project_inbox_shared
import ai_cache
from decisions import DecisionStore
from i18n_ar import tr
import project_log as plog
from relevance import (
    DEFAULT_RELEVANCE_THRESHOLD,
    UNRELATED_POLICIES,
    evaluate_relevance,
    is_file_locked,
    is_in_rename_journal,
    match_ignore_patterns,
    normalize_ar,
    parse_ignore_patterns,
    rename_in_place_with_journal,
    route_file_to_entity_project,
)
from security_utils import (
    atomic_write_json,
    normalize_win_path,
    now_iso,
    read_json_safe,
    resolve_local_or_unc,
    sha256_of_file,
)


_SKIP_DIR_PREFIXES = (
    "_System_Agent",
    "_Backups",
    "_Archived_",
    "_Auto_Projects_",
    "2_Archive",
    "3_Project_Logs",
    "4_Project_Templates",
    ".",
    "$RECYCLE.BIN",
    "System Volume Information",
    "DfsrPrivate",
    "#recycle",
    "#snapshot",
)


def _iter_inbox_supported_files(inbox_dir: Path, supported_exts: set[str] | frozenset[str]) -> list[Path]:
    """
    Walk `inbox_dir` safely across local disks and Windows Server UNC shares:
    - Uses `os.walk(..., onerror=...)` so a single unreadable/protected subfolder
      on a Windows Server share never aborts scanning the rest of the share.
    - Prunes system/hidden/internal folders before descending.
    """
    found: list[Path] = []
    if not inbox_dir.exists():
        return found

    def _on_err(_exc: OSError) -> None:
        return None

    for dirpath, dirnames, filenames in os.walk(str(inbox_dir), onerror=_on_err):
        dirnames[:] = sorted(
            d for d in dirnames
            if not any(d.lower().startswith(pref.lower()) for pref in _SKIP_DIR_PREFIXES)
        )
        dp = Path(dirpath)
        for fn in sorted(filenames):
            if fn.startswith(("~$", ".")):
                continue
            fp = dp / fn
            if fp.suffix.lower() in supported_exts:
                found.append(normalize_win_path(fp))
    return found


def archive_dir_for_project(proj: dict[str, Any]) -> Path:
    """Return configured output/archive directory for a project, defaulting to `<root>/2_Archive`."""
    raw = str(proj.get("archive_dir") or "").strip() if isinstance(proj, dict) else ""
    if raw:
        return resolve_local_or_unc(raw)
    return resolve_local_or_unc(str(proj.get("root") or ".")) / "2_Archive"


def rename_journal_path_for(inbox_dir: Path) -> Path:
    return resolve_local_or_unc(inbox_dir) / "_rename_journal.json"


def auto_projects_root_for(proj: dict[str, Any], global_settings: dict[str, Any] | None = None) -> Path:
    gcfg = global_settings or {}
    custom = str(gcfg.get("auto_projects_dir") or "").strip()
    if custom:
        return Path(custom).resolve()
    proj_root = Path(str(proj.get("root") or "")).resolve()
    return proj_root.parent / "_Auto_Projects_مشاريع_مستقلة"


def effective_project_relevance_config(proj: dict[str, Any], global_settings: dict[str, Any] | None = None) -> dict[str, Any]:
    gcfg = global_settings or {}
    pol = str(proj.get("unrelated_policy") or gcfg.get("default_unrelated_policy") or "rename_in_place").strip()
    if pol not in UNRELATED_POLICIES:
        pol = "rename_in_place"
    thr_raw = proj.get("relevance_threshold")
    if thr_raw in (None, ""):
        thr_raw = gcfg.get("relevance_threshold", DEFAULT_RELEVANCE_THRESHOLD)
    try:
        thr = max(0.1, min(0.95, float(thr_raw)))
    except Exception:
        thr = DEFAULT_RELEVANCE_THRESHOLD

    dry_run = bool(proj.get("dry_run", False) or gcfg.get("dry_run", False))
    log_unrelated = bool(proj.get("log_unrelated_sheet", False))
    patterns = parse_ignore_patterns(gcfg.get("ignore_patterns")) + parse_ignore_patterns(proj.get("ignore_patterns"))
    return {
        "unrelated_policy": pol,
        "relevance_threshold": thr,
        "dry_run": dry_run,
        "log_unrelated_sheet": log_unrelated,
        "ignore_patterns": patterns,
    }


def preflight_inbox_file(
    file_path: Path,
    *,
    inbox_root: Path,
    ignore_patterns: list[str] | None = None,
    decision_store: DecisionStore | None = None,
) -> tuple[bool, str]:
    """
    Fast deterministic pre-check before touching or analyzing an inbox file.
    Returns `(should_process, skip_reason)`.
    """
    fp = Path(file_path)
    if not fp.exists() or not fp.is_file():
        return False, "missing"

    ig = match_ignore_patterns(fp, ignore_patterns, inbox_root=inbox_root)
    if ig:
        return False, f"ignored:{ig}"

    if is_file_locked(fp):
        return False, "locked"

    try:
        fhash = sha256_of_file(fp)
    except OSError:
        return False, "unreadable"

    jpath = rename_journal_path_for(inbox_root)
    if is_in_rename_journal(jpath, fp, fhash):
        return False, "already_renamed_in_place"

    if decision_store is not None:
        if decision_store.is_ignored_hash(fhash):
            return False, "ignored_by_user"
        if decision_store.has_pending_for_file(fp, fhash):
            return False, "awaiting_user_decision"

    return True, "ok"


def archive_and_log_relevant_file(
    src_path: Path,
    proj: dict[str, Any],
    meta: dict[str, Any],
    *,
    inbox_root: Path | None = None,
    device_id: str = "",
    copy_only: bool = False,
    direct_move: bool = False,
) -> dict[str, Any]:
    """
    Archive a verified project-relevant file into the project's archive folder (`archive_dir`
    or `<proj_root>/2_Archive`) and record/update it inside the project's uploaded Excel log.
    When `direct_move=True`, copies cleanly to `dest_path`, verifies target integrity, then
    removes ONLY that approved file from the scanner folder (no extra duplicate copy left behind).
    """
    import scan_watcher as sw
    import template_profile as tprof

    src = Path(src_path).resolve()
    proj_root = Path(str(proj["root"])).resolve()
    archive_root = archive_dir_for_project(proj)
    sys_root = proj_root / "_System_Agent"
    archive_root.mkdir(parents=True, exist_ok=True)
    sys_root.mkdir(parents=True, exist_ok=True)

    ext = src.suffix.lower()
    raw_dt = str(meta.get("doc_type") or "other").strip().lower()
    is_img = ext in sw.IMAGE_EXTS
    final_ext = ".pdf" if (is_img and raw_dt != "site_photo" and meta.get("convert_images_to_pdf", True)) else ext

    # Normalize meta keys so compute_destination works for both AI output & manual resolution
    norm_meta = dict(meta)
    doc_dt = str(meta.get("submit_date") or meta.get("doc_date") or meta.get("date") or dt.date.today().isoformat()).strip()
    norm_meta["date"] = doc_dt
    norm_meta.setdefault("subject", str(meta.get("subject_ar") or meta.get("subject") or src.stem))
    norm_meta.setdefault("from_party", str(meta.get("party_name") or meta.get("from_party") or ""))
    if raw_dt in ("incoming", "outgoing"):
        norm_meta["doc_type"] = "letter"
        norm_meta["letter_direction"] = "Incoming" if raw_dt == "incoming" else "Outgoing"

    cfg_stub = {
        "default_project": proj.get("name", "عام"),
        "project_code": proj.get("code", "PRJ"),
        "archive_folder": str(archive_root),
        "folder_layout": "by_type",
        "default_currency": "KWD",
    }
    rel_target, resolved_dt = sw.compute_destination(norm_meta, final_ext, cfg_stub, "")

    # Respect custom section archive_folder if user picked/edited a specific section
    sec_override = str(meta.get("section") or "").strip()
    if sec_override:
        sec_map = tprof.sections_map_for_project(proj)
        if sec_override in sec_map:
            custom_af = str(sec_map[sec_override].get("archive_folder") or "").strip().strip("/\\")
            if custom_af:
                fname_only = Path(rel_target).name
                rel_target = Path(custom_af) / fname_only

    dest_path = archive_root / rel_target
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    if dest_path.exists():
        ts = dt.datetime.now().strftime("%H%M%S")
        cand = dest_path.with_name(f"{dest_path.stem}_{ts}{dest_path.suffix}")
        idx = 2
        while cand.exists():
            cand = dest_path.with_name(f"{dest_path.stem}_{ts}_{idx}{dest_path.suffix}")
            idx += 1
        dest_path = cand
    rel_path = dest_path.relative_to(archive_root).as_posix()

    if final_ext == ".pdf" and is_img:
        sw.convert_image_to_pdf(src, dest_path)
    else:
        src_sha = sha256_of_file(src)
        shutil.copy2(src, dest_path)
        if not dest_path.exists() or dest_path.stat().st_size == 0 or sha256_of_file(dest_path) != src_sha:
            raise OSError(tr("server.file_not_found"))

    if not dest_path.exists() or dest_path.stat().st_size == 0:
        raise OSError(tr("server.file_not_found"))

    raw_backup_path = None
    original_kept_readonly = False
    if not copy_only:
        if direct_move:
            try:
                if src.exists() and resolve_local_or_unc(src) != resolve_local_or_unc(dest_path):
                    src.unlink()
            except OSError:
                # Read-only Windows Server share: archived copy succeeded cleanly, source stays safely on share
                original_kept_readonly = True
        elif inbox_root is not None:
            try:
                raw_backup_path = str(sw.archive_raw_original(src, Path(inbox_root), sys_root))
            except OSError:
                original_kept_readonly = True

    # Map doc_type/direction/subtype cleanly to project's own uploaded Excel log!
    log_sec = raw_dt
    if sec_override in ("incoming", "outgoing", "contract_ipc", "variation_vo", "purchase_po", "petty_cash", "site_report", "general"):
        log_sec = sec_override
    elif sec_override.startswith("submittal"):
        log_sec = "submittal"
    elif resolved_dt == "letter" or raw_dt == "letter":
        ldir = str(norm_meta.get("letter_direction") or "").lower()
        log_sec = "outgoing" if ("out" in ldir or "صادر" in ldir) else "incoming"
    elif resolved_dt == "invoice":
        log_sec = "petty_cash"
    elif resolved_dt == "contract_ipc":
        csub = str(meta.get("contract_subtype") or meta.get("subtype") or "IPC").strip().upper()
        log_sec = "variation_vo" if csub == "VO" else ("purchase_po" if csub == "PO" else "contract_ipc")
    elif resolved_dt in ("submittal", "petty_cash", "site_photo"):
        log_sec = resolved_dt
    elif raw_dt not in ("incoming", "outgoing", "submittal", "contract_ipc", "variation_vo", "purchase_po", "petty_cash", "site_report"):
        log_sec = "general"

    subtype = str(meta.get("subtype") or meta.get("submittal_type") or meta.get("contract_subtype") or "").strip()
    if not subtype and sec_override:
        sec_map = tprof.sections_map_for_project(proj)
        subtype = str(sec_map.get(sec_override, {}).get("subtype") or "")
    authority = str(meta.get("authority") or "").strip()
    party = str(meta.get("party_name") or meta.get("from_party") or meta.get("to_party") or authority or "").strip()
    amt = meta.get("amount") if meta.get("amount") not in (None, "") else meta.get("net_amount")
    tot_amt = meta.get("total_amount") if meta.get("total_amount") not in (None, "") else meta.get("gross_amount")
    status_raw = str(meta.get("approval_code") or meta.get("status") or "").strip()
    # Keep explicit reply dates only — never copy submit_date/today into reply_date when stamp is undated!
    reply_dt = str(meta.get("reply_date") or "").strip()
    cons_reply_dt = str(meta.get("consultant_reply_date") or "").strip()
    cons_act = str(meta.get("consultant_action") or "").strip()
    moh_reply_dt = str(meta.get("moh_reply_date") or "").strip()
    moh_act = str(meta.get("moh_action") or "").strip()
    has_both = bool(meta.get("has_both_stamps", False))

    try:
        logs_dir = plog.logs_dir_for_project(proj_root)
        file_rel_link = Path(os.path.relpath(dest_path, logs_dir)).as_posix()
    except ValueError:
        file_rel_link = str(dest_path)

    log_entry = {
        "section": log_sec,
        "doc_type": log_sec,
        "subtype": subtype,
        "discipline": str(meta.get("discipline") or "").strip(),
        "authority": authority,
        "has_both_stamps": has_both,
        "is_superseded": bool(meta.get("is_superseded", False)),
        "ref_no": str(meta.get("ref_no") or "").strip(),
        "revision": str(meta.get("revision") or "").strip(),
        "doc_date": doc_dt,
        "submit_date": str(meta.get("submit_date") or "").strip(),
        "reply_date": reply_dt,
        "consultant_reply_date": cons_reply_dt,
        "consultant_action": cons_act,
        "moh_reply_date": moh_reply_dt,
        "moh_action": moh_act,
        "subject_ar": str(norm_meta["subject"]).strip(),
        "party_name": party,
        "approval_code": str(meta.get("approval_code") or "").strip(),
        "status": status_raw,
        "amount": amt,
        "total_amount": tot_amt,
        "payment_method": str(meta.get("payment_method") or "").strip(),
        "notes": str(meta.get("notes") or "").strip(),
        "file_rel_path": file_rel_link,
        "file_abs_path": str(dest_path),
    }
    log_res = plog.record_entry_in_project_log(proj, log_entry, device_id=device_id, queue_if_locked=True)

    return {
        "ok": True,
        "outcome": "archived_in_project",
        "project_id": proj.get("id"),
        "archive_path": str(dest_path),
        "archive_rel": rel_path,
        "raw_backup_path": raw_backup_path,
        "original_kept_readonly": original_kept_readonly,
        "log_result": log_res,
        "device_id": device_id,
    }


def handle_inbox_document(
    file_path: Path,
    *,
    extracted_text: str,
    ai_meta: dict[str, Any] | None,
    active_proj: dict[str, Any],
    all_projects: list[dict[str, Any]],
    global_settings: dict[str, Any],
    decision_store: DecisionStore,
    device_id: str,
    inbox_root: Path | None = None,
    copy_only: bool = False,
) -> dict[str, Any]:
    """
    Full deterministic + policy-driven router for a single incoming document:
      1) Pre-flight checks (ignore patterns, lock, already-renamed, pending decision)
      2) Relevance & cross-project conflict check (`relevance.evaluate_relevance`)
      3) Execute policy (`rename_in_place` | `new_project_by_name` | `ask` | `dry_run`)
         OR archive + update project's own uploaded Excel log if relevant!
    """
    src = Path(file_path).resolve()
    ibx = Path(inbox_root).resolve() if inbox_root else Path(str(active_proj.get("inbox_dir") or src.parent)).resolve()
    rcfg = effective_project_relevance_config(active_proj, global_settings)

    ok_pre, pre_reason = preflight_inbox_file(
        src,
        inbox_root=ibx,
        ignore_patterns=rcfg["ignore_patterns"],
        decision_store=decision_store,
    )
    if not ok_pre:
        return {"ok": True, "outcome": "skipped", "reason": pre_reason, "file": src.name}

    shared = is_project_inbox_shared(all_projects, str(active_proj.get("id") or ""))
    rules = decision_store.list_rules(include_disabled=False)
    rel = evaluate_relevance(
        src,
        extracted_text,
        ai_meta,
        active_proj,
        all_projects=all_projects,
        learned_rules=rules,
        threshold=rcfg["relevance_threshold"],
        shared_inbox=shared,
    )
    meta = dict(ai_meta or {})
    dry_run = rcfg["dry_run"]

    # Dry-Run mode: NEVER touch the file on disk — record preview in Pending Decisions!
    if dry_run:
        proposed = "attach_project" if rel["relevant"] else (rel.get("forced_action") or rcfg["unrelated_policy"])
        dec = decision_store.enqueue_decision(
            kind="dry_run_preview",
            file_path=src,
            source_project_id=str(active_proj.get("id") or ""),
            source_project_name=str(active_proj.get("name") or ""),
            reason=rel["reason"] or tr("server.trial_preview"),
            score=rel["score"],
            suggested_project_id=rel.get("target_project_id") or (str(active_proj.get("id")) if rel["relevant"] else None),
            conflict_project_id=rel.get("conflict_project_id"),
            conflict_project_name=rel.get("conflict_project_name", ""),
            suggested_entity=rel.get("suggested_entity", ""),
            suggested_clean_name=rel.get("suggested_clean_name", src.name),
            extracted_excerpt=extracted_text[:400],
            ai_meta=meta,
            signals=rel.get("signals"),
            dry_run=True,
        )
        return {"ok": True, "outcome": "dry_run", "proposed_action": proposed, "relevance": rel, "decision": dec}

    # Relevant to active project -> Archive + write into project's OWN uploaded Excel log!
    if rel["relevant"]:
        res = archive_and_log_relevant_file(
            src,
            active_proj,
            meta,
            inbox_root=ibx,
            device_id=device_id,
            copy_only=copy_only,
        )
        res["relevance"] = rel
        return res

    # Unrelated / Conflict / Low-confidence / Learned Rule
    action = rel.get("forced_action") or rcfg["unrelated_policy"]
    jpath = rename_journal_path_for(ibx)

    if action == "ignore":
        fhash = sha256_of_file(src)
        decision_store.add_ignored_hash(fhash, src.name, reason=rel["reason"])
        return {"ok": True, "outcome": "ignored_by_rule", "relevance": rel}

    if action == "attach_project" and rel.get("target_project_id"):
        target_p = next((p for p in all_projects if str(p.get("id")) == str(rel["target_project_id"])), None)
        proc_mode = str(global_settings.get("process_mode") or "active_only")
        if target_p is not None and (
            rel.get("decision_kind") == "learned_rule" or proc_mode == "all_projects"
        ):
            res = archive_and_log_relevant_file(
                src, target_p, meta, inbox_root=ibx, device_id=device_id, copy_only=copy_only
            )
            res["outcome"] = "routed_by_learned_rule"
            res["relevance"] = rel
            return res
        action = "ask"

    if action == "ask" or rel.get("decision_kind") in ("cross_project_conflict", "low_confidence"):
        dec = decision_store.enqueue_decision(
            kind=str(rel.get("decision_kind") or "unrelated_file"),
            file_path=src,
            source_project_id=str(active_proj.get("id") or ""),
            source_project_name=str(active_proj.get("name") or ""),
            reason=rel["reason"],
            score=rel["score"],
            suggested_project_id=rel.get("target_project_id"),
            conflict_project_id=rel.get("conflict_project_id"),
            conflict_project_name=rel.get("conflict_project_name", ""),
            suggested_entity=rel.get("suggested_entity", ""),
            suggested_clean_name=rel.get("suggested_clean_name", src.name),
            extracted_excerpt=extracted_text[:400],
            ai_meta=meta,
            signals=rel.get("signals"),
            dry_run=False,
        )
        return {"ok": True, "outcome": "pending_decision", "decision": dec, "relevance": rel}

    if action == "new_project_by_name":
        entity = str(rel.get("suggested_entity") or "").strip() or "غير مصنف"
        auto_root = auto_projects_root_for(active_proj, global_settings)
        # Do not match back to active_proj when unrelated
        other_projs = [p for p in all_projects if str(p.get("id")) != str(active_proj.get("id"))]
        routed = route_file_to_entity_project(
            src,
            entity,
            rel.get("suggested_clean_name") or src.name,
            auto_root,
            jpath,
            registered_projects=other_projs,
            device_id=device_id,
            source_project_id=str(active_proj.get("id") or ""),
            dry_run=False,
        )
        if rcfg["log_unrelated_sheet"]:
            plog.record_optional_unrelated_entry(
                active_proj,
                file_name=routed["old_name"],
                new_name=routed["new_path"],
                action_taken="new_project_by_name",
                reason=rel["reason"],
                entity=routed["entity_name"],
                device_id=device_id,
            )
        return {"ok": True, "outcome": "new_project_by_name", "journal_entry": routed, "relevance": rel}

    # Default Policy: `rename_in_place` — clean rename inside its own folder, ZERO archive move, ZERO main log row!
    ren = rename_in_place_with_journal(
        src,
        rel.get("suggested_clean_name") or src.name,
        jpath,
        reason=rel["reason"],
        device_id=device_id,
        project_id=str(active_proj.get("id") or ""),
        dry_run=False,
    )
    if rcfg["log_unrelated_sheet"]:
        plog.record_optional_unrelated_entry(
            active_proj,
            file_name=ren["old_name"],
            new_name=ren["new_name"],
            action_taken="rename_in_place",
            reason=rel["reason"],
            entity=str(rel.get("suggested_entity") or ""),
            device_id=device_id,
        )
    return {"ok": True, "outcome": "renamed_in_place", "journal_entry": ren, "relevance": rel}


def resolve_pending_decision(
    decision_id: str,
    *,
    action: str,
    params: dict[str, Any],
    decision_store: DecisionStore,
    all_projects: list[dict[str, Any]],
    global_settings: dict[str, Any],
    device_id: str,
) -> dict[str, Any]:
    """
    Execute the user's chosen resolution on a pending decision and optionally
    persist a reusable rule in `learned_rules.json`!
    """
    dec = decision_store.get_decision(decision_id)
    if not dec:
        raise KeyError(tr("server.decision_not_found"))
    if dec.get("status") == "resolved":
        raise ValueError(tr("server.decision_done"))

    fp = Path(str(dec.get("file_path") or ""))
    src_proj = next((p for p in all_projects if str(p.get("id")) == str(dec.get("source_project_id"))), None)
    if src_proj is None and all_projects:
        src_proj = all_projects[0]
    ibx = fp.parent if fp.exists() else Path(str((src_proj or {}).get("inbox_dir") or "."))
    jpath = rename_journal_path_for(ibx)

    result_info: dict[str, Any] = {}
    resolved_pid: str | None = None

    if action == "attach_project":
        target_pid = str(params.get("target_project_id") or dec.get("suggested_project_id") or dec.get("source_project_id") or "")
        target_p = next((p for p in all_projects if str(p.get("id")) == target_pid), None)
        if not target_p:
            raise ValueError(tr("validation.rule_project_req"))
        if not fp.exists():
            raise FileNotFoundError(tr("server.file_not_found"))
        meta = dict(dec.get("ai_meta") or {})
        if params.get("doc_type"):
            meta["doc_type"] = params["doc_type"]
        if params.get("subject_ar"):
            meta["subject_ar"] = params["subject_ar"]
        result_info = archive_and_log_relevant_file(fp, target_p, meta, inbox_root=ibx, device_id=device_id)
        resolved_pid = target_pid

    elif action == "rename_in_place":
        if not fp.exists():
            raise FileNotFoundError(tr("server.file_not_found"))
        new_nm = str(params.get("new_name") or dec.get("suggested_clean_name") or fp.name).strip()
        result_info = rename_in_place_with_journal(
            fp,
            new_nm,
            jpath,
            reason=f"decision_resolved:{decision_id}",
            device_id=device_id,
            project_id=str(dec.get("source_project_id") or ""),
        )
        if src_proj and src_proj.get("log_unrelated_sheet"):
            plog.record_optional_unrelated_entry(
                src_proj,
                file_name=result_info["old_name"],
                new_name=result_info["new_name"],
                action_taken="rename_in_place",
                reason=str(dec.get("reason") or ""),
                entity=str(dec.get("suggested_entity") or ""),
                device_id=device_id,
            )

    elif action == "new_project_by_name":
        if not fp.exists():
            raise FileNotFoundError(tr("server.file_not_found"))
        ent = str(params.get("entity_name") or dec.get("suggested_entity") or "غير مصنف").strip()
        new_nm = str(params.get("new_name") or dec.get("suggested_clean_name") or fp.name).strip()
        auto_root = auto_projects_root_for(src_proj or {"root": str(ibx)}, global_settings)
        result_info = route_file_to_entity_project(
            fp,
            ent,
            new_nm,
            auto_root,
            jpath,
            registered_projects=[p for p in all_projects if str(p.get("id")) != str(dec.get("source_project_id"))],
            device_id=device_id,
            source_project_id=str(dec.get("source_project_id") or ""),
        )

    elif action in ("ignore_once", "ignore_pattern"):
        fhash = str(dec.get("file_sha256") or (sha256_of_file(fp) if fp.exists() else ""))
        decision_store.add_ignored_hash(fhash, str(dec.get("file_name") or ""), reason=action)
        pat_added = ""
        if action == "ignore_pattern":
            pat_added = str(params.get("pattern") or f"*{fp.suffix.lower()}").strip()
        result_info = {"ignored": True, "pattern": pat_added}
    else:
        raise ValueError(tr("server.invalid_setting"))

    # Optional: remember rule for future documents!
    created_rule: dict[str, Any] | None = None
    if params.get("remember_rule"):
        rule_pattern = str(
            params.get("rule_pattern")
            or params.get("entity_name")
            or dec.get("suggested_entity")
            or (dec.get("ai_meta") or {}).get("party_name")
            or ""
        ).strip()
        rule_match_type = str(params.get("rule_match_type") or "entity").strip()
        rule_action = "ignore" if action in ("ignore_once", "ignore_pattern") else action
        if rule_pattern:
            created_rule = decision_store.add_rule(
                match_type=rule_match_type,
                pattern=rule_pattern,
                action=rule_action,
                target_project_id=resolved_pid or "",
                entity_name=str(params.get("entity_name") or dec.get("suggested_entity") or ""),
                note=tr("server.rule_from_decision"),
            )

    updated_dec = decision_store.mark_resolved(
        decision_id,
        action=action,
        result_info=result_info,
        resolved_project_id=resolved_pid,
        created_rule_id=created_rule["id"] if created_rule else None,
    )
    return {"ok": True, "decision": updated_dec, "result": result_info, "created_rule": created_rule}


def process_project_inbox_once(
    proj: dict[str, Any],
    *,
    all_projects: list[dict[str, Any]],
    global_settings: dict[str, Any],
    decision_store: DecisionStore,
    device_id: str,
    inbox_override: Path | None = None,
    copy_only: bool = False,
    max_files: int = 200,
    pace_seconds: float = 0.0,
    stop_event: Any = None,
    analyzer_override: Any = None,
) -> dict[str, Any]:
    """
    Scan and process a project's inbox folder once using the full v10 pipeline:
      - Ignore patterns + system/executable exclusions + locked-file guard
      - Deduplication via SHA-256 + `_rename_journal.json` + `pending_decisions.json`
      - Text extraction + Vision/AI classification (with relevance fields)
      - Deterministic + AI relevance scoring & cross-project conflict guard
      - Unrelated policies (`rename_in_place` + Undo / `new_project_by_name` / `ask` / `dry_run`)
      - Direct archiving + updating the project's OWN uploaded Excel log(s)!
    """
    import json
    import time
    import scan_watcher as sw

    proj_root = Path(str(proj["root"])).resolve()
    inbox_root = Path(inbox_override or proj.get("inbox_dir") or (proj_root / "1_Inbox")).resolve()
    if not inbox_root.exists():
        return {"ok": False, "error": tr("server.folder_not_found", path=str(inbox_root)), "processed": 0}

    sys_root = proj_root / "_System_Agent"
    sys_root.mkdir(parents=True, exist_ok=True)
    state_path = sys_root / "watcher_state.json"
    try:
        state = json.loads(state_path.read_text("utf-8")) if state_path.exists() else {}
    except Exception:
        state = {}
    hashes: dict[str, str] = state.setdefault("hashes", {})

    rcfg = effective_project_relevance_config(proj, global_settings)
    supported_exts = sw.SCANNED_EXTS | sw.NATIVE_DRAFT_EXTS | {".xlsm", ".txt"}
    triage_store = _load_triage_store(proj_root)
    triage_hashes = (
        {str(h) for h in (triage_store.get("rejected_hashes") or []) if h}
        | {str(h) for h in (triage_store.get("approved_hashes") or []) if h}
        | {
            str(it.get("file_sha256") or "")
            for it in triage_store.get("items", [])
            if it.get("file_sha256")
        }
    )
    triage_paths = {
        str(normalize_win_path(str(it.get("file_path") or ""))).lower()
        for it in triage_store.get("items", [])
        if it.get("file_path") and it.get("status") in ("rejected", "approved")
    } | {
        str(normalize_win_path(str(p))).lower()
        for p in (triage_store.get("rejected_paths") or []) + (triage_store.get("approved_paths") or [])
        if p
    }

    summary = {
        "ok": True,
        "processed": 0,
        "archived": 0,
        "renamed_in_place": 0,
        "new_project": 0,
        "pending_decisions": 0,
        "dry_run": 0,
        "skipped": 0,
        "warnings": [],
        "errors": [],
        "items": [],
    }

    # Walk inbox safely (supports Windows Server UNC shares & skips unreadable system dirs)
    candidates: list[Path] = []
    for p in _iter_inbox_supported_files(inbox_root, supported_exts):
        ig = match_ignore_patterns(p, rcfg["ignore_patterns"], inbox_root=inbox_root)
        if ig:
            summary["skipped"] += 1
            continue
        candidates.append(p)

    engine = str(global_settings.get("engine") or proj.get("engine") or "gemini").lower()
    ollama_url = str(global_settings.get("ollama_url") or "http://127.0.0.1:11434")
    ollama_model = str(global_settings.get("ollama_model") or "qwen2.5vl:7b")

    for src in candidates:
        if summary["processed"] >= max_files:
            break
        if stop_event is not None and getattr(stop_event, "is_set", lambda: False)():
            break

        ok_pre, _ = preflight_inbox_file(
            src,
            inbox_root=inbox_root,
            ignore_patterns=rcfg["ignore_patterns"],
            decision_store=decision_store,
        )
        if not ok_pre:
            summary["skipped"] += 1
            continue

        fhash = sha256_of_file(src)
        norm_src = str(normalize_win_path(src)).lower()
        if fhash in hashes or fhash in triage_hashes or norm_src in triage_paths:
            summary["skipped"] += 1
            continue

        ext = src.suffix.lower()
        extracted_text = ""
        meta: dict[str, Any] = {}
        pages = 1
        job_dir = sys_root / "_tmp_job"
        cache_dir = ai_cache.project_ai_cache_dir(proj_root)
        doc_cache_key = f"{fhash}:{ai_cache.stable_sig(proj.get('id'), proj.get('code'), proj.get('domain'), sorted(proj.get('custom_categories') or []))}"
        try:
            cached_doc = ai_cache.get_cached(cache_dir, "doc_classify", doc_cache_key)
            if cached_doc and isinstance(cached_doc.get("meta"), dict) and cached_doc["meta"]:
                meta = dict(cached_doc["meta"])
                extracted_text = str(cached_doc.get("extracted_text") or "")
                pages = int(cached_doc.get("pages") or 1)
            else:
                if ext in (".docx", ".xlsx", ".xlsm", ".doc", ".xls"):
                    extracted_text = sw.extract_native_office_text(src)
                    images: list[Path] = []
                    pages = 1
                elif ext == ".txt":
                    extracted_text = src.read_text("utf-8", errors="ignore")[:8000]
                    images = []
                    pages = 1
                else:
                    shutil.rmtree(job_dir, ignore_errors=True)
                    job_dir.mkdir(parents=True, exist_ok=True)
                    images, extracted_text, pages = sw.prepare_visual_pages(src, job_dir, dpi=150, stamp_dpi=200)

                prompt = sw.build_classification_prompt(
                    image_names=[im.name for im in images],
                    orig_name=src.name,
                    subfolder_hint="",
                    page_count=pages,
                    embedded_text=extracted_text,
                    existing_projects=[str(p.get("name") or "") for p in all_projects if p.get("name")],
                    default_currency="KWD",
                    include_at_refs=(engine == "gemini" and bool(images)),
                    domain=str(proj.get("domain") or "contracting"),
                    custom_categories=list(proj.get("custom_categories") or []),
                    active_project=proj,
                )

                if analyzer_override is not None:
                    meta = analyzer_override(src, extracted_text, prompt) or {}
                elif images:
                    if engine == "ollama":
                        meta = sw.analyze_with_ollama(prompt, images, ollama_url, ollama_model)
                    else:
                        # Never store API key in config dicts — loaded transiently inside client only!
                        meta = sw.analyze_with_gemini(prompt, job_dir, sw.load_saved_gemini_model(), images=images)
                else:
                    # Native text/Office file: use AI if key available, else pure deterministic fallback!
                    if engine == "ollama":
                        try:
                            meta = sw.analyze_with_ollama(prompt, [], ollama_url, ollama_model)
                        except Exception as exc_ai:
                            meta = {}
                            summary["warnings"].append(f"{src.name}: {exc_ai}")
                    elif sw.load_saved_gemini_key():
                        try:
                            meta = sw.analyze_with_gemini(prompt, job_dir if job_dir.exists() else sys_root, sw.load_saved_gemini_model(), images=[])
                        except Exception as exc_ai:
                            meta = {}
                            summary["warnings"].append(f"{src.name}: {exc_ai}")
                if isinstance(meta, dict) and meta.get("doc_type") and not meta.get("is_blank_or_unreadable"):
                    ai_cache.set_cached(cache_dir, "doc_classify", doc_cache_key, {
                        "meta": meta,
                        "extracted_text": extracted_text[:4000],
                        "pages": int(pages or 1),
                    })

            res = handle_inbox_document(
                src,
                extracted_text=extracted_text,
                ai_meta=meta,
                active_proj=proj,
                all_projects=all_projects,
                global_settings=global_settings,
                decision_store=decision_store,
                device_id=device_id,
                inbox_root=inbox_root,
                copy_only=copy_only,
            )
            outcome = res.get("outcome", "")
            summary["processed"] += 1
            if outcome in ("archived_in_project", "routed_by_learned_rule"):
                summary["archived"] += 1
                hashes[fhash] = str(res.get("archive_rel") or src.name)
                state["total_archived"] = int(state.get("total_archived", 0)) + 1
                from security_utils import atomic_write_json
                atomic_write_json(state_path, state)
            elif outcome == "renamed_in_place":
                summary["renamed_in_place"] += 1
            elif outcome == "new_project_by_name":
                summary["new_project"] += 1
                if copy_only:
                    hashes[fhash] = src.name
            elif outcome == "pending_decision":
                summary["pending_decisions"] += 1
            elif outcome == "dry_run":
                summary["dry_run"] += 1
            else:
                summary["skipped"] += 1

            summary["items"].append({"file": src.name, "outcome": outcome})
            if pace_seconds > 0:
                time.sleep(pace_seconds)
        except Exception as exc:
            summary["errors"].append(f"{src.name}: {exc}")
        finally:
            shutil.rmtree(job_dir, ignore_errors=True)

    from security_utils import atomic_write_json
    state["last_scan_at"] = now_iso()
    state["last_errors"] = list(summary["errors"])[-20:]
    state["last_warnings"] = list(summary["warnings"])[-20:]
    atomic_write_json(state_path, state)

    return summary


# ══════════════════════════════════════════════════════════════════════════════
# DEDICATED SCANNER TRIAGE QUEUE (Safe Read-Only Scan -> Approve/Reject UI)
# ══════════════════════════════════════════════════════════════════════════════
import threading

_TRIAGE_LOCK = threading.RLock()

_SEC_SUMMARY_KEY: dict[str, str] = {
    "submittal_mar": "scanner.sum_type_mar",
    "submittal_mir": "scanner.sum_type_mir",
    "submittal_ir": "scanner.sum_type_ir",
    "submittal_sd": "scanner.sum_type_sd",
    "submittal_ms": "scanner.sum_type_ms",
    "outgoing": "scanner.sum_type_outgoing",
    "incoming": "scanner.sum_type_incoming",
    "contract_ipc": "scanner.sum_type_ipc",
    "variation_vo": "scanner.sum_type_vo",
    "purchase_po": "scanner.sum_type_po",
    "petty_cash": "scanner.sum_type_petty",
    "site_report": "scanner.sum_type_report",
    "general": "scanner.sum_type_general",
}


def triage_store_path_for(proj_root: Path | str) -> Path:
    d = Path(proj_root).resolve() / "_System_Agent"
    d.mkdir(parents=True, exist_ok=True)
    return d / "scanner_triage.json"


def _load_triage_store(proj_root: Path | str) -> dict[str, Any]:
    with _TRIAGE_LOCK:
        p = triage_store_path_for(proj_root)
        raw = read_json_safe(p, {})
        if not isinstance(raw, dict):
            raw = {}
        items = raw.get("items")
        if not isinstance(items, list):
            items = []
        return {
            "schema_version": "1.0",
            "last_scan_at": raw.get("last_scan_at"),
            "last_errors": list(raw.get("last_errors") or []),
            "items": items,
            "rejected_hashes": [str(x) for x in (raw.get("rejected_hashes") or []) if x],
            "approved_hashes": [str(x) for x in (raw.get("approved_hashes") or []) if x],
            "rejected_paths": [str(x) for x in (raw.get("rejected_paths") or []) if x],
            "approved_paths": [str(x) for x in (raw.get("approved_paths") or []) if x],
        }


def _save_triage_store(proj_root: Path | str, data: dict[str, Any]) -> None:
    with _TRIAGE_LOCK:
        atomic_write_json(triage_store_path_for(proj_root), data)


def _infer_section_from_meta(
    meta: dict[str, Any],
    src_name: str,
    extracted_text: str,
    proj: dict[str, Any],
) -> str:
    import template_profile as tprof

    explicit = str(meta.get("section") or "").strip()
    sec_map = tprof.sections_map_for_project(proj)
    if explicit in sec_map:
        return explicit

    raw_dt = str(meta.get("doc_type") or "").strip().lower()
    sub = str(meta.get("subtype") or meta.get("submittal_type") or meta.get("contract_subtype") or "").strip().upper()
    ldir = str(meta.get("letter_direction") or "").strip().lower()

    if raw_dt in ("submittal", "submittals"):
        return {
            "MAR": "submittal_mar",
            "MIR": "submittal_mir",
            "IR": "submittal_ir",
            "WIR": "submittal_ir",
            "SD": "submittal_sd",
            "MS": "submittal_ms",
            "PQ": "submittal_ms",
        }.get(sub, "submittal_mar")
    if raw_dt in ("outgoing",):
        return "outgoing"
    if raw_dt in ("incoming",):
        return "incoming"
    if raw_dt == "letter":
        return "outgoing" if ("out" in ldir) else "incoming"
    if raw_dt == "contract_ipc":
        if sub == "VO":
            return "variation_vo"
        if sub == "PO":
            return "purchase_po"
        return "contract_ipc"
    if raw_dt in ("variation_vo", "purchase_po", "petty_cash", "site_report"):
        return raw_dt
    if raw_dt in ("invoice", "voucher", "receipt"):
        return "petty_cash"
    if raw_dt in ("site_photo", "report"):
        return "site_report"

    # Fallback deterministic filename/text hint using template_profile's strict token rules
    sugg = tprof.suggest_template_section(src_name)
    if sugg.get("suggested_section") in sec_map:
        return str(sugg["suggested_section"])

    txt_l = f"{src_name} {extracted_text[:600]}".lower()
    if "mar" in txt_l:
        return "submittal_mar"
    if "wir" in txt_l:
        return "submittal_ir"
    if "mir" in txt_l:
        return "submittal_mir"
    if "ipc" in txt_l:
        return "contract_ipc"
    if "invoice" in txt_l or "receipt" in txt_l or "kwd" in txt_l:
        return "petty_cash"
    return "general"


def _meta_to_section_meta(meta: dict[str, Any], section_id: str, src_name: str, proj: dict[str, Any]) -> dict[str, Any]:
    import template_profile as tprof

    sec_map = tprof.sections_map_for_project(proj)
    sec_info = sec_map.get(section_id) or {}
    out = dict(meta)
    out["section"] = section_id
    log_sec = str(sec_info.get("log_section") or "general")
    subtype = str(out.get("subtype") or out.get("submittal_type") or out.get("contract_subtype") or sec_info.get("subtype") or "")

    if section_id.startswith("submittal"):
        out["doc_type"] = "submittal"
        out["submittal_type"] = subtype or sec_info.get("subtype") or "MAR"
        out["subtype"] = out["submittal_type"]
    elif section_id == "outgoing":
        out["doc_type"] = "letter"
        out["letter_direction"] = "Outgoing"
        out["subtype"] = "OUT"
    elif section_id == "incoming":
        out["doc_type"] = "letter"
        out["letter_direction"] = "Incoming"
        out["subtype"] = "IN"
    elif section_id in ("contract_ipc", "variation_vo", "purchase_po"):
        out["doc_type"] = "contract_ipc"
        out["contract_subtype"] = {"contract_ipc": "IPC", "variation_vo": "VO", "purchase_po": "PO"}[section_id]
        out["subtype"] = out["contract_subtype"]
    elif section_id == "petty_cash":
        out["doc_type"] = "petty_cash"
    elif section_id == "site_report":
        out["doc_type"] = "site_report"
    else:
        out["doc_type"] = log_sec or "general"
        if subtype:
            out["subtype"] = subtype

    if not out.get("subject_ar") and not out.get("subject"):
        out["subject_ar"] = Path(src_name).stem.replace("_", " ").strip()
    auth_s = str(out.get("authority") or "").strip()
    is_moh = plog._is_moh_authority(auth_s) or bool(out.get("moh_reply_date") or out.get("moh_action"))
    has_cons = plog._has_consultant_stamp(auth_s, out)
    if is_moh and (has_cons or bool(out.get("has_both_stamps"))):
        out["authority_role"] = "MOH+CONSULTANT"
    elif is_moh:
        out["authority_role"] = "MOH"
    elif has_cons:
        out["authority_role"] = "SOOR"
    else:
        out["authority_role"] = ""
    return out


def _preview_target_rel(norm_meta: dict[str, Any], src: Path, proj: dict[str, Any], section_id: str) -> str:
    import scan_watcher as sw
    import template_profile as tprof

    archive_root = archive_dir_for_project(proj)
    sec_map = tprof.sections_map_for_project(proj)
    ext = src.suffix.lower()
    raw_dt = str(norm_meta.get("doc_type") or "other").strip().lower()
    final_ext = ".pdf" if (ext in sw.IMAGE_EXTS and raw_dt != "site_photo" and norm_meta.get("convert_images_to_pdf", True)) else ext
    prep = dict(norm_meta)
    prep.setdefault("date", str(prep.get("doc_date") or dt.date.today().isoformat()))
    prep.setdefault("subject", str(prep.get("subject_ar") or prep.get("subject") or src.stem))
    prep.setdefault("from_party", str(prep.get("party_name") or prep.get("from_party") or ""))
    cfg_stub = {
        "default_project": proj.get("name", ""),
        "project_code": proj.get("code", "PRJ"),
        "archive_folder": str(archive_root),
        "folder_layout": "by_type",
        "default_currency": "KWD",
    }
    try:
        rel_target, _ = sw.compute_destination(prep, final_ext, cfg_stub, "")
        custom_af = str((sec_map.get(section_id) or {}).get("archive_folder") or "").strip().strip("/\\")
        if custom_af:
            return (Path(custom_af) / Path(rel_target).name).as_posix()
        return Path(rel_target).as_posix()
    except Exception:
        custom_af = str((sec_map.get(section_id) or {}).get("archive_folder") or "09_General").strip().strip("/\\")
        return f"{custom_af}/{src.stem}{final_ext}"


def _build_human_summary_ar(
    src_name: str,
    extracted_text: str,
    meta: dict[str, Any],
    section_id: str,
    rel: dict[str, Any],
    proj: dict[str, Any],
) -> tuple[str, str]:
    import template_profile as tprof

    sec_map = tprof.sections_map_for_project(proj)
    sec_key = _SEC_SUMMARY_KEY.get(section_id)
    type_lbl = tr(sec_key) if sec_key else str((sec_map.get(section_id) or {}).get("label") or tr("scanner.sum_type_general"))

    subj = str(meta.get("subject_ar") or meta.get("subject") or "").strip()
    if not subj and extracted_text:
        for ln in extracted_text.splitlines():
            clean_ln = " ".join(ln.split()).strip()
            if len(clean_ln) >= 6 and not clean_ln.isdigit():
                subj = clean_ln[:110]
                break
    if not subj:
        subj = Path(src_name).stem.replace("_", " ").replace("-", " ").strip() or tr("scanner.unreadable_summary")

    headline = f"{type_lbl}: {subj}"
    parts: list[str] = []
    ref_no = str(meta.get("ref_no") or "").strip()
    disc = str(meta.get("discipline") or "").strip()
    party = str(meta.get("party_name") or meta.get("from_party") or meta.get("authority") or "").strip()
    doc_dt = str(meta.get("submit_date") or meta.get("doc_date") or meta.get("date") or "").strip()
    cons_rd = str(meta.get("consultant_reply_date") or "").strip()
    moh_rd = str(meta.get("moh_reply_date") or "").strip()
    amt = meta.get("amount") if meta.get("amount") not in (None, "") else meta.get("net_amount")
    st_val = str(meta.get("approval_code") or meta.get("status") or "").strip()

    if ref_no:
        parts.append(tr("scanner.sum_part_ref", val=ref_no))
    if disc:
        parts.append(tr("scanner.sum_part_disc", val=disc))
    if party:
        parts.append(tr("scanner.sum_part_party", val=party[:45]))
    if doc_dt:
        parts.append(tr("scanner.sum_part_date", val=doc_dt))
    if cons_rd:
        parts.append(tr("scanner.sum_part_cons_reply", val=cons_rd))
    if moh_rd:
        parts.append(tr("scanner.sum_part_moh_reply", val=moh_rd))
    if amt not in (None, "", 0):
        parts.append(tr("scanner.sum_part_amount", val=str(amt)))
    if st_val:
        parts.append(tr("scanner.sum_part_status", val=st_val))

    if rel and not rel.get("relevant", rel.get("is_relevant", True)):
        parts.append(tr("scanner.outside_match"))
    elif rel:
        parts.append(tr("scanner.project_match"))

    excerpt = ""
    if extracted_text:
        compact = " ".join(extracted_text.split())
        if compact and compact[:60] not in subj:
            excerpt = compact[:180]
    details = " | ".join(parts)
    if excerpt:
        details = f"{details}\n{excerpt}" if details else excerpt
    return headline, details


def get_scanner_triage_state(
    proj: dict[str, Any],
    *,
    watcher_running: bool = False,
    decision_store: Any = None,
) -> dict[str, Any]:
    """
    Return the current Scanner Triage state for a project (`pending`, `approved`, `rejected`),
    plus live folder status and number of unscanned files waiting in `inbox_dir`.
    """
    import scan_watcher as sw

    proj_root = resolve_local_or_unc(str(proj["root"]))
    inbox_raw = str(proj.get("inbox_dir") or "").strip()
    inbox_dir = resolve_local_or_unc(inbox_raw) if inbox_raw else (proj_root / "1_Inbox")
    archive_dir = archive_dir_for_project(proj)
    inbox_exists = inbox_dir.exists() and inbox_dir.is_dir()

    store = _load_triage_store(proj_root)
    rej_hashes = set(store.get("rejected_hashes") or [])
    app_hashes = set(store.get("approved_hashes") or [])
    known_paths = {
        str(normalize_win_path(str(it.get("file_path") or ""))).lower()
        for it in store["items"]
        if it.get("file_path")
    } | {
        str(normalize_win_path(str(p))).lower()
        for p in (store.get("rejected_paths") or []) + (store.get("approved_paths") or [])
        if p
    }

    supported_exts = sw.SCANNED_EXTS | sw.NATIVE_DRAFT_EXTS | {".xlsm", ".txt"}
    folder_files_count = 0
    unscanned_count = 0
    if inbox_exists:
        for p in _iter_inbox_supported_files(inbox_dir, supported_exts):
            folder_files_count += 1
            # Lightweight stat-free path match on UI render so SMB/network shares never lag!
            if str(normalize_win_path(p)).lower() not in known_paths:
                unscanned_count += 1

    pending: list[dict[str, Any]] = []
    approved: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []

    for it in store["items"]:
        item = dict(it)
        fp = Path(str(item.get("file_path") or ""))
        item["file_exists"] = fp.exists() and fp.is_file() if item.get("file_path") else False
        st = item.get("status", "pending")
        fhash = str(item.get("file_sha256") or "")
        if st == "pending":
            # Hide from pending queue if file no longer exists in inbox or was ignored/resolved elsewhere
            if not item["file_exists"]:
                continue
            if fhash and (fhash in rej_hashes or fhash in app_hashes):
                continue
            if decision_store is not None and fhash and decision_store.is_ignored_hash(fhash):
                continue
        if st == "pending" and str(item.get("ext") or fp.suffix).lower() == ".pdf" and item.get("page_count") is None:
            if item["file_exists"]:
                try:
                    import fitz  # PyMuPDF
                    with fitz.open(fp) as _d:
                        item["page_count"] = int(_d.page_count or 1)
                except Exception:
                    try:
                        raw_pdf = fp.read_bytes()
                        pgs_m = re.findall(rb"/Type\s*/Page(?![sA-Za-z])", raw_pdf)
                        item["page_count"] = max(1, len(pgs_m))
                    except Exception:
                        item["page_count"] = 1
            else:
                item["page_count"] = 1
        elif item.get("page_count") is None:
            item["page_count"] = 1

        item["is_multipage"] = bool(
            str(item.get("ext") or fp.suffix).lower() == ".pdf" and int(item.get("page_count") or 1) >= 2
        )
        if st == "approved":
            approved.append(item)
        elif st == "rejected":
            rejected.append(item)
        else:
            pending.append(item)

    # Surface multi-page PDFs (>= 2 pages) FIRST so user can AI-split or confirm before filing!
    pending.sort(
        key=lambda x: (
            0 if x.get("is_multipage") else 1,
            -int(x.get("page_count") or 1) if x.get("is_multipage") else 0,
            str(x.get("scanned_at") or ""),
        )
    )
    # Newest first for approved/rejected
    approved.sort(key=lambda x: str(x.get("resolved_at") or x.get("scanned_at") or ""), reverse=True)
    rejected.sort(key=lambda x: str(x.get("resolved_at") or x.get("scanned_at") or ""), reverse=True)
    multipage_count = sum(1 for x in pending if x.get("is_multipage"))

    return {
        "ok": True,
        "project_id": proj.get("id"),
        "inbox_dir": str(inbox_dir),
        "inbox_exists": inbox_exists,
        "archive_dir": str(archive_dir),
        "folder_files_count": folder_files_count,
        "unscanned_count": unscanned_count,
        "watcher_running": bool(watcher_running),
        "last_scan_at": store.get("last_scan_at"),
        "last_errors": store.get("last_errors") or [],
        "counts": {
            "pending": len(pending),
            "multipage": multipage_count,
            "approved": len(approved),
            "rejected": len(rejected),
        },
        "pending": pending,
        "approved": approved[:100],
        "rejected": rejected[:100],
    }


def scan_inbox_to_triage(
    proj: dict[str, Any],
    *,
    all_projects: list[dict[str, Any]],
    global_settings: dict[str, Any] | None = None,
    decision_store: DecisionStore | None = None,
    analyzer_override: Any = None,
    max_files: int = 30,
    stop_event: Any = None,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Scan the project's scanner folder (`inbox_dir`) in STRICT READ-ONLY MODE:
    - Never moves, renames, or deletes any original file in the scanner folder!
    - Skips files already recorded in triage (`pending`, `approved`, or `rejected`) by SHA-256.
    - Generates an Arabic summary ("ملخص هو بتاع إيه"), detected metadata, and suggested target path.
    - Places new items into `pending` ready for explicit user Approve / Reject.
    """
    import scan_watcher as sw
    import template_profile as tprof

    proj_root = resolve_local_or_unc(str(proj["root"]))
    sys_root = proj_root / "_System_Agent"
    sys_root.mkdir(parents=True, exist_ok=True)

    inbox_raw = str(proj.get("inbox_dir") or "").strip()
    inbox_dir = resolve_local_or_unc(inbox_raw) if inbox_raw else (proj_root / "1_Inbox")

    summary: dict[str, Any] = {
        "ok": True,
        "project_id": proj.get("id"),
        "inbox_dir": str(inbox_dir),
        "inbox_exists": inbox_dir.exists() and inbox_dir.is_dir(),
        "added": 0,
        "skipped": 0,
        "errors": [],
        "warnings": [],
        "items": [],
    }
    if not summary["inbox_exists"]:
        return summary

    rcfg = effective_project_relevance_config(proj, global_settings)
    gcfg = global_settings or {}
    supported_exts = sw.SCANNED_EXTS | sw.NATIVE_DRAFT_EXTS | {".xlsm", ".txt"}
    store = _load_triage_store(proj_root)
    known_hashes = (
        {
            str(it.get("file_sha256") or "")
            for it in store["items"]
            if it.get("file_sha256")
        }
        | set(store.get("rejected_hashes") or [])
        | set(store.get("approved_hashes") or [])
    )

    candidates: list[Path] = []
    for p in _iter_inbox_supported_files(inbox_dir, supported_exts):
        ig = match_ignore_patterns(p, rcfg["ignore_patterns"], inbox_root=inbox_dir)
        if ig:
            summary["skipped"] += 1
            continue
        if is_file_locked(p, allow_readonly=True):
            summary["skipped"] += 1
            continue
        candidates.append(p)
    # Process PDFs first (multi-page PDFs get prioritized in triage pending list)
    candidates.sort(key=lambda p: (0 if p.suffix.lower() == ".pdf" else 1, p.name.lower()))

    engine = str(gcfg.get("engine") or proj.get("engine") or "gemini").lower()
    ollama_url = str(gcfg.get("ollama_url") or "http://127.0.0.1:11434")
    ollama_model = str(gcfg.get("ollama_model") or "qwen2.5vl:7b")

    for src in candidates:
        if summary["added"] >= max_files:
            break
        if stop_event is not None and getattr(stop_event, "is_set", lambda: False)():
            break

        try:
            fhash = sha256_of_file(src)
        except OSError as exc:
            summary["errors"].append(f"{src.name}: {exc}")
            continue

        if fhash in known_hashes or (decision_store is not None and decision_store.is_ignored_hash(fhash)):
            summary["skipped"] += 1
            continue

        ext = src.suffix.lower()
        extracted_text = ""
        meta: dict[str, Any] = {}
        pages = 1
        ai_from_cache = False
        job_dir = sys_root / "_tmp_triage"
        cache_dir = ai_cache.project_ai_cache_dir(proj_root)
        doc_cache_key = f"{fhash}:{ai_cache.stable_sig(proj.get('id'), proj.get('code'), proj.get('domain'), sorted(proj.get('custom_categories') or []))}"
        try:
            cached_doc = ai_cache.get_cached(cache_dir, "doc_classify", doc_cache_key)
            if cached_doc and isinstance(cached_doc.get("meta"), dict) and cached_doc["meta"]:
                meta = dict(cached_doc["meta"])
                extracted_text = str(cached_doc.get("extracted_text") or "")
                pages = int(cached_doc.get("pages") or 1)
                ai_from_cache = True
            else:
                if ext in (".docx", ".xlsx", ".xlsm", ".doc", ".xls"):
                    extracted_text = sw.extract_native_office_text(src)
                    images: list[Path] = []
                    pages = 1
                elif ext == ".txt":
                    extracted_text = src.read_text("utf-8", errors="ignore")[:8000]
                    images = []
                    pages = 1
                else:
                    shutil.rmtree(job_dir, ignore_errors=True)
                    job_dir.mkdir(parents=True, exist_ok=True)
                    images, extracted_text, pages = sw.prepare_visual_pages(src, job_dir, dpi=140, stamp_dpi=180)

                prompt = sw.build_classification_prompt(
                    image_names=[im.name for im in images],
                    orig_name=src.name,
                    subfolder_hint="",
                    page_count=pages,
                    embedded_text=extracted_text,
                    existing_projects=[str(p.get("name") or "") for p in all_projects if p.get("name")],
                    default_currency="KWD",
                    include_at_refs=(engine == "gemini" and bool(images)),
                    domain=str(proj.get("domain") or "contracting"),
                    custom_categories=list(proj.get("custom_categories") or []),
                    active_project=proj,
                )

                if analyzer_override is not None:
                    meta = analyzer_override(src, extracted_text, prompt) or {}
                elif images:
                    if engine == "ollama":
                        try:
                            meta = sw.analyze_with_ollama(prompt, images, ollama_url, ollama_model)
                        except Exception as exc_ai:
                            meta = {}
                            summary["warnings"].append(f"{src.name}: {exc_ai}")
                    elif sw.load_saved_gemini_key():
                        try:
                            meta = sw.analyze_with_gemini(prompt, job_dir, sw.load_saved_gemini_model(), images=images)
                        except Exception as exc_ai:
                            meta = {}
                            summary["warnings"].append(f"{src.name}: {exc_ai}")
                else:
                    if engine == "ollama":
                        try:
                            meta = sw.analyze_with_ollama(prompt, [], ollama_url, ollama_model)
                        except Exception as exc_ai:
                            meta = {}
                            summary["warnings"].append(f"{src.name}: {exc_ai}")
                    elif sw.load_saved_gemini_key():
                        try:
                            meta = sw.analyze_with_gemini(prompt, job_dir if job_dir.exists() else sys_root, sw.load_saved_gemini_model(), images=[])
                        except Exception as exc_ai:
                            meta = {}
                            summary["warnings"].append(f"{src.name}: {exc_ai}")
                if isinstance(meta, dict) and meta.get("doc_type") and not meta.get("is_blank_or_unreadable"):
                    ai_cache.set_cached(cache_dir, "doc_classify", doc_cache_key, {
                        "meta": meta,
                        "extracted_text": extracted_text[:4000],
                        "pages": int(pages or 1),
                    })

            if not isinstance(meta, dict):
                meta = {}

            rel = evaluate_relevance(
                file_path=src,
                extracted_text=extracted_text,
                ai_meta=meta,
                active_proj=proj,
                all_projects=all_projects,
                threshold=rcfg["relevance_threshold"],
            )
            section_id = _infer_section_from_meta(meta, src.name, extracted_text, proj)
            norm_meta = _meta_to_section_meta(meta, section_id, src.name, proj)
            sec_map = tprof.sections_map_for_project(proj)
            sec_label = str((sec_map.get(section_id) or {}).get("label") or section_id)
            headline, details = _build_human_summary_ar(src.name, extracted_text, norm_meta, section_id, rel, proj)
            target_rel = _preview_target_rel(norm_meta, src, proj, section_id)

            suggested_links: list[dict[str, Any]] = []
            if section_id in ("incoming", "outgoing"):
                try:
                    import letter_links as llinks
                    suggested_links = llinks.suggest_letter_links(
                        proj,
                        ref_no=str(norm_meta.get("ref_no") or ""),
                        section=section_id,
                        subject=str(norm_meta.get("subject_ar") or norm_meta.get("subject") or ""),
                        doc_date=str(norm_meta.get("doc_date") or ""),
                        text=extracted_text[:2500],
                        party_name=str(norm_meta.get("party_name") or norm_meta.get("party") or ""),
                        limit=5,
                        min_score=0.48,
                        device_id=device_id,
                    )
                    if suggested_links and suggested_links[0].get("score", 0) >= 0.58:
                        top_m = suggested_links[0]
                        hint_txt = f"{top_m.get('party_label') or ''} {top_m.get('ref_no') or ''} ({top_m.get('pct') or 0}%)"
                        line = tr("scanner.link_hint_line", ref=hint_txt.strip())
                        details = f"{details}\n{line}" if details else line
                except Exception:
                    suggested_links = []

            item_id = f"tri_{secrets.token_hex(5)}"
            try:
                rel_in_inbox = src.relative_to(inbox_dir).as_posix()
            except ValueError:
                rel_in_inbox = src.name

            entry = {
                "id": item_id,
                "status": "pending",
                "file_name": src.name,
                "file_path": str(src),
                "rel_in_inbox": rel_in_inbox,
                "ext": ext,
                "page_count": int(pages or 1),
                "is_multipage": bool(ext == ".pdf" and int(pages or 1) >= 2),
                "ai_from_cache": ai_from_cache,
                "size_kb": round(src.stat().st_size / 1024, 1),
                "modified_at": dt.datetime.fromtimestamp(src.stat().st_mtime).isoformat(timespec="seconds"),
                "file_sha256": fhash,
                "section": section_id,
                "section_label": sec_label,
                "summary_headline": headline,
                "summary_details": details,
                "target_rel": target_rel,
                "suggested_links": suggested_links,
                "meta": norm_meta,
                "relevance": rel,
                "scanned_at": now_iso(),
            }
            with _TRIAGE_LOCK:
                cur_store = _load_triage_store(proj_root)
                if not any(x.get("file_sha256") == fhash for x in cur_store["items"]):
                    cur_store["items"].append(entry)
                    _save_triage_store(proj_root, cur_store)
                known_hashes.add(fhash)

            summary["added"] += 1
            summary["items"].append(entry)
        except Exception as exc:
            summary["errors"].append(f"{src.name}: {exc}")
        finally:
            shutil.rmtree(job_dir, ignore_errors=True)

    with _TRIAGE_LOCK:
        cur_store = _load_triage_store(proj_root)
        cur_store["last_scan_at"] = now_iso()
        cur_store["last_errors"] = list(summary["errors"])[-20:]
        _save_triage_store(proj_root, cur_store)

    return summary


def resolve_triage_item(
    proj: dict[str, Any],
    item_id: str,
    *,
    action: str,
    overrides: dict[str, Any] | None = None,
    device_id: str = "",
    decision_store: Any = None,
) -> dict[str, Any]:
    """
    Resolve a Scanner Triage item:
      - `approve`: moves ONLY this approved file from the scanner folder into `archive_dir`
        and records it in the project log.
      - `reject`: leaves the file completely untouched in the scanner folder and marks it rejected
        so future scans skip it unless user restores it.
      - `requeue`: restores a rejected item back into `pending` if the file still exists in scanner.
    """
    import template_profile as tprof

    proj_root = Path(str(proj["root"])).resolve()
    proj_id = str(proj.get("id") or "")
    act = str(action or "").strip().lower()
    if act not in ("approve", "reject", "requeue"):
        raise ValueError(tr("server.invalid_decision"))

    with _TRIAGE_LOCK:
        store = _load_triage_store(proj_root)
        idx = next((i for i, x in enumerate(store["items"]) if str(x.get("id")) == str(item_id)), None)
        if idx is None:
            raise KeyError(tr("server.decision_not_found"))
        item = dict(store["items"][idx])

    src = Path(str(item.get("file_path") or ""))
    fhash = str(item.get("file_sha256") or "")
    norm_src_path = str(normalize_win_path(str(src))).lower() if item.get("file_path") else ""

    if act == "reject":
        item["status"] = "rejected"
        item["resolved_at"] = now_iso()
        item["resolved_device"] = device_id
        with _TRIAGE_LOCK:
            store = _load_triage_store(proj_root)
            if fhash and fhash not in store["rejected_hashes"]:
                store["rejected_hashes"].append(fhash)
            if norm_src_path and norm_src_path not in store["rejected_paths"]:
                store["rejected_paths"].append(norm_src_path)
            for i, x in enumerate(store["items"]):
                if str(x.get("id")) == str(item_id):
                    store["items"][i] = item
                    break
            _save_triage_store(proj_root, store)
        if decision_store is not None:
            if fhash:
                decision_store.add_ignored_hash(fhash, src.name, reason="scanner_rejected")
            decision_store.resolve_by_file(
                file_sha256=fhash,
                file_path=str(src) if item.get("file_path") else "",
                project_id=proj_id,
                action="ignore_once",
                reason="scanner_rejected",
            )
        return {"ok": True, "action": "reject", "item": item}

    if act == "requeue":
        if not src.exists():
            raise FileNotFoundError(tr("scanner.file_missing_warn"))
        item["status"] = "pending"
        item.pop("resolved_at", None)
        with _TRIAGE_LOCK:
            store = _load_triage_store(proj_root)
            if fhash:
                store["rejected_hashes"] = [h for h in store["rejected_hashes"] if h != fhash]
                store["approved_hashes"] = [h for h in store["approved_hashes"] if h != fhash]
            if norm_src_path:
                store["rejected_paths"] = [p for p in store["rejected_paths"] if p != norm_src_path]
            for i, x in enumerate(store["items"]):
                if str(x.get("id")) == str(item_id):
                    store["items"][i] = item
                    break
            _save_triage_store(proj_root, store)
        if decision_store is not None and fhash:
            decision_store.remove_ignored_hash(fhash)
        return {"ok": True, "action": "requeue", "item": item}

    # action == "approve" -> verify file still exists in scanner, move ONLY it to archive_dir + log!
    if not src.exists() or not src.is_file():
        raise FileNotFoundError(tr("scanner.file_missing_warn"))

    ov = dict(overrides or {})
    section_id = str(ov.get("section") or item.get("section") or "general").strip()
    meta = dict(item.get("meta") or {})
    for k in (
        "subject_ar", "ref_no", "revision", "discipline", "party_name", "doc_date", "amount",
        "status", "notes", "authority", "reply_date",
        "consultant_reply_date", "consultant_action", "moh_reply_date", "moh_action",
    ):
        if k in ov and str(ov.get(k) if ov.get(k) is not None else "").strip() != "":
            meta[k] = ov[k]
    if "has_both_stamps" in ov:
        meta["has_both_stamps"] = bool(ov["has_both_stamps"])
    elif str(meta.get("authority") or "").upper() == "MOH+CONSULTANT" or (
        str(meta.get("consultant_reply_date") or "").strip() and str(meta.get("moh_reply_date") or "").strip()
    ):
        meta["has_both_stamps"] = True
    meta = _meta_to_section_meta(meta, section_id, src.name, proj)

    inbox_raw = str(proj.get("inbox_dir") or "").strip()
    inbox_dir = resolve_local_or_unc(inbox_raw) if inbox_raw else (proj_root / "1_Inbox").resolve()

    arch_res = archive_and_log_relevant_file(
        src,
        proj,
        meta,
        inbox_root=inbox_dir,
        device_id=device_id,
        copy_only=False,
        direct_move=True,
    )

    sec_map = tprof.sections_map_for_project(proj)
    item["section"] = section_id
    item["section_label"] = str((sec_map.get(section_id) or {}).get("label") or section_id)
    item["meta"] = meta
    item["status"] = "approved"
    item["archive_path"] = arch_res.get("archive_path")
    item["archive_rel"] = arch_res.get("archive_rel")
    item["log_result"] = arch_res.get("log_result")
    item["resolved_at"] = now_iso()
    item["resolved_device"] = device_id

    link_targets = ov.get("link_to_refs")
    if isinstance(link_targets, list) and link_targets and section_id in ("incoming", "outgoing"):
        try:
            import letter_links as llinks
            lres = arch_res.get("log_result") if isinstance(arch_res.get("log_result"), dict) else {}
            new_letter_item = {
                "ref_no": str(lres.get("ref_no") or meta.get("ref_no") or ""),
                "sheet": str(lres.get("sheet") or ""),
                "row": lres.get("row"),
                "log_name": str(lres.get("log_name") or ""),
                "section": section_id,
                "party_name": str(meta.get("party_name") or ""),
                "subject": str(meta.get("subject_ar") or meta.get("subject") or ""),
                "doc_date": str(meta.get("doc_date") or ""),
            }
            item["link_result"] = llinks.link_letters(
                proj,
                [new_letter_item, *link_targets],
                sync_excel=True,
                device_id=device_id,
            )
        except Exception:
            pass

    with _TRIAGE_LOCK:
        store = _load_triage_store(proj_root)
        if fhash and fhash not in store["approved_hashes"]:
            store["approved_hashes"].append(fhash)
        if norm_src_path and norm_src_path not in store["approved_paths"]:
            store["approved_paths"].append(norm_src_path)
        for i, x in enumerate(store["items"]):
            if str(x.get("id")) == str(item_id):
                store["items"][i] = item
                break
        _save_triage_store(proj_root, store)

    if decision_store is not None:
        decision_store.resolve_by_file(
            file_sha256=fhash,
            file_path=str(src) if item.get("file_path") else "",
            project_id=proj_id,
            action="attach_project",
            reason="scanner_approved",
        )

    return {"ok": True, "action": "approve", "item": item, "archive": arch_res}


def resolve_triage_bulk(
    proj: dict[str, Any],
    *,
    action: str,
    item_ids: list[str] | None = None,
    device_id: str = "",
    decision_store: Any = None,
) -> dict[str, Any]:
    """Approve or reject all pending triage items (or a specified subset of IDs)."""
    proj_root = Path(str(proj["root"])).resolve()
    store = _load_triage_store(proj_root)
    target_set = {str(x) for x in item_ids} if item_ids else None
    targets = [
        str(it.get("id"))
        for it in store["items"]
        if it.get("status", "pending") == "pending" and (target_set is None or str(it.get("id")) in target_set)
    ]

    succeeded = 0
    errors: list[str] = []
    results: list[dict[str, Any]] = []
    for tid in targets:
        try:
            r = resolve_triage_item(proj, tid, action=action, device_id=device_id, decision_store=decision_store)
            results.append(r)
            succeeded += 1
        except Exception as exc:
            errors.append(f"{tid}: {exc}")

    return {
        "ok": True,
        "action": action,
        "count": succeeded,
        "errors": errors,
        "results": results,
    }


def clear_scanner_resolved_items(
    proj: dict[str, Any],
    *,
    bucket: str = "all",
) -> dict[str, Any]:
    """
    Hide resolved (`approved` / `rejected`) items from the Scanner UI while keeping
    their hashes & paths recorded so future scans never re-surface them.
    """
    proj_root = Path(str(proj["root"])).resolve()
    b = str(bucket or "all").strip().lower()
    remove_statuses = {"approved", "rejected"} if b == "all" else {b}
    with _TRIAGE_LOCK:
        store = _load_triage_store(proj_root)
        kept: list[dict[str, Any]] = []
        removed = 0
        for it in store["items"]:
            st = str(it.get("status") or "pending")
            if st in remove_statuses:
                fh = str(it.get("file_sha256") or "")
                np = str(normalize_win_path(str(it.get("file_path") or ""))).lower() if it.get("file_path") else ""
                if st == "rejected":
                    if fh and fh not in store["rejected_hashes"]:
                        store["rejected_hashes"].append(fh)
                    if np and np not in store["rejected_paths"]:
                        store["rejected_paths"].append(np)
                elif st == "approved":
                    if fh and fh not in store["approved_hashes"]:
                        store["approved_hashes"].append(fh)
                    if np and np not in store["approved_paths"]:
                        store["approved_paths"].append(np)
                removed += 1
            else:
                kept.append(it)
        store["items"] = kept
        _save_triage_store(proj_root, store)
    return {"ok": True, "removed": removed, "bucket": b}


def sync_decision_to_triage(
    proj: dict[str, Any],
    *,
    file_sha256: str = "",
    file_path: str = "",
    action: str = "ignore_once",
) -> int:
    """
    When a file is resolved in Pending Decisions, remove/hide matching item(s)
    from this project's Scanner Triage pending queue so it disappears everywhere.
    """
    if not file_sha256 and not file_path:
        return 0
    proj_root = Path(str(proj["root"])).resolve()
    norm_p = str(normalize_win_path(file_path)).lower() if file_path else ""
    new_status = "rejected" if str(action).startswith("ignore") else "approved"
    updated = 0
    with _TRIAGE_LOCK:
        store = _load_triage_store(proj_root)
        changed = False
        for it in store["items"]:
            if it.get("status", "pending") != "pending":
                continue
            ih = str(it.get("file_sha256") or "")
            ip = str(normalize_win_path(str(it.get("file_path") or ""))).lower() if it.get("file_path") else ""
            if (file_sha256 and ih == file_sha256) or (norm_p and ip == norm_p):
                it["status"] = new_status
                it["resolved_at"] = now_iso()
                it["resolved_device"] = "decisions_sync"
                if new_status == "rejected":
                    if ih and ih not in store["rejected_hashes"]:
                        store["rejected_hashes"].append(ih)
                    if ip and ip not in store["rejected_paths"]:
                        store["rejected_paths"].append(ip)
                else:
                    if ih and ih not in store["approved_hashes"]:
                        store["approved_hashes"].append(ih)
                    if ip and ip not in store["approved_paths"]:
                        store["approved_paths"].append(ip)
                updated += 1
                changed = True
        if file_sha256:
            target_list = store["rejected_hashes"] if new_status == "rejected" else store["approved_hashes"]
            if file_sha256 not in target_list:
                target_list.append(file_sha256)
                changed = True
        if changed:
            _save_triage_store(proj_root, store)
    return updated


def _find_triage_item(proj_root: Path, item_id: str) -> dict[str, Any]:
    store = _load_triage_store(proj_root)
    for it in store.get("items") or []:
        if str(it.get("id") or "") == str(item_id):
            return dict(it)
    raise KeyError(tr("server.decision_not_found"))


def _require_fitz():
    try:
        import fitz  # PyMuPDF
        return fitz
    except ImportError as exc:
        raise RuntimeError(tr("watcher.pymupdf_missing")) from exc


IMAGE_PREVIEW_EXTS: frozenset[str] = frozenset({
    ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff",
})


def render_document_preview_page(
    src_path: Path | str,
    page_no: int = 1,
    *,
    dpi: int = 120,
    cache_dir: Path | None = None,
    cache_key: str = "",
) -> tuple[bytes, int]:
    """
    Render page `page_no` of a PDF or image file to PNG bytes plus total page count.
    Caches rendered pages safely when `cache_dir` is provided.
    """
    import io

    src = resolve_local_or_unc(str(src_path))
    if not src.exists() or not src.is_file():
        raise FileNotFoundError(tr("scanner.file_missing_warn"))

    ext = src.suffix.lower()
    dpi_c = max(60, min(int(dpi or 120), 220))
    req_page = max(1, int(page_no or 1))
    safe_key = re.sub(r"[^A-Za-z0-9_-]", "_", cache_key or src.stem)[:80] or "doc"

    if cache_dir is not None:
        cache_dir.mkdir(parents=True, exist_ok=True)
        try:
            src_mtime = src.stat().st_mtime
            for existing in cache_dir.glob(f"{safe_key}_p{req_page}_d{dpi_c}_n*.png"):
                if existing.stat().st_mtime >= src_mtime:
                    m_tot = re.search(r"_n(\d+)\.png$", existing.name)
                    if m_tot:
                        return existing.read_bytes(), max(1, int(m_tot.group(1)))
        except OSError:
            pass

    if ext in IMAGE_PREVIEW_EXTS:
        from PIL import Image, ImageOps

        with Image.open(src) as im:
            im = ImageOps.exif_transpose(im)
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGB")
            max_w = 1600 if dpi_c >= 110 else 640
            if im.width > max_w:
                ratio = max_w / float(im.width)
                im = im.resize((max_w, max(1, int(im.height * ratio))), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            png_bytes = buf.getvalue()
        total = 1
    elif ext == ".pdf":
        fitz = _require_fitz()
        with fitz.open(src) as doc:
            total = int(doc.page_count or 0)
            if total < 1 or req_page > total:
                raise ValueError(tr("scanner.split_bad_range"))
            pix = doc[req_page - 1].get_pixmap(dpi=dpi_c)
            png_bytes = pix.tobytes("png")
    else:
        raise ValueError(tr("preview.unsupported_ext"))

    if cache_dir is not None:
        try:
            out_p = cache_dir / f"{safe_key}_p{req_page}_d{dpi_c}_n{total}.png"
            out_p.write_bytes(png_bytes)
        except OSError:
            pass
    return png_bytes, total


def render_triage_pdf_page(
    proj: dict[str, Any],
    item_id: str,
    page_no: int,
    *,
    dpi: int = 110,
) -> bytes:
    """Render one page of a pending/resolved triage PDF or image to cached PNG bytes."""
    proj_root = resolve_local_or_unc(str(proj["root"]))
    item = _find_triage_item(proj_root, item_id)
    src = resolve_local_or_unc(str(item.get("file_path") or "")) if item.get("file_path") else Path("")
    cache_dir = proj_root / "_System_Agent" / "_tmp_split_thumbs"
    png_bytes, _ = render_document_preview_page(
        src, page_no, dpi=dpi, cache_dir=cache_dir, cache_key=str(item_id),
    )
    return png_bytes


def build_multipage_split_prompt(
    *,
    file_name: str,
    page_count: int,
    page_texts: list[str],
    proj: dict[str, Any],
) -> str:
    """Build the AI prompt for deciding if a multi-page PDF is 1 document or multiple documents."""
    import template_profile as tprof

    sec_map = tprof.sections_map_for_project(proj)
    sec_choices = ", ".join(f"{sid} ({(s.get('label') or sid)})" for sid, s in sec_map.items())
    page_blocks: list[str] = []
    for idx in range(1, page_count + 1):
        snippet = " ".join((page_texts[idx - 1] if idx - 1 < len(page_texts) else "").split())[:550]
        page_blocks.append(f"--- PAGE {idx} ---\n{snippet or '[Scanned page image — read image]'}")
    pages_joined = "\n".join(page_blocks)

    return (
        f"You are analyzing a {page_count}-page PDF file named '{file_name}' from a scanner inbox "
        f"for project '{proj.get('name') or ''}'.\n"
        f"Decide whether all {page_count} pages belong to ONE single continuous document "
        f"OR whether multiple separate documents were scanned together into one PDF bundle.\n"
        f"Look for clues on each page: new letterheads, 'Page 1 of', new Reference/Ref numbers, "
        f"new dates, new recipient greetings ('السادة', 'Dear'), or closing signatures/stamps at the end of one doc "
        f"followed by a fresh header on the next page.\n"
        f"Valid section IDs: {sec_choices}.\n"
        f"Return STRICT JSON ONLY with NO markdown wrapper:\n"
        f'{{"is_single_document": false, "documents": ['
        f'{{"start_page": 1, "end_page": 2, "title": "short Arabic title", "section": "incoming", '
        f'"ref_no": "", "party_name": "", "doc_date": "YYYY-MM-DD", "reason": "short Arabic reason"}}'
        f']}}\n'
        f"Every page from 1 to {page_count} MUST be covered in contiguous non-overlapping ascending ranges.\n\n"
        f"{pages_joined}"
    )


def _heuristic_split_pages(
    doc: Any,
    page_texts: list[str],
    proj: dict[str, Any],
    base_meta: dict[str, Any],
) -> dict[str, Any]:
    """Deterministic fallback page-boundary analyzer using text & layout cues (no API key required)."""
    import scan_watcher as sw

    n_pages = len(page_texts)
    if n_pages <= 1:
        return {
            "is_single_document": True,
            "documents": [{
                "start_page": 1,
                "end_page": max(1, n_pages),
                "title": str(base_meta.get("subject_ar") or base_meta.get("subject") or ""),
                "section": _infer_section_from_meta(base_meta, "", page_texts[0] if page_texts else "", proj),
                "ref_no": str(base_meta.get("ref_no") or ""),
                "party_name": str(base_meta.get("party_name") or ""),
                "doc_date": str(base_meta.get("doc_date") or ""),
                "reason": tr("scanner.split_single_hint"),
            }],
        }

    ref_pat = re.compile(
        r"(?:Ref(?:erence)?(?:\s*No\.?)?|Our\s*Ref|رقم(?:\s*الكتاب|\s*الإشارة)?|إشارة|مرجع)\s*[:#\-]?\s*([A-Za-z0-9\u0600-\u06FF/_-]{3,28})",
        re.I,
    )
    date_pat = re.compile(r"\b(20\d{2}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]20\d{2})\b")
    page1_pat = re.compile(r"(?:page\s*1\s*(?:of|/)\s*\d+|صفحة\s*(?:1|١)\s*(?:من|/)\s*\d+|\b1\s*/\s*[2-9]\b)", re.I)
    hdr_pat = re.compile(r"(?:السادة\s*[/\-:]|السيد\s*[/\-:]|الموضوع\s*:|subject\s*:|dear\s+(?:sir|madam)|to\s*:)", re.I)
    close_pat = re.compile(r"(?:وتفضلوا\s+بقبول|شاكرين\s+لكم|yours\s+faithfully|sincerely|best\s+regards|مدير\s+المشروع)", re.I)

    page_info: list[dict[str, Any]] = []
    for _i, txt in enumerate(page_texts):
        head = txt[:650]
        m_ref = ref_pat.search(head)
        ref_val = m_ref.group(1).strip() if m_ref else ""
        # Avoid false matches like "No" or digits only < 3
        if ref_val.lower() in ("no", "number", "date", "ref"):
            ref_val = ""
        m_dt = date_pat.search(head)
        dt_val = ""
        if m_dt:
            raw_d = m_dt.group(1)
            if hasattr(sw, "normalize_date_str"):
                dt_val = sw.normalize_date_str(raw_d)
            else:
                m_iso = re.match(r"^(20\d{2})[-/](\d{1,2})[-/](\d{1,2})$", raw_d)
                m_dmy = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](20\d{2})$", raw_d)
                if m_iso:
                    dt_val = f"{m_iso.group(1)}-{int(m_iso.group(2)):02d}-{int(m_iso.group(3)):02d}"
                elif m_dmy:
                    dt_val = f"{m_dmy.group(3)}-{int(m_dmy.group(2)):02d}-{int(m_dmy.group(1)):02d}"
                else:
                    dt_val = raw_d
        page_info.append({
            "ref": ref_val,
            "date": dt_val,
            "has_p1": bool(page1_pat.search(txt[:900] + "\n" + txt[-350:])),
            "has_hdr": bool(hdr_pat.search(head)),
            "has_close": bool(close_pat.search(txt[-650:])),
            "text_len": len(txt.strip()),
        })

    boundaries: list[tuple[int, str]] = [(1, "")]
    for idx in range(1, n_pages):
        cur = page_info[idx]
        prev = page_info[idx - 1]
        reasons: list[str] = []
        if cur["has_p1"]:
            reasons.append("ترقيم صفحة 1 جديد")
        if cur["ref"] and prev["ref"] and normalize_ar(cur["ref"]) != normalize_ar(prev["ref"]):
            reasons.append(f"مرجع مختلف ({cur['ref']})")
        elif cur["ref"] and prev["has_close"] and cur["has_hdr"]:
            reasons.append(f"مستند جديد ({cur['ref']})")
        elif prev["has_close"] and cur["has_hdr"]:
            reasons.append("ختام صفحة سابقة وترويسة جديدة")
        if reasons:
            boundaries.append((idx + 1, "، ".join(reasons)))

    docs_out: list[dict[str, Any]] = []
    for b_i, (sp, rsn) in enumerate(boundaries):
        ep = (boundaries[b_i + 1][0] - 1) if b_i + 1 < len(boundaries) else n_pages
        chunk_txt = "\n".join(page_texts[sp - 1 : ep])
        first_txt = page_texts[sp - 1] if sp - 1 < len(page_texts) else ""
        first_lines = [ln.strip() for ln in first_txt.splitlines() if len(ln.strip()) >= 6]
        title_guess = ""
        for ln in first_lines[:8]:
            if any(k in ln.lower() for k in ("الموضوع", "subject", "بشأن", "خطاب", "كتاب", "فاتورة", "مطالبة", "تقرير", "اعتماد")):
                title_guess = ln[:85]
                break
        if not title_guess and first_lines:
            title_guess = first_lines[0][:85]
        if not title_guess and len(boundaries) == 1:
            title_guess = str(base_meta.get("subject_ar") or base_meta.get("subject") or "")
        sec_guess = _infer_section_from_meta(base_meta if len(boundaries) == 1 else {}, title_guess, chunk_txt, proj)
        docs_out.append({
            "start_page": sp,
            "end_page": ep,
            "title": title_guess,
            "section": sec_guess,
            "ref_no": page_info[sp - 1]["ref"] or (str(base_meta.get("ref_no") or "") if len(boundaries) == 1 else ""),
            "party_name": str(base_meta.get("party_name") or "") if len(boundaries) == 1 else "",
            "doc_date": page_info[sp - 1]["date"] or (str(base_meta.get("doc_date") or "") if len(boundaries) == 1 else ""),
            "reason": rsn or tr("scanner.split_single_hint"),
        })

    return {
        "is_single_document": len(docs_out) <= 1,
        "documents": docs_out,
    }


def _normalize_split_ranges(
    raw_docs: list[dict[str, Any]],
    page_count: int,
    proj: dict[str, Any],
    base_meta: dict[str, Any],
    fallback_name: str,
) -> list[dict[str, Any]]:
    import template_profile as tprof

    sec_map = tprof.sections_map_for_project(proj)
    valid_secs = set(sec_map.keys())
    cleaned: list[dict[str, Any]] = []
    for d in raw_docs or []:
        if not isinstance(d, dict):
            continue
        try:
            sp = int(d.get("start_page") or 0)
            ep = int(d.get("end_page") or 0)
        except (TypeError, ValueError):
            continue
        sp = max(1, min(sp, page_count))
        ep = max(sp, min(ep, page_count))
        sec_v = str(d.get("section") or base_meta.get("section") or "general").strip()
        if sec_v not in valid_secs:
            sec_v = str(base_meta.get("section") or "general")
            if sec_v not in valid_secs:
                sec_v = "general"
        cleaned.append({
            "start_page": sp,
            "end_page": ep,
            "title": str(d.get("title") or "").strip(),
            "section": sec_v,
            "section_label": str((sec_map.get(sec_v) or {}).get("label") or sec_v),
            "ref_no": str(d.get("ref_no") or "").strip(),
            "party_name": str(d.get("party_name") or "").strip(),
            "doc_date": str(d.get("doc_date") or "").strip(),
            "reason": str(d.get("reason") or "").strip(),
        })

    if not cleaned:
        sec_v = str(base_meta.get("section") or "general")
        if sec_v not in valid_secs:
            sec_v = "general"
        cleaned = [{
            "start_page": 1,
            "end_page": page_count,
            "title": str(base_meta.get("subject_ar") or base_meta.get("subject") or Path(fallback_name).stem),
            "section": sec_v,
            "section_label": str((sec_map.get(sec_v) or {}).get("label") or sec_v),
            "ref_no": str(base_meta.get("ref_no") or ""),
            "party_name": str(base_meta.get("party_name") or ""),
            "doc_date": str(base_meta.get("doc_date") or ""),
            "reason": tr("scanner.split_single_hint"),
        }]

    cleaned.sort(key=lambda x: (x["start_page"], x["end_page"]))
    normalized: list[dict[str, Any]] = []
    cursor = 1
    for item in cleaned:
        sp = max(cursor, item["start_page"])
        if sp > page_count:
            break
        if sp > cursor and normalized:
            normalized[-1]["end_page"] = sp - 1
        elif sp > cursor:
            sp = cursor
        ep = max(sp, min(item["end_page"], page_count))
        row_copy = dict(item)
        row_copy["start_page"] = sp
        row_copy["end_page"] = ep
        if not row_copy["title"]:
            row_copy["title"] = f"{Path(fallback_name).stem} (p.{sp}-{ep})"
        normalized.append(row_copy)
        cursor = ep + 1

    if normalized and cursor <= page_count:
        normalized[-1]["end_page"] = page_count

    return normalized


def analyze_multipage_pdf_split(
    proj: dict[str, Any],
    item_id: str,
    *,
    global_settings: dict[str, Any] | None = None,
    analyzer_override: Any = None,
    force_refresh: bool = False,
) -> dict[str, Any]:
    """
    Analyze all pages of a multi-page triage PDF using AI Vision/Text (or deterministic fallback)
    and return proposed contiguous document split ranges (cached by file SHA-256 + project sections).
    """
    fitz = _require_fitz()
    import scan_watcher as sw
    import template_profile as tprof

    proj_root = resolve_local_or_unc(str(proj["root"]))
    item = _find_triage_item(proj_root, item_id)
    src = resolve_local_or_unc(str(item.get("file_path") or "")) if item.get("file_path") else Path("")
    if not src.exists() or not src.is_file():
        raise FileNotFoundError(tr("scanner.file_missing_warn"))
    if src.suffix.lower() != ".pdf":
        raise ValueError(tr("server.invalid_setting"))

    file_sha = sha256_of_file(src)
    sec_sig = ai_cache.stable_sig(sorted(tprof.sections_map_for_project(proj).keys()))
    cache_dir = ai_cache.project_ai_cache_dir(proj_root)
    split_cache_key = f"{file_sha}:{sec_sig}"
    if not force_refresh:
        cached_split = ai_cache.get_cached(cache_dir, "pdf_split", split_cache_key)
        if (
            cached_split
            and cached_split.get("mode") == "ai"
            and isinstance(cached_split.get("documents"), list)
            and cached_split["documents"]
        ):
            p_cnt_cached = int(cached_split.get("page_count") or item.get("page_count") or 1)
            docs_cached = _normalize_split_ranges(
                list(cached_split["documents"]),
                p_cnt_cached,
                proj,
                dict(item.get("meta") or {}, section=item.get("section") or "general"),
                src.name,
            )
            return {
                "ok": True,
                "item_id": item_id,
                "file_name": src.name,
                "page_count": p_cnt_cached,
                "mode": "ai",
                "from_cache": True,
                "cached_at": str(cached_split.get("_cached_at") or ""),
                "warning": "",
                "is_single_document": len(docs_cached) <= 1,
                "documents": docs_cached,
            }

    gcfg = global_settings or {}
    engine = str(gcfg.get("engine") or proj.get("engine") or "gemini").lower()
    ollama_url = str(gcfg.get("ollama_url") or "http://127.0.0.1:11434").strip()
    if ollama_url and "/api/" not in ollama_url:
        ollama_url = f"{ollama_url.rstrip('/')}/api/chat"
    ollama_model = str(gcfg.get("ollama_model") or "qwen2.5vl:7b")

    page_texts: list[str] = []
    with fitz.open(src) as doc:
        page_count = int(doc.page_count or 1)
        for p_idx in range(page_count):
            page_texts.append(doc[p_idx].get_text("text") or "")

        ai_res: dict[str, Any] | None = None
        mode = "heuristic"
        warn_msg = ""
        prompt = build_multipage_split_prompt(
            file_name=src.name,
            page_count=page_count,
            page_texts=page_texts,
            proj=proj,
        )

        # Render thumbnails only for AI Vision when pages are scanned (little embedded text) and <= 16 pages
        vis_imgs: list[Path] = []
        tmp_dir = proj_root / "_System_Agent" / "_tmp_split_ai" / re.sub(r"[^A-Za-z0-9_-]", "_", str(item_id))
        has_sparse_text = any(len(t.strip()) < 40 for t in page_texts[:16])
        if (analyzer_override is not None or (engine == "ollama" or sw.load_saved_gemini_key())) and has_sparse_text and page_count <= 16:
            try:
                tmp_dir.mkdir(parents=True, exist_ok=True)
                for p_idx in range(page_count):
                    im_p = tmp_dir / f"page_{p_idx + 1:03d}.png"
                    doc[p_idx].get_pixmap(dpi=100).save(str(im_p))
                    vis_imgs.append(im_p)
            except Exception:
                vis_imgs = []

        try:
            if analyzer_override is not None:
                ai_res = analyzer_override(src, page_texts, prompt)
                if isinstance(ai_res, dict) and ai_res.get("documents"):
                    mode = "ai"
            elif engine == "ollama":
                try:
                    ai_res = sw.analyze_with_ollama(prompt, vis_imgs, ollama_url, ollama_model)
                    if isinstance(ai_res, dict) and ai_res.get("documents"):
                        mode = "ai"
                except Exception as exc_ai:
                    warn_msg = str(exc_ai)
            elif sw.load_saved_gemini_key():
                try:
                    ai_res = sw.analyze_with_gemini(
                        prompt,
                        tmp_dir if tmp_dir.exists() else (proj_root / "_System_Agent"),
                        sw.load_saved_gemini_model(),
                        images=vis_imgs,
                    )
                    if isinstance(ai_res, dict) and ai_res.get("documents"):
                        mode = "ai"
                except Exception as exc_ai:
                    warn_msg = str(exc_ai)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        base_meta = dict(item.get("meta") or {})
        base_meta.setdefault("section", item.get("section") or "general")
        if mode != "ai" or not isinstance(ai_res, dict) or not ai_res.get("documents"):
            ai_res = _heuristic_split_pages(doc, page_texts, proj, base_meta)
            mode = "heuristic"

    docs_norm = _normalize_split_ranges(
        list(ai_res.get("documents") or []),
        page_count,
        proj,
        dict(item.get("meta") or {}, section=item.get("section") or "general"),
        src.name,
    )
    if mode == "ai" and docs_norm:
        ai_cache.set_cached(cache_dir, "pdf_split", split_cache_key, {
            "mode": "ai",
            "page_count": page_count,
            "documents": docs_norm,
        })
    return {
        "ok": True,
        "item_id": item_id,
        "file_name": src.name,
        "page_count": page_count,
        "mode": mode,
        "from_cache": False,
        "cached_at": now_iso() if mode == "ai" else "",
        "warning": warn_msg,
        "is_single_document": len(docs_norm) <= 1,
        "documents": docs_norm,
    }


def _safe_split_part_slug(text: str, max_len: int = 32) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", str(text or "")).strip(" ._")
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:max_len].strip(" ._") or "doc"


def apply_multipage_pdf_split(
    proj: dict[str, Any],
    item_id: str,
    ranges: list[dict[str, Any]],
    *,
    auto_approve: bool = False,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Split a multi-page triage PDF according to user-confirmed `ranges`
    (`[{start_page, end_page, title, section, ref_no, party_name, doc_date}, ...]`).
    - If a single range covers all pages (`1..page_count`), keeps it as one document
      (or approves directly when `auto_approve=True`).
    - Otherwise writes each part into `inbox_dir` (falling back to `_System_Agent/split_parts`
      if the scanner share is read-only), backs up the original bundle safely, marks the original
      triage entry resolved (`was_split=True`), registers each split part in triage, and
      optionally archives & logs all parts in one shot (`auto_approve=True`).
    """
    fitz = _require_fitz()
    import template_profile as tprof

    proj_root = resolve_local_or_unc(str(proj["root"]))
    sys_root = proj_root / "_System_Agent"
    sys_root.mkdir(parents=True, exist_ok=True)
    item = _find_triage_item(proj_root, item_id)
    src = resolve_local_or_unc(str(item.get("file_path") or "")) if item.get("file_path") else Path("")
    if not src.exists() or not src.is_file():
        raise FileNotFoundError(tr("scanner.file_missing_warn"))
    if src.suffix.lower() != ".pdf":
        raise ValueError(tr("server.invalid_setting"))

    with fitz.open(src) as doc:
        total_pages = int(doc.page_count or 1)

    if not isinstance(ranges, list) or not ranges:
        raise ValueError(tr("scanner.split_bad_range"))

    parsed_ranges: list[dict[str, Any]] = []
    used_pages: set[int] = set()
    for rg in ranges:
        if not isinstance(rg, dict):
            raise ValueError(tr("scanner.split_bad_range"))
        try:
            sp = int(rg.get("start_page") or 0)
            ep = int(rg.get("end_page") or 0)
        except (TypeError, ValueError):
            raise ValueError(tr("scanner.split_bad_range"))
        if sp < 1 or ep < sp or ep > total_pages:
            raise ValueError(tr("scanner.split_bad_range"))
        for p in range(sp, ep + 1):
            if p in used_pages:
                raise ValueError(tr("scanner.split_bad_range"))
            used_pages.add(p)
        parsed_ranges.append({
            "start_page": sp,
            "end_page": ep,
            "title": str(rg.get("title") or "").strip(),
            "section": str(rg.get("section") or item.get("section") or "general").strip(),
            "ref_no": str(rg.get("ref_no") or "").strip(),
            "party_name": str(rg.get("party_name") or "").strip(),
            "doc_date": str(rg.get("doc_date") or "").strip(),
        })

    # Keep as single document if user chose 1 full-document range
    if len(parsed_ranges) == 1 and parsed_ranges[0]["start_page"] == 1 and parsed_ranges[0]["end_page"] == total_pages:
        r0 = parsed_ranges[0]
        if auto_approve:
            ov = {
                "section": r0["section"],
                "subject_ar": r0["title"] or (item.get("meta") or {}).get("subject_ar") or "",
                "ref_no": r0["ref_no"] or (item.get("meta") or {}).get("ref_no") or "",
                "party_name": r0["party_name"] or (item.get("meta") or {}).get("party_name") or "",
                "doc_date": r0["doc_date"] or (item.get("meta") or {}).get("doc_date") or "",
            }
            res = resolve_triage_item(proj, item_id, action="approve", overrides=ov, device_id=device_id)
            return {
                "ok": True,
                "single_kept": True,
                "auto_approved": True,
                "split_count": 1,
                "results": [res],
            }
        # Update user edits on pending item without splitting
        sec_map = tprof.sections_map_for_project(proj)
        meta = dict(item.get("meta") or {})
        if r0["title"]:
            meta["subject_ar"] = r0["title"]
            meta["subject"] = r0["title"]
        if r0["ref_no"]:
            meta["ref_no"] = r0["ref_no"]
        if r0["party_name"]:
            meta["party_name"] = r0["party_name"]
        if r0["doc_date"]:
            meta["doc_date"] = r0["doc_date"]
        with _TRIAGE_LOCK:
            store = _load_triage_store(proj_root)
            for i, x in enumerate(store["items"]):
                if str(x.get("id")) == str(item_id):
                    store["items"][i]["section"] = r0["section"]
                    store["items"][i]["section_label"] = str((sec_map.get(r0["section"]) or {}).get("label") or r0["section"])
                    store["items"][i]["meta"] = meta
                    item = dict(store["items"][i])
                    _save_triage_store(proj_root, store)
                    break
        return {"ok": True, "single_kept": True, "auto_approved": False, "split_count": 1, "item": item}

    inbox_raw = str(proj.get("inbox_dir") or "").strip()
    inbox_dir = resolve_local_or_unc(inbox_raw) if inbox_raw else (proj_root / "1_Inbox").resolve()
    fallback_dir = sys_root / "split_parts" / re.sub(r"[^A-Za-z0-9_-]", "_", str(item_id))

    sec_map = tprof.sections_map_for_project(proj)
    created_children: list[dict[str, Any]] = []
    created_part_paths: list[Path] = []
    try:
        with fitz.open(src) as src_doc:
            for idx, rg in enumerate(parsed_ranges, 1):
                sp, ep = rg["start_page"], rg["end_page"]
                part_doc = fitz.open()
                part_txt_parts: list[str] = []
                try:
                    part_doc.insert_pdf(src_doc, from_page=sp - 1, to_page=ep - 1)
                    for p_i in range(sp - 1, ep):
                        part_txt_parts.append(src_doc[p_i].get_text("text") or "")

                    slug = _safe_split_part_slug(rg["title"] or rg["ref_no"] or f"p{sp}_{ep}")
                    base_name = f"{src.stem}_part{idx:02d}_{slug}.pdf"
                    target_dir = inbox_dir if (inbox_dir.exists() and inbox_dir.is_dir()) else fallback_dir
                    target_dir.mkdir(parents=True, exist_ok=True)
                    out_path = target_dir / base_name
                    c_n = 1
                    while out_path.exists():
                        c_n += 1
                        out_path = target_dir / f"{src.stem}_part{idx:02d}_{slug}_{c_n}.pdf"
                    try:
                        part_doc.save(str(out_path))
                    except (PermissionError, OSError):
                        fallback_dir.mkdir(parents=True, exist_ok=True)
                        out_path = fallback_dir / base_name
                        c_n = 1
                        while out_path.exists():
                            c_n += 1
                            out_path = fallback_dir / f"{src.stem}_part{idx:02d}_{slug}_{c_n}.pdf"
                        part_doc.save(str(out_path))
                finally:
                    part_doc.close()

                if not out_path.exists() or out_path.stat().st_size == 0:
                    out_path.unlink(missing_ok=True)
                    raise RuntimeError(tr("organizer.pdf_empty_err"))
                created_part_paths.append(out_path)

                part_text = "\n".join(part_txt_parts)
                sec_id = rg["section"] if rg["section"] in sec_map else str(item.get("section") or "general")
                base_meta = dict(item.get("meta") or {})
                child_raw_meta = {
                    "subject_ar": rg["title"] or base_meta.get("subject_ar") or out_path.stem,
                    "subject": rg["title"] or base_meta.get("subject") or out_path.stem,
                    "ref_no": rg["ref_no"],
                    "party_name": rg["party_name"] or base_meta.get("party_name") or "",
                    "doc_date": rg["doc_date"] or base_meta.get("doc_date") or "",
                    "discipline": base_meta.get("discipline") or "",
                    "status": base_meta.get("status") or "",
                }
                norm_meta = _meta_to_section_meta(child_raw_meta, sec_id, out_path.name, proj)
                rel = dict(item.get("relevance") or {"relevant": True, "score": 1.0, "reason": "split_part"})
                headline, details = _build_human_summary_ar(out_path.name, part_text, norm_meta, sec_id, rel, proj)
                target_rel = _preview_target_rel(norm_meta, out_path, proj, sec_id)
                p_cnt = ep - sp + 1
                fhash = sha256_of_file(out_path)
                child_links: list[dict[str, Any]] = []
                if sec_id in ("incoming", "outgoing"):
                    try:
                        import letter_links as llinks
                        child_links = llinks.suggest_letter_links(
                            proj,
                            ref_no=str(norm_meta.get("ref_no") or ""),
                            section=sec_id,
                            subject=str(norm_meta.get("subject_ar") or norm_meta.get("subject") or ""),
                            doc_date=str(norm_meta.get("doc_date") or ""),
                            text=part_text[:2500],
                            party_name=str(norm_meta.get("party_name") or ""),
                            limit=5,
                            min_score=0.48,
                            device_id=device_id,
                        )
                    except Exception:
                        child_links = []
                try:
                    rel_in_inbox = out_path.relative_to(inbox_dir).as_posix()
                except ValueError:
                    rel_in_inbox = out_path.name

                child_entry = {
                    "id": f"tri_{secrets.token_hex(5)}",
                    "status": "pending",
                    "file_name": out_path.name,
                    "file_path": str(out_path),
                    "rel_in_inbox": rel_in_inbox,
                    "ext": ".pdf",
                    "page_count": p_cnt,
                    "is_multipage": p_cnt >= 2,
                    "parent_split_id": item_id,
                    "split_pages": f"{sp}-{ep}",
                    "size_kb": round(out_path.stat().st_size / 1024, 1),
                    "modified_at": dt.datetime.fromtimestamp(out_path.stat().st_mtime).isoformat(timespec="seconds"),
                    "file_sha256": fhash,
                    "section": sec_id,
                    "section_label": str((sec_map.get(sec_id) or {}).get("label") or sec_id),
                    "summary_headline": headline,
                    "summary_details": details,
                    "target_rel": target_rel,
                    "suggested_links": child_links,
                    "meta": norm_meta,
                    "relevance": rel,
                    "scanned_at": now_iso(),
                }
                created_children.append(child_entry)
    except Exception:
        for cp in created_part_paths:
            cp.unlink(missing_ok=True)
        raise

    # Backup original multi-page bundle safely so it isn't re-scanned or lost (graceful on read-only share)
    orig_backup_dir = sys_root / "_01_Processed_Originals"
    try:
        orig_backup_dir.mkdir(parents=True, exist_ok=True)
        bk_dest = orig_backup_dir / src.name
        c_bk = 1
        while bk_dest.exists():
            c_bk += 1
            bk_dest = orig_backup_dir / f"{src.stem}_dup{c_bk}{src.suffix}"
        shutil.move(str(src), str(bk_dest))
    except (PermissionError, OSError):
        pass

    with _TRIAGE_LOCK:
        store = _load_triage_store(proj_root)
        for i, x in enumerate(store["items"]):
            if str(x.get("id")) == str(item_id):
                store["items"][i]["status"] = "approved"
                store["items"][i]["was_split"] = True
                store["items"][i]["split_children"] = [c["id"] for c in created_children]
                store["items"][i]["resolved_at"] = now_iso()
                store["items"][i]["resolved_device"] = device_id
                break
        store["items"].extend(created_children)
        _save_triage_store(proj_root, store)

    approved_results: list[dict[str, Any]] = []
    errors: list[str] = []
    if auto_approve:
        final_children: list[dict[str, Any]] = []
        for ch in created_children:
            try:
                res = resolve_triage_item(proj, ch["id"], action="approve", device_id=device_id)
                approved_results.append(res)
                final_children.append(res.get("item") or ch)
            except Exception as exc:
                errors.append(f"{ch['file_name']}: {exc}")
                final_children.append(ch)
        created_children = final_children

    return {
        "ok": True,
        "single_kept": False,
        "auto_approved": bool(auto_approve),
        "split_count": len(created_children),
        "children": created_children,
        "results": approved_results,
        "errors": errors,
    }

