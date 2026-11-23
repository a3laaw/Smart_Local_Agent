#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
project_log.py — Per-project uploaded Excel log engine (.xlsx & .xlsm with VBA intact).

Supports:
  - Uploading a project's own real Excel log(s) into `<project_root>/3_Project_Logs/`
    (no forced generic system log!)
  - Automatic deterministic inspection of both:
      1) Multi-revision horizontal sheets (e.g. REV. 0..REV. 7 with Submit/Date Reply/Action)
      2) Standard flat register sheets (Letters, Submittals, IPC, VO, PO, Petty Cash, Reports)
  - Sidecar `<logname>.logmap.json` with SHA-256 + header anchors + orphan-column warnings
  - Direct in-place row update & row append preserving VBA, formulas, borders, fonts, merges
  - Locked-file queue (`pending_log_ops.json`) with clean flush on Start/Sync & End Session
  - Accurate counter recalculation (`max(log)+1`) across IPC, VO, PO, OUT, IN, PC, RPT,
    MAR/MIR/IR/WIR/SD/MS/PQ, disciplines (CIV/EL/PL/HVAC/FF/ELEV/ARCH/STR/MEC), & Arabic refs
  - Optional isolated `99_Unrelated_غير_مرتبط` sheet when explicitly enabled in settings
"""

from __future__ import annotations

import copy
import datetime as dt
import os
import re
import secrets
import shutil
import tempfile
import threading
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

from i18n_ar import tr

from relevance import is_file_locked, normalize_ar, tokenize_words
from security_utils import (
    archive_instead_of_delete,
    atomic_write_json,
    normalize_win_path,
    now_iso,
    read_json_safe,
    resolve_local_or_unc,
    safe_resolve_within,
    safe_template_name,
    sanitize_upload_filename,
    sha256_of_file,
)

ALLOWED_LOG_EXTS: frozenset[str] = frozenset({".xlsx", ".xlsm"})
UNRELATED_SHEET_NAME = "99_Unrelated_غير_مرتبط"
_LINKED_LOGS_FILENAME = "_linked_logs.json"

# Canonical discipline aliases (including real-world site typos like HAVC / ELVE)
DISC_CANON: dict[str, str] = {
    "CIV": "CIV", "CIVIL": "CIV", "مدني": "CIV",
    "ARCH": "ARCH", "ARC": "ARCH", "ARCHITECTURAL": "ARCH", "معماري": "ARCH",
    "STR": "STR", "STRUCTURAL": "STR", "انشائي": "STR",
    "EL": "EL", "ELEC": "EL", "ELECTRICAL": "EL", "كهرباء": "EL",
    "PL": "PL", "PLUMB": "PL", "PLUMBING": "PL", "صحي": "PL",
    "HVAC": "HVAC", "HAVC": "HVAC", "AC": "HVAC", "تكييف": "HVAC",
    "MEC": "MEC", "MECH": "MEC", "MECHANICAL": "MEC", "ميكانيك": "MEC",
    "FF": "FF", "FIRE": "FF", "اطفاء": "FF", "حريق": "FF",
    "ELEV": "ELEV", "ELVE": "ELEV", "LIFT": "ELEV", "مصاعد": "ELEV",
    "LOW": "EL", "ELV": "EL",
    "GEN": "GEN", "GENERAL": "GEN",
}

STANDARD_COUNTER_KEYS: tuple[str, ...] = (
    "IPC", "VO", "PO", "OUT", "IN", "PC", "RPT",
    "MAR", "MIR", "IR", "WIR", "SD", "MS", "PQ", "NCR", "RFI", "GEN",
)

STANDARD_DISC_KEYS: tuple[str, ...] = (
    "CIV", "ARCH", "STR", "EL", "PL", "HVAC", "MEC", "FF", "ELEV",
)

# Column header recognition rules (normalized whole-word/phrase patterns)
_COL_PATTERNS: dict[str, tuple[str, ...]] = {
    "sn": ("#", "s n", "sn", "sr", "no", "item", "م", "ت", "تسلسل", "الرقم التسلسلي"),
    "ref_no": (
        "reference", "ref", "ref no", "reference no", "submittal no", "letter no", "doc no",
        "document no", "ipc no", "vo no", "po no", "voucher no", "invoice no", "transmittal no",
        "المرجع", "رقم المرجع", "رقم الكتاب", "رقم الاعتماد", "رقم المستند", "رقم المستخلص",
        "رقم الدفعه", "رقم الفاتوره", "رقم السند", "رقم الامر", "كود المرجع", "رقم الصادر", "رقم الوارد",
    ),
    "revision": ("rev", "revision", "rev no", "المراجعه", "رقم المراجعه", "اصدار"),
    "discipline": ("discipline", "disc", "trade", "التخصص", "القسم الفني"),
    "subtype": ("type", "submittal type", "doc type", "category", "النوع", "نوع الاعتماد", "نوع المستند", "التصنيف"),
    "subject": (
        "description", "subject", "title", "item description", "particulars", "details",
        "الموضوع", "البيان", "الوصف", "تفاصيل", "بيان الاعمال", "اسم الماده", "ملخص الموضوع",
    ),
    "party": (
        "party", "from", "to", "sender", "recipient", "supplier", "vendor", "manufacturer",
        "subcontractor", "company", "contractor", "client",
        "الجهه", "المرسل", "المرسل اليه", "المورد", "المقاول", "مقاول الباطن", "الشركه", "المحل", "المستفيد",
    ),
    "doc_date": (
        "date", "submit", "submit date", "submission date", "issue date", "letter date",
        "received date", "invoice date", "doc date",
        "التاريخ", "تاريخ التقديم", "تاريخ الكتاب", "تاريخ الاصدار", "تاريخ الاستلام",
        "تاريخ المستند", "تاريخ الفاتوره", "تاريخ الصادر", "تاريخ الوارد",
    ),
    "reply_date": (
        "date reply", "reply date", "response date", "return date", "approval date", "action date",
        "تاريخ الرد", "تاريخ الاعتماد", "تاريخ العوده", "تاريخ الاجراء",
    ),
    "moh_reply_date": (
        "moh reply", "moh date", "moh reply date", "ministry date", "ministry reply",
        "تاريخ رد الوزاره", "تاريخ رد الوزارة", "تاريخ الوزاره", "تاريخ الوزارة", "تاريخ الوكيل", "رد الوزاره", "رد الوزارة",
    ),
    "status": (
        "status", "action", "code", "approval status", "approval code", "result", "decision",
        "الحاله", "كود الاعتماد", "الاجراء", "النتيجه", "حاله الاعتماد", "القرار",
    ),
    "amount": (
        "amount", "net amount", "value", "amount kwd", "cost", "claimed amount",
        "المبلغ", "القيمه", "صافي المبلغ", "المبلغ د ك", "قيمه الدفعه", "قيمه الامر",
    ),
    "total_amount": (
        "total", "total amount", "gross amount", "certified amount",
        "الاجمالي", "المبلغ الاجمالي", "اجمالي المبلغ", "القيمه المعتمده",
    ),
    "payment_method": ("payment", "payment method", "طريقه الدفع", "وسيله الدفع"),
    "moh_status": ("moh", "الوزاره", "الوزارة", "moh status", "اعتماد الوزاره", "اعتماد الوزارة", "حاله الوزاره"),
    "related_letters": (
        "related letter", "related letters", "linked letter", "linked letters", "reply to", "in reply to",
        "الكتاب المرتبط", "الكتب المرتبطة", "مرتبط بكتاب", "الكتب المتصلة", "ردا على كتاب", "مرجع الرد",
    ),
    "notes": ("remarks", "notes", "comment", "comments", "ملاحظات", "الملاحظات", "التعليقات", "بيانات اضافيه"),
    "file_link": ("file", "link", "hyperlink", "attachment", "path", "الملف", "الرابط", "فتح الملف", "المرفق", "المسار"),
    "device_id": ("device id", "device", "معرف الجهاز", "الجهاز"),
}

_REV_HEADER_RE = re.compile(
    r"^\s*(?:rev(?:ision)?\.?|r|مراجعة|اصدار|إصدار)\s*[-_:#]?\s*0*(\d+)\s*$",
    re.IGNORECASE,
)


def canonical_discipline(raw: str) -> str:
    if not raw:
        return ""
    up = raw.strip().upper()
    if up in DISC_CANON:
        return DISC_CANON[up]
    na = normalize_ar(raw)
    for k, v in DISC_CANON.items():
        if normalize_ar(k) == na:
            return v
    return up[:8]


def logs_dir_for_project(proj_root: Path | str) -> Path:
    d = resolve_local_or_unc(proj_root) / "3_Project_Logs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def linked_logs_path_for(proj_root: Path | str) -> Path:
    return logs_dir_for_project(proj_root) / _LINKED_LOGS_FILENAME


def _load_linked_logs(proj_root: Path | str) -> list[dict[str, Any]]:
    raw = read_json_safe(linked_logs_path_for(proj_root), [])
    if not isinstance(raw, list):
        return []
    out: list[dict[str, Any]] = []
    for item in raw:
        if isinstance(item, dict) and item.get("name") and item.get("path"):
            out.append(dict(item))
    return out


def _save_linked_logs(proj_root: Path | str, items: list[dict[str, Any]]) -> None:
    atomic_write_json(linked_logs_path_for(proj_root), items)


def logmap_path_for(log_path: Path, proj_root: Path | str | None = None) -> Path:
    """
    Sidecar `.logmap.json` path:
    - For logs inside `<proj_root>/3_Project_Logs/`, stored right next to the workbook.
    - For directly linked external logs (on Desktop / Windows Server), stored cleanly inside
      `<proj_root>/3_Project_Logs/.linked_<name>.logmap.json` so the user's folder stays untouched.
    """
    lp = resolve_local_or_unc(log_path)
    if proj_root is not None:
        ldir = resolve_local_or_unc(logs_dir_for_project(proj_root))
        try:
            if not lp.is_relative_to(ldir):
                return ldir / f".linked_{lp.name}.logmap.json"
        except Exception:
            pass
    return lp.parent / f"{lp.name}.logmap.json"


def _iter_project_log_entries(proj_root: Path | str) -> list[dict[str, Any]]:
    """Return both uploaded workbooks in `3_Project_Logs/` and directly linked external logs."""
    ldir = logs_dir_for_project(proj_root)
    entries: list[dict[str, Any]] = []
    seen_names: set[str] = set()
    for f in sorted(ldir.iterdir()):
        if not f.is_file():
            continue
        if f.name.startswith(("~$", ".", "_")):
            continue
        if f.suffix.lower() not in ALLOWED_LOG_EXTS:
            continue
        seen_names.add(f.name.lower())
        entries.append({"name": f.name, "path": f, "is_linked": False, "exists": True})

    for item in _load_linked_logs(proj_root):
        lp = resolve_local_or_unc(str(item["path"]))
        nm = str(item.get("name") or lp.name)
        if nm.lower() in seen_names:
            continue
        seen_names.add(nm.lower())
        exists = False
        try:
            exists = lp.exists() and lp.is_file()
        except OSError:
            exists = False
        entries.append({
            "name": nm,
            "path": lp,
            "is_linked": True,
            "exists": exists,
            "linked_at": item.get("linked_at", ""),
        })
    return entries


def resolve_project_log_path(proj_root: Path | str, log_name: str) -> tuple[Path, bool]:
    """
    Safely resolve either an uploaded log inside `3_Project_Logs/` or a registered
    external linked log in `_linked_logs.json`. Returns `(path, is_linked)`.
    """
    clean_name = safe_template_name(log_name)
    ldir = logs_dir_for_project(proj_root)
    local_cand = safe_resolve_within(ldir, clean_name)
    if local_cand.exists() and local_cand.is_file():
        return local_cand, False
    for item in _load_linked_logs(proj_root):
        if str(item.get("name")) == clean_name:
            return resolve_local_or_unc(str(item["path"])), True
    raise FileNotFoundError(tr("server.log_not_found"))


def pending_queue_path_for(proj_root: Path | str) -> Path:
    return logs_dir_for_project(proj_root) / "pending_log_ops.json"


def audit_jsonl_path_for(proj_root: Path | str) -> Path:
    return logs_dir_for_project(proj_root) / "_log_audit.jsonl"


def _append_audit_event(proj_root: Path | str, event: dict[str, Any]) -> None:
    import json
    p = audit_jsonl_path_for(proj_root)
    p.parent.mkdir(parents=True, exist_ok=True)
    rec = {"timestamp": now_iso(), **event}
    with open(p, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def _match_col_role(header_text: str) -> str | None:
    raw = str(header_text or "").strip()
    if raw in ("#", "\\", "№", "م.", "ت.", "No.", "NO."):
        return "sn"
    hn = normalize_ar(raw)
    if not hn:
        return None
    # Exact match first
    for role, pats in _COL_PATTERNS.items():
        for p in pats:
            pn = normalize_ar(p)
            if pn and hn == pn:
                return role
    # Phrase containment for longer descriptive headers (e.g. 'Status / الحالة', 'المبلغ (د.ك)')
    priority_order = (
        "moh_reply_date", "moh_status", "ref_no", "reply_date", "doc_date", "revision", "discipline", "subtype",
        "subject", "party", "status", "total_amount", "amount", "payment_method",
        "file_link", "device_id", "related_letters", "notes", "sn",
    )
    htoks = set(hn.split())
    for role in priority_order:
        for p in _COL_PATTERNS[role]:
            pn = normalize_ar(p)
            if not pn or len(pn) < 3:
                continue
            pparts = pn.split()
            if len(pparts) == 1:
                if pparts[0] in htoks:
                    return role
            elif re.search(rf"(?<!\w){re.escape(pn)}(?!\w)", hn):
                return role
    return None


def _infer_sheet_route(sheet_name: str, title_hint: str = "") -> dict[str, Any]:
    """
    Deterministically classify a workbook sheet into its document routing rule.
    Handles real Kuwait site logs (including `LOG_Final.xlsm` sheets like CIV, EL, PL,
    HVAC, FF, ELVE, CHECK LIST, MOH, صادر -السور, وراد السور, صادر مقاولين الباطن,
    وارد مقاولين الباطن, CONT, INPUT, 00-Dashboard, plus standard Arabic/English logs).
    """
    sn_raw = sheet_name.strip()
    sn_up = sn_raw.upper()
    sn_norm = normalize_ar(sn_raw)
    sn_toks = set(tokenize_words(sn_raw))

    # Skip pure dashboard/summary/KPI/chart/unrelated sheets
    if sn_raw == UNRELATED_SHEET_NAME or "unrelated" in sn_toks or "غير مرتبط" in sn_norm:
        return {"skip": True, "reason": "unrelated_sheet", "routes": []}
    if any(w in sn_toks for w in ("dashboard", "summary", "kpi", "chart", "charts", "cover", "index")) or any(
        p in sn_norm for p in ("ملخص", "لوحه", "احصائيات", "غلاف", "سجل السكان الكامل")
    ):
        return {"skip": True, "reason": "dashboard_or_summary", "routes": []}

    # Central INPUT mirror sheet (common alongside per-discipline sheets)
    if sn_up in ("INPUT", "ALL_INPUT", "MASTER_INPUT"):
        return {
            "skip": False,
            "is_central_input": True,
            "routes": [{"section": "submittal", "disciplines": [], "subtypes": [], "party_role": ""}],
        }

    # Direct discipline submittal tabs
    exact_disc_map = {
        "CIV": ["CIV", "ARCH", "STR"],
        "CIVIL": ["CIV"],
        "ARCH": ["ARCH"],
        "ARC": ["ARCH"],
        "STR": ["STR"],
        "STRUCTURAL": ["STR"],
        "EL": ["EL"],
        "ELEC": ["EL"],
        "ELECTRICAL": ["EL"],
        "PL": ["PL"],
        "PLUMBING": ["PL"],
        "HVAC": ["HVAC", "MEC"],
        "HAVC": ["HVAC", "MEC"],
        "MEC": ["MEC"],
        "MECH": ["MEC"],
        "FF": ["FF"],
        "FIRE": ["FF"],
        "ELVE": ["ELEV"],
        "ELEV": ["ELEV"],
        "ELEVATOR": ["ELEV"],
    }
    if sn_up in exact_disc_map:
        return {
            "skip": False,
            "routes": [{"section": "submittal", "disciplines": exact_disc_map[sn_up], "subtypes": [], "party_role": ""}],
        }

    # Checklist / IR / WIR tab
    if "check" in sn_toks or "checklist" in sn_toks or sn_toks & {"ir", "wir"} or "فحص اعمال" in sn_norm:
        return {
            "skip": False,
            "routes": [{"section": "submittal", "disciplines": [], "subtypes": ["IR", "WIR", "CHECKLIST"], "party_role": ""}],
        }

    # Specific submittal type tabs
    if "mar" in sn_toks or "اعتماد مواد" in sn_norm:
        return {"skip": False, "routes": [{"section": "submittal", "disciplines": [], "subtypes": ["MAR", "SM"], "party_role": ""}]}
    if "mir" in sn_toks or "فحص مواد" in sn_norm:
        return {"skip": False, "routes": [{"section": "submittal", "disciplines": [], "subtypes": ["MIR"], "party_role": ""}]}
    if "sd" in sn_toks or "shop" in sn_toks or "مخططات" in sn_norm:
        return {"skip": False, "routes": [{"section": "submittal", "disciplines": [], "subtypes": ["SD"], "party_role": ""}]}
    if sn_toks & {"ms", "pq"} or "طريقه تنفيذ" in sn_norm or "تاهيل" in sn_norm:
        return {"skip": False, "routes": [{"section": "submittal", "disciplines": [], "subtypes": ["MS", "PQ"], "party_role": ""}]}

    # Incoming / Outgoing correspondence (note common Kuwait typo `وراد` for `وارد`)
    is_in = bool(sn_toks & {"incoming", "inbox", "in", "وارد", "وراد", "الوارد"})
    is_out = bool(sn_toks & {"outgoing", "outbox", "out", "صادر", "الصادر"})
    party_role = ""
    if any(w in sn_norm for w in ("مقاولين الباطن", "مقاول باطن", "مقاولي الباطن")) or (sn_toks & {"sub", "subcontractor", "subcontractors"}):
        party_role = "subcontractor"
    elif any(w in sn_norm for w in ("السور", "استشاري", "الاستشاري", "مكتب")) or (sn_toks & {"consultant", "engineer"}):
        party_role = "consultant"
    elif any(w in sn_norm for w in ("وزاره", "الصحه", "الاشغال", "المالك", "وكيل", "الشؤون الهندسيه", "الشئون الهندسيه")) or (sn_toks & {"moh", "mpw", "pahw", "mew", "owner", "client", "ministry", "undersecretary"}):
        party_role = "owner"
    elif sn_up in ("CONT", "CONTRACTOR") or "المقاول" in sn_norm:
        party_role = "contractor"

    if is_out and not is_in:
        return {"skip": False, "routes": [{"section": "outgoing", "disciplines": [], "subtypes": [], "party_role": party_role}]}
    if is_in and not is_out:
        return {"skip": False, "routes": [{"section": "incoming", "disciplines": [], "subtypes": [], "party_role": party_role}]}
    if is_in and is_out:
        return {
            "skip": False,
            "routes": [
                {"section": "outgoing", "disciplines": [], "subtypes": [], "party_role": party_role},
                {"section": "incoming", "disciplines": [], "subtypes": [], "party_role": party_role},
            ],
        }
    if sn_up in ("MOH", "MPW", "PAHW", "MEW", "OWNER", "CLIENT"):
        return {"skip": False, "routes": [{"section": "outgoing", "disciplines": [], "subtypes": [], "party_role": "owner"}]}
    if sn_up in ("CONT", "CONTRACTOR"):
        return {"skip": False, "routes": [{"section": "outgoing", "disciplines": [], "subtypes": [], "party_role": "contractor"}]}

    # Financial: IPC / VO / PO / Petty Cash
    has_ipc = bool(sn_toks & {"ipc", "ipcs", "payment", "certificate", "مستخلص", "مستخلصات", "المستخلصات", "دفعات", "عقود", "العقود"})
    has_vo = bool(sn_toks & {"vo", "vos", "variation", "variations"}) or "تغيير" in sn_norm or "تغييريه" in sn_norm
    has_po = bool(sn_toks & {"po", "pos", "purchase"}) or "شراء" in sn_norm
    if has_ipc or has_vo or has_po:
        routes = []
        if has_ipc:
            routes.append({"section": "contract_ipc", "disciplines": [], "subtypes": ["IPC"], "party_role": ""})
        if has_vo:
            routes.append({"section": "variation_vo", "disciplines": [], "subtypes": ["VO"], "party_role": ""})
        if has_po:
            routes.append({"section": "purchase_po", "disciplines": [], "subtypes": ["PO"], "party_role": ""})
        # Legacy combined contract_ipc sheet covers VO/PO too if only IPC mentioned with العقود
        if has_ipc and not has_vo and "العقود" in sn_norm:
            routes.append({"section": "variation_vo", "disciplines": [], "subtypes": ["VO"], "party_role": ""})
            routes.append({"section": "purchase_po", "disciplines": [], "subtypes": ["PO"], "party_role": ""})
        return {"skip": False, "routes": routes}

    if sn_toks & {"petty", "cash", "expense", "expenses", "pc"} or any(
        w in sn_norm for w in ("عهده", "العهد", "مصروفات", "المصروفات", "فواتير")
    ):
        return {"skip": False, "routes": [{"section": "petty_cash", "disciplines": [], "subtypes": ["PC"], "party_role": ""}]}

    if sn_toks & {"report", "reports", "daily", "weekly", "monthly", "rpt"} or any(
        w in sn_norm for w in ("تقرير", "تقارير", "التقارير", "محاضر")
    ):
        return {"skip": False, "routes": [{"section": "site_report", "disciplines": [], "subtypes": ["RPT"], "party_role": ""}]}

    if sn_toks & {"submittal", "submittals", "log", "register"} or any(
        w in sn_norm for w in ("اعتمادات", "الاعتمادات", "سجل")
    ):
        return {"skip": False, "routes": [{"section": "submittal", "disciplines": [], "subtypes": [], "party_role": ""}]}

    if sn_toks & {"letters", "correspondence"} or any(w in sn_norm for w in ("كتب", "مراسلات", "خطابات")):
        return {
            "skip": False,
            "routes": [
                {"section": "outgoing", "disciplines": [], "subtypes": [], "party_role": ""},
                {"section": "incoming", "disciplines": [], "subtypes": [], "party_role": ""},
            ],
        }

    return {"skip": False, "routes": [{"section": "general", "disciplines": [], "subtypes": [], "party_role": ""}]}


def inspect_workbook_to_logmap(log_path: Path, *, device_id: str = "") -> dict[str, Any]:
    """
    Deterministically inspect an uploaded `.xlsx`/`.xlsm` project log and build a
    complete `<logname>.logmap.json` profile (preserving formulas & multi-revision layout).
    """
    log_path = Path(log_path)
    ext = log_path.suffix.lower()
    fhash = sha256_of_file(log_path)

    wb = openpyxl.load_workbook(log_path, data_only=False, keep_vba=(ext == ".xlsm"))
    has_vba = bool(getattr(wb, "vba_archive", None) is not None)
    sheets_spec: list[dict[str, Any]] = []
    has_dedicated_disc_sheets = any(
        s.strip().upper() in ("CIV", "EL", "PL", "HVAC", "HAVC", "FF", "ELVE", "ELEV", "ARCH", "STR", "MEC")
        for s in wb.sheetnames
    )

    for sname in wb.sheetnames:
        ws = wb[sname]
        max_r = min(ws.max_row or 1, 25)
        max_c = min(ws.max_column or 1, 60)

        # Grab title hint from rows 1..3
        title_bits = []
        for r in range(1, min(4, max_r + 1)):
            for c in range(1, min(8, max_c + 1)):
                v = ws.cell(r, c).value
                if isinstance(v, str) and not v.startswith("=") and v.strip():
                    title_bits.append(v.strip()[:80])
        title_hint = " | ".join(title_bits[:3])

        route_info = _infer_sheet_route(sname, title_hint)
        if route_info.get("skip"):
            sheets_spec.append({
                "sheet": sname,
                "layout": "skip",
                "enabled": False,
                "skip_reason": route_info.get("reason", "skip"),
                "routes": [],
                "header_row": 1,
                "subheader_row": None,
                "data_start_row": 2,
                "columns": {},
                "rev_groups": [],
                "formula_cols": [],
                "status_vocab": {},
            })
            continue

        # Find best header row in rows 1..15 (score UNIQUE roles so repeated subheaders like Submit x8 don't beat main header!)
        best_hr = 1
        best_score = -1
        best_rev_starts: list[tuple[int, int, str]] = []
        for r in range(1, min(16, max_r + 1)):
            score = 0
            seen_roles: set[str] = set()
            rev_starts: list[tuple[int, int, str]] = []
            for c in range(1, max_c + 1):
                val = ws.cell(r, c).value
                if not isinstance(val, str) or val.startswith("="):
                    continue
                txt = val.strip()
                if not txt:
                    continue
                m_rev = _REV_HEADER_RE.match(txt)
                if m_rev:
                    rev_starts.append((int(m_rev.group(1)), c, txt))
                    score += 4
                    continue
                role = _match_col_role(txt)
                if role and role not in seen_roles:
                    seen_roles.add(role)
                    if role in ("ref_no", "subject"):
                        score += 8
                    else:
                        score += 3
            if score > best_score:
                best_score = score
                best_hr = r
                best_rev_starts = rev_starts

        # Detect multi-revision horizontal layout (e.g. R4: REV. 0..7, R5: Submit | Date Reply | Action)
        rev_groups: list[dict[str, Any]] = []
        subheader_row: int | None = None
        if best_rev_starts and best_hr + 1 <= (ws.max_row or 1):
            sub_r = best_hr + 1
            # Also catch REV. 0 when its group header cell was left blank (like sheet `CONT` in LOG_Final.xlsm!)
            submit_cols: list[int] = []
            for c in range(1, max_c + 1):
                sv = ws.cell(sub_r, c).value
                if isinstance(sv, str) and normalize_ar(sv) in ("submit", "تقديم", "تاريخ التقديم"):
                    submit_cols.append(c)
            if submit_cols:
                subheader_row = sub_r
                rev_by_col = {col: (rnum, lbl) for rnum, col, lbl in best_rev_starts}
                for idx, sc in enumerate(submit_cols):
                    if sc in rev_by_col:
                        rnum, lbl = rev_by_col[sc]
                    else:
                        rnum, lbl = idx, f"REV. {idx}"
                    rep_col = sc + 1 if sc + 1 <= max_c else None
                    act_col = sc + 2 if sc + 2 <= max_c else None
                    # Verify subheaders if present
                    rev_groups.append({
                        "rev": rnum,
                        "label": lbl,
                        "submit_col": sc,
                        "reply_col": rep_col,
                        "action_col": act_col,
                    })
                rev_groups.sort(key=lambda g: g["rev"])

        data_start_row = (subheader_row or best_hr) + 1

        # Detect formula columns from sample data rows (scan up to 500 rows so late formulas are caught)
        formula_cols_set: set[int] = set()
        for r in range(data_start_row, min(data_start_row + 500, (ws.max_row or data_start_row) + 1)):
            for c in range(1, max_c + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.startswith("="):
                    formula_cols_set.add(c)

        # Map non-revision columns
        rev_covered_cols: set[int] = set()
        for g in rev_groups:
            for k in ("submit_col", "reply_col", "action_col"):
                if g.get(k):
                    rev_covered_cols.add(int(g[k]))

        columns: dict[str, dict[str, Any]] = {}
        for c in range(1, max_c + 1):
            if c in rev_covered_cols:
                continue
            v = ws.cell(best_hr, c).value
            if (v is None or str(v).strip() == "") and subheader_row:
                v = ws.cell(subheader_row, c).value
            if not isinstance(v, str) or v.startswith("=") or not v.strip():
                continue
            htxt = v.strip()
            role = _match_col_role(htxt)
            if not role:
                continue
            # Keep formula status col noted as formula_cols; don't map writable status onto a formula col in multi_rev
            if c in formula_cols_set and role == "status" and rev_groups:
                continue
            if role not in columns:
                columns[role] = {
                    "col": c,
                    "col_letter": get_column_letter(c),
                    "header": htxt,
                    "header_norm": normalize_ar(htxt),
                    "is_formula": c in formula_cols_set,
                }

        # Extract DataValidation lists & overdue threshold from live sheet formulas
        dv_action_opts: list[str] = []
        dv_moh_opts: list[str] = []
        action_col_nums = {
            int(g["action_col"])
            for g in rev_groups if g.get("action_col")
        }
        if not action_col_nums and columns.get("status", {}).get("col"):
            action_col_nums.add(int(columns["status"]["col"]))
        moh_col_num = int(columns.get("moh_status", {}).get("col") or 0)

        try:
            from openpyxl.utils.cell import column_index_from_string
            for dv in getattr(getattr(ws, "data_validations", None), "dataValidation", []) or []:
                if getattr(dv, "type", "") != "list" or not getattr(dv, "formula1", None):
                    continue
                f1 = str(dv.formula1).strip()
                if f1.startswith("="):
                    continue
                if len(f1) >= 2 and f1[0] == '"' and f1[-1] == '"':
                    f1 = f1[1:-1]
                opts = [x.strip() for x in f1.split(",") if x.strip()]
                if not opts:
                    continue
                sq = str(getattr(dv, "sqref", "") or "")
                dv_col_nums: set[int] = set()
                for part in sq.replace(",", " ").split():
                    m_rng = re.match(r"^\$?([A-Z]{1,3})\$?\d*(?::\$?([A-Z]{1,3})\$?\d*)?$", part.strip(), re.I)
                    if m_rng:
                        c1 = column_index_from_string(m_rng.group(1).upper())
                        c2 = column_index_from_string((m_rng.group(2) or m_rng.group(1)).upper())
                        dv_col_nums.update(range(min(c1, c2), max(c1, c2) + 1))
                if moh_col_num and moh_col_num in dv_col_nums and not dv_moh_opts:
                    dv_moh_opts = opts
                elif (dv_col_nums & action_col_nums) and not dv_action_opts:
                    dv_action_opts = opts
        except Exception:
            pass

        # Detect overdue days threshold from formula columns (e.g. TODAY()-...>6)
        overdue_days: int = 14
        try:
            for chk_r in range(data_start_row, min(data_start_row + 5, (ws.max_row or data_start_row) + 1)):
                for fc in sorted(formula_cols_set):
                    fv = ws.cell(chk_r, int(fc)).value
                    if isinstance(fv, str) and "TODAY()" in fv.upper():
                        m_od = re.search(r"TODAY\s*\(\s*\)\s*-[^,)>]+>\s*(\d{1,3})\s*[,)]", fv, re.I)
                        if m_od:
                            od_val = int(m_od.group(1))
                            if 1 <= od_val <= 90:
                                overdue_days = od_val
                                break
                if overdue_days != 14:
                    break
        except Exception:
            pass

        # Detect status vocabulary used in this workbook
        status_vocab: dict[str, str] = {
            "approved": "Approved",
            "approved_noted": "Approved" if rev_groups else "Approved as Noted",
            "revise": "rejected" if rev_groups else "Rejected",
            "rejected": "rejected" if rev_groups else "Rejected",
            "pending": "pending" if rev_groups else "Under Review",
            "withdrawn": "withdrawn" if rev_groups else "Cancelled",
        }
        moh_vocab: dict[str, str] = {
            "approved": "✅ Approved",
            "approved_noted": "✅ Approved",
            "revise": "❌ Rejected",
            "rejected": "❌ Rejected",
            "pending": "⏳ Under Review",
            "withdrawn": "❌ Rejected",
        }
        if dv_moh_opts:
            for opt in dv_moh_opts:
                ol = opt.lower()
                if "approv" in ol or "معتمد" in ol:
                    moh_vocab["approved"] = opt
                    moh_vocab["approved_noted"] = opt
                elif "reject" in ol or "مرفوض" in ol:
                    moh_vocab["rejected"] = opt
                    moh_vocab["revise"] = opt
                    moh_vocab["withdrawn"] = opt
                elif "review" in ol or "pend" in ol or "انتظار" in ol or "مراجع" in ol:
                    moh_vocab["pending"] = opt

        if dv_action_opts:
            dv_lower = {o.lower(): o for o in dv_action_opts}
            if any("approved (c)" in k or "code c" in k for k in dv_lower):
                for o in dv_action_opts:
                    ol = o.lower()
                    if ol == "approved":
                        status_vocab["approved"] = o
                    elif "approved (c)" in ol or "code c" in ol or "noted" in ol:
                        status_vocab["approved_noted"] = o
                        status_vocab["revise"] = o
                    elif "reject" in ol:
                        status_vocab["rejected"] = o
                    elif "pend" in ol or "review" in ol:
                        status_vocab["pending"] = o
                    elif "withdraw" in ol or "cancel" in ol:
                        status_vocab["withdrawn"] = o
            elif all(o.upper() in ("A", "B", "C", "D", "UR", "UNDER REVIEW", "PENDING") for o in dv_action_opts):
                status_vocab = {
                    "approved": dv_lower.get("a", "A"),
                    "approved_noted": dv_lower.get("b", "B"),
                    "revise": dv_lower.get("c", "C"),
                    "rejected": dv_lower.get("d", dv_lower.get("c", "C")),
                    "pending": dv_lower.get("under review", dv_lower.get("pending", dv_lower.get("ur", "Under Review"))),
                    "withdrawn": dv_lower.get("d", "D"),
                }
            else:
                for o in dv_action_opts:
                    ol = o.lower()
                    if ol in ("approved", "معتمد"):
                        status_vocab["approved"] = o
                    elif "noted" in ol or "comment" in ol or "(c)" in ol or "ملاحظ" in ol:
                        status_vocab["approved_noted"] = o
                    elif "revise" in ol or "resubmit" in ol or "اعاد" in ol:
                        status_vocab["revise"] = o
                    elif "reject" in ol or "مرفوض" in ol:
                        status_vocab["rejected"] = o
                    elif "pend" in ol or "review" in ol or "مراجع" in ol or "انتظار" in ol:
                        status_vocab["pending"] = o
                    elif "withdraw" in ol or "cancel" in ol or "ملغ" in ol:
                        status_vocab["withdrawn"] = o
        else:
            # Inspect actual action values to detect A/B/C/D vs Approved/rejected
            sample_statuses: list[str] = []
            check_col = rev_groups[0]["action_col"] if rev_groups else (columns.get("status", {}).get("col"))
            if check_col:
                for r in range(data_start_row, min(data_start_row + 30, (ws.max_row or data_start_row) + 1)):
                    sv = ws.cell(r, int(check_col)).value
                    if isinstance(sv, str) and not sv.startswith("=") and sv.strip():
                        sample_statuses.append(sv.strip())
            if sample_statuses and all(s.upper() in ("A", "B", "C", "D", "UR", "UNDER REVIEW", "PENDING") for s in sample_statuses[:10]):
                if any(s.upper() in ("A", "B", "C", "D") for s in sample_statuses):
                    status_vocab = {
                        "approved": "A",
                        "approved_noted": "B",
                        "revise": "C",
                        "rejected": "D" if any(s.upper() == "D" for s in sample_statuses) else "C",
                        "pending": "Under Review",
                        "withdrawn": "D",
                    }
            elif sample_statuses and any("approved (c)" in s.lower() for s in sample_statuses):
                c_sample = next(s for s in sample_statuses if "approved (c)" in s.lower())
                status_vocab["approved"] = "approved" if any(s == "approved" for s in sample_statuses) else status_vocab["approved"]
                status_vocab["approved_noted"] = c_sample
                status_vocab["revise"] = c_sample

        layout = "multi_rev" if rev_groups else ("flat" if ("ref_no" in columns or "subject" in columns or "amount" in columns) else "skip")
        enabled = layout != "skip"
        # If sheet is central INPUT and dedicated discipline sheets exist, mark mirror=False by default to avoid double-counting
        is_central_input = bool(route_info.get("is_central_input"))
        if is_central_input and has_dedicated_disc_sheets:
            enabled = False

        sheets_spec.append({
            "sheet": sname,
            "layout": layout,
            "enabled": enabled,
            "is_central_input": is_central_input,
            "title_hint": title_hint,
            "routes": route_info.get("routes", []) if layout != "skip" else [],
            "header_row": best_hr,
            "subheader_row": subheader_row,
            "data_start_row": data_start_row,
            "columns": columns,
            "rev_groups": rev_groups,
            "formula_cols": sorted(formula_cols_set),
            "overdue_days": overdue_days,
            "status_vocab": status_vocab,
            "moh_vocab": moh_vocab if "moh_status" in columns or dv_moh_opts else {},
        })

    wb.close()
    return {
        "schema_version": "1.0",
        "log_name": log_path.name,
        "log_ext": ext,
        "has_vba": has_vba,
        "log_sha256": fhash,
        "orphan_columns": [],
        "updated_at": now_iso(),
        "device_id": device_id,
        "sheets": sheets_spec,
    }


def load_or_create_logmap(
    log_path: Path,
    *,
    device_id: str = "",
    proj_root: Path | str | None = None,
) -> dict[str, Any]:
    """
    Load `<logname>.logmap.json` (creating it deterministically if missing) and
    reconcile column header anchors if the Excel file was edited externally.
    """
    log_path = Path(log_path)
    mp = logmap_path_for(log_path, proj_root=proj_root)
    existing = read_json_safe(mp)
    if not isinstance(existing, dict) or not existing.get("sheets"):
        fresh = inspect_workbook_to_logmap(log_path, device_id=device_id)
        atomic_write_json(mp, fresh)
        return fresh

    cur_hash = sha256_of_file(log_path)
    if existing.get("log_sha256") == cur_hash:
        return existing

    # Workbook changed externally (e.g. user opened Excel and added rows/columns) -> rebind by header anchor!
    orphans: list[dict[str, Any]] = []
    try:
        wb = openpyxl.load_workbook(log_path, read_only=True, data_only=False)
        wb_sheets = set(wb.sheetnames)
        for s_cfg in existing.get("sheets", []):
            sname = s_cfg.get("sheet")
            if sname not in wb_sheets:
                if s_cfg.get("enabled"):
                    orphans.append({"sheet": sname, "column": "*", "reason": tr("validation.sheet_missing", sheet=sname)})
                continue
            if not s_cfg.get("enabled") or s_cfg.get("layout") == "skip":
                continue
            ws = wb[sname]
            hr = int(s_cfg.get("header_row") or 1)
            row_vals = list(next(ws.iter_rows(min_row=hr, max_row=hr, max_col=60, values_only=True), ()))
            norm_headers = {idx + 1: normalize_ar(str(v)) for idx, v in enumerate(row_vals) if v is not None and str(v).strip()}
            raw_headers = {idx + 1: str(v).strip() for idx, v in enumerate(row_vals) if v is not None and str(v).strip()}
            for role, cinfo in list((s_cfg.get("columns") or {}).items()):
                col_idx = int(cinfo.get("col") or 1)
                expected_norm = str(cinfo.get("header_norm") or normalize_ar(str(cinfo.get("header") or "")))
                if not expected_norm:
                    continue
                if norm_headers.get(col_idx) == expected_norm:
                    continue
                # Try rebinding to moved column with same normalized header
                moved_to = next((ci for ci, hn in norm_headers.items() if hn == expected_norm), None)
                if moved_to is not None:
                    cinfo["col"] = moved_to
                    cinfo["col_letter"] = get_column_letter(moved_to)
                    cinfo["header"] = raw_headers.get(moved_to, cinfo.get("header", ""))
                else:
                    orphans.append({
                        "sheet": sname,
                        "role": role,
                        "expected_header": cinfo.get("header", role),
                        "old_col": cinfo.get("col_letter", str(col_idx)),
                        "reason": tr("server.col_changed", col=cinfo.get("header", role), sheet=sname),
                    })
        # Also detect newly added sheets
        known_names = {s.get("sheet") for s in existing.get("sheets", [])}
        wb.close()
        if any(s not in known_names for s in wb_sheets):
            fresh = inspect_workbook_to_logmap(log_path, device_id=device_id)
            for fs in fresh.get("sheets", []):
                if fs["sheet"] not in known_names:
                    existing["sheets"].append(fs)
    except Exception:
        pass

    # Keep existing schema consistent: skip sheets never stay enabled
    for s_cfg in existing.get("sheets", []):
        if str(s_cfg.get("layout") or "") == "skip" and s_cfg.get("enabled"):
            s_cfg["enabled"] = False

    existing["log_sha256"] = cur_hash
    existing["orphan_columns"] = orphans
    existing["updated_at"] = now_iso()
    atomic_write_json(mp, existing)
    return existing


def save_logmap(
    log_path: Path,
    new_map: dict[str, Any],
    *,
    device_id: str = "",
    proj_root: Path | str | None = None,
) -> dict[str, Any]:
    """Validate and save user-edited `<logname>.logmap.json`."""
    log_path = Path(log_path)
    if not log_path.exists():
        raise FileNotFoundError(tr("server.log_not_found"))
    current = load_or_create_logmap(log_path, device_id=device_id, proj_root=proj_root)
    if isinstance(new_map.get("sheets"), list):
        sanitized_sheets = []
        for s in new_map["sheets"]:
            if not isinstance(s, dict) or not s.get("sheet"):
                continue
            cols = {}
            for role, cinfo in (s.get("columns") or {}).items():
                if isinstance(cinfo, int):
                    col_i = cinfo
                    cols[role] = {
                        "col": col_i,
                        "col_letter": get_column_letter(col_i),
                        "header": role,
                        "header_norm": normalize_ar(role),
                    }
                elif isinstance(cinfo, str) and cinfo.strip():
                    cs = cinfo.strip().upper()
                    col_i = int(cs) if cs.isdigit() else column_index_from_string(cs)
                    cols[role] = {
                        "col": col_i,
                        "col_letter": get_column_letter(col_i),
                        "header": role,
                        "header_norm": normalize_ar(role),
                    }
                elif isinstance(cinfo, dict) and cinfo.get("col"):
                    col_i = int(cinfo["col"])
                    hdr = str(cinfo.get("header") or role)
                    cols[role] = {
                        "col": col_i,
                        "col_letter": get_column_letter(col_i),
                        "header": hdr,
                        "header_norm": str(cinfo.get("header_norm") or normalize_ar(hdr)),
                        "is_formula": bool(cinfo.get("is_formula", False)),
                    }
            layout_val = str(s.get("layout") or "flat")
            enabled_val = False if layout_val == "skip" else bool(s.get("enabled", True))
            sanitized_sheets.append({
                "sheet": str(s["sheet"]),
                "layout": layout_val,
                "enabled": enabled_val,
                "is_central_input": bool(s.get("is_central_input", False)),
                "title_hint": str(s.get("title_hint") or ""),
                "routes": list(s.get("routes") or []),
                "header_row": max(1, int(s.get("header_row") or 1)),
                "subheader_row": int(s["subheader_row"]) if s.get("subheader_row") else None,
                "data_start_row": max(2, int(s.get("data_start_row") or 2)),
                "columns": cols,
                "rev_groups": list(s.get("rev_groups") or []),
                "formula_cols": [int(x) for x in (s.get("formula_cols") or [])],
                "overdue_days": s.get("overdue_days"),
                "status_vocab": dict(s.get("status_vocab") or {}),
                "moh_vocab": dict(s.get("moh_vocab") or {}),
            })
        current["sheets"] = sanitized_sheets
    current["log_sha256"] = sha256_of_file(log_path)
    current["orphan_columns"] = []
    current["updated_at"] = now_iso()
    if device_id:
        current["device_id"] = device_id
    atomic_write_json(logmap_path_for(log_path, proj_root=proj_root), current)
    return current


def list_project_logs(proj_root: Path | str, *, device_id: str = "") -> list[dict[str, Any]]:
    """Return all uploaded and directly linked `.xlsx`/`.xlsm` log files for the project."""
    out: list[dict[str, Any]] = []
    for ent in _iter_project_log_entries(proj_root):
        f = Path(ent["path"])
        is_linked = bool(ent.get("is_linked"))
        if not ent.get("exists"):
            out.append({
                "name": str(ent["name"]),
                "path": str(f),
                "ext": f.suffix.lower(),
                "is_linked": is_linked,
                "size_kb": 0.0,
                "locked": True,
                "missing": True,
                "error": tr("server.log_not_found"),
                "sheets_total": 0,
                "sheets_enabled": 0,
                "sections_covered": [],
                "orphan_columns": [],
            })
            continue
        try:
            lmap = load_or_create_logmap(f, device_id=device_id, proj_root=proj_root)
        except Exception as exc:
            out.append({
                "name": f.name,
                "path": str(f),
                "ext": f.suffix.lower(),
                "is_linked": is_linked,
                "size_kb": round(f.stat().st_size / 1024, 1),
                "locked": is_file_locked(f, allow_readonly=False),
                "error": str(exc),
                "sheets_total": 0,
                "sheets_enabled": 0,
                "sections_covered": [],
                "orphan_columns": [],
            })
            continue
        enabled_sheets = [s for s in lmap.get("sheets", []) if s.get("enabled")]
        sections = sorted({
            str(rt.get("section"))
            for s in enabled_sheets
            for rt in (s.get("routes") or [])
            if rt.get("section")
        })
        out.append({
            "name": f.name,
            "path": str(f),
            "ext": f.suffix.lower(),
            "is_linked": is_linked,
            "has_vba": bool(lmap.get("has_vba")),
            "size_kb": round(f.stat().st_size / 1024, 1),
            "modified_at": dt.datetime.fromtimestamp(f.stat().st_mtime).isoformat(timespec="seconds"),
            "locked": is_file_locked(f, allow_readonly=False),
            "sheets_total": len(lmap.get("sheets", [])),
            "sheets_enabled": len(enabled_sheets),
            "sections_covered": sections,
            "orphan_columns": lmap.get("orphan_columns", []),
            "sheets": lmap.get("sheets", []),
        })
    return out


def link_external_project_log(
    proj_root: Path | str,
    raw_file_path: str,
    *,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Directly link an existing `.xlsx`/`.xlsm` log file on the user's PC or Windows Server
    share so Smart Projects Hub reads from and writes directly into THAT original file
    in place without copying it into `3_Project_Logs/`.
    """
    raw = str(raw_file_path or "").strip().strip('"').strip("'")
    if not raw or "\x00" in raw:
        raise ValueError(tr("server.log_not_found"))
    target = resolve_local_or_unc(raw)
    if not target.exists() or not target.is_file():
        raise FileNotFoundError(tr("server.log_not_found"))
    if target.suffix.lower() not in ALLOWED_LOG_EXTS:
        raise ValueError(tr("server.log_corrupt"))
    if target.name.startswith(("~$", ".", "_")):
        raise ValueError(tr("server.log_corrupt"))

    # Validate workbook integrity before linking
    try:
        wb_check = openpyxl.load_workbook(target, read_only=True)
        _ = wb_check.sheetnames
        wb_check.close()
    except Exception as exc:
        raise ValueError(tr("server.log_corrupt")) from exc

    proj_root_p = resolve_local_or_unc(proj_root)
    safe_nm = safe_template_name(target.name)

    # Save a silent safety backup snapshot in `_Backups/Original_Logs/` (never modifies target)
    try:
        orig_bkp_dir = proj_root_p / "_Backups" / "Original_Logs"
        orig_bkp_dir.mkdir(parents=True, exist_ok=True)
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(target, orig_bkp_dir / f"{ts}_{safe_nm}")
    except OSError:
        pass

    linked = [
        it for it in _load_linked_logs(proj_root_p)
        if str(it.get("name") or "").lower() != safe_nm.lower()
        and str(normalize_win_path(str(it.get("path") or ""))).lower() != str(target).lower()
    ]
    linked.append({
        "name": safe_nm,
        "path": str(target),
        "linked_at": now_iso(),
        "device_id": device_id,
    })
    _save_linked_logs(proj_root_p, linked)

    lmap = load_or_create_logmap(target, device_id=device_id, proj_root=proj_root_p)
    _append_audit_event(proj_root_p, {
        "event": "log_linked",
        "log_name": target.name,
        "log_path": str(target),
        "sheets": len(lmap.get("sheets", [])),
        "has_vba": lmap.get("has_vba", False),
        "device_id": device_id,
    })
    return lmap


def upload_project_log(
    proj_root: Path | str,
    raw_filename: str,
    raw_bytes: bytes,
    *,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Save an uploaded `.xlsx`/`.xlsm` project log into `<proj_root>/3_Project_Logs/`,
    keep a safety backup of the pristine uploaded copy in `<proj_root>/_Backups/Original_Logs/`
    (zero deletion if replacing an existing log!), and build `<name>.logmap.json`.
    """
    if not raw_bytes or len(raw_bytes) < 100:
        raise ValueError(tr("server.log_corrupt"))
    safe_name = sanitize_upload_filename(raw_filename, allowed_exts=ALLOWED_LOG_EXTS)
    proj_root_p = Path(proj_root).resolve()
    ldir = logs_dir_for_project(proj_root_p)
    dest = safe_resolve_within(ldir, safe_name)

    # Validate workbook integrity before touching existing files
    fd, tmp_s = tempfile.mkstemp(prefix=".upload_log_", suffix=dest.suffix.lower(), dir=str(ldir))
    tmp_p = Path(tmp_s)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(raw_bytes)
        try:
            wb_check = openpyxl.load_workbook(tmp_p, read_only=True)
            _ = wb_check.sheetnames
            wb_check.close()
        except Exception as exc:
            raise ValueError(tr("server.log_corrupt")) from exc

        # Zero-deletion: archive prior log if same name already existed
        if dest.exists():
            archive_instead_of_delete(
                dest,
                proj_root_p / "_Backups" / "Replaced_Logs",
                reason="replaced_by_new_upload",
                device_id=device_id,
            )
            old_mp = logmap_path_for(dest)
            if old_mp.exists():
                archive_instead_of_delete(
                    old_mp,
                    proj_root_p / "_Backups" / "Replaced_Logs",
                    reason="replaced_logmap",
                    device_id=device_id,
                )

        os.replace(tmp_p, dest)
    finally:
        if tmp_p.exists():
            try:
                tmp_p.unlink()
            except OSError:
                pass

    # Save pristine untouched backup copy
    orig_bkp_dir = proj_root_p / "_Backups" / "Original_Logs"
    orig_bkp_dir.mkdir(parents=True, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy2(dest, orig_bkp_dir / f"{ts}_{dest.name}")

    lmap = inspect_workbook_to_logmap(dest, device_id=device_id)
    atomic_write_json(logmap_path_for(dest), lmap)
    _append_audit_event(proj_root_p, {
        "event": "log_uploaded",
        "log_name": dest.name,
        "sheets": len(lmap.get("sheets", [])),
        "has_vba": lmap.get("has_vba", False),
        "device_id": device_id,
    })
    return lmap


def archive_project_log(proj_root: Path | str, log_name: str, *, device_id: str = "") -> Path:
    """
    Safely archive (never delete!) an uploaded project log, OR unlink an external
    directly-linked log file without ever touching or deleting the user's original file!
    """
    proj_root_p = resolve_local_or_unc(proj_root)
    clean_name = safe_template_name(log_name)
    target, is_linked = resolve_project_log_path(proj_root_p, clean_name)
    if is_linked:
        linked = [it for it in _load_linked_logs(proj_root_p) if str(it.get("name")) != clean_name]
        _save_linked_logs(proj_root_p, linked)
        mp = logmap_path_for(target, proj_root=proj_root_p)
        if mp.exists():
            return archive_instead_of_delete(
                mp,
                proj_root_p / "_Backups" / "_Archived_Logs",
                reason="user_unlinked_logmap",
                device_id=device_id,
            )
        return target

    archived = archive_instead_of_delete(
        target,
        proj_root_p / "_Backups" / "_Archived_Logs",
        reason="user_archived_log",
        device_id=device_id,
    )
    mp = logmap_path_for(target, proj_root=proj_root_p)
    if mp.exists():
        archive_instead_of_delete(
            mp,
            proj_root_p / "_Backups" / "_Archived_Logs",
            reason="user_archived_logmap",
            device_id=device_id,
        )
    return archived


# ── Reference Normalization & Target-Sheet Resolution ───────────────────────
def normalize_ref_key(ref: str) -> str:
    """
    Normalize a reference code for duplicate/update matching across real site quirks:
      'EL-056' == 'EL-56' == 'el 56'
      'ELEV 162' == 'ELEV-162' == 'ELVE-162'
      'HAVC-030' == 'HVAC-030'
      'ج/2026/024' == 'ج/2026/24'
    """
    if not ref:
        return ""
    s = normalize_ar(str(ref), keep_punct=True).upper()
    s = s.replace("HAVC", "HVAC").replace("ELVE", "ELEV")
    # Split into alpha/Arabic and numeric runs; strip leading zeros on numbers
    parts = re.findall(r"[A-Z\u0600-\u06FF]+|\d+", s)
    if not parts:
        return ""
    norm_parts = [str(int(p)) if p.isdigit() else p for p in parts]
    return "-".join(norm_parts)


def parse_rev_number(rev: Any) -> int:
    if rev is None:
        return 0
    s = str(rev).strip()
    if not s:
        return 0
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else 0


def _normalize_section_for_log(section: str, subtype: str = "") -> tuple[str, str]:
    """Map granular UI sections (like `submittal_mar`, `variation_vo`) to log section + subtype."""
    sec = (section or "").strip().lower()
    sub = (subtype or "").strip().upper()
    mapping = {
        "submittal_mar": ("submittal", sub or "MAR"),
        "submittal_mir": ("submittal", sub or "MIR"),
        "submittal_ir": ("submittal", sub or "IR"),
        "submittal_sd": ("submittal", sub or "SD"),
        "submittal_ms": ("submittal", sub or "MS"),
        "submittal": ("submittal", sub),
        "outgoing": ("outgoing", sub or "OUT"),
        "incoming": ("incoming", sub or "IN"),
        "contract_ipc": ("contract_ipc", sub or "IPC"),
        "variation_vo": ("variation_vo", sub or "VO"),
        "purchase_po": ("purchase_po", sub or "PO"),
        "petty_cash": ("petty_cash", sub or "PC"),
        "site_report": ("site_report", sub or "RPT"),
        "site_photo": ("site_report", sub or "PHOTO"),
        "general": ("general", sub or "GEN"),
    }
    if sec in mapping:
        base_sec, eff_sub = mapping[sec]
        if base_sec == "contract_ipc" and eff_sub == "VO":
            return "variation_vo", "VO"
        if base_sec == "contract_ipc" and eff_sub == "PO":
            return "purchase_po", "PO"
        return base_sec, eff_sub
    return "general", sub or "GEN"


def _infer_party_role(party_raw: str, proj: dict[str, Any] | None = None) -> str:
    p = normalize_ar(party_raw or "")
    if not p:
        return ""
    if proj:
        if proj.get("consultant") and normalize_ar(proj["consultant"]) in p:
            return "consultant"
        if proj.get("owner") and normalize_ar(proj["owner"]) in p:
            return "owner"
    if any(w in p for w in ("وزاره", "الصحه", "الاشغال", "المالك", "وكيل", "الوكيل", "مهندس الوزاره", "الشؤون الهندسيه", "الشئون الهندسيه", "moh", "mpw", "pahw", "mew", "owner", "ministry", "undersecretary")):
        return "owner"
    if any(w in p for w in ("استشاري", "الاستشاري", "مكتب", "السور", "consultant", "engineer")):
        return "consultant"
    if any(w in p for w in ("مقاول باطن", "باطن", "مورد", "subcontractor", "supplier")):
        return "subcontractor"
    if any(w in p for w in ("مقاول", "contractor")):
        return "contractor"
    return ""


def _pick_target_log_and_sheet(
    proj_root: Path,
    entry: dict[str, Any],
    proj: dict[str, Any] | None = None,
    device_id: str = "",
) -> tuple[Path, dict[str, Any], dict[str, Any]] | None:
    """
    Choose the best matching `(log_path, logmap, sheet_spec)` among the project's
    uploaded logs for `entry`. Supports explicit `target_log` / `target_sheet` override too!
    """
    candidates: list[tuple[Path, dict[str, Any]]] = []
    for ent in _iter_project_log_entries(proj_root):
        if not ent.get("exists"):
            continue
        f = Path(ent["path"])
        try:
            candidates.append((f, load_or_create_logmap(f, device_id=device_id, proj_root=proj_root)))
        except Exception:
            continue
    if not candidates:
        return None

    # Explicit target override
    exp_log = str(entry.get("target_log") or "").strip()
    exp_sheet = str(entry.get("target_sheet") or "").strip()
    if exp_sheet:
        for lp, lm in candidates:
            if exp_log and lp.name != exp_log:
                continue
            for sc in lm.get("sheets", []):
                if sc.get("sheet") == exp_sheet and sc.get("layout") != "skip":
                    return lp, lm, sc

    base_sec, sub = _normalize_section_for_log(
        str(entry.get("section") or entry.get("doc_type") or ""),
        str(entry.get("subtype") or ""),
    )
    disc = canonical_discipline(str(entry.get("discipline") or ""))
    # Also infer discipline from ref_no prefix like 'EL-212' or 'CIV-206'
    if not disc and entry.get("ref_no"):
        m_d = re.match(r"^\s*([A-Za-z]+)[-\s_]*\d+", str(entry["ref_no"]))
        if m_d and m_d.group(1).upper() in DISC_CANON:
            disc = DISC_CANON[m_d.group(1).upper()]
    party_role = str(entry.get("party_role") or "") or _infer_party_role(str(entry.get("party_name") or entry.get("party") or ""), proj)

    best: tuple[int, Path, dict[str, Any], dict[str, Any]] | None = None
    for lp, lm in candidates:
        if exp_log and lp.name != exp_log:
            continue
        for sc in lm.get("sheets", []):
            if not sc.get("enabled") or sc.get("layout") == "skip":
                continue
            for rt in (sc.get("routes") or []):
                rsec = str(rt.get("section") or "")
                rdiscs = [canonical_discipline(d) for d in (rt.get("disciplines") or []) if d]
                rsubs = [str(x).upper() for x in (rt.get("subtypes") or []) if x]
                rparty = str(rt.get("party_role") or "")

                score = 0
                if rsec == base_sec:
                    score += 50
                elif base_sec in ("variation_vo", "purchase_po") and rsec == "contract_ipc" and not rsubs:
                    score += 30
                elif rsec == "general":
                    score += 10
                else:
                    continue

                if rsubs:
                    if sub and sub in rsubs:
                        score += 35
                    elif rdiscs:
                        pass
                    else:
                        score -= 25

                if rdiscs:
                    if disc and disc in rdiscs:
                        # Prefer primary discipline sheet (first in list)
                        score += 40 if rdiscs[0] == disc else 32
                    else:
                        score -= 30

                if rparty:
                    if party_role and rparty == party_role:
                        score += 25
                    elif not party_role and rparty == "consultant":
                        # Default consultant correspondence sheet when party unspecified
                        score += 8
                    else:
                        score -= 12

                if sc.get("is_central_input"):
                    score -= 5

                if best is None or score > best[0]:
                    best = (score, lp, lm, sc)

    if best and best[0] >= 10:
        return best[1], best[2], best[3]

    # Fallback: if only 1 single-sheet flat log is uploaded, use its enabled sheet
    enabled_all = [
        (lp, lm, sc)
        for lp, lm in candidates
        for sc in lm.get("sheets", [])
        if sc.get("enabled") and sc.get("layout") != "skip"
    ]
    if len(enabled_all) == 1:
        return enabled_all[0]
    return None


def _is_moh_authority(authority: str) -> bool:
    if not authority:
        return False
    an = normalize_ar(authority)
    toks = set(an.replace("+", " ").split())
    if toks & {"moh", "mpw", "pahw", "mew", "ministry", "owner", "client", "undersecretary"}:
        return True
    return any(
        w in an for w in (
            "وزاره", "الصحه", "الاشغال", "المالك",
            "وكيل", "الوكيل", "وكيل مساعد", "الوكيل المساعد",
            "مهندس الوزاره", "مهندس المشروع", "مهندس قطاع", "مدير الاداره",
            "الشؤون الهندسيه", "الشئون الهندسيه", "قطاع الهندسه", "اداره المشاريع",
            "ministry engineer", "assistant undersecretary", "client engineer",
        )
    )


def _has_consultant_stamp(authority: str, entry: dict[str, Any] | None = None) -> bool:
    ent = entry or {}
    if bool(ent.get("has_both_stamps")):
        return True
    if str(ent.get("consultant_reply_date") or "").strip() or str(ent.get("consultant_action") or "").strip():
        return True
    an = normalize_ar(authority or "")
    toks = set(an.replace("+", " ").split())
    if toks & {"consultant", "soor", "eng"}:
        return True
    return any(w in an for w in ("استشاري", "الاستشاري", "السور", "مكتب", "consultant"))


def _map_moh_status_to_vocab(raw_status: str, moh_vocab: dict[str, str]) -> str:
    s = (raw_status or "").strip()
    base = moh_vocab or {
        "approved": "✅ Approved",
        "approved_noted": "✅ Approved",
        "revise": "❌ Rejected",
        "rejected": "❌ Rejected",
        "pending": "⏳ Under Review",
        "withdrawn": "❌ Rejected",
    }
    if not s:
        return base.get("pending", "⏳ Under Review")
    sl = s.lower()
    if any(k in sl for k in ("code a", "code b", "approved", "معتمد", "مقبول")) or sl in ("a", "b"):
        return base.get("approved", "✅ Approved")
    if any(k in sl for k in ("code c", "code d", "reject", "revise", "مرفوض", "اعاده تقديم")) or sl in ("c", "d"):
        return base.get("rejected", "❌ Rejected")
    if any(k in sl for k in ("withdrawn", "cancel", "ملغي", "مسحوب")):
        return base.get("withdrawn", base.get("rejected", "❌ Rejected"))
    if any(k in sl for k in ("pending", "under review", "open", "submitted", "انتظار", "قيد المراجع")):
        return base.get("pending", "⏳ Under Review")
    return s


def _map_status_to_sheet_vocab(raw_status: str, vocab: dict[str, str], is_multi_rev: bool) -> str:
    s = (raw_status or "").strip()
    if not s:
        return vocab.get("pending", "pending" if is_multi_rev else "Under Review")
    sl = s.lower()
    # Check specific Code C / Approved as Noted / Revise before generic "approved"
    if (
        sl in ("c", "code c", "approved (c)", "revise", "revise & resubmit", "revise and resubmit", "revise & resubmit (code c)", "اعاده تقديم")
        or "(code c)" in sl
        or "approved (c)" in sl
        or "revise" in sl
    ):
        return vocab.get("revise", vocab.get("approved_noted", vocab.get("rejected", "rejected" if is_multi_rev else "Rejected")))
    if (
        sl in ("b", "code b", "approved as noted", "approved as noted (code b)", "approved with comments", "معتمد بملاحظات")
        or "(code b)" in sl
        or "approved as noted" in sl
        or "with comments" in sl
    ):
        return vocab.get("approved_noted", vocab.get("approved", "Approved as Noted"))
    if (
        sl in ("a", "code a", "approved", "approved (code a)", "معتمد", "مقبول")
        or "(code a)" in sl
    ):
        return vocab.get("approved", "Approved")
    if (
        sl in ("d", "code d", "rejected", "rejected (code d)", "مرفوض")
        or "(code d)" in sl
        or "reject" in sl
    ):
        return vocab.get("rejected", "rejected" if is_multi_rev else "Rejected")
    if sl in ("withdrawn", "cancelled", "canceled", "ملغي", "مسحوب") or "withdraw" in sl or "cancel" in sl:
        return vocab.get("withdrawn", "withdrawn")
    if (
        sl in ("pending", "under review", "open", "submitted", "قيد المراجعة", "قيد الاعتماد", "بانتظار الرد")
        or "under review" in sl
        or "pending" in sl
    ):
        return vocab.get("pending", "pending" if is_multi_rev else "Under Review")
    return s


def _copy_row_style_and_formulas(ws, src_row: int, dst_row: int, max_col: int) -> None:
    """Clone cell formatting and translate row-relative formulas from `src_row` to `dst_row`."""
    if ws.row_dimensions[src_row].height:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    row_ref_re = re.compile(rf"(\$?[A-Z]{{1,3}}){src_row}(?!\d)")
    for c in range(1, max_col + 1):
        sc = ws.cell(src_row, c)
        dc = ws.cell(dst_row, c)
        if sc.has_style:
            dc.font = copy.copy(sc.font)
            dc.fill = copy.copy(sc.fill)
            dc.border = copy.copy(sc.border)
            dc.alignment = copy.copy(sc.alignment)
            dc.number_format = sc.number_format
        sv = sc.value
        if isinstance(sv, str) and sv.startswith("="):
            dc.value = row_ref_re.sub(rf"\g<1>{dst_row}", sv)


def _is_totals_row(ws, row_idx: int, max_col: int) -> bool:
    for c in range(1, min(max_col + 1, 15)):
        v = ws.cell(row_idx, c).value
        if isinstance(v, str):
            if re.match(r"^\s*=\s*(?:SUM|SUBTOTAL)\s*\(", v, re.IGNORECASE):
                return True
            vn = normalize_ar(v)
            if vn in ("total", "grand total", "sub total", "الاجمالي", "المجموع", "الاجمالي الكلي", "اجمالي"):
                return True
    return False


def _update_totals_formulas_after_insert(ws, totals_row: int, old_end_row: int, new_end_row: int, max_col: int) -> None:
    pat = re.compile(rf"(\$?[A-Z]{{1,3}}\$?\d+):(\$?[A-Z]{{1,3}}\$?){old_end_row}(?!\d)")
    for c in range(1, max_col + 1):
        v = ws.cell(totals_row, c).value
        if isinstance(v, str) and v.startswith("="):
            ws.cell(totals_row, c).value = pat.sub(rf"\1:\g<2>{new_end_row}", v)


_MISSING_DATE_FILL = PatternFill("solid", fgColor="FEF3C7")
_CLEAR_CELL_FILL = PatternFill(fill_type=None)


def _resolve_target_cell(ws, row: int, col: int):
    cell = ws.cell(row, col)
    from openpyxl.cell.cell import MergedCell
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return ws.cell(rng.min_row, rng.min_col)
    return cell


def _mark_missing_date_cell(ws, row: int, col: int, formula_cols: set[int]) -> None:
    if not col or col in formula_cols:
        return
    cell = _resolve_target_cell(ws, row, col)
    if cell.value is None or str(cell.value).strip() == "":
        cell.fill = copy.copy(_MISSING_DATE_FILL)


def _clear_missing_date_cell(ws, row: int, col: int, formula_cols: set[int]) -> None:
    if not col or col in formula_cols:
        return
    cell = _resolve_target_cell(ws, row, col)
    if getattr(cell.fill, "start_color", None) and str(getattr(cell.fill.start_color, "rgb", "")).upper().endswith("FEF3C7"):
        cell.fill = copy.copy(_CLEAR_CELL_FILL)


def _is_pending_status_text(st: str) -> bool:
    sl = str(st or "").strip().lower()
    if not sl:
        return True
    return any(k in sl for k in ("pending", "under review", "open", "submitted", "waiting", "انتظار", "قيد المراجع", "معلق"))


def _strip_missing_flag(text: str, flag_phrase: str) -> str:
    if not text:
        return ""
    parts = [p.strip() for p in re.split(r"\s*\|\s*", text) if p.strip()]
    kept = [p for p in parts if flag_phrase not in p]
    return " | ".join(kept)


def _safe_set_cell(ws, row: int, col: int, value: Any, formula_cols: set[int], *, is_money: bool = False, is_date: bool = False) -> None:
    """Write `value` to `(row, col)` while NEVER overwriting any live formula cell."""
    if col in formula_cols:
        return
    cell = _resolve_target_cell(ws, row, col)
    if isinstance(cell.value, str) and cell.value.startswith("="):
        return
    if is_date and isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
        try:
            cell.value = dt.datetime.strptime(value.strip(), "%Y-%m-%d").date()
            if not cell.number_format or cell.number_format == "General":
                cell.number_format = "yyyy-mm-dd"
            if getattr(cell.fill, "start_color", None) and str(getattr(cell.fill.start_color, "rgb", "")).upper().endswith("FEF3C7"):
                cell.fill = copy.copy(_CLEAR_CELL_FILL)
            return
        except ValueError:
            pass
    if is_money and value not in (None, ""):
        try:
            num = float(str(value).replace(",", "").strip())
            cell.value = round(num, 3)
            if not cell.number_format or cell.number_format == "General":
                cell.number_format = "#,##0.000"
            return
        except ValueError:
            pass
    cell.value = value


_LOG_LOCKS_GUARD = threading.Lock()
_LOG_LOCKS: dict[str, threading.RLock] = {}


def _lock_for_path(p: Path | str) -> threading.RLock:
    key = str(resolve_local_or_unc(p)).lower()
    with _LOG_LOCKS_GUARD:
        lk = _LOG_LOCKS.get(key)
        if lk is None:
            lk = threading.RLock()
            _LOG_LOCKS[key] = lk
        return lk


def _save_pre_edit_snapshot(target_path: Path, proj_root: Path | str | None = None) -> None:
    if not target_path.exists() or not target_path.is_file():
        return
    try:
        base_bkp = (Path(proj_root).resolve() / "_Backups" / "Pre_Edit") if proj_root else (target_path.parent / "_Backups" / "Pre_Edit")
        base_bkp.mkdir(parents=True, exist_ok=True)
        existing = sorted(base_bkp.glob(f"*_{target_path.name}"), key=lambda p: p.stat().st_mtime, reverse=True)
        now_ts = dt.datetime.now().timestamp()
        if existing and (now_ts - existing[0].stat().st_mtime) < 300:
            return
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(target_path, base_bkp / f"{stamp}_{target_path.name}")
        for old_p in existing[9:]:
            try:
                old_p.unlink()
            except OSError:
                pass
    except OSError:
        pass


def _save_workbook_atomically(wb, target_path: Path, *, proj_root: Path | str | None = None) -> None:
    _save_pre_edit_snapshot(target_path, proj_root=proj_root)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{target_path.stem}_", suffix=target_path.suffix, dir=str(target_path.parent))
    os.close(fd)
    tmp_p = Path(tmp_name)
    try:
        wb.save(tmp_p)
        os.replace(tmp_p, target_path)
    finally:
        if tmp_p.exists():
            try:
                tmp_p.unlink()
            except OSError:
                pass


def _apply_entry_to_workbook(
    log_path: Path,
    lmap: dict[str, Any],
    sheet_spec: dict[str, Any],
    entry: dict[str, Any],
    *,
    device_id: str = "",
    proj_root: Path | str | None = None,
) -> dict[str, Any]:
    """Write or update a single record directly inside the user's uploaded project workbook."""
    ext = log_path.suffix.lower()
    with _lock_for_path(log_path):
        return _apply_entry_to_workbook_locked(
            log_path, lmap, sheet_spec, entry, ext=ext, device_id=device_id, proj_root=proj_root
        )


def _apply_entry_to_workbook_locked(
    log_path: Path,
    lmap: dict[str, Any],
    sheet_spec: dict[str, Any],
    entry: dict[str, Any],
    *,
    ext: str,
    device_id: str = "",
    proj_root: Path | str | None = None,
) -> dict[str, Any]:
    has_drawings = False
    wb = openpyxl.load_workbook(log_path, data_only=False, keep_vba=(ext == ".xlsm"))
    try:
        sname = sheet_spec["sheet"]
        if sname not in wb.sheetnames:
            raise KeyError(tr("validation.sheet_missing", sheet=sname))
        ws = wb[sname]
        layout = sheet_spec.get("layout", "flat")
        data_start = int(sheet_spec.get("data_start_row") or 2)
        cols = sheet_spec.get("columns") or {}
        rev_groups = sheet_spec.get("rev_groups") or []
        formula_cols = {int(x) for x in (sheet_spec.get("formula_cols") or [])}
        vocab = sheet_spec.get("status_vocab") or {}
        moh_vocab = sheet_spec.get("moh_vocab") or {}
        max_col = max(ws.max_column or 1, 8)

        ref_no = str(entry.get("ref_no") or "").strip()
        ref_key = normalize_ref_key(ref_no)
        rev_num = parse_rev_number(entry.get("revision"))
        has_explicit_rev = str(entry.get("revision") or "").strip() != ""
        is_superseded = bool(entry.get("is_superseded"))
        authority = str(entry.get("authority") or "").strip()
        cons_reply_date = str(entry.get("consultant_reply_date") or "").strip()
        cons_action = str(entry.get("consultant_action") or "").strip()
        moh_reply_date = str(entry.get("moh_reply_date") or "").strip()
        moh_action = str(entry.get("moh_action") or "").strip()
        has_both = bool(entry.get("has_both_stamps")) or (
            _is_moh_authority(authority) and _has_consultant_stamp(authority, entry)
        ) or (bool(cons_reply_date) and bool(moh_reply_date))
        is_moh_auth = _is_moh_authority(authority) or bool(moh_action) or bool(moh_reply_date)
        is_moh_only = is_moh_auth and not has_both
        is_moh_update = is_moh_auth and ("moh_status" in cols or "moh_reply_date" in cols or bool(moh_reply_date))

        subject = str(entry.get("subject_ar") or entry.get("subject") or "").strip()
        subtype = str(entry.get("subtype") or "").strip()
        disc = canonical_discipline(str(entry.get("discipline") or ""))
        party = str(entry.get("party_name") or entry.get("party") or "").strip()
        doc_date = str(entry.get("submit_date") or entry.get("doc_date") or dt.date.today().isoformat()).strip()
        reply_date = str(entry.get("reply_date") or "").strip()
        raw_status = str(entry.get("approval_code") or entry.get("status") or "").strip()
        eff_cons_reply_date = cons_reply_date or ("" if is_moh_only else reply_date)
        eff_cons_status_raw = cons_action or ("" if is_moh_only else raw_status)
        eff_moh_reply_date = moh_reply_date or (reply_date if is_moh_only else "")
        eff_moh_status_raw = moh_action or raw_status

        mapped_status = _map_status_to_sheet_vocab(eff_cons_status_raw or raw_status, vocab, layout == "multi_rev")
        mapped_moh = _map_moh_status_to_vocab(eff_moh_status_raw, moh_vocab) if is_moh_update else ""
        rs_low = str(eff_cons_status_raw or raw_status).strip().lower()
        is_reply_update = bool(
            eff_cons_reply_date
            or eff_moh_reply_date
            or reply_date
            or rs_low in (
                "a", "b", "c", "d", "code a", "code b", "code c", "code d",
                "approved", "approved as noted", "approved (c)", "rejected", "withdrawn",
                "معتمد", "معتمد بملاحظات", "مرفوض",
            )
            or any(k in rs_low for k in ("(code a)", "(code b)", "(code c)", "(code d)", "revise", "approved as noted"))
        )
        amount = entry.get("amount")
        total_amount = entry.get("total_amount") if entry.get("total_amount") not in (None, "") else amount
        notes = str(entry.get("notes") or "").strip()
        rel_file = str(entry.get("file_rel_path") or entry.get("file_path") or "").strip()

        ref_col = int(cols["ref_no"]["col"]) if "ref_no" in cols else None
        subj_col = int(cols["subject"]["col"]) if "subject" in cols else None
        rev_col = int(cols["revision"]["col"]) if "revision" in cols else None

        # 1. Search for existing matching row by ref_no
        matched_row: int | None = None
        if ref_col and ref_key:
            for r in range(data_start, (ws.max_row or data_start) + 1):
                rv = ws.cell(r, ref_col).value
                if rv is None or (isinstance(rv, str) and rv.startswith("=")):
                    continue
                if normalize_ref_key(str(rv)) == ref_key:
                    if layout == "flat" and rev_col and has_explicit_rev and not is_reply_update:
                        row_rev = parse_rev_number(ws.cell(r, rev_col).value)
                        if row_rev != rev_num:
                            continue
                    matched_row = r
                    break

        mode = "updated" if matched_row is not None else "inserted"

        # Guard: never overwrite an active/newer row with a SUPERSEDED revision scan!
        if is_superseded and matched_row is not None:
            return {
                "ok": True,
                "queued": False,
                "mode": "skipped_superseded",
                "log_name": log_path.name,
                "sheet": sname,
                "row": matched_row,
                "ref_no": ref_no,
                "status": "",
                "device_id": device_id,
            }

        # 2. If inserting, find first truly blank data row (handling pre-formatted sheets & totals rows!)
        target_row = matched_row
        if target_row is None:
            totals_row: int | None = None
            last_data_row = data_start - 1
            first_blank_row: int | None = None
            scan_end = ws.max_row or data_start
            for r in range(data_start, scan_end + 1):
                if _is_totals_row(ws, r, max_col):
                    totals_row = r
                    break
                r_val = ws.cell(r, ref_col).value if ref_col else None
                s_val = ws.cell(r, subj_col).value if subj_col else None
                a_val = ws.cell(r, int(cols["amount"]["col"])).value if "amount" in cols else None
                has_content = any(
                    v is not None and str(v).strip() != "" and not (isinstance(v, str) and v.startswith("="))
                    for v in (r_val, s_val, a_val)
                )
                if has_content:
                    last_data_row = r
                elif first_blank_row is None and r > last_data_row:
                    first_blank_row = r

            if first_blank_row is not None and (totals_row is None or first_blank_row < totals_row):
                # Also ensure there isn't non-empty data below first_blank_row when gaps exist:
                # prefer first blank row AFTER last_data_row so sequential order stays clean
                after_last = last_data_row + 1
                if totals_row is None and after_last <= scan_end:
                    target_row = after_last
                elif totals_row is not None and after_last < totals_row:
                    target_row = after_last
                else:
                    target_row = first_blank_row
                # Copy style from previous data row if target row is unstyled
                if last_data_row >= data_start and target_row > scan_end:
                    _copy_row_style_and_formulas(ws, last_data_row, target_row, max_col)
            elif totals_row is not None:
                # Insert a brand-new row right above totals_row and expand SUM formulas!
                ws.insert_rows(totals_row, amount=1)
                target_row = totals_row
                new_totals = totals_row + 1
                style_src = max(data_start, target_row - 1)
                _copy_row_style_and_formulas(ws, style_src, target_row, max_col)
                _update_totals_formulas_after_insert(ws, new_totals, totals_row - 1, target_row, max_col)
            else:
                target_row = max(data_start, last_data_row + 1)
                if target_row > data_start:
                    _copy_row_style_and_formulas(ws, target_row - 1, target_row, max_col)

        if target_row is None:
            raise RuntimeError("Unable to resolve target row in sheet")

        # 3. Populate standard mapped columns
        if mode == "inserted":
            if "sn" in cols:
                sn_c = int(cols["sn"]["col"])
                cur_sn = ws.cell(target_row, sn_c).value
                if cur_sn is None or str(cur_sn).strip() == "":
                    prev_sn = ws.cell(target_row - 1, sn_c).value if target_row > data_start else 0
                    try:
                        next_sn = int(float(str(prev_sn).strip())) + 1
                    except Exception:
                        next_sn = max(1, target_row - data_start + 1)
                    _safe_set_cell(ws, target_row, sn_c, next_sn, formula_cols)
            if ref_col and ref_no:
                _safe_set_cell(ws, target_row, ref_col, ref_no, formula_cols)
            if "discipline" in cols and disc:
                _safe_set_cell(ws, target_row, int(cols["discipline"]["col"]), disc, formula_cols)
            if "subtype" in cols and subtype:
                _safe_set_cell(ws, target_row, int(cols["subtype"]["col"]), subtype, formula_cols)
            if subj_col and subject:
                _safe_set_cell(ws, target_row, subj_col, subject, formula_cols)
            if "party" in cols and party:
                _safe_set_cell(ws, target_row, int(cols["party"]["col"]), party, formula_cols)
        else:
            # Update missing descriptive fields without clobbering existing description
            if subj_col and subject and not ws.cell(target_row, subj_col).value:
                _safe_set_cell(ws, target_row, subj_col, subject, formula_cols)
            if "party" in cols and party and not ws.cell(target_row, int(cols["party"]["col"])).value:
                _safe_set_cell(ws, target_row, int(cols["party"]["col"]), party, formula_cols)

        # 4. Handle MOH status/date + consultant multi-revision vs flat status/dates
        #    When reply decision exists without a readable reply date, leave the date cell BLANK
        #    and mark it with soft amber highlight + flag in Remarks so the user fills it manually!
        cons_missing_flag = tr("log.missing_cons_date_flag")
        moh_missing_flag = tr("log.missing_moh_date_flag")
        cons_has_decision = bool(not is_moh_only and (eff_cons_status_raw or mapped_status) and not _is_pending_status_text(eff_cons_status_raw or mapped_status))
        moh_has_decision = bool(is_moh_update and mapped_moh and not _is_pending_status_text(eff_moh_status_raw or mapped_moh))
        cons_missing_date = False
        moh_missing_date = False

        if is_moh_update and mapped_moh and "moh_status" in cols:
            _safe_set_cell(ws, target_row, int(cols["moh_status"]["col"]), mapped_moh, formula_cols)
        if "moh_reply_date" in cols and eff_moh_reply_date:
            _safe_set_cell(ws, target_row, int(cols["moh_reply_date"]["col"]), eff_moh_reply_date, formula_cols, is_date=True)
            notes = _strip_missing_flag(notes, moh_missing_flag)
        elif eff_moh_reply_date and "notes" in cols:
            notes = _strip_missing_flag(notes, moh_missing_flag)
            moh_note = tr("log.moh_remark_date", date=eff_moh_reply_date) + (f" ({mapped_moh})" if mapped_moh else "")
            if moh_note not in notes:
                notes = f"{notes} | {moh_note}" if notes else moh_note
        elif moh_has_decision and not eff_moh_reply_date:
            moh_missing_date = True
            if "moh_reply_date" in cols:
                _mark_missing_date_cell(ws, target_row, int(cols["moh_reply_date"]["col"]), formula_cols)
            if moh_missing_flag not in notes:
                notes = f"{notes} | {moh_missing_flag}" if notes else moh_missing_flag

        recorded_rev = ""
        if is_moh_only and mode == "updated":
            # Pure MOH / Undersecretary stamp updates MOH columns/remarks ONLY — never clobber consultant SOOR REV columns!
            pass
        elif layout == "multi_rev" and rev_groups:
            def _grp_has_data(g: dict[str, Any]) -> bool:
                for ck in ("submit_col", "reply_col", "action_col"):
                    ci = g.get(ck)
                    if ci:
                        v = ws.cell(target_row, int(ci)).value
                        if v is not None and str(v).strip() != "":
                            return True
                return False

            def _cell_iso(ci: Any) -> str:
                if not ci:
                    return ""
                v = ws.cell(target_row, int(ci)).value
                if isinstance(v, dt.datetime):
                    return v.date().isoformat()
                if isinstance(v, dt.date):
                    return v.isoformat()
                if isinstance(v, str) and v.strip():
                    return v.strip()[:10]
                return ""

            first_grp = rev_groups[0]
            chosen_grp = next((g for g in rev_groups if int(g.get("rev", 0)) == rev_num), None)
            if mode == "inserted":
                # REV. 0 (first slot) is ALWAYS the first recorded submission of a new reference,
                # even when a contractor cover sheet prints Rev.01 / R01.
                chosen_grp = first_grp
            elif not _grp_has_data(first_grp):
                # Existing row with empty REV. 0 -> must populate REV. 0 before using REV. 1+.
                chosen_grp = first_grp
            elif not has_explicit_rev:
                if is_reply_update:
                    # Find latest revision that has a submit/reply/action entry
                    for g in reversed(rev_groups):
                        if _grp_has_data(g):
                            chosen_grp = g
                            break
                else:
                    # New submission without explicit rev -> pick first empty rev slot
                    for g in rev_groups:
                        if not _grp_has_data(g):
                            chosen_grp = g
                            break
            else:
                if chosen_grp is None:
                    chosen_grp = rev_groups[min(rev_num, len(rev_groups) - 1)]
                chosen_idx = rev_groups.index(chosen_grp)
                if chosen_idx > 0 and not _grp_has_data(chosen_grp):
                    prev_grp = rev_groups[chosen_idx - 1]
                    if not _grp_has_data(prev_grp):
                        # Earlier slot is still empty -> fill sequentially so no blank revision gap is created
                        for g in rev_groups[:chosen_idx]:
                            if not _grp_has_data(g):
                                chosen_grp = g
                                break
                    elif chosen_idx == 1:
                        # REV. 0 populated & REV. 1 empty while R01 arrived:
                        # if this R01 is the reply/update to the open 1-based submission already in REV. 0, update REV. 0!
                        r0_act = str(ws.cell(target_row, int(first_grp["action_col"])).value or "").strip().lower() if first_grp.get("action_col") else ""
                        r0_rep = _cell_iso(first_grp.get("reply_col"))
                        r0_sub = _cell_iso(first_grp.get("submit_col"))
                        r0_needs_resubmit = any(k in r0_act for k in ("reject", "مرفوض", "(c)", "revise", "code c", "code d"))
                        explicit_doc = str(entry.get("submit_date") or entry.get("doc_date") or "").strip()
                        same_sub = bool(r0_sub and explicit_doc and r0_sub == explicit_doc[:10])
                        diff_sub = bool(r0_sub and explicit_doc and r0_sub != explicit_doc[:10])
                        if is_reply_update and not r0_needs_resubmit and not diff_sub and (same_sub or not r0_rep or _is_pending_status_text(r0_act)):
                            chosen_grp = first_grp

            if chosen_grp is None:
                chosen_grp = first_grp
            recorded_rev = f"R{int(chosen_grp.get('rev', 0)):02d}"

            s_col = chosen_grp.get("submit_col")
            r_col = chosen_grp.get("reply_col")
            a_col = chosen_grp.get("action_col")
            if s_col and doc_date and (not is_moh_only) and (mode == "inserted" or not is_reply_update or not ws.cell(target_row, int(s_col)).value):
                _safe_set_cell(ws, target_row, int(s_col), doc_date, formula_cols, is_date=True)
            if r_col and eff_cons_reply_date and (not is_moh_only):
                _safe_set_cell(ws, target_row, int(r_col), eff_cons_reply_date, formula_cols, is_date=True)
                notes = _strip_missing_flag(notes, cons_missing_flag)
            elif r_col and cons_has_decision and not eff_cons_reply_date:
                existing_r = ws.cell(target_row, int(r_col)).value
                if existing_r is None or str(existing_r).strip() == "":
                    cons_missing_date = True
                    _mark_missing_date_cell(ws, target_row, int(r_col), formula_cols)
                    if cons_missing_flag not in notes:
                        notes = f"{notes} | {cons_missing_flag}" if notes else cons_missing_flag
            if a_col and mapped_status and (not is_moh_only) and (eff_cons_status_raw or not is_moh_auth):
                _safe_set_cell(ws, target_row, int(a_col), mapped_status, formula_cols)
        else:
            recorded_rev = f"R{rev_num:02d}" if isinstance(rev_num, int) else str(rev_num)
            if rev_col and (mode == "inserted" or has_explicit_rev):
                _safe_set_cell(ws, target_row, rev_col, recorded_rev, formula_cols)
            if "doc_date" in cols and doc_date and (mode == "inserted" or not is_reply_update):
                _safe_set_cell(ws, target_row, int(cols["doc_date"]["col"]), doc_date, formula_cols, is_date=True)
            if "reply_date" in cols and eff_cons_reply_date:
                _safe_set_cell(ws, target_row, int(cols["reply_date"]["col"]), eff_cons_reply_date, formula_cols, is_date=True)
                notes = _strip_missing_flag(notes, cons_missing_flag)
            elif "reply_date" in cols and cons_has_decision and not eff_cons_reply_date:
                r_c = int(cols["reply_date"]["col"])
                existing_r = ws.cell(target_row, r_c).value
                if existing_r is None or str(existing_r).strip() == "":
                    cons_missing_date = True
                    _mark_missing_date_cell(ws, target_row, r_c, formula_cols)
                    if cons_missing_flag not in notes:
                        notes = f"{notes} | {cons_missing_flag}" if notes else cons_missing_flag
            if "status" in cols and mapped_status and (not is_moh_only or "moh_status" not in cols):
                _safe_set_cell(ws, target_row, int(cols["status"]["col"]), mapped_status, formula_cols)

        # Financial & auxiliary columns
        if "amount" in cols and amount not in (None, ""):
            _safe_set_cell(ws, target_row, int(cols["amount"]["col"]), amount, formula_cols, is_money=True)
        if "total_amount" in cols and total_amount not in (None, "") and cols["total_amount"]["col"] != cols.get("amount", {}).get("col"):
            _safe_set_cell(ws, target_row, int(cols["total_amount"]["col"]), total_amount, formula_cols, is_money=True)
        if "payment_method" in cols and entry.get("payment_method"):
            _safe_set_cell(ws, target_row, int(cols["payment_method"]["col"]), str(entry["payment_method"]), formula_cols)
        if "notes" in cols:
            notes_col = int(cols["notes"]["col"])
            existing_note = str(ws.cell(target_row, notes_col).value or "").strip()
            if eff_cons_reply_date:
                existing_note = _strip_missing_flag(existing_note, cons_missing_flag)
            if eff_moh_reply_date:
                existing_note = _strip_missing_flag(existing_note, moh_missing_flag)
            if notes:
                final_note = notes
                if mode == "updated" and existing_note and existing_note not in notes:
                    if notes in existing_note:
                        final_note = existing_note
                    else:
                        final_note = f"{existing_note} | {notes}"
                _safe_set_cell(ws, target_row, notes_col, final_note, formula_cols)
            elif existing_note != str(ws.cell(target_row, notes_col).value or "").strip():
                _safe_set_cell(ws, target_row, notes_col, existing_note, formula_cols)
        raw_abs_file = str(entry.get("file_abs_path") or "").strip()
        if raw_abs_file:
            try:
                rel_file = Path(os.path.relpath(Path(raw_abs_file), log_path.parent)).as_posix()
            except ValueError:
                rel_file = raw_abs_file
        if "file_link" in cols and rel_file:
            fc = int(cols["file_link"]["col"])
            safe_rel = rel_file.replace('"', '""')
            existing_fv = ws.cell(target_row, fc).value
            if fc not in formula_cols and not (isinstance(existing_fv, str) and existing_fv.startswith("=")):
                ws.cell(target_row, fc).value = f'=HYPERLINK("{safe_rel}","فتح الملف")'
        if "device_id" in cols and device_id:
            _safe_set_cell(ws, target_row, int(cols["device_id"]["col"]), device_id, formula_cols)

        has_drawings = any(
            bool(getattr(wb[sn], "_charts", None)) or bool(getattr(wb[sn], "_images", None))
            for sn in wb.sheetnames
        )
        _save_workbook_atomically(wb, log_path, proj_root=proj_root)
    finally:
        try:
            wb.close()
        except Exception:
            pass

    # Refresh stored sha256 in logmap so our own write doesn't trigger a false drift check
    lmap["log_sha256"] = sha256_of_file(log_path)
    lmap["updated_at"] = now_iso()
    atomic_write_json(logmap_path_for(log_path, proj_root=proj_root), lmap)
    _SUMMARY_CACHE.pop(str(proj_root.resolve()), None)

    return {
        "ok": True,
        "queued": False,
        "mode": mode,
        "log_name": log_path.name,
        "sheet": sheet_spec["sheet"],
        "row": target_row,
        "ref_no": ref_no,
        "recorded_revision": recorded_rev,
        "status": mapped_moh if is_moh_update else mapped_status,
        "is_moh_update": is_moh_update,
        "needs_manual_reply_date": bool(cons_missing_date or moh_missing_date),
        "drawings_present": has_drawings,
        "device_id": device_id,
    }


def _enqueue_pending_op(proj_root: Path, entry: dict[str, Any], device_id: str, reason: str) -> dict[str, Any]:
    qp = pending_queue_path_for(proj_root)
    with _lock_for_path(qp):
        queue = read_json_safe(qp, [])
        if not isinstance(queue, list):
            queue = []
        op = {
            "id": f"LOP-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}-{secrets.token_hex(2).upper()}",
            "entry": dict(entry),
            "device_id": device_id,
            "reason": reason,
            "queued_at": now_iso(),
        }
        queue.append(op)
        atomic_write_json(qp, queue)
        return {"ok": True, "queued": True, "reason": reason, "op_id": op["id"]}


def record_entry_in_project_log(
    proj: dict[str, Any],
    entry: dict[str, Any],
    *,
    device_id: str = "",
    queue_if_locked: bool = True,
) -> dict[str, Any]:
    """
    Record (insert or update) a document entry directly in the project's OWN uploaded
    Excel log workbook (`.xlsx` or `.xlsm`). Never creates a forced generic system log!
    """
    proj_root = Path(str(proj.get("root") or "")).resolve()
    picked = _pick_target_log_and_sheet(proj_root, entry, proj=proj, device_id=device_id)
    if picked is None:
        if queue_if_locked:
            res = _enqueue_pending_op(
                proj_root,
                entry,
                device_id,
                tr("server.queued_no_log"),
            )
            res["no_log_uploaded"] = True
            return res
        return {"ok": False, "queued": False, "no_log_uploaded": True, "error": tr("server.no_log_uploaded")}

    log_path, lmap, sheet_spec = picked
    if is_file_locked(log_path, allow_readonly=False):
        if queue_if_locked:
            return _enqueue_pending_op(
                proj_root,
                entry,
                device_id,
                tr("server.queued_log_open", name=log_path.name),
            )
        raise PermissionError(tr("server.log_open_locked", name=log_path.name))

    try:
        res = _apply_entry_to_workbook(log_path, lmap, sheet_spec, entry, device_id=device_id, proj_root=proj_root)
        _append_audit_event(proj_root, {
            "event": f"log_{res['mode']}",
            "log_name": res["log_name"],
            "sheet": res["sheet"],
            "row": res["row"],
            "ref_no": res["ref_no"],
            "status": res["status"],
            "device_id": device_id,
        })
        return res
    except PermissionError:
        if queue_if_locked:
            return _enqueue_pending_op(proj_root, entry, device_id, tr("server.queued_log_open", name=log_path.name))
        raise


def flush_project_log_queue(proj: dict[str, Any], *, device_id: str = "") -> dict[str, Any]:
    """Flush all queued operations in `pending_log_ops.json` into the project's uploaded log(s)."""
    proj_root = Path(str(proj.get("root") or "")).resolve()
    qp = pending_queue_path_for(proj_root)
    with _lock_for_path(qp):
        queue = read_json_safe(qp, [])
        if not isinstance(queue, list) or not queue:
            return {"flushed": 0, "remaining": 0, "errors": []}

        flushed = 0
        remaining: list[dict[str, Any]] = []
        errors: list[str] = []
        for op in queue:
            ent = op.get("entry") or {}
            dev = str(op.get("device_id") or device_id)
            picked = _pick_target_log_and_sheet(proj_root, ent, proj=proj, device_id=dev)
            if picked is None:
                remaining.append(op)
                continue
            log_path, lmap, sheet_spec = picked
            if is_file_locked(log_path, allow_readonly=False):
                remaining.append(op)
                errors.append(tr("server.log_open_locked", name=log_path.name))
                continue
            try:
                res = _apply_entry_to_workbook(log_path, lmap, sheet_spec, ent, device_id=dev, proj_root=proj_root)
                _append_audit_event(proj_root, {
                    "event": "queue_flushed",
                    "log_name": res["log_name"],
                    "sheet": res["sheet"],
                    "row": res["row"],
                    "ref_no": res["ref_no"],
                    "device_id": dev,
                })
                flushed += 1
            except Exception as exc:
                remaining.append(op)
                errors.append(str(exc))

        atomic_write_json(qp, remaining)
        return {"flushed": flushed, "remaining": len(remaining), "errors": errors}


def pending_queue_count(proj_root: Path | str) -> int:
    q = read_json_safe(pending_queue_path_for(proj_root), [])
    return len(q) if isinstance(q, list) else 0


def record_optional_unrelated_entry(
    proj: dict[str, Any],
    *,
    file_name: str,
    new_name: str,
    action_taken: str,
    reason: str,
    entity: str = "",
    device_id: str = "",
) -> dict[str, Any]:
    """
    ONLY called when `proj.get('log_unrelated_sheet')` is explicitly True.
    Writes to a strictly separated sheet `99_Unrelated_غير_مرتبط` (never mixed with main sheets).
    """
    if not proj.get("log_unrelated_sheet"):
        return {"logged": False}
    proj_root = Path(str(proj.get("root") or "")).resolve()
    logs = list_project_logs(proj_root, device_id=device_id)
    target_path: Path | None = None
    for lg in logs:
        p = Path(lg["path"])
        if not is_file_locked(p, allow_readonly=False):
            target_path = p
            break
    if target_path is None:
        target_path = logs_dir_for_project(proj_root) / "_Unrelated_Files_Register.xlsx"
        if not target_path.exists():
            wb_new = openpyxl.Workbook()
            ws0 = wb_new.active
            ws0.title = UNRELATED_SHEET_NAME
            wb_new.save(target_path)
            wb_new.close()

    ext = target_path.suffix.lower()
    with _lock_for_path(target_path):
        wb = openpyxl.load_workbook(target_path, data_only=False, keep_vba=(ext == ".xlsm"))
        try:
            if UNRELATED_SHEET_NAME in wb.sheetnames:
                ws = wb[UNRELATED_SHEET_NAME]
            else:
                ws = wb.create_sheet(title=UNRELATED_SHEET_NAME)
                ws.sheet_view.rightToLeft = True
                headers = ["#", "التاريخ", "الملف الأصلي", "الاسم/المسار الجديد", "الجهة المكتشفة", "الإجراء", "السبب", "الجهاز"]
                hdr_fill = PatternFill("solid", fgColor="475569")
                hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                thin = Side(style="thin", color="CBD5E1")
                bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
                for ci, h in enumerate(headers, 1):
                    cell = ws.cell(1, ci, h)
                    cell.fill = hdr_fill
                    cell.font = hdr_font
                    cell.border = bdr
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            if ws.max_row == 1 and ws.cell(1, 1).value is None:
                ws.sheet_view.rightToLeft = True
                for ci, h in enumerate(["#", "التاريخ", "الملف الأصلي", "الاسم/المسار الجديد", "الجهة المكتشفة", "الإجراء", "السبب", "الجهاز"], 1):
                    ws.cell(1, ci, h)
            nr = (ws.max_row or 1) + 1
            vals = [nr - 1, now_iso(), file_name, new_name, entity, action_taken, reason, device_id]
            for ci, val in enumerate(vals, 1):
                ws.cell(nr, ci, val)
            _save_workbook_atomically(wb, target_path, proj_root=proj_root)
        finally:
            wb.close()

    mp = logmap_path_for(target_path, proj_root=proj_root)
    if mp.exists():
        lm = read_json_safe(mp, {})
        if isinstance(lm, dict):
            lm["log_sha256"] = sha256_of_file(target_path)
            atomic_write_json(mp, lm)
    return {"logged": True, "file": target_path.name, "sheet": UNRELATED_SHEET_NAME, "row": nr}


# ── Counter Recalculation Directly from Uploaded Project Logs ───────────────
_PREFIX_NUM_RE = re.compile(r"(?:^|[-_/\\\s])([A-Za-z]{2,8})[-_\s/]*0*(\d{1,6})\s*$")
_MULTI_SEG_RE = re.compile(r"^\s*([A-Za-z]{2,8})(?:[-_/][A-Za-z0-9]{1,8})*?[-_/\s]0*(\d{1,6})\s*$")
_AR_YEAR_SEQ_RE = re.compile(r"^\s*([\u0600-\u06FF]{1,5})\s*/\s*(20\d{2})\s*/\s*0*(\d{1,6})\s*$")
_AR_SEQ_YEAR_RE = re.compile(r"^\s*([\u0600-\u06FF]{1,5})\s*/\s*0*(\d{1,6})\s*/\s*(20\d{2})\s*$")


def recalculate_counters_from_logs(proj: dict[str, Any], *, device_id: str = "") -> dict[str, Any]:
    """
    Scan ALL uploaded project logs and compute true maximum numbers directly from the
    log rows (ignoring stale cached counters, merging with user-set `baseline_counters`).
    Next reference for every type is strictly `max(log_actual, baseline) + 1`!
    """
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))

    log_max: dict[str, int] = {k: 0 for k in STANDARD_COUNTER_KEYS}
    disc_max: dict[str, int] = {k: 0 for k in STANDARD_DISC_KEYS}
    disc_nums_seen: dict[str, set[int]] = {}
    arabic_series: dict[str, int] = {}
    custom_prefixes: dict[str, int] = {}
    total_rows_scanned = 0

    for ent in _iter_project_log_entries(proj_root):
        if not ent.get("exists"):
            continue
        f = Path(ent["path"])
        try:
            lmap = load_or_create_logmap(f, device_id=device_id, proj_root=proj_root)
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        except Exception:
            continue
        try:
            for sc in lmap.get("sheets", []):
                sname = sc.get("sheet")
                if not sname or sname not in wb.sheetnames or sc.get("layout") == "skip":
                    continue
                if sname == UNRELATED_SHEET_NAME:
                    continue
                cols = sc.get("columns") or {}
                if "ref_no" not in cols:
                    continue
                ref_c = int(cols["ref_no"]["col"])
                data_start = int(sc.get("data_start_row") or 2)
                routes = sc.get("routes") or []
                primary_sec = str(routes[0].get("section") or "") if routes else ""
                ws = wb[sname]

                for row in ws.iter_rows(min_row=data_start, min_col=ref_c, max_col=ref_c, values_only=True):
                    val = row[0] if row else None
                    if val is None:
                        continue
                    s = normalize_ar(str(val), keep_punct=True).strip()
                    if not s or s.startswith("="):
                        continue
                    total_rows_scanned += 1

                    # Arabic letter refs: ج/2026/24 or م/39/2026
                    m_ay = _AR_YEAR_SEQ_RE.match(s)
                    if m_ay:
                        key = f"{m_ay.group(1)}/{m_ay.group(2)}"
                        seq = int(m_ay.group(3))
                        arabic_series[key] = max(arabic_series.get(key, 0), seq)
                        if primary_sec == "outgoing":
                            log_max["OUT"] = max(log_max["OUT"], seq)
                        elif primary_sec == "incoming":
                            log_max["IN"] = max(log_max["IN"], seq)
                        continue
                    m_sy = _AR_SEQ_YEAR_RE.match(s)
                    if m_sy:
                        key = f"{m_sy.group(1)}/*/{m_sy.group(3)}"
                        seq = int(m_sy.group(2))
                        arabic_series[key] = max(arabic_series.get(key, 0), seq)
                        if primary_sec == "incoming":
                            log_max["IN"] = max(log_max["IN"], seq)
                        elif primary_sec == "outgoing":
                            log_max["OUT"] = max(log_max["OUT"], seq)
                        continue

                    # Latin prefix refs: IPC-005, MAR-EL-012, EL-211, HAVC-072, ELEV 176, S.H.C/2026-05-017
                    matched_prefix = False
                    for rx in (_PREFIX_NUM_RE, _MULTI_SEG_RE):
                        m = rx.search(s)
                        if not m:
                            continue
                        pfx = m.group(1).upper()
                        seq = int(m.group(2))
                        if seq > 99999:
                            continue
                        matched_prefix = True
                        # Also catch leading standard key in multi-segment refs like MAR-EL-015 or OUT-CON-008
                        lead_m = re.match(r"^\s*([A-Za-z]{2,8})[-_/]", s)
                        lead_pfx = lead_m.group(1).upper() if lead_m else pfx

                        for cand in ({pfx, lead_pfx}):
                            if cand in log_max:
                                log_max[cand] = max(log_max[cand], seq)
                            elif cand in DISC_CANON:
                                dc = DISC_CANON[cand]
                                if dc in disc_max:
                                    disc_max[dc] = max(disc_max[dc], seq)
                                    disc_nums_seen.setdefault(dc, set()).add(seq)
                            else:
                                custom_prefixes[cand] = max(custom_prefixes.get(cand, 0), seq)
                        break

                    if not matched_prefix:
                        # Trailing sequence in complex ref like S.H.C/2026-05-017
                        m_tail = re.search(r"[-/]0*(\d{1,5})\s*$", s)
                        if m_tail:
                            seq = int(m_tail.group(1))
                            if seq < 2000:
                                if primary_sec == "incoming":
                                    log_max["IN"] = max(log_max["IN"], seq)
                                elif primary_sec == "outgoing":
                                    log_max["OUT"] = max(log_max["OUT"], seq)
        finally:
            wb.close()

    # Merge with user-defined baseline_counters (for fresh/empty logs on existing ongoing projects)
    base_c = proj.get("baseline_counters") or {}
    base_d = proj.get("baseline_disc_counters") or {}
    effective_counters: dict[str, int] = {}
    for k, v in log_max.items():
        effective_counters[k] = max(v, int(base_c.get(k) or 0))
    for k, v in custom_prefixes.items():
        effective_counters[k] = max(v, int(base_c.get(k) or 0))
    for k, bv in base_c.items():
        ku = str(k).upper()
        effective_counters[ku] = max(effective_counters.get(ku, 0), int(bv or 0))

    effective_disc: dict[str, int] = {}
    for k, v in disc_max.items():
        effective_disc[k] = max(v, int(base_d.get(k) or 0))
    for k, bv in base_d.items():
        dc = canonical_discipline(str(k))
        effective_disc[dc] = max(effective_disc.get(dc, 0), int(bv or 0))

    non_empty_discs = [s for s in disc_nums_seen.values() if s]
    all_disc_union = set().union(*non_empty_discs) if non_empty_discs else set()
    sum_disc_unique = sum(len(s) for s in non_empty_discs)
    is_global_disc_series = (
        len(non_empty_discs) >= 2
        and sum_disc_unique >= 5
        and len(all_disc_union) == sum_disc_unique
    )
    global_disc_max = max(effective_disc.values() or [0])

    next_preview: dict[str, str] = {}
    for k in ("IPC", "VO", "PO", "OUT", "IN", "PC", "RPT", "MAR", "MIR", "IR", "WIR", "SD", "MS", "PQ"):
        next_preview[k] = f"{k}-{effective_counters.get(k, 0) + 1:03d}"
    next_disc_preview: dict[str, str] = {}
    next_global_disc_preview: dict[str, str] = {}
    for d in STANDARD_DISC_KEYS:
        if effective_disc.get(d, 0) > 0 or d in ("CIV", "EL", "PL", "HVAC", "FF", "ELEV"):
            next_disc_preview[d] = f"{d}-{effective_disc.get(d, 0) + 1:03d}"
            next_global_disc_preview[d] = f"{d}-{global_disc_max + 1:03d}"
    next_arabic_preview: dict[str, str] = {}
    for ak, av in arabic_series.items():
        if "/*/" in ak:
            p1, p2 = ak.split("/*/", 1)
            next_arabic_preview[ak] = f"{p1}/{av + 1:02d}/{p2}"
        else:
            next_arabic_preview[ak] = f"{ak}/{av + 1}"

    return {
        "log_max": log_max,
        "disc_max": disc_max,
        "global_disc_max": global_disc_max,
        "is_global_disc_series": is_global_disc_series,
        "custom_prefixes": custom_prefixes,
        "arabic_series": arabic_series,
        "effective_counters": effective_counters,
        "effective_disc": effective_disc,
        "next_preview": next_preview,
        "next_disc_preview": next_disc_preview,
        "next_global_disc_preview": next_global_disc_preview,
        "next_arabic_preview": next_arabic_preview,
        "rows_scanned": total_rows_scanned,
        "synced_at": now_iso(),
    }


def next_reference_from_logs(
    proj: dict[str, Any],
    *,
    prefix: str,
    discipline: str = "",
    use_discipline_series: bool = False,
    global_series: bool | None = None,
    pad: int = 3,
    device_id: str = "",
) -> tuple[str, int, dict[str, Any]]:
    """
    Compute the exact next reference string directly from the project's uploaded log(s)
    (`max(log, baseline) + 1`), unified across Template Studio, Section screens, and Letters!
    """
    sync = recalculate_counters_from_logs(proj, device_id=device_id)
    pfx = (prefix or "DOC").strip().upper()

    # Arabic letter series support (e.g. prefix 'ج/2026' or 'م/2026' -> 'م/*/2026')
    if pfx in sync["arabic_series"]:
        nxt = sync["arabic_series"][pfx] + 1
        if "/*/" in pfx:
            return pfx.replace("/*/", f"/{nxt}/"), nxt, sync
        return f"{pfx}/{nxt}", nxt, sync
    if "/" in pfx:
        parts_ar = [x.strip() for x in pfx.split("/") if x.strip()]
        if len(parts_ar) == 2:
            star_key = f"{parts_ar[0]}/*/{parts_ar[1]}"
            if star_key in sync["arabic_series"]:
                nxt = sync["arabic_series"][star_key] + 1
                return f"{parts_ar[0]}/{nxt}/{parts_ar[1]}", nxt, sync

    if global_series is not None:
        use_global = bool(global_series)
    elif proj.get("global_submittal_numbering") is not None:
        use_global = bool(proj.get("global_submittal_numbering"))
    else:
        use_global = bool(sync.get("is_global_disc_series"))

    dc = canonical_discipline(discipline) if discipline else ""
    if use_discipline_series and dc:
        nxt = (sync["global_disc_max"] if use_global else sync["effective_disc"].get(dc, 0)) + 1
        return f"{dc}-{nxt:0{pad}d}", nxt, sync
    if pfx in sync["effective_disc"] and sync["effective_disc"][pfx] > 0 and pfx not in STANDARD_COUNTER_KEYS:
        nxt = (sync["global_disc_max"] if use_global else sync["effective_disc"][pfx]) + 1
        return f"{pfx}-{nxt:0{pad}d}", nxt, sync

    # If project log uses discipline numbering (like EL-211, CIV-205) and user picks discipline on a submittal with zero MAR-* refs
    if (
        dc
        and pfx in ("MAR", "MIR", "IR", "WIR", "SD", "MS", "PQ")
        and sync["effective_counters"].get(pfx, 0) == 0
        and sync["effective_disc"].get(dc, 0) > 0
    ):
        nxt = (sync["global_disc_max"] if use_global else sync["effective_disc"][dc]) + 1
        return f"{dc}-{nxt:0{pad}d}", nxt, sync

    nxt = sync["effective_counters"].get(pfx, 0) + 1
    return f"{pfx}-{nxt:0{pad}d}", nxt, sync


_SUMMARY_CACHE: dict[str, tuple[tuple[Any, ...], dict[str, Any]]] = {}


# ── Read Rows & KPI Stats Directly from Uploaded Project Logs ───────────────
_PREVIEWABLE_LOG_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tif", ".tiff", ".docx", ".xlsx", ".xlsm"}


def _build_archive_file_index(proj: dict[str, Any], proj_root: Path) -> dict[str, tuple[str, str, str]]:
    """
    Build a fast lookup `{normalized_ref_key: (rel_path, area, ext)}` from the project archive
    so log rows immediately gain Download and Inline Preview buttons.
    Prefers `.pdf` over editable Office files when both share the same reference.
    """
    arc_raw = str(proj.get("archive_dir") or "").strip()
    arc_root = Path(arc_raw).resolve() if arc_raw else (proj_root / "2_Archive").resolve()
    if not arc_root.exists() or not arc_root.is_dir():
        return {}
    try:
        arc_root.relative_to(proj_root)
        area = "root"
        rel_base = proj_root
    except ValueError:
        area = "archive"
        rel_base = arc_root

    best: dict[str, tuple[int, float, str, str, str]] = {}
    scanned = 0
    for fp in arc_root.rglob("*"):
        if scanned >= 3500:
            break
        if not fp.is_file() or fp.name.startswith("~$") or fp.name.startswith("."):
            continue
        if any(p.startswith("_System_Agent") or p == "_Backups" for p in fp.parts):
            continue
        scanned += 1
        ext = fp.suffix.lower()
        if ext.endswith(".json"):
            continue
        stem = fp.stem
        cands: list[str] = []
        head = stem.split("_")[0].strip()
        if head:
            cands.append(head)
        m_ref = re.match(r"^([A-Za-z\u0600-\u06FF]{1,8}(?:[-_/][A-Za-z0-9\u0600-\u06FF]{1,12}){1,4})", stem)
        if m_ref:
            cands.append(m_ref.group(1))
        try:
            rel_s = fp.relative_to(rel_base).as_posix()
            mtime = fp.stat().st_mtime
        except (ValueError, OSError):
            continue
        # Higher priority: PDF (3), Image (2), Office/Other (1)
        prio = 3 if ext == ".pdf" else (2 if ext in {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"} else 1)
        for cand in cands:
            rk = normalize_ref_key(cand)
            if not rk:
                continue
            prev = best.get(rk)
            if prev is None or (prio, mtime) > (prev[0], prev[1]):
                best[rk] = (prio, mtime, rel_s, area, ext)
    return {k: (v[2], v[3], v[4]) for k, v in best.items()}


def read_project_log_summary(
    proj: dict[str, Any],
    *,
    section_filter: str | None = None,
    query: str = "",
    limit: int = 200,
    recent_first: bool = False,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Read KPIs and document rows directly from the project's uploaded Excel log(s).
    Powers the main Dashboard, Section screens, and Search.
    Uses lightweight mtime/size caching when `limit == 0` (periodic card polling)
    so network-stored logs are not re-parsed every 12 seconds unless modified.
    """
    proj_root = Path(str(proj.get("root") or "")).resolve()
    logs = list_project_logs(proj_root, device_id=device_id)
    pq_count = pending_queue_count(proj_root)
    ov_days = int(proj.get("overdue_days") or 0) or 14
    links_p = logs_dir_for_project(proj_root) / "_letter_links.json"
    links_ns = links_p.stat().st_mtime_ns if links_p.exists() else 0

    cache_key: str | None = None
    cache_sig: tuple[Any, ...] = ()
    if limit == 0 and not query and not section_filter:
        sig_items: list[Any] = [ov_days, pq_count, links_ns, dt.date.today().isoformat()]
        for lg in logs:
            if lg.get("missing"):
                continue
            lp = Path(lg["path"])
            try:
                st = lp.stat()
                lm_p = logmap_path_for(lp, proj_root=proj_root)
                lm_ns = lm_p.stat().st_mtime_ns if lm_p.exists() else 0
                sig_items.append((lp.name, st.st_mtime_ns, st.st_size, lm_ns))
            except OSError:
                sig_items.append((lp.name, 0, 0, 0))
        cache_key = str(proj_root)
        cache_sig = tuple(sig_items)
        cached = _SUMMARY_CACHE.get(cache_key)
        if cached is not None and cached[0] == cache_sig:
            out_c = copy.deepcopy(cached[1])
            out_c["logs"] = logs
            return out_c

    arc_file_idx = _build_archive_file_index(proj, proj_root) if limit > 0 else {}

    # Load correspondence threads index cleanly (if present)
    raw_links = read_json_safe(links_p, {}) if links_p.exists() else {}
    threads_by_mk: dict[str, dict[str, Any]] = {}
    threads_by_rk: dict[str, dict[str, Any]] = {}
    if isinstance(raw_links, dict):
        for th in raw_links.get("threads") or []:
            if not isinstance(th, dict):
                continue
            for m in th.get("members") or []:
                if not isinstance(m, dict):
                    continue
                mk = str(m.get("member_key") or "")
                rk_m = str(m.get("ref_key") or normalize_ref_key(str(m.get("ref_no") or "")))
                if mk:
                    threads_by_mk[mk] = th
                if rk_m:
                    threads_by_rk[rk_m] = th

    def _corresp_party_lbl(sec_v: str, pr_v: str, sn_v: str, pt_v: str) -> tuple[str, str]:
        pr_c = str(pr_v or "").strip().lower()
        if not pr_c:
            comb = normalize_ar(f"{sn_v} {pt_v}")
            if any(k in comb for k in ("مقاولين الباطن", "مقاول الباطن", "مقاولي الباطن", "باطن", "subcontractor")):
                pr_c = "subcontractor"
            elif any(k in comb for k in ("السور", "استشاري", "الاستشاري", "مكتب", "consultant", "soor", "shc")):
                pr_c = "consultant"
            elif any(k in comb for k in ("وزاره", "الصحه", "الاشغال", "المالك", "وكيل", "moh", "mpw", "pahw", "mew", "ministry", "owner", "client")):
                pr_c = "owner"
            elif any(k in comb for k in ("مقاول", "المقاول", "contractor", "cont")):
                pr_c = "contractor"
        is_in_c = str(sec_v or "").strip().lower() == "incoming"
        lmap_k = {
            ("outgoing", "consultant"): "letters.role_out_cons",
            ("incoming", "consultant"): "letters.role_in_cons",
            ("outgoing", "subcontractor"): "letters.role_out_sub",
            ("incoming", "subcontractor"): "letters.role_in_sub",
            ("outgoing", "owner"): "letters.role_out_owner",
            ("incoming", "owner"): "letters.role_in_owner",
            ("outgoing", "contractor"): "letters.role_out_contr",
            ("incoming", "contractor"): "letters.role_in_contr",
        }
        return pr_c, tr(lmap_k.get(("incoming" if is_in_c else "outgoing", pr_c), "letters.role_in_gen" if is_in_c else "letters.role_out_gen"))

    stats = {
        "logs_count": len(logs),
        "has_uploaded_log": len(logs) > 0,
        "pending_queue": pq_count,
        "total_rows": 0,
        "incoming": 0,
        "outgoing": 0,
        "unreplied_incoming": 0,
        "submittals_total": 0,
        "approved": 0,
        "rejected": 0,
        "pending": 0,
        "withdrawn": 0,
        "overdue": 0,
        "overdue_days": ov_days,
        "moh_approved": 0,
        "moh_rejected": 0,
        "moh_pending": 0,
        "missing_reply_dates": 0,
        "ipc_count": 0,
        "ipc_total_kwd": 0.0,
        "vo_count": 0,
        "po_count": 0,
        "petty_count": 0,
        "petty_total_kwd": 0.0,
        "reports_count": 0,
    }
    rows_out: list[dict[str, Any]] = []
    missing_reply_rows: list[dict[str, Any]] = []
    q_norm = normalize_ar(query) if query else ""
    filter_base, filter_sub = _normalize_section_for_log(section_filter, "") if section_filter else ("", "")
    is_specific_sub_filter = bool(section_filter and section_filter.startswith("submittal_"))
    today = dt.date.today()

    for lg in logs:
        if lg.get("missing"):
            continue
        lp = Path(lg["path"])
        try:
            lmap = load_or_create_logmap(lp, device_id=device_id, proj_root=proj_root)
            wb = openpyxl.load_workbook(lp, read_only=True, data_only=True)
        except Exception:
            continue
        try:
            for sc in lmap.get("sheets", []):
                sname = sc.get("sheet")
                if not sc.get("enabled") or sc.get("layout") == "skip" or sname not in wb.sheetnames:
                    continue
                if sname == UNRELATED_SHEET_NAME:
                    continue
                routes = sc.get("routes") or []
                sec = str(routes[0].get("section") or "general") if routes else "general"
                cols = sc.get("columns") or {}
                rev_groups = sc.get("rev_groups") or []
                layout = sc.get("layout", "flat")
                data_start = int(sc.get("data_start_row") or 2)
                ws = wb[sname]

                def _g(r_tuple: tuple, role: str) -> Any:
                    ci = cols.get(role, {}).get("col")
                    if not ci or ci > len(r_tuple):
                        return None
                    return r_tuple[ci - 1]

                for r_idx, r_tuple in enumerate(
                    ws.iter_rows(min_row=data_start, max_col=min(ws.max_column or 60, 60), values_only=True),
                    start=data_start,
                ):
                    ref_v = _g(r_tuple, "ref_no")
                    subj_v = _g(r_tuple, "subject")
                    amt_v = _g(r_tuple, "total_amount") or _g(r_tuple, "amount")
                    ref_s = str(ref_v).strip() if ref_v is not None else ""
                    subj_s = str(subj_v).strip() if subj_v is not None else ""
                    if not ref_s and not subj_s:
                        continue
                    if normalize_ar(ref_s) in ("total", "الاجمالي", "المجموع") or normalize_ar(subj_s) in ("total", "الاجمالي", "المجموع"):
                        continue

                    sub_s = str(_g(r_tuple, "subtype") or "").strip()
                    disc_s = str(_g(r_tuple, "discipline") or "").strip()
                    if not disc_s and routes and routes[0].get("disciplines"):
                        disc_s = routes[0]["disciplines"][0]
                    party_s = str(_g(r_tuple, "party") or "").strip()
                    rev_s = str(_g(r_tuple, "revision") or "R00").strip()
                    date_v = _g(r_tuple, "doc_date")
                    reply_v = _g(r_tuple, "reply_date")
                    moh_reply_v = _g(r_tuple, "moh_reply_date")
                    stat_s = str(_g(r_tuple, "status") or "").strip()
                    moh_s = str(_g(r_tuple, "moh_status") or "").strip()
                    notes_s = str(_g(r_tuple, "notes") or "").strip()
                    sheet_overdue_days = int(proj.get("overdue_days") or sc.get("overdue_days") or 14)
                    if sc.get("overdue_days") and not proj.get("overdue_days") and sec == "submittal":
                        stats["overdue_days"] = int(sc["overdue_days"])

                    if layout == "multi_rev" and rev_groups:
                        latest_g = rev_groups[0]
                        for g in rev_groups:
                            sc_i = int(g.get("submit_col") or 0)
                            rc_i = int(g.get("reply_col") or 0)
                            ac_i = int(g.get("action_col") or 0)
                            sv = r_tuple[sc_i - 1] if sc_i and sc_i <= len(r_tuple) else None
                            rv = r_tuple[rc_i - 1] if rc_i and rc_i <= len(r_tuple) else None
                            av = r_tuple[ac_i - 1] if ac_i and ac_i <= len(r_tuple) else None
                            if (sv is not None and str(sv).strip() != "") or (rv is not None and str(rv).strip() != "") or (av is not None and str(av).strip() != ""):
                                latest_g = g
                        rev_s = f"R{int(latest_g.get('rev', 0)):02d}"
                        sc_i = int(latest_g.get("submit_col") or 0)
                        rc_i = int(latest_g.get("reply_col") or 0)
                        ac_i = int(latest_g.get("action_col") or 0)
                        date_v = r_tuple[sc_i - 1] if sc_i and sc_i <= len(r_tuple) else None
                        reply_v = r_tuple[rc_i - 1] if rc_i and rc_i <= len(r_tuple) else None
                        stat_s = str(r_tuple[ac_i - 1] or "").strip() if ac_i and ac_i <= len(r_tuple) else ""

                    def _fmt_d(dv: Any) -> tuple[str, dt.date | None]:
                        if isinstance(dv, dt.datetime):
                            return dv.date().isoformat(), dv.date()
                        if isinstance(dv, dt.date):
                            return dv.isoformat(), dv
                        if isinstance(dv, str) and dv.strip():
                            ds = dv.strip()[:10]
                            try:
                                return ds, dt.date.fromisoformat(ds)
                            except ValueError:
                                return dv.strip()[:20], None
                        return "", None

                    date_s, date_obj = _fmt_d(date_v)
                    reply_s, _ = _fmt_d(reply_v)
                    moh_reply_s, _ = _fmt_d(moh_reply_v)
                    if not moh_reply_s and notes_s:
                        m_mr = re.search(r"(?:20\d{2}-\d{2}-\d{2})", notes_s)
                        if m_mr and tr("log.missing_moh_date_flag") not in notes_s:
                            moh_reply_s = m_mr.group(0)

                    has_cons_reply_col = bool(
                        (layout == "multi_rev" and rev_groups)
                        or ("reply_date" in cols)
                    )
                    missing_cons_d = bool(
                        has_cons_reply_col
                        and not reply_s
                        and (
                            (stat_s and not _is_pending_status_text(stat_s))
                            or (tr("log.missing_cons_date_flag") in notes_s)
                        )
                    )
                    missing_moh_d = bool(
                        (
                            (moh_s and not _is_pending_status_text(moh_s) and not moh_reply_s and ("moh_reply_date" in cols or tr("log.missing_moh_date_flag") in notes_s))
                            or (tr("log.missing_moh_date_flag") in notes_s and not moh_reply_s)
                        )
                    )
                    if missing_cons_d or missing_moh_d:
                        stats["missing_reply_dates"] += 1
                        if len(missing_reply_rows) < 200:
                            missing_reply_rows.append({
                                "log_name": lp.name,
                                "sheet": sname,
                                "row": r_idx,
                                "ref_no": ref_s,
                                "revision": rev_s,
                                "subject": subj_s,
                                "status": stat_s,
                                "moh_status": moh_s,
                                "date": date_s,
                                "reply_date": reply_s,
                                "moh_reply_date": moh_reply_s,
                                "missing_consultant_date": missing_cons_d,
                                "missing_moh_date": missing_moh_d,
                                "has_moh_date_col": "moh_reply_date" in cols,
                            })
                    amt_f = 0.0
                    if amt_v not in (None, ""):
                        try:
                            amt_f = round(float(str(amt_v).replace(",", "").strip()), 3)
                        except ValueError:
                            amt_f = 0.0

                    # Accumulate KPIs
                    stats["total_rows"] += 1
                    row_sec = sec
                    up_sub = sub_s.upper()
                    if sec == "contract_ipc" and up_sub == "VO":
                        row_sec = "variation_vo"
                    elif sec == "contract_ipc" and up_sub == "PO":
                        row_sec = "purchase_po"

                    if row_sec == "incoming":
                        stats["incoming"] += 1
                    elif row_sec == "outgoing":
                        stats["outgoing"] += 1
                    elif row_sec == "submittal":
                        stats["submittals_total"] += 1
                        sl = stat_s.lower()
                        if "withdrawn" in sl or "cancel" in sl or "ملغي" in sl or "مسحوب" in sl:
                            stats["withdrawn"] += 1
                        elif any(k in sl for k in ("approved", "معتمد")) or sl in ("a", "b", "code a", "code b"):
                            stats["approved"] += 1
                        elif any(k in sl for k in ("reject", "مرفوض", "revise")) or sl in ("c", "d", "code c", "code d"):
                            stats["rejected"] += 1
                        else:
                            stats["pending"] += 1
                            if date_obj and not reply_s and (today - date_obj).days > sheet_overdue_days:
                                stats["overdue"] += 1
                        if moh_s:
                            ml = moh_s.lower()
                            if "approv" in ml or "معتمد" in ml:
                                stats["moh_approved"] += 1
                            elif "reject" in ml or "مرفوض" in ml:
                                stats["moh_rejected"] += 1
                            else:
                                stats["moh_pending"] += 1
                    elif row_sec == "contract_ipc":
                        stats["ipc_count"] += 1
                        stats["ipc_total_kwd"] = round(stats["ipc_total_kwd"] + amt_f, 3)
                    elif row_sec == "variation_vo":
                        stats["vo_count"] += 1
                    elif row_sec == "purchase_po":
                        stats["po_count"] += 1
                    elif row_sec == "petty_cash":
                        stats["petty_count"] += 1
                        stats["petty_total_kwd"] = round(stats["petty_total_kwd"] + amt_f, 3)
                    elif row_sec == "site_report":
                        stats["reports_count"] += 1

                    rel_letters_s = str(_g(r_tuple, "related_letters") or "").strip()
                    linked_items: list[dict[str, Any]] = []
                    th_obj: dict[str, Any] | None = None
                    row_pr = ""
                    row_plbl = ""
                    if row_sec in ("incoming", "outgoing"):
                        r_pr = str(routes[0].get("party_role") or "") if routes else ""
                        row_pr, row_plbl = _corresp_party_lbl(row_sec, r_pr, sname, party_s)
                        rk_row = normalize_ref_key(ref_s)
                        mk_row = f"{normalize_ar(sname)}::{rk_row}" if rk_row else f"{normalize_ar(sname)}::ROW:{r_idx}"
                        th_obj = threads_by_mk.get(mk_row) or (threads_by_rk.get(rk_row) if rk_row else None)
                        if th_obj:
                            parts_rl: list[str] = []
                            for om in th_obj.get("members") or []:
                                omk = str(om.get("member_key") or "")
                                if omk and omk == mk_row:
                                    continue
                                _or, olbl = _corresp_party_lbl(
                                    str(om.get("section") or ""),
                                    str(om.get("party_role") or ""),
                                    str(om.get("sheet") or ""),
                                    str(om.get("party") or ""),
                                )
                                oref = str(om.get("ref_no") or "").strip()
                                osubj = str(om.get("subject") or "").strip()
                                linked_items.append({
                                    "ref_no": oref,
                                    "sheet": str(om.get("sheet") or ""),
                                    "section": str(om.get("section") or ""),
                                    "party_role": _or,
                                    "party_label": olbl,
                                    "subject": osubj,
                                    "doc_date": str(om.get("doc_date") or ""),
                                    "thread_id": th_obj.get("id", ""),
                                })
                                if oref:
                                    parts_rl.append(f"{olbl}: {oref}")
                            if not rel_letters_s and parts_rl:
                                rel_letters_s = " | ".join(parts_rl)

                    inc_replied = False
                    inc_unreplied = False
                    days_waiting: int | None = None
                    if row_sec == "incoming":
                        has_out_link = any(str(m.get("section") or "").lower() == "outgoing" for m in linked_items)
                        rn_inc = normalize_ar(reply_s)
                        sn_inc = normalize_ar(stat_s)
                        replied_kw = any(
                            k in f"{rn_inc} {sn_inc}"
                            for k in ("تم الرد", "replied", "answered", "closed", "done", "مغلق")
                        )
                        inc_replied = bool(
                            has_out_link
                            or replied_kw
                            or (rn_inc and not _is_pending_status_text(reply_s))
                            or (sn_inc and not _is_pending_status_text(stat_s) and sn_inc not in ("وارد", "مستلم", "received", "incoming"))
                        )
                        inc_unreplied = not inc_replied
                        if inc_unreplied:
                            stats["unreplied_incoming"] += 1
                            if date_obj is not None:
                                days_waiting = max(0, (today - date_obj).days)

                    # Filter for requested view/search
                    if filter_base:
                        if row_sec != filter_base:
                            continue
                        if is_specific_sub_filter and filter_sub and up_sub:
                            equiv = {
                                "MAR": {"MAR", "SM", "OT", "MD", "MAT"},
                                "MIR": {"MIR"},
                                "IR": {"IR", "WIR", "CHECKLIST"},
                                "SD": {"SD", "DRG"},
                                "MS": {"MS", "PQ", "COMPANY PROFILE"},
                            }
                            allowed = equiv.get(filter_sub, {filter_sub})
                            if up_sub not in allowed and not ref_s.upper().startswith(filter_sub):
                                continue

                    if q_norm:
                        hay = normalize_ar(f"{ref_s} {subj_s} {party_s} {sub_s} {disc_s} {stat_s} {sname} {rel_letters_s}")
                        if q_norm not in hay:
                            continue

                    eff_status = stat_s
                    if not eff_status:
                        if row_sec == "incoming" and inc_replied:
                            eff_status = tr("letters.replied_status_val")
                        elif not reply_s:
                            eff_status = "pending"

                    if limit > 0 and (recent_first or len(rows_out) < limit):
                        row_rk = normalize_ref_key(ref_s)
                        f_match = arc_file_idx.get(row_rk) if row_rk else None
                        f_rel = ""
                        f_area = "root"
                        f_ext = ""
                        if f_match:
                            f_rel, f_area, f_ext = f_match
                        else:
                            fl_val = str(_g(r_tuple, "file_link") or "").strip()
                            if fl_val and ("/" in fl_val or "\\" in fl_val or "." in fl_val) and not fl_val.startswith("=HYPERLINK"):
                                clean_fl = fl_val.replace("\\", "/").lstrip("./")
                                while clean_fl.startswith("../"):
                                    clean_fl = clean_fl[3:]
                                cand_p = (proj_root / clean_fl).resolve()
                                if cand_p.exists() and cand_p.is_file():
                                    try:
                                        f_rel = cand_p.relative_to(proj_root).as_posix()
                                        f_area = "root"
                                        f_ext = cand_p.suffix.lower()
                                    except ValueError:
                                        pass
                        rows_out.append({
                            "log_name": lp.name,
                            "sheet": sname,
                            "row": r_idx,
                            "section": row_sec,
                            "ref_no": ref_s,
                            "revision": rev_s,
                            "discipline": disc_s,
                            "subtype": sub_s,
                            "subject": subj_s,
                            "party": party_s,
                            "party_role": row_pr,
                            "party_label": row_plbl,
                            "date": date_s,
                            "reply_date": reply_s,
                            "moh_reply_date": moh_reply_s,
                            "needs_manual_reply_date": bool(missing_cons_d or missing_moh_d),
                            "replied": inc_replied,
                            "unreplied": inc_unreplied,
                            "days_waiting": days_waiting,
                            "status": eff_status,
                            "moh_status": moh_s,
                            "related_letters": rel_letters_s,
                            "linked_letters": linked_items,
                            "thread_id": th_obj.get("id", "") if th_obj else "",
                            "amount_kwd": amt_f if amt_f else None,
                            "file_rel": f_rel,
                            "file_area": f_area,
                            "file_ext": f_ext,
                            "is_previewable": bool(f_rel and f_ext in _PREVIEWABLE_LOG_EXTS),
                        })
        finally:
            wb.close()

    if recent_first and limit > 0 and rows_out:
        rows_out.sort(key=lambda r: (str(r.get("date") or ""), int(r.get("row") or 0)), reverse=True)
        if len(rows_out) > limit:
            if section_filter:
                rows_out = rows_out[:limit]
            else:
                by_sec: dict[str, list[dict[str, Any]]] = {}
                for r in rows_out:
                    by_sec.setdefault(str(r.get("section") or "general"), []).append(r)
                per_sec = max(12, limit // max(len(by_sec), 1))
                picked: list[dict[str, Any]] = []
                picked_ids: set[int] = set()
                for grp in by_sec.values():
                    for r in grp[:per_sec]:
                        picked.append(r)
                        picked_ids.add(id(r))
                if len(picked) < limit:
                    for r in rows_out:
                        if id(r) not in picked_ids:
                            picked.append(r)
                            picked_ids.add(id(r))
                            if len(picked) >= limit:
                                break
                picked.sort(key=lambda r: (str(r.get("date") or ""), int(r.get("row") or 0)), reverse=True)
                rows_out = picked[:limit]

    res_summary = {"stats": stats, "rows": rows_out, "missing_reply_rows": missing_reply_rows, "logs": logs}
    if cache_key is not None:
        _SUMMARY_CACHE[cache_key] = (cache_sig, copy.deepcopy(res_summary))
    return res_summary


def update_log_row_reply_date(
    proj: dict[str, Any],
    *,
    log_name: str,
    sheet: str,
    row: int,
    consultant_reply_date: str = "",
    moh_reply_date: str = "",
    revision: str = "",
    device_id: str = "",
) -> dict[str, Any]:
    """
    Fill missing reply date(s) manually in a specific log sheet/row, clear the
    warning amber fill on those cells, and strip the pending-date flag from Notes.
    Works for both uploaded and directly linked external Excel logs.
    """
    proj_root = resolve_local_or_unc(str(proj["root"]))
    lp, _is_linked = resolve_project_log_path(proj_root, log_name)
    if not lp.exists():
        raise FileNotFoundError(tr("server.log_not_found"))
    if is_file_locked(lp):
        raise PermissionError(tr("server.queued_log_open", name=lp.name))

    lmap = load_or_create_logmap(lp, device_id=device_id, proj_root=proj_root)
    sheet_spec = next((s for s in (lmap.get("sheets") or []) if s.get("sheet") == sheet), None)
    if not sheet_spec:
        raise KeyError(tr("validation.sheet_missing", sheet=sheet))

    cons_d = str(consultant_reply_date or "").strip()
    moh_d = str(moh_reply_date or "").strip()
    if not cons_d and not moh_d:
        raise ValueError(tr("server.not_found"))

    cols = sheet_spec.get("columns") or {}
    rev_groups = sheet_spec.get("rev_groups") or []
    layout = sheet_spec.get("layout", "flat")
    formula_cols = {int(x) for x in (sheet_spec.get("formula_cols") or [])}
    rev_num = parse_rev_number(revision)

    keep_vba = lp.suffix.lower() == ".xlsm"
    with _lock_for_path(lp):
        wb = openpyxl.load_workbook(lp, keep_vba=keep_vba)
        try:
            if sheet not in wb.sheetnames:
                raise KeyError(tr("validation.sheet_missing", sheet=sheet))
            ws = wb[sheet]

            if cons_d:
                if layout == "multi_rev" and rev_groups:
                    def _row_grp_used(g: dict[str, Any]) -> bool:
                        return any(
                            ws.cell(row, int(g[ck])).value not in (None, "")
                            for ck in ("submit_col", "reply_col", "action_col")
                            if g.get(ck)
                        )

                    chosen_grp = next((g for g in rev_groups if int(g.get("rev", 0)) == rev_num), None)
                    if chosen_grp is None or not _row_grp_used(chosen_grp):
                        for g in reversed(rev_groups):
                            if _row_grp_used(g):
                                chosen_grp = g
                                break
                    if chosen_grp is None:
                        chosen_grp = rev_groups[0]
                    r_col = int(chosen_grp.get("reply_col") or 0)
                    if r_col:
                        _safe_set_cell(ws, row, r_col, cons_d, formula_cols, is_date=True)
                        _clear_missing_date_cell(ws, row, r_col, formula_cols)
                elif "reply_date" in cols:
                    r_col = int(cols["reply_date"]["col"])
                    _safe_set_cell(ws, row, r_col, cons_d, formula_cols, is_date=True)
                    _clear_missing_date_cell(ws, row, r_col, formula_cols)

            if moh_d:
                if "moh_reply_date" in cols:
                    m_col = int(cols["moh_reply_date"]["col"])
                    _safe_set_cell(ws, row, m_col, moh_d, formula_cols, is_date=True)
                    _clear_missing_date_cell(ws, row, m_col, formula_cols)
                elif "notes" in cols:
                    n_col = int(cols["notes"]["col"])
                    cur_n = str(ws.cell(row, n_col).value or "").strip()
                    cur_n = _strip_missing_flag(cur_n, tr("log.missing_moh_date_flag"))
                    moh_s = str(ws.cell(row, int(cols["moh_status"]["col"])).value or "").strip() if "moh_status" in cols else ""
                    tag = tr("log.moh_remark_date", date=moh_d) + (f" ({moh_s})" if moh_s else "")
                    new_n = f"{cur_n} | {tag}" if (cur_n and tag not in cur_n) else (tag if not cur_n else cur_n)
                    _safe_set_cell(ws, row, n_col, new_n, formula_cols)

            if "notes" in cols:
                n_col = int(cols["notes"]["col"])
                cur_n = str(ws.cell(row, n_col).value or "").strip()
                cleaned = cur_n
                if cons_d:
                    cleaned = _strip_missing_flag(cleaned, tr("log.missing_cons_date_flag"))
                if moh_d:
                    cleaned = _strip_missing_flag(cleaned, tr("log.missing_moh_date_flag"))
                if cleaned != cur_n:
                    _safe_set_cell(ws, row, n_col, cleaned, formula_cols)

            _save_workbook_atomically(wb, lp, proj_root=proj_root)
        finally:
            try:
                wb.close()
            except Exception:
                pass

    lmap["log_sha256"] = sha256_of_file(lp)
    lmap["updated_at"] = now_iso()
    atomic_write_json(logmap_path_for(lp, proj_root=proj_root), lmap)
    _SUMMARY_CACHE.pop(str(proj_root.resolve()), None)
    return {"ok": True, "log_name": lp.name, "sheet": sheet, "row": row}


def preview_log_sheet(proj_root: Path | str, log_name: str, sheet_name: str = "", max_rows: int = 45) -> dict[str, Any]:
    """Return a clean tabular preview of an uploaded or directly linked project log sheet for the UI mapper."""
    lp, is_linked = resolve_project_log_path(proj_root, log_name)
    lmap = load_or_create_logmap(lp, proj_root=proj_root)
    wb = openpyxl.load_workbook(lp, read_only=True, data_only=True)
    try:
        if sheet_name and sheet_name in wb.sheetnames:
            target = sheet_name
        else:
            lmap_sheets = lmap.get("sheets") or []
            active_cand = next(
                (s.get("sheet") for s in lmap_sheets
                 if s.get("enabled") and s.get("layout") != "skip" and s.get("sheet") in wb.sheetnames),
                None,
            )
            if not active_cand:
                active_cand = next(
                    (s.get("sheet") for s in lmap_sheets
                     if s.get("layout") != "skip" and s.get("sheet") in wb.sheetnames),
                    None,
                )
            target = active_cand or wb.sheetnames[0]
        ws = wb[target]
        max_c = min(ws.max_column or 1, 35)
        grid: list[dict[str, Any]] = []
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_c, values_only=True), 1):
            cells = []
            for v in row:
                if v is None:
                    cells.append("")
                elif isinstance(v, (dt.datetime, dt.date)):
                    cells.append(v.isoformat()[:10])
                else:
                    cells.append(str(v)[:120])
            if any(c.strip() for c in cells):
                grid.append({"row": r_idx, "cells": cells})
        col_letters = [get_column_letter(c) for c in range(1, max_c + 1)]
        return {
            "log_name": lp.name,
            "is_linked": is_linked,
            "file_path": str(lp),
            "sheets": wb.sheetnames,
            "active_sheet": target,
            "col_letters": col_letters,
            "rows": grid,
            "logmap": lmap,
        }
    finally:
        wb.close()
