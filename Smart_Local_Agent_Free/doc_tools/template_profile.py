#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
template_profile.py — Uploaded-Only Template Studio profiles (`<name>.tplmap.json`),
word-boundary section suggestion (no silent fallback), 6 field roles (including repeating
rows & formula protection), anchor auto-rebinding + orphan-field guard, validation,
and 100% deterministic `fill_by_profile` for `.xlsx`, `.xlsm` (VBA intact), and `.docx`.

NOTE: There are ZERO built-in/system-suggested templates — only files uploaded by the
user into `<project_root>/4_Project_Templates/`.
"""

from __future__ import annotations

import copy
import datetime as dt
import os
import re
import secrets
import tempfile
from pathlib import Path
from typing import Any

import docx
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, coordinate_to_tuple, get_column_letter

from relevance import normalize_ar, tokenize_words
from security_utils import (
    archive_instead_of_delete,
    atomic_write_json,
    now_iso,
    read_json_safe,
    safe_resolve_within,
    safe_template_name,
    sanitize_upload_filename,
    sha256_of_file,
)
from i18n_ar import auto_source_label, section_label, tr

ALLOWED_TEMPLATE_EXTS: frozenset[str] = frozenset({".xlsx", ".xlsm", ".docx"})

# ── Canonical Project Sections (each gets its own dedicated UI screen) ──────
SECTIONS: list[dict[str, Any]] = [
    {"id": "submittal_mar", "label": tr("sec.submittal_mar"), "short": "MAR", "ref_prefix": "MAR", "archive_folder": "03_Submittals - الاعتمادات الفنية/MAR - اعتماد مواد", "log_section": "submittal", "subtype": "MAR"},
    {"id": "submittal_mir", "label": tr("sec.submittal_mir"), "short": "MIR", "ref_prefix": "MIR", "archive_folder": "03_Submittals - الاعتمادات الفنية/MIR - فحص مواد", "log_section": "submittal", "subtype": "MIR"},
    {"id": "submittal_ir", "label": tr("sec.submittal_ir"), "short": "WIR", "ref_prefix": "IR", "archive_folder": "03_Submittals - الاعتمادات الفنية/IR_WIR - فحص أعمال", "log_section": "submittal", "subtype": "IR"},
    {"id": "submittal_sd", "label": tr("sec.submittal_sd"), "short": "SD", "ref_prefix": "SD", "archive_folder": "03_Submittals - الاعتمادات الفنية/SD - مخططات تنفيذية", "log_section": "submittal", "subtype": "SD"},
    {"id": "submittal_ms", "label": tr("sec.submittal_ms"), "short": "MS/PQ", "ref_prefix": "MS", "archive_folder": "03_Submittals - الاعتمادات الفنية/MS_PQ - طريقة تنفيذ وتأهيل", "log_section": "submittal", "subtype": "MS"},
    {"id": "outgoing", "label": tr("sec.outgoing"), "short": tr("sec.outgoing"), "ref_prefix": "OUT", "archive_folder": "02_Outgoing - الصادر", "log_section": "outgoing", "subtype": "OUT"},
    {"id": "incoming", "label": tr("sec.incoming"), "short": tr("sec.incoming"), "ref_prefix": "IN", "archive_folder": "01_Incoming - الوارد", "log_section": "incoming", "subtype": "IN"},
    {"id": "contract_ipc", "label": tr("sec.contract_ipc"), "short": "IPC", "ref_prefix": "IPC", "archive_folder": "04_Contracts_and_IPCs - العقود والمستخلصات/IPC - المستخلصات", "log_section": "contract_ipc", "subtype": "IPC"},
    {"id": "variation_vo", "label": tr("sec.variation_vo"), "short": "VO", "ref_prefix": "VO", "archive_folder": "04_Contracts_and_IPCs - العقود والمستخلصات/VO - الأوامر التغييرية", "log_section": "variation_vo", "subtype": "VO"},
    {"id": "purchase_po", "label": tr("sec.purchase_po"), "short": "PO", "ref_prefix": "PO", "archive_folder": "04_Contracts_and_IPCs - العقود والمستخلصات/PO - أوامر الشراء", "log_section": "purchase_po", "subtype": "PO"},
    {"id": "petty_cash", "label": tr("sec.petty_cash"), "short": tr("sec.petty_cash"), "ref_prefix": "PC", "archive_folder": "05_Petty_Cash_and_Invoices - العهد والفواتير", "log_section": "petty_cash", "subtype": "PC"},
    {"id": "site_report", "label": tr("sec.site_report"), "short": tr("sec.site_report"), "ref_prefix": "RPT", "archive_folder": "06_Site_Reports_and_MOM - التقارير والمحاضر", "log_section": "site_report", "subtype": "RPT"},
    {"id": "general", "label": tr("sec.general"), "short": tr("sec.general"), "ref_prefix": "GEN", "archive_folder": "09_General - مستندات عامة", "log_section": "general", "subtype": "GEN"},
]

SECTIONS_BY_ID: dict[str, dict[str, Any]] = {s["id"]: s for s in SECTIONS}


def get_localized_sections(style: str | None = None) -> list[dict[str, Any]]:
    return [{**s, "label": section_label(s["id"], style=style)} for s in SECTIONS]


def get_project_sections(
    proj: dict[str, Any] | None = None,
    *,
    style: str | None = None,
    include_hidden: bool = False,
) -> list[dict[str, Any]]:
    """
    Return the effective ordered list of sections for a project.
    Merges default SECTIONS with any per-project `custom_sections` (renamed labels,
    custom `ref_prefix`, custom `archive_folder`, `hidden` flags, or newly added sections).
    """
    base = get_localized_sections(style=style)
    by_id: dict[str, dict[str, Any]] = {s["id"]: dict(s) for s in base}
    order: list[str] = [s["id"] for s in base]

    raw_custom = (proj or {}).get("custom_sections") if isinstance(proj, dict) else None
    if isinstance(raw_custom, list):
        for item in raw_custom:
            if not isinstance(item, dict):
                continue
            sid = str(item.get("id") or "").strip()
            if not sid:
                continue
            if sid in by_id:
                cur = by_id[sid]
                if str(item.get("label") or "").strip():
                    cur["label"] = str(item["label"]).strip()
                    cur["short"] = str(item.get("short") or cur["label"]).strip()[:18]
                if str(item.get("short") or "").strip():
                    cur["short"] = str(item["short"]).strip()[:18]
                if str(item.get("ref_prefix") or "").strip():
                    cur["ref_prefix"] = str(item["ref_prefix"]).strip().upper()
                if str(item.get("archive_folder") or "").strip():
                    cur["archive_folder"] = str(item["archive_folder"]).strip().strip("/\\")
                if "hidden" in item:
                    cur["hidden"] = bool(item.get("hidden"))
                cur["custom"] = True
            else:
                lbl = str(item.get("label") or sid).strip()
                pfx = re.sub(r"[^A-Z0-9_-]", "", str(item.get("ref_prefix") or "DOC").upper()) or "DOC"
                safe_folder = re.sub(r'[<>:"|?*]', "_", str(item.get("archive_folder") or f"10_{lbl}")).strip("/\\. ")
                if not safe_folder:
                    safe_folder = f"10_{sid}"
                by_id[sid] = {
                    "id": sid,
                    "label": lbl,
                    "short": str(item.get("short") or pfx or lbl)[:18],
                    "ref_prefix": pfx,
                    "archive_folder": safe_folder,
                    "log_section": str(item.get("log_section") or "general"),
                    "subtype": str(item.get("subtype") or pfx),
                    "hidden": bool(item.get("hidden", False)),
                    "custom": True,
                    "user_added": True,
                }
                order.append(sid)

    out = [by_id[sid] for sid in order if sid in by_id]
    if not include_hidden:
        out = [s for s in out if not s.get("hidden")]
    return out


def sections_map_for_project(
    proj: dict[str, Any] | None = None,
    project_sections: list[dict[str, Any]] | dict[str, dict[str, Any]] | None = None,
    *,
    style: str | None = None,
) -> dict[str, dict[str, Any]]:
    if isinstance(project_sections, dict):
        merged = {s["id"]: dict(s) for s in get_localized_sections(style=style)}
        merged.update(project_sections)
        return merged
    if isinstance(project_sections, list):
        merged = {s["id"]: dict(s) for s in get_project_sections({"custom_sections": project_sections}, style=style, include_hidden=True)}
        return merged
    return {s["id"]: s for s in get_project_sections(proj, style=style, include_hidden=True)}


VALID_ROLES: frozenset[str] = frozenset({
    "fixed",
    "manual",
    "auto_date",
    "auto_counter",
    "auto_project",
    "auto_system",
    "formula_protected",
    "repeating_table",
})

VALID_INPUT_TYPES: frozenset[str] = frozenset({
    "text",
    "textarea",
    "date",
    "number",
    "kwd",
    "dropdown",
    "approval_code",
    "auto_index",
    "formula_protected",
})

DEFAULT_APPROVAL_OPTIONS: list[str] = ["A", "B", "C", "D", "Under Review"]

_AUTO_DATE_KEYS: tuple[str, ...] = (
    "date.today_dmy",
    "date.today_iso",
    "date.today_dmy_dash",
    "date.month_year",
    "date.day_name_ar",
    "date.year",
)

_AUTO_PROJECT_KEYS: tuple[str, ...] = (
    "project.name",
    "project.code",
    "project.contract_no",
    "project.owner",
    "project.consultant",
    "project.contractor",
    "project.location",
    "project.engineer",
)

_AUTO_SYSTEM_KEYS: tuple[str, ...] = (
    "system.next_ref",
    "system.today_iso",
    "system.today_dmy",
    "system.today_dmy_dash",
    "system.month_year",
    "system.day_name_ar",
    "system.revision",
    "system.year",
)

AUTO_DATE_SOURCES: dict[str, str] = {k: auto_source_label(k) for k in _AUTO_DATE_KEYS}
AUTO_PROJECT_SOURCES: dict[str, str] = {k: auto_source_label(k) for k in _AUTO_PROJECT_KEYS}
AUTO_SYSTEM_SOURCES: dict[str, str] = {k: auto_source_label(k) for k in _AUTO_SYSTEM_KEYS}


def get_localized_auto_date_sources(style: str | None = None) -> dict[str, str]:
    return {k: auto_source_label(k, style=style) for k in _AUTO_DATE_KEYS}


def get_localized_auto_project_sources(style: str | None = None) -> dict[str, str]:
    return {k: auto_source_label(k, style=style) for k in _AUTO_PROJECT_KEYS}


def get_localized_auto_system_sources(style: str | None = None) -> dict[str, str]:
    return {k: auto_source_label(k, style=style) for k in _AUTO_SYSTEM_KEYS}


def format_counter_value(fld: dict[str, Any], override_num: int | None = None) -> str:
    """Format a template custom auto_counter field like `<prefix><001><suffix>`."""
    prefix = str(fld.get("counter_prefix") or "")
    suffix = str(fld.get("counter_suffix") or "")
    raw_next = override_num if override_num is not None else fld.get("counter_next", 1)
    try:
        num = max(0, int(raw_next))
    except (TypeError, ValueError):
        num = 1
    try:
        digits = max(1, min(8, int(fld.get("counter_digits") or 3)))
    except (TypeError, ValueError):
        digits = 3
    num_str = f"{num:0{digits}d}" if digits > 1 else str(num)
    return f"{prefix}{num_str}{suffix}"

_AR_DAY_NAMES = ("الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد")


class ProfileValidationError(ValueError):
    """Raised when user form input violates template profile constraints."""

    def __init__(self, message: str, errors: list[dict[str, str]] | None = None):
        super().__init__(message)
        self.errors = errors or []


class ProfileOrphanError(ValueError):
    """Raised when a modified template has orphan fields that could not be safely rebound."""

    def __init__(self, message: str, orphan_fields: list[dict[str, Any]] | None = None):
        super().__init__(message)
        self.orphan_fields = orphan_fields or []


class FormulaOverwriteError(ValueError):
    """Raised if anything attempts to write over a protected Excel formula cell."""


# ── Stage 0: Strict Word-Boundary Section Suggestion (NO Silent Fallback!) ──
_SECTION_TOKEN_RULES: tuple[tuple[str, tuple[str, ...], tuple[str, ...]], ...] = (
    # (section_id, exact_single_tokens, exact_multiword_phrases)
    ("variation_vo", ("vo", "vos", "variation"), ("variation order", "امر تغييري", "اوامر تغييريه")),
    ("purchase_po", ("po", "pos", "lpo"), ("purchase order", "امر شراء", "اوامر شراء", "طلب شراء")),
    ("contract_ipc", ("ipc", "ipcs", "مستخلص", "مستخلصات"), ("payment certificate", "interim payment", "شهاده دفع", "دفعه جاريه")),
    ("submittal_mar", ("mar",), ("material approval", "material submittal", "اعتماد مواد", "اعتماد ماده")),
    ("submittal_mir", ("mir",), ("material inspection", "فحص مواد", "استلام مواد")),
    ("submittal_ir", ("ir", "wir", "itr", "checklist"), ("inspection request", "work inspection", "طلب فحص", "فحص اعمال")),
    ("submittal_sd", ("sd", "drg"), ("shop drawing", "shop drawings", "مخطط تنفيذي", "مخططات تنفيذيه")),
    ("submittal_ms", ("ms", "pq", "pqq", "mos"), ("method statement", "prequalification", "طريقه تنفيذ", "تاهيل مقاول")),
    ("petty_cash", ("pc", "petty", "عهده", "العهد"), ("petty cash", "كشف عهده", "تسويه عهده", "سند صرف", "مصروفات موقع")),
    ("site_report", ("report", "reports", "mom", "تقرير", "محضر"), ("daily report", "weekly report", "monthly report", "تقرير يومي", "تقرير اسبوعي", "تقرير شهري", "محضر اجتماع")),
    ("outgoing", ("outgoing", "letter", "transmittal", "صادر", "كتاب", "خطاب"), ("official letter", "كتاب رسمي", "خطاب صادر", "مذكره صادره")),
    ("incoming", ("incoming", "وارد"), ("incoming letter", "كتاب وارد", "مذكره وارده")),
)


def suggest_template_section(filename: str) -> dict[str, Any]:
    """
    Non-binding word-boundary suggestion for an uploaded template filename.
    NEVER matches substrings inside words:
      - `Invoice_2026.xlsx`   -> None (never `VO`!)
      - `Monthly_Report.xlsx` -> `site_report` (never `PO`!)
      - `Proposal.docx`       -> None (never `PO`!)
      - `Import_Permit.xlsx`  -> None (never `PO`!)
      - `Support_Form.docx`   -> None (never `PO`!)
      - `PO_001.xlsx` / `أمر_شراء.xlsx` -> `purchase_po` (unconfirmed suggestion)
      - `VO-03.xlsx` / `أمر_تغييري.xlsx` -> `variation_vo` (unconfirmed suggestion)
      - `Random_Form.xlsx`    -> None (ZERO silent fallback to MAR!)
    """
    stem = Path(filename or "").stem
    tokens = tokenize_words(stem)
    # Strip trailing digits ONLY from uppercase-style short codes like `PO01`? No — keep strict whole tokens!
    tok_set = set(tokens)
    joined = " ".join(tokens)

    # Multi-word phrases first (most specific)
    for sec_id, _, phrases in _SECTION_TOKEN_RULES:
        for ph in phrases:
            ph_norm = normalize_ar(ph)
            if ph_norm and re.search(rf"(?<!\w){re.escape(ph_norm)}(?!\w)", joined):
                sec = SECTIONS_BY_ID[sec_id]
                return {
                    "suggested_section": sec_id,
                    "suggested_label": section_label(sec_id),
                    "ref_prefix": sec["ref_prefix"],
                    "confirmed": False,
                    "matched_token": ph,
                }

    # Exact whole-word tokens second
    for sec_id, single_toks, _ in _SECTION_TOKEN_RULES:
        for tk in single_toks:
            tk_norm = normalize_ar(tk)
            if tk_norm in tok_set:
                sec = SECTIONS_BY_ID[sec_id]
                return {
                    "suggested_section": sec_id,
                    "suggested_label": section_label(sec_id),
                    "ref_prefix": sec["ref_prefix"],
                    "confirmed": False,
                    "matched_token": tk,
                }

    return {
        "suggested_section": None,
        "suggested_label": "",
        "ref_prefix": "",
        "confirmed": False,
        "matched_token": None,
    }


# ── Uploaded Template Directory & Sidecar `.tplmap.json` Helpers ────────────
def templates_dir_for_project(proj_root: Path | str) -> Path:
    d = Path(proj_root) / "4_Project_Templates"
    d.mkdir(parents=True, exist_ok=True)
    return d


def tplmap_path_for(tpl_path: Path) -> Path:
    return tpl_path.parent / f"{tpl_path.name}.tplmap.json"


def resolve_project_template(proj_root: Path | str, tpl_name: str) -> Path:
    tdir = templates_dir_for_project(proj_root)
    clean = safe_template_name(tpl_name)
    p = safe_resolve_within(tdir, clean)
    if p.suffix.lower() not in ALLOWED_TEMPLATE_EXTS:
        raise ValueError(tr("server.tpl_ext_unsupported"))
    return p


def upload_project_template(
    proj_root: Path | str,
    raw_filename: str,
    raw_bytes: bytes,
    *,
    section: str = "",
    display_name: str = "",
    device_id: str = "",
    project_sections: list[dict[str, Any]] | dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """
    Upload a user's own `.xlsx`, `.xlsm`, or `.docx` template into `4_Project_Templates/`.
    Zero deletion: if a file with the same name exists, archives the old copy safely first.
    """
    if not raw_bytes or len(raw_bytes) < 100:
        raise ValueError(tr("server.tpl_corrupt"))
    safe_name = sanitize_upload_filename(raw_filename, allowed_exts=ALLOWED_TEMPLATE_EXTS)
    proj_root_p = Path(proj_root).resolve()
    tdir = templates_dir_for_project(proj_root_p)
    dest = safe_resolve_within(tdir, safe_name)
    ext = dest.suffix.lower()

    fd, tmp_s = tempfile.mkstemp(prefix=".upload_tpl_", suffix=ext, dir=str(tdir))
    tmp_p = Path(tmp_s)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(raw_bytes)
        # Verify valid Office container
        try:
            if ext in (".xlsx", ".xlsm"):
                wb = openpyxl.load_workbook(tmp_p, read_only=True)
                _ = wb.sheetnames
                wb.close()
            else:
                d = docx.Document(str(tmp_p))
                _ = len(d.paragraphs)
        except Exception as exc:
            raise ValueError(tr("server.tpl_corrupt")) from exc

        prior_profile = read_json_safe(tplmap_path_for(dest))
        if dest.exists():
            archive_instead_of_delete(
                dest,
                tdir / "_Archived_Templates",
                reason="replaced_by_new_upload",
                device_id=device_id,
            )
        os.replace(tmp_p, dest)
    finally:
        if tmp_p.exists():
            try:
                tmp_p.unlink()
            except OSError:
                pass

    sec_map = sections_map_for_project(project_sections=project_sections)
    suggestion = suggest_template_section(dest.name)
    # If replacing an existing mapped template, automatically verify & rebind its `.tplmap.json`!
    if isinstance(prior_profile, dict) and prior_profile.get("fields") is not None:
        profile = get_template_profile(proj_root_p, dest.name, device_id=device_id)
    elif section and section in sec_map:
        sec_info = sec_map[section]
        init_prof = build_initial_profile(dest, device_id=device_id)
        init_prof["section"] = section
        init_prof["section_confirmed"] = True
        init_prof["ref_prefix"] = sec_info["ref_prefix"]
        if display_name.strip():
            init_prof["display_name"] = display_name.strip()
        atomic_write_json(tplmap_path_for(dest), init_prof)
        profile = init_prof
    else:
        profile = build_initial_profile(dest, device_id=device_id)
        if display_name.strip():
            profile["display_name"] = display_name.strip()

    return {
        "ok": True,
        "name": dest.name,
        "ext": ext,
        "size_kb": round(dest.stat().st_size / 1024, 1),
        "suggestion": suggestion,
        "profile": profile,
    }


def archive_project_template(proj_root: Path | str, tpl_name: str, *, device_id: str = "") -> Path:
    """Zero-delete removal: moves template + sidecar into `4_Project_Templates/_Archived_Templates/`."""
    tpl_path = resolve_project_template(proj_root, tpl_name)
    arch_dir = tpl_path.parent / "_Archived_Templates"
    archived = archive_instead_of_delete(tpl_path, arch_dir, reason="user_archived_template", device_id=device_id)
    mp = tplmap_path_for(tpl_path)
    if mp.exists():
        archive_instead_of_delete(mp, arch_dir, reason="user_archived_tplmap", device_id=device_id)
    return archived


# ── Deterministic Suggestion of Fields, Formula Locks & Repeating Tables ────
_LABEL_HINT_MAP: tuple[tuple[tuple[str, ...], str, str, str, str], ...] = (
    # (normalized_label_phrases, key, role, source, input_type)
    (("اسم المشروع", "المشروع", "project name", "project title", "project"), "project_name", "manual", "", "text"),
    (("رقم العقد", "رقم المناقصه", "العقد", "contract no", "contract number", "tender no"), "contract_no", "manual", "", "text"),
    (("الجهه المالكه", "المالك", "صاحب العمل", "client", "owner", "employer"), "owner", "manual", "", "text"),
    (("الاستشاري", "المكتب الاستشاري", "المهندس المشرف", "consultant", "engineer"), "consultant", "manual", "", "text"),
    (("المقاول", "المقاول الرئيسي", "الشركه المنفذه", "contractor", "main contractor"), "contractor", "manual", "", "text"),
    (("الموقع", "موقع المشروع", "location", "site location"), "location", "manual", "", "text"),
    (("رقم المرجع", "المرجع", "رقم النموذج", "رقم الطلب", "رقم الكتاب", "رقم المستخلص", "رقم الامر", "رقم السند", "ref no", "reference no", "submittal no", "ir no", "wir no", "mar no", "mir no", "ipc no", "vo no", "po no", "voucher no"), "ref_no", "auto_counter", "", "text"),
    (("رقم المراجعه", "المراجعه", "الاصدار", "rev", "rev no", "revision"), "revision", "manual", "", "text"),
    (("التاريخ", "تاريخ التقديم", "تاريخ الطلب", "تاريخ الاصدار", "تاريخ الكتاب", "date", "submission date", "issue date", "request date"), "doc_date", "auto_date", "date.today_dmy", "date"),
    (("اليوم", "day"), "day_name", "auto_date", "date.day_name_ar", "text"),
    (("كود الاعتماد", "قرار الاستشاري", "حاله الاعتماد", "نتيجه الفحص", "approval code", "action code", "consultant decision", "status code"), "approval_code", "manual", "", "approval_code"),
    (("التخصص", "القسم الفني", "discipline", "trade"), "discipline", "manual", "", "dropdown"),
    (("الموضوع", "عنوان الطلب", "وصف الاعمال", "بيان الاعمال", "وصف الماده", "subject", "description", "work description", "material description"), "subject", "manual", "", "text"),
    (("المورد", "الشركه الصانعه", "مقاول الباطن", "المرسل اليه", "supplier", "manufacturer", "vendor", "subcontractor", "attention"), "party_name", "manual", "", "text"),
    (("المبلغ", "القيمه", "صافي المبلغ", "قيمه الدفعه", "قيمه العهده", "amount", "value", "net amount", "amount kwd"), "amount", "manual", "", "kwd"),
    (("ملاحظات", "الملاحظات", "التوصيات", "التفاصيل", "نص الكتاب", "remarks", "notes", "comments", "details"), "notes", "manual", "", "textarea"),
)

_PLACEHOLDER_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}|\[\[\s*([^\]]+?)\s*\]\]")


def _classify_label(raw_label: str) -> dict[str, Any] | None:
    clean = re.sub(r"[:：\-_/\\*]+", " ", str(raw_label or "")).strip()
    if not clean or len(clean) > 65:
        return None
    ln = normalize_ar(clean)
    if not ln or len(ln) < 2:
        return None
    for phrases, key, role, source, itype in _LABEL_HINT_MAP:
        for ph in phrases:
            pn = normalize_ar(ph)
            if ln == pn or (len(pn) >= 5 and re.search(rf"(?<!\w){re.escape(pn)}(?!\w)", ln)):
                opts: list[str] = []
                if itype == "approval_code":
                    opts = list(DEFAULT_APPROVAL_OPTIONS)
                elif key == "discipline":
                    opts = ["CIV", "ARCH", "STR", "EL", "PL", "HVAC", "FF", "ELEV"]
                return {
                    "key": key,
                    "label": clean[:50],
                    "role": role,
                    "source": source,
                    "input_type": itype,
                    "required": key in ("subject",),
                    "options": opts,
                }
    return None


def _resolve_merge_master(ws, row: int, col: int) -> tuple[int, int, str | None]:
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col, str(rng)
    return row, col, None


def _merge_span_at(ws, row: int, col: int) -> tuple[int, int, str | None]:
    for rng in ws.merged_cells.ranges:
        if rng.min_row == row and rng.min_col == col:
            return rng.max_row, rng.max_col, str(rng)
    return row, col, None


def _find_xlsx_anchor(ws, target_row: int, target_col: int) -> dict[str, Any]:
    """Find the nearest static non-empty label cell neighboring `(target_row, target_col)`."""
    # Search left, right (for RTL sheets!), and up within 3 cells
    offsets = [
        (0, -1), (0, 1), (-1, 0), (0, -2), (0, 2), (0, -3), (0, 3), (-2, 0),
    ]
    for dr, dc in offsets:
        nr, nc = target_row + dr, target_col + dc
        if nr < 1 or nc < 1 or nr > (ws.max_row or 1) or nc > (ws.max_column or 1):
            continue
        mr, mc, _ = _resolve_merge_master(ws, nr, nc)
        if mr == target_row and mc == target_col:
            continue
        val = ws.cell(mr, mc).value
        if isinstance(val, str) and not val.startswith("="):
            txt = val.strip()
            if 2 <= len(txt) <= 80 and not _PLACEHOLDER_RE.search(txt):
                return {
                    "label_text": txt,
                    "label_norm": normalize_ar(txt),
                    "label_coord": f"{get_column_letter(mc)}{mr}",
                    "offset_row": target_row - mr,
                    "offset_col": target_col - mc,
                }
    return {}


def suggest_fields_deterministic(tpl_path: Path) -> dict[str, Any]:
    """
    Scan an uploaded `.xlsx`/`.xlsm`/`.docx` template and return deterministic
    CANDIDATE suggestions (fields, protected formula cells, and repeating item tables).
    Never mutates the template; user reviews & saves in Template Studio.
    """
    tpl_path = Path(tpl_path)
    ext = tpl_path.suffix.lower()
    suggested_fields: list[dict[str, Any]] = []
    protected_cells: list[dict[str, Any]] = []
    suggested_tables: list[dict[str, Any]] = []
    used_keys: dict[str, int] = {}

    def _uniq_key(base: str) -> str:
        c = used_keys.get(base, 0) + 1
        used_keys[base] = c
        return base if c == 1 else f"{base}_{c}"

    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(tpl_path, data_only=False, keep_vba=(ext == ".xlsm"))
        try:
            for sname in wb.sheetnames:
                ws = wb[sname]
                is_rtl = bool(getattr(ws.sheet_view, "rightToLeft", False))
                max_r = min(ws.max_row or 1, 120)
                max_c = min(ws.max_column or 1, 35)
                claimed_coords: set[str] = set()

                # 1. Collect all formula cells as protected
                for r in range(1, max_r + 1):
                    for c in range(1, max_c + 1):
                        v = ws.cell(r, c).value
                        if isinstance(v, str) and v.startswith("="):
                            coord = f"{get_column_letter(c)}{r}"
                            protected_cells.append({"sheet": sname, "coord": coord, "formula": v[:160]})
                            claimed_coords.add(coord)

                # 2. Detect repeating line-item tables (header row with S/N + Desc + Qty/Unit/Amount)
                table_rows_claimed: set[int] = set()
                for r in range(1, max_r):
                    hdr_cols: list[dict[str, Any]] = []
                    has_desc = False
                    has_qty_or_amt = False
                    c = 1
                    while c <= max_c:
                        v = ws.cell(r, c).value
                        end_r, end_c, _ = _merge_span_at(ws, r, c)
                        if isinstance(v, str) and not v.startswith("=") and v.strip():
                            ht = v.strip()
                            hn = normalize_ar(ht)
                            col_let = get_column_letter(c)
                            if ht in ("#", "م", "ت") or hn in ("sn", "s n", "sr", "no", "item", "م", "ت", "تسلسل", "البند"):
                                hdr_cols.append({"key": "sn", "label": ht, "col": col_let, "input_type": "auto_index", "required": False})
                            elif any(w in hn for w in ("البيان", "الوصف", "اسم الماده", "البنود", "description", "item description", "particulars", "material")):
                                hdr_cols.append({"key": "description", "label": ht, "col": col_let, "input_type": "text", "required": True})
                                has_desc = True
                            elif hn in ("الوحده", "unit", "uom"):
                                hdr_cols.append({"key": "unit", "label": ht, "col": col_let, "input_type": "text", "required": False})
                            elif any(w in hn for w in ("الكميه", "العدد", "qty", "quantity")):
                                hdr_cols.append({"key": "qty", "label": ht, "col": col_let, "input_type": "number", "required": False})
                                has_qty_or_amt = True
                            elif any(w in hn for w in ("سعر الوحده", "السعر", "الفئه", "unit price", "rate", "price")):
                                hdr_cols.append({"key": "unit_price", "label": ht, "col": col_let, "input_type": "kwd", "required": False})
                                has_qty_or_amt = True
                            elif any(w in hn for w in ("الاجمالي", "المبلغ", "القيمه", "total", "amount")):
                                tmpl_v = ws.cell(r + 1, c).value
                                if isinstance(tmpl_v, str) and tmpl_v.startswith("="):
                                    hdr_cols.append({
                                        "key": "line_total",
                                        "label": ht,
                                        "col": col_let,
                                        "input_type": "formula_protected",
                                        "required": False,
                                        "formula_template": re.sub(rf"(\$?[A-Z]{{1,3}}){r + 1}(?!\d)", r"\1{row}", tmpl_v),
                                    })
                                else:
                                    hdr_cols.append({"key": "amount", "label": ht, "col": col_let, "input_type": "kwd", "required": False})
                                has_qty_or_amt = True
                            elif any(w in hn for w in ("ملاحظات", "remarks", "notes", "المواصفات", "spec", "المنشا", "origin", "المورد")):
                                k = f"col_{col_let.lower()}"
                                hdr_cols.append({"key": k, "label": ht, "col": col_let, "input_type": "text", "required": False})
                        c = max(c + 1, end_c + 1)

                    if has_desc and has_qty_or_amt and len(hdr_cols) >= 3:
                        tmpl_row = r + 1
                        totals_row: int | None = None
                        blank_slots = 0
                        for row_i in range(tmpl_row, min(tmpl_row + 40, max_r + 1)):
                            row_has_sum = any(
                                isinstance(ws.cell(row_i, cc).value, str)
                                and re.match(r"^\s*=\s*(?:SUM|SUBTOTAL)\s*\(", str(ws.cell(row_i, cc).value), re.I)
                                for cc in range(1, max_c + 1)
                            )
                            row_has_tot_lbl = any(
                                isinstance(ws.cell(row_i, cc).value, str)
                                and normalize_ar(str(ws.cell(row_i, cc).value)) in ("الاجمالي", "المجموع", "الاجمالي الكلي", "total", "grand total", "subtotal")
                                for cc in range(1, max_c + 1)
                            )
                            if row_has_sum or row_has_tot_lbl:
                                totals_row = row_i
                                break
                            blank_slots += 1
                            table_rows_claimed.add(row_i)
                        table_rows_claimed.add(r)
                        if totals_row:
                            table_rows_claimed.add(totals_row)
                        suggested_tables.append({
                            "id": f"tbl_{len(suggested_tables) + 1}",
                            "key": f"items_{len(suggested_tables) + 1}" if suggested_tables else "items",
                            "label": tr("validation.table_default", index=len(suggested_tables) + 1),
                            "target_kind": "xlsx_rows",
                            "sheet": sname,
                            "header_row": r,
                            "template_row": tmpl_row,
                            "preformatted_rows": max(1, blank_slots if totals_row else min(blank_slots, 5)),
                            "totals_row": totals_row,
                            "columns": hdr_cols,
                            "anchor": {
                                "header_row": r,
                                "headers": {col["col"]: col["label"] for col in hdr_cols[:5]},
                            },
                            "suggested": True,
                        })
                        break

                # 3. Detect `{{placeholder}}` cells and `Label -> Empty Neighbor` fields
                for r in range(1, max_r + 1):
                    if r in table_rows_claimed:
                        continue
                    for c in range(1, max_c + 1):
                        cell = ws.cell(r, c)
                        if isinstance(cell, MergedCell):
                            continue
                        val = cell.value
                        if not isinstance(val, str) or val.startswith("=") or not val.strip():
                            continue
                        coord = f"{get_column_letter(c)}{r}"
                        end_r, end_c, m_rng = _merge_span_at(ws, r, c)

                        m_ph = _PLACEHOLDER_RE.search(val)
                        if m_ph and coord not in claimed_coords:
                            ph_txt = (m_ph.group(1) or m_ph.group(2) or "").strip()
                            cls = _classify_label(ph_txt) or {
                                "key": re.sub(r"\W+", "_", normalize_ar(ph_txt))[:24] or f"field_{coord.lower()}",
                                "label": ph_txt,
                                "role": "manual",
                                "source": "",
                                "input_type": "text",
                                "required": False,
                                "options": [],
                            }
                            fk = _uniq_key(cls["key"])
                            suggested_fields.append({
                                "id": f"fld_{sname}_{coord}",
                                "key": fk,
                                "label": cls["label"],
                                "role": cls["role"],
                                "source": cls["source"],
                                "input_type": cls["input_type"],
                                "required": cls["required"],
                                "default_value": "",
                                "options": cls["options"],
                                "target": {"kind": "xlsx_cell", "sheet": sname, "coord": coord, "merge_range": m_rng},
                                "anchor": _find_xlsx_anchor(ws, r, c),
                                "suggested": True,
                            })
                            claimed_coords.add(coord)
                            continue

                        cls = _classify_label(val)
                        if not cls:
                            continue
                        # Find empty neighbor cell (right first for LTR or visual Excel sheets, also check left if RTL, then below)
                        neighbor_candidates = [(r, end_c + 1), (r, c - 1), (end_r + 1, c)]
                        if is_rtl and not re.search(r"[A-Za-z]", val):
                            neighbor_candidates = [(r, end_c + 1), (r, c - 1), (end_r + 1, c)]
                        for nr, nc in neighbor_candidates:
                            if nr < 1 or nc < 1 or nr > max_r or nc > max_c:
                                continue
                            mr, mc, nm_rng = _resolve_merge_master(ws, nr, nc)
                            ncoord = f"{get_column_letter(mc)}{mr}"
                            if ncoord in claimed_coords or (mr == r and mc == c):
                                continue
                            nv = ws.cell(mr, mc).value
                            is_empty_or_dots = (
                                nv is None
                                or (isinstance(nv, str) and not nv.startswith("=") and re.fullmatch(r"[\s\._\-/…:]*", nv))
                            )
                            if is_empty_or_dots:
                                fk = _uniq_key(cls["key"])
                                suggested_fields.append({
                                    "id": f"fld_{sname}_{ncoord}",
                                    "key": fk,
                                    "label": cls["label"],
                                    "role": cls["role"],
                                    "source": cls["source"],
                                    "input_type": cls["input_type"],
                                    "required": cls["required"],
                                    "default_value": "",
                                    "options": cls["options"],
                                    "target": {"kind": "xlsx_cell", "sheet": sname, "coord": ncoord, "merge_range": nm_rng},
                                    "anchor": {
                                        "label_text": val.strip(),
                                        "label_norm": normalize_ar(val.strip()),
                                        "label_coord": coord,
                                        "offset_row": mr - r,
                                        "offset_col": mc - c,
                                    },
                                    "suggested": True,
                                })
                                claimed_coords.add(ncoord)
                                break
        finally:
            wb.close()

    elif ext == ".docx":
        doc = docx.Document(str(tpl_path))
        # 1. Paragraphs (`{{placeholder}}` or `Label: ____`)
        for p_idx, para in enumerate(doc.paragraphs):
            txt = para.text.strip()
            if not txt:
                continue
            for m_ph in _PLACEHOLDER_RE.finditer(txt):
                full_tok = m_ph.group(0)
                inner = (m_ph.group(1) or m_ph.group(2) or "").strip()
                cls = _classify_label(inner) or {
                    "key": re.sub(r"\W+", "_", normalize_ar(inner))[:24] or f"para_{p_idx}",
                    "label": inner,
                    "role": "manual",
                    "source": "",
                    "input_type": "text",
                    "required": False,
                    "options": [],
                }
                fk = _uniq_key(cls["key"])
                suggested_fields.append({
                    "id": f"fld_p{p_idx}_{fk}",
                    "key": fk,
                    "label": cls["label"],
                    "role": cls["role"],
                    "source": cls["source"],
                    "input_type": cls["input_type"],
                    "required": cls["required"],
                    "default_value": "",
                    "options": cls["options"],
                    "target": {
                        "kind": "docx_span",
                        "para_idx": p_idx,
                        "match_text": full_tok,
                    },
                    "anchor": {"label_text": full_tok, "label_norm": normalize_ar(inner), "para_idx": p_idx},
                    "suggested": True,
                })
            if not _PLACEHOLDER_RE.search(txt) and re.search(r"[:：]\s*[\._…\-]{0,30}\s*$", txt):
                lbl_part = re.split(r"[:：]", txt, maxsplit=1)[0].strip()
                cls = _classify_label(lbl_part)
                if cls:
                    fk = _uniq_key(cls["key"])
                    suggested_fields.append({
                        "id": f"fld_p{p_idx}",
                        "key": fk,
                        "label": cls["label"],
                        "role": cls["role"],
                        "source": cls["source"],
                        "input_type": cls["input_type"],
                        "required": cls["required"],
                        "default_value": "",
                        "options": cls["options"],
                        "target": {
                            "kind": "docx_paragraph",
                            "para_idx": p_idx,
                            "replace_mode": "after_colon",
                            "prefix_text": lbl_part + ": ",
                        },
                        "anchor": {"label_text": lbl_part, "label_norm": normalize_ar(lbl_part), "para_idx": p_idx},
                        "suggested": True,
                    })

        # 2. Tables (label->empty cell pairs, placeholders, and repeating item rows!)
        for t_idx, tbl in enumerate(doc.tables):
            claimed_cells: set[tuple[int, int]] = set()
            # Keep strong references to `_tc` elements so Python never reuses `id()`!
            tc_rows = [[cell._tc for cell in row.cells] for row in tbl.rows]

            # Detect repeating table if row 0 has S/N + Description + Qty/Amount
            if len(tbl.rows) >= 2:
                r0 = tbl.rows[0]
                seen_tc0: set[int] = set()
                tcols: list[dict[str, Any]] = []
                has_desc = False
                has_qa = False
                for c_idx, cell in enumerate(r0.cells):
                    tc_id = id(tc_rows[0][c_idx])
                    if tc_id in seen_tc0:
                        continue
                    seen_tc0.add(tc_id)
                    ht = cell.text.strip()
                    hn = normalize_ar(ht)
                    if ht in ("#", "م", "ت") or hn in ("sn", "s n", "no", "item", "م", "ت", "تسلسل"):
                        tcols.append({"key": "sn", "label": ht, "col_idx": c_idx, "input_type": "auto_index", "required": False})
                    elif any(w in hn for w in ("البيان", "الوصف", "اسم الماده", "description", "item description", "particulars")):
                        tcols.append({"key": "description", "label": ht, "col_idx": c_idx, "input_type": "text", "required": True})
                        has_desc = True
                    elif hn in ("الوحده", "unit", "uom"):
                        tcols.append({"key": "unit", "label": ht, "col_idx": c_idx, "input_type": "text", "required": False})
                    elif any(w in hn for w in ("الكميه", "العدد", "qty", "quantity")):
                        tcols.append({"key": "qty", "label": ht, "col_idx": c_idx, "input_type": "number", "required": False})
                        has_qa = True
                    elif any(w in hn for w in ("السعر", "المبلغ", "الاجمالي", "القيمه", "price", "rate", "amount", "total")):
                        tcols.append({"key": f"amt_{c_idx}", "label": ht, "col_idx": c_idx, "input_type": "kwd", "required": False})
                        has_qa = True
                    elif ht:
                        tcols.append({"key": f"col_{c_idx}", "label": ht[:30], "col_idx": c_idx, "input_type": "text", "required": False})

                if has_desc and has_qa and len(tcols) >= 3:
                    totals_r: int | None = None
                    last_r_txt = normalize_ar(" ".join(c.text for c in tbl.rows[-1].cells))
                    if len(tbl.rows) >= 3 and any(w in last_r_txt for w in ("الاجمالي", "المجموع", "total")):
                        totals_r = len(tbl.rows) - 1
                    suggested_tables.append({
                        "id": f"dtbl_{t_idx}",
                        "key": f"items_{t_idx + 1}" if suggested_tables else "items",
                        "label": tr("validation.table_default", index=t_idx + 1),
                        "target_kind": "docx_rows",
                        "table_idx": t_idx,
                        "header_row": 0,
                        "template_row": 1,
                        "preformatted_rows": (totals_r - 1) if totals_r else (len(tbl.rows) - 1),
                        "totals_row": totals_r,
                        "columns": tcols,
                        "anchor": {"table_idx": t_idx, "header_row": 0, "headers": {str(c["col_idx"]): c["label"] for c in tcols[:4]}},
                        "suggested": True,
                    })
                    continue

            for r_idx, row in enumerate(tbl.rows):
                seen_tc: set[int] = set()
                unique_cols: list[int] = []
                for c_idx in range(len(row.cells)):
                    tid = id(tc_rows[r_idx][c_idx])
                    if tid not in seen_tc:
                        seen_tc.add(tid)
                        unique_cols.append(c_idx)

                for pos, c_idx in enumerate(unique_cols):
                    txt = row.cells[c_idx].text.strip()
                    if not txt:
                        continue
                    m_ph = _PLACEHOLDER_RE.search(txt)
                    if m_ph and (r_idx, c_idx) not in claimed_cells:
                        full_tok = m_ph.group(0)
                        inner = (m_ph.group(1) or m_ph.group(2) or "").strip()
                        cls = _classify_label(inner) or {
                            "key": re.sub(r"\W+", "_", normalize_ar(inner))[:24] or f"t{t_idx}_r{r_idx}_c{c_idx}",
                            "label": inner,
                            "role": "manual",
                            "source": "",
                            "input_type": "text",
                            "required": False,
                            "options": [],
                        }
                        fk = _uniq_key(cls["key"])
                        suggested_fields.append({
                            "id": f"fld_t{t_idx}_r{r_idx}_c{c_idx}",
                            "key": fk,
                            "label": cls["label"],
                            "role": cls["role"],
                            "source": cls["source"],
                            "input_type": cls["input_type"],
                            "required": cls["required"],
                            "default_value": "",
                            "options": cls["options"],
                            "target": {"kind": "docx_cell", "table_idx": t_idx, "row_idx": r_idx, "col_idx": c_idx, "replace_mode": "span", "match_text": full_tok},
                            "anchor": {"label_text": full_tok, "label_norm": normalize_ar(inner), "table_idx": t_idx, "label_row": r_idx, "label_col": c_idx, "offset_row": 0, "offset_col": 0},
                            "suggested": True,
                        })
                        claimed_cells.add((r_idx, c_idx))
                        continue

                    cls = _classify_label(txt)
                    if not cls:
                        continue
                    # Try adjacent unique cell (next column or previous column for RTL tables)
                    adj_cols = []
                    if pos + 1 < len(unique_cols):
                        adj_cols.append(unique_cols[pos + 1])
                    if pos - 1 >= 0:
                        adj_cols.append(unique_cols[pos - 1])
                    for nc in adj_cols:
                        if (r_idx, nc) in claimed_cells:
                            continue
                        ntxt = row.cells[nc].text.strip()
                        if not ntxt or re.fullmatch(r"[\s\._\-/…:]*", ntxt):
                            fk = _uniq_key(cls["key"])
                            suggested_fields.append({
                                "id": f"fld_t{t_idx}_r{r_idx}_c{nc}",
                                "key": fk,
                                "label": cls["label"],
                                "role": cls["role"],
                                "source": cls["source"],
                                "input_type": cls["input_type"],
                                "required": cls["required"],
                                "default_value": "",
                                "options": cls["options"],
                                "target": {"kind": "docx_cell", "table_idx": t_idx, "row_idx": r_idx, "col_idx": nc, "replace_mode": "cell"},
                                "anchor": {
                                    "label_text": txt,
                                    "label_norm": normalize_ar(txt),
                                    "table_idx": t_idx,
                                    "label_row": r_idx,
                                    "label_col": c_idx,
                                    "offset_row": 0,
                                    "offset_col": nc - c_idx,
                                },
                                "suggested": True,
                            })
                            claimed_cells.add((r_idx, nc))
                            break

    return {
        "suggested_fields": suggested_fields,
        "protected_cells": protected_cells,
        "suggested_tables": suggested_tables,
    }


def build_initial_profile(tpl_path: Path, *, device_id: str = "") -> dict[str, Any]:
    """Create an unsaved starter profile structure with deterministic suggestions attached."""
    tpl_path = Path(tpl_path)
    sugg_sec = suggest_template_section(tpl_path.name)
    det = suggest_fields_deterministic(tpl_path)
    return {
        "schema_version": "2.0",
        "template_name": tpl_path.name,
        "template_ext": tpl_path.suffix.lower(),
        "template_sha256": sha256_of_file(tpl_path),
        "display_name": tpl_path.stem.replace("_", " "),
        "section": sugg_sec["suggested_section"] or "",
        "section_confirmed": False,
        "section_suggestion": sugg_sec,
        "ref_prefix": sugg_sec["ref_prefix"] or "",
        "use_discipline_series": False,
        "default_discipline": "",
        "log_on_issue": True,
        "default_status": "Under Review",
        "has_saved_profile": False,
        "fields": [],
        "protected_cells": det["protected_cells"],
        "repeating_tables": [],
        "suggestions": {
            "fields": det["suggested_fields"],
            "repeating_tables": det["suggested_tables"],
        },
        "orphan_fields": [],
        "rebound_fields": [],
        "updated_at": now_iso(),
        "device_id": device_id,
    }


# ── SHA-256 Drift Check + Neighbor-Anchor Auto-Rebind + Orphan Blocker ──────
def verify_and_rebind_profile(tpl_path: Path, profile: dict[str, Any], *, persist_if_clean: bool = True) -> dict[str, Any]:
    """
    Verify `profile["template_sha256"]` against the current `tpl_path`.
    If the template file was modified (e.g. rows/columns inserted or a cell deleted):
      - Automatically rebinds fields & repeating tables via their stored label anchors!
      - Flags any field whose anchor vanished or now collides with a formula cell in `orphan_fields`!
    """
    tpl_path = Path(tpl_path)
    prof = copy.deepcopy(profile)
    cur_hash = sha256_of_file(tpl_path)
    if prof.get("template_sha256") == cur_hash and not prof.get("orphan_fields"):
        prof["drift_detected"] = False
        prof.setdefault("rebound_fields", [])
        return prof

    prof["drift_detected"] = True
    orphans: list[dict[str, Any]] = []
    rebound: list[dict[str, Any]] = []
    ext = tpl_path.suffix.lower()

    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(tpl_path, data_only=False, keep_vba=(ext == ".xlsm"))
        try:
            # Refresh formula protection index across all sheets
            fresh_protected: list[dict[str, Any]] = []
            formula_set: set[tuple[str, str]] = set()
            label_index: dict[str, dict[str, list[tuple[int, int, str]]]] = {}
            for sname in wb.sheetnames:
                ws = wb[sname]
                s_idx: dict[str, list[tuple[int, int, str]]] = {}
                for r in range(1, min((ws.max_row or 1) + 1, 200)):
                    for c in range(1, min((ws.max_column or 1) + 1, 50)):
                        v = ws.cell(r, c).value
                        if isinstance(v, str):
                            if v.startswith("="):
                                cd = f"{get_column_letter(c)}{r}"
                                fresh_protected.append({"sheet": sname, "coord": cd, "formula": v[:160]})
                                formula_set.add((sname, cd))
                            elif v.strip():
                                ln = normalize_ar(v.strip())
                                if ln:
                                    s_idx.setdefault(ln, []).append((r, c, v.strip()))
                label_index[sname] = s_idx
            # Preserve any user-explicitly-protected cells too
            seen_prot = {(p["sheet"], p["coord"]) for p in fresh_protected}
            for up in (prof.get("protected_cells") or []):
                k = (up.get("sheet"), up.get("coord"))
                if k not in seen_prot and k[0] in wb.sheetnames:
                    fresh_protected.append(up)
            prof["protected_cells"] = fresh_protected

            # Rebind fields
            for fld in (prof.get("fields") or []):
                if fld.get("role") in ("fixed", "formula_protected"):
                    continue
                tgt = fld.get("target") or {}
                if tgt.get("kind") != "xlsx_cell":
                    continue
                sname = str(tgt.get("sheet") or "")
                old_coord = str(tgt.get("coord") or "").upper()
                if sname not in wb.sheetnames:
                    orphans.append({
                        "field_id": fld.get("id"),
                        "key": fld.get("key"),
                        "label": fld.get("label"),
                        "old_target": f"{sname}!{old_coord}",
                        "reason": tr("validation.sheet_missing", sheet=sname),
                    })
                    continue

                ws = wb[sname]
                anchor = fld.get("anchor") or {}
                lbl_norm = str(anchor.get("label_norm") or normalize_ar(str(anchor.get("label_text") or "")))
                if lbl_norm:
                    old_lbl_coord = str(anchor.get("label_coord") or "")
                    off_r = int(anchor.get("offset_row") or 0)
                    off_c = int(anchor.get("offset_col") or 0)
                    hits = label_index.get(sname, {}).get(lbl_norm, [])
                    if not hits:
                        orphans.append({
                            "field_id": fld.get("id"),
                            "key": fld.get("key"),
                            "label": fld.get("label"),
                            "old_target": f"{sname}!{old_coord}",
                            "reason": tr("server.field_moved", field=fld.get("label")),
                        })
                        continue
                    # Choose hit closest to previous label coordinate
                    if old_lbl_coord:
                        olr, olc = coordinate_to_tuple(old_lbl_coord)
                        hits = sorted(hits, key=lambda h: abs(h[0] - olr) + abs(h[1] - olc))
                    nr_lbl, nc_lbl, raw_lbl = hits[0]
                    new_r = nr_lbl + off_r
                    new_c = nc_lbl + off_c
                    if new_r < 1 or new_c < 1:
                        orphans.append({
                            "field_id": fld.get("id"),
                            "key": fld.get("key"),
                            "label": fld.get("label"),
                            "old_target": f"{sname}!{old_coord}",
                            "reason": tr("server.field_moved", field=fld.get("label")),
                        })
                        continue
                    mr, mc, mrng = _resolve_merge_master(ws, new_r, new_c)
                    new_coord = f"{get_column_letter(mc)}{mr}"
                    if (sname, new_coord) in formula_set:
                        orphans.append({
                            "field_id": fld.get("id"),
                            "key": fld.get("key"),
                            "label": fld.get("label"),
                            "old_target": f"{sname}!{old_coord}",
                            "reason": tr("server.tpl_formula_blocked", cell=f"{sname}!{new_coord}"),
                        })
                        continue
                    if new_coord != old_coord:
                        tgt["coord"] = new_coord
                        tgt["merge_range"] = mrng
                        anchor["label_coord"] = f"{get_column_letter(nc_lbl)}{nr_lbl}"
                        anchor["label_text"] = raw_lbl
                        rebound.append({
                            "field_id": fld.get("id"),
                            "label": fld.get("label"),
                            "from": f"{sname}!{old_coord}",
                            "to": f"{sname}!{new_coord}",
                        })
                else:
                    # Unanchored cell: block if it turned into a formula or static conflicting label
                    if (sname, old_coord) in formula_set:
                        orphans.append({
                            "field_id": fld.get("id"),
                            "key": fld.get("key"),
                            "label": fld.get("label"),
                            "old_target": f"{sname}!{old_coord}",
                            "reason": tr("server.tpl_formula_blocked", cell=f"{sname}!{old_coord}"),
                        })

            # Rebind repeating tables via header row anchor
            for tbl in (prof.get("repeating_tables") or []):
                if tbl.get("target_kind") != "xlsx_rows":
                    continue
                sname = str(tbl.get("sheet") or "")
                if sname not in wb.sheetnames:
                    orphans.append({
                        "field_id": tbl.get("id"),
                        "key": tbl.get("key"),
                        "label": tbl.get("label"),
                        "old_target": sname,
                        "reason": tr("validation.sheet_missing", sheet=sname),
                    })
                    continue
                ws = wb[sname]
                anch_hdrs = (tbl.get("anchor") or {}).get("headers") or {}
                if anch_hdrs:
                    expected = {c: normalize_ar(h) for c, h in anch_hdrs.items() if h}
                    found_hr: int | None = None
                    for r in range(1, min((ws.max_row or 1) + 1, 150)):
                        matches = 0
                        for col_let, exp_n in expected.items():
                            ci = column_index_from_string(col_let)
                            v = ws.cell(r, ci).value
                            if isinstance(v, str) and normalize_ar(v) == exp_n:
                                matches += 1
                        if matches >= max(1, len(expected) - 1):
                            found_hr = r
                            break
                    if found_hr is None:
                        orphans.append({
                            "field_id": tbl.get("id"),
                            "key": tbl.get("key"),
                            "label": tbl.get("label"),
                            "old_target": f"{sname}!R{tbl.get('template_row')}",
                            "reason": tr("server.table_moved", table=tbl.get("label")),
                        })
                    else:
                        old_hr = int(tbl.get("header_row") or (int(tbl.get("template_row") or 2) - 1))
                        delta = found_hr - old_hr
                        if delta != 0:
                            old_tr = int(tbl["template_row"])
                            tbl["header_row"] = found_hr
                            tbl["template_row"] = old_tr + delta
                            if tbl.get("totals_row"):
                                tbl["totals_row"] = int(tbl["totals_row"]) + delta
                            tbl["anchor"]["header_row"] = found_hr
                            rebound.append({
                                "field_id": tbl.get("id"),
                                "label": tbl.get("label"),
                                "from": f"{sname}!R{old_tr}",
                                "to": f"{sname}!R{tbl['template_row']}",
                            })
        finally:
            wb.close()

    elif ext == ".docx":
        doc = docx.Document(str(tpl_path))
        for fld in (prof.get("fields") or []):
            if fld.get("role") == "fixed":
                continue
            tgt = fld.get("target") or {}
            kind = tgt.get("kind")
            anchor = fld.get("anchor") or {}
            lbl_norm = str(anchor.get("label_norm") or normalize_ar(str(anchor.get("label_text") or "")))

            if kind == "docx_cell":
                ti = int(tgt.get("table_idx") or 0)
                ri = int(tgt.get("row_idx") or 0)
                ci = int(tgt.get("col_idx") or 0)
                if ti >= len(doc.tables):
                    orphans.append({
                        "field_id": fld.get("id"),
                        "key": fld.get("key"),
                        "label": fld.get("label"),
                        "old_target": f"Table[{ti}] R{ri}C{ci}",
                        "reason": tr("server.field_moved", field=fld.get("label")),
                    })
                    continue
                tbl = doc.tables[ti]
                if lbl_norm:
                    off_r = int(anchor.get("offset_row") or 0)
                    off_c = int(anchor.get("offset_col") or 0)
                    found_pos: tuple[int, int] | None = None
                    for rr, row in enumerate(tbl.rows):
                        for cc, cell in enumerate(row.cells):
                            if normalize_ar(cell.text.strip()) == lbl_norm or (
                                tgt.get("match_text") and str(tgt["match_text"]) in cell.text
                            ):
                                found_pos = (rr, cc)
                                break
                        if found_pos:
                            break
                    if not found_pos:
                        orphans.append({
                            "field_id": fld.get("id"),
                            "key": fld.get("key"),
                            "label": fld.get("label"),
                            "old_target": f"Table[{ti}] R{ri}C{ci}",
                            "reason": tr("server.field_moved", field=fld.get("label")),
                        })
                        continue
                    nr = found_pos[0] + off_r
                    nc = found_pos[1] + off_c
                    if nr < 0 or nr >= len(tbl.rows) or nc < 0 or nc >= len(tbl.rows[nr].cells):
                        orphans.append({
                            "field_id": fld.get("id"),
                            "key": fld.get("key"),
                            "label": fld.get("label"),
                            "old_target": f"Table[{ti}] R{ri}C{ci}",
                            "reason": tr("server.field_moved", field=fld.get("label")),
                        })
                        continue
                    if (nr, nc) != (ri, ci):
                        tgt["row_idx"] = nr
                        tgt["col_idx"] = nc
                        anchor["label_row"], anchor["label_col"] = found_pos
                        rebound.append({
                            "field_id": fld.get("id"),
                            "label": fld.get("label"),
                            "from": f"T{ti}:R{ri}C{ci}",
                            "to": f"T{ti}:R{nr}C{nc}",
                        })
                elif ri >= len(tbl.rows) or ci >= len(tbl.rows[ri].cells):
                    orphans.append({
                        "field_id": fld.get("id"),
                        "key": fld.get("key"),
                        "label": fld.get("label"),
                        "old_target": f"Table[{ti}] R{ri}C{ci}",
                        "reason": tr("server.field_moved", field=fld.get("label")),
                    })

            elif kind in ("docx_paragraph", "docx_span"):
                pi = int(tgt.get("para_idx") or 0)
                match_txt = str(tgt.get("match_text") or "")
                found_pi: int | None = None
                for idx_p, para in enumerate(doc.paragraphs):
                    pt = para.text
                    if match_txt and match_txt in pt:
                        found_pi = idx_p
                        break
                    if not match_txt and lbl_norm and lbl_norm in normalize_ar(pt):
                        found_pi = idx_p
                        break
                if found_pi is None:
                    if not lbl_norm and not match_txt and 0 <= pi < len(doc.paragraphs):
                        found_pi = pi
                    else:
                        orphans.append({
                            "field_id": fld.get("id"),
                            "key": fld.get("key"),
                            "label": fld.get("label"),
                            "old_target": f"Paragraph[{pi}]",
                            "reason": tr("server.field_moved", field=fld.get("label")),
                        })
                        continue
                if found_pi != pi:
                    tgt["para_idx"] = found_pi
                    anchor["para_idx"] = found_pi
                    rebound.append({
                        "field_id": fld.get("id"),
                        "label": fld.get("label"),
                        "from": f"P{pi}",
                        "to": f"P{found_pi}",
                    })

    prof["orphan_fields"] = orphans
    prof["rebound_fields"] = rebound
    if not orphans:
        prof["template_sha256"] = cur_hash
        if persist_if_clean and tplmap_path_for(tpl_path).exists():
            prof["updated_at"] = now_iso()
            atomic_write_json(tplmap_path_for(tpl_path), prof)
    return prof


def get_template_profile(proj_root: Path | str, tpl_name: str, *, device_id: str = "") -> dict[str, Any]:
    """
    Load `<name>.tplmap.json` for an uploaded template (migrating legacy v9.0 `_template_maps.json`
    if present), verify SHA-256 + rebind anchors, or return an unsaved starter profile with suggestions.
    """
    tpl_path = resolve_project_template(proj_root, tpl_name)
    mp = tplmap_path_for(tpl_path)
    saved = read_json_safe(mp)

    # Backward compatibility: migrate legacy v9.0 `_template_maps.json` if present
    if not isinstance(saved, dict):
        legacy = read_json_safe(tpl_path.parent / "_template_maps.json", {})
        leg_key = f"custom:{tpl_path.name}"
        if isinstance(legacy, dict) and leg_key in legacy and isinstance(legacy[leg_key], dict):
            leg_entry = legacy[leg_key]
            saved = _migrate_legacy_v9_entry(tpl_path, leg_entry, device_id=device_id)
            atomic_write_json(mp, saved)

    if isinstance(saved, dict):
        saved["has_saved_profile"] = True
        verified = verify_and_rebind_profile(tpl_path, saved, persist_if_clean=True)
        # Attach fresh deterministic suggestions on demand if profile has zero fields yet
        if not verified.get("fields") and not verified.get("repeating_tables"):
            det = suggest_fields_deterministic(tpl_path)
            verified["suggestions"] = {
                "fields": det["suggested_fields"],
                "repeating_tables": det["suggested_tables"],
            }
        return verified

    return build_initial_profile(tpl_path, device_id=device_id)


def _migrate_legacy_v9_entry(tpl_path: Path, leg: dict[str, Any], *, device_id: str = "") -> dict[str, Any]:
    ext = tpl_path.suffix.lower()
    fields: list[dict[str, Any]] = []
    for k, target_str in (leg.get("field_map") or {}).items():
        if not target_str:
            continue
        ts = str(target_str).strip()
        if ext in (".xlsx", ".xlsm") and "!" in ts:
            sh, cd = ts.split("!", 1)
            tgt = {"kind": "xlsx_cell", "sheet": sh, "coord": cd.upper()}
        elif ext == ".docx" and ts.startswith("T"):
            m = re.match(r"T(\d+):R(\d+)C(\d+)", ts)
            if not m:
                continue
            tgt = {"kind": "docx_cell", "table_idx": int(m.group(1)), "row_idx": int(m.group(2)), "col_idx": int(m.group(3)), "replace_mode": "cell"}
        else:
            continue
        fields.append({
            "id": f"fld_mig_{k}",
            "key": k,
            "label": k,
            "role": "auto_project" if k in ("project_name", "contract_no", "owner", "consultant", "contractor") else ("auto_system" if k in ("ref_no", "doc_date", "revision") else "manual"),
            "source": f"project.{k}" if k in ("contract_no", "owner", "consultant", "contractor") else ("project.name" if k == "project_name" else ("system.next_ref" if k == "ref_no" else ("system.today_iso" if k == "doc_date" else ("system.revision" if k == "revision" else "")))),
            "input_type": "date" if "date" in k else ("kwd" if "amount" in k else "text"),
            "required": k == "subject",
            "default_value": "",
            "options": [],
            "target": tgt,
            "anchor": {},
        })
    sugg = suggest_template_section(tpl_path.name)
    return {
        "schema_version": "2.0",
        "template_name": tpl_path.name,
        "template_ext": ext,
        "template_sha256": sha256_of_file(tpl_path),
        "display_name": str(leg.get("display_name") or tpl_path.stem),
        "section": str(leg.get("section") or sugg["suggested_section"] or ""),
        "section_confirmed": bool(leg.get("section") in SECTIONS_BY_ID),
        "ref_prefix": str(leg.get("ref_prefix") or sugg["ref_prefix"] or ""),
        "use_discipline_series": False,
        "default_discipline": "",
        "log_on_issue": True,
        "default_status": "Under Review",
        "has_saved_profile": True,
        "fields": fields,
        "protected_cells": [],
        "repeating_tables": [],
        "orphan_fields": [],
        "rebound_fields": [],
        "updated_at": now_iso(),
        "device_id": device_id,
    }


def save_template_profile(
    proj_root: Path | str,
    tpl_name: str,
    payload: dict[str, Any],
    *,
    device_id: str = "",
    project_sections: list[dict[str, Any]] | dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """
    Validate and persist `<template_name>.tplmap.json` beside the uploaded template.
    Automatically computes neighbor anchors for every mapped field & blocks mapping
    any writable field onto a live Excel formula cell!
    """
    tpl_path = resolve_project_template(proj_root, tpl_name)
    ext = tpl_path.suffix.lower()

    sec_map = sections_map_for_project(project_sections=project_sections)
    section = str(payload.get("section") or "").strip()
    if not section or section not in sec_map:
        raise ValueError(tr("studio.section_required"))
    sec_info = sec_map[section]

    raw_fields = payload.get("fields") or []
    raw_tables = payload.get("repeating_tables") or []
    raw_protected = payload.get("protected_cells") or []

    clean_fields: list[dict[str, Any]] = []
    clean_protected: list[dict[str, Any]] = []
    clean_tables: list[dict[str, Any]] = []
    used_keys: set[str] = set()

    wb = openpyxl.load_workbook(tpl_path, data_only=False, keep_vba=(ext == ".xlsm")) if ext in (".xlsx", ".xlsm") else None
    doc = docx.Document(str(tpl_path)) if ext == ".docx" else None
    try:
        formula_coords: set[tuple[str, str]] = set()
        if wb is not None:
            for sn in wb.sheetnames:
                ws = wb[sn]
                for r in range(1, min((ws.max_row or 1) + 1, 200)):
                    for c in range(1, min((ws.max_column or 1) + 1, 50)):
                        v = ws.cell(r, c).value
                        if isinstance(v, str) and v.startswith("="):
                            cd = f"{get_column_letter(c)}{r}"
                            formula_coords.add((sn, cd))
                            clean_protected.append({"sheet": sn, "coord": cd, "formula": v[:160]})

        seen_prot = {(p["sheet"], p["coord"]) for p in clean_protected}
        for p in raw_protected:
            if isinstance(p, dict) and p.get("sheet") and p.get("coord"):
                pk = (str(p["sheet"]), str(p["coord"]).upper())
                if pk not in seen_prot:
                    clean_protected.append({"sheet": pk[0], "coord": pk[1], "formula": str(p.get("formula") or "user_locked")})
                    seen_prot.add(pk)

        for idx, rf in enumerate(raw_fields):
            if not isinstance(rf, dict):
                continue
            role = str(rf.get("role") or "manual").strip()
            if role not in VALID_ROLES or role == "repeating_table":
                raise ValueError(tr("server.invalid_setting"))
            tgt = dict(rf.get("target") or {})
            label = str(rf.get("label") or rf.get("key") or tr("validation.field_default", index=idx + 1)).strip()
            raw_key = str(rf.get("key") or "").strip()
            if not raw_key:
                raw_key = re.sub(r"\W+", "_", normalize_ar(label)).strip("_")[:28] or f"field_{idx + 1}"
            key = raw_key
            k_i = 2
            while key in used_keys:
                key = f"{raw_key}_{k_i}"
                k_i += 1
            used_keys.add(key)

            itype = str(rf.get("input_type") or "text").strip()
            if itype not in VALID_INPUT_TYPES:
                itype = "text"
            options = [str(o).strip() for o in (rf.get("options") or []) if str(o).strip()]
            if itype == "approval_code" and not options:
                options = list(DEFAULT_APPROVAL_OPTIONS)
            if itype == "dropdown" and not options and role == "manual":
                raise ValueError(tr("studio.need_opts_toast"))

            source = str(rf.get("source") or "").strip()
            if role == "auto_project" and source not in AUTO_PROJECT_SOURCES:
                source = "project.name"
            elif role == "auto_system" and source not in AUTO_SYSTEM_SOURCES:
                source = "system.today_iso"
            elif role == "auto_date" and source not in AUTO_DATE_SOURCES:
                source = "date.today_dmy"

            # Optional custom auto_counter settings
            c_prefix = str(rf.get("counter_prefix") or "")
            c_suffix = str(rf.get("counter_suffix") or "")
            c_sync_log = bool(rf.get("counter_sync_log", False))
            try:
                c_next = max(0, int(rf.get("counter_next") if rf.get("counter_next") not in (None, "") else 1))
            except (TypeError, ValueError):
                c_next = 1
            try:
                c_digits = max(1, min(8, int(rf.get("counter_digits") or 3)))
            except (TypeError, ValueError):
                c_digits = 3

            anchor = dict(rf.get("anchor") or {})
            if ext in (".xlsx", ".xlsm") and wb is not None:
                tgt["kind"] = "xlsx_cell"
                sh = str(tgt.get("sheet") or wb.sheetnames[0])
                if sh not in wb.sheetnames:
                    raise ValueError(tr("validation.sheet_missing", sheet=sh))
                ws = wb[sh]
                cd_raw = str(tgt.get("coord") or "").strip().upper()
                if not re.fullmatch(r"[A-Z]{1,3}\d{1,5}", cd_raw):
                    raise ValueError(tr("validation.bad_cell", field=label))
                rr, cc = coordinate_to_tuple(cd_raw)
                mr, mc, mrng = _resolve_merge_master(ws, rr, cc)
                cd = f"{get_column_letter(mc)}{mr}"
                tgt["sheet"] = sh
                tgt["coord"] = cd
                tgt["merge_range"] = mrng

                if role == "formula_protected":
                    if (sh, cd) not in seen_prot:
                        clean_protected.append({"sheet": sh, "coord": cd, "formula": str(ws.cell(mr, mc).value or "locked")})
                        seen_prot.add((sh, cd))
                    continue
                if role != "fixed" and ((sh, cd) in formula_coords or (sh, cd) in seen_prot):
                    raise FormulaOverwriteError(
                        tr("server.tpl_formula_blocked", cell=f"{sh}!{cd}")
                    )
                if not anchor.get("label_norm"):
                    anchor = _find_xlsx_anchor(ws, mr, mc)

            elif ext == ".docx" and doc is not None:
                kind = str(tgt.get("kind") or "docx_cell")
                if kind == "docx_cell":
                    ti = int(tgt.get("table_idx") or 0)
                    ri = int(tgt.get("row_idx") or 0)
                    ci = int(tgt.get("col_idx") or 0)
                    if ti < 0 or ti >= len(doc.tables):
                        raise ValueError(tr("validation.bad_cell", field=label))
                    tbl = doc.tables[ti]
                    if ri < 0 or ri >= len(tbl.rows) or ci < 0 or ci >= len(tbl.rows[ri].cells):
                        raise ValueError(tr("validation.bad_cell", field=label))
                    tgt.setdefault("replace_mode", "cell")
                    if not anchor.get("label_norm"):
                        # Auto-capture adjacent label cell in same row
                        for dc in (-1, 1):
                            nc = ci + dc
                            if 0 <= nc < len(tbl.rows[ri].cells):
                                ltxt = tbl.rows[ri].cells[nc].text.strip()
                                if ltxt and tbl.rows[ri].cells[nc]._tc is not tbl.rows[ri].cells[ci]._tc:
                                    anchor = {
                                        "label_text": ltxt[:80],
                                        "label_norm": normalize_ar(ltxt[:80]),
                                        "table_idx": ti,
                                        "label_row": ri,
                                        "label_col": nc,
                                        "offset_row": 0,
                                        "offset_col": ci - nc,
                                    }
                                    break
                elif kind in ("docx_paragraph", "docx_span"):
                    pi = int(tgt.get("para_idx") or 0)
                    if pi < 0 or pi >= len(doc.paragraphs):
                        raise ValueError(tr("validation.bad_cell", field=label))
                    if not anchor.get("label_norm"):
                        ptxt = str(tgt.get("match_text") or tgt.get("prefix_text") or doc.paragraphs[pi].text[:50]).strip()
                        if ptxt:
                            anchor = {"label_text": ptxt, "label_norm": normalize_ar(ptxt), "para_idx": pi}
                else:
                    raise ValueError(tr("validation.bad_cell", field=label))

            fld_entry: dict[str, Any] = {
                "id": str(rf.get("id") or f"fld_{secrets.token_hex(3)}"),
                "key": key,
                "label": label,
                "role": role,
                "source": source,
                "input_type": itype,
                "required": bool(rf.get("required", False)) if role == "manual" else False,
                "default_value": rf.get("default_value", ""),
                "options": options,
                "target": tgt,
                "anchor": anchor,
            }
            if role == "auto_counter":
                fld_entry.update({
                    "counter_prefix": c_prefix,
                    "counter_next": c_next,
                    "counter_digits": c_digits,
                    "counter_suffix": c_suffix,
                    "counter_sync_log": c_sync_log,
                })
            clean_fields.append(fld_entry)

        # Validate & normalize repeating tables
        for tidx, rt in enumerate(raw_tables):
            if not isinstance(rt, dict):
                continue
            tkey = str(rt.get("key") or f"items_{tidx + 1}").strip()
            tlabel = str(rt.get("label") or tr("validation.table_default", index=tidx + 1)).strip()
            tcols_in = rt.get("columns") or []
            if not tcols_in:
                raise ValueError(tr("validation.table_needs_col"))

            if ext in (".xlsx", ".xlsm") and wb is not None:
                sh = str(rt.get("sheet") or wb.sheetnames[0])
                if sh not in wb.sheetnames:
                    raise ValueError(tr("validation.sheet_missing", sheet=sh))
                ws = wb[sh]
                tmpl_row = max(1, int(rt.get("template_row") or 2))
                hdr_row = int(rt.get("header_row") or max(1, tmpl_row - 1))
                tot_row = int(rt["totals_row"]) if rt.get("totals_row") not in (None, "", 0, "0") else None
                pref_rows = max(1, int(rt.get("preformatted_rows") or ((tot_row - tmpl_row) if tot_row and tot_row > tmpl_row else 1)))
                norm_cols = []
                anch_hdrs = {}
                for _ci_idx, col_d in enumerate(tcols_in):
                    col_let = str(col_d.get("col") or "").strip().upper()
                    if col_let.isdigit():
                        col_let = get_column_letter(int(col_let))
                    if not re.fullmatch(r"[A-Z]{1,3}", col_let):
                        raise ValueError(tr("validation.bad_cell", field=tlabel))
                    c_num = column_index_from_string(col_let)
                    c_it = str(col_d.get("input_type") or "text")
                    tv = ws.cell(tmpl_row, c_num).value
                    ftmpl = str(col_d.get("formula_template") or "")
                    if isinstance(tv, str) and tv.startswith("="):
                        c_it = "formula_protected"
                        if not ftmpl:
                            ftmpl = re.sub(rf"(\$?[A-Z]{{1,3}}){tmpl_row}(?!\d)", r"\1{row}", tv)
                    clbl = str(col_d.get("label") or ws.cell(hdr_row, c_num).value or col_let).strip()
                    ckey = str(col_d.get("key") or f"col_{col_let.lower()}").strip()
                    norm_cols.append({
                        "key": ckey,
                        "label": clbl,
                        "col": col_let,
                        "input_type": c_it if c_it in VALID_INPUT_TYPES else "text",
                        "required": bool(col_d.get("required", False)) and c_it not in ("auto_index", "formula_protected"),
                        "options": [str(o).strip() for o in (col_d.get("options") or []) if str(o).strip()],
                        "formula_template": ftmpl,
                    })
                    hv = ws.cell(hdr_row, c_num).value
                    if isinstance(hv, str) and hv.strip():
                        anch_hdrs[col_let] = hv.strip()
                clean_tables.append({
                    "id": str(rt.get("id") or f"tbl_{secrets.token_hex(3)}"),
                    "key": tkey,
                    "label": tlabel,
                    "target_kind": "xlsx_rows",
                    "sheet": sh,
                    "header_row": hdr_row,
                    "template_row": tmpl_row,
                    "preformatted_rows": pref_rows,
                    "totals_row": tot_row,
                    "columns": norm_cols,
                    "anchor": {"header_row": hdr_row, "headers": anch_hdrs},
                })
            elif ext == ".docx" and doc is not None:
                ti = int(rt.get("table_idx") or 0)
                if ti < 0 or ti >= len(doc.tables):
                    raise ValueError(tr("validation.bad_cell", field=tlabel))
                dtbl = doc.tables[ti]
                tmpl_row = max(0, int(rt.get("template_row") or 1))
                hdr_row = max(0, int(rt.get("header_row") or 0))
                tot_row = int(rt["totals_row"]) if rt.get("totals_row") not in (None, "") else None
                pref_rows = max(1, int(rt.get("preformatted_rows") or ((tot_row - tmpl_row) if tot_row and tot_row > tmpl_row else max(1, len(dtbl.rows) - tmpl_row))))
                norm_cols = []
                anch_hdrs = {}
                for col_d in tcols_in:
                    c_i = int(col_d.get("col_idx") or 0)
                    c_it = str(col_d.get("input_type") or "text")
                    clbl = str(col_d.get("label") or tr("validation.col_default", index=c_i + 1)).strip()
                    ckey = str(col_d.get("key") or f"col_{c_i}").strip()
                    norm_cols.append({
                        "key": ckey,
                        "label": clbl,
                        "col_idx": c_i,
                        "input_type": c_it if c_it in VALID_INPUT_TYPES else "text",
                        "required": bool(col_d.get("required", False)) and c_it != "auto_index",
                        "options": [str(o).strip() for o in (col_d.get("options") or []) if str(o).strip()],
                    })
                    if hdr_row < len(dtbl.rows) and c_i < len(dtbl.rows[hdr_row].cells):
                        ht = dtbl.rows[hdr_row].cells[c_i].text.strip()
                        if ht:
                            anch_hdrs[str(c_i)] = ht
                clean_tables.append({
                    "id": str(rt.get("id") or f"dtbl_{secrets.token_hex(3)}"),
                    "key": tkey,
                    "label": tlabel,
                    "target_kind": "docx_rows",
                    "table_idx": ti,
                    "header_row": hdr_row,
                    "template_row": tmpl_row,
                    "preformatted_rows": pref_rows,
                    "totals_row": tot_row,
                    "columns": norm_cols,
                    "anchor": {"table_idx": ti, "header_row": hdr_row, "headers": anch_hdrs},
                })
    finally:
        if wb is not None:
            wb.close()

    saved_profile = {
        "schema_version": "2.0",
        "template_name": tpl_path.name,
        "template_ext": ext,
        "template_sha256": sha256_of_file(tpl_path),
        "display_name": str(payload.get("display_name") or tpl_path.stem).strip(),
        "section": section,
        "section_confirmed": True,
        "ref_prefix": str(payload.get("ref_prefix") or sec_info["ref_prefix"]).strip().upper(),
        "use_discipline_series": bool(payload.get("use_discipline_series", False)),
        "default_discipline": str(payload.get("default_discipline") or "").strip().upper(),
        "log_on_issue": bool(payload.get("log_on_issue", True)),
        "default_status": str(payload.get("default_status") or "Under Review").strip(),
        "target_log": str(payload.get("target_log") or "").strip(),
        "target_sheet": str(payload.get("target_sheet") or "").strip(),
        "has_saved_profile": True,
        "fields": clean_fields,
        "protected_cells": clean_protected,
        "repeating_tables": clean_tables,
        "orphan_fields": [],
        "rebound_fields": [],
        "updated_at": now_iso(),
        "device_id": device_id,
    }
    atomic_write_json(tplmap_path_for(tpl_path), saved_profile)
    return saved_profile


def list_project_templates(
    proj_root: Path | str,
    *,
    section: str | None = None,
    device_id: str = "",
    project_sections: list[dict[str, Any]] | dict[str, dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """
    List ONLY user-uploaded templates in `<proj_root>/4_Project_Templates/`.
    ZERO system/built-in templates are ever returned or suggested!
    """
    sec_map = sections_map_for_project(project_sections=project_sections)
    tdir = templates_dir_for_project(proj_root)
    items: list[dict[str, Any]] = []
    for f in sorted(tdir.iterdir()):
        if not f.is_file():
            continue
        if f.name.startswith(("~$", ".", "_")):
            continue
        if f.suffix.lower() not in ALLOWED_TEMPLATE_EXTS:
            continue
        mp = tplmap_path_for(f)
        has_prof = mp.exists()
        prof_data: dict[str, Any] = {}
        if has_prof:
            try:
                prof_data = get_template_profile(proj_root, f.name, device_id=device_id)
            except Exception:
                prof_data = read_json_safe(mp, {}) or {}
        sugg = suggest_template_section(f.name)
        sec_id = str(prof_data.get("section") or "") if prof_data.get("section_confirmed") else ""
        if section and sec_id != section:
            continue
        sec_meta = sec_map.get(sec_id, {})
        active_fields = [x for x in (prof_data.get("fields") or []) if x.get("role") not in ("fixed", "formula_protected")]
        items.append({
            "name": f.name,
            "display_name": str(prof_data.get("display_name") or f.stem),
            "ext": f.suffix.lower(),
            "size_kb": round(f.stat().st_size / 1024, 1),
            "modified_at": dt.datetime.fromtimestamp(f.stat().st_mtime).isoformat(timespec="seconds"),
            "has_profile": bool(has_prof and prof_data.get("section_confirmed")),
            "section": sec_id,
            "section_label": str(sec_meta.get("label") or (section_label(sec_id) if sec_id in SECTIONS_BY_ID else tr("templates.no_section"))),
            "section_confirmed": bool(prof_data.get("section_confirmed")),
            "suggested_section": sugg["suggested_section"],
            "suggested_label": sugg["suggested_label"],
            "ref_prefix": str(prof_data.get("ref_prefix") or sec_meta.get("ref_prefix") or ""),
            "fields_count": len(active_fields),
            "repeating_tables_count": len(prof_data.get("repeating_tables") or []),
            "protected_cells_count": len(prof_data.get("protected_cells") or []),
            "orphan_fields": prof_data.get("orphan_fields") or [],
            "rebound_fields": prof_data.get("rebound_fields") or [],
            "ready_to_issue": bool(
                has_prof
                and prof_data.get("section_confirmed")
                and not (prof_data.get("orphan_fields") or [])
                and (len(active_fields) > 0 or len(prof_data.get("repeating_tables") or []) > 0)
            ),
        })
    return items


# ── Strict Input Validation + Deterministic `fill_by_profile` ───────────────
def _coerce_and_validate_value(
    raw_val: Any,
    *,
    label: str,
    key: str,
    input_type: str,
    required: bool,
    default_value: Any = "",
    options: list[str] | None = None,
) -> tuple[Any, str | None]:
    """Validate and normalize a single field/cell value. Returns `(typed_value, error_msg)`."""
    val = raw_val
    if val is None or (isinstance(val, str) and val.strip() == ""):
        val = default_value
    is_empty = val is None or (isinstance(val, str) and val.strip() == "")
    if is_empty:
        if required:
            return None, tr("validation.required")
        return "", None

    s = str(val).strip()
    if input_type == "date":
        if isinstance(val, (dt.date, dt.datetime)):
            return (val.date() if isinstance(val, dt.datetime) else val).isoformat(), None
        # Accept YYYY-MM-DD or DD/MM/YYYY or DD-MM-YYYY
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                parsed = dt.datetime.strptime(s, fmt).date()
                return parsed.isoformat(), None
            except ValueError:
                continue
        return None, tr("validation.bad_date")

    if input_type in ("number", "kwd"):
        cleaned = s.replace(",", "").replace("د.ك", "").replace("KWD", "").strip()
        try:
            num = float(cleaned)
        except ValueError:
            return None, tr("validation.bad_kwd" if input_type == "kwd" else "validation.bad_number")
        if input_type == "kwd":
            return round(num, 3), None
        return int(num) if num.is_integer() and "." not in cleaned else num, None

    if input_type in ("dropdown", "approval_code"):
        allowed = list(options or (DEFAULT_APPROVAL_OPTIONS if input_type == "approval_code" else []))
        if allowed:
            # Case-insensitive exact match, returning canonical option spelling
            match = next((opt for opt in allowed if opt.strip().lower() == s.lower()), None)
            if match is None:
                return None, tr("validation.bad_option")
            return match, None
        return s, None

    return s, None


def resolve_auto_value(
    role: str,
    source: str,
    project: dict[str, Any],
    system_ctx: dict[str, Any],
    override_val: Any = None,
    field_def: dict[str, Any] | None = None,
) -> Any:
    """Resolve `auto_date`, `auto_counter`, `auto_project`, or `auto_system` values cleanly."""
    if override_val is not None and str(override_val).strip() != "":
        return str(override_val).strip()
    today = dt.date.today()
    if system_ctx.get("today_iso"):
        try:
            today = dt.date.fromisoformat(str(system_ctx["today_iso"]))
        except ValueError:
            pass

    if role == "auto_date":
        fmt_key = source or "date.today_dmy"
        if fmt_key in ("date.today_iso", "system.today_iso"):
            return today.isoformat()
        if fmt_key in ("date.today_dmy", "system.today_dmy"):
            return today.strftime("%d/%m/%Y")
        if fmt_key in ("date.today_dmy_dash", "system.today_dmy_dash"):
            return today.strftime("%d-%m-%Y")
        if fmt_key in ("date.month_year", "system.month_year"):
            return today.strftime("%m/%Y")
        if fmt_key in ("date.day_name_ar", "system.day_name_ar"):
            return _AR_DAY_NAMES[today.weekday()]
        if fmt_key in ("date.year", "system.year"):
            return str(today.year)
        return today.strftime("%d/%m/%Y")

    if role == "auto_counter":
        fdef = field_def or {}
        if fdef.get("counter_sync_log") and str(system_ctx.get("next_ref") or "").strip():
            return str(system_ctx["next_ref"]).strip()
        return format_counter_value(fdef)

    if role == "auto_project":
        field = source.split(".", 1)[1] if "." in source else source
        return str(project.get(field) or "").strip()
    if role == "auto_system":
        if source == "system.next_ref":
            return str(system_ctx.get("next_ref") or "").strip()
        if source == "system.today_iso":
            return today.isoformat()
        if source == "system.today_dmy":
            return today.strftime("%d/%m/%Y")
        if source == "system.today_dmy_dash":
            return today.strftime("%d-%m-%Y")
        if source == "system.month_year":
            return today.strftime("%m/%Y")
        if source == "system.day_name_ar":
            return _AR_DAY_NAMES[today.weekday()]
        if source == "system.revision":
            return str(system_ctx.get("revision") or "R00").strip()
        if source == "system.year":
            return str(today.year)
    return ""


def validate_and_resolve_payload(
    profile: dict[str, Any],
    payload: dict[str, Any],
    project: dict[str, Any],
    system_ctx: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Validate all manual fields and repeating table rows against `profile`, and resolve
    all `auto_date`, `auto_counter`, `auto_project`, and `auto_system` values.
    """
    sys_ctx = dict(system_ctx or {})
    values_in = payload.get("values") if isinstance(payload.get("values"), dict) else payload
    tables_in = payload.get("tables") if isinstance(payload.get("tables"), dict) else {}

    errors: list[dict[str, str]] = []
    resolved_fields: dict[str, Any] = {}

    for fld in (profile.get("fields") or []):
        role = fld.get("role", "manual")
        key = str(fld.get("key") or "")
        label = str(fld.get("label") or key)
        if role in ("fixed", "formula_protected"):
            continue
        if role in ("auto_date", "auto_counter", "auto_project", "auto_system"):
            resolved_fields[key] = resolve_auto_value(
                role, str(fld.get("source") or ""), project, sys_ctx, values_in.get(key), field_def=fld
            )
            continue
        val, err = _coerce_and_validate_value(
            values_in.get(key),
            label=label,
            key=key,
            input_type=str(fld.get("input_type") or "text"),
            required=bool(fld.get("required", False)),
            default_value=fld.get("default_value", ""),
            options=fld.get("options"),
        )
        if err:
            errors.append({"key": key, "label": label, "error": err})
        else:
            resolved_fields[key] = val

    resolved_tables: dict[str, list[dict[str, Any]]] = {}
    for tbl in (profile.get("repeating_tables") or []):
        tkey = str(tbl.get("key") or "items")
        raw_rows = tables_in.get(tkey)
        if raw_rows is None and isinstance(payload.get(tkey), list):
            raw_rows = payload.get(tkey)
        if not isinstance(raw_rows, list):
            resolved_tables[tkey] = []
            continue
        clean_rows: list[dict[str, Any]] = []
        cols = tbl.get("columns") or []
        for r_i, row_obj in enumerate(raw_rows):
            if not isinstance(row_obj, dict):
                continue
            # Skip completely blank UI rows
            has_any_user_val = any(
                col.get("input_type") not in ("auto_index", "formula_protected")
                and row_obj.get(col["key"]) not in (None, "")
                and str(row_obj.get(col["key"])).strip() != ""
                for col in cols
            )
            if not has_any_user_val:
                continue
            out_row: dict[str, Any] = {}
            for col in cols:
                ck = str(col["key"])
                cit = str(col.get("input_type") or "text")
                clbl = tr("validation.row_prefix", field=col.get("label", ck), row=len(clean_rows) + 1)
                if cit == "auto_index":
                    out_row[ck] = len(clean_rows) + 1
                    continue
                if cit == "formula_protected":
                    continue
                v, err = _coerce_and_validate_value(
                    row_obj.get(ck),
                    label=clbl,
                    key=f"{tkey}[{r_i}].{ck}",
                    input_type=cit,
                    required=bool(col.get("required", False)),
                    options=col.get("options"),
                )
                if err:
                    errors.append({"key": f"{tkey}[{r_i}].{ck}", "label": clbl, "error": err})
                else:
                    out_row[ck] = v
            clean_rows.append(out_row)
        resolved_tables[tkey] = clean_rows

    if errors:
        summary = " | ".join(e["error"] for e in errors[:4])
        raise ProfileValidationError(summary, errors=errors)

    return {"fields": resolved_fields, "tables": resolved_tables}


def _copy_xlsx_template_row(ws, src_row: int, dst_row: int, max_col: int, row_merges: list[tuple[int, int]]) -> None:
    """Clone full Excel row appearance (height, fonts, borders, fills, number formats, row merges & formulas)."""
    if src_row == dst_row:
        return
    if ws.row_dimensions[src_row].height is not None:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    row_ref_re = re.compile(rf"(\$?[A-Z]{{1,3}}){src_row}(?!\d)")
    for c in range(1, max_col + 1):
        sc = ws.cell(src_row, c)
        dc = ws.cell(dst_row, c)
        if isinstance(dc, MergedCell):
            continue
        if isinstance(sc, MergedCell):
            # Copy border/fill/font from the underlying cell style if accessible
            mr, mc, _ = _resolve_merge_master(ws, src_row, c)
            sc_style = ws.cell(mr, mc)
            if sc_style.has_style:
                dc.font = copy.copy(sc_style.font)
                dc.fill = copy.copy(sc_style.fill)
                dc.border = copy.copy(sc.border) if sc.has_style else copy.copy(sc_style.border)
                dc.alignment = copy.copy(sc_style.alignment)
                dc.number_format = sc_style.number_format
            continue
        if sc.has_style:
            dc.font = copy.copy(sc.font)
            dc.fill = copy.copy(sc.fill)
            dc.border = copy.copy(sc.border)
            dc.alignment = copy.copy(sc.alignment)
            dc.number_format = sc.number_format
        if isinstance(sc.value, str) and sc.value.startswith("="):
            dc.value = row_ref_re.sub(rf"\g<1>{dst_row}", sc.value)
    # Re-apply horizontal merges inside this row
    for min_c, max_c in row_merges:
        rng_str = f"{get_column_letter(min_c)}{dst_row}:{get_column_letter(max_c)}{dst_row}"
        already = any(
            rg.min_row == dst_row and rg.max_row == dst_row and rg.min_col == min_c and rg.max_col == max_c
            for rg in ws.merged_cells.ranges
        )
        if not already:
            ws.merge_cells(rng_str)


def _shift_merges_and_formulas_on_insert(ws, insert_at: int, count: int, tmpl_row: int, old_last_item_row: int) -> None:
    """
    Safely insert `count` rows at `insert_at`, shifting merges at/below `insert_at`
    and expanding `=SUM(X{tmpl_row}:X{old_last_item_row})` -> `=SUM(X{tmpl_row}:X{old_last_item_row + count})`,
    plus shifting single-cell references at/below `insert_at` in footer formulas!
    """
    if count <= 0:
        return
    # Snapshot merges that sit at or below insert_at before insert_rows mutates sheet
    below_merges: list[tuple[int, int, int, int]] = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= insert_at:
            below_merges.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
            ws.unmerge_cells(str(rng))

    ws.insert_rows(insert_at, amount=count)

    # Restore shifted merges below insert_at
    for r1, c1, r2, c2 in below_merges:
        ws.merge_cells(start_row=r1 + count, start_column=c1, end_row=r2 + count, end_column=c2)

    # Update formulas in rows at or below `insert_at + count` (totals, VAT, grand total, etc.)
    new_last_item = old_last_item_row + count
    cell_or_range_re = re.compile(
        r"(\$?[A-Z]{1,3})(\$?)(\d+)(?::(\$?[A-Z]{1,3})(\$?)(\d+))?"
    )

    def _rewrite_formula(f_str: str) -> str:
        def _repl(m: re.Match) -> str:
            c1, d1, r1_s, c2, d2, r2_s = m.groups()
            r1 = int(r1_s)
            if r2_s is not None:
                r2 = int(r2_s)
                nr1 = r1 + count if r1 >= insert_at else r1
                if r1 <= tmpl_row and old_last_item_row <= r2 < insert_at:
                    nr2 = new_last_item
                elif r2 >= insert_at:
                    nr2 = r2 + count
                else:
                    nr2 = r2
                return f"{c1}{d1}{nr1}:{c2}{d2}{nr2}"
            nr1 = r1 + count if r1 >= insert_at else r1
            return f"{c1}{d1}{nr1}"
        return cell_or_range_re.sub(_repl, f_str)

    max_c = ws.max_column or 1
    for r in range(insert_at + count, (ws.max_row or insert_at + count) + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("="):
                ws.cell(r, c).value = _rewrite_formula(v)


def _set_docx_paragraph_text_preserving_style(para, new_text: str) -> None:
    """Replace a paragraph's text while keeping the first run's exact font, size, bold, color & RTL."""
    if para.runs:
        para.runs[0].text = new_text
        for extra in para.runs[1:]:
            extra.text = ""
    else:
        para.add_run(new_text)


def _replace_docx_span_preserving_style(para, match_text: str, replacement: str) -> None:
    """Replace a specific text span (even when split across multiple Word XML runs!) keeping run style."""
    for run in para.runs:
        if match_text in run.text:
            run.text = run.text.replace(match_text, replacement)
            return
    full = para.text
    if match_text in full:
        _set_docx_paragraph_text_preserving_style(para, full.replace(match_text, replacement))


def fill_by_profile(
    tpl_path: Path | str,
    profile: dict[str, Any],
    payload: dict[str, Any],
    project: dict[str, Any],
    output_path: Path | str,
    *,
    system_ctx: dict[str, Any] | None = None,
    verify_drift: bool = True,
    project_sections: list[dict[str, Any]] | dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """
    100% deterministic template filling engine — ZERO guessing!
    Touches STRICTLY AND ONLY the cells/paragraphs/spans/repeating-rows defined in `profile`.
    Preserves every other cell, formula, border, font, fill, merge, embedded logo, and VBA macro!
    """
    src_path = Path(tpl_path).resolve()
    out_path = Path(output_path).resolve()
    if not src_path.exists():
        raise FileNotFoundError(tr("server.file_not_found"))

    active_prof = verify_and_rebind_profile(src_path, profile, persist_if_clean=False) if verify_drift else copy.deepcopy(profile)
    if active_prof.get("orphan_fields"):
        orph = active_prof["orphan_fields"]
        raise ProfileOrphanError(
            tr("server.tpl_moved_fields", count=len(orph), name=src_path.name),
            orphan_fields=orph,
        )

    sec_map = sections_map_for_project(project, project_sections=project_sections)
    if not active_prof.get("section_confirmed") or active_prof.get("section") not in sec_map:
        raise ProfileValidationError(
            tr("studio.section_required"),
            errors=[{"key": "section", "label": tr("studio.pick_section"), "error": tr("studio.section_required")}],
        )

    resolved = validate_and_resolve_payload(active_prof, payload, project, system_ctx=system_ctx)
    fvals = resolved["fields"]
    tvals = resolved["tables"]
    ext = src_path.suffix.lower()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    touched: list[str] = []

    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(src_path, data_only=False, keep_vba=(ext == ".xlsm"))
        try:
            protected_set = {
                (str(p.get("sheet")), str(p.get("coord")).upper())
                for p in (active_prof.get("protected_cells") or [])
            }

            # 1. Write scalar fields FIRST (before repeating-row insertion shifts rows below table)
            # Track fields sitting below any repeating table so we only write each field to its true cell:
            # Actually if we write scalar fields BEFORE inserting repeating rows, `ws.insert_rows`
            # automatically shifts any scalar cell below the table along with its row!
            for fld in (active_prof.get("fields") or []):
                role = fld.get("role")
                if role in ("fixed", "formula_protected"):
                    continue
                key = str(fld.get("key") or "")
                if key not in fvals:
                    continue
                tgt = fld.get("target") or {}
                sh = str(tgt.get("sheet") or "")
                cd = str(tgt.get("coord") or "").upper()
                if sh not in wb.sheetnames or not cd:
                    continue
                ws = wb[sh]
                rr, cc = coordinate_to_tuple(cd)
                mr, mc, _ = _resolve_merge_master(ws, rr, cc)
                mcd = f"{get_column_letter(mc)}{mr}"
                cell = ws.cell(mr, mc)
                if (sh, mcd) in protected_set or (isinstance(cell.value, str) and cell.value.startswith("=")):
                    raise FormulaOverwriteError(tr("server.tpl_formula_blocked", cell=f"{sh}!{mcd}"))

                val = fvals[key]
                itype = str(fld.get("input_type") or "text")
                # Avoid overwriting empty non-provided optional fields with None; write only resolved value
                if itype == "kwd" and isinstance(val, (int, float)):
                    cell.value = round(float(val), 3)
                    if not cell.number_format or cell.number_format == "General":
                        cell.number_format = "#,##0.000"
                else:
                    cell.value = val
                touched.append(f"{sh}!{mcd}")

            # 2. Expand & populate repeating tables (process bottom-to-top per sheet so row shifts never collide!)
            sorted_tables = sorted(
                active_prof.get("repeating_tables") or [],
                key=lambda t: int(t.get("template_row") or 0),
                reverse=True,
            )
            for tbl in sorted_tables:
                if tbl.get("target_kind") != "xlsx_rows":
                    continue
                tkey = str(tbl.get("key") or "items")
                rows_data = tvals.get(tkey) or []
                if not rows_data:
                    continue
                sh = str(tbl.get("sheet") or "")
                if sh not in wb.sheetnames:
                    continue
                ws = wb[sh]
                tmpl_row = int(tbl["template_row"])
                tot_row = int(tbl["totals_row"]) if tbl.get("totals_row") else None
                pref_slots = int(tbl.get("preformatted_rows") or ((tot_row - tmpl_row) if tot_row and tot_row > tmpl_row else 1))
                pref_slots = max(1, pref_slots)
                max_c = ws.max_column or 12

                # Capture horizontal merges inside template_row
                row_merges = [
                    (rg.min_col, rg.max_col)
                    for rg in ws.merged_cells.ranges
                    if rg.min_row == tmpl_row and rg.max_row == tmpl_row
                ]

                needed = len(rows_data)
                extra = max(0, needed - pref_slots)
                insert_at = tot_row if (tot_row and tot_row > tmpl_row) else (tmpl_row + pref_slots)
                old_last_slot = insert_at - 1
                if extra > 0:
                    _shift_merges_and_formulas_on_insert(ws, insert_at, extra, tmpl_row, old_last_slot)

                for i, rdata in enumerate(rows_data):
                    cur_r = tmpl_row + i
                    if i > 0:
                        _copy_xlsx_template_row(ws, tmpl_row, cur_r, max_c, row_merges)
                    for col in (tbl.get("columns") or []):
                        col_let = str(col["col"])
                        c_num = column_index_from_string(col_let)
                        mr, mc, _ = _resolve_merge_master(ws, cur_r, c_num)
                        c_it = str(col.get("input_type") or "text")
                        target_cell = ws.cell(mr, mc)
                        if c_it == "formula_protected":
                            ftmpl = str(col.get("formula_template") or "")
                            if ftmpl and not (isinstance(target_cell.value, str) and target_cell.value.startswith("=")):
                                target_cell.value = ftmpl.replace("{row}", str(cur_r))
                            continue
                        if isinstance(target_cell.value, str) and target_cell.value.startswith("="):
                            raise FormulaOverwriteError(tr("server.tpl_formula_blocked", cell=f"{sh}!{col_let}{cur_r}"))
                        v = rdata.get(col["key"], "")
                        if c_it == "kwd" and isinstance(v, (int, float)):
                            target_cell.value = round(float(v), 3)
                            if not target_cell.number_format or target_cell.number_format == "General":
                                target_cell.number_format = "#,##0.000"
                        else:
                            target_cell.value = v
                        touched.append(f"{sh}!{get_column_letter(mc)}{mr}")

            wb.save(out_path)
        finally:
            wb.close()

    elif ext == ".docx":
        doc = docx.Document(str(src_path))

        # 1. Fill scalar fields (cells, paragraphs, spans)
        for fld in (active_prof.get("fields") or []):
            if fld.get("role") in ("fixed", "formula_protected"):
                continue
            key = str(fld.get("key") or "")
            if key not in fvals:
                continue
            raw_v = fvals[key]
            itype = str(fld.get("input_type") or "text")
            str_v = f"{float(raw_v):,.3f}" if (itype == "kwd" and isinstance(raw_v, (int, float))) else str(raw_v if raw_v is not None else "")
            tgt = fld.get("target") or {}
            kind = tgt.get("kind")

            if kind == "docx_cell":
                ti = int(tgt.get("table_idx") or 0)
                ri = int(tgt.get("row_idx") or 0)
                ci = int(tgt.get("col_idx") or 0)
                cell = doc.tables[ti].rows[ri].cells[ci]
                rmode = str(tgt.get("replace_mode") or "cell")
                if rmode == "span" and tgt.get("match_text"):
                    for p in cell.paragraphs:
                        _replace_docx_span_preserving_style(p, str(tgt["match_text"]), str_v)
                else:
                    if cell.paragraphs:
                        _set_docx_paragraph_text_preserving_style(cell.paragraphs[0], str_v)
                        for ep in cell.paragraphs[1:]:
                            _set_docx_paragraph_text_preserving_style(ep, "")
                touched.append(f"T{ti}:R{ri}C{ci}")

            elif kind == "docx_paragraph":
                pi = int(tgt.get("para_idx") or 0)
                para = doc.paragraphs[pi]
                if tgt.get("replace_mode") == "after_colon":
                    pfx = str(tgt.get("prefix_text") or (re.split(r"[:：]", para.text, maxsplit=1)[0] + ": "))
                    _set_docx_paragraph_text_preserving_style(para, f"{pfx}{str_v}")
                else:
                    _set_docx_paragraph_text_preserving_style(para, str_v)
                touched.append(f"P{pi}")

            elif kind == "docx_span":
                pi = int(tgt.get("para_idx") or 0)
                mtxt = str(tgt.get("match_text") or "")
                if mtxt:
                    _replace_docx_span_preserving_style(doc.paragraphs[pi], mtxt, str_v)
                    touched.append(f"P{pi}[{mtxt}]")

        # 2. Expand & fill repeating Word tables (deep-copy `<w:tr>` XML keeping exact borders & RTL!)
        for tbl_spec in (active_prof.get("repeating_tables") or []):
            if tbl_spec.get("target_kind") != "docx_rows":
                continue
            tkey = str(tbl_spec.get("key") or "items")
            rows_data = tvals.get(tkey) or []
            if not rows_data:
                continue
            ti = int(tbl_spec["table_idx"])
            dtbl = doc.tables[ti]
            tmpl_r = int(tbl_spec["template_row"])
            tot_r = int(tbl_spec["totals_row"]) if tbl_spec.get("totals_row") is not None else None
            pref_slots = max(1, int(tbl_spec.get("preformatted_rows") or ((tot_r - tmpl_r) if tot_r and tot_r > tmpl_r else max(1, len(dtbl.rows) - tmpl_r))))
            needed = len(rows_data)
            extra = max(0, needed - pref_slots)
            proto_tr = copy.deepcopy(dtbl.rows[tmpl_r]._tr)

            if extra > 0:
                anchor_idx = tot_r if (tot_r is not None and tot_r < len(dtbl.rows)) else None
                for _ in range(extra):
                    cloned_tr = copy.deepcopy(proto_tr)
                    if anchor_idx is not None:
                        dtbl.rows[anchor_idx]._tr.addprevious(cloned_tr)
                        anchor_idx += 1
                    else:
                        dtbl._tbl.append(cloned_tr)

            for i, rdata in enumerate(rows_data):
                row_obj = dtbl.rows[tmpl_r + i]
                for col in (tbl_spec.get("columns") or []):
                    ci = int(col["col_idx"])
                    if ci >= len(row_obj.cells):
                        continue
                    rv = rdata.get(col["key"], "")
                    cit = str(col.get("input_type") or "text")
                    sv = f"{float(rv):,.3f}" if (cit == "kwd" and isinstance(rv, (int, float))) else str(rv if rv is not None else "")
                    if row_obj.cells[ci].paragraphs:
                        _set_docx_paragraph_text_preserving_style(row_obj.cells[ci].paragraphs[0], sv)
                    touched.append(f"T{ti}:R{tmpl_r + i}C{ci}")

        doc.save(str(out_path))

    return {
        "ok": True,
        "output_path": str(out_path),
        "output_name": out_path.name,
        "section": active_prof["section"],
        "resolved_fields": fvals,
        "resolved_tables": tvals,
        "touched_targets": touched,
        "rebound_fields": active_prof.get("rebound_fields") or [],
    }


def advance_template_counters(
    proj_root: Path | str,
    tpl_name: str,
    used_keys: set[str] | None = None,
    *,
    device_id: str = "",
) -> dict[str, int]:
    """
    Increment `counter_next` by 1 for each `auto_counter` field in `<tpl_name>.tplmap.json`
    after a document is successfully issued. Returns `{field_key: new_next_value}`.
    """
    tpl_path = resolve_project_template(proj_root, tpl_name)
    mp = tplmap_path_for(tpl_path)
    saved = read_json_safe(mp)
    if not isinstance(saved, dict):
        return {}
    updated: dict[str, int] = {}
    changed = False
    for fld in (saved.get("fields") or []):
        if fld.get("role") != "auto_counter":
            continue
        k = str(fld.get("key") or "")
        if used_keys is not None and k not in used_keys:
            continue
        try:
            cur = max(0, int(fld.get("counter_next") if fld.get("counter_next") not in (None, "") else 1))
        except (TypeError, ValueError):
            cur = 1
        fld["counter_next"] = cur + 1
        updated[k] = cur + 1
        changed = True
    if changed:
        saved["updated_at"] = now_iso()
        if device_id:
            saved["device_id"] = device_id
        atomic_write_json(mp, saved)
    return updated

