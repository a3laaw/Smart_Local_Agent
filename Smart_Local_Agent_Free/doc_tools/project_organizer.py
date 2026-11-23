#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
project_organizer.py — Large-scale Existing Project Organizer, Deduplicator,
Log-Archive Auditor, Weekly Follow-up Exporter, and PDF Split/Merge Utility.

Designed cleanly for huge real-world contracting folders (1,000 - 10,000+ files)
on local drives or Windows Server SMB shares (`\\\\Server\\Share`):
  - Fast `os.walk` scanning with protected/system folder pruning & optional ZIP/RAR/7z extraction
  - Two-stage deduplication (`file size` grouping -> `SHA-256` only on size collisions)
  - Fast zero-cost deterministic classification + optional preview snippet extraction
  - Automatic `Superseded` detection for older revisions (`R00/R01` vs `R02`)
  - Log cross-audit against uploaded OR directly linked Excel logs (`missing_from_log`,
    `log_rows_without_files`, `sequence_gaps`)
  - Paginated plan inspection & safe background execution (`copy` or `move`) with live progress
  - Rich multi-sheet RTL Excel reports (Organization Report + Audit & Weekly Follow-up Report)
  - Optional PDF Split / Merge via `pymupdf` (`fitz`) when available
"""

from __future__ import annotations

import datetime as dt
import hashlib
import os
import re
import shutil
import threading
import time
import uuid
from collections import defaultdict
from pathlib import Path
from typing import Any, Callable

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from i18n_ar import tr
import ai_cache
import project_log as plog
from relevance import normalize_ar
import scan_archive as sa
import scan_watcher as sw
from security_utils import (
    atomic_write_json,
    normalize_win_path,
    now_iso,
    read_json_safe,
    resolve_local_or_unc,
    safe_resolve_within,
    sha256_of_file,
)
import template_profile as tprof

try:
    import pymupdf as fitz  # type: ignore
except ImportError:
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None


SUPPORTED_DOC_EXTS: frozenset[str] = (
    sw.SCANNED_EXTS | sw.NATIVE_DRAFT_EXTS | {".xlsm", ".xls", ".csv", ".txt"}
)
ARCHIVE_EXTS: frozenset[str] = frozenset({".zip", ".rar", ".7z"})

_PRUNE_DIR_PREFIXES = (
    "_System_Agent",
    "_Backups",
    "_Archived_",
    "_Auto_Projects_",
    "_Duplicates",
    "_Superseded",
    "_Reports",
    "Organized_Archive",
    ".",
    "$RECYCLE.BIN",
    "System Volume Information",
    "DfsrPrivate",
    "#recycle",
    "#snapshot",
    "__pycache__",
    "node_modules",
)

_INVALID_FS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
_REV_RE = re.compile(r"(?:^|[-_\s.(/])(?:REV\.?\s*|R)0*(\d{1,2})(?=$|[-_\s.)/])", re.I)
_REF_PATTERNS = [
    re.compile(r"(?<![A-Za-z0-9])((?:MAR|MIR|WIR|IR|SD|MS|PQ|IPC|VO|PO|RFI|NCR|OUT|IN|PC|RPT)[-\s_./]*(?:(?:CIV|STR|ARCH|EL|ELEC|PL|HVAC|HAVC|FF|ELEV|ELVE|MEC|MEP|GEN|[A-Z]{2,4})[-\s_./]*)?0*\d{1,6})(?![A-Za-z0-9])", re.I),
    re.compile(r"(?<![A-Za-z0-9])((?:CIV|ARCH|STR|EL|ELEC|PL|HVAC|HAVC|FF|ELEV|ELVE|MEC)[-\s_./]+0*\d{1,6})(?![A-Za-z0-9])", re.I),
    re.compile(r"([\u0600-\u06FF]{1,4}\s*/\s*(?:20\d{2}\s*/\s*0*\d{1,6}|0*\d{1,6}\s*/\s*20\d{2}))"),
]
_SEQ_NUM_RE = re.compile(r"^\s*([A-Za-z]{2,8})(?:[-\s_]+([A-Za-z]{2,5}))?[-\s_]*0*(\d{1,6})\s*$")


def _clean_seg(s: str, max_len: int = 60) -> str:
    c = _INVALID_FS_CHARS.sub("-", str(s or "")).strip(" ._-")
    c = re.sub(r"\s+", " ", c)
    return c[:max_len].strip() if c else "DOC"


def organizer_plan_path(proj_root: Path | str) -> Path:
    return resolve_local_or_unc(proj_root) / "_System_Agent" / "organizer_plan.json"


def reports_dir_for_project(proj_root: Path | str) -> Path:
    d = resolve_local_or_unc(proj_root) / "_Reports"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _walk_folder_files(
    src_root: Path,
    *,
    include_archives: bool = True,
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
    emit_offset: int = 0,
    folders_offset: int = 0,
    archives_offset: int = 0,
    counters_out: dict[str, int] | None = None,
) -> tuple[list[Path], list[Path]]:
    """
    Fast `os.walk` over local or Windows Server UNC folders (`onerror` resilient),
    emitting live folder/file discovery counters so large folders show real-time activity.
    """
    docs: list[Path] = []
    archives: list[Path] = []
    folders_scanned = 0
    last_emit = time.monotonic()

    def _on_err(_exc: OSError) -> None:
        return None

    def _emit_walk(cur_label: str, force: bool = False) -> None:
        nonlocal last_emit
        if not progress_cb:
            return
        now = time.monotonic()
        if not force and (now - last_emit) < 0.25 and (len(docs) % 20 != 0):
            return
        last_emit = now
        progress_cb({
            "phase": "walking",
            "scanned": emit_offset + len(docs),
            "discovered": emit_offset + len(docs),
            "archives_found": archives_offset + len(archives),
            "folders_scanned": folders_offset + folders_scanned,
            "total": 0,
            "current_file": cur_label,
        })

    for dirpath, dirnames, filenames in os.walk(str(src_root), onerror=_on_err):
        if stop_event is not None and stop_event.is_set():
            break
        folders_scanned += 1
        dirnames[:] = sorted(
            d for d in dirnames
            if not any(d.lower().startswith(p.lower()) for p in _PRUNE_DIR_PREFIXES)
        )
        dp = Path(dirpath)
        try:
            rel_dir = dp.relative_to(src_root).as_posix()
        except ValueError:
            rel_dir = dp.name
        cur_dir_label = src_root.name if rel_dir in ("", ".") else rel_dir
        _emit_walk(cur_dir_label)

        for fn in sorted(filenames):
            if stop_event is not None and stop_event.is_set():
                break
            if fn.startswith(("~$", ".")):
                continue
            fp = normalize_win_path(dp / fn)
            ext = fp.suffix.lower()
            if ext in SUPPORTED_DOC_EXTS:
                docs.append(fp)
                _emit_walk(f"{cur_dir_label}/{fn}" if rel_dir not in ("", ".") else fn)
            elif include_archives and ext in ARCHIVE_EXTS:
                archives.append(fp)
                _emit_walk(f"{cur_dir_label}/{fn}" if rel_dir not in ("", ".") else fn)

    _emit_walk(src_root.name, force=True)
    if counters_out is not None:
        counters_out["folders_scanned"] = folders_scanned
    return docs, archives


def _read_fast_text_snippet(fp: Path, max_chars: int = 1500) -> str:
    ext = fp.suffix.lower()
    try:
        if ext == ".txt":
            return fp.read_text("utf-8", errors="ignore")[:max_chars]
        if ext == ".pdf" and fitz is not None:
            doc = fitz.open(str(fp))
            parts: list[str] = []
            for i in range(min(len(doc), 2)):
                t = doc[i].get_text("text").strip()
                if t:
                    parts.append(t)
                if sum(len(x) for x in parts) >= max_chars:
                    break
            doc.close()
            return "\n".join(parts)[:max_chars]
        if ext in (".docx", ".xlsx", ".xlsm"):
            return (sw.extract_native_office_text(fp) or "")[:max_chars]
    except Exception:
        return ""
    return ""


def _canonical_ref(raw: str) -> str:
    s = re.sub(r"[\s_./]+", "-", raw.strip().upper())
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s


_COPY_PREFIX_RE = re.compile(r"^(?:copy\s+of|نسخة\s+من|نسخه\s+من|مكرر)\s+", re.I)
_COPY_SUFFIX_RE = re.compile(
    r"(?:\s*[\(\[\{]\s*\d+\s*[\)\]\}]|\s*[-_–]\s*(?:copy|dup|old|backup|final|new|نسخة|نسخه|مكرر)(?:\s*\d*)?|\s+(?:copy|dup|نسخة|نسخه|مكرر)\s*\d*)$",
    re.I,
)


def _normalize_stem_key(stem: str) -> str:
    """Normalize a filename stem for smart near-duplicate matching while keeping reference/rev numbers intact."""
    s = str(stem or "").strip()
    if not s:
        return ""
    # Strip bracketed copy indices like `(1)` and copy prefixes/suffixes BEFORE normalize_ar removes punctuation
    prev = None
    while prev != s:
        prev = s
        s = _COPY_PREFIX_RE.sub("", s).strip()
        s = _COPY_SUFFIX_RE.sub("", s).strip()
    s = normalize_ar(s).strip()
    prev = None
    while prev != s:
        prev = s
        s = _COPY_PREFIX_RE.sub("", s).strip()
        s = _COPY_SUFFIX_RE.sub("", s).strip()
    s = re.sub(r"[\s_\-]+", " ", s).strip()
    return s


def _normalize_content_sig(text: str) -> str:
    """Compute compact SHA-256 signature of normalized document content (>= 50 meaningful chars)."""
    raw = normalize_ar(str(text or ""))
    compact = re.sub(r"[\W_]+", "", raw)
    if len(compact) < 50:
        return ""
    return hashlib.sha256(compact[:1400].encode("utf-8")).hexdigest()[:24]


_SECTION_SHORT_TAGS: dict[str, str] = {
    "submittal_mar": "MAR",
    "submittal_mir": "MIR",
    "submittal_ir": "WIR",
    "submittal_sd": "SD",
    "submittal_ms": "MS",
    "outgoing": "OUT",
    "incoming": "IN",
    "contract_ipc": "IPC",
    "variation_vo": "VO",
    "purchase_po": "PO",
    "petty_cash": "PC",
    "site_report": "RPT",
    "ncr": "NCR",
    "rfi": "RFI",
    "permit": "PRM",
    "general": "GEN",
}
_DATE_YMD_RE = re.compile(r"\b(20[0-3]\d)[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])\b")
_DATE_DMY_RE = re.compile(r"\b(0?[1-9]|[12]\d|3[01])[-/.](0?[1-9]|1[0-2])[-/.](20[0-3]\d)\b")
_DATE_COMPACT_RE = re.compile(r"(?<!\d)(20[0-3]\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d)")
_WA_NAME_RE = re.compile(r"whatsapp\s+(?:image|video|audio|document)\s+\d{4}[-.\d]+\s+at\s+[\d.\s]+(?:am|pm)?", re.I)
_GENERIC_STEM_WORDS = frozenset({
    "scan", "scanned", "img", "image", "photo", "pic", "picture", "doc", "document",
    "file", "page", "untitled", "new", "old", "copy", "dup", "backup", "final",
    "whatsapp", "camscanner", "adobe", "pdf", "extra", "misc", "other", "temp",
    "letter", "memo", "book", "sheet", "note",
    "سكان", "صوره", "صورة", "ملف", "مستند", "ورقه", "ورقة", "صفحه", "صفحة",
    "كتاب", "خطاب", "مذكره", "مذكرة", "نموذج", "ورق", "بدون", "اسم",
    "اضافي", "إضافي", "قديم", "جديد", "نسخه", "نسخة", "مكرر", "عام", "متنوع",
})
_SUBJ_LINE_RE = re.compile(r"^(?:الموضوع|بخصوص|البيان|العنوان|Subject|RE|Title)\s*[:：\-–]\s*(.+)$", re.I)
_SKIP_TITLE_LINE_RE = re.compile(
    r"^(?:بسم الله|دولة الكويت|وزارة|التاريخ|الموافق|الرقم|الإشارة|رقم الكتاب|السادة|الموقرين|تحية طيبة|السلام عليكم|Page\s+\d+|Date\s*:|Ref\s*:)",
    re.I,
)


def _section_short_tag(sec_id: str) -> str:
    s = str(sec_id or "general").strip().lower()
    if s in _SECTION_SHORT_TAGS:
        return _SECTION_SHORT_TAGS[s]
    clean = re.sub(r"[^A-Za-z0-9]", "", s).upper()
    return (clean[:4] or "GEN")


def _clean_subject_phrase(raw: str) -> str:
    s = str(raw or "").strip()
    if not s:
        return ""
    s = _WA_NAME_RE.sub(" ", s)
    prev = None
    while prev != s:
        prev = s
        s = _COPY_PREFIX_RE.sub("", s).strip()
        s = _COPY_SUFFIX_RE.sub("", s).strip()
    s = _REV_RE.sub(" ", s)
    s = re.sub(r"[_\-]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" ._-")
    return s


def _is_generic_stem(stem: str) -> bool:
    s = _clean_subject_phrase(stem)
    if not s:
        return True
    s_norm = normalize_ar(s).lower()
    tokens = [t for t in re.split(r"[\s_\-]+", s_norm) if t]
    meaningful = [
        t for t in tokens
        if t not in _GENERIC_STEM_WORDS and not t.isdigit() and len(t) >= 2
    ]
    return len(meaningful) == 0 or len("".join(meaningful)) < 4


def _extract_doc_date(corpus: str, fp: Path | None = None) -> str:
    txt = str(corpus or "")
    m = _DATE_YMD_RE.search(txt)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m2 = _DATE_DMY_RE.search(txt)
    if m2:
        return f"{int(m2.group(3)):04d}-{int(m2.group(2)):02d}-{int(m2.group(1)):02d}"
    m3 = _DATE_COMPACT_RE.search(txt)
    if m3:
        return f"{int(m3.group(1)):04d}-{int(m3.group(2)):02d}-{int(m3.group(3)):02d}"
    if fp is not None:
        try:
            return dt.date.fromtimestamp(fp.stat().st_mtime).isoformat()
        except OSError:
            pass
    return ""


def _extract_subject_from_text(snippet: str) -> tuple[str, bool]:
    """Return `(best_subject, is_explicit_subject_header)` from document text snippet."""
    if not snippet:
        return "", False
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in snippet.splitlines()]
    lines = [ln for ln in lines if ln]
    for ln in lines[:25]:
        m = _SUBJ_LINE_RE.match(ln)
        if m:
            cand = _clean_subject_phrase(m.group(1))
            if len(cand) >= 4 and re.search(r"[A-Za-z\u0600-\u06FF]{2,}", cand):
                return cand[:90], True
    for ln in lines[:18]:
        if _SKIP_TITLE_LINE_RE.match(ln):
            continue
        cand = _clean_subject_phrase(ln)
        if 8 <= len(cand) <= 95 and re.search(r"[A-Za-z\u0600-\u06FF]{4,}", cand) and not _is_generic_stem(cand):
            return cand[:90], False
    return "", False


def _apply_custom_filename(raw_name: str, ext: str) -> str:
    """Sanitize a user/AI-supplied filename while keeping internal dots like `2026.05.14` or `Rev.01` intact."""
    s = str(raw_name or "").strip()
    if ext and s.lower().endswith(ext.lower()):
        s = s[: -len(ext)].rstrip(" ._-")
    return _clean_seg(s, 110) + ext


def _extract_metadata_fast(
    fp: Path,
    src_root: Path,
    proj: dict[str, Any],
    *,
    read_text: bool = True,
) -> dict[str, Any]:
    """Deterministic, high-speed classifier suitable for 2,000+ project files."""
    try:
        rel_parts = fp.relative_to(src_root).parts
        folder_hint = " / ".join(rel_parts[:-1])
    except ValueError:
        folder_hint = fp.parent.name

    fname = fp.stem
    quick_corpus = f"{folder_hint} {fname}"
    snippet = ""
    # Prefer filename first, then folder hint, then fast document text snippet
    ref_no = ""
    for src_txt in (fname, folder_hint):
        if not src_txt:
            continue
        for pat in _REF_PATTERNS:
            m = pat.search(src_txt)
            if m:
                ref_no = _canonical_ref(m.group(1))
                break
        if ref_no:
            break

    if read_text and (not ref_no or len(fname) < 12 or _is_generic_stem(fname) or fp.suffix.lower() in (".txt", ".docx", ".xlsx", ".xlsm")):
        snippet = _read_fast_text_snippet(fp, 1400)
        if not ref_no and snippet:
            for pat in _REF_PATTERNS:
                m = pat.search(snippet)
                if m:
                    ref_no = _canonical_ref(m.group(1))
                    break

    corpus = f"{quick_corpus} {snippet}".strip()
    corpus_l = corpus.lower()
    corpus_ar = normalize_ar(corpus)
    doc_date = _extract_doc_date(f"{fname} {snippet[:900]}", fp)

    # Revision number
    rev_num = 0
    for rm in _REV_RE.finditer(quick_corpus + " " + snippet[:500]):
        try:
            rev_num = max(rev_num, int(rm.group(1)))
        except ValueError:
            pass

    # Discipline
    disc = ""
    for m_d in re.finditer(r"\b(CIV|STR|ARCH|EL|ELEC|PL|HVAC|HAVC|FF|ELEV|ELVE|MEC)\b", corpus, re.I):
        disc = plog.canonical_discipline(m_d.group(1))
        if disc:
            break
    if not disc:
        for kw, cd in (("كهرب", "EL"), ("مدن", "CIV"), ("انشائ", "STR"), ("معمار", "ARCH"), ("صح", "PL"), ("تكييف", "HVAC"), ("حريق", "FF"), ("مصعد", "ELEV")):
            if kw in corpus_ar:
                disc = cd
                break

    # Section inference
    sec_map = tprof.sections_map_for_project(proj)
    section = ""
    ref_up = ref_no.upper()
    if ref_up.startswith("MAR"):
        section = "submittal_mar"
    elif ref_up.startswith("MIR"):
        section = "submittal_mir"
    elif ref_up.startswith(("WIR", "IR")):
        section = "submittal_ir"
    elif ref_up.startswith("SD"):
        section = "submittal_sd"
    elif ref_up.startswith(("MS", "PQ")):
        section = "submittal_ms"
    elif ref_up.startswith("OUT"):
        section = "outgoing"
    elif ref_up.startswith("IN"):
        section = "incoming"
    elif ref_up.startswith("IPC"):
        section = "contract_ipc"
    elif ref_up.startswith("VO"):
        section = "variation_vo"
    elif ref_up.startswith("PO"):
        section = "purchase_po"
    elif ref_up.startswith("PC"):
        section = "petty_cash"
    elif ref_up.startswith("RPT"):
        section = "site_report"

    if not section:
        sugg = tprof.suggest_template_section(f"{folder_hint} {fname} {snippet[:400]}".strip())
        if sugg.get("suggested_section") in sec_map and float(sugg.get("confidence", 0)) >= 0.35:
            section = str(sugg["suggested_section"])

    if not section:
        if any(w in corpus_ar for w in ("صادر", "مرسل", "outgoing", "sent")):
            section = "outgoing"
        elif any(w in corpus_ar for w in ("وارد", "مستلم", "incoming", "received")):
            section = "incoming"
        elif any(w in corpus_ar for w in ("مستخلص", "دفعه", "payment", "ipc")):
            section = "contract_ipc"
        elif any(w in corpus_ar for w in ("تغييري", "variation", "vo")):
            section = "variation_vo"
        elif any(w in corpus_ar for w in ("امر شراء", "purchase order", "مقارنه اسعار", "عروض اسعار", "عرض سعر", "تسعير", "quotation", "price comparison")):
            section = "purchase_po"
        elif any(w in corpus_ar for w in ("عهد", "فواتير", "فاتور", "ايصال", "petty", "invoice", "receipt")):
            section = "petty_cash"
        elif any(w in corpus_ar for w in ("تقرير", "يومي", "اسبوعي", "شهري", "محضر", "report", "minutes")):
            section = "site_report"
        elif any(w in corpus_ar for w in ("اعتماد مواد", "عينة", "material")):
            section = "submittal_mar"
        elif any(w in corpus_ar for w in ("فحص مواد", "توريد")):
            section = "submittal_mir"
        elif any(w in corpus_ar for w in ("استلام اعمال", "تفتيش", "inspection")):
            section = "submittal_ir"
        elif any(w in corpus_ar for w in ("مخطط", "لوح", "shop drawing")):
            section = "submittal_sd"
        elif any(w in corpus_ar for w in ("طريقة عمل", "منهجية", "method statement")):
            section = "submittal_ms"

    confident = bool(section and section in sec_map)
    if not confident:
        section = "general"

    # Approval / MOH hint
    approval = ""
    if re.search(r"\bcode\s*a\b|\bapproved\b", corpus_l) and "not" not in corpus_l:
        approval = "A"
    elif re.search(r"\bcode\s*b\b|approved as noted", corpus_l):
        approval = "B"
    elif re.search(r"\bcode\s*c\b|revise", corpus_l):
        approval = "C"
    elif re.search(r"\bcode\s*d\b|\brejected\b", corpus_l):
        approval = "D"
    elif "معتمد" in corpus_ar and "غير" not in corpus_ar:
        approval = "A"
    elif "مرفوض" in corpus_ar:
        approval = "D"

    authority = "MOH" if any(k in corpus_ar for k in ("الوزاره", "الاشغال", "moh", "mpw", "mow")) else ""

    # Clean subject: prefer document content subject when filename is generic/noisy or has explicit Subject: header
    subj_clean = fname
    if ref_no:
        subj_clean = re.sub(re.escape(ref_no), "", subj_clean, flags=re.I)
    subj_clean = _clean_subject_phrase(subj_clean)
    txt_subj, is_explicit_subj = _extract_subject_from_text(snippet)
    if txt_subj and (_is_generic_stem(subj_clean) or is_explicit_subj or len(subj_clean) < 4):
        subj_clean = txt_subj
    if len(subj_clean) < 3:
        subj_clean = _clean_subject_phrase(fname) or fname

    return {
        "section": section,
        "confident": confident,
        "ref_no": ref_no,
        "ref_key": plog.normalize_ref_key(ref_no) if ref_no else "",
        "revision": f"R{rev_num:02d}",
        "rev_num": rev_num,
        "doc_date": doc_date,
        "discipline": disc,
        "approval_code": approval,
        "authority": authority,
        "subject": subj_clean[:120],
        "folder_hint": folder_hint,
        "stem_key": _normalize_stem_key(fname),
        "text_sig": _normalize_content_sig(snippet),
        "needs_ai": bool(not snippet and (not confident or not ref_no or _is_generic_stem(fname))),
    }


def _build_dest_rel(
    meta: dict[str, Any],
    src_name: str,
    ext: str,
    proj: dict[str, Any],
    *,
    status: str = "keep",
) -> tuple[str, str]:
    """Compute `(dest_folder_rel, proposed_filename)` cleanly inside project sections."""
    sec_map = tprof.sections_map_for_project(proj)
    sec = str(meta.get("section") or "general")
    sec_info = sec_map.get(sec) or sec_map.get("general") or {"archive_folder": "09_General_عام"}
    base_folder = str(sec_info.get("archive_folder") or "09_General_عام").strip("/\\")

    if status == "duplicate":
        return "_Duplicates_ملفات_مكررة", _clean_seg(Path(src_name).stem, 75) + ext
    if status == "near_duplicate":
        return "_Duplicates_ملفات_مكررة/By_Name_Data", _clean_seg(Path(src_name).stem, 75) + ext
    if status == "superseded":
        sub = f"{base_folder}/_Superseded_مراجعات_سابقة"
    else:
        sub = base_folder
        if meta.get("discipline") and sec.startswith("submittal"):
            sub = f"{base_folder}/{_clean_seg(meta['discipline'], 12)}"

    parts_nm: list[str] = []
    has_ref = bool(str(meta.get("ref_no") or "").strip())
    if has_ref:
        parts_nm.append(_clean_seg(meta["ref_no"], 30))
        if meta.get("rev_num", 0) > 0 or sec.startswith("submittal"):
            parts_nm.append(_clean_seg(meta.get("revision") or "R00", 8))
    else:
        ddate = str(meta.get("doc_date") or "").strip()
        if ddate:
            parts_nm.append(_clean_seg(ddate, 16))
        sec_tag = _section_short_tag(sec)
        parts_nm.append(sec_tag)
        if meta.get("rev_num", 0) > 0:
            parts_nm.append(_clean_seg(meta.get("revision") or "R00", 8))

    if meta.get("authority") == "MOH":
        parts_nm.append("MOH")

    raw_subj = str(meta.get("subject") or "").strip()
    if not raw_subj:
        raw_subj = _clean_subject_phrase(Path(src_name).stem) or Path(src_name).stem
    subj_seg = _clean_seg(raw_subj, 48).replace(" ", "_")
    # Avoid duplicating date/ref/prefix already placed in parts_nm
    if subj_seg:
        subj_low = subj_seg.lower()
        already_covered = (
            (has_ref and subj_low in parts_nm[0].lower())
            or (not has_ref and meta.get("doc_date") and subj_low == str(meta["doc_date"]).lower())
        )
        if not already_covered:
            parts_nm.append(subj_seg)
    new_stem = "_".join(p for p in parts_nm if p) or _clean_seg(Path(src_name).stem, 60)
    return sub, f"{new_stem[:110]}{ext}"


def _collect_log_refs_and_rows(proj: dict[str, Any], *, device_id: str = "") -> dict[str, Any]:
    """Read all existing reference numbers & rows from uploaded + directly linked logs."""
    sum_data = plog.read_project_log_summary(proj, limit=50000, device_id=device_id)
    rows = sum_data.get("rows") or []
    ref_map: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for r in rows:
        rk = plog.normalize_ref_key(str(r.get("ref_no") or ""))
        if rk:
            ref_map[rk].append(r)
    return {"rows": rows, "ref_map": ref_map, "stats": sum_data.get("stats") or {}}


def _find_sequence_gaps(ref_strings: list[str]) -> list[dict[str, Any]]:
    """Detect missing numbers inside numeric series like MAR-EL-001..140 or OUT-001..095."""
    series: dict[str, set[int]] = defaultdict(set)
    for rs in ref_strings:
        m = _SEQ_NUM_RE.match(str(rs or ""))
        if not m:
            continue
        pfx = m.group(1).upper()
        disc = (m.group(2) or "").upper()
        num = int(m.group(3))
        if 0 < num < 50000:
            key = f"{pfx}-{disc}" if disc else pfx
            series[key].add(num)

    gaps: list[dict[str, Any]] = []
    for sk, nums in sorted(series.items()):
        if len(nums) < 3:
            continue
        hi = max(nums)
        lo = min(nums)
        if hi - lo > 5000:
            continue
        missing = [n for n in range(1, hi + 1) if n not in nums]
        if missing:
            gaps.append({
                "series": sk,
                "min": lo,
                "max": hi,
                "count_present": len(nums),
                "missing_count": len(missing),
                "missing_numbers": missing[:80],
            })
    return gaps


def build_organization_plan(
    proj: dict[str, Any],
    src_folder: str,
    *,
    extract_archives: bool = True,
    read_text: bool = True,
    device_id: str = "",
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """
    Scan a messy/large project folder (1,000 - 10,000+ files), detect exact duplicates
    and superseded revisions, classify files deterministically, compare against project
    Excel logs, and save the reviewable plan to `_System_Agent/organizer_plan.json`.
    """
    proj_root = resolve_local_or_unc(str(proj["root"]))
    src_root = resolve_local_or_unc(str(src_folder or "").strip())
    if not str(src_folder or "").strip() or not src_root.exists() or not src_root.is_dir():
        raise ValueError(tr("server.folder_not_found", path=str(src_folder)))

    walk_counts: dict[str, int] = {}
    doc_files, arc_files = _walk_folder_files(
        src_root,
        include_archives=extract_archives,
        stop_event=stop_event,
        progress_cb=progress_cb,
        counters_out=walk_counts,
    )
    total_folders = int(walk_counts.get("folders_scanned", 0))

    extracted_dir = proj_root / "_System_Agent" / "_org_extracted"
    extracted_archives: list[dict[str, Any]] = []
    if extract_archives and arc_files:
        shutil.rmtree(extracted_dir, ignore_errors=True)
        extracted_dir.mkdir(parents=True, exist_ok=True)
        for idx, af in enumerate(arc_files):
            if stop_event is not None and stop_event.is_set():
                break
            if progress_cb:
                progress_cb({
                    "phase": "extracting",
                    "scanned": idx + 1,
                    "total": len(arc_files),
                    "discovered": len(doc_files),
                    "archives_found": len(arc_files),
                    "folders_scanned": total_folders,
                    "current_file": af.name,
                })
            sub_out = extracted_dir / f"{idx + 1:02d}_{_clean_seg(af.stem, 32)}"
            ok_ex, tool_used = sa.extract_archive(af, sub_out)
            extracted_archives.append({"archive": af.name, "ok": ok_ex, "tool": tool_used})
            if ok_ex:
                ex_counts: dict[str, int] = {}
                ex_docs, _ = _walk_folder_files(
                    sub_out,
                    include_archives=False,
                    stop_event=stop_event,
                    progress_cb=progress_cb,
                    emit_offset=len(doc_files),
                    folders_offset=total_folders,
                    archives_offset=len(arc_files),
                    counters_out=ex_counts,
                )
                doc_files.extend(ex_docs)
                total_folders += int(ex_counts.get("folders_scanned", 0))

    total_files = len(doc_files)
    if progress_cb:
        progress_cb({
            "phase": "hashing",
            "scanned": 0,
            "total": total_files,
            "discovered": total_files,
            "archives_found": len(arc_files),
            "folders_scanned": total_folders,
            "current_file": "",
        })

    # Stage 1: Group by file size so SHA-256 runs ONLY on size collisions!
    by_size: dict[int, list[Path]] = defaultdict(list)
    file_sizes: dict[Path, int] = {}
    file_mtimes: dict[Path, float] = {}
    last_stat_emit = time.monotonic()
    for s_idx, fp in enumerate(doc_files, 1):
        if stop_event is not None and stop_event.is_set():
            break
        if progress_cb and (s_idx == 1 or s_idx == total_files or s_idx % 40 == 0 or (time.monotonic() - last_stat_emit) >= 0.25):
            last_stat_emit = time.monotonic()
            progress_cb({
                "phase": "hashing",
                "scanned": s_idx,
                "total": total_files,
                "discovered": total_files,
                "current_file": fp.name,
            })
        try:
            st = fp.stat()
            file_sizes[fp] = int(st.st_size)
            file_mtimes[fp] = float(st.st_mtime)
            by_size[int(st.st_size)].append(fp)
        except OSError:
            continue

    hashes: dict[Path, str] = {}
    by_hash: dict[str, list[Path]] = defaultdict(list)
    hash_cands = [fp for grp in by_size.values() if len(grp) >= 2 for fp in grp]
    hash_total = len(hash_cands)
    last_hash_emit = time.monotonic()
    for h_idx, fp in enumerate(hash_cands, 1):
        if stop_event is not None and stop_event.is_set():
            break
        if progress_cb and (h_idx == 1 or h_idx == hash_total or h_idx % 8 == 0 or (time.monotonic() - last_hash_emit) >= 0.25):
            last_hash_emit = time.monotonic()
            progress_cb({
                "phase": "hashing",
                "scanned": h_idx,
                "total": hash_total,
                "discovered": total_files,
                "current_file": fp.name,
            })
        try:
            fh = sha256_of_file(fp)
            hashes[fp] = fh
            by_hash[fh].append(fp)
        except OSError:
            continue

    def _dup_rank(p: Path) -> tuple[int, int, float]:
        low = str(p).lower()
        penalty = 1 if any(k in low for k in ("copy", "نسخة", "old", "backup", "dup", "temp", "مكرر")) else 0
        return (penalty, len(str(p)), -file_mtimes.get(p, 0.0))

    dup_winner_of: dict[Path, str] = {}
    for fh, grp in by_hash.items():
        if len(grp) < 2:
            continue
        grp_sorted = sorted(grp, key=_dup_rank)
        winner = grp_sorted[0]
        for loser in grp_sorted[1:]:
            dup_winner_of[loser] = winner.name

    # Stage 2: Fast deterministic classification
    if progress_cb:
        progress_cb({
            "phase": "analyzing",
            "scanned": 0,
            "total": total_files,
            "discovered": total_files,
            "current_file": proj.get("name") or "",
        })
    log_ctx = _collect_log_refs_and_rows(proj, device_id=device_id)
    log_ref_map = log_ctx["ref_map"]

    items: list[dict[str, Any]] = []
    last_an_emit = time.monotonic()
    for i, fp in enumerate(doc_files, 1):
        if stop_event is not None and stop_event.is_set():
            break
        if progress_cb and (i == 1 or i % 10 == 0 or i == total_files or (time.monotonic() - last_an_emit) >= 0.25):
            last_an_emit = time.monotonic()
            progress_cb({
                "phase": "analyzing",
                "scanned": i,
                "total": total_files,
                "discovered": total_files,
                "current_file": fp.name,
            })

        root_base = extracted_dir if fp.is_relative_to(extracted_dir) else src_root
        try:
            src_rel = fp.relative_to(root_base).as_posix()
        except ValueError:
            src_rel = fp.name

        meta = _extract_metadata_fast(fp, root_base, proj, read_text=read_text)
        is_dup = fp in dup_winner_of
        status = "duplicate" if is_dup else "keep"
        dest_folder, new_name = _build_dest_rel(meta, fp.name, fp.suffix.lower(), proj, status=status)

        rk = meta["ref_key"]
        if not rk:
            log_status = "no_ref"
        elif rk in log_ref_map:
            log_status = "matched"
        else:
            log_status = "missing_from_log"

        items.append({
            "id": f"ORG-{i:05d}",
            "src_path": str(fp),
            "src_rel": src_rel,
            "file_name": fp.name,
            "ext": fp.suffix.lower(),
            "size_bytes": int(file_sizes.get(fp, 0)),
            "size_kb": round(file_sizes.get(fp, 0) / 1024, 1),
            "sha256": hashes.get(fp, ""),
            "status": status,
            "decision": "pending" if is_dup else "",
            "include": not is_dup,
            "duplicate_of": dup_winner_of.get(fp, ""),
            "section": meta["section"],
            "confident": meta["confident"],
            "ref_no": meta["ref_no"],
            "ref_key": rk,
            "revision": meta["revision"],
            "rev_num": meta["rev_num"],
            "doc_date": meta.get("doc_date") or "",
            "discipline": meta["discipline"],
            "approval_code": meta["approval_code"],
            "authority": meta["authority"],
            "subject": meta["subject"],
            "dest_folder": dest_folder,
            "new_name": new_name,
            "dest_rel": f"{dest_folder}/{new_name}",
            "log_status": log_status,
            "executed": False,
            "custom_name": False,
            "ai_renamed": False,
            "ai_from_cache": False,
            "needs_ai": bool(meta.get("needs_ai")),
            "stem_key": meta.get("stem_key") or "",
            "text_sig": meta.get("text_sig") or "",
        })

    # Stage 2.5: Hydrate any previously cached AI results by exact file SHA-256 (0 AI calls!)
    _org_cache_dir = ai_cache.project_ai_cache_dir(proj_root)
    _org_cached_sizes = ai_cache.cached_sizes(_org_cache_dir, "organizer_ai")
    if _org_cached_sizes:
        _sec_map_org = tprof.sections_map_for_project(proj)
        _org_sec_sig = ai_cache.stable_sig(sorted(_sec_map_org.keys()), proj.get("domain"))
        for it in items:
            if it.get("status") != "keep":
                continue
            sb = int(it.get("size_bytes") or 0)
            if sb not in _org_cached_sizes and not it.get("sha256"):
                continue
            fh = str(it.get("sha256") or "")
            if not fh:
                try:
                    fh = sha256_of_file(resolve_local_or_unc(str(it.get("src_path") or "")))
                    it["sha256"] = fh
                except OSError:
                    continue
            c_rec = ai_cache.get_cached(_org_cache_dir, "organizer_ai", f"{fh}:{_org_sec_sig}")
            if c_rec and isinstance(c_rec.get("meta_ai"), dict):
                if _apply_organizer_ai_meta(it, c_rec["meta_ai"], proj, _sec_map_org, log_ref_map):
                    it["ai_from_cache"] = True

    # Stage 3: Mark older revisions as `superseded`, or equal revisions as `near_duplicate`
    by_ref_sec: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for it in items:
        if it["status"] == "keep" and it["ref_key"]:
            by_ref_sec[(it["section"], it["ref_key"])].append(it)

    for (_sec, _rk), grp in by_ref_sec.items():
        if len(grp) < 2:
            continue
        grp.sort(key=lambda x: (int(x.get("rev_num") or 0), float(x.get("size_kb") or 0)), reverse=True)
        top_rev = int(grp[0].get("rev_num") or 0)
        for older in grp[1:]:
            if int(older.get("rev_num") or 0) < top_rev:
                older["status"] = "superseded"
                older["decision"] = "pending"
                older["superseded_by"] = grp[0]["file_name"]
                df, nn = _build_dest_rel(older, older["file_name"], older["ext"], proj, status="superseded")
                older["dest_folder"] = df
                older["new_name"] = nn
                older["dest_rel"] = f"{df}/{nn}"
            else:
                older["status"] = "near_duplicate"
                older["decision"] = "pending"
                older["include"] = False
                older["near_dup_of"] = grp[0]["file_name"]
                older["near_dup_reason"] = "ref"
                df, nn = _build_dest_rel(older, older["file_name"], older["ext"], proj, status="near_duplicate")
                older["dest_folder"] = df
                older["new_name"] = nn
                older["dest_rel"] = f"{df}/{nn}"

    # Stage 3.5: Smart near-duplicate detection by normalized filename + document text content
    keep_items = [it for it in items if it["status"] == "keep"]
    if len(keep_items) >= 2:
        parent = list(range(len(keep_items)))

        def _uf_find(x: int) -> int:
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        def _uf_union(a: int, b: int) -> None:
            ra, rb = _uf_find(a), _uf_find(b)
            if ra != rb:
                parent[rb] = ra

        by_stem: dict[str, list[int]] = defaultdict(list)
        by_tsig: dict[str, list[int]] = defaultdict(list)
        for idx_k, it_k in enumerate(keep_items):
            sk = str(it_k.get("stem_key") or "")
            if len(sk) >= 4:
                by_stem[sk].append(idx_k)
            ts = str(it_k.get("text_sig") or "")
            if ts:
                by_tsig[ts].append(idx_k)

        for idxs in by_stem.values():
            if len(idxs) >= 2:
                for other in idxs[1:]:
                    _uf_union(idxs[0], other)
        for idxs in by_tsig.values():
            if len(idxs) >= 2:
                for other in idxs[1:]:
                    _uf_union(idxs[0], other)

        comps: dict[int, list[dict[str, Any]]] = defaultdict(list)
        for idx_k, it_k in enumerate(keep_items):
            comps[_uf_find(idx_k)].append(it_k)

        ext_prio = {".pdf": 0, ".docx": 1, ".xlsx": 1, ".xlsm": 1, ".doc": 2, ".xls": 2, ".txt": 3}

        def _near_rank(it_x: dict[str, Any]) -> tuple[int, int, float, int]:
            low = str(it_x.get("file_name") or "").lower()
            has_copy_kw = 1 if any(k in low for k in ("copy", "نسخة", "مكرر", "old", "backup", "(1)", "(2)", "(3)")) else 0
            ep = ext_prio.get(str(it_x.get("ext") or "").lower(), 4)
            return (has_copy_kw, ep, -float(it_x.get("size_kb") or 0), len(low))

        for comp_items in comps.values():
            if len(comp_items) < 2:
                continue
            comp_sorted = sorted(comp_items, key=_near_rank)
            winner = comp_sorted[0]
            for loser in comp_sorted[1:]:
                same_name = bool(loser.get("stem_key") and loser.get("stem_key") == winner.get("stem_key"))
                same_cont = bool(loser.get("text_sig") and loser.get("text_sig") == winner.get("text_sig"))
                if same_name and same_cont:
                    reason = "name_and_content"
                elif same_cont:
                    reason = "content"
                else:
                    reason = "name"
                loser["status"] = "near_duplicate"
                loser["decision"] = "pending"
                loser["include"] = False
                loser["near_dup_of"] = winner["file_name"]
                loser["near_dup_reason"] = reason
                df, nn = _build_dest_rel(loser, loser["file_name"], loser["ext"], proj, status="near_duplicate")
                loser["dest_folder"] = df
                loser["new_name"] = nn
                loser["dest_rel"] = f"{df}/{nn}"

    # Audit against log: rows in Excel that have NO file in the scanned folder!
    seen_ref_keys = {it["ref_key"] for it in items if it["ref_key"]}
    log_rows_missing_file = [
        {
            "log_name": r.get("log_name", ""),
            "sheet": r.get("sheet", ""),
            "row": r.get("row", 0),
            "ref_no": r.get("ref_no", ""),
            "subject": r.get("subject", ""),
            "status": r.get("status", ""),
        }
        for r in log_ctx["rows"]
        if r.get("ref_no") and plog.normalize_ref_key(str(r["ref_no"])) not in seen_ref_keys
    ]
    all_refs = [it["ref_no"] for it in items if it["ref_no"]] + [str(r.get("ref_no") or "") for r in log_ctx["rows"] if r.get("ref_no")]
    seq_gaps = _find_sequence_gaps(all_refs)

    plan = {
        "version": 1,
        "project_id": proj.get("id"),
        "source_folder": str(src_root),
        "created_at": now_iso(),
        "extracted_archives": extracted_archives,
        "items": items,
        "log_rows_missing_file": log_rows_missing_file[:500],
        "sequence_gaps": seq_gaps,
        "executed": False,
        "last_execution": None,
    }
    plan_p = organizer_plan_path(proj_root)
    plan_p.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_json(plan_p, plan)
    return summarize_organizer_plan(plan, proj_root=proj_root)


def _validated_org_last_execution(plan: dict[str, Any] | None, proj_root: Path | str | None = None) -> dict[str, Any] | None:
    if not isinstance(plan, dict) or not plan.get("executed"):
        return None
    raw = plan.get("last_execution")
    if not isinstance(raw, dict) or raw.get("stopped"):
        return None
    if not raw.get("cleared"):
        return None
    out = dict(raw)
    if proj_root is not None:
        pr = resolve_local_or_unc(str(proj_root))
        rep_rel = str(out.get("report_rel") or "").strip()
        rep_exists = False
        if rep_rel:
            try:
                rp = safe_resolve_within(pr, rep_rel, allow_base=False, reject_symlinks=True)
                rep_exists = bool(rp.exists() and rp.is_file() and rp.stat().st_size > 0)
            except Exception:
                rep_exists = False
        out["report_exists"] = rep_exists
        if not rep_exists:
            out["report_rel"] = ""
    return out


def summarize_organizer_plan(plan: dict[str, Any] | None, proj_root: Path | str | None = None) -> dict[str, Any]:
    validated_exec = _validated_org_last_execution(plan, proj_root=proj_root)
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        return {
            "has_plan": False,
            "source_folder": str(plan.get("source_folder") or "") if isinstance(plan, dict) else "",
            "total": 0,
            "counts": {
                "keep": 0,
                "duplicate": 0, "duplicate_total": 0, "duplicate_resolved": 0,
                "near_duplicate": 0, "near_duplicate_total": 0, "near_duplicate_resolved": 0,
                "superseded": 0, "superseded_total": 0, "superseded_resolved": 0,
                "skip": 0, "excluded": 0, "unclassified": 0, "missing_from_log": 0, "included": 0,
                "active_total": 0,
            },
            "log_rows_missing_file": [],
            "sequence_gaps": [],
            "last_execution": validated_exec,
        }
    items = plan["items"]
    counts = {
        "keep": 0,
        "duplicate": 0,
        "duplicate_total": 0,
        "duplicate_resolved": 0,
        "near_duplicate": 0,
        "near_duplicate_total": 0,
        "near_duplicate_resolved": 0,
        "superseded": 0,
        "superseded_total": 0,
        "superseded_resolved": 0,
        "skip": 0,
        "excluded": 0,
        "unclassified": 0,
        "missing_from_log": 0,
        "included": 0,
        "active_total": 0,
    }
    for it in items:
        st = str(it.get("status") or "keep")
        dec = str(it.get("decision") or ("pending" if st in ("duplicate", "near_duplicate", "superseded") else ""))
        is_excluded = (not it.get("include")) or (dec == "excluded")
        is_resolved_special = st in ("duplicate", "near_duplicate", "superseded") and dec not in ("", "pending")
        if is_excluded:
            counts["excluded"] += 1
        if not is_excluded and not is_resolved_special:
            counts["active_total"] += 1
        if st == "duplicate":
            counts["duplicate_total"] += 1
            if dec in ("", "pending"):
                counts["duplicate"] += 1
            else:
                counts["duplicate_resolved"] += 1
        elif st == "near_duplicate":
            counts["near_duplicate_total"] += 1
            if dec in ("", "pending"):
                counts["near_duplicate"] += 1
            else:
                counts["near_duplicate_resolved"] += 1
        elif st == "superseded":
            counts["superseded_total"] += 1
            if dec in ("", "pending"):
                counts["superseded"] += 1
            else:
                counts["superseded_resolved"] += 1
        elif st in counts:
            if not is_excluded:
                counts[st] += 1
        if not it.get("confident") and st == "keep" and not is_excluded:
            counts["unclassified"] += 1
        if it.get("log_status") == "missing_from_log" and st not in ("duplicate", "near_duplicate") and not is_excluded:
            counts["missing_from_log"] += 1
        if it.get("include"):
            counts["included"] += 1
    return {
        "has_plan": len(items) > 0,
        "source_folder": plan.get("source_folder", ""),
        "created_at": plan.get("created_at", ""),
        "total": len(items),
        "counts": counts,
        "extracted_archives": plan.get("extracted_archives") or [],
        "log_rows_missing_file": plan.get("log_rows_missing_file") or [],
        "log_rows_missing_file_count": len(plan.get("log_rows_missing_file") or []),
        "sequence_gaps": plan.get("sequence_gaps") or [],
        "last_execution": validated_exec,
    }


def get_organizer_plan_page(
    proj_root: Path | str,
    *,
    page: int = 1,
    per_page: int = 50,
    status_filter: str = "all",
    section_filter: str = "",
    query: str = "",
    show_resolved: bool = False,
) -> dict[str, Any]:
    plan = read_json_safe(organizer_plan_path(proj_root), {})
    summary = summarize_organizer_plan(plan, proj_root=proj_root)
    if not summary["has_plan"]:
        return {**summary, "page": 1, "per_page": per_page, "total_filtered": 0, "total_pages": 1, "items": []}

    items = plan["items"]
    sf = str(status_filter or "all").strip().lower()
    sec_f = str(section_filter or "").strip()
    qn = normalize_ar(query) if query else ""

    filtered: list[dict[str, Any]] = []
    for it in items:
        st = str(it.get("status") or "keep")
        dec = str(it.get("decision") or ("pending" if st in ("duplicate", "near_duplicate", "superseded") else ""))
        is_excluded = (not it.get("include")) or (dec == "excluded")
        is_resolved_special = st in ("duplicate", "near_duplicate", "superseded") and dec not in ("", "pending")
        if sf == "all":
            if not show_resolved and (is_excluded or is_resolved_special):
                continue
        elif sf == "keep":
            if st != "keep":
                continue
            if not show_resolved and is_excluded:
                continue
        elif sf == "duplicate":
            if st != "duplicate":
                continue
            if not show_resolved and dec not in ("", "pending"):
                continue
        elif sf == "near_duplicate":
            if st != "near_duplicate":
                continue
            if not show_resolved and dec not in ("", "pending"):
                continue
        elif sf == "superseded":
            if st != "superseded":
                continue
            if not show_resolved and dec not in ("", "pending"):
                continue
        elif sf == "excluded":
            if not is_excluded:
                continue
        elif sf == "skip" and (st != "skip" and it.get("include")):
            continue
        elif sf == "unclassified":
            if it.get("confident") or st != "keep":
                continue
            if not show_resolved and is_excluded:
                continue
        elif sf == "missing_from_log":
            if it.get("log_status") != "missing_from_log" or st in ("duplicate", "near_duplicate"):
                continue
            if not show_resolved and is_excluded:
                continue

        if sec_f and it.get("section") != sec_f:
            continue
        if qn:
            hay = normalize_ar(f"{it.get('file_name','')} {it.get('ref_no','')} {it.get('subject','')} {it.get('src_rel','')} {it.get('dest_rel','')} {it.get('duplicate_of','')} {it.get('near_dup_of','')}")
            if qn not in hay:
                continue
        filtered.append(it)

    per_p = max(10, min(200, int(per_page)))
    total_f = len(filtered)
    total_pages = max(1, (total_f + per_p - 1) // per_p)
    cur_p = max(1, min(int(page), total_pages))
    start = (cur_p - 1) * per_p
    return {
        **summary,
        "page": cur_p,
        "per_page": per_p,
        "total_filtered": total_f,
        "total_pages": total_pages,
        "items": filtered[start:start + per_p],
    }


def update_organizer_plan_items(
    proj: dict[str, Any],
    updates: list[dict[str, Any]] | None = None,
    *,
    bulk_action: str = "",
) -> dict[str, Any]:
    proj_root = resolve_local_or_unc(str(proj["root"]))
    pp = organizer_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        raise FileNotFoundError(tr("server.not_found"))

    items = plan["items"]
    by_id = {str(it["id"]): it for it in items}

    if bulk_action == "exclude_duplicates":
        for it in items:
            if it.get("status") == "duplicate":
                it["include"] = False
                it["decision"] = "excluded"
    elif bulk_action == "isolate_duplicates":
        for it in items:
            if it.get("status") == "duplicate":
                it["include"] = True
                it["decision"] = "isolated"
                df, nn = _build_dest_rel(it, it["file_name"], it["ext"], proj, status="duplicate")
                it["dest_folder"] = df
                it["new_name"] = nn
                it["dest_rel"] = f"{df}/{nn}"
    elif bulk_action == "exclude_near_duplicates":
        for it in items:
            if it.get("status") == "near_duplicate":
                it["include"] = False
                it["decision"] = "excluded"
    elif bulk_action == "isolate_near_duplicates":
        for it in items:
            if it.get("status") == "near_duplicate":
                it["include"] = True
                it["decision"] = "isolated"
                df, nn = _build_dest_rel(it, it["file_name"], it["ext"], proj, status="near_duplicate")
                it["dest_folder"] = df
                it["new_name"] = nn
                it["dest_rel"] = f"{df}/{nn}"
    elif bulk_action == "exclude_superseded":
        for it in items:
            if it.get("status") == "superseded":
                it["include"] = False
                it["decision"] = "excluded"
    elif bulk_action == "include_superseded":
        for it in items:
            if it.get("status") == "superseded":
                it["include"] = True
                it["decision"] = "isolated"
                df, nn = _build_dest_rel(it, it["file_name"], it["ext"], proj, status="superseded")
                it["dest_folder"] = df
                it["new_name"] = nn
                it["dest_rel"] = f"{df}/{nn}"

    for upd in (updates or []):
        if not isinstance(upd, dict) or not upd.get("id"):
            continue
        it = by_id.get(str(upd["id"]))
        if not it:
            continue
        if "decision" in upd:
            dec_v = str(upd["decision"] or "").strip()
            if dec_v == "excluded":
                it["include"] = False
                it["decision"] = "excluded"
            elif dec_v == "isolated":
                it["include"] = True
                it["decision"] = "isolated"
            elif dec_v == "kept":
                it["status"] = "keep"
                it["include"] = True
                it["decision"] = "kept"
            elif dec_v == "pending":
                it["decision"] = "pending"
                it["include"] = (it.get("status") not in ("duplicate", "near_duplicate"))
        if "include" in upd:
            it["include"] = bool(upd["include"])
            if "decision" not in upd:
                if not it["include"]:
                    it["decision"] = "excluded"
                elif it.get("status") in ("duplicate", "near_duplicate", "superseded"):
                    it["decision"] = "isolated"
                else:
                    it["decision"] = "kept"
        if "status" in upd and str(upd["status"]) in ("keep", "duplicate", "near_duplicate", "superseded", "skip"):
            it["status"] = str(upd["status"])
            if it["status"] == "skip":
                it["include"] = False
                it["decision"] = "excluded"
            elif it["status"] == "keep" and "decision" not in upd:
                it["include"] = True
                it["decision"] = "kept"
        if "section" in upd and str(upd["section"]).strip():
            it["section"] = str(upd["section"]).strip()
            it["confident"] = True
        if "ref_no" in upd:
            it["ref_no"] = _canonical_ref(str(upd["ref_no"] or "")) if str(upd["ref_no"] or "").strip() else ""
            it["ref_key"] = plog.normalize_ref_key(it["ref_no"]) if it["ref_no"] else ""
        if "subject" in upd and str(upd["subject"]).strip():
            it["subject"] = str(upd["subject"]).strip()[:120]

        df, nn = _build_dest_rel(it, it["file_name"], it["ext"], proj, status=it["status"])
        it["dest_folder"] = df
        if "new_name" in upd:
            raw_nm = str(upd["new_name"] or "").strip()
            if raw_nm:
                nm = _apply_custom_filename(raw_nm, it["ext"])
                it["new_name"] = nm
                it["custom_name"] = True
            else:
                it["custom_name"] = False
                it["ai_renamed"] = False
                it["new_name"] = nn
        elif it.get("custom_name") and "ref_no" not in upd and it.get("status") not in ("duplicate", "near_duplicate"):
            # Preserve user/AI-assigned filename when toggling include/section/decision
            pass
        else:
            it["new_name"] = nn
        it["dest_rel"] = f"{it['dest_folder']}/{it['new_name']}"

    atomic_write_json(pp, plan)
    return summarize_organizer_plan(plan)


def _write_organizer_excel_report(proj: dict[str, Any], plan: dict[str, Any], exec_res: dict[str, Any]) -> Path:
    proj_root = resolve_local_or_unc(str(proj["root"]))
    rdir = reports_dir_for_project(proj_root)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_xlsx = rdir / f"Organizer_Report_{ts}.xlsx"

    wb = openpyxl.Workbook()
    hdr_fill = PatternFill("solid", fgColor="0F766E")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)

    def _style_hdr(ws):
        ws.views.sheetView[0].rightToLeft = True
        for c in ws[1]:
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Sheet 1: Summary
    ws1 = wb.active
    ws1.title = "1_الملخص_Summary"
    ws1.append(["البند", "القيمة"])
    _style_hdr(ws1)
    sum_info = summarize_organizer_plan(plan)
    cnt = sum_info["counts"]
    rows_s1 = [
        ("المشروع", f"[{proj.get('code','')}] {proj.get('name','')}"),
        ("المجلد المصدر", plan.get("source_folder", "")),
        ("تاريخ التنفيذ", now_iso()),
        ("إجمالي الملفات المفحوصة", sum_info["total"]),
        ("ملفات رئيسية مرتبة (Keep)", cnt["keep"]),
        ("ملفات مكررة متطابقة (Duplicates)", cnt.get("duplicate_total", cnt["duplicate"])),
        ("مكرر بالاسم والبيانات (Near-Duplicates)", cnt.get("near_duplicate_total", 0)),
        ("مراجعات سابقة معزولة (Superseded)", cnt.get("superseded_total", cnt["superseded"])),
        ("تم نسخها/نقلها فعلياً", exec_res.get("copied", 0)),
        ("تم تسجيلها/تحديثها في اللوج", exec_res.get("logged", 0)),
        ("صفوف لوج ينقصها ملف", len(plan.get("log_rows_missing_file") or [])),
        ("أخطاء", len(exec_res.get("errors") or [])),
    ]
    for k, v in rows_s1:
        ws1.append([k, v])
    ws1.column_dimensions["A"].width = 36
    ws1.column_dimensions["B"].width = 48

    # Sheet 2: All files mapping
    ws2 = wb.create_sheet("2_خطة_الترتيب_Files")
    ws2.append(["#", "اسم الملف الأصلي", "المسار النسبي القديم", "الإجراء", "القسم", "المرجع", "المراجعة", "الاسم الجديد", "المسار الجديد", "حالة اللوج", "التنفيذ"])
    _style_hdr(ws2)
    for idx, it in enumerate(plan.get("items") or [], 1):
        ws2.append([
            idx,
            it.get("file_name", ""),
            it.get("src_rel", ""),
            it.get("status", ""),
            it.get("section", ""),
            it.get("ref_no", ""),
            it.get("revision", ""),
            it.get("new_name", ""),
            it.get("dest_rel", ""),
            it.get("log_status", ""),
            "تم" if it.get("executed") else "مستبعد/معلق",
        ])
    for ci, w in enumerate((7, 30, 38, 14, 18, 16, 10, 32, 42, 18, 14), 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w

    # Sheet 3: Duplicates, Near-Duplicates & Superseded
    ws3 = wb.create_sheet("3_التكرار_والمراجعات")
    ws3.append(["النوع", "الملف", "المسار القديم", "السبب / النسخة المعتمدة", "القرار"])
    _style_hdr(ws3)
    for it in plan.get("items") or []:
        if it.get("status") in ("duplicate", "near_duplicate", "superseded"):
            reason = it.get("duplicate_of") or it.get("near_dup_of") or it.get("superseded_by") or ""
            if it.get("near_dup_reason"):
                reason = f"{it['near_dup_reason']}: {reason}"
            ws3.append([it.get("status"), it.get("file_name"), it.get("src_rel"), reason, it.get("decision") or ""])
    for ci, w in enumerate((16, 34, 44, 38, 14), 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w

    # Sheet 4: Log Audit & Missing Sequence Gaps
    ws4 = wb.create_sheet("4_تدقيق_اللوج_والنواقص")
    ws4.append(["نوع الملاحظة", "التفاصيل / السلسلة", "المرجع / الأرقام المفقودة", "الشيت / الملاحظة"])
    _style_hdr(ws4)
    for gap in (plan.get("sequence_gaps") or []):
        ws4.append([
            "أرقام مفقودة في التسلسل",
            f"{gap['series']} (من {gap['min']} إلى {gap['max']})",
            ", ".join(str(x) for x in gap.get("missing_numbers", [])),
            f"نقص {gap.get('missing_count', 0)}",
        ])
    for mr in (plan.get("log_rows_missing_file") or []):
        ws4.append([
            "صف في اللوج بلا ملف",
            mr.get("ref_no", ""),
            mr.get("subject", ""),
            f"{mr.get('log_name','')} / {mr.get('sheet','')} (صف {mr.get('row','')})",
        ])
    for it in plan.get("items") or []:
        if it.get("log_status") == "missing_from_log" and it.get("status") not in ("duplicate", "near_duplicate"):
            ws4.append([
                "ملف غير مسجل باللوج",
                it.get("ref_no", ""),
                it.get("file_name", ""),
                it.get("dest_rel", ""),
            ])
    for ci, w in enumerate((24, 30, 42, 34), 1):
        ws4.column_dimensions[get_column_letter(ci)].width = w

    for cell in ws1["A"]:
        if cell.row > 1:
            cell.font = bold_font

    wb.save(out_xlsx)
    return out_xlsx


def execute_organizer_plan(
    proj: dict[str, Any],
    *,
    mode: str = "copy",
    update_excel_log: bool = True,
    device_id: str = "",
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """
    Execute the saved `organizer_plan.json` over all included files (`copy` or `move`),
    optionally register missing items / update links inside the project's Excel log,
    and write a comprehensive Excel report in `_Reports/`.
    """
    import inbox_pipeline as ip

    proj_root = resolve_local_or_unc(str(proj["root"]))
    pp = organizer_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        raise FileNotFoundError(tr("server.not_found"))

    arc_root = ip.archive_dir_for_project(proj)
    arc_root.mkdir(parents=True, exist_ok=True)
    logs_dir = plog.logs_dir_for_project(proj_root)
    has_logs = bool(plog.list_project_logs(proj_root, device_id=device_id))

    included = [it for it in plan["items"] if it.get("include") and str(it.get("status")) != "skip"]
    total = len(included)
    if total == 0:
        raise ValueError(tr("organizer.no_items_to_exec_err"))
    move_mode = str(mode or "copy").strip().lower() == "move"

    copied = 0
    skipped = 0
    logged = 0
    errors: list[str] = []

    last_exec_emit = time.monotonic()
    for idx, it in enumerate(included, 1):
        if stop_event is not None and stop_event.is_set():
            break
        src = resolve_local_or_unc(str(it.get("src_path") or ""))
        if progress_cb and (idx == 1 or idx % 10 == 0 or idx == total or (time.monotonic() - last_exec_emit) >= 0.25):
            last_exec_emit = time.monotonic()
            progress_cb({
                "phase": "executing",
                "processed": idx,
                "total": total,
                "copied": copied,
                "logged": logged,
                "current_file": src.name,
            })

        if not src.exists() or not src.is_file():
            skipped += 1
            continue

        dest_folder_parts = [p for p in str(it.get("dest_folder") or "09_General_عام").split("/") if p.strip()]
        dest_dir = arc_root.joinpath(*dest_folder_parts)
        dest_dir.mkdir(parents=True, exist_ok=True)
        ext0 = src.suffix.lower()
        nm_full = _apply_custom_filename(str(it.get("new_name") or src.name), ext0)
        dest_path = dest_dir / nm_full
        if dest_path.exists():
            # Avoid overwriting distinct content; append short suffix
            k = 2
            stem0 = nm_full[: -len(ext0)] if ext0 and nm_full.lower().endswith(ext0) else nm_full
            while dest_path.exists():
                dest_path = dest_dir / f"{stem0}_{k}{ext0}"
                k += 1

        try:
            src_sha = sha256_of_file(src) if move_mode else ""
            shutil.copy2(src, dest_path)
            if move_mode:
                if not dest_path.exists() or sha256_of_file(dest_path) != src_sha:
                    raise OSError(f"Integrity verification failed after copy: {dest_path.name}")
                if resolve_local_or_unc(src) != resolve_local_or_unc(dest_path):
                    try:
                        src.unlink()
                    except OSError:
                        pass
            it["executed"] = True
            it["final_dest_rel"] = dest_path.relative_to(arc_root).as_posix()
            copied += 1
        except Exception as exc:
            errors.append(f"{src.name}: {exc}")
            continue

        # Record/update in project's uploaded OR directly linked Excel log
        if update_excel_log and has_logs and it.get("status") == "keep" and it.get("ref_no"):
            try:
                try:
                    rel_for_log = Path(os.path.relpath(dest_path, logs_dir)).as_posix()
                except ValueError:
                    rel_for_log = str(dest_path)
                lr = plog.record_entry_in_project_log(
                    proj,
                    {
                        "section": it.get("section") or "general",
                        "ref_no": it.get("ref_no"),
                        "revision": it.get("revision") or "R00",
                        "discipline": it.get("discipline") or "",
                        "subject": it.get("subject") or dest_path.stem,
                        "approval_code": it.get("approval_code") or "",
                        "authority": it.get("authority") or "",
                        "file_rel_path": rel_for_log,
                        "file_abs_path": str(dest_path),
                    },
                    device_id=device_id,
                    queue_if_locked=True,
                )
                if lr.get("ok"):
                    logged += 1
            except Exception as exc_l:
                errors.append(f"Log({it.get('ref_no')}): {exc_l}")

    was_stopped = bool(stop_event is not None and stop_event.is_set())
    if was_stopped:
        # Stopped before completion: never generate final report or completion banner
        plan["executed"] = False
        plan["last_execution"] = None
        atomic_write_json(pp, plan)
        return {
            "ok": False,
            "stopped": True,
            "mode": "move" if move_mode else "copy",
            "total_included": total,
            "copied": copied,
            "skipped": skipped,
            "logged": logged,
            "errors": errors[:50],
            "cleared": False,
        }

    if copied == 0:
        raise ValueError(tr("organizer.no_items_to_exec_err"))

    exec_summary = {
        "mode": "move" if move_mode else "copy",
        "total_included": total,
        "excluded": max(0, len(plan["items"]) - len(included)),
        "copied": copied,
        "skipped": skipped,
        "logged": logged,
        "errors": errors[:50],
        "finished_at": now_iso(),
        "executed": True,
    }
    report_path = _write_organizer_excel_report(proj, plan, exec_summary)
    try:
        if not report_path.exists() or not report_path.is_file() or report_path.stat().st_size == 0:
            raise RuntimeError("Failed to generate organizer report")
        exec_summary["report_rel"] = report_path.relative_to(proj_root).as_posix()
    except ValueError:
        exec_summary["report_rel"] = report_path.name
    except Exception:
        report_path.unlink(missing_ok=True)
        raise
    exec_summary["report_name"] = report_path.name
    exec_summary["report_exists"] = True

    # Clear pending organizer items once execution completes cleanly so the UI resets to a clean state
    exec_summary["cleared"] = True
    plan["items"] = []
    plan["log_rows_missing_file"] = []
    plan["sequence_gaps"] = []
    plan["executed"] = True
    plan["executed_at"] = now_iso()
    shutil.rmtree(proj_root / "_System_Agent" / "_org_extracted", ignore_errors=True)

    plan["last_execution"] = exec_summary
    atomic_write_json(pp, plan)
    return exec_summary


def audit_archive_and_generate_weekly_report(
    proj: dict[str, Any],
    *,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Standalone Project Audit + Management Weekly Follow-up Report:
      1) Overdue submittals (> overdue_days with no reply)
      2) Rejected items needing resubmission (`Code C / D`)
      3) Log rows whose files are missing from `archive_dir`
      4) Files sitting in `archive_dir` that are not registered in the project Excel log
      5) Missing numeric sequence gaps (`MAR-EL`, `OUT`, ...)
    """
    import inbox_pipeline as ip

    proj_root = resolve_local_or_unc(str(proj["root"]))
    arc_root = ip.archive_dir_for_project(proj)
    arc_files, _ = _walk_folder_files(arc_root, include_archives=False) if arc_root.exists() else ([], [])

    log_ctx = _collect_log_refs_and_rows(proj, device_id=device_id)
    rows = log_ctx["rows"]
    stats = log_ctx["stats"]
    overdue_days = int(stats.get("overdue_days") or proj.get("overdue_days") or 14)
    today = dt.date.today()

    # Fast index of archive files by extracted ref_key
    arc_by_ref: dict[str, list[dict[str, str]]] = defaultdict(list)
    arc_unlogged: list[dict[str, Any]] = []
    arc_refs_all: list[str] = []
    for fp in arc_files:
        meta = _extract_metadata_fast(fp, arc_root, proj, read_text=False)
        rk = meta["ref_key"]
        try:
            rel_p = fp.relative_to(arc_root).as_posix()
        except ValueError:
            rel_p = fp.name
        if rk:
            arc_refs_all.append(meta["ref_no"])
            arc_by_ref[rk].append({"file_name": fp.name, "rel": rel_p})
            if rk not in log_ctx["ref_map"]:
                arc_unlogged.append({
                    "file_name": fp.name,
                    "rel": rel_p,
                    "section": meta["section"],
                    "ref_no": meta["ref_no"],
                    "revision": meta["revision"],
                    "discipline": meta["discipline"],
                    "subject": meta["subject"],
                    "path": str(fp),
                })

    log_missing_files: list[dict[str, Any]] = []
    overdue_items: list[dict[str, Any]] = []
    rejected_items: list[dict[str, Any]] = []

    for r in rows:
        ref_s = str(r.get("ref_no") or "").strip()
        rk = plog.normalize_ref_key(ref_s)
        if rk and rk not in arc_by_ref:
            log_missing_files.append({
                "ref_no": ref_s,
                "subject": r.get("subject", ""),
                "sheet": r.get("sheet", ""),
                "row": r.get("row", ""),
            })
        sl = str(r.get("status") or "").lower()
        ds = str(r.get("date") or "")[:10]
        reply_s = str(r.get("reply_date") or "").strip()
        if any(k in sl for k in ("reject", "مرفوض", "revise")) or sl in ("c", "d", "code c", "code d"):
            rejected_items.append(r)
        elif not reply_s and not any(k in sl for k in ("approv", "معتمد", "withdraw", "cancel", "ملغي")) and ds:
            try:
                d_obj = dt.date.fromisoformat(ds)
                age = (today - d_obj).days
                if age > overdue_days:
                    overdue_items.append({**r, "days_open": age})
            except ValueError:
                pass

    all_refs = arc_refs_all + [str(r.get("ref_no") or "") for r in rows if r.get("ref_no")]
    seq_gaps = _find_sequence_gaps(all_refs)

    # Export Excel Audit & Follow-up Report
    rdir = reports_dir_for_project(proj_root)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_xlsx = rdir / f"Audit_Followup_{ts}.xlsx"
    wb = openpyxl.Workbook()
    hdr_fill = PatternFill("solid", fgColor="1E40AF")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)

    def _hdr(ws):
        ws.views.sheetView[0].rightToLeft = True
        for c in ws[1]:
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center")

    ws1 = wb.active
    ws1.title = "1_ملخص_المتابعة"
    ws1.append(["مؤشر المتابعة", "العدد"])
    _hdr(ws1)
    for lbl, val in (
        ("إجمالي سجلات اللوج", stats.get("total_rows", 0)),
        ("الاعتمادات المعتمدة", stats.get("approved", 0)),
        ("المرفوضات / تحتاج إعادة تقديم", len(rejected_items)),
        (f"المتأخرات أكثر من {overdue_days} يوم", len(overdue_items)),
        ("صفوف لوج بدون ملف بالأرشيف", len(log_missing_files)),
        ("ملفات أرشيف غير مسجلة باللوج", len(arc_unlogged)),
        ("سلاسل بها أرقام مفقودة", len(seq_gaps)),
    ):
        ws1.append([lbl, val])
    ws1.column_dimensions["A"].width = 38
    ws1.column_dimensions["B"].width = 18

    ws2 = wb.create_sheet("2_متأخرات_ومرفوضات")
    ws2.append(["التصنيف", "المرجع", "التخصص", "الموضوع", "التاريخ", "الحالة", "أيام الانتظار"])
    _hdr(ws2)
    for it in overdue_items:
        ws2.append(["متأخر عن الرد", it.get("ref_no"), it.get("discipline"), it.get("subject"), it.get("date"), it.get("status"), it.get("days_open")])
    for it in rejected_items:
        ws2.append(["مرفوض / إعادة تقديم", it.get("ref_no"), it.get("discipline"), it.get("subject"), it.get("date"), it.get("status"), ""])
    for ci, w in enumerate((22, 16, 12, 42, 14, 18, 14), 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w

    ws3 = wb.create_sheet("3_فجوات_اللوج_والأرشيف")
    ws3.append(["النوع", "المرجع", "التفاصيل / اسم الملف", "الموقع / الشيت"])
    _hdr(ws3)
    for x in log_missing_files:
        ws3.append(["صف باللوج بلا ملف", x["ref_no"], x["subject"], f"{x['sheet']} ({x['row']})"])
    for x in arc_unlogged:
        ws3.append(["ملف أرشيف غير مسجل باللوج", x["ref_no"], x["file_name"], x["rel"]])
    for g in seq_gaps:
        ws3.append(["أرقام ساقطة بالتسلسل", g["series"], ", ".join(str(n) for n in g["missing_numbers"]), f"نقص {g['missing_count']}"])
    for ci, w in enumerate((26, 18, 42, 36), 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w

    try:
        wb.save(out_xlsx)
        if not out_xlsx.exists() or not out_xlsx.is_file() or out_xlsx.stat().st_size == 0:
            raise RuntimeError("Failed to generate weekly audit report")
    except Exception:
        out_xlsx.unlink(missing_ok=True)
        raise
    try:
        rep_rel = out_xlsx.relative_to(proj_root).as_posix()
    except ValueError:
        rep_rel = out_xlsx.name

    return {
        "ok": True,
        "archive_files_count": len(arc_files),
        "overdue_count": len(overdue_items),
        "rejected_count": len(rejected_items),
        "log_missing_files_count": len(log_missing_files),
        "archive_unlogged_count": len(arc_unlogged),
        "sequence_gaps_count": len(seq_gaps),
        "overdue_items": overdue_items[:80],
        "rejected_items": rejected_items[:80],
        "log_missing_files": log_missing_files[:100],
        "archive_unlogged": arc_unlogged[:200],
        "sequence_gaps": seq_gaps,
        "report_name": out_xlsx.name,
        "report_rel": rep_rel,
        "report_exists": True,
    }


def generate_weekly_followup_report(
    proj: dict[str, Any],
    *,
    device_id: str = "",
    export_pdf: bool = True,
    days_window: int = 7,
) -> dict[str, Any]:
    """
    Generate a 1-click Management Weekly Follow-up Report (`Excel` + `PDF`) covering:
      1) Overdue submittals waiting on Consultant / Ministry beyond `overdue_days` (with delay days)
      2) Unreplied incoming letters (with waiting days)
      3) Rejected / Revise-and-Resubmit items (Code C / D)
      4) Documents & archive files added/modified within the last `days_window` (default 7) days
    """
    import inbox_pipeline as ip
    import template_render as trender

    proj_root = resolve_local_or_unc(str(proj["root"]))
    arc_root = ip.archive_dir_for_project(proj)
    arc_files, _ = _walk_folder_files(arc_root, include_archives=False) if arc_root.exists() else ([], [])

    log_ctx = _collect_log_refs_and_rows(proj, device_id=device_id)
    rows = log_ctx["rows"]
    stats = log_ctx["stats"]
    overdue_days = int(stats.get("overdue_days") or proj.get("overdue_days") or 14)
    today = dt.date.today()
    cutoff_date = today - dt.timedelta(days=max(1, int(days_window or 7)))
    cutoff_ts = dt.datetime.combine(cutoff_date, dt.time.min).timestamp()

    overdue_submittals: list[dict[str, Any]] = []
    unreplied_incoming: list[dict[str, Any]] = []
    rejected_items: list[dict[str, Any]] = []
    recent_7d_items: list[dict[str, Any]] = []
    seen_recent_keys: set[str] = set()

    for r in rows:
        sec = str(r.get("section") or "")
        ref_s = str(r.get("ref_no") or "").strip()
        ds = str(r.get("date") or "")[:10]
        sl = str(r.get("status") or "").lower()
        ml = str(r.get("moh_status") or "").lower()
        reply_s = str(r.get("reply_date") or "").strip()

        d_obj: dt.date | None = None
        if ds:
            try:
                d_obj = dt.date.fromisoformat(ds)
            except ValueError:
                d_obj = None

        # 1) Unreplied incoming letters
        if sec == "incoming" and r.get("unreplied"):
            wait_d = r.get("days_waiting")
            if wait_d is None and d_obj is not None:
                wait_d = max(0, (today - d_obj).days)
            unreplied_incoming.append({
                "ref_no": ref_s,
                "sheet": r.get("sheet", ""),
                "party": r.get("party", ""),
                "party_label": r.get("party_label") or r.get("party") or "",
                "subject": r.get("subject", ""),
                "date": ds,
                "days_waiting": int(wait_d or 0),
                "file_rel": r.get("file_rel", ""),
                "file_area": r.get("file_area", "root"),
                "is_previewable": bool(r.get("is_previewable")),
            })

        # 2) Rejected / Resubmission needed
        if any(k in sl for k in ("reject", "مرفوض", "revise")) or sl in ("c", "d", "code c", "code d"):
            rejected_items.append({
                "ref_no": ref_s,
                "discipline": r.get("discipline", ""),
                "subject": r.get("subject", ""),
                "date": ds,
                "status": r.get("status", ""),
                "sheet": r.get("sheet", ""),
            })
        # 3) Overdue Consultant or Ministry Submittals
        elif sec == "submittal" or sec.startswith("submittal_"):
            is_open_cons = (not reply_s) and not any(k in sl for k in ("approv", "معتمد", "withdraw", "cancel", "ملغي", "مسحوب"))
            is_open_moh = bool(ml) and not any(k in ml for k in ("approv", "معتمد", "reject", "مرفوض", "withdraw", "cancel", "ملغي"))
            if (is_open_cons or is_open_moh) and d_obj is not None:
                age = max(0, (today - d_obj).days)
                if age > overdue_days:
                    overdue_submittals.append({
                        "ref_no": ref_s,
                        "discipline": r.get("discipline", ""),
                        "subject": r.get("subject", ""),
                        "date": ds,
                        "status": r.get("status", ""),
                        "moh_status": r.get("moh_status", ""),
                        "days_open": age,
                        "delay_days": age - overdue_days,
                        "file_rel": r.get("file_rel", ""),
                        "file_area": r.get("file_area", "root"),
                        "is_previewable": bool(r.get("is_previewable")),
                    })

        # 4) Log entries added in the last 7 days
        if d_obj is not None and d_obj >= cutoff_date:
            rk = plog.normalize_ref_key(ref_s)
            key_u = rk or f"{r.get('sheet','')}|{r.get('row','')}"
            seen_recent_keys.add(key_u)
            recent_7d_items.append({
                "kind": "log",
                "kind_label": r.get("sheet") or sec or "سجل",
                "ref_no": ref_s,
                "title": r.get("subject") or ref_s,
                "date": ds,
                "status": r.get("status", ""),
                "file_rel": r.get("file_rel", ""),
                "file_area": r.get("file_area", "root"),
                "is_previewable": bool(r.get("is_previewable")),
            })

    # Include newly modified/archived files from the last 7 days not already represented by log rows
    for fp in arc_files:
        try:
            st = fp.stat()
            if st.st_mtime < cutoff_ts:
                continue
            m_date = dt.date.fromtimestamp(st.st_mtime).isoformat()
        except OSError:
            continue
        meta = _extract_metadata_fast(fp, arc_root, proj, read_text=False)
        rk = meta.get("ref_key") or ""
        if rk and rk in seen_recent_keys:
            continue
        try:
            rel_r = fp.relative_to(proj_root).as_posix()
            f_area = "root"
        except ValueError:
            try:
                rel_r = fp.relative_to(arc_root).as_posix()
                f_area = "archive"
            except ValueError:
                rel_r = fp.name
                f_area = "archive"
        recent_7d_items.append({
            "kind": "archive",
            "kind_label": "ملف أرشيف",
            "ref_no": meta.get("ref_no") or "",
            "title": fp.name,
            "date": m_date,
            "status": meta.get("section") or "",
            "file_rel": rel_r,
            "file_area": f_area,
            "is_previewable": fp.suffix.lower() in {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".docx", ".xlsx", ".xlsm"},
        })

    overdue_submittals.sort(key=lambda x: int(x.get("delay_days") or 0), reverse=True)
    unreplied_incoming.sort(key=lambda x: int(x.get("days_waiting") or 0), reverse=True)
    recent_7d_items.sort(key=lambda x: str(x.get("date") or ""), reverse=True)

    summary = {
        "total_rows": int(stats.get("total_rows") or 0),
        "approved": int(stats.get("approved") or 0),
        "pending": int(stats.get("pending") or 0),
        "overdue_days_limit": overdue_days,
        "overdue_count": len(overdue_submittals),
        "unreplied_count": len(unreplied_incoming),
        "rejected_count": len(rejected_items),
        "recent_7d_count": len(recent_7d_items),
    }

    rdir = reports_dir_for_project(proj_root)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_xlsx = rdir / f"Weekly_Followup_{ts}.xlsx"
    out_pdf = rdir / f"Weekly_Followup_{ts}.pdf"

    wb = openpyxl.Workbook()
    hdr_fill = PatternFill("solid", fgColor="035D89")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)

    def _style_ws(ws):
        ws.views.sheetView[0].rightToLeft = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        for c in ws[1]:
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center")

    ws1 = wb.active
    ws1.title = "1_ملخص_المتابعة"
    ws1.append(["مؤشر المتابعة الأسبوعي", "القيمة"])
    _style_ws(ws1)
    for lbl, val in (
        ("المشروع", f"[{proj.get('code') or ''}] {proj.get('name') or ''}"),
        ("تاريخ التقرير", today.isoformat()),
        ("إجمالي السجلات", summary["total_rows"]),
        ("الاعتمادات المعتمدة", summary["approved"]),
        ("قيد المراجعة", summary["pending"]),
        (f"متأخرات الاستشاري والوزارة (> {overdue_days} يوم)", summary["overdue_count"]),
        ("كتب واردة بلا رد", summary["unreplied_count"]),
        ("مرفوضات تحتاج إعادة تقديم", summary["rejected_count"]),
        (f"مستندات مضافة آخر {days_window} أيام", summary["recent_7d_count"]),
    ):
        ws1.append([lbl, val])
    ws1.column_dimensions["A"].width = 44
    ws1.column_dimensions["B"].width = 28

    ws2 = wb.create_sheet("2_متأخرات_الاستشاري")
    ws2.append(["المرجع", "التخصص", "الموضوع", "تاريخ التقديم", "الحالة", "عمر الطلب (يوم)", "أيام التأخير"])
    _style_ws(ws2)
    for it in overdue_submittals:
        ws2.append([it["ref_no"], it["discipline"], it["subject"], it["date"], it["status"], it["days_open"], it["delay_days"]])
    for ci, w in enumerate((18, 12, 42, 14, 18, 16, 14), 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w

    ws3 = wb.create_sheet("3_وارد_بلا_رد")
    ws3.append(["المرجع", "الجهة", "الموضوع", "تاريخ الورود", "أيام الانتظار"])
    _style_ws(ws3)
    for it in unreplied_incoming:
        ws3.append([it["ref_no"], it["party_label"], it["subject"], it["date"], it["days_waiting"]])
    for ci, w in enumerate((18, 22, 46, 14, 16), 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w

    ws4 = wb.create_sheet("4_مستندات_آخر_7_أيام")
    ws4.append(["التصنيف", "المرجع", "العنوان / اسم الملف", "التاريخ", "الحالة"])
    _style_ws(ws4)
    for it in recent_7d_items:
        ws4.append([it["kind_label"], it["ref_no"], it["title"], it["date"], it["status"]])
    for ci, w in enumerate((22, 18, 44, 14, 18), 1):
        ws4.column_dimensions[get_column_letter(ci)].width = w

    try:
        wb.save(out_xlsx)
        if not out_xlsx.exists() or not out_xlsx.is_file() or out_xlsx.stat().st_size == 0:
            raise RuntimeError("Failed to generate weekly report workbook")
    except Exception:
        out_xlsx.unlink(missing_ok=True)
        raise

    try:
        xlsx_rel = out_xlsx.relative_to(proj_root).as_posix()
    except ValueError:
        xlsx_rel = out_xlsx.name

    res_payload: dict[str, Any] = {
        "ok": True,
        "project_id": proj.get("id", ""),
        "project_name": proj.get("name", ""),
        "project_code": proj.get("code", ""),
        "generated_date": today.isoformat(),
        "summary": summary,
        "overdue_submittals": overdue_submittals[:100],
        "unreplied_incoming": unreplied_incoming[:100],
        "recent_7d_items": recent_7d_items[:100],
        "rejected_items": rejected_items[:60],
        "xlsx_name": out_xlsx.name,
        "xlsx_rel": xlsx_rel,
        "report_name": out_xlsx.name,
        "report_rel": xlsx_rel,
        "pdf_name": "",
        "pdf_rel": "",
        "pdf_pages": 0,
    }

    if export_pdf:
        pdf_res = trender.build_weekly_report_pdf(res_payload, out_xlsx, out_pdf)
        if out_pdf.exists() and out_pdf.stat().st_size > 0:
            try:
                pdf_rel = out_pdf.relative_to(proj_root).as_posix()
            except ValueError:
                pdf_rel = out_pdf.name
            res_payload["pdf_name"] = out_pdf.name
            res_payload["pdf_rel"] = pdf_rel
            res_payload["pdf_pages"] = int(pdf_res.get("pages") or 1)
            res_payload["pdf_engine"] = pdf_res.get("engine") or ""

    return res_payload


def register_unlogged_archive_files(proj: dict[str, Any], *, device_id: str = "") -> dict[str, Any]:
    """Batch-register archive files that carry a reference code and are missing from the project Excel log."""
    audit = audit_archive_and_generate_weekly_report(proj, device_id=device_id)
    proj_root = resolve_local_or_unc(str(proj["root"]))
    logs_dir = plog.logs_dir_for_project(proj_root)
    added = 0
    queued = 0
    errors: list[str] = []

    for item in audit.get("archive_unlogged") or []:
        fp = Path(str(item["path"]))
        try:
            rel_link = Path(os.path.relpath(fp, logs_dir)).as_posix()
        except ValueError:
            rel_link = str(fp)
        try:
            res = plog.record_entry_in_project_log(
                proj,
                {
                    "section": item.get("section") or "general",
                    "ref_no": item["ref_no"],
                    "revision": item.get("revision") or "R00",
                    "discipline": item.get("discipline") or "",
                    "subject": item.get("subject") or fp.stem,
                    "file_rel_path": rel_link,
                },
                device_id=device_id,
                queue_if_locked=True,
            )
            if res.get("queued"):
                queued += 1
            elif res.get("ok"):
                added += 1
        except Exception as exc:
            errors.append(f"{item['ref_no']}: {exc}")
    return {"ok": True, "added": added, "queued": queued, "errors": errors[:20]}


def _render_organizer_ai_inputs(fp: Path, tmp_dir: Path) -> tuple[list[Path], str]:
    """Prepare up to 2 preview PNGs (for images/PDFs) plus a fast text excerpt for AI renaming."""
    ext = fp.suffix.lower()
    images: list[Path] = []
    snippet = _read_fast_text_snippet(fp, 2400)
    if ext in sw.IMAGE_EXTS:
        images = [fp]
    elif ext == ".pdf" and fitz is not None:
        try:
            doc = fitz.open(str(fp))
            for pg_i in range(min(len(doc), 2)):
                pix = doc[pg_i].get_pixmap(dpi=130)
                out_png = tmp_dir / f"{_clean_seg(fp.stem, 25)}_p{pg_i + 1}.png"
                pix.save(str(out_png))
                images.append(out_png)
            doc.close()
        except Exception:
            pass
    return images, snippet


def _apply_organizer_ai_meta(
    it: dict[str, Any],
    meta_ai: dict[str, Any],
    proj: dict[str, Any],
    sec_map: dict[str, Any],
    log_ref_map: dict[str, Any],
) -> bool:
    if not isinstance(meta_ai, dict) or not meta_ai:
        return False
    sec_new = str(meta_ai.get("section_id") or "").strip()
    if sec_new in sec_map:
        it["section"] = sec_new
        it["confident"] = True

    ref_raw = str(meta_ai.get("ref_no") or "").strip()
    if ref_raw:
        it["ref_no"] = _canonical_ref(ref_raw)
        it["ref_key"] = plog.normalize_ref_key(it["ref_no"])
        it["log_status"] = "matched" if it["ref_key"] in log_ref_map else "missing_from_log"

    rev_raw = str(meta_ai.get("revision") or "").strip()
    if rev_raw:
        m_rev = re.search(r"(\d{1,2})", rev_raw)
        if m_rev:
            it["rev_num"] = int(m_rev.group(1))
            it["revision"] = f"R{it['rev_num']:02d}"

    dd_raw = str(meta_ai.get("doc_date") or "").strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", dd_raw):
        it["doc_date"] = dd_raw

    auth_raw = str(meta_ai.get("authority") or "").strip()
    if auth_raw:
        auth_ar = normalize_ar(auth_raw)
        if any(k in auth_ar for k in ("وزار", "اشغال", "moh", "mpw", "mow")):
            it["authority"] = "MOH"

    subj_raw = _clean_subject_phrase(str(meta_ai.get("subject_ar") or meta_ai.get("subject") or "").strip())
    if subj_raw and len(subj_raw) >= 3:
        it["subject"] = subj_raw[:120]

    df, auto_nn = _build_dest_rel(it, it["file_name"], it["ext"], proj, status=it.get("status") or "keep")
    it["dest_folder"] = df
    sugg_fn = str(meta_ai.get("suggested_filename") or "").strip()
    if sugg_fn and not _is_generic_stem(sugg_fn):
        sugg_clean = _apply_custom_filename(sugg_fn, it["ext"])
        if not it.get("ext") or len(sugg_clean) > len(it["ext"]) + 3:
            it["new_name"] = sugg_clean
        else:
            it["new_name"] = auto_nn
    else:
        it["new_name"] = auto_nn
    it["dest_rel"] = f"{df}/{it['new_name']}"
    it["custom_name"] = True
    it["ai_renamed"] = True
    it["needs_ai"] = False
    return True


def ai_rename_organizer_items(
    proj: dict[str, Any],
    *,
    item_ids: list[str] | None = None,
    scope: str = "unclassified",
    limit: int = 25,
    engine: str = "gemini",
    ollama_url: str = "",
    ollama_model: str = "",
    gemini_api_key: str = "",
    device_id: str = "",
    force_refresh: bool = False,
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
    analyzer_fn: Callable[[str, Path, list[Path], str], dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """
    Use AI (Gemini / Ollama) to inspect document images/PDFs/text in an organizer plan,
    infer true project section, reference code, revision, date, and descriptive filename.
    Results are cached per project by exact file SHA-256 so re-scans never burn duplicate quota.
    """
    proj_root = resolve_local_or_unc(str(proj["root"]))
    pp = organizer_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        raise FileNotFoundError(tr("server.not_found"))

    items = plan["items"]
    sec_map = tprof.sections_map_for_project(proj)
    sec_list_txt = "\n".join(
        f"  - {sid}: {info.get('label', sid)}"
        for sid, info in sec_map.items()
    )

    if item_ids:
        wanted = {str(x) for x in item_ids if str(x).strip()}
        targets = [it for it in items if str(it.get("id")) in wanted]
    elif scope == "all_keep":
        targets = [
            it for it in items
            if it.get("status") == "keep" and (force_refresh or not it.get("ai_renamed"))
        ][: max(1, int(limit))]
    else:
        targets = [
            it for it in items
            if it.get("status") == "keep"
            and (force_refresh or not it.get("ai_renamed"))
            and (not it.get("confident") or not it.get("ref_no") or it.get("needs_ai"))
        ][: max(1, int(limit))]

    total_targets = len(targets)
    if total_targets == 0:
        return {
            "ok": True,
            "processed": 0,
            "updated": 0,
            "total_targets": 0,
            "errors": [],
            "summary": summarize_organizer_plan(plan),
        }

    log_ctx = _collect_log_refs_and_rows(proj, device_id=device_id)
    log_ref_map = log_ctx["ref_map"]

    cache_dir = ai_cache.project_ai_cache_dir(proj_root)
    sec_sig = ai_cache.stable_sig(sorted(sec_map.keys()), proj.get("domain"))

    tmp_dir = proj_root / "_System_Agent" / "_org_ai_tmp"
    shutil.rmtree(tmp_dir, ignore_errors=True)
    tmp_dir.mkdir(parents=True, exist_ok=True)

    processed = 0
    updated = 0
    cache_hits = 0
    errors: list[str] = []

    try:
        for idx, it in enumerate(targets, 1):
            if stop_event is not None and stop_event.is_set():
                break
            fp = resolve_local_or_unc(str(it.get("src_path") or ""))
            if progress_cb:
                progress_cb({
                    "phase": "ai_renaming",
                    "processed": idx,
                    "total": total_targets,
                    "updated": updated,
                    "current_file": it.get("file_name") or fp.name,
                })
            if not fp.exists() or not fp.is_file():
                errors.append(f"{it.get('file_name','')}: missing")
                continue

            fhash = str(it.get("sha256") or "")
            if not fhash:
                try:
                    fhash = sha256_of_file(fp)
                    it["sha256"] = fhash
                except OSError:
                    fhash = ""
            cache_key = f"{fhash}:{sec_sig}" if fhash else ""
            if cache_key and not force_refresh:
                c_rec = ai_cache.get_cached(cache_dir, "organizer_ai", cache_key)
                if c_rec and isinstance(c_rec.get("meta_ai"), dict):
                    processed += 1
                    if _apply_organizer_ai_meta(it, c_rec["meta_ai"], proj, sec_map, log_ref_map):
                        it["ai_from_cache"] = True
                        updated += 1
                        cache_hits += 1
                    continue

            images, snippet = _render_organizer_ai_inputs(fp, tmp_dir)
            prompt = (
                "أنت خبير أرشفة مستندات مشاريع هندسية وإنشائية. اقرأ هذا المستند بعناية واستخرج بياناته الحقيقية "
                "واقترح اسم ملف رسمي نظيف ومميز يعكس محتواه بدقة (تجاهل أي اسم عشوائي أو قديم للملف).\n"
                f"اسم الملف الحالي: {it.get('file_name','')}\n"
                f"الأقسام المتاحة في المشروع (اختر أفضل section_id من القائمة):\n{sec_list_txt}\n"
                + (f"\nمقتطف نصي من المستند:\n{snippet[:1800]}\n" if snippet else "")
                + "\nأرجع JSON فقط بهذه المفاتيح:\n"
                '{"section_id": "...", "ref_no": "", "revision": "R00", "doc_date": "YYYY-MM-DD", '
                '"authority": "", "subject_ar": "موضوع موجز واضح بالعربية أو الإنجليزية", '
                '"suggested_filename": "اسم ملف نظيف بدون امتداد مثل 2026-11-18_OUT_اعتماد_صب_السقف أو MAR-EL-045_R01_كابلات_الضغط"}'
            )

            try:
                if analyzer_fn is not None:
                    meta_ai = analyzer_fn(prompt, tmp_dir, images, snippet) or {}
                elif str(engine or "").strip().lower() == "ollama":
                    meta_ai = sw.analyze_with_ollama(
                        prompt,
                        images,
                        ollama_url or "http://127.0.0.1:11434/api/chat",
                        ollama_model or "qwen3-vl:8b",
                    ) or {}
                else:
                    meta_ai = sw.analyze_with_gemini(
                        prompt,
                        tmp_dir,
                        sw.load_saved_gemini_model(),
                        images=images,
                        api_key=gemini_api_key,
                    ) or {}
            except Exception as exc:
                errors.append(f"{it.get('file_name','')}: {exc}")
                processed += 1
                continue

            processed += 1
            if not isinstance(meta_ai, dict):
                continue

            if _apply_organizer_ai_meta(it, meta_ai, proj, sec_map, log_ref_map):
                it["ai_from_cache"] = False
                updated += 1
                if cache_key and not meta_ai.get("error"):
                    try:
                        sb = int(it.get("size_bytes") or fp.stat().st_size)
                    except OSError:
                        sb = int(it.get("size_bytes") or 0)
                    ai_cache.set_cached(cache_dir, "organizer_ai", cache_key, {
                        "meta_ai": meta_ai,
                        "size_bytes": sb,
                    })
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    atomic_write_json(pp, plan)
    return {
        "ok": True,
        "processed": processed,
        "updated": updated,
        "cache_hits": cache_hits,
        "total_targets": total_targets,
        "errors": errors[:30],
        "summary": summarize_organizer_plan(plan),
    }


def split_pdf_document(
    src_pdf: str | Path,
    dest_dir: str | Path,
    *,
    every_n_pages: int = 1,
    ranges_spec: str = "",
) -> dict[str, Any]:
    """Split a bulk multi-document scanner PDF into individual PDF files."""
    if fitz is None:
        raise RuntimeError(tr("organizer.pdf_lib_missing"))
    sp = resolve_local_or_unc(str(src_pdf))
    if not sp.exists() or not sp.is_file() or sp.suffix.lower() != ".pdf":
        raise FileNotFoundError(tr("server.file_not_found"))
    ddir = resolve_local_or_unc(str(dest_dir))
    ddir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(sp))
    total_pages = len(doc)
    groups: list[tuple[int, int]] = []  # 0-based inclusive (start, end)
    if ranges_spec.strip():
        for part in re.split(r"[,،;]+", ranges_spec.strip()):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a_s, b_s = part.split("-", 1)
                a = max(1, int(a_s.strip()))
                b = min(total_pages, int(b_s.strip()))
                if a <= b:
                    groups.append((a - 1, b - 1))
            else:
                p = int(part)
                if 1 <= p <= total_pages:
                    groups.append((p - 1, p - 1))
    else:
        step = max(1, int(every_n_pages))
        for start in range(0, total_pages, step):
            groups.append((start, min(total_pages - 1, start + step - 1)))

    created: list[str] = []
    created_paths: list[Path] = []
    try:
        for idx, (a, b) in enumerate(groups, 1):
            out_p = ddir / f"{_clean_seg(sp.stem, 50)}_part_{idx:03d}_p{a + 1}-{b + 1}.pdf"
            tmp_p = out_p.with_name(f".__tmp_{uuid.uuid4().hex[:6]}_{out_p.name}")
            out_doc = fitz.open()
            try:
                out_doc.insert_pdf(doc, from_page=a, to_page=b)
                out_doc.save(str(tmp_p))
            finally:
                out_doc.close()
            if not tmp_p.exists() or tmp_p.stat().st_size == 0:
                tmp_p.unlink(missing_ok=True)
                raise RuntimeError(tr("organizer.pdf_empty_err"))
            tmp_p.replace(out_p)
            created_paths.append(out_p)
            created.append(out_p.name)
    except Exception:
        for cp in created_paths:
            cp.unlink(missing_ok=True)
        raise
    finally:
        doc.close()
    if not created:
        raise ValueError(tr("organizer.pdf_empty_err"))
    return {"ok": True, "total_pages": total_pages, "parts_created": len(created), "files": created, "dest_dir": str(ddir)}


def merge_pdf_documents(src_pdfs: list[str | Path], out_pdf: str | Path) -> dict[str, Any]:
    """Merge multiple PDF files cleanly in sequence into a single PDF."""
    if fitz is None:
        raise RuntimeError(tr("organizer.pdf_lib_missing"))
    if not src_pdfs:
        raise ValueError(tr("server.file_not_found"))
    out_p = resolve_local_or_unc(str(out_pdf))
    out_p.parent.mkdir(parents=True, exist_ok=True)
    merged = fitz.open()
    tmp_p = out_p.with_name(f".__tmp_{uuid.uuid4().hex[:6]}_{out_p.name}")
    count = 0
    try:
        for item in src_pdfs:
            sp = resolve_local_or_unc(str(item))
            if sp.exists() and sp.is_file() and sp.suffix.lower() == ".pdf":
                sub = fitz.open(str(sp))
                try:
                    merged.insert_pdf(sub)
                finally:
                    sub.close()
                count += 1
        if count == 0 or len(merged) == 0:
            raise FileNotFoundError(tr("server.file_not_found"))
        total_pages = len(merged)
        merged.save(str(tmp_p))
    finally:
        merged.close()
    try:
        if not tmp_p.exists() or tmp_p.stat().st_size == 0:
            raise RuntimeError(tr("organizer.pdf_empty_err"))
        tmp_p.replace(out_p)
    except Exception:
        tmp_p.unlink(missing_ok=True)
        raise
    return {"ok": True, "merged_files": count, "total_pages": total_pages, "output_path": str(out_p), "name": out_p.name}
