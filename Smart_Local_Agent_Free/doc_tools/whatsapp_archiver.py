#!/usr/bin/env python3
"""
whatsapp_archiver.py — أرشفة وتسمية وسائط وملفات واتساب بالتاريخ والكلام المرتبط بها
====================================================================================
يستخرج رسائل ومرفقات تصدير واتساب (.zip أو مجلد يحتوي _chat.txt + صور/فيديو/مستندات)،
ويربط كل صورة/فيديو/مستند بالشرطين:
  1) تاريخ ووقت الصورة/الملف (من رسالة المحادثة، أو اسم ملف واتساب، أو EXIF/تاريخ التعديل).
  2) الكلام المكتوب قبل الصورة أو بعدها في نفس الدقيقة (أو ضمن نافذة دقائق قابلة للتعديل).

ويفصل المخرجات بوضوح إلى:
  - متطابق بالشرطين (تاريخ + كلام قبل/بعد/تعليق)
  - بتاريخ فقط (بدون كلام مطابق في نفس الدقيقة)
  - غير مطابق / مكرر (SHA256)

مع وضع التنفيذ الافتراضي: إعادة التسمية في نفس المكان (بدون حذف أي ملف + سجل تراجع + تقرير Excel).
"""
from __future__ import annotations

import datetime as dt
import re
import shutil
import threading
import time
import uuid
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any, Callable

try:
    from PIL import Image, ExifTags  # type: ignore
except Exception:
    Image = None  # type: ignore
    ExifTags = None  # type: ignore

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import ai_cache
from i18n_ar import tr
from security_utils import PathTraversalError, atomic_write_json, read_json_safe, safe_resolve_within, sha256_of_file

try:
    import scan_watcher as sw
except Exception:
    sw = None  # type: ignore


IMAGE_EXTS = {".jpg", ".jpeg", ".jfif", ".png", ".webp", ".gif", ".bmp", ".heic", ".heif", ".tif", ".tiff"}
VIDEO_EXTS = {".mp4", ".mov", ".avi", ".mkv", ".3gp", ".m4v", ".webm", ".wmv", ".mpeg", ".mpg"}
AUDIO_EXTS = {".opus", ".ogg", ".mp3", ".m4a", ".aac", ".wav", ".amr"}
DOC_EXTS = {".pdf", ".docx", ".doc", ".xlsx", ".xlsm", ".xls", ".pptx", ".ppt", ".txt", ".csv", ".rtf", ".dwg", ".dxf", ".zip", ".rar"}
SUPPORTED_MEDIA_EXTS = IMAGE_EXTS | VIDEO_EXTS | AUDIO_EXTS | DOC_EXTS

SKIP_FILENAMES_PREFIX = ("~$", ".", "WhatsApp_Archive_Report_", "WhatsApp_Undo_")
SKIP_DIR_NAMES = {"_System_Agent", "_Reports", "_Backups", "_WhatsApp_Sorted", "__MACOSX", ".git"}

_AR_DIGIT_TRANS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_BIDI_STRIP_RE = re.compile(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]")

# iOS export: [18/11/2026, 2:15:30 PM] Sender: Body   OR   [2026-11-18, 14:15:30] Sender: Body
_AMPM_PAT = r"(?:[APap]\.?[Mm]\.?|صباحاً|صباحا|مساءً|مساء|ق\.ظ|ب\.ظ|ص|م)"
_IOS_LINE_RE = re.compile(
    rf"^\[\s*(\d{{1,4}}[./\-]\d{{1,2}}[./\-]\d{{1,4}})\s*[,،]?\s*((?:{_AMPM_PAT}\s*)?\d{{1,2}}:\d{{2}}(?::\d{{2}})?(?:\s*{_AMPM_PAT})?)\s*\]\s*(.*)$"
)
# Android export: 18/11/2026, 14:15 - Sender: Body   OR   11/18/26, 2:15 pm - Sender: Body
_ANDROID_LINE_RE = re.compile(
    rf"^(\d{{1,4}}[./\-]\d{{1,2}}[./\-]\d{{1,4}})\s*[,،]?\s*((?:{_AMPM_PAT}\s*)?\d{{1,2}}:\d{{2}}(?::\d{{2}})?(?:\s*{_AMPM_PAT})?)\s*[-–—]\s*(.*)$"
)

_ATTACH_IOS_RE = re.compile(r"<\s*(?:attached|مرفق|ملف مرفق|المرفق|Anhang|pièce jointe|adjunto)\s*:\s*([^<>]+?)\s*>", re.I)
_ATTACH_BARE_RE = re.compile(r"<\s*(?:attached|مرفق|ملف مرفق|المرفق|Anhang|pièce jointe|adjunto)\s*>", re.I)
_ATTACH_ANDROID_RE = re.compile(
    r"([^\n<>:\"/\\|?*]+?\.[A-Za-z0-9]{2,6})\s*\((?:file attached|تم إرفاق الملف|تم إرفاق ملف|الملف مرفق|ملف مرفق|مرفق|Datei angehängt|archivo adjunto|fichier joint)\)",
    re.I,
)
_FILE_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Za-z0-9_\-.\u0600-\u06FF ]{3,120}\.(?:jpg|jpeg|jfif|png|webp|gif|bmp|heic|heif|tif|tiff|mp4|mov|avi|mkv|3gp|webm|m4v|pdf|docx|doc|xlsx|xlsm|xls|pptx|ppt|txt|csv|opus|ogg|mp3|m4a|dwg|dxf|zip|rar))(?![A-Za-z0-9_])", re.I)

_OMITTED_PATTERNS = (
    "<media omitted>",
    "media omitted",
    "image omitted",
    "video omitted",
    "audio omitted",
    "sticker omitted",
    "gif omitted",
    "document omitted",
    "contact card omitted",
    "الوسائط غير مدرجة",
    "الوسائط غير مضمنة",
    "<الوسائط غير مدرجة>",
    "تم حذف هذه الرسالة",
    "this message was deleted",
    "you deleted this message",
    "waiting for this message",
    "في انتظار هذه الرسالة",
    "messages and calls are end-to-end encrypted",
    "الرسائل والمكالمات مشفرة تماما",
    "الرسائل والمكالمات محمية بتشفير",
)

_EDITED_MARKER_RE = re.compile(r"\s*<\s*(?:This message was edited|تم تعديل هذه الرسالة)\s*>\s*", re.I)

# Filename date patterns: IMG-20261118-WA0005, PHOTO-2026-11-18-14-15-30, WhatsApp Image 2026-11-18 at 14.15.30
_FN_PHOTO_FULL_RE = re.compile(r"(?:PHOTO|VIDEO|AUDIO|DOC)-(\d{4})-(\d{2})-(\d{2})-(\d{2})-(\d{2})-(\d{2})", re.I)
_FN_WA_TEXT_RE = re.compile(r"(?:WhatsApp\s+(?:Image|Video|Audio|Document)|VID|IMG|AUD|DOC)[_\s-]+(\d{4})[-_.]?(\d{2})[-_.]?(\d{2})(?:\s+(?:at|في)\s+(\d{1,2})[.:](\d{2})(?:[.:](\d{2}))?(?:\s*([APap][Mm]|ص|م))?)?", re.I)
_FN_COMPACT_DT_RE = re.compile(r"\b(20\d{2})(\d{2})(\d{2})[_-](\d{2})(\d{2})(\d{2})?\b")
_FN_COMPACT_D_RE = re.compile(r"(?:IMG|VID|PTT|AUD|DOC|PXL|Screenshot)[_-]?(20\d{2})(\d{2})(\d{2})(?:[_-]|$)", re.I)
_FN_ISO_D_RE = re.compile(r"\b(20\d{2})[-_.](\d{2})[-_.](\d{2})\b")

_WIN_BAD_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_URL_RE = re.compile(r"https?://\S+", re.I)


def _clean_line(raw: str) -> str:
    s = _BIDI_STRIP_RE.sub("", str(raw or ""))
    s = s.replace("\u00a0", " ")
    return s.translate(_AR_DIGIT_TRANS)


def _resolve_local_path(p: str | Path) -> Path:
    raw = str(p or "").strip()
    if raw.startswith("\\\\") or raw.startswith("//"):
        return Path(raw)
    return Path(raw).resolve()


def _detect_media_kind(ext: str) -> str:
    e = str(ext or "").lower()
    if e in IMAGE_EXTS:
        return "image"
    if e in VIDEO_EXTS:
        return "video"
    if e in AUDIO_EXTS:
        return "audio"
    return "document"


def _parse_time_parts(t_str: str) -> tuple[int, int, int]:
    t = _clean_line(t_str).strip()
    ampm = ""
    m_ap = re.search(rf"({_AMPM_PAT})\s*$", t)
    if m_ap:
        ampm = m_ap.group(1).replace(".", "").lower()
        t = t[: m_ap.start()].strip()
    else:
        m_ap_pre = re.match(rf"^\s*({_AMPM_PAT})\s+", t)
        if m_ap_pre:
            ampm = m_ap_pre.group(1).replace(".", "").lower()
            t = t[m_ap_pre.end() :].strip()
    parts = [int(x) for x in t.split(":") if x.strip().isdigit()]
    hh = parts[0] if len(parts) >= 1 else 0
    mm = parts[1] if len(parts) >= 2 else 0
    ss = parts[2] if len(parts) >= 3 else 0
    if ampm in ("pm", "p", "م", "مساءً", "مساء", "بظ"):
        if hh < 12:
            hh += 12
    elif ampm in ("am", "a", "ص", "صباحاً", "صباحا", "قظ"):
        if hh == 12:
            hh = 0
    hh = max(0, min(23, hh))
    mm = max(0, min(59, mm))
    ss = max(0, min(59, ss))
    return hh, mm, ss


def _split_date_nums(d_str: str) -> tuple[int, int, int] | None:
    nums = [int(x) for x in re.split(r"[./\-]", _clean_line(d_str).strip()) if x.isdigit()]
    if len(nums) != 3:
        return None
    return nums[0], nums[1], nums[2]


def _normalize_year(y: int) -> int:
    if y < 100:
        return 2000 + y if y <= 79 else 1900 + y
    return y


def _infer_date_order(raw_dates: list[str], media_filenames: set[str] | None = None) -> str:
    """Detect whether numeric slash dates are DMY, MDY, or YMD."""
    dmy_only = 0
    mdy_only = 0
    ymd_hits = 0
    hint_dates: set[tuple[int, int, int]] = set()
    for fn in (media_filenames or set()):
        d_iso, _t, _src = extract_date_from_filename(fn)
        if d_iso:
            try:
                p = [int(x) for x in d_iso.split("-")]
                hint_dates.add((p[0], p[1], p[2]))
            except Exception:
                pass

    dmy_fn_matches = 0
    mdy_fn_matches = 0
    for rd in raw_dates:
        nums = _split_date_nums(rd)
        if not nums:
            continue
        a, b, c = nums
        if a >= 1000:
            ymd_hits += 1
            continue
        yr = _normalize_year(c)
        if a > 12 and 1 <= b <= 12:
            dmy_only += 1
        elif b > 12 and 1 <= a <= 12:
            mdy_only += 1
        if hint_dates:
            if 1 <= b <= 12 and 1 <= a <= 31 and (yr, b, a) in hint_dates:
                dmy_fn_matches += 1
            if 1 <= a <= 12 and 1 <= b <= 31 and (yr, a, b) in hint_dates:
                mdy_fn_matches += 1

    if ymd_hits > max(dmy_only, mdy_only) and ymd_hits > len(raw_dates) // 2:
        return "YMD"
    if mdy_only > dmy_only:
        return "MDY"
    if dmy_only > mdy_only:
        return "DMY"
    if mdy_fn_matches > dmy_fn_matches:
        return "MDY"
    return "DMY"


def _build_datetime(d_str: str, t_str: str, order: str = "DMY") -> dt.datetime | None:
    nums = _split_date_nums(d_str)
    if not nums:
        return None
    a, b, c = nums
    if a >= 1000:
        yr, mo, dy = a, b, c
    elif order == "MDY":
        mo, dy, yr = a, b, _normalize_year(c)
    else:
        dy, mo, yr = a, b, _normalize_year(c)
    # Fallback flip if one order produces invalid month (>12)
    if mo > 12 and 1 <= dy <= 12:
        dy, mo = mo, dy
    hh, mm, ss = _parse_time_parts(t_str)
    try:
        return dt.datetime(yr, mo, dy, hh, mm, ss)
    except ValueError:
        return None


def _is_ignored_body(text: str) -> bool:
    low = _clean_line(text).strip().lower()
    if not low:
        return True
    for pat in _OMITTED_PATTERNS:
        if pat in low:
            return True
    return False


def _is_omitted_media_body(text: str) -> bool:
    low = _clean_line(text).strip().lower()
    return any(
        p in low
        for p in (
            "media omitted",
            "image omitted",
            "video omitted",
            "document omitted",
            "الوسائط غير مدرجة",
            "الوسائط غير مضمنة",
        )
    )


def _split_sender_and_body(rest: str) -> tuple[str, str, bool]:
    """Separate 'Sender: body' from system notification messages."""
    if ":" not in rest:
        return "", rest.strip(), True
    idx = rest.find(":")
    sender = rest[:idx].strip()
    body = rest[idx + 1 :].strip()
    # Guard against accidental colon inside a URL or long system sentence
    if not sender or len(sender) > 85 or sender.lower().startswith(("http", "https")):
        return "", rest.strip(), True
    return sender, body, False


_PAGE_BADGE_RE = re.compile(r"\s*[•·\-]\s*\d+\s*(?:صفحة|صفحات|pages?|sheets?)\s*(?:[•·\-]\s*[A-Za-z0-9]+)?", re.I)


def _strip_ios_prefix(name: str) -> str:
    return re.sub(r"^\d{6,10}-", "", str(name or "").strip())


def _is_caption_just_filename(caption_txt: str, file_name: str) -> bool:
    cap = _PAGE_BADGE_RE.sub("", _clean_line(caption_txt)).strip(" \t\r\n._-•·()")
    fn = _clean_line(file_name).strip()
    if not cap or not fn:
        return False
    cap_l = cap.lower()
    cap_no_pfx = _strip_ios_prefix(cap_l)
    fn_l = fn.lower()
    fn_no_pfx = _strip_ios_prefix(fn_l)
    stem_l = Path(fn_l).stem.strip(" ._-")
    stem_no_pfx = Path(fn_no_pfx).stem.strip(" ._-")
    candidates = {fn_l, fn_no_pfx, stem_l, stem_no_pfx} - {""}
    if cap_l in candidates or cap_no_pfx in candidates:
        return True
    for cp in (cap_l, cap_no_pfx):
        if any(Path(cp).stem.strip(" ._-") == c and Path(cp).suffix.lower() == Path(fn_l).suffix.lower() for c in (stem_l, stem_no_pfx) if c):
            return True
    return False


def _strip_redundant_date_from_title(title: str, doc_date: str) -> str:
    """Remove leading/trailing duplicate date tokens matching `doc_date` (YYYY-MM-DD or DD-MM-YYYY)."""
    s = str(title or "").strip(" _-.")
    if not s or not doc_date or not re.match(r"^\d{4}-\d{2}-\d{2}$", doc_date):
        return s
    yr, mo, dy = doc_date.split("-")
    date_variants = [
        f"{yr}[-_.]{mo}[-_.]{dy}",
        f"{dy}[-_.]{mo}[-_.]{yr}",
        f"{int(dy)}[-_.]{int(mo)}[-_.]{yr}",
        f"{dy}[-_.]{mo}[-_.]{yr[2:]}",
    ]
    pat = r"(?:" + "|".join(date_variants) + r")"
    s = re.sub(rf"^{pat}[_\s.-]*", "", s)
    s = re.sub(rf"[_\s.-]*{pat}$", "", s)
    return s.strip(" _-.")


def _extract_attachments_and_caption(body: str, known_names_lower: dict[str, str] | None = None) -> tuple[list[str], str]:
    """Extract attached filenames and any accompanying caption text from a single message body."""
    cleaned = _EDITED_MARKER_RE.sub(" ", _clean_line(body)).strip()
    attachments: list[str] = []
    seen_att: set[str] = set()
    work = cleaned

    def _add_att(raw_fn: str) -> None:
        fn = _clean_line(raw_fn).strip().strip("\"'")
        fn = Path(fn.replace("\\", "/")).name
        if not fn:
            return
        kl = fn.lower()
        if known_names_lower and kl in known_names_lower:
            fn = known_names_lower[kl]
        if kl not in seen_att:
            seen_att.add(kl)
            attachments.append(fn)

    for m in _ATTACH_IOS_RE.finditer(work):
        _add_att(m.group(1))
    work = _ATTACH_IOS_RE.sub(" ", work)

    for m in _ATTACH_ANDROID_RE.finditer(work):
        _add_att(m.group(1))
    work = _ATTACH_ANDROID_RE.sub(" ", work)

    had_bare_attach = bool(_ATTACH_BARE_RE.search(work))
    if had_bare_attach:
        work = _ATTACH_BARE_RE.sub(" ", work)

    # Strip any duplicate occurrence of detected attachment filenames / page-count badges from caption
    for att_fn in list(attachments):
        for cand in {att_fn, _strip_ios_prefix(att_fn)}:
            if cand and cand.lower() in work.lower():
                work = re.sub(re.escape(cand), " ", work, flags=re.I)

    # Check if a line in the body directly names a known exported file
    if known_names_lower:
        stripped_known = {_strip_ios_prefix(k): v for k, v in known_names_lower.items()}
        kept_lines: list[str] = []
        for ln in work.splitlines():
            ln_s = _PAGE_BADGE_RE.sub(" ", ln).strip()
            ln_low = ln_s.lower()
            matched_in_line = False
            if ln_low in known_names_lower:
                _add_att(known_names_lower[ln_low])
                ln_s = ""
                matched_in_line = True
            elif _strip_ios_prefix(ln_low) in stripped_known:
                _add_att(stripped_known[_strip_ios_prefix(ln_low)])
                ln_s = ""
                matched_in_line = True
            else:
                for tm in _FILE_TOKEN_RE.finditer(ln_s):
                    cand = tm.group(1).strip()
                    cl = cand.lower()
                    if cl in known_names_lower:
                        _add_att(known_names_lower[cl])
                        ln_s = ln_s.replace(cand, " ")
                        matched_in_line = True
                    elif _strip_ios_prefix(cl) in stripped_known:
                        _add_att(stripped_known[_strip_ios_prefix(cl)])
                        ln_s = ln_s.replace(cand, " ")
                        matched_in_line = True
            ln_after = re.sub(r"\s*\((?:file attached|تم إرفاق الملف|تم إرفاق ملف|الملف مرفق|ملف مرفق|مرفق)\)\s*", " ", ln_s, flags=re.I).strip()
            if ln_after and not (matched_in_line and len(ln_after) <= 1):
                kept_lines.append(ln_after)
        work = "\n".join(kept_lines)
    else:
        for tm in _FILE_TOKEN_RE.finditer(work):
            tok = tm.group(1).strip()
            if had_bare_attach or tok.lower().startswith(("img-", "vid-", "aud-", "doc-", "ptt-", "0000", "photo-", "video-", "whatsapp ")):
                _add_att(tok)
                work = work.replace(tok, " ")

    work = _PAGE_BADGE_RE.sub(" ", work)
    caption_lines = [ln.strip() for ln in work.splitlines() if ln.strip() and not _is_ignored_body(ln)]
    caption = " ".join(caption_lines).strip()
    if attachments and any(_is_caption_just_filename(caption, att) for att in attachments):
        caption = ""
    return attachments, caption


def parse_whatsapp_chat_text(
    raw_text: str,
    *,
    media_filenames: set[str] | None = None,
    date_order: str = "auto",
) -> dict[str, Any]:
    """
    Parse WhatsApp exported chat text (`_chat.txt` / `WhatsApp Chat with ...txt`)
    across iOS, Android, Arabic, and English formats.
    """
    lines = str(raw_text or "").splitlines()
    raw_entries: list[tuple[str, str, str]] = []  # (date_str, time_str, rest)

    for raw_ln in lines:
        ln = _clean_line(raw_ln).rstrip("\r\n")
        if not ln.strip():
            if raw_entries:
                d0, t0, r0 = raw_entries[-1]
                raw_entries[-1] = (d0, t0, r0 + "\n")
            continue
        m_ios = _IOS_LINE_RE.match(ln.strip())
        if m_ios:
            raw_entries.append((m_ios.group(1), m_ios.group(2), m_ios.group(3)))
            continue
        m_and = _ANDROID_LINE_RE.match(ln.strip())
        if m_and:
            raw_entries.append((m_and.group(1), m_and.group(2), m_and.group(3)))
            continue
        if raw_entries:
            d0, t0, r0 = raw_entries[-1]
            raw_entries[-1] = (d0, t0, r0 + "\n" + ln)

    order = date_order if date_order in ("DMY", "MDY", "YMD") else _infer_date_order(
        [e[0] for e in raw_entries], media_filenames=media_filenames
    )

    known_names_lower = {fn.lower(): fn for fn in (media_filenames or set())}
    messages: list[dict[str, Any]] = []

    for d_str, t_str, rest in raw_entries:
        msg_dt = _build_datetime(d_str, t_str, order=order)
        if msg_dt is None:
            continue
        sender, body, is_sys = _split_sender_and_body(rest)
        if is_sys:
            attachments, caption = ([], "")
            clean_text = ""
        else:
            attachments, caption = _extract_attachments_and_caption(body, known_names_lower=known_names_lower)
            clean_text = caption if attachments else (
                "" if _is_ignored_body(body) else _EDITED_MARKER_RE.sub(" ", _clean_line(body)).strip()
            )
            clean_text = re.sub(r"\s+", " ", clean_text).strip()

        has_sec = ":" in t_str and t_str.count(":") == 2
        messages.append({
            "idx": len(messages),
            "dt": msg_dt,
            "ts": msg_dt.timestamp(),
            "date_iso": msg_dt.strftime("%Y-%m-%d"),
            "time_str": msg_dt.strftime("%H:%M:%S") if has_sec else msg_dt.strftime("%H:%M"),
            "has_seconds": has_sec,
            "minute_key": msg_dt.strftime("%Y-%m-%d %H:%M"),
            "sender": sender,
            "is_system": is_sys,
            "is_omitted_media": (not is_sys and _is_omitted_media_body(body)),
            "attachments": attachments,
            "caption": caption if attachments else "",
            "clean_text": clean_text,
            "raw_body": body.strip(),
        })

    return {
        "date_order": order,
        "messages": messages,
        "message_count": len(messages),
    }


def extract_date_from_filename(filename: str) -> tuple[str, str, str]:
    """Return (YYYY-MM-DD, HH:MM, source_tag) parsed from WhatsApp or camera filenames."""
    fn = _clean_line(filename)
    m = _FN_PHOTO_FULL_RE.search(fn)
    if m:
        yr, mo, dy, hh, mm, ss = [int(x) for x in m.groups()]
        try:
            d = dt.datetime(yr, mo, dy, hh, mm, ss)
            return d.strftime("%Y-%m-%d"), d.strftime("%H:%M"), "filename"
        except ValueError:
            pass

    m = _FN_WA_TEXT_RE.search(fn)
    if m:
        yr, mo, dy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        hh = int(m.group(4) or 0)
        mm = int(m.group(5) or 0)
        ap = (m.group(7) or "").lower()
        if ap in ("pm", "م") and hh < 12:
            hh += 12
        elif ap in ("am", "ص") and hh == 12:
            hh = 0
        try:
            d = dt.datetime(yr, mo, dy, hh, mm, 0)
            t_str = d.strftime("%H:%M") if m.group(4) is not None else ""
            return d.strftime("%Y-%m-%d"), t_str, "filename"
        except ValueError:
            pass

    m = _FN_COMPACT_DT_RE.search(fn)
    if m:
        yr, mo, dy, hh, mm = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5))
        try:
            d = dt.datetime(yr, mo, dy, hh, mm, 0)
            return d.strftime("%Y-%m-%d"), d.strftime("%H:%M"), "filename"
        except ValueError:
            pass

    m = _FN_COMPACT_D_RE.search(fn)
    if m:
        yr, mo, dy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            d = dt.date(yr, mo, dy)
            return d.strftime("%Y-%m-%d"), "", "filename"
        except ValueError:
            pass

    m = _FN_ISO_D_RE.search(fn)
    if m:
        yr, mo, dy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            d = dt.date(yr, mo, dy)
            return d.strftime("%Y-%m-%d"), "", "filename"
        except ValueError:
            pass

    return "", "", ""


def extract_date_from_file_or_meta(fp: Path) -> tuple[str, str, str]:
    d_iso, t_str, src = extract_date_from_filename(fp.name)
    if d_iso:
        return d_iso, t_str, src

    if Image is not None and fp.suffix.lower() in {".jpg", ".jpeg", ".jfif", ".tif", ".tiff"}:
        try:
            with Image.open(fp) as im:
                exif = im.getexif()
                if exif:
                    for tag_id in (36867, 36868, 306):  # DateTimeOriginal, DateTimeDigitized, DateTime
                        val = exif.get(tag_id)
                        if isinstance(val, str) and len(val) >= 10:
                            m = re.match(r"(\d{4})[:\-](\d{2})[:\-](\d{2})(?:\s+(\d{2}):(\d{2}))?", val.strip())
                            if m:
                                d_s = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                                t_s = f"{m.group(4)}:{m.group(5)}" if m.group(4) else ""
                                return d_s, t_s, "exif"
        except Exception:
            pass

    try:
        st = fp.stat()
        mtime = dt.datetime.fromtimestamp(st.st_mtime)
        return mtime.strftime("%Y-%m-%d"), mtime.strftime("%H:%M"), "mtime"
    except Exception:
        return "", "", ""


def _decode_zip_member_name(info: zipfile.ZipInfo) -> str:
    fn = info.filename
    if info.flag_bits & 0x800:
        return fn
    try:
        raw = fn.encode("cp437")
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            return raw.decode("cp1256", errors="replace")
    except Exception:
        return fn


def extract_zip_safely(
    zip_path: Path,
    dest_dir: Path,
    *,
    max_files: int = 5000,
    max_total_bytes: int = 2 * 1024 * 1024 * 1024,
) -> dict[str, Any]:
    """Extract a WhatsApp export ZIP safely (Zip-Slip, symlink, and Zip-Bomb guarded)."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_res = dest_dir.resolve()
    extracted_files: list[str] = []
    skipped = 0
    total_written = 0

    with zipfile.ZipFile(zip_path, "r") as zf:
        infos = [i for i in zf.infolist() if not i.is_dir()]
        if len(infos) > max_files:
            raise ValueError(tr("dups.zip_too_large_err"))
        declared_total = sum(max(0, int(getattr(i, "file_size", 0) or 0)) for i in infos)
        if declared_total > max_total_bytes:
            raise ValueError(tr("dups.zip_too_large_err"))
        for info in infos:
            member_name = _decode_zip_member_name(info).replace("\\", "/").lstrip("/")
            if not member_name or member_name.startswith("__MACOSX/"):
                skipped += 1
                continue
            # Reject symlinks inside ZIP (Unix external_attr S_IFLNK = 0o120000)
            if ((info.external_attr >> 16) & 0o170000) == 0o120000:
                skipped += 1
                continue
            parts = [p for p in member_name.split("/") if p not in ("", ".")]
            if not parts or any(p == ".." for p in parts):
                raise PathTraversalError(tr("server.path_blocked"))
            safe_parts = [_WIN_BAD_CHARS_RE.sub("_", p).strip() or "_" for p in parts]
            target = (dest_res / Path(*safe_parts)).resolve()
            try:
                target.relative_to(dest_res)
            except ValueError as exc:
                raise PathTraversalError(tr("server.path_blocked")) from exc
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info, "r") as src_f, open(target, "wb") as out_f:
                while True:
                    chunk = src_f.read(64 * 1024)
                    if not chunk:
                        break
                    total_written += len(chunk)
                    if total_written > max_total_bytes:
                        raise ValueError(tr("dups.zip_too_large_err"))
                    out_f.write(chunk)
            extracted_files.append(target.relative_to(dest_res).as_posix())

    return {
        "ok": True,
        "dest_dir": str(dest_res),
        "extracted_count": len(extracted_files),
        "skipped_count": skipped,
        "files": extracted_files,
    }


def _minute_diff(dt_a: dt.datetime, dt_b: dt.datetime) -> int:
    ma = dt_a.replace(second=0, microsecond=0)
    mb = dt_b.replace(second=0, microsecond=0)
    return abs(int((ma - mb).total_seconds() // 60))


def _messages_within_window(
    dt_a: dt.datetime,
    has_sec_a: bool,
    dt_b: dt.datetime,
    has_sec_b: bool,
    window_minutes: int,
) -> bool:
    if window_minutes <= 0:
        if dt_a.replace(second=0, microsecond=0) == dt_b.replace(second=0, microsecond=0):
            return True
        if (has_sec_a or has_sec_b) and abs((dt_a - dt_b).total_seconds()) <= 59:
            return True
        return False
    if _minute_diff(dt_a, dt_b) <= window_minutes:
        return True
    return abs((dt_a - dt_b).total_seconds()) <= window_minutes * 60 + 59


def _sender_matches(msg_sender: str, target_sender: str, sender_mode: str) -> bool:
    if sender_mode == "any_sender" or not target_sender:
        return True
    return msg_sender.strip().lower() == target_sender.strip().lower()


def _effective_msg_caption(mj: dict[str, Any]) -> str:
    cap_j = str(mj.get("caption") or "").strip()
    if not cap_j:
        return ""
    atts = [str(a) for a in (mj.get("attachments") or [])]
    if atts and any(_is_caption_just_filename(cap_j, a) for a in atts):
        return ""
    return cap_j


def _collect_context_texts(
    messages: list[dict[str, Any]],
    msg_idx: int,
    *,
    window_minutes: int = 0,
    sender_mode: str = "same_sender",
) -> tuple[str, str, str]:
    """
    Collect (text_before, caption, text_after) for attachment message at `msg_idx`.
    Default `window_minutes=0` requires the exact same minute (or <= 59s when seconds exist).
    Respects attachment burst boundaries and shares single album captions across consecutive
    attachment-only messages from the same sender.
    """
    cur = messages[msg_idx]
    cur_dt: dt.datetime = cur["dt"]
    cur_sec = bool(cur.get("has_seconds"))
    cur_sender = cur.get("sender") or ""
    caption = _effective_msg_caption(cur)

    # Walk backward: step along contiguous album attachments from same sender, then capture contiguous text block
    before_chunks: list[str] = []
    album_caps_before: list[str] = []
    started_text_block = False
    anchor_dt_b = cur_dt
    anchor_sec_b = cur_sec
    for j in range(msg_idx - 1, -1, -1):
        mj = messages[j]
        mj_sec = bool(mj.get("has_seconds"))
        if not _messages_within_window(anchor_dt_b, anchor_sec_b, mj["dt"], mj_sec, window_minutes):
            break
        if mj.get("is_system"):
            continue
        if not _sender_matches(mj.get("sender") or "", cur_sender, sender_mode):
            if sender_mode == "same_sender":
                continue
            break
        has_att = bool(mj.get("attachments")) or bool(mj.get("is_omitted_media"))
        txt = str(mj.get("clean_text") or "").strip()
        if has_att:
            if started_text_block:
                break
            if _sender_matches(mj.get("sender") or "", cur_sender, "same_sender"):
                anchor_dt_b = mj["dt"]
                anchor_sec_b = mj_sec
                cap_j = _effective_msg_caption(mj)
                if cap_j:
                    album_caps_before.append(cap_j)
            continue
        if txt:
            started_text_block = True
            before_chunks.append(txt)
        else:
            if started_text_block:
                break

    before_chunks.reverse()
    text_before = " - ".join(before_chunks).strip()

    # Walk forward: step along contiguous album attachments from same sender, then capture contiguous text block
    after_chunks: list[str] = []
    album_caps_after: list[str] = []
    started_after = False
    anchor_dt_a = cur_dt
    anchor_sec_a = cur_sec
    for j in range(msg_idx + 1, len(messages)):
        mj = messages[j]
        mj_sec = bool(mj.get("has_seconds"))
        if not _messages_within_window(anchor_dt_a, anchor_sec_a, mj["dt"], mj_sec, window_minutes):
            break
        if mj.get("is_system"):
            continue
        if not _sender_matches(mj.get("sender") or "", cur_sender, sender_mode):
            if sender_mode == "same_sender":
                continue
            break
        has_att = bool(mj.get("attachments")) or bool(mj.get("is_omitted_media"))
        txt = str(mj.get("clean_text") or "").strip()
        if has_att:
            if started_after:
                break
            if _sender_matches(mj.get("sender") or "", cur_sender, "same_sender"):
                anchor_dt_a = mj["dt"]
                anchor_sec_a = mj_sec
                cap_j = _effective_msg_caption(mj)
                if cap_j:
                    album_caps_after.append(cap_j)
            continue
        if txt:
            started_after = True
            after_chunks.append(txt)
        else:
            if started_after:
                break

    text_after = " - ".join(after_chunks).strip()

    # Share a single album-level caption across consecutive uncaptioned attachments in the same burst
    if not caption:
        all_album_caps = album_caps_before + album_caps_after
        unique_album_caps = list(dict.fromkeys(all_album_caps))
        if len(unique_album_caps) == 1:
            if not text_after and album_caps_after:
                text_after = album_caps_after[0]
            elif not text_before and album_caps_before:
                text_before = album_caps_before[0]

    return text_before, caption, text_after


def _collect_context_texts_by_time(
    messages: list[dict[str, Any]],
    doc_date: str,
    msg_time: str,
    *,
    window_minutes: int = 0,
    sender_mode: str = "same_sender",
    target_sender: str = "",
) -> tuple[str, str, str]:
    """Fallback context lookup by timestamp when a media file has no direct `msg_idx`."""
    if not messages or not doc_date or not msg_time:
        return "", "", ""
    try:
        t_clean = msg_time.strip()
        fmt = "%Y-%m-%d %H:%M:%S" if t_clean.count(":") == 2 else "%Y-%m-%d %H:%M"
        target_dt = dt.datetime.strptime(f"{doc_date} {t_clean}", fmt)
        has_sec = t_clean.count(":") == 2
    except Exception:
        return "", "", ""

    in_win: list[dict[str, Any]] = []
    for mj in messages:
        if mj.get("is_system"):
            continue
        if _messages_within_window(target_dt, has_sec, mj["dt"], bool(mj.get("has_seconds")), window_minutes):
            if target_sender and not _sender_matches(mj.get("sender") or "", target_sender, sender_mode):
                continue
            in_win.append(mj)

    if not in_win:
        return "", "", ""
    if sender_mode == "same_sender" and not target_sender:
        senders = {str(m.get("sender") or "").strip() for m in in_win if str(m.get("sender") or "").strip()}
        if len(senders) > 1:
            # Prefer texts from the nearest message's sender
            nearest = min(in_win, key=lambda m: abs((m["dt"] - target_dt).total_seconds()))
            ns = str(nearest.get("sender") or "").strip()
            if ns:
                in_win = [m for m in in_win if (m.get("sender") or "").strip().lower() == ns.lower()]

    bef_chunks = [str(m.get("clean_text") or "").strip() for m in in_win if m["dt"] <= target_dt and str(m.get("clean_text") or "").strip()]
    aft_chunks = [str(m.get("clean_text") or "").strip() for m in in_win if m["dt"] > target_dt and str(m.get("clean_text") or "").strip()]
    return " - ".join(bef_chunks).strip(), "", " - ".join(aft_chunks).strip()


def _pick_chosen_text(
    text_before: str,
    caption: str,
    text_after: str,
    *,
    text_priority: str = "before_then_after",
) -> tuple[str, str]:
    cap = caption.strip()
    bef = text_before.strip()
    aft = text_after.strip()

    if text_priority == "combine":
        parts: list[str] = []
        srcs: list[str] = []
        if cap:
            parts.append(cap)
            srcs.append("caption")
        if bef and bef not in parts:
            parts.append(bef)
            srcs.append("before")
        if aft and aft not in parts:
            parts.append(aft)
            srcs.append("after")
        if not parts:
            return "", "none"
        src_tag = "before_and_after" if ("before" in srcs and "after" in srcs) else srcs[0]
        return " - ".join(parts), src_tag

    if text_priority == "after_then_before":
        if cap:
            return cap, "caption"
        if aft:
            return aft, "after"
        if bef:
            return bef, "before"
        return "", "none"

    # Default: before_then_after
    if cap:
        return cap, "caption"
    if bef:
        return bef, "before"
    if aft:
        return aft, "after"
    return "", "none"


def sanitize_wa_filename_part(text: str, *, max_len: int = 75) -> str:
    s = _clean_line(text)
    s = _URL_RE.sub(" ", s)
    s = _WIN_BAD_CHARS_RE.sub(" ", s)
    # Remove emoji / dingbats while preserving Arabic, Latin, digits, underscores, dashes, dots
    s = re.sub(r"[\U00010000-\U0010ffff\u2600-\u27bf]", " ", s)
    s = re.sub(r"[،,;؛!؟\"'`~@#$%^&*+=\[\]{}()]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" ._-")
    if len(s) > max_len:
        cut = s[:max_len].rsplit(" ", 1)[0].strip(" ._-")
        s = cut if len(cut) >= 12 else s[:max_len].strip(" ._-")
    s = re.sub(r"\s+", "_", s).strip("._-")
    return s


def build_proposed_filename(
    *,
    doc_date: str,
    msg_time: str,
    sender: str,
    chosen_text: str,
    orig_name: str,
    ext: str,
    name_format: str = "date_name",
) -> str:
    clean_ext = ext if str(ext).startswith(".") else (f".{ext}" if ext else "")
    d_part = doc_date or "undated"
    t_part = msg_time.replace(":", "-")[:5] if msg_time else ""
    sender_part = sanitize_wa_filename_part(sender, max_len=24)
    title_part = sanitize_wa_filename_part(chosen_text, max_len=75)
    title_part = _strip_redundant_date_from_title(title_part, doc_date)

    if title_part:
        if name_format == "date_time_name" and t_part:
            base = f"{d_part}_{t_part}_{title_part}"
        elif name_format == "date_sender_name" and sender_part:
            base = f"{d_part}_{sender_part}_{title_part}"
        else:
            base = f"{d_part}_{title_part}"
    else:
        orig_stem = sanitize_wa_filename_part(Path(orig_name).stem, max_len=45)
        orig_stem = _strip_redundant_date_from_title(orig_stem, doc_date) or "media"
        if t_part:
            base = f"{d_part}_{t_part}_{orig_stem}"
        else:
            base = f"{d_part}_{orig_stem}"

    base = _WIN_BAD_CHARS_RE.sub("_", base).strip(" ._-")[:140] or "whatsapp_file"
    return f"{base}{clean_ext}"


def _assign_unique_names(items: list[dict[str, Any]]) -> None:
    """
    Assign deterministic non-colliding `new_name` values for items inside the same directory.
    If multiple photos share `2026-11-18_صب_السقف.jpg`, suffix them `_01`, `_02`, ...
    """
    by_dir_stem: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for it in items:
        parent_rel = str(Path(str(it.get("src_rel") or "")).parent.as_posix())
        raw_new = str(
            it.get("new_name")
            if it.get("custom_name")
            else (it.get("base_new_name") or it.get("new_name") or it.get("file_name") or "file")
        )
        p = Path(raw_new)
        stem = p.stem
        ext = p.suffix.lower() or str(it.get("ext") or "").lower()
        by_dir_stem[(parent_rel, stem.lower(), ext)].append(it)

    for (_dir_rel, _stem_low, ext), grp in by_dir_stem.items():
        if len(grp) == 1:
            it = grp[0]
            if not it.get("custom_name"):
                it["new_name"] = f"{Path(str(it.get('base_new_name') or it.get('new_name') or it.get('file_name'))).stem}{ext}"
                it["_auto_suffixed"] = False
            continue
        grp.sort(key=lambda x: (str(x.get("doc_date") or ""), str(x.get("msg_time") or ""), int(x["msg_idx"]) if x.get("msg_idx") is not None else 10**9, str(x.get("file_name") or "")))
        used_custom = {str(x.get("new_name") or "").lower() for x in grp if x.get("custom_name")}
        seq = 1
        for it in grp:
            if it.get("custom_name"):
                continue
            stem_base = Path(str(it.get("base_new_name") or it.get("new_name") or it.get("file_name"))).stem
            while True:
                cand = f"{stem_base}_{seq:02d}{ext}"
                seq += 1
                if cand.lower() not in used_custom:
                    used_custom.add(cand.lower())
                    it["new_name"] = cand
                    it["_auto_suffixed"] = True
                    break


def whatsapp_plan_path(proj_root: Path) -> Path:
    return proj_root / "_System_Agent" / "whatsapp_archive_plan.json"


def whatsapp_undo_path(proj_root: Path) -> Path:
    return proj_root / "_System_Agent" / "whatsapp_archive_undo.json"


def whatsapp_chats_root(proj_root: Path) -> Path:
    return proj_root / "00_WhatsApp_Chats"


def _find_chat_txt_files(src_dir: Path) -> list[Path]:
    cands: list[Path] = []
    for f in src_dir.rglob("*.txt"):
        if not f.is_file():
            continue
        if any(part in SKIP_DIR_NAMES for part in f.relative_to(src_dir).parts):
            continue
        if f.name.startswith(SKIP_FILENAMES_PREFIX):
            continue
        try:
            head = f.read_text(encoding="utf-8", errors="replace")[:4000]
        except Exception:
            continue
        for ln in head.splitlines()[:25]:
            cl = _clean_line(ln).strip()
            if _IOS_LINE_RE.match(cl) or _ANDROID_LINE_RE.match(cl):
                cands.append(f)
                break
    cands.sort(key=lambda p: (0 if "_chat" in p.stem.lower() else 1, len(p.parts), -p.stat().st_size))
    return cands


def _check_date_in_range(doc_date: str, date_from: str, date_to: str) -> bool:
    if not date_from and not date_to:
        return True
    if not doc_date:
        return False
    if date_from and doc_date < date_from:
        return False
    if date_to and doc_date > date_to:
        return False
    return True


def summarize_whatsapp_plan(plan: dict[str, Any] | None) -> dict[str, int]:
    items = (plan or {}).get("items") or []
    counts = {
        "total": 0,
        "total_all": len(items),
        "matched": 0,
        "date_only": 0,
        "unmatched": 0,
        "duplicates": 0,
        "included": 0,
        "excluded": 0,
        "in_range": 0,
        "out_of_range": 0,
        "images": 0,
        "videos": 0,
        "documents": 0,
        "executed": 0,
    }
    for it in items:
        in_rng = bool(it.get("in_date_range", True))
        if in_rng:
            counts["in_range"] += 1
            counts["total"] += 1
            ms = it.get("match_status") or "unmatched"
            if ms in counts:
                counts[ms] += 1
            if it.get("is_duplicate"):
                counts["duplicates"] += 1
            if it.get("include"):
                counts["included"] += 1
            else:
                counts["excluded"] += 1
            mk = it.get("media_kind") or "document"
            if mk == "image":
                counts["images"] += 1
            elif mk == "video":
                counts["videos"] += 1
            else:
                counts["documents"] += 1
            if it.get("executed"):
                counts["executed"] += 1
        else:
            counts["out_of_range"] += 1
    return counts


def _apply_matching_to_items(
    items: list[dict[str, Any]],
    messages: list[dict[str, Any]],
    *,
    window_minutes: int = 0,
    sender_mode: str = "same_sender",
    text_priority: str = "before_then_after",
    name_format: str = "date_name",
    date_from: str = "",
    date_to: str = "",
    include_docs: bool = True,
    preserve_custom: bool = False,
) -> None:
    for it in items:
        msg_idx = it.get("msg_idx")
        if isinstance(msg_idx, int) and 0 <= msg_idx < len(messages):
            msg = messages[msg_idx]
            it["doc_date"] = msg["date_iso"]
            it["msg_time"] = msg["time_str"]
            it["minute_key"] = msg["minute_key"]
            it["sender"] = msg.get("sender") or ""
            it["date_source"] = it.get("date_source") or "chat"
            tb, cap, ta = _collect_context_texts(
                messages, msg_idx, window_minutes=window_minutes, sender_mode=sender_mode
            )
            if cap and _is_caption_just_filename(cap, str(it.get("file_name") or "")):
                cap = ""
            it["text_before"] = tb
            it["caption"] = cap
            it["text_after"] = ta
            chosen, tsrc = _pick_chosen_text(tb, cap, ta, text_priority=text_priority)
            it["chosen_text"] = chosen
            it["text_source"] = tsrc
        else:
            tb = str(it.get("text_before") or "")
            cap = str(it.get("caption") or "")
            if cap and _is_caption_just_filename(cap, str(it.get("file_name") or "")):
                cap = ""
                it["caption"] = ""
            ta = str(it.get("text_after") or "")
            if (
                messages
                and not (tb or cap or ta)
                and it.get("doc_date")
                and it.get("msg_time")
                and it.get("date_source") in ("filename", "exif", "chat", "chat_time")
            ):
                tb, cap, ta = _collect_context_texts_by_time(
                    messages,
                    str(it.get("doc_date") or ""),
                    str(it.get("msg_time") or ""),
                    window_minutes=window_minutes,
                    sender_mode=sender_mode,
                    target_sender=str(it.get("sender") or ""),
                )
                it["text_before"] = tb
                it["caption"] = cap
                it["text_after"] = ta
            chosen, tsrc = _pick_chosen_text(tb, cap, ta, text_priority=text_priority)
            it["chosen_text"] = chosen
            it["text_source"] = tsrc

        has_date = bool(it.get("doc_date"))
        has_text = bool(str(it.get("chosen_text") or "").strip())
        if has_date and has_text:
            it["match_status"] = "matched"
        elif has_date:
            it["match_status"] = "date_only"
        else:
            it["match_status"] = "unmatched"

        in_range = _check_date_in_range(str(it.get("doc_date") or ""), date_from, date_to)
        it["in_date_range"] = in_range
        kind_ok = include_docs or (it.get("media_kind") in ("image", "video"))

        if not (preserve_custom and (it.get("custom_name") or (it.get("ai_renamed") and it.get("text_source") == "ai"))):
            base_nm = build_proposed_filename(
                doc_date=str(it.get("doc_date") or ""),
                msg_time=str(it.get("msg_time") or ""),
                sender=str(it.get("sender") or ""),
                chosen_text=str(it.get("chosen_text") or ""),
                orig_name=str(it.get("file_name") or ""),
                ext=str(it.get("ext") or ""),
                name_format=name_format,
            )
            it["base_new_name"] = base_nm
            it["new_name"] = base_nm
            it["custom_name"] = False
            it["ai_renamed"] = False

        if not in_range or not kind_ok:
            it["include"] = False
        elif preserve_custom and "manual_include" in it:
            it["include"] = bool(it["manual_include"])
        else:
            it["include"] = bool(not it.get("is_duplicate"))

    items.sort(key=lambda x: (
        str(x.get("doc_date") or "9999-99-99"),
        str(x.get("msg_time") or "99:99"),
        int(x["msg_idx"]) if x.get("msg_idx") is not None else 10**9,
        str(x.get("file_name") or "").lower(),
    ))
    _assign_unique_names(items)


class WhatsappStoppedError(Exception):
    """Raised when user stops a running WhatsApp archive scan job."""


def _emit_wa_prog(
    progress_cb: Callable[[dict[str, Any]], None] | None,
    last_ts: list[float],
    info: dict[str, Any],
    *,
    force: bool = False,
) -> None:
    if not progress_cb:
        return
    now = time.monotonic()
    if force or (now - last_ts[0]) >= 0.2:
        last_ts[0] = now
        try:
            progress_cb(info)
        except Exception:
            pass


def scan_and_match_whatsapp_source(
    proj: dict[str, Any],
    source_path: str | Path,
    *,
    window_minutes: int = 0,
    sender_mode: str = "same_sender",
    text_priority: str = "before_then_after",
    name_format: str = "date_name",
    date_from: str = "",
    date_to: str = "",
    include_docs: bool = True,
    batch_label: str = "",
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """
    Scan an exported WhatsApp folder or `.zip` file, parse chat logs,
    match every media/document item by Date + Text (before/after in the same minute),
    and save the reviewable plan for the project.
    """
    proj_root = _resolve_local_path(str(proj["root"]))
    last_emit = [0.0]

    def _chk_stop() -> None:
        if stop_event is not None and stop_event.is_set():
            raise WhatsappStoppedError(tr("jobs.stopped_by_user"))

    src_in = _resolve_local_path(source_path)
    if not src_in.exists():
        raise FileNotFoundError(tr("server.folder_not_found", path=str(source_path)))

    _emit_wa_prog(progress_cb, last_emit, {
        "phase": "walking", "processed": 0, "total": 0, "discovered": 0, "current_file": src_in.name,
    }, force=True)

    source_kind = "folder"
    extract_info: dict[str, Any] = {}
    work_dir = src_in

    if src_in.is_file() and src_in.suffix.lower() == ".zip":
        source_kind = "zip"
        _emit_wa_prog(progress_cb, last_emit, {
            "phase": "extracting", "processed": 0, "total": 0, "discovered": 0, "current_file": src_in.name,
        }, force=True)
        safe_stem = sanitize_wa_filename_part(src_in.stem, max_len=40) or "WhatsApp_Batch"
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        work_dir = whatsapp_chats_root(proj_root) / f"{stamp}_{safe_stem}"
        extract_info = extract_zip_safely(src_in, work_dir)
    elif not src_in.is_dir():
        raise ValueError(tr("server.folder_not_found", path=str(source_path)))

    # Collect all media/document files inside work_dir
    media_paths: list[Path] = []
    for fp in sorted(work_dir.rglob("*"), key=lambda p: p.as_posix().lower()):
        _chk_stop()
        if not fp.is_file():
            continue
        rel = fp.relative_to(work_dir)
        if any(part in SKIP_DIR_NAMES for part in rel.parts[:-1]):
            continue
        if fp.name.startswith(SKIP_FILENAMES_PREFIX):
            continue
        if fp.suffix.lower() not in SUPPORTED_MEDIA_EXTS:
            continue
        # Skip the chat transcript `.txt` itself when matching media items
        if fp.suffix.lower() == ".txt":
            try:
                head = fp.read_text(encoding="utf-8", errors="replace")[:2000]
                if any(_IOS_LINE_RE.match(_clean_line(ln).strip()) or _ANDROID_LINE_RE.match(_clean_line(ln).strip()) for ln in head.splitlines()[:15]):
                    continue
            except Exception:
                pass
        media_paths.append(fp)
        if len(media_paths) == 1 or len(media_paths) % 20 == 0:
            _emit_wa_prog(progress_cb, last_emit, {
                "phase": "walking", "processed": 0, "total": 0,
                "discovered": len(media_paths), "current_file": fp.name,
            })

    chat_files = _find_chat_txt_files(work_dir)
    media_names_set = {p.name for p in media_paths}

    _emit_wa_prog(progress_cb, last_emit, {
        "phase": "analyzing", "processed": 0, "total": max(1, len(media_paths)),
        "discovered": len(media_paths), "current_file": chat_files[0].name if chat_files else "",
    }, force=True)

    all_messages: list[dict[str, Any]] = []
    date_order_used = "DMY"
    for cf in chat_files:
        _chk_stop()
        try:
            raw_txt = cf.read_text(encoding="utf-8-sig", errors="replace")
        except Exception:
            continue
        parsed = parse_whatsapp_chat_text(raw_txt, media_filenames=media_names_set)
        date_order_used = parsed.get("date_order") or date_order_used
        base_offset = len(all_messages)
        for m in parsed["messages"]:
            m["idx"] = base_offset + m["idx"]
            all_messages.append(m)

    # Sort combined messages chronologically if multiple chat text files existed
    if len(chat_files) > 1:
        all_messages.sort(key=lambda m: (m["ts"], m["idx"]))
        for i, m in enumerate(all_messages):
            m["idx"] = i

    # Index attachments by filename -> list of message indices (in case same filename referenced once/more)
    att_index: dict[str, list[int]] = defaultdict(list)
    att_index_stripped: dict[str, list[int]] = defaultdict(list)
    by_minute_omitted: dict[str, list[int]] = defaultdict(list)
    by_minute_any: dict[str, list[int]] = defaultdict(list)

    def _norm_att_key(name_low: str) -> str:
        st_name = _strip_ios_prefix(name_low)
        p = Path(st_name)
        stem = re.sub(r"\s*\(\d+\)$", "", p.stem).strip()
        return f"{stem}{p.suffix.lower()}"

    for m in all_messages:
        for att in m.get("attachments") or []:
            al = att.lower()
            att_index[al].append(m["idx"])
            att_index_stripped[_norm_att_key(al)].append(m["idx"])
        if m.get("is_omitted_media"):
            by_minute_omitted[m["minute_key"]].append(m["idx"])
        if not m.get("is_system"):
            by_minute_any[m["minute_key"]].append(m["idx"])

    seen_hashes: dict[str, str] = {}
    items: list[dict[str, Any]] = []
    used_att_idx: set[int] = set()
    total_media = max(1, len(media_paths))

    for m_idx, fp in enumerate(media_paths):
        _chk_stop()
        _emit_wa_prog(progress_cb, last_emit, {
            "phase": "hashing", "processed": m_idx + 1, "total": total_media,
            "discovered": len(media_paths), "current_file": fp.name,
        })
        rel_str = fp.relative_to(work_dir).as_posix()
        try:
            fhash = sha256_of_file(fp)
        except Exception:
            fhash = ""
        is_dup = bool(fhash and fhash in seen_hashes)
        dup_of = seen_hashes.get(fhash, "") if is_dup else ""
        if fhash and fhash not in seen_hashes:
            seen_hashes[fhash] = fp.name

        size_kb = round(fp.stat().st_size / 1024.0, 1)
        ext = fp.suffix.lower()
        mkind = _detect_media_kind(ext)

        # Match message: 1) direct attachment filename match (or stripped iOS/copy-suffix match)
        matched_idx: int | None = None
        date_source = ""
        fn_low = fp.name.lower()
        fn_norm = _norm_att_key(fn_low)
        cand_list = att_index.get(fn_low) or att_index_stripped.get(fn_norm) or []
        if cand_list:
            for cand_i in cand_list:
                if cand_i not in used_att_idx:
                    matched_idx = cand_i
                    used_att_idx.add(cand_i)
                    break
            if matched_idx is None:
                matched_idx = cand_list[0]
            date_source = "chat"

        d_iso, t_str, meta_src = extract_date_from_file_or_meta(fp)
        # 2) Fallback: if chat had `<Media omitted>` or same-minute message matching reliable filename/EXIF timestamp
        if matched_idx is None and d_iso and t_str and all_messages:
            mkey = f"{d_iso} {t_str[:5]}"
            cands = by_minute_omitted.get(mkey) or []
            for cand_i in cands:
                if cand_i not in used_att_idx:
                    matched_idx = cand_i
                    used_att_idx.add(cand_i)
                    date_source = "chat_time"
                    break
            if matched_idx is None and meta_src in ("filename", "exif"):
                cands_any = by_minute_any.get(mkey) or []
                for cand_i in cands_any:
                    if cand_i not in used_att_idx and (all_messages[cand_i].get("attachments") or all_messages[cand_i].get("is_omitted_media")):
                        matched_idx = cand_i
                        used_att_idx.add(cand_i)
                        date_source = "chat_time"
                        break

        if matched_idx is not None:
            msg = all_messages[matched_idx]
            d_iso = msg["date_iso"]
            t_str = msg["time_str"]
            mkey = msg["minute_key"]
            sender = msg.get("sender") or ""
            date_source = date_source or "chat"
        else:
            mkey = f"{d_iso} {t_str[:5]}".strip() if d_iso else ""
            sender = ""
            date_source = meta_src or ""

        items.append({
            "id": uuid.uuid4().hex[:10],
            "file_name": fp.name,
            "src_path": str(fp),
            "src_rel": rel_str,
            "ext": ext,
            "media_kind": mkind,
            "size_kb": size_kb,
            "sha256": fhash,
            "is_duplicate": is_dup,
            "duplicate_of": dup_of,
            "msg_idx": matched_idx,
            "doc_date": d_iso,
            "msg_time": t_str,
            "minute_key": mkey,
            "date_source": date_source,
            "sender": sender,
            "text_before": "",
            "caption": "",
            "text_after": "",
            "chosen_text": "",
            "text_source": "none",
            "match_status": "unmatched",
            "in_date_range": True,
            "include": True,
            "base_new_name": fp.name,
            "new_name": fp.name,
            "custom_name": False,
            "ai_renamed": False,
            "executed": False,
        })

    win_m = max(0, min(60, int(window_minutes or 0)))
    s_mode = sender_mode if sender_mode in ("same_sender", "any_sender") else "same_sender"
    t_prio = text_priority if text_priority in ("before_then_after", "after_then_before", "combine") else "before_then_after"
    n_fmt = name_format if name_format in ("date_name", "date_time_name", "date_sender_name") else "date_name"

    _apply_matching_to_items(
        items,
        all_messages,
        window_minutes=win_m,
        sender_mode=s_mode,
        text_priority=t_prio,
        name_format=n_fmt,
        date_from=str(date_from or "").strip(),
        date_to=str(date_to or "").strip(),
        include_docs=bool(include_docs),
        preserve_custom=False,
    )

    # Hydrate any previously cached AI names by file SHA-256 + nearby chat context
    _wa_cache_dir = ai_cache.project_ai_cache_dir(proj_root)
    if ai_cache.cache_file_path(_wa_cache_dir, "whatsapp_ai").exists():
        hydrated_any = False
        for it in items:
            if it.get("match_status") not in ("date_only", "unmatched") or not it.get("sha256"):
                continue
            ckey = _wa_ai_cache_key(it)
            crec = ai_cache.get_cached(_wa_cache_dir, "whatsapp_ai", ckey)
            if crec and _apply_wa_ai_meta(it, crec, name_format=n_fmt):
                it["ai_from_cache"] = True
                hydrated_any = True
        if hydrated_any:
            _assign_unique_names(items)

    # Serialize messages (ISO string instead of raw datetime object)
    ser_messages = [
        {
            "idx": m["idx"],
            "date_iso": m["date_iso"],
            "time_str": m["time_str"],
            "has_seconds": bool(m.get("has_seconds")),
            "minute_key": m["minute_key"],
            "sender": m["sender"],
            "is_system": m["is_system"],
            "is_omitted_media": m["is_omitted_media"],
            "attachments": m["attachments"],
            "caption": m["caption"],
            "clean_text": m["clean_text"],
        }
        for m in all_messages
    ]

    plan = {
        "version": 1,
        "batch_id": uuid.uuid4().hex[:10],
        "batch_label": batch_label or src_in.name,
        "source_input": str(src_in),
        "source_dir": str(work_dir),
        "source_kind": source_kind,
        "chat_files": [cf.relative_to(work_dir).as_posix() for cf in chat_files],
        "date_order": date_order_used,
        "scanned_at": dt.datetime.now().isoformat(timespec="seconds"),
        "options": {
            "window_minutes": win_m,
            "sender_mode": s_mode,
            "text_priority": t_prio,
            "name_format": n_fmt,
            "date_from": str(date_from or "").strip(),
            "date_to": str(date_to or "").strip(),
            "include_docs": bool(include_docs),
        },
        "extract_info": extract_info,
        "messages": ser_messages,
        "items": items,
        "executed": False,
        "last_execution": None,
    }
    atomic_write_json(whatsapp_plan_path(proj_root), plan)
    return {
        "ok": True,
        "batch_id": plan["batch_id"],
        "source_dir": plan["source_dir"],
        "source_kind": source_kind,
        "chat_files_count": len(chat_files),
        "messages_count": len(ser_messages),
        "counts": summarize_whatsapp_plan(plan),
        "options": plan["options"],
    }


def recompute_whatsapp_plan(
    proj: dict[str, Any],
    *,
    window_minutes: int | None = None,
    sender_mode: str | None = None,
    text_priority: str | None = None,
    name_format: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    include_docs: bool | None = None,
) -> dict[str, Any]:
    """Re-apply matching rules and date filtering on the already parsed messages without re-scanning."""
    proj_root = _resolve_local_path(str(proj["root"]))
    pp = whatsapp_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        raise FileNotFoundError(tr("wa.no_plan_err"))

    opts = dict(plan.get("options") or {})
    if window_minutes is not None:
        opts["window_minutes"] = max(0, min(60, int(window_minutes)))
    if sender_mode in ("same_sender", "any_sender"):
        opts["sender_mode"] = sender_mode
    if text_priority in ("before_then_after", "after_then_before", "combine"):
        opts["text_priority"] = text_priority
    if name_format in ("date_name", "date_time_name", "date_sender_name"):
        opts["name_format"] = name_format
    if date_from is not None:
        opts["date_from"] = str(date_from).strip()
    if date_to is not None:
        opts["date_to"] = str(date_to).strip()
    if include_docs is not None:
        opts["include_docs"] = bool(include_docs)

    msgs_hydrated: list[dict[str, Any]] = []
    for m in plan.get("messages") or []:
        try:
            t_s = str(m.get("time_str") or "00:00")
            has_sec = bool(m.get("has_seconds") or t_s.count(":") == 2)
            fmt = "%Y-%m-%d %H:%M:%S" if t_s.count(":") == 2 else "%Y-%m-%d %H:%M"
            mdt = dt.datetime.strptime(f"{m['date_iso']} {t_s}", fmt)
        except Exception:
            continue
        msgs_hydrated.append({**m, "dt": mdt, "ts": mdt.timestamp(), "has_seconds": has_sec})

    _apply_matching_to_items(
        plan["items"],
        msgs_hydrated,
        window_minutes=int(opts.get("window_minutes") or 0),
        sender_mode=str(opts.get("sender_mode") or "same_sender"),
        text_priority=str(opts.get("text_priority") or "before_then_after"),
        name_format=str(opts.get("name_format") or "date_name"),
        date_from=str(opts.get("date_from") or ""),
        date_to=str(opts.get("date_to") or ""),
        include_docs=bool(opts.get("include_docs", True)),
        preserve_custom=True,
    )
    plan["options"] = opts
    atomic_write_json(pp, plan)
    return {
        "ok": True,
        "options": opts,
        "counts": summarize_whatsapp_plan(plan),
    }


def _validated_wa_last_execution(proj_root: Path, plan: dict[str, Any]) -> dict[str, Any] | None:
    if not isinstance(plan, dict) or not plan.get("executed"):
        return None
    raw = plan.get("last_execution")
    if not isinstance(raw, dict):
        return None
    out = dict(raw)
    rep_exists = False
    rep_rel = str(out.get("report_rel") or "").strip()
    rep_local = str(out.get("report_local") or "").strip()
    cands: list[Path] = []
    if rep_rel:
        try:
            cands.append(safe_resolve_within(proj_root, rep_rel, allow_base=False, reject_symlinks=True))
        except Exception:
            pass
    if rep_local:
        cands.append(_resolve_local_path(rep_local))
    for cand in cands:
        try:
            if cand.exists() and cand.is_file() and cand.stat().st_size > 0:
                rep_exists = True
                break
        except Exception:
            pass
    out["report_exists"] = rep_exists
    if not rep_exists:
        out["report_rel"] = ""
    out["executed"] = True
    t_dir_raw = str(out.get("target_dir") or plan.get("source_dir") or "").strip()
    t_dir = _resolve_local_path(t_dir_raw) if t_dir_raw else None
    zip_ready = bool(t_dir and t_dir.exists() and t_dir.is_dir())
    out["zip_ready"] = zip_ready
    if not rep_exists and not zip_ready:
        return None
    return out


def get_whatsapp_plan_page(
    proj_root: Path,
    *,
    page: int = 1,
    per_page: int = 50,
    status_filter: str = "all",
    media_filter: str = "",
    query: str = "",
    hide_excluded: bool = False,
) -> dict[str, Any]:
    pp = whatsapp_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        return {
            "has_plan": False,
            "items": [],
            "total": 0,
            "total_all": 0,
            "total_filtered": 0,
            "page": 1,
            "total_pages": 1,
            "counts": summarize_whatsapp_plan(None),
            "options": {},
            "source_dir": "",
            "source_kind": "",
            "last_execution": None,
            "has_undo": whatsapp_undo_path(proj_root).exists(),
        }

    items = plan["items"]
    counts = summarize_whatsapp_plan(plan)

    if status_filter == "out_of_range":
        filtered = [it for it in items if not it.get("in_date_range", True)]
    elif status_filter == "excluded":
        filtered = [it for it in items if it.get("in_date_range", True) and not it.get("include")]
    else:
        filtered = [it for it in items if it.get("in_date_range", True)]
        if hide_excluded:
            filtered = [it for it in filtered if it.get("include")]
        if status_filter == "matched":
            filtered = [it for it in filtered if it.get("match_status") == "matched"]
        elif status_filter == "date_only":
            filtered = [it for it in filtered if it.get("match_status") == "date_only"]
        elif status_filter == "unmatched":
            filtered = [it for it in filtered if it.get("match_status") == "unmatched"]
        elif status_filter == "duplicate":
            filtered = [it for it in filtered if it.get("is_duplicate")]

    if media_filter in ("image", "video", "document", "audio"):
        filtered = [it for it in filtered if it.get("media_kind") == media_filter]

    q_s = str(query or "").strip().lower()
    if q_s:
        filtered = [
            it for it in filtered
            if q_s in str(it.get("file_name") or "").lower()
            or q_s in str(it.get("new_name") or "").lower()
            or q_s in str(it.get("sender") or "").lower()
            or q_s in str(it.get("doc_date") or "").lower()
            or q_s in str(it.get("chosen_text") or "").lower()
            or q_s in str(it.get("text_before") or "").lower()
            or q_s in str(it.get("text_after") or "").lower()
        ]

    per_p = max(10, min(200, int(per_page or 50)))
    total_f = len(filtered)
    total_pages = max(1, (total_f + per_p - 1) // per_p)
    pg = max(1, min(total_pages, int(page or 1)))
    start = (pg - 1) * per_p

    return {
        "has_plan": len(items) > 0,
        "batch_id": plan.get("batch_id") or "",
        "batch_label": plan.get("batch_label") or "",
        "source_dir": plan.get("source_dir") or "",
        "source_input": plan.get("source_input") or "",
        "source_kind": plan.get("source_kind") or "folder",
        "chat_files": plan.get("chat_files") or [],
        "messages_count": len(plan.get("messages") or []),
        "scanned_at": plan.get("scanned_at") or "",
        "options": plan.get("options") or {},
        "counts": counts,
        "total": counts["in_range"],
        "total_all": counts["total_all"],
        "total_filtered": total_f,
        "page": pg,
        "per_page": per_p,
        "total_pages": total_pages,
        "items": filtered[start : start + per_p],
        "last_execution": _validated_wa_last_execution(proj_root, plan),
        "has_undo": bool(read_json_safe(whatsapp_undo_path(proj_root), {}).get("entries")),
    }


def _sanitize_user_new_name(raw_name: str, orig_ext: str) -> str:
    clean = _WIN_BAD_CHARS_RE.sub("_", str(raw_name or "").strip()).replace("/", "_").replace("\\", "_")
    clean = clean.strip(" .")
    if not clean:
        return ""
    p = Path(clean)
    if not p.suffix and orig_ext:
        clean = f"{clean}{orig_ext}"
    elif orig_ext and p.suffix.lower() != orig_ext.lower():
        clean = f"{p.stem}{orig_ext}"
    return clean[:180]


def update_whatsapp_plan_items(
    proj: dict[str, Any],
    updates: list[dict[str, Any]] | None = None,
    *,
    bulk_action: str = "",
) -> dict[str, Any]:
    proj_root = _resolve_local_path(str(proj["root"]))
    pp = whatsapp_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        raise FileNotFoundError(tr("wa.no_plan_err"))

    items = plan["items"]
    by_id = {str(it.get("id")): it for it in items}

    if bulk_action == "include_all":
        for it in items:
            if it.get("in_date_range", True):
                it["include"] = True
                it["manual_include"] = True
    elif bulk_action == "exclude_duplicates":
        for it in items:
            if it.get("is_duplicate"):
                it["include"] = False
                it["manual_include"] = False
    elif bulk_action == "include_matched_only":
        for it in items:
            val = bool(it.get("match_status") == "matched" and not it.get("is_duplicate") and it.get("in_date_range", True))
            it["include"] = val
            it["manual_include"] = val
    elif bulk_action == "exclude_date_only":
        for it in items:
            if it.get("match_status") != "matched":
                it["include"] = False
                it["manual_include"] = False
    elif bulk_action == "reset_names":
        opts = plan.get("options") or {}
        for it in items:
            it["custom_name"] = False
            it["ai_renamed"] = False
            it.pop("manual_include", None)
            base_nm = build_proposed_filename(
                doc_date=str(it.get("doc_date") or ""),
                msg_time=str(it.get("msg_time") or ""),
                sender=str(it.get("sender") or ""),
                chosen_text=str(it.get("chosen_text") or ""),
                orig_name=str(it.get("file_name") or ""),
                ext=str(it.get("ext") or ""),
                name_format=str(opts.get("name_format") or "date_name"),
            )
            it["base_new_name"] = base_nm
            it["new_name"] = base_nm
        _assign_unique_names(items)

    for upd in (updates or []):
        if not isinstance(upd, dict):
            continue
        it = by_id.get(str(upd.get("id") or ""))
        if not it:
            continue
        if "include" in upd:
            it["include"] = bool(upd["include"])
            it["manual_include"] = bool(upd["include"])
        if "doc_date" in upd:
            d_val = str(upd.get("doc_date") or "").strip()
            if re.match(r"^\d{4}-\d{2}-\d{2}$", d_val):
                it["doc_date"] = d_val
        if "new_name" in upd:
            clean_nn = _sanitize_user_new_name(str(upd["new_name"] or ""), str(it.get("ext") or ""))
            if clean_nn:
                it["new_name"] = clean_nn
                it["base_new_name"] = clean_nn
                it["custom_name"] = True
                it["ai_renamed"] = False
                it["include"] = True

    atomic_write_json(pp, plan)
    return {"ok": True, "counts": summarize_whatsapp_plan(plan)}


def _build_unique_disk_target(dest_dir: Path, desired_name: str, *, allow_existing_same_file: Path | None = None) -> Path:
    p = Path(desired_name)
    stem = p.stem or "whatsapp_file"
    ext = p.suffix
    cand = dest_dir / f"{stem}{ext}"
    if allow_existing_same_file is not None:
        try:
            if cand.resolve() == allow_existing_same_file.resolve():
                return cand
        except Exception:
            pass
    if not cand.exists():
        return cand
    for i in range(1, 1000):
        cand_i = dest_dir / f"{stem}_{i:02d}{ext}"
        if allow_existing_same_file is not None:
            try:
                if cand_i.resolve() == allow_existing_same_file.resolve():
                    return cand_i
            except Exception:
                pass
        if not cand_i.exists():
            return cand_i
    return dest_dir / f"{stem}_{uuid.uuid4().hex[:5]}{ext}"


def _write_whatsapp_excel_report(
    report_path: Path,
    plan: dict[str, Any],
    items: list[dict[str, Any]],
    *,
    link_base_dir: Path | None = None,
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "أرشيف الواتساب"
    ws.views.sheetView[0].rightToLeft = True

    hdr_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    matched_fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    date_only_fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")

    headers = [
        "#",
        "تاريخ الصورة/الملف",
        "الوقت",
        "المرسل",
        "النوع",
        "حالة الفصل",
        "مصدر التسمية",
        "الكلام قبل الصورة",
        "التعليق المباشر",
        "الكلام بعد الصورة",
        "الاسم الأصلي",
        "الاسم الجديد",
        "فتح الملف",
    ]
    ws.append(headers)
    for col_i in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_i)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    status_ar = {
        "matched": "متطابق (تاريخ + كلام)",
        "date_only": "تاريخ فقط (بدون كلام)",
        "unmatched": "غير مطابق",
    }
    kind_ar = {"image": "صورة", "video": "فيديو", "audio": "صوت", "document": "مستند"}
    src_ar = {
        "caption": "تعليق مباشر",
        "before": "الكلام قبل الصورة",
        "after": "الكلام بعد الصورة",
        "before_and_after": "قبل + بعد الصورة",
        "ai": "ذكاء اصطناعي",
        "manual": "تعديل يدوي",
        "none": "بدون كلام",
    }

    for row_idx, it in enumerate(items, 2):
        tsrc = "manual" if it.get("custom_name") else ("ai" if it.get("ai_renamed") else (it.get("text_source") or "none"))
        row_vals = [
            row_idx - 1,
            it.get("doc_date") or "",
            it.get("msg_time") or "",
            it.get("sender") or "",
            kind_ar.get(it.get("media_kind") or "document", "ملف"),
            status_ar.get(it.get("match_status") or "unmatched", "غير مطابق"),
            src_ar.get(tsrc, tsrc),
            it.get("text_before") or "",
            it.get("caption") or "",
            it.get("text_after") or "",
            it.get("file_name") or "",
            it.get("final_name") or it.get("new_name") or it.get("file_name") or "",
            "فتح الملف",
        ]
        ws.append(row_vals)
        fill = matched_fill if it.get("match_status") == "matched" else (date_only_fill if it.get("match_status") == "date_only" else None)
        for col_i in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_i)
            c.alignment = Alignment(horizontal="right" if col_i in (4, 8, 9, 10, 11, 12) else "center", vertical="center", wrap_text=True)
            if fill is not None:
                c.fill = fill

        final_path_s = str(it.get("final_path") or it.get("src_path") or "")
        link_cell = ws.cell(row=row_idx, column=13)
        if final_path_s:
            try:
                fp_obj = Path(final_path_s)
                if link_base_dir is not None:
                    try:
                        rel_lnk = fp_obj.resolve().relative_to(link_base_dir.resolve()).as_posix()
                        link_cell.hyperlink = rel_lnk
                    except Exception:
                        link_cell.hyperlink = fp_obj.as_uri()
                else:
                    link_cell.hyperlink = fp_obj.as_uri()
                link_cell.font = Font(name="Calibri", size=11, color="0563C1", underline="single")
            except Exception:
                pass

    widths = [6, 15, 10, 20, 10, 22, 18, 32, 26, 32, 30, 38, 13]
    for idx_w, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx_w)].width = w
    ws.freeze_panes = "A2"

    # Sheet 2: Parsed chat messages timeline
    msgs = plan.get("messages") or []
    if msgs:
        ws2 = wb.create_sheet("سجل المحادثة")
        ws2.views.sheetView[0].rightToLeft = True
        ws2.append(["#", "التاريخ", "الوقت", "المرسل", "الرسالة", "المرفقات"])
        for c_i in range(1, 7):
            c = ws2.cell(row=1, column=c_i)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center")
        for m_i, m in enumerate(msgs[:15000], 2):
            ws2.append([
                m_i - 1,
                m.get("date_iso") or "",
                m.get("time_str") or "",
                m.get("sender") or "",
                m.get("clean_text") or m.get("caption") or "",
                ", ".join(m.get("attachments") or []),
            ])
        for idx_w, w in enumerate([7, 14, 10, 22, 60, 34], 1):
            ws2.column_dimensions[get_column_letter(idx_w)].width = w
        ws2.freeze_panes = "A2"

    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(report_path)


def execute_whatsapp_archive_plan(
    proj: dict[str, Any],
    *,
    mode: str = "rename_in_place",
    dest_folder: str = "",
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """
    Execute the WhatsApp archive plan.
    Default `mode="rename_in_place"`: renames included files directly inside their current folder,
    with collision protection, zero deletion, undo journal, and Excel report.
    """
    proj_root = _resolve_local_path(str(proj["root"]))
    pp = whatsapp_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        raise FileNotFoundError(tr("wa.no_plan_err"))

    items = plan["items"]
    exec_mode = mode if mode in ("rename_in_place", "copy_by_date") else "rename_in_place"
    src_dir = _resolve_local_path(str(plan.get("source_dir") or proj_root))

    if exec_mode == "copy_by_date":
        target_root = _resolve_local_path(dest_folder) if str(dest_folder or "").strip() else (src_dir / "_WhatsApp_Sorted")
        target_root.mkdir(parents=True, exist_ok=True)
    else:
        target_root = src_dir

    last_emit = [0.0]
    included_items = [x for x in items if x.get("include")]
    if not included_items:
        raise ValueError(tr("wa.no_items_to_exec_err"))
    total_inc = max(1, len(included_items))
    _emit_wa_prog(progress_cb, last_emit, {
        "phase": "executing", "processed": 0, "total": total_inc, "current_file": "",
    }, force=True)

    renamed_count = 0
    copied_count = 0
    already_ok_count = 0
    skipped_count = 0
    done_inc = 0
    errors: list[str] = []
    undo_entries: list[dict[str, str]] = []

    for it in items:
        if stop_event is not None and stop_event.is_set():
            if undo_entries:
                existing_undo = read_json_safe(whatsapp_undo_path(proj_root), {})
                merged = list(existing_undo.get("entries") or []) if isinstance(existing_undo, dict) else []
                atomic_write_json(whatsapp_undo_path(proj_root), {
                    "created_at": dt.datetime.now().isoformat(timespec="seconds"),
                    "source_dir": str(src_dir),
                    "entries": merged + undo_entries,
                })
                atomic_write_json(pp, plan)
            raise WhatsappStoppedError(tr("jobs.stopped_by_user"))
        if not it.get("include"):
            skipped_count += 1
            continue
        done_inc += 1
        _emit_wa_prog(progress_cb, last_emit, {
            "phase": "executing", "processed": done_inc, "total": total_inc,
            "current_file": str(it.get("file_name") or ""),
        })
        cur_path_s = str(it.get("final_path") or it.get("src_path") or "")
        cur_p = _resolve_local_path(cur_path_s)
        if not cur_p.exists() or not cur_p.is_file():
            errors.append(f"{it.get('file_name', '')}: missing")
            continue

        desired = str(it.get("new_name") or it.get("file_name") or cur_p.name)
        desired = _sanitize_user_new_name(desired, cur_p.suffix) or cur_p.name

        if exec_mode == "copy_by_date":
            day_dir = target_root / (str(it.get("doc_date") or "") or "Undated")
            day_dir.mkdir(parents=True, exist_ok=True)
            dest_p = _build_unique_disk_target(day_dir, desired)
            shutil.copy2(cur_p, dest_p)
            it["final_name"] = dest_p.name
            it["final_path"] = str(dest_p)
            it["executed"] = True
            copied_count += 1
        else:
            dest_dir = cur_p.parent
            if desired == cur_p.name:
                it["final_name"] = cur_p.name
                it["final_path"] = str(cur_p)
                it["executed"] = True
                already_ok_count += 1
                continue
            dest_p = _build_unique_disk_target(dest_dir, desired, allow_existing_same_file=cur_p)
            old_str = str(cur_p)
            # Handle case-only rename safely on case-insensitive filesystems (Windows)
            if cur_p.name.lower() == dest_p.name.lower() and cur_p.name != dest_p.name:
                tmp_p = dest_dir / f".__tmp_wa_{uuid.uuid4().hex[:6]}{cur_p.suffix}"
                cur_p.rename(tmp_p)
                tmp_p.rename(dest_p)
            else:
                cur_p.rename(dest_p)
            undo_entries.append({
                "item_id": str(it.get("id") or ""),
                "old_path": old_str,
                "new_path": str(dest_p),
                "old_name": cur_p.name,
                "new_name": dest_p.name,
            })
            it["src_path"] = str(dest_p)
            it["final_name"] = dest_p.name
            it["final_path"] = str(dest_p)
            it["new_name"] = dest_p.name
            try:
                it["src_rel"] = dest_p.relative_to(src_dir).as_posix()
            except Exception:
                pass
            it["executed"] = True
            renamed_count += 1

    if stop_event is not None and stop_event.is_set():
        if undo_entries:
            existing_undo = read_json_safe(whatsapp_undo_path(proj_root), {})
            merged = list(existing_undo.get("entries") or []) if isinstance(existing_undo, dict) else []
            atomic_write_json(whatsapp_undo_path(proj_root), {
                "created_at": dt.datetime.now().isoformat(timespec="seconds"),
                "source_dir": str(src_dir),
                "entries": merged + undo_entries,
            })
            atomic_write_json(pp, plan)
        raise WhatsappStoppedError(tr("jobs.stopped_by_user"))

    if (renamed_count + copied_count + already_ok_count) == 0:
        raise ValueError(tr("wa.no_items_to_exec_err"))

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    rep_name = f"WhatsApp_Archive_Report_{stamp}.xlsx"
    reports_dir = proj_root / "_Reports"
    reports_dir.mkdir(parents=True, exist_ok=True)
    rep_proj = reports_dir / rep_name
    try:
        _write_whatsapp_excel_report(rep_proj, plan, items, link_base_dir=target_root)
        if not rep_proj.exists() or rep_proj.stat().st_size == 0:
            raise RuntimeError("Failed to generate WhatsApp report")
    except Exception:
        rep_proj.unlink(missing_ok=True)
        raise

    # Also save a local Excel report right inside the source/target folder for convenience
    rep_local = target_root / rep_name
    try:
        if rep_local.resolve() != rep_proj.resolve():
            _write_whatsapp_excel_report(rep_local, plan, items, link_base_dir=target_root)
    except Exception:
        rep_local = rep_proj

    if undo_entries:
        atomic_write_json(whatsapp_undo_path(proj_root), {
            "created_at": dt.datetime.now().isoformat(timespec="seconds"),
            "source_dir": str(src_dir),
            "entries": undo_entries,
        })

    exec_sum = {
        "mode": exec_mode,
        "renamed": renamed_count,
        "copied": copied_count,
        "skipped": skipped_count,
        "errors": errors,
        "report_rel": rep_proj.relative_to(proj_root).as_posix(),
        "report_local": str(rep_local),
        "target_dir": str(target_root),
        "executed_at": dt.datetime.now().isoformat(timespec="seconds"),
        "cleared": True,
        "executed": True,
        "report_exists": True,
        "zip_ready": True,
    }
    # Clear pending plan items once execution finishes cleanly so the UI resets to a clean screen
    plan["items"] = []
    plan["messages"] = []
    plan["executed"] = True
    plan["executed_at"] = exec_sum["executed_at"]
    plan["last_execution"] = exec_sum
    atomic_write_json(pp, plan)
    return {"ok": True, **exec_sum, "counts": summarize_whatsapp_plan(plan)}


def clear_whatsapp_plan(proj: dict[str, Any]) -> dict[str, Any]:
    """Reset the WhatsApp review table and summary banner in the project workspace."""
    proj_root = _resolve_local_path(str(proj["root"]))
    pp = whatsapp_plan_path(proj_root)
    if pp.exists():
        plan = read_json_safe(pp, {})
        if isinstance(plan, dict):
            plan["items"] = []
            plan["messages"] = []
            plan["executed"] = False
            plan["last_execution"] = None
            atomic_write_json(pp, plan)
    return {"ok": True, "counts": summarize_whatsapp_plan(None)}


def undo_last_whatsapp_rename(proj: dict[str, Any]) -> dict[str, Any]:
    """Restore filenames renamed during the last `rename_in_place` execution."""
    proj_root = _resolve_local_path(str(proj["root"]))
    up = whatsapp_undo_path(proj_root)
    undo_data = read_json_safe(up, {})
    entries = list(undo_data.get("entries") or []) if isinstance(undo_data, dict) else []
    if not entries:
        raise FileNotFoundError(tr("wa.no_undo_err"))

    pp = whatsapp_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    items = plan.get("items") or [] if isinstance(plan, dict) else []
    by_id = {str(it.get("id")): it for it in items}

    restored = 0
    errors: list[str] = []
    failed_entries: list[dict[str, Any]] = []
    for ent in reversed(entries):
        new_p = _resolve_local_path(str(ent.get("new_path") or ""))
        old_p = _resolve_local_path(str(ent.get("old_path") or ""))
        if not new_p.exists():
            errors.append(f"{new_p.name}: missing")
            failed_entries.append(ent)
            continue
        target_p = old_p if not old_p.exists() else _build_unique_disk_target(old_p.parent, old_p.name, allow_existing_same_file=new_p)
        try:
            if new_p.name.lower() == target_p.name.lower() and new_p.name != target_p.name:
                tmp_p = new_p.parent / f".__tmp_wau_{uuid.uuid4().hex[:6]}{new_p.suffix}"
                new_p.rename(tmp_p)
                tmp_p.rename(target_p)
            else:
                new_p.rename(target_p)
        except OSError as exc:
            errors.append(f"{new_p.name}: {exc}")
            failed_entries.append(ent)
            continue
        it = by_id.get(str(ent.get("item_id") or ""))
        if it:
            it["src_path"] = str(target_p)
            it["final_path"] = str(target_p)
            it["file_name"] = target_p.name
            it["final_name"] = target_p.name
            it["executed"] = False
            try:
                src_dir = _resolve_local_path(str(plan.get("source_dir") or proj_root))
                it["src_rel"] = target_p.relative_to(src_dir).as_posix()
            except Exception:
                pass
        restored += 1

    if failed_entries:
        atomic_write_json(up, {
            "entries": list(reversed(failed_entries)),
            "partial_restore_at": dt.datetime.now().isoformat(timespec="seconds"),
        })
    else:
        atomic_write_json(up, {"entries": [], "restored_at": dt.datetime.now().isoformat(timespec="seconds")})
    if isinstance(plan, dict):
        old_ex = plan.get("last_execution") or {}
        if isinstance(old_ex, dict):
            rep_rel = str(old_ex.get("report_rel") or "").strip()
            if rep_rel:
                try:
                    safe_resolve_within(proj_root, rep_rel, allow_base=False, reject_symlinks=True).unlink(missing_ok=True)
                except Exception:
                    pass
            rep_local_raw = str(old_ex.get("report_local") or "").strip()
            if rep_local_raw:
                try:
                    _resolve_local_path(rep_local_raw).unlink(missing_ok=True)
                except Exception:
                    pass
        plan["executed"] = False
        plan["last_execution"] = None
        atomic_write_json(pp, plan)
    return {"ok": True, "restored": restored, "errors": errors, "counts": summarize_whatsapp_plan(plan)}


def _wa_ai_cache_key(it: dict[str, Any]) -> str:
    fh = str(it.get("sha256") or "")
    if not fh:
        return ""
    ctx_sig = ai_cache.stable_sig(
        it.get("text_before") or "",
        it.get("caption") or "",
        it.get("text_after") or "",
        it.get("doc_date") or "",
    )
    return f"{fh}:{ctx_sig}"


def _apply_wa_ai_meta(it: dict[str, Any], meta_ai: dict[str, Any], *, name_format: str) -> bool:
    if not isinstance(meta_ai, dict):
        return False
    subj_ai = str(meta_ai.get("subject_ar") or meta_ai.get("subject") or "").strip()
    d_ai = str(meta_ai.get("doc_date") or "").strip()
    if not it.get("doc_date") and re.match(r"^\d{4}-\d{2}-\d{2}$", d_ai):
        it["doc_date"] = d_ai
    if not subj_ai:
        return False
    it["chosen_text"] = subj_ai
    it["text_source"] = "ai"
    if it.get("doc_date"):
        it["match_status"] = "matched"
    base_nm = build_proposed_filename(
        doc_date=str(it.get("doc_date") or ""),
        msg_time=str(it.get("msg_time") or ""),
        sender=str(it.get("sender") or ""),
        chosen_text=subj_ai,
        orig_name=str(it.get("file_name") or ""),
        ext=str(it.get("ext") or ""),
        name_format=name_format or "date_name",
    )
    it["base_new_name"] = base_nm
    it["new_name"] = base_nm
    it["custom_name"] = False
    it["ai_renamed"] = True
    return True


def ai_rename_whatsapp_items(
    proj: dict[str, Any],
    *,
    item_ids: list[str] | None = None,
    scope: str = "date_only",
    limit: int = 25,
    engine: str = "gemini",
    ollama_url: str = "",
    ollama_model: str = "",
    gemini_api_key: str = "",
    analyzer_fn: Callable[[str, Path, list[Path], str], dict[str, Any]] | None = None,
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
    force_refresh: bool = False,
) -> dict[str, Any]:
    """Use AI vision/text to generate clean names for WhatsApp media items lacking chat captions (cached by SHA-256 + context)."""
    proj_root = _resolve_local_path(str(proj["root"]))
    pp = whatsapp_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("items"), list):
        raise FileNotFoundError(tr("wa.no_plan_err"))

    items = plan["items"]
    if item_ids:
        wanted = {str(x) for x in item_ids if str(x).strip()}
        targets = [it for it in items if str(it.get("id")) in wanted]
    elif scope == "all":
        targets = [it for it in items if (force_refresh or not it.get("ai_renamed"))][: max(1, int(limit))]
    else:
        targets = [
            it for it in items
            if it.get("match_status") in ("date_only", "unmatched")
            and (force_refresh or not it.get("ai_renamed"))
            and not it.get("custom_name")
        ][: max(1, int(limit))]

    if not targets:
        return {"ok": True, "processed": 0, "updated": 0, "total_targets": 0, "errors": [], "counts": summarize_whatsapp_plan(plan)}

    tmp_dir = proj_root / "_System_Agent" / "_wa_ai_tmp"
    shutil.rmtree(tmp_dir, ignore_errors=True)
    tmp_dir.mkdir(parents=True, exist_ok=True)
    opts = plan.get("options") or {}
    wa_cache_dir = ai_cache.project_ai_cache_dir(proj_root)

    last_emit = [0.0]
    total_targets = len(targets)
    _emit_wa_prog(progress_cb, last_emit, {
        "phase": "ai_renaming", "processed": 0, "total": total_targets, "updated": 0, "current_file": "",
    }, force=True)

    processed = 0
    updated = 0
    cache_hits = 0
    stopped_flag = False
    errors: list[str] = []

    try:
        for it in targets:
            if stop_event is not None and stop_event.is_set():
                stopped_flag = True
                break
            fp = _resolve_local_path(str(it.get("src_path") or ""))
            _emit_wa_prog(progress_cb, last_emit, {
                "phase": "ai_renaming", "processed": processed, "total": total_targets,
                "updated": updated, "current_file": fp.name or str(it.get("file_name") or ""),
            }, force=True)
            if not fp.exists() or not fp.is_file():
                errors.append(f"{it.get('file_name', '')}: missing")
                continue

            if not it.get("sha256"):
                try:
                    it["sha256"] = sha256_of_file(fp)
                except OSError:
                    pass
            ckey = _wa_ai_cache_key(it)
            if ckey and not force_refresh:
                crec = ai_cache.get_cached(wa_cache_dir, "whatsapp_ai", ckey)
                if crec and _apply_wa_ai_meta(it, crec, name_format=str(opts.get("name_format") or "date_name")):
                    it["ai_from_cache"] = True
                    processed += 1
                    updated += 1
                    cache_hits += 1
                    _emit_wa_prog(progress_cb, last_emit, {
                        "phase": "ai_renaming", "processed": processed, "total": total_targets,
                        "updated": updated, "current_file": fp.name,
                    }, force=True)
                    continue

            images: list[Path] = []
            snippet = ""
            if fp.suffix.lower() in {".jpg", ".jpeg", ".jfif", ".png", ".webp", ".bmp"}:
                images = [fp]
            elif fp.suffix.lower() == ".pdf" and sw is not None:
                try:
                    if hasattr(sw, "prepare_visual_pages"):
                        pdf_dir = tmp_dir / f"pdf_{uuid.uuid4().hex[:6]}"
                        images, snippet, _pc = sw.prepare_visual_pages(fp, pdf_dir)
                except Exception:
                    pass
            elif fp.suffix.lower() in {".txt", ".csv"}:
                try:
                    snippet = fp.read_text(encoding="utf-8", errors="replace")[:1800]
                except Exception:
                    pass

            if not images and not snippet:
                errors.append(f"{fp.name}: unsupported for AI")
                processed += 1
                continue

            ctx_txt = " | ".join(x for x in (it.get("text_before"), it.get("caption"), it.get("text_after")) if x)
            prompt = (
                "أنت مهندس مراقبة مستندات وموقع إنشائي. انظر إلى هذه الصورة أو المستند المرسل في واتساب، "
                "واقترح اسماً وصفياً موجزاً جداً وواضحاً يعبر عما يظهر فيها (مثل: صب_أعمدة_الدور_الثاني أو عزل_الأساسات أو اعتماد_مواد_سباكة).\n"
                f"التاريخ الحالي للملف: {it.get('doc_date') or ''}\n"
                + (f"سياق المحادثة القريب: {ctx_txt}\n" if ctx_txt else "")
                + (f"مقتطف نصي: {snippet[:1000]}\n" if snippet else "")
                + 'أرجع JSON فقط: {"doc_date": "YYYY-MM-DD أو فارغ", "subject_ar": "عنوان وصفي موجز من 3 إلى 6 كلمات"}'
            )

            try:
                if analyzer_fn is not None:
                    meta_ai = analyzer_fn(prompt, tmp_dir, images, snippet) or {}
                elif str(engine or "").lower() == "ollama" and sw is not None:
                    meta_ai = sw.analyze_with_ollama(
                        prompt, images,
                        ollama_url or "http://127.0.0.1:11434/api/chat",
                        ollama_model or "qwen3-vl:8b",
                    ) or {}
                elif sw is not None:
                    meta_ai = sw.analyze_with_gemini(
                        prompt, tmp_dir, sw.load_saved_gemini_model(),
                        images=images, api_key=gemini_api_key,
                    ) or {}
                else:
                    meta_ai = {}
            except Exception as exc:
                errors.append(f"{fp.name}: {exc}")
                processed += 1
                continue

            processed += 1
            if not isinstance(meta_ai, dict):
                continue

            if _apply_wa_ai_meta(it, meta_ai, name_format=str(opts.get("name_format") or "date_name")):
                it["ai_from_cache"] = False
                updated += 1
                if ckey and not meta_ai.get("error"):
                    ai_cache.set_cached(wa_cache_dir, "whatsapp_ai", ckey, {
                        "subject_ar": str(meta_ai.get("subject_ar") or meta_ai.get("subject") or "").strip(),
                        "doc_date": str(meta_ai.get("doc_date") or "").strip(),
                    })
            _emit_wa_prog(progress_cb, last_emit, {
                "phase": "ai_renaming", "processed": processed, "total": total_targets,
                "updated": updated, "current_file": fp.name,
            }, force=True)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    _assign_unique_names(items)
    atomic_write_json(pp, plan)
    return {
        "ok": True,
        "processed": processed,
        "updated": updated,
        "cache_hits": cache_hits,
        "total_targets": len(targets),
        "stopped": stopped_flag,
        "errors": errors,
        "counts": summarize_whatsapp_plan(plan),
    }


def build_whatsapp_export_zip(proj: dict[str, Any]) -> Path:
    """Create a downloadable ZIP containing the executed archive contents and Excel report."""
    proj_root = _resolve_local_path(str(proj["root"]))
    pp = whatsapp_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not plan.get("executed") or not isinstance(plan.get("last_execution"), dict):
        raise ValueError(tr("wa.exec_required_zip_err"))
    if not plan.get("source_dir"):
        raise FileNotFoundError(tr("wa.no_plan_err"))

    src_dir = _resolve_local_path(str(plan["source_dir"]))
    if not src_dir.exists() or not src_dir.is_dir():
        raise FileNotFoundError(tr("server.folder_not_found", path=str(src_dir)))

    zip_dir = proj_root / "_Reports"
    zip_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_zip = zip_dir / f"WhatsApp_Archive_Bundle_{stamp}.zip"

    last_ex = plan.get("last_execution") or {}
    target_dir = _resolve_local_path(str(last_ex.get("target_dir") or src_dir))
    if not target_dir.exists():
        target_dir = src_dir

    written_files = 0
    try:
        with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=5) as zf:
            for fp in sorted(target_dir.rglob("*")):
                if not fp.is_file():
                    continue
                rel = fp.relative_to(target_dir)
                if any(part in SKIP_DIR_NAMES for part in rel.parts[:-1]):
                    continue
                if fp.name.startswith("~$") or fp.suffix.lower() == ".zip":
                    continue
                zf.write(fp, arcname=rel.as_posix())
                written_files += 1
            rep_rel = str(last_ex.get("report_rel") or "")
            if rep_rel:
                rep_p = safe_resolve_within(proj_root, rep_rel, allow_base=False, reject_symlinks=True)
                if rep_p.exists() and rep_p.is_file() and rep_p.name not in {p.name for p in target_dir.glob("*.xlsx")}:
                    zf.write(rep_p, arcname=rep_p.name)
        if written_files == 0:
            raise ValueError(tr("wa.no_items_to_exec_err"))
    except Exception:
        out_zip.unlink(missing_ok=True)
        raise

    return out_zip
