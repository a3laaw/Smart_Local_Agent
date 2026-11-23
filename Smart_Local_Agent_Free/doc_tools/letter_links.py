#!/usr/bin/env python3
"""
letter_links.py — Multi-party correspondence linking engine for Smart_Local_Agent_Free.

Discovers, suggests, links, unlinks, and writes cross-party correspondence chains
(Consultant <-> Contractor/User <-> Subcontractor <-> Owner) even when a reply or
follow-up letter discusses the same topic (e.g. CCTV cameras) WITHOUT citing the
original letter's reference number.

Persists confirmed links cleanly in `<proj_root>/3_Project_Logs/_letter_links.json`
and writes the linked letter summary ("الكتاب المرتبط") directly in front of each
letter row inside the project's Excel log(s).
"""
from __future__ import annotations

import copy
import datetime as dt
import re
import secrets
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

from i18n_ar import tr
from relevance import is_file_locked, normalize_ar
from security_utils import (
    atomic_write_json,
    now_iso,
    read_json_safe,
    resolve_local_or_unc,
    sha256_of_file,
)
import project_log as plog

_LINKS_FILENAME = "_letter_links.json"
_RELATED_CELL_FILL = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
_REMARK_TAG_RE = re.compile(r"\s*\|?\s*\[(?:الكتب المرتبطة|Related Letters?)\s*:[^\]]*\]\s*", re.IGNORECASE)

# ── Arabic + English Correspondence Topic Stopwords & Synonym Concepts ────────
_LETTER_STOPWORDS: set[str] = {
    # Correspondence boilerplate (AR normalized via normalize_ar -> ه / ي / ا)
    "كتاب", "كتب", "خطاب", "مراسله", "مراسلات", "رساله", "رد", "ردا", "الرد", "العوده",
    "بشان", "بشأن", "بخصوص", "موضوع", "الموضوع", "اشاره", "الاشاره", "اشارتنا", "كتابكم", "كتابنا",
    "رقم", "تاريخ", "بتاريخ", "مشروع", "المشروع", "بالمشروع", "الموقع", "بالموقع", "موقع",
    "اعمال", "الاعمال", "عمل", "ملاحظات", "ملاحظه", "طلب", "الطلب", "تقديم", "تقرير",
    "خاص", "خاصه", "الخاص", "الخاصه", "متعلق", "متعلقه", "شان", "امر", "المعروض",
    "الساده", "ساده", "شركه", "مؤسسه", "مكتب", "محترمين", "المحترمين", "تحيه", "طيبه", "وبعد",
    "الاستشاري", "استشاري", "السور", "مقاول", "المقاول", "الباطن", "مقاولي", "مقاولين",
    "الوزاره", "وزاره", "الصحه", "الاشغال", "المالك", "وكيل", "الشؤون", "الهندسيه",
    "تم", "عدم", "لعدم", "علي", "على", "في", "من", "الي", "الى", "عن", "مع", "او", "و", "ان",
    "كل", "جميع", "بعض", "حول", "خلال", "لمده", "مده", "يوم", "ايام", "اسبوع",
    "ماده", "مواد", "الماده", "المواد", "نظام", "انظمه", "النظام", "الانظمه",
    "اجراء", "اتخاذ", "اللازم", "لازم", "يرجي", "نرجو", "افاده", "الافاده", "توضيح",
    "توريد", "تركيب", "تنفيذ", "تجهيز", "اعتماد", "عرض", "حاله", "شروط", "متطلبات",
    "مركز", "صباح", "السالم", "الجنوبي", "صحي", "المركز", "الصحي", "مبني", "المبني",
    # English correspondence boilerplate
    "ref", "no", "re", "fw", "fwd", "letter", "letters", "correspondence", "subject",
    "regarding", "about", "reply", "response", "follow", "followup", "reminder", "notice",
    "project", "site", "work", "works", "supply", "install", "installation", "submission",
    "request", "approval", "notes", "comments", "subcontractor", "contractor", "consultant",
    "owner", "client", "ministry", "moh", "soor", "shc", "eng", "dear", "sir", "the",
    "and", "for", "with", "from", "to", "of", "in", "on", "at", "by", "as", "per", "dated",
    "date", "mr", "ms", "ltd", "co", "company", "engineering", "center", "health",
}

# Canonical technical/contracting domain concept clusters (bridging Arabic & English!)
_CONCEPT_CLUSTERS: dict[str, tuple[str, ...]] = {
    "_c_cctv": (
        "كاميرا", "كاميرات", "مراقبه", "تلفزيونيه", "cctv", "camera", "cameras",
        "surveillance", "nvr", "dvr", "iptv", "security camera",
    ),
    "_c_safety": (
        "سلامه", "امن", "امان", "وقايه", "خوذه", "سقالات", "سقاله", "safety",
        "hse", "ppe", "hazard", "scaffolding",
    ),
    "_c_fire": (
        "حريق", "اطفاء", "مكافحه", "رشاشات", "حساسات", "دخان", "fire", "sprinkler",
        "alarm", "fighting", "hydrant", "kff", "nfpa",
    ),
    "_c_hvac": (
        "تكييف", "تبريد", "تهويه", "مكيفات", "مكيف", "دكت", "شيلر", "hvac",
        "chiller", "ahu", "fcu", "duct", "vrf", "air conditioning",
    ),
    "_c_elec": (
        "كهرباء", "كهربائيه", "كابلات", "كابل", "لوحات", "لوحه", "محول", "مولد",
        "اضاءه", "اناره", "electrical", "cable", "cables", "switchgear", "panel",
        "transformer", "generator", "lighting", "ups", "earthing",
    ),
    "_c_plumb": (
        "سباكه", "صحي", "صحيه", "مضخات", "مضخه", "خزان", "خزانات", "مواسير",
        "مجاري", "صرف", "تغذيه", "مياه", "plumbing", "pump", "pumps", "drainage",
        "water tank", "piping", "sanitary",
    ),
    "_c_elev": (
        "مصعد", "مصاعد", "elevator", "elevators", "lift", "lifts", "escalator",
    ),
    "_c_earth": (
        "دفان", "ردم", "حفر", "تربه", "اساسات", "خوازيق", "backfill", "excavation",
        "soil", "compaction", "subbase",
    ),
    "_c_concrete": (
        "خرسانه", "صب", "حديد", "تسليح", "قواعد", "اعمدة", "سقف", "concrete",
        "rebar", "steel", "casting", "slab",
    ),
    "_c_payment": (
        "دفعه", "مستخلص", "صرف", "فاتوره", "مطالبه", "مالي", "ماليه", "ipc",
        "payment", "invoice", "certificate", "claim",
    ),
    "_c_replace": (
        "استبدال", "تبديل", "بديل", "اسبدال", "تغيير", "replace", "replacement", "substitute",
    ),
    "_c_staff": (
        "مهندس", "مراقب", "فني", "كادر", "طاقم", "مدير", "جهاز", "موظف", "تعيين",
        "engineer", "supervisor", "foreman", "staff", "manpower", "cv",
    ),
    "_c_schedule": (
        "جدول", "زمني", "برنامج", "تمديد", "تاخير", "مهله", "تقدم", "سير",
        "schedule", "baseline", "cpm", "extension", "eot", "delay", "progress",
    ),
    "_c_xray": (
        "اشعه", "ترصيص", "رصاص", "وقايه اشعاعيه", "xray", "x-ray", "radiation", "lead",
    ),
    "_c_dental": (
        "اسنان", "عيادات", "عياده", "dental", "chair",
    ),
    "_c_bms": (
        "bms", "تحكم الي", "مباني ذكيه", "building management", "automation",
    ),
}

_SPECIFIC_HIGH_CONCEPTS = {
    "_c_cctv", "_c_elev", "_c_xray", "_c_dental", "_c_bms", "_c_earth", "_c_hvac",
}

_DISTINCT_TRADES: dict[str, tuple[str, ...]] = {
    "civil": ("مدني", "مدنيه", "المدنيه", "خرسانه", "انشائي", "انشائيه"),
    "elec": ("كهرباء", "كهربائيه", "الكهرباء", "الكهربائيه", "كابلات"),
    "mech": ("ميكانيك", "ميكانيكيه", "الميكانيك"),
    "plumb": ("صحي", "صحيه", "سباكه"),
    "arch": ("معماري", "معماريه"),
    "hvac": ("تكييف", "تبريد"),
}
_TRADE_BY_WORD: dict[str, str] = {}
for _tk, _tws in _DISTINCT_TRADES.items():
    for _tw in _tws:
        _TRADE_BY_WORD[normalize_ar(_tw)] = _tk

_TOKEN_SYNONYM_MAP: dict[str, set[str]] = {}
for _cid, _cwords in _CONCEPT_CLUSTERS.items():
    for _cw in _cwords:
        _cn = normalize_ar(_cw)
        if " " not in _cn:
            _TOKEN_SYNONYM_MAP.setdefault(_cn, set()).add(_cid)
            if _cn.startswith("ال") and len(_cn) > 3:
                _TOKEN_SYNONYM_MAP.setdefault(_cn[2:], set()).add(_cid)


# ── Storage Path & JSON Helpers ──────────────────────────────────────────────
def links_path_for(proj_root: Path | str) -> Path:
    return plog.logs_dir_for_project(proj_root) / _LINKS_FILENAME


def _load_links_store(proj_root: Path | str) -> dict[str, Any]:
    raw = read_json_safe(links_path_for(proj_root), {})
    if not isinstance(raw, dict):
        raw = {}
    threads = raw.get("threads")
    if not isinstance(threads, list):
        threads = []
    ignored = raw.get("ignored_pairs")
    if not isinstance(ignored, list):
        ignored = []
    return {
        "version": int(raw.get("version") or 1),
        "threads": [dict(t) for t in threads if isinstance(t, dict) and isinstance(t.get("members"), list)],
        "ignored_pairs": [list(p[:2]) for p in ignored if isinstance(p, (list, tuple)) and len(p) >= 2],
    }


def _save_links_store(proj_root: Path | str, store: dict[str, Any]) -> None:
    atomic_write_json(links_path_for(proj_root), store)


# ── Party & Member Key Normalization ─────────────────────────────────────────
def _infer_party_role_from_hints(party_role: str = "", sheet_name: str = "", party_name: str = "") -> str:
    pr = str(party_role or "").strip().lower()
    if pr in ("consultant", "subcontractor", "owner", "contractor"):
        return pr
    combo = normalize_ar(f"{sheet_name} {party_name}")
    if any(k in combo for k in ("مقاولين الباطن", "مقاول الباطن", "مقاولي الباطن", "باطن", "subcontractor")):
        return "subcontractor"
    if any(k in combo for k in ("السور", "استشاري", "الاستشاري", "مكتب", "consultant", "soor", "shc")):
        return "consultant"
    if any(k in combo for k in ("وزاره", "الصحه", "الاشغال", "المالك", "وكيل", "moh", "mpw", "pahw", "mew", "ministry", "owner", "client")):
        return "owner"
    if any(k in combo for k in ("مقاول", "المقاول", "contractor", "cont")):
        return "contractor"
    return ""


def party_label_for(
    section: str,
    party_role: str = "",
    sheet_name: str = "",
    party_name: str = "",
) -> tuple[str, str]:
    """Return `(canonical_party_role, localized_role_label)` for a correspondence item."""
    role = _infer_party_role_from_hints(party_role, sheet_name, party_name)
    is_in = str(section or "").strip().lower() == "incoming"
    key_map = {
        ("outgoing", "consultant"): "letters.role_out_cons",
        ("incoming", "consultant"): "letters.role_in_cons",
        ("outgoing", "subcontractor"): "letters.role_out_sub",
        ("incoming", "subcontractor"): "letters.role_in_sub",
        ("outgoing", "owner"): "letters.role_out_owner",
        ("incoming", "owner"): "letters.role_in_owner",
        ("outgoing", "contractor"): "letters.role_out_contr",
        ("incoming", "contractor"): "letters.role_in_contr",
    }
    dir_k = "incoming" if is_in else "outgoing"
    tkey = key_map.get((dir_k, role), "letters.role_in_gen" if is_in else "letters.role_out_gen")
    return role, tr(tkey)


def make_member_key(
    ref_no: str = "",
    *,
    sheet: str = "",
    section: str = "",
    row: int | None = None,
) -> str:
    """Stable key identifying a letter across sheets & logs."""
    rk = plog.normalize_ref_key(ref_no or "")
    sn = normalize_ar(sheet or "")
    sc = str(section or "").strip().lower()
    if rk:
        if sn:
            return f"{sn}::{rk}"
        if sc:
            return f"{sc}::{rk}"
        return rk
    if sn and row:
        return f"{sn}::ROW:{int(row)}"
    if row:
        return f"ROW:{int(row)}"
    return ""


def _canonical_pair(key_a: str, key_b: str) -> tuple[str, str]:
    a, b = str(key_a or "").strip(), str(key_b or "").strip()
    return (a, b) if a <= b else (b, a)


def _member_matches_target(m: dict[str, Any], key: str, ref_key: str, sheet: str = "") -> bool:
    mk = str(m.get("member_key") or "")
    mrk = str(m.get("ref_key") or plog.normalize_ref_key(m.get("ref_no") or ""))
    ms = normalize_ar(str(m.get("sheet") or ""))
    sn = normalize_ar(sheet or "")
    if key and mk and mk == key:
        return True
    if ref_key and mrk and mrk == ref_key:
        if sn and ms and sn != ms:
            return False
        return True
    return False


# ── NLP Topic Extraction & Similarity Scoring ─────────────────────────────────
def _arabic_stem_variants(word: str) -> set[str]:
    """Generate clean normalized Arabic/English stem variants for topic matching."""
    w = normalize_ar(word)
    if not w:
        return set()
    out: set[str] = {w}
    # Strip definite article `ال`
    if w.startswith("ال") and len(w) > 3:
        w = w[2:]
        out.add(w)
    # Strip common prefixed prepositions/conjunctions before `ال` or `لل`
    for pfx in ("وال", "بال", "فال", "كال", "لال"):
        if w.startswith(pfx) and len(w) > len(pfx) + 2:
            w = w[len(pfx):]
            out.add(w)
            break
    if w.startswith("لل") and len(w) > 4:
        w = w[2:]
        out.add(w)
    # Also try stripping leading single `و`/`ب`/`ل` on longer Arabic words (`بمخالفة`, `للمشروع`, `ومراقبة`)
    if w[0] in ("و", "ب", "ل", "ف") and len(w) >= 5 and re.match(r"[\u0600-\u06FF]", w):
        cand = w[1:]
        if cand.startswith("ال") and len(cand) > 3:
            cand = cand[2:]
        out.add(cand)

    # Add sound-feminine plural <-> singular forms without turning nouns like `مراقبة` into job title `مراقب`
    more: set[str] = set()
    for s in list(out):
        if not re.match(r"[\u0600-\u06FF]", s):
            continue
        if len(s) >= 5 and s.endswith("ات"):
            stem_at = s[:-2]
            more.add(stem_at + "ه")
            more.add(stem_at + "ا")
            if len(stem_at) >= 4:
                more.add(stem_at)
        elif len(s) >= 5 and s.endswith(("ين", "ون")):
            more.add(s[:-2])
        elif len(s) >= 6 and s.endswith("يه"):
            more.add(s[:-2])
    out |= {x for x in more if len(x) >= 3}
    return out


def extract_topic_signature(text: str, ref_no: str = "") -> dict[str, Any]:
    """
    Build a structured topic signature (stems, concepts, specific ordinal numbers)
    from a letter's subject + optional body/notes text.
    """
    raw = str(text or "")
    norm = normalize_ar(raw)
    # Strip explicit reference codes like `S.H.C/2026-05-018` or `ج/2026/24` before topic word analysis
    cleaned = re.sub(r"\b[a-z0-9._-]{1,10}[/-]\d{2,4}(?:[/-][a-z0-9._-]+)?\b", " ", norm, flags=re.IGNORECASE)
    cleaned = re.sub(r"[\u0600-\u06FF]/[0-9]{2,4}(?:/[0-9]+)?", " ", cleaned)

    raw_words = re.findall(r"[a-z0-9]+|[\u0600-\u06FF]+", cleaned)
    surface_tokens: list[str] = []
    stems: set[str] = set()
    concepts: set[str] = set()
    ordinals: set[str] = set()
    codes: set[str] = set()
    trades: set[str] = set()

    # Check multi-word concept phrases first
    for cid, cwords in _CONCEPT_CLUSTERS.items():
        for cw in cwords:
            if " " in cw:
                cwn = normalize_ar(cw)
                if cwn and cwn in norm:
                    concepts.add(cid)

    for rw in raw_words:
        if rw.isdigit():
            val = int(rw)
            if 2018 <= val <= 2039:
                continue  # ignore calendar years as topic discriminators
            if 1 <= val <= 999:
                ordinals.add(str(val))
            continue
        if len(rw) <= 1:
            continue
        variants = _arabic_stem_variants(rw)
        for v in variants:
            if v in _TRADE_BY_WORD:
                trades.add(_TRADE_BY_WORD[v])
        # Skip boilerplate stopwords if any primary variant is a stopword
        base_variants = {rw, rw[2:] if rw.startswith("ال") and len(rw) > 3 else rw}
        if base_variants & _LETTER_STOPWORDS or variants & _LETTER_STOPWORDS:
            # Still check concept clusters ONLY if it wasn't a generic action word
            for v in variants:
                if v in _TOKEN_SYNONYM_MAP:
                    concepts.update(_TOKEN_SYNONYM_MAP[v])
            continue

        if re.match(r"^[a-z]{2,8}$", rw) and rw in ("cctv", "hvac", "bms", "ipc", "cpm", "ups", "vrf", "kff", "hse", "ppe", "xray"):
            codes.add(rw)

        surface_tokens.append(rw)
        stems.update(variants)
        for v in variants:
            if v in _TOKEN_SYNONYM_MAP:
                concepts.update(_TOKEN_SYNONYM_MAP[v])

    return {
        "norm": norm,
        "surface": surface_tokens,
        "stems": stems,
        "concepts": concepts,
        "ordinals": ordinals,
        "codes": codes,
        "trades": trades,
        "ref_no": str(ref_no or "").strip(),
    }


def _check_explicit_ref_mention(letter: dict[str, Any], other_ref: str) -> bool:
    """True when `letter`'s subject, notes, or extracted text cites `other_ref`."""
    ref_raw = str(other_ref or "").strip()
    if len(ref_raw) < 3 or ref_raw.isdigit():
        return False
    # Require either >= 4 chars or punctuation/letters so single integers never false-trigger
    if len(ref_raw) < 4 and not re.search(r"[/\-\\.]", ref_raw):
        return False
    hay = normalize_ar(" ".join(str(letter.get(k) or "") for k in ("subject", "notes", "text", "summary")))
    if not hay:
        return False
    rn = normalize_ar(ref_raw)
    if rn and re.search(rf"(?<![\w/]){re.escape(rn)}(?![\w/])", hay):
        return True
    # Also check separator-agnostic normalized form (`ج 2026 24` vs `ج/2026/24`)
    rk = normalize_ar(ref_raw.replace("/", " ").replace("-", " "))
    hay_sp = normalize_ar(hay.replace("/", " ").replace("-", " "))
    if rk and len(rk) >= 5 and re.search(rf"(?<!\w){re.escape(rk)}(?!\w)", hay_sp):
        return True
    return False


def _parse_iso_date(val: str) -> dt.date | None:
    s = str(val or "").strip()[:10]
    if not s:
        return None
    try:
        return dt.date.fromisoformat(s)
    except ValueError:
        return None


def score_letter_pair(
    a: dict[str, Any],
    b: dict[str, Any],
    *,
    sig_a: dict[str, Any] | None = None,
    sig_b: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Score topic & correspondence similarity `[0.0, 0.99]` between two letters `a` and `b`.
    Works BOTH when reference numbers are cited AND when NO reference number is cited!
    """
    ref_a = str(a.get("ref_no") or "").strip()
    ref_b = str(b.get("ref_no") or "").strip()

    # Never self-match identical row/key
    ka = str(a.get("member_key") or make_member_key(ref_a, sheet=str(a.get("sheet") or ""), section=str(a.get("section") or ""), row=a.get("row")))
    kb = str(b.get("member_key") or make_member_key(ref_b, sheet=str(b.get("sheet") or ""), section=str(b.get("section") or ""), row=b.get("row")))
    if ka and kb and ka == kb:
        return {"score": 0.0, "matched_keywords": [], "reason": ""}

    # 1. Explicit reference citation beats everything
    if (ref_b and _check_explicit_ref_mention(a, ref_b)) or (ref_a and _check_explicit_ref_mention(b, ref_a)):
        return {
            "score": 0.98,
            "matched_keywords": [ref_a or ref_b],
            "reason": tr("letters.reason_explicit"),
            "explicit_ref": True,
        }

    sa = sig_a or extract_topic_signature(f"{a.get('subject') or ''} {a.get('notes') or ''} {a.get('text') or ''}", ref_a)
    sb = sig_b or extract_topic_signature(f"{b.get('subject') or ''} {b.get('notes') or ''} {b.get('text') or ''}", ref_b)

    stems_b: set[str] = sb["stems"]
    concepts_shared = sa["concepts"] & sb["concepts"]
    codes_shared = sa["codes"] & sb["codes"]
    ord_a: set[str] = sa["ordinals"]
    ord_b: set[str] = sb["ordinals"]
    ord_shared = ord_a & ord_b

    # Exact stem overlap (group stems belonging to the same surface word so a word doesn't double-count)
    matched_surface: list[str] = []
    exact_hits = 0
    fuzzy_hits = 0
    used_b_words: set[str] = set()

    surf_a = sa["surface"]
    surf_b = sb["surface"]
    for wa in surf_a:
        va = _arabic_stem_variants(wa)
        found_exact = False
        for wb in surf_b:
            vb = _arabic_stem_variants(wb)
            if va & vb:
                if wb not in used_b_words:
                    used_b_words.add(wb)
                    exact_hits += 1
                    matched_surface.append(wa)
                found_exact = True
                break
        if found_exact:
            continue
        # Fuzzy match for minor spelling typos (e.g. `اسبدال` vs `استبدال`), guarding short/unrelated stems
        wa_core = wa[2:] if wa.startswith("ال") and len(wa) > 3 else wa
        if len(wa_core) >= 6:
            for wb in surf_b:
                if wb in used_b_words:
                    continue
                wb_core = wb[2:] if wb.startswith("ال") and len(wb) > 3 else wb
                if len(wb_core) < 6 or wa_core[0] != wb_core[0]:
                    continue
                if SequenceMatcher(None, wa_core, wb_core).ratio() >= 0.80:
                    used_b_words.add(wb)
                    fuzzy_hits += 1
                    matched_surface.append(wa)
                    break

    total_surface_matches = exact_hits + (0.85 * fuzzy_hits)
    has_specific_concept = bool(concepts_shared & _SPECIFIC_HIGH_CONCEPTS) or bool(codes_shared)

    # Require solid domain evidence:
    # - Either >= 2 overlapping meaningful words/stems (e.g. `كاميرات` + `مراقبه`, `امن` + `سلامه`), OR
    # - >= 1 shared surface word + >= 1 matching ordinal (e.g. `الدفعه` + `02`), OR
    # - A specific cross-lingual concept (e.g. Arabic `كاميرات` <-> English `CCTV`)!
    if (
        total_surface_matches < 1.7
        and not (exact_hits >= 1 and ord_shared)
        and not has_specific_concept
    ):
        return {"score": 0.0, "matched_keywords": [], "reason": ""}

    # Prevent false matches on conflicting ordinal numbers (`الدفعة 01` vs `الدفعة 02`)
    ord_adj = 0.0
    if ord_a and ord_b:
        if ord_shared:
            ord_adj = 0.22
            matched_surface.extend(sorted(ord_shared))
        else:
            ord_adj = -0.40

    # Prevent false matches across conflicting engineering trades (`الكهرباء` vs `المدنية`)
    trade_adj = 0.0
    trades_a: set[str] = sa.get("trades") or set()
    trades_b: set[str] = sb.get("trades") or set()
    if trades_a and trades_b and not (trades_a & trades_b):
        trade_adj = -0.50

    denom = max(1.0, min(len(surf_a), len(surf_b)) * 0.55 + (len(surf_a) + len(surf_b)) * 0.225)
    sim_ratio = min(1.0, (total_surface_matches + 0.55 * len(concepts_shared) + 0.5 * len(codes_shared)) / denom)

    score = 0.28 + (0.46 * sim_ratio) + min(0.18, 0.045 * total_surface_matches)
    if concepts_shared:
        if has_specific_concept and exact_hits == 0:
            score += 0.16
        else:
            score += 0.07 if exact_hits == 0 else 0.04
    if codes_shared:
        score += 0.10
    score += ord_adj + trade_adj

    # Cross-direction / cross-party boost (Consultant <-> My Outgoing <-> Subcontractor)
    sec_a = str(a.get("section") or "").lower()
    sec_b = str(b.get("section") or "").lower()
    pr_a = str(a.get("party_role") or "").lower()
    pr_b = str(b.get("party_role") or "").lower()
    if (sec_a and sec_b and sec_a != sec_b) or (pr_a and pr_b and pr_a != pr_b):
        score += 0.09
    elif sec_a and sec_a == sec_b and pr_a == pr_b and total_surface_matches < 3:
        # Same party & same direction requires higher overlap unless clearly same thread
        score -= 0.05

    # Date proximity nudge
    da = _parse_iso_date(str(a.get("doc_date") or a.get("date") or ""))
    db = _parse_iso_date(str(b.get("doc_date") or b.get("date") or ""))
    if da and db:
        days = abs((da - db).days)
        if days <= 21:
            score += 0.06
        elif days <= 60:
            score += 0.03
        elif days > 180:
            score -= 0.06

    score = round(max(0.0, min(0.97, score)), 3)
    if score < 0.42:
        return {"score": 0.0, "matched_keywords": [], "reason": ""}

    # Deduplicate matched keywords nicely (prefer cleaner stem without `و`/`ب`/`ل` prefix)
    kw_clean: list[str] = []
    for kw in matched_surface:
        disp = kw
        if re.match(r"^[\u0600-\u06FF]", kw) and kw[0] in ("و", "ب", "ل", "ف") and len(kw) >= 5:
            cand = kw[1:]
            if cand.startswith("ال") and len(cand) > 3:
                disp = cand
            elif cand in stems_b or (cand.startswith("ل") and len(cand) > 4):
                disp = cand[1:] if cand.startswith("ل") else cand
        if disp not in kw_clean:
            kw_clean.append(disp)
    if not kw_clean and concepts_shared:
        for cid in sorted(concepts_shared):
            words = _CONCEPT_CLUSTERS.get(cid, ())
            if words:
                kw_clean.append(words[0])

    return {
        "score": score,
        "matched_keywords": kw_clean[:6],
        "reason": tr("letters.reason_topic"),
        "explicit_ref": False,
    }


# ── Collect Correspondence Rows & Thread Formatting ──────────────────────────
def _format_thread_summary_for_member(
    thread: dict[str, Any],
    self_member: dict[str, Any],
    *,
    include_subject: bool = True,
    joiner: str = " | ",
) -> str:
    """Format the 'الكتاب المرتبط' cell text shown in front of `self_member`."""
    self_key = str(self_member.get("member_key") or "")
    parts: list[str] = []
    for m in thread.get("members") or []:
        mk = str(m.get("member_key") or "")
        if mk and self_key and mk == self_key:
            continue
        _role, plabel = party_label_for(
            str(m.get("section") or ""),
            str(m.get("party_role") or ""),
            str(m.get("sheet") or ""),
            str(m.get("party") or ""),
        )
        ref = str(m.get("ref_no") or "").strip()
        subj = str(m.get("subject") or "").strip()
        if len(subj) > 36:
            subj = subj[:34].rstrip() + "…"
        if ref and subj and include_subject:
            parts.append(f"{plabel}: {ref} ({subj})")
        elif ref:
            parts.append(f"{plabel}: {ref}")
        elif subj:
            parts.append(f"{plabel}: {subj}")
    return joiner.join(parts)


def _index_threads_by_key(store: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    by_member_key: dict[str, dict[str, Any]] = {}
    by_ref_key: dict[str, dict[str, Any]] = {}
    for th in store.get("threads") or []:
        for m in th.get("members") or []:
            mk = str(m.get("member_key") or make_member_key(str(m.get("ref_no") or ""), sheet=str(m.get("sheet") or ""), section=str(m.get("section") or ""), row=m.get("row")))
            rk = str(m.get("ref_key") or plog.normalize_ref_key(m.get("ref_no") or ""))
            if mk:
                by_member_key[mk] = th
            if rk:
                by_ref_key[rk] = th
    return by_member_key, by_ref_key


def collect_project_letters(
    proj: dict[str, Any],
    *,
    device_id: str = "",
) -> list[dict[str, Any]]:
    """
    Collect all correspondence letters (`incoming` / `outgoing`) across the project's
    uploaded & directly linked Excel logs, enriched with party labels and current threads.
    """
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))
    logs = plog.list_project_logs(proj_root, device_id=device_id)
    store = _load_links_store(proj_root)
    by_mk, by_rk = _index_threads_by_key(store)

    letters: list[dict[str, Any]] = []
    for lg in logs:
        if lg.get("missing"):
            continue
        lp = Path(lg["path"])
        try:
            lmap = plog.load_or_create_logmap(lp, device_id=device_id, proj_root=proj_root)
            wb = openpyxl.load_workbook(lp, read_only=True, data_only=True)
        except Exception:
            continue
        try:
            for sc in lmap.get("sheets") or []:
                sname = str(sc.get("sheet") or "")
                if not sc.get("enabled") or sc.get("layout") == "skip" or sname not in wb.sheetnames:
                    continue
                routes = sc.get("routes") or []
                sec = str(routes[0].get("section") or "general") if routes else "general"
                if sec not in ("incoming", "outgoing"):
                    continue
                route_pr = str(routes[0].get("party_role") or "") if routes else ""
                cols = sc.get("columns") or {}
                rev_groups = sc.get("rev_groups") or []
                layout = sc.get("layout", "flat")
                data_start = int(sc.get("data_start_row") or 2)
                ws = wb[sname]

                def _g(rtup: tuple, role: str) -> Any:
                    ci = cols.get(role, {}).get("col")
                    if not ci or ci > len(rtup):
                        return None
                    return rtup[ci - 1]

                for r_idx, rtup in enumerate(
                    ws.iter_rows(min_row=data_start, max_col=min(ws.max_column or 40, 45), values_only=True),
                    start=data_start,
                ):
                    ref_v = _g(rtup, "ref_no")
                    subj_v = _g(rtup, "subject")
                    ref_s = str(ref_v).strip() if ref_v is not None else ""
                    subj_s = str(subj_v).strip() if subj_v is not None else ""
                    if not ref_s and not subj_s:
                        continue
                    if normalize_ar(ref_s) in ("total", "الاجمالي", "المجموع") or normalize_ar(subj_s) in ("total", "الاجمالي", "المجموع"):
                        continue

                    party_s = str(_g(rtup, "party") or "").strip()
                    notes_s = str(_g(rtup, "notes") or "").strip()
                    rel_col_s = str(_g(rtup, "related_letters") or "").strip()
                    date_v = _g(rtup, "doc_date")
                    reply_v = _g(rtup, "reply_date")
                    stat_s = str(_g(rtup, "status") or "").strip()

                    if layout == "multi_rev" and rev_groups:
                        latest_g = rev_groups[0]
                        for g in rev_groups:
                            sc_i = int(g.get("submit_col") or 0)
                            rc_i = int(g.get("reply_col") or 0)
                            ac_i = int(g.get("action_col") or 0)
                            sv = rtup[sc_i - 1] if sc_i and sc_i <= len(rtup) else None
                            rv = rtup[rc_i - 1] if rc_i and rc_i <= len(rtup) else None
                            av = rtup[ac_i - 1] if ac_i and ac_i <= len(rtup) else None
                            if any(x is not None and str(x).strip() != "" for x in (sv, rv, av)):
                                latest_g = g
                        sc_i = int(latest_g.get("submit_col") or 0)
                        rc_i = int(latest_g.get("reply_col") or 0)
                        ac_i = int(latest_g.get("action_col") or 0)
                        date_v = rtup[sc_i - 1] if sc_i and sc_i <= len(rtup) else None
                        reply_v = rtup[rc_i - 1] if rc_i and rc_i <= len(rtup) else None
                        stat_s = str(rtup[ac_i - 1] or "").strip() if ac_i and ac_i <= len(rtup) else ""

                    def _d_str(dv: Any) -> str:
                        if isinstance(dv, dt.datetime):
                            return dv.date().isoformat()
                        if isinstance(dv, dt.date):
                            return dv.isoformat()
                        if isinstance(dv, str) and dv.strip():
                            return dv.strip()[:10]
                        return ""

                    date_s = _d_str(date_v)
                    reply_s = _d_str(reply_v)
                    pr_canon, p_lbl = party_label_for(sec, route_pr, sname, party_s)
                    rk = plog.normalize_ref_key(ref_s)
                    mk = make_member_key(ref_s, sheet=sname, section=sec, row=r_idx)

                    th = by_mk.get(mk) or (by_rk.get(rk) if rk else None)
                    linked_list: list[dict[str, Any]] = []
                    related_txt = rel_col_s
                    if th:
                        for om in th.get("members") or []:
                            omk = str(om.get("member_key") or "")
                            if omk and omk == mk:
                                continue
                            _orole, olbl = party_label_for(
                                str(om.get("section") or ""),
                                str(om.get("party_role") or ""),
                                str(om.get("sheet") or ""),
                                str(om.get("party") or ""),
                            )
                            linked_list.append({
                                "ref_no": str(om.get("ref_no") or ""),
                                "sheet": str(om.get("sheet") or ""),
                                "section": str(om.get("section") or ""),
                                "party_role": str(om.get("party_role") or ""),
                                "party_label": olbl,
                                "subject": str(om.get("subject") or ""),
                                "doc_date": str(om.get("doc_date") or ""),
                                "thread_id": th.get("id", ""),
                            })
                        if not related_txt:
                            related_txt = _format_thread_summary_for_member(th, {"member_key": mk})

                    letters.append({
                        "member_key": mk,
                        "ref_key": rk,
                        "ref_no": ref_s,
                        "sheet": sname,
                        "log_name": lp.name,
                        "row": r_idx,
                        "section": sec,
                        "party": party_s,
                        "party_role": pr_canon,
                        "party_label": p_lbl,
                        "subject": subj_s,
                        "doc_date": date_s,
                        "date": date_s,
                        "reply_date": reply_s,
                        "status": stat_s,
                        "notes": notes_s,
                        "related_letters": related_txt,
                        "linked_letters": linked_list,
                        "thread_id": th.get("id", "") if th else "",
                    })
        finally:
            wb.close()

    return letters


def _ignored_set_from_store(store: dict[str, Any]) -> set[tuple[str, str]]:
    out: set[tuple[str, str]] = set()
    for p in store.get("ignored_pairs") or []:
        if len(p) >= 2:
            out.add(_canonical_pair(str(p[0]), str(p[1])))
    return out


def suggest_letter_links(
    proj: dict[str, Any],
    *,
    ref_no: str = "",
    sheet: str = "",
    section: str = "",
    subject: str = "",
    doc_date: str = "",
    text: str = "",
    party_role: str = "",
    party_name: str = "",
    limit: int = 8,
    min_score: float = 0.48,
    device_id: str = "",
) -> list[dict[str, Any]]:
    """Suggest candidate related letters in the project for a given or newly scanned letter."""
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))
    letters = collect_project_letters(proj, device_id=device_id)
    if not letters:
        return []

    store = _load_links_store(proj_root)
    ignored = _ignored_set_from_store(store)

    # Locate existing entry if it already lives in the log
    rk = plog.normalize_ref_key(ref_no or "")
    target_mk = make_member_key(ref_no, sheet=sheet, section=section)
    existing = next((x for x in letters if _member_matches_target(x, target_mk, rk, sheet)), None)
    if existing:
        pr_canon = existing["party_role"]
        eff_sec = existing["section"]
        eff_subj = subject or existing["subject"]
        eff_date = doc_date or existing["doc_date"]
        eff_notes = existing["notes"]
        target_mk = existing["member_key"]
        already_linked_keys = {make_member_key(m["ref_no"], sheet=m.get("sheet", ""), section=m.get("section", "")) for m in existing.get("linked_letters") or []}
    else:
        pr_canon, _ = party_label_for(section, party_role, sheet, party_name)
        eff_sec = section
        eff_subj = subject
        eff_date = doc_date
        eff_notes = ""
        already_linked_keys = set()

    query_obj = {
        "member_key": target_mk,
        "ref_no": ref_no,
        "sheet": sheet,
        "section": eff_sec,
        "party_role": pr_canon,
        "subject": eff_subj,
        "doc_date": eff_date,
        "notes": eff_notes,
        "text": text,
    }
    sig_q = extract_topic_signature(f"{eff_subj} {eff_notes} {text}", ref_no)

    scored: list[dict[str, Any]] = []
    for cand in letters:
        cmk = cand["member_key"]
        if target_mk and cmk == target_mk:
            continue
        if cmk in already_linked_keys:
            continue
        if target_mk and _canonical_pair(target_mk, cmk) in ignored:
            continue
        res = score_letter_pair(query_obj, cand, sig_a=sig_q)
        if res["score"] >= min_score:
            scored.append({
                "ref_no": cand["ref_no"],
                "sheet": cand["sheet"],
                "log_name": cand["log_name"],
                "row": cand["row"],
                "section": cand["section"],
                "party_role": cand["party_role"],
                "party_label": cand["party_label"],
                "subject": cand["subject"],
                "doc_date": cand["doc_date"],
                "member_key": cmk,
                "score": res["score"],
                "pct": int(round(res["score"] * 100)),
                "matched_keywords": res["matched_keywords"],
                "reason": res["reason"],
                "explicit_ref": bool(res.get("explicit_ref")),
            })

    scored.sort(key=lambda x: (x["score"], x["doc_date"]), reverse=True)
    return scored[:limit]


def discover_unlinked_matches(
    proj: dict[str, Any],
    *,
    min_score: float = 0.58,
    cross_party_only: bool = True,
    limit: int = 80,
    device_id: str = "",
) -> list[dict[str, Any]]:
    """Find high-confidence unlinked correspondence pairs across parties in the project."""
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))
    letters = collect_project_letters(proj, device_id=device_id)
    if len(letters) < 2:
        return []

    store = _load_links_store(proj_root)
    ignored = _ignored_set_from_store(store)
    sigs = [extract_topic_signature(f"{x.get('subject') or ''} {x.get('notes') or ''}", x.get("ref_no") or "") for x in letters]

    pairs: list[dict[str, Any]] = []
    n = len(letters)
    for i in range(n):
        a = letters[i]
        for j in range(i + 1, n):
            b = letters[j]
            # Already in the exact same thread -> skip
            if a["thread_id"] and a["thread_id"] == b["thread_id"]:
                continue
            if _canonical_pair(a["member_key"], b["member_key"]) in ignored:
                continue
            if cross_party_only and a["sheet"] == b["sheet"] and a["section"] == b["section"] and a["party_role"] == b["party_role"]:
                # Allow same-sheet auto-pair ONLY when one explicitly mentions the other's ref_no
                if not (_check_explicit_ref_mention(a, b["ref_no"]) or _check_explicit_ref_mention(b, a["ref_no"])):
                    continue
            res = score_letter_pair(a, b, sig_a=sigs[i], sig_b=sigs[j])
            if res["score"] >= min_score:
                pairs.append({
                    "pair_id": f"{a['member_key']}__{b['member_key']}",
                    "score": res["score"],
                    "pct": int(round(res["score"] * 100)),
                    "matched_keywords": res["matched_keywords"],
                    "reason": res["reason"],
                    "explicit_ref": bool(res.get("explicit_ref")),
                    "a": {
                        "member_key": a["member_key"],
                        "ref_no": a["ref_no"],
                        "sheet": a["sheet"],
                        "section": a["section"],
                        "party_role": a["party_role"],
                        "party_label": a["party_label"],
                        "subject": a["subject"],
                        "doc_date": a["doc_date"],
                        "log_name": a["log_name"],
                        "row": a["row"],
                    },
                    "b": {
                        "member_key": b["member_key"],
                        "ref_no": b["ref_no"],
                        "sheet": b["sheet"],
                        "section": b["section"],
                        "party_role": b["party_role"],
                        "party_label": b["party_label"],
                        "subject": b["subject"],
                        "doc_date": b["doc_date"],
                        "log_name": b["log_name"],
                        "row": b["row"],
                    },
                })

    pairs.sort(key=lambda x: x["score"], reverse=True)
    return pairs[:limit]


# ── Mutating Link / Unlink Operations & Excel Synchronization ─────────────────
def _normalize_item_to_member(
    raw_item: dict[str, Any] | str,
    letters_index: dict[str, dict[str, Any]],
    ref_index: dict[str, dict[str, Any]],
) -> dict[str, Any] | None:
    if isinstance(raw_item, str):
        ref = raw_item.strip()
        if not ref:
            return None
        rk = plog.normalize_ref_key(ref)
        found = ref_index.get(rk)
        if found:
            return {
                "member_key": found["member_key"],
                "ref_key": found["ref_key"],
                "ref_no": found["ref_no"],
                "sheet": found["sheet"],
                "section": found["section"],
                "party_role": found["party_role"],
                "party_label": found["party_label"],
                "subject": found["subject"],
                "doc_date": found["doc_date"],
                "log_name": found["log_name"],
                "row": found["row"],
            }
        return {
            "member_key": make_member_key(ref),
            "ref_key": rk,
            "ref_no": ref,
            "sheet": "",
            "section": "outgoing",
            "party_role": "",
            "party_label": tr("letters.role_out_gen"),
            "subject": "",
            "doc_date": "",
        }

    if not isinstance(raw_item, dict):
        return None
    ref = str(raw_item.get("ref_no") or "").strip()
    sheet = str(raw_item.get("sheet") or "").strip()
    section = str(raw_item.get("section") or "").strip()
    row = raw_item.get("row")
    rk = plog.normalize_ref_key(ref)
    mk = str(raw_item.get("member_key") or make_member_key(ref, sheet=sheet, section=section, row=row))
    found = letters_index.get(mk) or (ref_index.get(rk) if rk and not sheet else None)
    if found:
        return {
            "member_key": found["member_key"],
            "ref_key": found["ref_key"],
            "ref_no": found["ref_no"],
            "sheet": found["sheet"],
            "section": found["section"],
            "party_role": found["party_role"],
            "party_label": found["party_label"],
            "subject": str(raw_item.get("subject") or found["subject"] or ""),
            "doc_date": str(raw_item.get("doc_date") or found["doc_date"] or ""),
            "log_name": found["log_name"],
            "row": found["row"],
        }
    pr, plbl = party_label_for(
        section or "outgoing",
        str(raw_item.get("party_role") or ""),
        sheet,
        str(raw_item.get("party") or raw_item.get("party_name") or ""),
    )
    if not ref and not raw_item.get("subject"):
        return None
    return {
        "member_key": mk,
        "ref_key": rk,
        "ref_no": ref,
        "sheet": sheet,
        "section": section or "outgoing",
        "party_role": pr,
        "party_label": plbl,
        "subject": str(raw_item.get("subject") or ""),
        "doc_date": str(raw_item.get("doc_date") or raw_item.get("date") or ""),
        "log_name": str(raw_item.get("log_name") or ""),
        "row": row,
    }


def _merge_members_into_store(
    store: dict[str, Any],
    members: list[dict[str, Any]],
    *,
    topic: str = "",
) -> dict[str, Any]:
    """Merge a list of correspondence members into ONE unified thread inside `store`."""
    clean_members: list[dict[str, Any]] = []
    seen_keys: set[str] = set()
    for m in members:
        mk = str(m.get("member_key") or "")
        if not mk or mk in seen_keys:
            continue
        seen_keys.add(mk)
        clean_members.append(m)

    # Clear ignored status among newly confirmed links
    ig = _ignored_set_from_store(store)
    for i in range(len(clean_members)):
        for j in range(i + 1, len(clean_members)):
            ig.discard(_canonical_pair(clean_members[i]["member_key"], clean_members[j]["member_key"]))
    store["ignored_pairs"] = [list(p) for p in sorted(ig)]

    # Find any existing threads containing any of these members -> merge them all!
    matching_threads: list[dict[str, Any]] = []
    remaining_threads: list[dict[str, Any]] = []
    for th in store.get("threads") or []:
        tkeys = {str(x.get("member_key") or "") for x in th.get("members") or []}
        trks = {str(x.get("ref_key") or "") for x in th.get("members") or [] if x.get("ref_key")}
        new_rks = {str(m.get("ref_key") or "") for m in clean_members if m.get("ref_key") and not m.get("sheet")}
        if (tkeys & seen_keys) or (trks & new_rks):
            matching_threads.append(th)
        else:
            remaining_threads.append(th)

    merged_by_key: dict[str, dict[str, Any]] = {}
    chosen_topic = topic.strip()
    created_at = now_iso()
    chosen_id = f"LTH-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}-{secrets.token_hex(2).upper()}"

    if matching_threads:
        chosen_id = str(matching_threads[0].get("id") or chosen_id)
        created_at = str(matching_threads[0].get("created_at") or created_at)
        for th in matching_threads:
            if not chosen_topic and th.get("topic"):
                chosen_topic = str(th["topic"])
            for om in th.get("members") or []:
                omk = str(om.get("member_key") or "")
                if omk:
                    merged_by_key[omk] = dict(om)

    for m in clean_members:
        mk = m["member_key"]
        if mk in merged_by_key:
            # Refresh metadata without losing existing filled fields
            cur = dict(merged_by_key[mk])
            for k, v in m.items():
                if v not in (None, ""):
                    cur[k] = v
            merged_by_key[mk] = cur
        else:
            merged_by_key[mk] = m

    if not chosen_topic:
        chosen_topic = next((str(m.get("subject") or "")[:60] for m in merged_by_key.values() if m.get("subject")), "")

    new_thread = {
        "id": chosen_id,
        "topic": chosen_topic,
        "created_at": created_at,
        "updated_at": now_iso(),
        "members": list(merged_by_key.values()),
    }
    remaining_threads.append(new_thread)
    store["threads"] = remaining_threads
    return new_thread


def _ensure_related_letters_col(ws, sheet_spec: dict[str, Any]) -> tuple[int, str]:
    """
    Determine which Excel column receives 'الكتاب المرتبط' on a correspondence sheet:
      1) Existing `related_letters` column in `sheet_spec["columns"]`,
      2) Existing `notes` column (fallback when sheet already has Remarks/Notes),
      3) Otherwise, cleanly append `الكتاب المرتبط / Related Letter` immediately after
         the rightmost header column (preserving merged header formatting like `AC4:AC5`).
    Returns `(col_idx, mode)` where `mode` is `'related_letters'` or `'notes'`.
    """
    cols = sheet_spec.setdefault("columns", {})
    if "related_letters" in cols and cols["related_letters"].get("col"):
        return int(cols["related_letters"]["col"]), "related_letters"
    if "notes" in cols and cols["notes"].get("col"):
        return int(cols["notes"]["col"]), "notes"

    hdr_row = int(sheet_spec.get("header_row") or 1)
    sub_row = sheet_spec.get("subheader_row")
    max_scan = max(ws.max_column or 1, 1)
    last_hdr_col = 0
    for c in range(1, max_scan + 1):
        v1 = ws.cell(hdr_row, c).value
        v2 = ws.cell(int(sub_row), c).value if sub_row else None
        if (v1 is not None and str(v1).strip() != "") or (v2 is not None and str(v2).strip() != ""):
            last_hdr_col = c
    if last_hdr_col < 1:
        last_hdr_col = max_scan

    new_col = last_hdr_col + 1
    new_letter = get_column_letter(new_col)
    hdr_label = f"{tr('letters.excel_col_hdr')} / Related Letter"

    src_hdr = ws.cell(hdr_row, last_hdr_col)
    dst_hdr = ws.cell(hdr_row, new_col)
    dst_hdr.value = hdr_label
    if src_hdr.has_style:
        dst_hdr.font = copy.copy(src_hdr.font)
        dst_hdr.fill = copy.copy(src_hdr.fill)
        dst_hdr.border = copy.copy(src_hdr.border)
        dst_hdr.alignment = copy.copy(src_hdr.alignment)

    if sub_row:
        sr = int(sub_row)
        src_sub = ws.cell(sr, last_hdr_col)
        dst_sub = ws.cell(sr, new_col)
        if src_sub.has_style:
            dst_sub.font = copy.copy(src_sub.font)
            dst_sub.fill = copy.copy(src_sub.fill)
            dst_sub.border = copy.copy(src_sub.border)
            dst_sub.alignment = copy.copy(src_sub.alignment)
        # Merge vertically if previous status column merged header..subheader (e.g. AC4:AC5)
        was_merged = any(
            mr.min_col <= last_hdr_col <= mr.max_col and mr.min_row == hdr_row and mr.max_row == sr
            for mr in ws.merged_cells.ranges
        )
        if was_merged:
            ws.merge_cells(start_row=hdr_row, start_column=new_col, end_row=sr, end_column=new_col)

    ws.column_dimensions[new_letter].width = 44
    cols["related_letters"] = {
        "col": new_col,
        "letter": new_letter,
        "header": hdr_label,
    }
    return new_col, "related_letters"


def _parse_iso_date_str(val: Any) -> str:
    if val is None or val == "":
        return ""
    if isinstance(val, dt.datetime):
        return val.date().isoformat()
    if isinstance(val, dt.date):
        return val.isoformat()
    s = str(val).strip()
    if not s:
        return ""
    m_iso = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m_iso:
        try:
            return dt.date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))).isoformat()
        except ValueError:
            return ""
    m_dmy = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})", s)
    if m_dmy:
        try:
            return dt.date(int(m_dmy.group(3)), int(m_dmy.group(2)), int(m_dmy.group(1))).isoformat()
        except ValueError:
            return ""
    return ""


def _sync_incoming_reply_status(
    ws,
    sc: dict[str, Any],
    row: int,
    mk: str,
    th: dict[str, Any] | None,
    formula_cols: set[int],
) -> bool:
    """
    When an incoming correspondence row is linked to an outgoing reply (`out_date >= in_date`,
    or either date blank), auto-fill the blank/pending Reply Date & Status ('تم الرد')
    in Excel without ever overwriting formulas or an existing real reply date.
    """
    if not th or not sc.get("routes"):
        return False
    sec_s = str(sc["routes"][0].get("section") or "").lower()
    if sec_s != "incoming":
        return False

    cols = sc.get("columns") or {}
    rev_groups = sc.get("rev_groups") or []
    date_col = int(cols["date"]["col"]) if "date" in cols else 0
    reply_col = int(cols["reply_date"]["col"]) if "reply_date" in cols else 0
    status_col = int(cols["status"]["col"]) if "status" in cols else 0

    if rev_groups:
        chosen_rg = None
        for rg in rev_groups:
            sd_c = int(rg.get("submit_col") or rg.get("submit_date") or 0)
            if sd_c and ws.cell(row, sd_c).value not in (None, ""):
                chosen_rg = rg
        if chosen_rg is None:
            chosen_rg = rev_groups[0]
        if chosen_rg:
            date_col = int(chosen_rg.get("submit_col") or chosen_rg.get("submit_date") or 0) or date_col
            reply_col = int(chosen_rg.get("reply_col") or chosen_rg.get("reply_date") or 0) or reply_col
            if not status_col:
                status_col = int(chosen_rg.get("action_col") or chosen_rg.get("status") or 0) or status_col

    if not reply_col and not status_col:
        return False

    in_date_iso = _parse_iso_date_str(ws.cell(row, date_col).value) if date_col else ""

    qualifying_dates: list[str] = []
    has_qualifying_out = False
    for om in th.get("members") or []:
        if not isinstance(om, dict):
            continue
        if str(om.get("member_key") or "") == mk:
            continue
        if str(om.get("section") or "").lower() != "outgoing":
            continue
        out_iso = _parse_iso_date_str(om.get("doc_date") or om.get("date"))
        if in_date_iso and out_iso and out_iso < in_date_iso:
            continue
        has_qualifying_out = True
        if out_iso:
            qualifying_dates.append(out_iso)

    if not has_qualifying_out:
        return False

    best_reply_iso = max(qualifying_dates) if qualifying_dates else dt.date.today().isoformat()
    changed = False

    if reply_col and reply_col not in formula_cols:
        rc = plog._resolve_target_cell(ws, row, reply_col)
        cur_rv = rc.value
        cur_rs = str(cur_rv or "").strip()
        is_empty_or_pending = (cur_rv is None or cur_rs == "" or plog._is_pending_status_text(cur_rs))
        if not isinstance(cur_rv, str) or not cur_rv.startswith("="):
            if is_empty_or_pending:
                plog._safe_set_cell(ws, row, reply_col, best_reply_iso, formula_cols, is_date=True)
                plog._clear_missing_date_cell(ws, row, reply_col, formula_cols)
                changed = True

    if status_col and status_col != reply_col and status_col not in formula_cols:
        scell = plog._resolve_target_cell(ws, row, status_col)
        cur_sv = scell.value
        cur_ss = str(cur_sv or "").strip()
        cur_sn = normalize_ar(cur_ss)
        if not (isinstance(cur_sv, str) and cur_sv.startswith("=")):
            if not cur_ss or plog._is_pending_status_text(cur_ss) or cur_sn in ("وارد", "مستلم", "received", "incoming"):
                plog._safe_set_cell(ws, row, status_col, tr("letters.replied_status_val"), formula_cols)
                changed = True

    return changed


def sync_letter_links_to_excel(
    proj: dict[str, Any],
    *,
    affected_sheets: set[str] | None = None,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Write or clear 'الكتاب المرتبط' right in front of every linked correspondence row
    across all project Excel log workbooks (preserving VBA & formulas).
    Also auto-syncs reply date & 'تم الرد' status onto linked incoming rows.
    """
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))
    store = _load_links_store(proj_root)
    by_mk, by_rk = _index_threads_by_key(store)

    entries = plog._iter_project_log_entries(proj_root)
    updated_rows = 0
    reply_synced_rows = 0
    locked_logs: list[str] = []
    touched_logs: list[str] = []

    for ent in entries:
        if not ent.get("exists"):
            continue
        lp = Path(ent["path"])
        try:
            lmap = plog.load_or_create_logmap(lp, device_id=device_id, proj_root=proj_root)
        except Exception:
            continue

        corr_specs = [
            sc for sc in (lmap.get("sheets") or [])
            if sc.get("enabled")
            and sc.get("layout") != "skip"
            and sc.get("routes")
            and str(sc["routes"][0].get("section") or "") in ("incoming", "outgoing")
        ]
        if not corr_specs:
            continue

        if is_file_locked(lp, allow_readonly=False):
            locked_logs.append(lp.name)
            continue

        keep_vba = lp.suffix.lower() == ".xlsm"
        wb = openpyxl.load_workbook(lp, keep_vba=keep_vba)
        wb_changed = False
        try:
            for sc in corr_specs:
                sname = str(sc.get("sheet") or "")
                if sname not in wb.sheetnames:
                    continue
                cols = sc.get("columns") or {}
                ref_col = int(cols["ref_no"]["col"]) if "ref_no" in cols else 0
                subj_col = int(cols["subject"]["col"]) if "subject" in cols else 0
                if not ref_col and not subj_col:
                    continue

                ws = wb[sname]
                data_start = int(sc.get("data_start_row") or 2)
                formula_cols = {int(x) for x in (sc.get("formula_cols") or [])}
                has_rel_col = "related_letters" in cols and bool(cols["related_letters"].get("col"))

                row_targets: list[tuple[int, str, dict[str, Any] | None]] = []
                sheet_has_links = False
                for r in range(data_start, (ws.max_row or data_start) + 1):
                    ref_v = ws.cell(r, ref_col).value if ref_col else None
                    subj_v = ws.cell(r, subj_col).value if subj_col else None
                    ref_s = str(ref_v).strip() if ref_v is not None else ""
                    subj_s = str(subj_v).strip() if subj_v is not None else ""
                    if not ref_s and not subj_s:
                        continue
                    rk = plog.normalize_ref_key(ref_s)
                    mk = make_member_key(ref_s, sheet=sname, section=str(sc["routes"][0].get("section") or ""), row=r)
                    th = by_mk.get(mk) or (by_rk.get(rk) if rk else None)
                    if th:
                        sheet_has_links = True
                    row_targets.append((r, mk, th))

                should_touch_sheet = (
                    sheet_has_links
                    or has_rel_col
                    or (affected_sheets is not None and sname in affected_sheets)
                )
                if not should_touch_sheet or not row_targets:
                    continue

                existed_before = has_rel_col
                target_col, mode = _ensure_related_letters_col(ws, sc)
                if mode == "related_letters" and not existed_before:
                    wb_changed = True
                if target_col in formula_cols:
                    continue

                # Copy side borders from previous column when writing new related_letters cells
                ref_style_col = max(1, target_col - 1)
                for r, mk, th in row_targets:
                    cell = ws.cell(r, target_col)
                    if mode == "related_letters":
                        summary_txt = _format_thread_summary_for_member(th, {"member_key": mk}, joiner="\n") if th else ""
                        cur_val = str(cell.value or "").strip()
                        if summary_txt:
                            if cur_val != summary_txt:
                                cell.value = summary_txt
                                wb_changed = True
                                updated_rows += 1
                            src_c = ws.cell(r, ref_style_col)
                            if src_c.has_style:
                                cell.font = copy.copy(src_c.font)
                                cell.border = copy.copy(src_c.border)
                            cell.fill = copy.copy(_RELATED_CELL_FILL)
                            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
                            line_cnt = summary_txt.count("\n") + 1
                            if line_cnt >= 2:
                                need_h = min(140, 18 * line_cnt + 6)
                                cur_h = ws.row_dimensions[r].height or 0
                                if cur_h < need_h:
                                    ws.row_dimensions[r].height = need_h
                        elif cur_val:
                            cell.value = None
                            cell.fill = PatternFill(fill_type=None)
                            wb_changed = True
                    else:
                        # Notes/Remarks fallback: preserve existing notes cleanly, swap only `[الكتب المرتبطة: ...]`
                        summary_txt = _format_thread_summary_for_member(th, {"member_key": mk}, joiner=" | ") if th else ""
                        cur_n = str(cell.value or "").strip()
                        base_n = _REMARK_TAG_RE.sub("", cur_n).strip(" |")
                        if summary_txt:
                            tag = f"[{tr('letters.remark_prefix')}: {summary_txt}]"
                            desired = f"{base_n} | {tag}" if base_n else tag
                        else:
                            desired = base_n
                        if cur_n != desired:
                            cell.value = desired or None
                            wb_changed = True
                            if summary_txt:
                                updated_rows += 1
                    if th and _sync_incoming_reply_status(ws, sc, r, mk, th, formula_cols):
                        wb_changed = True
                        reply_synced_rows += 1

            if wb_changed:
                plog._save_workbook_atomically(wb, lp)
                touched_logs.append(lp.name)
        finally:
            try:
                wb.close()
            except Exception:
                pass

        if lp.name in touched_logs:
            lmap["log_sha256"] = sha256_of_file(lp)
            lmap["updated_at"] = now_iso()
            atomic_write_json(plog.logmap_path_for(lp, proj_root=proj_root), lmap)

    plog._SUMMARY_CACHE.pop(str(proj_root.resolve()), None)
    return {
        "updated_rows": updated_rows,
        "reply_synced_rows": reply_synced_rows,
        "touched_logs": touched_logs,
        "locked_logs": locked_logs,
    }


def link_letters(
    proj: dict[str, Any],
    items: list[dict[str, Any] | str],
    *,
    topic: str = "",
    sync_excel: bool = True,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Create or extend a multi-party correspondence thread linking 2+ letters across
    Consultant / User-Contractor / Subcontractor / Owner, and write the linked
    references right into the Excel project log(s).
    """
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))
    letters = collect_project_letters(proj, device_id=device_id)
    by_mk = {x["member_key"]: x for x in letters}
    by_rk = {x["ref_key"]: x for x in letters if x.get("ref_key")}

    normalized: list[dict[str, Any]] = []
    for raw in items:
        m = _normalize_item_to_member(raw, by_mk, by_rk)
        if m is not None:
            normalized.append(m)

    if len(normalized) < 2:
        raise ValueError(tr("letters.need_two"))

    store = _load_links_store(proj_root)
    thread = _merge_members_into_store(store, normalized, topic=topic)
    _save_links_store(proj_root, store)

    affected = {str(m.get("sheet") or "") for m in thread.get("members") or [] if m.get("sheet")}
    excel_res = sync_letter_links_to_excel(proj, affected_sheets=affected, device_id=device_id) if sync_excel else {"updated_rows": 0, "touched_logs": [], "locked_logs": []}

    return {
        "ok": True,
        "thread": thread,
        "excel": excel_res,
        "excel_locked": bool(excel_res.get("locked_logs")),
    }


def unlink_letter(
    proj: dict[str, Any],
    *,
    ref_no: str = "",
    sheet: str = "",
    target_ref_no: str = "",
    target_sheet: str = "",
    sync_excel: bool = True,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Unlink a specific pair (or remove a letter from its thread), remember the rejection
    so auto-discovery will not re-link that pair, and update Excel immediately.
    """
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))
    store = _load_links_store(proj_root)

    rk_a = plog.normalize_ref_key(ref_no)
    mk_a = make_member_key(ref_no, sheet=sheet)
    rk_b = plog.normalize_ref_key(target_ref_no)
    mk_b = make_member_key(target_ref_no, sheet=target_sheet)

    affected_sheets: set[str] = set()
    if sheet:
        affected_sheets.add(sheet)
    if target_sheet:
        affected_sheets.add(target_sheet)

    ig = _ignored_set_from_store(store)
    new_threads: list[dict[str, Any]] = []
    removed_any = False

    for th in store.get("threads") or []:
        members = th.get("members") or []
        idx_a = next((i for i, m in enumerate(members) if _member_matches_target(m, mk_a, rk_a, sheet)), None)
        if idx_a is None:
            new_threads.append(th)
            continue

        removed_any = True
        for m in members:
            if m.get("sheet"):
                affected_sheets.add(str(m["sheet"]))
        actual_mk_a = str(members[idx_a].get("member_key") or mk_a)

        if target_ref_no:
            idx_b = next((i for i, m in enumerate(members) if _member_matches_target(m, mk_b, rk_b, target_sheet)), None)
            if idx_b is not None:
                actual_mk_b = str(members[idx_b].get("member_key") or mk_b)
                ig.add(_canonical_pair(actual_mk_a, actual_mk_b))
                kept = [m for i, m in enumerate(members) if i != idx_b]
                if len(kept) >= 2:
                    th["members"] = kept
                    th["updated_at"] = now_iso()
                    new_threads.append(th)
                else:
                    # Also ignore remaining pair if dissolving
                    continue
            else:
                new_threads.append(th)
        else:
            # Remove member A from thread and ignore A against all remaining thread members
            kept = [m for i, m in enumerate(members) if i != idx_a]
            for km in kept:
                ig.add(_canonical_pair(actual_mk_a, str(km.get("member_key") or "")))
            if len(kept) >= 2:
                th["members"] = kept
                th["updated_at"] = now_iso()
                new_threads.append(th)

    store["threads"] = new_threads
    store["ignored_pairs"] = [list(p) for p in sorted(ig)]
    _save_links_store(proj_root, store)

    excel_res = sync_letter_links_to_excel(proj, affected_sheets=affected_sheets, device_id=device_id) if (sync_excel and removed_any) else {"updated_rows": 0, "touched_logs": [], "locked_logs": []}
    return {
        "ok": True,
        "removed": removed_any,
        "excel": excel_res,
    }


def auto_link_project_letters(
    proj: dict[str, Any],
    *,
    min_score: float = 0.58,
    apply: bool = True,
    sync_excel: bool = True,
    pairs: list[dict[str, Any]] | None = None,
    device_id: str = "",
) -> dict[str, Any]:
    """
    Discover and optionally apply cross-party topic links across all correspondence in the project.
    Transitively groups connected matches into multi-party threads (e.g. Consultant + User + Subcontractor).
    """
    proj_root = resolve_local_or_unc(str(proj.get("root") or ""))
    candidates = discover_unlinked_matches(proj, min_score=min_score, device_id=device_id) if pairs is None else list(pairs)

    if not apply:
        return {
            "ok": True,
            "applied": False,
            "count": len(candidates),
            "matches": candidates,
        }

    if not candidates:
        return {
            "ok": True,
            "applied": True,
            "linked_pairs": 0,
            "threads_updated": 0,
            "matches": [],
            "excel": {"updated_rows": 0, "touched_logs": [], "locked_logs": []},
        }

    letters = collect_project_letters(proj, device_id=device_id)
    by_mk = {x["member_key"]: x for x in letters}
    by_rk = {x["ref_key"]: x for x in letters if x.get("ref_key")}
    store = _load_links_store(proj_root)

    # Group connected components among selected candidate pairs using union-find
    parent: dict[str, str] = {}

    def _find(x: str) -> str:
        while parent.get(x, x) != x:
            parent[x] = parent.get(parent[x], parent[x])
            x = parent[x]
        return x

    def _union(x: str, y: str) -> None:
        rx, ry = _find(x), _find(y)
        if rx != ry:
            parent[ry] = rx

    member_pool: dict[str, dict[str, Any]] = {}
    for p in candidates:
        ma = _normalize_item_to_member(p.get("a") or {}, by_mk, by_rk)
        mb = _normalize_item_to_member(p.get("b") or {}, by_mk, by_rk)
        if not ma or not mb:
            continue
        ka, kb = ma["member_key"], mb["member_key"]
        member_pool[ka] = ma
        member_pool[kb] = mb
        parent.setdefault(ka, ka)
        parent.setdefault(kb, kb)
        _union(ka, kb)

    groups: dict[str, list[dict[str, Any]]] = {}
    for mk, m in member_pool.items():
        groups.setdefault(_find(mk), []).append(m)

    affected: set[str] = set()
    threads_updated = 0
    for grp in groups.values():
        if len(grp) < 2:
            continue
        th = _merge_members_into_store(store, grp)
        threads_updated += 1
        for m in th.get("members") or []:
            if m.get("sheet"):
                affected.add(str(m["sheet"]))

    _save_links_store(proj_root, store)
    excel_res = sync_letter_links_to_excel(proj, affected_sheets=affected, device_id=device_id) if (sync_excel and threads_updated > 0) else {"updated_rows": 0, "touched_logs": [], "locked_logs": []}

    return {
        "ok": True,
        "applied": True,
        "linked_pairs": len(candidates),
        "threads_updated": threads_updated,
        "matches": candidates,
        "excel": excel_res,
        "excel_locked": bool(excel_res.get("locked_logs")),
    }
