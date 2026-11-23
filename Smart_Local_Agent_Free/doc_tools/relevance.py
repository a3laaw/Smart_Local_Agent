#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
relevance.py — Unified Arabic text normalization, project-relevance scoring,
cross-project conflict detection, ignore-pattern filtering, fuzzy entity deduplication,
and reversible rename-in-place journaling (Smart_Local_Agent_Free).
"""

from __future__ import annotations

import datetime as dt
import difflib
import fnmatch
import os
import re
import secrets
import shutil
from pathlib import Path
from typing import Any

from i18n_ar import tr
from security_utils import (
    DANGEROUS_EXTENSIONS,
    SYSTEM_TEMP_FILENAMES,
    atomic_write_json,
    now_iso,
    read_json_safe,
    sha256_of_file,
)

# ── Unified Arabic + English Normalization ──────────────────────────────────
_AR_DIACRITICS_RE = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED\u0640]")
_AR_DIGIT_TRANS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_AR_CHAR_TRANS = str.maketrans({
    "أ": "ا",
    "إ": "ا",
    "آ": "ا",
    "ٱ": "ا",
    "ة": "ه",
    "ى": "ي",
    "ؤ": "و",
    "ئ": "ي",
    "ک": "ك",
    "ی": "ي",
})

UNRELATED_POLICIES: tuple[str, ...] = ("rename_in_place", "new_project_by_name", "ask")
DEFAULT_RELEVANCE_THRESHOLD: float = 0.6

# Corporate/project filler words stripped when comparing entity names for deduplication
_ENTITY_STOPWORDS: frozenset[str] = frozenset({
    "شركه", "موسسه", "مكتب", "مجموعه", "دار", "مشروع", "مقاولات", "للمقاولات",
    "التجاريه", "العامه", "للتجاره", "والمقاولات", "الهندسي", "الهندسيه",
    "للاستشارات", "الاستشاري", "المحدوده", "الكويتيه", "الكويت", "ذ", "م", "ذمم",
    "فرع", "قسم", "اداره", "انشاء", "وانجاز", "وصيانه", "توريد", "وتركيب",
    "الساده", "المحترمين", "المحترم",
    "company", "co", "corp", "corporation", "llc", "wll", "ltd", "inc", "est",
    "group", "general", "trading", "contracting", "construction", "engineering",
    "consultants", "consultant", "consulting", "office", "project", "kuwait",
    "for", "and", "the", "of", "al",
})

# Generic construction words that should not by themselves identify a specific project
_GENERIC_PROJECT_WORDS: frozenset[str] = frozenset({
    "مشروع", "عقد", "مناقصه", "ممارسه", "وزاره", "شركه", "مكتب", "انشاء", "انجاز",
    "صيانه", "مبني", "مركز", "منطقه", "الكويت", "دوله", "العامه", "الهندسي",
    "للمقاولات", "التجاريه", "رقم", "سنه", "قسم", "اداره", "اعمال", "توريد",
    "project", "contract", "tender", "ministry", "company", "center", "centre",
    "building", "kuwait", "state", "construction", "maintenance", "general", "no",
})

# Explicit personal / non-project document signals
_PERSONAL_DOC_KEYWORDS: tuple[tuple[str, str], ...] = (
    ("سيره ذاتيه", "سيرة_ذاتية"),
    ("curriculum vitae", "CV"),
    ("resume", "CV"),
    ("جواز سفر", "جواز_سفر"),
    ("passport", "جواز_سفر"),
    ("بطاقه مدنيه", "بطاقة_مدنية"),
    ("civil id", "بطاقة_مدنية"),
    ("تذكره سفر", "تذكرة_سفر"),
    ("boarding pass", "تذكرة_سفر"),
    ("حجز فندق", "حجز_فندق"),
    ("شهاده راتب شخصيه", "شهادة_راتب"),
    ("كشف حساب شخصي", "كشف_حساب_شخصي"),
    ("personal bank statement", "كشف_حساب_شخصي"),
    ("وصفه طبيه", "مستند_طبي"),
    ("تحليل طبي", "مستند_طبي"),
    ("رسوم دراسيه", "رسوم_دراسية"),
    ("مدرسه خاصه", "رسوم_دراسية"),
    ("عقد ايجار شقه", "عقد_إيجار_شخصي"),
    ("مخالفه مروريه", "مخالفة_مرورية"),
    ("رخصه قياده", "رخصة_قيادة"),
    ("فاتوره هاتف شخصي", "فاتورة_شخصية"),
    ("فاتوره كهرباء منزل", "فاتورة_منزل"),
)

DEFAULT_IGNORE_PATTERNS: tuple[str, ...] = (
    "~$*",
    "*.tmp",
    "*.temp",
    "*.part",
    "*.crdownload",
    "*.partial",
    "*.bak",
    "*.swp",
    "*.lock",
    "*.lck",
    "Thumbs.db",
    "desktop.ini",
    ".DS_Store",
    "*.tplmap.json",
    "*.logmap.json",
    "_rename_journal.json",
    "_archive_journal.json",
    "_Archived_*",
    "_Backups",
    "_Backups/*",
    "_System_Agent",
    "_System_Agent/*",
    "_Auto_Projects_*",
    "_Auto_Projects_*/*",
)


def normalize_ar(text: str, *, keep_punct: bool = False) -> str:
    """
    Canonical Arabic + English text normalizer used everywhere:
    - strips harakat/diacritics and tatweel
    - unifies Alef (أ إ آ ٱ -> ا), Taa Marbuta (ة -> ه), Alef Maqsura (ى -> ي), Hamza (ؤ->و, ئ->ي)
    - converts Eastern Arabic / Persian digits to ASCII 0-9
    - lowercases Latin characters and collapses whitespace
    """
    if not text:
        return ""
    s = str(text)
    s = _AR_DIACRITICS_RE.sub("", s)
    s = s.translate(_AR_DIGIT_TRANS)
    s = s.translate(_AR_CHAR_TRANS)
    s = s.lower()
    if not keep_punct:
        s = re.sub(r"[^\w\u0600-\u06FF]+", " ", s, flags=re.UNICODE)
        s = s.replace("_", " ")
    return re.sub(r"\s+", " ", s).strip()


def tokenize_words(text: str) -> list[str]:
    """
    Split text or filenames into true word tokens (supports camelCase, underscores,
    hyphens, dots, slashes, Arabic and English words) — never raw substrings.
    """
    if not text:
        return []
    # Split camelCase first (e.g. MonthlyReport -> Monthly Report, VOForm -> VO Form)
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", str(text))
    s = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", s)
    # Separate letters and digits when directly adjacent if helpful? Keep tokens clean:
    norm = normalize_ar(s, keep_punct=False)
    return [tok for tok in norm.split() if tok]


def normalize_entity_core(name: str) -> tuple[str, list[str]]:
    """
    Extract core distinguishing tokens of a company/project/entity name.
    Example:
      'شركة النور'              -> ('النور', ['النور'])
      'شركة النور للمقاولات'    -> ('النور', ['النور'])
      'شركه النور التجاريه ذ.م.م' -> ('النور', ['النور'])
    """
    full_norm = normalize_ar(name)
    raw_tokens = [t for t in full_norm.split() if len(t) > 1 or t.isdigit()]
    core: list[str] = []
    for tok in raw_tokens:
        if tok in _ENTITY_STOPWORDS:
            continue
        # Strip leading Arabic definite article 'ال' for secondary comparison while keeping meaning
        core.append(tok)
    if not core:
        core = raw_tokens
    return " ".join(core), core


def _strip_al(tok: str) -> str:
    if tok.startswith("ال") and len(tok) >= 5:
        return tok[2:]
    return tok


def entity_similarity(name_a: str, name_b: str) -> float:
    """
    Compute deterministic similarity [0.0..1.0] between two Arabic/English entity names.
    Handles spelling variants (أ/ا, ة/ه, ى/ي) and corporate suffixes ('للمقاولات', 'Co. WLL').
    """
    na = normalize_ar(name_a)
    nb = normalize_ar(name_b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0

    core_a, toks_a = normalize_entity_core(name_a)
    core_b, toks_b = normalize_entity_core(name_b)
    if core_a and core_b and core_a == core_b:
        return 0.96

    set_a = {_strip_al(t) for t in toks_a if t}
    set_b = {_strip_al(t) for t in toks_b if t}
    if not set_a or not set_b:
        return difflib.SequenceMatcher(None, na, nb).ratio()

    if set_a == set_b:
        return 0.95

    inter = set_a & set_b
    union = set_a | set_b
    jaccard = len(inter) / max(len(union), 1)
    # Containment score when shorter core is non-trivial (e.g. {'نور'} subset of {'نور', 'حديثه'})
    smaller = min(len(set_a), len(set_b))
    larger = max(len(set_a), len(set_b))
    containment = len(inter) / max(smaller, 1)
    if containment == 1.0 and inter:
        # E.g. 'شركة النور' vs 'شركة النور للمقاولات والإنشاءات'
        first_match = (_strip_al(toks_a[0]) == _strip_al(toks_b[0])) if (toks_a and toks_b) else False
        if smaller >= 2 or (first_match and larger <= 2):
            return 0.90
        if first_match and larger <= 3:
            return 0.84

    seq = difflib.SequenceMatcher(None, core_a, core_b).ratio()
    return round(max(jaccard * 0.65 + seq * 0.35, seq * 0.92), 4)


def fuzzy_entity_match(
    candidate: str,
    existing: list[dict[str, Any]],
    *,
    name_key: str = "name",
    aliases_key: str = "aliases",
    threshold: float = 0.80,
) -> tuple[dict[str, Any] | None, float]:
    """
    Find the best matching existing project/entity above `threshold`.
    Returns `(matched_dict, score)` or `(None, best_score)`.
    """
    best_item: dict[str, Any] | None = None
    best_score = 0.0
    for item in existing:
        names = [str(item.get(name_key) or "")]
        raw_aliases = item.get(aliases_key) or []
        if isinstance(raw_aliases, str):
            raw_aliases = [x.strip() for x in re.split(r"[,،;\n]+", raw_aliases) if x.strip()]
        names.extend(str(a) for a in raw_aliases if a)
        for nm in names:
            sc = entity_similarity(candidate, nm)
            if sc > best_score:
                best_score = sc
                best_item = item
    if best_item is not None and best_score >= threshold:
        return best_item, best_score
    return None, best_score


# ── Ignore Patterns & Locked-File Protection ────────────────────────────────
def parse_ignore_patterns(raw: Any) -> list[str]:
    if not raw:
        return []
    if isinstance(raw, str):
        lines = raw.splitlines()
    elif isinstance(raw, (list, tuple, set)):
        lines = [str(x) for x in raw]
    else:
        return []
    out: list[str] = []
    for ln in lines:
        s = ln.strip()
        if not s or s.startswith("#"):
            continue
        out.append(s)
    return out


def match_ignore_patterns(
    file_path: Path,
    custom_patterns: list[str] | str | None = None,
    *,
    inbox_root: Path | None = None,
) -> str | None:
    """
    Return the reason/pattern string if `file_path` must be ignored, else `None`.
    Enforces hardcoded safety exclusions (system files, temp/lock files, executables)
    PLUS user-supplied `.gitignore`-style glob patterns (supports `!negation` for user rules).
    """
    p = Path(file_path)
    name = p.name
    name_low = name.lower()
    ext_low = p.suffix.lower()

    # 1. Hardcoded non-overridable safety exclusions
    if not name or name.startswith("~$") or name.startswith(".~"):
        return "office_lock_file"
    if name.startswith(".") and name_low != ".keep":
        return "hidden_dot_file"
    if name_low in SYSTEM_TEMP_FILENAMES:
        return f"system_file:{name}"
    if ext_low in DANGEROUS_EXTENSIONS:
        return f"executable_or_system_ext:{ext_low}"
    for pat in DEFAULT_IGNORE_PATTERNS:
        if fnmatch.fnmatchcase(name_low, pat.lower()):
            return f"default_ignore:{pat}"

    rel_posix = name
    if inbox_root is not None:
        try:
            rel_posix = p.resolve().relative_to(Path(inbox_root).resolve()).as_posix()
        except Exception:
            rel_posix = name
        # Check parent parts against internal system folders
        for part in rel_posix.split("/")[:-1]:
            pl = part.lower()
            if pl.startswith(("_system_agent", "_backups", "_archived_", "_auto_projects_", ".")):
                return f"internal_dir:{part}"

    # 2. User-configured .gitignore-style patterns
    patterns = parse_ignore_patterns(custom_patterns)
    matched_pat: str | None = None
    rel_low = rel_posix.lower()
    for pat in patterns:
        negate = pat.startswith("!")
        rule = pat[1:].strip() if negate else pat
        if not rule:
            continue
        rule_low = rule.replace("\\", "/").lower()
        is_hit = False
        if rule_low.endswith("/"):
            dir_pat = rule_low.rstrip("/")
            parts = rel_low.split("/")[:-1]
            is_hit = any(fnmatch.fnmatchcase(pt, dir_pat) for pt in parts) or rel_low.startswith(dir_pat + "/")
        elif "/" in rule_low:
            trimmed = rule_low.lstrip("/")
            is_hit = fnmatch.fnmatchcase(rel_low, trimmed) or fnmatch.fnmatchcase(rel_low, f"*/{trimmed}")
        else:
            is_hit = fnmatch.fnmatchcase(name_low, rule_low) or any(
                fnmatch.fnmatchcase(pt, rule_low) for pt in rel_low.split("/")
            )
        if is_hit:
            matched_pat = None if negate else f"user_pattern:{pat}"
    return matched_pat


def is_file_locked(
    path: Path,
    *,
    min_age_seconds: float = 0.0,
    allow_readonly: bool = True,
) -> bool:
    """
    Check if a file is currently open/locked by another process (Word/Excel/Scanner)
    or still being written.
    Handles Windows Server SMB / Read-Only shares gracefully when `allow_readonly=True`:
    a sharing violation (WinError 32/33) means genuinely locked/writing, whereas a
    read-only share (PermissionError on `"r+b"` when `"rb"` succeeds) is readable!
    """
    p = Path(path)
    if not p.exists() or not p.is_file():
        return True
    # Check adjacent MS Office lock file `~$filename.xlsx`
    office_lock = p.parent / f"~${p.name}"
    if office_lock.exists():
        return True
    if p.suffix.lower() in (".xlsx", ".xlsm", ".docx") and len(p.name) > 2:
        alt_lock = p.parent / f"~${p.name[2:]}"
        if alt_lock.exists():
            return True
    try:
        st = p.stat()
        if st.st_size == 0 and min_age_seconds > 0:
            return True
        if min_age_seconds > 0 and (dt.datetime.now().timestamp() - st.st_mtime) < min_age_seconds:
            return True
    except OSError:
        return True

    def _can_read_safely() -> bool:
        if not allow_readonly:
            return False
        try:
            with open(p, "rb") as rfh:
                rfh.read(16)
            return True
        except OSError:
            return False

    try:
        if os.name == "nt":
            import msvcrt  # type: ignore
            try:
                with open(p, "r+b") as fh:
                    try:
                        msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                        msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
                    except OSError as lock_exc:
                        # On SMB/Windows Server shares, msvcrt byte-range locking can fail with
                        # WinError 50/59/1/5 even when the file was opened r+b cleanly.
                        # Only WinError 32 (sharing) or 33 (lock violation) indicates an active lock.
                        if getattr(lock_exc, "winerror", None) in (32, 33):
                            return True
                        return not _can_read_safely()
            except OSError as open_exc:
                if getattr(open_exc, "winerror", None) in (32, 33):
                    return True
                return not _can_read_safely()
        else:
            try:
                import fcntl
                try:
                    with open(p, "r+b") as fh:
                        fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                        fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
                except PermissionError:
                    return not _can_read_safely()
            except OSError as exc:
                import errno
                if exc.errno in (errno.EAGAIN, errno.EACCES, errno.EWOULDBLOCK):
                    return True
                return not _can_read_safely()
    except OSError:
        return True
    return False


# ── Project Profile & Deterministic + AI Relevance Scoring ──────────────────
def _split_list_field(raw: Any) -> list[str]:
    if not raw:
        return []
    if isinstance(raw, (list, tuple, set)):
        return [str(x).strip() for x in raw if str(x).strip()]
    return [x.strip() for x in re.split(r"[,،;\n|]+", str(raw)) if x.strip()]


def build_project_profile(proj: dict[str, Any]) -> dict[str, Any]:
    """Build normalized match fingerprints from a project's metadata."""
    name = str(proj.get("name") or "").strip()
    code = str(proj.get("code") or "").strip()
    contract_no = str(proj.get("contract_no") or "").strip()
    owner = str(proj.get("owner") or "").strip()
    consultant = str(proj.get("consultant") or "").strip()
    contractor = str(proj.get("contractor") or "").strip()
    aliases = _split_list_field(proj.get("aliases"))
    keywords = _split_list_field(proj.get("keywords"))
    subcontractors = _split_list_field(proj.get("subcontractors"))
    # Include contact names/roles if configured
    for c in (proj.get("contacts") or []):
        if isinstance(c, dict) and c.get("name"):
            subcontractors.append(str(c["name"]).strip())

    name_norm = normalize_ar(name)
    distinctive_name_tokens = [
        t for t in name_norm.split()
        if len(t) >= 3 and t not in _GENERIC_PROJECT_WORDS and not t.isdigit()
    ]
    # Build 2-gram phrases from distinctive name tokens (e.g. 'صباح السالم')
    name_bigrams = [
        f"{distinctive_name_tokens[i]} {distinctive_name_tokens[i + 1]}"
        for i in range(len(distinctive_name_tokens) - 1)
    ]

    contract_norm = normalize_ar(contract_no, keep_punct=True)
    contract_compact = re.sub(r"[\s\-_/\\]+", "", normalize_ar(contract_no, keep_punct=True))

    code_norm = normalize_ar(code)
    # Ignore ultra-generic auto codes like 'prj' alone
    code_is_specific = bool(code_norm and len(code_norm) >= 3 and code_norm not in ("prj", "proj", "pro", "doc", "new"))

    return {
        "id": str(proj.get("id") or ""),
        "name": name,
        "name_norm": name_norm,
        "distinctive_tokens": distinctive_name_tokens,
        "name_bigrams": name_bigrams,
        "code": code,
        "code_norm": code_norm,
        "code_is_specific": code_is_specific,
        "contract_no": contract_no,
        "contract_norm": contract_norm,
        "contract_compact": contract_compact if len(contract_compact) >= 4 else "",
        "owner_norm": normalize_ar(owner),
        "consultant_norm": normalize_ar(consultant),
        "contractor_norm": normalize_ar(contractor),
        "owner_core": normalize_entity_core(owner)[0] if owner else "",
        "consultant_core": normalize_entity_core(consultant)[0] if consultant else "",
        "contractor_core": normalize_entity_core(contractor)[0] if contractor else "",
        "aliases_norm": [normalize_ar(a) for a in aliases if len(normalize_ar(a)) >= 3],
        "keywords_norm": [normalize_ar(k) for k in keywords if len(normalize_ar(k)) >= 2],
        "subs_norm": [normalize_entity_core(s)[0] for s in subcontractors if len(normalize_entity_core(s)[0]) >= 3],
    }


def _contains_whole_phrase(haystack_norm: str, haystack_tokens: set[str], phrase_norm: str) -> bool:
    if not phrase_norm or not haystack_norm:
        return False
    parts = phrase_norm.split()
    if len(parts) == 1:
        return parts[0] in haystack_tokens
    return bool(re.search(rf"(?<!\w){re.escape(phrase_norm)}(?!\w)", haystack_norm))


def score_text_against_project(
    corpus_norm: str,
    corpus_compact: str,
    corpus_tokens: set[str],
    profile: dict[str, Any],
) -> tuple[float, list[str]]:
    """
    Deterministically score how strongly a document's text/filename/metadata
    belongs to `profile`. Returns `(score in 0..1, matched_signals)`.
    """
    score = 0.0
    signals: list[str] = []

    # 1. Exact contract number match (strongest identifier)
    cc = profile["contract_compact"]
    if cc and cc in corpus_compact:
        score = max(score, 0.92)
        signals.append(f"contract_no:{profile['contract_no']}")

    # 2. Full project name or explicit alias
    if profile["name_norm"] and len(profile["name_norm"]) >= 5:
        if _contains_whole_phrase(corpus_norm, corpus_tokens, profile["name_norm"]):
            score = max(score, 0.90)
            signals.append(f"project_name:{profile['name']}")
        else:
            # Distinctive name tokens & bigrams (e.g. 'صباح السالم الصحي')
            dtoks = profile["distinctive_tokens"]
            if dtoks:
                matched_dt = [t for t in dtoks if t in corpus_tokens or _strip_al(t) in {_strip_al(x) for x in corpus_tokens}]
                ratio = len(matched_dt) / len(dtoks)
                if len(matched_dt) >= 2 and ratio >= 0.65:
                    score = max(score, 0.82)
                    signals.append(f"project_name_tokens:{' '.join(matched_dt)}")
                elif profile["name_bigrams"]:
                    for bg in profile["name_bigrams"]:
                        if _contains_whole_phrase(corpus_norm, corpus_tokens, bg):
                            score = max(score, 0.78)
                            signals.append(f"project_phrase:{bg}")
                            break
                elif len(dtoks) == 1 and len(matched_dt) == 1 and len(dtoks[0]) >= 4:
                    score = max(score, 0.75)
                    signals.append(f"project_distinctive:{dtoks[0]}")

    for al in profile["aliases_norm"]:
        if _contains_whole_phrase(corpus_norm, corpus_tokens, al):
            score = max(score, 0.85)
            signals.append(f"alias:{al}")
            break

    # 3. Specific project code as whole token
    if profile["code_is_specific"] and _contains_whole_phrase(corpus_norm, corpus_tokens, profile["code_norm"]):
        score = max(score, 0.76)
        signals.append(f"project_code:{profile['code']}")

    # 4. Parties (consultant, owner, contractor, subcontractors)
    party_hits = 0
    for role_label, core_key, full_key, boost in (
        ("consultant", "consultant_core", "consultant_norm", 0.48),
        ("owner", "owner_core", "owner_norm", 0.42),
        ("contractor", "contractor_core", "contractor_norm", 0.40),
    ):
        full_v = profile[full_key]
        core_v = profile[core_key]
        if full_v and len(full_v) >= 4 and _contains_whole_phrase(corpus_norm, corpus_tokens, full_v):
            party_hits += 1
            score = min(1.0, max(score, boost) + (0.18 if score >= 0.35 else 0.0))
            signals.append(f"{role_label}:{full_v}")
        elif core_v and len(core_v) >= 4 and _contains_whole_phrase(corpus_norm, corpus_tokens, core_v):
            party_hits += 1
            score = min(1.0, max(score, boost - 0.04) + (0.15 if score >= 0.35 else 0.0))
            signals.append(f"{role_label}:{core_v}")

    if party_hits >= 2:
        score = max(score, 0.72)

    for sub in profile["subs_norm"]:
        if _contains_whole_phrase(corpus_norm, corpus_tokens, sub):
            score = min(1.0, max(score, 0.62) + (0.10 if score >= 0.4 else 0.0))
            signals.append(f"subcontractor:{sub}")
            break

    # 5. User-defined project keywords
    kw_hits = [kw for kw in profile["keywords_norm"] if _contains_whole_phrase(corpus_norm, corpus_tokens, kw)]
    if kw_hits:
        kw_boost = 0.68 if len(kw_hits) >= 2 else 0.62
        score = min(1.0, max(score, kw_boost))
        signals.append(f"keywords:{','.join(kw_hits[:3])}")

    return round(min(score, 1.0), 3), signals


def match_learned_rule(
    rules: list[dict[str, Any]],
    *,
    file_name: str,
    corpus_norm: str,
    entity_hint: str = "",
    sender_hint: str = "",
) -> dict[str, Any] | None:
    """Match a document against persistent user-approved `learned_rules.json`."""
    fname_norm = normalize_ar(Path(file_name).stem)
    fname_low = file_name.lower()
    ent_norm = normalize_ar(entity_hint)
    snd_norm = normalize_ar(sender_hint)
    corpus_tokens = set(corpus_norm.split())

    for rule in rules:
        if not isinstance(rule, dict) or not rule.get("enabled", True):
            continue
        mtype = str(rule.get("match_type") or "entity").strip().lower()
        pattern = str(rule.get("pattern") or "").strip()
        if not pattern:
            continue
        pnorm = normalize_ar(pattern)
        if mtype == "filename_glob":
            if fnmatch.fnmatchcase(fname_low, pattern.lower()):
                return rule
        elif mtype == "entity":
            if ent_norm and entity_similarity(pattern, entity_hint) >= 0.82:
                return rule
            if pnorm and len(pnorm) >= 3 and _contains_whole_phrase(corpus_norm, corpus_tokens, pnorm):
                return rule
        elif mtype == "sender":
            if snd_norm and (pnorm in snd_norm or entity_similarity(pattern, sender_hint) >= 0.82):
                return rule
        elif mtype == "keyword":
            if pnorm and (_contains_whole_phrase(corpus_norm, corpus_tokens, pnorm) or pnorm in fname_norm):
                return rule
    return None


def _detect_personal_doc(corpus_norm: str) -> str | None:
    for kw, label in _PERSONAL_DOC_KEYWORDS:
        kw_norm = normalize_ar(kw)
        if kw_norm and re.search(rf"(?<!\w){re.escape(kw_norm)}(?!\w)", corpus_norm):
            return label
    return None


def extract_entity_candidate(extracted_text: str, ai_meta: dict[str, Any] | None = None, file_name: str = "") -> str:
    """Extract a clean organization/project/client name from AI metadata or document text."""
    meta = ai_meta or {}
    for key in ("detected_project_or_entity", "unrelated_entity", "project", "party_name", "from_party", "sender", "vendor"):
        val = str(meta.get(key) or "").strip()
        if val and val.lower() not in ("unknown", "none", "n/a", "null", "عام", "غير معروف", "غير محدد"):
            return val[:70]

    if extracted_text:
        # Look for explicit Arabic/English entity patterns
        patterns = [
            r"(?:مشروع|عقد)\s*[:\-]?\s*([^\n\r،,;|]{4,55})",
            r"((?:شركة|مؤسسة|مكتب)\s+[\u0600-\u06FF0-9\s]{3,45})",
            r"(?:Project|Client|Company)\s*[:\-]\s*([A-Za-z0-9 &.\-]{4,55})",
        ]
        for pat in patterns:
            m = re.search(pat, extracted_text, flags=re.IGNORECASE)
            if m:
                cand = re.sub(r"\s+", " ", m.group(1)).strip(" .:-_،,")
                if len(cand) >= 4:
                    return cand[:65]
    return ""


def build_unrelated_filename(
    file_path: Path,
    extracted_text: str = "",
    ai_meta: dict[str, Any] | None = None,
    personal_kind: str | None = None,
) -> str:
    """
    Build a clean, deterministic, human-readable filename for an unrelated file:
      `<YYYY-MM-DD>_<kind>_<entity_or_subject>.<ext>`
    """
    meta = ai_meta or {}
    ext = file_path.suffix.lower()

    # Date
    doc_date = str(meta.get("doc_date") or meta.get("date") or "").strip()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", doc_date):
        m = re.search(r"(20\d{2})[-/.](0[1-9]|1[0-2])[-/.](0[1-9]|[12]\d|3[01])", f"{file_path.stem} {extracted_text[:600]}")
        if m:
            doc_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        else:
            doc_date = dt.date.today().isoformat()

    # Kind
    kind_map = {
        "incoming": "وارد",
        "outgoing": "صادر",
        "petty_cash": "فاتورة",
        "contract_ipc": "مالي",
        "submittal": "اعتماد",
        "site_photo": "صورة",
        "site_report": "تقرير",
        "general": "مستند",
        "unrelated": "مستند",
    }
    if personal_kind:
        kind = personal_kind
    else:
        raw_kind = str(meta.get("unrelated_kind") or meta.get("subtype") or meta.get("doc_type") or "").strip()
        kind = kind_map.get(raw_kind.lower(), raw_kind or "مستند")

    # Entity or subject
    subject = str(meta.get("subject_ar") or meta.get("subject") or "").strip()
    entity = extract_entity_candidate(extracted_text, meta, file_path.name)
    desc = subject or entity or file_path.stem
    if entity and subject and normalize_ar(entity) not in normalize_ar(subject):
        desc = f"{entity}_{subject}"

    def _clean(seg: str, limit: int) -> str:
        s = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", seg)
        s = re.sub(r"\s+", "_", s.strip(" ._-"))
        return s[:limit].rstrip("._-") or "ملف"

    stem = f"{doc_date}_{_clean(kind, 22)}_{_clean(desc, 55)}"
    return f"{stem}{ext}"


def evaluate_relevance(
    file_path: Path,
    extracted_text: str,
    ai_meta: dict[str, Any] | None,
    active_proj: dict[str, Any],
    all_projects: list[dict[str, Any]] | None = None,
    learned_rules: list[dict[str, Any]] | None = None,
    *,
    threshold: float = DEFAULT_RELEVANCE_THRESHOLD,
    shared_inbox: bool = False,
) -> dict[str, Any]:
    """
    Evaluate whether `file_path` belongs to `active_proj`, belongs to another registered
    project (conflict!), matches a learned rule, is low-confidence, or is unrelated.
    """
    meta = ai_meta or {}
    all_projs = all_projects or [active_proj]
    rules = learned_rules or []
    thr = max(0.1, min(0.95, float(threshold if threshold is not None else DEFAULT_RELEVANCE_THRESHOLD)))

    # Assemble searchable corpus from filename + extracted text + AI fields
    meta_bits = " ".join(
        str(meta.get(k) or "")
        for k in (
            "subject_ar", "subject", "party_name", "from_party", "to_party", "authority",
            "project", "po_or_contract_no", "ref_no", "notes",
            "detected_project_or_entity", "unrelated_entity", "contract_no",
        )
    )
    raw_corpus = f"{file_path.stem}\n{extracted_text or ''}\n{meta_bits}"
    corpus_norm = normalize_ar(raw_corpus, keep_punct=False)
    corpus_punct = normalize_ar(raw_corpus, keep_punct=True)
    corpus_compact = re.sub(r"[\s\-_/\\]+", "", corpus_punct)
    corpus_tokens = set(corpus_norm.split())

    entity_hint = extract_entity_candidate(extracted_text, meta, file_path.name)
    sender_hint = str(meta.get("party_name") or meta.get("from_party") or meta.get("sender") or "").strip()
    personal_kind = _detect_personal_doc(corpus_norm)
    clean_name = build_unrelated_filename(file_path, extracted_text, meta, personal_kind=personal_kind)

    # 1. Learned Rules take top deterministic priority
    matched_rule = match_learned_rule(
        rules,
        file_name=file_path.name,
        corpus_norm=corpus_norm,
        entity_hint=entity_hint,
        sender_hint=sender_hint,
    )
    if matched_rule:
        r_action = str(matched_rule.get("action") or "rename_in_place")
        r_proj_id = str(matched_rule.get("target_project_id") or "")
        active_id = str(active_proj.get("id") or "")
        if r_action == "attach_project" and (not r_proj_id or r_proj_id == active_id):
            return {
                "relevant": True,
                "score": 0.99,
                "reason": tr("server.rule_matched", pattern=matched_rule.get("pattern")),
                "decision_kind": "relevant",
                "forced_action": "attach_project",
                "target_project_id": active_id,
                "conflict_project_id": None,
                "conflict_project_name": "",
                "suggested_entity": entity_hint,
                "suggested_clean_name": clean_name,
                "low_confidence": False,
                "matched_rule_id": matched_rule.get("id"),
                "signals": [f"learned_rule:{matched_rule.get('id')}"],
            }
        return {
            "relevant": False,
            "score": 0.05,
            "reason": tr("server.rule_matched", pattern=matched_rule.get("pattern")),
            "decision_kind": "learned_rule",
            "forced_action": r_action,
            "target_project_id": r_proj_id or None,
            "conflict_project_id": r_proj_id if (r_action == "attach_project" and r_proj_id != active_id) else None,
            "conflict_project_name": "",
            "suggested_entity": str(matched_rule.get("entity_name") or entity_hint or "").strip(),
            "suggested_clean_name": clean_name,
            "low_confidence": False,
            "matched_rule_id": matched_rule.get("id"),
            "signals": [f"learned_rule:{matched_rule.get('id')}"],
        }

    # 2. Score against active project AND every other registered project on this device
    active_profile = build_project_profile(active_proj)
    active_det_score, active_signals = score_text_against_project(
        corpus_norm, corpus_compact, corpus_tokens, active_profile
    )

    best_other_proj: dict[str, Any] | None = None
    best_other_score = 0.0
    best_other_signals: list[str] = []
    for p in all_projs:
        if str(p.get("id") or "") == active_profile["id"]:
            continue
        op = build_project_profile(p)
        osc, osig = score_text_against_project(corpus_norm, corpus_compact, corpus_tokens, op)
        if osc > best_other_score:
            best_other_score = osc
            best_other_proj = p
            best_other_signals = osig

    # 3. Cross-project conflict guard (CRITICAL on a 2-project PC!)
    if best_other_proj is not None and best_other_score >= max(thr, 0.65):
        if best_other_score >= active_det_score + 0.12 or (
            any(s.startswith(("contract_no:", "project_name:", "project_code:", "alias:")) for s in best_other_signals)
            and not any(s.startswith(("contract_no:", "project_name:", "alias:")) for s in active_signals)
        ):
            other_name = str(best_other_proj.get("name") or best_other_proj.get("id"))
            return {
                "relevant": False,
                "score": active_det_score,
                "conflict_score": best_other_score,
                "reason": tr("server.conflict_warn", project=other_name),
                "decision_kind": "cross_project_conflict",
                "forced_action": "ask",
                "target_project_id": str(best_other_proj.get("id")),
                "conflict_project_id": str(best_other_proj.get("id")),
                "conflict_project_name": other_name,
                "suggested_entity": other_name,
                "suggested_clean_name": clean_name,
                "low_confidence": False,
                "matched_rule_id": None,
                "signals": best_other_signals,
            }

    # 4. Personal / clearly non-project document check
    if personal_kind and active_det_score < 0.72:
        return {
            "relevant": False,
            "score": round(min(active_det_score, 0.15), 3),
            "reason": tr("server.unrelated_personal"),
            "decision_kind": "unrelated_file",
            "forced_action": None,
            "target_project_id": None,
            "conflict_project_id": None,
            "conflict_project_name": "",
            "suggested_entity": entity_hint,
            "suggested_clean_name": clean_name,
            "low_confidence": False,
            "matched_rule_id": None,
            "signals": [f"personal_doc:{personal_kind}"],
        }

    # 5. Blend AI relevance metadata when provided
    ai_related = meta.get("is_project_related")
    ai_rel_score = meta.get("relevance_score")
    ai_conf = meta.get("confidence")
    doc_type = str(meta.get("doc_type") or "").strip().lower()

    if ai_related is False and active_det_score < 0.72:
        sc = float(ai_rel_score) if isinstance(ai_rel_score, (int, float)) else min(active_det_score, 0.25)
        unrel_msg = (
            tr("server.unrelated_entity", entity=entity_hint)
            if entity_hint
            else tr("server.unrelated_none", project=active_profile["name"])
        )
        return {
            "relevant": False,
            "score": round(max(0.0, min(sc, thr - 0.05)), 3),
            "reason": unrel_msg,
            "decision_kind": "unrelated_file",
            "forced_action": None,
            "target_project_id": None,
            "conflict_project_id": None,
            "conflict_project_name": "",
            "suggested_entity": entity_hint,
            "suggested_clean_name": clean_name,
            "low_confidence": False,
            "matched_rule_id": None,
            "signals": active_signals + ["ai_marked_unrelated"],
        }

    combined_score = active_det_score
    if isinstance(ai_rel_score, (int, float)):
        ai_s = max(0.0, min(1.0, float(ai_rel_score)))
        if active_det_score >= thr:
            combined_score = max(active_det_score, round(active_det_score * 0.7 + ai_s * 0.3, 3))
        else:
            combined_score = round(max(active_det_score, ai_s * 0.85 if ai_related is not False else ai_s * 0.5), 3)

    # Explicit foreign entity check: document names a specific project/company that doesn't match active project at all
    if entity_hint and active_det_score < 0.40:
        sim_to_proj = max(
            entity_similarity(entity_hint, active_proj.get("name", "")),
            entity_similarity(entity_hint, active_proj.get("owner", "")),
            entity_similarity(entity_hint, active_proj.get("consultant", "")),
            entity_similarity(entity_hint, active_proj.get("contractor", "")),
        )
        if sim_to_proj < 0.45 and doc_type not in ("petty_cash", "site_photo"):
            combined_score = min(combined_score, 0.30)
            active_signals.append(f"foreign_entity:{entity_hint}")

    # 6. Site Petty-Cash / K-Net / WAMD / Site-Photo contextual safeguard
    # Shop receipts & site photos rarely carry the project title, so accept them in a
    # dedicated project inbox when there is zero foreign-project conflict and good AI confidence.
    if (
        combined_score < thr
        and doc_type in ("petty_cash", "site_photo")
        and not shared_inbox
        and ai_related is not False
        and not any(s.startswith("foreign_entity:") for s in active_signals)
    ):
        conf_val = float(ai_conf) if isinstance(ai_conf, (int, float)) else 0.75
        if conf_val >= 0.65:
            combined_score = max(combined_score, thr)
            active_signals.append(f"contextual_{doc_type}")
        else:
            return {
                "relevant": False,
                "score": round(combined_score, 3),
                "reason": tr("server.unclear_warn"),
                "decision_kind": "low_confidence",
                "forced_action": "ask",
                "target_project_id": active_profile["id"],
                "conflict_project_id": None,
                "conflict_project_name": "",
                "suggested_entity": entity_hint,
                "suggested_clean_name": clean_name,
                "low_confidence": True,
                "matched_rule_id": None,
                "signals": active_signals + ["low_confidence_guard"],
            }

    # 7. Low-confidence guard (never guess on weak OCR / borderline score)
    conf_num = float(ai_conf) if isinstance(ai_conf, (int, float)) else None
    if conf_num is not None and conf_num < 0.50:
        return {
            "relevant": False,
            "score": round(combined_score, 3),
            "reason": tr("server.unclear_warn"),
            "decision_kind": "low_confidence",
            "forced_action": "ask",
            "target_project_id": active_profile["id"] if combined_score >= thr * 0.8 else None,
            "conflict_project_id": None,
            "conflict_project_name": "",
            "suggested_entity": entity_hint,
            "suggested_clean_name": clean_name,
            "low_confidence": True,
            "matched_rule_id": None,
            "signals": active_signals + ["low_ai_confidence"],
        }

    # Shared inbox with borderline score -> ask instead of risking wrong project
    if shared_inbox and thr * 0.75 <= combined_score < thr + 0.08 and not any(
        s.startswith(("contract_no:", "project_name:", "project_code:", "alias:")) for s in active_signals
    ):
        return {
            "relevant": False,
            "score": round(combined_score, 3),
            "reason": tr("server.shared_unclear_warn"),
            "decision_kind": "low_confidence",
            "forced_action": "ask",
            "target_project_id": active_profile["id"],
            "conflict_project_id": None,
            "conflict_project_name": "",
            "suggested_entity": entity_hint,
            "suggested_clean_name": clean_name,
            "low_confidence": True,
            "matched_rule_id": None,
            "signals": active_signals + ["shared_inbox_guard"],
        }

    if combined_score >= thr:
        return {
            "relevant": True,
            "score": round(combined_score, 3),
            "reason": tr("server.matched_active"),
            "decision_kind": "relevant",
            "forced_action": None,
            "target_project_id": active_profile["id"],
            "conflict_project_id": None,
            "conflict_project_name": "",
            "suggested_entity": entity_hint,
            "suggested_clean_name": clean_name,
            "low_confidence": False,
            "matched_rule_id": None,
            "signals": active_signals,
        }

    # Unrelated file
    reason_msg = (
        tr("server.unrelated_entity", entity=entity_hint)
        if entity_hint
        else tr("server.unrelated_none", project=active_profile["name"])
    )
    return {
        "relevant": False,
        "score": round(combined_score, 3),
        "reason": reason_msg,
        "decision_kind": "unrelated_file",
        "forced_action": None,
        "target_project_id": None,
        "conflict_project_id": None,
        "conflict_project_name": "",
        "suggested_entity": entity_hint,
        "suggested_clean_name": clean_name,
        "low_confidence": False,
        "matched_rule_id": None,
        "signals": active_signals,
    }


# ── Reversible Rename-in-Place Journal + Undo ───────────────────────────────
def _unique_sibling_path(parent: Path, desired_name: str, current_path: Path) -> Path:
    cand = parent / desired_name
    if cand == current_path or not cand.exists():
        return cand
    stem = cand.stem
    ext = cand.suffix
    idx = 2
    while True:
        alt = parent / f"{stem}_{idx}{ext}"
        if alt == current_path or not alt.exists():
            return alt
        idx += 1


def load_rename_journal(journal_path: Path) -> list[dict[str, Any]]:
    data = read_json_safe(journal_path, [])
    return data if isinstance(data, list) else []


def is_in_rename_journal(journal_path: Path, file_path: Path, file_sha256: str = "") -> bool:
    """Prevent infinite re-processing of files already renamed in place."""
    entries = load_rename_journal(journal_path)
    p_res = str(Path(file_path).resolve())
    for item in entries:
        if item.get("undone") or item.get("dry_run"):
            continue
        if str(item.get("new_path") or "") == p_res:
            return True
        if file_sha256 and item.get("sha256") == file_sha256 and Path(str(item.get("new_path") or "")).exists():
            return True
    return False


def rename_in_place_with_journal(
    file_path: Path,
    new_name: str,
    journal_path: Path,
    *,
    reason: str = "unrelated_file",
    device_id: str = "",
    project_id: str = "",
    dry_run: bool = False,
) -> dict[str, Any]:
    """
    Rename an unrelated file cleanly IN ITS OWN FOLDER without touching the project
    archive or project log, recording a 100% reversible entry in `_rename_journal.json`.
    """
    src = Path(file_path).resolve()
    if not src.exists():
        raise FileNotFoundError(tr("server.file_not_found"))
    if is_file_locked(src):
        raise PermissionError(tr("server.file_in_use", name=src.name))

    # Preserve original extension strictly
    wanted = Path(new_name).name
    if Path(wanted).suffix.lower() != src.suffix.lower():
        wanted = f"{Path(wanted).stem}{src.suffix.lower()}"
    dest = _unique_sibling_path(src.parent, wanted, src)
    fhash = sha256_of_file(src)

    entry = {
        "id": f"REN-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}-{secrets.token_hex(3).upper()}",
        "action": "rename_in_place",
        "old_path": str(src),
        "new_path": str(dest),
        "old_name": src.name,
        "new_name": dest.name,
        "folder": str(src.parent),
        "sha256": fhash,
        "reason": reason,
        "project_id": project_id,
        "device_id": device_id,
        "dry_run": bool(dry_run),
        "undone": False,
        "timestamp": now_iso(),
    }
    if not dry_run and dest != src:
        os.replace(src, dest)

    entries = load_rename_journal(journal_path)
    entries.append(entry)
    atomic_write_json(journal_path, entries)
    return entry


def undo_rename_entry(journal_path: Path, entry_id: str, *, device_id: str = "") -> dict[str, Any]:
    """Revert a previous `rename_in_place` or `new_project_by_name` move to its exact original path."""
    entries = load_rename_journal(journal_path)
    target_entry: dict[str, Any] | None = None
    for item in entries:
        if item.get("id") == entry_id:
            target_entry = item
            break
    if target_entry is None:
        raise KeyError(tr("server.journal_not_found"))
    if target_entry.get("undone"):
        raise ValueError(tr("server.undo_done"))
    if target_entry.get("dry_run"):
        target_entry["undone"] = True
        target_entry["undone_at"] = now_iso()
        atomic_write_json(journal_path, entries)
        return target_entry

    cur_path = Path(str(target_entry["new_path"]))
    old_path = Path(str(target_entry["old_path"]))
    if not cur_path.exists():
        raise FileNotFoundError(tr("server.file_not_found"))
    if is_file_locked(cur_path):
        raise PermissionError(tr("server.file_in_use", name=cur_path.name))
    old_path.parent.mkdir(parents=True, exist_ok=True)
    restore_to = _unique_sibling_path(old_path.parent, old_path.name, cur_path)
    shutil.move(str(cur_path), str(restore_to))

    target_entry["undone"] = True
    target_entry["undone_at"] = now_iso()
    target_entry["undone_by_device"] = device_id
    target_entry["restored_path"] = str(restore_to)
    atomic_write_json(journal_path, entries)
    return target_entry


# ── Policy `new_project_by_name` with Fuzzy Arabic Entity Deduplication ─────
def resolve_or_create_entity_project_dir(
    entity_name: str,
    auto_projects_root: Path,
    *,
    registered_projects: list[dict[str, Any]] | None = None,
    threshold: float = 0.80,
    dry_run: bool = False,
) -> dict[str, Any]:
    """
    Find an existing independent entity folder (or registered project) matching `entity_name`
    via fuzzy Arabic/English normalization, or create a new one cleanly.
    Guarantees 'شركة النور' and 'شركة النور للمقاولات' map to the SAME single folder!
    """
    auto_root = Path(auto_projects_root)
    index_path = auto_root / "_entities_index.json"
    entities: list[dict[str, Any]] = read_json_safe(index_path, [])
    if not isinstance(entities, list):
        entities = []

    # Also discover any existing folders on disk not yet in index
    known_dirs = {str(e.get("folder") or "") for e in entities}
    if auto_root.exists():
        for child in sorted(auto_root.iterdir()):
            if child.is_dir() and not child.name.startswith(("_", ".")):
                if str(child.resolve()) not in known_dirs:
                    entities.append({
                        "name": child.name,
                        "aliases": [],
                        "folder": str(child.resolve()),
                        "created_at": now_iso(),
                    })

    clean_entity = (entity_name or "").strip() or "غير_مصنف_Unclassified"
    is_unclassified = normalize_ar(clean_entity) in (
        "غير مصنف unclassified", "غير مصنف", "unclassified", "unknown", "عام", ""
    )

    if is_unclassified:
        folder = (auto_root / "غير_مصنف").resolve()
        if not dry_run:
            folder.mkdir(parents=True, exist_ok=True)
        return {
            "entity_name": "غير مصنف",
            "folder": str(folder),
            "created": not folder.exists(),
            "merged_with_existing": folder.exists(),
            "similarity": 1.0,
            "matched_registered_project_id": None,
        }

    # 1. Check registered projects first
    if registered_projects:
        reg_match, reg_score = fuzzy_entity_match(clean_entity, registered_projects, threshold=threshold)
        if reg_match is not None:
            reg_inbox = Path(str(reg_match.get("root") or "")) / "2_Archive" / "09_General - مستندات عامة"
            if not dry_run:
                reg_inbox.mkdir(parents=True, exist_ok=True)
            return {
                "entity_name": reg_match.get("name", clean_entity),
                "folder": str(reg_inbox.resolve()),
                "created": False,
                "merged_with_existing": True,
                "similarity": reg_score,
                "matched_registered_project_id": reg_match.get("id"),
            }

    # 2. Check existing auto-created entity folders
    ent_match, ent_score = fuzzy_entity_match(clean_entity, entities, threshold=threshold)
    if ent_match is not None:
        folder = Path(str(ent_match["folder"])).resolve()
        if not dry_run:
            folder.mkdir(parents=True, exist_ok=True)
            aliases = list(ent_match.get("aliases") or [])
            if clean_entity != ent_match.get("name") and clean_entity not in aliases:
                aliases.append(clean_entity)
                ent_match["aliases"] = aliases
                atomic_write_json(index_path, entities)
        return {
            "entity_name": str(ent_match["name"]),
            "folder": str(folder),
            "created": False,
            "merged_with_existing": True,
            "similarity": ent_score,
            "matched_registered_project_id": None,
        }

    # 3. Create fresh entity project folder
    safe_folder_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", clean_entity).strip(" ._-")[:65] or "غير_مصنف"
    folder = (auto_root / safe_folder_name).resolve()
    if not dry_run:
        folder.mkdir(parents=True, exist_ok=True)
        entities.append({
            "name": clean_entity,
            "aliases": [],
            "folder": str(folder),
            "created_at": now_iso(),
        })
        atomic_write_json(index_path, entities)
    return {
        "entity_name": clean_entity,
        "folder": str(folder),
        "created": True,
        "merged_with_existing": False,
        "similarity": 1.0,
        "matched_registered_project_id": None,
    }


def route_file_to_entity_project(
    file_path: Path,
    entity_name: str,
    clean_name: str,
    auto_projects_root: Path,
    journal_path: Path,
    *,
    registered_projects: list[dict[str, Any]] | None = None,
    device_id: str = "",
    source_project_id: str = "",
    dry_run: bool = False,
) -> dict[str, Any]:
    """
    Move an unrelated file into its deduplicated entity project folder (or `غير_مصنف`)
    and log a reversible entry in `_rename_journal.json` (Undo supported!).
    """
    src = Path(file_path).resolve()
    if not src.exists():
        raise FileNotFoundError(tr("server.file_not_found"))
    if is_file_locked(src):
        raise PermissionError(tr("server.file_in_use", name=src.name))

    target_info = resolve_or_create_entity_project_dir(
        entity_name,
        auto_projects_root,
        registered_projects=registered_projects,
        dry_run=dry_run,
    )
    dest_dir = Path(target_info["folder"])
    wanted = Path(clean_name).name if clean_name else src.name
    if Path(wanted).suffix.lower() != src.suffix.lower():
        wanted = f"{Path(wanted).stem}{src.suffix.lower()}"
    dest = _unique_sibling_path(dest_dir, wanted, src) if dest_dir.exists() else (dest_dir / wanted)
    fhash = sha256_of_file(src)

    if not dry_run:
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = _unique_sibling_path(dest_dir, wanted, src)
        shutil.move(str(src), str(dest))

    entry = {
        "id": f"REN-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}-{secrets.token_hex(3).upper()}",
        "action": "new_project_by_name",
        "entity_name": target_info["entity_name"],
        "merged_with_existing": target_info["merged_with_existing"],
        "created_entity": target_info["created"],
        "old_path": str(src),
        "new_path": str(dest),
        "old_name": src.name,
        "new_name": dest.name,
        "folder": str(dest_dir),
        "sha256": fhash,
        "reason": f"new_project_by_name:{target_info['entity_name']}",
        "project_id": source_project_id,
        "device_id": device_id,
        "dry_run": bool(dry_run),
        "undone": False,
        "timestamp": now_iso(),
    }
    entries = load_rename_journal(journal_path)
    entries.append(entry)
    atomic_write_json(journal_path, entries)
    return {**entry, "entity_info": target_info}
