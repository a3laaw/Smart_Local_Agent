#!/usr/bin/env python3
"""
lint_strings.py — Automated UX-copy linter for Smart_Local_Agent_Free (strings_ar.json & source code).

Fails (exit code 1) if ANY rule is violated:
  1. Missing `msa` or `eg` translation, missing `type`, or mismatched `{placeholders}`.
  2. Word/sentence length limits:
       - btn:      <= 3 words, single line
       - label:    <= 4 words, single line
       - tab:      <= 4 words, single line
       - title:    <= 5 words, single line
       - progress: <= 5 words, single line
       - status:   <= 4 words, single line
       - option:   <= 6 words, single line
       - tooltip:  <= 12 words, single line
       - tip:      <= 14 words, single line
       - toast:    single line (<= 12 words)
       - empty:    1 sentence (<= 10 words)
       - warn:     <= 2 sentences
       - error:    <= 2 sentences
       - confirm:  <= 2 sentences
  3. Forbidden technical jargon in ANY user-facing string (JSON, API, SHA-256, hash, watcher,
     inbox, registry, schema, engine, token, Dry-Run, VBA, CSRF, Origin, Host, HTTP, symlink,
     CLI, CDN, glob, flat, profile, URL, PID, OCR, بروفايل, بصمة, حقل يتيم, مرساة, نقطة النهاية, ...).
  4. Colloquial Egyptian words inside `msa` (ده, دي, دلوقتي, عايز, شغال, لسه, بتاع, كده, مفيش, عشان, ...).
  5. Pompous / archaic words in `msa` or `eg` (ريثما, سيّما, حيث إن, بناءً عليه, يتعيّن, ينبغي, آنفاً, البتّ, ...).
  6. Gendered imperatives or vague dialog buttons (`موافق`, `إلغاء`, `اختر`, `ارفع`, `افتح`, `احفظ`, `اضغط`, ...) in `btn`/`title`/`label`/`tab`,
     and second-person pronouns (`لديك`, `تأكد`, `أغلِقها`, `افتحه`) in `msa`.
  7. Optional `--check-code <doc_tools_dir>`: fails if hardcoded user-facing Arabic remains in
     `dashboard_ui.html`, `project_dashboard.py`, or `scan_watcher.py`.
"""
from __future__ import annotations

import argparse
import ast
import json
import re
import sys
from pathlib import Path

PLACEHOLDER_RE = re.compile(r"\{[a-zA-Z0-9_]+\}")
DIACRITICS_RE = re.compile(r"[\u064B-\u065F\u0670\u0640]")
ARABIC_RE = re.compile(r"[\u0600-\u06FF]")
FUNC_CALL_RE = re.compile(r"\b[a-z]+_[a-z0-9_]+\s*\(")
ERR_CODE_RE = re.compile(r"(?<!\.)\b(?:ERR_[A-Z_]+|0x[0-9A-Fa-f]+|HTTP\s*\d{3}|40[013459]|413|422|50[0234])\b")

WORD_LIMITS: dict[str, int] = {
    "btn": 3,
    "label": 4,
    "tab": 4,
    "title": 5,
    "progress": 5,
    "status": 4,
    "option": 6,
    "tooltip": 12,
    "tip": 14,
    "toast": 12,
    "empty": 10,
}

SENTENCE_LIMITS: dict[str, int] = {
    "empty": 1,
    "error": 2,
    "warn": 2,
    "confirm": 2,
}

SINGLE_LINE_TYPES = {
    "btn", "label", "tab", "title", "progress", "status", "option", "tooltip", "tip", "toast", "empty"
}

# Allowed site acronyms/units (not treated as technical jargon)
ALLOWED_SITE_TERMS = {
    "MAR", "MIR", "WIR", "IR", "SD", "MS", "PQ", "IPC", "VO", "PO",
    "RFI", "NCR", "OUT", "IN", "PC", "RPT", "GEN", "REV", "R00",
    "CIV", "ARCH", "STR", "EL", "PL", "HVAC", "MEC", "FF", "ELEV",
    "A", "B", "C", "D", "Excel", "Word", "PDF", "Windows", "Ollama", "Gemini",
}

FORBIDDEN_TECH_EN = [
    "json", "api", "sha-256", "sha256", "hash", "watcher", "inbox", "1_inbox",
    "2_archive", "3_project_logs", "4_project_templates", "00_editable_originals",
    "registry", "schema", "engine", "token", "dry-run", "dry_run", "vba",
    "csrf", "origin", "host", "http", "https", "symlink", "cli", "cdn",
    "glob", "gitignore", "flat", "multi_rev", "profile", "tplmap", "logmap",
    "url", "uri", "pid", "ocr", "endpoint", "payload", "regex", "base64",
    "libreoffice", "png", "device_id", "learned_rules", "rename_journal",
]

FORBIDDEN_TECH_AR = [
    "بروفايل", "البروفايل", "بصمة", "البصمة", "هاش", "واتشر", "إنبوكس", "انبوكس",
    "ريجستري", "سكيما", "توكن", "نقطة النهاية", "حقل يتيم", "حقول يتيمة",
    "الحقول اليتيمة", "مرساة", "المرساة", "مراسي", "محرك", "المحرك",
    "رابط رمزي", "محارف", "خادم", "الخادم", "الطرفية", "كود الخطأ",
]

COLLOQUIAL_IN_MSA = [
    "ده", "دي", "دول", "دلوقتي", "عايز", "عاوز", "عايزة", "شغال", "شغالة",
    "لسه", "لسة", "بتاع", "بتاعة", "بتوع", "كده", "كدا", "إزاي", "ازاي",
    "مفيش", "مافيش", "عشان", "علشان", "خالص", "برضه", "برضو", "ليه",
    "إيه", "ايه", "أوي", "اوي", "حاجة", "حاجات", "كمان", "خلاص",
    "فين", "مين", "إحنا", "احنا", "جوه", "بره", "أهو", "اهو",
    "مستني", "مستنية", "تظبيط", "اتحفظ", "اتقفل", "اتغير", "اتسجل", "اتمسح",
]

ARCHAIC_POMPOUS = [
    "ريثما", "سيما", "لا سيما", "حيث إن", "حيث ان", "بناء عليه",
    "إذ إن", "اذ ان", "يتعين", "ينبغي", "آنفا", "المزمع", "تلقاء",
    "بيد أن", "علاوة على", "بموجب", "حيال", "يتسنى", "مفاده",
    "السالف", "المشار إليه", "البت", "الآنف", "حري", "قيد الإجراء",
]

VAGUE_BUTTONS = {"موافق", "إلغاء", "الغاء", "حسنا", "نعم", "لا", "تم فقط"}

GENDERED_IMPERATIVES_START = {
    "اختر", "اختاري", "ارفع", "ارفعي", "افتح", "افتحي", "احفظ", "احفظي",
    "أضف", "اضف", "أضيفي", "أدخل", "ادخل", "أدخلي", "اكتب", "اكتبي",
    "اضغط", "اضغطي", "أغلق", "اغلق", "أغلقي", "راجع", "راجعي",
    "حدد", "حددي", "تأكد", "تأكدي", "تحقق", "تحققي", "انقر", "انقري",
    "اسحب", "اسحبي", "ابدأ", "ابدئي", "جرّب", "جرب", "أعد", "اعد",
}

SECOND_PERSON_MSA = [
    "لديك", "عندك", "أمامك", "بانتظارك", "تأكد", "تأكدي", "تحقق", "تحققي",
    "افتحه", "افتحها", "ارفعه", "ارفعها", "أغلقها", "أغلقه", "اغلقها",
    "قرارك", "موافقتك", "حفظك", "نماذجك", "مشاريعك", "ملفاتك", "أنت",
]


def strip_diacritics(s: str) -> str:
    return DIACRITICS_RE.sub("", s)


def count_words(text: str) -> int:
    # Replace placeholders with a single word token, ignore standalone punctuation/slashes
    cleaned = PLACEHOLDER_RE.sub("X", text)
    cleaned = re.sub(r"[—–·•|/\\()\[\]:،,.؟!+«»\"']+", " ", cleaned)
    tokens = [t for t in cleaned.split() if t.strip()]
    return len(tokens)


def count_sentences(text: str) -> int:
    cleaned = PLACEHOLDER_RE.sub("X", text).strip()
    if not cleaned:
        return 0
    # Protect decimal KWD format like 0.000 and extensions like .xlsx
    cleaned = re.sub(r"\d+\.\d+", "NUM", cleaned)
    cleaned = re.sub(r"\.[a-zA-Z0-9]{2,5}\b", "EXT", cleaned)
    parts = [p.strip() for p in re.split(r"[.؟!]+", cleaned) if p.strip()]
    return max(1, len(parts))


def has_whole_term(text_plain: str, term: str) -> bool:
    t = strip_diacritics(term)
    return bool(re.search(rf"(?<!\w){re.escape(t)}(?!\w)", text_plain))


def lint_dictionary(path: Path) -> list[str]:
    errors: list[str] = []
    raw = json.loads(path.read_text(encoding="utf-8"))
    strings = raw.get("strings") if isinstance(raw.get("strings"), dict) else raw

    if not strings:
        return ["strings_ar.json فارغ أو لا يحتوي على مفاتيح."]

    for key, entry in strings.items():
        if key.startswith("_"):
            continue
        if not isinstance(entry, dict):
            errors.append(f"[{key}] القيمة ليست كائناً صالحاً.")
            continue

        kind = str(entry.get("type") or "").strip()
        if kind not in (set(WORD_LIMITS) | set(SENTENCE_LIMITS)):
            errors.append(f"[{key}] نوع النص غير معروف أو مفقود: {kind!r}")
            continue

        msa = entry.get("msa")
        eg = entry.get("eg")
        if not isinstance(msa, str) or not msa.strip():
            errors.append(f"[{key}] نص الفصحى المبسطة (msa) مفقود.")
            continue
        if not isinstance(eg, str) or not eg.strip():
            errors.append(f"[{key}] نص العامية المصرية (eg) مفقود.")
            continue

        # Placeholder parity
        ph_msa = set(PLACEHOLDER_RE.findall(msa))
        ph_eg = set(PLACEHOLDER_RE.findall(eg))
        if ph_msa != ph_eg:
            errors.append(f"[{key}] عدم تطابق المتغيرات بين الأسلوبين: msa={ph_msa} vs eg={ph_eg}")

        for style, text in (("msa", msa), ("eg", eg)):
            no_ph = PLACEHOLDER_RE.sub("", text)
            low = no_ph.lower()
            plain_ar = strip_diacritics(no_ph)

            # Single-line check
            if kind in SINGLE_LINE_TYPES and "\n" in text:
                errors.append(f"[{key}.{style}] ({kind}) يجب أن يكون سطراً واحداً فقط: {text!r}")

            # Word limit check
            if kind in WORD_LIMITS:
                wc = count_words(text)
                lim = WORD_LIMITS[kind]
                if wc > lim:
                    errors.append(f"[{key}.{style}] ({kind}) تجاوز حد الكلمات ({wc} > {lim}): {text!r}")

            # Sentence limit check
            if kind in SENTENCE_LIMITS:
                sc = count_sentences(text)
                slim = SENTENCE_LIMITS[kind]
                if sc > slim:
                    errors.append(f"[{key}.{style}] ({kind}) تجاوز حد الجمل ({sc} > {slim}): {text!r}")

            # Forbidden technical jargon (English & Arabic)
            for bad_en in FORBIDDEN_TECH_EN:
                if re.search(rf"(?<![a-z0-9]){re.escape(bad_en)}(?![a-z0-9])", low):
                    errors.append(f"[{key}.{style}] مصطلح تقني ممنوع ({bad_en!r}): {text!r}")
            for bad_ar in FORBIDDEN_TECH_AR:
                if has_whole_term(plain_ar, bad_ar):
                    errors.append(f"[{key}.{style}] مصطلح تقني ممنوع ({bad_ar!r}): {text!r}")

            # Function names / raw error codes
            if FUNC_CALL_RE.search(no_ph) or ERR_CODE_RE.search(no_ph):
                errors.append(f"[{key}.{style}] يحتوي على كود خطأ أو اسم دالة: {text!r}")

            # Archaic / pompous words
            for arc in ARCHAIC_POMPOUS:
                if has_whole_term(plain_ar, arc):
                    errors.append(f"[{key}.{style}] لفظ متقعّر/مهجور ({arc!r}): {text!r}")

            # Vague button labels
            if kind == "btn" and plain_ar.strip() in VAGUE_BUTTONS:
                errors.append(f"[{key}.{style}] زر عام غير محدد الفعل ({text!r}) — سمِّ الفعل بدقة أو استخدم «تراجع».")

            # MSA-specific purity checks (zero dialect, zero gendered imperative, zero 2nd person)
            if style == "msa":
                for col in COLLOQUIAL_IN_MSA:
                    if has_whole_term(plain_ar, col):
                        errors.append(f"[{key}.msa] كلمة عامية داخل الفصحى المبسطة ({col!r}): {text!r}")
                words = [w for w in re.split(r"[\s،,.؟!:()«»—–/]+", plain_ar) if w]
                if words and words[0] in GENDERED_IMPERATIVES_START:
                    errors.append(f"[{key}.msa] صيغة أمر غير محايدة في البداية ({words[0]!r}): {text!r}")
                for sp in SECOND_PERSON_MSA:
                    if has_whole_term(plain_ar, sp):
                        errors.append(f"[{key}.msa] ضمير مخاطب غير محايد ({sp!r}): {text!r}")

    return errors


# Scope-excluded non-UI constructs (formal folder/file names, AI prompts, Excel log headers, document text matchers)
CODE_ALLOW_SUBSTRINGS = (
    "01_Incoming", "02_Outgoing", "03_Submittals", "04_Contracts", "05_Petty_Cash",
    "06_Site_Reports", "06_Permits", "07_Site_Photos", "08_Custom", "09_General",
    "01_Submittals", "02_Petty_Cash", "03_Invoices", "04_Letters", "05_Contracts_IPCs",
    "00_Editable_Originals", "00_WhatsApp_Chats", "_01_Processed_Originals",
    "_02_Duplicates", "_03_Needs_Review", "_قبل التعديل", "Master_Archive_Log",
    "99_Unrelated_", "_Auto_Projects_", "Smart_Projects_", "غير_مصنف",
    "CIV - ", "ARCH - ", "STR - ", "EL - ", "PL - ", "HVAC - ", "MEC - ", "FF - ", "ELVE - ", "ELV - ",
    "JSON", "relevance_reason", "[,،;", ".:-_،,", "HYPERLINK", '","فتح الملف")',
    "هذا ملف أصلي قابل للتعديل", "تم تجهيزه قبل الطباعة/الختم",
    "تم الاستلام / مختوم", "صادر — بانتظار ختم الاستلام", "Under Review (مقدم",
)

SCOPE_EXCLUDED_ASSIGN_NAMES = {
    # AI prompt profiles, formal Excel register headers, & document-parsing token tables (scope-excluded)
    "DOMAIN_PROFILES", "DEFAULT_CONFIG", "CATEGORY_FOLDERS", "DISC_LABELS",
    "_GENERIC_PARTY_WORDS", "SHEETS_SPEC",
    "_AR_DIGIT_TRANS", "_AR_CHAR_TRANS", "_ENTITY_STOPWORDS",
    "_GENERIC_PROJECT_WORDS", "_PERSONAL_DOC_KEYWORDS",
    "UNRELATED_SHEET_NAME", "DISC_CANON", "_COL_PATTERNS",
    "_AR_DAY_NAMES", "_SECTION_TOKEN_RULES", "_LABEL_HINT_MAP", "_REV_HEADER_RE",
    "_LETTER_STOPWORDS", "_CONCEPT_CLUSTERS", "_DISTINCT_TRADES", "_REMARK_TAG_RE",
}

SCOPE_EXCLUDED_FUNC_NAMES = {
    # Internal AI prompts, formal archive paths, Excel register writers, & document header/status parsers
    "build_classification_prompt", "suggest_fields_with_ai",
    "_norm_arabic_text", "_core_party_tokens", "resolve_party_folder", "compute_destination",
    "ensure_master_workbook", "_try_update_existing_native_row", "build_excel_records",
    "extract_entity_candidate", "build_unrelated_filename", "resolve_or_create_entity_project_dir",
    "_infer_sheet_route", "_infer_party_role", "_map_status_to_sheet_vocab", "_is_totals_row",
    "_match_col_role", "inspect_workbook_to_logmap", "read_project_log_summary", "_is_moh_authority", "_has_consultant_stamp", "_is_pending_status_text", "_map_moh_status_to_vocab",
    "record_optional_unrelated_entry", "suggest_fields_deterministic", "detect_multi_rev_groups",
    "_infer_party_role_from_hints", "_arabic_stem_variants", "extract_topic_signature", "score_letter_pair",
    "build_multipage_split_prompt", "_heuristic_split_pages", "_sync_incoming_reply_status",
}

CHECKED_PY_FILES = (
    "project_dashboard.py", "scan_watcher.py", "security_utils.py",
    "decisions.py", "active_project.py", "inbox_pipeline.py",
    "relevance.py", "project_log.py", "template_profile.py", "letter_links.py",
)


def lint_source_code(doc_tools_dir: Path) -> list[str]:
    errors: list[str] = []
    html_p = doc_tools_dir / "dashboard_ui.html"
    if html_p.exists():
        html = html_p.read_text(encoding="utf-8")
        stripped = re.sub(r"<!--.*?-->", "", html, flags=re.S)
        stripped = re.sub(r"/\*.*?\*/", "", stripped, flags=re.S)
        stripped = re.sub(r"//[^\n]*", "", stripped)
        for ln_no, line in enumerate(stripped.splitlines(), 1):
            if ARABIC_RE.search(line):
                errors.append(f"[dashboard_ui.html:{ln_no}] نص عربي مكتوب مباشرةً خارج strings_ar.json: {line.strip()[:90]}")

    for py_name in CHECKED_PY_FILES:
        py_p = doc_tools_dir / py_name
        if not py_p.exists():
            continue
        src = py_p.read_text(encoding="utf-8")
        tree = ast.parse(src)
        skip_ids: set[int] = set()
        for node in ast.walk(tree):
            if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
                if node.body and isinstance(node.body[0], ast.Expr) and isinstance(node.body[0].value, ast.Constant):
                    skip_ids.add(id(node.body[0].value))
                if isinstance(node, ast.FunctionDef) and node.name in SCOPE_EXCLUDED_FUNC_NAMES:
                    for sub in ast.walk(node):
                        skip_ids.add(id(sub))
            elif isinstance(node, (ast.Assign, ast.AnnAssign)):
                targets = node.targets if isinstance(node, ast.Assign) else [node.target]
                for t in targets:
                    if isinstance(t, ast.Name) and t.id in SCOPE_EXCLUDED_ASSIGN_NAMES:
                        for sub in ast.walk(node):
                            skip_ids.add(id(sub))
        for node in ast.walk(tree):
            if id(node) in skip_ids:
                continue
            if isinstance(node, ast.Constant) and isinstance(node.value, str) and ARABIC_RE.search(node.value):
                val = node.value.strip()
                if len(val) < 2 or val in (
                    "عام", "صادر", "وارد", "د.ك", "غير مصنف", "ال",
                    "تقديم", "تاريخ التقديم", "معتمد", "معتمد بملاحظات",
                    "مرفوض", "ملغي", "الاجمالي", "المجموع",
                ):
                    continue
                if any(sub in val for sub in CODE_ALLOW_SUBSTRINGS):
                    continue
                if re.search(r"\\u06", repr(val)):
                    continue
                errors.append(f"[{py_name}:{node.lineno}] نص عربي مباشر خارج strings_ar.json: {val[:80]!r}")
    return errors


def main() -> int:
    ap = argparse.ArgumentParser(description="فحص جودة نصوص واجهة Smart_Local_Agent_Free")
    ap.add_argument("strings_file", help="مسار ملف strings_ar.json")
    ap.add_argument("--check-code", default="", help="مسار doc_tools للتحقق من خلو الكود وHTML من النصوص العربية المباشرة")
    args = ap.parse_args()

    sp = Path(args.strings_file)
    data = json.loads(sp.read_text(encoding="utf-8"))
    total_keys = sum(1 for k in (data.get("strings") or data) if not k.startswith("_"))

    errs = lint_dictionary(sp)
    if args.check_code:
        errs.extend(lint_source_code(Path(args.check_code)))

    if errs:
        print(f"FAIL: وُجدت {len(errs)} مخالفة في النصوص (من أصل {total_keys} مفتاح):")
        for e in errs:
            print("  -", e)
        return 1

    print(f"PASS: جميع المفاتيح ({total_keys}) اجتازت الفحص بنجاح (0 مخالفات طول أو مصطلحات تقنية أو عامية أو نقص ترجمة).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
