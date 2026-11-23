#!/usr/bin/env python3
"""
scan_watcher.py — المراقب التلقائي لمجلد السكانر (Auto Scan-to-Archive & Excel Agent)
مجاني 100% (يعمل مع Google Gemini CLI المجاني أو Ollama أوفلاين).

طريقة العمل:
  ماكينة السكانر / الموبايل -> مجلد Scan_Inbox
    1) ينتظر اكتمال كتابة الملف وفك قفل السكانر عنه بالكامل.
    2) يفحص بصمة SHA-256 فوراً لمنع تكرار نفس الورقة بدون استهلاك أي طلب.
    3) يحوّل صفحات الـ PDF أو الصورة إلى PNG واضحة محلياً (مع تكبير منطقة الأختام).
    4) يقرأ المستند بصرياً (في وضع القراءة الآمن فقط) ويستخرج المشروع والنوع والرقم والتاريخ والمبلغ والختم.
    5) ينسخ الملف بالاسم القياسي النظيف داخل المجلد الصحيح (اعتمادات / عُهدة / فواتير / صادر ووارد / عقود ومستخلصات).
    6) يحتفظ بالأصل الخام في _01_Processed_Originals (حماية تلقائية من الحذف) ويفرغ Scan_Inbox للسكان التالي.
    7) يحدّث السجل الشامل Master_Archive_Log.xlsx تلقائياً مع رابط HYPERLINK مباشر (وإن كان الإكسل مفتوحاً يحفظ الصف في طابور انتظار ويكتبه فور إغلاقك للملف!).
"""
import argparse
import base64
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import sys
import threading
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import pymupdf as fitz  # type: ignore
except ImportError:
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None  # type: ignore

from PIL import Image
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from i18n_ar import tr as _tr

SUPPORTED_EXTS = {".pdf", ".jpg", ".jpeg", ".jfif", ".png", ".tif", ".tiff", ".bmp", ".webp", ".gif"}
IMAGE_EXTS = SUPPORTED_EXTS - {".pdf"}
SCANNED_EXTS = SUPPORTED_EXTS
NATIVE_DRAFT_EXTS = {".docx", ".xlsx", ".xlsm"}
NATIVE_DOC_EXTS = {".docx", ".xlsx", ".xlsm"}
CHAT_LOG_EXTS = {".txt"}
TEMP_EXTS = {".tmp", ".part", ".partial", ".crdownload", ".lock", ".lck", ".dat"}
INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

DOMAIN_PROFILES = {
    "contracting": {
        "label": "مقاولات وتنفيذ مشاريع",
        "prompt_hint": "النشاط: شركة مقاولات وتنفيذ مشاريع (اعتمادات MAR/MIR/WIR/Shop Drawings، عُهدة موقع وإيصالات K-Net/ومض، فواتير موردين، كتب صادرة وواردة، عقود ومستخلصات مقاولي باطن IPC وأوامر تغييرية VO وأوامر شراء PO، تراخيص، وصور تقدم أعمال الموقع من واتساب).",
    },
    "consulting": {
        "label": "مكتب استشارات هندسية وإشراف",
        "prompt_hint": "النشاط: مكتب استشارات هندسية وتصميم وإشراف (مراجعة اعتمادات ومخططات MAR/MIR/WIR/Shop Drawings، استفسارات فنية RFI وتقارير عدم مطابقة NCR وتعليمات موقع SI -> تصنف submittal مع كتابة النوع RFI/NCR/SI، تراخيص وموافقات البلدية وقوة الإطفاء KFF والكهرباء والماء MEW والوزارات -> تصنف permit، عقود التصميم والإشراف ودفعات الأتعاب المرحلية -> contract_ipc، الكتب الرسمية ومحاضر الاجتماعات -> letter، فواتير ومصاريف المكتب -> invoice/petty_cash، وصور معاينات الإشراف والموقع من واتساب -> site_photo).",
    },
    "trading": {
        "label": "تجارة وتوريدات ومحاسبة",
        "prompt_hint": "النشاط: شركة تجارة وتوريدات ومحاسبة (فواتير مبيعات ومشتريات وموردين وعملاء -> invoice بمجلد لكل عميل/مورد، إيصالات نقدية وK-Net وتحويلات ومض Wamd وسندات قبض/صرف وعُهدة -> petty_cash، عقود توريد وعروض أسعار وأوامر شراء PO وكفالات -> contract_ipc، كتب رسمية وإشعارات تسليم Delivery Notes -> letter، رخص تجارية/جمركية/شهادات منشأ وبوالص شحن -> permit، صور بضائع/مخازن من واتساب -> site_photo).",
    },
    "custom": {
        "label": "إداري / عام (أقسام مخصصة)",
        "prompt_hint": "النشاط: مكتب إداري / نشاط مخصص — طابق المستند أولاً مع الأقسام المخصصة التي حددها المستخدم (doc_type='custom' وضع اسم القسم في custom_category)، أو استخدم الفواتير/العُهدة/الكتب/العقود/التراخيص/الصور حسب الأنسب.",
    },
}

DEFAULT_CONFIG = {
    "inbox_folder": str(Path.home() / "Desktop" / "Smart_Scan" / "1_Scan_Inbox"),
    "drafts_folder": str(Path.home() / "Desktop" / "Smart_Scan" / "2_Drafts_Word_Excel"),
    "archive_folder": str(Path.home() / "Desktop" / "Smart_Scan" / "3_Organized_Archive"),
    "templates_folder": str(Path.home() / "Desktop" / "Smart_Scan" / "4_Project_Templates"),
    "folder_layout": "by_type",  # "by_type" أو "by_project_then_type"
    "domain": "contracting",     # contracting | consulting | trading | custom
    "custom_categories": [],     # أقسام إضافية يحددها المستخدم لأي مجال
    "default_project": "عام",
    "default_currency": "KWD",
    "known_projects": [],
    "engine": "gemini",          # "gemini" (سحابي مجاني 1000 طلب/يوم) أو "ollama" (محلي أوفلاين 100%)
    "gemini_model": "",          # اتركه فارغاً للاختيار التلقائي الذكي في Gemini CLI
    "ollama_url": "http://127.0.0.1:11434/api/chat",
    "ollama_model": "qwen3-vl:8b",
    "poll_seconds": 4,
    "settle_seconds": 6,
    "min_confidence": 0.58,
    "render_dpi": 150,
    "stamp_zoom_dpi": 200,
    "convert_images_to_pdf": False,
}

# ألوان وتنسيق السجل الشامل
NAVY = "0F2942"
SLATE = "1E4976"
SOFT = "EFF6FF"
ZEBRA = "F8FAFC"
LINK_CLR = "0D6EFD"
THIN = Side(style="thin", color="334155")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CATEGORY_FOLDERS = {
    "submittal": "01_Submittals - الاعتمادات والمخططات",
    "petty_cash": "02_Petty_Cash - العهدة والإيصالات",
    "invoice": "03_Invoices - الفواتير",
    "letter": "04_Letters - الصادر والوارد",
    "contract_ipc": "05_Contracts_IPCs - العقود والمستخلصات",
    "permit": "06_Permits - التراخيص والموافقات",
    "site_photo": "07_Site_Photos - صور الموقع والواتساب",
    "custom": "08_Custom - أقسام مخصصة",
    "other": "09_General - مستندات عامة",
}


def log(msg: str, log_file: Path | None = None):
    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    if log_file:
        try:
            log_file.parent.mkdir(parents=True, exist_ok=True)
            with log_file.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass


def clean_seg(text: str, max_len: int = 55) -> str:
    if not text:
        return ""
    s = INVALID_CHARS.sub("-", str(text))
    s = re.sub(r"\s+", " ", s).strip().strip(".-_ ")
    return s[:max_len].strip(".-_ ")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def is_file_unlocked_and_valid(path: Path) -> tuple[bool, int]:
    """يتأكد أن السكانر أنهى كتابة الملف وأغلق القفل وأن الـ PDF/الصورة سليمة وغير مقطوعة.
    يدعم مجلدات الشبكة في Windows Server حتى عند القراءة فقط (Read-Only Share)."""
    if not path.exists():
        return False, 0
    try:
        size = path.stat().st_size
        if size == 0:
            return False, 0
        try:
            with path.open("r+b"):
                pass
        except OSError as open_exc:
            if getattr(open_exc, "winerror", None) in (32, 33):
                return False, 0
            # Fallback for read-only SMB shares: verify file is genuinely readable
            with path.open("rb") as rf:
                rf.read(16)
    except OSError:
        return False, 0

    ext = path.suffix.lower()
    if ext == ".pdf":
        if fitz is None:
            try:
                with path.open("rb") as pf:
                    hdr = pf.read(8)
                return hdr.startswith(b"%PDF"), 1
            except Exception:
                return False, 0
        try:
            doc = fitz.open(str(path))
            pages = len(doc)
            if pages < 1:
                doc.close()
                return False, 0
            _ = doc[0].rect
            doc.close()
            return True, pages
        except Exception:
            return False, 0
    elif ext in NATIVE_DOC_EXTS:
        try:
            with zipfile.ZipFile(path, "r") as zf:
                bad = zf.testzip()
                return (bad is None), 1
        except Exception:
            return False, 0
    elif ext in CHAT_LOG_EXTS:
        return True, 1
    else:
        try:
            with Image.open(path) as im:
                im.verify()
            return True, 1
        except Exception:
            return False, 0


def extract_native_office_text(src: Path) -> str:
    """يستخرج النص الكامل والجداول من ملفات Word (.docx) و Excel (.xlsx/.xlsm) الأصلية لقراءتها وتصنيفها تلقائياً."""
    ext = src.suffix.lower()
    lines: list[str] = []
    if ext == ".docx":
        from docx import Document
        doc = Document(str(src))
        for sec in doc.sections:
            for p in sec.header.paragraphs:
                if p.text.strip():
                    lines.append(f"[Header] {p.text.strip()}")
            for tbl in sec.header.tables:
                for row in tbl.rows:
                    vals = list(dict.fromkeys(c.text.strip() for c in row.cells if c.text.strip()))
                    if vals:
                        lines.append("[HeaderTbl] " + " | ".join(vals))
        for p in doc.paragraphs:
            if p.text.strip():
                lines.append(p.text.strip())
        for tbl in doc.tables:
            for row in tbl.rows:
                vals = list(dict.fromkeys(c.text.strip() for c in row.cells if c.text.strip()))
                if vals:
                    lines.append(" | ".join(vals))
    elif ext in {".xlsx", ".xlsm"}:
        wb = openpyxl.load_workbook(str(src), data_only=True, read_only=True)
        for sname in wb.sheetnames[:5]:
            ws = wb[sname]
            lines.append(f"[Sheet: {sname}]")
            row_cnt = 0
            for row in ws.iter_rows(values_only=True):
                vals = [str(v).strip() for v in row[:18] if v is not None and str(v).strip() != ""]
                if vals:
                    lines.append(" | ".join(vals))
                    row_cnt += 1
                    if row_cnt >= 65:
                        break
        wb.close()
    return "\n".join(lines)[:5500]


def prepare_visual_pages(src: Path, job_dir: Path, dpi: int = 150, stamp_dpi: int = 200) -> tuple[list[Path], str, int]:
    """يحوّل المستند الممسوح إلى صور PNG واضحة (مع تكبير منطقة الأختام) ويستخرج أي نص رقمي."""
    job_dir.mkdir(parents=True, exist_ok=True)
    images: list[Path] = []
    embedded_text = ""
    page_count = 1
    ext = src.suffix.lower()

    if ext == ".pdf":
        if fitz is None:
            raise RuntimeError(_tr("watcher.pymupdf_missing"))
        doc = fitz.open(str(src))
        page_count = len(doc)
        texts = []
        # نلتقط الصفحة 1، والصفحة 2 (المرفق/خطاب الاعتماد)، والصفحة الأخيرة إن زادت الصفحات (التوقيعات/الإجمالي)
        idxs = [0]
        if page_count >= 2:
            idxs.append(1)
        if page_count >= 3:
            idxs.append(page_count - 1)

        for idx in idxs:
            page = doc[idx]
            t = page.get_text("text").strip()
            if t:
                texts.append(f"[Page {idx+1}]\n{t[:1200]}")
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0), alpha=False)
            out_p = job_dir / f"page_{idx+1:02d}.png"
            pix.save(str(out_p))
            images.append(out_p)

        # تكبير النصف السفلي للصفحة الأولى لقراءة كود الاعتماد (A/B/C/D) والأختام والتاريخ اليدوي بدقة عالية
        p0 = doc[0]
        r = p0.rect
        clip = fitz.Rect(r.x0, r.y0 + r.height * 0.50, r.x1, r.y1)
        pix_stamp = p0.get_pixmap(matrix=fitz.Matrix(stamp_dpi / 72.0, stamp_dpi / 72.0), clip=clip, alpha=False)
        stamp_p = job_dir / "page_01_stamps_zoom.png"
        pix_stamp.save(str(stamp_p))
        images.append(stamp_p)
        doc.close()
        embedded_text = "\n".join(texts).strip()
    else:
        # ملف صورة (JPG/PNG/TIFF متعدد الصفحات من السكانر)
        with Image.open(src) as im:
            n_frames = getattr(im, "n_frames", 1)
            page_count = n_frames
            for idx in range(min(n_frames, 3)):
                try:
                    im.seek(idx)
                except EOFError:
                    break
                rgb = im.convert("RGB")
                max_side = max(rgb.size)
                if max_side > 2200:
                    scale = 2200.0 / max_side
                    rgb = rgb.resize((int(rgb.width * scale), int(rgb.height * scale)), Image.LANCZOS)
                out_p = job_dir / f"page_{idx+1:02d}.png"
                rgb.save(str(out_p), "PNG")
                images.append(out_p)

    return images, embedded_text, page_count


def build_classification_prompt(
    image_names: list[str],
    orig_name: str,
    subfolder_hint: str,
    page_count: int,
    embedded_text: str,
    existing_projects: list[str],
    default_currency: str,
    include_at_refs: bool = True,
    domain: str = "contracting",
    custom_categories: list[str] | None = None,
    active_project: dict | None = None,
) -> str:
    refs = " ".join(f"@./{n}" for n in image_names) if include_at_refs else ""
    dom_info = DOMAIN_PROFILES.get(domain or "contracting", DOMAIN_PROFILES["contracting"])
    dom_line = f"- مجال العمل: {dom_info['prompt_hint']}\n"
    active_line = ""
    if isinstance(active_project, dict) and active_project.get("name"):
        ap_bits = [f"الاسم: {active_project.get('name')}"]
        if active_project.get("code"):
            ap_bits.append(f"الكود: {active_project.get('code')}")
        if active_project.get("contract_no"):
            ap_bits.append(f"رقم العقد: {active_project.get('contract_no')}")
        if active_project.get("owner"):
            ap_bits.append(f"المالك: {active_project.get('owner')}")
        if active_project.get("consultant"):
            ap_bits.append(f"الاستشاري: {active_project.get('consultant')}")
        if active_project.get("contractor"):
            ap_bits.append(f"المقاول: {active_project.get('contractor')}")
        active_line = f"- المشروع النشط الحالي (تحقق بدقة هل الملف يخص هذا المشروع أم مشروعاً/جهة أخرى أم هو مستند شخصي غير مرتبط): {' | '.join(ap_bits)}\n"
    cats = [c.strip() for c in (custom_categories or []) if str(c).strip()]
    custom_line = (
        f"- أقسام مخصصة أنشأها المستخدم (إذا كان المستند ينتمي لأحدها اجعل doc_type=\"custom\" وضع اسم القسم كما هو في custom_category): {', '.join(cats)}\n"
        if cats else ""
    )
    is_wa = "whatsapp" in orig_name.lower() or "wa0" in orig_name.lower()
    wa_line = "- ملاحظة واتساب: الملف وارد من WhatsApp؛ إذا كان صورة إيصال K-Net أو تحويل ومض (Wamd) أو تحويل بنكي أو فاتورة صُنّف petty_cash أو invoice واستخرج المبلغ بدقة، أما إذا كان صورة ميدانية لموقع/أعمال/بضاعة وليست مستنداً ورقياً فاجعل doc_type=\"site_photo\" وصف ما يظهر في الصورة بوضوح.\n" if is_wa else "- إذا كانت الصورة لقطة شاشة لتحويل K-Net أو ومض (Wamd) أو إشعار بنكي استخرج المبلغ والمرجع والمستفيد بدقة، وإذا كانت صورة فوتوغرافية لموقع/معاينة/أعمال وليست ورقة فاجعل doc_type=\"site_photo\".\n"
    proj_hint_line = f"- تلميح مجلد السكانر الفرعي (مهم جداً): {subfolder_hint}\n" if subfolder_hint else ""
    known_line = f"- مشاريع حالية معروفة (طابق الاسم معها إن كان المستند يخص أحدها): {', '.join(existing_projects)}\n" if existing_projects else ""
    text_block = f"\nنص رقمي مستخرج من الملف (للمساعدة، وتحقق دائماً من الصورة والأختام):\n{embedded_text[:2000]}\n" if embedded_text else ""

    return f"""افحص صور المستند/الملف المرفقة بدقة عالية {refs}
معلومات مساعدة:
- اسم الملف الأصلي: {orig_name}
- عدد الصفحات: {page_count}
{dom_line}{active_line}{custom_line}{wa_line}{proj_hint_line}{known_line}- العملة الافتراضية عند عدم وضوح الرمز: {default_currency} (الدينار الكويتي KWD بـ 3 خانات عشرية 0.000).
{text_block}
المطلوب: اقرأ الترويسة، الجداول، خط اليد، التواريخ، المبالغ، علامات الاختيار (A/B/C/D أو معتمد/مرفوض)، والأختام، ثم أرجع كائن JSON واحد فقط (بدون أي شرح قبله أو بعده) بالهيكل التالي تماماً:

{{
  "doc_type": "submittal | petty_cash | invoice | letter | contract_ipc | permit | site_photo | custom | other",
  "custom_category": "اسم القسم المخصص من قائمة المستخدم إن كان doc_type=custom وإلا فارغ",
  "confidence": 0.95,
  "is_project_related": true,
  "relevance_score": 0.90,
  "detected_project_or_entity": "اسم المشروع أو الشركة أو الجهة المستقلة إن كان الملف لا يخص المشروع النشط، وإلا فارغ",
  "relevance_reason": "سبب مختصر لارتباط الملف بالمشروع النشط أو عدم ارتباطه",
  "is_blank_or_unreadable": false,
  "project": "اسم أو كود المشروع إن ظهر وإلا اتركه فارغاً",
  "date": "YYYY-MM-DD (التاريخ الأساسي للمستند أو تاريخ الرد إن كان ورقة رد معتمدة) أو فارغ",
  "submit_date": "YYYY-MM-DD (تاريخ التقديم/الإرسال المكتوب بخانة Submit إن ظهر) أو فارغ",
  "reply_date": "YYYY-MM-DD (تاريخ الختم/الرد الفعلي الأخير إن وُجد؛ اتركه فارغاً تماماً إذا كان الختم غير مؤرخ أو بخط غير واضح — لا تخمّن أي تاريخ أبداً) أو فارغ",
  "consultant_reply_date": "YYYY-MM-DD (تاريخ ختم/رد المكتب الاستشاري فقط إذا ظهر بوضوح؛ وإلا اتركه فارغاً تماماً دون تخمين)",
  "consultant_action": "قرار الاستشاري حسب ختمه: Approved (Code A) | Approved as Noted (Code B) | Revise & Resubmit (Code C) | Rejected (Code D) | فارغ",
  "moh_reply_date": "YYYY-MM-DD (تاريخ ختم أو توقيع مهندس الوزارة أو وكيل الوزارة فقط إذا ظهر بوضوح؛ وإلا اتركه فارغاً تماماً دون تخمين)",
  "moh_action": "قرار جهة الوزارة/المالك (مهندس الوزارة أو الوكيل): Approved | Approved as Noted | Revise & Resubmit | Rejected | فارغ",
  "has_both_stamps": false,
  "ref_no": "رقم الاعتماد/RFI/NCR أو الفاتورة أو الكتاب أو المستخلص أو الترخيص أو عملية K-Net/ومض",
  "revision": "رقم المراجعة للاعتمادات/المخططات مثل R00 أو R01 أو فارغ",
  "is_superseded": false,
  "discipline": "للاعتمادات/الاستشارات: CIV أو EL أو PL أو HVAC أو FF أو ELVE أو ARCH أو STR أو كود القسم",
  "authority": "تنبيه مهم جداً للأختام: قد تحمل الورقة العائدة من الوزارة ختم أو توقيع مهندس الوزارة أو مدير الإدارة أو وكيل الوزارة/الوكيل المساعد بدلاً من كلمة الوزارة صراحة، وكثيراً ما تحمل الورقة ختم الاستشاري وختم مهندس/وكيل الوزارة معاً على نفس الصفحة — لذا اكتب: MOH+CONSULTANT إذا وُجد الختمان معاً، وMOH إذا كان الختم/التوقيع لمهندس أو وكيل الوزارة أو الوزارة فقط، وSOOR إذا كان للاستشاري فقط، أو البلدية/KFF/MEW",
  "submittal_type": "MAR | MIR | WIR | Shop Drawing | RFI | NCR | SI | Method Statement | Prequalification | فارغ",
  "letter_direction": "للرسائل: Incoming (وارد) أو Outgoing (صادر) أو Internal",
  "from_party": "الجهة المرسلة أو المورد أو المقاول أو العميل أو الموظف",
  "to_party": "الجهة المرسل إليها أو المستفيد",
  "subject": "وصف مختصر ودقيق للمادة أو البند أو موضوع الكتاب أو الفاتورة أو وصف صورة الموقع",
  "expense_category": "للعهدة: مواد موقع | عدد وأدوات | محروقات ونقل | قرطاسية | ضيافة | صيانة | رسوم | تحويل ومض/بنكي | أخرى",
  "payment_method": "نقدي | K-Net | ومض Wamd | تحويل بنكي | شيك | آجل | فارغ",
  "po_or_contract_no": "رقم أمر الشراء PO أو رقم العقد أو عرض السعر إن وجد",
  "contract_subtype": "Contract | IPC | VO | PO | Guarantee | BOQ | فارغ",
  "gross_amount": null,
  "deductions_amount": null,
  "net_amount": null,
  "currency": "KWD أو USD أو SAR أو EGP أو AED أو فارغ",
  "status": "الحالة الفعلية حسب الختم/الإشارة: Approved (Code A) | Approved as Noted (Code B) | Revise & Resubmit (Code C) | Rejected (Code D) | Under Review | ساري | مدفوع | قيد السداد | بانتظار الرد | موثق",
  "needs_reply": false,
  "suggested_filename": "اسم ملف نظيف وواضح بدون امتداد (تجاهل اسم WhatsApp Image تماماً واستبدله باسم رسمي مثل: EL-135-R01-MOH أو 2026-11-17_المورد_12.500KWD أو 2026-11-17_معاينة_صب_السقف)",
  "notes": "أي ملاحظة تدقيق مهمة (ختم، تحفظ، تاريخ انتهاء ترخيص، رقم مرجع ومض/K-Net، إلخ)"
}}
"""


def extract_json_object(raw: str) -> dict:
    """يستخرج كائن JSON بأمان سواء جاء ضمن غلاف Gemini CLI أو داخل ```json```."""
    text = raw.strip()
    # إذا كان غلاف Gemini CLI بصيغة {"response": "..."}
    try:
        outer = json.loads(text)
        if isinstance(outer, dict):
            if "doc_type" in outer:
                return outer
            if isinstance(outer.get("response"), str):
                text = outer["response"].strip()
    except Exception:
        pass

    # إزالة أسوار Markdown إن وجدت
    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.S | re.I)
    if fence:
        return json.loads(fence.group(1))

    # إيجاد أول { وآخر }
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return json.loads(text[start : end + 1])

    raise ValueError(_tr("watcher.read_failed_retry"))


GEMINI_REST_BASE = "https://generativelanguage.googleapis.com/v1beta"
AI_KEY_FILE = Path.home() / ".gemini" / "ai_key.json"
_CACHED_MODELS: dict[str, list[str]] = {}
_KEY_STATE_LOCK = threading.Lock()
_QUOTA_EXHAUSTED_KEYS: dict[str, float] = {}  # key_hash -> epoch timestamp when marked exhausted
_QUOTA_COOLDOWN_SEC = 900.0
_MODEL_CACHE_TTL_SEC = 86400.0


def _runtime_cache_path() -> Path:
    return AI_KEY_FILE.parent / "gemini_runtime_cache.json"


def _read_runtime_cache() -> dict:
    from security_utils import read_json_safe
    raw = read_json_safe(_runtime_cache_path(), {})
    if not isinstance(raw, dict):
        return {"exhausted_keys": {}, "models": {}}
    ex = raw.get("exhausted_keys") if isinstance(raw.get("exhausted_keys"), dict) else {}
    md = raw.get("models") if isinstance(raw.get("models"), dict) else {}
    return {"exhausted_keys": dict(ex), "models": dict(md)}


def _write_runtime_cache(data: dict) -> None:
    try:
        from security_utils import atomic_write_json
        rcp = _runtime_cache_path()
        rcp.parent.mkdir(parents=True, exist_ok=True)
        atomic_write_json(rcp, data)
    except Exception:
        pass


def _key_fp(api_key: str) -> str:
    k = str(api_key or "").strip()
    return hashlib.sha256(k.encode("utf-8")).hexdigest()[:16] if k else ""


def _is_key_in_cooldown(api_key: str) -> bool:
    fp = _key_fp(api_key)
    if not fp:
        return False
    now_w = time.time()
    with _KEY_STATE_LOCK:
        ts = _QUOTA_EXHAUSTED_KEYS.get(fp)
        if ts is None:
            rt = _read_runtime_cache()
            disk_ts = rt["exhausted_keys"].get(fp)
            if disk_ts is not None:
                try:
                    ts = float(disk_ts)
                    _QUOTA_EXHAUSTED_KEYS[fp] = ts
                except (TypeError, ValueError):
                    ts = None
        if ts is None:
            return False
        if now_w - ts > _QUOTA_COOLDOWN_SEC:
            _QUOTA_EXHAUSTED_KEYS.pop(fp, None)
            rt = _read_runtime_cache()
            if fp in rt["exhausted_keys"]:
                rt["exhausted_keys"].pop(fp, None)
                _write_runtime_cache(rt)
            return False
        return True


def mark_gemini_key_quota_exhausted(api_key: str) -> None:
    fp = _key_fp(api_key)
    if not fp:
        return
    now_w = time.time()
    with _KEY_STATE_LOCK:
        _QUOTA_EXHAUSTED_KEYS[fp] = now_w
        rt = _read_runtime_cache()
        rt["exhausted_keys"][fp] = now_w
        _write_runtime_cache(rt)


def clear_gemini_key_quota_state(api_key: str = "") -> None:
    with _KEY_STATE_LOCK:
        rt = _read_runtime_cache()
        if not api_key:
            _QUOTA_EXHAUSTED_KEYS.clear()
            if rt["exhausted_keys"]:
                rt["exhausted_keys"] = {}
                _write_runtime_cache(rt)
        else:
            fp = _key_fp(api_key)
            _QUOTA_EXHAUSTED_KEYS.pop(fp, None)
            if fp in rt["exhausted_keys"]:
                rt["exhausted_keys"].pop(fp, None)
                _write_runtime_cache(rt)


def _is_quota_error(http_code: int, message: str = "") -> bool:
    if int(http_code or 0) == 429:
        return True
    low = str(message or "").lower()
    return any(
        tok in low
        for tok in (
            "quota",
            "resource_exhausted",
            "rate limit",
            "rate_limit",
            "too many requests",
            "exceeded",
            "limit: 0",
        )
    )


def _read_ai_key_file_dict() -> dict:
    if AI_KEY_FILE.exists():
        try:
            d = json.loads(AI_KEY_FILE.read_text(encoding="utf-8"))
            if isinstance(d, dict):
                return d
        except Exception:
            pass
    return {}


def _mask_key(k: str) -> str:
    s = str(k or "").strip()
    if not s:
        return ""
    if len(s) >= 12:
        return f"{s[:6]}...{s[-4:]}"
    return "*" * len(s)


def _parse_key_tokens(raw) -> list[str]:
    if raw is None:
        return []
    out: list[str] = []

    def _walk(item) -> None:
        if item is None:
            return
        if isinstance(item, (list, tuple)):
            for sub in item:
                _walk(sub)
            return
        for p in re.split(r"[\s,;]+", str(item)):
            tok = p.strip()
            if tok and tok not in out:
                out.append(tok)

    _walk(raw)
    return out


def load_saved_primary_gemini_key() -> str:
    env_k = (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or "").strip()
    if env_k:
        return env_k
    d = _read_ai_key_file_dict()
    return str(d.get("gemini_api_key") or "").strip()


def load_saved_gemini_backup_keys() -> list[str]:
    primary = load_saved_primary_gemini_key()
    env_b1 = (os.environ.get("GEMINI_BACKUP_API_KEY") or "").strip()
    env_b_many = (os.environ.get("GEMINI_BACKUP_API_KEYS") or "").strip()
    d = _read_ai_key_file_dict()
    tokens = _parse_key_tokens([env_b1, env_b_many, d.get("backup_gemini_api_key"), d.get("backup_gemini_api_keys")])
    return [k for k in tokens if k != primary]


def load_saved_gemini_backup_key() -> str:
    baks = load_saved_gemini_backup_keys()
    if not baks:
        return ""
    for k in baks:
        if not _is_key_in_cooldown(k):
            return k
    return baks[0]


def load_saved_gemini_keys() -> list[str]:
    """Return all configured Gemini keys `[primary, backup1, backup2, ...]` deduplicated, prioritizing non-exhausted keys."""
    primary = load_saved_primary_gemini_key()
    backups = load_saved_gemini_backup_keys()
    ordered: list[str] = []
    for k in [primary, *backups]:
        if k and k not in ordered:
            ordered.append(k)
    if len(ordered) <= 1:
        return ordered
    fresh = [k for k in ordered if not _is_key_in_cooldown(k)]
    stale = [k for k in ordered if _is_key_in_cooldown(k)]
    return fresh + stale


def load_saved_gemini_key() -> str:
    keys = load_saved_gemini_keys()
    return keys[0] if keys else ""


def load_saved_gemini_model() -> str:
    d = _read_ai_key_file_dict()
    m = str(d.get("top_model") or "").strip()
    if m.startswith("gemini-"):
        return m
    return ""


def get_gemini_keys_status() -> dict:
    primary = load_saved_primary_gemini_key()
    backups = load_saved_gemini_backup_keys()
    p_exhausted = bool(primary and _is_key_in_cooldown(primary))

    # Determine which key is currently first in rotation (`load_saved_gemini_keys()[0]`)
    ordered = load_saved_gemini_keys()
    active_raw = ordered[0] if ordered else ""

    if primary and active_raw == primary:
        slot = "primary"
        active_idx = 0
    elif backups and active_raw in backups:
        slot = "backup"
        active_idx = backups.index(active_raw) + 1
    else:
        slot = "none"
        active_idx = -1

    masked_p = _mask_key(primary)
    masked_b_first = _mask_key(active_raw if slot == "backup" else (backups[0] if backups else ""))
    backup_list = [
        {
            "index": idx,
            "masked": _mask_key(bk),
            "exhausted": bool(_is_key_in_cooldown(bk)),
            "active": bool(active_raw == bk),
        }
        for idx, bk in enumerate(backups)
    ]

    return {
        "has_key": bool(primary or backups),
        "has_primary": bool(primary),
        "has_backup_key": bool(backups),
        "backup_count": len(backups),
        "backup_keys": backup_list,
        "masked_key": masked_p or masked_b_first,
        "masked_primary_key": masked_p,
        "masked_backup_key": masked_b_first,
        "active_key_slot": slot,
        "active_key_index": active_idx,
        "primary_quota_exhausted": p_exhausted,
        "model": load_saved_gemini_model() or "auto",
    }


def save_gemini_key(
    api_key: str | None = None,
    top_model: str = "",
    backup_api_key: str | None = None,
    backup_api_keys: list[str] | None = None,
) -> None:
    existing = _read_ai_key_file_dict()
    clean_k = str(existing.get("gemini_api_key") or "").strip() if api_key is None else str(api_key or "").strip()

    existing_baks = _parse_key_tokens([existing.get("backup_gemini_api_key"), existing.get("backup_gemini_api_keys")])
    if backup_api_keys is not None:
        clean_baks = [k for k in _parse_key_tokens(backup_api_keys) if k != clean_k]
    elif backup_api_key is not None:
        clean_baks = [k for k in _parse_key_tokens(backup_api_key) if k != clean_k]
    else:
        clean_baks = [k for k in existing_baks if k != clean_k]

    if api_key is not None:
        clear_gemini_key_quota_state(str(existing.get("gemini_api_key") or ""))
        clear_gemini_key_quota_state(clean_k)
    if backup_api_key is not None or backup_api_keys is not None:
        for old_b in existing_baks:
            clear_gemini_key_quota_state(old_b)
        for new_b in clean_baks:
            clear_gemini_key_quota_state(new_b)

    gemini_dir = AI_KEY_FILE.parent
    gemini_dir.mkdir(parents=True, exist_ok=True)
    payload: dict = {"gemini_api_key": clean_k}
    if clean_baks:
        payload["backup_gemini_api_key"] = clean_baks[0]
        payload["backup_gemini_api_keys"] = clean_baks
    eff_model = top_model.strip() if (top_model and top_model.startswith("gemini-")) else str(existing.get("top_model") or "").strip()
    if eff_model and eff_model.startswith("gemini-") and (clean_k or clean_baks):
        payload["top_model"] = eff_model
    AI_KEY_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # مزامنة المفتاح تلقائياً مع الوكيل التفاعلي الكامل (Gemini CLI) ليعمل كـ Agent كامل بدون IneligibleTierError!
    sync_k = clean_k or (clean_baks[0] if clean_baks else "")
    try:
        if gemini_dir.name == ".gemini":
            env_f = gemini_dir / ".env"
            env_f.write_text(f"GEMINI_API_KEY={sync_k}\n", encoding="utf-8")
            settings_f = gemini_dir / "settings.json"
            sdata: dict = {}
            if settings_f.exists():
                try:
                    loaded = json.loads(settings_f.read_text(encoding="utf-8"))
                    if isinstance(loaded, dict):
                        sdata = loaded
                except Exception:
                    sdata = {}
            sec = sdata.setdefault("security", {})
            if not isinstance(sec, dict):
                sec = {}
                sdata["security"] = sec
            auth = sec.setdefault("auth", {})
            if not isinstance(auth, dict):
                auth = {}
                sec["auth"] = auth
            auth["selectedType"] = "gemini-api-key"
            settings_f.write_text(json.dumps(sdata, ensure_ascii=False, indent=2), encoding="utf-8")
            if os.name != "nt":
                os.chmod(AI_KEY_FILE, 0o600)
                os.chmod(env_f, 0o600)
        elif os.name != "nt":
            os.chmod(AI_KEY_FILE, 0o600)
    except Exception:
        pass


def save_gemini_backup_key(backup_api_key: str, top_model: str = "") -> None:
    save_gemini_key(api_key=None, top_model=top_model, backup_api_key=backup_api_key)


def add_gemini_backup_keys(raw_keys: str | list[str], top_model: str = "") -> list[str]:
    """Append one or more backup keys to the existing backup list (deduplicated)."""
    existing = _read_ai_key_file_dict()
    prim = str(existing.get("gemini_api_key") or "").strip()
    cur_baks = _parse_key_tokens([existing.get("backup_gemini_api_key"), existing.get("backup_gemini_api_keys")])
    new_tokens = _parse_key_tokens(raw_keys)
    merged = list(cur_baks)
    for tok in new_tokens:
        if tok and tok != prim and tok not in merged:
            merged.append(tok)
    save_gemini_key(api_key=None, top_model=top_model, backup_api_keys=merged)
    return merged


def remove_gemini_backup_key(index: int) -> bool:
    existing = _read_ai_key_file_dict()
    cur_baks = _parse_key_tokens([existing.get("backup_gemini_api_key"), existing.get("backup_gemini_api_keys")])
    if not (0 <= int(index) < len(cur_baks)):
        return False
    removed = cur_baks.pop(int(index))
    clear_gemini_key_quota_state(removed)
    save_gemini_key(api_key=None, backup_api_keys=cur_baks)
    return True


def promote_gemini_backup_key(index: int) -> bool:
    """Make the backup key at `index` the primary key, moving previous primary to backup list."""
    existing = _read_ai_key_file_dict()
    prim = str(existing.get("gemini_api_key") or "").strip()
    cur_baks = _parse_key_tokens([existing.get("backup_gemini_api_key"), existing.get("backup_gemini_api_keys")])
    if not (0 <= int(index) < len(cur_baks)):
        return False
    new_prim = cur_baks.pop(int(index))
    if prim and prim != new_prim and prim not in cur_baks:
        cur_baks.insert(0, prim)
    save_gemini_key(api_key=new_prim, backup_api_keys=cur_baks)
    return True


def discover_free_gemini_models(api_key: str, api_base: str = GEMINI_REST_BASE, force_refresh: bool = False) -> list[str]:
    """يكتشف تلقائياً أحدث نماذج Gemini المجانية المتاحة على مفتاح المستخدم ويرتبها حسب أعلى حصة يومية مجانية ودقة بصرية."""
    cache_k = api_key[-8:]
    if not force_refresh and cache_k in _CACHED_MODELS and _CACHED_MODELS[cache_k]:
        return list(_CACHED_MODELS[cache_k])
    disk_k = f"{_key_fp(api_key)}:{api_base}"
    if not force_refresh:
        with _KEY_STATE_LOCK:
            rt = _read_runtime_cache()
            rec = rt["models"].get(disk_k)
            if isinstance(rec, dict) and isinstance(rec.get("models"), list) and rec["models"]:
                ts_m = float(rec.get("ts") or 0.0)
                if ts_m > 0 and (time.time() - ts_m) <= _MODEL_CACHE_TTL_SEC:
                    _CACHED_MODELS[cache_k] = list(rec["models"])
                    return list(rec["models"])

    skip_tokens = (
        "tts", "image", "imagen", "veo", "lyria", "embed", "aqa", "audio",
        "native-audio", "live", "robotics", "computer-use", "thinking-exp", "learnlm", "1.0", "1.5-pro",
    )
    req = urllib.request.Request(
        f"{api_base}/models?pageSize=100",
        headers={"x-goog-api-key": api_key},
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=25) as resp:
        data = json.loads(resp.read().decode("utf-8", errors="replace"))

    flash_lite: list[tuple[tuple[int, int, int], str]] = []
    flash_std: list[tuple[tuple[int, int, int], str]] = []
    gemma_it: list[tuple[tuple[int, int, int], str]] = []

    def _ver_key(mid: str) -> tuple[int, int, int]:
        m = re.search(r"(\d+)(?:\.(\d+))?", mid)
        maj = int(m.group(1)) if m else 0
        mn = int(m.group(2)) if (m and m.group(2)) else 0
        stable = 0 if ("preview" in mid or "exp" in mid) else 1
        return (maj, mn, stable)

    for m in data.get("models") or []:
        methods = m.get("supportedGenerationMethods") or []
        if "generateContent" not in methods:
            continue
        name = str(m.get("name") or "").replace("models/", "").strip()
        low = name.lower()
        if not name or any(tok in low for tok in skip_tokens):
            continue
        if low.startswith("gemini-") and "flash-lite" in low and "latest" not in low:
            flash_lite.append((_ver_key(low), name))
        elif low.startswith("gemini-") and "flash" in low and "pro" not in low and "latest" not in low:
            flash_std.append((_ver_key(low), name))
        elif low.startswith("gemma-") and low.endswith("-it") and any(sz in low for sz in ("31b", "27b", "26b", "12b")):
            gemma_it.append((_ver_key(low), name))

    flash_lite.sort(reverse=True)
    flash_std.sort(reverse=True)
    gemma_it.sort(reverse=True)

    ranked = [n for _, n in flash_lite[:3]] + [n for _, n in flash_std[:3]] + [n for _, n in gemma_it[:2]]
    if not ranked:
        ranked = ["gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-3-flash-preview", "gemini-2.5-flash-lite", "gemini-2.5-flash"]
    _CACHED_MODELS[cache_k] = ranked
    with _KEY_STATE_LOCK:
        rt = _read_runtime_cache()
        rt["models"][disk_k] = {"models": ranked, "ts": time.time()}
        _write_runtime_cache(rt)
    return list(ranked)


def analyze_with_gemini_api(
    prompt: str,
    images: list[Path],
    api_key: str = "",
    preferred_model: str = "",
    api_base: str = GEMINI_REST_BASE,
    backup_api_key: str = "",
) -> dict:
    """يقرأ المستندات والصور والأختام مباشرة عبر Google Gemini API المجاني مع تبديل تلقائي للنماذج والمفتاح الاحتياطي عند نفاد الكوتة."""
    explicit_k = str(api_key or "").strip()
    explicit_bak = str(backup_api_key or "").strip()
    keys_to_try: list[str] = []
    if explicit_k:
        keys_to_try.append(explicit_k)
        if explicit_bak and explicit_bak not in keys_to_try:
            keys_to_try.append(explicit_bak)
        saved_all = load_saved_gemini_keys()
        if not explicit_bak and explicit_k in saved_all:
            for sk in saved_all:
                if sk not in keys_to_try:
                    keys_to_try.append(sk)
    else:
        keys_to_try = load_saved_gemini_keys()
        if explicit_bak and explicit_bak not in keys_to_try:
            keys_to_try.append(explicit_bak)

    if not keys_to_try:
        raise RuntimeError(_tr("watcher.need_ai_key"))

    clean_prompt = re.sub(r"@\./page_\d+\.png\s*", "", prompt)

    img_parts = []
    for p in images:
        if p.exists():
            b64 = base64.b64encode(p.read_bytes()).decode("ascii")
            mime = "image/png" if p.suffix.lower() == ".png" else "image/jpeg"
            img_parts.append({"inline_data": {"mime_type": mime, "data": b64}})

    default_fallback_models = [
        "gemini-3.5-flash-lite",
        "gemini-3.1-flash-lite",
        "gemini-3-flash-preview",
        "gemini-2.5-flash-lite",
        "gemini-2.5-flash",
    ]

    last_err = ""
    had_bad_key_only = True
    for key_idx, cur_key in enumerate(keys_to_try):
        has_more_keys = key_idx + 1 < len(keys_to_try)
        try:
            candidates = discover_free_gemini_models(cur_key, api_base=api_base)
        except urllib.error.HTTPError as e:
            err_raw = e.read().decode("utf-8", errors="replace") if hasattr(e, "read") else ""
            if _is_quota_error(e.code, err_raw) and has_more_keys:
                mark_gemini_key_quota_exhausted(cur_key)
                last_err = f"key[{key_idx+1}] quota ({e.code})"
                had_bad_key_only = False
                continue
            if e.code in (400, 401, 403):
                last_err = f"key[{key_idx+1}] auth ({e.code})"
                if has_more_keys:
                    continue
                raise RuntimeError(_tr("watcher.bad_ai_key"))
            candidates = list(default_fallback_models)
        except Exception:
            candidates = list(default_fallback_models)

        if preferred_model and preferred_model.strip():
            pm = preferred_model.strip().replace("models/", "")
            candidates = [pm] + [c for c in candidates if c != pm]

        quota_hits_on_key = 0
        for model_id in candidates[:6]:
            is_gemma = model_id.lower().startswith("gemma")
            gen_cfg: dict = {"temperature": 0.1}
            if not is_gemma:
                gen_cfg["responseMimeType"] = "application/json"

            payload = {
                "contents": [{"parts": img_parts + [{"text": clean_prompt}]}],
                "generationConfig": gen_cfg,
            }
            req = urllib.request.Request(
                f"{api_base}/models/{model_id}:generateContent",
                data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                headers={"Content-Type": "application/json", "x-goog-api-key": cur_key},
                method="POST",
            )
            try:
                with urllib.request.urlopen(req, timeout=120) as resp:
                    body = json.loads(resp.read().decode("utf-8", errors="replace"))
                cands = body.get("candidates") or []
                if not cands:
                    had_bad_key_only = False
                    last_err = f"{model_id}: empty"
                    continue
                parts = (cands[0].get("content") or {}).get("parts") or []
                text_out = "".join(str(pt.get("text") or "") for pt in parts if not pt.get("thought"))
                if not text_out.strip():
                    text_out = "".join(str(pt.get("text") or "") for pt in parts)
                # Keep primary key clean if it succeeded; keep backup key clean if used
                clear_gemini_key_quota_state(cur_key)
                return extract_json_object(text_out)
            except urllib.error.HTTPError as e:
                err_raw = e.read().decode("utf-8", errors="replace")
                try:
                    emsg = json.loads(err_raw).get("error", {}).get("message") or err_raw[:200]
                except Exception:
                    emsg = err_raw[:200]
                last_err = f"{model_id} ({e.code}): {emsg}"
                if e.code in (401, 403) or (e.code == 400 and "api key" in emsg.lower()):
                    if has_more_keys:
                        break
                    if had_bad_key_only:
                        raise RuntimeError(_tr("watcher.bad_ai_key"))
                    break
                had_bad_key_only = False
                if _is_quota_error(e.code, emsg):
                    quota_hits_on_key += 1
                    low_em = emsg.lower()
                    project_wide_quota = ("exceeded your current quota" in low_em) or ("check your plan" in low_em) or ("limit: 0" in low_em)
                    if has_more_keys and (project_wide_quota or quota_hits_on_key >= 2):
                        # Primary key quota exhausted -> switch immediately to backup key!
                        mark_gemini_key_quota_exhausted(cur_key)
                        break
                    continue
                if e.code in (400, 404, 500, 502, 503):
                    continue
                raise RuntimeError(last_err)
            except Exception as e:
                had_bad_key_only = False
                last_err = f"{model_id}: {e}"
                continue
        else:
            if quota_hits_on_key > 0 and has_more_keys:
                mark_gemini_key_quota_exhausted(cur_key)
            continue

        # Broke out early due to key auth/quota failure -> try next backup key
        if quota_hits_on_key > 0:
            mark_gemini_key_quota_exhausted(cur_key)

    if had_bad_key_only and last_err:
        raise RuntimeError(_tr("watcher.bad_ai_key"))
    detail = f" ({last_err[:140]})" if last_err else ""
    raise RuntimeError(f"{_tr('watcher.ai_busy')}{detail}")


def analyze_with_gemini(
    prompt: str,
    job_dir: Path,
    model: str = "",
    images: list[Path] | None = None,
    api_key: str = "",
    backup_api_key: str = "",
) -> dict:
    """اتصال مباشر فائق السرعة بـ Google Gemini API المجاني (مع تبديل تلقائي للمفتاح الاحتياطي عند انتهاء الكوتة)."""
    imgs = list(images) if images is not None else sorted(job_dir.glob("page_*.png"))
    return analyze_with_gemini_api(prompt, imgs, api_key=api_key, preferred_model=model, backup_api_key=backup_api_key)


def analyze_with_ollama(prompt: str, images: list[Path], url: str, model: str) -> dict:
    b64_imgs = [base64.b64encode(p.read_bytes()).decode("ascii") for p in images]
    payload = {
        "model": model,
        "stream": False,
        "format": "json",
        "messages": [{"role": "user", "content": prompt, "images": b64_imgs}],
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=300) as resp:
        body = json.loads(resp.read().decode("utf-8", errors="replace"))
    content = body.get("message", {}).get("content", "")
    return extract_json_object(content)


DISC_LABELS = {
    "CIV": "CIV - المدني",
    "CIVIL": "CIV - المدني",
    "ARCH": "ARCH - المعماري",
    "STR": "STR - الإنشائي",
    "EL": "EL - الكهرباء",
    "ELEC": "EL - الكهرباء",
    "PL": "PL - الصحي",
    "PLUMB": "PL - الصحي",
    "HVAC": "HVAC - التكييف",
    "MEC": "MEC - الميكانيك",
    "MECH": "MEC - الميكانيك",
    "FF": "FF - مكافحة الحريق",
    "FIRE": "FF - مكافحة الحريق",
    "ELVE": "ELVE - المصاعد",
    "ELEV": "ELVE - المصاعد",
    "LOW": "ELV - التيار الخفيف",
    "ELV": "ELV - التيار الخفيف",
}


_GENERIC_PARTY_WORDS = {
    "شركه", "موسسه", "مكتب", "مقاول", "المقاول", "مورد", "المورد", "الساده",
    "ساده", "مهندس", "المهندس", "م", "أ", "ا", "للمقاولات", "مقاولات", "العامه",
    "عامه", "للتجاره", "تجاره", "والمقاولات", "الحديثه", "المتحده", "الدوليه",
    "company", "co", "est", "llc", "wll", "ltd", "inc", "group", "contracting",
    "general", "trading", "contractor", "subcontractor", "mr", "eng",
}


def _norm_arabic_text(s: str) -> str:
    """توحيد الحروف العربية المتشابهة عبر دالة التطبيع الموحّدة relevance.normalize_ar."""
    from relevance import normalize_ar
    t = normalize_ar(str(s or ""))
    t = re.sub(r"\bعبد\s+", "عبد", t)
    return re.sub(r"\s+", " ", t).strip()


def _core_party_tokens(s: str) -> list[str]:
    toks = [w for w in _norm_arabic_text(s).split() if w and w not in _GENERIC_PARTY_WORDS]
    if not toks:
        toks = _norm_arabic_text(s).split()
    return [w[2:] if w.startswith("ال") and len(w) > 4 else w for w in toks]


def resolve_party_folder(parent_dir: Path, raw_name: str, fallback: str = "عام", max_len: int = 40) -> str:
    """يبحث أولاً عن مجلد موجود لنفس الشخص/المقاول/المورد (حتى مع اختلاف علي/على أو مؤسسة علي)
    فيضيف بداخله بدلاً من تكرار المجلد، وإن كان اسماً جديداً لأول مرة ينشئ له مجلداً باسمه."""
    cleaned = clean_seg(raw_name or fallback, max_len) or fallback
    if not parent_dir or not parent_dir.exists():
        return cleaned

    existing = [
        d.name for d in sorted(parent_dir.iterdir())
        if d.is_dir() and not d.name.startswith(("_", ".", "00_"))
    ]
    if not existing:
        return cleaned

    target_norm = _norm_arabic_text(cleaned)
    target_core = _core_party_tokens(cleaned)

    # 1) تطابق مباشر أو بعد توحيد الحروف العربية (علي == على، أحمد == احمد)
    for name in existing:
        if name == cleaned or _norm_arabic_text(name) == target_norm:
            return name

    # 2) تطابق الاسم الجوهري بعد إزالة الألقاب والكلمات العامة (علي == المقاول علي == مؤسسة علي للمقاولات)
    if target_core:
        for name in existing:
            if _core_party_tokens(name) == target_core:
                return name

        # 3) تطابق بادئة الاسم شريطة عدم وجود التباس (لا يخلط أبداً بين "علي حسن" و"علي محمود"!)
        prefix_hits = []
        for name in existing:
            ex_core = _core_party_tokens(name)
            if not ex_core:
                continue
            shorter, longer = (target_core, ex_core) if len(target_core) <= len(ex_core) else (ex_core, target_core)
            # التطابق الجزئي مسموح فقط إذا كان الطرف الأقصر يطابق بداية اسم مجلد واحد فقط بلا أي تشابه آخر
            if len(target_core) == 1 and len(ex_core) == 1 and target_core[0] == ex_core[0]:
                prefix_hits.append(name)
            elif len(shorter) >= 2 and longer[: len(shorter)] == shorter:
                prefix_hits.append(name)
        if len(prefix_hits) == 1:
            return prefix_hits[0]

    return cleaned


def compute_destination(meta: dict, ext: str, cfg: dict, subfolder_hint: str) -> tuple[Path, str]:
    """يبني المسار الفرعي واسم الملف القياسي بمرونة تامة لأي مشروع وأي نوع مستند."""
    doc_type = str(meta.get("doc_type") or "other").strip().lower()
    if doc_type not in CATEGORY_FOLDERS:
        doc_type = "other"

    archive_root = Path(cfg["archive_folder"]) if cfg.get("archive_folder") else Path(".")
    cat_root = CATEGORY_FOLDERS[doc_type]
    project = clean_seg(meta.get("project") or subfolder_hint or cfg.get("default_project") or "عام", 45)
    date_str = clean_seg(meta.get("date") or dt.date.today().isoformat(), 12)
    month_str = date_str[:7] if re.match(r"^\d{4}-\d{2}", date_str) else dt.date.today().strftime("%Y-%m")
    currency = clean_seg(meta.get("currency") or cfg.get("default_currency") or "KWD", 6).upper()
    decimals = 3 if currency in {"KWD", "BHD", "OMR", "JOD"} else 2

    parts: list[str] = []
    if cfg.get("folder_layout") == "by_project_then_type":
        proj_folder = resolve_party_folder(archive_root, project or "عام", "عام", 45)
        parts.append(proj_folder)
    parts.append(cat_root)
    base_cat_dir = archive_root.joinpath(*parts)

    if doc_type == "submittal":
        raw_disc = clean_seg(meta.get("discipline") or "GEN", 16).upper()
        disc_folder = DISC_LABELS.get(raw_disc, raw_disc or "General")
        parts.append(disc_folder)
        if meta.get("is_superseded"):
            parts.append("_قبل التعديل (Superseded)")
        else:
            raw_auth = meta.get("authority") or "General"
            auth = resolve_party_folder(base_cat_dir / disc_folder, raw_auth, "General", 25)
            parts.append(auth)
    elif doc_type == "petty_cash":
        parts.append(month_str)
    elif doc_type == "invoice":
        vendor = resolve_party_folder(base_cat_dir, meta.get("from_party") or "مورد_عام", "مورد_عام", 40)
        parts.append(vendor)
    elif doc_type == "letter":
        d = str(meta.get("letter_direction") or "").lower()
        if "out" in d or "صادر" in d:
            dir_seg = "02_Outgoing - الصادر"
            raw_party = meta.get("to_party") or meta.get("from_party") or ""
        elif "int" in d or "داخل" in d:
            dir_seg = "03_Internal - داخلي"
            raw_party = ""
        else:
            dir_seg = "01_Incoming - الوارد"
            raw_party = meta.get("from_party") or ""
        parts.append(dir_seg)
        if clean_seg(raw_party, 38):
            party = resolve_party_folder(base_cat_dir / dir_seg, raw_party, "", 38)
            if party:
                parts.append(party)
    elif doc_type == "contract_ipc":
        raw_contractor = meta.get("from_party") or meta.get("to_party") or "مقاول_عام"
        contractor = resolve_party_folder(base_cat_dir, raw_contractor, "مقاول_عام", 40)
        sub = str(meta.get("contract_subtype") or "").upper()
        sub_map = {
            "CONTRACT": "01_Contract - العقد",
            "BOQ": "01_Contract - العقد",
            "IPC": "02_IPCs - المستخلصات",
            "VO": "03_Variations - الأوامر التغييرية",
            "PO": "05_POs - أوامر الشراء",
            "GUARANTEE": "04_Guarantees - الكفالات",
        }
        parts.append(contractor)
        parts.append(sub_map.get(sub, "06_Docs - مستندات تعاقدية"))
    elif doc_type == "permit":
        auth = resolve_party_folder(base_cat_dir, meta.get("authority") or meta.get("from_party") or "جهات_رسمية", "جهات_رسمية", 38)
        parts.append(auth)
    elif doc_type == "site_photo":
        topic = meta.get("discipline") or month_str
        sub_f = resolve_party_folder(base_cat_dir, topic, month_str, 35) if topic != month_str else month_str
        parts.append(sub_f)
    elif doc_type == "custom":
        raw_cat = meta.get("custom_category") or (cfg.get("custom_categories") or ["عام"])[0]
        cat_f = resolve_party_folder(base_cat_dir, raw_cat, "عام", 40)
        parts.append(cat_f)
        raw_p = meta.get("from_party") or meta.get("to_party") or ""
        if clean_seg(raw_p, 38):
            parts.append(resolve_party_folder(base_cat_dir / cat_f, raw_p, "", 38))
    else:
        parts.append(month_str)

    # بناء اسم الملف النظيف
    suggested = clean_seg(meta.get("suggested_filename") or "", 85)
    # إزالة أي امتداد مكرر داخل الاسم المقترح
    if suggested.lower().endswith(ext.lower()):
        suggested = suggested[: -len(ext)].rstrip(" .-_")

    if not suggested:
        ref = clean_seg(meta.get("ref_no") or "", 25)
        party = clean_seg(meta.get("from_party") or meta.get("authority") or "", 28)
        subj = clean_seg(meta.get("subject") or "مستند", 32)
        amt = meta.get("net_amount")
        amt_str = f"{float(amt):.{decimals}f}{currency}" if isinstance(amt, (int, float)) else ""
        tokens = [x for x in [date_str, party, ref, amt_str or subj] if x]
        suggested = "_".join(tokens) or f"Scan_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}"

    rel_dir = Path(*parts)
    filename = f"{suggested}{ext.lower()}"
    return rel_dir / filename, doc_type


def convert_image_to_pdf(img_path: Path, out_pdf: Path):
    with Image.open(img_path) as im:
        frames = []
        for idx in range(getattr(im, "n_frames", 1)):
            try:
                im.seek(idx)
            except EOFError:
                break
            frames.append(im.convert("RGB"))
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        frames[0].save(str(out_pdf), "PDF", resolution=150.0, save_all=True, append_images=frames[1:])


# =====================================================================
# إدارة السجل الشامل التلقائي (Master_Archive_Log.xlsx) مع حماية القفل
# =====================================================================
SHEETS_SPEC = [
    ("01_الاعتمادات", "سجل الاعتمادات الهندسية — ربط أصل Word/Excel بالنسخة المختومة (Submittals)", [
        ("م", 6), ("تاريخ التقديم", 13), ("تاريخ الختم/الرد", 14), ("المشروع", 17), ("القسم", 10),
        ("كود الاعتماد", 15), ("Rev", 7), ("النوع", 12), ("الجهة", 12),
        ("وصف المادة / البند", 32), ("المورد / المقاول", 22), ("الحالة / الكود", 20),
        ("ملاحظات", 24), ("أصل Word/Excel", 15), ("النسخة المختومة PDF", 17),
    ], []),
    ("02_العهدة", "سجل فواتير وإيصالات العُهدة (Petty Cash)", [
        ("م", 6), ("التاريخ", 13), ("المشروع", 18), ("رقم الفاتورة", 15), ("المورد / المحل", 24),
        ("البيان", 32), ("بند الصرف", 16), ("طريقة الدفع", 13), ("المبلغ", 15),
        ("العملة", 9), ("الحالة", 14), ("ملاحظات", 24), ("فتح الفاتورة", 14),
    ], [9]),
    ("03_الفواتير", "سجل فواتير الموردين والمقاولين (Invoices)", [
        ("م", 6), ("التاريخ", 13), ("المشروع", 18), ("رقم الفاتورة", 16), ("المورد / المقاول", 25),
        ("رقم PO", 14), ("البيان", 32), ("الإجمالي", 15), ("الخصم/الضريبة", 14),
        ("الصافي المستحق", 16), ("العملة", 9), ("الحالة", 15), ("ملاحظات", 24), ("فتح الفاتورة", 14),
    ], [8, 9, 10]),
    ("04_الصادر_والوارد", "سجل الكتب والمراسلات — ربط مسودة Word بكتاب السكان المختوم", [
        ("م", 6), ("التاريخ", 13), ("النوع", 12), ("رقم الكتاب", 18), ("الجهة المرسلة", 22),
        ("الجهة المرسل إليها", 22), ("المشروع", 17), ("الموضوع", 34),
        ("يحتاج رد؟", 11), ("الحالة", 16), ("ملاحظات", 22),
        ("أصل Word/Excel", 15), ("النسخة المختومة PDF", 17),
    ], []),
    ("05_العقود_والمستخلصات", "سجل العقود والمستخلصات — ربط شيت Excel بالنسخة المعتمدة المختومة", [
        ("م", 6), ("التاريخ", 13), ("المشروع", 17), ("رقم العقد", 15), ("المقاول / الطرف الثاني", 24),
        ("نوع المستند", 13), ("رقم المستخلص/VO", 16), ("البيان", 28), ("إجمالي الأعمال", 15),
        ("الاستقطاعات", 15), ("الصافي المستحق", 16), ("العملة", 9), ("الحالة", 16),
        ("ملاحظات", 22), ("أصل Word/Excel", 15), ("النسخة المختومة PDF", 17),
    ], [9, 10, 11]),
    ("06_تراخيص_وصور_وأقسام", "سجل التراخيص الرسمية وصور الموقع/الواتساب والأقسام المخصصة", [
        ("م", 6), ("التاريخ", 13), ("المشروع", 17), ("القسم / النوع", 22), ("الرقم المرجعي", 16),
        ("الجهة / الشخص", 24), ("الوصف / البيان", 36), ("المبلغ", 14), ("العملة", 9),
        ("الحالة", 16), ("ملاحظات", 26), ("فتح الملف", 16),
    ], [8]),
    ("07_سجل_السكان_الكامل", "السجل الزمني الكامل لكل الملفات الأصلية والواتساب وعمليات السكانر", [
        ("م", 6), ("وقت الأرشفة", 19), ("الملف الأصلي", 26), ("الإجراء", 18),
        ("التصنيف", 15), ("المشروع", 18), ("عدد الصفحات", 11), ("درجة الثقة", 11),
        ("المسار النهائي المنظم", 42), ("ملاحظات", 28), ("فتح الملف", 14),
    ], []),
]


def ensure_master_workbook(xlsx_path: Path):
    if xlsx_path.exists():
        return
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = "00_الملخص"
    ws0.views.sheetView[0].rightToLeft = True
    ws0.merge_cells("A1:E1")
    t = ws0["A1"]
    t.value = "لوحة المتابعة التلقائية — أرشيف السكانر الذكي (تتحدث تلقائياً مع كل ورقة)"
    t.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws0.row_dimensions[1].height = 34

    dash_headers = [("القسم / السجل", 30), ("عدد المستندات المؤرشفة", 22), ("إجمالي المبالغ المسجلة", 24), ("حالة المراقبة", 20), ("ملاحظة", 34)]
    ws0.row_dimensions[3].height = 26
    for i, (h, w) in enumerate(dash_headers, 1):
        c = ws0.cell(3, i, h)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=SLATE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
        ws0.column_dimensions[get_column_letter(i)].width = w

    dash_rows = [
        ("01 - الاعتمادات الهندسية (Submittals)", "=COUNTA('01_الاعتمادات'!B4:B5000)", "—", "تلقائي نشط", "مرتبة حسب القسم والجهة مع حفظ قبل التعديل"),
        ("02 - العُهدة النقدية (Petty Cash)", "=COUNTA('02_العهدة'!B4:B5000)", "=SUM('02_العهدة'!I4:I5000)", "تلقائي نشط", "مرتبة شهرياً مع معادلة إجمالي فورية"),
        ("03 - فواتير الموردين (Invoices)", "=COUNTA('03_الفواتير'!B4:B5000)", "=SUM('03_الفواتير'!J4:J5000)", "تلقائي نشط", "مرتبة بمجلد مستقل لكل مورد"),
        ("04 - الكتب (صادر ووارد)", "=COUNTA('04_الصادر_والوارد'!B4:B5000)", "—", "تلقائي نشط", "مفصولة وارد وصادر مع تتبع الكتب التي تنتظر رداً"),
        ("05 - العقود والمستخلصات (IPCs/VOs)", "=COUNTA('05_العقود_والمستخلصات'!B4:B5000)", "=SUM('05_العقود_والمستخلصات'!K4:K5000)", "تلقائي نشط", "مرتبة لكل مقاول (عقد / مستخلصات / VOs / كفالات)"),
        ("06 - التراخيص وصور الواتساب والأقسام", "=COUNTA('06_تراخيص_وصور_وأقسام'!B4:B5000)", "=SUM('06_تراخيص_وصور_وأقسام'!H4:H5000)", "تلقائي نشط", "تراخيص رسمية + صور موقع/واتساب + أقسام مخصصة"),
        ("07 - إجمالي عمليات السكانر والواتساب", "=COUNTA('07_سجل_السكان_الكامل'!B4:B10000)", "—", "صفر فقدان 100%", "يشمل المعتمد + صور واتساب + المكرر + قيد المراجعة"),
    ]
    for r_idx, row_vals in enumerate(dash_rows, start=4):
        ws0.row_dimensions[r_idx].height = 24
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws0.cell(r_idx, c_idx, val)
            cell.font = Font(name="Calibri", size=11, bold=(c_idx in (1, 2, 3)))
            cell.alignment = Alignment(horizontal="center" if c_idx in (2, 3, 4) else "right", vertical="center")
            cell.border = BORDER
            if c_idx == 3 and str(val).startswith("="):
                cell.number_format = "#,##0.000"
            if r_idx % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=ZEBRA)

    for sname, stitle, headers, _mcols in SHEETS_SPEC:
        ws = wb.create_sheet(sname)
        ws.views.sheetView[0].rightToLeft = True
        ncols = len(headers)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        c1 = ws.cell(1, 1, stitle)
        c1.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        c1.fill = PatternFill("solid", fgColor=NAVY)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[3].height = 26
        for idx, (h, w) in enumerate(headers, 1):
            hc = ws.cell(3, idx, h)
            hc.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
            hc.fill = PatternFill("solid", fgColor=SLATE)
            hc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            hc.border = BORDER
            ws.column_dimensions[get_column_letter(idx)].width = w
        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}3"

    wb.save(str(xlsx_path))


def _append_styled_row(ws, values: list, money_cols: list[int], currency: str):
    row_idx = max(ws.max_row + 1, 4)
    # إذا كان الصف 4 فارغاً تماماً استخدمه
    if row_idx == 4 and ws.cell(4, 2).value is not None:
        row_idx = 5
    seq = row_idx - 3
    values[0] = seq
    num_fmt = "#,##0.000" if str(currency).upper() in {"KWD", "BHD", "OMR", "JOD", ""} else "#,##0.00"
    ws.row_dimensions[row_idx].height = 22
    for c_idx, val in enumerate(values, start=1):
        cell = ws.cell(row_idx, c_idx, val)
        is_link = isinstance(val, str) and val.startswith("=HYPERLINK(")
        cell.font = Font(
            name="Calibri",
            size=10.5,
            bold=is_link,
            color=LINK_CLR if is_link else "1E293B",
            underline="single" if is_link else None,
        )
        cell.alignment = Alignment(
            horizontal="center" if (c_idx == 1 or is_link or c_idx in money_cols) else "right",
            vertical="center",
            wrap_text=False,
        )
        cell.border = BORDER
        if c_idx in money_cols and isinstance(val, (int, float)):
            cell.number_format = num_fmt
        elif row_idx % 2 == 1:
            cell.fill = PatternFill("solid", fgColor=ZEBRA)


def _num_or_none(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None


def _norm_ref(val: str) -> str:
    """يوحّد صيغة الأرقام المرجعية للمطابقة الذكية (مثلاً EL-0228 == el 228 == EL-228)."""
    if not val:
        return ""
    s = re.sub(r"[\s_/\\]+", "-", str(val).strip().upper())
    s = re.sub(r"-+", "-", s).strip("-")
    s = re.sub(r"([A-Z])-?0+(\d+)", r"\1-\2", s)
    return s


def _norm_rev(val: str) -> str:
    if not val:
        return "R0"
    m = re.search(r"(\d+)", str(val))
    return f"R{int(m.group(1))}" if m else str(val).strip().upper()


def _make_link(rel_path: str, label: str) -> str:
    if not rel_path:
        return ""
    safe = rel_path.replace("\\", "/").replace('"', '""')
    return f'=HYPERLINK("{safe}","{label}")'


def find_matching_native_file(archive_root: Path, ref_no: str, revision: str = "") -> str:
    """يبحث تلقائياً في مجلد أصول الوورد والإكسل عن الملف الأصلي المطابق لنفس الرقم المرجعي لربطهما معاً!"""
    native_root = archive_root / "00_Editable_Originals - أصل الوورد والإكسل"
    if not native_root.exists() or not ref_no:
        return ""
    target_ref = _norm_ref(ref_no)
    target_rev = _norm_rev(revision) if revision else ""
    best = ""
    for p in native_root.rglob("*"):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in {".docx", ".doc", ".xlsx", ".xlsm", ".xls"}:
            continue
        stem_norm = _norm_ref(p.stem)
        if target_ref and target_ref in stem_norm:
            rel = str(p.relative_to(archive_root)).replace("\\", "/")
            if target_rev and target_rev in _norm_rev(p.stem):
                return rel
            best = rel
    return best


def _try_update_existing_native_row(wb, record: dict) -> bool:
    """إذا كان المستند مسجلاً مسبقاً عند إنشاء ملف Word/Excel، يحدّث نفس الصف بالنسخة المختومة وحالة الاعتماد!"""
    if record.get("is_native"):
        return False
    meta = record.get("meta") or {}
    ref = _norm_ref(meta.get("ref_no") or "")
    if not ref:
        return False
    doc_type = record.get("doc_type", "")
    scan_link = _make_link(record.get("dest_rel", ""), "فتح المختوم PDF")
    native_link = _make_link(record.get("native_rel", ""), "فتح أصل Word/Excel")
    date_v = meta.get("date") or dt.date.today().isoformat()
    status_v = meta.get("status") or ""
    note_v = meta.get("notes") or ""

    def _style_link(cell, formula: str):
        cell.value = formula
        cell.font = Font(name="Calibri", size=10.5, bold=True, color=LINK_CLR, underline="single")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if doc_type == "submittal" and "01_الاعتمادات" in wb.sheetnames:
        ws = wb["01_الاعتمادات"]
        rev = _norm_rev(meta.get("revision") or "R00")
        for r in range(4, ws.max_row + 1):
            row_ref = _norm_ref(ws.cell(r, 6).value or "")
            row_rev = _norm_rev(ws.cell(r, 7).value or "R00")
            if row_ref == ref and row_rev == rev:
                ws.cell(r, 3, date_v)
                if meta.get("authority"):
                    ws.cell(r, 9, meta["authority"])
                if status_v:
                    ws.cell(r, 12, status_v)
                if note_v:
                    old_n = str(ws.cell(r, 13).value or "").strip()
                    ws.cell(r, 13, f"{old_n} | {note_v}".strip(" |") if old_n else note_v)
                if native_link and not ws.cell(r, 14).value:
                    _style_link(ws.cell(r, 14), native_link)
                if scan_link:
                    _style_link(ws.cell(r, 15), scan_link)
                return True

    elif doc_type == "letter" and "04_الصادر_والوارد" in wb.sheetnames:
        ws = wb["04_الصادر_والوارد"]
        for r in range(4, ws.max_row + 1):
            if _norm_ref(ws.cell(r, 4).value or "") == ref:
                if status_v:
                    ws.cell(r, 10, status_v)
                else:
                    ws.cell(r, 10, "تم الاستلام / مختوم")
                if note_v:
                    old_n = str(ws.cell(r, 11).value or "").strip()
                    ws.cell(r, 11, f"{old_n} | {note_v}".strip(" |") if old_n else note_v)
                if native_link and not ws.cell(r, 12).value:
                    _style_link(ws.cell(r, 12), native_link)
                if scan_link:
                    _style_link(ws.cell(r, 13), scan_link)
                return True

    elif doc_type == "invoice" and "03_الفواتير" in wb.sheetnames:
        ws = wb["03_الفواتير"]
        in_party = _core_party_tokens(meta.get("from_party") or "")
        for r in range(4, ws.max_row + 1):
            if _norm_ref(ws.cell(r, 4).value or "") == ref:
                row_party = _core_party_tokens(ws.cell(r, 5).value or "")
                if in_party and row_party and in_party != row_party:
                    continue
                if _num_or_none(meta.get("gross_amount")) is not None:
                    ws.cell(r, 8, _num_or_none(meta.get("gross_amount")))
                if _num_or_none(meta.get("deductions_amount")) is not None:
                    ws.cell(r, 9, _num_or_none(meta.get("deductions_amount")))
                if _num_or_none(meta.get("net_amount")) is not None:
                    ws.cell(r, 10, _num_or_none(meta.get("net_amount")))
                if status_v:
                    ws.cell(r, 12, status_v)
                if note_v:
                    old_n = str(ws.cell(r, 13).value or "").strip()
                    ws.cell(r, 13, f"{old_n} | {note_v}".strip(" |") if old_n else note_v)
                if scan_link:
                    _style_link(ws.cell(r, 14), scan_link)
                return True

    elif doc_type == "petty_cash" and "02_العهدة" in wb.sheetnames:
        ws = wb["02_العهدة"]
        for r in range(4, ws.max_row + 1):
            if _norm_ref(ws.cell(r, 4).value or "") == ref and not ws.cell(r, 13).value:
                amt = _num_or_none(meta.get("net_amount")) or _num_or_none(meta.get("gross_amount"))
                if amt is not None and ws.cell(r, 9).value is None:
                    ws.cell(r, 9, amt)
                if status_v:
                    ws.cell(r, 11, status_v)
                if scan_link:
                    _style_link(ws.cell(r, 13), scan_link)
                return True

    elif doc_type == "contract_ipc" and "05_العقود_والمستخلصات" in wb.sheetnames:
        ws = wb["05_العقود_والمستخلصات"]
        in_party = _core_party_tokens(meta.get("from_party") or meta.get("to_party") or "")
        for r in range(4, ws.max_row + 1):
            row_ref = str(ws.cell(r, 7).value or "")
            base_row_ref = row_ref.split("(")[0].strip()
            if _norm_ref(row_ref) == ref or _norm_ref(base_row_ref) == ref:
                row_party = _core_party_tokens(ws.cell(r, 5).value or "")
                if in_party and row_party and in_party != row_party:
                    continue  # لا يخلط مستخلص IPC-01 للمقاول علي مع IPC-01 لمقاول آخر!
                if meta.get("subject") and not ws.cell(r, 8).value:
                    ws.cell(r, 8, meta.get("subject"))
                if _num_or_none(meta.get("gross_amount")) is not None:
                    ws.cell(r, 9, _num_or_none(meta.get("gross_amount")))
                if _num_or_none(meta.get("deductions_amount")) is not None:
                    ws.cell(r, 10, _num_or_none(meta.get("deductions_amount")))
                if _num_or_none(meta.get("net_amount")) is not None:
                    ws.cell(r, 11, _num_or_none(meta.get("net_amount")))
                if status_v:
                    ws.cell(r, 13, status_v)
                if note_v:
                    old_n = str(ws.cell(r, 14).value or "").strip()
                    ws.cell(r, 14, f"{old_n} | {note_v}".strip(" |") if old_n else note_v)
                if native_link and not ws.cell(r, 15).value:
                    _style_link(ws.cell(r, 15), native_link)
                if scan_link:
                    _style_link(ws.cell(r, 16), scan_link)
                return True

    return False


def build_excel_records(record: dict, xlsx_path: Path, skip_domain_sheet: bool = False) -> list[tuple[str, list, list[int], str]]:
    """يجهّز الصفوف المطلوب إضافتها مع ربط أصل Word/Excel والنسخة المختومة PDF جنباً إلى جنب."""
    out = []
    is_native = bool(record.get("is_native"))
    dest_rel = record.get("dest_rel", "")
    native_rel = dest_rel if is_native else record.get("native_rel", "")
    scan_rel = "" if is_native else dest_rel

    native_link = _make_link(native_rel, "فتح أصل Word/Excel")
    scan_link = _make_link(scan_rel, "فتح المختوم PDF")
    any_link = native_link if is_native else _make_link(dest_rel, "فتح الملف")

    meta = record.get("meta") or {}
    doc_type = record.get("doc_type", "other")
    action = record.get("action", "ARCHIVED")
    currency = str(meta.get("currency") or "KWD").upper()
    project = meta.get("project") or record.get("project") or "عام"
    date_v = meta.get("date") or dt.date.today().isoformat()

    if not skip_domain_sheet and action in ("ARCHIVED", "UPGRADED", "NATIVE_CREATED"):
        if doc_type == "submittal":
            sub_date = date_v if is_native else ""
            stamp_date = "" if is_native else date_v
            out.append(("01_الاعتمادات", [
                0, sub_date or date_v, stamp_date, project, meta.get("discipline") or "",
                meta.get("ref_no") or "", meta.get("revision") or "R00",
                meta.get("submittal_type") or "", meta.get("authority") or "",
                meta.get("subject") or "", meta.get("from_party") or "",
                meta.get("status") or ("Under Review (مقدم)" if is_native else ""),
                meta.get("notes") or "", native_link, scan_link,
            ], [], currency))
        elif doc_type == "petty_cash":
            amt = _num_or_none(meta.get("net_amount")) or _num_or_none(meta.get("gross_amount"))
            out.append(("02_العهدة", [
                0, date_v, project, meta.get("ref_no") or "", meta.get("from_party") or "",
                meta.get("subject") or "", meta.get("expense_category") or "أخرى",
                meta.get("payment_method") or "", amt, currency,
                meta.get("status") or "مسجل", meta.get("notes") or "", any_link,
            ], [9], currency))
        elif doc_type == "invoice":
            out.append(("03_الفواتير", [
                0, date_v, project, meta.get("ref_no") or "", meta.get("from_party") or "",
                meta.get("po_or_contract_no") or "", meta.get("subject") or "",
                _num_or_none(meta.get("gross_amount")), _num_or_none(meta.get("deductions_amount")),
                _num_or_none(meta.get("net_amount")), currency,
                meta.get("status") or "قيد المراجعة", meta.get("notes") or "", any_link,
            ], [8, 9, 10], currency))
        elif doc_type == "letter":
            d = str(meta.get("letter_direction") or ("Outgoing" if is_native else "Incoming"))
            d_ar = "صادر" if ("out" in d.lower() or "صادر" in d) else ("داخلي" if "int" in d.lower() else "وارد")
            out.append(("04_الصادر_والوارد", [
                0, date_v, d_ar, meta.get("ref_no") or "", meta.get("from_party") or "",
                meta.get("to_party") or "", project, meta.get("subject") or "",
                "نعم" if meta.get("needs_reply") else "لا",
                meta.get("status") or ("صادر — بانتظار ختم الاستلام" if is_native else ("بانتظار الرد" if meta.get("needs_reply") else "للعلم والحفظ")),
                meta.get("notes") or "", native_link, scan_link,
            ], [], currency))
        elif doc_type == "contract_ipc":
            out.append(("05_العقود_والمستخلصات", [
                0, date_v, project, meta.get("po_or_contract_no") or "",
                meta.get("from_party") or meta.get("to_party") or "",
                meta.get("contract_subtype") or "IPC", meta.get("ref_no") or "",
                meta.get("subject") or "", _num_or_none(meta.get("gross_amount")),
                _num_or_none(meta.get("deductions_amount")), _num_or_none(meta.get("net_amount")),
                currency, meta.get("status") or ("مقدم للاعتماد" if is_native else "قيد المراجعة"),
                meta.get("notes") or "", native_link, scan_link,
            ], [9, 10, 11], currency))

        elif doc_type in ("permit", "site_photo", "custom", "whatsapp_chat", "other"):
            cat_label = {
                "permit": "ترخيص / موافقة رسمية",
                "site_photo": "صورة موقع / واتساب",
                "whatsapp_chat": "محادثة واتساب",
                "other": "مستند عام",
            }.get(doc_type, meta.get("custom_category") or "قسم مخصص")
            out.append(("06_تراخيص_وصور_وأقسام", [
                0, date_v, project, cat_label, meta.get("ref_no") or "",
                meta.get("authority") or meta.get("from_party") or meta.get("to_party") or "",
                meta.get("subject") or record.get("orig_name", ""),
                _num_or_none(meta.get("net_amount")) or _num_or_none(meta.get("gross_amount")),
                currency, meta.get("status") or "موثق", meta.get("notes") or "", any_link,
            ], [8], currency))

    audit_sheet = "07_سجل_السكان_الكامل"
    out.append((audit_sheet, [
        0, record.get("timestamp", ""), record.get("orig_name", ""), action,
        doc_type, project, record.get("pages", 1),
        round(float(meta.get("confidence") or 1.0), 2) if meta else "",
        dest_rel, record.get("note") or meta.get("notes") or "", any_link,
    ], [], currency))
    return out


def flush_excel_queue(xlsx_path: Path, queue_path: Path, log_file: Path):
    """يكتب كل السجلات المعلقة ويحدّث الصفوف المفتوحة (Word/Excel -> PDF مختوم)؛ وينتظر بهدوء إن كان الإكسل مفتوحاً."""
    if not queue_path.exists():
        return
    pending = json.loads(queue_path.read_text(encoding="utf-8"))
    if not pending:
        return
    try:
        ensure_master_workbook(xlsx_path)
        wb = openpyxl.load_workbook(str(xlsx_path))
        for rec in pending:
            matched_existing = _try_update_existing_native_row(wb, rec)
            if matched_existing:
                rec["action"] = "MATCHED_NATIVE"
                log(_tr("watcher.linked_stamped", ref=rec.get("meta", {}).get("ref_no", "")), log_file)
            for sheet_name, vals, mcols, curr in build_excel_records(rec, xlsx_path, skip_domain_sheet=matched_existing):
                if sheet_name in wb.sheetnames:
                    _append_styled_row(wb[sheet_name], vals, mcols, curr)
        wb.save(str(xlsx_path))
        queue_path.write_text("[]", encoding="utf-8")
    except PermissionError:
        log(_tr("watcher.excel_open_queued"), log_file)


def enqueue_and_flush(record: dict, xlsx_path: Path, queue_path: Path, log_file: Path):
    pending = []
    if queue_path.exists():
        try:
            pending = json.loads(queue_path.read_text(encoding="utf-8"))
        except Exception:
            pending = []
    pending.append(record)
    queue_path.write_text(json.dumps(pending, ensure_ascii=False, indent=2), encoding="utf-8")
    flush_excel_queue(xlsx_path, queue_path, log_file)


# =====================================================================
# المعالجة الكاملة لملف واحد قادم من السكانر
# =====================================================================
def archive_raw_original(src: Path, inbox_root: Path, sys_root: Path) -> Path:
    day = dt.date.today().isoformat()
    raw_dir = sys_root / "_01_Processed_Originals - أصل السكان" / day
    raw_dir.mkdir(parents=True, exist_ok=True)
    dest = raw_dir / src.name
    if dest.exists():
        stamp = dt.datetime.now().strftime("%H%M%S")
        dest = raw_dir / f"{src.stem}_{stamp}{src.suffix}"
        idx = 2
        while dest.exists():
            dest = raw_dir / f"{src.stem}_{stamp}_{idx}{src.suffix}"
            idx += 1
    shutil.move(str(src), str(dest))
    # إذا كان داخل مجلد فرعي فارغ داخل Scan_Inbox اترك المجلد الفرعي كما هو (لأزرار السكانر السريعة)
    return dest


def process_one_file(src: Path, cfg: dict, state: dict, state_path: Path, sys_root: Path, log_file: Path, analyzer_override=None):
    inbox_root = Path(cfg["inbox_folder"]).resolve()
    archive_root = Path(cfg["archive_folder"]).resolve()
    xlsx_path = archive_root / "Master_Archive_Log - السجل الشامل التلقائي.xlsx"
    queue_path = sys_root / "pending_excel_rows.json"

    rel_in_inbox = src.relative_to(inbox_root)
    subfolder_hint = rel_in_inbox.parts[0] if len(rel_in_inbox.parts) > 1 else ""
    orig_name = src.name
    size_bytes = src.stat().st_size
    digest = sha256_file(src)
    now_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    log(f"{_tr('watcher.new_paper')}: {rel_in_inbox}", log_file)

    # 1) كشف التكرار الحرفي (SHA-256) بدون استهلاك أي طلب AI
    hashes = state.setdefault("hashes", {})
    if digest in hashes:
        dup_dir = archive_root / "_02_Duplicates - مكررات"
        dup_dir.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        dup_path = dup_dir / f"{src.stem}_DUP_{stamp}{src.suffix}"
        idx = 2
        while dup_path.exists():
            dup_path = dup_dir / f"{src.stem}_DUP_{stamp}_{idx}{src.suffix}"
            idx += 1
        shutil.move(str(src), str(dup_path))
        rel_dup = str(dup_path.relative_to(archive_root)).replace("\\", "/")
        log(_tr("watcher.duplicate_moved"), log_file)
        enqueue_and_flush({
            "timestamp": now_str, "orig_name": orig_name, "action": "DUPLICATE_SKIP",
            "doc_type": "duplicate", "project": subfolder_hint or "—", "pages": 1,
            "dest_rel": rel_dup, "note": _tr("watcher.duplicate_moved"), "meta": {},
        }, xlsx_path, queue_path, log_file)
        return

    # 2) تجهيز الصور للقراءة البصرية (مع قراءة الكاش إن سبق تحليل نفس الملف)
    import ai_cache
    job_dir = sys_root / "_tmp_job"
    if job_dir.exists():
        shutil.rmtree(job_dir, ignore_errors=True)
    job_dir.mkdir(parents=True, exist_ok=True)
    w_cache_dir = ai_cache.project_ai_cache_dir(archive_root)
    w_cache_key = f"{digest}:{ai_cache.stable_sig(subfolder_hint, cfg.get('domain'), sorted(cfg.get('custom_categories') or []))}"

    try:
        cached_doc = ai_cache.get_cached(w_cache_dir, "doc_classify", w_cache_key)
        if cached_doc and isinstance(cached_doc.get("meta"), dict) and cached_doc["meta"]:
            meta = dict(cached_doc["meta"])
            pages = int(cached_doc.get("pages") or 1)
        else:
            images, embedded_text, pages = prepare_visual_pages(
                src, job_dir, dpi=int(cfg.get("render_dpi", 150)), stamp_dpi=int(cfg.get("stamp_zoom_dpi", 200))
            )
            known_projects = list(dict.fromkeys(
                list(cfg.get("known_projects") or []) + list(state.get("seen_projects") or [])
            ))

            engine = str(cfg.get("engine", "gemini")).lower()
            prompt = build_classification_prompt(
                image_names=[p.name for p in images],
                orig_name=orig_name,
                subfolder_hint=subfolder_hint,
                page_count=pages,
                embedded_text=embedded_text,
                existing_projects=known_projects[:30],
                default_currency=cfg.get("default_currency", "KWD"),
                include_at_refs=(engine == "gemini"),
                domain=str(cfg.get("domain") or "contracting"),
                custom_categories=list(cfg.get("custom_categories") or []),
            )

            if analyzer_override is not None:
                meta = analyzer_override(src, images, prompt)
            elif engine == "ollama":
                meta = analyze_with_ollama(prompt, images, cfg["ollama_url"], cfg["ollama_model"])
            else:
                meta = analyze_with_gemini(prompt, job_dir, cfg.get("gemini_model", ""), images=images, api_key=cfg.get("gemini_api_key", ""))
            if isinstance(meta, dict) and meta.get("doc_type") and not meta.get("is_blank_or_unreadable"):
                ai_cache.set_cached(w_cache_dir, "doc_classify", w_cache_key, {
                    "meta": meta,
                    "pages": int(pages or 1),
                })

    except Exception:
        log(_tr("watcher.read_failed_retry"), log_file)
        shutil.rmtree(job_dir, ignore_errors=True)
        return
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)

    conf = float(meta.get("confidence") or 0.0)
    unreadable = bool(meta.get("is_blank_or_unreadable"))

    # 3) فحص الثقة / الأوراق الفارغة أو المقلوبة
    if unreadable or conf < float(cfg.get("min_confidence", 0.58)):
        rev_dir = archive_root / "_03_Needs_Review - يحتاج مراجعة"
        rev_dir.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        rev_dest = rev_dir / f"{stamp}_{clean_seg(src.stem, 40)}{src.suffix.lower()}"
        shutil.copy2(str(src), str(rev_dest))
        archive_raw_original(src, inbox_root, sys_root)
        rel_rev = str(rev_dest.relative_to(archive_root)).replace("\\", "/")
        hashes[digest] = rel_rev
        state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        log(_tr("watcher.unclear_saved"), log_file)
        enqueue_and_flush({
            "timestamp": now_str, "orig_name": orig_name, "action": "NEEDS_REVIEW",
            "doc_type": str(meta.get("doc_type") or "unknown"), "project": subfolder_hint or "—",
            "pages": pages, "dest_rel": rel_rev,
            "note": meta.get("notes") or _tr("watcher.unclear_saved"), "meta": meta,
        }, xlsx_path, queue_path, log_file)
        return

    # 4) تحديد المسار النهائي — صور الموقع/الواتساب الميدانية تبقى صورة أصلية عالية الدقة ولا تُحوَّل لـ PDF!
    raw_dt = str(meta.get("doc_type") or "other").strip().lower()
    want_pdf = bool(cfg.get("convert_images_to_pdf")) and src.suffix.lower() != ".pdf" and raw_dt != "site_photo"
    out_ext = ".pdf" if want_pdf else src.suffix.lower()
    rel_target, doc_type = compute_destination(meta, out_ext, cfg, subfolder_hint)
    final_path = archive_root / rel_target
    final_path.parent.mkdir(parents=True, exist_ok=True)

    action = "ARCHIVED"
    extra_note = ""
    if final_path.exists():
        # قارن عدد الصفحات والحجم (النسخة ذات الصفحات الأكثر تفوز دائماً!)
        old_ok, old_pages = is_file_unlocked_and_valid(final_path)
        old_size = final_path.stat().st_size if old_ok else 0
        if (pages, size_bytes) > (old_pages, old_size) and doc_type == "submittal":
            sup_dir = final_path.parent.parent / "_قبل التعديل (Superseded)"
            sup_dir.mkdir(parents=True, exist_ok=True)
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            sup_path = sup_dir / f"{final_path.stem}_prev_{old_pages}p_{stamp}{final_path.suffix}"
            shutil.move(str(final_path), str(sup_path))
            action = "UPGRADED"
            extra_note = _tr("watcher.version_upgraded")
        else:
            # مستند مختلف بنفس التاريخ والمورد (مثلاً فاتورتان لنفس المحل في نفس اليوم) -> أضف رقماً تسلسلياً
            counter = 2
            while True:
                cand = final_path.with_name(f"{final_path.stem}_({counter}){final_path.suffix}")
                if not cand.exists():
                    final_path = cand
                    break
                counter += 1

    if out_ext == ".pdf" and src.suffix.lower() != ".pdf":
        convert_image_to_pdf(src, final_path)
    else:
        shutil.copy2(str(src), str(final_path))

    # حفظ الأصل الخام في أرشيف السكانر وتفريغ Scan_Inbox
    archive_raw_original(src, inbox_root, sys_root)

    rel_final = str(final_path.relative_to(archive_root)).replace("\\", "/")
    hashes[digest] = rel_final
    proj_clean = clean_seg(meta.get("project") or "", 45)
    if proj_clean and proj_clean != "عام":
        seen = state.setdefault("seen_projects", [])
        if proj_clean not in seen:
            seen.append(proj_clean)
    state["total_archived"] = int(state.get("total_archived", 0)) + 1
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    # 5) البحث التلقائي عن أصل Word/Excel المطابق + تحديث السجل الشامل في الإكسل فوراً
    native_rel = find_matching_native_file(archive_root, meta.get("ref_no") or "", meta.get("revision") or "")
    if extra_note:
        meta["notes"] = f"{extra_note} | {meta.get('notes') or ''}".strip(" |")
    enqueue_and_flush({
        "timestamp": now_str, "orig_name": orig_name, "action": action,
        "doc_type": doc_type, "project": proj_clean or subfolder_hint or cfg.get("default_project", "عام"),
        "pages": pages, "dest_rel": rel_final, "native_rel": native_rel,
        "is_native": False, "note": meta.get("notes") or "", "meta": meta,
    }, xlsx_path, queue_path, log_file)

    log(_tr("watcher.archived_ok", name=final_path.name), log_file)


def register_native_document(native_file: Path, meta: dict, cfg: dict):
    """يسجّل ملف Word أو Excel أصلي تم إصداره حديثاً ويحفظه في 00_Editable_Originals بانتظار عودة السكان المختوم!"""
    archive_root = Path(cfg["archive_folder"]).resolve()
    sys_root = archive_root / "_00_Scanner_System"
    archive_root.mkdir(parents=True, exist_ok=True)
    sys_root.mkdir(parents=True, exist_ok=True)
    xlsx_path = archive_root / "Master_Archive_Log - السجل الشامل التلقائي.xlsx"
    queue_path = sys_root / "pending_excel_rows.json"
    log_file = sys_root / "watcher.log"

    doc_type = str(meta.get("doc_type") or "submittal").strip().lower()
    nat_resolved = native_file.resolve()
    try:
        # إذا كان الملف محفوظاً أصلاً داخل مجلد أرشيف المشروع (مثلاً داخل فولدر المقاول علي)، اربطه في مكانه مباشرة دون تكرار!
        rel_native = str(nat_resolved.relative_to(archive_root)).replace("\\", "/")
    except ValueError:
        sub_dir = clean_seg(meta.get("discipline") or doc_type, 30)
        target_dir = archive_root / "00_Editable_Originals - أصل الوورد والإكسل" / sub_dir
        target_dir.mkdir(parents=True, exist_ok=True)
        dest_path = target_dir / clean_seg(native_file.name, 95)
        if nat_resolved != dest_path.resolve():
            shutil.copy2(str(native_file), str(dest_path))
        rel_native = str(dest_path.relative_to(archive_root)).replace("\\", "/")
    enqueue_and_flush({
        "timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "orig_name": native_file.name,
        "action": "NATIVE_CREATED",
        "doc_type": doc_type,
        "project": meta.get("project") or cfg.get("default_project", "عام"),
        "pages": 1,
        "dest_rel": rel_native,
        "native_rel": rel_native,
        "is_native": True,
        "note": meta.get("notes") or _tr("watcher.original_saved", ref=meta.get("ref_no", "")),
        "meta": meta,
    }, xlsx_path, queue_path, log_file)
    log(_tr("watcher.original_saved", ref=meta.get("ref_no", "")), log_file)


def process_native_draft_file(src: Path, cfg: dict, state: dict, state_path: Path, sys_root: Path, log_file: Path, analyzer_override=None):
    """يلتقط أي ملف Word (.docx) أو Excel (.xlsx/.xlsm) يضعه المستخدم في 2_Drafts_Word_Excel،
    يقرأ محتواه، يسمّيه ويحفظه في 00_Editable_Originals، ويسجّله فوراً في اللوج بانتظار عودة السكان المختوم!"""
    drafts_root = Path(cfg["drafts_folder"]).resolve()
    archive_root = Path(cfg["archive_folder"]).resolve()
    xlsx_path = archive_root / "Master_Archive_Log - السجل الشامل التلقائي.xlsx"
    queue_path = sys_root / "pending_excel_rows.json"

    rel_in_drafts = src.relative_to(drafts_root)
    subfolder_hint = rel_in_drafts.parts[0] if len(rel_in_drafts.parts) > 1 else ""
    digest = sha256_file(src)
    hashes = state.setdefault("native_hashes", {})
    if digest in hashes:
        dup_dir = archive_root / "_02_Duplicates - مكررات"
        dup_dir.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.move(str(src), str(dup_dir / f"{src.stem}_DUP_{stamp}{src.suffix}"))
        log(_tr("watcher.duplicate_moved"), log_file)
        return

    log(f"{_tr('watcher.new_office_file')}: {rel_in_drafts}", log_file)
    import ai_cache
    job_dir = sys_root / "_tmp_native_job"
    shutil.rmtree(job_dir, ignore_errors=True)
    job_dir.mkdir(parents=True, exist_ok=True)
    w_cache_dir = ai_cache.project_ai_cache_dir(archive_root)
    w_cache_key = f"{digest}:{ai_cache.stable_sig(subfolder_hint, cfg.get('domain'), sorted(cfg.get('custom_categories') or []))}"

    try:
        cached_doc = ai_cache.get_cached(w_cache_dir, "doc_classify", w_cache_key)
        if cached_doc and isinstance(cached_doc.get("meta"), dict) and cached_doc["meta"]:
            meta = dict(cached_doc["meta"])
        else:
            office_text = extract_native_office_text(src)
            known_projects = list(dict.fromkeys(
                list(cfg.get("known_projects") or []) + list(state.get("seen_projects") or [])
            ))
            prompt = build_classification_prompt(
                image_names=[],
                orig_name=src.name,
                subfolder_hint=subfolder_hint,
                page_count=1,
                embedded_text=f"هذا ملف أصلي قابل للتعديل ({src.suffix}) تم تجهيزه قبل الطباعة/الختم:\n{office_text}",
                existing_projects=known_projects[:30],
                default_currency=cfg.get("default_currency", "KWD"),
                include_at_refs=False,
                domain=str(cfg.get("domain") or "contracting"),
                custom_categories=list(cfg.get("custom_categories") or []),
            )
            engine = str(cfg.get("engine", "gemini")).lower()
            if analyzer_override is not None:
                meta = analyzer_override(src, [], prompt)
            elif engine == "ollama":
                meta = analyze_with_ollama(prompt, [], cfg["ollama_url"], cfg["ollama_model"])
            else:
                meta = analyze_with_gemini(prompt, job_dir, cfg.get("gemini_model", ""), images=[], api_key=cfg.get("gemini_api_key", ""))
            if isinstance(meta, dict) and meta.get("doc_type") and not meta.get("is_blank_or_unreadable"):
                ai_cache.set_cached(w_cache_dir, "doc_classify", w_cache_key, {"meta": meta, "pages": 1})
    except Exception:
        log(_tr("watcher.read_failed_retry"), log_file)
        shutil.rmtree(job_dir, ignore_errors=True)
        return
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)

    doc_type = str(meta.get("doc_type") or "other").strip().lower()
    if doc_type not in CATEGORY_FOLDERS:
        doc_type = "other"
    meta["doc_type"] = doc_type
    if not meta.get("project") and subfolder_hint:
        meta["project"] = subfolder_hint
    if not meta.get("status"):
        meta["status"] = "صادر — بانتظار ختم الاستلام" if doc_type == "letter" else "Under Review (مقدم للاعتماد)"

    if doc_type in ("contract_ipc", "invoice"):
        # يبحث عن فولدر المقاول/المورد الموجود (مثال: "علي") ويضيف ملف الإكسل/الوورد بداخله مباشرة!
        rel_dir, final_name = compute_destination(meta, src.suffix.lower(), cfg, subfolder_hint)
        target_dir = archive_root / rel_dir
        target_dir.mkdir(parents=True, exist_ok=True)
        dest_path = target_dir / final_name
    else:
        sub_group = clean_seg(
            (meta.get("discipline") if doc_type == "submittal" else None)
            or CATEGORY_FOLDERS.get(doc_type, "General"),
            40,
        )
        target_dir = archive_root / "00_Editable_Originals - أصل الوورد والإكسل" / sub_group
        target_dir.mkdir(parents=True, exist_ok=True)
        suggested = clean_seg(meta.get("suggested_filename") or src.stem, 85)
        if suggested.lower().endswith(src.suffix.lower()):
            suggested = suggested[: -len(src.suffix)].rstrip(" .-_")
        dest_path = target_dir / f"{suggested}{src.suffix.lower()}"
    if dest_path.exists():
        c = 2
        while True:
            cand = dest_path.with_name(f"{dest_path.stem}_({c}){dest_path.suffix}")
            if not cand.exists():
                dest_path = cand
                break
            c += 1

    shutil.copy2(str(src), str(dest_path))
    # حفظ نسخة من الملف الخام في أرشيف النظام وتفريغ مجلد المسودات
    raw_dir = sys_root / "_01_Processed_Originals - أصل السكان" / f"{dt.date.today().isoformat()}_Native"
    raw_dir.mkdir(parents=True, exist_ok=True)
    raw_dest = raw_dir / src.name
    if raw_dest.exists():
        raw_dest = raw_dir / f"{src.stem}_{dt.datetime.now().strftime('%H%M%S')}{src.suffix}"
    shutil.move(str(src), str(raw_dest))

    rel_native = str(dest_path.relative_to(archive_root)).replace("\\", "/")
    hashes[digest] = rel_native
    proj_clean = clean_seg(meta.get("project") or "", 45)
    if proj_clean and proj_clean != "عام":
        seen = state.setdefault("seen_projects", [])
        if proj_clean not in seen:
            seen.append(proj_clean)
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    enqueue_and_flush({
        "timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "orig_name": src.name,
        "action": "NATIVE_CREATED",
        "doc_type": doc_type,
        "project": proj_clean or subfolder_hint or cfg.get("default_project", "عام"),
        "pages": 1,
        "dest_rel": rel_native,
        "native_rel": rel_native,
        "is_native": True,
        "note": meta.get("notes") or _tr("watcher.original_saved", ref=meta.get("ref_no", "") or doc_type),
        "meta": meta,
    }, xlsx_path, queue_path, log_file)
    log(_tr("watcher.original_saved", ref=meta.get("ref_no", "") or doc_type), log_file)


def process_whatsapp_chat_txt(src: Path, cfg: dict, state: dict, state_path: Path, sys_root: Path, log_file: Path):
    """يحفظ سجل محادثة واتساب المصدَّر (_chat.txt / WhatsApp Chat with ....txt) ويوثّقه في اللوج بالهايبر لينك."""
    inbox_root = Path(cfg["inbox_folder"]).resolve()
    archive_root = Path(cfg["archive_folder"]).resolve()
    xlsx_path = archive_root / "Master_Archive_Log - السجل الشامل التلقائي.xlsx"
    queue_path = sys_root / "pending_excel_rows.json"
    digest = sha256_file(src)
    hashes = state.setdefault("hashes", {})
    if digest in hashes:
        archive_raw_original(src, inbox_root, sys_root)
        return
    today = dt.date.today().isoformat()
    chat_dir = archive_root / CATEGORY_FOLDERS["site_photo"] / "00_WhatsApp_Chats - محادثات واتساب"
    chat_dir.mkdir(parents=True, exist_ok=True)
    clean_n = clean_seg(src.stem, 55) or "WhatsApp_Chat"
    if clean_n.lower() == "chat":
        clean_n = f"WhatsApp_Chat_{clean_seg(src.parent.name, 40)}"
    dest = chat_dir / f"{today}_{clean_n}.txt"
    if dest.exists():
        dest = chat_dir / f"{today}_{clean_n}_{dt.datetime.now().strftime('%H%M%S')}.txt"
    shutil.copy2(str(src), str(dest))
    try:
        lines_cnt = len(src.read_text(encoding="utf-8", errors="replace").splitlines())
    except Exception:
        lines_cnt = 0
    archive_raw_original(src, inbox_root, sys_root)
    rel = str(dest.relative_to(archive_root)).replace("\\", "/")
    hashes[digest] = rel
    state["total_archived"] = int(state.get("total_archived", 0)) + 1
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    enqueue_and_flush({
        "timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "orig_name": src.name,
        "action": "ARCHIVED",
        "doc_type": "whatsapp_chat",
        "project": cfg.get("default_project", "عام"),
        "pages": 1,
        "dest_rel": rel,
        "is_native": False,
        "note": _tr("watcher.whatsapp_saved"),
        "meta": {"date": today, "project": cfg.get("default_project", "عام"), "from_party": "WhatsApp", "subject": clean_n, "status": "OK", "notes": str(lines_cnt)},
    }, xlsx_path, queue_path, log_file)
    log(_tr("watcher.whatsapp_saved"), log_file)


def scan_inbox_candidates(inbox_root: Path, exclude_root: Path | None = None, allowed_exts: set[str] | None = None) -> list[Path]:
    exts = allowed_exts if allowed_exts is not None else SUPPORTED_EXTS
    found = []
    if not inbox_root.exists():
        return found
    ex = exclude_root.resolve() if exclude_root else None
    for dp, dns, fns in os.walk(inbox_root):
        cur = Path(dp).resolve()
        filtered = []
        for d in dns:
            if d.startswith(".") or d.startswith("_"):
                continue
            cand_dir = (cur / d).resolve()
            if ex and (cand_dir == ex or ex in cand_dir.parents):
                continue
            if d.startswith("0") and " - " in d:
                continue
            filtered.append(d)
        dns[:] = filtered
        for fn in sorted(fns):
            if fn.startswith("~$") or fn.startswith(".") or fn.lower() in {"thumbs.db", "desktop.ini"}:
                continue
            p = Path(dp) / fn
            ext = p.suffix.lower()
            if ext in TEMP_EXTS or ext not in exts:
                continue
            found.append(p)
    return found


def load_or_create_config(
    config_path: Path,
    inbox_arg: str | None = None,
    archive_arg: str | None = None,
    layout_arg: str | None = None,
    drafts_arg: str | None = None,
) -> dict:
    cfg = dict(DEFAULT_CONFIG)
    if config_path.exists():
        try:
            saved = json.loads(config_path.read_text(encoding="utf-8"))
            cfg.update(saved)
        except Exception:
            pass
    if inbox_arg:
        cfg["inbox_folder"] = str(Path(inbox_arg).resolve())
    if drafts_arg:
        cfg["drafts_folder"] = str(Path(drafts_arg).resolve())
    if archive_arg:
        cfg["archive_folder"] = str(Path(archive_arg).resolve())
        if not drafts_arg and "drafts_folder" not in (json.loads(config_path.read_text(encoding="utf-8")) if config_path.exists() else {}):
            cfg["drafts_folder"] = str(Path(archive_arg).resolve().parent / "2_Drafts_Word_Excel")
    if layout_arg:
        cfg["folder_layout"] = layout_arg
    from security_utils import atomic_write_json, scrub_secrets
    cfg = scrub_secrets(cfg)
    atomic_write_json(config_path, cfg, scrub=True)
    return cfg


def run_watcher(cfg: dict, once: bool = False, analyzer_override=None):
    inbox_root = Path(cfg["inbox_folder"]).resolve()
    drafts_root = Path(cfg.get("drafts_folder") or (Path(cfg["archive_folder"]).resolve().parent / "2_Drafts_Word_Excel")).resolve()
    archive_root = Path(cfg["archive_folder"]).resolve()
    sys_root = archive_root / "_00_Scanner_System"
    inbox_root.mkdir(parents=True, exist_ok=True)
    drafts_root.mkdir(parents=True, exist_ok=True)
    archive_root.mkdir(parents=True, exist_ok=True)
    sys_root.mkdir(parents=True, exist_ok=True)

    state_path = sys_root / "archive_state.json"
    log_file = sys_root / "watcher.log"
    xlsx_path = archive_root / "Master_Archive_Log - السجل الشامل التلقائي.xlsx"
    queue_path = sys_root / "pending_excel_rows.json"

    state = {}
    if state_path.exists():
        try:
            state = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            state = {}

    ensure_master_workbook(xlsx_path)

    log("=" * 70, log_file)
    log(_tr("watcher.started"), log_file)
    log(f"{_tr('overview.folder_scanner')}: {inbox_root}", log_file)
    log(f"{_tr('overview.folder_archive')}: {archive_root}", log_file)
    log(f"{_tr('nav.logs')}: {xlsx_path.name}", log_file)
    log("=" * 70, log_file)

    size_tracker: dict[str, tuple[int, float]] = {}
    settle = float(cfg.get("settle_seconds", 6))
    poll = float(cfg.get("poll_seconds", 4))

    while True:
        # محاولة تفريغ أي صفوف إكسل معلقة (إذا أغلق المستخدم ملف الإكسل)
        flush_excel_queue(xlsx_path, queue_path, log_file)

        # نعالج أصول Word/Excel أولاً ثم ملفات السكانر المختومة لتتم المطابقة الفورية!
        native_cands = [("native", p) for p in scan_inbox_candidates(drafts_root, exclude_root=archive_root, allowed_exts=NATIVE_DOC_EXTS)]
        # دعم إضافي: إذا حُفظ ملف Word/Excel مباشرة داخل فولدر السكانر يعالجه كأصل تلقائياً!
        inbox_native = (
            [("native", p) for p in scan_inbox_candidates(inbox_root, exclude_root=archive_root, allowed_exts=NATIVE_DOC_EXTS)]
            if inbox_root != drafts_root else []
        )
        scan_cands = [("scan", p) for p in scan_inbox_candidates(inbox_root, exclude_root=archive_root, allowed_exts=SUPPORTED_EXTS)]
        chat_cands = [("chat", p) for p in scan_inbox_candidates(inbox_root, exclude_root=archive_root, allowed_exts=CHAT_LOG_EXTS)]

        all_jobs = native_cands + inbox_native + scan_cands + chat_cands
        active_keys = set()
        now = time.time()

        for kind, p in all_jobs:
            key = str(p)
            active_keys.add(key)
            try:
                sz = p.stat().st_size
            except OSError:
                continue
            if sz == 0:
                continue

            def _dispatch(target_p: Path, job_kind: str):
                if job_kind == "chat":
                    process_whatsapp_chat_txt(target_p, cfg, state, state_path, sys_root, log_file)
                elif job_kind == "native":
                    # ضبط الجذر مؤقتاً إن جاء ملف Word/Excel من inbox_root
                    if inbox_root in target_p.parents:
                        tmp_cfg = dict(cfg, drafts_folder=str(inbox_root))
                        process_native_draft_file(target_p, tmp_cfg, state, state_path, sys_root, log_file, analyzer_override=analyzer_override)
                    else:
                        process_native_draft_file(target_p, cfg, state, state_path, sys_root, log_file, analyzer_override=analyzer_override)
                else:
                    process_one_file(target_p, cfg, state, state_path, sys_root, log_file, analyzer_override=analyzer_override)

            if once:
                ok, _ = is_file_unlocked_and_valid(p)
                if ok:
                    _dispatch(p, kind)
                continue

            prev = size_tracker.get(key)
            if prev is None or prev[0] != sz:
                size_tracker[key] = (sz, now)
                continue

            if (now - prev[1]) >= settle:
                ok, _ = is_file_unlocked_and_valid(p)
                if ok:
                    size_tracker.pop(key, None)
                    _dispatch(p, kind)

        for k in list(size_tracker.keys()):
            if k not in active_keys:
                size_tracker.pop(k, None)

        if once:
            break
        time.sleep(poll)


def main():
    default_cfg_path = Path.home() / ".gemini" / "scan_watcher_config.json"
    ap = argparse.ArgumentParser(description="Auto Scan-Folder Watcher & AI Document Organizer ($0 Free)")
    ap.add_argument("--config", default=str(default_cfg_path), help="Path to JSON config file")
    ap.add_argument("--inbox", default=None, help="Set/override scanner inbox folder (PDF/JPG/PNG)")
    ap.add_argument("--drafts", default=None, help="Set/override native Word/Excel drafts folder (.docx/.xlsx/.xlsm)")
    ap.add_argument("--archive", default=None, help="Set/override organized archive destination folder")
    ap.add_argument("--layout", choices=["by_type", "by_project_then_type"], default=None, help="Folder hierarchy mode")
    ap.add_argument("--engine", choices=["gemini", "ollama"], default=None, help="Free AI vision engine")
    ap.add_argument("--setup-only", action="store_true", help="Save config + create folders & Excel log, then exit")
    ap.add_argument("--once", action="store_true", help="Process all ready files in inbox & drafts once and exit")
    ap.add_argument("--register-native", default=None, help="Copy & register a newly created .docx/.xlsx original")
    ap.add_argument("--meta-json", default=None, help="Metadata JSON file for --register-native")
    args = ap.parse_args()

    cfg = load_or_create_config(Path(args.config), args.inbox, args.archive, args.layout, args.drafts)
    if args.engine:
        cfg["engine"] = args.engine
        Path(args.config).write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.register_native:
        meta = json.loads(Path(args.meta_json).read_text(encoding="utf-8")) if args.meta_json else {}
        register_native_document(Path(args.register_native), meta, cfg)
        return

    if args.setup_only:
        inbox = Path(cfg["inbox_folder"]).resolve()
        drafts = Path(cfg["drafts_folder"]).resolve()
        arch = Path(cfg["archive_folder"]).resolve()
        inbox.mkdir(parents=True, exist_ok=True)
        drafts.mkdir(parents=True, exist_ok=True)
        arch.mkdir(parents=True, exist_ok=True)
        ensure_master_workbook(arch / "Master_Archive_Log - السجل الشامل التلقائي.xlsx")
        print(f"[OK] Config saved       : {args.config}")
        print(f"[OK] Scan Inbox (PDF)   : {inbox}")
        print(f"[OK] Word/Excel Drafts  : {drafts}")
        print(f"[OK] Organized Archive  : {arch}")
        return

    try:
        run_watcher(cfg, once=args.once)
    except KeyboardInterrupt:
        print(f"\n{_tr('watcher.stopped')}")


if __name__ == "__main__":
    main()
