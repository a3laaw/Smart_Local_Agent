#!/usr/bin/env python3
"""
whatsapp_bridge.py — جسر واتساب المزدوج الآمن 100% لبوابة المشاريع الذكية
===========================================================================
يدعم وضعين معاً بدون أي مكتبات خارجية إضافية (Python stdlib فقط) وبدون أي خطر حظر:

1) الوضع الأول — WhatsApp Web / Desktop السريع (مجاني 100% $0، بدون فيزا، أفراد وجروبات):
   - ينسخ الملف الفعلي (PDF / Word / Excel / صورة) تلقائياً إلى حافظة النظام (Clipboard) على Windows/macOS.
   - يولّد رابط WhatsApp Web الرسمي مع تنسيق رقم الهاتف (يدعم الكويت +965 تلقائياً) والرسالة الرسمية جاهزة.
   - يضغط المستخدم Ctrl+V ثم Enter فقط فيُرفق الملف ويُرسل في ثانية واحدة!

2) الوضع الثاني — Meta WhatsApp Business Cloud API الرسمي (إرسال واستلام أوتوماتيكي كامل):
   - إرسال ملفات PDF / Word (.docx) / Excel (.xlsx) / صور مباشرة عبر Graph API (أول 1,000 رسالة خدمة/شهر مجاناً).
   - دعم القوالب المعتمدة (Utility Template) عند الإرسال خارج نافذة الـ 24 ساعة.
   - استقبال الصور والملفات الواردة (Webhook مجاني بلا حدود) وتوجيهها تلقائياً لفولدر المشروع المناسب
     لتُقرأ وتُفرز وتُسجَّل في شيت Excel بالهايبر لينك مع رد استلام تلقائي اختياري!
"""
from __future__ import annotations

import hashlib
import hmac
import json
import mimetypes
import os
import re
import secrets
import subprocess
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

GRAPH_API_BASE = "https://graph.facebook.com/v21.0"

MIME_OVERRIDES = {
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc": "application/msword",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    ".xls": "application/vnd.ms-excel",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".jfif": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".txt": "text/plain",
}

MIME_TO_EXT = {
    "application/pdf": ".pdf",
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-excel": ".xls",
    "text/plain": ".txt",
}

DEFAULT_WA_CONFIG = {
    "default_country_code": "965",   # الكويت افتراضياً
    "phone_number_id": "",
    "access_token": "",
    "app_secret": "",
    "verify_token": "",
    "template_name": "",
    "template_lang": "ar",
    "auto_reply_ack": True,
    "default_project_id": "",
}


def load_wa_config(cfg_path: Path) -> dict:
    cfg = dict(DEFAULT_WA_CONFIG)
    if cfg_path.exists():
        try:
            saved = json.loads(cfg_path.read_text(encoding="utf-8"))
            if isinstance(saved, dict):
                cfg.update(saved)
        except Exception:
            pass
    if not cfg.get("verify_token"):
        cfg["verify_token"] = "smart_hub_" + secrets.token_hex(6)
    return cfg


def save_wa_config(cfg_path: Path, current: dict, updates: dict) -> dict:
    merged = dict(current)
    for k in ("default_country_code", "phone_number_id", "verify_token", "template_name", "template_lang", "default_project_id"):
        if k in updates and updates[k] is not None:
            merged[k] = str(updates[k]).strip()
    merged["default_country_code"] = re.sub(r"\D", "", merged.get("default_country_code") or "965") or "965"

    # لا تمسح التوكن المحفوظ إذا أرسل المستخدم الحقل فارغاً أو مقنّعاً (****)
    for secret_k in ("access_token", "app_secret"):
        if secret_k in updates and updates[secret_k] is not None:
            val = str(updates[secret_k]).strip()
            if val == "__CLEAR__":
                merged[secret_k] = ""
            elif val and "..." not in val and "***" not in val:
                merged[secret_k] = val

    if "auto_reply_ack" in updates:
        merged["auto_reply_ack"] = bool(updates["auto_reply_ack"])
    if not merged.get("verify_token"):
        merged["verify_token"] = "smart_hub_" + secrets.token_hex(6)

    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    cfg_path.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        if os.name != "nt":
            os.chmod(cfg_path, 0o600)
    except Exception:
        pass
    return merged


def mask_secret(val: str) -> str:
    s = str(val or "").strip()
    if not s:
        return ""
    if len(s) <= 10:
        return "****"
    return f"{s[:5]}...{s[-4:]}"


def public_wa_config(cfg: dict) -> dict:
    """إعدادات آمنة للعرض في الواجهة دون كشف التوكن السري الكامل."""
    configured = bool(cfg.get("phone_number_id") and cfg.get("access_token"))
    return {
        "configured": configured,
        "default_country_code": cfg.get("default_country_code") or "965",
        "phone_number_id": cfg.get("phone_number_id") or "",
        "access_token_masked": mask_secret(cfg.get("access_token") or ""),
        "has_app_secret": bool(cfg.get("app_secret")),
        "verify_token": cfg.get("verify_token") or "",
        "template_name": cfg.get("template_name") or "",
        "template_lang": cfg.get("template_lang") or "ar",
        "auto_reply_ack": bool(cfg.get("auto_reply_ack", True)),
        "default_project_id": cfg.get("default_project_id") or "",
    }


def normalize_phone(raw: str, default_cc: str = "965") -> str:
    """ينظف رقم الهاتف ويضيف كود الدولة تلقائياً (مثال: 55123456 -> 96555123456)."""
    if not raw:
        return ""
    # تحويل الأرقام العربية/الفارسية إلى إنجليزية
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    s = str(raw).translate(trans).strip()
    had_plus = s.startswith("+")
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    if digits.startswith("00"):
        digits = digits[2:]
        had_plus = True
    cc = re.sub(r"\D", "", default_cc or "965") or "965"
    if had_plus or digits.startswith(cc):
        return digits
    # رقم كويتي محلي من 8 أرقام (يبدأ عادة بـ 2, 4, 5, 6, 9)
    if cc == "965" and len(digits) == 8:
        return cc + digits
    # أرقام خليجية/عربية محلية تبدأ بصفر (05... / 01...)
    if digits.startswith("0") and 9 <= len(digits) <= 11:
        return cc + digits[1:]
    if len(digits) <= 9:
        return cc + digits
    return digits


# =====================================================================
# الوضع (1): WhatsApp Web / Desktop السريع (مجاني 100% + نسخ الملف للحافظة)
# =====================================================================
def build_whatsapp_web_url(phone: str, message: str, default_cc: str = "965") -> str:
    clean_phone = normalize_phone(phone, default_cc)
    params = {}
    if clean_phone:
        params["phone"] = clean_phone
    if message:
        params["text"] = message
    qs = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    return f"https://web.whatsapp.com/send?{qs}"


def copy_file_to_os_clipboard(file_path: Path) -> tuple[bool, str]:
    """ينسخ الملف الفعلي إلى حافظة النظام (Windows / macOS) بحيث يكفي الضغط Ctrl+V داخل شات واتساب!"""
    p = file_path.resolve()
    if not p.exists() or not p.is_file():
        return False, "الملف غير موجود"
    try:
        if os.name == "nt":
            # نمرر المسار عبر متغير بيئة لتجنب أي مشاكل مع الأسماء العربية أو علامات الاقتباس
            env = dict(os.environ, SMART_CLIP_FILE=str(p))
            kw = {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
            res = subprocess.run(
                ["powershell", "-NoProfile", "-STA", "-Command", "Set-Clipboard -LiteralPath $env:SMART_CLIP_FILE"],
                env=env, capture_output=True, text=True, timeout=10, **kw,
            )
            if res.returncode == 0:
                return True, "تم نسخ الملف في الحافظة — اضغط Ctrl+V ثم Enter داخل شات واتساب!"
            return False, (res.stderr or "").strip()[:160]
        elif os.uname().sysname == "Darwin":
            escaped = str(p).replace("\\", "\\\\").replace('"', '\\"')
            res = subprocess.run(
                ["osascript", "-e", f'set the clipboard to (POSIX file "{escaped}")'],
                capture_output=True, text=True, timeout=10,
            )
            if res.returncode == 0:
                return True, "تم نسخ الملف في الحافظة — اضغط Cmd+V ثم Enter داخل شات واتساب!"
            return False, (res.stderr or "").strip()[:160]
    except Exception as e:
        return False, str(e)[:160]
    return False, "افتح شات واتساب وأرفق الملف (أو استخدم Cloud API للإرسال التلقائي)"


# =====================================================================
# الوضع (2): Meta WhatsApp Business Cloud API الرسمي (إرسال + استقبال)
# =====================================================================
def guess_wa_mime(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    if ext in MIME_OVERRIDES:
        return MIME_OVERRIDES[ext]
    mt, _ = mimetypes.guess_type(str(file_path))
    return mt or "application/octet-stream"


def _encode_multipart(fields: dict[str, str], file_field: str, filename: str, file_bytes: bytes, mime_type: str) -> tuple[bytes, str]:
    boundary = f"----SmartHubBoundary{uuid.uuid4().hex}"
    chunks: list[bytes] = []
    for k, v in fields.items():
        chunks.append(f"--{boundary}\r\n".encode("utf-8"))
        chunks.append(f'Content-Disposition: form-data; name="{k}"\r\n\r\n'.encode("utf-8"))
        chunks.append(f"{v}\r\n".encode("utf-8"))
    safe_fn = filename.replace('"', "_").replace("\r", "").replace("\n", "")
    chunks.append(f"--{boundary}\r\n".encode("utf-8"))
    chunks.append(f'Content-Disposition: form-data; name="{file_field}"; filename="{safe_fn}"\r\n'.encode("utf-8"))
    chunks.append(f"Content-Type: {mime_type}\r\n\r\n".encode("utf-8"))
    chunks.append(file_bytes)
    chunks.append(f"\r\n--{boundary}--\r\n".encode("utf-8"))
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def _graph_request(url: str, access_token: str, data: bytes | None = None, content_type: str = "application/json", method: str = "POST") -> dict:
    headers = {"Authorization": f"Bearer {access_token}"}
    if data is not None:
        headers["Content-Type"] = content_type
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            return json.loads(raw) if raw.strip() else {}
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        try:
            j = json.loads(err_body)
            msg = j.get("error", {}).get("message") or err_body[:250]
            code = j.get("error", {}).get("code")
            raise RuntimeError(f"Meta API ({code or e.code}): {msg}")
        except json.JSONDecodeError:
            raise RuntimeError(f"Meta API HTTP {e.code}: {err_body[:220]}")


def cloud_upload_media(file_path: Path, phone_number_id: str, access_token: str, api_base: str = GRAPH_API_BASE) -> tuple[str, str]:
    """يرفع الملف (PDF / Word / Excel / صورة) إلى خادم Meta WhatsApp ويعيد media_id."""
    mime = guess_wa_mime(file_path)
    file_bytes = file_path.read_bytes()
    if len(file_bytes) > 100 * 1024 * 1024:
        raise ValueError("حجم الملف يتجاوز حد واتساب الرسمي (100 MB)")
    body, ctype = _encode_multipart(
        {"messaging_product": "whatsapp", "type": mime},
        "file", file_path.name, file_bytes, mime,
    )
    res = _graph_request(f"{api_base}/{phone_number_id}/media", access_token, data=body, content_type=ctype, method="POST")
    media_id = str(res.get("id") or "")
    if not media_id:
        raise RuntimeError(f"لم يُرجع Meta معرّف الملف: {res}")
    return media_id, mime


def cloud_send_file(
    file_path: Path,
    to_phone: str,
    caption: str,
    wa_cfg: dict,
    use_template: bool = False,
    api_base: str = GRAPH_API_BASE,
) -> dict:
    """يرسل ملفاً رسمياً مع رسالة عبر WhatsApp Business Cloud API."""
    phone_id = str(wa_cfg.get("phone_number_id") or "").strip()
    token = str(wa_cfg.get("access_token") or "").strip()
    if not phone_id or not token:
        raise ValueError("يرجى إدخال Phone Number ID و Access Token في إعدادات واتساب أولاً (أو استخدم زر WhatsApp Web المجاني).")

    clean_to = normalize_phone(to_phone, wa_cfg.get("default_country_code") or "965")
    if not clean_to or len(clean_to) < 8:
        raise ValueError("يرجى كتابة رقم واتساب صحيح (مثال: 55123456 أو +96555123456).")

    media_id, mime = cloud_upload_media(file_path, phone_id, token, api_base=api_base)
    tpl_name = str(wa_cfg.get("template_name") or "").strip()

    if use_template and tpl_name:
        payload = {
            "messaging_product": "whatsapp",
            "to": clean_to,
            "type": "template",
            "template": {
                "name": tpl_name,
                "language": {"code": wa_cfg.get("template_lang") or "ar"},
                "components": [
                    {"type": "header", "parameters": [{"type": "document", "document": {"id": media_id, "filename": file_path.name}}]},
                    {"type": "body", "parameters": [{"type": "text", "text": (caption or file_path.name)[:900]}]},
                ],
            },
        }
    elif mime in {"image/jpeg", "image/png"} and file_path.stat().st_size <= 5 * 1024 * 1024:
        payload = {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": clean_to,
            "type": "image",
            "image": {"id": media_id, "caption": (caption or "")[:1024]},
        }
    else:
        payload = {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": clean_to,
            "type": "document",
            "document": {"id": media_id, "filename": file_path.name, "caption": (caption or "")[:1024]},
        }

    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    res = _graph_request(f"{api_base}/{phone_id}/messages", token, data=data, content_type="application/json", method="POST")
    msg_id = ""
    if isinstance(res.get("messages"), list) and res["messages"]:
        msg_id = str(res["messages"][0].get("id") or "")
    return {"ok": True, "to": clean_to, "media_id": media_id, "message_id": msg_id, "file_name": file_path.name}


def cloud_send_text(to_phone: str, text: str, wa_cfg: dict, api_base: str = GRAPH_API_BASE) -> dict:
    """يرسل رسالة نصية سريعة (مثل إشعار استلام الملف المجاني داخل نافذة 24 ساعة)."""
    phone_id = str(wa_cfg.get("phone_number_id") or "").strip()
    token = str(wa_cfg.get("access_token") or "").strip()
    clean_to = normalize_phone(to_phone, wa_cfg.get("default_country_code") or "965")
    if not (phone_id and token and clean_to and text):
        return {"ok": False}
    payload = {
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": clean_to,
        "type": "text",
        "text": {"body": text[:4000]},
    }
    res = _graph_request(
        f"{api_base}/{phone_id}/messages", token,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        content_type="application/json", method="POST",
    )
    msg_id = res["messages"][0].get("id", "") if res.get("messages") else ""
    return {"ok": True, "to": clean_to, "message_id": msg_id}


def verify_webhook_signature(raw_body: bytes, sig_header: str, app_secret: str) -> bool:
    """يتحقق من توقيع X-Hub-Signature-256 القادم من Meta إن كان App Secret مضبوطاً."""
    secret = (app_secret or "").strip()
    if not secret:
        return True
    if not sig_header or not sig_header.startswith("sha256="):
        return False
    expected = "sha256=" + hmac.new(secret.encode("utf-8"), raw_body, hashlib.sha256).hexdigest()
    return hmac.compare_digest(expected, sig_header.strip())


def cloud_download_media(media_id: str, access_token: str, dest_dir: Path, orig_filename: str = "", mime_hint: str = "", api_base: str = GRAPH_API_BASE) -> Path:
    """ينزّل صورة أو مستند وارد من واتساب Cloud API (مجاني 100%) ويحفظه في مجلد استلام المشروع."""
    info = _graph_request(f"{api_base}/{media_id}", access_token, data=None, method="GET")
    dl_url = str(info.get("url") or "")
    mime = str(info.get("mime_type") or mime_hint or "").split(";")[0].strip().lower()
    if not dl_url:
        raise RuntimeError(f"تعذر جلب رابط تنزيل ملف واتساب ({media_id})")

    req = urllib.request.Request(dl_url, headers={"Authorization": f"Bearer {access_token}"}, method="GET")
    with urllib.request.urlopen(req, timeout=120) as resp:
        content = resp.read()

    dest_dir.mkdir(parents=True, exist_ok=True)
    ext = Path(orig_filename).suffix.lower() if orig_filename else ""
    if not ext:
        ext = MIME_TO_EXT.get(mime, ".bin")
    try:
        from security_utils import DANGEROUS_EXTENSIONS
        if ext in DANGEROUS_EXTENSIONS:
            raise ValueError(f"Blocked dangerous file extension: {ext}")
    except ImportError:
        pass
    stem = Path(orig_filename).stem if orig_filename else f"WhatsApp_{media_id[-8:]}"
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem).strip(" ._")[:65] or "WhatsApp_File"
    out_p = dest_dir / f"{stem}{ext}"
    if out_p.exists():
        out_p = dest_dir / f"{stem}_{uuid.uuid4().hex[:5]}{ext}"
    out_p.write_bytes(content)
    return out_p


def extract_incoming_media_events(webhook_payload: dict) -> list[dict]:
    """يستخرج كافة الصور والملفات (PDF / Word / Excel / صور) الواردة في إشعار Webhook من Meta."""
    events: list[dict] = []
    for entry in webhook_payload.get("entry") or []:
        for change in entry.get("changes") or []:
            val = change.get("value") or {}
            contacts_map = {}
            for c in val.get("contacts") or []:
                wa_id = str(c.get("wa_id") or "")
                pname = str((c.get("profile") or {}).get("name") or "")
                if wa_id:
                    contacts_map[wa_id] = pname
            for msg in val.get("messages") or []:
                mtype = str(msg.get("type") or "")
                sender = str(msg.get("from") or "")
                sender_name = contacts_map.get(sender, "")
                if mtype in ("document", "image"):
                    obj = msg.get(mtype) or {}
                    mid = str(obj.get("id") or "")
                    if not mid:
                        continue
                    events.append({
                        "message_id": str(msg.get("id") or ""),
                        "from": sender,
                        "sender_name": sender_name,
                        "media_type": mtype,
                        "media_id": mid,
                        "mime_type": str(obj.get("mime_type") or ""),
                        "filename": str(obj.get("filename") or ""),
                        "caption": str(obj.get("caption") or ""),
                    })
    return events
