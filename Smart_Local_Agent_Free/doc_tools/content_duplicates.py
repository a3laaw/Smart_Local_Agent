#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
content_duplicates.py — Content-based Duplicate & Visual Similarity Reviewer.

Detects duplicate and near-duplicate images and documents purely by their
ACTUAL CONTENT (ignoring filenames completely):
  1. Exact content match (SHA-256) across all file types (images, documents, video, audio, etc.)
  2. Visual content similarity for images (multi-signal dHash horizontal + vertical + aHash + 4x4 RGB
     color signature with EXIF orientation normalization; catches resized/recompressed WhatsApp photos,
     format conversions JPG<->PNG<->WEBP, and near-identical shots)
  3. Document text-content similarity (PDF via fitz/stream fallback, DOCX, XLSX/XLSM, PPTX, TXT/CSV/MD/JSON)
     using normalized Arabic/English word shingles (SimHash + Jaccard verification)
  4. Optional visual rendering of scanned PDFs via PyMuPDF (`fitz`) when available

Supports local folder scans AND browser file/folder/ZIP uploads, side-by-side group review,
custom "set as original", safe quarantine with full undo, unique-only ZIP export, permanent deletion
with double confirmation, RTL Excel reporting, and automatic screen clearing after execution.
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
import zipfile
import zlib
from collections import defaultdict
from pathlib import Path
from typing import Any, Callable
try:
    from defusedxml import ElementTree as ET
except ImportError:
    from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from i18n_ar import tr
from relevance import normalize_ar
from security_utils import (
    atomic_write_json,
    read_json_safe,
    resolve_local_or_unc,
    safe_resolve_within,
    sha256_of_file,
)

try:
    from PIL import Image, ImageOps
except ImportError:  # pragma: no cover
    Image = None  # type: ignore
    ImageOps = None  # type: ignore

try:
    import docx as _python_docx  # type: ignore
except ImportError:
    _python_docx = None  # type: ignore

try:
    import pymupdf as _fitz  # type: ignore
except ImportError:
    try:
        import fitz as _fitz  # type: ignore
    except ImportError:
        _fitz = None  # type: ignore

IMAGE_EXTS = {
    ".jpg", ".jpeg", ".jfif", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff", ".heic",
}
DOC_EXTS = {
    ".pdf", ".docx", ".doc", ".xlsx", ".xlsm", ".xls", ".pptx", ".ppt",
    ".txt", ".csv", ".tsv", ".md", ".json", ".rtf", ".log",
}
VIDEO_EXTS = {".mp4", ".mov", ".avi", ".mkv", ".3gp", ".webm", ".m4v"}
AUDIO_EXTS = {".opus", ".ogg", ".mp3", ".m4a", ".wav", ".aac"}

SKIP_DIR_NAMES = {
    "_system_agent", "_duplicates_quarantine", "_organized_archive",
    "_backups", ".git", "__pycache__", "node_modules", ".venv",
}

MAX_DUPS_ZIP_MEMBERS = 5000
MAX_DUPS_ZIP_BYTES = 2 * 1024 * 1024 * 1024  # 2 GB
_WIN_BAD_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WORD_RE = re.compile(r"[\w\u0600-\u06FF]{2,}", re.UNICODE)


def _sys_dir(proj_root: Path) -> Path:
    d = proj_root / "_System_Agent"
    d.mkdir(parents=True, exist_ok=True)
    return d


def dups_plan_path(proj_root: Path) -> Path:
    return _sys_dir(proj_root) / "duplicates_plan.json"


def dups_undo_path(proj_root: Path) -> Path:
    return _sys_dir(proj_root) / "duplicates_undo.json"


def dups_uploads_root(proj_root: Path) -> Path:
    return proj_root / "00_Duplicates_Review"


def dups_quarantine_root(proj_root: Path) -> Path:
    return proj_root / "_Duplicates_Quarantine"


def _classify_media_kind(ext: str) -> str:
    el = ext.lower()
    if el in IMAGE_EXTS:
        return "image"
    if el in DOC_EXTS:
        return "document"
    if el in VIDEO_EXTS:
        return "video"
    if el in AUDIO_EXTS:
        return "audio"
    return "other"


def _fmt_bytes(n: int) -> str:
    v = float(max(0, int(n)))
    if v < 1024:
        return f"{int(v)} B"
    if v < 1024 * 1024:
        return f"{v / 1024:.1f} KB"
    if v < 1024 * 1024 * 1024:
        return f"{v / (1024 * 1024):.2f} MB"
    return f"{v / (1024 * 1024 * 1024):.2f} GB"


def _safe_extract_zip(zip_path: Path, dest_dir: Path) -> int:
    dest_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    total_bytes = 0
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            raw_name = info.filename.replace("\\", "/")
            if "\x00" in raw_name or raw_name.startswith("/"):
                raise ValueError(tr("dups.zip_unsafe_err"))
            segs = [s for s in raw_name.split("/") if s not in ("", ".")]
            if not segs or any(s == ".." for s in segs):
                raise ValueError(tr("dups.zip_unsafe_err"))
            # Block symlinks inside zip (external_attr Unix symlink bit)
            unix_mode = (info.external_attr >> 16) & 0o170000
            if unix_mode == 0o120000:
                raise ValueError(tr("dups.zip_unsafe_err"))
            count += 1
            total_bytes += max(0, int(info.file_size))
            if count > MAX_DUPS_ZIP_MEMBERS or total_bytes > MAX_DUPS_ZIP_BYTES:
                raise ValueError(tr("dups.zip_too_large_err"))
            clean_segs = [_WIN_BAD_CHARS_RE.sub("_", s).strip() or "file" for s in segs]
            target = safe_resolve_within(dest_dir, "/".join(clean_segs), allow_base=False, reject_symlinks=True)
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info, "r") as src_f, open(target, "wb") as dst_f:
                shutil.copyfileobj(src_f, dst_f, length=65536)
    return count


def _image_signature_from_pil(img_raw: Any) -> dict[str, Any] | None:
    if Image is None:
        return None
    try:
        img = ImageOps.exif_transpose(img_raw) if ImageOps is not None else img_raw
        w, h = img.size
        if w <= 0 or h <= 0:
            return None
        resamp = getattr(getattr(Image, "Resampling", Image), "BILINEAR", Image.BILINEAR)
        rgb = img.convert("RGB")

        # 1) 4x4 RGB spatial color matrix (48 ints)
        small_rgb = rgb.resize((4, 4), resamp)
        rgb_bytes = list(small_rgb.tobytes())  # 48 values 0..255
        mean_r = sum(rgb_bytes[0::3]) // 16
        mean_g = sum(rgb_bytes[1::3]) // 16
        mean_b = sum(rgb_bytes[2::3]) // 16

        # 2) Horizontal dHash (9x8 -> 64 bits)
        gray = rgb.convert("L")
        g_h = gray.resize((9, 8), resamp).tobytes()
        dhash_h = 0
        bit_idx = 0
        for row in range(8):
            off = row * 9
            for col in range(8):
                if g_h[off + col] > g_h[off + col + 1]:
                    dhash_h |= (1 << bit_idx)
                bit_idx += 1

        # 3) Vertical dHash (8x9 -> 64 bits)
        g_v = gray.resize((8, 9), resamp).tobytes()
        dhash_v = 0
        bit_idx = 0
        for row in range(8):
            off = row * 8
            for col in range(8):
                if g_v[off + col] > g_v[off + 8 + col]:
                    dhash_v |= (1 << bit_idx)
                bit_idx += 1

        # 4) Average Hash aHash (8x8 -> 64 bits) + contrast range
        g_a = gray.resize((8, 8), resamp).tobytes()
        avg_g = sum(g_a) / 64.0
        ahash = 0
        for idx, px in enumerate(g_a):
            if px >= avg_g:
                ahash |= (1 << idx)
        contrast = (max(g_a) - min(g_a)) if g_a else 0

        return {
            "width": int(w),
            "height": int(h),
            "aspect": float(w) / float(max(1, h)),
            "dhash_h": dhash_h,
            "dhash_v": dhash_v,
            "ahash": ahash,
            "rgb4": rgb_bytes,
            "mean_rgb": (mean_r, mean_g, mean_b),
            "mean_luma": int(round(avg_g)),
            "is_flat": contrast < 5,
        }
    except Exception:
        return None


def _extract_image_signature(fp: Path) -> dict[str, Any] | None:
    if Image is None:
        return None
    try:
        with Image.open(fp) as im:
            return _image_signature_from_pil(im)
    except Exception:
        return None


def _compute_image_similarity(sig_a: dict[str, Any], sig_b: dict[str, Any]) -> int:
    ar_a = sig_a["aspect"]
    ar_b = sig_b["aspect"]
    if max(ar_a / max(1e-6, ar_b), ar_b / max(1e-6, ar_a)) > 1.45:
        return 0

    col_diff = sum(abs(x - y) for x, y in zip(sig_a["rgb4"], sig_b["rgb4"]))
    col_sim = 1.0 - (col_diff / (48.0 * 255.0))

    if sig_a["is_flat"] and sig_b["is_flat"]:
        return int(round(col_sim * 100))
    if sig_a["is_flat"] != sig_b["is_flat"]:
        return int(round(min(0.65, col_sim) * 100))

    if col_sim < 0.70:
        return 0

    dh_h = 1.0 - ((sig_a["dhash_h"] ^ sig_b["dhash_h"]).bit_count() / 64.0)
    dh_v = 1.0 - ((sig_a["dhash_v"] ^ sig_b["dhash_v"]).bit_count() / 64.0)
    ah = 1.0 - ((sig_a["ahash"] ^ sig_b["ahash"]).bit_count() / 64.0)

    if max(dh_h, dh_v) < 0.62:
        return 0

    score = 0.40 * dh_h + 0.25 * dh_v + 0.20 * ah + 0.15 * col_sim
    return max(0, min(100, int(round(score * 100))))


def _extract_pdf_text_and_sig(fp: Path) -> tuple[str, dict[str, Any] | None]:
    txt_parts: list[str] = []
    vis_sig: dict[str, Any] | None = None
    if _fitz is not None:
        try:
            with _fitz.open(fp) as doc:
                for idx, page in enumerate(doc):
                    if idx >= 6:
                        break
                    t = page.get_text("text") or ""
                    if t.strip():
                        txt_parts.append(t)
                full_t = "\n".join(txt_parts).strip()
                # If scanned PDF (very little text), render page 1 for visual similarity
                if len(full_t) < 40 and len(doc) >= 1 and Image is not None:
                    try:
                        pix = doc[0].get_pixmap(matrix=_fitz.Matrix(1, 1), alpha=False)
                        pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        vis_sig = _image_signature_from_pil(pil_img)
                    except Exception:
                        pass
                return full_t[:80000], vis_sig
        except Exception:
            pass

    # Lightweight pure-Python PDF text stream fallback when PyMuPDF is absent
    try:
        with fp.open("rb") as _fh:
            raw = _fh.read(2 * 1024 * 1024)
        chunks: list[str] = []
        for m in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", raw, re.DOTALL):
            data = m.group(1)
            try:
                dec = zlib.decompress(data)
            except Exception:
                dec = data
            for tm in re.finditer(rb"\(([^()]{2,200})\)\s*Tj", dec):
                try:
                    chunks.append(tm.group(1).decode("utf-8", errors="ignore"))
                except Exception:
                    pass
        return " ".join(chunks)[:80000], None
    except Exception:
        return "", None


def _extract_document_text_and_sig(fp: Path) -> tuple[str, dict[str, Any] | None]:
    ext = fp.suffix.lower()
    if ext in {".txt", ".csv", ".tsv", ".md", ".json", ".rtf", ".log"}:
        try:
            with fp.open("rb") as _fh:
                raw = _fh.read(120 * 1024)
            for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
                try:
                    return raw.decode(enc)[:80000], None
                except UnicodeDecodeError:
                    continue
        except Exception:
            return "", None

    if ext == ".docx":
        if _python_docx is not None:
            try:
                doc = _python_docx.Document(str(fp))
                parts = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
                for tbl in doc.tables[:8]:
                    for row in tbl.rows[:30]:
                        row_t = " ".join(c.text.strip() for c in row.cells if c.text and c.text.strip())
                        if row_t:
                            parts.append(row_t)
                return "\n".join(parts)[:80000], None
            except Exception:
                pass
        try:
            with zipfile.ZipFile(fp, "r") as zf:
                xml_bytes = zf.read("word/document.xml")
            root = ET.fromstring(xml_bytes)
            texts = [el.text for el in root.iter() if el.text and el.tag.endswith("}t")]
            return " ".join(texts)[:80000], None
        except Exception:
            return "", None

    if ext in {".xlsx", ".xlsm"}:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            parts: list[str] = []
            try:
                for s_name in wb.sheetnames[:4]:
                    ws = wb[s_name]
                    for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
                        if r_idx >= 250:
                            break
                        vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
                        if vals:
                            parts.append(" ".join(vals))
            finally:
                wb.close()
            return "\n".join(parts)[:80000], None
        except Exception:
            return "", None

    if ext == ".pptx":
        try:
            parts = []
            with zipfile.ZipFile(fp, "r") as zf:
                slide_names = sorted(
                    n for n in zf.namelist()
                    if n.startswith("ppt/slides/slide") and n.endswith(".xml")
                )[:30]
                for sn in slide_names:
                    root = ET.fromstring(zf.read(sn))
                    texts = [el.text for el in root.iter() if el.text and el.tag.endswith("}t")]
                    if texts:
                        parts.append(" ".join(texts))
            return "\n".join(parts)[:80000], None
        except Exception:
            return "", None

    if ext == ".pdf":
        return _extract_pdf_text_and_sig(fp)

    return "", None


def _build_text_shingles(raw_text: str) -> tuple[set[str], int, str]:
    norm = normalize_ar(raw_text)
    words = _WORD_RE.findall(norm)
    if len(words) < 6 or len(norm) < 35:
        return set(), 0, ""
    n = 3 if len(words) >= 8 else 2
    shingles = {" ".join(words[i : i + n]) for i in range(len(words) - n + 1)}
    if not shingles:
        return set(), 0, ""
    # 64-bit SimHash for fast bucketing
    acc = [0] * 64
    for sh in shingles:
        digest = int.from_bytes(hashlib.blake2b(sh.encode("utf-8"), digest_size=8).digest(), "little")
        for b in range(64):
            acc[b] += 1 if (digest & (1 << b)) else -1
    simhash = 0
    for b in range(64):
        if acc[b] > 0:
            simhash |= (1 << b)
    snippet = " ".join(words[:18])[:140]
    return shingles, simhash, snippet


def _compute_text_similarity(sh_a: set[str], sh_b: set[str]) -> int:
    if not sh_a or not sh_b:
        return 0
    inter = len(sh_a & sh_b)
    if inter == 0:
        return 0
    union = len(sh_a | sh_b)
    jacc = inter / max(1, union)
    overlap = inter / max(1, min(len(sh_a), len(sh_b)))
    score = 0.70 * jacc + 0.30 * overlap
    return max(0, min(100, int(round(score * 100))))


class _UnionFind:
    def __init__(self, n: int) -> None:
        self.parent = list(range(n))

    def find(self, x: int) -> int:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            if ra < rb:
                self.parent[rb] = ra
            else:
                self.parent[ra] = rb


def _original_sort_key(rec: dict[str, Any]) -> tuple[Any, ...]:
    pixels = int(rec.get("width") or 0) * int(rec.get("height") or 0)
    tlen = int(rec.get("text_len") or 0)
    size_b = int(rec.get("size_bytes") or 0)
    mtime_ts = float(rec.get("mtime_ts") or 0.0)
    rel_len = len(str(rec.get("rel_path") or ""))
    # Highest resolution / text richness / largest file size / earliest mtime / shortest path
    return (-pixels, -tlen, -size_b, mtime_ts, rel_len, str(rec.get("rel_path") or ""))


class DuplicatesStoppedError(Exception):
    """Raised cleanly when the user stops a running duplicates scan/execution job."""


def _emit_dup_prog(
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


def scan_content_duplicates(
    proj: dict[str, Any],
    *,
    source_kind: str = "folder",
    source_path: str = "",
    batch_id: str = "",
    scope: str = "all",
    min_similarity: int = 85,
    include_subfolders: bool = True,
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """Scan a folder or uploaded batch for exact and content-similar duplicates (ignoring names)."""
    proj_root = resolve_local_or_unc(str(proj["root"]))
    last_emit = [0.0]

    def _chk_stop() -> None:
        if stop_event is not None and stop_event.is_set():
            raise DuplicatesStoppedError(tr("jobs.stopped_by_user"))

    _emit_dup_prog(progress_cb, last_emit, {
        "phase": "walking", "processed": 0, "total": 0, "discovered": 0, "current_file": "",
    }, force=True)

    s_kind = str(source_kind or "folder").strip().lower()
    if s_kind not in ("folder", "upload", "whatsapp_current", "project_archive"):
        s_kind = "folder"

    clean_batch = re.sub(r"[^A-Za-z0-9_-]", "", str(batch_id or ""))[:40]
    if s_kind == "upload":
        if not clean_batch:
            raise ValueError(tr("dups.no_files_selected"))
        base_dir = safe_resolve_within(dups_uploads_root(proj_root), clean_batch, allow_base=False, reject_symlinks=True)
        if not base_dir.exists():
            raise ValueError(tr("dups.no_files_selected"))
        # Unpack any uploaded ZIPs safely in-place under _unpacked_<name>
        for zp in sorted(base_dir.rglob("*.zip")):
            _chk_stop()
            if zp.is_file():
                _emit_dup_prog(progress_cb, last_emit, {
                    "phase": "extracting", "processed": 0, "total": 0, "discovered": 0, "current_file": zp.name,
                }, force=True)
                unpack_dir = zp.parent / f"_zip_{zp.stem[:24]}"
                if not unpack_dir.exists():
                    _safe_extract_zip(zp, unpack_dir)
                try:
                    zp.unlink()
                except Exception:
                    pass
    elif s_kind == "whatsapp_current":
        wa_plan = read_json_safe(proj_root / "_System_Agent" / "whatsapp_archive_plan.json", {}) or {}
        wa_src = str(wa_plan.get("source_dir") or "").strip()
        if wa_src:
            base_dir = resolve_local_or_unc(wa_src)
        else:
            base_dir = proj_root / "00_WhatsApp_Chats"
        if not base_dir.exists():
            raise ValueError(tr("dups.source_missing_err"))
    elif s_kind == "project_archive":
        arc_raw = str(proj.get("archive_dir") or "").strip()
        base_dir = resolve_local_or_unc(arc_raw) if arc_raw else (proj_root / "02_Project_Archive")
        if not base_dir.exists():
            raise ValueError(tr("dups.source_missing_err"))
    else:
        raw_p = str(source_path or "").strip()
        if not raw_p:
            raise ValueError(tr("dups.source_missing_err"))
        base_dir = resolve_local_or_unc(raw_p)
        if not base_dir.exists() or not base_dir.is_dir():
            raise ValueError(tr("dups.source_missing_err"))

    sc_mode = str(scope or "all").strip().lower()
    if sc_mode not in ("all", "images", "documents"):
        sc_mode = "all"
    thresh = max(65, min(100, int(min_similarity if min_similarity is not None else 85)))

    # 1) Walk files safely
    records: list[dict[str, Any]] = []
    if include_subfolders:
        for cur_root, dirs, files in os.walk(base_dir):
            _chk_stop()
            dirs[:] = sorted(
                d for d in dirs
                if d.lower() not in SKIP_DIR_NAMES and not d.startswith(".")
            )
            for fn in sorted(files):
                _chk_stop()
                if fn.startswith(".") or fn.startswith("~$"):
                    continue
                fp = Path(cur_root) / fn
                if fp.is_symlink() or not fp.is_file():
                    continue
                mk = _classify_media_kind(fp.suffix)
                if sc_mode == "images" and mk != "image":
                    continue
                if sc_mode == "documents" and mk != "document":
                    continue
                st = fp.stat()
                if st.st_size == 0:
                    continue
                rel_p = fp.relative_to(base_dir).as_posix()
                records.append({
                    "id": f"d_{len(records):05d}_{uuid.uuid4().hex[:6]}",
                    "file_name": fp.name,
                    "rel_path": rel_p,
                    "abs_path": str(fp),
                    "ext": fp.suffix.lower(),
                    "media_kind": mk,
                    "size_bytes": int(st.st_size),
                    "size_label": _fmt_bytes(st.st_size),
                    "mtime": dt.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
                    "mtime_ts": float(st.st_mtime),
                    "width": 0,
                    "height": 0,
                    "sha256": "",
                    "text_len": 0,
                    "snippet": "",
                })
                if len(records) == 1 or len(records) % 15 == 0:
                    _emit_dup_prog(progress_cb, last_emit, {
                        "phase": "walking", "processed": 0, "total": 0,
                        "discovered": len(records), "current_file": fp.name,
                    })
    else:
        for fp in sorted(base_dir.iterdir()):
            _chk_stop()
            if fp.name.startswith(".") or fp.name.startswith("~$") or fp.is_symlink() or not fp.is_file():
                continue
            mk = _classify_media_kind(fp.suffix)
            if sc_mode == "images" and mk != "image":
                continue
            if sc_mode == "documents" and mk != "document":
                continue
            st = fp.stat()
            if st.st_size == 0:
                continue
            records.append({
                "id": f"d_{len(records):05d}_{uuid.uuid4().hex[:6]}",
                "file_name": fp.name,
                "rel_path": fp.name,
                "abs_path": str(fp),
                "ext": fp.suffix.lower(),
                "media_kind": mk,
                "size_bytes": int(st.st_size),
                "size_label": _fmt_bytes(st.st_size),
                "mtime": dt.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
                "mtime_ts": float(st.st_mtime),
                "width": 0,
                "height": 0,
                "sha256": "",
                "text_len": 0,
                "snippet": "",
            })
            if len(records) == 1 or len(records) % 15 == 0:
                _emit_dup_prog(progress_cb, last_emit, {
                    "phase": "walking", "processed": 0, "total": 0,
                    "discovered": len(records), "current_file": fp.name,
                })

    n_total = len(records)
    _emit_dup_prog(progress_cb, last_emit, {
        "phase": "hashing", "processed": 0, "total": max(1, n_total),
        "discovered": n_total, "current_file": "",
    }, force=True)
    uf = _UnionFind(n_total)
    edge_info: dict[tuple[int, int], tuple[str, int, str]] = {}  # (u, v) -> (kind, pct, detail)

    def _add_edge(i: int, j: int, kind: str, pct: int, detail: str = "") -> None:
        u, v = (i, j) if i < j else (j, i)
        prev = edge_info.get((u, v))
        if prev is None or pct > prev[1]:
            edge_info[(u, v)] = (kind, pct, detail)
        uf.union(u, v)

    # 2) Stage A: Exact byte-level content match (SHA-256) on same-size files
    by_size: dict[int, list[int]] = defaultdict(list)
    for idx, rec in enumerate(records):
        by_size[rec["size_bytes"]].append(idx)

    by_sha: dict[str, list[int]] = defaultdict(list)
    hash_candidates = sum(len(idxs) for idxs in by_size.values() if len(idxs) >= 2)
    hashed_done = 0
    for _sz, idxs in by_size.items():
        if len(idxs) < 2:
            continue
        for idx in idxs:
            _chk_stop()
            try:
                fh = sha256_of_file(Path(records[idx]["abs_path"]))
            except Exception:
                fh = ""
            records[idx]["sha256"] = fh
            if fh:
                by_sha[fh].append(idx)
            hashed_done += 1
            _emit_dup_prog(progress_cb, last_emit, {
                "phase": "hashing",
                "processed": hashed_done if thresh >= 100 else 0,
                "total": max(1, hash_candidates if thresh >= 100 else n_total),
                "discovered": n_total,
                "current_file": records[idx]["file_name"],
            })

    for fh, idxs in by_sha.items():
        if len(idxs) >= 2:
            anchor = idxs[0]
            for other in idxs[1:]:
                _add_edge(anchor, other, "exact_content", 100, "")

    # 3) Stage B & C: Extract visual signatures (images + scanned PDFs) & text shingles (docs)
    vis_sigs: dict[int, dict[str, Any]] = {}
    doc_shingles: dict[int, tuple[set[str], int, str]] = {}

    if thresh < 100:
        _emit_dup_prog(progress_cb, last_emit, {
            "phase": "analyzing", "processed": 0, "total": max(1, n_total),
            "discovered": n_total, "current_file": records[0]["file_name"] if records else "",
        }, force=True)
        for idx, rec in enumerate(records):
            _chk_stop()
            fp = Path(rec["abs_path"])
            mk = rec["media_kind"]
            _emit_dup_prog(progress_cb, last_emit, {
                "phase": "analyzing", "processed": idx + 1, "total": max(1, n_total),
                "discovered": n_total, "current_file": rec["file_name"],
            })
            if mk == "image" and sc_mode in ("all", "images"):
                sig = _extract_image_signature(fp)
                if sig:
                    rec["width"] = sig["width"]
                    rec["height"] = sig["height"]
                    vis_sigs[idx] = sig
            elif mk == "document" and sc_mode in ("all", "documents"):
                raw_t, scanned_sig = _extract_document_text_and_sig(fp)
                if raw_t:
                    sh, sh_hash, snip = _build_text_shingles(raw_t)
                    if sh:
                        rec["text_len"] = len(raw_t)
                        rec["snippet"] = snip
                        doc_shingles[idx] = (sh, sh_hash, snip)
                if scanned_sig is not None and sc_mode == "all":
                    rec["width"] = scanned_sig["width"]
                    rec["height"] = scanned_sig["height"]
                    vis_sigs[idx] = scanned_sig

    # Visual similarity pairwise with smart luma/color bucketing for large batches
    vis_indices = sorted(vis_sigs.keys())
    if len(vis_indices) >= 2 and thresh < 100:
        _emit_dup_prog(progress_cb, last_emit, {
            "phase": "comparing", "processed": 0, "total": len(vis_indices),
            "discovered": n_total, "current_file": records[vis_indices[0]]["file_name"],
        }, force=True)
        # Group by 16-bit dhash bands when large; direct compare up to 800
        if len(vis_indices) <= 800:
            for pos_a in range(len(vis_indices)):
                _chk_stop()
                ia = vis_indices[pos_a]
                sa = vis_sigs[ia]
                _emit_dup_prog(progress_cb, last_emit, {
                    "phase": "comparing", "processed": pos_a + 1, "total": len(vis_indices),
                    "discovered": n_total, "current_file": records[ia]["file_name"],
                })
                for pos_b in range(pos_a + 1, len(vis_indices)):
                    ib = vis_indices[pos_b]
                    if records[ia]["sha256"] and records[ia]["sha256"] == records[ib]["sha256"]:
                        continue
                    sb = vis_sigs[ib]
                    # Quick pre-filter on mean RGB channel difference (<~ 77 per channel avg)
                    ma, mb = sa["mean_rgb"], sb["mean_rgb"]
                    if (abs(ma[0] - mb[0]) + abs(ma[1] - mb[1]) + abs(ma[2] - mb[2])) > 235:
                        continue
                    pct = _compute_image_similarity(sa, sb)
                    if pct >= thresh:
                        _add_edge(ia, ib, "visual_similar", pct, "")
        else:
            # Multi-index 16-bit bands on dhash_h + dhash_v + luma bucket to avoid O(N^2) on thousands of images
            bands: dict[tuple[int, int], list[int]] = defaultdict(list)
            for idx in vis_indices:
                dh = vis_sigs[idx]["dhash_h"]
                dv = vis_sigs[idx]["dhash_v"]
                for b in range(4):
                    bands[(b, (dh >> (b * 16)) & 0xFFFF)].append(idx)
                    bands[(b + 4, (dv >> (b * 16)) & 0xFFFF)].append(idx)
                lu = vis_sigs[idx]["mean_luma"] // 24
                bands[(99, lu)].append(idx)
            seen_pairs: set[tuple[int, int]] = set()
            total_buckets = max(1, len(bands))
            for b_pos, bucket in enumerate(bands.values()):
                _chk_stop()
                _emit_dup_prog(progress_cb, last_emit, {
                    "phase": "comparing", "processed": b_pos + 1, "total": total_buckets,
                    "discovered": n_total, "current_file": "",
                })
                if len(bucket) < 2 or len(bucket) > 1500:
                    continue
                for pa in range(len(bucket)):
                    ia = bucket[pa]
                    sa = vis_sigs[ia]
                    for pb in range(pa + 1, len(bucket)):
                        ib = bucket[pb]
                        pair = (ia, ib) if ia < ib else (ib, ia)
                        if pair in seen_pairs:
                            continue
                        seen_pairs.add(pair)
                        if records[ia]["sha256"] and records[ia]["sha256"] == records[ib]["sha256"]:
                            continue
                        pct = _compute_image_similarity(sa, vis_sigs[ib])
                        if pct >= thresh:
                            _add_edge(ia, ib, "visual_similar", pct, "")

    # Document text similarity pairwise
    doc_indices = sorted(doc_shingles.keys())
    if len(doc_indices) >= 2 and thresh < 100:
        for pa in range(len(doc_indices)):
            _chk_stop()
            ia = doc_indices[pa]
            sh_a, sim_a, snip_a = doc_shingles[ia]
            _emit_dup_prog(progress_cb, last_emit, {
                "phase": "comparing", "processed": pa + 1, "total": len(doc_indices),
                "discovered": n_total, "current_file": records[ia]["file_name"],
            })
            for pb in range(pa + 1, len(doc_indices)):
                ib = doc_indices[pb]
                if records[ia]["sha256"] and records[ia]["sha256"] == records[ib]["sha256"]:
                    continue
                sh_b, sim_b, _ = doc_shingles[ib]
                # Fast length & SimHash filter before exact Jaccard
                if len(doc_indices) > 250 and ((sim_a ^ sim_b).bit_count() > 24):
                    continue
                pct = _compute_text_similarity(sh_a, sh_b)
                if pct >= thresh:
                    _add_edge(ia, ib, "doc_content_similar", pct, snip_a)

    # 4) Assemble connected components into groups
    comp_map: dict[int, list[int]] = defaultdict(list)
    for idx in range(n_total):
        comp_map[uf.find(idx)].append(idx)

    groups: list[dict[str, Any]] = []
    flat_items: list[dict[str, Any]] = []

    grp_seq = 0
    for _root_id, members in comp_map.items():
        if len(members) < 2:
            continue
        grp_seq += 1
        gid = f"grp_{grp_seq:04d}"

        # Pick best representative (original) by resolution / richness / size / earliest mtime
        members_sorted = sorted(members, key=lambda idx: _original_sort_key(records[idx]))
        orig_idx = members_sorted[0]
        orig_rec = records[orig_idx]

        # Determine dominant match type across edges in component
        grp_edges = [
            info for (u, v), info in edge_info.items()
            if u in members and v in members
        ]
        kinds_set = {e[0] for e in grp_edges}
        if kinds_set == {"exact_content"}:
            grp_kind = "exact_content"
        elif "visual_similar" in kinds_set:
            grp_kind = "visual_similar"
        elif "doc_content_similar" in kinds_set:
            grp_kind = "doc_content_similar"
        else:
            grp_kind = "exact_content"

        member_dicts: list[dict[str, Any]] = []
        savable = 0
        max_pct = 0
        grp_snippet = orig_rec.get("snippet") or ""

        for _pos, idx in enumerate(members_sorted):
            rec = dict(records[idx])
            is_orig = (idx == orig_idx)
            if is_orig:
                sim_pct = 100
                m_kind = grp_kind
            else:
                u, v = (orig_idx, idx) if orig_idx < idx else (idx, orig_idx)
                direct = edge_info.get((u, v))
                if direct:
                    m_kind, sim_pct, ed_snip = direct
                else:
                    # Transitive connection: compute direct similarity against original if available
                    if idx in vis_sigs and orig_idx in vis_sigs:
                        sim_pct = max(thresh, _compute_image_similarity(vis_sigs[orig_idx], vis_sigs[idx]))
                        m_kind = "visual_similar"
                    elif idx in doc_shingles and orig_idx in doc_shingles:
                        sim_pct = max(thresh, _compute_text_similarity(doc_shingles[orig_idx][0], doc_shingles[idx][0]))
                        m_kind = "doc_content_similar"
                    else:
                        best_e = max(
                            (info for (eu, ev), info in edge_info.items() if eu == idx or ev == idx),
                            key=lambda x: x[1],
                            default=(grp_kind, thresh, ""),
                        )
                        m_kind, sim_pct, ed_snip = best_e
                    ed_snip = ""
                if not grp_snippet and ed_snip:
                    grp_snippet = ed_snip
                savable += int(rec["size_bytes"] or 0)
                max_pct = max(max_pct, sim_pct)

            rec["group_id"] = gid
            rec["role"] = "original" if is_orig else "duplicate"
            rec["decision"] = "keep" if is_orig else "exclude"
            rec["match_kind"] = m_kind
            rec["similarity_pct"] = int(sim_pct)
            rec["duplicate_of"] = "" if is_orig else orig_rec["file_name"]
            member_dicts.append(rec)
            flat_items.append(rec)

        if max_pct == 0:
            max_pct = 100
        groups.append({
            "group_id": gid,
            "match_kind": grp_kind,
            "similarity_pct": max_pct if grp_kind != "exact_content" else 100,
            "original_id": orig_rec["id"],
            "original_name": orig_rec["file_name"],
            "count": len(member_dicts),
            "duplicate_count": len(member_dicts) - 1,
            "savable_bytes": savable,
            "savable_label": _fmt_bytes(savable),
            "snippet": grp_snippet[:140],
            "reviewed": False,
            "items": member_dicts,
        })

    # Sort groups: exact matches & highest savings first
    groups.sort(key=lambda g: (0 if g["match_kind"] == "exact_content" else 1, -g["savable_bytes"], -g["similarity_pct"]))

    # Also store unique (non-duplicate) file paths so unique ZIP export can include everything clean
    dup_member_ids = {it["id"] for it in flat_items}
    clean_unique_paths = [
        {"id": r["id"], "file_name": r["file_name"], "rel_path": r["rel_path"], "abs_path": r["abs_path"]}
        for r in records if r["id"] not in dup_member_ids
    ]

    # New scan starts fresh: never leak previous execution reports/ZIP outputs into a new scan
    plan = {
        "version": 1,
        "source_kind": s_kind,
        "source_dir": str(base_dir),
        "batch_id": clean_batch,
        "scanned_at": dt.datetime.now().isoformat(timespec="seconds"),
        "options": {
            "scope": sc_mode,
            "min_similarity": thresh,
            "include_subfolders": bool(include_subfolders),
        },
        "total_files_scanned": n_total,
        "groups": groups,
        "clean_unique_files": clean_unique_paths,
        "export_unique_snapshot": [],
        "executed": False,
        "last_execution": None,
    }
    atomic_write_json(dups_plan_path(proj_root), plan)
    return {
        "ok": True,
        "counts": summarize_duplicates_plan(plan),
        "source_dir": str(base_dir),
        "groups_count": len(groups),
        "total_files_scanned": n_total,
    }


def summarize_duplicates_plan(plan: dict[str, Any] | None) -> dict[str, Any]:
    out = {
        "total_files_scanned": 0,
        "groups": 0,
        "active_groups": 0,
        "reviewed_groups": 0,
        "exact_groups": 0,
        "visual_groups": 0,
        "doc_groups": 0,
        "duplicate_files": 0,
        "selected_exclude": 0,
        "kept_files": 0,
        "savable_bytes": 0,
        "savable_label": "0 B",
    }
    if not isinstance(plan, dict) or not isinstance(plan.get("groups"), list):
        return out
    groups = plan["groups"]
    out["total_files_scanned"] = int(plan.get("total_files_scanned") or 0)
    out["groups"] = len(groups)
    sav_bytes = 0
    for g in groups:
        is_rev = bool(g.get("reviewed"))
        if is_rev:
            out["reviewed_groups"] += 1
        else:
            out["active_groups"] += 1
            mk = g.get("match_kind")
            if mk == "exact_content":
                out["exact_groups"] += 1
            elif mk == "visual_similar":
                out["visual_groups"] += 1
            elif mk == "doc_content_similar":
                out["doc_groups"] += 1
        for it in g.get("items") or []:
            if it.get("role") == "duplicate":
                out["duplicate_files"] += 1
            if it.get("decision") == "exclude":
                out["selected_exclude"] += 1
                sav_bytes += int(it.get("size_bytes") or 0)
            else:
                out["kept_files"] += 1
    out["savable_bytes"] = sav_bytes
    out["savable_label"] = _fmt_bytes(sav_bytes)
    return out


def _validated_dups_last_execution(proj_root: Path, plan: dict[str, Any]) -> dict[str, Any] | None:
    if not isinstance(plan, dict) or not plan.get("executed"):
        return None
    raw = plan.get("last_execution")
    if not isinstance(raw, dict):
        return None
    out = dict(raw)
    rep_exists = False
    rep_local = str(out.get("report_local") or "").strip()
    rep_rel = str(out.get("report_rel") or "").strip()
    for cand in ([Path(rep_local)] if rep_local else []) + ([proj_root / rep_rel] if rep_rel else []):
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
    out["zip_ready"] = bool(plan.get("export_unique_snapshot"))
    if not rep_exists and not out["zip_ready"]:
        return None
    return out


def get_duplicates_plan_page(
    proj_root: Path,
    *,
    page: int = 1,
    per_page: int = 20,
    match_filter: str = "all",
    query: str = "",
    show_reviewed: bool = False,
) -> dict[str, Any]:
    plan = read_json_safe(dups_plan_path(proj_root), {})
    has_undo = bool(read_json_safe(dups_undo_path(proj_root), {}).get("entries"))
    if not isinstance(plan, dict) or not isinstance(plan.get("groups"), list):
        return {
            "has_plan": False,
            "groups": [],
            "total_groups": 0,
            "total_filtered": 0,
            "page": 1,
            "per_page": 20,
            "total_pages": 1,
            "counts": summarize_duplicates_plan(None),
            "options": {},
            "source_dir": "",
            "source_kind": "",
            "last_execution": None,
            "has_undo": has_undo,
        }

    groups = plan["groups"]
    counts = summarize_duplicates_plan(plan)

    filtered = list(groups)
    if match_filter == "reviewed":
        filtered = [g for g in filtered if g.get("reviewed")]
    else:
        if not show_reviewed:
            filtered = [g for g in filtered if not g.get("reviewed")]
        if match_filter in ("exact_content", "visual_similar", "doc_content_similar"):
            filtered = [g for g in filtered if g.get("match_kind") == match_filter]
        elif match_filter == "has_selected":
            filtered = [g for g in filtered if any(it.get("decision") == "exclude" for it in (g.get("items") or []))]

    q_s = str(query or "").strip().lower()
    if q_s:
        filtered = [
            g for g in filtered
            if q_s in str(g.get("snippet") or "").lower()
            or any(
                q_s in str(it.get("file_name") or "").lower()
                or q_s in str(it.get("rel_path") or "").lower()
                for it in (g.get("items") or [])
            )
        ]

    per_p = max(5, min(100, int(per_page or 20)))
    total_f = len(filtered)
    total_pages = max(1, (total_f + per_p - 1) // per_p)
    pg = max(1, min(total_pages, int(page or 1)))
    start = (pg - 1) * per_p

    return {
        "has_plan": len(groups) > 0,
        "source_dir": plan.get("source_dir") or "",
        "source_kind": plan.get("source_kind") or "folder",
        "batch_id": plan.get("batch_id") or "",
        "scanned_at": plan.get("scanned_at") or "",
        "options": plan.get("options") or {},
        "counts": counts,
        "total_groups": len(groups),
        "total_filtered": total_f,
        "page": pg,
        "per_page": per_p,
        "total_pages": total_pages,
        "groups": filtered[start : start + per_p],
        "last_execution": _validated_dups_last_execution(proj_root, plan),
        "has_undo": has_undo,
    }


def _rebalance_group_keeps(grp: dict[str, Any]) -> None:
    items = grp.get("items") or []
    if not items:
        return
    # Guarantee at least one item stays "keep" (the original)
    has_orig = any(it.get("role") == "original" for it in items)
    if not has_orig:
        items[0]["role"] = "original"
        items[0]["decision"] = "keep"
    for it in items:
        if it.get("role") == "original":
            it["decision"] = "keep"
            grp["original_id"] = it["id"]
            grp["original_name"] = it["file_name"]
    if not any(it.get("decision") == "keep" for it in items):
        items[0]["role"] = "original"
        items[0]["decision"] = "keep"
    sav = sum(int(it.get("size_bytes") or 0) for it in items if it.get("decision") == "exclude")
    grp["savable_bytes"] = sav
    grp["savable_label"] = _fmt_bytes(sav)


def update_duplicates_plan(
    proj: dict[str, Any],
    *,
    item_id: str = "",
    group_id: str = "",
    action: str = "",
    bulk_action: str = "",
) -> dict[str, Any]:
    proj_root = resolve_local_or_unc(str(proj["root"]))
    pp = dups_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("groups"), list):
        raise ValueError(tr("dups.no_plan_err"))

    groups = plan["groups"]
    changed = 0

    if bulk_action == "select_all_dups":
        for g in groups:
            g["reviewed"] = False
            for it in g.get("items") or []:
                val = "keep" if it.get("role") == "original" else "exclude"
                if it.get("decision") != val:
                    it["decision"] = val
                    changed += 1
            _rebalance_group_keeps(g)
    elif bulk_action == "clear_selection":
        for g in groups:
            for it in g.get("items") or []:
                if it.get("decision") != "keep":
                    it["decision"] = "keep"
                    changed += 1
            _rebalance_group_keeps(g)
    elif group_id and action in ("keep_all_hide", "mark_reviewed", "reopen_group"):
        for g in groups:
            if str(g.get("group_id") or "") != str(group_id):
                continue
            if action == "keep_all_hide":
                for it in g.get("items") or []:
                    it["decision"] = "keep"
                g["reviewed"] = True
                _rebalance_group_keeps(g)
            elif action == "mark_reviewed":
                g["reviewed"] = True
            elif action == "reopen_group":
                g["reviewed"] = False
            changed += 1
            break
    elif item_id and action in ("set_original", "keep", "exclude", "toggle_exclude"):
        for g in groups:
            items = g.get("items") or []
            target = next((it for it in items if str(it.get("id")) == str(item_id)), None)
            if not target:
                continue
            if action == "set_original":
                for it in items:
                    if it["id"] == target["id"]:
                        it["role"] = "original"
                        it["decision"] = "keep"
                        it["duplicate_of"] = ""
                    else:
                        it["role"] = "duplicate"
                        it["duplicate_of"] = target["file_name"]
                        if it.get("decision") not in ("keep", "exclude"):
                            it["decision"] = "exclude"
                changed += 1
            elif action == "keep":
                target["decision"] = "keep"
                changed += 1
            elif action == "exclude":
                if target.get("role") != "original":
                    target["decision"] = "exclude"
                    changed += 1
            elif action == "toggle_exclude":
                if target.get("role") != "original":
                    target["decision"] = "keep" if target.get("decision") == "exclude" else "exclude"
                    changed += 1
            _rebalance_group_keeps(g)
            break

    atomic_write_json(pp, plan)
    return {"ok": True, "changed": changed, "counts": summarize_duplicates_plan(plan)}


def _write_duplicates_excel_report(
    report_path: Path,
    plan: dict[str, Any],
    exec_sum: dict[str, Any],
) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "ملخص المكررات"
    ws1.views.sheetView[0].rightToLeft = True

    hdr_fill = PatternFill("solid", fgColor="0F766E")
    hdr_font = Font(color="FFFFFF", bold=True, size=11)
    bold_font = Font(bold=True, size=11)
    keep_fill = PatternFill("solid", fgColor="ECFDF5")
    excl_fill = PatternFill("solid", fgColor="FEF3C7")

    cnt = summarize_duplicates_plan(plan)
    summary_rows = [
        ("المسار المفحوص", str(plan.get("source_dir") or "")),
        ("إجمالي الملفات المفحوصة", cnt["total_files_scanned"]),
        ("مجموعات التشابه المكتشفة", cnt["groups"]),
        ("تطابق تام بالمحتوى (100%)", cnt["exact_groups"]),
        ("صور متشابهة بصرياً", cnt["visual_groups"]),
        ("مستندات متشابهة بالمحتوى", cnt["doc_groups"]),
        ("وضع التنفيذ", exec_sum.get("mode_label", "")),
        ("ملفات معزولة أو محذوفة", exec_sum.get("processed", 0)),
        ("المساحة الموفرة", exec_sum.get("saved_label", "0 B")),
        ("تاريخ التنفيذ", exec_sum.get("executed_at", "")),
    ]
    ws1.append(["البند", "القيمة"])
    for c in range(1, 3):
        cell = ws1.cell(1, c)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for k, v in summary_rows:
        ws1.append([k, v])
        ws1.cell(ws1.max_row, 1).font = bold_font
    ws1.column_dimensions["A"].width = 34
    ws1.column_dimensions["B"].width = 48

    ws2 = wb.create_sheet("تفاصيل المجموعات")
    ws2.views.sheetView[0].rightToLeft = True
    headers = [
        "المجموعة", "نوع التشابه", "نسبة التشابه %", "الدور", "الإجراء",
        "اسم الملف", "المسار النسبي", "الأبعاد", "الحجم", "محتوى مشترك / الأصل",
    ]
    ws2.append(headers)
    for col_i in range(1, len(headers) + 1):
        cell = ws2.cell(1, col_i)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    kind_labels = {
        "exact_content": "تطابق تام بالمحتوى",
        "visual_similar": "تشابه بصري للصور",
        "doc_content_similar": "تشابه محتوى المستند",
    }
    for g in plan.get("groups") or []:
        g_kind = kind_labels.get(g.get("match_kind", ""), "تشابه")
        for it in g.get("items") or []:
            dims = f"{it['width']}x{it['height']}" if it.get("width") and it.get("height") else ""
            role_lbl = "الأصل (احتفاظ)" if it.get("role") == "original" else "نسخة مكررة"
            act_lbl = "احتفاظ" if it.get("decision") == "keep" else exec_sum.get("action_verb", "عزل")
            note = g.get("snippet") or it.get("duplicate_of") or ""
            ws2.append([
                g.get("group_id", ""),
                kind_labels.get(it.get("match_kind", ""), g_kind),
                int(it.get("similarity_pct") or 0),
                role_lbl,
                act_lbl,
                it.get("file_name", ""),
                it.get("rel_path", ""),
                dims,
                it.get("size_label", ""),
                note,
            ])
            row_idx = ws2.max_row
            rfill = keep_fill if it.get("decision") == "keep" else excl_fill
            for ci in range(1, len(headers) + 1):
                c = ws2.cell(row_idx, ci)
                c.fill = rfill
                c.alignment = Alignment(vertical="center", wrap_text=True)

    widths = [13, 22, 15, 16, 13, 30, 38, 13, 12, 34]
    for ci, w in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    wb.save(report_path)


def execute_duplicates_plan(
    proj: dict[str, Any],
    *,
    mode: str = "quarantine",
    confirm_delete: bool = False,
    stop_event: threading.Event | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    proj_root = resolve_local_or_unc(str(proj["root"]))
    pp = dups_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if not isinstance(plan, dict) or not isinstance(plan.get("groups"), list) or not plan["groups"]:
        raise ValueError(tr("dups.no_plan_err"))

    exec_mode = str(mode or "quarantine").strip().lower()
    if exec_mode not in ("quarantine", "delete_permanent", "export_unique_zip"):
        exec_mode = "quarantine"
    if exec_mode == "delete_permanent" and not confirm_delete:
        raise ValueError(tr("dups.confirm_delete_required"))

    base_dir = resolve_local_or_unc(str(plan.get("source_dir") or ""))
    ts_tag = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    # Safety pass: make sure every group keeps at least 1 file
    for g in plan["groups"]:
        _rebalance_group_keeps(g)

    last_emit = [0.0]
    targets = [
        it for g in plan["groups"] for it in (g.get("items") or [])
        if it.get("decision") == "exclude"
    ]
    if exec_mode in ("quarantine", "delete_permanent") and not targets:
        raise ValueError(tr("dups.no_items_to_exec_err"))
    if exec_mode == "export_unique_zip":
        kept_cnt = sum(1 for g in plan["groups"] for it in (g.get("items") or []) if it.get("decision") == "keep")
        if kept_cnt + len(plan.get("clean_unique_files") or []) == 0:
            raise ValueError(tr("dups.no_items_to_exec_err"))

    total_targets = max(1, len(targets))
    _emit_dup_prog(progress_cb, last_emit, {
        "phase": "executing", "processed": 0, "total": total_targets, "current_file": "",
    }, force=True)

    processed = 0
    saved_bytes = 0
    errors = 0
    undo_entries: list[dict[str, str]] = []
    q_dir = dups_quarantine_root(proj_root) / ts_tag

    if exec_mode in ("quarantine", "delete_permanent"):
        for t_idx, it in enumerate(targets):
            if stop_event is not None and stop_event.is_set():
                if undo_entries:
                    existing_undo = read_json_safe(dups_undo_path(proj_root), {})
                    merged = list(existing_undo.get("entries") or []) if isinstance(existing_undo, dict) else []
                    atomic_write_json(dups_undo_path(proj_root), {
                        "executed_at": dt.datetime.now().isoformat(timespec="seconds"),
                        "quarantine_dir": str(q_dir),
                        "entries": merged + undo_entries,
                    })
                raise DuplicatesStoppedError(tr("jobs.stopped_by_user"))
            src_p = resolve_local_or_unc(str(it.get("abs_path") or ""))
            _emit_dup_prog(progress_cb, last_emit, {
                "phase": "executing", "processed": t_idx + 1, "total": total_targets,
                "current_file": str(it.get("file_name") or src_p.name),
            })
            try:
                rel_check = src_p.relative_to(base_dir)
                safe_src = safe_resolve_within(base_dir, rel_check.as_posix(), allow_base=False, reject_symlinks=True)
            except Exception:
                errors += 1
                continue
            if not safe_src.exists() or not safe_src.is_file():
                errors += 1
                continue
            sz = int(it.get("size_bytes") or safe_src.stat().st_size or 0)
            if exec_mode == "quarantine":
                dest_p = safe_resolve_within(q_dir, rel_check.as_posix(), allow_base=False, reject_symlinks=True)
                dest_p.parent.mkdir(parents=True, exist_ok=True)
                # Avoid collision if needed
                if dest_p.exists():
                    dest_p = dest_p.with_name(f"{dest_p.stem}_{uuid.uuid4().hex[:5]}{dest_p.suffix}")
                shutil.move(str(safe_src), str(dest_p))
                undo_entries.append({"src_path": str(safe_src), "quarantine_path": str(dest_p)})
                processed += 1
                saved_bytes += sz
            else:
                safe_src.unlink()
                processed += 1
                saved_bytes += sz

        if processed == 0:
            raise ValueError(tr("dups.no_items_to_exec_err"))

    if stop_event is not None and stop_event.is_set():
        raise DuplicatesStoppedError(tr("jobs.stopped_by_user"))

    # Build Excel report only after genuine execution completion
    rep_dir = proj_root / "_System_Agent" / "reports"
    rep_dir.mkdir(parents=True, exist_ok=True)
    rep_path = rep_dir / f"تقرير_المكررات_{ts_tag}.xlsx"

    mode_labels = {
        "quarantine": "عزل المكررات إلى مجلد العزل",
        "delete_permanent": "حذف المكررات نهائياً",
        "export_unique_zip": "تصدير النسخ الفريدة في حزمة مضغوطة",
    }
    action_verbs = {
        "quarantine": "تم العزل",
        "delete_permanent": "تم الحذف",
        "export_unique_zip": "مستبعد من الحزمة",
    }
    if exec_mode == "export_unique_zip":
        for g in plan["groups"]:
            for it in g.get("items") or []:
                if it.get("decision") == "exclude":
                    processed += 1
                    saved_bytes += int(it.get("size_bytes") or 0)

    exec_sum: dict[str, Any] = {
        "mode": exec_mode,
        "mode_label": mode_labels[exec_mode],
        "action_verb": action_verbs[exec_mode],
        "processed": processed,
        "errors": errors,
        "saved_bytes": saved_bytes,
        "saved_label": _fmt_bytes(saved_bytes),
        "quarantine_dir": str(q_dir) if exec_mode == "quarantine" else "",
        "report_rel": rep_path.relative_to(proj_root).as_posix(),
        "report_local": str(rep_path),
        "executed_at": dt.datetime.now().isoformat(timespec="seconds"),
        "cleared": True,
        "executed": True,
    }
    try:
        _write_duplicates_excel_report(rep_path, plan, exec_sum)
        if not rep_path.exists() or rep_path.stat().st_size == 0:
            raise RuntimeError("Failed to generate report")
    except Exception:
        rep_path.unlink(missing_ok=True)
        raise
    exec_sum["report_exists"] = True

    if undo_entries:
        atomic_write_json(dups_undo_path(proj_root), {
            "executed_at": exec_sum["executed_at"],
            "quarantine_dir": str(q_dir),
            "entries": undo_entries,
        })

    # Keep snapshots of clean unique files for download-zip even after clearing review UI
    kept_from_groups = [
        {
            "id": it["id"],
            "file_name": it["file_name"],
            "rel_path": it["rel_path"],
            "abs_path": it["abs_path"],
        }
        for g in plan["groups"]
        for it in (g.get("items") or [])
        if it.get("decision") == "keep"
    ]
    snapshot = list(plan.get("clean_unique_files") or []) + kept_from_groups
    exec_sum["zip_ready"] = bool(snapshot)
    plan["export_unique_snapshot"] = snapshot
    # Automatically reset review screen after execution
    plan["groups"] = []
    plan["clean_unique_files"] = []
    plan["total_files_scanned"] = 0
    plan["executed"] = True
    plan["last_execution"] = exec_sum
    atomic_write_json(pp, plan)

    return {
        "ok": True,
        **exec_sum,
        "has_undo": bool(undo_entries or read_json_safe(dups_undo_path(proj_root), {}).get("entries")),
        "counts": summarize_duplicates_plan(plan),
    }


def build_duplicates_unique_zip(proj: dict[str, Any]) -> Path:
    proj_root = resolve_local_or_unc(str(proj["root"]))
    plan = read_json_safe(dups_plan_path(proj_root), {})
    if not isinstance(plan, dict) or not plan.get("executed") or not isinstance(plan.get("last_execution"), dict):
        raise ValueError(tr("dups.exec_required_zip_err"))

    files_to_pack: list[dict[str, Any]] = list(plan.get("export_unique_snapshot") or [])
    if not files_to_pack:
        raise ValueError(tr("dups.exec_required_zip_err"))

    base_dir = resolve_local_or_unc(str(plan.get("source_dir") or ""))
    out_dir = proj_root / "_System_Agent" / "exports"
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = out_dir / f"unique_files_{ts}.zip"

    used_rels: set[str] = set()
    written_files = 0
    try:
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for rec in files_to_pack:
                fp = resolve_local_or_unc(str(rec.get("abs_path") or ""))
                try:
                    rel_p = fp.relative_to(base_dir)
                    safe_fp = safe_resolve_within(base_dir, rel_p.as_posix(), allow_base=False, reject_symlinks=True)
                except Exception:
                    continue
                if not safe_fp.exists() or not safe_fp.is_file():
                    continue
                arc_name = str(rec.get("rel_path") or safe_fp.name).replace("\\", "/")
                if arc_name in used_rels:
                    arc_name = f"{Path(arc_name).stem}_{uuid.uuid4().hex[:4]}{Path(arc_name).suffix}"
                used_rels.add(arc_name)
                zf.write(safe_fp, arcname=arc_name)
                written_files += 1

            last_ex = plan.get("last_execution") or {}
            rep_local = str(last_ex.get("report_local") or "").strip()
            if rep_local:
                rp = Path(rep_local)
                if rp.exists() and rp.is_file():
                    zf.write(rp, arcname=rp.name)
        if written_files == 0:
            raise ValueError(tr("dups.no_items_to_exec_err"))
    except Exception:
        zip_path.unlink(missing_ok=True)
        raise

    return zip_path


def undo_last_duplicates_quarantine(proj: dict[str, Any]) -> dict[str, Any]:
    proj_root = resolve_local_or_unc(str(proj["root"]))
    up = dups_undo_path(proj_root)
    undo_data = read_json_safe(up, {})
    entries = undo_data.get("entries") if isinstance(undo_data, dict) else None
    if not entries:
        raise ValueError(tr("dups.no_undo_err"))

    restored = 0
    errors = 0
    failed_entries: list[dict[str, Any]] = []
    q_base = dups_quarantine_root(proj_root)
    for ent in reversed(entries):
        src_p = resolve_local_or_unc(str(ent.get("src_path") or ""))
        q_p = resolve_local_or_unc(str(ent.get("quarantine_path") or ""))
        try:
            rel_q = q_p.relative_to(q_base)
            safe_q = safe_resolve_within(q_base, rel_q.as_posix(), allow_base=False, reject_symlinks=True)
        except Exception:
            errors += 1
            failed_entries.append(ent)
            continue
        if not safe_q.exists() or not safe_q.is_file():
            errors += 1
            failed_entries.append(ent)
            continue
        if src_p.exists():
            errors += 1
            failed_entries.append(ent)
            continue
        src_p.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(safe_q), str(src_p))
        restored += 1

    q_dir_raw = str(undo_data.get("quarantine_dir") or "").strip()
    if q_dir_raw:
        try:
            q_sub = Path(q_dir_raw)
            if q_sub.exists() and q_sub.is_dir() and not any(q_sub.rglob("*")):
                shutil.rmtree(q_sub, ignore_errors=True)
        except Exception:
            pass

    if failed_entries:
        atomic_write_json(up, {
            "executed_at": undo_data.get("executed_at", ""),
            "quarantine_dir": q_dir_raw,
            "entries": list(reversed(failed_entries)),
        })
    else:
        try:
            up.unlink(missing_ok=True)
        except Exception:
            pass

    pp = dups_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if isinstance(plan, dict) and not failed_entries:
        old_ex = plan.get("last_execution") or {}
        rep_local = str(old_ex.get("report_local") or "").strip() if isinstance(old_ex, dict) else ""
        if rep_local:
            try:
                Path(rep_local).unlink(missing_ok=True)
            except Exception:
                pass
        plan["groups"] = []
        plan["clean_unique_files"] = []
        plan["export_unique_snapshot"] = []
        plan["executed"] = False
        plan["last_execution"] = None
        atomic_write_json(pp, plan)

    return {"ok": True, "restored": restored, "errors": errors, "has_undo": bool(failed_entries)}


def clear_duplicates_plan(proj: dict[str, Any]) -> dict[str, Any]:
    proj_root = resolve_local_or_unc(str(proj["root"]))
    pp = dups_plan_path(proj_root)
    plan = read_json_safe(pp, {})
    if isinstance(plan, dict):
        plan["groups"] = []
        plan["clean_unique_files"] = []
        plan["export_unique_snapshot"] = []
        plan["total_files_scanned"] = 0
        plan["executed"] = False
        plan["last_execution"] = None
        atomic_write_json(pp, plan)
    return {"ok": True, "counts": summarize_duplicates_plan(None)}
