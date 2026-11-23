#!/usr/bin/env python3
"""
scan_archive.py — فحص شامل لأي مجلد مشروع + فك RAR/ZIP/7Z + كشف التكرار + حساب الصفحات والنصوص.
يعمل محلياً 100% بدون إنترنت وبدون أي تكلفة.
"""
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

try:
    from PIL import Image
except ImportError:
    Image = None

ARCHIVE_EXTS = {".rar", ".zip", ".7z"}
DOC_EXTS = {
    ".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp",
    ".xlsx", ".xlsm", ".xls", ".csv", ".docx", ".doc",
}
IGNORE_DIRS = {"_agent_work", "_Backups", "Organized_Archive", ".git", ".gemini", "__pycache__", "node_modules"}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def extract_archive(arc_path: Path, dest_dir: Path) -> tuple[bool, str]:
    dest_dir.mkdir(parents=True, exist_ok=True)
    ext = arc_path.suffix.lower()

    # 1) Python built-in zipfile for standard .zip (Zip-Slip, symlink, and Zip-Bomb guarded)
    if ext == ".zip":
        import zipfile
        max_files = 5000
        max_total_bytes = 2 * 1024 * 1024 * 1024
        dest_res = dest_dir.resolve()
        try:
            with zipfile.ZipFile(arc_path, "r") as zf:
                infos = [i for i in zf.infolist() if not i.is_dir()]
                if len(infos) > max_files:
                    return False, "حجم الأرشيف أو عدد ملفاته يتجاوز الحد المسموح"
                if sum(max(0, int(getattr(i, "file_size", 0) or 0)) for i in infos) > max_total_bytes:
                    return False, "حجم الأرشيف يتجاوز الحد المسموح"
                total_written = 0
                for info in infos:
                    raw_name = info.filename.replace("\\", "/").lstrip("/")
                    if not raw_name or raw_name.startswith("__MACOSX/"):
                        continue
                    if ((info.external_attr >> 16) & 0o170000) == 0o120000:
                        continue
                    parts = [p for p in raw_name.split("/") if p not in ("", ".")]
                    if not parts or any(p == ".." for p in parts):
                        return False, f"مسار غير آمن داخل الأرشيف: {raw_name}"
                    target = (dest_res / Path(*parts)).resolve()
                    try:
                        target.relative_to(dest_res)
                    except ValueError:
                        return False, f"مسار غير آمن داخل الأرشيف: {raw_name}"
                    target.parent.mkdir(parents=True, exist_ok=True)
                    with zf.open(info, "r") as src_f, open(target, "wb") as out_f:
                        while True:
                            buf = src_f.read(64 * 1024)
                            if not buf:
                                break
                            total_written += len(buf)
                            if total_written > max_total_bytes:
                                return False, "حجم الأرشيف يتجاوز الحد المسموح"
                            out_f.write(buf)
            return True, "zipfile"
        except Exception:
            pass

    # 2) tar / bsdtar (built into Windows 10/11, macOS, and Linux — supports RAR5, ZIP, 7Z!)
    for cmd_name in ("bsdtar", "tar"):
        exe = shutil.which(cmd_name)
        if exe:
            res = subprocess.run(
                [exe, "-xf", str(arc_path), "-C", str(dest_dir)],
                capture_output=True, text=True, errors="replace"
            )
            if res.returncode == 0:
                return True, cmd_name

    # 3) 7-Zip fallback on Windows/Linux
    seven_z_candidates = [
        shutil.which("7z"),
        shutil.which("7za"),
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        r"C:\Program Files\WinRAR\UnRAR.exe",
    ]
    for exe in seven_z_candidates:
        if exe and os.path.exists(exe):
            if exe.lower().endswith("unrar.exe"):
                cmd = [exe, "x", "-o+", "-y", str(arc_path), str(dest_dir) + os.sep]
            else:
                cmd = [exe, "x", "-y", f"-o{dest_dir}", str(arc_path)]
            res = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
            if res.returncode == 0:
                return True, Path(exe).name

    return False, "لم يتم العثور على أداة فك مناسبة (ثبّت 7-Zip المجاني إذا فشل فك RAR)"


def inspect_pdf(path: Path) -> dict:
    info = {"pages": None, "has_text": False, "text_preview": "", "is_scanned": True}
    if fitz is None:
        return info
    try:
        doc = fitz.open(str(path))
        info["pages"] = len(doc)
        texts = []
        for i in range(min(len(doc), 3)):
            t = doc[i].get_text("text").strip()
            if t:
                texts.append(t)
        joined = "\n".join(texts).strip()
        if len(joined) >= 40:
            info["has_text"] = True
            info["is_scanned"] = False
            info["text_preview"] = joined[:600]
        doc.close()
    except Exception as e:
        info["error"] = str(e)
    return info


def inspect_image(path: Path) -> dict:
    if Image is None:
        return {}
    try:
        with Image.open(path) as im:
            return {"width": im.width, "height": im.height, "mode": im.mode}
    except Exception as e:
        return {"error": str(e)}


def main():
    ap = argparse.ArgumentParser(description="Scan project folder, extract archives safely, detect duplicates.")
    ap.add_argument("folder", nargs="?", default=".", help="Target folder to scan")
    ap.add_argument("--extract", action="store_true", help="Extract RAR/ZIP/7Z into _agent_work/extracted/")
    ap.add_argument("--out", default=None, help="Output inventory JSON path")
    args = ap.parse_args()

    root = Path(args.folder).resolve()
    work_dir = root / "_agent_work"
    work_dir.mkdir(parents=True, exist_ok=True)
    extract_root = work_dir / "extracted"
    out_json = Path(args.out) if args.out else (work_dir / "inventory.json")

    archives_report = []
    if args.extract:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in IGNORE_DIRS]
            for fn in sorted(filenames):
                p = Path(dirpath) / fn
                if p.suffix.lower() in ARCHIVE_EXTS:
                    target = extract_root / p.stem
                    ok, tool = extract_archive(p, target)
                    count = sum(1 for x in target.rglob("*") if x.is_file()) if ok else 0
                    archives_report.append({
                        "archive": str(p.relative_to(root)),
                        "size_bytes": p.stat().st_size,
                        "extracted_to": str(target.relative_to(root)),
                        "success": ok,
                        "tool": tool,
                        "extracted_files": count,
                    })

    scan_roots = [("folder", root)]
    if extract_root.exists():
        scan_roots.append(("extracted", extract_root))

    files_list = []
    by_hash = defaultdict(list)

    for source_kind, sroot in scan_roots:
        for dirpath, dirnames, filenames in os.walk(sroot):
            if source_kind == "folder":
                dirnames[:] = [d for d in dirnames if d not in IGNORE_DIRS]
            for fn in sorted(filenames):
                if fn.startswith("~$") or fn == ".DS_Store" or fn.lower() == "thumbs.db":
                    continue
                p = Path(dirpath) / fn
                ext = p.suffix.lower()
                if ext in ARCHIVE_EXTS:
                    continue
                if ext not in DOC_EXTS:
                    continue
                rel = str(p.relative_to(root)).replace("\\", "/")
                size = p.stat().st_size
                digest = sha256_file(p)
                entry = {
                    "id": len(files_list),
                    "rel_path": rel,
                    "name": fn,
                    "ext": ext,
                    "source": source_kind,
                    "size_bytes": size,
                    "sha256": digest[:16],
                }
                if ext == ".pdf":
                    entry.update(inspect_pdf(p))
                elif ext in {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}:
                    entry.update(inspect_image(p))
                files_list.append(entry)
                by_hash[digest].append(rel)

    exact_dupes = [
        {"keep_suggestion": paths[0], "duplicates": paths[1:], "count": len(paths)}
        for paths in by_hash.values() if len(paths) > 1
    ]

    ext_counts = defaultdict(int)
    for f in files_list:
        ext_counts[f["ext"]] += 1

    report = {
        "root_folder": str(root),
        "archives": archives_report,
        "summary": {
            "total_files": len(files_list),
            "by_extension": dict(ext_counts),
            "exact_duplicate_groups": len(exact_dupes),
            "exact_duplicate_extra_files": sum(g["count"] - 1 for g in exact_dupes),
            "scanned_pdfs": sum(1 for f in files_list if f["ext"] == ".pdf" and f.get("is_scanned")),
            "digital_pdfs": sum(1 for f in files_list if f["ext"] == ".pdf" and f.get("has_text")),
            "multi_page_pdfs": sum(1 for f in files_list if (f.get("pages") or 0) > 1),
            "excel_files": [f["rel_path"] for f in files_list if f["ext"] in {".xlsx", ".xlsm", ".xls", ".csv"}],
        },
        "exact_duplicates": exact_dupes,
        "files": files_list,
    }

    out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    s = report["summary"]
    print("=" * 65)
    print(f"فحص المجلد اكتمل: {root}")
    if archives_report:
        for a in archives_report:
            st = "OK" if a["success"] else "FAIL"
            print(f"  [{st}] أرشيف {a['archive']} -> {a['extracted_files']} ملف ({a['tool']})")
    print(f"إجمالي الملفات المكتشفة : {s['total_files']} | التوزيع: {s['by_extension']}")
    print(f"ملفات PDF ممسوحة (صور)  : {s['scanned_pdfs']} | نصية: {s['digital_pdfs']} | متعددة الصفحات: {s['multi_page_pdfs']}")
    print(f"نسخ متطابقة تماماً (SHA) : {s['exact_duplicate_extra_files']} ملف مكرر في {s['exact_duplicate_groups']} مجموعة")
    if s["excel_files"]:
        print(f"ملفات الإكسل/السجلات     : {s['excel_files']}")
    print(f"تم حفظ الجرد الكامل في  : {out_json}")
    print("=" * 65)


if __name__ == "__main__":
    main()
