#!/usr/bin/env python3
"""
apply_plan.py — تنفيذ آمن لخطة التنظيم وإعادة التسمية وإزالة التكرار (نسخ دون حذف الأصل + تحقق 100%).

صيغة ملف الخطة JSON:
[
  {"src": "path/to/old.pdf", "action": "KEEP",       "dest_folder": "EL - الكهرباء/MOH", "new_name": "EL-135-R01-MOH.pdf", "note": "تم تصحيح الرقم من 136 إلى 135"},
  {"src": "path/to/sup.pdf", "action": "SUPERSEDED", "dest_folder": "EL - الكهرباء/_قبل التعديل (Superseded)", "new_name": "15-قبل_التعديل_(EL-019).pdf", "note": "ترقيم قديم"},
  {"src": "path/to/dup.pdf", "action": "SKIP",       "reason": "نسخة مكررة من EL-098-MOH.pdf (تم الاحتفاظ بالنسخة ذات الصفحتين)"}
]
"""
import argparse
import json
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

INVALID_WIN_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def clean_filename(name: str) -> str:
    cleaned = INVALID_WIN_CHARS.sub("-", name).strip().rstrip(".")
    return re.sub(r"\s+", " ", cleaned)


def pdf_quality_score(path: Path, priority: int = 1) -> tuple:
    size = path.stat().st_size if path.exists() else 0
    pages = 1
    if fitz is not None and path.suffix.lower() == ".pdf":
        try:
            doc = fitz.open(str(path))
            pages = len(doc)
            doc.close()
        except Exception:
            pages = 1
    # يفضّل الأولوية الأعلى، ثم الصفحات الأكثر (المرفقات الكاملة)، ثم الحجم/الدقة الأعلى
    return (priority, pages, size)


def main():
    ap = argparse.ArgumentParser(description="Safely execute organization/deduplication plan without touching originals.")
    ap.add_argument("plan_json", help="Path to JSON plan")
    ap.add_argument("--root", default=".", help="Project root folder")
    ap.add_argument("--out", default="Organized_Archive", help="Destination archive folder")
    ap.add_argument("--expected-total", type=int, default=None, help="Verify total processed equals inventory count")
    args = ap.parse_args()

    root = Path(args.root).resolve()
    out_root = Path(args.out)
    if not out_root.is_absolute():
        out_root = root / out_root
    out_root.mkdir(parents=True, exist_ok=True)

    items = json.loads(Path(args.plan_json).read_text(encoding="utf-8"))

    # تجميع المرشحين لنفس المسار النهائي واختيار النسخة الأفضل تلقائياً عند التعارض
    grouped = {}
    skipped = []
    missing = []

    for item in items:
        action = str(item.get("action", "KEEP")).upper()
        src_raw = item["src"]
        src = Path(src_raw) if Path(src_raw).is_absolute() else (root / src_raw)

        if action == "SKIP":
            skipped.append({"src": src_raw, "reason": item.get("reason", "duplicate")})
            continue

        if not src.exists():
            missing.append(src_raw)
            continue

        dest_folder = item.get("dest_folder", "General").strip("/\\")
        raw_parts = [p.strip() for p in re.split(r"[/\\]+", dest_folder) if p.strip()]
        if any(p in (".", "..") for p in raw_parts):
            raise ValueError(f"Unsafe path segment in dest_folder: {dest_folder}")
        safe_parts = [clean_filename(p) for p in raw_parts]
        safe_parts = [p for p in safe_parts if p and p not in (".", "..")]
        if not safe_parts:
            safe_parts = ["General"]
        new_name = clean_filename(item.get("new_name") or src.name)
        if not new_name or new_name in (".", ".."):
            new_name = clean_filename(src.name) or "document"
        target_rel = "/".join(safe_parts + [new_name])
        prio = int(item.get("priority", 1))
        score = pdf_quality_score(src, prio)

        grouped.setdefault(target_rel, []).append({
            "score": score,
            "src": src,
            "src_raw": src_raw,
            "action": action,
            "dest_folder": "/".join(safe_parts),
            "new_name": new_name,
            "note": item.get("note", ""),
        })

    copied = []
    superseded = []
    folder_counts = Counter()

    for target_rel, candidates in sorted(grouped.items()):
        candidates.sort(key=lambda c: c["score"], reverse=True)
        winner = candidates[0]
        for loser in candidates[1:]:
            skipped.append({
                "src": loser["src_raw"],
                "reason": f"تم استبعادها تلقائياً لصالح النسخة الأكمل/الأوضح ({winner['src_raw']} -> {target_rel})",
            })

        dest_path = out_root / target_rel
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(winner["src"], dest_path)

        rec = {
            "src": winner["src_raw"],
            "dest": str(dest_path.relative_to(root)).replace("\\", "/") if dest_path.is_relative_to(root) else str(dest_path),
            "folder": winner["dest_folder"],
            "name": winner["new_name"],
            "pages": winner["score"][1],
            "size_bytes": winner["score"][2],
            "note": winner["note"],
        }
        folder_counts[winner["dest_folder"]] += 1
        if winner["action"] == "SUPERSEDED":
            superseded.append(rec)
        else:
            copied.append(rec)

    accounted = len(copied) + len(superseded) + len(skipped)
    audit = {
        "output_root": str(out_root),
        "total_in_plan": len(items),
        "kept_clean": len(copied),
        "kept_superseded": len(superseded),
        "skipped_duplicates": len(skipped),
        "missing_files": missing,
        "sum_verified": (accounted == len(items) and len(missing) == 0),
        "folder_counts": dict(sorted(folder_counts.items())),
        "kept_files": copied,
        "superseded_files": superseded,
        "skipped_files": skipped,
    }

    work_dir = root / "_agent_work"
    work_dir.mkdir(parents=True, exist_ok=True)
    audit_path = work_dir / "organization_audit.json"
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")

    print("=" * 65)
    print(f"تم تنفيذ الأرشفة بنجاح في: {out_root}")
    print(f"  ملفات نظيفة معتمدة     : {len(copied)}")
    print(f"  ملفات قبل التعديل محفوظة: {len(superseded)}")
    print(f"  نسخ مكررة مستبعدة       : {len(skipped)}")
    print(f"  المجموع المتحقق منه     : {accounted} / {len(items)} | مفقود: {len(missing)}")
    if args.expected_total is not None and accounted != args.expected_total:
        print(f"  [تحذير] العدد المتوقع {args.expected_total} يختلف عن {accounted}!")
    print("\nتوزيع المجلدات:")
    for fld, cnt in sorted(folder_counts.items()):
        print(f"  - {fld}: {cnt} ملف")
    print(f"\nسجل المراجعة الكامل (لاستخدامه في ربط الإكسل): {audit_path}")
    print("=" * 65)


if __name__ == "__main__":
    main()
