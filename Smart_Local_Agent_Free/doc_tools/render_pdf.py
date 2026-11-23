#!/usr/bin/env python3
"""
render_pdf.py — تحويل صفحات PDF الممسوحة ضوئياً إلى صور PNG واضحة لقراءة الأختام والتوقيعات والفواتير.
يعتمد على PyMuPDF (يعمل على Windows/Mac/Linux مباشرة بدون Poppler).
"""
import argparse
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None


def parse_pages(spec: str, total: int) -> list[int]:
    if not spec or spec.lower() == "all":
        return list(range(total))
    result = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start = max(1, int(a.strip()))
                end = min(total, int(b.strip()))
            except ValueError as exc:
                raise ValueError(f"نطاق صفحات غير صالح: {part!r}") from exc
            if start <= end:
                result.extend(range(start - 1, end))
        else:
            try:
                idx = int(part) - 1
            except ValueError as exc:
                raise ValueError(f"رقم صفحة غير صالح: {part!r}") from exc
            if 0 <= idx < total:
                result.append(idx)
    return sorted(set(result))


def render_one(pdf_path: Path, out_dir: Path, dpi: int, pages_spec: str, zone: str) -> list[str]:
    if fitz is None:
        raise RuntimeError("مكتبة PyMuPDF غير مثبتة؛ يلزم تثبيتها عبر requirements.txt.")
    doc = fitz.open(str(pdf_path))
    indices = parse_pages(pages_spec, len(doc))
    saved = []
    safe_stem = pdf_path.stem.replace(" ", "_")

    for idx in indices:
        page = doc[idx]
        rect = page.rect
        if zone == "header":
            clip = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.42)
        elif zone == "footer":
            clip = fitz.Rect(rect.x0, rect.y0 + rect.height * 0.48, rect.x1, rect.y1)
        elif zone == "middle":
            clip = fitz.Rect(rect.x0, rect.y0 + rect.height * 0.25, rect.x1, rect.y0 + rect.height * 0.75)
        else:
            clip = rect

        zoom = dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
        suffix = f"_{zone}" if zone != "full" else ""
        out_path = out_dir / f"{safe_stem}_p{idx + 1:02d}{suffix}.png"
        pix.save(str(out_path))
        saved.append(str(out_path))

    doc.close()
    return saved


def main():
    ap = argparse.ArgumentParser(description="Render scanned PDF pages/zones to PNG images.")
    ap.add_argument("pdfs", nargs="+", help="One or more PDF file paths")
    ap.add_argument("--out", default="_agent_work/pages", help="Output folder for PNGs")
    ap.add_argument("--dpi", type=int, default=150, help="Resolution (default 150; use 220 for small handwriting/stamps)")
    ap.add_argument("--pages", default="1-2", help="Page range e.g. '1', '1-2', '1,3', or 'all'")
    ap.add_argument("--zone", choices=["full", "header", "middle", "footer"], default="full",
                    help="Crop region: full page, header (ref/date/subject), middle (table), footer (stamps/signatures)")
    args = ap.parse_args()

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    for raw in args.pdfs:
        p = Path(raw)
        if not p.exists():
            print(f"[SKIP] غير موجود: {p}")
            continue
        imgs = render_one(p, out_dir, args.dpi, args.pages, args.zone)
        for im in imgs:
            print(f"[OK] {p.name} -> {im}")


if __name__ == "__main__":
    main()
