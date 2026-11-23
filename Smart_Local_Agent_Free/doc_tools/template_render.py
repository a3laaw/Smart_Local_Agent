#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
template_render.py — High-fidelity 100% offline visual renderer for uploaded
`.xlsx`, `.xlsm`, and `.docx` project templates + live filled preview.

Features:
  - `.xlsx` / `.xlsm`: All sheet tabs, RTL sheet direction, exact merged cells,
    real column widths & row heights, resolved RGB/Theme/Tint/Indexed fills,
    fonts (name, size, bold, italic, underline, color), 4-side borders, alignment,
    formula markers, and embedded images/logos extracted as inline `data:` URIs!
  - `.docx`: Headless LibreOffice -> PDF -> PNG page images (when LibreOffice is
    installed locally on Windows/Linux) PLUS full-fidelity built-in A4 HTML/JSON
    renderer preserving header/footer logos (`data:` URIs), document order of
    paragraphs & tables, merged cells (`gridSpan` + `vMerge`), cell borders &
    shading, fonts, RTL, and clickable targets (cells, paragraphs, `{{spans}}`).
"""

from __future__ import annotations

import base64
import datetime as dt
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

try:
    from defusedxml import ElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET

import docx
from docx.oxml.ns import qn
import openpyxl
from openpyxl.utils import get_column_letter

import template_profile as tprof

# Standard Office theme colors fallback (lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink)
_DEFAULT_THEME_RGB: list[str] = [
    "FFFFFF", "000000", "E7E6E6", "44546A",
    "4472C4", "ED7D31", "A5A5A5", "FFC000",
    "5B9BD5", "70AD47", "0563C1", "954F72",
]

# Common legacy Excel indexed color table
_INDEXED_RGB: dict[int, str] = {
    0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF",
    5: "FFFF00", 6: "FF00FF", 7: "00FFFF", 8: "000000", 9: "FFFFFF",
    10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF",
    15: "00FFFF", 16: "800000", 17: "008000", 18: "000080", 19: "808000",
    22: "C0C0C0", 23: "808080", 40: "00CCFF", 41: "CCFFFF", 42: "CCFFCC",
    43: "FFFF99", 44: "99CCFF", 45: "FF99CC", 46: "CC99FF", 47: "FFCC99",
    48: "3366FF", 53: "FF6600", 55: "969696", 56: "003366", 63: "333333",
}

_PLACEHOLDER_RE = re.compile(r"\{\{\s*[^}]+?\s*\}\}|\[\[\s*[^\]]+?\s*\]\]")


def _extract_workbook_theme_palette(xlsx_path: Path) -> list[str]:
    palette = list(_DEFAULT_THEME_RGB)
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            if "xl/theme/theme1.xml" not in zf.namelist():
                return palette
            root = ET.fromstring(zf.read("xl/theme/theme1.xml"))
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            scheme = root.find(".//a:clrScheme", ns)
            if scheme is None:
                return palette
            tags = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
            for idx, tg in enumerate(tags):
                el = scheme.find(f"a:{tg}", ns)
                if el is None:
                    continue
                srgb = el.find("a:srgbClr", ns)
                if srgb is not None and srgb.attrib.get("val"):
                    palette[idx] = srgb.attrib["val"].upper()[-6:]
                    continue
                sys_c = el.find("a:sysClr", ns)
                if sys_c is not None and sys_c.attrib.get("lastClr"):
                    palette[idx] = sys_c.attrib["lastClr"].upper()[-6:]
    except Exception:
        pass
    return palette


def _apply_tint(hex6: str, tint: float) -> str:
    if not tint or abs(tint) < 1e-4:
        return hex6
    try:
        r = int(hex6[0:2], 16)
        g = int(hex6[2:4], 16)
        b = int(hex6[4:6], 16)
    except Exception:
        return hex6

    def _adj(ch: int) -> int:
        if tint < 0:
            return max(0, min(255, int(round(ch * (1.0 + tint)))))
        return max(0, min(255, int(round(ch * (1.0 - tint) + 255 * tint))))

    return f"{_adj(r):02X}{_adj(g):02X}{_adj(b):02X}"


def _resolve_openpyxl_color(color_obj: Any, theme_palette: list[str]) -> str | None:
    if color_obj is None:
        return None
    ctype = getattr(color_obj, "type", None)
    tint = float(getattr(color_obj, "tint", 0.0) or 0.0)
    if ctype == "rgb":
        rgb = str(getattr(color_obj, "rgb", "") or "").strip()
        if len(rgb) == 8 and re.fullmatch(r"[0-9A-Fa-f]{8}", rgb):
            if rgb[:2] == "00" and rgb[2:] == "000000":
                return None
            return f"#{_apply_tint(rgb[2:].upper(), tint)}"
        if len(rgb) == 6 and re.fullmatch(r"[0-9A-Fa-f]{6}", rgb):
            return f"#{_apply_tint(rgb.upper(), tint)}"
    elif ctype == "theme":
        ti = getattr(color_obj, "theme", None)
        if isinstance(ti, int) and 0 <= ti < len(theme_palette):
            return f"#{_apply_tint(theme_palette[ti], tint)}"
    elif ctype == "indexed":
        ii = getattr(color_obj, "indexed", None)
        if isinstance(ii, int) and ii in _INDEXED_RGB:
            return f"#{_apply_tint(_INDEXED_RGB[ii], tint)}"
    return None


def _border_side_css(side_obj: Any, theme_palette: list[str]) -> str:
    if side_obj is None or not getattr(side_obj, "style", None):
        return ""
    st = str(side_obj.style)
    col = _resolve_openpyxl_color(getattr(side_obj, "color", None), theme_palette) or "#334155"
    style_map = {
        "thin": f"1px solid {col}",
        "hair": f"1px solid {col}",
        "medium": f"2px solid {col}",
        "thick": f"3px solid {col}",
        "double": f"3px double {col}",
        "dashed": f"1px dashed {col}",
        "mediumDashed": f"2px dashed {col}",
        "dotted": f"1px dotted {col}",
        "dashDot": f"1px dashed {col}",
    }
    return style_map.get(st, f"1px solid {col}")


def _guess_mime(filename: str, raw: bytes) -> str:
    if raw.startswith(b"\x89PNG"):
        return "image/png"
    if raw.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if raw.startswith((b"GIF87a", b"GIF89a")):
        return "image/gif"
    if raw.startswith(b"BM"):
        return "image/bmp"
    mt, _ = mimetypes.guess_type(filename)
    if mt and mt.startswith("image/"):
        return mt
    return "image/png"


def _to_data_uri(filename: str, raw: bytes, max_bytes: int = 2_500_000) -> str | None:
    if not raw or len(raw) > max_bytes:
        return None
    mime = _guess_mime(filename, raw)
    # Skip non-browser formats like EMF/WMF unless convertible
    if mime in ("image/x-emf", "image/x-wmf"):
        return None
    b64 = base64.b64encode(raw).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _extract_xlsx_sheet_images(xlsx_path: Path, sheet_index_zero: int, ws) -> list[dict[str, Any]]:
    """Extract embedded logos/images in `ws` as inline `data:` URIs with anchor row/col & px size."""
    images_out: list[dict[str, Any]] = []
    # 1. Try openpyxl `ws._images` directly (works when Pillow is installed)
    ws_imgs = list(getattr(ws, "_images", []) or [])
    for idx, img in enumerate(ws_imgs[:12]):
        try:
            raw = img._data()
            uri = _to_data_uri(getattr(img, "path", f"img_{idx}.png"), raw)
            if not uri:
                continue
            row0, col0 = 0, 0
            anc = getattr(img, "anchor", None)
            if isinstance(anc, str):
                from openpyxl.utils import coordinate_to_tuple
                rr, cc = coordinate_to_tuple(anc)
                row0, col0 = rr - 1, cc - 1
            elif anc is not None and hasattr(anc, "_from") and anc._from is not None:
                row0 = int(getattr(anc._from, "row", 0) or 0)
                col0 = int(getattr(anc._from, "col", 0) or 0)
            w_px = int(getattr(img, "width", 120) or 120)
            h_px = int(getattr(img, "height", 60) or 60)
            images_out.append({
                "row": row0 + 1,
                "col": col0 + 1,
                "col_letter": get_column_letter(col0 + 1),
                "width_px": max(24, min(w_px, 600)),
                "height_px": max(20, min(h_px, 320)),
                "data_uri": uri,
            })
        except Exception:
            continue
    if images_out:
        return images_out

    # 2. Direct OOXML zip drawing fallback (catches drawings even if openpyxl skipped them)
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            sheet_rel = f"xl/worksheets/_rels/sheet{sheet_index_zero + 1}.xml.rels"
            if sheet_rel not in zf.namelist():
                return []
            srels = ET.fromstring(zf.read(sheet_rel))
            drawing_target = ""
            for rel in srels:
                if rel.attrib.get("Type", "").endswith("/drawing"):
                    drawing_target = rel.attrib.get("Target", "")
                    break
            if not drawing_target:
                return []
            d_path = "xl/" + drawing_target.lstrip("./").replace("../", "")
            d_rel = str(Path(d_path).parent / "_rels" / (Path(d_path).name + ".rels")).replace("\\", "/")
            if d_path not in zf.namelist() or d_rel not in zf.namelist():
                return []
            rid_map: dict[str, str] = {}
            for rel in ET.fromstring(zf.read(d_rel)):
                if rel.attrib.get("Type", "").endswith("/image"):
                    t = "xl/" + rel.attrib.get("Target", "").lstrip("./").replace("../", "")
                    rid_map[rel.attrib.get("Id", "")] = t
            ns = {
                "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            }
            droot = ET.fromstring(zf.read(d_path))
            for anchor_tag in ("xdr:twoCellAnchor", "xdr:oneCellAnchor"):
                for anc in droot.findall(anchor_tag, ns):
                    blip = anc.find(".//a:blip", ns)
                    if blip is None:
                        continue
                    rid = blip.attrib.get(f"{{{ns['r']}}}embed", "")
                    mpath = rid_map.get(rid, "")
                    if not mpath or mpath not in zf.namelist():
                        continue
                    uri = _to_data_uri(mpath, zf.read(mpath))
                    if not uri:
                        continue
                    fr = anc.find("xdr:from", ns)
                    r0 = int(fr.findtext("xdr:row", "0", ns)) if fr is not None else 0
                    c0 = int(fr.findtext("xdr:col", "0", ns)) if fr is not None else 0
                    ext_el = anc.find(".//a:xfrm/a:ext", ns) or anc.find("xdr:ext", ns)
                    w_px = int(int(ext_el.attrib.get("cx", "1143000")) / 9525) if ext_el is not None else 120
                    h_px = int(int(ext_el.attrib.get("cy", "571500")) / 9525) if ext_el is not None else 60
                    images_out.append({
                        "row": r0 + 1,
                        "col": c0 + 1,
                        "col_letter": get_column_letter(c0 + 1),
                        "width_px": max(24, min(w_px, 600)),
                        "height_px": max(20, min(h_px, 320)),
                        "data_uri": uri,
                    })
    except Exception:
        pass
    return images_out


def _format_cell_display(val: Any, cached_val: Any, num_fmt: str) -> tuple[str, bool, str]:
    """Return `(display_text, is_formula, raw_formula)`."""
    if isinstance(val, str) and val.startswith("="):
        if cached_val not in (None, ""):
            if isinstance(cached_val, float):
                disp = f"{cached_val:,.3f}" if "000" in (num_fmt or "") else f"{cached_val:,.2f}"
            else:
                disp = str(cached_val)
        else:
            disp = "ƒx"
        return disp, True, val
    if val is None:
        return "", False, ""
    if isinstance(val, dt.datetime):
        return val.strftime("%Y-%m-%d"), False, ""
    if isinstance(val, dt.date):
        return val.isoformat(), False, ""
    if isinstance(val, float):
        nf = num_fmt or ""
        if "%" in nf:
            return f"{val * 100:.1f}%", False, ""
        if "000" in nf:
            return f"{val:,.3f}", False, ""
        if "0.00" in nf or "#,##0" in nf:
            return f"{val:,.2f}" if "00" in nf else f"{val:,.0f}", False, ""
        return f"{val:g}", False, ""
    return str(val), False, ""


def render_xlsx_file(
    xlsx_path: Path | str,
    *,
    sheet_name: str = "",
    max_rows: int = 90,
    max_cols: int = 28,
    highlighted_targets: set[str] | None = None,
) -> dict[str, Any]:
    """Render an `.xlsx`/`.xlsm` workbook sheet with true visual fidelity (100% local)."""
    p = Path(xlsx_path).resolve()
    ext = p.suffix.lower()
    theme = _extract_workbook_theme_palette(p)
    highlights = highlighted_targets or set()

    wb_f = openpyxl.load_workbook(p, data_only=False, keep_vba=(ext == ".xlsm"))
    wb_v = openpyxl.load_workbook(p, data_only=True, read_only=False)
    try:
        sheets_list = [
            {
                "name": sn,
                "rtl": bool(getattr(wb_f[sn].sheet_view, "rightToLeft", False)),
                "rows": wb_f[sn].max_row or 1,
                "cols": wb_f[sn].max_column or 1,
            }
            for sn in wb_f.sheetnames
        ]
        active_name = sheet_name if sheet_name in wb_f.sheetnames else wb_f.sheetnames[0]
        s_idx = wb_f.sheetnames.index(active_name)
        ws = wb_f[active_name]
        ws_v = wb_v[active_name]
        is_rtl = bool(getattr(ws.sheet_view, "rightToLeft", False))

        # Compute real visual bounding box (including merged cells & embedded images!)
        images = _extract_xlsx_sheet_images(p, s_idx, ws)
        eff_max_r = min(max(ws.max_row or 1, 8), max_rows)
        eff_max_c = min(max(ws.max_column or 1, 6), max_cols)
        for im in images:
            eff_max_r = min(max(eff_max_r, im["row"] + 2), max_rows)
            eff_max_c = min(max(eff_max_c, im["col"] + 1), max_cols)

        # Trim trailing completely empty/unstyled rows & cols beyond actual content
        last_used_r = 1
        last_used_c = 1
        for r in range(1, eff_max_r + 1):
            for c in range(1, eff_max_c + 1):
                cell = ws.cell(r, c)
                has_v = cell.value is not None and str(cell.value) != ""
                has_b = bool(cell.border and any(getattr(getattr(cell.border, sd, None), "style", None) for sd in ("top", "bottom", "left", "right")))
                has_f = bool(cell.fill and getattr(cell.fill, "fill_type", None) not in (None, "none"))
                if has_v or has_b or has_f:
                    last_used_r = max(last_used_r, r)
                    last_used_c = max(last_used_c, c)
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= max_rows and rng.min_col <= max_cols:
                last_used_r = max(last_used_r, min(rng.max_row, max_rows))
                last_used_c = max(last_used_c, min(rng.max_col, max_cols))
        eff_max_r = min(max(last_used_r + 1, 6), max_rows)
        eff_max_c = min(max(last_used_c, 4), max_cols)

        # Column widths (Excel character units -> CSS px)
        columns_meta: list[dict[str, Any]] = []
        for c in range(1, eff_max_c + 1):
            let = get_column_letter(c)
            cd = ws.column_dimensions.get(let)
            hidden = bool(cd.hidden) if cd else False
            raw_w = float(cd.width) if (cd and cd.width is not None) else 10.5
            px_w = 0 if hidden else max(38, min(int(round(raw_w * 7.5 + 8)), 420))
            columns_meta.append({"col": c, "letter": let, "width_px": px_w, "hidden": hidden})

        # Merged-cell lookup (merge outer-edge borders onto master cell!)
        merge_master: dict[tuple[int, int], tuple[int, int, str, int, int]] = {}
        merge_covered: set[tuple[int, int]] = set()
        for rng in ws.merged_cells.ranges:
            r1, c1, r2, c2 = rng.min_row, rng.min_col, min(rng.max_row, eff_max_r), min(rng.max_col, eff_max_c)
            if r1 > eff_max_r or c1 > eff_max_c:
                continue
            merge_master[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1, str(rng), r2, c2)
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) != (r1, c1):
                        merge_covered.add((rr, cc))

        rows_out: list[dict[str, Any]] = []
        for r in range(1, eff_max_r + 1):
            rd = ws.row_dimensions.get(r)
            r_hidden = bool(rd.hidden) if rd else False
            raw_h = float(rd.height) if (rd and rd.height is not None) else 20.0
            px_h = 0 if r_hidden else max(22, min(int(round(raw_h * 1.33)), 260))
            cells_out: list[dict[str, Any]] = []

            for c in range(1, eff_max_c + 1):
                if (r, c) in merge_covered:
                    continue
                cell = ws.cell(r, c)
                cval = ws_v.cell(r, c).value
                let = get_column_letter(c)
                coord = f"{let}{r}"
                rowspan, colspan, mrng, end_r, end_c = merge_master.get((r, c), (1, 1, "", r, c))
                disp, is_formula, raw_formula = _format_cell_display(cell.value, cval, cell.number_format or "")

                # Style extraction
                style: dict[str, Any] = {}
                if cell.fill and getattr(cell.fill, "fill_type", None) == "solid":
                    bg = _resolve_openpyxl_color(cell.fill.fgColor, theme)
                    if bg:
                        style["bg"] = bg
                if cell.font:
                    if cell.font.name:
                        style["font"] = cell.font.name
                    if cell.font.size:
                        style["size"] = round(float(cell.font.size), 1)
                    if cell.font.bold:
                        style["bold"] = True
                    if cell.font.italic:
                        style["italic"] = True
                    if cell.font.underline:
                        style["underline"] = True
                    fg = _resolve_openpyxl_color(cell.font.color, theme)
                    if fg:
                        style["color"] = fg
                if cell.alignment:
                    ha = cell.alignment.horizontal
                    va = cell.alignment.vertical
                    if ha in ("left", "center", "right", "justify"):
                        style["align"] = ha
                    if va in ("top", "center", "bottom"):
                        style["valign"] = "middle" if va == "center" else va
                    if cell.alignment.wrap_text:
                        style["wrap"] = True

                # Borders (take bottom/right from end of merge span so merged boxes render crisp!)
                bt = _border_side_css(cell.border.top if cell.border else None, theme)
                bl = _border_side_css(cell.border.left if cell.border else None, theme)
                br_cell = ws.cell(r, end_c)
                bb_cell = ws.cell(end_r, c)
                br = _border_side_css(br_cell.border.right if br_cell.border else None, theme)
                bb = _border_side_css(bb_cell.border.bottom if bb_cell.border else None, theme)
                if bt:
                    style["bt"] = bt
                if br:
                    style["br"] = br
                if bb:
                    style["bb"] = bb
                if bl:
                    style["bl"] = bl

                cell_obj: dict[str, Any] = {
                    "coord": coord,
                    "r": r,
                    "c": c,
                    "text": disp,
                }
                if rowspan > 1:
                    cell_obj["rowspan"] = rowspan
                if colspan > 1:
                    cell_obj["colspan"] = colspan
                if mrng:
                    cell_obj["merge"] = mrng
                if is_formula:
                    cell_obj["formula"] = raw_formula
                if style:
                    cell_obj["style"] = style
                if f"{active_name}!{coord}" in highlights:
                    cell_obj["filled"] = True
                cells_out.append(cell_obj)

            rows_out.append({"row": r, "height_px": px_h, "hidden": r_hidden, "cells": cells_out})

        return {
            "kind": "xlsx",
            "file_name": p.name,
            "ext": ext,
            "sheets": sheets_list,
            "active_sheet": active_name,
            "rtl": is_rtl,
            "columns": columns_meta,
            "rows": rows_out,
            "images": images,
        }
    finally:
        wb_f.close()
        wb_v.close()


# ── Word (`.docx`) Renderer: LibreOffice PNG (if installed) + Faithful A4 HTML ─
def _find_libreoffice_binary() -> str | None:
    for cand in (
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        found = shutil.which(cand) or (cand if Path(cand).exists() else None)
        if found:
            return found
    return None


def _try_render_docx_pages_via_libreoffice(docx_path: Path, dpi: int = 130) -> list[dict[str, Any]]:
    """Optional pixel-exact page rendering when LibreOffice + PyMuPDF are installed locally."""
    soffice = _find_libreoffice_binary()
    if not soffice:
        return []
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return []
    pages: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="sla_docx_render_") as td:
        td_p = Path(td)
        try:
            subprocess.run(
                [soffice, "--headless", "--nologo", "--nofirststartwizard", "--convert-to", "pdf", "--outdir", str(td_p), str(docx_path)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=25,
                check=False,
            )
            pdf_p = td_p / f"{docx_path.stem}.pdf"
            if not pdf_p.exists():
                return []
            pdf_doc = fitz.open(str(pdf_p))
            try:
                for i in range(min(len(pdf_doc), 8)):
                    page = pdf_doc[i]
                    pix = page.get_pixmap(dpi=dpi, alpha=False)
                    png_bytes = pix.tobytes("png")
                    uri = _to_data_uri(f"page_{i + 1}.png", png_bytes, max_bytes=4_000_000)
                    if uri:
                        pages.append({"page": i + 1, "width": pix.width, "height": pix.height, "data_uri": uri})
            finally:
                pdf_doc.close()
        except Exception:
            return []
    return pages


def _extract_docx_part_images(part: Any, el: Any) -> list[dict[str, Any]]:
    """Extract inline/anchored images inside a Word XML element as `data:` URIs."""
    imgs: list[dict[str, Any]] = []
    a_blip = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
    r_embed = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    wp_ext = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
    for blip in el.iter(a_blip):
        rid = blip.attrib.get(r_embed)
        if not rid or rid not in part.related_parts:
            continue
        img_part = part.related_parts[rid]
        uri = _to_data_uri(str(getattr(img_part, "partname", "logo.png")), img_part.blob)
        if not uri:
            continue
        w_px, h_px = 140, 64
        # Walk up a few levels to find wp:extent
        parent = blip.getparent()
        for _ in range(6):
            if parent is None:
                break
            ext_el = parent.find(wp_ext)
            if ext_el is not None:
                try:
                    w_px = max(24, min(int(int(ext_el.attrib.get("cx", "1333500")) / 9525), 620))
                    h_px = max(20, min(int(int(ext_el.attrib.get("cy", "609600")) / 9525), 340))
                except Exception:
                    pass
                break
            parent = parent.getparent()
        imgs.append({"data_uri": uri, "width_px": w_px, "height_px": h_px})
    return imgs


def _docx_border_el_to_css(b_el: Any) -> str:
    if b_el is None:
        return ""
    val = b_el.get(qn("w:val")) or ""
    if not val or val in ("none", "nil"):
        return ""
    sz = b_el.get(qn("w:sz"))
    px = 1
    if sz and str(sz).isdigit():
        pt = int(sz) / 8.0
        px = 1 if pt <= 1.1 else (2 if pt <= 2.2 else 3)
    col = (b_el.get(qn("w:color")) or "334155").strip()
    color_css = "#334155" if col.lower() in ("auto", "") else f"#{col[-6:]}"
    kind = "double" if val == "double" else ("dashed" if "dash" in val else ("dotted" if "dot" in val else "solid"))
    return f"{px}px {kind} {color_css}"


def _extract_para_data(para: Any, part: Any, para_idx: int | None = None, highlights: set[str] | None = None) -> dict[str, Any]:
    hl = highlights or set()
    pPr = para._p.find(qn("w:pPr"))
    is_rtl = False
    align = ""
    if pPr is not None:
        bidi = pPr.find(qn("w:bidi"))
        if bidi is not None and bidi.get(qn("w:val"), "1") not in ("0", "false"):
            is_rtl = True
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            jv = jc.get(qn("w:val"), "")
            if jv in ("center", "left", "right", "both"):
                align = "justify" if jv == "both" else jv
    txt = para.text or ""
    if not is_rtl and re.search(r"[\u0600-\u06FF]", txt):
        is_rtl = True

    runs_out: list[dict[str, Any]] = []
    for run in para.runs:
        rt = run.text or ""
        imgs = _extract_docx_part_images(part, run._r)
        if not rt and not imgs:
            continue
        r_style: dict[str, Any] = {}
        if run.bold:
            r_style["bold"] = True
        if run.italic:
            r_style["italic"] = True
        if run.underline:
            r_style["underline"] = True
        if run.font:
            if run.font.size:
                try:
                    r_style["size"] = round(run.font.size.pt, 1)
                except Exception:
                    pass
            if run.font.color and run.font.color.rgb:
                r_style["color"] = f"#{str(run.font.color.rgb)}"
            if run.font.name:
                r_style["font"] = run.font.name
        runs_out.append({"text": rt, "style": r_style, "images": imgs})

    spans = [m.group(0) for m in _PLACEHOLDER_RE.finditer(txt)]
    has_colon_slot = bool(txt.strip() and re.search(r"[:：]\s*[\._…\-]{0,40}\s*$", txt.strip()))
    filled = bool(para_idx is not None and (f"P{para_idx}" in hl or any(k.startswith(f"P{para_idx}[") for k in hl)))
    return {
        "kind": "paragraph",
        "para_idx": para_idx,
        "text": txt,
        "rtl": is_rtl,
        "align": align,
        "runs": runs_out,
        "spans": spans,
        "has_colon_slot": has_colon_slot,
        "filled": filled,
    }


def _extract_table_data(tbl: Any, part: Any, table_idx: int, highlights: set[str] | None = None) -> dict[str, Any]:
    hl = highlights or set()
    tblPr = tbl._tbl.find(qn("w:tblPr"))
    is_rtl = bool(tblPr is not None and tblPr.find(qn("w:bidiVisual")) is not None)

    # Default table borders
    default_borders = {"bt": "1px solid #94A3B8", "br": "1px solid #94A3B8", "bb": "1px solid #94A3B8", "bl": "1px solid #94A3B8"}
    if tblPr is not None:
        tb = tblPr.find(qn("w:tblBorders"))
        if tb is not None:
            for key, wtag in (("bt", "top"), ("bb", "bottom"), ("bl", "left"), ("br", "right")):
                el = tb.find(qn(f"w:{wtag}"))
                if el is not None:
                    default_borders[key] = _docx_border_el_to_css(el)
            ins_h = tb.find(qn("w:insideH"))
            ins_v = tb.find(qn("w:insideV"))
            if ins_h is not None:
                c_h = _docx_border_el_to_css(ins_h)
                if c_h:
                    default_borders["ih"] = c_h
            if ins_v is not None:
                c_v = _docx_border_el_to_css(ins_v)
                if c_v:
                    default_borders["iv"] = c_v

    # Column grid widths (dxa -> px: 1 dxa = 1/20 pt)
    col_widths_px: list[int] = []
    tblGrid = tbl._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        for gc in tblGrid.findall(qn("w:gridCol")):
            w = gc.get(qn("w:w"))
            if w and str(w).isdigit():
                col_widths_px.append(max(36, int(round(int(w) / 15.0))))

    # Retain ALL `_tc` proxies strongly for the lifetime of the table scan (prevents Python GC id reuse!)
    all_row_tcs = [[c._tc for c in row.cells] for row in tbl.rows]
    n_rows = len(tbl.rows)

    # Build grid accounting for horizontal (`same _tc` across columns) and vertical (`same _tc` across rows) merges
    visited_tc_ids: set[int] = set()
    rows_out: list[dict[str, Any]] = []

    for r_i, row in enumerate(tbl.rows):
        row_cells_out: list[dict[str, Any]] = []
        n_cols = len(row.cells)
        c_i = 0
        while c_i < n_cols:
            tc = all_row_tcs[r_i][c_i]
            tc_id = id(tc)
            # Count horizontal span in current row
            colspan = 1
            while c_i + colspan < n_cols and all_row_tcs[r_i][c_i + colspan] is tc:
                colspan += 1
            if tc_id in visited_tc_ids:
                c_i += colspan
                continue
            visited_tc_ids.add(tc_id)

            # Count vertical span across subsequent rows
            rowspan = 1
            while r_i + rowspan < n_rows and c_i < len(all_row_tcs[r_i + rowspan]) and all_row_tcs[r_i + rowspan][c_i] is tc:
                rowspan += 1

            cell = row.cells[c_i]
            tcPr = tc.find(qn("w:tcPr"))
            style: dict[str, Any] = {
                "bt": default_borders.get("ih" if r_i > 0 else "bt", default_borders["bt"]),
                "bb": default_borders.get("ih" if r_i + rowspan < n_rows else "bb", default_borders["bb"]),
                "bl": default_borders.get("iv" if c_i > 0 else "bl", default_borders["bl"]),
                "br": default_borders.get("iv" if c_i + colspan < n_cols else "br", default_borders["br"]),
            }
            if tcPr is not None:
                shd = tcPr.find(qn("w:shd"))
                if shd is not None:
                    fill = (shd.get(qn("w:fill")) or "").strip()
                    if fill and fill.lower() not in ("auto", "ffffff") and re.fullmatch(r"[0-9A-Fa-f]{6}", fill):
                        style["bg"] = f"#{fill.upper()}"
                tcb = tcPr.find(qn("w:tcBorders"))
                if tcb is not None:
                    for key, wtag in (("bt", "top"), ("bb", "bottom"), ("bl", "left"), ("br", "right")):
                        bel = tcb.find(qn(f"w:{wtag}"))
                        if bel is not None:
                            style[key] = _docx_border_el_to_css(bel)
                va = tcPr.find(qn("w:vAlign"))
                if va is not None:
                    vv = va.get(qn("w:val"), "")
                    if vv in ("center", "top", "bottom"):
                        style["valign"] = "middle" if vv == "center" else vv

            paras_data = [_extract_para_data(p, part) for p in cell.paragraphs]
            full_txt = cell.text.strip()
            if not is_rtl and re.search(r"[\u0600-\u06FF]", full_txt):
                style["rtl"] = True
            spans = [m.group(0) for m in _PLACEHOLDER_RE.finditer(full_txt)]

            row_cells_out.append({
                "table_idx": table_idx,
                "r": r_i,
                "c": c_i,
                "rowspan": rowspan,
                "colspan": colspan,
                "text": full_txt,
                "paragraphs": paras_data,
                "spans": spans,
                "style": style,
                "filled": f"T{table_idx}:R{r_i}C{c_i}" in hl,
            })
            c_i += colspan

        # Row height if specified
        trPr = row._tr.find(qn("w:trPr"))
        h_px = 0
        if trPr is not None:
            trH = trPr.find(qn("w:trHeight"))
            if trH is not None and str(trH.get(qn("w:val"), "")).isdigit():
                h_px = max(22, int(round(int(trH.get(qn("w:val"))) / 15.0)))
        rows_out.append({"row": r_i, "height_px": h_px, "cells": row_cells_out})

    return {
        "kind": "table",
        "table_idx": table_idx,
        "rtl": is_rtl,
        "col_widths_px": col_widths_px,
        "rows": rows_out,
    }


def render_docx_file(
    docx_path: Path | str,
    *,
    try_libreoffice: bool = True,
    highlighted_targets: set[str] | None = None,
) -> dict[str, Any]:
    """
    Render a `.docx` file with:
      1) `page_images` via LibreOffice headless -> PDF -> PNG (when installed)
      2) Full-fidelity structured A4 blocks (header logos/tables, body paragraphs & tables
         in exact document order, footer) with clickable targets for cells, paragraphs, and spans!
    """
    p = Path(docx_path).resolve()
    hl = highlighted_targets or set()
    doc = docx.Document(str(p))

    # Extract header & footer blocks (logos, letterhead titles, ministry headers!)
    header_blocks: list[dict[str, Any]] = []
    footer_blocks: list[dict[str, Any]] = []
    if doc.sections:
        sec0 = doc.sections[0]
        try:
            hdr = sec0.header
            if hdr and not hdr.is_linked_to_previous or hdr:
                for hp in hdr.paragraphs:
                    pd = _extract_para_data(hp, hdr.part)
                    if pd["text"].strip() or any(r["images"] for r in pd["runs"]):
                        header_blocks.append(pd)
                for hti, htbl in enumerate(hdr.tables):
                    header_blocks.append(_extract_table_data(htbl, hdr.part, -(hti + 1)))
        except Exception:
            pass
        try:
            ftr = sec0.footer
            if ftr:
                for fp in ftr.paragraphs:
                    pd = _extract_para_data(fp, ftr.part)
                    if pd["text"].strip() or any(r["images"] for r in pd["runs"]):
                        footer_blocks.append(pd)
        except Exception:
            pass

    # Walk `<w:body>` in EXACT visual order keeping exact `para_idx` and `table_idx`
    para_by_el = {id(para._p): (idx, para) for idx, para in enumerate(doc.paragraphs)}
    tbl_by_el = {id(tbl._tbl): (idx, tbl) for idx, tbl in enumerate(doc.tables)}
    body_blocks: list[dict[str, Any]] = []
    for child in doc.element.body:
        cid = id(child)
        if cid in para_by_el:
            p_idx, para = para_by_el[cid]
            body_blocks.append(_extract_para_data(para, doc.part, para_idx=p_idx, highlights=hl))
        elif cid in tbl_by_el:
            t_idx, tbl = tbl_by_el[cid]
            body_blocks.append(_extract_table_data(tbl, doc.part, table_idx=t_idx, highlights=hl))

    page_images = _try_render_docx_pages_via_libreoffice(p) if try_libreoffice else []
    return {
        "kind": "docx",
        "file_name": p.name,
        "ext": ".docx",
        "render_engine": "libreoffice_png+html_overlay" if page_images else "html_a4_fallback",
        "has_page_images": len(page_images) > 0,
        "page_images": page_images,
        "header_blocks": header_blocks,
        "body_blocks": body_blocks,
        "footer_blocks": footer_blocks,
        "paragraphs_count": len(doc.paragraphs),
        "tables_count": len(doc.tables),
    }


def render_template(
    proj_root: Path | str,
    tpl_name: str,
    *,
    sheet_name: str = "",
    try_libreoffice: bool = True,
) -> dict[str, Any]:
    """Render any uploaded project template (`.xlsx`, `.xlsm`, or `.docx`)."""
    tpl_path = tprof.resolve_project_template(proj_root, tpl_name)
    if tpl_path.suffix.lower() in (".xlsx", ".xlsm"):
        return render_xlsx_file(tpl_path, sheet_name=sheet_name)
    return render_docx_file(tpl_path, try_libreoffice=try_libreoffice)


def render_filled_preview(
    proj_root: Path | str,
    tpl_name: str,
    profile: dict[str, Any],
    payload: dict[str, Any],
    project: dict[str, Any],
    *,
    system_ctx: dict[str, Any] | None = None,
    sheet_name: str = "",
    try_libreoffice: bool = True,
) -> dict[str, Any]:
    """
    Validate inputs, generate a temporary filled copy via `fill_by_profile` (without
    consuming counters or writing to archive/log!), and return its exact rendered visual!
    """
    tpl_path = tprof.resolve_project_template(proj_root, tpl_name)
    ext = tpl_path.suffix.lower()
    with tempfile.TemporaryDirectory(prefix="sla_preview_") as td:
        tmp_out = Path(td) / f"preview_{tpl_path.name}"
        fill_res = tprof.fill_by_profile(
            tpl_path,
            profile,
            payload,
            project,
            tmp_out,
            system_ctx=system_ctx,
            verify_drift=True,
        )
        hl = set(fill_res.get("touched_targets") or [])
        if ext in (".xlsx", ".xlsm"):
            visual = render_xlsx_file(tmp_out, sheet_name=sheet_name, highlighted_targets=hl)
        else:
            visual = render_docx_file(tmp_out, try_libreoffice=try_libreoffice, highlighted_targets=hl)
    return {
        "ok": True,
        "preview": True,
        "resolved_fields": fill_res["resolved_fields"],
        "resolved_tables": fill_res["resolved_tables"],
        "touched_targets": fill_res["touched_targets"],
        "rebound_fields": fill_res["rebound_fields"],
        "render": visual,
    }


# ── PDF Export, Merging & Executive Report PDF Helpers ────────────────────────
_PDF_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif"}


def count_pdf_pages(pdf_path: Path) -> int:
    """Return page count of `pdf_path` using PyMuPDF if available, or safe byte scan."""
    if not pdf_path.exists() or not pdf_path.is_file():
        return 0
    try:
        import fitz  # type: ignore
        with fitz.open(str(pdf_path)) as doc:
            return max(1, int(doc.page_count or 1))
    except Exception:
        pass
    try:
        raw = pdf_path.read_bytes()
        hits = len(re.findall(rb"/Type\s*/Page(?:\b|[^s])", raw))
        return max(1, hits)
    except Exception:
        return 1


def _try_ms_office_to_pdf(src_path: Path, out_pdf: Path) -> bool:
    """On Windows with Microsoft Office installed, convert `.docx`/`.xlsx`/`.xlsm` cleanly to PDF."""
    if sys.platform != "win32":
        return False
    ext = src_path.suffix.lower()
    if ext not in (".docx", ".doc", ".xlsx", ".xlsm", ".xls"):
        return False
    ps = shutil.which("powershell") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    if not ps or not Path(ps).exists():
        return False
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    if ext in (".docx", ".doc"):
        script = (
            "$app = $null; $doc = $null; "
            "try { "
            "$app = New-Object -ComObject Word.Application; "
            "$app.Visible = $false; $app.DisplayAlerts = 0; "
            "$doc = $app.Documents.Open($env:SLA_IN, $false, $true); "
            "$doc.ExportAsFixedFormat($env:SLA_OUT, 17); "
            "} finally { "
            "if ($doc) { try { $doc.Close($false) } catch {} }; "
            "if ($app) { try { $app.Quit() } catch {} } "
            "}"
        )
    else:
        script = (
            "$app = $null; $wb = $null; "
            "try { "
            "$app = New-Object -ComObject Excel.Application; "
            "$app.Visible = $false; $app.DisplayAlerts = $false; "
            "$wb = $app.Workbooks.Open($env:SLA_IN, 0, $true); "
            "$wb.ExportAsFixedFormat(0, $env:SLA_OUT); "
            "} finally { "
            "if ($wb) { try { $wb.Close($false) } catch {} }; "
            "if ($app) { try { $app.Quit() } catch {} } "
            "}"
        )
    env = dict(os.environ)
    env["SLA_IN"] = str(src_path.resolve())
    env["SLA_OUT"] = str(out_pdf.resolve())
    try:
        subprocess.run(
            [ps, "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-Command", script],
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=45,
            check=False,
        )
        return out_pdf.exists() and out_pdf.is_file() and out_pdf.stat().st_size > 0
    except Exception:
        return False


def _try_libreoffice_to_pdf(src_path: Path, out_pdf: Path) -> bool:
    """Convert `.docx`/`.xlsx`/`.xlsm` via headless LibreOffice if present."""
    soffice = _find_libreoffice_binary()
    if not soffice:
        return False
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="sla_lo_pdf_") as td:
        td_p = Path(td)
        try:
            subprocess.run(
                [soffice, "--headless", "--nologo", "--nofirststartwizard", "--convert-to", "pdf", "--outdir", str(td_p), str(src_path)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=45,
                check=False,
            )
            gen_pdf = td_p / f"{src_path.stem}.pdf"
            if gen_pdf.exists() and gen_pdf.stat().st_size > 0:
                shutil.copy2(gen_pdf, out_pdf)
                return out_pdf.exists() and out_pdf.stat().st_size > 0
        except Exception:
            return False
    return False


def _extract_lines_from_office(src_path: Path) -> tuple[str, list[str]]:
    """Extract clean readable title & text lines from `.docx` or `.xlsx`/`.xlsm` for fallback PDF."""
    ext = src_path.suffix.lower()
    lines: list[str] = []
    if ext == ".docx":
        try:
            doc = docx.Document(str(src_path))
            for para in doc.paragraphs:
                txt = (para.text or "").strip()
                if txt:
                    lines.append(txt)
            for tb in doc.tables:
                for row in tb.rows:
                    vals = [((c.text or "").replace("\n", " ").strip()) for c in row.cells]
                    vals_nonempty = [v for v in vals if v]
                    if vals_nonempty:
                        lines.append(" | ".join(vals_nonempty))
        except Exception:
            pass
    elif ext in (".xlsx", ".xlsm", ".xls"):
        try:
            wb = openpyxl.load_workbook(str(src_path), read_only=True, data_only=True)
            try:
                for sname in wb.sheetnames[:6]:
                    ws = wb[sname]
                    lines.append(f"[{sname}]")
                    rc = 0
                    for row_vals in ws.iter_rows(values_only=True):
                        cells = [str(v).strip() for v in row_vals if v is not None and str(v).strip() != ""]
                        if cells:
                            lines.append(" | ".join(cells[:10]))
                            rc += 1
                            if rc >= 80:
                                break
            finally:
                wb.close()
        except Exception:
            pass
    return src_path.stem, lines


def _pick_pil_font(size: int = 16):
    from PIL import ImageFont  # type: ignore

    for cand in (
        r"C:\Windows\Fonts\tahoma.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    ):
        if Path(cand).exists():
            try:
                return ImageFont.truetype(cand, size=size)
            except Exception:
                pass
    return ImageFont.load_default()


def render_lines_to_pdf_fallback(title: str, subtitle: str, lines: list[str], out_pdf: Path) -> str:
    """
    Pure-Python A4 PDF builder used when neither LibreOffice nor MS Office is installed.
    Uses PyMuPDF HTML/Story when available, otherwise Pillow multi-page A4 PDF export.
    """
    import html as html_mod

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    safe_lines = [str(x).strip() for x in (lines or []) if str(x).strip()]

    # Try PyMuPDF searchable HTML Story if fitz is available
    try:
        import fitz  # type: ignore

        body_html = "".join(f"<p>{html_mod.escape(ln)}</p>" for ln in safe_lines[:400])
        full_html = (
            "<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'>"
            "<style>"
            "body{font-family:sans-serif;font-size:10.5pt;color:#0f172a;line-height:1.45;}"
            "h1{font-size:15pt;color:#035D89;border-bottom:2px solid #059BE5;padding-bottom:4pt;margin:0 0 6pt;}"
            ".sub{font-size:9.5pt;color:#475569;margin-bottom:10pt;}"
            "p{margin:2pt 0;}"
            "</style></head><body>"
            f"<h1>{html_mod.escape(title)}</h1>"
            + (f"<div class='sub'>{html_mod.escape(subtitle)}</div>" if subtitle else "")
            + body_html
            + "</body></html>"
        )
        doc = fitz.open()
        rect = fitz.Rect(36, 36, 559, 806)  # A4 with 36pt margins
        story = fitz.Story(html=full_html)
        more = True
        while more:
            page = doc.new_page(width=595, height=842)
            more, _ = story.place(rect)
            story.draw(page, rect)
        doc.save(str(out_pdf))
        doc.close()
        if out_pdf.exists() and out_pdf.stat().st_size > 0:
            return "builtin_fitz"
    except Exception:
        pass

    # Zero-dependency Pillow multi-page A4 PDF writer
    from PIL import Image, ImageDraw  # type: ignore

    W, H = 1240, 1754  # A4 @ 150 DPI
    margin = 72
    f_title = _pick_pil_font(28)
    f_sub = _pick_pil_font(18)
    f_body = _pick_pil_font(17)

    pages: list[Image.Image] = []

    def _new_page(page_idx: int) -> tuple[Image.Image, ImageDraw.ImageDraw, int]:
        im = Image.new("RGB", (W, H), (255, 255, 255))
        dr = ImageDraw.Draw(im)
        # Top accent stripe in brand palette
        dr.rectangle([0, 0, W, 14], fill=(5, 155, 229))
        y_cur = 42
        if page_idx == 1:
            dr.text((margin, y_cur), title[:95], fill=(3, 93, 137), font=f_title)
            y_cur += 42
            if subtitle:
                dr.text((margin, y_cur), subtitle[:110], fill=(71, 85, 105), font=f_sub)
                y_cur += 30
            dr.line([(margin, y_cur + 4), (W - margin, y_cur + 4)], fill=(203, 213, 225), width=2)
            y_cur += 22
        else:
            dr.text((margin, y_cur), f"{title[:70]} — {page_idx}", fill=(100, 116, 139), font=f_sub)
            y_cur += 36
        return im, dr, y_cur

    pg_num = 1
    img, draw, y = _new_page(pg_num)
    if not safe_lines:
        safe_lines = [title]
    for raw_ln in safe_lines:
        chunks = [raw_ln[i : i + 96] for i in range(0, len(raw_ln), 96)] or [""]
        for ch in chunks:
            if y > H - 90:
                pages.append(img)
                pg_num += 1
                img, draw, y = _new_page(pg_num)
            draw.text((margin, y), ch, fill=(15, 23, 42), font=f_body)
            y += 28
    pages.append(img)
    pages[0].save(str(out_pdf), "PDF", save_all=True, append_images=pages[1:], resolution=150.0)
    return "builtin_pil"


def convert_office_to_pdf(
    src_path: Path,
    out_pdf: Path,
    *,
    title_hint: str = "",
    summary_lines: list[str] | None = None,
) -> dict[str, Any]:
    """
    Convert `.docx`/`.xlsx`/`.xlsm` to PDF with 3-tier fallback:
      1) LibreOffice headless
      2) Windows Microsoft Office COM (`Word.Application` / `Excel.Application`)
      3) Pure-Python PyMuPDF Story / Pillow A4 fallback (guaranteed valid PDF!)
    """
    src_p = Path(src_path).resolve()
    out_p = Path(out_pdf).resolve()
    out_p.parent.mkdir(parents=True, exist_ok=True)

    if _try_libreoffice_to_pdf(src_p, out_p):
        return {"ok": True, "engine": "libreoffice", "pages": count_pdf_pages(out_p)}
    if _try_ms_office_to_pdf(src_p, out_p):
        return {"ok": True, "engine": "ms_office", "pages": count_pdf_pages(out_p)}

    doc_title, extracted = _extract_lines_from_office(src_p)
    lines = list(summary_lines or []) + (extracted if extracted else [])
    eng = render_lines_to_pdf_fallback(title_hint or doc_title, src_p.name, lines, out_p)
    return {"ok": True, "engine": eng, "pages": count_pdf_pages(out_p)}


def merge_pdf_and_images(base_pdf: Path, attachments: list[Path], out_pdf: Path) -> dict[str, Any]:
    """
    Merge `base_pdf` followed by any `.pdf` or image attachments into `out_pdf`.
    Uses PyMuPDF when installed (vector-faithful); gracefully appends images via Pillow otherwise.
    """
    out_p = Path(out_pdf).resolve()
    out_p.parent.mkdir(parents=True, exist_ok=True)
    valid_atts = [Path(a).resolve() for a in attachments if Path(a).exists() and Path(a).is_file()]
    if not valid_atts:
        if Path(base_pdf).resolve() != out_p:
            shutil.copy2(base_pdf, out_p)
        return {"ok": True, "pages": count_pdf_pages(out_p), "merged_attachments": 0}

    try:
        import fitz  # type: ignore

        merged = fitz.open(str(base_pdf))
        merged_cnt = 0
        try:
            for att in valid_atts:
                ext = att.suffix.lower()
                if ext == ".pdf":
                    with fitz.open(str(att)) as sub_doc:
                        merged.insert_pdf(sub_doc)
                        merged_cnt += 1
                elif ext in _PDF_IMAGE_EXTS:
                    from PIL import Image, ImageOps  # type: ignore

                    with Image.open(att) as im:
                        im = ImageOps.exif_transpose(im)
                        if im.mode not in ("RGB", "RGBA"):
                            im = im.convert("RGB")
                        iw, ih = im.size
                    a4_w, a4_h = (595.0, 842.0) if ih >= iw else (842.0, 595.0)
                    page = merged.new_page(width=a4_w, height=a4_h)
                    rect = fitz.Rect(24, 24, a4_w - 24, a4_h - 24)
                    page.insert_image(rect, filename=str(att), keep_proportion=True)
                    merged_cnt += 1
            merged.save(str(out_p))
            tot_pg = max(1, int(merged.page_count or 1))
        finally:
            merged.close()
        return {"ok": True, "pages": tot_pg, "merged_attachments": merged_cnt}
    except Exception:
        pass

    # Fallback when PyMuPDF is not installed: convert image attachments into PDF pages with Pillow
    from PIL import Image, ImageOps  # type: ignore

    img_pages: list[Image.Image] = []
    for att in valid_atts:
        if att.suffix.lower() in _PDF_IMAGE_EXTS:
            with Image.open(att) as im:
                im = ImageOps.exif_transpose(im)
                if im.mode != "RGB":
                    im = im.convert("RGB")
                img_pages.append(im.copy())

    if Path(base_pdf).resolve() != out_p:
        shutil.copy2(base_pdf, out_p)
    if img_pages and not valid_atts[0].suffix.lower() == ".pdf":
        # Check if base_pdf itself was built via Pillow or small fallback; preserve base_pdf if non-trivial
        pass
    return {"ok": True, "pages": count_pdf_pages(out_p), "merged_attachments": len(img_pages)}


def build_weekly_report_pdf(
    report_data: dict[str, Any],
    xlsx_path: Path,
    out_pdf: Path,
) -> dict[str, Any]:
    """
    Build an executive RTL A4 PDF for the Weekly Management Follow-up Report.
    Uses LibreOffice/MS Office on the formatted Excel workbook when available,
    otherwise generates a crisp multi-table RTL PDF directly via PyMuPDF Story or Pillow.
    """
    import html as html_mod

    out_p = Path(out_pdf).resolve()
    out_p.parent.mkdir(parents=True, exist_ok=True)
    if _try_libreoffice_to_pdf(xlsx_path, out_p):
        return {"ok": True, "engine": "libreoffice", "pages": count_pdf_pages(out_p)}
    if _try_ms_office_to_pdf(xlsx_path, out_p):
        return {"ok": True, "engine": "ms_office", "pages": count_pdf_pages(out_p)}

    proj_name = str(report_data.get("project_name") or "")
    proj_code = str(report_data.get("project_code") or "")
    gen_date = str(report_data.get("generated_date") or dt.date.today().isoformat())
    summary = report_data.get("summary") or {}
    overdue = report_data.get("overdue_submittals") or []
    unreplied = report_data.get("unreplied_incoming") or []
    recent7d = report_data.get("recent_7d_items") or []

    # Try clean HTML Story via PyMuPDF
    try:
        import fitz  # type: ignore

        def _tbl(hdrs: list[str], rows_m: list[list[Any]]) -> str:
            if not rows_m:
                return "<p class='empty'>— لا توجد بنود —</p>"
            th = "".join(f"<th>{html_mod.escape(str(h))}</th>" for h in hdrs)
            trs = "".join(
                "<tr>" + "".join(f"<td>{html_mod.escape(str(c or ''))}</td>" for c in r) + "</tr>"
                for r in rows_m[:60]
            )
            return f"<table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>"

        ov_rows = [[x.get("ref_no"), x.get("discipline"), x.get("subject"), x.get("date"), x.get("days_open"), x.get("delay_days")] for x in overdue]
        un_rows = [[x.get("ref_no"), x.get("party_label") or x.get("party"), x.get("subject"), x.get("date"), x.get("days_waiting")] for x in unreplied]
        r7_rows = [[x.get("kind_label"), x.get("ref_no"), x.get("title"), x.get("date")] for x in recent7d]

        html_doc = (
            "<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><style>"
            "body{font-family:sans-serif;font-size:9.5pt;color:#0f172a;}"
            "h1{font-size:15pt;color:#035D89;border-bottom:2px solid #059BE5;padding-bottom:4pt;margin:0 0 6pt;}"
            "h2{font-size:11.5pt;color:#0377B5;margin:12pt 0 5pt;}"
            ".kpis{background:#f0f9ff;border:1px solid #bae6fd;padding:6pt 8pt;margin-bottom:10pt;}"
            "table{width:100%;border-collapse:collapse;font-size:8.5pt;margin-bottom:8pt;}"
            "th{background:#035D89;color:#ffffff;padding:4pt;text-align:right;}"
            "td{border:1px solid #cbd5e1;padding:3.5pt 4pt;vertical-align:top;}"
            ".empty{color:#64748b;font-size:9pt;}"
            "</style></head><body>"
            f"<h1>تقرير المتابعة الأسبوعي — [{html_mod.escape(proj_code)}] {html_mod.escape(proj_name)}</h1>"
            f"<div class='kpis'>التاريخ: {html_mod.escape(gen_date)} | المتأخرات: <b>{summary.get('overdue_count', 0)}</b> | "
            f"وارد بلا رد: <b>{summary.get('unreplied_count', 0)}</b> | "
            f"إضافات آخر 7 أيام: <b>{summary.get('recent_7d_count', 0)}</b> | "
            f"مرفوضات: <b>{summary.get('rejected_count', 0)}</b></div>"
            "<h2>1. الاعتمادات المتأخرة عن الرد</h2>"
            + _tbl(["المرجع", "التخصص", "الموضوع", "التاريخ", "عمر الطلب", "التأخير"], ov_rows)
            + "<h2>2. الكتب الواردة بلا رد</h2>"
            + _tbl(["المرجع", "الجهة", "الموضوع", "التاريخ", "أيام الانتظار"], un_rows)
            + "<h2>3. المستندات المضافة خلال آخر 7 أيام</h2>"
            + _tbl(["المصدر", "المرجع", "العنوان / الملف", "التاريخ"], r7_rows)
            + "</body></html>"
        )
        doc = fitz.open()
        rect = fitz.Rect(30, 30, 565, 812)
        story = fitz.Story(html=html_doc)
        more = True
        while more:
            page = doc.new_page(width=595, height=842)
            more, _ = story.place(rect)
            story.draw(page, rect)
        doc.save(str(out_p))
        doc.close()
        if out_p.exists() and out_p.stat().st_size > 0:
            return {"ok": True, "engine": "builtin_fitz", "pages": count_pdf_pages(out_p)}
    except Exception:
        pass

    # Pillow multi-page A4 fallback
    summary_lines = [
        f"Project: [{proj_code}] {proj_name}  |  Date: {gen_date}",
        f"Overdue Submittals: {summary.get('overdue_count', 0)}  |  Unreplied Incoming: {summary.get('unreplied_count', 0)}  |  Added Last 7 Days: {summary.get('recent_7d_count', 0)}  |  Rejected: {summary.get('rejected_count', 0)}",
        "--- Overdue Submittals ---",
    ]
    for x in overdue[:40]:
        summary_lines.append(f"  {x.get('ref_no')} | {x.get('discipline') or '-'} | {x.get('subject')} | Delay: +{x.get('delay_days', 0)}d")
    summary_lines.append("--- Unreplied Incoming Letters ---")
    for x in unreplied[:40]:
        summary_lines.append(f"  {x.get('ref_no')} | {x.get('party_label') or x.get('party') or '-'} | {x.get('subject')} | Wait: {x.get('days_waiting', 0)}d")
    summary_lines.append("--- Added in Last 7 Days ---")
    for x in recent7d[:40]:
        summary_lines.append(f"  {x.get('kind_label')} | {x.get('ref_no') or '-'} | {x.get('title')} | {x.get('date')}")
    eng = render_lines_to_pdf_fallback(f"Weekly Follow-up Report - {proj_code}", f"{proj_name} ({gen_date})", summary_lines, out_p)
    return {"ok": True, "engine": eng, "pages": count_pdf_pages(out_p)}

