#!/usr/bin/env python3
"""
template_filler.py — فحص وتعبئة نماذج المشاريع الأصلية (Word .docx و Excel .xlsx/.xlsm)
مع الحفاظ 100% على الشعارات (Logos)، الترويسة، التذييل، الجداول، الحدود، الخطوط، الخلايا المدمجة، والماكرو!
+ استخراج الرقم التسلسلي التالي (next-ref) تلقائياً من لوج المشروع لمنع تكرار أو سقوط الأرقام.

الأوامر:
  1) فحص هيكل أي نموذج Word أو Excel (حتى لو كان ملفاً قديماً معبّأً مسبقاً بدون Placeholders):
     python template_filler.py inspect "نموذج_اعتماد_SOOR.docx"
     python template_filler.py inspect "نموذج_مستخلص.xlsx"

  2) تعبئة نموذج Word (.docx) وإصدار ملف جديد جاهز للطباعة دون المساس بالقالب الأصلي:
     python template_filler.py fill-docx "template.docx" "edits.json" "EL-228-R00.docx"

  3) تعبئة نموذج Excel (.xlsx / .xlsm) مع حماية المعادلات والخلايا المدمجة والماكرو:
     python template_filler.py fill-xlsx "template.xlsx" "edits.json" "IPC-05.xlsx"

  4) معرفة الرقم المرجعي التالي تلقائياً من سجل المشروع (يدعم الترقيم المتسلسل العام أو المستقل لكل قسم):
     python template_filler.py next-ref "LOG_Final_Updated.xlsm" --prefix EL
"""
import argparse
import json
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


# =====================================================================
# 1) فحص النماذج الأصلية (Word & Excel) بدقة إحداثية كاملة
# =====================================================================
def _unique_row_cells(row):
    """يتجنب تكرار الخلايا المدمجة أفقياً في جداول Word."""
    seen = set()
    out = []
    for c_idx, cell in enumerate(row.cells):
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
        out.append((c_idx, cell))
    return out


def inspect_docx(path: Path) -> dict:
    doc = Document(str(path))
    data = {
        "file": str(path),
        "type": "docx",
        "headers": [],
        "paragraphs": [],
        "tables": [],
        "footers": [],
    }

    for s_i, sec in enumerate(doc.sections):
        hdr_texts = [p.text.strip() for p in sec.header.paragraphs if p.text.strip()]
        for t_i, tbl in enumerate(sec.header.tables):
            for r_i, row in enumerate(tbl.rows):
                vals = [f"c{c_i}:{c.text.strip()}" for c_i, c in _unique_row_cells(row) if c.text.strip()]
                if vals:
                    hdr_texts.append(f"[HeaderTable {t_i} Row {r_i}] " + " | ".join(vals))
        if hdr_texts:
            data["headers"].append({"section": s_i, "content": hdr_texts})

        ftr_texts = [p.text.strip() for p in sec.footer.paragraphs if p.text.strip()]
        if ftr_texts:
            data["footers"].append({"section": s_i, "content": ftr_texts})

    for p_i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            data["paragraphs"].append({"index": p_i, "text": p.text})

    for t_i, tbl in enumerate(doc.tables):
        rows_out = []
        seen_vmerge = set()
        for r_i, row in enumerate(tbl.rows):
            cells_out = []
            for c_i, cell in _unique_row_cells(row):
                tc_id = id(cell._tc)
                if tc_id in seen_vmerge:
                    continue
                seen_vmerge.add(tc_id)
                cells_out.append({"col": c_i, "text": cell.text.strip()})
            rows_out.append({"row": r_i, "cells": cells_out})
        data["tables"].append({
            "table": t_i,
            "row_count": len(tbl.rows),
            "col_count": len(tbl.columns),
            "rows": rows_out,
        })
    return data


def inspect_xlsx_template(path: Path) -> dict:
    with zipfile.ZipFile(path, "r") as zf:
        has_vba = any("vbaProject.bin" in n for n in zf.namelist())
        has_images = any(n.startswith("xl/media/") for n in zf.namelist())

    wb = openpyxl.load_workbook(str(path), data_only=False, keep_vba=has_vba)
    sheets = []
    for name in wb.sheetnames:
        ws = wb[name]
        cells = []
        formulas = []
        for r in range(1, min(ws.max_row + 1, 120)):
            for c in range(1, min(ws.max_column + 1, 35)):
                cell = ws.cell(r, c)
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                coord = f"{get_column_letter(c)}{r}"
                val = str(cell.value)
                if val.startswith("="):
                    formulas.append({"cell": coord, "formula": val})
                elif val.strip():
                    cells.append({"cell": coord, "value": val.strip()[:120]})
        sheets.append({
            "name": name,
            "rows": ws.max_row,
            "cols": ws.max_column,
            "merged_ranges": [str(rng) for rng in list(ws.merged_cells.ranges)[:60]],
            "non_empty_cells": cells[:200],
            "formula_cells": formulas[:80],
        })
    return {
        "file": str(path),
        "type": path.suffix.lower().lstrip("."),
        "has_vba": has_vba,
        "has_embedded_logos_or_images": has_images,
        "sheets": sheets,
    }


# =====================================================================
# 2) تعبئة نماذج Word (.docx) مع حفظ التنسيق والشعارات وRTL 100%
# =====================================================================
ARABIC_RE = re.compile(r"[\u0600-\u06FF]")


def _replace_para_text_keep_style(para, new_text: str):
    """يستبدل نص الفقرة مع الحفاظ الكامل على خط وحجم ولون واتجاه أول Run فيها."""
    if para.runs:
        first = para.runs[0]
        first.text = new_text
        for extra in para.runs[1:]:
            extra.text = ""
        if ARABIC_RE.search(new_text):
            rPr = first._r.get_or_add_rPr()
            if rPr.find(qn("w:rtl")) is None:
                rPr.append(OxmlElement("w:rtl"))
    else:
        run = para.add_run(new_text)
        if ARABIC_RE.search(new_text):
            run._r.get_or_add_rPr().append(OxmlElement("w:rtl"))


def _replace_substring_in_para(para, old: str, new: str) -> int:
    """يستبدل عبارة داخل فقرة حتى لو قسّمها Word عبر عدة Runs متجاورة!"""
    full = para.text
    if old not in full:
        return 0
    count = full.count(old)
    # محاولة الاستبدال داخل نفس الـ Run أولاً للحفاظ على التنسيق الموضعي (مثل كلمة غامقة وسط سطر)
    replaced_in_run = False
    for r in para.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
            replaced_in_run = True
    if not replaced_in_run:
        _replace_para_text_keep_style(para, full.replace(old, new))
    return count


def _iter_all_docx_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        seen = set()
        for row in tbl.rows:
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for p in cell.paragraphs:
                    yield p
    for sec in doc.sections:
        for container in (sec.header, sec.footer):
            for p in container.paragraphs:
                yield p
            for tbl in container.tables:
                seen = set()
                for row in tbl.rows:
                    for cell in row.cells:
                        if id(cell._tc) in seen:
                            continue
                        seen.add(id(cell._tc))
                        for p in cell.paragraphs:
                            yield p


def fill_docx(template_path: Path, spec: dict, out_path: Path):
    """
    صيغة edits.json المرنة:
    {
      "replacements": {"EL-120": "EL-228", "2026/10/01": "2026/11/17", "{{SUBJECT}}": "اعتماد كابلات"},
      "cells": [
        {"table": 0, "row": 2, "col": 1, "value": "EL-228"},
        {"table": 0, "row": 2, "col": 3, "value": "R00"}
      ],
      "header_cells": [
        {"section": 0, "table": 0, "row": 1, "col": 2, "value": "EL-228"}
      ],
      "paragraphs": [
        {"index": 4, "value": "الموضوع: اعتماد كابلات ضغط منخفض"}
      ],
      "append_table_rows": [
        {"table": 1, "rows": [["1", "كابلات 4x16mm2", "شركة الخليج", "500 م"]]}
      ]
    }
    """
    doc = Document(str(template_path))
    changes = 0

    # 1) استبدال النصوص والـ Placeholders في المتن والجداول والهيدر والفوتر
    for old, new in (spec.get("replacements") or {}).items():
        for p in _iter_all_docx_paragraphs(doc):
            changes += _replace_substring_in_para(p, str(old), str(new))

    # 2) تعديل خلايا الجداول بالإحداثيات المباشرة
    for c_edit in spec.get("cells") or []:
        tbl = doc.tables[int(c_edit["table"])]
        cell = tbl.rows[int(c_edit["row"])].cells[int(c_edit["col"])]
        lines = str(c_edit["value"]).split("\n")
        _replace_para_text_keep_style(cell.paragraphs[0], lines[0])
        for extra_p in cell.paragraphs[1:]:
            _replace_para_text_keep_style(extra_p, "")
        if len(lines) > 1:
            _replace_para_text_keep_style(cell.paragraphs[0], str(c_edit["value"]))
        changes += 1

    # 3) تعديل خلايا جداول الترويسة (Header) إن لزم
    for h_edit in spec.get("header_cells") or []:
        sec = doc.sections[int(h_edit.get("section", 0))]
        tbl = sec.header.tables[int(h_edit["table"])]
        cell = tbl.rows[int(h_edit["row"])].cells[int(h_edit["col"])]
        _replace_para_text_keep_style(cell.paragraphs[0], str(h_edit["value"]))
        changes += 1

    # 4) تعديل فقرات محددة بالرقم
    for p_edit in spec.get("paragraphs") or []:
        para = doc.paragraphs[int(p_edit["index"])]
        _replace_para_text_keep_style(para, str(p_edit["value"]))
        changes += 1

    # 5) إضافة صفوف لجدول قائم (مثلاً حافظة Transmittal أو بنود اعتماد)
    for app in spec.get("append_table_rows") or []:
        tbl = doc.tables[int(app["table"])]
        for r_vals in app.get("rows") or []:
            new_row = tbl.add_row()
            for ci, val in enumerate(r_vals):
                if ci < len(new_row.cells):
                    _replace_para_text_keep_style(new_row.cells[ci].paragraphs[0], str(val))
            changes += 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"[OK] تم إصدار ملف Word من النموذج الأصلي بنجاح ({changes} حقل محدّث): {out_path}")


# =====================================================================
# 3) تعبئة نماذج Excel (.xlsx / .xlsm) مع حماية المعادلات والدمج والماكرو
# =====================================================================
def _resolve_writable_cell(ws, coord: str):
    """إذا كانت الخلية ضمن نطاق مدمج يرجع الخلية العلوية اليسرى القابلة للكتابة تلقائياً."""
    cell = ws[coord]
    if not isinstance(cell, MergedCell):
        return cell
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            return ws.cell(rng.min_row, rng.min_col)
    return cell


def fill_xlsx(template_path: Path, spec: dict, out_path: Path):
    """
    صيغة edits.json:
    {
      "sheet": "اسم الشيت (اختياري — الافتراضي الشيت النشط)",
      "allow_overwrite_formulas": false,
      "cells": {
        "C4": "EL-228",
        "G4": "2026-11-17",
        "C6": "اعتماد كابلات ضغط منخفض"
      },
      "multi_sheet_cells": {
        "Sheet2": {"B3": 1500.500}
      },
      "replacements": {"{{PROJECT}}": "مركز صحي السالمية"}
    }
    """
    with zipfile.ZipFile(template_path, "r") as zf:
        has_vba = any("vbaProject.bin" in n for n in zf.namelist())

    wb = openpyxl.load_workbook(str(template_path), keep_vba=has_vba)
    allow_fml = bool(spec.get("allow_overwrite_formulas", False))
    changes = 0
    skipped_formulas = []

    def apply_map(ws, cmap: dict):
        nonlocal changes
        for coord, val in cmap.items():
            target = _resolve_writable_cell(ws, coord)
            cur = target.value
            if isinstance(cur, str) and cur.startswith("=") and not allow_fml:
                if not (isinstance(val, str) and val.startswith("=")):
                    skipped_formulas.append(f"{ws.title}!{target.coordinate}")
                    continue
            target.value = val
            changes += 1

    # استبدال النصوص العامة (لا يلمس المعادلات أبداً)
    repls = spec.get("replacements") or {}
    if repls:
        for ws in wb.worksheets:
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    cell = ws.cell(r, c)
                    if isinstance(cell.value, str) and not cell.value.startswith("="):
                        txt = cell.value
                        for old, new in repls.items():
                            if str(old) in txt:
                                txt = txt.replace(str(old), str(new))
                        if txt != cell.value:
                            cell.value = txt
                            changes += 1

    if spec.get("cells"):
        sname = spec.get("sheet")
        ws = wb[sname] if (sname and sname in wb.sheetnames) else wb.active
        apply_map(ws, spec["cells"])

    for sname, cmap in (spec.get("multi_sheet_cells") or {}).items():
        if sname in wb.sheetnames:
            apply_map(wb[sname], cmap)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    msg = f"[OK] تم إصدار ملف Excel من النموذج الأصلي ({changes} خلية محدّثة"
    if has_vba:
        msg += " + تم الحفاظ على VBA"
    msg += f"): {out_path}"
    print(msg)
    if skipped_formulas:
        print(f"   [حماية تلقائية] لم يتم الكتابة فوق خلايا المعادلات: {skipped_formulas}")


# =====================================================================
# 4) استخراج الرقم التسلسلي التالي (next-ref) تلقائياً من لوج المشروع
# =====================================================================
REF_PATTERN = re.compile(r"^([A-Za-z]{2,6})\s*[-_/]\s*0*(\d{1,6})(?:\s*[-_/]\s*R\d+)?$")


def compute_next_ref(log_path: Path, prefix: str | None = None, sheet: str | None = None) -> dict:
    if not log_path.exists():
        p = (prefix or "REF").upper()
        return {"numbering_scheme": "new_log", "suggested_next_ref": f"{p}-001", "global_max": 0, "prefix_max": 0}
    wb = openpyxl.load_workbook(str(log_path), data_only=True, read_only=True)
    per_prefix: dict[str, set[int]] = defaultdict(set)
    all_nums: dict[int, set[str]] = defaultdict(set)
    pad_widths: list[int] = []

    req_upper = (prefix or "").upper()
    if sheet and sheet in wb.sheetnames:
        target_sheets = [sheet]
    elif "01_الاعتمادات" in wb.sheetnames and req_upper not in {"INV", "IPC", "VO", "PO", "PC", "OUT", "IN"}:
        # في Master_Archive_Log افحص شيت الاعتمادات فقط حتى لا تختلط أرقام الفواتير والعهدة بأرقام الاعتمادات!
        target_sheets = ["01_الاعتمادات"]
    else:
        target_sheets = wb.sheetnames

    for name in target_sheets:
        ws = wb[name]
        for row in ws.iter_rows(values_only=True):
            for val in row[:12]:
                if not isinstance(val, str):
                    continue
                s = val.strip()
                m = re.match(r"^([A-Za-z]{2,6})\s*[-_/]\s*(0*\d{1,6})$", s)
                if m:
                    pfx = m.group(1).upper()
                    raw_num = m.group(2)
                    num = int(raw_num)
                    if 0 < num < 100000:
                        per_prefix[pfx].add(num)
                        all_nums[num].add(pfx)
                        pad_widths.append(len(raw_num))
    wb.close()

    if not all_nums:
        p = (prefix or "REF").upper()
        return {
            "numbering_scheme": "new_log",
            "suggested_next_ref": f"{p}-001",
            "global_max": 0,
            "prefix_max": 0,
        }

    # كشف ذكي: هل المشروع يرقّم تسلسلياً موحداً عبر كل الأقسام (CIV-001, EL-002, FF-227)
    # أم يرقّم كل قسم على حدة (CIV-001, EL-001, PL-001)؟
    shared_nums = sum(1 for pfxs in all_nums.values() if len(pfxs) > 1)
    multi_prefix = len(per_prefix) >= 2
    is_global_seq = multi_prefix and (shared_nums / max(len(all_nums), 1) < 0.08)

    pad = max(3, min(max(set(pad_widths), key=pad_widths.count), 5)) if pad_widths else 3
    global_max = max(all_nums.keys())
    req_pfx = (prefix or next(iter(per_prefix.keys()))).upper()
    pfx_max = max(per_prefix.get(req_pfx, {0}))

    next_num = (global_max + 1) if is_global_seq else (pfx_max + 1)
    suggested = f"{req_pfx}-{next_num:0{pad}d}"

    return {
        "file": str(log_path),
        "numbering_scheme": "global_sequential_across_disciplines" if is_global_seq else "independent_per_discipline",
        "requested_prefix": req_pfx,
        "global_max_number": global_max,
        "prefix_max_number": pfx_max,
        "existing_prefixes": {k: {"count": len(v), "max": max(v)} for k, v in sorted(per_prefix.items())},
        "suggested_next_ref": suggested,
        "suggested_revision": "R00",
    }


# =====================================================================
# 5) التسطير الرسمي الواضح + محرك المصمم المرئي لخانات النماذج (Visual Template Studio)
# =====================================================================
AUTO_FIELD_CATALOG = [
    ("project_name", "اسم المشروع"),
    ("project_code", "كود المشروع"),
    ("contract_no", "رقم العقد"),
    ("client", "المالك / الوزارة / الجهة"),
    ("consultant", "المكتب الاستشاري"),
    ("contractor", "الشركة المقاولة"),
    ("auto_ref", "الرقم المرجعي التسلسلي التلقائي (EL-001 / IPC-01...)"),
    ("revision", "رقم المراجعة (R00)"),
    ("discipline", "القسم الهندسي (EL / CIV / ARCH...)"),
    ("today_iso", "تاريخ اليوم (2026-11-17)"),
    ("today_dmy", "تاريخ اليوم (17/11/2026)"),
    ("currency", "عملة المشروع (KWD)"),
]


_TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
    "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
    "tblLook", "tblCaption", "tblDescription", "tblPrChange",
]
_TCPR_ORDER = [
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
    "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
]
_SECTPR_ORDER = [
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
    "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
    "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]


def _insert_ooxml_ordered(parent, elem, order_list: list[str]):
    """يدرج عنصر OOXML ويرتّب كافة عناصر الأب بالترتيب القياسي الصارم (ISO-29500) ليظهر التسطير 100% في Word."""
    tag_local = elem.tag.split("}")[-1]
    for old in list(parent.findall(qn(f"w:{tag_local}"))):
        parent.remove(old)
    parent.append(elem)
    children = list(parent)
    rank_map = {name: i for i, name in enumerate(order_list)}
    children.sort(key=lambda el: rank_map.get(el.tag.split("}")[-1], len(order_list)))
    for ch in children:
        parent.remove(ch)
        parent.append(ch)


def _set_docx_table_crisp_borders(table, inner_color: str = "1E293B", inner_sz: str = "6", outer_color: str = "0F172A", outer_sz: str = "12"):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    try:
        table.style = "Table Grid"
    except Exception:
        pass
    tblPr = table._tbl.tblPr
    # إعادة ترتيب bidiVisual إن وُجد ليكون في مكانه القياسي قبل tblW/tblLook
    if tblPr.find(qn("w:bidiVisual")) is not None:
        _insert_ooxml_ordered(tblPr, parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'), _TBLPR_ORDER)
    _insert_ooxml_ordered(tblPr, parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:bottom w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:left w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:right w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:insideH w:val="single" w:sz="{inner_sz}" w:space="0" w:color="{inner_color}"/>'
        f'<w:insideV w:val="single" w:sz="{inner_sz}" w:space="0" w:color="{inner_color}"/>'
        f"</w:tblBorders>"
    ), _TBLPR_ORDER)
    nrows = len(table.rows)
    ncols = len(table.columns)
    for r_i, row in enumerate(table.rows):
        seen_row = set()
        for c_i, cell in enumerate(row.cells):
            tc = cell._tc
            if id(tc) in seen_row:
                continue
            seen_row.add(id(tc))
            last_c = c_i
            while last_c + 1 < ncols and row.cells[last_c + 1]._tc is tc:
                last_c += 1
            tcPr = tc.get_or_add_tcPr()
            t_sz, t_cl = (outer_sz, outer_color) if r_i == 0 else (inner_sz, inner_color)
            b_sz, b_cl = (outer_sz, outer_color) if r_i == nrows - 1 else (inner_sz, inner_color)
            l_sz, l_cl = (outer_sz, outer_color) if c_i == 0 else (inner_sz, inner_color)
            r_sz, r_cl = (outer_sz, outer_color) if last_c == ncols - 1 else (inner_sz, inner_color)
            _insert_ooxml_ordered(tcPr, parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="{t_sz}" w:space="0" w:color="{t_cl}"/>'
                f'<w:bottom w:val="single" w:sz="{b_sz}" w:space="0" w:color="{b_cl}"/>'
                f'<w:left w:val="single" w:sz="{l_sz}" w:space="0" w:color="{l_cl}"/>'
                f'<w:right w:val="single" w:sz="{r_sz}" w:space="0" w:color="{r_cl}"/>'
                f"</w:tcBorders>"
            ), _TCPR_ORDER)


def apply_crisp_borders_docx(src: Path, out_path: Path | None = None, page_frame: bool = True):
    """يطبّق تسطيراً رسمياً داكناً وواضحاً 100% للطباعة على كافة جداول وخلايا ملف Word (+ إطار صفحة A4 اختياري)."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    target = out_path or src
    doc = Document(str(src))
    for sec in doc.sections:
        for htbl in sec.header.tables:
            _set_docx_table_crisp_borders(htbl)
        if page_frame:
            sectPr = sec._sectPr
            _insert_ooxml_ordered(sectPr, parse_xml(
                f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
                f'<w:top w:val="single" w:sz="12" w:space="18" w:color="1B365D"/>'
                f'<w:bottom w:val="single" w:sz="12" w:space="18" w:color="1B365D"/>'
                f'<w:left w:val="single" w:sz="12" w:space="18" w:color="1B365D"/>'
                f'<w:right w:val="single" w:sz="12" w:space="18" w:color="1B365D"/>'
                f"</w:pgBorders>"
            ), _SECTPR_ORDER)
    for tbl in doc.tables:
        _set_docx_table_crisp_borders(tbl)
    target.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(target))


def apply_crisp_borders_xlsx(src: Path, out_path: Path | None = None):
    """يطبّق تسطيراً رسمياً واضحاً للطباعة والشاشة على ملف Excel (.xlsx/.xlsm) دون المساس بالمعادلات أو الماكرو."""
    from openpyxl.styles import Border, Side

    target = out_path or src
    with zipfile.ZipFile(src, "r") as zf:
        has_vba = any("vbaProject.bin" in n for n in zf.namelist())
    wb = openpyxl.load_workbook(str(src), keep_vba=has_vba)
    thin = Side(style="thin", color="1E293B")
    med = Side(style="medium", color="0F172A")

    for ws in wb.worksheets:
        if ws.views and ws.views.sheetView:
            ws.views.sheetView[0].showGridLines = True
        max_r = min(ws.max_row or 1, 200)
        max_c = min(ws.max_column or 1, 30)
        # نحدد الصفوف النشطة (التي تحتوي نصاً أو معادلة أو لوناً أو حدوداً أو دمجاً)
        merged_rows = set()
        for mr in ws.merged_cells.ranges:
            for rr in range(mr.min_row, mr.max_row + 1):
                merged_rows.add(rr)
        active_cols_end = 1
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cl = ws.cell(r, c)
                if cl.value is not None and str(cl.value).strip() != "":
                    active_cols_end = max(active_cols_end, c)
        for mr in ws.merged_cells.ranges:
            active_cols_end = max(active_cols_end, min(mr.max_col, max_c))

        active_rows = []
        for r in range(1, max_r + 1):
            has_any = r in merged_rows
            if not has_any:
                for c in range(1, active_cols_end + 1):
                    cl = ws.cell(r, c)
                    b = cl.border
                    hb = bool(b and any(getattr(b, s) and getattr(b, s).style for s in ("top", "bottom", "left", "right")))
                    if (cl.value is not None and str(cl.value).strip() != "") or hb:
                        has_any = True
                        break
            if has_any:
                active_rows.append(r)
                for c in range(1, active_cols_end + 1):
                    ws.cell(r, c).border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # تصحيح حدود الخلايا المدمجة (إزالة الخطوط الداخلية وإبقاء المحيط الواضح)
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row not in active_rows:
                continue
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    ws.cell(r, c).border = Border(
                        top=thin if r == mr.min_row else None,
                        bottom=thin if r == mr.max_row else None,
                        left=thin if c == mr.min_col else None,
                        right=thin if c == mr.max_col else None,
                    )

        # إطار خارجي سميك حول كل جدول متصل
        if active_rows:
            blocks = []
            s_r = active_rows[0]
            p_r = active_rows[0]
            for rr in active_rows[1:]:
                if rr == p_r + 1:
                    p_r = rr
                else:
                    blocks.append((s_r, p_r))
                    s_r = p_r = rr
            blocks.append((s_r, p_r))
            for b1, b2 in blocks:
                for r in range(b1, b2 + 1):
                    for c in range(1, active_cols_end + 1):
                        cb = ws.cell(r, c).border
                        ws.cell(r, c).border = Border(
                            top=med if r == b1 else (cb.top if cb else None),
                            bottom=med if r == b2 else (cb.bottom if cb else None),
                            left=med if c == 1 else (cb.left if cb else None),
                            right=med if c == active_cols_end else (cb.right if cb else None),
                        )

    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(target))


def _extract_docx_table_grid(table, key_prefix: str, title: str) -> dict:
    nrows = len(table.rows)
    ncols = len(table.columns)
    row_tcs = [[cell._tc for cell in row.cells] for row in table.rows]
    seen_tc_ids: set[int] = set()
    grid_rows = []
    for r_i, row in enumerate(table.rows):
        row_cells = []
        c_i = 0
        while c_i < ncols:
            cell = row.cells[c_i]
            tc = row_tcs[r_i][c_i]
            tc_id = id(tc)
            colspan = 1
            while c_i + colspan < ncols and row_tcs[r_i][c_i + colspan] is tc:
                colspan += 1
            if tc_id in seen_tc_ids:
                c_i += colspan
                continue
            # فحص الدمج الرأسي الحقيقي (vMerge) أو مشاركة نفس عنصر tc
            rowspan = 1
            while r_i + rowspan < nrows and row_tcs[r_i + rowspan][c_i] is tc:
                rowspan += 1
            seen_tc_ids.add(tc_id)
            bg = ""
            tcPr = cell._tc.tcPr
            if tcPr is not None:
                shd = tcPr.find(qn("w:shd"))
                if shd is not None:
                    fill_v = str(shd.get(qn("w:fill")) or "").upper()
                    if fill_v and fill_v not in ("AUTO", "FFFFFF"):
                        bg = "#" + fill_v[-6:]
            bold = any(bool(run.bold) for p in cell.paragraphs for run in p.runs if run.text.strip())
            txt = "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
            row_cells.append({
                "key": f"{key_prefix}:{r_i}:{c_i}",
                "r": r_i,
                "c": c_i,
                "colspan": colspan,
                "rowspan": rowspan,
                "text": txt,
                "bg": bg,
                "bold": bold,
                "is_formula": False,
            })
            c_i += colspan
        grid_rows.append(row_cells)
    return {"id": key_prefix, "title": title, "ncols": ncols, "rows": grid_rows}


def extract_visual_grid(path: Path) -> dict:
    """يحوّل أي نموذج Word (.docx) أو Excel (.xlsx/.xlsm) إلى شبكة مرئية مطابقة للأصل لعرضها وتحديد الخانات بالنقر."""
    ext = path.suffix.lower()
    if ext == ".docx":
        doc = Document(str(path))
        sections_out = []
        for s_i, sec in enumerate(doc.sections):
            for ht_i, htbl in enumerate(sec.header.tables):
                sections_out.append(_extract_docx_table_grid(htbl, f"h:{s_i}:{ht_i}", f"ترويسة الصفحة (جدول {ht_i + 1})"))
        for t_i, tbl in enumerate(doc.tables):
            sections_out.append(_extract_docx_table_grid(tbl, f"t:{t_i}", f"جدول {t_i + 1}"))
        paras = []
        for p_i, p in enumerate(doc.paragraphs):
            tx = p.text.strip()
            if tx:
                paras.append({"key": f"p:{p_i}", "index": p_i, "text": tx})
        return {
            "file_name": path.name,
            "ext": ext,
            "kind": "docx",
            "tables": sections_out,
            "paragraphs": paras[:60],
            "auto_catalog": [{"key": k, "label": lbl} for k, lbl in AUTO_FIELD_CATALOG],
        }

    # Excel (.xlsx / .xlsm)
    wb = openpyxl.load_workbook(str(path), data_only=False)
    sheets_out = []
    for ws in wb.worksheets[:8]:
        max_r = min(ws.max_row or 1, 65)
        max_c = min(ws.max_column or 1, 18)
        # تقليم الصفوف والأعمدة الفارغة في النهاية
        real_r, real_c = 1, 1
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                v = ws.cell(r, c).value
                if v is not None and str(v).strip() != "":
                    real_r = max(real_r, r)
                    real_c = max(real_c, c)
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= max_r and mr.min_col <= max_c:
                real_r = max(real_r, min(mr.max_row, max_r))
                real_c = max(real_c, min(mr.max_col, max_c))

        merge_start: dict[tuple[int, int], tuple[int, int]] = {}
        merge_skip: set[tuple[int, int]] = set()
        for mr in ws.merged_cells.ranges:
            cs = mr.max_col - mr.min_col + 1
            rs = mr.max_row - mr.min_row + 1
            merge_start[(mr.min_row, mr.min_col)] = (rs, cs)
            for rr in range(mr.min_row, mr.max_row + 1):
                for cc in range(mr.min_col, mr.max_col + 1):
                    if (rr, cc) != (mr.min_row, mr.min_col):
                        merge_skip.add((rr, cc))

        rows_out = []
        for r in range(1, real_r + 1):
            rcells = []
            for c in range(1, real_c + 1):
                if (r, c) in merge_skip:
                    continue
                rs, cs = merge_start.get((r, c), (1, 1))
                cell = ws.cell(r, c)
                coord = f"{get_column_letter(c)}{r}"
                raw_v = cell.value
                is_fml = isinstance(raw_v, str) and raw_v.startswith("=")
                txt = "" if raw_v is None else str(raw_v)
                bg = ""
                if cell.fill and cell.fill.fill_type == "solid" and cell.fill.fgColor:
                    rgb = str(cell.fill.fgColor.rgb or "").upper()
                    if len(rgb) in (6, 8) and rgb[-6:] not in ("000000", "FFFFFF") and not rgb.startswith("Values"):
                        bg = "#" + rgb[-6:]
                fg = ""
                if cell.font and cell.font.color and cell.font.color.type == "rgb":
                    frgb = str(cell.font.color.rgb or "").upper()
                    if len(frgb) in (6, 8) and frgb[-6:] == "FFFFFF":
                        fg = "#FFFFFF"
                rcells.append({
                    "key": f"x:{ws.title}:{coord}",
                    "coord": coord,
                    "r": r,
                    "c": c,
                    "rowspan": rs,
                    "colspan": cs,
                    "text": txt,
                    "bg": bg,
                    "fg": fg,
                    "bold": bool(cell.font and cell.font.bold),
                    "is_formula": is_fml,
                })
            rows_out.append({"r": r, "cells": rcells})

        rtl = True
        if ws.views and ws.views.sheetView and ws.views.sheetView[0].rightToLeft is False:
            rtl = False
        sheets_out.append({
            "title": ws.title,
            "rtl": rtl,
            "col_letters": [get_column_letter(c) for c in range(1, real_c + 1)],
            "rows": rows_out,
        })

    return {
        "file_name": path.name,
        "ext": ext,
        "kind": "xlsx",
        "sheets": sheets_out,
        "auto_catalog": [{"key": k, "label": lbl} for k, lbl in AUTO_FIELD_CATALOG],
    }


def auto_suggest_field_map(grid: dict) -> dict[str, dict]:
    """يقترح تلقائياً تصنيف كل خانة (تلقائي من النظام / حقل إدخال / كتابة يدوية / ثابت) بدقة فورية ليسهّل على المستخدم."""
    fmap: dict[str, dict] = {}
    ph_auto = {
        "{{PROJECT_NAME}}": "project_name", "{{PROJECT}}": "project_name",
        "{{CONTRACT_NO}}": "contract_no",
        "{{CLIENT}}": "client", "[اسم المالك]": "client", "[شعار واسم المالك / الوزارة]": "client",
        "{{CONSULTANT}}": "consultant", "[اسم الاستشاري]": "consultant", "[شعار واسم المكتب الاستشاري]": "consultant",
        "{{CONTRACTOR}}": "contractor", "[اسم المقاول]": "contractor", "[شعار واسم الشركة المقاولة]": "contractor",
        "{{REF_NO}}": "auto_ref", "{{REF}}": "auto_ref",
        "{{REVISION}}": "revision",
        "{{DISCIPLINE}}": "discipline",
        "{{DATE}}": "today_iso",
    }
    ph_input = {
        "{{MATERIAL_DESCRIPTION}}": ("وصف المادة / البند", "textarea"),
        "{{INSPECTION_DESCRIPTION}}": ("وصف الأعمال / الفحص المطلوب", "textarea"),
        "{{SUBJECT}}": ("الموضوع / البيان", "text"),
        "{{MANUFACTURER}}": ("الشركة المصنّعة (Manufacturer)", "text"),
        "{{COUNTRY_OF_ORIGIN}}": ("بلد المنشأ (Origin)", "text"),
        "{{LOCAL_SUPPLIER}}": ("المورد المحلي / الوكيل", "text"),
        "{{BOQ_ITEM}}": ("رقم بند جدول الكميات (BOQ)", "text"),
        "{{SPEC_REF}}": ("مرجع المواصفات (Spec Ref)", "text"),
        "{{DWG_REF}}": ("رقم المخطط المرجعي (Drawing)", "text"),
        "{{LOCATION_ZONE_LEVEL}}": ("الموقع / المبنى / الدور", "text"),
        "{{INSPECTION_DATETIME}}": ("موعد الفحص المقترح", "text"),
        "{{APPROVED_REF_DWG}}": ("رقم المخطط / الاعتماد المرجعي", "text"),
        "{{DELIVERY_NOTE_NO}}": ("رقم إشعار التوريد", "text"),
        "{{QUANTITY}}": ("الكمية", "text"),
        "{{NEXT_DATE}}": ("موعد الاجتماع القادم", "text"),
    }

    def _classify_by_label(lbl_raw: str) -> dict | None:
        l = lbl_raw.strip().lower()
        if not l or len(l) > 65:
            return None
        if any(w in l for w in ("اسم المشروع", "المشروع /", "المشروع:", "project name", "project:")):
            return {"mode": "auto", "auto_key": "project_name"}
        if any(w in l for w in ("رقم العقد", "contract no", "contract #")):
            return {"mode": "auto", "auto_key": "contract_no"}
        if any(w in l for w in ("اسم المالك", "الجهة المالكة", "client", "employer", "owner")) and "اعتماد" not in l:
            return {"mode": "auto", "auto_key": "client"}
        if any(w in l for w in ("الاستشاري:", "المكتب الاستشاري", "consultant:")) and "اعتماد" not in l and "مراجعة" not in l:
            return {"mode": "auto", "auto_key": "consultant"}
        if any(w in l for w in ("المقاول:", "المقاول / الطرف", "الشركة المقاولة", "contractor:")) and "إعداد" not in l:
            return {"mode": "auto", "auto_key": "contractor"}
        if any(w in l for w in ("رقم الاعتماد", "رقم المستخلص", "رقم الأمر", "رقم أمر الشراء", "رقم الحافظة", "رقم التقرير", "الرقم المرجعي", "رقم الطلب", "submittal no", "ref no", "ref.", "ipc no", "po no", "vo no")):
            return {"mode": "auto", "auto_key": "auto_ref"}
        if any(w in l for w in ("المراجعة", "revision", "rev.")):
            return {"mode": "auto", "auto_key": "revision"}
        if any(w in l for w in ("التاريخ", "تاريخ المستخلص", "تاريخ التقديم", "date")):
            return {"mode": "auto", "auto_key": "today_iso"}
        if any(w in l for w in ("التخصص", "القسم", "discipline")) and len(l) < 30:
            return {"mode": "auto", "auto_key": "discipline"}
        if any(w in l for w in ("السادة المورد", "المورد", "المصنّع", "المصنع", "supplier", "vendor", "manufacturer", "مرسل إلى", "إلى السادة")):
            return {"mode": "input", "label": lbl_raw.strip(" :："), "input_type": "text"}
        if any(w in l for w in ("وصف المادة", "الموضوع", "الفترة من", "سبب التغيير", "شروط الدفع", "عرض السعر", "مهندس الموقع", "الموقع", "بلد المنشأ", "بند المقايسة", "العهدة المستلمة", "مبلغ العهدة", "subject", "description", "location", "origin", "boq")):
            return {"mode": "input", "label": lbl_raw.strip(" :："), "input_type": "text"}
        return None

    def _scan_rows(rows_list):
        for rcells in rows_list:
            for idx, cell in enumerate(rcells):
                if cell.get("is_formula"):
                    continue
                k = cell["key"]
                tx = (cell.get("text") or "").strip()
                # 1) فحص الـ Placeholders الصريحة
                matched = False
                for ph, ak in ph_auto.items():
                    if ph == tx or (ph in tx and len(tx) <= len(ph) + 6):
                        fmap[k] = {"mode": "auto", "auto_key": ak}
                        matched = True
                        break
                if matched:
                    continue
                for ph, (lbl, itype) in ph_input.items():
                    if ph in tx:
                        fmap[k] = {"mode": "input", "label": lbl, "input_type": itype}
                        matched = True
                        break
                if matched:
                    continue
                # 2) فحص الخلية المجاورة لعنوان واضح (عربي/إنجليزي) حتى لو لم يكن العنوان ملوّناً أو غامقاً
                if idx > 0 and not cell.get("bg") and not _classify_by_label(tx):
                    prev = rcells[idx - 1]
                    prev_tx = (prev.get("text") or "").strip()
                    if prev_tx:
                        sug = _classify_by_label(prev_tx)
                        if sug and len(tx) < 55:
                            fmap[k] = sug
                            continue
                # 3) خانات الملاحظات والتوقيعات تُترك للكتابة اليدوية
                if any(w in tx for w in ("........", "ملاحظات الاستشاري:", "التوقيع والختم")):
                    fmap[k] = {"mode": "manual", "label": "يُكتب / يُوقّع يدوياً"}

    if grid.get("kind") == "docx":
        for tb in grid.get("tables") or []:
            _scan_rows(tb.get("rows") or [])
    else:
        for sh in grid.get("sheets") or []:
            _scan_rows([r["cells"] for r in (sh.get("rows") or [])])
    return fmap


def resolve_auto_values(proj: dict, ref_no: str, discipline: str = "EL", revision: str = "R00") -> dict[str, str]:
    import datetime as _dt
    today = _dt.date.today()
    return {
        "project_name": proj.get("name") or "",
        "project_code": proj.get("code") or "",
        "contract_no": proj.get("contract_no") or "",
        "client": proj.get("client") or "",
        "consultant": proj.get("consultant") or "",
        "contractor": proj.get("contractor") or "",
        "auto_ref": ref_no,
        "revision": revision,
        "discipline": discipline,
        "today_iso": today.isoformat(),
        "today_dmy": today.strftime("%d/%m/%Y"),
        "currency": proj.get("currency") or "KWD",
    }


def apply_fixed_edits_to_template(template_path: Path, fixed_edits: dict[str, str]):
    """يحفظ التعديلات النصية الثابتة التي أجراها المستخدم مباشرةً على القالب الأصلي لهذا المشروع."""
    if not fixed_edits:
        return
    ext = template_path.suffix.lower()
    if ext == ".docx":
        cells, hcells, paras = [], [], []
        for k, val in fixed_edits.items():
            parts = k.split(":")
            if parts[0] == "t" and len(parts) == 4:
                cells.append({"table": int(parts[1]), "row": int(parts[2]), "col": int(parts[3]), "value": val})
            elif parts[0] == "h" and len(parts) == 5:
                hcells.append({"section": int(parts[1]), "table": int(parts[2]), "row": int(parts[3]), "col": int(parts[4]), "value": val})
            elif parts[0] == "p" and len(parts) == 2:
                paras.append({"index": int(parts[1]), "value": val})
        fill_docx(template_path, {"cells": cells, "header_cells": hcells, "paragraphs": paras}, template_path)
    else:
        multi: dict[str, dict[str, object]] = {}
        for k, val in fixed_edits.items():
            if k.startswith("x:"):
                _, sh_name, coord = k.split(":", 2)
                multi.setdefault(sh_name, {})[coord] = val
        if multi:
            fill_xlsx(template_path, {"allow_overwrite_formulas": False, "multi_sheet_cells": multi}, template_path)


def apply_visual_map(
    template_path: Path,
    out_path: Path,
    field_map: dict[str, dict],
    auto_values: dict[str, str],
    user_inputs: dict[str, object],
):
    """يطبّق خريطة الخانات المحفوظة (تلقائي / حقل إدخال / كتابة يدوية / ثابت) بدقة 100% دون المساس بالثوابت أو المعادلات أو الشعارات."""
    ext = template_path.suffix.lower()

    def _compute_val(k: str, rule: dict) -> tuple[bool, object]:
        mode = rule.get("mode") or "fixed"
        if mode == "auto":
            ak = str(rule.get("auto_key") or "")
            return True, auto_values.get(ak, "")
        if mode == "input":
            raw = user_inputs.get(k)
            if raw is None or str(raw).strip() == "":
                raw = rule.get("default", "")
            if rule.get("input_type") == "number" and str(raw).strip() != "":
                try:
                    return True, float(str(raw).replace(",", "").strip())
                except Exception:
                    pass
            return True, str(raw)
        if mode == "manual":
            # نفرّغ أي نص نموذجي/Placeholder قديم ونترك الخانة نظيفة ومسطّرة للكتابة اليدوية (إلا إن طلب إبقاء نقاط التوقيع)
            if rule.get("keep_dots"):
                return False, ""
            return True, str(rule.get("hint_text") or "")
        return False, ""

    if ext == ".docx":
        cells, hcells, paras = [], [], []
        for k, rule in (field_map or {}).items():
            ok, val = _compute_val(k, rule)
            if not ok:
                continue
            parts = k.split(":")
            if parts[0] == "t" and len(parts) == 4:
                cells.append({"table": int(parts[1]), "row": int(parts[2]), "col": int(parts[3]), "value": str(val)})
            elif parts[0] == "h" and len(parts) == 5:
                hcells.append({"section": int(parts[1]), "table": int(parts[2]), "row": int(parts[3]), "col": int(parts[4]), "value": str(val)})
            elif parts[0] == "p" and len(parts) == 2:
                paras.append({"index": int(parts[1]), "value": str(val)})
        fill_docx(template_path, {"cells": cells, "header_cells": hcells, "paragraphs": paras}, out_path)
    else:
        multi: dict[str, dict[str, object]] = {}
        for k, rule in (field_map or {}).items():
            ok, val = _compute_val(k, rule)
            if not ok:
                continue
            if k.startswith("x:"):
                _, sh_name, coord = k.split(":", 2)
                multi.setdefault(sh_name, {})[coord] = val
        fill_xlsx(template_path, {"allow_overwrite_formulas": False, "multi_sheet_cells": multi}, out_path)


def main():
    ap = argparse.ArgumentParser(description="Inspect and fill native Word (.docx) & Excel (.xlsx/.xlsm) project forms.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_ins = sub.add_parser("inspect", help="Inspect all fields/tables/cells of a Word or Excel form")
    p_ins.add_argument("file")

    p_fd = sub.add_parser("fill-docx", help="Fill a Word .docx template preserving logos, layout, and RTL")
    p_fd.add_argument("template")
    p_fd.add_argument("edits_json")
    p_fd.add_argument("out_file")

    p_fx = sub.add_parser("fill-xlsx", help="Fill an Excel .xlsx/.xlsm template preserving formulas, merges, logos, and VBA")
    p_fx.add_argument("template")
    p_fx.add_argument("edits_json")
    p_fx.add_argument("out_file")

    p_nr = sub.add_parser("next-ref", help="Detect numbering scheme and get next reference number from Excel Log")
    p_nr.add_argument("log_file")
    p_nr.add_argument("--prefix", default=None, help="Discipline or document prefix, e.g. EL, CIV, FF, OUT, IPC")

    args = ap.parse_args()
    if args.cmd == "inspect":
        p = Path(args.file)
        res = inspect_docx(p) if p.suffix.lower() == ".docx" else inspect_xlsx_template(p)
        print(json.dumps(res, ensure_ascii=False, indent=2))
    elif args.cmd == "fill-docx":
        spec = json.loads(Path(args.edits_json).read_text(encoding="utf-8"))
        fill_docx(Path(args.template), spec, Path(args.out_file))
    elif args.cmd == "fill-xlsx":
        spec = json.loads(Path(args.edits_json).read_text(encoding="utf-8"))
        fill_xlsx(Path(args.template), spec, Path(args.out_file))
    elif args.cmd == "next-ref":
        print(json.dumps(compute_next_ref(Path(args.log_file), args.prefix), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
