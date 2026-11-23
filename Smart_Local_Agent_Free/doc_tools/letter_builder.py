#!/usr/bin/env python3
"""
letter_builder.py — إنشاء كتب ومراسلات رسمية احترافية بصيغة Word (.docx) جاهزة للطباعة والتوقيع.
يدعم:
  - اتجاه عربي حقيقي 100% (RTL XML: w:bidi, w:rtl, w:bidiVisual) أو إنجليزي (LTR).
  - الطباعة على ورق الشركة المروّس الجاهز (--preprinted يترك هامش علوي 4.2 سم للترويسة)
    أو استخدام ملف ترويسة الشركة (--template letterhead.docx)
    أو إنشاء ترويسة رسمية أنيقة تلقائياً.
  - الجداول المنسّقة داخل الكتاب، الإشارة للكتب السابقة، المرفقات، التوقيع، ونسخة إلى (CC).

الاستخدام:
  python letter_builder.py spec.json "كتاب_صادر.docx" [--template letterhead.docx] [--preprinted]
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

NAVY_HEX = "1B365D"
SOFT_HEX = "F1F5F9"
BORDER_HEX = "1E293B"
NAVY_RGB = RGBColor(0x1B, 0x36, 0x5D)
DARK_RGB = RGBColor(0x1E, 0x29, 0x3B)
MUTED_RGB = RGBColor(0x47, 0x55, 0x69)


def set_para_rtl(para, is_rtl: bool = True, align=None, space_before: float = 0, space_after: float = 6, line_spacing: float = 1.35):
    pPr = para._p.get_or_add_pPr()
    if is_rtl:
        bidi = OxmlElement("w:bidi")
        bidi.set(qn("w:val"), "1")
        pPr.append(bidi)
    para.alignment = align if align is not None else (WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT)
    pf = para.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing


def add_styled_run(para, text: str, is_rtl: bool = True, font_name: str = "Arial", size: float = 12, bold: bool = False, underline: bool = False, color: RGBColor = DARK_RGB):
    run = para.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    run.underline = underline
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), font_name)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    if bold:
        bCs = OxmlElement("w:bCs")
        rPr.append(bCs)
    if is_rtl:
        rtl_el = OxmlElement("w:rtl")
        rPr.append(rtl_el)
    return run


def _insert_ordered(parent, elem, order: list[str]):
    loc = elem.tag.split("}")[-1]
    for old in list(parent.findall(qn(f"w:{loc}"))):
        parent.remove(old)
    parent.append(elem)
    rmap = {n: i for i, n in enumerate(order)}
    ch_all = sorted(list(parent), key=lambda el: rmap.get(el.tag.split("}")[-1], len(order)))
    for ch in ch_all:
        parent.remove(ch)
        parent.append(ch)


_TBL_ORD = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]
_TC_ORD = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]


def set_cell_shading(cell, hex_color: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    _insert_ordered(cell._tc.get_or_add_tcPr(), shading, _TC_ORD)


def set_table_borders(table, color: str = BORDER_HEX, sz: str = "6", outer_color: str = "0F172A", outer_sz: str = "10"):
    try:
        table.style = "Table Grid"
    except Exception:
        pass
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:bottom w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:left w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:right w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{outer_color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f"</w:tblBorders>"
    )
    _insert_ordered(tblPr, borders, _TBL_ORD)
    nrows = len(table.rows)
    ncols = len(table.columns)
    for r_i, row in enumerate(table.rows):
        seen_tc = set()
        for c_i, cell in enumerate(row.cells):
            tc = cell._tc
            if id(tc) in seen_tc:
                continue
            seen_tc.add(id(tc))
            last_c = c_i
            while last_c + 1 < ncols and row.cells[last_c + 1]._tc is tc:
                last_c += 1
            tcPr = tc.get_or_add_tcPr()
            t_sz, t_cl = (outer_sz, outer_color) if r_i == 0 else (sz, color)
            b_sz, b_cl = (outer_sz, outer_color) if r_i == nrows - 1 else (sz, color)
            l_sz, l_cl = (outer_sz, outer_color) if c_i == 0 else (sz, color)
            r_sz, r_cl = (outer_sz, outer_color) if last_c == ncols - 1 else (sz, color)
            _insert_ordered(tcPr, parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="{t_sz}" w:space="0" w:color="{t_cl}"/>'
                f'<w:bottom w:val="single" w:sz="{b_sz}" w:space="0" w:color="{b_cl}"/>'
                f'<w:left w:val="single" w:sz="{l_sz}" w:space="0" w:color="{l_cl}"/>'
                f'<w:right w:val="single" w:sz="{r_sz}" w:space="0" w:color="{r_cl}"/>'
                f"</w:tcBorders>"
            ), _TC_ORD)


def set_table_rtl(table, is_rtl: bool = True):
    if is_rtl:
        tblPr = table._tbl.tblPr
        _insert_ordered(tblPr, parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'), _TBL_ORD)


def add_bottom_border(para, color: str = NAVY_HEX, sz: str = "12"):
    pPr = para._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/>'
        f"</w:pBdr>"
    )
    pPr.append(pBdr)


def build_letter(spec: dict, out_path: Path, template_path: Path | None = None, preprinted: bool = False):
    lang = str(spec.get("lang", "ar")).lower()
    is_rtl = lang != "en"
    font_name = spec.get("font") or ("Traditional Arabic" if False else "Arial")

    if template_path and template_path.exists():
        doc = Document(str(template_path))
    else:
        doc = Document()
        sec = doc.sections[0]
        sec.page_width = Cm(21.0)   # A4
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(4.2 if preprinted else 2.0)
        sec.bottom_margin = Cm(2.5 if preprinted else 2.0)
        sec.right_margin = Cm(2.2)
        sec.left_margin = Cm(2.2)
        # تفعيل اتجاه الصفحة RTL للعربي
        if is_rtl:
            sectPr = sec._sectPr
            sectPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))

        # ترويسة الشركة إن لم يكن ورقاً مروّساً جاهزاً
        sender_org = spec.get("sender_org", "")
        if not preprinted and sender_org:
            hp = doc.add_paragraph()
            set_para_rtl(hp, is_rtl, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, line_spacing=1.15)
            add_styled_run(hp, sender_org, is_rtl, font_name, size=15, bold=True, color=NAVY_RGB)
            sender_sub = spec.get("sender_subtitle", "")
            if sender_sub:
                sp = doc.add_paragraph()
                set_para_rtl(sp, is_rtl, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8, line_spacing=1.15)
                add_styled_run(sp, sender_sub, is_rtl, font_name, size=10, color=MUTED_RGB)
                add_bottom_border(sp)
            else:
                add_bottom_border(hp)

    # 1) شريط المرجع والتاريخ والمرفقات
    date_str = spec.get("date") or dt.date.today().strftime("%Y/%m/%d")
    ref_no = spec.get("ref_no") or ("—" if is_rtl else "TBD")
    att_summary = spec.get("attachments_count") or ""

    meta_lines = []
    if is_rtl:
        meta_lines.append(f"الإشارة / الرقم : {ref_no}")
        meta_lines.append(f"التاريـــــــــخ : {date_str}")
        if att_summary:
            meta_lines.append(f"المرفقــــــــات : {att_summary}")
    else:
        meta_lines.append(f"Ref. No. : {ref_no}")
        meta_lines.append(f"Date     : {date_str}")
        if att_summary:
            meta_lines.append(f"Encl.    : {att_summary}")

    for idx, line in enumerate(meta_lines):
        mp = doc.add_paragraph()
        set_para_rtl(mp, is_rtl, space_before=(6 if idx == 0 else 0), space_after=(10 if idx == len(meta_lines) - 1 else 2), line_spacing=1.2)
        add_styled_run(mp, line, is_rtl, font_name, size=11, bold=True, color=DARK_RGB)

    # 2) المرسل إليه (السادة / المحترمين)
    to_org = spec.get("to_org", "")
    honorific = spec.get("to_honorific", "المحترمين" if is_rtl else "")
    if to_org:
        tp = doc.add_paragraph()
        set_para_rtl(tp, is_rtl, space_before=4, space_after=3, line_spacing=1.3)
        prefix = spec.get("to_prefix", "السادة / " if is_rtl else "To: ")
        main_txt = f"{prefix}{to_org}"
        if honorific:
            main_txt += f"    {honorific}"
        add_styled_run(tp, main_txt, is_rtl, font_name, size=13, bold=True, color=NAVY_RGB)

    attention = spec.get("attention", "")
    if attention:
        ap = doc.add_paragraph()
        set_para_rtl(ap, is_rtl, space_after=6, line_spacing=1.25)
        lbl = "عناية: " if is_rtl else "Attention: "
        add_styled_run(ap, f"{lbl}{attention}", is_rtl, font_name, size=11.5, bold=True, color=DARK_RGB)

    # التحية
    greeting = spec.get("greeting") or ("تحية طيبة وبعد،،،" if is_rtl else "Dear Sirs,")
    gp = doc.add_paragraph()
    set_para_rtl(gp, is_rtl, space_before=4, space_after=8)
    add_styled_run(gp, greeting, is_rtl, font_name, size=12, bold=False)

    # 3) كتلة المشروع / العقد / الموضوع / الإشارة
    project = spec.get("project", "")
    contract_no = spec.get("contract_no", "")
    subject = spec.get("subject", "")
    ref_letter = spec.get("reference_letter", "")

    if project:
        pp = doc.add_paragraph()
        set_para_rtl(pp, is_rtl, align=WD_ALIGN_PARAGRAPH.CENTER if is_rtl else WD_ALIGN_PARAGRAPH.LEFT, space_after=2, line_spacing=1.25)
        lbl = "المشروع: " if is_rtl else "Project: "
        txt = f"{lbl}{project}" + (f"  (عقد رقم: {contract_no})" if (contract_no and is_rtl) else (f"  (Contract No.: {contract_no})" if contract_no else ""))
        add_styled_run(pp, txt, is_rtl, font_name, size=11.5, bold=True, color=NAVY_RGB)

    if subject:
        subp = doc.add_paragraph()
        set_para_rtl(subp, is_rtl, align=WD_ALIGN_PARAGRAPH.CENTER if is_rtl else WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=6, line_spacing=1.3)
        lbl = "الموضوع: " if is_rtl else "Subject: "
        add_styled_run(subp, f"{lbl}{subject}", is_rtl, font_name, size=12.5, bold=True, underline=True, color=NAVY_RGB)

    if ref_letter:
        rp = doc.add_paragraph()
        set_para_rtl(rp, is_rtl, space_after=8, line_spacing=1.3)
        lbl = "المرجع: " if is_rtl else "Reference: "
        add_styled_run(rp, lbl, is_rtl, font_name, size=11, bold=True, color=MUTED_RGB)
        add_styled_run(rp, ref_letter, is_rtl, font_name, size=11, bold=False, color=MUTED_RGB)

    # 4) فقرات المتن
    for p_text in spec.get("paragraphs", []):
        bp = doc.add_paragraph()
        set_para_rtl(bp, is_rtl, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, line_spacing=1.4)
        add_styled_run(bp, str(p_text), is_rtl, font_name, size=12)

    # نقاط مرقمة إن وجدت
    bullets = spec.get("bullet_points") or []
    for idx, b_text in enumerate(bullets, start=1):
        blp = doc.add_paragraph()
        set_para_rtl(blp, is_rtl, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=4, line_spacing=1.35)
        if is_rtl:
            blp.paragraph_format.right_indent = Cm(0.6)
        else:
            blp.paragraph_format.left_indent = Cm(0.6)
        add_styled_run(blp, f"{idx}- ", is_rtl, font_name, size=11.5, bold=True, color=NAVY_RGB)
        add_styled_run(blp, str(b_text), is_rtl, font_name, size=11.5)

    # 5) جدول اختياري داخل الكتاب (مثلاً كشف مرفقات / مطالبات / بنود / اعتمادات متأخرة)
    tbl_spec = spec.get("table")
    if isinstance(tbl_spec, dict) and tbl_spec.get("headers") and tbl_spec.get("rows"):
        headers = tbl_spec["headers"]
        rows = tbl_spec["rows"]
        tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_rtl(tbl, is_rtl)
        set_table_borders(tbl)

        for c_i, h in enumerate(headers):
            cell = tbl.cell(0, c_i)
            set_cell_shading(cell, NAVY_HEX)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cp = cell.paragraphs[0]
            set_para_rtl(cp, is_rtl, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=3, space_after=3, line_spacing=1.15)
            add_styled_run(cp, str(h), is_rtl, font_name, size=10.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

        for r_i, r_vals in enumerate(rows, start=1):
            for c_i in range(len(headers)):
                val = r_vals[c_i] if c_i < len(r_vals) else ""
                cell = tbl.cell(r_i, c_i)
                if r_i % 2 == 0:
                    set_cell_shading(cell, SOFT_HEX)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                cp = cell.paragraphs[0]
                set_para_rtl(cp, is_rtl, align=WD_ALIGN_PARAGRAPH.CENTER if c_i == 0 else (WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT), space_before=2, space_after=2, line_spacing=1.15)
                add_styled_run(cp, str(val), is_rtl, font_name, size=10.5)

        sp_after = doc.add_paragraph()
        set_para_rtl(sp_after, is_rtl, space_after=4)

    # فقرة ختامية بعد الجدول إن وجدت
    closing_para = spec.get("closing_paragraph", "")
    if closing_para:
        cp = doc.add_paragraph()
        set_para_rtl(cp, is_rtl, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=4, space_after=8, line_spacing=1.4)
        add_styled_run(cp, closing_para, is_rtl, font_name, size=12)

    # 6) الخاتمة الرسمية والتوقيع
    valediction = spec.get("closing") or ("وتفضلوا بقبول فائق الاحترام والتقدير،،،" if is_rtl else "Yours faithfully,")
    vp = doc.add_paragraph()
    set_para_rtl(vp, is_rtl, align=WD_ALIGN_PARAGRAPH.CENTER if is_rtl else WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=16)
    add_styled_run(vp, valediction, is_rtl, font_name, size=12, bold=True)

    signer_title = spec.get("signer_title", "")
    signer_name = spec.get("signer_name", "")
    signer_org = spec.get("signer_org") or spec.get("sender_org", "")
    sig_align = WD_ALIGN_PARAGRAPH.LEFT

    for s_line in [x for x in [signer_org, signer_title, signer_name] if x]:
        sgp = doc.add_paragraph()
        set_para_rtl(sgp, is_rtl, align=sig_align, space_after=2, line_spacing=1.2)
        if is_rtl:
            sgp.paragraph_format.left_indent = Cm(1.0)
        add_styled_run(sgp, s_line, is_rtl, font_name, size=11.5, bold=True, color=NAVY_RGB)

    # 7) المرفقات ونسخة إلى (CC)
    attachments = spec.get("attachments") or []
    if attachments:
        at_h = doc.add_paragraph()
        set_para_rtl(at_h, is_rtl, space_before=14, space_after=2)
        add_styled_run(at_h, "المرفقات:" if is_rtl else "Attachments:", is_rtl, font_name, size=10, bold=True, underline=True, color=MUTED_RGB)
        for att in attachments:
            at_p = doc.add_paragraph()
            set_para_rtl(at_p, is_rtl, space_after=1, line_spacing=1.15)
            add_styled_run(at_p, f"- {att}", is_rtl, font_name, size=9.5, color=MUTED_RGB)

    cc_list = spec.get("cc") or []
    if cc_list:
        cc_h = doc.add_paragraph()
        set_para_rtl(cc_h, is_rtl, space_before=8, space_after=2)
        add_styled_run(cc_h, "نسخة إلى:" if is_rtl else "CC:", is_rtl, font_name, size=10, bold=True, underline=True, color=MUTED_RGB)
        for c_item in cc_list:
            cc_p = doc.add_paragraph()
            set_para_rtl(cc_p, is_rtl, space_after=1, line_spacing=1.15)
            add_styled_run(cc_p, f"- {c_item}", is_rtl, font_name, size=9.5, color=MUTED_RGB)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"[OK] تم إنشاء الكتاب الرسمي الجاهز للطباعة: {out_path}")


def main():
    ap = argparse.ArgumentParser(description="Build professional RTL Arabic / English official letters (.docx)")
    ap.add_argument("spec_json", help="Path to JSON letter specification")
    ap.add_argument("out_docx", help="Output .docx file path")
    ap.add_argument("--template", default=None, help="Optional company letterhead .docx")
    ap.add_argument("--preprinted", action="store_true", help="Leave 4.2cm top margin for pre-printed company letterhead")
    args = ap.parse_args()

    spec = json.loads(Path(args.spec_json).read_text(encoding="utf-8"))
    tpl = Path(args.template) if args.template else None
    # استخدام ترويسة الشركة المحفوظة تلقائياً إن وجدت
    if tpl is None and not args.preprinted:
        default_tpl = Path.home() / ".gemini" / "templates" / "letterhead.docx"
        if default_tpl.exists():
            tpl = default_tpl
    build_letter(spec, Path(args.out_docx), template_path=tpl, preprinted=args.preprinted)


if __name__ == "__main__":
    main()
