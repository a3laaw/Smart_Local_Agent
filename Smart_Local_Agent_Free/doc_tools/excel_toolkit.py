#!/usr/bin/env python3
"""
excel_toolkit.py — فحص + نسخ احتياطي + إنشاء قوالب سجلات Excel احترافية عربية (RTL) بالمعادلات والروابط
مع دعم كامل للحفاظ على الماكرو في ملفات .xlsm.
"""
import argparse
import datetime as dt
import json
import os
import shutil
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY = "1B365D"
SLATE = "2C5282"
SOFT_BG = "F0F4F8"
ZEBRA = "F8FAFC"
BORDER_CLR = "1E293B"
THIN = Side(style="thin", color=BORDER_CLR)
MEDIUM_OUT = Side(style="medium", color="0F172A")
GRID_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _apply_box_frame(ws, r1: int, c1: int, r2: int, c2: int):
    """يرسم إطاراً خارجياً سميكاً واضحاً حول الجدول مع تسطير داخلي كامل لكل خلية (للطباعة الرسمية الواضحة 100%)."""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cb = cell.border
            cell.border = Border(
                top=MEDIUM_OUT if r == r1 else (cb.top if (cb and cb.top and cb.top.style) else THIN),
                bottom=MEDIUM_OUT if r == r2 else (cb.bottom if (cb and cb.bottom and cb.bottom.style) else THIN),
                left=MEDIUM_OUT if c == c1 else (cb.left if (cb and cb.left and cb.left.style) else THIN),
                right=MEDIUM_OUT if c == c2 else (cb.right if (cb and cb.right and cb.right.style) else THIN),
            )


def _finalize_printable_borders(ws):
    """يضمن اكتمال التسطير الرسمي الواضح حول كل الخلايا المدمجة وكل الجداول وعند الطباعة."""
    ws.views.sheetView[0].showGridLines = True
    ws.print_options.gridLines = False
    for mr in list(ws.merged_cells.ranges):
        tl = ws.cell(mr.min_row, mr.min_col)
        tl_b = tl.border
        has_b = bool(tl_b and any(getattr(tl_b, s) and getattr(tl_b, s).style for s in ("top", "bottom", "left", "right")))
        has_f = bool(tl.fill and tl.fill.fill_type == "solid" and tl.fill.fgColor and str(tl.fill.fgColor.rgb or "") not in ("", "00000000", "FFFFFFFF"))
        if not (has_b or has_f or (tl.value is not None and str(tl.value).strip() != "")):
            continue
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                ws.cell(r, c).border = Border(
                    top=THIN if r == mr.min_row else None,
                    bottom=THIN if r == mr.max_row else None,
                    left=THIN if c == mr.min_col else None,
                    right=THIN if c == mr.max_col else None,
                )
    # إطار خارجي واضح لكل كتلة متصلة من الصفوف المسطرة
    max_r, max_c = ws.max_row, ws.max_column
    block_start = None
    for r in range(1, max_r + 2):
        row_cols = []
        if r <= max_r:
            for c in range(1, max_c + 1):
                b = ws.cell(r, c).border
                if b and any(getattr(b, s) and getattr(b, s).style for s in ("top", "bottom", "left", "right")):
                    row_cols.append(c)
        if row_cols and block_start is None:
            block_start = (r, min(row_cols), max(row_cols))
        elif row_cols and block_start is not None:
            block_start = (block_start[0], min(block_start[1], min(row_cols)), max(block_start[2], max(row_cols)))
        elif not row_cols and block_start is not None:
            _apply_box_frame(ws, block_start[0], block_start[1], r - 1, block_start[2])
            block_start = None


def inspect_workbook(path: Path) -> dict:
    has_vba = False
    with zipfile.ZipFile(path, "r") as zf:
        has_vba = any("vbaProject.bin" in n for n in zf.namelist())

    wb_f = openpyxl.load_workbook(str(path), data_only=False, keep_vba=has_vba)
    wb_v = openpyxl.load_workbook(str(path), data_only=True)
    info = {
        "file": str(path),
        "ext": path.suffix.lower(),
        "has_vba_macros": has_vba,
        "sheet_count": len(wb_f.sheetnames),
        "sheets": [],
    }
    for name in wb_f.sheetnames:
        ws = wb_f[name]
        wsv = wb_v[name]
        headers = []
        header_row = None
        for r in range(1, min(12, ws.max_row + 1)):
            vals = [wsv.cell(r, c).value for c in range(1, min(40, ws.max_column + 1))]
            non_empty = [v for v in vals if v is not None and str(v).strip() != ""]
            if len(non_empty) >= 4:
                header_row = r
                headers = [
                    {"col": get_column_letter(c), "title": str(wsv.cell(r, c).value).strip()}
                    for c in range(1, min(40, ws.max_column + 1))
                    if wsv.cell(r, c).value is not None and str(wsv.cell(r, c).value).strip() != ""
                ]
                break
        formula_cols = set()
        if header_row and header_row < ws.max_row:
            for r in range(header_row + 1, min(header_row + 6, ws.max_row + 1)):
                for c in range(1, min(40, ws.max_column + 1)):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.startswith("="):
                        formula_cols.add(get_column_letter(c))
        info["sheets"].append({
            "name": name,
            "state": ws.sheet_state,
            "rows": ws.max_row,
            "cols": ws.max_column,
            "header_row": header_row,
            "headers": headers,
            "formula_columns": sorted(formula_cols),
        })
    return info


def backup_workbook(path: Path) -> Path:
    bdir = path.parent / "_Backups"
    bdir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = bdir / f"{path.stem}_BACKUP_{stamp}{path.suffix}"
    shutil.copy2(path, dest)
    return dest


TEMPLATES = {
    "petty_cash": {
        "title": "كشف تسوية العُهدة والمصروفات النقدية (Petty Cash Settlement)",
        "sheet": "تسوية العهدة",
        "headers": [
            ("م", 6), ("التاريخ", 13), ("رقم الفاتورة", 15), ("المورد / الجهة", 26),
            ("البيان / تفاصيل الصرف", 36), ("بند الصرف", 18), ("المشروع / القسم", 16),
            ("طريقة الدفع", 13), ("المبلغ (د.ك)", 15), ("حالة الفاتورة", 15),
            ("ملاحظات المراجعة", 26), ("فتح الفاتورة", 14),
        ],
        "money_cols": [9],
        "num_fmt": "#,##0.000",
        "kpis": [
            ("قيمة العُهدة المستلمة (د.ك)", "0"),
            ("إجمالي المصروفات المعتمدة (د.ك)", "=SUM(I6:I500)"),
            ("المتبقي / (العجز) في العُهدة (د.ك)", "=B3-D3"),
            ("عدد الفواتير المسجلة", "=COUNTA(D6:D500)"),
        ],
    },
    "invoices": {
        "title": "سجل مراجعة وأرشفة فواتير الموردين والمقاولين (Invoices Register)",
        "sheet": "سجل الفواتير",
        "headers": [
            ("م", 6), ("رقم الفاتورة", 16), ("تاريخ الفاتورة", 13), ("المورد / المقاول", 28),
            ("المشروع", 18), ("رقم أمر الشراء PO", 16), ("البيان", 34),
            ("الإجمالي قبل الخصم", 16), ("الخصم / الضريبة", 15), ("الصافي المستحق", 16),
            ("حالة السداد", 15), ("ملاحظات", 24), ("فتح الفاتورة", 14),
        ],
        "money_cols": [8, 9, 10],
        "num_fmt": "#,##0.000",
        "kpis": [
            ("عدد الفواتير", "=COUNTA(B6:B1000)"),
            ("إجمالي الصافي المستحق", "=SUM(J6:J1000)"),
            ("مسدد", '=SUMIF(K6:K1000,"مسدد",J6:J1000)'),
            ("قيد السداد / معلق", "=D3-F3"),
        ],
    },
    "letters": {
        "title": "سجل الكتب والمراسلات الرسمية — صادر ووارد (Correspondence Log)",
        "sheet": "الصادر والوارد",
        "headers": [
            ("م", 6), ("النوع (صادر/وارد)", 15), ("رقم الكتاب", 18), ("تاريخ الكتاب", 13),
            ("الجهة المرسلة", 24), ("الجهة المرسل إليها", 24), ("المشروع", 18),
            ("الموضوع / الملخص", 38), ("إشارة لكتاب سابق", 18), ("يتطلب رد؟", 12),
            ("آخر موعد للرد", 14), ("الحالة", 15), ("ملاحظات", 24), ("فتح الكتاب", 14),
        ],
        "money_cols": [],
        "num_fmt": "#,##0",
        "kpis": [
            ("إجمالي الكتب", "=COUNTA(C6:C1000)"),
            ("كتب الوارد", '=COUNTIF(B6:B1000,"وارد")'),
            ("كتب الصادر", '=COUNTIF(B6:B1000,"صادر")'),
            ("بانتظار الرد", '=COUNTIF(L6:L1000,"بانتظار الرد")'),
        ],
    },
    "contracts": {
        "title": "سجل العقود والمستخلصات والأوامر التغييرية (Contracts & IPCs Tracker)",
        "sheet": "العقود والمستخلصات",
        "headers": [
            ("م", 6), ("رقم العقد", 16), ("المقاول / الطرف الثاني", 26), ("نطاق الأعمال", 30),
            ("نوع المستند", 16), ("رقم المستخلص / VO", 16), ("التاريخ / الفترة", 15),
            ("قيمة العقد المعدلة", 17), ("إجمالي الأعمال المنفذة", 18), ("استقطاع ضمان Retention", 18),
            ("استرداد دفعة مقدمة / خصم", 18), ("صافي المستخلص الحالي", 18),
            ("حالة الاعتماد / الصرف", 16), ("ملاحظات المراجعة", 26), ("فتح المستند", 14),
        ],
        "money_cols": [8, 9, 10, 11, 12],
        "num_fmt": "#,##0.000",
        "kpis": [
            ("عدد السجلات", "=COUNTA(C6:C1000)"),
            ("إجمالي صافي المستخلصات", "=SUM(L6:L1000)"),
            ("إجمالي الاستقطاعات المحتجزة", "=SUM(J6:J1000)"),
            ("قيد المراجعة", '=COUNTIF(M6:M1000,"قيد المراجعة")'),
        ],
    },
    "submittals": {
        "title": "سجل متابعة الاعتمادات الهندسية (Submittal Control Log)",
        "sheet": "Submittals Log",
        "headers": [
            ("م", 6), ("كود الاعتماد", 16), ("Rev", 7), ("القسم", 12),
            ("نوع الاعتماد", 15), ("وصف المادة / البند", 36), ("المورد / المصنّع / المقاول", 26),
            ("تاريخ التقديم", 13), ("تاريخ رد الاستشاري", 15), ("كود الاستشاري", 14),
            ("حالة الجهة المالكة", 16), ("تاريخ رد المالك", 14), ("الحالة النهائية", 15),
            ("ملاحظات", 26), ("ملف الاستشاري", 14), ("ملف المالك", 14),
        ],
        "money_cols": [],
        "num_fmt": "#,##0",
        "kpis": [
            ("إجمالي الاعتمادات", "=COUNTA(B6:B2000)"),
            ("معتمد (Approved)", '=COUNTIF(M6:M2000,"Approved")+COUNTIF(M6:M2000,"معتمد")'),
            ("مرفوض (Rejected)", '=COUNTIF(M6:M2000,"Rejected")+COUNTIF(M6:M2000,"مرفوض")'),
            ("قيد المراجعة / متأخر", '=COUNTIF(M6:M2000,"Under Review")+COUNTIF(M6:M2000,"Overdue")+COUNTIF(M6:M2000,"قيد المراجعة")'),
        ],
    },
    "amendments": {
        "title": "تقرير التعديلات والمطابقة وتنظيف الأرشيف (Amendments & Audit Report)",
        "sheet": "سجل التعديلات",
        "headers": [
            ("م", 6), ("الشيت / القسم", 14), ("المرجع / الكود", 16), ("نوع التعديل", 18),
            ("الحقل المعدَّل", 18), ("القيمة القديمة", 24), ("القيمة الجديدة", 24),
            ("المستند المرجعي", 26), ("السبب / الملاحظة", 34),
        ],
        "money_cols": [],
        "num_fmt": "#,##0",
        "kpis": [
            ("إجمالي التعديلات الموثّقة", "=COUNTA(C6:C2000)"),
            ("سجلات جديدة مضافة", '=COUNTIF(D6:D2000,"إضافة جديد")'),
            ("تحديثات حالة/تاريخ", '=COUNTIF(D6:D2000,"تحديث")'),
            ("تصحيح تسمية/قسم", '=COUNTIF(D6:D2000,"تصحيح")'),
        ],
    },
}


def create_template(kind: str, out_path: Path):
    spec = TEMPLATES[kind]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = spec["sheet"]
    ws.views.sheetView[0].rightToLeft = True
    ws.views.sheetView[0].showGridLines = True

    ncols = len(spec["headers"])
    # Row 1: Title banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c1 = ws.cell(1, 1, spec["title"])
    c1.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    c1.fill = PatternFill("solid", fgColor=NAVY)
    c1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    # Row 3: Live KPI bar
    ws.row_dimensions[2].height = 8
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 8
    col_ptr = 1
    for label, formula in spec["kpis"]:
        lbl_cell = ws.cell(3, col_ptr, label)
        lbl_cell.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        lbl_cell.fill = PatternFill("solid", fgColor=SOFT_BG)
        lbl_cell.alignment = Alignment(horizontal="center", vertical="center")
        lbl_cell.border = GRID_BORDER

        val_val = float(formula) if not str(formula).startswith("=") else formula
        val_cell = ws.cell(3, col_ptr + 1, val_val)
        val_cell.font = Font(name="Calibri", size=11, bold=True, color=SLATE)
        val_cell.fill = PatternFill("solid", fgColor="FFFFFF")
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.border = GRID_BORDER
        if spec["money_cols"] and "عدد" not in label:
            val_cell.number_format = spec["num_fmt"]
        col_ptr += 2

    # Row 5: Table headers
    ws.row_dimensions[5].height = 28
    for idx, (h_title, width) in enumerate(spec["headers"], start=1):
        cell = ws.cell(5, idx, h_title)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=SLATE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = GRID_BORDER
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:{get_column_letter(ncols)}5"
    wb.save(str(out_path))
    print(f"[OK] تم إنشاء السجل الاحترافي: {out_path} ({spec['sheet']})")


def make_hyperlink_formula(excel_path: Path, target_file: Path, label: str = "فتح PDF") -> str:
    rel = os.path.relpath(target_file.resolve(), excel_path.resolve().parent).replace("\\", "/")
    safe_rel = rel.replace('"', '""')
    safe_lbl = label.replace('"', '""')
    return f'=HYPERLINK("{safe_rel}","{safe_lbl}")'


# =====================================================================
# نماذج رسمية جاهزة للطباعة A4 بالمعادلات وخانات التوقيع (create-form)
# =====================================================================
def _setup_a4_form(ws, title_ar: str, title_en: str, ncols: int, landscape: bool = False):
    ws.views.sheetView[0].rightToLeft = True
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if landscape else ws.ORIENTATION_PORTRAIT
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c1 = ws.cell(1, 1, f"{title_ar}  |  {title_en}")
    c1.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c1.fill = PatternFill("solid", fgColor=NAVY)
    c1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34


def _draw_info_grid(ws, start_row: int, pairs: list[tuple[str, str]], ncols: int) -> int:
    """يرسم جدول بيانات المشروع والمرجع أعلى النموذج مع دمج كافٍ يمنع قص العناوين تماماً."""
    half = max(ncols // 2, 3)
    lbl1_end = 2 if ws.column_dimensions["A"].width < 16 else 1
    val1_start = lbl1_end + 1
    r_lbl_start = half + 1
    r_lbl_end = min(r_lbl_start + 1, ncols - 1) if ncols >= 8 else r_lbl_start
    r_val_start = r_lbl_end + 1

    r = start_row
    for i in range(0, len(pairs), 2):
        ws.row_dimensions[r].height = 23
        left_lbl, left_val = pairs[i]
        if lbl1_end > 1:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=lbl1_end)
        for cc in range(1, lbl1_end + 1):
            ws.cell(r, cc).fill = PatternFill("solid", fgColor=SOFT_BG)
        ws.cell(r, 1, left_lbl).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        ws.cell(r, 1).alignment = Alignment(horizontal="right", vertical="center")

        if val1_start <= half:
            if val1_start < half:
                ws.merge_cells(start_row=r, start_column=val1_start, end_row=r, end_column=half)
            ws.cell(r, val1_start, left_val).alignment = Alignment(horizontal="right", vertical="center")
            ws.cell(r, val1_start).font = Font(name="Calibri", size=10, color="1E293B")

        if i + 1 < len(pairs) and r_lbl_start <= ncols:
            right_lbl, right_val = pairs[i + 1]
            if r_lbl_end > r_lbl_start:
                ws.merge_cells(start_row=r, start_column=r_lbl_start, end_row=r, end_column=r_lbl_end)
            for cc in range(r_lbl_start, r_lbl_end + 1):
                ws.cell(r, cc).fill = PatternFill("solid", fgColor=SOFT_BG)
            ws.cell(r, r_lbl_start, right_lbl).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
            ws.cell(r, r_lbl_start).alignment = Alignment(horizontal="right", vertical="center")

            if r_val_start <= ncols:
                if r_val_start < ncols:
                    ws.merge_cells(start_row=r, start_column=r_val_start, end_row=r, end_column=ncols)
                ws.cell(r, r_val_start, right_val).alignment = Alignment(horizontal="right", vertical="center")
                ws.cell(r, r_val_start).font = Font(name="Calibri", size=10, bold=("رقم" in right_lbl), color="1E293B")

        for c in range(1, ncols + 1):
            ws.cell(r, c).border = GRID_BORDER
        r += 1
    return r + 1


def _draw_signatures(ws, start_row: int, roles: list[str], ncols: int):
    step = max(ncols // len(roles), 2)
    ws.row_dimensions[start_row].height = 24
    ws.row_dimensions[start_row + 1].height = 52
    for idx, role in enumerate(roles):
        c1 = idx * step + 1
        c2 = ncols if idx == len(roles) - 1 else min((idx + 1) * step, ncols)
        ws.merge_cells(start_row=start_row, start_column=c1, end_row=start_row, end_column=c2)
        ws.merge_cells(start_row=start_row + 1, start_column=c1, end_row=start_row + 1, end_column=c2)
        h = ws.cell(start_row, c1, role)
        h.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
        h.fill = PatternFill("solid", fgColor=SLATE)
        h.alignment = Alignment(horizontal="center", vertical="center")
        b = ws.cell(start_row + 1, c1, "الاسم: ............................\nالتوقيع والختم: ................   التاريخ: ..../..../2026")
        b.font = Font(name="Calibri", size=9.5, color="334155")
        b.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        for rr in (start_row, start_row + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).border = GRID_BORDER


def create_printable_form(kind: str, out_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active

    if kind == "ipc":
        ws.title = "مستخلص جاري IPC"
        cols = [
            ("م", 6), ("بيان الأعمال / البند", 30), ("الوحدة", 9), ("كمية العقد", 12),
            ("الفئة (د.ك)", 13), ("إجمالي العقد", 15), ("كمية سابقة", 12), ("كمية حالية", 12),
            ("إجمالي المنفذ", 13), ("نسبة الإنجاز", 11), ("قيمة المنفذ الكلي", 16), ("قيمة الحالي (د.ك)", 16),
        ]
        _setup_a4_form(ws, "مستخلص أعمال جاري", "Interim Payment Certificate (IPC)", len(cols), landscape=True)
        for i, (_, w) in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        r = _draw_info_grid(ws, 3, [
            ("اسم المشروع:", ""), ("رقم المستخلص:", "IPC-01"),
            ("المقاول / الطرف الثاني:", ""), ("تاريخ المستخلص:", dt.date.today().isoformat()),
            ("رقم العقد:", ""), ("الفترة من / إلى:", ""),
        ], len(cols))

        ws.row_dimensions[r].height = 28
        for i, (h, _) in enumerate(cols, 1):
            c = ws.cell(r, i, h)
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=SLATE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = GRID_BORDER

        first_data = r + 1
        last_data = first_data + 11
        for dr in range(first_data, last_data + 1):
            ws.row_dimensions[dr].height = 21
            ws.cell(dr, 1, dr - first_data + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(dr, 6, f"=IF(OR(D{dr}=\"\",E{dr}=\"\"),0,D{dr}*E{dr})").number_format = "#,##0.000"
            ws.cell(dr, 9, f"=SUM(G{dr}:H{dr})").number_format = "#,##0.00"
            ws.cell(dr, 10, f"=IFERROR(IF(D{dr}>0,I{dr}/D{dr},0),0)").number_format = "0.0%"
            ws.cell(dr, 11, f"=I{dr}*E{dr}").number_format = "#,##0.000"
            ws.cell(dr, 12, f"=H{dr}*E{dr}").number_format = "#,##0.000"
            ws.cell(dr, 5).number_format = "#,##0.000"
            for ci in range(1, len(cols) + 1):
                ws.cell(dr, ci).border = GRID_BORDER

        tr = last_data + 1
        ws.row_dimensions[tr].height = 24
        ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=5)
        ws.cell(tr, 1, "الإجمــــالــــي").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        ws.cell(tr, 1).alignment = Alignment(horizontal="center", vertical="center")
        for col_i in (6, 11, 12):
            col_l = get_column_letter(col_i)
            tc = ws.cell(tr, col_i, f"=SUM({col_l}{first_data}:{col_l}{last_data})")
            tc.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
            tc.number_format = "#,##0.000"
        for ci in range(1, len(cols) + 1):
            ws.cell(tr, ci).fill = PatternFill("solid", fgColor=SOFT_BG)
            ws.cell(tr, ci).border = GRID_BORDER

        # ملخص الاستقطاعات وصافي المستحق
        sr = tr + 2
        summary_lines = [
            ("إجمالي الأعمال المنفذة في هذا المستخلص (د.ك)", f"=L{tr}", False),
            ("يُخصم: استقطاع ضمان حسن التنفيذ Retention (10%)", f"=ROUND(L{sr}*0.10,3)", False),
            ("يُخصم: استرداد الدفعة المقدمة (10%)", f"=ROUND(L{sr}*0.10,3)", False),
            ("يُخصم: خصومات / غرامات أخرى", 0.0, False),
            ("صافي المبلغ المستحق للصرف في هذا المستخلص (د.ك)", f"=L{sr}-L{sr+1}-L{sr+2}-L{sr+3}", True),
        ]
        for off, (lbl, fml, is_net) in enumerate(summary_lines):
            rr = sr + off
            ws.row_dimensions[rr].height = 23
            ws.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=11)
            lc = ws.cell(rr, 7, lbl)
            vc = ws.cell(rr, 12, fml)
            lc.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF" if is_net else NAVY)
            vc.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF" if is_net else NAVY)
            vc.number_format = "#,##0.000"
            lc.alignment = Alignment(horizontal="right", vertical="center")
            vc.alignment = Alignment(horizontal="center", vertical="center")
            for cc in range(7, 13):
                ws.cell(rr, cc).fill = PatternFill("solid", fgColor=NAVY if is_net else SOFT_BG)
                ws.cell(rr, cc).border = GRID_BORDER

        _draw_signatures(ws, sr + len(summary_lines) + 2, ["إعداد المقاول (Contractor)", "مراجعة الاستشاري (Consultant)", "اعتماد المالك / الإدارة (Client)"], len(cols))

    elif kind == "po":
        ws.title = "أمر شراء PO"
        cols = [("م", 6), ("كود الصنف", 14), ("وصف المادة / البند", 34), ("الوحدة", 10), ("الكمية", 11), ("سعر الوحدة (د.ك)", 15), ("الإجمالي (د.ك)", 16)]
        _setup_a4_form(ws, "أمر شراء رسمي", "Purchase Order (PO)", len(cols))
        for i, (_, w) in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        r = _draw_info_grid(ws, 3, [
            ("السادة المورد:", ""), ("رقم أمر الشراء:", "PO-2026-001"),
            ("المشروع / موقع التوريد:", ""), ("التاريخ:", dt.date.today().isoformat()),
            ("عرض السعر المرجعي:", ""), ("شروط الدفع والتوريد:", ""),
        ], len(cols))
        ws.row_dimensions[r].height = 26
        for i, (h, _) in enumerate(cols, 1):
            c = ws.cell(r, i, h)
            c.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=SLATE)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = GRID_BORDER
        f_r, l_r = r + 1, r + 14
        for dr in range(f_r, l_r + 1):
            ws.row_dimensions[dr].height = 21
            ws.cell(dr, 1, dr - f_r + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(dr, 6).number_format = "#,##0.000"
            ws.cell(dr, 7, f"=IF(OR(E{dr}=\"\",F{dr}=\"\"),0,E{dr}*F{dr})").number_format = "#,##0.000"
            for ci in range(1, len(cols) + 1):
                ws.cell(dr, ci).border = GRID_BORDER
        sr = l_r + 1
        for off, (lbl, fml, is_net) in enumerate([
            ("الإجمالي قبل الخصم (د.ك)", f"=SUM(G{f_r}:G{l_r})", False),
            ("الخصم (د.ك)", 0.0, False),
            ("الصافي الإجمالي لأمر الشراء (د.ك)", f"=G{sr}-G{sr+1}", True),
        ]):
            rr = sr + off
            ws.row_dimensions[rr].height = 23
            ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=6)
            ws.cell(rr, 4, lbl).font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF" if is_net else NAVY)
            ws.cell(rr, 7, fml).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF" if is_net else NAVY)
            ws.cell(rr, 7).number_format = "#,##0.000"
            for cc in range(4, 8):
                ws.cell(rr, cc).fill = PatternFill("solid", fgColor=NAVY if is_net else SOFT_BG)
                ws.cell(rr, cc).border = GRID_BORDER
        _draw_signatures(ws, sr + 5, ["إدارة المشتريات", "مدير المشروع", "الاعتماد المالي / المدير العام"], len(cols))

    elif kind == "vo":
        ws.title = "أمر تغييري VO"
        cols = [("م", 6), ("نوع التغيير (إضافة/حذف)", 18), ("بيان البند", 32), ("الوحدة", 10), ("الكمية (+/-)", 13), ("سعر الوحدة (د.ك)", 15), ("القيمة الصافية (د.ك)", 17)]
        _setup_a4_form(ws, "نموذج أمر تغييري", "Variation Order (VO)", len(cols))
        for i, (_, w) in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        r = _draw_info_grid(ws, 3, [
            ("المشروع:", ""), ("رقم الأمر التغييري:", "VO-01"),
            ("المقاول:", ""), ("التاريخ:", dt.date.today().isoformat()),
            ("قيمة العقد الأصلية (د.ك):", "0.000"), ("الأوامر التغييرية السابقة (د.ك):", "0.000"),
        ], len(cols))
        ws.row_dimensions[r].height = 26
        for i, (h, _) in enumerate(cols, 1):
            c = ws.cell(r, i, h)
            c.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=SLATE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = GRID_BORDER
        f_r, l_r = r + 1, r + 10
        for dr in range(f_r, l_r + 1):
            ws.row_dimensions[dr].height = 21
            ws.cell(dr, 1, dr - f_r + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(dr, 6).number_format = "#,##0.000"
            ws.cell(dr, 7, f"=IF(OR(E{dr}=\"\",F{dr}=\"\"),0,E{dr}*F{dr})").number_format = "#,##0.000"
            for ci in range(1, len(cols) + 1):
                ws.cell(dr, ci).border = GRID_BORDER
        sr = l_r + 1
        for off, (lbl, fml, is_net) in enumerate([
            ("صافي قيمة هذا الأمر التغييري (إضافة / حذف) (د.ك)", f"=SUM(G{f_r}:G{l_r})", False),
            ("قيمة العقد الأصلية (د.ك)", 0.0, False),
            ("صافي الأوامر التغييرية السابقة المعتمدة (د.ك)", 0.0, False),
            ("قيمة العقد المعدّلة الجديدة (د.ك)", f"=G{sr}+G{sr+1}+G{sr+2}", True),
        ]):
            rr = sr + off
            ws.row_dimensions[rr].height = 23
            ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=6)
            ws.cell(rr, 3, lbl).font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF" if is_net else NAVY)
            ws.cell(rr, 3).alignment = Alignment(horizontal="right", vertical="center")
            ws.cell(rr, 7, fml).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF" if is_net else NAVY)
            ws.cell(rr, 7).number_format = "#,##0.000"
            for cc in range(3, 8):
                ws.cell(rr, cc).fill = PatternFill("solid", fgColor=NAVY if is_net else SOFT_BG)
                ws.cell(rr, cc).border = GRID_BORDER
        _draw_signatures(ws, sr + 6, ["المقاول", "الاستشاري", "اعتماد المالك"], len(cols))

    elif kind == "mir":
        ws.title = "نموذج اعتماد وفحص MIR"
        cols = [("م", 6), ("وصف المادة / الأعمال المطلوب اعتمادها أو فحصها", 36), ("المصنّع / المنشأ", 20), ("الكمية / الموقع", 16), ("مرجع المواصفات / المخطط", 20), ("ملاحظات", 18)]
        _setup_a4_form(ws, "طلب اعتماد / فحص مواد وأعمال", "Material / Work Inspection & Submittal (MAR/MIR/WIR)", len(cols))
        for i, (_, w) in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        r = _draw_info_grid(ws, 3, [
            ("اسم المشروع:", ""), ("رقم الطلب / المراجعة:", "EL-001 / R00"),
            ("المقاول:", ""), ("تاريخ التقديم:", dt.date.today().isoformat()),
            ("الاستشاري:", ""), ("القسم التخصصي:", "مدني / كهرباء / صحي / تكييف / حريق"),
            ("المالك / الجهة:", ""), ("نوع الطلب:", "[ ] MAR   [ ] MIR   [ ] WIR   [ ] Shop Drawing"),
        ], len(cols))
        ws.row_dimensions[r].height = 26
        for i, (h, _) in enumerate(cols, 1):
            c = ws.cell(r, i, h)
            c.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=SLATE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = GRID_BORDER
        for dr in range(r + 1, r + 9):
            ws.row_dimensions[dr].height = 24
            ws.cell(dr, 1, dr - r).alignment = Alignment(horizontal="center", vertical="center")
            for ci in range(1, len(cols) + 1):
                ws.cell(dr, ci).border = GRID_BORDER
        cr = r + 10
        ws.merge_cells(start_row=cr, start_column=1, end_row=cr, end_column=len(cols))
        ws.row_dimensions[cr].height = 28
        code_c = ws.cell(cr, 1, "قرار الاستشاري:   [ ] Code A معتمد     [ ] Code B معتمد بملاحظات     [ ] Code C يُعاد التقديم     [ ] Code D مرفوض")
        code_c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        code_c.fill = PatternFill("solid", fgColor=SOFT_BG)
        code_c.alignment = Alignment(horizontal="center", vertical="center")
        for ci in range(1, len(cols) + 1):
            ws.cell(cr, ci).border = GRID_BORDER
        ws.merge_cells(start_row=cr + 1, start_column=1, end_row=cr + 1, end_column=len(cols))
        ws.row_dimensions[cr + 1].height = 55
        cm = ws.cell(cr + 1, 1, "ملاحظات الاستشاري / المهندس المشرف:")
        cm.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        cm.alignment = Alignment(horizontal="right", vertical="top")
        for ci in range(1, len(cols) + 1):
            ws.cell(cr + 1, ci).border = GRID_BORDER
        _draw_signatures(ws, cr + 3, ["تقديم المقاول", "قرار الاستشاري", "اعتماد المالك"], len(cols))

    elif kind == "transmittal":
        ws.title = "حافظة تسليم Transmittal"
        cols = [("م", 6), ("رقم المرجع", 18), ("Rev", 8), ("وصف المستند / المخطط / العينة", 38), ("عدد النسخ", 12), ("الغرض (للاعتماد/للعلم)", 20), ("ملاحظات", 18)]
        _setup_a4_form(ws, "حافظة إرسال وتسليم مستندات", "Document Transmittal", len(cols))
        for i, (_, w) in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        r = _draw_info_grid(ws, 3, [
            ("إلى السادة:", ""), ("رقم الحافظة:", "TR-2026-001"),
            ("عناية:", ""), ("التاريخ:", dt.date.today().isoformat()),
            ("المشروع:", ""), ("طريقة التسليم:", "باليد / بريد رسمي"),
        ], len(cols))
        ws.row_dimensions[r].height = 26
        for i, (h, _) in enumerate(cols, 1):
            c = ws.cell(r, i, h)
            c.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=SLATE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = GRID_BORDER
        f_r, l_r = r + 1, r + 12
        for dr in range(f_r, l_r + 1):
            ws.row_dimensions[dr].height = 22
            ws.cell(dr, 1, dr - f_r + 1).alignment = Alignment(horizontal="center", vertical="center")
            for ci in range(1, len(cols) + 1):
                ws.cell(dr, ci).border = GRID_BORDER
        tr = l_r + 1
        ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=4)
        ws.cell(tr, 1, "إجمالي عدد النسخ المسلّمة").font = Font(name="Calibri", size=10.5, bold=True, color=NAVY)
        ws.cell(tr, 5, f"=SUM(E{f_r}:E{l_r})").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        for ci in range(1, len(cols) + 1):
            ws.cell(tr, ci).fill = PatternFill("solid", fgColor=SOFT_BG)
            ws.cell(tr, ci).border = GRID_BORDER
        _draw_signatures(ws, tr + 2, ["المُرسِل (Delivered By)", "المُستلِم وختم الاستلام (Received By)"], len(cols))

    elif kind == "daily_report":
        ws.title = "التقرير اليومي للموقع"
        cols = [("م", 6), ("التخصص / الفئة", 20), ("عدد العمالة", 13), ("المعدات بالموقع", 18), ("الأعمال المنفذة اليوم وموقعها", 38), ("ملاحظات / عوائق", 22)]
        _setup_a4_form(ws, "التقرير اليومي لأعمال الموقع", "Daily Site Progress Report", len(cols))
        for i, (_, w) in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        r = _draw_info_grid(ws, 3, [
            ("المشروع:", ""), ("التاريخ / اليوم:", dt.date.today().isoformat()),
            ("المقاول:", ""), ("حالة الطقس / الحرارة:", "صحو"),
            ("مهندس الموقع:", ""), ("رقم التقرير:", "DSR-001"),
        ], len(cols))
        ws.row_dimensions[r].height = 26
        for i, (h, _) in enumerate(cols, 1):
            c = ws.cell(r, i, h)
            c.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=SLATE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = GRID_BORDER
        trades = ["مهندسون ومشرفون", "مدني / خرسانة / مباني", "تشطيبات", "كهرباء", "صحي وميكانيك", "تكييف HVAC", "مكافحة حريق", "عمالة مساعدة ونظافة", "مقاولو باطن", "أخرى"]
        f_r = r + 1
        for idx, tr_name in enumerate(trades):
            dr = f_r + idx
            ws.row_dimensions[dr].height = 22
            ws.cell(dr, 1, idx + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(dr, 2, tr_name)
            for ci in range(1, len(cols) + 1):
                ws.cell(dr, ci).border = GRID_BORDER
        l_r = f_r + len(trades) - 1
        tr = l_r + 1
        ws.cell(tr, 2, "إجمالي العمالة بالموقع").font = Font(name="Calibri", size=10.5, bold=True, color=NAVY)
        ws.cell(tr, 3, f"=SUM(C{f_r}:C{l_r})").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        for ci in range(1, len(cols) + 1):
            ws.cell(tr, ci).fill = PatternFill("solid", fgColor=SOFT_BG)
            ws.cell(tr, ci).border = GRID_BORDER
        _draw_signatures(ws, tr + 2, ["مهندس الموقع", "مدير المشروع", "اعتماد الاستشاري"], len(cols))

    for sheet_ws in wb.worksheets:
        _finalize_printable_borders(sheet_ws)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    print(f"[OK] تم إنشاء النموذج الجاهز للطباعة A4 بالمعادلات: {out_path} ({ws.title})")


def main():
    ap = argparse.ArgumentParser(description="Inspect, backup, and generate RTL Arabic Excel logs.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_ins = sub.add_parser("inspect", help="Inspect workbook structure, headers, formulas, and VBA")
    p_ins.add_argument("file")

    p_bak = sub.add_parser("backup", help="Create timestamped backup in _Backups/")
    p_bak.add_argument("file")

    p_tpl = sub.add_parser("create-template", help="Create a ready RTL Excel register template")
    p_tpl.add_argument("kind", choices=list(TEMPLATES.keys()))
    p_tpl.add_argument("out_file")

    p_frm = sub.add_parser("create-form", help="Create a print-ready A4 official form with live formulas")
    p_frm.add_argument("kind", choices=["ipc", "po", "vo", "mir", "transmittal", "daily_report"])
    p_frm.add_argument("out_file")

    args = ap.parse_args()
    if args.cmd == "inspect":
        data = inspect_workbook(Path(args.file))
        print(json.dumps(data, ensure_ascii=False, indent=2))
    elif args.cmd == "backup":
        dest = backup_workbook(Path(args.file))
        print(f"[OK] تم حفظ النسخة الاحتياطية: {dest}")
    elif args.cmd == "create-template":
        create_template(args.kind, Path(args.out_file))
    elif args.cmd == "create-form":
        create_printable_form(args.kind, Path(args.out_file))


if __name__ == "__main__":
    main()
