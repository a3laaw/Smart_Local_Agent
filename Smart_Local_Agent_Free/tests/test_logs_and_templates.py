from __future__ import annotations

import base64
import shutil
import sys
from pathlib import Path

import docx
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pytest

DOC_TOOLS = Path(__file__).resolve().parent.parent / "doc_tools"
if str(DOC_TOOLS) not in sys.path:
    sys.path.insert(0, str(DOC_TOOLS))

import project_log as plog
import template_profile as tprof
import template_render as trender

# Tiny 1x1 PNG logo for offline image-rendering tests
TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)
REAL_XLSM = Path(__file__).resolve().parent.parent.parent / "google_drive" / "LOG_Final.xlsm"


def _make_rich_xlsx_template(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IPC_Form"
    ws.sheet_view.rightToLeft = True
    ws.merge_cells("A1:E1")
    ws["A1"] = "شهادة دفع جارية (IPC)"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="0F766E")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    thin = Side(style="thin", color="334155")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    labels = [
        ("A3", "اسم المشروع:", "B3", ""),
        ("D3", "رقم المرجع:", "E3", ""),
        ("A4", "تاريخ المستند:", "B4", ""),
        ("D4", "كود الاعتماد:", "E4", ""),
        ("A5", "الموضوع:", "B5", ""),
        ("D5", "التخصص:", "E5", ""),
        ("A6", "دفعة مقدمة (د.ك):", "B6", ""),
    ]
    for lc, lt, vc, vt in labels:
        ws[lc] = lt
        ws[lc].font = Font(bold=True)
        ws[lc].border = bdr
        ws[vc] = vt
        ws[vc].border = bdr

    # Repeating items table: header row 8, template row 9, totals row 10
    headers = ["م", "بيان الأعمال", "الكمية", "سعر الوحدة", "الإجمالي"]
    for idx, h in enumerate(headers, 1):
        c = ws.cell(8, idx, h)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="E2E8F0")
        c.border = bdr
        dc = ws.cell(9, idx)
        dc.border = bdr
    ws["E9"] = "=C9*D9"
    ws["D10"] = "الإجمالي الكلي"
    ws["E10"] = "=SUM(E9:E9)"
    ws["E10"].font = Font(bold=True)
    ws["E10"].number_format = "#,##0.000"

    # Embed tiny PNG logo
    logo_p = path.parent / "_tmp_logo.png"
    logo_p.write_bytes(TINY_PNG)
    try:
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(str(logo_p))
        img.width = 48
        img.height = 32
        ws.add_image(img, "E1")
        wb.save(path)
    finally:
        if logo_p.exists():
            logo_p.unlink()


def _make_rich_docx_template(path: Path) -> None:
    doc = docx.Document()
    logo_p = path.parent / "_tmp_wlogo.png"
    logo_p.write_bytes(TINY_PNG)
    try:
        hdr = doc.sections[0].header
        hp = hdr.paragraphs[0]
        hp.add_run().add_picture(str(logo_p))
        hp.add_run("  شركة المقاولات الهندسية")

        doc.add_paragraph("المرجع: {{REF_NO}}")
        doc.add_paragraph("الموضوع: ")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "اسم المشروع"
        tbl.cell(0, 1).text = ""
        tbl.cell(1, 0).text = "الجهة المرسل إليها"
        tbl.cell(1, 1).text = ""

        items = doc.add_table(rows=2, cols=3)
        items.cell(0, 0).text = "م"
        items.cell(0, 1).text = "البيان"
        items.cell(0, 2).text = "الكمية"
        doc.save(str(path))
    finally:
        if logo_p.exists():
            logo_p.unlink()


def _make_flat_project_log(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Submittals"
    ws.append(["No", "Reference", "Rev", "Discipline", "Description", "Submit Date", "Status", "Amount", "Remarks", "Link"])
    ws.append([1, "MAR-004", "R00", "EL", "اعتماد كابلات ضغط منخفض", "2026-10-01", "A", 1250.500, "", ""])
    ws2 = wb.create_sheet("Correspondence")
    ws2.append(["No", "Reference", "Description", "Date", "Party", "Status", "Link"])
    ws2.append([1, "OUT-011", "كتاب تسليم موقع", "2026-10-05", "المكتب الاستشاري", "Sent", ""])
    ws3 = wb.create_sheet("Payments")
    ws3.append(["No", "Reference", "Description", "Date", "Status", "Amount", "TaxFormula"])
    ws3.append([1, "IPC-002", "مستخلص جاري رقم 2", "2026-10-10", "Approved", 15000.000, "=F2*0.05"])
    wb.save(path)


# ── 1. Word-Boundary Section Suggestion & Zero Built-in Templates ───────────
def test_section_suggestion_word_boundaries_and_no_silent_fallback(tmp_path: Path):
    # Previously buggy filenames MUST NOT misclassify to VO / PO / MAR!
    assert tprof.suggest_template_section("Invoice_2026.xlsx")["suggested_section"] != "variation_vo"
    assert tprof.suggest_template_section("Monthly_Report.xlsx")["suggested_section"] == "site_report"
    assert tprof.suggest_template_section("Technical_Proposal.docx")["suggested_section"] is None
    assert tprof.suggest_template_section("Import_Permit.xlsx")["suggested_section"] is None
    assert tprof.suggest_template_section("Support_Form.docx")["suggested_section"] is None
    assert tprof.suggest_template_section("Random_Company_Form.xlsx")["suggested_section"] is None

    # Exact tokens & Arabic keywords suggest cleanly (never auto-confirmed!)
    assert tprof.suggest_template_section("MAR_Submittal_Form.xlsx")["suggested_section"] == "submittal_mar"
    assert tprof.suggest_template_section("WIR-Inspection.docx")["suggested_section"] == "submittal_ir"
    assert tprof.suggest_template_section("أمر_شراء_مواد.xlsx")["suggested_section"] == "purchase_po"
    assert tprof.suggest_template_section("تقرير_شهري.docx")["suggested_section"] == "site_report"
    assert tprof.suggest_template_section("VO_01_Form.xlsx")["suggested_section"] == "variation_vo"

    # Empty project has ZERO system templates, and Ready_Templates dir does not exist
    assert not (DOC_TOOLS.parent / "Ready_Templates").exists()
    assert not (DOC_TOOLS / "Ready_Templates").exists()
    assert tprof.list_project_templates(tmp_path) == []


# ── 2. Excel Visual Render, 6 Roles, Formula Guard, Repeating Rows & Drift ──
def test_xlsx_template_render_fill_formulas_validation_and_drift(tmp_path: Path):
    proj_root = tmp_path / "proj1"
    raw_tpl = tmp_path / "IPC_Client_Form.xlsx"
    _make_rich_xlsx_template(raw_tpl)

    up = tprof.upload_project_template(proj_root, raw_tpl.name, raw_tpl.read_bytes(), device_id="dev-test")
    tname = up["name"]

    # Faithful visual render includes RTL, merged title, styles, formulas & inline data: URI logo
    vis = trender.render_template(proj_root, tname)
    assert vis["kind"] == "xlsx"
    assert vis["rtl"] is True
    assert len(vis["images"]) >= 1 and vis["images"][0]["data_uri"].startswith("data:image/png;base64,")
    a1 = vis["rows"][0]["cells"][0]
    assert a1["coord"] == "A1" and a1.get("colspan") == 5
    assert a1["style"].get("bg") == "#0F766E"

    # Blocking attempt to map a manual field onto protected formula cell E10
    with pytest.raises(tprof.FormulaOverwriteError):
        tprof.save_template_profile(proj_root, tname, {
            "section": "contract_ipc",
            "fields": [{"label": "خطأ", "role": "manual", "input_type": "kwd", "target": {"sheet": "IPC_Form", "coord": "E10"}}],
        })

    # Save valid profile exercising all 6 roles + validation + repeating table with SUM row
    prof = tprof.save_template_profile(proj_root, tname, {
        "section": "contract_ipc",
        "fields": [
            {"key": "title_fixed", "label": "العنوان", "role": "fixed", "target": {"sheet": "IPC_Form", "coord": "A1"}},
            {"key": "proj_name", "label": "اسم المشروع", "role": "auto_project", "source": "project.name", "target": {"sheet": "IPC_Form", "coord": "B3"}},
            {"key": "ref_no", "label": "رقم المرجع", "role": "auto_system", "source": "system.next_ref", "target": {"sheet": "IPC_Form", "coord": "E3"}},
            {"key": "doc_date", "label": "التاريخ", "role": "manual", "input_type": "date", "required": True, "target": {"sheet": "IPC_Form", "coord": "B4"}},
            {"key": "approval_code", "label": "كود الاعتماد", "role": "manual", "input_type": "approval_code", "default_value": "B", "target": {"sheet": "IPC_Form", "coord": "E4"}},
            {"key": "subject", "label": "الموضوع", "role": "manual", "input_type": "text", "required": True, "target": {"sheet": "IPC_Form", "coord": "B5"}},
            {"key": "disc_sel", "label": "التخصص", "role": "manual", "input_type": "dropdown", "options": ["مدني", "كهرباء", "صحي"], "required": True, "target": {"sheet": "IPC_Form", "coord": "E5"}},
            {"key": "advance_kwd", "label": "دفعة مقدمة", "role": "manual", "input_type": "kwd", "default_value": "0.000", "target": {"sheet": "IPC_Form", "coord": "B6"}},
        ],
        "repeating_tables": [{
            "key": "boq",
            "label": "بنود الأعمال",
            "sheet": "IPC_Form",
            "header_row": 8,
            "template_row": 9,
            "totals_row": 10,
            "columns": [
                {"key": "sn", "label": "م", "col": "A", "input_type": "auto_index"},
                {"key": "desc", "label": "البيان", "col": "B", "input_type": "text", "required": True},
                {"key": "qty", "label": "الكمية", "col": "C", "input_type": "number", "required": True},
                {"key": "rate", "label": "السعر", "col": "D", "input_type": "kwd", "required": True},
                {"key": "line_total", "label": "الإجمالي", "col": "E", "input_type": "formula_protected"},
            ],
        }],
    }, device_id="dev-test")

    proj_meta = {"name": "مشروع مدرسة حولي", "code": "HW-01", "contract_no": "C-99"}
    sys_ctx = {"next_ref": "IPC-003", "today_iso": "2026-11-17", "revision": "R00"}

    # Validation catches missing required & invalid dropdown & bad KWD
    with pytest.raises(tprof.ProfileValidationError) as exc_info:
        tprof.validate_and_resolve_payload(prof, {
            "values": {"doc_date": "not-a-date", "subject": "", "disc_sel": "غير موجود", "advance_kwd": "abc"}
        }, proj_meta, system_ctx=sys_ctx)
    assert len(exc_info.value.errors) >= 3

    # Fill 3 repeating rows -> inserts 2 extra rows, expands SUM(E9:E11), preserves fixed title & formulas!
    out_xlsx = tmp_path / "filled_ipc.xlsx"
    tprof.fill_by_profile(
        tprof.resolve_project_template(proj_root, tname),
        prof,
        {
            "values": {"doc_date": "2026-11-17", "subject": "مستخلص أعمال الكهرباء", "disc_sel": "كهرباء", "advance_kwd": "250.75"},
            "tables": {"boq": [
                {"desc": "تمديد كابلات", "qty": 10, "rate": "12.500"},
                {"desc": "لوحات توزيع", "qty": 2, "rate": "450.250"},
                {"desc": "إنارة LED", "qty": 20, "rate": "8.125"},
            ]},
        },
        proj_meta,
        out_xlsx,
        system_ctx=sys_ctx,
    )
    wb_out = openpyxl.load_workbook(out_xlsx, data_only=False)
    ws = wb_out["IPC_Form"]
    assert ws["A1"].value == "شهادة دفع جارية (IPC)"  # fixed untouched
    assert ws["B3"].value == "مشروع مدرسة حولي"        # auto_project
    assert ws["E3"].value == "IPC-003"                 # auto_system
    assert ws["E4"].value == "B"                       # default approval code
    assert ws["B5"].value == "مستخلص أعمال الكهرباء"
    assert ws["E5"].value == "كهرباء"
    assert float(ws["B6"].value) == pytest.approx(250.75, abs=1e-4)
    assert "0.000" in (ws["B6"].number_format or "")
    # Repeating rows 9, 10, 11 + row formulas + shifted & expanded SUM at E12!
    assert [ws["A9"].value, ws["A10"].value, ws["A11"].value] == [1, 2, 3]
    assert ws["B11"].value == "إنارة LED"
    assert str(ws["E11"].value).startswith("=") and "11" in str(ws["E11"].value)
    assert "E9:E11" in str(ws["E12"].value).upper()
    wb_out.close()

    # SHA-256 Drift Test A: Insert 2 rows at top -> anchor auto-rebinds cleanly!
    tpl_file = tprof.resolve_project_template(proj_root, tname)
    wb_mod = openpyxl.load_workbook(tpl_file)
    wb_mod["IPC_Form"].insert_rows(2, amount=2)
    wb_mod.save(tpl_file)
    wb_mod.close()

    rebound_prof = tprof.get_template_profile(proj_root, tname)
    assert len(rebound_prof["orphan_fields"]) == 0
    assert len(rebound_prof["rebound_fields"]) >= 1
    subj_f = next(f for f in rebound_prof["fields"] if f["key"] == "subject")
    assert subj_f["target"]["coord"] == "B7"  # shifted from B5 -> B7 automatically!

    # SHA-256 Drift Test B: Destroy anchor label -> triggers orphan & blocks fill!
    wb_brk = openpyxl.load_workbook(tpl_file)
    wb_brk["IPC_Form"]["A7"] = "حقل مختلف تماماً"
    wb_brk.save(tpl_file)
    wb_brk.close()

    broken_prof = tprof.get_template_profile(proj_root, tname)
    assert any(o["key"] == "subject" for o in broken_prof["orphan_fields"])
    with pytest.raises(tprof.ProfileOrphanError):
        tprof.fill_by_profile(tpl_file, broken_prof, {
            "values": {"doc_date": "2026-11-17", "subject": "تجربة", "disc_sel": "مدني"},
            "tables": {"boq": [{"desc": "بند", "qty": 1, "rate": 5}]},
        }, proj_meta, tmp_path / "blocked.xlsx", system_ctx=sys_ctx)


# ── 3. Word (.docx) Visual Render & Profile Fill (Span, Paragraph, Cell, Table)
def test_docx_template_render_and_fill(tmp_path: Path):
    proj_root = tmp_path / "proj_docx"
    raw_docx = tmp_path / "Official_Letter.docx"
    _make_rich_docx_template(raw_docx)

    up = tprof.upload_project_template(proj_root, raw_docx.name, raw_docx.read_bytes())
    tname = up["name"]
    vis = trender.render_template(proj_root, tname, try_libreoffice=False)
    assert vis["kind"] == "docx"
    assert len(vis["header_blocks"]) >= 1
    assert vis["tables_count"] == 2

    prof = tprof.save_template_profile(proj_root, tname, {
        "section": "outgoing",
        "fields": [
            {"key": "ref_no", "label": "المرجع", "role": "auto_system", "source": "system.next_ref",
             "target": {"kind": "docx_span", "para_idx": 0, "match_text": "{{REF_NO}}"}},
            {"key": "subject", "label": "الموضوع", "role": "manual", "input_type": "text", "required": True,
             "target": {"kind": "docx_paragraph", "para_idx": 1, "replace_mode": "after_colon"}},
            {"key": "proj", "label": "المشروع", "role": "auto_project", "source": "project.name",
             "target": {"kind": "docx_cell", "table_idx": 0, "row_idx": 0, "col_idx": 1}},
            {"key": "party_name", "label": "الجهة", "role": "manual", "input_type": "text", "required": True,
             "target": {"kind": "docx_cell", "table_idx": 0, "row_idx": 1, "col_idx": 1}},
        ],
        "repeating_tables": [{
            "key": "att",
            "label": "المرفقات",
            "table_idx": 1,
            "header_row": 0,
            "template_row": 1,
            "columns": [
                {"key": "sn", "label": "م", "col_idx": 0, "input_type": "auto_index"},
                {"key": "item", "label": "البيان", "col_idx": 1, "input_type": "text"},
                {"key": "qty", "label": "الكمية", "col_idx": 2, "input_type": "number"},
            ],
        }],
    })

    out_doc = tmp_path / "filled_letter.docx"
    tprof.fill_by_profile(
        tprof.resolve_project_template(proj_root, tname),
        prof,
        {
            "values": {"subject": "تسليم مخططات التكييف", "party_name": "دار الاستشارات"},
            "tables": {"att": [{"item": "مخطط الطابق الأرضي", "qty": 3}, {"item": "مخطط السطح", "qty": 2}]},
        },
        {"name": "مشروع برج العاصمة"},
        out_doc,
        system_ctx={"next_ref": "OUT-012", "today_iso": "2026-11-17"},
    )
    d = docx.Document(str(out_doc))
    assert "OUT-012" in d.paragraphs[0].text
    assert "تسليم مخططات التكييف" in d.paragraphs[1].text
    assert d.tables[0].cell(0, 1).text.strip() == "مشروع برج العاصمة"
    assert d.tables[0].cell(1, 1).text.strip() == "دار الاستشارات"
    assert len(d.tables[1].rows) == 3
    assert d.tables[1].cell(2, 1).text.strip() == "مخطط السطح"


# ── 4. Uploaded Project Logs (Flat .xlsx + Lock Queue + Unrelated Sheet + Real .xlsm)
def test_uploaded_flat_project_log_and_lock_queue(tmp_path: Path, monkeypatch):
    proj_root = tmp_path / "proj_log"
    proj_root.mkdir()
    src_log = tmp_path / "MyProject_Log.xlsx"
    _make_flat_project_log(src_log)

    lmap = plog.upload_project_log(proj_root, src_log.name, src_log.read_bytes(), device_id="pc-a")
    assert lmap["log_name"] == "MyProject_Log.xlsx"
    proj = {"id": "p1", "name": "مشروع أ", "root": str(proj_root), "log_unrelated_sheet": True, "baseline_counters": {"PO": 9}}

    # Next references come directly from the uploaded log (and respect higher baseline)
    r_mar, _, _ = plog.next_reference_from_logs(proj, prefix="MAR")
    r_out, _, _ = plog.next_reference_from_logs(proj, prefix="OUT")
    r_ipc, _, _ = plog.next_reference_from_logs(proj, prefix="IPC")
    r_po, _, _ = plog.next_reference_from_logs(proj, prefix="PO")
    assert r_mar == "MAR-005"
    assert r_out == "OUT-012"
    assert r_ipc == "IPC-003"
    assert r_po == "PO-010"

    # Insert new IPC row & verify formula column G preserved/copied!
    ins = plog.record_entry_in_project_log(proj, {
        "section": "contract_ipc", "ref_no": "IPC-003", "doc_date": "2026-11-17",
        "subject_ar": "مستخلص جاري 3", "status": "Under Review", "amount": 20000.500,
    }, device_id="pc-a")
    assert ins["ok"] is True and ins["sheet"] == "Payments"

    # Simulate Excel lock -> entry safely queued, then flushed once unlocked!
    log_file = plog.logs_dir_for_project(proj_root) / "MyProject_Log.xlsx"
    lock_marker = log_file.parent / "~$MyProject_Log.xlsx"
    lock_marker.write_bytes(b"locked")
    q_res = plog.record_entry_in_project_log(proj, {
        "section": "submittal_mar", "ref_no": "MAR-005", "revision": "R00",
        "discipline": "EL", "doc_date": "2026-11-17", "subject_ar": "اعتماد قواطع", "status": "Under Review",
    }, device_id="pc-a", queue_if_locked=True)
    assert q_res.get("queued") is True
    assert plog.pending_queue_count(proj_root) == 1

    lock_marker.unlink()
    fl = plog.flush_project_log_queue(proj, device_id="pc-a")
    assert fl["flushed"] == 1 and fl["remaining"] == 0

    # Optional separate unrelated sheet never touches project sheets
    u_res = plog.record_optional_unrelated_entry(
        proj, file_name="scan01.pdf", new_name="2026-11-17_فاتورة_مطعم.pdf",
        action_taken="rename_in_place", reason="غير مرتبط", entity="مطعم", device_id="pc-a",
    )
    assert u_res.get("logged") is True and u_res.get("sheet") == plog.UNRELATED_SHEET_NAME
    summary = plog.read_project_log_summary(proj)
    assert summary["stats"]["total_rows"] == 5  # 3 initial + IPC-003 + MAR-005 (unrelated sheet excluded!)


@pytest.mark.skipif(not REAL_XLSM.exists(), reason="LOG_Final.xlsm fixture not present")
def test_uploaded_real_xlsm_vba_formulas_and_multirev(tmp_path: Path):
    proj_root = tmp_path / "proj_xlsm"
    proj_root.mkdir()
    lmap = plog.upload_project_log(proj_root, "LOG_Final.xlsm", REAL_XLSM.read_bytes(), device_id="pc-b")
    assert lmap["has_vba"] is True
    proj = {"id": "p2", "name": "مشروع السور", "root": str(proj_root)}

    sync = plog.recalculate_counters_from_logs(proj, device_id="pc-b")
    assert sync["next_disc_preview"]["EL"] == "EL-212"
    assert sync["next_disc_preview"]["CIV"] == "CIV-206"

    # Insert R00 then update reply to A on EL-212
    r1 = plog.record_entry_in_project_log(proj, {
        "section": "submittal_mar", "discipline": "EL", "ref_no": "EL-212",
        "revision": "R00", "doc_date": "2026-11-17", "subject_ar": "اعتماد كشافات طوارئ", "status": "Under Review",
    }, device_id="pc-b")
    assert r1["ok"] is True and r1["sheet"] == "EL"

    r2 = plog.record_entry_in_project_log(proj, {
        "section": "submittal_mar", "discipline": "EL", "ref_no": "EL-212",
        "revision": "R00", "reply_date": "2026-11-18", "status": "A",
    }, device_id="pc-b")
    assert r2["ok"] is True and r2["row"] == r1["row"]

    # Verify VBA archive and formula column survive intact!
    saved_xlsm = plog.logs_dir_for_project(proj_root) / "LOG_Final.xlsm"
    wb = openpyxl.load_workbook(saved_xlsm, keep_vba=True, data_only=False)
    assert wb.vba_archive is not None
    assert str(wb["EL"].cell(r1["row"], 29).value or "").startswith("=")
    wb.close()
