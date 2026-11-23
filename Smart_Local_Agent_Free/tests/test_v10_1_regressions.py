from __future__ import annotations

import datetime as dt
import io
import json
import shutil
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
import pytest

DOC_TOOLS = Path(__file__).resolve().parent.parent / "doc_tools"
if str(DOC_TOOLS) not in sys.path:
    sys.path.insert(0, str(DOC_TOOLS))

import active_project as ap
import inbox_pipeline as ip
import lint_strings
import project_dashboard as pdash
import project_log as plog
import scan_watcher as sw


def _http_json(url: str, *, method: str = "GET", body: dict | None = None) -> tuple[int, dict]:
    hdrs = {"Accept": "application/json"}
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        hdrs["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, method=method, headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, {"ok": False, "raw": raw}


def _build_realistic_multirev_log(path: Path) -> None:
    """Build a realistic multi-rev engineering log resembling real-world MOH submittals workbooks."""
    wb = openpyxl.Workbook()

    def _setup_disc(ws, title: str, rows_data: list[tuple[int, str, str, dt.date | None, dt.date | None, str, str]]):
        ws.title = title
        # Row 5: headers (matching real multi-rev layout with #, MOH Status in AD, formula in AC)
        ws.cell(5, 1, "#")
        ws.cell(5, 2, "Submittal No.")
        ws.cell(5, 3, "Discipline")
        ws.cell(5, 4, "Description")
        ws.cell(5, 7, "Submit Date (R0)")
        ws.cell(5, 8, "Reply Date (R0)")
        ws.cell(5, 9, "Action (R0)")
        ws.cell(5, 10, "Submit Date (R1)")
        ws.cell(5, 11, "Reply Date (R1)")
        ws.cell(5, 12, "Action (R1)")
        ws.cell(5, 29, "Final Status")  # AC
        ws.cell(5, 30, "MOH Status")    # AD
        ws.cell(5, 31, "Link")          # AE

        dv_act = DataValidation(
            type="list",
            formula1='"approved,approved (C),rejected,pending,withdrawn"',
            allow_blank=True,
        )
        ws.add_data_validation(dv_act)
        dv_act.add("G6:AB100")

        dv_moh = DataValidation(
            type="list",
            formula1='"✅ Approved,❌ Rejected,⏳ Under Review"',
            allow_blank=True,
        )
        ws.add_data_validation(dv_moh)
        dv_moh.add("AD6:AD100")

        for idx, (sn, ref, desc, s0, r0, act0, moh) in enumerate(rows_data, start=6):
            ws.cell(idx, 1, sn)
            ws.cell(idx, 2, ref)
            ws.cell(idx, 3, title)
            ws.cell(idx, 4, desc)
            ws.cell(idx, 7, s0)
            ws.cell(idx, 8, r0)
            ws.cell(idx, 9, act0)
            ws.cell(idx, 29, f'=IF(I{idx}="","",IF(AND(LOWER(I{idx})="pending",TODAY()-G{idx}>6),"⏳ Overdue",I{idx}))')
            ws.cell(idx, 30, moh)

    old_day = dt.date.today() - dt.timedelta(days=10)
    fresh_day = dt.date.today() - dt.timedelta(days=2)

    ws_civ = wb.active
    _setup_disc(ws_civ, "CIV", [
        (1, "CIV-001", "Concrete Mix", old_day, old_day, "approved", "✅ Approved"),
        (2, "CIV-002", "Old Cancelled Spec", old_day, None, "withdrawn", ""),
    ])
    ws_el = wb.create_sheet("EL")
    _setup_disc(ws_el, "EL", [
        (3, "EL-003", "Main Switchgear", old_day, None, "pending", "⏳ Under Review"),
        (4, "EL-004", "Cable Tray", fresh_day, None, "pending", ""),
    ])
    ws_ff = wb.create_sheet("FF")
    _setup_disc(ws_ff, "FF", [
        (5, "FF-005", "Sprinkler Heads", fresh_day, fresh_day, "approved (C)", ""),
    ])
    wb.save(path)


def test_global_serial_moh_dv_overdue_superseded_and_reply_date(tmp_path: Path):
    proj_root = tmp_path / "Proj_A"
    logs_dir = proj_root / "3_Project_Logs"
    logs_dir.mkdir(parents=True)
    log_path = logs_dir / "Master_Log.xlsx"
    _build_realistic_multirev_log(log_path)

    proj = {"id": "p_a", "name": "Health Center", "code": "HC", "root": str(proj_root)}

    # 1) Inspection captures MOH column AD, #->sn, DV vocab (incl. approved (C)), and overdue_days=6
    lm = plog.inspect_workbook_to_logmap(log_path)
    el_sc = next(s for s in lm["sheets"] if s["sheet"] == "EL")
    assert el_sc["columns"]["sn"]["col"] == 1
    assert el_sc["columns"]["moh_status"]["col"] == 30
    assert el_sc["overdue_days"] == 6
    assert el_sc["status_vocab"]["approved_noted"] == "approved (C)"
    assert el_sc["status_vocab"]["withdrawn"] == "withdrawn"
    assert el_sc["moh_vocab"]["approved"] == "✅ Approved"

    # 2) Global submittal numbering across CIV / EL / FF (max=5 -> next EL-006, not EL-005!)
    cnt = plog.recalculate_counters_from_logs(proj)
    assert cnt["is_global_disc_series"] is True
    assert cnt["global_disc_max"] == 5
    assert cnt["next_global_disc_preview"]["EL"] == "EL-006"
    assert cnt["next_global_disc_preview"]["CIV"] == "CIV-006"
    assert cnt["next_disc_preview"]["EL"] == "EL-005"  # backward-compatible per-sheet max+1

    nxt, _, _ = plog.next_reference_from_logs(proj, prefix="MAR", discipline="EL", use_discipline_series=True)
    assert nxt == "EL-006"

    # 3) Stats detect overdue_days=6 (EL-003 is 10 days old -> overdue; EL-004 is 2 days old -> not overdue) & withdrawn
    summary = plog.read_project_log_summary(proj)
    st = summary["stats"]
    assert st["overdue_days"] == 6
    assert st["overdue"] == 1
    assert st["withdrawn"] == 1
    assert st["approved"] == 2  # CIV-001 (approved) + FF-005 (approved (C))
    assert st["moh_approved"] == 1
    assert st["moh_pending"] == 1

    # 4) Superseded guard: older revision does NOT overwrite newer status/dates on existing row
    res_sup = plog.record_entry_in_project_log(proj, {
        "ref_no": "FF-005",
        "section": "submittal",
        "discipline": "FF",
        "revision": "R00",
        "status": "Rejected",
        "is_superseded": True,
        "reply_date": "2026-01-01",
    })
    assert res_sup.get("mode") == "skipped_superseded"
    wb = openpyxl.load_workbook(log_path)
    assert wb["FF"].cell(6, 9).value == "approved (C)"

    # 5) Code C update sets exact DV string "approved (C)" and keeps reply_date distinct from submit_date
    plog.record_entry_in_project_log(proj, {
        "ref_no": "EL-003",
        "section": "submittal",
        "discipline": "EL",
        "revision": "R00",
        "status": "Approved as Noted (Code C)",
        "doc_date": "2026-11-10",
        "reply_date": "2026-11-15",
    })
    wb = openpyxl.load_workbook(log_path)
    assert wb["EL"].cell(6, 9).value == "approved (C)"
    r_val = wb["EL"].cell(6, 8).value
    assert (r_val.strftime("%Y-%m-%d") if hasattr(r_val, "strftime") else str(r_val)) == "2026-11-15"  # reply date untouched by submit date
    assert wb["EL"].cell(6, 30).value == "⏳ Under Review"  # MOH column AD untouched!

    # 6) MOH authority update updates ONLY column AD and leaves consultant column I intact
    plog.record_entry_in_project_log(proj, {
        "ref_no": "EL-003",
        "section": "submittal",
        "discipline": "EL",
        "revision": "R00",
        "authority": "MOH",
        "status": "Code A",
    })
    wb = openpyxl.load_workbook(log_path)
    assert wb["EL"].cell(6, 9).value == "approved (C)"
    assert wb["EL"].cell(6, 30).value == "✅ Approved"

    # 7) Explicit submit_date does NOT leak into reply_date on new row insertion
    plog.record_entry_in_project_log(proj, {
        "ref_no": "EL-006",
        "section": "submittal",
        "discipline": "EL",
        "revision": "R00",
        "subject": "Emergency Lighting",
        "status": "pending",
        "doc_date": "2026-11-18",
        "reply_date": "",
    })
    wb = openpyxl.load_workbook(log_path)
    assert wb["EL"].cell(8, 2).value == "EL-006"
    s_val = wb["EL"].cell(8, 7).value
    assert (s_val.strftime("%Y-%m-%d") if hasattr(s_val, "strftime") else str(s_val)) == "2026-11-18"
    assert wb["EL"].cell(8, 8).value in (None, "")


def test_real_log_final_updated_smoke_if_present(tmp_path: Path):
    real_log = Path(__file__).resolve().parents[2] / "google_drive" / "LOG_Final_Updated.xlsm"
    if not real_log.exists():
        pytest.skip("Optional real reference workbook not present in workspace")

    proj_root = tmp_path / "Real_Smoke"
    (proj_root / "3_Project_Logs").mkdir(parents=True)
    copy_log = proj_root / "3_Project_Logs" / "LOG_Final_Updated.xlsm"
    shutil.copy2(real_log, copy_log)
    proj = {"id": "p_real", "name": "Real", "code": "SR", "root": str(proj_root)}

    cnt = plog.recalculate_counters_from_logs(proj)
    assert cnt["is_global_disc_series"] is True
    assert cnt["global_disc_max"] == 227
    assert cnt["next_global_disc_preview"]["EL"] == "EL-228"

    res = plog.record_entry_in_project_log(proj, {
        "ref_no": "EL-215",
        "section": "submittal",
        "discipline": "EL",
        "revision": "R00",
        "authority": "MOH",
        "status": "Approved",
    })
    assert res["ok"] and res.get("is_moh_update") is True

    summary = plog.read_project_log_summary(proj)
    assert summary["stats"]["withdrawn"] == 19
    assert summary["stats"]["moh_approved"] >= 1


def test_scan_watcher_optional_fitz_gemini_errors_and_watcher_state(tmp_path: Path, monkeypatch):
    # 1) Simulate missing fitz (PyMuPDF) safely: module imports, %PDF check passes, visual page step raises clear translated error
    orig_fitz = sw.fitz
    monkeypatch.setattr(sw, "fitz", None)
    pdf_p = tmp_path / "sample.pdf"
    pdf_p.write_bytes(b"%PDF-1.4\n%fake\n")
    ok_valid, sz = sw.is_file_unlocked_and_valid(pdf_p)
    assert ok_valid is True and sz > 0
    with pytest.raises(RuntimeError) as fitz_err:
        sw.prepare_visual_pages(pdf_p, tmp_path / "out")
    assert "PDF" in str(fitz_err.value)

    # Restore fitz for remaining tests if available
    monkeypatch.setattr(sw, "fitz", orig_fitz)

    # 2) Gemini 401/403 raises RuntimeError with human-readable message instead of silent None
    def _fake_urlopen(req, timeout=90):
        raise urllib.error.HTTPError(
            req.full_url, 403, "Forbidden", {}, io.BytesIO(b'{"error":{"message":"API key invalid"}}')
        )

    monkeypatch.setattr(urllib.request, "urlopen", _fake_urlopen)
    with pytest.raises(RuntimeError) as exc_info:
        sw.analyze_with_gemini_api(
            api_key="AIzaDummyInvalidKey1234567890",
            prompt="Analyze",
            images=[],
            preferred_model="gemini-2.5-flash",
        )
    assert len(str(exc_info.value)) > 5

    # 3) Pipeline scan records errors/warnings into watcher_state.json for dashboard visibility
    from PIL import Image

    proj_root = tmp_path / "Proj_W"
    inbox_dir = proj_root / "1_Inbox"
    inbox_dir.mkdir(parents=True)
    bad_file = inbox_dir / "20261118_001.png"
    Image.new("RGB", (120, 120), color=(240, 240, 240)).save(bad_file)

    reg = tmp_path / "_hub_state" / "registry.json"
    proot = tmp_path / "Projects"
    hub = pdash.HubState(reg, proot)
    p_rec = hub.create_project({"name": "Watcher Proj", "code": "WP", "root": str(proj_root), "inbox_dir": str(inbox_dir)})

    def _boom_analyzer(prompt, job_dir, model="", images=None, api_key="", **kw):
        raise RuntimeError("Gemini key rejected (403)")

    summary = ip.process_project_inbox_once(
        p_rec,
        all_projects=hub.list_projects(),
        global_settings=hub.get_settings(),
        decision_store=hub.decisions,
        device_id=hub.device_id,
        analyzer_override=_boom_analyzer,
    )
    assert len(summary["errors"]) == 1
    assert "403" in summary["errors"][0]

    detail = pdash.build_project_detail(hub, hub.get_project(p_rec["id"]))
    assert len(detail["last_scan_errors"]) == 1
    assert "403" in detail["last_scan_errors"][0]


def test_whatsapp_endpoints_i18n_and_lint_strings(tmp_path: Path):
    # 1) Full i18n + source lint must pass with 0 violations (including dashboard_ui.html!)
    errs = lint_strings.lint_dictionary(DOC_TOOLS / "strings_ar.json") + lint_strings.lint_source_code(DOC_TOOLS)
    assert errs == [], f"Unexpected i18n/code lint errors: {errs}"

    # 2) Live server serves i18n bundle, dual styles, and WhatsApp config/send safely
    reg = tmp_path / "_hub_state" / "registry.json"
    proot = tmp_path / "Projects"
    hub = pdash.HubState(reg, proot)
    srv = pdash.create_server(hub, host="127.0.0.1", port=0, allow_remote_preview=False)
    port = srv.server_address[1]
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    try:
        base = f"http://127.0.0.1:{port}"
        with urllib.request.urlopen(f"{base}/", timeout=10) as r:
            html = r.read().decode("utf-8")
            assert "window.__I18N__" not in html or "__I18N_BOOT__" in html
            assert "t('topbar.app_title')" in html or "t(\"topbar.app_title\")" in html
            assert "__I18N_BUNDLE__" not in html  # injected cleanly by handler!
            assert "settings.global_submittal_numbering" in html

        st, i18n_eg = _http_json(f"{base}/api/i18n?style=eg")
        assert st == 200 and i18n_eg["active_style"] == "eg"
        assert "watcher.pymupdf_missing" in i18n_eg["strings"]

        # Save WhatsApp config + secret masking + non-destructive partial update
        st, wa1 = _http_json(f"{base}/api/whatsapp/config", method="POST", body={
            "default_country_code": "+965",
            "phone_number_id": "1234567890",
            "access_token": "EAAB_SECRET_WA_TOKEN_9999",
        })
        assert st == 200 and wa1["config"]["configured"] is True
        assert "EAAB_SECRET_WA_TOKEN_9999" not in json.dumps(wa1)

        st, wa2 = _http_json(f"{base}/api/whatsapp/config", method="POST", body={
            "default_country_code": "965",
            "access_token": "",  # must NOT wipe existing token
        })
        assert st == 200 and wa2["config"]["configured"] is True

        # Project WhatsApp web send + path traversal guard
        _, pr = _http_json(f"{base}/api/projects", method="POST", body={"name": "WA Proj", "code": "WA1"})
        pid = pr["project"]["id"]

        st, wres = _http_json(f"{base}/api/projects/{pid}/whatsapp/send", method="POST", body={
            "mode": "web",
            "phone": "55123456",
            "message": "Test Submittal EL-001",
        })
        assert st == 200 and wres["mode"] == "web"
        assert "phone=96555123456" in wres["url"]

        st_trav, _ = _http_json(f"{base}/api/projects/{pid}/whatsapp/send", method="POST", body={
            "mode": "web",
            "phone": "55123456",
            "rel": "../../etc/passwd",
        })
        assert st_trav == 403

        # Run start_sync includes global_disc fields
        st_s, sres = _http_json(f"{base}/api/projects/{pid}/start", method="POST", body={})
        assert st_s == 200
        assert "next_global_disc_preview" in sres["counters"]
        assert "is_global_disc_series" in sres["counters"]
    finally:
        hub.toggle_watcher(False)
        srv.shutdown()
        srv.server_close()
