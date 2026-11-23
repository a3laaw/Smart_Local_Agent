#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""End-to-end tests for v10.12.0's 5 ordered productivity enhancements."""

from __future__ import annotations

import base64
import datetime as dt
import io
import json
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

import docx
import openpyxl
from PIL import Image, ImageDraw
import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import active_project as actp  # noqa: E402
import project_dashboard as pdash  # noqa: E402
import security_utils as secu  # noqa: E402
import whatsapp_archiver as waa  # noqa: E402


@pytest.fixture()
def live_server(tmp_path: Path):
    reg = tmp_path / "_hub_state" / "registry.json"
    proot = tmp_path / "Projects"
    hub = pdash.HubState(reg, proot)
    srv = pdash.create_server(hub, host="127.0.0.1", port=0, allow_remote_preview=False)
    port = srv.server_address[1]
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    try:
        yield f"http://127.0.0.1:{port}", hub, tmp_path
    finally:
        hub.stop_all_background()
        srv.shutdown()
        srv.server_close()


def _http_json(url: str, *, method: str = "GET", body: dict | None = None) -> tuple[int, dict]:
    data = json.dumps(body).encode("utf-8") if body is not None else None
    headers = {"Content-Type": "application/json"} if body is not None else {}
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as he:
        return he.code, json.loads(he.read().decode("utf-8"))


def _http_raw(url: str) -> tuple[int, bytes, dict[str, str]]:
    try:
        with urllib.request.urlopen(url, timeout=20) as resp:
            return resp.status, resp.read(), {k.lower(): v for k, v in resp.headers.items()}
    except urllib.error.HTTPError as he:
        return he.code, he.read(), {k.lower(): v for k, v in he.headers.items()}


def _make_png_bytes(color: tuple[int, int, int] = (5, 155, 229)) -> bytes:
    im = Image.new("RGB", (240, 160), color)
    d = ImageDraw.Draw(im)
    d.rectangle([20, 20, 220, 140], outline=(255, 203, 43), width=4)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


def _create_sample_log_xlsx(path: Path) -> None:
    today = dt.date.today()
    old_25 = (today - dt.timedelta(days=25)).isoformat()
    old_20 = (today - dt.timedelta(days=20)).isoformat()
    old_10 = (today - dt.timedelta(days=10)).isoformat()
    recent_2 = (today - dt.timedelta(days=2)).isoformat()

    wb = openpyxl.Workbook()
    ws_sub = wb.active
    ws_sub.title = "MAR"
    ws_sub.append(["Ref No", "Discipline", "Subject", "Date", "Reply Date", "Status"])
    ws_sub.append(["MAR-EL-001", "EL", "Switchgear Power Cables", old_25, "", "Under Review"])
    ws_sub.append(["MAR-EL-002", "EL", "Lighting Fixture Sample", old_20, old_10, "Code C - Rejected"])
    ws_sub.append(["MAR-HVAC-003", "HVAC", "Chilled Water Valve", recent_2, "", "Under Review"])

    ws_in = wb.create_sheet("Incoming")
    ws_in.append(["Ref No", "Party", "Subject", "Date", "Status"])
    ws_in.append(["IN-014", "Consultant", "Site Safety Notice", old_10, ""])
    ws_in.append(["IN-015", "Ministry", "Permit Extension Approved", recent_2, "Replied"])

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _create_sample_docx_tpl(path: Path) -> None:
    d = docx.Document()
    d.add_heading("Material Submittal", level=1)
    tb = d.add_table(rows=2, cols=2)
    tb.cell(0, 0).text = "Ref No"
    tb.cell(0, 1).text = "{{ref_no}}"
    tb.cell(1, 0).text = "Subject"
    tb.cell(1, 1).text = "{{subject}}"
    path.parent.mkdir(parents=True, exist_ok=True)
    d.save(path)


def test_feature_1_global_search_and_feature_4_inline_preview(live_server):
    base, hub, tmp_path = live_server
    p1_root = tmp_path / "P1"
    p2_root = tmp_path / "P2"

    _, r1 = _http_json(f"{base}/api/projects", method="POST", body={"name": "Hawalli Tower", "code": "HW-1", "root": str(p1_root)})
    _, r2 = _http_json(f"{base}/api/projects", method="POST", body={"name": "Salmiya Mall", "code": "SM-2", "root": str(p2_root)})
    pid1 = r1["project"]["id"]
    pid2 = r2["project"]["id"]

    # Populate log & matching archive PNG in Project 1
    _create_sample_log_xlsx(p1_root / "3_Project_Logs" / "HW_Log.xlsx")
    arc_img = p1_root / "2_Archive" / "01_Submittals" / "MAR-EL-001_2026-11-01_Switchgear.png"
    arc_img.parent.mkdir(parents=True, exist_ok=True)
    arc_img.write_bytes(_make_png_bytes())

    # Populate archive file & WhatsApp plan item in Project 2
    arc_doc = p2_root / "2_Archive" / "02_Letters" / "IN-099_Site_Generator.png"
    arc_doc.parent.mkdir(parents=True, exist_ok=True)
    arc_doc.write_bytes(_make_png_bytes((246, 130, 13)))
    secu.atomic_write_json(
        waa.whatsapp_plan_path(p2_root),
        {
            "items": [
                {
                    "id": "wa-item-1",
                    "final_name": "2026-11-16_Generator_Inspection.png",
                    "orig_name": "IMG-1.png",
                    "final_path": str(arc_doc),
                    "context_text": "Switchgear generator test photo",
                    "sender": "Eng. Tariq",
                    "ref_no": "IN-099",
                    "date_iso": "2026-11-16",
                }
            ]
        },
    )

    # 1) Global search across ALL projects finds log_row + archive_file (P1) + whatsapp (P2)
    st_s, res_all = _http_json(f"{base}/api/search?q=Switchgear")
    assert st_s == 200 and res_all["ok"] is True
    assert res_all["counts"]["log_row"] >= 1
    assert res_all["counts"]["archive_file"] >= 1
    assert res_all["counts"]["whatsapp"] >= 1

    # Filter by project_id=pid2 -> only P2 results
    st_p2, res_p2 = _http_json(f"{base}/api/search?q=Switchgear&project_id={pid2}")
    assert st_p2 == 200
    assert all(x["project_id"] == pid2 for x in res_p2["results"])

    # WhatsApp share resolves archive files cleanly via area='archive'
    st_wa_arc, r_wa_arc = _http_json(
        f"{base}/api/projects/{pid2}/whatsapp/send",
        method="POST",
        body={"mode": "web", "phone": "+96599990001", "message": "Photo", "rel": "02_Letters/IN-099_Site_Generator.png", "area": "archive", "copy_to_clipboard": False},
    )
    assert st_wa_arc == 200 and r_wa_arc["ok"] is True

    # 2) Feature 4: Project detail rows automatically link `file_rel` & preview endpoints work
    st_d, detail = _http_json(f"{base}/api/projects/{pid1}")
    assert st_d == 200
    row_001 = next(r for r in detail["recent_rows"] if r.get("ref_no") == "MAR-EL-001")
    assert row_001["file_rel"].endswith("MAR-EL-001_2026-11-01_Switchgear.png")
    assert row_001["is_previewable"] is True

    rel_q = urllib.request.quote(row_001["file_rel"], safe="")
    st_pi, p_info = _http_json(f"{base}/api/projects/{pid1}/preview-info?rel={rel_q}")
    assert st_pi == 200 and p_info["ok"] is True
    assert p_info["kind"] == "image" and p_info["total_pages"] == 1

    st_pp, png_data, hdrs = _http_raw(f"{base}/api/projects/{pid1}/preview-page?rel={rel_q}&page=1&dpi=110")
    assert st_pp == 200 and png_data.startswith(b"\x89PNG")
    assert hdrs.get("x-total-pages") == "1"

    # Security: path traversal blocked
    st_bad, _ = _http_json(f"{base}/api/projects/{pid1}/preview-info?rel=../../etc/passwd")
    assert st_bad in (400, 403, 404)


def test_feature_2_issue_pdf_merge_and_feature_3_weekly_report(live_server):
    base, hub, tmp_path = live_server
    p_root = tmp_path / "Proj_PDF_Weekly"
    _, r = _http_json(f"{base}/api/projects", method="POST", body={"name": "Substation Alpha", "code": "SUB-A", "root": str(p_root)})
    pid = r["project"]["id"]

    _create_sample_log_xlsx(p_root / "3_Project_Logs" / "LOG.xlsx")
    tpl_path = p_root / "4_Project_Templates" / "MAR_Template.docx"
    _create_sample_docx_tpl(tpl_path)

    # Save ready template profile for MAR_Template.docx
    prof_payload = {
        "section": "submittal_mar",
        "section_confirmed": True,
        "display_name": "Submittal Form",
        "ref_prefix": "MAR",
        "default_discipline": "EL",
        "fields": [
            {"key": "ref_no", "label": "Ref", "role": "auto_system", "source": "system.next_ref", "target": {"kind": "docx_cell", "table_idx": 0, "row_idx": 0, "col_idx": 1}},
            {"key": "subject", "label": "Subject", "role": "manual", "input_type": "text", "required": True, "target": {"kind": "docx_cell", "table_idx": 0, "row_idx": 1, "col_idx": 1}},
        ],
    }
    st_pr, r_pr = _http_json(f"{base}/api/projects/{pid}/templates/MAR_Template.docx/profile", method="POST", body=prof_payload)
    assert st_pr == 200 and r_pr["ok"] is True

    # Feature 2: Issue template WITH export_pdf=True + PNG attachment
    att_b64 = base64.b64encode(_make_png_bytes((255, 166, 17))).decode("ascii")
    st_is, r_is = _http_json(
        f"{base}/api/projects/{pid}/templates/MAR_Template.docx/issue",
        method="POST",
        body={
            "discipline": "EL",
            "subject": "Main Breaker Relay",
            "values": {"subject": "Main Breaker Relay"},
            "export_pdf": True,
            "attachments": [{"name": "datasheet.png", "data_b64": att_b64}],
        },
    )
    assert st_is == 200 and r_is["ok"] is True
    assert r_is["output_rel"].endswith(".docx")
    assert r_is["pdf_rel"].endswith(".pdf") and r_is["pdf_pages"] >= 1
    assert r_is["pdf_download_url"]

    st_dl, pdf_bytes, _ = _http_raw(f"{base}{r_is['pdf_download_url']}")
    assert st_dl == 200 and pdf_bytes.startswith(b"%PDF-")

    # WhatsApp Web link generation accepts the issued PDF rel
    st_wa, r_wa = _http_json(
        f"{base}/api/projects/{pid}/whatsapp/send",
        method="POST",
        body={"mode": "web", "phone": "+96599990000", "message": "Submittal Ready", "rel": r_is["pdf_rel"], "copy_to_clipboard": False},
    )
    assert st_wa == 200 and r_wa["ok"] is True and "whatsapp" in (r_wa.get("url") or "").lower()

    # Feature 4 also previews the issued .docx cleanly as kind='office'
    docx_q = urllib.request.quote(r_is["output_rel"], safe="")
    st_ox, ox_info = _http_json(f"{base}/api/projects/{pid}/preview-info?rel={docx_q}")
    assert st_ox == 200 and ox_info["kind"] == "office" and ox_info["render"]

    # Feature 3: 1-Click Weekly Follow-up Report (Excel + PDF)
    st_wk, r_wk = _http_json(f"{base}/api/projects/{pid}/reports/weekly", method="POST", body={"export_pdf": True, "days_window": 7})
    assert st_wk == 200 and r_wk["ok"] is True
    sm = r_wk["summary"]
    assert sm["overdue_count"] >= 1
    assert sm["unreplied_count"] >= 1
    assert sm["rejected_count"] >= 1
    assert sm["recent_7d_count"] >= 2  # MAR-HVAC-003 + newly issued MAR-EL-004 + IN-015
    assert r_wk["xlsx_download_url"] and r_wk["pdf_download_url"]

    st_wx, wx_bytes, _ = _http_raw(f"{base}{r_wk['xlsx_download_url']}")
    assert st_wx == 200 and len(wx_bytes) > 500
    wb = openpyxl.load_workbook(io.BytesIO(wx_bytes))
    assert len(wb.sheetnames) == 4

    st_wp, wp_bytes, _ = _http_raw(f"{base}{r_wk['pdf_download_url']}")
    assert st_wp == 200 and wp_bytes.startswith(b"%PDF-")

    # Weekly PDF report opens inside inline document preview cleanly
    wk_q = urllib.request.quote(r_wk["pdf_rel"], safe="")
    st_wpi, wk_info = _http_json(f"{base}/api/projects/{pid}/preview-info?rel={wk_q}")
    assert st_wpi == 200 and wk_info["total_pages"] >= 1


def test_feature_5_backup_alert_and_google_drive_flow(live_server):
    base, hub, tmp_path = live_server
    p_root = tmp_path / "Proj_Bkp"
    fake_drive = tmp_path / "GoogleDrive" / "My Drive"
    fake_drive.mkdir(parents=True, exist_ok=True)

    # 1) Cloud folder detection discovers Google Drive candidate without mutating disk
    detected = actp.detect_cloud_backup_folders(extra_candidates=[fake_drive])
    assert any(d["provider"] == "google_drive" and str(fake_drive) in d["path"] for d in detected)
    assert not (fake_drive / "SmartProjects_Backups").exists()

    # 2) New project starts with status='never'
    _, r = _http_json(f"{base}/api/projects", method="POST", body={"name": "Drive Project", "code": "DRV1", "root": str(p_root)})
    pid = r["project"]["id"]
    _, st0 = _http_json(f"{base}/api/state")
    card0 = next(c for c in st0["projects"] if c["id"] == pid)
    assert card0["backup"]["status"] == "never" and card0["backup"]["stale"] is False

    # 3) Simulate 10-day-old project with no backup -> overdue alert triggered
    proj_obj = hub.get_project(pid)
    stale_check = actp.check_backup_status(
        proj_obj,
        warn_days=7,
        now_override=dt.datetime.now() + dt.timedelta(days=10),
    )
    assert stale_check["stale"] is True and stale_check["status"] == "overdue"

    # 4) Quick Backup with Google Drive destination saves local + external mirror and sets drive provider
    target_drive = str(fake_drive / "SmartProjects_Backups")
    st_bk, r_bk = _http_json(
        f"{base}/api/projects/{pid}/backup",
        method="POST",
        body={"copy_external": True, "external_backup_dir": target_drive, "save_external_dir": True},
    )
    assert st_bk == 200 and r_bk["ok"] is True
    assert r_bk["backup"]["external"]["copied"] is True
    assert Path(r_bk["backup"]["external"]["path"]).exists()

    _, st1 = _http_json(f"{base}/api/state")
    card1 = next(c for c in st1["projects"] if c["id"] == pid)
    assert card1["backup"]["status"] == "ok"
    assert card1["backup"]["external_provider"] == "google_drive"
    assert st1["settings"]["external_backup_dir"] == target_drive

    # 5) auto_backup_on_end setting persists cleanly
    st_gs, r_gs = _http_json(f"{base}/api/settings", method="POST", body={"auto_backup_on_end": True})
    assert st_gs == 200 and r_gs["settings"]["auto_backup_on_end"] is True


def test_v10_12_ui_markers():
    ui_txt = (Path(__file__).resolve().parent.parent / "doc_tools" / "dashboard_ui.html").read_text(encoding="utf-8")
    for sym in (
        "renderSearchView",
        "openInlineDocPreview",
        "openWeeklyReportModal",
        "_backupBadgeHtml",
        "quickBackupProject",
        "iss_export_pdf",
        "iss_merge_files",
        "btnGlobalSearch",
    ):
        assert sym in ui_txt, f"Missing v10.12.0 UI symbol: {sym}"
