from __future__ import annotations

import base64
import datetime as dt
import io
import json
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import openpyxl
import pytest

DOC_TOOLS = Path(__file__).resolve().parent.parent / "doc_tools"
if str(DOC_TOOLS) not in sys.path:
    sys.path.insert(0, str(DOC_TOOLS))

import inbox_pipeline as ip
import project_dashboard as pdash
import project_log as plog
import template_profile as tprof


def _xlsx_bytes(rows: list[list], sheet_title: str = "Submittals") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _http_json(url: str, *, method: str = "GET", body: dict | None = None) -> tuple[int, dict]:
    hdrs = {"Accept": "application/json"}
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        hdrs["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, method=method, headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, {"ok": False, "raw": raw}


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
        hub.toggle_watcher(False)
        srv.shutdown()
        srv.server_close()


def test_scanner_triage_readonly_approve_reject_requeue_bulk(live_server):
    base, hub, tmp_path = live_server
    inbox = tmp_path / "Shared_Scanner_Inbox"
    out_dir = tmp_path / "Custom_Project_Output"
    root = tmp_path / "Proj_Alpha"
    inbox.mkdir(parents=True)

    # Three distinct scanner documents (never moved during scan!)
    (inbox / "MAR_EL_042_Switchgear.txt").write_text(
        "Project Alpha PRJ-AL\nMaterial Approval Request MAR-EL-042\nMain Switchgear Panel 400A\nStatus: Approved",
        encoding="utf-8",
    )
    (inbox / "Letter_MEW_011.txt").write_text(
        "Project Alpha PRJ-AL\nOutgoing Letter OUT-011 to Ministry of Electricity\nSubject: Power Connection",
        encoding="utf-8",
    )
    (inbox / "IPC_03_Contractors.txt").write_text(
        "Project Alpha PRJ-AL\nInterim Payment Certificate IPC-003\nAmount: 12500.500 KWD",
        encoding="utf-8",
    )

    log_b = _xlsx_bytes([
        ["Ref", "Rev", "Subject", "Date", "Status", "File"],
        ["MAR-EL-001", "R00", "Initial", "2026-10-01", "A", ""],
    ])
    st, r_proj = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Project Alpha",
        "code": "PRJ-AL",
        "root": str(root),
        "inbox_dir": str(inbox),
        "archive_dir": str(out_dir),
        "set_active": True,
        "initial_log": {"name": "Alpha_Log.xlsx", "data_b64": base64.b64encode(log_b).decode()},
    })
    assert st == 201 and r_proj["ok"] is True
    pid = r_proj["project"]["id"]
    assert Path(r_proj["project"]["archive_dir"]) == out_dir.resolve()

    # 1) Scan Inbox -> strict read-only! Originals remain in inbox untouched
    st_s, res_s = _http_json(f"{base}/api/projects/{pid}/scanner/scan", method="POST", body={"blocking": True})
    assert st_s == 200 and res_s["ok"] is True
    assert res_s["scan_summary"]["added"] == 3
    assert len(res_s["pending"]) == 3
    assert res_s["inbox_exists"] is True
    assert res_s["folder_files_count"] == 3
    assert (inbox / "MAR_EL_042_Switchgear.txt").exists()
    assert (inbox / "Letter_MEW_011.txt").exists()
    assert (inbox / "IPC_03_Contractors.txt").exists()

    # 2) Re-scan does not duplicate items (SHA-256 cache)
    st_s2, res_s2 = _http_json(f"{base}/api/projects/{pid}/scanner/scan", method="POST", body={"blocking": True})
    assert st_s2 == 200 and res_s2["scan_summary"]["added"] == 0
    assert len(res_s2["pending"]) == 3

    # Check human summary and target path preview
    mar_item = next(x for x in res_s2["pending"] if "MAR_EL_042" in x["file_name"])
    assert mar_item["summary_headline"]
    assert mar_item["target_rel"]

    # 3) Inline file preview directly from scanner folder
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/scanner/file?id={mar_item['id']}", timeout=10) as resp:
        assert resp.status == 200
        assert b"Switchgear" in resp.read()

    # 4) Reject one item -> file stays in scanner folder, moves to rejected bucket
    ipc_item = next(x for x in res_s2["pending"] if "IPC_03" in x["file_name"])
    st_rej, res_rej = _http_json(f"{base}/api/projects/{pid}/scanner/reject", method="POST", body={"id": ipc_item["id"]})
    assert st_rej == 200 and res_rej["counts"]["rejected"] == 1
    assert (inbox / "IPC_03_Contractors.txt").exists()

    # Requeue puts it back to pending
    st_rq, res_rq = _http_json(f"{base}/api/projects/{pid}/scanner/requeue", method="POST", body={"id": ipc_item["id"]})
    assert st_rq == 200 and res_rq["counts"]["rejected"] == 0 and res_rq["counts"]["pending"] == 3

    # 5) Approve ONLY MAR item -> moves ONLY that file into custom archive_dir, leaves others in inbox!
    st_app, res_app = _http_json(
        f"{base}/api/projects/{pid}/scanner/approve",
        method="POST",
        body={
            "id": mar_item["id"],
            "overrides": {
                "section": "submittal_mar",
                "subject_ar": "اعتماد لوحة كهرباء رئيسية 400 أمبير",
                "ref_no": "MAR-EL-042",
            },
        },
    )
    assert st_app == 200 and res_app["counts"]["approved"] == 1 and res_app["counts"]["pending"] == 2
    assert not (inbox / "MAR_EL_042_Switchgear.txt").exists()
    assert (inbox / "Letter_MEW_011.txt").exists()
    assert (inbox / "IPC_03_Contractors.txt").exists()

    approved_entry = res_app["approved"][0]
    assert approved_entry["archive_path"]
    assert str(out_dir.resolve()) in approved_entry["archive_path"]
    assert Path(approved_entry["archive_path"]).exists()

    # File preview still works from its new archived location + download endpoint supports area=archive and fallback
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/scanner/file?id={mar_item['id']}", timeout=10) as resp:
        assert resp.status == 200
        assert b"Switchgear" in resp.read()
    arc_rel_q = urllib.parse.quote(approved_entry["archive_rel"], safe="/")
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/download?area=archive&rel={arc_rel_q}", timeout=10) as resp:
        assert resp.status == 200
        assert b"Switchgear" in resp.read()
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/download?rel={arc_rel_q}", timeout=10) as resp:
        assert resp.status == 200

    # 6) Bulk reject remaining items leaves them safe in scanner
    rem_ids = [x["id"] for x in res_app["pending"]]
    st_b, res_b = _http_json(f"{base}/api/projects/{pid}/scanner/reject", method="POST", body={"ids": rem_ids})
    assert st_b == 200 and res_b["resolution"]["count"] == 2
    assert res_b["counts"]["pending"] == 0 and res_b["counts"]["rejected"] == 2
    assert (inbox / "Letter_MEW_011.txt").exists()
    assert (inbox / "IPC_03_Contractors.txt").exists()


def test_custom_archive_dir_and_sections_crud(live_server):
    base, hub, tmp_path = live_server
    root = tmp_path / "Proj_Sec"
    out_dir = tmp_path / "Sec_Output"
    new_out = tmp_path / "Sec_Output_Updated"

    st, r_proj = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Sections Project",
        "code": "SEC-1",
        "root": str(root),
        "archive_dir": str(out_dir),
    })
    assert st == 201
    pid = r_proj["project"]["id"]

    # Update archive_dir via project settings
    st_u, r_u = _http_json(f"{base}/api/projects/{pid}/settings", method="POST", body={"archive_dir": str(new_out)})
    assert st_u == 200
    assert Path(r_u["project"]["archive_dir"]) == new_out.resolve()

    # Fetch default sections, hide one, add a custom section
    st_g, r_g = _http_json(f"{base}/api/projects/{pid}/sections")
    assert st_g == 200 and len(r_g["sections"]) >= 10
    draft = list(r_g["all_sections"])
    for item in draft:
        if item["id"] == "petty_cash":
            item["enabled"] = False
    draft.append({
        "id": "custom_rfis",
        "label": "استفسارات فنية خاصة",
        "ref_prefix": "RFI",
        "folder": "Custom/01_RFI",
        "enabled": True,
        "is_custom": True,
    })
    st_p, r_p = _http_json(f"{base}/api/projects/{pid}/sections", method="POST", body={"sections": draft})
    assert st_p == 200
    vis_ids = [s["id"] for s in r_p["sections"]]
    assert "custom_rfis" in vis_ids
    assert "petty_cash" not in vis_ids

    # Custom section route resolves cleanly without 404
    st_cs, r_cs = _http_json(f"{base}/api/projects/{pid}/sections/custom_rfis")
    assert st_cs == 200 and r_cs["section"]["ref_prefix"] == "RFI"

    # Non-existent section returns 404 without raw KeyError quotes
    st_404, r_404 = _http_json(f"{base}/api/projects/{pid}/sections/does_not_exist")
    assert st_404 == 404 and not str(r_404.get("error", "")).startswith("'")


def test_template_auto_date_and_auto_counter_roles_and_issue(live_server):
    base, hub, tmp_path = live_server
    root = tmp_path / "Proj_Tpl"
    out_dir = tmp_path / "Tpl_Output"
    st, r_proj = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Sabah Al-Ahmad Project",
        "code": "SABAH",
        "root": str(root),
        "archive_dir": str(out_dir),
    })
    assert st == 201
    pid = r_proj["project"]["id"]

    # Upload an Excel template
    tpl_bytes = _xlsx_bytes([
        ["Ref:", "", "Date:", ""],
        ["Subject:", "", "", ""],
    ], sheet_title="Form")
    st_up, r_up = _http_json(
        f"{base}/api/projects/{pid}/templates/upload",
        method="POST",
        body={
            "name": "Electrical_MAR_Form.xlsx",
            "data_b64": base64.b64encode(tpl_bytes).decode(),
            "section": "submittal_mar",
            "display_name": "نموذج اعتماد كهرباء",
        },
    )
    assert st_up == 201

    # Configure auto_counter + auto_date + manual fields
    st_prof, r_prof = _http_json(
        f"{base}/api/projects/{pid}/templates/Electrical_MAR_Form.xlsx/profile",
        method="POST",
        body={
            "section": "submittal_mar",
            "section_confirmed": True,
            "ref_prefix": "SABAH-EL",
            "log_on_issue": False,
            "fields": [
                {
                    "key": "custom_ref",
                    "label": "الرقم المتسلسل",
                    "role": "auto_counter",
                    "source": "counter.custom",
                    "counter_prefix": "SABAH-EL-",
                    "counter_next": 5,
                    "counter_digits": 3,
                    "counter_suffix": "",
                    "sync_with_log": False,
                    "target": {"kind": "xlsx_cell", "sheet": "Form", "coord": "B1"},
                },
                {
                    "key": "issue_dt",
                    "label": "تاريخ الإصدار",
                    "role": "auto_date",
                    "source": "date.dd_mm_yyyy",
                    "target": {"kind": "xlsx_cell", "sheet": "Form", "coord": "D1"},
                },
                {
                    "key": "subj",
                    "label": "الموضوع",
                    "role": "manual",
                    "input_type": "text",
                    "required": True,
                    "target": {"kind": "xlsx_cell", "sheet": "Form", "coord": "B2"},
                },
            ],
        },
    )
    assert st_prof == 200 and r_prof["ok"] is True

    # Preview verifies auto_counter + auto_date rendering
    st_pv, r_pv = _http_json(
        f"{base}/api/projects/{pid}/templates/Electrical_MAR_Form.xlsx/preview",
        method="POST",
        body={"values": {"subj": "لوحة توزيع فرعية"}, "try_libreoffice": False},
    )
    assert st_pv == 200
    assert r_pv["resolved_fields"]["custom_ref"] == "SABAH-EL-005"
    assert r_pv["resolved_fields"]["issue_dt"] == dt.date.today().strftime("%d/%m/%Y")

    # Issue writes to custom archive_dir AND increments auto_counter from 5 -> 6!
    st_iss, r_iss = _http_json(
        f"{base}/api/projects/{pid}/templates/Electrical_MAR_Form.xlsx/issue",
        method="POST",
        body={"subject": "لوحة توزيع فرعية", "values": {"subj": "لوحة توزيع فرعية"}},
    )
    assert st_iss == 200
    assert Path(r_iss["output_path"]).exists()
    assert str(out_dir.resolve()) in r_iss["output_path"]
    assert "area=archive" in r_iss["download_url"]

    wb = openpyxl.load_workbook(r_iss["output_path"])
    ws = wb["Form"]
    assert ws["B1"].value == "SABAH-EL-005"
    assert ws["D1"].value == dt.date.today().strftime("%d/%m/%Y")
    assert ws["B2"].value == "لوحة توزيع فرعية"

    # Re-fetch profile -> next counter is now 6!
    st_p2, r_p2 = _http_json(f"{base}/api/projects/{pid}/templates/Electrical_MAR_Form.xlsx/profile")
    assert st_p2 == 200
    c_fld = next(f for f in r_p2["profile"]["fields"] if f["role"] == "auto_counter")
    assert int(c_fld["counter_next"]) == 6


def test_logmap_skip_sheet_forces_disabled_and_preview_picks_enabled(tmp_path: Path):
    proj_root = tmp_path / "Proj_Log"
    logs_dir = plog.logs_dir_for_project(proj_root)
    logs_dir.mkdir(parents=True, exist_ok=True)
    lp = logs_dir / "Master.xlsx"

    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Dashboard_Summary"
    ws_sum.append(["Total Submittals", 42])
    ws_data = wb.create_sheet("MAR_Log")
    ws_data.append(["Ref", "Subject", "Date", "Status", "Link"])
    ws_data.append(["MAR-001", "Cables", "2026-10-01", "A", ""])
    wb.save(lp)

    # Even if enabled=True is sent with layout="skip", save_logmap enforces enabled=False
    saved = plog.save_logmap(lp, {
        "sheets": [
            {"sheet": "Dashboard_Summary", "layout": "skip", "enabled": True, "header_row": 1, "data_start_row": 2},
            {"sheet": "MAR_Log", "layout": "flat", "enabled": True, "header_row": 1, "data_start_row": 2},
        ]
    })
    sh_map = {s["sheet"]: s for s in saved["sheets"]}
    assert sh_map["Dashboard_Summary"]["enabled"] is False
    assert sh_map["MAR_Log"]["enabled"] is True

    # Default preview auto-selects the first enabled non-skip sheet ("MAR_Log")
    prev = plog.preview_log_sheet(proj_root, "Master.xlsx")
    assert prev["active_sheet"] == "MAR_Log"


def test_hub_build_tag_start_sync_keys_and_404_error_message(live_server):
    base, hub, tmp_path = live_server

    # Hub state carries v10.13.2 build_tag and auto_date_sources
    st_h, r_h = _http_json(f"{base}/api/state")
    assert st_h == 200
    assert r_h["build_tag"] == "2026-11-18-v10.13.2"
    assert isinstance(r_h.get("auto_date_sources"), dict) and "date.today_iso" in r_h["auto_date_sources"]

    # Create project and call active-project with sync=True -> uses uploaded_logs & uploaded_count
    st_c, r_c = _http_json(f"{base}/api/projects", method="POST", body={"name": "Sync Proj", "code": "SP"})
    assert st_c == 201
    pid = r_c["project"]["id"]

    st_sy, r_sy = _http_json(f"{base}/api/active-project", method="POST", body={"project_id": pid, "sync": True})
    assert st_sy == 200 and r_sy["ok"] is True
    assert "uploaded_logs" in r_sy
    assert "uploaded_count" in (r_sy.get("templates") or {})

    # 404 project not found returns clean string without single quotes from KeyError.__str__
    st_404, r_404 = _http_json(f"{base}/api/projects/nonexistent_id_999")
    assert st_404 == 404
    err_msg = str(r_404.get("error") or "")
    assert err_msg and not (err_msg.startswith("'") and err_msg.endswith("'"))

    # UI HTML contains new Scanner view, browse folder, doc viewer expand/zoom, and zero legacy "بلا ربط" text
    with urllib.request.urlopen(f"{base}/", timeout=10) as resp:
        html = resp.read().decode("utf-8")
        assert "renderScannerView" in html
        assert "openManageSectionsModal" in html
        assert "browseFolderForInput" in html
        assert "toggleDocViewerFullscreen" in html
        assert "stepDocViewerZoom" in html
        assert "doc-viewer-fs" in html
        assert "بلا ربط" not in html


def test_windows_server_network_share_and_readonly_resilience(tmp_path: Path, monkeypatch):
    import builtins
    import relevance as rel
    import scan_watcher as sw
    import security_utils as su

    # 1) Windows extended UNC normalization
    assert str(su.normalize_win_path(r"\\?\UNC\WIN-SRV01\Scans\ProjectA\doc.pdf")) == r"\\WIN-SRV01\Scans\ProjectA\doc.pdf"
    assert str(su.normalize_win_path(r"\\?\C:\Projects\doc.pdf")) == r"C:\Projects\doc.pdf"

    # 2) Read-only SMB share behavior: PermissionError on "r+b" while "rb" succeeds -> readable for scanner, locked for Excel write
    f_scan = tmp_path / "scan_01.txt"
    f_scan.write_text("MAR-001 Main Switchgear", encoding="utf-8")
    real_open = builtins.open

    def _ro_open(file, mode="r", *args, **kwargs):
        if Path(str(file)) == f_scan and "+" in str(mode):
            raise PermissionError(13, "Read-only SMB share")
        return real_open(file, mode, *args, **kwargs)

    monkeypatch.setattr(builtins, "open", _ro_open)
    assert rel.is_file_locked(f_scan, allow_readonly=True) is False
    assert rel.is_file_locked(f_scan, allow_readonly=False) is True
    monkeypatch.undo()

    # 3) _iter_inbox_supported_files skips Windows Server system folders ($RECYCLE.BIN, System Volume Information, DfsrPrivate)
    net_inbox = tmp_path / "Server_Scan_Share"
    (net_inbox / "$RECYCLE.BIN").mkdir(parents=True)
    (net_inbox / "$RECYCLE.BIN" / "deleted_scan.txt").write_text("junk", encoding="utf-8")
    (net_inbox / "System Volume Information").mkdir(parents=True)
    (net_inbox / "System Volume Information" / "index.txt").write_text("sys", encoding="utf-8")
    (net_inbox / "DfsrPrivate").mkdir(parents=True)
    (net_inbox / "DfsrPrivate" / "tmp.txt").write_text("sys", encoding="utf-8")
    (net_inbox / "Valid_Subfolder").mkdir(parents=True)
    (net_inbox / "Valid_Subfolder" / "Letter_01.txt").write_text("Project X OUT-001", encoding="utf-8")

    found = ip._iter_inbox_supported_files(net_inbox, {".txt", ".pdf"})
    assert [p.name for p in found] == ["Letter_01.txt"]

    # 4) Scanner scan + approve succeeds even if src.unlink() fails on read-only server share
    proj_root = tmp_path / "Proj_Net"
    arc_dir = tmp_path / "Arc_Net"
    proj_root.mkdir()
    proj = {
        "id": "prj_net",
        "name": "Network Test",
        "code": "NET",
        "root": str(proj_root),
        "inbox_dir": str(net_inbox),
        "archive_dir": str(arc_dir),
    }
    log_dir = plog.logs_dir_for_project(proj_root)
    log_dir.mkdir(parents=True, exist_ok=True)
    (log_dir / "Net_Log.xlsx").write_bytes(_xlsx_bytes([
        ["Ref", "Rev", "Subject", "Date", "Status", "File"],
        ["OUT-000", "R00", "Seed", "2026-10-01", "A", ""],
    ], sheet_title="Letters"))

    scan_res = ip.scan_inbox_to_triage(proj, all_projects=[proj], max_files=20)
    assert scan_res["added"] == 1
    tstate = ip.get_scanner_triage_state(proj)
    tid = tstate["pending"][0]["id"]

    orig_unlink = Path.unlink
    def _deny_unlink(self, missing_ok=False):
        if self.name == "Letter_01.txt":
            raise PermissionError(13, "Read-only share cannot delete")
        return orig_unlink(self, missing_ok=missing_ok)

    monkeypatch.setattr(Path, "unlink", _deny_unlink)
    app_res = ip.resolve_triage_item(
        proj, tid, action="approve",
        overrides={"section": "outgoing", "subject": "Test Net Letter", "ref_no": "OUT-001"},
    )
    assert app_res["ok"] is True
    assert app_res["item"]["status"] == "approved"
    assert ip.get_scanner_triage_state(proj)["counts"]["approved"] == 1
    assert any(arc_dir.rglob("*.txt"))
    monkeypatch.undo()

    # 5) read_project_log_summary(limit=0) caches network Excel reads until file changes
    s1 = plog.read_project_log_summary(proj, limit=0)
    assert s1["stats"]["total_rows"] >= 1
    s2 = plog.read_project_log_summary(proj, limit=0)
    assert s2["stats"]["total_rows"] == s1["stats"]["total_rows"]
    # Verify scan_watcher helper accepts read-only share files too
    assert sw.SCANNED_EXTS


def test_direct_linked_external_excel_log_updates_in_place(live_server):
    base, hub, tmp_path = live_server
    ext_dir = tmp_path / "User_Desktop_Folder"
    ext_dir.mkdir()
    ext_log = ext_dir / "Desktop_Master_Log.xlsx"
    ext_log.write_bytes(_xlsx_bytes([
        ["Ref", "Rev", "Subject", "Date", "Status", "File"],
        ["MAR-EL-001", "R00", "Main Panel", "2026-10-01", "Pending", ""],
    ], sheet_title="Submittals"))

    # Create project without uploading log
    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Direct Link Proj",
        "code": "DLP",
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]
    proj = hub.get_project(pid)
    proj_root = Path(proj["root"])

    # Link the external file directly via API
    st_l, r_l = _http_json(f"{base}/api/projects/{pid}/logs/link", method="POST", body={"path": str(ext_log)})
    assert st_l == 200 and r_l["ok"] is True
    assert r_l["log_name"] == "Desktop_Master_Log.xlsx"

    # Original file is NOT copied into 3_Project_Logs, and user's desktop folder is NOT polluted with sidecars
    assert not (proj_root / "3_Project_Logs" / "Desktop_Master_Log.xlsx").exists()
    assert not (ext_dir / "Desktop_Master_Log.xlsx.logmap.json").exists()
    assert (proj_root / "3_Project_Logs" / ".linked_Desktop_Master_Log.xlsx.logmap.json").exists()

    # Listed as linked with exact path
    st_list, r_list = _http_json(f"{base}/api/projects/{pid}/logs")
    assert st_list == 200 and len(r_list["logs"]) == 1
    assert r_list["logs"][0]["is_linked"] is True
    assert Path(r_list["logs"][0]["path"]) == ext_log.resolve()

    # Map endpoint preview & save work seamlessly on linked file
    st_m, r_m = _http_json(f"{base}/api/projects/{pid}/logs/Desktop_Master_Log.xlsx/map")
    assert st_m == 200 and r_m["is_linked"] is True
    assert len(r_m["rows"]) >= 2

    # Writing a record updates the ORIGINAL external Excel file IN PLACE!
    rec_res = plog.record_entry_in_project_log(proj, {
        "section": "submittal",
        "ref_no": "MAR-EL-002",
        "revision": "R00",
        "subject": "Backup Generator",
        "doc_date": "2026-11-18",
        "status": "Under Review",
    })
    assert rec_res["ok"] is True and rec_res["queued"] is False

    wb_after = openpyxl.load_workbook(ext_log, data_only=True)
    vals = [[c for c in row] for row in wb_after["Submittals"].iter_rows(values_only=True)]
    wb_after.close()
    assert any("MAR-EL-002" in str(r) for r in vals), f"Expected MAR-EL-002 written into {ext_log}, got {vals}"

    # Unlinking removes it from project WITHOUT deleting or moving the user's original Excel file!
    st_u, r_u = _http_json(f"{base}/api/projects/{pid}/logs/Desktop_Master_Log.xlsx/archive", method="POST", body={})
    assert st_u == 200 and r_u["ok"] is True
    assert ext_log.exists()
    st_list2, r_list2 = _http_json(f"{base}/api/projects/{pid}/logs")
    assert st_list2 == 200 and len(r_list2["logs"]) == 0


def test_bulk_organizer_over_1200_files_dedup_superseded_execute_and_report(live_server):
    import zipfile
    base, hub, tmp_path = live_server

    # Linked external Excel log on user's machine
    ext_log = tmp_path / "External_Project_Log.xlsx"
    ext_log.write_bytes(_xlsx_bytes([
        ["Ref", "Rev", "Subject", "Date", "Status", "File"],
        ["MAR-EL-001", "R00", "Switchgear", "2026-10-01", "Approved", ""],
    ], sheet_title="Submittals"))

    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Mega Messy Project",
        "code": "MMP",
        "linked_log_path": str(ext_log),
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]

    # Create 1,200+ scattered files + 1 exact duplicate + 1 superseded revision + 1 ZIP archive
    messy_dir = tmp_path / "Messy_2000_Files"
    for b in range(12):
        sub = messy_dir / f"Old_PC_{b}" / "Misc"
        sub.mkdir(parents=True, exist_ok=True)
        for i in range(100):
            idx = b * 100 + i + 1
            (sub / f"MAR-EL-{idx:04d}_R00_Cable_{idx}.txt").write_bytes(f"doc-{idx}".encode())

    extra_dir = messy_dir / "Unsorted"
    extra_dir.mkdir(parents=True, exist_ok=True)
    dup_payload = b"UNIQUE-DUPLICATE-PAYLOAD-XYZ-999"
    (extra_dir / "MIR-PL-0099_Pipe_Orig.txt").write_bytes(dup_payload)
    (messy_dir / "Old_PC_0" / "MIR-PL-0099_Copy.txt").write_bytes(dup_payload)
    (extra_dir / "WIR-CIV-0500_R00_Slab.txt").write_bytes(b"slab rev0")
    (extra_dir / "WIR-CIV-0500_R01_Slab.txt").write_bytes(b"slab rev1")

    zpath = messy_dir / "Batch_Bundle.zip"
    with zipfile.ZipFile(zpath, "w") as zf:
        zf.writestr("Nested/SD-STR-0700_R00_Beam.txt", b"beam shop drawing")

    # 1) Scan unlimited files via API
    st_s, r_s = _http_json(f"{base}/api/projects/{pid}/organizer/scan", method="POST", body={
        "folder": str(messy_dir),
        "extract_archives": True,
        "read_text": False,
        "blocking": True,
    })
    assert st_s == 200 and r_s["ok"] is True
    sm = r_s["summary"]
    assert sm["total"] >= 1205
    assert sm["counts"]["duplicate"] >= 1
    assert sm["counts"]["superseded"] >= 1

    # 2) Paginated review fetch (page 2 of 50)
    st_pg, r_pg = _http_json(f"{base}/api/projects/{pid}/organizer?page=2&per_page=50&filter=all")
    assert st_pg == 200 and r_pg["ok"] is True
    assert r_pg["page"] == 2 and len(r_pg["items"]) == 50
    assert r_pg["total_pages"] >= 25

    # Check duplicate & superseded filters work and start as `pending`
    st_dp, r_dp = _http_json(f"{base}/api/projects/{pid}/organizer?filter=duplicate")
    assert st_dp == 200 and len(r_dp["items"]) >= 1
    assert all(it.get("decision") == "pending" for it in r_dp["items"])
    st_sp, r_sp = _http_json(f"{base}/api/projects/{pid}/organizer?filter=superseded")
    assert st_sp == 200 and any("WIR-CIV-0500_R00" in it["file_name"] for it in r_sp["items"])

    # 3) Bulk exclude duplicates & isolate superseded -> clears pending duplicate/superseded view
    st_u1, r_u1 = _http_json(f"{base}/api/projects/{pid}/organizer/update", method="POST", body={
        "bulk_action": "exclude_duplicates",
    })
    assert st_u1 == 200 and r_u1["counts"]["duplicate"] == 0
    assert r_u1["counts"]["duplicate_resolved"] >= 1
    st_dp_after, r_dp_after = _http_json(f"{base}/api/projects/{pid}/organizer?filter=duplicate")
    assert st_dp_after == 200 and r_dp_after["items"] == []
    st_dp_res, r_dp_res = _http_json(f"{base}/api/projects/{pid}/organizer?filter=duplicate&show_resolved=1")
    assert st_dp_res == 200 and len(r_dp_res["items"]) >= 1 and r_dp_res["items"][0]["decision"] == "excluded"

    st_u2, r_u2 = _http_json(f"{base}/api/projects/{pid}/organizer/update", method="POST", body={
        "bulk_action": "exclude_superseded",
    })
    assert st_u2 == 200 and r_u2["counts"]["superseded"] == 0
    st_sp_after, r_sp_after = _http_json(f"{base}/api/projects/{pid}/organizer?filter=superseded")
    assert st_sp_after == 200 and r_sp_after["items"] == []

    # Edit one item manually to verify inline editing persists
    target_id = r_pg["items"][0]["id"]
    st_u3, _ = _http_json(f"{base}/api/projects/{pid}/organizer/update", method="POST", body={
        "updates": [{"id": target_id, "section": "ncr", "ref_no": "NCR-9999"}],
    })
    assert st_u3 == 200

    # 4) Execute safe copy + report -> resets plan items to empty while preserving last_execution
    st_ex, r_ex = _http_json(f"{base}/api/projects/{pid}/organizer/execute", method="POST", body={
        "mode": "copy",
        "update_excel_log": False,
        "blocking": True,
    })
    assert st_ex == 200 and r_ex["ok"] is True
    ex = r_ex["execution"]
    assert ex["copied"] >= 1203
    assert ex["excluded"] >= 2
    assert ex.get("cleared") is True
    st_post_ex, r_post_ex = _http_json(f"{base}/api/projects/{pid}/organizer")
    assert st_post_ex == 200 and r_post_ex["has_plan"] is False and r_post_ex["total"] == 0
    assert r_post_ex["last_execution"]["copied"] == ex["copied"]
    arc_root = Path(hub.get_project(pid)["archive_dir"])
    assert not any(arc_root.rglob("*WIR-CIV-0500_R00*"))
    assert any(arc_root.rglob("*NCR-9999*"))
    rep_path = Path(hub.get_project(pid)["root"]) / ex["report_rel"]
    assert rep_path.exists()
    wb_rep = openpyxl.load_workbook(rep_path, data_only=True)
    assert len(wb_rep.sheetnames) == 4
    wb_rep.close()

    # 5) Verify live walking/extracting/hashing/analyzing progress counters & UI host
    import project_organizer as porg
    events = []
    porg.build_organization_plan(
        hub.get_project(pid),
        str(messy_dir),
        extract_archives=True,
        read_text=False,
        progress_cb=lambda ev: events.append(dict(ev)),
    )
    walk_evs = [e for e in events if e.get("phase") == "walking"]
    assert walk_evs and walk_evs[-1]["discovered"] >= 1200 and walk_evs[-1]["folders_scanned"] >= 12
    assert any(e.get("current_file") for e in walk_evs)
    assert any(e.get("phase") == "extracting" and e.get("current_file") == "Batch_Bundle.zip" for e in events)
    assert any(e.get("phase") == "analyzing" and e.get("scanned", 0) >= 1200 for e in events)
    with urllib.request.urlopen(f"{base}/", timeout=5) as idx:
        html_src = idx.read().decode("utf-8")
        assert "orgProgressHost" in html_src and "_renderOrgProgressHtml" in html_src
        assert "organizer.phase_walking" in html_src
    # Original messy files remain untouched in copy mode
    assert zpath.exists()


def test_audit_weekly_followup_sequence_gaps_register_unlogged_and_pdf_tools(live_server):
    base, hub, tmp_path = live_server

    ext_log = tmp_path / "Audit_Log.xlsx"
    ext_log.write_bytes(_xlsx_bytes([
        ["Ref", "Rev", "Subject", "Date", "Status", "File"],
        ["MAR-EL-001", "R00", "Panel Board", "2026-10-01", "Under Review", ""],
        ["MAR-EL-003", "R00", "Lighting", "2026-11-10", "Code C - Revise", ""],
        ["MAR-EL-005", "R00", "Earthing", "2026-11-11", "Approved", ""],
    ], sheet_title="Submittals"))

    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Audit PDF Proj",
        "code": "AUD",
        "linked_log_path": str(ext_log),
        "overdue_days": 5,
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]
    proj = hub.get_project(pid)
    arc = Path(proj["archive_dir"])

    # Put MAR-EL-001 in archive AND one unlogged file MAR-EL-004
    sub_dir = arc / "01_Submittals"
    sub_dir.mkdir(parents=True, exist_ok=True)
    (sub_dir / "MAR-EL-001_R00_Panel.txt").write_bytes(b"panel")
    (sub_dir / "MAR-EL-004_R00_Socket.txt").write_bytes(b"socket unlogged")

    # 1) Run standalone audit & weekly report
    st_au, r_au = _http_json(f"{base}/api/projects/{pid}/organizer/audit", method="POST", body={})
    assert st_au == 200 and r_au["ok"] is True
    assert r_au["overdue_count"] >= 1
    assert r_au["rejected_count"] >= 1
    assert r_au["log_missing_files_count"] >= 1  # MAR-EL-003 / 005 have no archive file
    assert r_au["archive_unlogged_count"] >= 1   # MAR-EL-004 is not in Excel log
    el_gaps = [g for g in (r_au["sequence_gaps"] or []) if g["series"] == "MAR-EL"]
    assert el_gaps and 2 in el_gaps[0]["missing_numbers"]
    assert (Path(proj["root"]) / r_au["report_rel"]).exists()

    # 2) Batch register unlogged files into linked Excel log in-place
    st_reg, r_reg = _http_json(f"{base}/api/projects/{pid}/organizer/register-unlogged", method="POST", body={})
    assert st_reg == 200 and r_reg["ok"] is True
    assert r_reg["added"] >= 1

    wb = openpyxl.load_workbook(ext_log, data_only=True)
    rows_txt = "\n".join(str(r) for r in wb["Submittals"].iter_rows(values_only=True))
    wb.close()
    assert "MAR-EL-004" in rows_txt

    # 3) PDF split & merge test (uses PyMuPDF when available, otherwise clean error)
    try:
        import fitz  # type: ignore
        has_fitz = True
    except Exception:
        has_fitz = False

    sample_pdf = tmp_path / "scan_3pages.pdf"
    if has_fitz:
        doc = fitz.open()
        for pg_i in range(3):
            page = doc.new_page()
            page.insert_text((72, 72), f"Page {pg_i + 1}")
        doc.save(str(sample_pdf))
        doc.close()

        split_out = tmp_path / "SplitOut"
        st_sp, r_sp = _http_json(f"{base}/api/projects/{pid}/organizer/pdf-split", method="POST", body={
            "src_pdf": str(sample_pdf),
            "dest_dir": str(split_out),
            "every_n_pages": 1,
        })
        assert st_sp == 200 and r_sp["ok"] is True
        assert r_sp["parts_created"] == 3 and len(r_sp["files"]) == 3

        merged_pdf = tmp_path / "Merged.pdf"
        st_mg, r_mg = _http_json(f"{base}/api/projects/{pid}/organizer/pdf-merge", method="POST", body={
            "src_pdfs": [str(split_out / fn) for fn in r_sp["files"]],
            "out_pdf": str(merged_pdf),
        })
        assert st_mg == 200 and r_mg["ok"] is True
        assert r_mg["total_pages"] == 3 and merged_pdf.exists()


def test_dual_stamps_consultant_and_ministry_undersecretary_engineer_dates(live_server):
    """
    Field workflow test:
      1) Returned sheet stamped by Ministry Engineer / Undersecretary (not just 'MOH' text)
      2) Returned sheet carrying BOTH Consultant + Ministry stamps with two separate reply dates
      3) Optional standalone `MOH Reply Date` column vs Remarks fallback when absent
    """
    base, hub, tmp_path = live_server

    # Multi-rev workbook matching SABAH ALSALEM style + optional MOH Reply Date on sheet EL
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EL"
    ws.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة", "MOH / الوزارة", "MOH Reply Date", "Remarks / ملاحظات"])
    ws.append(["", "", "", "", "Submit", "Date Reply", "Action", "", "", "", ""])
    ws.append([1, "SD", "EL-101", "Main Breaker", "2026-11-01", "2026-11-03", "approved", "", "", "", "ملاحظة سابقة"])
    wb_path = tmp_path / "DualStamp_Log.xlsx"
    wb.save(wb_path)
    wb.close()

    # Second workbook WITHOUT MOH Reply Date column (like LOG_Final_Updated.xlsm)
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "CIV"
    ws2.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة", "MOH / الوزارة", "Remarks / ملاحظات"])
    ws2.append(["", "", "", "", "Submit", "Date Reply", "Action", "", "", ""])
    ws2.append([1, "SD", "CIV-202", "Foundation", "2026-11-01", "", "", "", "", ""])
    wb2_path = tmp_path / "DualStamp_NoMohDateCol.xlsx"
    wb2.save(wb2_path)
    wb2.close()

    # Case A: Returned from Undersecretary / Ministry Engineer ONLY -> updates MOH columns, preserves REV. 0 consultant date!
    st_p1, r_p1 = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Dual Stamp A",
        "code": "DSA",
        "linked_log_path": str(wb_path),
        "set_active": True,
    })
    assert st_p1 == 201
    proj1 = hub.get_project(r_p1["project"]["id"])

    res_a = plog.record_entry_in_project_log(proj1, {
        "section": "submittal",
        "discipline": "EL",
        "ref_no": "EL-101",
        "revision": "R00",
        "authority": "وكيل الوزارة المساعد - مهندس الوزارة",
        "reply_date": "2026-11-16",
        "status": "Approved",
    })
    assert res_a["ok"] is True and res_a["is_moh_update"] is True
    wba = openpyxl.load_workbook(wb_path, data_only=True)
    wsa = wba["EL"]
    assert "2026-11-03" in str(wsa.cell(3, 6).value)  # Consultant date preserved!
    assert "approved" in str(wsa.cell(3, 9).value).lower() or "✅" in str(wsa.cell(3, 9).value)
    assert "2026-11-16" in str(wsa.cell(3, 10).value)  # MOH Reply Date populated!
    assert "ملاحظة سابقة" in str(wsa.cell(3, 11).value)
    wba.close()

    # Case B: Single returned sheet carrying BOTH Consultant + Ministry stamps on LOG_Final-style sheet (no dedicated MOH date col)
    st_p2, r_p2 = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Dual Stamp B",
        "code": "DSB",
        "linked_log_path": str(wb2_path),
    })
    assert st_p2 == 201
    proj2 = hub.get_project(r_p2["project"]["id"])

    res_b = plog.record_entry_in_project_log(proj2, {
        "section": "submittal",
        "discipline": "CIV",
        "ref_no": "CIV-202",
        "revision": "R00",
        "authority": "MOH+CONSULTANT",
        "has_both_stamps": True,
        "consultant_reply_date": "2026-11-04",
        "consultant_action": "Code B",
        "moh_reply_date": "2026-11-15",
        "moh_action": "Approved",
    })
    assert res_b["ok"] is True and res_b["is_moh_update"] is True
    wbb = openpyxl.load_workbook(wb2_path, data_only=True)
    wsb = wbb["CIV"]
    assert "2026-11-04" in str(wsb.cell(3, 6).value)   # Consultant reply date written in REV. 0!
    assert str(wsb.cell(3, 7).value).strip() != ""     # Consultant action written!
    assert "✅" in str(wsb.cell(3, 9).value) or "approved" in str(wsb.cell(3, 9).value).lower()
    assert "2026-11-15" in str(wsb.cell(3, 10).value)  # Ministry reply date preserved cleanly in Remarks!
    wbb.close()


def test_blank_reply_date_highlighted_amber_flagged_and_completed_manually(live_server):
    base, hub, tmp_path = live_server

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EL"
    ws.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة", "MOH / الوزارة", "Remarks / ملاحظات"])
    ws.append(["", "", "", "", "Submit", "Date Reply", "Action", "", "", ""])
    ws.append([1, "MAR", "MAR-EL-777", "Cable Tray", "2026-11-10", "", "", "", "", ""])
    log_p = tmp_path / "ManualDate_Log.xlsx"
    wb.save(log_p)
    wb.close()

    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Manual Reply Date Proj",
        "code": "MRD",
        "linked_log_path": str(log_p),
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]
    proj = hub.get_project(pid)

    # 1) Record reply decision WITHOUT a readable reply date -> cell stays blank + amber fill + flag!
    rec = plog.record_entry_in_project_log(proj, {
        "section": "submittal",
        "discipline": "EL",
        "ref_no": "MAR-EL-777",
        "revision": "R00",
        "authority": "SOOR",
        "reply_date": "",
        "status": "Code B",
    })
    assert rec["ok"] is True and rec["needs_manual_reply_date"] is True

    wb1 = openpyxl.load_workbook(log_p)
    ws1 = wb1["EL"]
    assert ws1.cell(3, 6).value in (None, "")  # Date Reply left BLANK (never guessed!)
    assert str(ws1.cell(3, 6).fill.start_color.rgb).upper().endswith("FEF3C7")  # Highlighted amber!
    assert "مطلوب تاريخ رد" in str(ws1.cell(3, 10).value or "")
    wb1.close()

    # 2) Appears in API missing_reply_rows for manual entry
    st_lg, r_lg = _http_json(f"{base}/api/projects/{pid}/logs")
    assert st_lg == 200 and r_lg["missing_reply_dates"] >= 1
    row_info = next(x for x in r_lg["missing_reply_rows"] if x["ref_no"] == "MAR-EL-777")
    assert row_info["missing_consultant_date"] is True

    # 3) User enters the date manually via UI API -> fills cell, clears amber fill & flag!
    st_fix, r_fix = _http_json(f"{base}/api/projects/{pid}/logs/set-reply-date", method="POST", body={
        "log_name": row_info["log_name"],
        "sheet": row_info["sheet"],
        "row": row_info["row"],
        "revision": row_info["revision"],
        "consultant_reply_date": "2026-11-17",
    })
    assert st_fix == 200 and r_fix["ok"] is True and r_fix["missing_reply_dates"] == 0

    wb2 = openpyxl.load_workbook(log_p)
    ws2 = wb2["EL"]
    assert "2026-11-17" in str(ws2.cell(3, 6).value)
    assert not str(getattr(ws2.cell(3, 6).fill.start_color, "rgb", "") or "").upper().endswith("FEF3C7")
    assert "مطلوب تاريخ رد" not in str(ws2.cell(3, 10).value or "")
    wb2.close()


def test_rev0_is_always_first_submission_slot_not_rev1(live_server):
    """
    Ensure REV. 0 is always recorded as the first submission for a reference number:
      1) Brand-new ref tagged R01/Rev.01 (like PL-232) goes to REV. 0, leaving REV. 1 blank.
      2) Existing placeholder row with blank REV. 0 goes to REV. 0 even if tagged R01.
      3) Follow-up reply on an open 1-based REV. 0 updates REV. 0 cleanly (not REV. 1).
      4) Genuine resubmission after rejection fills REV. 1 sequentially (whether R01 or R02),
         and subsequent R02 after REV. 1 rejection fills REV. 2.
    """
    base, hub, tmp_path = live_server

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PL"
    ws.append([
        "#", "Type", "Reference", "Description",
        "REV. 0", "", "", "REV. 1", "", "", "REV. 2", "", "",
        "Status / الحالة", "MOH / الوزارة", "Remarks / ملاحظات",
    ])
    ws.append(["", "", "", "", "Submit", "Date Reply", "Action", "Submit", "Date Reply", "Action", "Submit", "Date Reply", "Action", "", "", ""])
    # Pre-existing placeholder row with blank REV. 0
    ws.append([1, "MD", "PL-240", "Placeholder Valve", None, None, None, None, None, None, None, None, None, "", "", ""])
    log_p = tmp_path / "Rev0_First_Log.xlsx"
    wb.save(log_p)
    wb.close()

    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Rev0 First Proj",
        "code": "R0F",
        "linked_log_path": str(log_p),
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]
    proj = hub.get_project(pid)

    # 1) Brand-new PL-232 arriving tagged Rev.01 -> must go to REV. 0 (cols 5..7), NOT REV. 1 (cols 8..10)!
    r1 = plog.record_entry_in_project_log(proj, {
        "section": "submittal",
        "discipline": "PL",
        "subtype": "MD",
        "ref_no": "PL-232",
        "revision": "Rev.01",
        "submit_date": "2026-11-10",
        "reply_date": "2026-11-10",
        "approval_code": "B",
        "subject": "SAND FILTER - PENTAIR",
    })
    assert r1["ok"] is True and r1["mode"] == "inserted" and r1["recorded_revision"] == "R00"

    # 2) Existing row PL-240 with blank REV. 0 arriving as pending R01 -> goes to REV. 0;
    #    then follow-up return tagged R01 updates that same REV. 0!
    r2a = plog.record_entry_in_project_log(proj, {
        "section": "submittal",
        "discipline": "PL",
        "ref_no": "PL-240",
        "revision": "R01",
        "submit_date": "2026-11-01",
        "status": "pending",
    })
    assert r2a["ok"] is True and r2a["mode"] == "updated" and r2a["recorded_revision"] == "R00"

    r2b = plog.record_entry_in_project_log(proj, {
        "section": "submittal",
        "discipline": "PL",
        "ref_no": "PL-240",
        "revision": "R01",
        "reply_date": "2026-11-04",
        "status": "Rejected",
    })
    assert r2b["ok"] is True and r2b["recorded_revision"] == "R00"

    # 3) Resubmission of PL-240 tagged R02 (1-based) after REV. 0 rejection -> fills REV. 1 sequentially (no blank gap!)
    r2c = plog.record_entry_in_project_log(proj, {
        "section": "submittal",
        "discipline": "PL",
        "ref_no": "PL-240",
        "revision": "R02",
        "submit_date": "2026-11-08",
        "reply_date": "2026-11-09",
        "status": "Rejected",
    })
    assert r2c["ok"] is True and r2c["recorded_revision"] == "R01"

    # 4) Third submission tagged R02/R03 after both REV. 0 and REV. 1 are populated -> fills REV. 2!
    r2d = plog.record_entry_in_project_log(proj, {
        "section": "submittal",
        "discipline": "PL",
        "ref_no": "PL-240",
        "revision": "R02",
        "submit_date": "2026-11-12",
        "reply_date": "2026-11-14",
        "approval_code": "A",
    })
    assert r2d["ok"] is True and r2d["recorded_revision"] == "R02"

    wb_chk = openpyxl.load_workbook(log_p, data_only=True)
    ws_chk = wb_chk["PL"]
    # PL-240 row 3: REV.0 (rejected), REV.1 (rejected), REV.2 (approved)
    assert "2026-11-01" in str(ws_chk.cell(3, 5).value)
    assert "2026-11-04" in str(ws_chk.cell(3, 6).value)
    assert "reject" in str(ws_chk.cell(3, 7).value).lower()
    assert "2026-11-08" in str(ws_chk.cell(3, 8).value)
    assert "reject" in str(ws_chk.cell(3, 10).value).lower()
    assert "2026-11-12" in str(ws_chk.cell(3, 11).value)
    assert "approv" in str(ws_chk.cell(3, 13).value).lower()
    # PL-232 row 4: REV.0 populated, REV.1 completely blank!
    assert ws_chk.cell(4, 3).value == "PL-232"
    assert "2026-11-10" in str(ws_chk.cell(4, 5).value)
    assert "2026-11-10" in str(ws_chk.cell(4, 6).value)
    assert "approv" in str(ws_chk.cell(4, 7).value).lower()
    assert ws_chk.cell(4, 8).value is None
    assert ws_chk.cell(4, 9).value is None
    assert ws_chk.cell(4, 10).value is None
    wb_chk.close()


def test_correspondence_cross_party_topic_linking_without_ref_and_excel_column(live_server):
    """
    End-to-end correspondence cross-linking:
      - Consultant incoming, User outgoing, and Subcontractor incoming discuss the same
        topic ('كاميرات المراقبة' / CCTV) WITHOUT mentioning each other's ref numbers.
      - Conflicting ordinals (الدفعة 01 vs 02) and conflicting trades (كهرباء vs مدنية)
        must NOT falsely auto-link.
      - Auto-link transitively groups all 3 parties into one thread and writes
        'الكتاب المرتبط / Related Letter' in front of every letter row in Excel.
      - Unlink clears Excel and prevents auto-link from re-linking the unlinked pair.
    """
    base, hub, tmp_path = live_server
    import letter_links as ll

    wb = openpyxl.Workbook()
    ws_out = wb.active
    ws_out.title = "صادر -السور"
    ws_out.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة"])
    ws_out.append(["", "", "", "", "Submit", "Date Reply", "Action", ""])
    ws_out.merge_cells("H1:H2")
    ws_out.append([1, "OUT", "ج/2026/90", "بخصوص كاميرات المراقبة والتجهيزات الأمنية", "2026-11-12", "", "pending", "Under Review"])
    ws_out.append([2, "OUT", "ج/2026/27", "طلب صرف الدفعة رقم 01", "2026-11-01", "", "pending", "Under Review"])
    ws_out.append([3, "OUT", "ج/2026/37", "استبدال مراقب اعمال الكهرباء", "2026-11-02", "", "pending", "Under Review"])

    ws_in = wb.create_sheet("وراد السور")
    ws_in.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة"])
    ws_in.append(["", "", "", "", "Submit", "Date Reply", "Action", ""])
    ws_in.merge_cells("H1:H2")
    ws_in.append([1, "IN", "S.H.C/2026-05-099", "بشأن تركيب كاميرات المراقبة بالمركز الصحي", "2026-11-10", "", "", "Open"])
    ws_in.append([2, "IN", "S.H.C/2026-05-019", "اعتماد واستبدال مراقب الاعمال المدنية", "2026-11-03", "", "", "Open"])

    ws_sub = wb.create_sheet("وارد مقاولين الباطن")
    ws_sub.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة"])
    ws_sub.append(["", "", "", "", "Submit", "Date Reply", "Action", ""])
    ws_sub.merge_cells("H1:H2")
    ws_sub.append([1, "IN", "م/55/2026", "توريد نظام كاميرات المراقبة CCTV", "2026-11-14", "", "", "Open"])
    ws_sub.append([2, "IN", "م/36/2026", "رفض الدفعة رقم (02)", "2026-11-05", "", "", "Open"])

    log_p = tmp_path / "Letters_CrossLink.xlsx"
    wb.save(log_p)
    wb.close()

    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "CCTV Link Test",
        "code": "CLT",
        "linked_log_path": str(log_p),
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]

    # 1) Query suggestions for Consultant CCTV letter -> both Outgoing and Subcontractor suggested!
    st_q, r_q = _http_json(
        f"{base}/api/projects/{pid}/letters/links?ref_no=S.H.C/2026-05-099&sheet=" + urllib.parse.quote("وراد السور")
    )
    assert st_q == 200 and r_q["ok"] is True
    sug_refs = {s["ref_no"] for s in r_q["suggestions"]}
    assert "ج/2026/90" in sug_refs and "م/55/2026" in sug_refs
    assert "S.H.C/2026-05-019" not in sug_refs

    # 2) Auto-link preview -> finds CCTV pairs, ignores 01 vs 02 and elec vs civil
    st_prev, r_prev = _http_json(f"{base}/api/projects/{pid}/letters/auto-link", method="POST", body={
        "apply": False,
        "min_score": 0.58,
    })
    assert st_prev == 200 and r_prev["applied"] is False
    pair_sets = [{m["a"]["ref_no"], m["b"]["ref_no"]} for m in r_prev["matches"]]
    assert {"S.H.C/2026-05-099", "ج/2026/90"} in pair_sets
    assert {"ج/2026/27", "م/36/2026"} not in pair_sets
    assert {"ج/2026/37", "S.H.C/2026-05-019"} not in pair_sets

    # 3) Apply auto-link -> unites all 3 CCTV letters into ONE thread & writes Col I ('الكتاب المرتبط') in Excel!
    st_ap, r_ap = _http_json(f"{base}/api/projects/{pid}/letters/auto-link", method="POST", body={
        "apply": True,
        "min_score": 0.58,
        "sync_excel": True,
    })
    assert st_ap == 200 and r_ap["ok"] is True
    assert r_ap["threads_updated"] == 1
    assert r_ap["excel"]["updated_rows"] >= 3

    wb_chk = openpyxl.load_workbook(log_p, data_only=True)
    # Col I (col 9) added right after Status / الحالة (col 8), header merged I1:I2!
    assert "الكتاب المرتبط" in str(wb_chk["صادر -السور"].cell(1, 9).value)
    assert "I1:I2" in {str(r) for r in wb_chk["صادر -السور"].merged_cells.ranges}
    out_cell = str(wb_chk["صادر -السور"].cell(3, 9).value or "")
    in_cell = str(wb_chk["وراد السور"].cell(3, 9).value or "")
    sub_cell = str(wb_chk["وارد مقاولين الباطن"].cell(3, 9).value or "")
    assert "S.H.C/2026-05-099" in out_cell and "م/55/2026" in out_cell
    assert "ج/2026/90" in in_cell and "م/55/2026" in in_cell
    assert "ج/2026/90" in sub_cell and "S.H.C/2026-05-099" in sub_cell
    # Unrelated rows must stay empty
    assert wb_chk["صادر -السور"].cell(4, 9).value is None
    assert wb_chk["وارد مقاولين الباطن"].cell(4, 9).value is None
    wb_chk.close()

    # 4) Verify section endpoint exposes structured linked_letters & party_label
    st_sec, r_sec = _http_json(f"{base}/api/projects/{pid}/sections/outgoing")
    assert st_sec == 200
    row90 = next(x for x in r_sec["rows"] if x["ref_no"] == "ج/2026/90")
    assert {m["ref_no"] for m in row90["linked_letters"]} == {"S.H.C/2026-05-099", "م/55/2026"}
    assert row90["party_label"]

    # 5) Unlink subcontractor letter -> Excel updates & cleared for م/55/2026, AND ignored against re-suggesting
    st_un, r_un = _http_json(f"{base}/api/projects/{pid}/letters/unlink", method="POST", body={
        "ref_no": "ج/2026/90",
        "sheet": "صادر -السور",
        "target_ref_no": "م/55/2026",
        "target_sheet": "وارد مقاولين الباطن",
    })
    assert st_un == 200 and r_un["removed"] is True
    wb_after = openpyxl.load_workbook(log_p, data_only=True)
    assert "م/55/2026" not in str(wb_after["صادر -السور"].cell(3, 9).value or "")
    assert wb_after["وارد مقاولين الباطن"].cell(3, 9).value is None
    wb_after.close()

    # Scanner triage integration: scanned follow-up letter suggests matching Consultant letter and links on approve
    proj_obj = hub.get_project(pid)
    inbox_dir = Path(proj_obj["inbox_dir"])
    (inbox_dir / "Letter_CCTV_Reply.txt").write_text(
        "Ref: ج/2026/95\nالتاريخ: 2026-11-17\nالموضوع: بخصوص توريد وتركيب كاميرات المراقبة في المشروع\n",
        encoding="utf-8",
    )
    st_sc, r_sc = _http_json(f"{base}/api/projects/{pid}/scanner/scan", method="POST", body={})
    assert st_sc == 200 and len(r_sc["pending"]) == 1
    tri_it = r_sc["pending"][0]
    tri_id = tri_it["id"]
    # Even if scanner initially routes to outgoing/incoming, force outgoing + check suggest helper directly
    sugs_sc = ll.suggest_letter_links(
        proj_obj,
        section="outgoing",
        subject="بخصوص توريد وتركيب كاميرات المراقبة في المشروع",
        ref_no="ج/2026/95",
    )
    assert any(x["ref_no"] == "S.H.C/2026-05-099" for x in sugs_sc)

    st_ap_sc, r_ap_sc = _http_json(f"{base}/api/projects/{pid}/scanner/approve", method="POST", body={
        "id": tri_id,
        "overrides": {
            "section": "outgoing",
            "ref_no": "ج/2026/95",
            "subject_ar": "بخصوص توريد وتركيب كاميرات المراقبة في المشروع",
            "link_to_refs": [{"ref_no": "S.H.C/2026-05-099", "sheet": "وراد السور"}],
        },
    })
    assert st_ap_sc == 200 and r_ap_sc["ok"] is True
    wb_sc = openpyxl.load_workbook(log_p, data_only=True)
    in_text_final = str(wb_sc["وراد السور"].cell(3, 9).value or "")
    assert "ج/2026/95" in in_text_final and "ج/2026/90" in in_text_final
    wb_sc.close()


def test_recent_first_balanced_log_summary_and_ui_fixes(live_server):
    """
    Verify:
      1) read_project_log_summary(recent_first=True, limit=6) avoids starving later sheets
         when an earlier sheet has many rows, and returns rows sorted newest-first.
      2) UI HTML contains CSS definitions for all component classes used across Organizer,
         Logs, Scanner, and Overview, plus updated layout/ordering helpers.
    """
    import project_log as plog

    base, hub, tmp_path = live_server
    log_p = tmp_path / "LOG_BSU.xlsx"

    wb = openpyxl.Workbook()
    ws_sub = wb.active
    ws_sub.title = "MAR"
    ws_sub.append(["Ref", "Subject", "Date", "Status"])
    for i in range(1, 21):
        ws_sub.append([f"MAR-{i:03d}", f"Material Submittal {i}", f"2026-01-{i:02d}", "Pending"])

    ws_in = wb.create_sheet("Incoming")
    ws_in.append(["Ref", "Subject", "Date", "Status"])
    ws_in.append(["IN-001", "Old Incoming Letter", "2025-05-01", "Received"])
    ws_in.append(["IN-002", "Newest Incoming Letter", "2026-11-18", "Received"])

    ws_out = wb.create_sheet("Outgoing")
    ws_out.append(["Ref", "Subject", "Date", "Status"])
    ws_out.append(["OUT-001", "Newest Outgoing Letter", "2026-11-17", "Sent"])
    wb.save(log_p)
    wb.close()

    st, created = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Balanced Summary UI Project",
        "code": "BSU",
        "linked_log_path": str(log_p),
        "set_active": True,
    })
    assert st == 201 and created["ok"] is True
    pid = created["project"]["id"]
    proj = hub.get_project(pid)

    # Confirm default vs recent_first balanced sampling
    summ = plog.read_project_log_summary(proj, limit=6, recent_first=True)
    rows = summ["rows"]
    assert len(rows) == 6
    secs_seen = {r["section"] for r in rows}
    assert {"submittal", "incoming", "outgoing"}.issubset(secs_seen)
    assert rows[0]["ref_no"] == "IN-002"
    assert rows[1]["ref_no"] == "OUT-001"
    # Even within submittal, the newest rows (MAR-020 / MAR-019) are kept rather than MAR-001
    sub_refs = {r["ref_no"] for r in rows if r["section"] == "submittal"}
    assert "MAR-020" in sub_refs
    assert "MAR-001" not in sub_refs

    # API detail endpoint uses recent_first=True too
    st_det, det = _http_json(f"{base}/api/projects/{pid}")
    assert st_det == 200 and det["recent_rows"][0]["ref_no"] == "IN-002"

    # Verify UI asset serves updated CSS & JS helpers without missing CSS vars
    import urllib.request
    with urllib.request.urlopen(f"{base}/") as resp:
        html = resp.read().decode("utf-8")
    for needle in (
        "--muted:",
        ".sidebar-footer",
        ".grid-2",
        ".badge",
        ".b-amber",
        ".empty",
        ".search-in",
        ".filter-pills",
        ".ov-tools-tray",
        "#toast.warn",
        ".banner.warn",
        "onScannerSectionChange",
        "setOverviewSecFilter",
        "openScannerSplitModal",
        "openReplyFromLetter",
        "openQuickWhatsAppForRow",
        "__unreplied__",
    ):
        assert needle in html, f"Missing expected UI fix marker: {needle}"


def test_multipage_pdf_scanner_priority_ai_split_and_apply_archive(live_server, monkeypatch):
    """
    Verify:
      1) PDFs with >= 3 pages are detected (`page_count`, `is_multipage`) and sorted FIRST in Scanner pending list.
      2) `/scanner/page-thumb` serves valid PNG page thumbnails.
      3) `/scanner/analyze-split` detects document boundaries (heuristic fallback + AI override mode) covering 1..N.
      4) `/scanner/apply-split` splits into child PDFs (both pending review and `auto_approve=True` archive+log).
    """
    import io
    import json
    import re
    import sys
    import types
    import scan_watcher as sw
    from PIL import Image

    try:
        import fitz  # type: ignore
    except ImportError:
        # Lightweight JSON-backed PyMuPDF stub when PyMuPDF wheel isn't preinstalled in test container
        class _FakeRect:
            x0, y0, x1, y1, width, height = 0.0, 0.0, 595.0, 842.0, 595.0, 842.0

        class _FakePixmap:
            def __init__(self):
                im = Image.new("RGB", (40, 56), (255, 255, 255))
                buf = io.BytesIO()
                im.save(buf, format="PNG")
                self._png = buf.getvalue()

            def tobytes(self, fmt="png"):
                return self._png

            def save(self, path):
                Path(path).write_bytes(self._png)

        class _FakePage:
            def __init__(self, doc, idx):
                self._doc = doc
                self._idx = idx
                self.rect = _FakeRect()

            def insert_text(self, pos, text):
                self._doc._pages[self._idx] = (self._doc._pages[self._idx] + "\n" + str(text)).strip()

            def get_text(self, mode="text"):
                return self._doc._pages[self._idx]

            def get_pixmap(self, **kwargs):
                return _FakePixmap()

        class _FakeDoc:
            def __init__(self, path=None):
                self._pages = []
                if path is not None:
                    raw = Path(path).read_bytes()
                    m = re.search(rb"%%FAKEFITZ:(.*?)%%ENDFAKEFITZ", raw, re.S)
                    if m:
                        self._pages = json.loads(m.group(1).decode("utf-8"))
                    else:
                        self._pages = [""]

            @property
            def page_count(self):
                return len(self._pages)

            def __len__(self):
                return len(self._pages)

            def __getitem__(self, idx):
                return _FakePage(self, idx)

            def new_page(self):
                self._pages.append("")
                return _FakePage(self, len(self._pages) - 1)

            def insert_pdf(self, src_doc, from_page=0, to_page=0):
                for idx in range(from_page, to_page + 1):
                    self._pages.append(src_doc._pages[idx])

            def save(self, path):
                payload = json.dumps(self._pages, ensure_ascii=False).encode("utf-8")
                Path(path).write_bytes(b"%PDF-1.4\n%%FAKEFITZ:" + payload + b"%%ENDFAKEFITZ\n")

            def close(self):
                pass

            def __enter__(self):
                return self

            def __exit__(self, *args):
                self.close()

        fake_fitz = types.SimpleNamespace(
            open=lambda p=None: _FakeDoc(p),
            Matrix=lambda a, b: (a, b),
            Rect=lambda *args: _FakeRect(),
        )
        monkeypatch.setitem(sys.modules, "fitz", fake_fitz)
        monkeypatch.setattr(sw, "fitz", fake_fitz)
        fitz = fake_fitz

    base, hub, tmp_path = live_server
    log_p = tmp_path / "LOG_MultiSplit.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Incoming"
    ws.append(["Ref", "Subject", "Date", "Status"])
    wb.save(log_p)
    wb.close()

    st_p, created = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Multipage Split Project",
        "code": "MSP",
        "linked_log_path": str(log_p),
        "set_active": True,
    })
    assert st_p == 201
    pid = created["project"]["id"]
    proj = hub.get_project(pid)
    inbox_dir = Path(proj["inbox_dir"])
    inbox_dir.mkdir(parents=True, exist_ok=True)

    # 1) Create a 1-page PDF (named alphabetically first so priority sort proves multi-page wins!)
    d1 = fitz.open()
    p1 = d1.new_page()
    p1.insert_text((72, 72), "Ref: IN-001\nSingle page notice\nDate: 2026-11-10")
    d1.save(str(inbox_dir / "01_single_notice.pdf"))
    d1.close()

    # 2) Create a 4-page PDF bundle containing 2 distinct documents (p1-2 and p3-4)
    d4 = fitz.open()
    b1 = d4.new_page()
    b1.insert_text((72, 72), "Ref: CON-101\nDate: 2026-11-11\nSubject: Plumbing Material Approval\nPage 1 of 2")
    b2 = d4.new_page()
    b2.insert_text((72, 72), "Continuation of CON-101\nYours faithfully,\nConsultant Team")
    b3 = d4.new_page()
    b3.insert_text((72, 72), "Ref: SUB-202\nDate: 2026-11-12\nSubject: Electrical Cable Tray Delivery\nPage 1 of 2")
    b4 = d4.new_page()
    b4.insert_text((72, 72), "Details for SUB-202\nBest regards,\nSubcontractor")
    d4.save(str(inbox_dir / "02_bundle_4pages.pdf"))
    d4.close()

    # Deterministic mock for scanner triage scan (avoids external AI calls)
    hub.analyzer_override = lambda src, txt, prompt: {
        "doc_type": "letter",
        "letter_direction": "incoming",
        "subject_ar": src.stem,
        "ref_no": "IN-001" if "single" in src.name else "CON-101",
        "doc_date": "2026-11-11",
    }
    st_sc, r_sc = _http_json(f"{base}/api/projects/{pid}/scanner/scan", method="POST", body={"blocking": True})
    assert st_sc == 200 and r_sc["counts"]["pending"] == 2
    assert r_sc["counts"]["multipage"] == 1
    # Multi-page PDF MUST appear FIRST in pending list!
    assert r_sc["pending"][0]["file_name"] == "02_bundle_4pages.pdf"
    assert r_sc["pending"][0]["page_count"] == 4
    assert r_sc["pending"][0]["is_multipage"] is True
    assert r_sc["pending"][1]["file_name"] == "01_single_notice.pdf"
    assert r_sc["pending"][1]["is_multipage"] is False
    multi_id = r_sc["pending"][0]["id"]

    # Check page thumbnail endpoint
    import urllib.request
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/scanner/page-thumb?id={multi_id}&page=1&dpi=90") as trsp:
        png_data = trsp.read()
        assert png_data.startswith(b"\x89PNG")

    # Analyze split with heuristic fallback (override returns empty dict so heuristic engages cleanly)
    hub.analyzer_override = None
    st_an, r_an = _http_json(f"{base}/api/projects/{pid}/scanner/analyze-split", method="POST", body={"id": multi_id})
    assert st_an == 200 and r_an["ok"] is True
    assert r_an["page_count"] == 4
    assert r_an["mode"] == "heuristic"
    assert len(r_an["documents"]) == 2
    assert (r_an["documents"][0]["start_page"], r_an["documents"][0]["end_page"]) == (1, 2)
    assert (r_an["documents"][1]["start_page"], r_an["documents"][1]["end_page"]) == (3, 4)

    # Also verify AI split mode when analyzer_override supplies structured documents
    hub.analyzer_override = lambda src, texts, prompt: {
        "is_single_document": False,
        "documents": [
            {"start_page": 1, "end_page": 2, "title": "اعتماد مواد سباكة", "section": "incoming", "ref_no": "CON-101"},
            {"start_page": 3, "end_page": 4, "title": "توريد كيبلات كهرباء", "section": "incoming", "ref_no": "SUB-202"},
        ],
    }
    st_ai, r_ai = _http_json(f"{base}/api/projects/{pid}/scanner/analyze-split", method="POST", body={"id": multi_id})
    assert st_ai == 200 and r_ai["mode"] == "ai"
    assert len(r_ai["documents"]) == 2

    # Apply split into Scanner Inbox (auto_approve=False)
    st_sp, r_sp = _http_json(f"{base}/api/projects/{pid}/scanner/apply-split", method="POST", body={
        "id": multi_id,
        "ranges": r_ai["documents"],
        "auto_approve": False,
    })
    assert st_sp == 200 and r_sp["split_result"]["split_count"] == 2
    pend_names = [it["file_name"] for it in r_sp["pending"]]
    assert any("part01" in nm for nm in pend_names)
    assert any("part02" in nm for nm in pend_names)
    assert "02_bundle_4pages.pdf" not in pend_names

    # Create a 3-page PDF and split + archive immediately (`auto_approve=True`)
    d3 = fitz.open()
    for p_i in range(1, 4):
        pg = d3.new_page()
        pg.insert_text((72, 72), f"Page {p_i} of bundle")
    d3.save(str(inbox_dir / "03_bundle_3pages.pdf"))
    d3.close()
    hub.analyzer_override = lambda src, txt, prompt: {
        "doc_type": "letter",
        "letter_direction": "incoming",
        "subject_ar": "Bundle 3",
        "ref_no": "IN-300",
    }
    st_sc2, r_sc2 = _http_json(f"{base}/api/projects/{pid}/scanner/scan", method="POST", body={"blocking": True})
    assert st_sc2 == 200
    tri3 = next(x for x in r_sc2["pending"] if x["file_name"] == "03_bundle_3pages.pdf")
    assert tri3["is_multipage"] is True and tri3["page_count"] == 3

    st_auto, r_auto = _http_json(f"{base}/api/projects/{pid}/scanner/apply-split", method="POST", body={
        "id": tri3["id"],
        "ranges": [
            {"start_page": 1, "end_page": 1, "title": "كتاب مرفق 1", "section": "incoming", "ref_no": "IN-301", "doc_date": "2026-11-15"},
            {"start_page": 2, "end_page": 3, "title": "كتاب مرفق 2", "section": "incoming", "ref_no": "IN-302", "doc_date": "2026-11-15"},
        ],
        "auto_approve": True,
    })
    assert st_auto == 200 and r_auto["split_result"]["auto_approved"] is True
    assert r_auto["split_result"]["split_count"] == 2
    wb_after = openpyxl.load_workbook(log_p, data_only=True)
    refs_logged = {str(wb_after["Incoming"].cell(r, 1).value or "") for r in range(2, wb_after["Incoming"].max_row + 1)}
    wb_after.close()
    assert {"IN-301", "IN-302"}.issubset(refs_logged)


def test_correspondence_reply_sync_unreplied_kpi_inline_reply_and_whatsapp(live_server):
    """
    Verify:
      1) `unreplied_incoming` KPI accurately counts open incoming letters (respecting `'تم الرد'`).
      2) Linking an outgoing letter (`out_date >= in_date`) auto-fills incoming `Date Reply` + `'تم الرد'` in Excel.
      3) Issuing an outgoing reply from a template with `link_to_refs` auto-syncs reply date & clears unreplied KPI.
      4) Contacts + quick WhatsApp URL generation works with message & attachment.
    """
    import datetime as dt
    import docx

    base, hub, tmp_path = live_server
    log_p = tmp_path / "LOG_ReplySync.xlsx"

    wb = openpyxl.Workbook()
    ws_in = wb.active
    ws_in.title = "وراد السور"
    ws_in.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة"])
    ws_in.append(["", "", "", "", "Submit", "Date Reply", "Action", ""])
    ws_in.merge_cells("H1:H2")
    ws_in.append([1, "IN", "IN-100", "بخصوص اعتماد العينات الصحية", "2026-11-10", "", "", ""])
    ws_in.append([2, "IN", "IN-101", "بخصوص لوحات الكهرباء", "2026-11-05", "تم الرد", "", ""])
    ws_in.append([3, "IN", "IN-102", "بخصوص فحص التربة", "2026-11-15", "", "", "Open"])

    ws_out = wb.create_sheet("صادر -السور")
    ws_out.append(["#", "Type", "Reference", "Description", "REV. 0", "", "", "Status / الحالة"])
    ws_out.append(["", "", "", "", "Submit", "Date Reply", "Action", ""])
    ws_out.merge_cells("H1:H2")
    ws_out.append([1, "OUT", "OUT-200", "رد بخصوص العينات الصحية", "2026-11-12", "", "", "Sent"])
    wb.save(log_p)
    wb.close()

    st_p, created = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Reply Sync Project",
        "code": "RSP",
        "linked_log_path": str(log_p),
        "set_active": True,
    })
    assert st_p == 201
    pid = created["project"]["id"]

    # 1) Initial summary: IN-100 and IN-102 are unreplied (2), IN-101 has 'تم الرد' -> replied!
    st_det, det = _http_json(f"{base}/api/projects/{pid}")
    assert st_det == 200
    assert det["stats"]["unreplied_incoming"] == 2
    row101 = next(x for x in det["recent_rows"] if x["ref_no"] == "IN-101")
    row100 = next(x for x in det["recent_rows"] if x["ref_no"] == "IN-100")
    assert row101["replied"] is True and row101["unreplied"] is False
    assert row100["unreplied"] is True and isinstance(row100["days_waiting"], int)

    # 2) Link IN-100 (2026-11-10) <-> OUT-200 (2026-11-12) -> auto-syncs Date Reply & Status in Excel!
    st_lk, r_lk = _http_json(f"{base}/api/projects/{pid}/letters/link", method="POST", body={
        "items": [
            {"ref_no": "IN-100", "sheet": "وراد السور", "section": "incoming", "doc_date": "2026-11-10"},
            {"ref_no": "OUT-200", "sheet": "صادر -السور", "section": "outgoing", "doc_date": "2026-11-12"},
        ],
        "sync_excel": True,
    })
    assert st_lk == 200 and r_lk["excel"]["reply_synced_rows"] >= 1

    wb_c = openpyxl.load_workbook(log_p, data_only=True)
    rep_cell = wb_c["وراد السور"].cell(3, 6).value
    st_cell = wb_c["وراد السور"].cell(3, 8).value
    assert "2026-11-12" in str(rep_cell)
    assert "تم الرد" in str(st_cell)
    wb_c.close()

    # Unreplied KPI drops from 2 to 1!
    _, det2 = _http_json(f"{base}/api/projects/{pid}")
    assert det2["stats"]["unreplied_incoming"] == 1

    # 3) Upload an outgoing .docx template, map it, and issue a reply linked to IN-102!
    doc_p = tmp_path / "Outgoing_Reply_Tpl.docx"
    d = docx.Document()
    d.add_paragraph("Ref: {{REF}}")
    d.add_paragraph("Date: {{DATE}}")
    d.add_paragraph("Subject: {{SUBJECT}}")
    d.save(str(doc_p))

    st_up, r_up = _http_json(f"{base}/api/projects/{pid}/templates/upload", method="POST", body={
        "name": "Outgoing_Reply_Tpl.docx",
        "section": "outgoing",
        "display_name": "نموذج كتاب صادر للرد",
        "data_b64": base64.b64encode(doc_p.read_bytes()).decode("ascii"),
    })
    assert st_up == 201

    st_pr, r_pr = _http_json(f"{base}/api/projects/{pid}/templates/Outgoing_Reply_Tpl.docx/profile", method="POST", body={
        "section": "outgoing",
        "display_name": "نموذج كتاب صادر للرد",
        "target_sheet": "صادر -السور",
        "fields": [
            {"key": "ref_no", "label": "المرجع", "role": "auto_counter", "target": {"kind": "docx_span", "para_idx": 0, "match_text": "{{REF}}"}, "counter_prefix": "OUT-", "counter_digits": 3, "counter_next": 201},
            {"key": "doc_date", "label": "التاريخ", "role": "auto_date", "source": "today_iso", "target": {"kind": "docx_span", "para_idx": 1, "match_text": "{{DATE}}"}},
            {"key": "subject", "label": "الموضوع", "role": "manual", "input_type": "text", "target": {"kind": "docx_span", "para_idx": 2, "match_text": "{{SUBJECT}}"}},
        ],
    })
    assert st_pr == 200

    st_is, r_is = _http_json(f"{base}/api/projects/{pid}/templates/Outgoing_Reply_Tpl.docx/issue", method="POST", body={
        "ref_no": "OUT-201",
        "subject": "رداً على: بخصوص فحص التربة",
        "party_name": "المكتب الاستشاري",
        "values": {"subject": "رداً على: بخصوص فحص التربة"},
        "link_to_refs": [{"ref_no": "IN-102", "sheet": "وراد السور", "section": "incoming", "subject": "بخصوص فحص التربة"}],
    })
    assert st_is == 200 and r_is["ok"] is True
    assert r_is["link_result"]["excel"]["reply_synced_rows"] >= 1

    # IN-102 now has today's date in Date Reply & 'تم الرد' status, and unreplied_incoming == 0!
    wb_c2 = openpyxl.load_workbook(log_p, data_only=True)
    assert dt.date.today().isoformat() in str(wb_c2["وراد السور"].cell(5, 6).value)
    assert "تم الرد" in str(wb_c2["وراد السور"].cell(5, 8).value)
    wb_c2.close()

    _, det3 = _http_json(f"{base}/api/projects/{pid}")
    assert det3["stats"]["unreplied_incoming"] == 0

    # 4) Verify saved contact & quick WhatsApp message URL generation
    st_ct, r_ct = _http_json(f"{base}/api/projects/{pid}/contacts", method="POST", body={
        "name": "م. الاستشاري",
        "phone": "+96550001234",
        "role": "استشاري",
    })
    assert st_ct == 200 and len(r_ct["contacts"]) == 1
    st_wa, r_wa = _http_json(f"{base}/api/projects/{pid}/whatsapp/send", method="POST", body={
        "mode": "web",
        "phone": "+96550001234",
        "message": "المشروع: Reply Sync Project | IN-102 — بخصوص فحص التربة",
    })
    assert st_wa == 200 and "web.whatsapp.com" in r_wa["url"] and "96550001234" in r_wa["url"]


def test_organizer_near_duplicate_section_by_name_and_content(live_server):
    base, hub, tmp_path = live_server

    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Near Dup Test Project",
        "code": "NDP-1",
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]

    src = tmp_path / "Messy_Near_Dups"
    src.mkdir(parents=True, exist_ok=True)

    # 1) Same name across different formats / copy suffixes -> winner is PDF, others go to near_duplicate
    (src / "اشعار بموعد الصب.pdf").write_bytes(b"%PDF-1.4 main pdf bytes")
    (src / "اشعار بموعد الصب.jpg").write_bytes(b"JPG scan image bytes")
    (src / "اشعار بموعد الصب (1).pdf").write_bytes(b"%PDF-1.4 slightly re-saved copy bytes")

    # 2) Different filenames but identical normalized Arabic document content -> near_duplicate (content)
    txt_body = "الموضوع: اعتماد مخططات التكييف والكهرباء للدور الأرضي بمستوصف صباح السالم الصحي وفقاً لملاحظات الاستشاري المعتمدة."
    (src / "كتاب_اعتماد_التكييف.txt").write_text(txt_body, encoding="utf-8")
    (src / "صورة_ممسوحة_غير_مرتبة.txt").write_text(f"  {txt_body}   \n", encoding="utf-8")

    # 3) Equal-revision same reference across two different files -> near_duplicate (ref)
    (src / "MAR-EL-005_R01_Original.pdf").write_bytes(b"%PDF-1.4 mar el 005 r01 pdf version")
    (src / "MAR-EL-005_R01_Photo.jpg").write_bytes(b"JPG scan of mar el 005 r01")

    # 4) Distinct numbered files must NOT be flagged as near_duplicate
    (src / "MAR-EL-001_Cable.txt").write_text("توصيف كابلات الضغط المنخفض رقم واحد للمشروع بشكل تفصيلي وكامل.", encoding="utf-8")
    (src / "MAR-EL-002_Cable.txt").write_text("توصيف لوحات التوزيع الرئيسية رقم اثنين للمشروع بشكل مختلف تماماً.", encoding="utf-8")

    st_s, r_s = _http_json(f"{base}/api/projects/{pid}/organizer/scan", method="POST", body={
        "folder": str(src),
        "extract_archives": False,
        "read_text": True,
        "blocking": True,
    })
    assert st_s == 200 and r_s["ok"] is True
    cnt = r_s["summary"]["counts"]
    assert cnt["near_duplicate"] == 4
    assert cnt["keep"] == 5  # pdf winner + content winner + ref winner + 2 distinct MAR-EL files

    # Dedicated near_duplicate filter lists all 4 pending matches with reason & winner name
    st_nd, r_nd = _http_json(f"{base}/api/projects/{pid}/organizer?filter=near_duplicate")
    assert st_nd == 200 and len(r_nd["items"]) == 4
    reasons = {it["file_name"]: it.get("near_dup_reason") for it in r_nd["items"]}
    assert reasons["اشعار بموعد الصب.jpg"] == "name"
    assert reasons["اشعار بموعد الصب (1).pdf"] == "name"
    assert reasons["MAR-EL-005_R01_Photo.jpg"] == "ref"
    assert "content" in set(reasons.values())
    assert all(it.get("decision") == "pending" for it in r_nd["items"])

    # Per-row decisions: exclude one, keep one as regular file, then isolate remaining two
    id_by_name = {it["file_name"]: it["id"] for it in r_nd["items"]}
    st_u1, r_u1 = _http_json(f"{base}/api/projects/{pid}/organizer/update", method="POST", body={
        "updates": [
            {"id": id_by_name["اشعار بموعد الصب (1).pdf"], "decision": "excluded"},
            {"id": id_by_name["MAR-EL-005_R01_Photo.jpg"], "decision": "kept"},
        ],
    })
    assert st_u1 == 200 and r_u1["counts"]["near_duplicate"] == 2

    # Isolate all remaining near duplicates (bulk applies across near_duplicate items -> clears pending tab)
    st_iso, r_iso = _http_json(f"{base}/api/projects/{pid}/organizer/update", method="POST", body={
        "bulk_action": "isolate_near_duplicates",
    })
    assert st_iso == 200 and r_iso["counts"]["near_duplicate"] == 0 and r_iso["counts"]["near_duplicate_resolved"] == 3
    st_nd_empty, r_nd_empty = _http_json(f"{base}/api/projects/{pid}/organizer?filter=near_duplicate")
    assert st_nd_empty == 200 and r_nd_empty["items"] == []
    st_nd_res, r_nd_res = _http_json(f"{base}/api/projects/{pid}/organizer?filter=near_duplicate&show_resolved=1")
    assert st_nd_res == 200 and len(r_nd_res["items"]) == 3 and all(it["decision"] == "isolated" for it in r_nd_res["items"])

    st_ex, r_ex = _http_json(f"{base}/api/projects/{pid}/organizer/execute", method="POST", body={
        "mode": "copy",
        "update_excel_log": False,
        "blocking": True,
    })
    # 5 original keeps + 1 kept near_dup + 3 isolated near_dups = 9 files copied
    assert st_ex == 200 and r_ex["execution"]["copied"] == 9
    arc_root = Path(hub.get_project(pid)["archive_dir"])
    assert (arc_root / "_Duplicates_ملفات_مكررة" / "By_Name_Data").is_dir()
    assert len(list((arc_root / "_Duplicates_ملفات_مكررة" / "By_Name_Data").iterdir())) == 3

    # Verify Excel report includes near_duplicate summary row & Sheet 3 rows
    rep_file = Path(hub.get_project(pid)["root"]) / r_ex["execution"]["report_rel"]
    wb_rep = openpyxl.load_workbook(rep_file, data_only=True)
    ws1_vals = [str(r[0]) for r in wb_rep["1_الملخص_Summary"].iter_rows(values_only=True) if r and r[0]]
    assert any("Near-Duplicates" in v for v in ws1_vals)
    ws3_rows = list(wb_rep["3_التكرار_والمراجعات"].iter_rows(values_only=True))
    assert any(r and r[0] == "near_duplicate" for r in ws3_rows[1:])
    wb_rep.close()


def test_organizer_smart_renaming_inline_persistence_and_ai_rename(live_server, monkeypatch):
    import scan_watcher as sw

    base, hub, tmp_path = live_server

    st_p, r_p = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "Organizer Smart Rename Project",
        "code": "OSR-1",
        "set_active": True,
    })
    assert st_p == 201
    pid = r_p["project"]["id"]

    src = tmp_path / "Messy_Rename_Dir"
    src.mkdir(parents=True, exist_ok=True)

    # 1) Generic filename whose true date & subject live inside document content
    (src / "scan_001.txt").write_text(
        "التاريخ: 2025-04-12\nالموضوع: اعتماد مخطط التكييف للدور الأول\nمرسل للاعتماد من المكتب الاستشاري.",
        encoding="utf-8",
    )
    # 2) Uncoded Arabic file without ref_no -> should receive standardized YYYY-MM-DD_TAG_Subject name
    (src / "اذن عمل مؤقت مدير المشروع.pdf").write_bytes(b"%PDF-1.4 permit body")
    # 3) Scanned image with generic name -> ready for AI rename
    (src / "IMG_20261118_001.jpg").write_bytes(b"\xff\xd8\xff\xe0FAKEJPG")

    st_s, r_s = _http_json(f"{base}/api/projects/{pid}/organizer/scan", method="POST", body={
        "folder": str(src),
        "extract_archives": False,
        "read_text": True,
        "blocking": True,
    })
    assert st_s == 200 and r_s["ok"] is True

    st_all, r_all = _http_json(f"{base}/api/projects/{pid}/organizer?filter=all")
    assert st_all == 200
    by_orig = {it["file_name"]: it for it in r_all["items"]}

    scan_it = by_orig["scan_001.txt"]
    assert scan_it["new_name"] != "scan_001.txt"
    assert "2025-04-12" in scan_it["new_name"]
    assert "اعتماد_مخطط_التكييف" in scan_it["new_name"]

    permit_it = by_orig["اذن عمل مؤقت مدير المشروع.pdf"]
    assert permit_it["new_name"] != "اذن عمل مؤقت مدير المشروع.pdf"
    assert permit_it["new_name"].endswith(".pdf")
    assert "_GEN_" in permit_it["new_name"] or "_PRM_" in permit_it["new_name"]

    # Manual inline rename with internal dots (must NOT be truncated or overwritten on subsequent section/include changes)
    st_ren, _ = _http_json(f"{base}/api/projects/{pid}/organizer/update", method="POST", body={
        "updates": [{"id": permit_it["id"], "new_name": "2025.05.14_خطاب اذن عمل مهم.pdf"}],
    })
    assert st_ren == 200

    st_sec, _ = _http_json(f"{base}/api/projects/{pid}/organizer/update", method="POST", body={
        "updates": [{"id": permit_it["id"], "section": "outgoing", "include": True}],
    })
    assert st_sec == 200

    st_chk, r_chk = _http_json(f"{base}/api/projects/{pid}/organizer?filter=all")
    permit_after = {it["file_name"]: it for it in r_chk["items"]}["اذن عمل مؤقت مدير المشروع.pdf"]
    assert permit_after["new_name"] == "2025.05.14_خطاب اذن عمل مهم.pdf"
    assert permit_after["custom_name"] is True
    assert "02_Outgoing" in permit_after["dest_folder"]

    # AI Smart Rename on scanned image (mock Gemini)
    img_it = by_orig["IMG_20261118_001.jpg"]

    def _fake_gemini(prompt, job_dir, model="", images=None, api_key=""):
        assert images and len(images) == 1
        return {
            "section_id": "outgoing",
            "ref_no": "OUT-2026-088",
            "revision": "R00",
            "doc_date": "2026-11-18",
            "authority": "MOH",
            "subject_ar": "اشعار بموعد صب السقف",
            "suggested_filename": "OUT-2026-088_MOH_اشعار_بموعد_صب_السقف",
        }

    monkeypatch.setattr(sw, "analyze_with_gemini", _fake_gemini)
    st_ai, r_ai = _http_json(f"{base}/api/projects/{pid}/organizer/ai-rename", method="POST", body={
        "item_ids": [img_it["id"]],
        "blocking": True,
    })
    assert st_ai == 200 and r_ai["ok"] is True and r_ai["updated"] == 1

    st_chk2, r_chk2 = _http_json(f"{base}/api/projects/{pid}/organizer?filter=all")
    img_after = {it["file_name"]: it for it in r_chk2["items"]}["IMG_20261118_001.jpg"]
    assert img_after["new_name"] == "OUT-2026-088_MOH_اشعار_بموعد_صب_السقف.jpg"
    assert img_after["section"] == "outgoing"
    assert img_after["ref_no"] == "OUT-2026-088"
    assert img_after["ai_renamed"] is True and img_after["custom_name"] is True

    # Execute and verify archive files carry the new custom / AI names
    st_ex, r_ex = _http_json(f"{base}/api/projects/{pid}/organizer/execute", method="POST", body={
        "mode": "copy",
        "update_excel_log": False,
        "blocking": True,
    })
    assert st_ex == 200 and r_ex["execution"]["copied"] == 3
    arc_root = Path(hub.get_project(pid)["archive_dir"])
    assert any(p.name == "2025.05.14_خطاب اذن عمل مهم.pdf" for p in arc_root.rglob("*.pdf"))
    assert any(p.name == "OUT-2026-088_MOH_اشعار_بموعد_صب_السقف.jpg" for p in arc_root.rglob("*.jpg"))





