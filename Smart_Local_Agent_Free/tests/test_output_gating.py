#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regression tests guaranteeing output files/downloads NEVER appear unless a command genuinely completed."""

from __future__ import annotations

import base64
import io
import json
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

import openpyxl
from PIL import Image, ImageDraw
import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import content_duplicates as dups  # noqa: E402
import project_dashboard as pdash  # noqa: E402
import project_organizer as porg  # noqa: E402
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
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as he:
        return he.code, json.loads(he.read().decode("utf-8"))


def _create_project(base: str, root: Path, code: str = "GATE1") -> str:
    st, res = _http_json(
        f"{base}/api/projects",
        method="POST",
        body={
            "name": f"Project {code}",
            "code": code,
            "root": str(root),
            "inbox_dir": str(root / "1_Inbox"),
            "archive_dir": str(root / "2_Archive"),
        },
    )
    assert st in (200, 201) and res["ok"] is True
    return res["project"]["id"]


def _make_img(path: Path, color: tuple[int, int, int] = (40, 120, 210)) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.new("RGB", (64, 64), color)
    d = ImageDraw.Draw(im)
    d.rectangle([10, 10, 54, 54], outline=(255, 255, 255), width=3)
    im.save(path)


def test_duplicates_strict_output_gating(live_server):
    base, hub, tmp_path = live_server
    proj_root = tmp_path / "Proj_Dups_Gate"
    pid = _create_project(base, proj_root, "DG01")
    proj = hub.get_project(pid)

    src = tmp_path / "DupSrc"
    _make_img(src / "orig.png", (10, 80, 190))
    (src / "copy.png").write_bytes((src / "orig.png").read_bytes())
    _make_img(src / "unique.png", (210, 40, 40))

    # 1) Scan only (NO execution yet): ZERO output files, ZERO last_execution, download-zip blocked!
    dups.scan_content_duplicates(proj, source_kind="folder", source_path=str(src))
    st_pg, pg_scan = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert st_pg == 200
    assert pg_scan["has_plan"] is True
    assert pg_scan["last_execution"] is None
    reports_dir = proj_root / "_System_Agent" / "reports"
    assert not reports_dir.exists() or not list(reports_dir.glob("*.xlsx"))
    st_zip_pre, r_zip_pre = _http_json(f"{base}/api/projects/{pid}/duplicates/download-zip")
    assert st_zip_pre == 400 and r_zip_pre["ok"] is False

    # 2) Clear selection (0 items excluded) -> execute must refuse without generating any report!
    dups.update_duplicates_plan(proj, bulk_action="clear_selection")
    st_zero, r_zero = _http_json(
        f"{base}/api/projects/{pid}/duplicates/execute",
        method="POST",
        body={"mode": "quarantine", "blocking": True},
    )
    assert st_zero == 400 and r_zero["ok"] is False
    assert not reports_dir.exists() or not list(reports_dir.glob("*.xlsx"))
    _, pg_zero = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_zero["last_execution"] is None

    # 3) Re-select duplicates, then abort with stop_event -> must NOT mark executed or generate report!
    dups.update_duplicates_plan(proj, bulk_action="select_all_dups")
    stop_ev = threading.Event()
    stop_ev.set()
    with pytest.raises(dups.DuplicatesStoppedError):
        dups.execute_duplicates_plan(proj, mode="quarantine", stop_event=stop_ev)
    assert not reports_dir.exists() or not list(reports_dir.glob("*.xlsx"))
    _, pg_stop = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_stop["has_plan"] is True and pg_stop["last_execution"] is None

    # 4) Execute completely -> report exists, last_execution valid, ZIP downloads 200
    st_ok, r_ok = _http_json(
        f"{base}/api/projects/{pid}/duplicates/execute",
        method="POST",
        body={"mode": "quarantine", "blocking": True},
    )
    assert st_ok == 200 and r_ok["ok"] is True and r_ok["report_exists"] is True
    rep_p = proj_root / r_ok["report_rel"]
    assert rep_p.exists()
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/duplicates/download-zip", timeout=10) as rz:
        assert rz.status == 200 and len(rz.read()) > 0

    # 5) Manual report deletion hides report_rel & flips report_exists=False
    rep_p.unlink()
    _, pg_missing_rep = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_missing_rep["last_execution"]["report_exists"] is False
    assert pg_missing_rep["last_execution"]["report_rel"] == ""

    # 6) Restore via Undo -> last_execution must become None AND download-zip must return 400!
    st_un, r_un = _http_json(f"{base}/api/projects/{pid}/duplicates/undo", method="POST", body={})
    assert st_un == 200 and r_un["restored"] == 1
    _, pg_after_undo = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_after_undo["last_execution"] is None
    st_zip_undo, _ = _http_json(f"{base}/api/projects/{pid}/duplicates/download-zip")
    assert st_zip_undo == 400

    # 7) Fresh scan after prior execution never inherits old last_execution
    dups.scan_content_duplicates(proj, source_kind="folder", source_path=str(src))
    dups.execute_duplicates_plan(proj, mode="export_unique_zip")
    _, pg_ex2 = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_ex2["last_execution"] is not None
    dups.scan_content_duplicates(proj, source_kind="folder", source_path=str(src))
    _, pg_rescan = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_rescan["last_execution"] is None
    st_zip_rescan, _ = _http_json(f"{base}/api/projects/{pid}/duplicates/download-zip")
    assert st_zip_rescan == 400


def test_whatsapp_strict_output_gating(live_server):
    base, hub, tmp_path = live_server
    proj_root = tmp_path / "Proj_WA_Gate"
    pid = _create_project(base, proj_root, "WG01")
    proj = hub.get_project(pid)

    wa_dir = tmp_path / "WA_Gate_Src"
    _make_img(wa_dir / "IMG-20261118-WA0001.jpg", (50, 150, 50))
    (wa_dir / "_chat.txt").write_text(
        "[18/11/2026, 09:10:00] م. أحمد: <attached: IMG-20261118-WA0001.jpg>\n"
        "استلام صب الأعمدة\n",
        encoding="utf-8",
    )

    # 1) Scan only -> no report, no last_execution, download-zip returns 400
    waa.scan_and_match_whatsapp_source(proj, wa_dir)
    _, pg0 = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert pg0["has_plan"] is True and pg0["last_execution"] is None
    rep_dir = proj_root / "_Reports"
    assert not rep_dir.exists() or not list(rep_dir.glob("*.xlsx"))
    st_z0, _ = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive/download-zip")
    assert st_z0 == 400

    # 2) Exclude all items (0 included) -> execute refuses without report
    item_id = pg0["items"][0]["id"]
    waa.update_whatsapp_plan_items(proj, updates=[{"id": item_id, "include": False}])
    st_zero, _ = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/execute",
        method="POST",
        body={"mode": "rename_in_place", "blocking": True},
    )
    assert st_zero == 400
    assert not rep_dir.exists() or not list(rep_dir.glob("*.xlsx"))

    # 3) Re-include, stop mid-execution -> raises WhatsappStoppedError, NO report & NO last_execution
    waa.update_whatsapp_plan_items(proj, updates=[{"id": item_id, "include": True}])
    stop_ev = threading.Event()
    stop_ev.set()
    with pytest.raises(waa.WhatsappStoppedError):
        waa.execute_whatsapp_archive_plan(proj, mode="rename_in_place", stop_event=stop_ev)
    assert not rep_dir.exists() or not list(rep_dir.glob("*.xlsx"))
    _, pg_stop = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert pg_stop["has_plan"] is True and pg_stop["last_execution"] is None

    # 4) Execute genuinely -> report created, ZIP 200
    st_ex, r_ex = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/execute",
        method="POST",
        body={"mode": "rename_in_place", "blocking": True},
    )
    assert st_ex == 200 and r_ex["ok"] is True and r_ex["renamed"] == 1
    rep_proj = proj_root / r_ex["report_rel"]
    rep_local = Path(r_ex["report_local"])
    assert rep_proj.exists() and rep_local.exists()
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/whatsapp-archive/download-zip", timeout=10) as rz:
        assert rz.status == 200 and len(rz.read()) > 0

    # 5) Undo after execution (when plan["items"] was already cleared!):
    #    MUST clear last_execution, remove output reports, and block download-zip!
    st_un, r_un = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive/undo", method="POST", body={})
    assert st_un == 200 and r_un["restored"] == 1
    assert not rep_proj.exists()
    assert not rep_local.exists()
    _, pg_un = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert pg_un["last_execution"] is None
    st_z_un, _ = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive/download-zip")
    assert st_z_un == 400


def test_organizer_pdf_and_issue_strict_output_gating(live_server, monkeypatch):
    import template_profile as tprof

    base, hub, tmp_path = live_server
    proj_root = tmp_path / "Proj_Org_Gate"
    pid = _create_project(base, proj_root, "OG01")
    proj = hub.get_project(pid)

    src_dir = tmp_path / "OrgSrc"
    src_dir.mkdir(parents=True)
    (src_dir / "MAR-EL-010_R00_Switchgear.txt").write_bytes(b"Switchgear submittal")

    # 1) Organizer scan -> no execution report
    porg.build_organization_plan(proj, src_dir, read_text=False)
    _, og0 = _http_json(f"{base}/api/projects/{pid}/organizer")
    assert og0["has_plan"] is True and og0["last_execution"] is None
    rep_dir = proj_root / "_Reports"
    assert not rep_dir.exists() or not list(rep_dir.glob("*.xlsx"))

    # 2) Organizer stopped via stop_event -> NO report, NO last_execution, plan stays intact
    stop_ev = threading.Event()
    stop_ev.set()
    r_stop = porg.execute_organizer_plan(proj, mode="copy", update_excel_log=False, stop_event=stop_ev)
    assert r_stop["stopped"] is True and r_stop["cleared"] is False
    assert not rep_dir.exists() or not list(rep_dir.glob("*.xlsx"))
    _, og_stop = _http_json(f"{base}/api/projects/{pid}/organizer")
    assert og_stop["has_plan"] is True and og_stop["last_execution"] is None

    # 3) Complete organizer execution, then delete report -> report_exists flips to False
    r_done = porg.execute_organizer_plan(proj, mode="copy", update_excel_log=False)
    assert r_done["cleared"] is True and r_done["copied"] == 1 and r_done["report_exists"] is True
    rep_org = proj_root / r_done["report_rel"]
    assert rep_org.exists()
    rep_org.unlink()
    _, og_del = _http_json(f"{base}/api/projects/{pid}/organizer")
    assert og_del["last_execution"]["report_exists"] is False
    assert og_del["last_execution"]["report_rel"] == ""

    # 4) Failed template issue must NOT leave partial output file on disk
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MAR"
    ws["A1"] = "Ref"
    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    tname = "FailTpl.xlsx"
    st_up, _ = _http_json(
        f"{base}/api/projects/{pid}/templates/upload",
        method="POST",
        body={"name": tname, "data_b64": base64.b64encode(buf.getvalue()).decode("ascii")},
    )
    assert st_up in (200, 201)
    st_pr, _ = _http_json(
        f"{base}/api/projects/{pid}/templates/{tname}/profile",
        method="POST",
        body={
            "section": "submittal_mar",
            "section_confirmed": True,
            "ref_prefix": "MAR",
            "log_on_issue": False,
            "fields": [
                {"key": "ref_no", "label": "المرجع", "role": "auto_system", "source": "system.next_ref", "target": {"sheet": "MAR", "coord": "B1"}},
            ],
        },
    )
    assert st_pr == 200

    def _broken_fill(tpl_path, profile, payload, proj_d, out_path, **kwargs):
        Path(out_path).write_bytes(b"PARTIAL_CORRUPT")
        raise RuntimeError("Simulated write crash")

    monkeypatch.setattr(tprof, "fill_by_profile", _broken_fill)
    st_iss, r_iss = _http_json(
        f"{base}/api/projects/{pid}/templates/{tname}/issue",
        method="POST",
        body={"values": {}},
    )
    assert st_iss == 400 and r_iss["ok"] is False
    out_folder = proj_root / "2_Archive"
    assert not list(out_folder.rglob("MAR-*.xlsx"))

    # 5) Zero included items in organizer must reject without generating report
    src2 = tmp_path / "OrgSrcEmpty"
    src2.mkdir(parents=True)
    (src2 / "doc.txt").write_bytes(b"doc")
    porg.build_organization_plan(proj, src2, read_text=False)
    pp = porg.organizer_plan_path(proj_root)
    pl_d = json.loads(pp.read_text(encoding="utf-8"))
    for it in pl_d["items"]:
        it["include"] = False
    pp.write_text(json.dumps(pl_d, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(ValueError):
        porg.execute_organizer_plan(proj, mode="copy", update_excel_log=False)
    assert not rep_dir.exists() or not list(rep_dir.glob("*.xlsx"))

    # 6) Verify UI HTML includes strict output-gating guards
    with urllib.request.urlopen(f"{base}/", timeout=10) as r_ui:
        ui_html = r_ui.read().decode("utf-8")
    for marker in ("waExecBlocked", "dupExecBlocked", "jobStoppedOrErr", "report_exists !== false", "zip_ready !== false", "pdfSplitOut", "pdfMergeOut"):
        assert marker in ui_html

    # 7) Optional PyMuPDF atomic failure cleanup + download_url check
    fitz = getattr(porg, "fitz", None)
    if fitz is not None:
        p1 = tmp_path / "p1.pdf"
        p2 = tmp_path / "p2.pdf"
        for p, txt in ((p1, "One"), (p2, "Two")):
            d = fitz.open()
            pg = d.new_page()
            pg.insert_text((72, 72), txt)
            d.save(str(p))
            d.close()

        corrupt = tmp_path / "bad.pdf"
        corrupt.write_bytes(b"NOT A VALID PDF")
        bad_out = tmp_path / "MergedFail.pdf"
        with pytest.raises(Exception):
            porg.merge_pdf_documents([p1, corrupt], bad_out)
        assert not bad_out.exists()
        assert not list(tmp_path.glob(".__tmp_*"))

        st_mg, r_mg = _http_json(
            f"{base}/api/projects/{pid}/organizer/pdf-merge",
            method="POST",
            body={"src_pdfs": [str(p1), str(p2)]},
        )
        assert st_mg == 200 and r_mg["ok"] is True and r_mg["download_url"]
        with urllib.request.urlopen(f"{base}{r_mg['download_url']}", timeout=10) as r_dl:
            assert r_dl.status == 200 and r_dl.read().startswith(b"%PDF")


def test_all_ui_action_buttons_have_busy_feedback_and_duplicate_click_protection():
    import re
    base_dir = Path(__file__).resolve().parent.parent
    ui_path = base_dir / "doc_tools" / "dashboard_ui.html"
    strings_path = base_dir / "doc_tools" / "strings_ar.json"
    html = ui_path.read_text(encoding="utf-8")
    strings_d = json.loads(strings_path.read_text(encoding="utf-8"))
    busy_entry = strings_d["strings"]["jobs.busy_wait"]
    assert busy_entry.get("msa") and busy_entry.get("eg")

    # Guard against double-click inside withBusyButton and confirm modals
    assert "b.disabled || b.classList.contains('is-busy')" in html
    assert "window.__confirmCb = async (btnEl) => {\n    await withBusyButton(btnEl," in html

    # Every async action handler bound to user buttons must wrap its work with withBusyButton
    required_busy_handlers = [
        "stopScopedJob",
        "startWorkSync",
        "confirmEndSession",
        "toggleWatcher",
        "submitNewProject",
        "scanScannerNow",
        "submitScannerSplit",
        "toggleScannerWatcher",
        "approveScannerItem",
        "rejectScannerItem",
        "requeueScannerItem",
        "clearScannerResolved",
        "bulkScannerAction",
        "saveCustomSections",
        "scanInboxNow",
        "launchAgentCli",
        "runLegacyAction",
        "addContact",
        "delContact",
        "sendWhatsAppMsg",
        "saveMissingReplyDate",
        "linkExternalProjectLog",
        "flushLogQueue",
        "saveLogMapper",
        "saveBaselineCounters",
        "_execLinkPair",
        "unlinkTargetLetter",
        "openAutoLinkModal",
        "applySelectedAutoLinks",
        "runStudioAiSuggest",
        "saveStudioProfile",
        "runIssuePreview",
        "submitIssueDocument",
        "resolveDecision",
        "undoRename",
        "addLearnedRule",
        "saveGlobalSettings",
        "saveAiKey",
        "saveBackupAiKey",
        "removeBackupAiKey",
        "promoteBackupAiKey",
        "saveProjectSettings",
        "saveWaSettings",
        "startOrganizerScan",
        "updateOrganizerItem",
        "bulkOrganizerAction",
        "startOrganizerExecute",
        "aiRenameSingle",
        "aiRenameOrganizerBatch",
        "runOrganizerAudit",
        "registerUnloggedFiles",
        "browsePdfForSplit",
        "runPdfSplit",
        "browsePdfAddMerge",
        "runPdfMerge",
        "waScanLocalFolder",
        "waRecomputePlan",
        "waUpdateItem",
        "waBulkAction",
        "waAiRenameBatch",
        "waAiRenameSingle",
        "waExecutePlan",
        "waUndoRename",
        "waClearScreen",
        "dupScanLocalFolder",
        "dupQuickScan",
        "dupUpdateItem",
        "dupGroupAction",
        "dupBulkAction",
        "dupExecutePlan",
        "dupUndoQuarantine",
        "dupClearScreen",
        "toggleProjectEnabled",
        "deleteProjectConfirmed",
        "installDesktopAppShortcut",
        "openStandaloneWindowNow",
        "executeAppShutdown",
        "executeSystemReset",
        "addQuickTask",
        "toggleManualTaskDone",
        "deleteManualTaskItem",
        "snoozeAutoTask",
        "restoreAutoTask",
        "clearDoneManualTasks",
    ]
    for fn in required_busy_handlers:
        m = re.search(rf"async function {re.escape(fn)}\s*\([^)]*\)\s*\{{([\s\S]{{0,3200}}?)\n\}}", html)
        assert m is not None, f"Missing async function {fn}"
        assert "withBusyButton" in m.group(1), f"Handler {fn} does not use withBusyButton"

    # Long-running view buttons also surface busy text and disabled state during active jobs
    for label_var in ("orgScanLbl", "orgAiLbl", "orgExecLbl", "waScanLbl", "waAiLbl", "waExecLbl", "dupScanLbl", "dupExecLbl", "scanBtnLbl"):
        assert label_var in html, f"Missing running-state button label {label_var}"


def test_ui_layout_css_variables_and_dom_ids_integrity():
    import re
    root_dir = Path(__file__).resolve().parent.parent
    html = (root_dir / "doc_tools" / "dashboard_ui.html").read_text(encoding="utf-8")
    css = html.split("<style>", 1)[1].split("</style>", 1)[0]
    js = html.split("<script>", 1)[1].rsplit("</script>", 1)[0]

    # 1) All CSS custom properties referenced via var(--...) must be defined in :root
    defined_vars = set(re.findall(r"--([a-zA-Z0-9_-]+)\s*:", css))
    used_vars = set(re.findall(r"var\(--([a-zA-Z0-9_-]+)\)", html))
    missing_vars = sorted(used_vars - defined_vars)
    assert not missing_vars, f"Undefined CSS variables: {missing_vars}"

    # 2) Modal footer bug check: never reference non-existent 'modalFt', must use 'modalFoot'
    assert "modalFt" not in js, "Broken modal footer ID 'modalFt' found; expected 'modalFoot'"
    assert 'id="modalFoot"' in html

    # 3) No duplicate single-class .kpi definition
    assert len(re.findall(r"^\.kpi\s*\{", css, re.MULTILINE)) == 1

    # 4) Dynamic topbar height sync + non-jumping withBusyButton width guard
    assert "--topbar-h" in css and "_syncTopbarHeight" in js
    assert "b.style.minWidth" in js

    # 5) Brand theme palette from attached spec (#FFA611, #F6820D, #FFCB2B, #059BE5, #FF8964) + coordinated topbar
    for hex_clr in ("#FFA611", "#F6820D", "#FFCB2B", "#059BE5", "#FF8964"):
        assert hex_clr in css, f"Missing brand palette color {hex_clr} in dashboard CSS"
    assert '<meta name="theme-color" content="#059BE5">' in html
    assert "linear-gradient(125deg, #035D89 0%, #0377B5 54%, #059BE5 100%)" in css
    assert ".tbtn.start { background: #FFCB2B;" in css

    # 6) Install/Update scripts must NOT copy Smart_Folder_Agent.bat onto Desktop (and must clean legacy copies)
    for script_name in ("1_INSTALL_FIRST_TIME.bat", "UPDATE_QUICK.bat"):
        txt = (root_dir / script_name).read_text(encoding="utf-8")
        assert not re.search(r"copy\s+/Y\b[^\r\n]*Smart_Folder_Agent\.bat", txt, re.I), script_name
        assert re.search(r"del\s+/F\s+/Q\b[^\r\n]*Desktop\\Smart_Folder_Agent\.bat", txt, re.I), script_name

    # 7) Collapsible sidebar + collapsible sections group
    assert "toggleSidebarCollapse" in js and "toggleNavGroup" in js
    assert "sla_nav_grp_sections" in js and "sla_sidebar_collapsed" in js
    assert ".shell.sb-collapsed" in css and ".nav-group-body.is-collapsed" in css


def test_uuid_imports_for_uploads_without_batch_id_and_pdf_tools(live_server, monkeypatch):
    import base64
    import types
    import urllib.request
    import project_organizer as porg

    base, hub, tmp_path = live_server
    _, pr = _http_json(f"{base}/api/projects", method="POST", body={"name": "اختبار UUID", "code": "UID-1"})
    pid = pr["project"]["id"]

    # 1) Duplicates/WhatsApp raw binary upload WITHOUT X-Upload-Batch header (exercises uuid.uuid4 at project_dashboard.py:3146)
    req = urllib.request.Request(
        f"{base}/api/projects/{pid}/duplicates/upload-file",
        data=b"%PDF-1.4\n%%EOF\n",
        method="POST",
        headers={"Content-Type": "application/octet-stream", "X-Upload-Filename": "no_batch.pdf"},
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        assert resp.status == 200
        up_res = json.loads(resp.read().decode("utf-8"))
        assert up_res["ok"] is True and up_res["batch_id"]

    # 2) WhatsApp upload WITHOUT batch_id (exercises uuid.uuid4 at project_dashboard.py:3740)
    st_wa, r_wa = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/upload-file",
        method="POST",
        body={"name": "_chat.txt", "data_b64": base64.b64encode(b"[1/1/26, 10:00:00] A: hi\n").decode()},
    )
    assert st_wa == 200 and r_wa["ok"] is True and r_wa["batch_id"]

    # 3) PDF split and merge WITHOUT real PyMuPDF installed (stub fitz to exercise uuid.uuid4 in project_organizer.py)
    class _FakeDoc:
        def __init__(self, path=None):
            self.page_count = 3 if path else 0
        def __len__(self):
            return self.page_count
        def insert_pdf(self, src, from_page=0, to_page=None):
            if to_page is None:
                self.page_count += len(src)
            else:
                self.page_count += (to_page - from_page + 1)
        def save(self, out_path):
            Path(out_path).write_bytes(b"%PDF-1.4\n%%EOF\n")
        def close(self):
            pass

    fake_fitz = types.SimpleNamespace(open=lambda path=None: _FakeDoc(path))
    monkeypatch.setattr(porg, "fitz", fake_fitz)

    src1 = tmp_path / "sample_a.pdf"
    src2 = tmp_path / "sample_b.pdf"
    src1.write_bytes(b"%PDF-1.4\n%%EOF\n")
    src2.write_bytes(b"%PDF-1.4\n%%EOF\n")

    sp_dir = tmp_path / "split_out"
    sp_res = porg.split_pdf_document(src1, sp_dir, every_n_pages=2)
    assert sp_res["ok"] is True and sp_res["parts_created"] == 2
    assert len(list(sp_dir.glob("*.pdf"))) == 2

    mg_out = tmp_path / "merged_out.pdf"
    mg_res = porg.merge_pdf_documents([src1, src2], mg_out)
    assert mg_res["ok"] is True and mg_out.exists()



