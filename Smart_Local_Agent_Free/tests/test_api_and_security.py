from __future__ import annotations

import base64
import io
import json
import os
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

import openpyxl
import pytest

DOC_TOOLS = Path(__file__).resolve().parent.parent / "doc_tools"
if str(DOC_TOOLS) not in sys.path:
    sys.path.insert(0, str(DOC_TOOLS))

import project_dashboard as pdash
from security_utils import scrub_secrets


def _xlsx_bytes(rows: list[list], sheet_title: str = "Submittals") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _http_json(url: str, *, method: str = "GET", body: dict | None = None, headers: dict | None = None) -> tuple[int, dict]:
    hdrs = {"Accept": "application/json"}
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        hdrs["Content-Type"] = "application/json"
    if headers:
        hdrs.update(headers)
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


def test_end_to_end_api_uploaded_log_template_studio_issue_and_sessions(live_server):
    base, hub, tmp_path = live_server

    # 1) Offline UI HTML served cleanly with zero external CDN links
    with urllib.request.urlopen(f"{base}/", timeout=10) as r:
        html = r.read().decode("utf-8")
        assert r.status == 200
        assert 'dir="rtl"' in html
        assert "cdn." not in html.lower()
        assert "googleapis" not in html.lower()

    # 2) Create Project 1 with uploaded initial_log + Project 2 (verifies active_only auto-switch)
    log_b = _xlsx_bytes([
        ["No", "Reference", "Rev", "Discipline", "Description", "Submit Date", "Status", "Amount", "Link"],
        [1, "MAR-009", "R00", "EL", "اعتماد أسلاك", "2026-10-01", "A", "", ""],
    ])
    st1, r1 = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "مشروع المركز الصحي",
        "code": "HC-01",
        "contract_no": "MOH-2026-07",
        "owner": "وزارة الصحة",
        "consultant": "دار الهندسة",
        "contractor": "شركة الخليج للمقاولات",
        "initial_log": {"name": "HC01_Log.xlsx", "data_b64": base64.b64encode(log_b).decode()},
    })
    assert st1 == 201 and r1["ok"] is True
    p1_id = r1["project"]["id"]

    st2, r2 = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع المخازن", "code": "WH-02"})
    assert st2 == 201
    p2_id = r2["project"]["id"]

    # State shows process_mode == active_only when >1 project, and P2 has ZERO generic logs & ZERO system templates!
    _, state = _http_json(f"{base}/api/state")
    assert state["version"] == "10.13.2"
    assert state["settings"]["process_mode"] == "active_only"
    p2_card = next(p for p in state["projects"] if p["id"] == p2_id)
    assert p2_card["logs"] == []
    assert p2_card["templates_count"] == 0

    # 3) Start Work / Sync on Project 1 -> activates P1 & computes next MAR = MAR-010 from uploaded log!
    st_s, sync = _http_json(f"{base}/api/projects/{p1_id}/start", method="POST", body={})
    assert st_s == 200 and sync["counters"]["next_preview"]["MAR"] == "MAR-010"

    # 4) Section screen starts with ZERO templates (strictly no system templates)
    _, sec0 = _http_json(f"{base}/api/projects/{p1_id}/sections/submittal_mar")
    assert sec0["templates"] == []
    assert sec0["next_ref"] == "MAR-010"

    # 5) Upload a user MAR .xlsx template -> render -> save profile -> preview -> issue!
    tpl_b = _xlsx_bytes([
        ["نموذج اعتماد مواد (MAR)", "", "", ""],
        ["المشروع:", "", "المرجع:", ""],
        ["المادة المقدمة:", "", "المبلغ (د.ك):", ""],
    ], sheet_title="MAR")
    st_u, up = _http_json(f"{base}/api/projects/{p1_id}/templates/upload", method="POST", body={
        "name": "Company_MAR_Form.xlsx",
        "data_b64": base64.b64encode(tpl_b).decode(),
    })
    assert st_u == 201
    tname = up["name"]

    st_r, ren = _http_json(f"{base}/api/projects/{p1_id}/templates/{tname}/render")
    assert st_r == 200 and ren["render"]["kind"] == "xlsx"

    st_p, saved_p = _http_json(f"{base}/api/projects/{p1_id}/templates/{tname}/profile", method="POST", body={
        "section": "submittal_mar",
        "ref_prefix": "MAR",
        "fields": [
            {"key": "proj", "label": "المشروع", "role": "auto_project", "source": "project.name", "target": {"sheet": "MAR", "coord": "B2"}},
            {"key": "ref_no", "label": "المرجع", "role": "auto_system", "source": "system.next_ref", "target": {"sheet": "MAR", "coord": "D2"}},
            {"key": "subject", "label": "المادة المقدمة", "role": "manual", "input_type": "text", "required": True, "target": {"sheet": "MAR", "coord": "B3"}},
            {"key": "amount", "label": "المبلغ", "role": "manual", "input_type": "kwd", "target": {"sheet": "MAR", "coord": "D3"}},
        ],
    })
    assert st_p == 200 and saved_p["profile"]["section_confirmed"] is True

    # Live Preview does NOT consume counter or write to log
    st_pv, pv = _http_json(f"{base}/api/projects/{p1_id}/templates/{tname}/preview", method="POST", body={
        "values": {"subject": "قواطع كهربائية شنايدر", "amount": "320.500"}
    })
    assert st_pv == 200 and pv["preview"] is True and pv["next_ref"] == "MAR-010"

    # Official Issue saves into 00_Editable_Originals AND writes MAR-010 (Under Review) into HC01_Log.xlsx!
    st_i, issued = _http_json(f"{base}/api/projects/{p1_id}/templates/{tname}/issue", method="POST", body={
        "discipline": "EL",
        "values": {"subject": "قواطع كهربائية شنايدر", "amount": "320.500"},
    })
    assert st_i == 200 and issued["ref_no"] == "MAR-010"
    assert "00_Editable_Originals" in issued["output_rel"]
    assert issued["log_result"]["ok"] is True

    # Download issued document + confirm section screen now has 1 uploaded template, 2 rows & next_ref == MAR-011!
    with urllib.request.urlopen(f"{base}{issued['download_url']}", timeout=10) as dl:
        wb_dl = openpyxl.load_workbook(io.BytesIO(dl.read()))
        assert wb_dl["MAR"]["D2"].value == "MAR-010"
        assert wb_dl["MAR"]["B2"].value == "مشروع المركز الصحي"
        assert wb_dl["MAR"]["B3"].value == "قواطع كهربائية شنايدر"
        wb_dl.close()

    _, sec1 = _http_json(f"{base}/api/projects/{p1_id}/sections/submittal_mar")
    assert len(sec1["templates"]) == 1
    assert sec1["next_ref"] == "MAR-011"
    assert any(r["ref_no"] == "MAR-010" and r["status"] == "Under Review" for r in sec1["rows"])

    # 6) Learned Rules CRUD
    _, r_add = _http_json(f"{base}/api/learned-rules", method="POST", body={
        "match_type": "entity", "pattern": "مطعم الساحل", "action": "rename_in_place"
    })
    rid = r_add["rule"]["id"]
    _, r_list = _http_json(f"{base}/api/learned-rules")
    assert any(x["id"] == rid for x in r_list["rules"])
    st_del, _ = _http_json(f"{base}/api/learned-rules/{rid}", method="DELETE")
    assert st_del == 200

    # 7) Archive template (zero deletion — moved to _Archived_Templates)
    st_arc, arc_r = _http_json(f"{base}/api/projects/{p1_id}/templates/{tname}/archive", method="POST", body={})
    assert st_arc == 200
    p1_root = Path(r1["project"]["root"])
    assert (p1_root / "4_Project_Templates" / "_Archived_Templates" / arc_r["archived_to"]).exists()

    # 8) End Session creates dated backup & clears active_project_id
    st_end, end_r = _http_json(f"{base}/api/projects/{p1_id}/end", method="POST", body={"copy_external": False})
    assert st_end == 200 and end_r["active_cleared"] is True
    _, state_after = _http_json(f"{base}/api/state")
    assert state_after["active_project_id"] is None

    # 9) Unlink Project 2 preserves folder on disk (zero deletion)
    p2_root = Path(r2["project"]["root"])
    st_un, _ = _http_json(f"{base}/api/projects/{p2_id}/unlink", method="POST", body={})
    assert st_un == 200 and p2_root.exists()


def test_security_bind_origin_traversal_symlink_upload_and_secrets(live_server):
    base, hub, tmp_path = live_server

    # 1) Non-local bind blocked unless --allow-remote-preview explicitly passed
    with pytest.raises((ValueError, PermissionError)):
        pdash.create_server(hub, host="0.0.0.0", port=0, allow_remote_preview=False)

    # Create a project to test project-scoped security boundaries
    _, pr = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع أمان", "code": "SEC-1"})
    pid = pr["project"]["id"]
    proot = Path(pr["project"]["root"])

    # 2) Cross-site POST Origin AND GET Host/Origin (DNS rebinding) rejected with 403
    code_csrf, _ = _http_json(
        f"{base}/api/settings", method="POST",
        body={"dry_run": True},
        headers={"Origin": "https://evil-attacker.example"},
    )
    assert code_csrf == 403

    import http.client
    from urllib.parse import urlparse
    parsed_base = urlparse(base)
    for get_path in ("/api/state", "/api/settings", "/"):
        conn = http.client.HTTPConnection(parsed_base.hostname, parsed_base.port, timeout=5)
        conn.request("GET", get_path, headers={"Host": "evil.example.com"})
        resp = conn.getresponse()
        assert resp.status == 403, f"DNS rebinding GET {get_path} must be 403, got {resp.status}"
        conn.close()
    st_get_orig, _ = _http_json(f"{base}/api/state", headers={"Origin": "https://evil.example.com"})
    assert st_get_orig == 403

    # 3) Path traversal attempts (raw ../, URL-encoded %2e%2e, null byte %00) -> 403
    secret_outside = tmp_path / "outside_secret.txt"
    secret_outside.write_text("TOP_SECRET", encoding="utf-8")
    for bad_url in (
        f"{base}/api/projects/{pid}/download?rel=../../outside_secret.txt",
        f"{base}/api/projects/{pid}/download?rel=%2e%2e%2f%2e%2e%2foutside_secret.txt",
        f"{base}/api/projects/{pid}/download?rel=1_Inbox/a%00.pdf",
        f"{base}/api/projects/{pid}/templates/%2e%2e/render",
    ):
        st, _ = _http_json(bad_url)
        assert st == 403, f"Expected 403 for {bad_url}, got {st}"

    # 4) Symlink inside project pointing outside -> rejected with 403!
    sym = proot / "1_Inbox" / "escape_link.txt"
    os.symlink(secret_outside, sym)
    st_sym, _ = _http_json(f"{base}/api/projects/{pid}/download?rel=1_Inbox/escape_link.txt")
    assert st_sym == 403

    # 5) Dangerous executable upload rejected with 400
    st_exe, _ = _http_json(f"{base}/api/projects/{pid}/upload", method="POST", body={
        "target": "inbox",
        "files": [{"name": "malware.exe", "data_b64": base64.b64encode(b"MZ").decode()}],
    })
    assert st_exe == 400

    # 6) File DELETE rejected with 405 (zero deletion policy)
    st_del, _ = _http_json(f"{base}/api/projects/{pid}", method="DELETE")
    assert st_del == 405

    # 7) Zero secret storage in settings, registry.json, or project files
    fake_secret = "AIzaSyD1234567890abcdefghijklmnopqrstuv"
    _http_json(f"{base}/api/settings", method="POST", body={"api_key": fake_secret, "gemini_api_key": fake_secret})
    reg_raw = hub.registry_path.read_text(encoding="utf-8")
    assert fake_secret not in reg_raw
    scrubbed = scrub_secrets({"api_key": fake_secret, "note": f"key={fake_secret}"})
    assert fake_secret not in json.dumps(scrubbed)


def test_project_disable_delete_standalone_app_and_close_guard(live_server):
    base, hub, tmp_path = live_server

    # 1) Create two projects and verify disabling stops watcher/active processing
    _, pr1 = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع تعطيل", "code": "DIS-1"})
    _, pr2 = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع حذف", "code": "DEL-2"})
    p1 = pr1["project"]["id"]
    p2 = pr2["project"]["id"]
    p1_root = Path(pr1["project"]["root"])
    p2_root = Path(pr2["project"]["root"])
    assert p1_root.exists() and p2_root.exists()

    # Enable all_projects mode and make p1 active
    _http_json(f"{base}/api/settings", method="POST", body={"process_mode": "all_projects"})
    _http_json(f"{base}/api/active-project", method="POST", body={"project_id": p1})

    # Disable p1 -> active_project_id cleared AND skipped by watcher cycle even in all_projects mode
    st_dis, r_dis = _http_json(f"{base}/api/projects/{p1}/toggle-enabled", method="POST", body={"enabled": False})
    assert st_dis == 200 and r_dis["enabled"] is False
    _, st_hub = _http_json(f"{base}/api/state")
    assert st_hub["active_project_id"] is None
    cycle_disabled = hub.run_watcher_cycle_once()
    assert p1 not in cycle_disabled["processed_project_ids"]
    assert p2 in cycle_disabled["processed_project_ids"]

    # Re-enable p1 -> included in watcher cycle again
    st_en, r_en = _http_json(f"{base}/api/projects/{p1}/toggle-enabled", method="POST", body={"enabled": True})
    assert st_en == 200 and r_en["enabled"] is True
    cycle_enabled = hub.run_watcher_cycle_once()
    assert p1 in cycle_enabled["processed_project_ids"]

    # 2) Delete mode="unlink" (delete_files=False) removes project from registry while keeping disk folder
    st_d1, r_d1 = _http_json(f"{base}/api/projects/{p1}/delete", method="POST", body={"delete_files": False})
    assert st_d1 == 200 and r_d1["ok"] is True and r_d1["deleted_files"] is False
    assert p1_root.exists()

    # Delete mode="purge" (delete_files=True) REQUIRES explicit confirm + matching project code/name
    st_d2_noack, _ = _http_json(f"{base}/api/projects/{p2}/delete", method="POST", body={"delete_files": True})
    assert st_d2_noack == 400 and p2_root.exists()
    st_d2_wrong, _ = _http_json(
        f"{base}/api/projects/{p2}/delete", method="POST",
        body={"delete_files": True, "confirm": True, "confirm_name": "WRONG-CODE"},
    )
    assert st_d2_wrong == 400 and p2_root.exists()
    st_d2, r_d2 = _http_json(
        f"{base}/api/projects/{p2}/delete", method="POST",
        body={"delete_files": True, "confirm": True, "confirm_name": "DEL-2"},
    )
    assert st_d2 == 200 and r_d2["ok"] is True and r_d2["deleted_files"] is True
    assert not p2_root.exists()

    # Safety guard: refuse deleting a project whose root matches projects_root or shares root with another
    _, pr_safe = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "حماية المسار", "code": "SAFE-1", "root": str(hub.projects_root),
    })
    st_bad, _ = _http_json(
        f"{base}/api/projects/{pr_safe['project']['id']}/delete",
        method="POST", body={"delete_files": True, "confirm": True, "confirm_name": "SAFE-1"},
    )
    assert st_bad == 400 and hub.projects_root.exists()

    # 3) PWA manifest, icons, desktop shortcut endpoint, and silent launcher --app mode
    with urllib.request.urlopen(f"{base}/manifest.webmanifest", timeout=5) as mf:
        assert mf.status == 200
        mf_data = json.loads(mf.read().decode("utf-8"))
        assert mf_data["display"] == "standalone"
        assert mf_data["start_url"] == "/"
        assert len(mf_data["icons"]) >= 2

    for ep in ("/favicon.ico", "/app-icon-192.png", "/app-icon-512.png"):
        with urllib.request.urlopen(f"{base}{ep}", timeout=5) as resp:
            assert resp.status == 200 and len(resp.read()) > 100

    st_sc, r_sc = _http_json(f"{base}/api/app/install-shortcut", method="POST", body={})
    assert st_sc == 200 and r_sc["ok"] is True

    vbs_txt = (Path(__file__).resolve().parent.parent / "0_OPEN_SILENT.vbs").read_text(encoding="utf-8")
    assert "--app=" in vbs_txt and "msedge.exe" in vbs_txt

    # 4) Close-confirmation guard setting + UI hooks + CSRF-protected shutdown endpoint
    _, s0 = _http_json(f"{base}/api/state")
    assert s0["settings"]["confirm_before_close"] is True
    _http_json(f"{base}/api/settings", method="POST", body={"confirm_before_close": False})
    _, s1 = _http_json(f"{base}/api/state")
    assert s1["settings"]["confirm_before_close"] is False

    with urllib.request.urlopen(f"{base}/", timeout=5) as idx:
        html = idx.read().decode("utf-8")
        for token in ("beforeunload", "openInstallAppModal", "confirmAndShutdownApp", "openDeleteProjectModal", "toggleProjectEnabled"):
            assert token in html

    st_evil_shut, _ = _http_json(
        f"{base}/api/app/shutdown", method="POST",
        body={"simulate": True}, headers={"Origin": "https://evil.example"},
    )
    assert st_evil_shut == 403

    hub.toggle_watcher(True)
    st_shut, r_shut = _http_json(f"{base}/api/app/shutdown", method="POST", body={"simulate": True})
    assert st_shut == 200 and r_shut["ok"] is True and r_shut["shutting_down"] is True
    assert hub.watcher_running is False


def test_system_reset_wipes_trial_data_safely(live_server):
    base, hub, tmp_path = live_server

    # External scanner directory outside project roots MUST NEVER be deleted by system reset
    ext_inbox = tmp_path / "Shared_External_Scanner"
    ext_inbox.mkdir(parents=True, exist_ok=True)
    important_scan = ext_inbox / "real_scan.pdf"
    important_scan.write_bytes(b"%PDF-1.4\n%%EOF\n")

    _, pr1 = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "مشروع تجريبي ١", "code": "TST-1", "inbox_dir": str(ext_inbox),
    })
    _, pr2 = _http_json(f"{base}/api/projects", method="POST", body={
        "name": "مشروع تجريبي ٢", "code": "TST-2",
    })
    p1_root = Path(pr1["project"]["root"])
    p2_root = Path(pr2["project"]["root"])
    assert p1_root.exists() and p2_root.exists()

    # Add a learned rule, a pending decision, custom settings, and WhatsApp config
    _http_json(f"{base}/api/learned-rules", method="POST", body={
        "match_type": "entity", "pattern": "شركة تجريبية", "action": "rename_in_place",
    })
    hub.decisions.enqueue_decision(
        kind="unrelated_file",
        source_project_id=pr1["project"]["id"],
        file_path=important_scan,
    )
    hub.decisions.add_ignored_hash("deadbeef" * 8, "ignored_test.pdf")
    _http_json(f"{base}/api/settings", method="POST", body={
        "ui_style": "eg", "backup_warn_days": 22, "dry_run": True,
    })
    _http_json(f"{base}/api/whatsapp/config", method="POST", body={
        "default_country_code": "965", "phone_number_id": "12345",
    })

    # 1) Missing confirm -> 400; evil Origin -> 403
    st_no_ack, _ = _http_json(f"{base}/api/system/reset", method="POST", body={"confirm": False})
    assert st_no_ack == 400
    st_evil, _ = _http_json(
        f"{base}/api/system/reset", method="POST",
        body={"confirm": True}, headers={"Origin": "https://evil.example"},
    )
    assert st_evil == 403

    # 2) Execute full reset wiping test projects/decisions/rules/folders while preserving external scanner
    st_ok, r_ok = _http_json(f"{base}/api/system/reset", method="POST", body={
        "confirm": True,
        "delete_project_folders": True,
        "keep_ai_key": True,
        "keep_whatsapp": False,
        "reset_settings": True,
    })
    assert st_ok == 200 and r_ok["ok"] is True
    assert r_ok["projects_removed"] == 2
    assert r_ok["folders_deleted"] == 2
    assert r_ok["decisions_cleared"] >= 1
    assert r_ok["rules_cleared"] >= 1
    assert not p1_root.exists() and not p2_root.exists()
    assert important_scan.exists()  # external scanner folder completely untouched!

    _, st_after = _http_json(f"{base}/api/state")
    assert st_after["projects"] == []
    assert st_after["active_project_id"] is None
    assert st_after["pending_decisions_count"] == 0
    assert st_after["learned_rules_count"] == 0
    assert st_after["settings"]["backup_warn_days"] == 7
    assert st_after["settings"]["dry_run"] is False
    assert st_after["settings"]["ui_style"] == "eg"

    wa_after = json.loads((hub.state_dir / "whatsapp_config.json").read_text(encoding="utf-8"))
    assert wa_after.get("phone_number_id") in (None, "")

    with urllib.request.urlopen(f"{base}/", timeout=5) as idx:
        html = idx.read().decode("utf-8")
        assert "openSystemResetModal" in html and "executeSystemReset" in html
        hub_fn = html.split("function renderHubView()", 1)[1].split("function openNewProjectModal()", 1)[0]
        assert "proj-card is-clickable" in hub_fn
        assert "onclick=\"navigate('overview',${jsStr(p.id)})\"" in hub_fn
        assert "event.stopPropagation();" in hub_fn
        assert "📊 ${esc(t('common.open'))}" not in hub_fn


def test_decisions_side_image_preview_pdf_image_and_ignore_all(live_server, monkeypatch):
    base, hub, tmp_path = live_server
    import sys
    from PIL import Image

    _, pr = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع معاينة", "code": "PRV-1"})
    pid = pr["project"]["id"]

    # 1) Create a real image file & a 3-page PDF stubbed via fake fitz
    scan_dir = tmp_path / "scans"
    scan_dir.mkdir(parents=True, exist_ok=True)
    img_path = scan_dir / "letter_scan.jpg"
    Image.new("RGB", (400, 560), color=(240, 248, 255)).save(img_path, format="JPEG")

    pdf_path = scan_dir / "multipage_letter.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n/Type /Page\n/Type /Page\n/Type /Page\n%%EOF\n")

    buf_png = io.BytesIO()
    Image.new("RGB", (320, 440), color=(255, 250, 240)).save(buf_png, format="PNG")
    sample_png = buf_png.getvalue()

    class _FakePix:
        def tobytes(self, fmt="png"):
            return sample_png

    class _FakePage:
        def get_pixmap(self, dpi=120):
            return _FakePix()

    class _FakeDoc:
        page_count = 3
        def __enter__(self): return self
        def __exit__(self, *a): pass
        def __getitem__(self, idx):
            if idx < 0 or idx >= 3:
                raise IndexError(idx)
            return _FakePage()

    class _FakeFitz:
        def open(self, p):
            return _FakeDoc()

    monkeypatch.setitem(sys.modules, "fitz", _FakeFitz())

    d_img = hub.decisions.enqueue_decision(kind="unrelated_file", source_project_id=pid, file_path=img_path)
    d_pdf = hub.decisions.enqueue_decision(kind="unrelated_file", source_project_id=pid, file_path=pdf_path)

    # 2) Image preview returns PNG with X-Total-Pages=1
    with urllib.request.urlopen(f"{base}/api/decisions/{d_img['id']}/preview?page=1&dpi=120", timeout=5) as r1:
        assert r1.status == 200
        assert r1.headers.get("Content-Type") == "image/png"
        assert r1.headers.get("X-Total-Pages") == "1"
        raw1 = r1.read()
        assert raw1.startswith(b"\x89PNG")

    # 3) PDF preview returns PNG with X-Total-Pages=3 and supports page navigation
    with urllib.request.urlopen(f"{base}/api/decisions/{d_pdf['id']}/preview?page=2&dpi=135", timeout=5) as r2:
        assert r2.status == 200
        assert r2.headers.get("X-Total-Pages") == "3"
        assert r2.headers.get("X-Page-Number") == "2"
        assert r2.read().startswith(b"\x89PNG")

    # 4) Original file download works & symlink is blocked with 403
    with urllib.request.urlopen(f"{base}/api/decisions/{d_img['id']}/file", timeout=5) as rf:
        assert rf.status == 200 and len(rf.read()) > 100

    sym_file = scan_dir / "bad_sym.jpg"
    os.symlink(img_path, sym_file)
    cur_decs = json.loads(hub.decisions.decisions_path.read_text(encoding="utf-8"))
    cur_decs.append({
        "id": "DEC-SYM-TEST",
        "kind": "unrelated_file",
        "status": "pending",
        "source_project_id": pid,
        "file_path": str(sym_file),
        "file_name": "bad_sym.jpg",
    })
    hub.decisions.decisions_path.write_text(json.dumps(cur_decs), encoding="utf-8")
    st_sym, _ = _http_json(f"{base}/api/decisions/DEC-SYM-TEST/preview")
    assert st_sym == 403

    # 5) UI contains side-by-side preview pane, zoom/page controls & collapsible extra log box
    with urllib.request.urlopen(f"{base}/", timeout=5) as idx:
        html = idx.read().decode("utf-8")
        for tok in ("decisions-split", "decPreviewPane", "loadDecisionSidePreview", "openScannerPagePreviewModal", "ignoreAllPendingDecisions", "extraLogBox"):
            assert tok in html

    # 6) Ignore-all blocks CSRF and clears pending decisions without deleting files
    st_evil, _ = _http_json(
        f"{base}/api/decisions/ignore-all", method="POST",
        body={}, headers={"Origin": "https://evil.example"},
    )
    assert st_evil == 403

    st_ia, r_ia = _http_json(f"{base}/api/decisions/ignore-all", method="POST", body={})
    assert st_ia == 200 and r_ia["ok"] is True and r_ia["cleared"] >= 2
    assert img_path.exists() and pdf_path.exists()
    _, st_after = _http_json(f"{base}/api/state")
    assert st_after["pending_decisions_count"] == 0


def test_arabic_folder_picker_utf8_b64_and_invalid_question_marks_guard(live_server, monkeypatch):
    base, hub, tmp_path = live_server
    import subprocess
    import types

    share_root = tmp_path / "ICBS SHARE"
    arabic_dir = share_root / "مستوصف صباح السالم - agrnt"
    arabic_dir.mkdir(parents=True, exist_ok=True)
    arabic_excel = arabic_dir / "لوج مستوصف صباح السالم.xlsm"
    arabic_excel.write_bytes(b"placeholder")

    # 1) _resolve_picker_initial_dir steps back to existing parent when path has `?`
    corrupted = f"{share_root}/????? ?????? ???? ?????? - agrnt"
    assert pdash._resolve_picker_initial_dir(corrupted) == str(share_root)
    assert pdash._has_invalid_win_question_marks(corrupted) is True
    assert pdash._has_invalid_win_question_marks(r"\\?\C:\Projects\مستوصف") is False

    # 2) Simulated Windows PowerShell folder/file picker preserves Arabic UTF-8 via B64 + env vars
    captured_calls = []

    def _fake_run(cmd, **kwargs):
        captured_calls.append({"cmd": cmd, "env": kwargs.get("env") or {}})
        target_str = str(arabic_excel if "SLA_PICK_FILTER" in (kwargs.get("env") or {}) else arabic_dir)
        b64 = base64.b64encode(target_str.encode("utf-8")).decode("ascii")
        return types.SimpleNamespace(returncode=0, stdout=f"B64:{b64}\n", stderr="")

    monkeypatch.setattr(pdash.sys, "platform", "win32")
    monkeypatch.setattr(subprocess, "run", _fake_run)

    ok_d, chosen_d, err_d = pdash.browse_local_folder(initial_dir=corrupted, title="اختيار مجلد")
    assert ok_d is True and err_d == "" and chosen_d == str(arabic_dir)
    assert captured_calls[-1]["env"]["SLA_PICK_INIT"] == str(share_root)
    assert captured_calls[-1]["env"]["SLA_PICK_TITLE"] == "اختيار مجلد"
    assert "Convert]::ToBase64String" in captured_calls[-1]["cmd"][-1]
    assert "$owner.TopMost = $true" in captured_calls[-1]["cmd"][-1]
    assert "$d.ShowDialog($owner)" in captured_calls[-1]["cmd"][-1]

    ok_f, chosen_f, err_f = pdash.browse_local_file(initial_dir=str(arabic_dir), title="تصفح ملف")
    assert ok_f is True and err_f == "" and chosen_f == str(arabic_excel)
    assert "$owner.TopMost = $true" in captured_calls[-1]["cmd"][-1]
    assert "$d.ShowDialog($owner)" in captured_calls[-1]["cmd"][-1]

    # 3) Creating project with corrupted `?` path returns clear 400 error instead of WinError 123
    st_bad, r_bad = _http_json(
        f"{base}/api/projects",
        method="POST",
        body={
            "name": "مستوصف صباح السالم",
            "code": "PRJ-01",
            "root": corrupted,
            "archive_dir": corrupted,
        },
    )
    assert st_bad == 400 and r_bad["ok"] is False and "?" not in r_bad["error"]

    # 4) Creating project with real Arabic paths succeeds cleanly
    ar_inbox = tmp_path / "سكانر المستوصف"
    st_ok, r_ok = _http_json(
        f"{base}/api/projects",
        method="POST",
        body={
            "name": "مستوصف صباح السالم",
            "code": "PRJ-01",
            "root": str(arabic_dir),
            "archive_dir": str(arabic_dir),
            "inbox_dir": str(ar_inbox),
        },
    )
    assert st_ok in (200, 201) and r_ok["ok"] is True
    assert r_ok["project"]["root"] == str(arabic_dir.resolve())
    assert (arabic_dir / "1_Inbox").is_dir() and (arabic_dir / "2_Archive").is_dir()


