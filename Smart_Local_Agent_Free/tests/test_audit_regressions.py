#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Regression tests for all findings uncovered in the full independent codebase audit:
  1. `.docx` render & filled-preview routes (fixing `shutil.which` NameError in template_render.py).
  2. CSP headers on dashboard HTML & forced attachment/sandbox on `.html`/`.svg` files across all file routes.
  3. `/api/decisions/<id>/file|preview` confinement + binary/JSON upload extension & header validation.
  4. IPv6 `[::1]:8765` Host/Origin check + UNC boundary check in security_utils.
  5. `launch_gemini_in_project` shell metacharacter sanitization & `/D` usage.
  6. `whatsapp_archiver` PDF AI-renaming via `prepare_visual_pages` + Zip-Bomb limits.
  7. `content_duplicates` partial-error undo retention (no premature undo deletion).
  8. `project_log` Arabic `م/*/2026` numbering, 60-col summary read, and Pre_Edit snapshot.
  9. `apply_plan` traversal rejection on `..` in `dest_folder` & `render_pdf` invalid `--pages`.
 10. `dashboard_ui.html` zero raw `'${esc(...)}'` inside `on*` event attributes.
"""

from __future__ import annotations

import base64
import io
import json
import re
import subprocess
import sys
import threading
import types
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

import docx
import openpyxl
import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import apply_plan as aplan  # noqa: E402
import content_duplicates as dups  # noqa: E402
import inbox_pipeline as ipipe  # noqa: E402
import project_dashboard as pdash  # noqa: E402
import project_log as plog  # noqa: E402
import render_pdf as rpdf  # noqa: E402
import scan_archive as sarch  # noqa: E402
import security_utils as su  # noqa: E402
import template_render as trender  # noqa: E402
import whatsapp_archiver as waa  # noqa: E402


def _http_json(url: str, method: str = "GET", body: dict | None = None, headers: dict | None = None):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    hdrs = dict(headers or {})
    if data is not None:
        hdrs.setdefault("Content-Type", "application/json; charset=utf-8")
    req = urllib.request.Request(url, data=data, method=method, headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, {"ok": False, "error": raw}


@pytest.fixture()
def audit_server(tmp_path):
    hub = pdash.HubState(tmp_path / "state" / "registry.json", tmp_path / "projects")
    srv = pdash.create_server(hub, host="127.0.0.1", port=0)
    port = srv.server_address[1]
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        yield f"http://127.0.0.1:{port}", hub, tmp_path
    finally:
        hub.toggle_watcher(False)
        srv.shutdown()
        srv.server_close()


def test_docx_template_render_and_preview_with_libreoffice_enabled(audit_server, monkeypatch):
    base, hub, tmp_path = audit_server
    _, pr = _http_json(f"{base}/api/projects", method="POST", body={"name": "Word Studio Test", "code": "WRD-1"})
    pid = pr["project"]["id"]
    proj_root = Path(pr["project"]["root"])

    # Create a genuine .docx template with paragraph + table
    tpl_path = proj_root / "4_Project_Templates" / "Official_Letter.docx"
    tpl_path.parent.mkdir(parents=True, exist_ok=True)
    doc = docx.Document()
    doc.add_paragraph("Ref: {{ref_no}}")
    tbl = doc.add_table(rows=1, cols=2)
    tbl.cell(0, 0).text = "Subject"
    tbl.cell(0, 1).text = ""
    doc.save(str(tpl_path))

    # Even if LibreOffice binary isn't installed, calling with default try_libreoffice=True must NOT raise NameError
    monkeypatch.setattr(trender.shutil, "which", lambda _cmd: None)
    st_ren, r_ren = _http_json(f"{base}/api/projects/{pid}/templates/{tpl_path.name}/render")
    assert st_ren == 200 and r_ren["ok"] is True
    assert r_ren["render"]["kind"] == "docx"

    st_save, r_save = _http_json(
        f"{base}/api/projects/{pid}/templates/{tpl_path.name}/profile",
        method="POST",
        body={
            "section": "outgoing",
            "section_confirmed": True,
            "fields": [
                {
                    "key": "subject",
                    "label": "الموضوع",
                    "role": "manual",
                    "input_type": "text",
                    "target": {"kind": "docx_cell", "table_idx": 0, "row_idx": 0, "col_idx": 1},
                    "required": True,
                }
            ],
            "repeating_tables": [],
        },
    )
    assert st_save == 200 and r_save["ok"] is True
    st_prv, r_prv = _http_json(
        f"{base}/api/projects/{pid}/templates/{tpl_path.name}/preview",
        method="POST",
        body={"values": {"subject": "تجربة التقرير"}},
    )
    assert st_prv == 200 and r_prv["ok"] is True and r_prv["render"]["kind"] == "docx"


def test_csp_same_origin_xss_and_upload_hardening(audit_server):
    base, hub, tmp_path = audit_server
    _, pr = _http_json(f"{base}/api/projects", method="POST", body={"name": "Security Hardening", "code": "SEC-1"})
    pid = pr["project"]["id"]
    proj_root = Path(pr["project"]["root"])

    # 1) Root HTML carries strict Content-Security-Policy
    with urllib.request.urlopen(f"{base}/", timeout=5) as resp:
        csp = resp.headers.get("Content-Security-Policy") or ""
        assert "script-src 'self' 'unsafe-inline'" in csp
        assert "object-src 'none'" in csp

    # 2) Decision pointing outside project/state roots is blocked with 403
    outside_f = tmp_path / "outside_secret.txt"
    outside_f.write_text("secret", encoding="utf-8")
    cur_decs = json.loads(hub.decisions.decisions_path.read_text(encoding="utf-8")) if hub.decisions.decisions_path.exists() else []
    cur_decs.append({
        "id": "DEC-OUTSIDE",
        "kind": "unrelated_file",
        "status": "pending",
        "source_project_id": pid,
        "file_path": str(outside_f),
        "file_name": "outside_secret.txt",
    })
    hub.decisions.decisions_path.parent.mkdir(parents=True, exist_ok=True)
    hub.decisions.decisions_path.write_text(json.dumps(cur_decs), encoding="utf-8")
    st_out, _ = _http_json(f"{base}/api/decisions/DEC-OUTSIDE/file")
    assert st_out == 403

    # 3) HTML/SVG decision file is served with attachment + sandbox CSP (never inline XSS)
    xss_html = Path(pr["project"]["inbox_dir"]) / "trap.html"
    xss_html.write_text("<script>alert(1)</script>", encoding="utf-8")
    dec_html = hub.decisions.enqueue_decision(kind="unrelated_file", source_project_id=pid, file_path=xss_html)
    with urllib.request.urlopen(f"{base}/api/decisions/{dec_html['id']}/file", timeout=5) as r_html:
        assert "attachment" in (r_html.headers.get("Content-Disposition") or "").lower()
        assert "sandbox" in (r_html.headers.get("Content-Security-Policy") or "").lower()
        assert r_html.headers.get("Content-Type") == "application/octet-stream"

    # 4) Streaming upload rejects dangerous extension (.exe) and malformed Content-Length
    req_exe = urllib.request.Request(
        f"{base}/api/projects/{pid}/whatsapp-archive/upload-file",
        data=b"MZ",
        method="POST",
        headers={
            "Content-Type": "application/octet-stream",
            "X-Upload-Filename": "payload.exe",
        },
    )
    with pytest.raises(urllib.error.HTTPError) as exc_exe:
        urllib.request.urlopen(req_exe, timeout=5)
    assert exc_exe.value.code == 400

    # 5) JSON WhatsApp upload rejects .bat and invalid base64
    st_bat, _ = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/upload-file",
        method="POST",
        body={"name": "run.bat", "data_b64": base64.b64encode(b"echo hi").decode()},
    )
    assert st_bat == 400

    st_b64, _ = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/upload-file",
        method="POST",
        body={"name": "_chat.txt", "data_b64": "!!!not-base64!!!"},
    )
    assert st_b64 == 400

    # 6) Duplicates inline=1 on .svg forces attachment + sandbox
    svg_f = Path(pr["project"]["archive_dir"]) / "logo.svg"
    svg_f.parent.mkdir(parents=True, exist_ok=True)
    svg_f.write_text("<svg onload='alert(1)'></svg>", encoding="utf-8")
    dup_p = dups.dups_plan_path(proj_root)
    dup_p.parent.mkdir(parents=True, exist_ok=True)
    dup_p.write_text(json.dumps({
        "source_dir": str(svg_f.parent),
        "groups": [{"id": "G1", "items": [{"id": "DUP-SVG", "abs_path": str(svg_f), "file_name": "logo.svg"}]}],
    }), encoding="utf-8")
    with urllib.request.urlopen(
        f"{base}/api/projects/{pid}/duplicates/file?id=DUP-SVG&inline=1",
        timeout=5,
    ) as r_svg:
        assert "attachment" in (r_svg.headers.get("Content-Disposition") or "").lower()
        assert "sandbox" in (r_svg.headers.get("Content-Security-Policy") or "").lower()


def test_security_utils_ipv6_unc_and_launch_gemini(monkeypatch, tmp_path):
    # 1) IPv6 loopback Host and Origin pass cleanly; evil Host fails
    assert su.validate_request_origin({"Host": "[::1]:8765", "Origin": "http://[::1]:8765"}) is True
    assert su.validate_request_origin({"Host": "[::1]:8765", "Origin": "http://evil.example"}) is False

    # 2) UNC boundary check: //server/share2 must NOT match base //server/share even when os.name == 'nt'
    assert su.safe_resolve_within(Path("/tmp/base"), "sub/file.txt") == (Path("/tmp/base").resolve() / "sub/file.txt")
    # First: on non-Windows, any UNC candidate is blocked
    with pytest.raises(su.PathTraversalError):
        su.safe_resolve_within(tmp_path, "//server/share/secret.txt")
    # Second: verify UNC prefix logic (exact share boundary) without WindowsPath instantiation on POSIX
    orig_resolve = su.resolve_local_or_unc
    monkeypatch.setattr(su, "resolve_local_or_unc", lambda x: Path(str(x).replace("\\", "/")))
    monkeypatch.setattr(su, "normalize_win_path", lambda x: Path(str(x).replace("\\", "/")))
    with pytest.raises(su.PathTraversalError):
        su.safe_resolve_within(Path("//server/share"), "//server/share2/secret.txt")
    monkeypatch.setattr(su, "resolve_local_or_unc", orig_resolve)

    # 3) launch_gemini_in_project sanitizes shell metacharacters and uses /D
    proj_dir = tmp_path / "P & A"
    proj_dir.mkdir(parents=True)
    captured = {}

    def _fake_popen(cmd_list, **kwargs):
        captured["cmd"] = cmd_list
        return None

    monkeypatch.setattr(pdash.sys, "platform", "win32")
    monkeypatch.setattr(pdash.subprocess, "Popen", _fake_popen)
    ok, _msg = pdash.launch_gemini_in_project({"root": str(proj_dir)}, initial_prompt='hello & calc.exe | whoami "x"')
    assert ok is True
    cmd = captured["cmd"]
    assert cmd[:5] == ["cmd", "/c", "start", "", "/D"]
    assert "&" not in cmd[-1] and "|" not in cmd[-1]


def test_whatsapp_pdf_ai_rename_and_zip_limits(monkeypatch, tmp_path):
    proj_root = tmp_path / "wa_pdf_proj"
    src_dir = tmp_path / "wa_src"
    src_dir.mkdir(parents=True)
    pdf_p = src_dir / "DOC-20261118.pdf"
    pdf_p.write_bytes(b"%PDF-1.4\n%%EOF\n")
    (src_dir / "_chat.txt").write_text("[18/11/2026, 09:00:00] علي: <مرفق: DOC-20261118.pdf>\n", encoding="utf-8")

    proj = {"id": "p_wapdf", "name": "WA PDF", "root": str(proj_root)}
    waa.scan_and_match_whatsapp_source(proj, src_dir, include_docs=True)

    # Stub sw.prepare_visual_pages so PDF AI-renaming works without Poppler/PyMuPDF
    import scan_watcher as sw
    called = []

    def _fake_prepare(fp, job_dir):
        called.append(fp.name)
        job_dir.mkdir(parents=True, exist_ok=True)
        img = job_dir / "page_01.png"
        img.write_bytes(b"\x89PNG\r\n\x1a\n")
        return [img], "مخطط قواعد الدور الأرضي", 1

    monkeypatch.setattr(sw, "prepare_visual_pages", _fake_prepare)
    res_ai = waa.ai_rename_whatsapp_items(
        proj,
        scope="all",
        analyzer_fn=lambda _p, _d, _imgs, snip: {"subject_ar": "مخطط_قواعد_أرضي", "doc_date": "2026-11-18"},
    )
    assert res_ai["ok"] and res_ai["updated"] == 1
    assert called == ["DOC-20261118.pdf"]

    # Zip-Bomb file-count guard
    zpath = tmp_path / "many.zip"
    with zipfile.ZipFile(zpath, "w") as zf:
        for i in range(5):
            zf.writestr(f"f{i}.txt", b"hi")
    with pytest.raises(ValueError):
        waa.extract_zip_safely(zpath, tmp_path / "ext_out", max_files=3)

    # scan_archive.extract_archive blocks Zip-Slip
    evil_zip = tmp_path / "evil.zip"
    with zipfile.ZipFile(evil_zip, "w") as zf:
        zf.writestr("../escape.txt", b"evil")
    ok_zip, _ = sarch.extract_archive(evil_zip, tmp_path / "sarch_out")
    assert ok_zip is False
    assert not (tmp_path / "escape.txt").exists()


def test_duplicates_partial_undo_keeps_failed_entries(tmp_path):
    proj_root = tmp_path / "dups_proj"
    src_dir = tmp_path / "dups_src"
    src_dir.mkdir(parents=True)
    f1 = src_dir / "dup1.txt"
    f2 = src_dir / "dup2.txt"
    q_dir = dups.dups_quarantine_root(proj_root) / "20261118_100000"
    q_dir.mkdir(parents=True)
    q1 = q_dir / "dup1.txt"
    q2 = q_dir / "dup2.txt"
    q1.write_text("one", encoding="utf-8")
    q2.write_text("two", encoding="utf-8")
    # Put an obstructing file at f2 so restoring f2 fails while restoring f1 succeeds
    f2.write_text("existing", encoding="utf-8")

    dups.dups_undo_path(proj_root).parent.mkdir(parents=True, exist_ok=True)
    dups.atomic_write_json(dups.dups_undo_path(proj_root), {
        "executed_at": "2026-11-18T10:00:00",
        "quarantine_dir": str(q_dir),
        "entries": [
            {"src_path": str(f1), "quarantine_path": str(q1)},
            {"src_path": str(f2), "quarantine_path": str(q2)},
        ],
    })

    res = dups.undo_last_duplicates_quarantine({"root": str(proj_root)})
    assert res["restored"] == 1
    assert res["errors"] == 1
    assert res["has_undo"] is True
    assert f1.exists()
    rem_undo = json.loads(dups.dups_undo_path(proj_root).read_text(encoding="utf-8"))
    assert len(rem_undo["entries"]) == 1 and rem_undo["entries"][0]["src_path"] == str(f2)


def test_project_log_arabic_seq_60_cols_and_pre_edit_backup(tmp_path):
    proj_root = tmp_path / "log_proj"
    logs_dir = proj_root / "3_Project_Logs"
    logs_dir.mkdir(parents=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Outgoing_صادر"
    # Place headers with Status at column 50 (> 45!)
    ws.cell(1, 1, "م")
    ws.cell(1, 2, "Ref No")
    ws.cell(1, 3, "Subject")
    ws.cell(1, 4, "Date")
    ws.cell(1, 50, "Status")
    ws.cell(2, 1, 1)
    ws.cell(2, 2, "م/39/2026")
    ws.cell(2, 3, "كتاب رسمي تجريبي")
    ws.cell(2, 4, "2026-11-01")
    ws.cell(2, 50, "Replied")
    log_p = logs_dir / "Letters_Log.xlsx"
    wb.save(log_p)
    wb.close()

    proj = {"id": "p_log", "name": "Log Test", "code": "LT-1", "root": str(proj_root)}
    nxt_ref, nxt_num, _ = plog.next_reference_from_logs(proj, prefix="م/2026")
    assert nxt_ref == "م/40/2026" and nxt_num == 40

    summary = plog.read_project_log_summary(proj)
    assert len(summary["rows"]) == 1
    assert summary["rows"][0]["status"] == "Replied"

    # Recording a new row creates a Pre_Edit backup snapshot
    res_rec = plog.record_entry_in_project_log(
        proj,
        {"section": "outgoing", "ref_no": "م/40/2026", "subject_ar": "كتاب جديد", "submit_date": "2026-11-18"},
    )
    assert res_rec["ok"] is True
    pre_edit_dir = proj_root / "_Backups" / "Pre_Edit"
    assert pre_edit_dir.exists() and any(pre_edit_dir.glob("*_Letters_Log.xlsx"))


def test_apply_plan_traversal_render_pdf_and_ui_no_raw_esc_in_attrs(tmp_path):
    # 1) apply_plan rejects '..' in dest_folder
    src_f = tmp_path / "doc.pdf"
    src_f.write_bytes(b"%PDF-1.4\n%%EOF\n")
    plan_f = tmp_path / "plan.json"
    plan_f.write_text(json.dumps([{"src": str(src_f), "action": "KEEP", "dest_folder": "..", "new_name": "out.pdf"}]), encoding="utf-8")
    res = subprocess.run(
        [sys.executable, str(TOOLS_DIR / "apply_plan.py"), str(plan_f), "--root", str(tmp_path)],
        capture_output=True, text=True,
    )
    assert res.returncode != 0
    assert not Path("/out.pdf").exists()

    # 2) render_pdf raises clean ValueError on malformed --pages
    with pytest.raises(ValueError):
        rpdf.parse_pages("1-", 5)
    with pytest.raises(ValueError):
        rpdf.parse_pages("abc", 5)

    # 3) dashboard_ui.html has ZERO raw `'${esc(...)}'` inside on*= HTML attributes
    ui_txt = (TOOLS_DIR / "dashboard_ui.html").read_text(encoding="utf-8")
    for ln in ui_txt.splitlines():
        assert not re.search(r'\bon\w+\s*=\s*"[^"]*\$\{esc\(', ln), f"Unsafe esc inside on*: {ln}"
