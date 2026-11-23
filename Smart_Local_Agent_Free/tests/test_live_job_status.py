#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for system-wide live job progress, completion, error, cancellation, and UI status hosts."""

from __future__ import annotations

import json
import sys
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

from PIL import Image, ImageDraw
import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import content_duplicates as cdup  # noqa: E402
import project_dashboard as pdash  # noqa: E402
import whatsapp_archiver as wa  # noqa: E402


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


def _make_sample_img(path: Path, color: tuple[int, int, int] = (20, 120, 200)) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.new("RGB", (64, 64), color)
    d = ImageDraw.Draw(im)
    d.rectangle([8, 8, 56, 56], outline=(255, 255, 255), width=3)
    im.save(path)


def _wait_job_done(url: str, *, timeout_sec: float = 15.0) -> dict:
    deadline = time.monotonic() + timeout_sec
    last_data: dict = {}
    while time.monotonic() < deadline:
        st, last_data = _http_json(url)
        assert st == 200
        job = last_data.get("job") or {}
        if not job.get("running") and job.get("phase") in ("done", "error", "stopped"):
            return last_data
        time.sleep(0.05)
    raise AssertionError(f"Job at {url} did not finish in {timeout_sec}s; last={last_data.get('job')}")


def test_duplicates_async_progress_done_error_stop_and_dismiss(live_server, monkeypatch):
    base, _hub, tmp_path = live_server
    st, p_res = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع التشابه", "code": "DUP"})
    assert st in (200, 201)
    pid = p_res["project"]["id"]

    src_dir = tmp_path / "dup_src"
    _make_sample_img(src_dir / "img_a.png", (30, 140, 210))
    (src_dir / "img_a_copy.png").write_bytes((src_dir / "img_a.png").read_bytes())

    # 1) Start non-blocking duplicate scan -> returns immediately with running job
    st, start_res = _http_json(
        f"{base}/api/projects/{pid}/duplicates/scan",
        method="POST",
        body={"source_kind": "folder", "source_path": str(src_dir), "blocking": False},
    )
    assert st == 200
    assert start_res["ok"] is True
    assert start_res["started"] is True
    assert "job" in start_res

    # 2) Poll GET /duplicates until phase == 'done'
    done_data = _wait_job_done(f"{base}/api/projects/{pid}/duplicates")
    job = done_data["job"]
    assert job["running"] is False
    assert job["phase"] == "done"
    assert job["kind"] == "scan"
    assert job["summary"]["groups_count"] == 1
    assert done_data["total_groups"] == 1

    # 3) Dismiss job -> cleared from GET /duplicates
    st, dsm = _http_json(f"{base}/api/projects/{pid}/duplicates/job-dismiss", method="POST", body={})
    assert st == 200 and dsm["ok"] is True
    _, after_dsm = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert after_dsm["job"].get("phase", "") == ""

    # 4) Async execute (quarantine) transitions to done
    st, exec_start = _http_json(
        f"{base}/api/projects/{pid}/duplicates/execute",
        method="POST",
        body={"mode": "quarantine", "blocking": False},
    )
    assert st == 200 and exec_start["started"] is True
    exec_done = _wait_job_done(f"{base}/api/projects/{pid}/duplicates")
    assert exec_done["job"]["phase"] == "done"
    assert exec_done["job"]["kind"] == "execute"
    assert exec_done["job"]["summary"]["processed"] >= 1

    # 5) Error case: non-existent folder produces visible phase=='error' + message
    st, _ = _http_json(
        f"{base}/api/projects/{pid}/duplicates/scan",
        method="POST",
        body={"source_kind": "folder", "source_path": str(tmp_path / "does_not_exist"), "blocking": False},
    )
    assert st == 200
    err_data = _wait_job_done(f"{base}/api/projects/{pid}/duplicates")
    assert err_data["job"]["phase"] == "error"
    assert err_data["job"]["error"]

    # 6) Cancellation via POST /duplicates/stop
    entered = threading.Event()
    orig_scan = cdup.scan_content_duplicates

    def slow_scan(*args, **kwargs):
        stop_ev = kwargs.get("stop_event")
        cb = kwargs.get("progress_cb")
        if cb:
            cb({"phase": "hashing", "processed": 1, "total": 10, "current_file": "slow_file.png"})
        entered.set()
        for _ in range(100):
            if stop_ev is not None and stop_ev.is_set():
                raise cdup.DuplicatesStoppedError("stopped")
            time.sleep(0.03)
        return orig_scan(*args, **kwargs)

    monkeypatch.setattr(cdup, "scan_content_duplicates", slow_scan)
    st, _ = _http_json(
        f"{base}/api/projects/{pid}/duplicates/scan",
        method="POST",
        body={"source_kind": "folder", "source_path": str(src_dir), "blocking": False},
    )
    assert st == 200
    assert entered.wait(timeout=5)
    # While running, /api/state exposes it in active_jobs
    st_hub, hub_state = _http_json(f"{base}/api/state")
    assert st_hub == 200
    assert any(j.get("scope") == "duplicates" and j.get("running") for j in hub_state.get("active_jobs", []))

    st, stop_res = _http_json(f"{base}/api/projects/{pid}/duplicates/stop", method="POST", body={})
    assert st == 200 and stop_res["ok"] is True
    stopped_data = _wait_job_done(f"{base}/api/projects/{pid}/duplicates")
    assert stopped_data["job"]["phase"] == "stopped"
    assert stopped_data["job"]["running"] is False


def test_whatsapp_and_scanner_async_jobs_error_stop_and_dismiss(live_server, monkeypatch):
    base, hub, tmp_path = live_server
    st, p_res = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع الواتساب", "code": "WA"})
    assert st in (200, 201)
    pid = p_res["project"]["id"]
    proj = hub.get_project(pid)

    wa_dir = tmp_path / "wa_src"
    wa_dir.mkdir(parents=True, exist_ok=True)
    (wa_dir / "_chat.txt").write_text("[18/11/2026, 10:00:00 AM] Ali: IMG-20261118-WA0001.jpg (file attached)\n", encoding="utf-8")
    _make_sample_img(wa_dir / "IMG-20261118-WA0001.jpg", (200, 80, 40))

    # 1) Non-blocking WhatsApp scan -> done
    st, wa_start = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/scan",
        method="POST",
        body={"source_path": str(wa_dir), "blocking": False},
    )
    assert st == 200 and wa_start["started"] is True
    wa_done = _wait_job_done(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert wa_done["job"]["phase"] == "done"
    assert wa_done["job"]["kind"] == "scan"
    assert wa_done["counts"]["total"] == 1

    # 2) Non-blocking WhatsApp execute -> done
    st, wa_exec = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/execute",
        method="POST",
        body={"mode": "copy_by_date", "blocking": False},
    )
    assert st == 200 and wa_exec["started"] is True
    wa_exec_done = _wait_job_done(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert wa_exec_done["job"]["phase"] == "done"
    assert wa_exec_done["job"]["kind"] == "execute"

    # 3) WhatsApp error on bad source path + dismiss
    st, _ = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/scan",
        method="POST",
        body={"source_path": str(tmp_path / "missing_wa_folder"), "blocking": False},
    )
    assert st == 200
    wa_err = _wait_job_done(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert wa_err["job"]["phase"] == "error"
    assert wa_err["job"]["error"]

    st, _ = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive/job-dismiss", method="POST", body={})
    assert st == 200
    _, wa_cleared = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert wa_cleared["job"].get("phase", "") == ""

    # 4) WhatsApp cancellation
    entered_wa = threading.Event()

    def slow_wa_scan(*args, **kwargs):
        stop_ev = kwargs.get("stop_event")
        cb = kwargs.get("progress_cb")
        if cb:
            cb({"phase": "analyzing", "processed": 1, "total": 5, "current_file": "chat.txt"})
        entered_wa.set()
        for _ in range(100):
            if stop_ev is not None and stop_ev.is_set():
                raise wa.WhatsappStoppedError("stopped")
            time.sleep(0.03)
        raise AssertionError("Expected stop")

    monkeypatch.setattr(wa, "scan_and_match_whatsapp_source", slow_wa_scan)
    st, _ = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/scan",
        method="POST",
        body={"source_path": str(wa_dir), "blocking": False},
    )
    assert st == 200 and entered_wa.wait(timeout=5)
    st, _ = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive/stop", method="POST", body={})
    assert st == 200
    wa_stopped = _wait_job_done(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert wa_stopped["job"]["phase"] == "stopped"

    # 5) Scanner non-blocking scan & dismiss
    inbox_dir = Path(proj["inbox_dir"])
    inbox_dir.mkdir(parents=True, exist_ok=True)
    (inbox_dir / "sample_scan.txt").write_text("مشروع الواتساب\nمذكرة اختبار", encoding="utf-8")
    st, sc_start = _http_json(
        f"{base}/api/projects/{pid}/scanner/scan",
        method="POST",
        body={"blocking": False},
    )
    assert st == 200 and sc_start["started"] is True
    sc_done = _wait_job_done(f"{base}/api/projects/{pid}/scanner")
    assert sc_done["job"]["phase"] == "done"
    assert sc_done["job"]["summary"]["added"] >= 1
    st, _ = _http_json(f"{base}/api/projects/{pid}/scanner/job-dismiss", method="POST", body={})
    assert st == 200
    _, sc_after = _http_json(f"{base}/api/projects/{pid}/scanner")
    assert sc_after["job"].get("phase", "") == ""

    # 6) Organizer job dismiss endpoint
    st, org_dsm = _http_json(f"{base}/api/projects/{pid}/organizer/job-dismiss", method="POST", body={})
    assert st == 200 and org_dsm["ok"] is True


def test_dashboard_ui_contains_live_status_hosts_and_no_stale_plain_text():
    ui_path = TOOLS_DIR / "dashboard_ui.html"
    strings_path = TOOLS_DIR / "strings_ar.json"
    html = ui_path.read_text(encoding="utf-8")
    strings = json.loads(strings_path.read_text(encoding="utf-8"))["strings"]

    for marker in (
        'id="hdrGlobalActivity"',
        'id="dupJobHost"',
        'id="waJobHost"',
        'id="scannerJobHost"',
        'id="ovJobHost"',
        'id="orgProgressHost"',
        "function _renderUnifiedJobHtml(",
        "function updateGlobalActivityBar(",
        "function withBusyButton(",
        "AbortController",
        "S.serverOffline",
        "blocking: false",
    ):
        assert marker in html, f"Missing UI live-status element: {marker}"

    for k in (
        "jobs.running_badge",
        "jobs.done_badge",
        "jobs.stopped_badge",
        "jobs.error_badge",
        "jobs.offline_error",
        "jobs.timeout_error",
        "jobs.stop_btn",
        "jobs.dismiss_btn",
        "jobs.busy_wait",
        "jobs.topbar_running",
    ):
        assert k in strings
        entry = strings[k]
        if isinstance(entry, dict):
            assert str(entry.get("msa") or "").strip(), f"Missing msa text for: {k}"
            assert str(entry.get("eg") or "").strip(), f"Missing eg text for: {k}"
        else:
            assert str(entry).strip(), f"Missing Arabic string: {k}"
