#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""End-to-end tests for v10.13.0 Smart Taskbar (Option C: Side Peek + Pinned Tray + Ctrl+K)."""

from __future__ import annotations

import datetime as dt
import json
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

import openpyxl
import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import project_dashboard as pdash  # noqa: E402
import task_bar as tbar  # noqa: E402


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


def _make_log_with_actions(path: Path) -> None:
    today = dt.date.today()
    d25 = (today - dt.timedelta(days=25)).isoformat()
    d18 = (today - dt.timedelta(days=18)).isoformat()
    d8 = (today - dt.timedelta(days=8)).isoformat()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MAR"
    ws.append(["Ref No", "Discipline", "Subject", "Date", "Reply Date", "Status"])
    ws.append(["MAR-EL-101", "EL", "Main Switchboard", d25, "", "Under Review"])
    ws.append(["MAR-EL-102", "EL", "Busbar Duct Sample", d18, d8, "Code C - Rejected"])

    ws_in = wb.create_sheet("Incoming")
    ws_in.append(["Ref No", "Party", "Subject", "Date", "Status"])
    ws_in.append(["IN-501", "Consultant", "Site Inspection Comment", d8, ""])

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def test_taskbar_manual_tasks_quick_tags_crud_and_zero_project_writes(live_server):
    base, hub, tmp_path = live_server
    p_root = tmp_path / "Proj_Tasks"
    _, r = _http_json(f"{base}/api/projects", method="POST", body={"name": "Marina Mall", "code": "MRN", "root": str(p_root)})
    pid = r["project"]["id"]

    # Empty title returns 400
    st_bad, _ = _http_json(f"{base}/api/tasks", method="POST", body={"title": "   "})
    assert st_bad == 400

    # Quick-add parses #عاجل and #غدا automatically and links project metadata cleanly
    tomorrow = (dt.date.today() + dt.timedelta(days=1)).isoformat()
    st1, r1 = _http_json(
        f"{base}/api/tasks",
        method="POST",
        body={"title": "متابعة مقاول التكييف #عاجل #غدا", "project_id": pid},
    )
    assert st1 == 201 and r1["ok"] is True
    tk1 = r1["task"]
    assert tk1["priority"] == "urgent"
    assert tk1["due_date"] == tomorrow
    assert tk1["project_code"] == "MRN"
    assert "#عاجل" not in tk1["title"] and "#غدا" not in tk1["title"]

    # Normal task without project
    st2, r2 = _http_json(f"{base}/api/tasks", method="POST", body={"title": "تجهيز جدول الاجتماعات"})
    assert st2 == 201 and r2["task"]["priority"] == "normal"

    # Stored strictly in _hub_state/tasks.json, never inside project root
    assert tbar.tasks_file_path(hub.state_dir).exists()
    assert not (p_root / "tasks.json").exists()

    # Hub state exposes tasks_summary badge counts
    _, state = _http_json(f"{base}/api/state")
    assert state["version"] == "10.13.2"
    assert state["build_tag"] == "2026-11-18-v10.13.2"
    assert state["tasks_summary"]["manual_open"] == 2
    assert state["tasks_summary"]["manual_urgent"] == 1

    # Toggle done -> archived into done_tasks ("اللي تم") with done_at & 1-click undo
    st_u, r_u = _http_json(f"{base}/api/tasks/{tk1['id']}", method="POST", body={"done": True})
    assert st_u == 200 and r_u["task"]["done"] is True

    _, list_open = _http_json(f"{base}/api/tasks?include_done=0")
    assert len(list_open["manual_tasks"]) == 1
    assert list_open["counts"]["manual_open"] == 1
    assert list_open["counts"]["manual_done"] == 1
    assert list_open["counts"]["done_total"] == 1
    assert any(d["id"] == tk1["id"] and d["source"] == "manual" for d in list_open["done_tasks"])

    # Undo from "اللي تم" -> back in open tasks, removed from done_tasks
    st_undo, r_undo = _http_json(f"{base}/api/tasks/{tk1['id']}", method="POST", body={"done": False})
    assert st_undo == 200 and r_undo["task"]["done"] is False
    _, list_reopened = _http_json(f"{base}/api/tasks")
    assert list_reopened["counts"]["manual_open"] == 2
    assert list_reopened["counts"]["done_total"] == 0

    # Complete again and test clear-done
    _http_json(f"{base}/api/tasks/{tk1['id']}", method="POST", body={"done": True})
    st_clr, r_clr = _http_json(f"{base}/api/tasks/clear-done", method="POST", body={})
    assert st_clr == 200 and r_clr["removed"] == 1

    st_404, _ = _http_json(f"{base}/api/tasks/nonexistent", method="POST", body={"done": True})
    assert st_404 == 404


def test_taskbar_auto_extraction_focus_cap_and_snooze_today(live_server):
    base, hub, tmp_path = live_server
    p_root = tmp_path / "Proj_AutoTasks"
    _, r = _http_json(f"{base}/api/projects", method="POST", body={"name": "Sabah Clinic", "code": "SBC", "root": str(p_root)})
    pid = r["project"]["id"]
    _make_log_with_actions(p_root / "3_Project_Logs" / "SBC_Log.xlsx")

    # Add 2 urgent manual tasks + 1 normal
    _http_json(f"{base}/api/tasks", method="POST", body={"title": "مراجعة خطة الصب #عاجل", "project_id": pid})
    _http_json(f"{base}/api/tasks", method="POST", body={"title": "اتصال بالاستشاري #عاجل", "project_id": pid})

    st, res = _http_json(f"{base}/api/tasks?project_id={pid}")
    assert st == 200 and res["ok"] is True

    kinds = {a["kind"] for a in res["auto_tasks"]}
    assert {"rejected", "unreplied", "overdue"}.issubset(kinds)
    assert len(res["today_focus"]) <= 5  # Todoist 3-5 Today focus rule
    assert res["counts"]["total_open"] >= 5

    # Snooze one auto task with metadata -> disappears from auto_tasks AND archives into done_tasks ("اللي تم")!
    rej_task = next(a for a in res["auto_tasks"] if a["kind"] == "rejected")
    st_sn, _ = _http_json(f"{base}/api/tasks/auto-dismiss", method="POST", body={"auto_key": rej_task["auto_key"], "meta": rej_task})
    assert st_sn == 200

    _, res_after = _http_json(f"{base}/api/tasks?project_id={pid}")
    assert all(a["auto_key"] != rej_task["auto_key"] for a in res_after["auto_tasks"])
    assert any(d.get("auto_key") == rej_task["auto_key"] and d.get("source") == "auto" for d in res_after["done_tasks"])

    # Undo/Restore archived auto task -> returns cleanly to auto_tasks & leaves done_tasks!
    _http_json(f"{base}/api/tasks/auto-dismiss", method="POST", body={"auto_key": rej_task["auto_key"], "restore": True})
    _, res_restored = _http_json(f"{base}/api/tasks?project_id={pid}")
    assert any(a["auto_key"] == rej_task["auto_key"] for a in res_restored["auto_tasks"])
    assert all(d.get("auto_key") != rej_task["auto_key"] for d in res_restored["done_tasks"])


def test_v10_13_taskbar_ui_symbols():
    ui_txt = (Path(__file__).resolve().parent.parent / "doc_tools" / "dashboard_ui.html").read_text(encoding="utf-8")
    for sym in (
        "smartTaskbar",
        "taskbarRestoreFab",
        "taskPeekDrawer",
        "cmdPaletteBackdrop",
        "pinToTray",
        "pinCurrentContext",
        "toggleTaskPeek",
        "mountTaskPeekShell",
        "_syncTaskDraftFromDom",
        "_isUserTypingOrPicking",
        "_renderDoneTaskCard",
        "restoreAutoTask",
        "clearDoneManualTasks",
        "tpBodyHost",
        "taskDraft",
        "openCommandPalette",
        "sla_pinned_tray",
        "sla_taskbar_hidden",
        "body.has-taskbar #toast",
    ):
        assert sym in ui_txt, f"Missing v10.13.x taskbar UI symbol: {sym}"
