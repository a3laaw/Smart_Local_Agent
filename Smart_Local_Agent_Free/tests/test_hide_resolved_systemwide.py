#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""System-wide tests verifying that rejected/approved/resolved items hide immediately
across Scanner, Pending Decisions, Overview Inbox Sync, Project Organizer,
WhatsApp Archiver, and Content Duplicates."""

from __future__ import annotations

import json
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

from PIL import Image, ImageDraw
import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import inbox_pipeline as ip  # noqa: E402
import project_dashboard as pdash  # noqa: E402


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


def _http_json(url: str, *, method: str = "GET", body: dict | None = None) -> tuple[int, dict]:
    data = json.dumps(body).encode("utf-8") if body is not None else None
    headers = {"Content-Type": "application/json"} if body is not None else {}
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as he:
        return he.code, json.loads(he.read().decode("utf-8"))


def test_scanner_reject_and_approve_hide_and_sync_across_decisions_and_overview(live_server):
    base, hub, tmp_path = live_server
    st, p_res = _http_json(
        f"{base}/api/projects",
        method="POST",
        body={"name": "مشروع السالمية", "code": "SAL"},
    )
    assert st in (200, 201)
    pid = p_res["project"]["id"]
    proj = hub.get_project(pid)
    inbox_dir = Path(proj["inbox_dir"])
    inbox_dir.mkdir(parents=True, exist_ok=True)

    f_keep_untouched = inbox_dir / "keep_unchanged_contract.txt"
    f_keep_untouched.write_text("مشروع السالمية\nمذكرة داخلية خاصة لا تغير اسمها", encoding="utf-8")

    f_to_approve = inbox_dir / "incoming_letter_sample.txt"
    f_to_approve.write_text("مشروع السالمية\nكتاب وارد رقم IN-001 بتاريخ 2026-11-18", encoding="utf-8")

    f_from_decisions = inbox_dir / "other_file_resolved_in_decisions.txt"
    f_from_decisions.write_text("مشروع السالمية\nملف قيد المراجعة من القرارات", encoding="utf-8")

    # 1) Scan into Scanner Triage
    st, scan_res = _http_json(f"{base}/api/projects/{pid}/scanner/scan", method="POST", body={"blocking": True})
    assert st == 200
    assert len(scan_res["pending"]) == 3

    by_name = {it["file_name"]: it for it in scan_res["pending"]}
    id_reject = by_name[f_keep_untouched.name]["id"]
    id_approve = by_name[f_to_approve.name]["id"]

    # Also simulate a matching pending decision in DecisionStore for the rejected & decisions-bound files
    dec_rej = hub.decisions.enqueue_decision(
        kind="low_confidence",
        file_path=f_keep_untouched,
        source_project_id=pid,
        source_project_name=proj["name"],
        reason="needs_review",
    )
    dec_other = hub.decisions.enqueue_decision(
        kind="low_confidence",
        file_path=f_from_decisions,
        source_project_id=pid,
        source_project_name=proj["name"],
        reason="needs_review",
    )
    assert hub.decisions.pending_count() == 2

    # 2) Reject first file in Scanner -> hides from pending, stays untouched on disk, closes its pending decision
    st, rej_res = _http_json(f"{base}/api/projects/{pid}/scanner/reject", method="POST", body={"id": id_reject})
    assert st == 200
    assert f_keep_untouched.exists()  # Never modified or moved on disk
    assert all(x["id"] != id_reject for x in rej_res["pending"])
    assert any(x["id"] == id_reject for x in rej_res["rejected"])
    assert hub.decisions.pending_count() == 1  # dec_rej closed automatically!

    # 3) Approve second file in Scanner -> hides from pending and moves to archive
    st, app_res = _http_json(f"{base}/api/projects/{pid}/scanner/approve", method="POST", body={"id": id_approve})
    assert st == 200
    assert not f_to_approve.exists()
    assert all(x["id"] != id_approve for x in app_res["pending"])

    # 4) Resolve third file from Pending Decisions screen -> must vanish from Scanner pending too!
    st, dec_res = _http_json(
        f"{base}/api/decisions/{dec_other['id']}/resolve",
        method="POST",
        body={"action": "ignore_once"},
    )
    assert st == 200
    assert hub.decisions.pending_count() == 0

    st, sc_after_dec = _http_json(f"{base}/api/projects/{pid}/scanner")
    assert st == 200
    assert len(sc_after_dec["pending"]) == 0

    # 5) Running Overview organize sync must NOT re-scan or enqueue the rejected file!
    org_res = ip.process_project_inbox_once(
        proj,
        all_projects=hub.list_projects(),
        decision_store=hub.decisions,
        global_settings=hub.get_settings(),
        device_id="test",
    )
    assert org_res["ok"] is True
    assert hub.decisions.pending_count() == 0
    assert f_keep_untouched.exists()

    # 6) Clear resolved tabs from Scanner screen -> approved & rejected vanish from UI while hash stays protected
    st, clr_res = _http_json(f"{base}/api/projects/{pid}/scanner/clear-resolved", method="POST", body={"bucket": "all"})
    assert st == 200
    assert clr_res["pending"] == []
    assert clr_res["approved"] == []
    assert clr_res["rejected"] == []

    # Re-scanning folder afterwards must NOT bring the rejected file back on screen
    st, rescan = _http_json(f"{base}/api/projects/{pid}/scanner/scan", method="POST", body={"blocking": True})
    assert st == 200
    assert rescan["pending"] == []
    assert rescan["rejected"] == []


def test_organizer_whatsapp_and_duplicates_hide_rejected_and_resolved_items(live_server):
    base, hub, tmp_path = live_server
    st, p_res = _http_json(
        f"{base}/api/projects",
        method="POST",
        body={"name": "مشروع حولي", "code": "HAW"},
    )
    assert st in (200, 201)
    pid = p_res["project"]["id"]

    # ── A) Project Organizer: rejecting/excluding a 'keep' item hides it immediately from active tabs ──
    org_src = tmp_path / "org_src"
    org_src.mkdir()
    (org_src / "file_keep_1.txt").write_text("مستند 1", encoding="utf-8")
    (org_src / "file_reject_2.txt").write_text("مستند 2", encoding="utf-8")

    st, plan_res = _http_json(
        f"{base}/api/projects/{pid}/organizer/scan",
        method="POST",
        body={"folder": str(org_src), "read_text": False, "blocking": True},
    )
    assert st == 200
    st, pg1 = _http_json(f"{base}/api/projects/{pid}/organizer?filter=all")
    assert st == 200
    assert len(pg1["items"]) == 2

    target_id = next(it["id"] for it in pg1["items"] if it["file_name"] == "file_reject_2.txt")
    st, _ = _http_json(
        f"{base}/api/projects/{pid}/organizer/update",
        method="POST",
        body={"updates": [{"id": target_id, "include": False, "decision": "excluded"}]},
    )
    assert st == 200

    st, pg_all = _http_json(f"{base}/api/projects/{pid}/organizer?filter=all")
    assert st == 200
    assert [it["file_name"] for it in pg_all["items"]] == ["file_keep_1.txt"]
    assert pg_all["counts"]["excluded"] == 1

    st, pg_excl = _http_json(f"{base}/api/projects/{pid}/organizer?filter=excluded")
    assert st == 200
    assert [it["file_name"] for it in pg_excl["items"]] == ["file_reject_2.txt"]

    # ── B) WhatsApp Archiver: unchecking / rejecting a file hides it with hide_excluded=1 ──
    wa_dir = tmp_path / "wa_src"
    wa_dir.mkdir()
    (wa_dir / "_chat.txt").write_text(
        "[18/11/2026, 09:00:00] م. أحمد: مخطط السطح المعتمد\n"
        "[18/11/2026, 09:00:05] م. أحمد: \u200e<attached: IMG-20261118-WA0001.jpg>\n"
        "[18/11/2026, 09:05:00] م. أحمد: صورة موقع عامة\n"
        "[18/11/2026, 09:05:05] م. أحمد: \u200e<attached: IMG-20261118-WA0002.jpg>\n",
        encoding="utf-8",
    )
    Image.new("RGB", (60, 60), (220, 40, 40)).save(wa_dir / "IMG-20261118-WA0001.jpg")
    Image.new("RGB", (60, 60), (40, 220, 40)).save(wa_dir / "IMG-20261118-WA0002.jpg")

    st, wa_scan = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/scan",
        method="POST",
        body={"source_path": str(wa_dir)},
    )
    assert st == 200
    assert wa_scan["counts"]["total"] == 2

    st, wa_pg1 = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive?filter=all&hide_excluded=1")
    assert st == 200
    assert len(wa_pg1["items"]) == 2
    wa_rej_id = next(it["id"] for it in wa_pg1["items"] if it["file_name"] == "IMG-20261118-WA0002.jpg")

    st, _ = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/update",
        method="POST",
        body={"updates": [{"id": wa_rej_id, "include": False}]},
    )
    assert st == 200

    st, wa_pg_after = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive?filter=all&hide_excluded=1")
    assert st == 200
    assert [it["file_name"] for it in wa_pg_after["items"]] == ["IMG-20261118-WA0001.jpg"]
    assert wa_pg_after["counts"]["excluded"] == 1

    st, wa_pg_excl = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive?filter=excluded")
    assert st == 200
    assert [it["file_name"] for it in wa_pg_excl["items"]] == ["IMG-20261118-WA0002.jpg"]

    # ── C) Content Duplicates: 'keep_all_hide' hides the resolved group from active screen ──
    dup_dir = tmp_path / "dup_src"
    dup_dir.mkdir()
    im = Image.new("RGB", (120, 120), (255, 255, 255))
    dr = ImageDraw.Draw(im)
    dr.rectangle([10, 10, 110, 110], fill=(20, 80, 160))
    im.save(dup_dir / "photo_orig.png")
    im.save(dup_dir / "photo_copy.png")

    st, dup_scan = _http_json(
        f"{base}/api/projects/{pid}/duplicates/scan",
        method="POST",
        body={"source_kind": "folder", "source_path": str(dup_dir), "min_similarity": 85},
    )
    assert st == 200
    assert dup_scan["groups_count"] == 1

    st, dup_pg1 = _http_json(f"{base}/api/projects/{pid}/duplicates?filter=all&show_reviewed=0")
    assert st == 200
    assert len(dup_pg1["groups"]) == 1
    gid = dup_pg1["groups"][0]["group_id"]

    st, _ = _http_json(
        f"{base}/api/projects/{pid}/duplicates/update",
        method="POST",
        body={"group_id": gid, "action": "keep_all_hide"},
    )
    assert st == 200

    st, dup_pg_active = _http_json(f"{base}/api/projects/{pid}/duplicates?filter=all&show_reviewed=0")
    assert st == 200
    assert dup_pg_active["groups"] == []
    assert dup_pg_active["counts"]["active_groups"] == 0
    assert dup_pg_active["counts"]["reviewed_groups"] == 1

    st, dup_pg_rev = _http_json(f"{base}/api/projects/{pid}/duplicates?filter=reviewed&show_reviewed=1")
    assert st == 200
    assert len(dup_pg_rev["groups"]) == 1
    assert all(it["decision"] == "keep" for it in dup_pg_rev["groups"][0]["items"])
