#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for Dual Gemini API Keys & Automatic Quota Failover (`v10.7`)."""

from __future__ import annotations

import http.server
import json
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import project_dashboard as pdash  # noqa: E402
import scan_watcher as sw  # noqa: E402


@pytest.fixture()
def isolated_ai_key_file(tmp_path: Path, monkeypatch):
    kf = tmp_path / ".gemini" / "ai_key.json"
    monkeypatch.setattr(sw, "AI_KEY_FILE", kf)
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    monkeypatch.delenv("GEMINI_BACKUP_API_KEY", raising=False)
    sw.clear_gemini_key_quota_state()
    sw._CACHED_MODELS.clear()
    yield kf
    sw.clear_gemini_key_quota_state()
    sw._CACHED_MODELS.clear()


@pytest.fixture()
def live_server(tmp_path: Path, isolated_ai_key_file):
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
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as he:
        return he.code, json.loads(he.read().decode("utf-8"))


def test_backup_gemini_key_crud_via_http(live_server):
    base, hub, tmp_path = live_server

    prim = "AIzaSyPrimaryKeyExample1234567890"
    bak = "AIzaSyBackupKeyExample0987654321"

    # 1) Save primary key only
    st, r1 = _http_json(f"{base}/api/ai-key", method="POST", body={"api_key": prim})
    assert st == 200 and r1["ok"] is True
    assert r1["has_primary"] is True
    assert r1["has_backup_key"] is False
    assert r1["active_key_slot"] == "primary"

    # 2) Save backup key separately -> primary remains intact!
    st, r2 = _http_json(f"{base}/api/ai-key", method="POST", body={"backup_api_key": bak})
    assert st == 200 and r2["ok"] is True
    assert r2["has_primary"] is True
    assert r2["has_backup_key"] is True
    assert sw.load_saved_primary_gemini_key() == prim
    assert sw.load_saved_gemini_backup_key() == bak

    # 3) Hub state reflects both keys cleanly + v10.13.2
    st, state = _http_json(f"{base}/api/state")
    assert st == 200
    assert state["version"] == "10.13.2"
    assert state["ai_status"]["has_primary"] is True
    assert state["ai_status"]["has_backup_key"] is True
    assert state["ai_status"]["masked_backup_key"].startswith("AIzaSy")

    # 4) Add multiple extra backup keys via newline-separated input (`add_backup_keys`)
    bak2 = "AIzaSySecondBackupKey2222222222222"
    bak3 = "AIzaSyThirdBackupKey33333333333333"
    st, r_multi = _http_json(
        f"{base}/api/ai-key",
        method="POST",
        body={"add_backup_keys": f"{bak2}\n{bak3}"},
    )
    assert st == 200
    assert r_multi["backup_count"] == 3
    assert len(r_multi["backup_keys"]) == 3
    assert sw.load_saved_gemini_backup_keys() == [bak, bak2, bak3]

    # 5) Promote 2nd backup key (index 1 -> bak2) to become primary
    st, r_prom = _http_json(f"{base}/api/ai-key", method="POST", body={"promote_backup_index": 1})
    assert st == 200
    assert sw.load_saved_primary_gemini_key() == bak2
    assert prim in sw.load_saved_gemini_backup_keys()

    # 6) Remove one specific backup key by index
    cur_cnt = r_prom["backup_count"]
    st, r_rem = _http_json(f"{base}/api/ai-key", method="POST", body={"remove_backup_index": 0})
    assert st == 200
    assert r_rem["backup_count"] == cur_cnt - 1

    # Restore exact state for clear-sequence check
    _http_json(f"{base}/api/ai-key", method="POST", body={"api_key": prim, "backup_api_key": bak})

    # 7) Clear primary -> backup becomes the active fallback automatically!
    st, r3 = _http_json(f"{base}/api/ai-key", method="POST", body={"clear": True})
    assert st == 200
    assert r3["has_primary"] is False
    assert r3["has_backup_key"] is True
    assert r3["has_key"] is True
    assert r3["active_key_slot"] == "backup"
    assert sw.load_saved_gemini_key() == bak

    # 8) Clear all backups -> no keys remain
    st, r4 = _http_json(f"{base}/api/ai-key", method="POST", body={"clear_backup": True})
    assert st == 200
    assert r4["has_key"] is False
    assert r4["active_key_slot"] == "none"


def test_auto_failover_to_backup_key_when_primary_quota_exhausted(isolated_ai_key_file):
    prim = "AIzaSyPrimaryQuotaExhausted9999999"
    bak1 = "AIzaSyBackupOneAlsoExhausted888888"
    bak2 = "AIzaSyBackupTwoWorkingKey111111111"
    sw.save_gemini_key(prim, backup_api_keys=[bak1, bak2])

    seen_calls: list[tuple[str, str]] = []

    class FakeGeminiHandler(http.server.BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            return

        def do_GET(self):
            key = self.headers.get("x-goog-api-key", "")
            seen_calls.append(("GET", key))
            payload = {
                "models": [
                    {"name": "models/gemini-2.5-flash-lite", "supportedGenerationMethods": ["generateContent"]},
                    {"name": "models/gemini-2.5-flash", "supportedGenerationMethods": ["generateContent"]},
                ]
            }
            raw = json.dumps(payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

        def do_POST(self):
            ln = int(self.headers.get("Content-Length") or 0)
            _ = self.rfile.read(ln)
            key = self.headers.get("x-goog-api-key", "")
            seen_calls.append(("POST", key))
            if key in (prim, bak1):
                err = {
                    "error": {
                        "code": 429,
                        "message": "You exceeded your current quota, please check your plan and billing details.",
                        "status": "RESOURCE_EXHAUSTED",
                    }
                }
                raw = json.dumps(err).encode("utf-8")
                self.send_response(429)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
                return

            ok_body = {
                "candidates": [
                    {
                        "content": {
                            "parts": [
                                {"text": json.dumps({"doc_type": "letter", "subject_ar": "كتاب اعتماد مقاول"})}
                            ]
                        }
                    }
                ]
            }
            raw = json.dumps(ok_body).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), FakeGeminiHandler)
    port = srv.server_address[1]
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        base_url = f"http://127.0.0.1:{port}"
        res1 = sw.analyze_with_gemini_api("اقرأ المستند", [], api_base=base_url)
        assert res1["doc_type"] == "letter"
        assert res1["subject_ar"] == "كتاب اعتماد مقاول"
        assert ("POST", prim) in seen_calls
        assert ("POST", bak1) in seen_calls
        assert ("POST", bak2) in seen_calls

        # Status now reports primary & backup#1 exhausted and backup#2 active!
        st_now = sw.get_gemini_keys_status()
        assert st_now["primary_quota_exhausted"] is True
        assert st_now["active_key_slot"] == "backup"
        assert st_now["active_key_index"] == 2
        assert st_now["backup_keys"][0]["exhausted"] is True
        assert st_now["backup_keys"][1]["active"] is True

        # Next request starts directly on backup#2 without wasting calls on exhausted keys!
        seen_calls.clear()
        res2 = sw.analyze_with_gemini_api("اقرأ مستنداً ثانياً", [], api_base=base_url)
        assert res2["doc_type"] == "letter"
        assert ("POST", prim) not in seen_calls
        assert ("POST", bak1) not in seen_calls
        assert ("POST", bak2) in seen_calls
    finally:
        srv.shutdown()
        srv.server_close()
