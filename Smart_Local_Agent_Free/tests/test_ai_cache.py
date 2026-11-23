#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for system-wide persistent AI cache & 2-page PDF Smart Split support."""

from __future__ import annotations

import io
import json
import sys
import time
import types
from pathlib import Path

import openpyxl
from PIL import Image

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import ai_cache  # noqa: E402
import inbox_pipeline as ip  # noqa: E402
import project_dashboard as pdash  # noqa: E402
import project_organizer as porg  # noqa: E402
import scan_watcher as sw  # noqa: E402
import whatsapp_archiver as waa  # noqa: E402


def _make_fake_fitz(monkeypatch):
    try:
        import fitz  # type: ignore
        return fitz
    except ImportError:
        class _FakeRect:
            x0, y0, x1, y1, width, height = 0.0, 0.0, 595.0, 842.0, 595.0, 842.0

        class _FakePixmap:
            def __init__(self):
                im = Image.new("RGB", (32, 44), (255, 255, 255))
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
                    import re
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
        monkeypatch.setattr(porg, "fitz", fake_fitz)
        return fake_fitz


def test_ai_cache_core_lru_ttl_and_invalid_skip(tmp_path: Path):
    cdir = tmp_path / "_ai_cache"

    # Rejects empty, error, and blank/unreadable payloads
    assert ai_cache.set_cached(cdir, "test_ns", "k0", {}) is False
    assert ai_cache.set_cached(cdir, "test_ns", "k1", {"error": "429"}) is False
    assert ai_cache.set_cached(cdir, "test_ns", "k2", {"is_blank_or_unreadable": True}) is False

    # Stores & retrieves valid entries
    assert ai_cache.set_cached(cdir, "test_ns", "k_ok", {"subject": "Test"}, max_entries=10) is True
    got = ai_cache.get_cached(cdir, "test_ns", "k_ok")
    assert got is not None and got["subject"] == "Test"

    # TTL expiration
    assert ai_cache.get_cached(cdir, "test_ns", "k_ok", ttl_seconds=60) is not None
    time.sleep(0.05)
    assert ai_cache.get_cached(cdir, "test_ns", "k_ok", ttl_seconds=0.01) is None

    # LRU pruning at max_entries=10
    for i in range(15):
        ai_cache.set_cached(cdir, "lru_ns", f"k_{i:02d}", {"n": i}, max_entries=10)
        time.sleep(0.002)
    st = ai_cache.cache_stats(cdir)
    assert st["namespaces"]["lru_ns"] == 10
    assert ai_cache.get_cached(cdir, "lru_ns", "k_00") is None
    assert ai_cache.get_cached(cdir, "lru_ns", "k_14")["n"] == 14

    assert ai_cache.delete_cached(cdir, "lru_ns", "k_14") is True
    assert ai_cache.clear_cache(cdir, "lru_ns") == 9


def test_scanner_2page_multipage_and_split_cache_reuse(tmp_path: Path, monkeypatch):
    fitz = _make_fake_fitz(monkeypatch)
    reg = tmp_path / "_hub" / "registry.json"
    proot = tmp_path / "Projects"
    hub = pdash.HubState(reg, proot)
    proj = hub.create_project({"name": "TwoPage Cache Proj", "code": "TPC"})
    inbox_dir = Path(proj["inbox_dir"])
    inbox_dir.mkdir(parents=True, exist_ok=True)

    # Create a 2-page PDF (2 distinct letters inside one PDF)
    pdf_2p = inbox_dir / "two_letters.pdf"
    doc = fitz.open()
    p1 = doc.new_page()
    p1.insert_text((72, 72), "Ref: IN-101\nDate: 2026-11-18\nSubject: Letter One")
    p2 = doc.new_page()
    p2.insert_text((72, 72), "Ref: IN-102\nDate: 2026-11-18\nSubject: Letter Two")
    doc.save(str(pdf_2p))
    doc.close()

    scan_calls = [0]

    def _scan_ai(src, text, prompt):
        scan_calls[0] += 1
        return {"doc_type": "letter", "letter_direction": "incoming", "subject_ar": "خطاب وارد", "ref_no": "IN-101"}

    res_scan = ip.scan_inbox_to_triage(proj, all_projects=[proj], analyzer_override=_scan_ai)
    assert res_scan["added"] == 1 and scan_calls[0] == 1
    st = ip.get_scanner_triage_state(proj)
    assert st["counts"]["multipage"] == 1
    it = st["pending"][0]
    assert it["page_count"] == 2
    assert it["is_multipage"] is True

    # 1) Heuristic fallback is NOT cached so AI still runs next
    r_heur = ip.analyze_multipage_pdf_split(proj, it["id"], analyzer_override=None)
    assert r_heur["mode"] == "heuristic" and r_heur["from_cache"] is False

    # 2) First AI split call -> invokes AI once and saves to persistent cache
    split_calls = [0]

    def _split_ai(src, texts, prompt):
        split_calls[0] += 1
        return {
            "is_single_document": False,
            "documents": [
                {"start_page": 1, "end_page": 1, "title": "كتاب 1", "section": "incoming", "ref_no": "IN-101"},
                {"start_page": 2, "end_page": 2, "title": "كتاب 2", "section": "incoming", "ref_no": "IN-102"},
            ],
        }

    r1 = ip.analyze_multipage_pdf_split(proj, it["id"], analyzer_override=_split_ai)
    assert r1["mode"] == "ai" and r1["from_cache"] is False and split_calls[0] == 1

    # 3) Second AI split call (simulates closing modal & reopening) -> instant cache HIT!
    r2 = ip.analyze_multipage_pdf_split(proj, it["id"], analyzer_override=_split_ai)
    assert r2["mode"] == "ai" and r2["from_cache"] is True
    assert split_calls[0] == 1  # Zero extra AI calls!
    assert len(r2["documents"]) == 2

    # 4) Explicit force_refresh=True -> re-runs AI
    r3 = ip.analyze_multipage_pdf_split(proj, it["id"], analyzer_override=_split_ai, force_refresh=True)
    assert r3["mode"] == "ai" and r3["from_cache"] is False and split_calls[0] == 2


def test_doc_classify_cache_shared_between_scanner_and_inbox(tmp_path: Path):
    reg = tmp_path / "_hub" / "registry.json"
    proot = tmp_path / "Projects"
    hub = pdash.HubState(reg, proot)
    proj = hub.create_project({"name": "Shared Classify Cache", "code": "SCC"})
    inbox_dir = Path(proj["inbox_dir"])
    inbox_dir.mkdir(parents=True, exist_ok=True)

    f1 = inbox_dir / "sample_doc.txt"
    f1.write_text("مشروع حولي\nاعتماد مواد كهرباء رقم EL-045 بتاريخ 2026-11-18", encoding="utf-8")

    calls = [0]

    def _clf(src, text, prompt):
        calls[0] += 1
        return {
            "doc_type": "submittal",
            "discipline": "EL",
            "ref_no": "EL-045",
            "subject": "اعتماد مواد كهرباء",
            "confidence": 0.95,
            "is_project_related": True,
            "relevance_score": 0.95,
        }

    r_sc = ip.scan_inbox_to_triage(proj, all_projects=[proj], analyzer_override=_clf)
    assert r_sc["added"] == 1 and calls[0] == 1

    # Place an identical copy under a new filename in a subfolder for process_project_inbox_once
    sub_f = inbox_dir / "copy_folder" / "same_content_renamed.txt"
    sub_f.parent.mkdir(parents=True, exist_ok=True)
    sub_f.write_bytes(f1.read_bytes())
    # Remove hash guard from triage temporarily so process_project_inbox_once classifies the duplicate copy
    proj_root = Path(proj["root"])
    tp = proj_root / "_System_Agent" / "scanner_triage.json"
    tp.unlink()

    r_in = ip.process_project_inbox_once(
        proj,
        all_projects=[proj],
        global_settings={"engine": "gemini"},
        decision_store=hub.decisions,
        device_id=hub.device_id,
        analyzer_override=_clf,
    )
    assert r_in["processed"] >= 1
    assert calls[0] == 1  # Reused cached classification with 0 extra AI calls!


def test_organizer_and_whatsapp_ai_cache_persist_across_rescans(tmp_path: Path):
    reg = tmp_path / "_hub" / "registry.json"
    proot = tmp_path / "Projects"
    hub = pdash.HubState(reg, proot)
    proj = hub.create_project({"name": "Org WA Cache", "code": "OWC"})

    # 1) Organizer cache across re-scan
    org_src = tmp_path / "MessyFolder"
    org_src.mkdir(parents=True)
    doc_p = org_src / "scan_001.txt"
    doc_p.write_text("خطاب وارد بخصوص صب سقف الدور الأول", encoding="utf-8")

    porg.build_organization_plan(proj, str(org_src), extract_archives=False)
    org_calls = [0]

    def _org_ai(prompt, tmp_d, imgs, snip):
        org_calls[0] += 1
        return {
            "section_id": "incoming",
            "ref_no": "IN-777",
            "doc_date": "2026-11-18",
            "subject_ar": "اعتماد صب السقف",
            "suggested_filename": "2026-11-18_IN-777_اعتماد_صب_السقف",
        }

    r_org1 = porg.ai_rename_organizer_items(proj, scope="all_keep", analyzer_fn=_org_ai)
    assert r_org1["updated"] == 1 and org_calls[0] == 1

    # Re-scan entire folder from scratch -> cached AI rename must auto-hydrate!
    porg.build_organization_plan(proj, str(org_src), extract_archives=False)
    plan_after = porg.read_json_safe(porg.organizer_plan_path(Path(proj["root"])), {})
    it0 = plan_after["items"][0]
    assert it0["ai_renamed"] is True and it0["ai_from_cache"] is True
    assert it0["ref_no"] == "IN-777"

    # Explicit ai_rename with scope="all_keep" skips cached items cleanly too (0 extra calls)
    r_org2 = porg.ai_rename_organizer_items(proj, scope="all_keep", analyzer_fn=_org_ai)
    assert r_org2["total_targets"] == 0 and org_calls[0] == 1

    # 2) WhatsApp cache across re-scan
    wa_dir = tmp_path / "WA_Export"
    wa_dir.mkdir(parents=True)
    (wa_dir / "_chat.txt").write_text("[18/11/2026, 09:00] Ahmad: صباح الخير\n", encoding="utf-8")
    img_wa = wa_dir / "IMG-20261118-WA0001.jpg"
    Image.new("RGB", (60, 60), color=(200, 210, 220)).save(img_wa)

    waa.scan_and_match_whatsapp_source(proj, str(wa_dir))
    wa_calls = [0]

    def _wa_ai(prompt, tmp_d, imgs, snip):
        wa_calls[0] += 1
        return {"doc_date": "2026-11-18", "subject_ar": "صب_أعمدة_الدور_الأول"}

    r_wa1 = waa.ai_rename_whatsapp_items(proj, scope="date_only", analyzer_fn=_wa_ai)
    assert r_wa1["updated"] == 1 and wa_calls[0] == 1

    # Re-scan WhatsApp source from scratch -> cached AI name auto-hydrates!
    waa.scan_and_match_whatsapp_source(proj, str(wa_dir))
    wa_plan2 = waa.read_json_safe(waa.whatsapp_plan_path(Path(proj["root"])), {})
    wa_it = wa_plan2["items"][0]
    assert wa_it["ai_renamed"] is True and wa_it["ai_from_cache"] is True
    assert "صب_أعمدة_الدور_الأول" in wa_it["new_name"]
    assert wa_calls[0] == 1


def test_template_suggest_cache_and_gemini_runtime_cache(tmp_path: Path, monkeypatch):
    # Isolate Gemini key & runtime cache
    kf = tmp_path / ".gemini" / "ai_key.json"
    monkeypatch.setattr(sw, "AI_KEY_FILE", kf)
    sw.clear_gemini_key_quota_state()
    sw._CACHED_MODELS.clear()

    sw.save_gemini_key("AIzaSyRuntimeCacheTestKey1234567890")
    sw.mark_gemini_key_quota_exhausted("AIzaSyRuntimeCacheTestKey1234567890")
    # Clear in-memory dict to simulate server restart -> must persist on disk!
    sw._QUOTA_EXHAUSTED_KEYS.clear()
    assert sw._is_key_in_cooldown("AIzaSyRuntimeCacheTestKey1234567890") is True
    sw.clear_gemini_key_quota_state()
    assert sw._is_key_in_cooldown("AIzaSyRuntimeCacheTestKey1234567890") is False

    # Template AI suggestion cache
    reg = tmp_path / "_hub" / "registry.json"
    proot = tmp_path / "Projects"
    hub = pdash.HubState(reg, proot)
    proj = hub.create_project({"name": "Tpl Cache Proj", "code": "TCP"})
    proj_root = Path(proj["root"])
    tpl_dir = proj_root / "4_Project_Templates"
    tpl_dir.mkdir(parents=True, exist_ok=True)
    tpl_p = tpl_dir / "MAR.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Ref No"
    ws["B1"] = ""
    wb.save(tpl_p)
    wb.close()

    tpl_calls = [0]

    def _fake_gem(prompt, job_dir, model="", images=None, api_key="", backup_api_key=""):
        tpl_calls[0] += 1
        return {"suggested_section": "submittal_mar", "fields": [{"label": "المرجع", "key": "ref_no", "role": "auto_system", "input_type": "text", "target_hint": "B1"}]}

    monkeypatch.setattr(sw, "analyze_with_gemini", _fake_gem)

    r1 = pdash.suggest_fields_with_ai(proj_root, "MAR.xlsx")
    assert r1["ai_used"] is True and r1["from_cache"] is False and tpl_calls[0] == 1

    r2 = pdash.suggest_fields_with_ai(proj_root, "MAR.xlsx")
    assert r2["ai_used"] is True and r2["from_cache"] is True and tpl_calls[0] == 1

    r3 = pdash.suggest_fields_with_ai(proj_root, "MAR.xlsx", force_refresh=True)
    assert r3["ai_used"] is True and r3["from_cache"] is False and tpl_calls[0] == 2
