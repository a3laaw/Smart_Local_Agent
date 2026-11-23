from __future__ import annotations

import io
import json
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

import openpyxl
import pytest

DOC_TOOLS = Path(__file__).resolve().parent.parent / "doc_tools"
if str(DOC_TOOLS) not in sys.path:
    sys.path.insert(0, str(DOC_TOOLS))

import project_dashboard as pdash
import whatsapp_archiver as waa
from security_utils import PathTraversalError


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


def test_parse_ios_and_android_mixed_chats():
    ios_txt = (
        "[١٨‏/١١‏/٢٠٢٦، ١٠:١٤:٠٥ ص] م. أحمد: صب أعمدة الدور الثاني\n"
        "[١٨‏/١١‏/٢٠٢٦، ١٠:١٤:٢٢ ص] م. أحمد: \u200e<attached: 00000001-PHOTO-2026-11-18-10-14-22.jpg>\n"
        "[18/11/2026, 2:30:10 PM] Sara: \u200e<attached: 00000002-VIDEO-2026-11-18-14-30-10.mp4>\n"
        "[18/11/2026, 2:30:45 PM] Sara: اختبار ضغط الشبكة الدور الأول\n"
    )
    parsed = waa.parse_whatsapp_chat_text(ios_txt)
    assert parsed["date_order"] == "DMY"
    assert parsed["message_count"] == 4
    m0, m1, m2, m3 = parsed["messages"]
    assert m0["date_iso"] == "2026-11-18"
    assert m0["minute_key"] == "2026-11-18 10:14"
    assert m1["attachments"] == ["00000001-PHOTO-2026-11-18-10-14-22.jpg"]
    assert m2["date_iso"] == "2026-11-18"
    assert m2["minute_key"] == "2026-11-18 14:30"
    assert m3["clean_text"] == "اختبار ضغط الشبكة الدور الأول"


def test_whatsapp_two_conditions_separation_before_after_same_minute_and_in_place_rename(tmp_path: Path):
    proj_root = tmp_path / "Proj_WA"
    wa_dir = tmp_path / "WA_Export_Folder"
    wa_dir.mkdir(parents=True)

    # Create sample media files
    (wa_dir / "IMG-20261118-WA0001.jpg").write_bytes(b"\xff\xd8\xff\xe0img1")
    (wa_dir / "IMG-20261118-WA0002.jpg").write_bytes(b"\xff\xd8\xff\xe0img2")
    (wa_dir / "VID-20261118-WA0003.mp4").write_bytes(b"\x00\x00\x00\x18ftypmp42video3")
    (wa_dir / "IMG-20261118-WA0004.jpg").write_bytes(b"\xff\xd8\xff\xe0img4")
    (wa_dir / "DOC-20261117-WA0005.pdf").write_bytes(b"%PDF-1.4\n1 0 obj\n<<>>\nendobj\ntrailer\n<<>>\n%%EOF")
    # Exact duplicate of img1
    (wa_dir / "IMG-20261118-WA0006.jpg").write_bytes(b"\xff\xd8\xff\xe0img1")

    chat_txt = (
        "18/11/2026, 09:00 - Messages and calls are end-to-end encrypted.\n"
        # Same minute (10:14) text BEFORE two photos by Ahmed
        "18/11/2026, 10:14 - م. أحمد: صب قواعد مبنى الإدارة\n"
        "18/11/2026, 10:14 - م. أحمد: IMG-20261118-WA0001.jpg (file attached)\n"
        "18/11/2026, 10:14 - م. أحمد: IMG-20261118-WA0002.jpg (file attached)\n"
        # Different sender in same minute must NOT leak into Ahmed's photos under same_sender
        "18/11/2026, 10:14 - شخص آخر: كلام غير مرتبط\n"
        # Same minute (11:20) video first, text AFTER video by Khaled
        "18/11/2026, 11:20 - م. خالد: VID-20261118-WA0003.mp4 (file attached)\n"
        "18/11/2026, 11:20 - م. خالد: اختبار تكييف الدور الأرضي\n"
        # Photo with NO text in same minute (12:00), text is 1 minute later (12:01)
        "18/11/2026, 12:00 - م. سارة: IMG-20261118-WA0004.jpg (file attached)\n"
        "18/11/2026, 12:01 - م. سارة: عزل السطح النهائي\n"
        # Previous day document with caption
        "17/11/2026, 16:05 - م. أحمد: DOC-20261117-WA0005.pdf (file attached)\nاعتماد توريد حديد التسليح\n"
    )
    (wa_dir / "_chat.txt").write_text(chat_txt, encoding="utf-8")

    proj = {"id": "p1", "name": "WA Test", "code": "WAT", "root": str(proj_root)}
    scan_res = waa.scan_and_match_whatsapp_source(proj, wa_dir, window_minutes=0, sender_mode="same_sender")
    assert scan_res["ok"]
    cnts = scan_res["counts"]
    assert cnts["total"] == 6
    # matched: WA0001, WA0002, WA0003, WA0005, plus WA0006 (duplicate, unmatched in chat) -> 4 matched, 2 date_only
    assert cnts["matched"] == 4
    assert cnts["date_only"] == 2
    assert cnts["duplicates"] == 1

    pg = waa.get_whatsapp_plan_page(proj_root, per_page=50)
    by_fn = {it["file_name"]: it for it in pg["items"]}

    # Burst photos share base name and get _01, _02 cleanly
    assert by_fn["IMG-20261118-WA0001.jpg"]["match_status"] == "matched"
    assert by_fn["IMG-20261118-WA0001.jpg"]["text_before"] == "صب قواعد مبنى الإدارة"
    assert by_fn["IMG-20261118-WA0001.jpg"]["new_name"] == "2026-11-18_صب_قواعد_مبنى_الإدارة_01.jpg"
    assert by_fn["IMG-20261118-WA0002.jpg"]["new_name"] == "2026-11-18_صب_قواعد_مبنى_الإدارة_02.jpg"

    # Video matched by text AFTER in same minute
    v3 = by_fn["VID-20261118-WA0003.mp4"]
    assert v3["match_status"] == "matched"
    assert v3["text_source"] == "after"
    assert v3["text_after"] == "اختبار تكييف الدور الأرضي"
    assert v3["new_name"] == "2026-11-18_اختبار_تكييف_الدور_الأرضي.mp4"

    # WA0004 separated as date_only under exact same minute (12:00 vs 12:01)
    assert by_fn["IMG-20261118-WA0004.jpg"]["match_status"] == "date_only"
    assert by_fn["IMG-20261118-WA0004.jpg"]["doc_date"] == "2026-11-18"
    assert by_fn["IMG-20261118-WA0004.jpg"]["chosen_text"] == ""

    # Document matched with inline caption
    d5 = by_fn["DOC-20261117-WA0005.pdf"]
    assert d5["match_status"] == "matched"
    assert d5["text_source"] == "caption"
    assert d5["new_name"] == "2026-11-17_اعتماد_توريد_حديد_التسليح.pdf"

    # Duplicate excluded by default
    assert by_fn["IMG-20261118-WA0006.jpg"]["is_duplicate"] is True
    assert by_fn["IMG-20261118-WA0006.jpg"]["include"] is False

    # Recompute with +-1 minute picks up 12:01 text for WA0004!
    rec = waa.recompute_whatsapp_plan(proj, window_minutes=1)
    assert rec["counts"]["matched"] == 5
    pg2 = waa.get_whatsapp_plan_page(proj_root, per_page=50)
    by_fn2 = {it["file_name"]: it for it in pg2["items"]}
    assert by_fn2["IMG-20261118-WA0004.jpg"]["match_status"] == "matched"
    assert by_fn2["IMG-20261118-WA0004.jpg"]["new_name"] == "2026-11-18_عزل_السطح_النهائي.jpg"

    # Date filter: isolate only 2026-11-18 — hidden from 'all' and scoped KPIs, shown in 'out_of_range'
    rec_d = waa.recompute_whatsapp_plan(proj, date_from="2026-11-18", date_to="2026-11-18")
    assert rec_d["counts"]["out_of_range"] == 1
    assert rec_d["counts"]["total"] == 5
    assert rec_d["counts"]["total_all"] == 6
    pg3 = waa.get_whatsapp_plan_page(proj_root, per_page=50, status_filter="all")
    assert len(pg3["items"]) == 5
    assert "DOC-20261117-WA0005.pdf" not in {it["file_name"] for it in pg3["items"]}

    pg3_out = waa.get_whatsapp_plan_page(proj_root, per_page=50, status_filter="out_of_range")
    assert len(pg3_out["items"]) == 1
    assert pg3_out["items"][0]["file_name"] == "DOC-20261117-WA0005.pdf"
    assert pg3_out["items"][0]["in_date_range"] is False
    assert pg3_out["items"][0]["include"] is False

    # Restore full date range and execute rename_in_place
    waa.recompute_whatsapp_plan(proj, date_from="", date_to="")
    ex = waa.execute_whatsapp_archive_plan(proj, mode="rename_in_place")
    assert ex["ok"] and ex["renamed"] == 5 and ex.get("cleared") is True
    pg_after = waa.get_whatsapp_plan_page(proj_root)
    assert pg_after["has_plan"] is False and pg_after["total"] == 0 and pg_after["has_undo"] is True
    assert (wa_dir / "2026-11-18_صب_قواعد_مبنى_الإدارة_01.jpg").exists()
    assert (wa_dir / "2026-11-18_صب_قواعد_مبنى_الإدارة_02.jpg").exists()
    assert (wa_dir / "2026-11-18_اختبار_تكييف_الدور_الأرضي.mp4").exists()
    assert (wa_dir / "2026-11-18_عزل_السطح_النهائي.jpg").exists()
    assert (wa_dir / "2026-11-17_اعتماد_توريد_حديد_التسليح.pdf").exists()
    # Duplicate left untouched in place (zero deletion!)
    assert (wa_dir / "IMG-20261118-WA0006.jpg").exists()

    # Check Excel report generated with both sheets
    rep_p = Path(ex["report_local"])
    assert rep_p.exists()
    wb = openpyxl.load_workbook(rep_p)
    assert "أرشيف الواتساب" in wb.sheetnames
    assert "سجل المحادثة" in wb.sheetnames

    # Undo restores original filenames cleanly
    und = waa.undo_last_whatsapp_rename(proj)
    assert und["ok"] and und["restored"] == 5
    assert (wa_dir / "IMG-20261118-WA0001.jpg").exists()
    assert (wa_dir / "VID-20261118-WA0003.mp4").exists()


def test_whatsapp_ai_rename_on_date_only_items(tmp_path: Path):
    proj_root = tmp_path / "Proj_WA_AI"
    wa_dir = tmp_path / "WA_NoText"
    wa_dir.mkdir(parents=True)
    img_p = wa_dir / "IMG-20261118-WA0010.jpg"
    img_p.write_bytes(b"\xff\xd8\xff\xe0fakejpeg")

    proj = {"id": "p_ai", "name": "WA AI", "code": "WAAI", "root": str(proj_root)}
    waa.scan_and_match_whatsapp_source(proj, wa_dir)
    pg = waa.get_whatsapp_plan_page(proj_root)
    assert pg["items"][0]["match_status"] == "date_only"

    def _fake_ai(prompt, tmp_d, images, snippet):
        return {"doc_date": "2026-11-18", "subject_ar": "تشطيب واجهة الحجر"}

    ai_res = waa.ai_rename_whatsapp_items(proj, scope="date_only", analyzer_fn=_fake_ai)
    assert ai_res["ok"] and ai_res["updated"] == 1
    pg2 = waa.get_whatsapp_plan_page(proj_root)
    it = pg2["items"][0]
    assert it["ai_renamed"] is True
    assert it["match_status"] == "matched"
    assert it["new_name"] == "2026-11-18_تشطيب_واجهة_الحجر.jpg"


def test_zip_slip_blocked_on_whatsapp_extract(tmp_path: Path):
    zpath = tmp_path / "evil.zip"
    with zipfile.ZipFile(zpath, "w") as zf:
        zf.writestr("../../outside.txt", "pwn")
    with pytest.raises(PathTraversalError):
        waa.extract_zip_safely(zpath, tmp_path / "out")


def test_whatsapp_archive_full_http_streaming_zip_preview_range_and_download(live_server):
    base, hub, tmp_path = live_server
    st, cp = _http_json(f"{base}/api/projects", method="POST", body={"name": "WA API Proj", "code": "WAPI"})
    assert st == 201
    pid = cp["project"]["id"]

    # Build a real WhatsApp export ZIP in memory
    zbuf = io.BytesIO()
    chat = (
        "18/11/2026, 15:10 - م. طارق: استلام نجارة الميدة\n"
        "18/11/2026, 15:10 - م. طارق: IMG-20261118-WA0021.jpg (file attached)\n"
        "18/11/2026, 15:12 - م. طارق: VID-20261118-WA0022.mp4 (file attached)\n"
        "18/11/2026, 15:12 - م. طارق: تجربة تشغيل المضخة\n"
    )
    with zipfile.ZipFile(zbuf, "w") as zf:
        zf.writestr("_chat.txt", chat.encode("utf-8"))
        zf.writestr("IMG-20261118-WA0021.jpg", b"\xff\xd8\xff\xe0JPGDATA_123456")
        zf.writestr("VID-20261118-WA0022.mp4", b"0123456789ABCDEF")
    zip_bytes = zbuf.getvalue()

    # Path traversal header blocked cleanly (403)
    bad_req = urllib.request.Request(
        f"{base}/api/projects/{pid}/whatsapp-archive/upload-file",
        data=b"abc",
        method="POST",
        headers={
            "Content-Type": "application/octet-stream",
            "X-Upload-Batch": "b1",
            "X-Upload-Filename": urllib.parse.quote("../evil.txt"),
            "X-Upload-Relpath": urllib.parse.quote("../evil.txt"),
        },
    )
    with pytest.raises(urllib.error.HTTPError) as exc_i:
        urllib.request.urlopen(bad_req, timeout=10)
    assert exc_i.value.code == 403

    # Upload ZIP via raw binary streaming endpoint
    req_up = urllib.request.Request(
        f"{base}/api/projects/{pid}/whatsapp-archive/upload-file",
        data=zip_bytes,
        method="POST",
        headers={
            "Content-Type": "application/octet-stream",
            "X-Upload-Batch": "batch_site1",
            "X-Upload-Filename": urllib.parse.quote("WhatsApp Chat - Site1.zip"),
            "X-Upload-Relpath": urllib.parse.quote("WhatsApp Chat - Site1.zip"),
        },
    )
    with urllib.request.urlopen(req_up, timeout=15) as r_up:
        up_j = json.loads(r_up.read().decode("utf-8"))
    assert up_j["ok"] and up_j["is_zip"] is True

    # Scan the uploaded batch
    st_sc, sc_j = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/scan",
        method="POST",
        body={"batch_id": "batch_site1", "window_minutes": 0, "sender_mode": "same_sender"},
    )
    assert st_sc == 200 and sc_j["ok"]
    assert sc_j["counts"]["matched"] == 2

    # Fetch plan page + test inline edit + Range request on video
    st_pg, pg_j = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert st_pg == 200 and pg_j["total"] == 2
    vid_it = next(it for it in pg_j["items"] if it["media_kind"] == "video")
    assert vid_it["new_name"] == "2026-11-18_تجربة_تشغيل_المضخة.mp4"

    rng_req = urllib.request.Request(
        f"{base}/api/projects/{pid}/whatsapp-archive/file?id={vid_it['id']}",
        headers={"Range": "bytes=0-5"},
    )
    with urllib.request.urlopen(rng_req, timeout=10) as r_rng:
        assert r_rng.status == 206
        assert r_rng.read() == b"012345"

    # Inline rename update & execute in place
    img_it = next(it for it in pg_j["items"] if it["media_kind"] == "image")
    st_u, u_j = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/update",
        method="POST",
        body={"updates": [{"id": img_it["id"], "new_name": "2026-11-18_نجارة_مبنى_A.jpg"}]},
    )
    assert st_u == 200 and u_j["ok"]

    st_ex, ex_j = _http_json(
        f"{base}/api/projects/{pid}/whatsapp-archive/execute",
        method="POST",
        body={"mode": "rename_in_place"},
    )
    assert st_ex == 200 and ex_j["ok"] and ex_j["renamed"] == 2 and ex_j.get("cleared") is True

    # Screen/plan cleared after execution while summary & ZIP export remain available
    st_pg_clr, pg_clr = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert st_pg_clr == 200 and pg_clr["has_plan"] is False and pg_clr["total"] == 0
    assert pg_clr["last_execution"]["cleared"] is True and pg_clr["has_undo"] is True

    # Download final ZIP bundle with renamed files + Excel report
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/whatsapp-archive/download-zip", timeout=15) as r_z:
        assert r_z.status == 200
        out_zip_bytes = r_z.read()
    with zipfile.ZipFile(io.BytesIO(out_zip_bytes), "r") as zout:
        names = zout.namelist()
        assert "2026-11-18_نجارة_مبنى_A.jpg" in names
        assert "2026-11-18_تجربة_تشغيل_المضخة.mp4" in names
        assert any(n.endswith(".xlsx") for n in names)

    # Clear endpoint resets the last_execution banner as well
    st_cl, cl_j = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive/clear", method="POST", body={})
    assert st_cl == 200 and cl_j["ok"]
    _, pg_cl2 = _http_json(f"{base}/api/projects/{pid}/whatsapp-archive")
    assert pg_cl2["has_plan"] is False and pg_cl2["last_execution"] is None

    # Verify UI HTML includes WhatsApp Archive navigation & view
    with urllib.request.urlopen(f"{base}/", timeout=10) as r_ui:
        html_ui = r_ui.read().decode("utf-8")
    assert "nav.whatsapp_archive" in html_ui
    assert "renderWhatsappView" in html_ui


def test_whatsapp_filename_not_treated_as_caption_and_no_redundant_date(tmp_path: Path):
    proj_root = tmp_path / "Proj_Redundant"
    wa_dir = tmp_path / "WA_Docs"
    wa_dir.mkdir(parents=True)
    fn_june = "01-06-2026 - التقرير اليومي -.pdf"
    fn_may = "11-05-2026 - التقرير اليومي -.pdf"
    (wa_dir / fn_june).write_bytes(b"%PDF-1.4\njune")
    (wa_dir / fn_may).write_bytes(b"%PDF-1.4\nmay")
    (wa_dir / "_chat.txt").write_text(
        f"[11/05/2026, 08:15:00] a3laa wahib: \u200e<attached: {fn_may}>\n"
        f"[01/06/2026, 09:32:00] a3laa wahib: \u200e<attached: {fn_june}>\n{fn_june} • 4 صفحات\n",
        encoding="utf-8",
    )
    proj = {"id": "p_r", "name": "Redundant", "code": "RED", "root": str(proj_root)}
    waa.scan_and_match_whatsapp_source(proj, wa_dir)
    pg = waa.get_whatsapp_plan_page(proj_root)
    assert len(pg["items"]) == 2
    # Chronological order: May comes before June even though "01-..." sorts before "11-..." alphabetically
    assert pg["items"][0]["file_name"] == fn_may
    assert pg["items"][1]["file_name"] == fn_june
    it_june = pg["items"][1]
    assert it_june["caption"] == ""
    assert it_june["match_status"] == "date_only"
    assert it_june["new_name"] == "2026-06-01_09-32_التقرير_اليومي.pdf"

    # Empty-in-range date filter keeps has_plan=True and exposes out_of_range count
    waa.recompute_whatsapp_plan(proj, date_from="2026-11-01", date_to="2026-11-21")
    pg_empty = waa.get_whatsapp_plan_page(proj_root, status_filter="all")
    assert pg_empty["has_plan"] is True
    assert pg_empty["total"] == 0
    assert pg_empty["total_all"] == 2
    assert pg_empty["counts"]["total"] == 0
    assert pg_empty["counts"]["out_of_range"] == 2
    assert pg_empty["items"] == []


def test_whatsapp_album_shared_caption_cross_minute_seconds_and_arabic_formats(tmp_path: Path):
    proj_root = tmp_path / "Proj_Album"
    wa_dir = tmp_path / "WA_Album"
    wa_dir.mkdir(parents=True)

    # 3 album photos, 1 cross-minute photo+text, 1 bare <مرفق> photo, 1 unreferenced timestamped photo
    files = [
        "00000101-PHOTO-2026-06-01-09-10-10.jpg",
        "00000102-PHOTO-2026-06-01-09-10-10.jpg",
        "00000103-PHOTO-2026-06-01-09-10-11.jpg",
        "00000104-PHOTO-2026-06-01-09-32-50.jpg",
        "PHOTO-2026-06-01-10-05-00.jpg",  # bare <مرفق> & stripped prefix in folder
        "PHOTO-2026-06-01-11-15-00.jpg",  # not listed in _chat.txt, matched by filename timestamp
    ]
    for f in files:
        (wa_dir / f).write_bytes(b"\xff\xd8\xff\xe0" + f.encode("utf-8"))

    chat_txt = (
        # Album of 3 photos where only the 3rd carries the caption below
        "[01/06/2026, 9:10:10 صباحاً] علاء: \u200e<مرفق: 00000101-PHOTO-2026-06-01-09-10-10.jpg>\n"
        "[01/06/2026, 9:10:10 صباحاً] علاء: \u200e<مرفق: 00000102-PHOTO-2026-06-01-09-10-10.jpg>\n"
        "[01/06/2026, 9:10:11 صباحاً] علاء: \u200e<مرفق: 00000103-PHOTO-2026-06-01-09-10-11.jpg>\n"
        "صب خرسانة السقف\n"
        # Photo at 09:32:50 and text below at 09:33:12 (22 seconds later across minute boundary)
        "[01/06/2026, 9:32:50 صباحاً] علاء: \u200e<مرفق: 00000104-PHOTO-2026-06-01-09-32-50.jpg>\n"
        "[01/06/2026, 9:33:12 صباحاً] علاء: تركيب دكتات التكييف\n"
        # Bare <مرفق> without colon + prefix 00000105- in chat, stripped in folder + text above
        "[01/06/2026, 10:05:00 صباحاً] علاء: اختبار ضغط المياه\n"
        "[01/06/2026, 10:05:00 صباحاً] علاء: \u200e00000105-PHOTO-2026-06-01-10-05-00.jpg \u200e<مرفق>\n"
        # Timestamped file matched by time fallback with text written below it
        "[01/06/2026, 11:15:10 صباحاً] علاء: استلام أعمال البلاط\n"
    )
    (wa_dir / "_chat.txt").write_text(chat_txt, encoding="utf-8")

    proj = {"id": "p_alb", "name": "Album", "code": "ALB", "root": str(proj_root)}
    res = waa.scan_and_match_whatsapp_source(proj, wa_dir, window_minutes=0, sender_mode="same_sender")
    assert res["ok"]
    assert res["counts"]["matched"] == 6
    assert res["counts"]["date_only"] == 0

    pg = waa.get_whatsapp_plan_page(proj_root, per_page=50)
    by_fn = {it["file_name"]: it for it in pg["items"]}

    assert by_fn["00000101-PHOTO-2026-06-01-09-10-10.jpg"]["new_name"] == "2026-06-01_صب_خرسانة_السقف_01.jpg"
    assert by_fn["00000102-PHOTO-2026-06-01-09-10-10.jpg"]["new_name"] == "2026-06-01_صب_خرسانة_السقف_02.jpg"
    assert by_fn["00000103-PHOTO-2026-06-01-09-10-11.jpg"]["new_name"] == "2026-06-01_صب_خرسانة_السقف_03.jpg"
    assert by_fn["00000104-PHOTO-2026-06-01-09-32-50.jpg"]["new_name"] == "2026-06-01_تركيب_دكتات_التكييف.jpg"
    assert by_fn["PHOTO-2026-06-01-10-05-00.jpg"]["new_name"] == "2026-06-01_اختبار_ضغط_المياه.jpg"
    assert by_fn["PHOTO-2026-06-01-11-15-00.jpg"]["new_name"] == "2026-06-01_استلام_أعمال_البلاط.jpg"



