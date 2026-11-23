#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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

from PIL import Image, ImageDraw
import openpyxl
import docx
import pytest

TOOLS_DIR = Path(__file__).resolve().parent.parent / "doc_tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import content_duplicates as dups  # noqa: E402
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
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as he:
        return he.code, json.loads(he.read().decode("utf-8"))


def _make_rich_image(path: Path, size: tuple[int, int], *, variant: str = "site") -> None:
    img = Image.new("RGB", size, (235, 240, 245))
    dr = ImageDraw.Draw(img)
    w, h = size
    if variant == "site":
        # Sky gradient-like top, building structure, dark base, red crane accent
        dr.rectangle([0, 0, w, int(h * 0.35)], fill=(120, 175, 230))
        dr.rectangle([int(w * 0.12), int(h * 0.25), int(w * 0.55), int(h * 0.85)], fill=(210, 195, 165))
        dr.rectangle([int(w * 0.18), int(h * 0.32), int(w * 0.30), int(h * 0.50)], fill=(45, 70, 95))
        dr.rectangle([int(w * 0.34), int(h * 0.32), int(w * 0.46), int(h * 0.50)], fill=(45, 70, 95))
        dr.rectangle([0, int(h * 0.82), w, h], fill=(75, 70, 65))
        dr.rectangle([int(w * 0.62), int(h * 0.15), int(w * 0.72), int(h * 0.82)], fill=(195, 60, 45))
    else:
        # Completely different visual layout: dark night grid with bright yellow diagonal
        img = Image.new("RGB", size, (20, 24, 34))
        dr = ImageDraw.Draw(img)
        dr.polygon([(0, h), (w, 0), (w, int(h * 0.35)), (int(w * 0.35), h)], fill=(245, 200, 35))
        dr.ellipse([int(w * 0.08), int(h * 0.08), int(w * 0.36), int(h * 0.36)], fill=(40, 180, 120))
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() in (".jpg", ".jpeg"):
        img.save(path, format="JPEG", quality=45)
    else:
        img.save(path, format="PNG")


def test_content_duplicates_visual_exact_and_doc_ignoring_filenames(tmp_path: Path):
    proj_root = tmp_path / "ProjA"
    proj_root.mkdir(parents=True)
    proj = {"id": "p1", "code": "P1", "name": "مشروع تجريبي", "root": str(proj_root)}

    src = tmp_path / "messy_folder"
    src.mkdir()

    # 1) Visual similarity across DIFFERENT filenames, formats, and resolutions
    hi_png = src / "original_blueprint_hi.png"
    lo_jpg = src / "whatsapp_recompressed_copy_999.jpg"
    other_img = src / "completely_different_night.png"
    _make_rich_image(hi_png, (800, 600), variant="site")
    _make_rich_image(lo_jpg, (400, 300), variant="site")
    _make_rich_image(other_img, (600, 450), variant="night")

    # 2) Exact byte match across DIFFERENT filenames
    raw_pdf_bytes = b"%PDF-1.4\n%fake identical binary payload 1234567890\n"
    (src / "final_signed_contract.pdf").write_bytes(raw_pdf_bytes)
    (src / "scan_0088.pdf").write_bytes(raw_pdf_bytes)

    # 3) Document text-content similarity across DIFFERENT filenames (.docx)
    d1 = docx.Document()
    d1.add_paragraph("تقرير متابعة أعمال الحفر والخرسانة المسلحة في مبنى الإدارة الرئيسي بالمشروع رقم ٤٥.")
    d1.add_paragraph("تم استلام حديد التسليح واعتماد الخلطة التصميمية حسب المواصفات الفنية المعتمدة في العقد.")
    d1.save(str(src / "concrete_report_A.docx"))

    d2 = docx.Document()
    d2.add_paragraph("تقرير متابعة أعمال الحفر والخرسانة المسلحة في مبنى الإدارة الرئيسي بالمشروع رقم 45.")
    d2.add_paragraph("تم استلام حديد التسليح واعتماد الخلطة التصميمية حسب المواصفات الفنية المعتمدة في العقد.")
    d2.add_paragraph("ملاحظة طفيفة.")
    d2.save(str(src / "draft_copy_unrelated_name.docx"))

    d3 = docx.Document()
    d3.add_paragraph("كشف حساب أجور العمالة والمعدات الشهرية الخاصة بأعمال البستنة والحدائق الخارجية.")
    d3.add_paragraph("قيمة المستخلص الشهري تشمل النقل والصيانة الدورية للآليات الثقيلة في الموقع الشرقي.")
    d3.save(str(src / "gardening_payroll.docx"))

    res = dups.scan_content_duplicates(
        proj,
        source_kind="folder",
        source_path=str(src),
        scope="all",
        min_similarity=82,
    )
    assert res["ok"] is True
    assert res["total_files_scanned"] == 8
    assert res["groups_count"] == 3

    pg = dups.get_duplicates_plan_page(proj_root)
    assert pg["has_plan"] is True
    kinds = {g["match_kind"]: g for g in pg["groups"]}
    assert "exact_content" in kinds
    assert "visual_similar" in kinds
    assert "doc_content_similar" in kinds

    # Visual group picked the high-res PNG as original and did NOT include the night image
    vis_grp = kinds["visual_similar"]
    vis_names = {it["file_name"] for it in vis_grp["items"]}
    assert vis_names == {"original_blueprint_hi.png", "whatsapp_recompressed_copy_999.jpg"}
    assert vis_grp["original_name"] == "original_blueprint_hi.png"
    assert vis_grp["similarity_pct"] >= 85

    # Doc group matched the two concrete reports by content regardless of filename
    doc_grp = kinds["doc_content_similar"]
    doc_names = {it["file_name"] for it in doc_grp["items"]}
    assert doc_names == {"concrete_report_A.docx", "draft_copy_unrelated_name.docx"}

    # Exact group matched the two identically-contented PDFs regardless of filename
    ex_grp = kinds["exact_content"]
    ex_names = {it["file_name"] for it in ex_grp["items"]}
    assert ex_names == {"final_signed_contract.pdf", "scan_0088.pdf"}


def test_duplicates_http_upload_set_original_quarantine_clear_zip_and_undo(live_server):
    base, hub, tmp_path = live_server
    st_cr, cp = _http_json(f"{base}/api/projects", method="POST", body={"name": "مشروع المكررات", "code": "DUP"})
    assert st_cr == 201
    proj = cp["project"]
    pid = proj["id"]
    proj_root = Path(proj["root"])

    # Prepare two visually identical images (different names & sizes) + one unique image
    up_tmp = tmp_path / "to_upload"
    up_tmp.mkdir()
    img_hi = up_tmp / "cam_photo_01.png"
    img_lo = up_tmp / "renamed_copy_02.jpg"
    img_unq = up_tmp / "unique_03.png"
    _make_rich_image(img_hi, (640, 480), variant="site")
    _make_rich_image(img_lo, (320, 240), variant="site")
    _make_rich_image(img_unq, (400, 300), variant="night")

    batch_id = "batch_dup_1"
    for f in (img_hi, img_lo, img_unq):
        req = urllib.request.Request(
            f"{base}/api/projects/{pid}/duplicates/upload-file",
            data=f.read_bytes(),
            method="POST",
            headers={
                "Content-Type": "application/octet-stream",
                "X-Upload-Batch": batch_id,
                "X-Upload-Filename": urllib.parse.quote(f.name),
                "X-Upload-Relpath": urllib.parse.quote(f.name),
            },
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            assert resp.status == 200 and json.loads(resp.read().decode("utf-8"))["ok"]

    st_sc, sc_j = _http_json(
        f"{base}/api/projects/{pid}/duplicates/scan",
        method="POST",
        body={"source_kind": "upload", "batch_id": batch_id, "scope": "images", "min_similarity": 85},
    )
    assert st_sc == 200 and sc_j["ok"] and sc_j["groups_count"] == 1

    # Hub state reports dup_groups_count == 1
    _, st_hub = _http_json(f"{base}/api/state")
    p_card = next(p for p in st_hub["projects"] if p["id"] == pid)
    assert p_card["dup_groups_count"] == 1

    st_pg, pg = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert st_pg == 200 and pg["has_plan"] is True
    grp = pg["groups"][0]
    assert grp["original_name"] == "cam_photo_01.png"
    copy_item = next(it for it in grp["items"] if it["file_name"] == "renamed_copy_02.jpg")
    hi_item = next(it for it in grp["items"] if it["file_name"] == "cam_photo_01.png")

    # Preview file endpoint works safely
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/duplicates/file?id={hi_item['id']}", timeout=10) as rf:
        assert rf.status == 200 and len(rf.read()) > 100

    # Permanent delete without explicit confirm_delete must be rejected
    st_del_rej, _ = _http_json(
        f"{base}/api/projects/{pid}/duplicates/execute",
        method="POST",
        body={"mode": "delete_permanent", "confirm_delete": False},
    )
    assert st_del_rej == 400

    # Switch original to copy_item, then switch back to hi_item
    st_up1, _ = _http_json(
        f"{base}/api/projects/{pid}/duplicates/update",
        method="POST",
        body={"item_id": copy_item["id"], "action": "set_original"},
    )
    assert st_up1 == 200
    _, pg_sw = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_sw["groups"][0]["original_name"] == "renamed_copy_02.jpg"
    _http_json(
        f"{base}/api/projects/{pid}/duplicates/update",
        method="POST",
        body={"item_id": hi_item["id"], "action": "set_original"},
    )
    _http_json(
        f"{base}/api/projects/{pid}/duplicates/update",
        method="POST",
        body={"bulk_action": "select_all_dups"},
    )

    # Execute quarantine -> screen clears automatically
    st_ex, ex_j = _http_json(
        f"{base}/api/projects/{pid}/duplicates/execute",
        method="POST",
        body={"mode": "quarantine"},
    )
    assert st_ex == 200 and ex_j["ok"] and ex_j["processed"] == 1 and ex_j["cleared"] is True

    _, pg_after = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_after["has_plan"] is False and pg_after["total_groups"] == 0
    assert pg_after["has_undo"] is True and pg_after["last_execution"]["cleared"] is True

    # Excel report exists & is readable
    rep_p = proj_root / ex_j["report_rel"]
    assert rep_p.exists()
    wb = openpyxl.load_workbook(rep_p, read_only=True)
    assert "تفاصيل المجموعات" in wb.sheetnames
    wb.close()

    # Unique ZIP export contains cam_photo_01.png + unique_03.png + report (without quarantined copy)
    with urllib.request.urlopen(f"{base}/api/projects/{pid}/duplicates/download-zip", timeout=15) as rz:
        assert rz.status == 200
        z_bytes = rz.read()
    with zipfile.ZipFile(io.BytesIO(z_bytes), "r") as zin:
        names = set(zin.namelist())
        assert "cam_photo_01.png" in names
        assert "unique_03.png" in names
        assert "renamed_copy_02.jpg" not in names

    # Quarantined file moved out of upload batch and restored cleanly on undo
    batch_dir = dups.dups_uploads_root(proj_root) / batch_id
    assert not (batch_dir / "renamed_copy_02.jpg").exists()

    st_un, un_j = _http_json(f"{base}/api/projects/{pid}/duplicates/undo", method="POST", body={})
    assert st_un == 200 and un_j["ok"] and un_j["restored"] == 1
    assert (batch_dir / "renamed_copy_02.jpg").exists()

    st_cl, cl_j = _http_json(f"{base}/api/projects/{pid}/duplicates/clear", method="POST", body={})
    assert st_cl == 200 and cl_j["ok"]
    _, pg_clr = _http_json(f"{base}/api/projects/{pid}/duplicates")
    assert pg_clr["last_execution"] is None
