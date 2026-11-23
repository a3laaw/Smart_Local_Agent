from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import openpyxl

DOC_TOOLS = Path(__file__).resolve().parent.parent / "doc_tools"
if str(DOC_TOOLS) not in sys.path:
    sys.path.insert(0, str(DOC_TOOLS))

import active_project as ap
from decisions import DecisionStore
import inbox_pipeline as ip
import project_log as plog
import relevance as rel


def _mk_proj(tmp_path: Path, pid: str, name: str, code: str, contract: str, **extra) -> dict:
    root = tmp_path / f"{code}_{pid}"
    ibx = Path(extra.pop("inbox_dir", root / "1_Inbox"))
    (root / "2_Archive").mkdir(parents=True, exist_ok=True)
    (root / "3_Project_Logs").mkdir(parents=True, exist_ok=True)
    (root / "4_Project_Templates").mkdir(parents=True, exist_ok=True)
    ibx.mkdir(parents=True, exist_ok=True)
    # Attach a real uploaded project log
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append(["No", "Reference", "Description", "Date", "Party", "Status", "Amount", "Link"])
    ws.append([1, "OUT-005", "كتاب افتتاحي", "2026-09-01", "المالك", "Sent", "", ""])
    lp = tmp_path / f"{code}_log.xlsx"
    wb.save(lp)
    plog.upload_project_log(root, lp.name, lp.read_bytes(), device_id="pc-test")
    lp.unlink()

    base = {
        "id": pid,
        "name": name,
        "code": code,
        "contract_no": contract,
        "owner": "وزارة الأشغال العامة",
        "consultant": "مكتب الخليج الاستشاري",
        "contractor": "شركة البناء الحديث",
        "location": "المطلاع",
        "root": str(root),
        "inbox_dir": str(ibx),
        "enabled": True,
        "aliases": [],
        "keywords": [],
        "subcontractors": ["شركة النور للكهرباء"],
        "unrelated_policy": "rename_in_place",
        "relevance_threshold": 0.6,
        "dry_run": False,
        "ignore_patterns": ["*.bak", "drafts/*"],
        "log_unrelated_sheet": False,
        "baseline_counters": {},
        "baseline_disc_counters": {},
    }
    base.update(extra)
    return base


def test_arabic_normalizer_unification():
    s1 = "إِشْعَارُ أَمْرِ شِرَاءٍ — شَرِكَةُ الخَلِيجِ لِلتِّجَارَةِ ١٤٢٥"
    s2 = "اشعار امر شراء شركه الخليج للتجاره 1425"
    assert rel.normalize_ar(s1) == rel.normalize_ar(s2)
    assert "1425" in rel.normalize_ar(s1)


def test_rename_in_place_no_rescan_and_full_undo(tmp_path: Path):
    proj = _mk_proj(tmp_path, "p1", "مشروع مدرسة المطلاع", "MT-01", "هـ ع/2026/55")
    ds = DecisionStore(tmp_path / "state", device_id="pc-a")
    gcfg = {"default_unrelated_policy": "rename_in_place", "relevance_threshold": 0.6, "process_mode": "active_only"}

    f = Path(proj["inbox_dir"]) / "IMG_0091.pdf"
    f.write_bytes(b"%PDF-1.4 personal car insurance renewal 2026")

    res = ip.handle_inbox_document(
        f,
        extracted_text="وثيقة تجديد تأمين سيارة خاصة - شركة التأمين الأهلية - رقم اللوحة 99881",
        ai_meta={
            "is_project_related": False,
            "relevance_score": 0.08,
            "detected_project_or_entity": "شركة التأمين الأهلية",
            "relevance_reason": "تأمين سيارة شخصي لا يخص المشروع",
            "subject_ar": "تجديد تأمين سيارة",
            "doc_date": "2026-11-17",
        },
        active_proj=proj,
        all_projects=[proj],
        global_settings=gcfg,
        decision_store=ds,
        device_id="pc-a",
    )
    assert res["outcome"] == "renamed_in_place"
    renamed_path = Path(proj["inbox_dir"]) / res["journal_entry"]["new_name"]
    assert renamed_path.exists()
    assert not f.exists()
    # Zero pollution of project's uploaded log!
    assert plog.read_project_log_summary(proj)["stats"]["total_rows"] == 1

    # Re-scanning the renamed file MUST be skipped via SHA-256 journal (no infinite loop!)
    res2 = ip.handle_inbox_document(
        renamed_path, extracted_text="نفس النص", ai_meta={"is_project_related": False, "relevance_score": 0.08},
        active_proj=proj, all_projects=[proj], global_settings=gcfg, decision_store=ds, device_id="pc-a",
    )
    assert res2["outcome"] == "skipped"

    # Full Undo restores exact original filename IMG_0091.pdf!
    jpath = ip.rename_journal_path_for(Path(proj["inbox_dir"]))
    entries = rel.load_rename_journal(jpath)
    assert len(entries) == 1
    rel.undo_rename_entry(jpath, entries[0]["id"], device_id="pc-a")
    assert f.exists()
    assert not renamed_path.exists()


def test_unclassified_entity_low_confidence_and_legacy_tplmap_migration(tmp_path: Path):
    import json
    import template_profile as tprof

    proj = _mk_proj(tmp_path, "pu", "مشروع الصيانة", "MN-01", "MN-77", unrelated_policy="new_project_by_name")
    ds = DecisionStore(tmp_path / "state", device_id="pc-a")
    gcfg = {"default_unrelated_policy": "new_project_by_name", "relevance_threshold": 0.6, "process_mode": "active_only"}

    # 1) Unrelated file with NO detected entity -> routed cleanly to unclassified folder & undoable!
    u_none = Path(proj["inbox_dir"]) / "mystery.pdf"
    u_none.write_bytes(b"%PDF-1.4 random brochure")
    r_un = ip.handle_inbox_document(
        u_none, extracted_text="نشرة إعلانية عامة للسفر والسياحة",
        ai_meta={"is_project_related": False, "relevance_score": 0.05, "detected_project_or_entity": ""},
        active_proj=proj, all_projects=[proj], global_settings=gcfg, decision_store=ds, device_id="pc-a",
    )
    je = r_un.get("journal_entry") or r_un.get("route") or r_un
    dest_p = Path(je["new_path"])
    assert dest_p.exists()
    rel.undo_rename_entry(ip.rename_journal_path_for(Path(proj["inbox_dir"])), je["id"], device_id="pc-a")
    assert u_none.exists()

    # 2) Low-confidence ambiguous document -> routed safely to Pending Decisions (never misfiled!)
    amb = Path(proj["inbox_dir"]) / "blurry.pdf"
    amb.write_bytes(b"%PDF-1.4 ambiguous scan")
    r_amb = ip.handle_inbox_document(
        amb, extracted_text="إيصال استلام غير واضح",
        ai_meta={"is_project_related": True, "relevance_score": 0.52, "confidence": 0.45},
        active_proj=proj, all_projects=[proj], global_settings=gcfg, decision_store=ds, device_id="pc-a",
    )
    assert r_amb["outcome"] == "pending_decision"
    assert amb.exists()

    # 3) Legacy v9 `_template_maps.json` auto-migrates seamlessly to v10 `<name>.tplmap.json`
    tdir = Path(proj["root"]) / "4_Project_Templates"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Form"
    ws["A1"] = "الموضوع:"
    ws["B1"] = ""
    tpl_p = tdir / "Old_MAR.xlsx"
    wb.save(tpl_p)
    (tdir / "_template_maps.json").write_text(json.dumps({
        "custom:Old_MAR.xlsx": {
            "category": "submittal", "subtype": "MAR",
            "fields": [{"key": "subject", "label": "الموضوع", "sheet": "Form", "cell": "B1", "role": "manual"}],
        }
    }, ensure_ascii=False), encoding="utf-8")
    mig = tprof.get_template_profile(Path(proj["root"]), "Old_MAR.xlsx", device_id="pc-a")
    assert mig["has_saved_profile"] is True
    assert (tdir / "Old_MAR.xlsx.tplmap.json").exists()


def test_contextual_petty_cash_and_new_project_by_name_fuzzy_dedup(tmp_path: Path):
    proj = _mk_proj(
        tmp_path, "p1", "مشروع مدرسة المطلاع", "MT-01", "C-2026-55",
        unrelated_policy="new_project_by_name",
    )
    ds = DecisionStore(tmp_path / "state", device_id="pc-a")
    gcfg = {"default_unrelated_policy": "new_project_by_name", "relevance_threshold": 0.6, "process_mode": "active_only"}

    # 1) Petty cash invoice from known subcontractor -> ACCEPTED as relevant & recorded in uploaded log!
    pc_file = Path(proj["inbox_dir"]) / "receipt_01.pdf"
    pc_file.write_bytes(b"%PDF-1.4 subcontractor petty cash")
    r_pc = ip.handle_inbox_document(
        pc_file,
        extracted_text="فاتورة نقدية عهدة موقع - شركة النور للكهرباء - توريد أسلاك لمشروع مدرسة المطلاع عقد C-2026-55",
        ai_meta={
            "category": "petty_cash",
            "is_project_related": True,
            "relevance_score": 0.88,
            "party_name": "شركة النور للكهرباء",
            "subject_ar": "توريد أسلاك كهربائية",
            "doc_date": "2026-11-17",
            "amount": 45.750,
        },
        active_proj=proj, all_projects=[proj], global_settings=gcfg, decision_store=ds, device_id="pc-a",
    )
    assert r_pc.get("ok") is True and r_pc.get("archive_path")
    assert plog.read_project_log_summary(proj)["stats"]["total_rows"] == 2

    # 2) Two unrelated files with orthographic variations of the SAME Arabic entity -> SAME folder (no duplicates)!
    u1 = Path(proj["inbox_dir"]) / "u1.pdf"
    u2 = Path(proj["inbox_dir"]) / "u2.pdf"
    u1.write_bytes(b"%PDF-1.4 other entity doc 1")
    u2.write_bytes(b"%PDF-1.4 other entity doc 2")

    res1 = ip.handle_inbox_document(
        u1, extracted_text="عرض سعر لعيادة الأسنان - شركة الصفا للتجهيزات الطبية",
        ai_meta={"is_project_related": False, "relevance_score": 0.1, "detected_project_or_entity": "شركة الصفا للتجهيزات الطبية", "subject_ar": "عرض سعر 1", "doc_date": "2026-11-17"},
        active_proj=proj, all_projects=[proj], global_settings=gcfg, decision_store=ds, device_id="pc-a",
    )
    res2 = ip.handle_inbox_document(
        u2, extracted_text="فاتورة صيانة - شركه الصفا للتجهيزات الطبيه",
        ai_meta={"is_project_related": False, "relevance_score": 0.1, "detected_project_or_entity": "شركه الصفا للتجهيزات الطبيه", "subject_ar": "فاتورة صيانة 2", "doc_date": "2026-11-17"},
        active_proj=proj, all_projects=[proj], global_settings=gcfg, decision_store=ds, device_id="pc-a",
    )
    e1 = res1.get("journal_entry") or res1.get("route") or res1.get("result") or res1
    e2 = res2.get("journal_entry") or res2.get("route") or res2.get("result") or res2
    p1_out = Path(e1.get("new_path") or e1.get("dest_path"))
    p2_out = Path(e2.get("new_path") or e2.get("dest_path"))
    assert p1_out.parent == p2_out.parent
    assert p1_out.exists() and p2_out.exists()


def test_cross_project_conflict_ask_policy_and_learned_rules(tmp_path: Path):
    p_a = _mk_proj(tmp_path, "pa", "مشروع برج السالمية", "SAL-01", "SAL-100", unrelated_policy="ask")
    p_b = _mk_proj(tmp_path, "pb", "مشروع مستشفى الجهراء", "JAH-02", "JAH-200", unrelated_policy="ask")
    ds = DecisionStore(tmp_path / "state", device_id="pc-b")
    gcfg = {"default_unrelated_policy": "ask", "relevance_threshold": 0.6, "process_mode": "active_only"}

    # File belonging to Project B accidentally scanned into active Project A's inbox!
    wrong_ibx_file = Path(p_a["inbox_dir"]) / "scan_jahra.pdf"
    wrong_ibx_file.write_bytes(b"%PDF-1.4 jahra hospital letter")
    r_conf = ip.handle_inbox_document(
        wrong_ibx_file,
        extracted_text="كتاب رسمي بخصوص مشروع مستشفى الجهراء عقد رقم JAH-200 اعتماد مخططات",
        ai_meta={
            "category": "outgoing",
            "is_project_related": False,
            "relevance_score": 0.15,
            "detected_project_or_entity": "مشروع مستشفى الجهراء",
            "subject_ar": "اعتماد مخططات مستشفى الجهراء",
            "doc_date": "2026-11-17",
        },
        active_proj=p_a, all_projects=[p_a, p_b], global_settings=gcfg, decision_store=ds, device_id="pc-b",
    )
    assert r_conf["outcome"] == "pending_decision"
    dec_id = r_conf["decision"]["id"]
    assert wrong_ibx_file.exists()  # untouched in inbox until user decides!

    # Resolve -> attach to Project B + remember_rule=True
    resolved = ip.resolve_pending_decision(
        dec_id,
        action="attach_project",
        params={"target_project_id": "pb", "remember_rule": True, "rule_match_type": "keyword", "rule_pattern": "مستشفى الجهراء"},
        decision_store=ds,
        all_projects=[p_a, p_b],
        global_settings=gcfg,
        device_id="pc-b",
    )
    assert resolved["ok"] is True
    assert resolved["created_rule"] is not None
    assert plog.read_project_log_summary(p_a)["stats"]["total_rows"] == 1  # Project A unpolluted
    assert plog.read_project_log_summary(p_b)["stats"]["total_rows"] == 2  # Recorded in Project B's log!

    # Next document matching the learned rule auto-routes to Project B directly!
    f_next = Path(p_a["inbox_dir"]) / "scan_jahra_2.pdf"
    f_next.write_bytes(b"%PDF-1.4 second jahra doc")
    r_auto = ip.handle_inbox_document(
        f_next, extracted_text="مذكرة موقع مشروع مستشفى الجهراء تسليم أعمال",
        ai_meta={"category": "outgoing", "subject_ar": "مذكرة تسليم", "doc_date": "2026-11-17"},
        active_proj=p_a, all_projects=[p_a, p_b], global_settings=gcfg, decision_store=ds, device_id="pc-b",
    )
    assert r_auto["outcome"] == "routed_by_learned_rule"
    assert plog.read_project_log_summary(p_b)["stats"]["total_rows"] == 3


def test_dry_run_ignore_patterns_and_locked_files(tmp_path: Path):
    proj = _mk_proj(tmp_path, "p1", "مشروع أ", "A-01", "C-1", dry_run=True)
    ds = DecisionStore(tmp_path / "state", device_id="pc-a")
    ibx = Path(proj["inbox_dir"])

    # Ignored system/exe/temp/glob files
    for bad in ("Thumbs.db", "~$open.xlsx", "installer.exe", "notes.bak"):
        bp = ibx / bad
        bp.write_bytes(b"ignore me")
        ok, _ = ip.preflight_inbox_file(bp, inbox_root=ibx, ignore_patterns=proj["ignore_patterns"], decision_store=ds)
        assert ok is False
        assert bp.exists()  # zero deletion!

    # Locked Office file skipped safely
    real_doc = ibx / "report.xlsx"
    lock_doc = ibx / "~$report.xlsx"
    real_doc.write_bytes(b"PK\x03\x04content")
    lock_doc.write_bytes(b"owner")
    ok_l, reason_l = ip.preflight_inbox_file(real_doc, inbox_root=ibx, ignore_patterns=[], decision_store=ds)
    assert ok_l is False and real_doc.exists()
    lock_doc.unlink()

    # Dry-run mode NEVER moves or renames valid file
    sample = ibx / "valid_scan.pdf"
    sample.write_bytes(b"%PDF-1.4 dry run sample")
    dr = ip.handle_inbox_document(
        sample, extracted_text="فاتورة مطعم غير مرتبطة", ai_meta={"is_project_related": False, "relevance_score": 0.1},
        active_proj=proj, all_projects=[proj], global_settings={"dry_run": True}, decision_store=ds, device_id="pc-a",
    )
    assert dr["outcome"] == "dry_run"
    assert sample.exists()


def test_active_project_modes_shared_inbox_start_end_and_backup_warning(tmp_path: Path):
    shared_ibx = tmp_path / "Shared_Scanner"
    p1 = _mk_proj(tmp_path, "p1", "مشروع 1", "P1", "C-1", inbox_dir=shared_ibx)
    p2 = _mk_proj(tmp_path, "p2", "مشروع 2", "P2", "C-2", inbox_dir=shared_ibx)
    ext_bkp = tmp_path / "USB_Backup"

    # Multi-project defaults to active_only + detects shared inbox
    assert ap.effective_process_mode({}, 2) == "active_only"
    selected = ap.get_projects_to_process([p1, p2], "p2", "active_only")
    assert [x["id"] for x in selected] == ["p2"]
    assert len(ap.detect_shared_inboxes([p1, p2])) == 1

    # >7-day stale backup warning fires when last_backup_at is 10 days old
    p2["last_backup_at"] = (dt.datetime.now() - dt.timedelta(days=10)).isoformat()
    bstat = ap.check_backup_status(p2, {"backup_warn_days": 7}, warn_days=7)
    assert len(bstat["warnings"]) >= 1

    # Start Work / Sync recalculates counters directly from uploaded log (OUT-005 -> OUT-006)
    gcfg = {"external_backup_dir": str(ext_bkp), "backup_warn_days": 7}
    start_res = ap.run_start_sync(
        p2, all_projects=[p1, p2], global_settings=gcfg, pending_decisions_count=0, device_id="pc-b"
    )
    assert start_res["ok"] is True
    assert start_res["counters"]["next_preview"]["OUT"] == "OUT-006"

    # End Session creates dated _Backups snapshot + external USB mirror & marks active_cleared=True
    end_res = ap.run_end_session(
        p2, global_settings=gcfg, device_id="pc-b", copy_external=True, include_archive_external=False
    )
    assert end_res["ok"] is True
    assert end_res["active_cleared"] is True
    assert end_res["all_writes_closed"] is True
    assert (Path(p2["root"]) / "_Backups").exists()
    assert len(list((Path(p2["root"]) / "_Backups").iterdir())) >= 1
    assert end_res["backup"]["external"]["copied"] is True
