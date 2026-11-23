#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
decisions.py — Persistent Pending Decisions queue + user-editable Learned Rules store
for unrelated files, cross-project conflicts, and low-confidence scans (Smart_Local_Agent_Free).
"""

from __future__ import annotations

import datetime as dt
import secrets
import threading
from pathlib import Path
from typing import Any

from i18n_ar import tr
from security_utils import atomic_write_json, get_or_create_device_id, now_iso, read_json_safe, sha256_of_file

VALID_DECISION_ACTIONS: frozenset[str] = frozenset({
    "attach_project",
    "rename_in_place",
    "new_project_by_name",
    "ignore_once",
    "ignore_pattern",
})

VALID_RULE_MATCH_TYPES: frozenset[str] = frozenset({
    "entity",
    "sender",
    "keyword",
    "filename_glob",
})


class DecisionStore:
    """Thread-safe persistent store for pending decisions and learned routing rules."""

    def __init__(self, state_dir: Path, device_id: str = ""):
        self.state_dir = Path(state_dir)
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.decisions_path = self.state_dir / "pending_decisions.json"
        self.rules_path = self.state_dir / "learned_rules.json"
        self.ignored_hashes_path = self.state_dir / "ignored_file_hashes.json"
        self.device_id = device_id or get_or_create_device_id(self.state_dir)
        self._lock = threading.RLock()

    # ── Pending Decisions ────────────────────────────────────────────────────
    def _load_decisions_unlocked(self) -> list[dict[str, Any]]:
        data = read_json_safe(self.decisions_path, [])
        return data if isinstance(data, list) else []

    def _save_decisions_unlocked(self, items: list[dict[str, Any]]) -> None:
        atomic_write_json(self.decisions_path, items)

    def list_decisions(
        self,
        *,
        status: str | None = None,
        project_id: str | None = None,
        limit: int = 300,
    ) -> list[dict[str, Any]]:
        with self._lock:
            items = self._load_decisions_unlocked()
        out: list[dict[str, Any]] = []
        for d in reversed(items):
            if status and d.get("status") != status:
                continue
            if project_id and project_id not in (
                d.get("source_project_id"),
                d.get("suggested_project_id"),
                d.get("conflict_project_id"),
                d.get("resolved_project_id"),
            ):
                continue
            item_copy = dict(d)
            fp = item_copy.get("file_path")
            item_copy["file_exists"] = bool(fp and Path(str(fp)).exists())
            out.append(item_copy)
            if len(out) >= limit:
                break
        return out

    def pending_count(self, project_id: str | None = None) -> int:
        return len(self.list_decisions(status="pending", project_id=project_id, limit=10000))

    def get_decision(self, decision_id: str) -> dict[str, Any] | None:
        with self._lock:
            for d in self._load_decisions_unlocked():
                if d.get("id") == decision_id:
                    return dict(d)
        return None

    def has_pending_for_file(self, file_path: Path, file_sha256: str = "") -> bool:
        p_str = str(Path(file_path).resolve())
        with self._lock:
            for d in self._load_decisions_unlocked():
                if d.get("status") != "pending":
                    continue
                if d.get("file_path") == p_str:
                    return True
                if file_sha256 and d.get("file_sha256") == file_sha256:
                    # Only treat as duplicate if original file path still exists
                    if Path(str(d.get("file_path") or "")).exists():
                        return True
        return False

    def is_ignored_hash(self, file_sha256: str) -> bool:
        if not file_sha256:
            return False
        with self._lock:
            data = read_json_safe(self.ignored_hashes_path, {})
            return isinstance(data, dict) and file_sha256 in data

    def add_ignored_hash(self, file_sha256: str, file_name: str = "", reason: str = "user_ignored") -> None:
        if not file_sha256:
            return
        with self._lock:
            data = read_json_safe(self.ignored_hashes_path, {})
            if not isinstance(data, dict):
                data = {}
            data[file_sha256] = {
                "file_name": file_name,
                "reason": reason,
                "device_id": self.device_id,
                "timestamp": now_iso(),
            }
            atomic_write_json(self.ignored_hashes_path, data)

    def remove_ignored_hash(self, file_sha256: str) -> None:
        if not file_sha256:
            return
        with self._lock:
            data = read_json_safe(self.ignored_hashes_path, {})
            if isinstance(data, dict) and file_sha256 in data:
                data.pop(file_sha256, None)
                atomic_write_json(self.ignored_hashes_path, data)

    def resolve_by_file(
        self,
        *,
        file_sha256: str = "",
        file_path: str = "",
        project_id: str = "",
        action: str = "ignore_once",
        reason: str = "scanner_sync",
    ) -> int:
        """Close any matching pending decision(s) when resolved from another screen (e.g. Scanner)."""
        if not file_sha256 and not file_path:
            return 0
        norm_p = str(Path(file_path).resolve()).lower() if file_path else ""
        if action not in VALID_DECISION_ACTIONS:
            action = "ignore_once"
        closed = 0
        ts = now_iso()
        with self._lock:
            items = self._load_decisions_unlocked()
            for d in items:
                if d.get("status") != "pending":
                    continue
                if project_id and str(d.get("source_project_id") or "") not in ("", str(project_id)):
                    continue
                d_hash = str(d.get("file_sha256") or "")
                d_path = str(d.get("file_path") or "").lower()
                if (file_sha256 and d_hash == file_sha256) or (norm_p and d_path == norm_p):
                    d["status"] = "resolved"
                    d["resolved_at"] = ts
                    d["resolved_by_device"] = self.device_id
                    d["resolution"] = {
                        "action": action,
                        "result": {"synced_from": reason},
                        "created_rule_id": None,
                        "device_id": self.device_id,
                        "timestamp": ts,
                    }
                    closed += 1
            if closed:
                self._save_decisions_unlocked(items)
        return closed

    def enqueue_decision(
        self,
        *,
        kind: str,
        file_path: Path,
        source_project_id: str,
        source_project_name: str = "",
        reason: str = "",
        score: float = 0.0,
        suggested_project_id: str | None = None,
        conflict_project_id: str | None = None,
        conflict_project_name: str = "",
        suggested_entity: str = "",
        suggested_clean_name: str = "",
        extracted_excerpt: str = "",
        ai_meta: dict[str, Any] | None = None,
        signals: list[str] | None = None,
        dry_run: bool = False,
    ) -> dict[str, Any]:
        """Add a pending decision (deduplicated by path + sha256)."""
        fp = Path(file_path).resolve()
        fhash = ""
        if fp.exists() and fp.is_file():
            try:
                fhash = sha256_of_file(fp)
            except Exception:
                fhash = ""

        with self._lock:
            items = self._load_decisions_unlocked()
            for existing in items:
                if existing.get("status") == "pending" and (
                    existing.get("file_path") == str(fp)
                    or (fhash and existing.get("file_sha256") == fhash and Path(str(existing.get("file_path") or "")).exists())
                ):
                    return dict(existing)

            ts = dt.datetime.now().strftime("%Y%m%d%H%M%S")
            entry: dict[str, Any] = {
                "id": f"DEC-{ts}-{secrets.token_hex(3).upper()}",
                "kind": kind,  # unrelated_file | cross_project_conflict | low_confidence | dry_run_preview
                "status": "pending",
                "file_path": str(fp),
                "trusted_source_dir": str(fp.parent),
                "file_name": fp.name,
                "file_ext": fp.suffix.lower(),
                "file_sha256": fhash,
                "source_project_id": source_project_id,
                "source_project_name": source_project_name,
                "suggested_project_id": suggested_project_id,
                "conflict_project_id": conflict_project_id,
                "conflict_project_name": conflict_project_name,
                "suggested_entity": suggested_entity,
                "suggested_clean_name": suggested_clean_name or fp.name,
                "score": round(float(score or 0.0), 3),
                "reason": reason,
                "extracted_excerpt": (extracted_excerpt or "")[:500],
                "ai_meta": dict(ai_meta or {}),
                "signals": list(signals or []),
                "dry_run": bool(dry_run),
                "device_id": self.device_id,
                "created_at": now_iso(),
                "resolved_at": None,
                "resolution": None,
            }
            items.append(entry)
            self._save_decisions_unlocked(items)
            return dict(entry)

    def mark_resolved(
        self,
        decision_id: str,
        *,
        action: str,
        result_info: dict[str, Any] | None = None,
        resolved_project_id: str | None = None,
        created_rule_id: str | None = None,
    ) -> dict[str, Any]:
        if action not in VALID_DECISION_ACTIONS:
            raise ValueError(tr("server.invalid_setting"))
        with self._lock:
            items = self._load_decisions_unlocked()
            for d in items:
                if d.get("id") == decision_id:
                    d["status"] = "resolved"
                    d["resolved_at"] = now_iso()
                    d["resolved_by_device"] = self.device_id
                    d["resolved_project_id"] = resolved_project_id
                    d["resolution"] = {
                        "action": action,
                        "result": result_info or {},
                        "created_rule_id": created_rule_id,
                        "device_id": self.device_id,
                        "timestamp": now_iso(),
                    }
                    self._save_decisions_unlocked(items)
                    return dict(d)
        raise KeyError(tr("server.decision_not_found"))

    # ── Learned Rules (view / add / edit / delete) ───────────────────────────
    def _load_rules_unlocked(self) -> list[dict[str, Any]]:
        data = read_json_safe(self.rules_path, [])
        return data if isinstance(data, list) else []

    def _save_rules_unlocked(self, rules: list[dict[str, Any]]) -> None:
        atomic_write_json(self.rules_path, rules)

    def list_rules(self, *, include_disabled: bool = True) -> list[dict[str, Any]]:
        with self._lock:
            rules = self._load_rules_unlocked()
        if include_disabled:
            return [dict(r) for r in rules]
        return [dict(r) for r in rules if r.get("enabled", True)]

    def add_rule(
        self,
        *,
        match_type: str,
        pattern: str,
        action: str,
        target_project_id: str = "",
        entity_name: str = "",
        note: str = "",
        enabled: bool = True,
    ) -> dict[str, Any]:
        mt = (match_type or "entity").strip().lower()
        if mt not in VALID_RULE_MATCH_TYPES:
            raise ValueError(tr("server.invalid_setting"))
        pat = (pattern or "").strip()
        if not pat:
            raise ValueError(tr("validation.rule_pattern_req"))
        act = (action or "").strip().lower()
        if act not in ("attach_project", "rename_in_place", "new_project_by_name", "ignore"):
            raise ValueError(tr("server.invalid_setting"))
        if act == "attach_project" and not target_project_id:
            raise ValueError(tr("validation.rule_project_req"))

        with self._lock:
            rules = self._load_rules_unlocked()
            # Update in-place if identical match_type + pattern already exists
            for r in rules:
                if r.get("match_type") == mt and str(r.get("pattern", "")).strip().lower() == pat.lower():
                    r["action"] = act
                    r["target_project_id"] = target_project_id
                    r["entity_name"] = entity_name
                    r["note"] = note
                    r["enabled"] = bool(enabled)
                    r["updated_at"] = now_iso()
                    r["device_id"] = self.device_id
                    self._save_rules_unlocked(rules)
                    return dict(r)

            rule = {
                "id": f"RUL-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}-{secrets.token_hex(2).upper()}",
                "match_type": mt,
                "pattern": pat,
                "action": act,
                "target_project_id": target_project_id,
                "entity_name": entity_name,
                "note": note,
                "enabled": bool(enabled),
                "device_id": self.device_id,
                "created_at": now_iso(),
                "updated_at": now_iso(),
            }
            rules.append(rule)
            self._save_rules_unlocked(rules)
            return dict(rule)

    def update_rule(self, rule_id: str, patch: dict[str, Any]) -> dict[str, Any]:
        with self._lock:
            rules = self._load_rules_unlocked()
            for r in rules:
                if r.get("id") == rule_id:
                    if "match_type" in patch and patch["match_type"] in VALID_RULE_MATCH_TYPES:
                        r["match_type"] = patch["match_type"]
                    if "pattern" in patch and str(patch["pattern"]).strip():
                        r["pattern"] = str(patch["pattern"]).strip()
                    if "action" in patch and patch["action"] in (
                        "attach_project", "rename_in_place", "new_project_by_name", "ignore"
                    ):
                        r["action"] = patch["action"]
                    if "target_project_id" in patch:
                        r["target_project_id"] = str(patch["target_project_id"] or "")
                    if "entity_name" in patch:
                        r["entity_name"] = str(patch["entity_name"] or "")
                    if "note" in patch:
                        r["note"] = str(patch["note"] or "")
                    if "enabled" in patch:
                        r["enabled"] = bool(patch["enabled"])
                    r["updated_at"] = now_iso()
                    r["device_id"] = self.device_id
                    self._save_rules_unlocked(rules)
                    return dict(r)
        raise KeyError(tr("server.rule_not_found"))

    def delete_rule(self, rule_id: str) -> bool:
        with self._lock:
            rules = self._load_rules_unlocked()
            kept = [r for r in rules if r.get("id") != rule_id]
            if len(kept) == len(rules):
                return False
            self._save_rules_unlocked(kept)
            return True

    def clear_all(self) -> dict[str, int]:
        """Clear all pending decisions, learned rules, and ignored file hashes."""
        with self._lock:
            dec_cnt = len(self._load_decisions_unlocked())
            rules_raw = read_json_safe(self.rules_path, [])
            rule_cnt = len(rules_raw) if isinstance(rules_raw, list) else 0
            ign_raw = read_json_safe(self.ignored_hashes_path, [])
            ign_cnt = len(ign_raw) if isinstance(ign_raw, list) else 0
            atomic_write_json(self.decisions_path, [])
            atomic_write_json(self.rules_path, [])
            atomic_write_json(self.ignored_hashes_path, {})
            return {
                "decisions_cleared": dec_cnt,
                "rules_cleared": rule_cnt,
                "ignored_cleared": ign_cnt,
            }

    def ignore_all_pending(self) -> int:
        """Mark all currently pending decisions as resolved (`ignore_once`) in one atomic write."""
        closed_items: list[dict[str, str]] = []
        with self._lock:
            items = self._load_decisions_unlocked()
            cleared = 0
            ts = now_iso()
            ign_data = read_json_safe(self.ignored_hashes_path, {})
            if not isinstance(ign_data, dict):
                ign_data = {}
            for d in items:
                if d.get("status") == "pending":
                    d["status"] = "resolved"
                    d["resolved_at"] = ts
                    d["resolved_by_device"] = self.device_id
                    d["resolution"] = {
                        "action": "ignore_once",
                        "result": {"ignored": True, "bulk": True},
                        "created_rule_id": None,
                        "device_id": self.device_id,
                        "timestamp": ts,
                    }
                    fh = str(d.get("file_sha256") or "")
                    if fh:
                        ign_data[fh] = {
                            "file_name": str(d.get("file_name") or ""),
                            "reason": "ignore_all_pending",
                            "device_id": self.device_id,
                            "timestamp": ts,
                        }
                    closed_items.append({
                        "file_sha256": fh,
                        "file_path": str(d.get("file_path") or ""),
                        "project_id": str(d.get("source_project_id") or ""),
                    })
                    cleared += 1
            if cleared:
                self._save_decisions_unlocked(items)
                atomic_write_json(self.ignored_hashes_path, ign_data)
            self._last_bulk_closed = closed_items
            return cleared
