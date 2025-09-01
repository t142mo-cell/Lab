# -*- coding: utf-8 -*-
from __future__ import annotations
import json, hashlib
from typing import List, Dict, Any
from models import Item
from constants import ITEMS_JSON, USERS_JSON, NEEDS_JSON, DATA_DIR, NEEDS_DEPARTMENTS

def _ensure_data_dir():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

def load_items() -> List[Item]:
    _ensure_data_dir()
    if not ITEMS_JSON.exists():
        ITEMS_JSON.write_text("[]", encoding="utf-8")
    data = json.loads(ITEMS_JSON.read_text(encoding="utf-8"))
    return [Item.from_dict(x) for x in data]

def save_items(items: List[Item]) -> None:
    _ensure_data_dir()
    with open(ITEMS_JSON, "w", encoding="utf-8") as f:
        json.dump([i.to_dict() for i in items], f, ensure_ascii=False, indent=2)

def get_next_seq_id(items: List[Item]) -> int:
    return (max((i.seq_id for i in items), default=0)) + 1

def hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

def load_users() -> List[Dict[str, Any]]:
    _ensure_data_dir()
    if not USERS_JSON.exists():
        USERS_JSON.write_text("[]", encoding="utf-8")
    return json.loads(USERS_JSON.read_text(encoding="utf-8"))

def save_users(users: List[Dict[str, Any]]) -> None:
    _ensure_data_dir()
    with open(USERS_JSON, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def ensure_default_admin() -> None:
    users = load_users()
    if not any(u.get("role") == "admin" for u in users):
        users.append({
            "username": "admin",
            "password_hash": hash_password("admin"),
            "role": "admin",
            "department": "Учет, хранение и выдача"
        })
        save_users(users)

def load_needs() -> Dict[str, Any]:
    _ensure_data_dir()
    if not NEEDS_JSON.exists():
        NEEDS_JSON.write_text("{}", encoding="utf-8")
    data = json.loads(NEEDS_JSON.read_text(encoding="utf-8"))
    from datetime import date
    next_year = date.today().year + 1
    data.setdefault("plan_year", next_year)
    data.setdefault("locked", False)
    data.setdefault("departments", { d: [] for d in NEEDS_DEPARTMENTS })
    for d in NEEDS_DEPARTMENTS:
        data["departments"].setdefault(d, [])
    data.setdefault("qa_overflow_requests", [])
    data.setdefault("issues", [])
    data.setdefault("store_requests", [])
    return data

def save_needs(needs: Dict[str, Any]) -> None:
    _ensure_data_dir()
    with open(NEEDS_JSON, "w", encoding="utf-8") as f:
        json.dump(needs, f, ensure_ascii=False, indent=2)

def next_need_id(needs: Dict[str, Any]) -> int:
    max_id = 0
    for lst in needs.get("departments", {}).values():
        for n in lst:
            max_id = max(max_id, int(n.get("need_id", 0)))
    return max_id + 1

def next_qa_request_id(needs: Dict[str, Any]) -> int:
    max_id = 0
    for r in needs.get("qa_overflow_requests", []):
        max_id = max(max_id, int(r.get("request_id", 0)))
    return max_id + 1

def next_issue_id(needs: Dict[str, Any]) -> int:
    max_id = 0
    for r in needs.get("issues", []):
        max_id = max(max_id, int(r.get("issue_id", 0)))
    return max_id + 1

def next_store_request_id(needs: Dict[str, Any]) -> int:
    max_id = 0
    for r in needs.get("store_requests", []):
        max_id = max(max_id, int(r.get("request_id", 0)))
    return max_id + 1
