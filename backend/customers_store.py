"""
Simple JSON-file backed CRUD store for the customer list.

No real database is needed for ~50-a-few-hundred customers; a single JSON file
(data/customers.json) that's bind-mounted into the container (see
docker-compose.yml) is enough, is human-readable/editable, and survives
container rebuilds.

Each customer record: {"id": str, "name": str, "full_text": str}
`full_text` is the exact multi-line address block that gets dropped into
documents wherever "the customer's information" is requested (name + address
+ phone etc, however that particular customer's block is written).
"""
from __future__ import annotations

import json
import re
import threading
from pathlib import Path
from typing import List, Optional

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
CUSTOMERS_PATH = DATA_DIR / "customers.json"
_lock = threading.RLock()  # reentrant: list_customers() is called from inside
                           # add/update/delete's own `with _lock:` blocks below


def _ensure_file():
    DATA_DIR.mkdir(exist_ok=True)
    if not CUSTOMERS_PATH.exists():
        CUSTOMERS_PATH.write_text("[]", encoding="utf-8")


def _slugify(name: str, existing_ids: set) -> str:
    base = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")[:40] or "customer"
    cid = base
    n = 2
    while cid in existing_ids:
        cid = f"{base}-{n}"
        n += 1
    return cid


def list_customers() -> List[dict]:
    _ensure_file()
    with _lock:
        return json.loads(CUSTOMERS_PATH.read_text(encoding="utf-8"))


def _save(customers: List[dict]) -> None:
    CUSTOMERS_PATH.write_text(json.dumps(customers, ensure_ascii=False, indent=2), encoding="utf-8")


def get_customer(customer_id: str) -> Optional[dict]:
    for c in list_customers():
        if c["id"] == customer_id:
            return c
    return None


def add_customer(name: str, full_text: str) -> dict:
    name = (name or "").strip()
    full_text = (full_text or "").strip()
    if not name:
        raise ValueError("Customer name is required.")
    if not full_text:
        full_text = name
    with _lock:
        customers = list_customers()
        existing_ids = {c["id"] for c in customers}
        record = {"id": _slugify(name, existing_ids), "name": name, "full_text": full_text}
        customers.append(record)
        _save(customers)
        return record


def update_customer(customer_id: str, name: str, full_text: str) -> dict:
    with _lock:
        customers = list_customers()
        for c in customers:
            if c["id"] == customer_id:
                c["name"] = (name or c["name"]).strip()
                c["full_text"] = (full_text or c["full_text"]).strip()
                _save(customers)
                return c
        raise KeyError(f"No customer with id {customer_id}")


def delete_customer(customer_id: str) -> None:
    with _lock:
        customers = list_customers()
        remaining = [c for c in customers if c["id"] != customer_id]
        if len(remaining) == len(customers):
            raise KeyError(f"No customer with id {customer_id}")
        _save(remaining)
