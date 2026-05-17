"""Persistent long-term memory — key/value store in persistent_mem/memories.json.

All writes are atomic (write-temp-then-rename) to avoid corruption.
"""
from __future__ import annotations

import json
import os
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

STORE_DIR = Path("persistent_mem")
STORE_FILE = STORE_DIR / "memories.json"


# ── I/O ──────────────────────────────────────────────────────────────────────

def _load() -> list[dict]:
    STORE_DIR.mkdir(parents=True, exist_ok=True)
    if STORE_FILE.exists():
        try:
            return json.loads(STORE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _save(items: list[dict]) -> None:
    STORE_DIR.mkdir(parents=True, exist_ok=True)
    tmp = STORE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(items, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(STORE_FILE)


# ── Public API ────────────────────────────────────────────────────────────────

def save(
    key: str,
    value: str,
    category: str = "fact",
    tags: list[str] | None = None,
    notes: str = "",
) -> dict:
    """Upsert a memory entry by key. Returns the saved entry."""
    items = _load()
    now = datetime.now().isoformat(timespec="seconds")

    for item in items:
        if item.get("key") == key:
            item["value"] = value
            item["category"] = category
            item["tags"] = tags or []
            item["notes"] = notes
            item["updated_at"] = now
            _save(items)
            return item

    entry: dict = {
        "id": uuid.uuid4().hex[:12],
        "key": key,
        "value": value,
        "category": category,
        "tags": tags or [],
        "notes": notes,
        "created_at": now,
        "updated_at": now,
    }
    items.append(entry)
    _save(items)
    return entry


def load(key: str) -> dict | None:
    """Return entry for key, or None if not found."""
    for item in _load():
        if item.get("key") == key:
            return item
    return None


def search(query: str) -> list[dict]:
    """Keyword search across key, value, tags, notes. Returns ranked results."""
    stopwords = {"the", "a", "an", "is", "in", "on", "at", "to", "for", "of", "and", "or"}
    tokens = {w.lower() for w in re.findall(r"\w+", query) if w.lower() not in stopwords}
    if not tokens:
        return _load()
    results: list[tuple[int, dict]] = []
    for item in _load():
        haystack = " ".join([
            item.get("key", ""),
            item.get("value", ""),
            " ".join(item.get("tags", [])),
            item.get("notes", ""),
            item.get("category", ""),
        ]).lower()
        score = len(tokens & set(re.findall(r"\w+", haystack)))
        if score > 0:
            results.append((score, item))
    results.sort(key=lambda x: x[0], reverse=True)
    return [item for _, item in results]


def list_all() -> list[dict]:
    """Return all memory entries sorted by most-recently updated."""
    items = _load()
    items.sort(key=lambda x: x.get("updated_at", ""), reverse=True)
    return items


def delete(key: str) -> bool:
    """Delete by key. Returns True if something was removed."""
    items = _load()
    new_items = [i for i in items if i.get("key") != key]
    if len(new_items) < len(items):
        _save(new_items)
        return True
    return False
