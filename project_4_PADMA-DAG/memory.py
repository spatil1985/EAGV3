"""Memory role — typed service over state/memory.json."""
from __future__ import annotations

import json
import os
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx
from dotenv import load_dotenv

from schemas import GatewayMessage, MemoryItem, ToolCall

load_dotenv()

MEMORY_PATH = Path("state/memory.json")
GATEWAY_URL = os.getenv("GATEWAY_URL", "http://localhost:8101")
TOP_K = 5


def _load() -> list[dict]:
    if MEMORY_PATH.exists():
        try:
            return json.loads(MEMORY_PATH.read_text())
        except Exception:
            return []
    return []


def _save(items: list[dict]) -> None:
    MEMORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    MEMORY_PATH.write_text(json.dumps(items, indent=2, default=str))


def _keyword_score(item: MemoryItem, query_tokens: set[str]) -> float:
    item_tokens = {k.lower() for k in item.keywords}
    item_tokens |= set(re.findall(r"\w+", item.descriptor.lower()))
    return len(query_tokens & item_tokens) / max(len(query_tokens), 1)


def _query_tokens(query: str) -> set[str]:
    stopwords = {"the", "a", "an", "is", "in", "on", "at", "to", "for", "of", "and", "or"}
    return {w.lower() for w in re.findall(r"\w+", query) if w.lower() not in stopwords}


def read(
    query: str,
    history: list[dict] | None = None,
    kinds: list[str] | None = None,
    top_k: int = TOP_K,
) -> list[MemoryItem]:
    raw = _load()
    tokens = _query_tokens(query)
    items = [MemoryItem(**r) for r in raw]

    if kinds:
        items = [i for i in items if i.kind in kinds]

    scored = [(i, _keyword_score(i, tokens)) for i in items]
    scored.sort(key=lambda x: x[1], reverse=True)
    return [i for i, s in scored[:top_k] if s > 0]


def remember(query: str, source: str = "user_query", run_id: str = "") -> MemoryItem | None:
    """Classify query and store durable fact/preference if worthy."""
    prompt = f"""You are a memory classifier. Given a user query, decide if it contains a durable fact or user preference worth remembering across sessions.

User query: {query}

Respond with JSON only:
{{
  "worth_remembering": true/false,
  "kind": "fact" or "preference",
  "descriptor": "one short line",
  "keywords": ["keyword1", "keyword2", ...],
  "value": {{ ... relevant structured data ... }}
}}

If not worth remembering, set worth_remembering=false and leave other fields empty strings/lists."""

    try:
        resp = httpx.post(
            f"{GATEWAY_URL}/chat",
            json={
                "messages": [{"role": "user", "content": prompt}],
                "auto_route": "memory",
                "response_format": {
                    "type": "json_schema",
                    "schema_": {
                        "type": "object",
                        "properties": {
                            "worth_remembering": {"type": "boolean"},
                            "kind": {"type": "string"},
                            "descriptor": {"type": "string"},
                            "keywords": {"type": "array", "items": {"type": "string"}},
                            "value": {"type": "object"},
                        },
                        "required": ["worth_remembering"],
                    },
                },
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        content = data.get("content", {})
        if isinstance(content, str):
            content = json.loads(content)

        if not content.get("worth_remembering"):
            return None

        item = MemoryItem(
            id=uuid.uuid4().hex[:12],
            kind=content.get("kind", "fact"),
            keywords=content.get("keywords", []),
            descriptor=content.get("descriptor", query[:80]),
            value=content.get("value", {"raw": query}),
            source=source,
            run_id=run_id,
            created_at=datetime.now(),
        )
        raw = _load()
        raw.append(json.loads(item.model_dump_json()))
        _save(raw)
        return item

    except Exception as e:
        print(f"[memory.remember] error: {e}")
        return None


def record_outcome(
    tool_call: ToolCall,
    result_text: str,
    artifact_id: str | None,
    run_id: str,
    goal_id: str | None,
) -> MemoryItem:
    """Record a tool outcome in memory."""
    keywords = [tool_call.name] + list(tool_call.arguments.values())[:3]
    keywords = [str(k)[:30] for k in keywords if k]

    item = MemoryItem(
        id=uuid.uuid4().hex[:12],
        kind="tool_outcome",
        keywords=keywords,
        descriptor=f"{tool_call.name}: {result_text[:80]}",
        value={"tool": tool_call.name, "args": tool_call.arguments, "result": result_text[:500]},
        artifact_id=artifact_id,
        source="action",
        run_id=run_id,
        goal_id=goal_id,
        created_at=datetime.now(),
    )
    raw = _load()
    raw.append(json.loads(item.model_dump_json()))
    _save(raw)
    return item
