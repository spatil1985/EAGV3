"""Perception role — orchestrator that maintains goal list across iterations."""
from __future__ import annotations

import json
import os
from collections.abc import Callable
from typing import Any

import httpx
from dotenv import load_dotenv

from schemas import Goal, MemoryItem, Observation

load_dotenv()
GATEWAY_URL = os.getenv("GATEWAY_URL", "http://localhost:8101")

SYSTEM = """You are the Perception module of an agentic system. Your role is to maintain a goal list.

You receive:
- The original user query
- Relevant memory hits from past runs
- The run history so far (tool results, interim answers)
- The prior goal list (empty on first iteration)

You must output a JSON goal list. Follow these five rules strictly:

1. DECOMPOSE: On first run (prior goals empty), break the query into 1-5 ordered goals.
2. MARK DONE: If history already contains a satisfying action for a goal, set done=true.
3. ATTACH: If a goal requires reading the *content* of a fetched resource (not just its URL),
   set attach_artifact_id to the artifact id (art:...) from the most recent relevant tool_result in history.
4. PRESERVE ORDER: Do not reorder or remove goals. Only update done flags and attach_artifact_id.
5. STALL DETECTION: If the same [answer] text appears 2 or more times for a goal in the history,
   the goal is impossible with available tools — set done=true and move on. Do not loop forever.

IMPORTANT: attach_artifact_id must be either null (JSON null, not the string "null") or a valid art:... id.

Respond ONLY with this JSON (no other text):
{
  "goals": [
    {
      "id": "g1",
      "text": "short imperative description",
      "done": false,
      "attach_artifact_id": null
    }
  ]
}
"""


def observe(
    query: str,
    memory_hits: list[MemoryItem],
    history: list[dict],
    prior_goals: list[Goal],
    run_id: str,
    emit: Callable[[str], None] | None = None,
) -> Observation:
    def _log(msg: str) -> None:
        if emit:
            try:
                emit(msg)
            except Exception:
                pass
    memory_text = ""
    if memory_hits:
        memory_text = "MEMORY HITS:\n"
        for m in memory_hits:
            memory_text += f"  [{m.kind}] {m.descriptor} — {json.dumps(m.value)[:200]}\n"

    history_text = ""
    for h in history:
        role = h.get("role", "?")
        if role == "answer":
            history_text += f"[answer] {h.get('text','')[:400]}\n"
        elif role == "tool_result":
            art_id = h.get("artifact_id", "")
            history_text += f"[tool:{h.get('tool')}] {h.get('result_descriptor','')} artifact={art_id}\n"

    prior_text = ""
    if prior_goals:
        prior_text = "PRIOR GOALS:\n"
        for g in prior_goals:
            status = "DONE" if g.done else "pending"
            prior_text += f"  [{status}] {g.id}: {g.text}"
            if g.attach_artifact_id:
                prior_text += f" (attach:{g.attach_artifact_id})"
            prior_text += "\n"

    _log(f"[perception.in] history={len(history)} entries, prior_goals={len(prior_goals)}, memory_hits={len(memory_hits)}")
    if history:
        _log(f"[perception.history]")
        for h in history[-5:]:
            role = h.get("role", "?")
            if role == "answer":
                _log(f"  [answer] {h.get('text', '')[:120]}")
            elif role == "tool_result":
                _log(f"  [tool:{h.get('tool')}] {h.get('result_descriptor', '')[:100]}")

    user_msg = f"""USER QUERY: {query}

{memory_text}
HISTORY:
{history_text or "(none yet)"}

{prior_text}
Update the goal list now."""

    resp = httpx.post(
        f"{GATEWAY_URL}/chat",
        json={
            "messages": [
                {"role": "system", "content": SYSTEM},
                {"role": "user", "content": user_msg},
            ],
            "provider": "g",
            "response_format": {
                "type": "json_schema",
                "schema_": {
                    "type": "object",
                    "properties": {
                        "goals": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "id": {"type": "string"},
                                    "text": {"type": "string"},
                                    "done": {"type": "boolean"},
                                    "attach_artifact_id": {"type": "string"},
                                },
                                "required": ["id", "text", "done"],
                            },
                        }
                    },
                    "required": ["goals"],
                },
            },
        },
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json().get("content", {})
    if isinstance(content, str):
        content = json.loads(content)

    _log(f"[perception.raw] {json.dumps(content)[:400]}")

    _INVALID_ATTACH = {None, "null", "None", "none", "undefined", ""}

    def _clean_attach(v: Any) -> str | None:
        return v if (v and str(v) not in _INVALID_ATTACH and str(v).startswith("art:")) else None

    raw_goals = content.get("goals", [])
    goals = [
        Goal(
            id=g["id"],
            text=g["text"],
            done=g.get("done", False),
            attach_artifact_id=_clean_attach(g.get("attach_artifact_id")),
        )
        for g in raw_goals
    ]
    for g in goals:
        status = "DONE" if g.done else "pending"
        attach = f" attach={g.attach_artifact_id}" if g.attach_artifact_id else ""
        _log(f"[perception.goal] [{status}] {g.id}: {g.text[:80]}{attach}")

    # Safety: if Perception returns empty on first iteration, create a single default goal
    if not goals:
        goals = [Goal(id="g1", text=query, done=False)]

    return Observation(goals=goals)
