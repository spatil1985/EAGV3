"""Decision role — one LLM call per step, returns answer xor tool_call."""
from __future__ import annotations

import json
import os
from collections.abc import Callable
from typing import Any

import httpx
from dotenv import load_dotenv

from schemas import DecisionOutput, Goal, ToolCall

load_dotenv()
GATEWAY_URL = os.getenv("GATEWAY_URL", "http://localhost:8101")

SYSTEM = """You are the Decision module of an agentic system. Your job is to choose the next action for ONE goal.

You have two possible outputs — choose EXACTLY ONE:

Option A — Answer (use when you already have enough information to satisfy the goal):
Respond with plain prose. Do NOT use JSON.

Option B — Tool call (use when you need to gather information or take an action):
Respond with ONLY this JSON (no other text):
{
  "FUNCTION_CALL": {
    "name": "<tool_name>",
    "arguments": { ... }
  }
}

Rules:
- Never pick more than one tool per response.
- Never narrate. Never explain your reasoning.
- If attached artifact content is provided, read it carefully before deciding.
- Do not hallucinate tool names. Only use tools from the available list.
- Artifact handles (strings starting with "art:") are NOT valid tool arguments — never pass them.
"""


def _tool_schema(tools: list[Any]) -> str:
    lines = []
    for t in tools:
        name = getattr(t, "name", str(t))
        desc = getattr(t, "description", "")
        schema = getattr(t, "inputSchema", {})
        props = schema.get("properties", {}) if isinstance(schema, dict) else {}
        args = ", ".join(f"{k}: {v.get('type','any')}" for k, v in props.items())
        lines.append(f"- {name}({args}): {desc}")
    return "\n".join(lines)


def next_step(
    goal: Goal,
    attached_bytes: list[bytes],
    history: list[dict],
    tools: list[Any],
    emit: Callable[[str], None] | None = None,
) -> DecisionOutput:
    def _log(msg: str) -> None:
        if emit:
            try:
                emit(msg)
            except Exception:
                pass

    tool_list = _tool_schema(tools)

    history_text = ""
    for h in history[-6:]:
        role = h.get("role", "?")
        if role == "answer":
            history_text += f"[answer] {h.get('text','')[:300]}\n"
        elif role == "tool_result":
            history_text += f"[tool:{h.get('tool')}] {h.get('result_descriptor','')[:200]}\n"

    attachment_text = ""
    if attached_bytes:
        for i, b in enumerate(attached_bytes):
            text = b.decode("utf-8", errors="replace")[:3000]
            attachment_text += f"\n--- ATTACHED ARTIFACT {i+1} ---\n{text}\n"

    _log(f"[decision.in] goal={goal.text[:80]} | history={len(history)} entries | attachment={sum(len(b) for b in attached_bytes)} bytes")

    user_msg = f"""GOAL: {goal.text}

AVAILABLE TOOLS:
{tool_list}

RECENT HISTORY:
{history_text or "(none)"}
{attachment_text}
Now choose: answer in plain text OR call exactly one tool with JSON."""

    resp = httpx.post(
        f"{GATEWAY_URL}/chat",
        json={
            "messages": [
                {"role": "system", "content": SYSTEM},
                {"role": "user", "content": user_msg},
            ],
            "auto_route": "decision",
        },
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json().get("content", "")
    if isinstance(content, dict):
        content = json.dumps(content)

    content = content.strip()
    _log(f"[decision.raw] {content[:500]}")

    # Strip markdown code fences — LLM sometimes wraps JSON in ```json ... ```
    if content.startswith("```"):
        first_nl = content.find("\n")
        if first_nl != -1:
            content = content[first_nl + 1:]
        if content.rstrip().endswith("```"):
            content = content.rstrip()[:-3].rstrip()
        content = content.strip()
        _log(f"[decision.stripped] {content[:300]}")

    # Try to parse as tool call
    try:
        parsed = json.loads(content)
        if "FUNCTION_CALL" in parsed:
            fc = parsed["FUNCTION_CALL"]
            _log(f"[decision.out] tool_call: {fc['name']}({fc.get('arguments', {})})")
            return DecisionOutput(tool_call=ToolCall(name=fc["name"], arguments=fc.get("arguments", {})))
    except (json.JSONDecodeError, KeyError):
        pass

    # Treat as plain answer
    _log(f"[decision.out] plain answer ({len(content)} chars): {content[:200]}")
    return DecisionOutput(answer=content)
