"""Pure parsing helpers shared between flow.py and tests. No heavy imports."""
from __future__ import annotations

import json


def parse_critic_output(raw: str, upstream_content: str) -> tuple[str, str | None]:
    """
    Parse critic JSON and return (output_str, rationale).
    output_str starts with CRITIC_PASS or CRITIC_FAIL.
    """
    text = raw.strip()
    if text.startswith("```"):
        nl = text.find("\n")
        if nl != -1:
            text = text[nl + 1:]
        if text.rstrip().endswith("```"):
            text = text.rstrip()[:-3].rstrip()
    try:
        data = json.loads(text)
        verdict = data.get("verdict", "fail")
        rationale = data.get("rationale", "")
        if verdict == "pass":
            return f"CRITIC_PASS\n{upstream_content}", rationale
        return f"CRITIC_FAIL\n{rationale}", rationale
    except (json.JSONDecodeError, TypeError):
        return f"CRITIC_FAIL\n{text}", None


def parse_coder_output(raw: str) -> tuple[str, str | None, str | None]:
    """Return (output_summary, code, summary) from coder JSON."""
    text = raw.strip()
    if text.startswith("```"):
        nl = text.find("\n")
        if nl != -1:
            text = text[nl + 1:]
        if text.rstrip().endswith("```"):
            text = text.rstrip()[:-3].rstrip()
    try:
        data = json.loads(text)
        code = data.get("code", "")
        summary = data.get("summary", text)
        return summary, code, summary
    except (json.JSONDecodeError, TypeError):
        return text, None, None


def parse_planner_output(raw: str) -> tuple[str, list[dict]]:
    """Return (rationale, nodes_list) from planner JSON."""
    text = raw.strip()
    if text.startswith("```"):
        nl = text.find("\n")
        if nl != -1:
            text = text[nl + 1:]
        if text.rstrip().endswith("```"):
            text = text.rstrip()[:-3].rstrip()
    data = json.loads(text)
    return data.get("rationale", ""), data.get("nodes", [])
