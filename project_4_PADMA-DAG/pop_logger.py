"""PoP (Prompt of Prompts) logger.

For every agent iteration, saves two files per module (perception / decision)
inside  PromptsOfPrompt/<run_id>/:

  iter<N>_perception_prompt.txt   — full system + user prompt text
  iter<N>_perception_pop.json     — structural validation of the prompt

  iter<N>_<goal_id>_decision_prompt.txt
  iter<N>_<goal_id>_decision_pop.json
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

POP_DIR = Path("PromptsOfPrompt")


def _ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _run_dir(run_id: str) -> Path:
    d = POP_DIR / run_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def save_perception(
    run_id: str,
    iteration: int,
    system_prompt: str,
    user_message: str,
    query: str,
    memory_hits_count: int,
    history_entries_count: int,
    prior_goals_count: int,
) -> None:
    try:
        d = _run_dir(run_id)
        prefix = f"iter{iteration:02d}_perception"

        prompt_text = (
            "=== SYSTEM PROMPT ===\n"
            + system_prompt
            + "\n\n=== USER MESSAGE ===\n"
            + user_message
            + "\n"
        )
        (d / f"{prefix}_prompt.txt").write_text(prompt_text, encoding="utf-8")

        pop = {
            "module": "perception",
            "run_id": run_id,
            "iteration": iteration,
            "timestamp": _ts(),
            "prompt_structure": {
                "system_prompt_chars": len(system_prompt),
                "user_message_chars": len(user_message),
                "total_chars": len(system_prompt) + len(user_message),
            },
            "context_inputs": {
                "user_query": query,
                "user_query_chars": len(query),
                "memory_hits_count": memory_hits_count,
                "history_entries_count": history_entries_count,
                "prior_goals_count": prior_goals_count,
            },
            "rules_present": {
                "decompose": "DECOMPOSE" in system_prompt,
                "mark_done": "MARK DONE" in system_prompt,
                "attach_artifact": "ATTACH" in system_prompt,
                "preserve_order": "PRESERVE ORDER" in system_prompt,
                "stall_detection": "STALL DETECTION" in system_prompt,
            },
            "schema_validation": {
                "json_response_required": "JSON" in system_prompt,
                "required_fields_specified": '"required"' in system_prompt or "required" in system_prompt,
                "output_format_specified": '"goals"' in system_prompt,
                "null_safety_note_present": 'attach_artifact_id must be either null' in system_prompt,
            },
        }
        (d / f"{prefix}_pop.json").write_text(
            json.dumps(pop, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except Exception:
        pass  # never crash the agent due to logging


def save_decision(
    run_id: str,
    iteration: int,
    goal_id: str,
    goal_text: str,
    system_prompt: str,
    user_message: str,
    tool_names: list[str],
    history_entries_count: int,
    attached_bytes: int,
) -> None:
    try:
        d = _run_dir(run_id)
        safe_goal = re.sub(r"[^\w]", "_", goal_id)
        prefix = f"iter{iteration:02d}_{safe_goal}_decision"

        prompt_text = (
            "=== SYSTEM PROMPT ===\n"
            + system_prompt
            + "\n\n=== USER MESSAGE ===\n"
            + user_message
            + "\n"
        )
        (d / f"{prefix}_prompt.txt").write_text(prompt_text, encoding="utf-8")

        pop = {
            "module": "decision",
            "run_id": run_id,
            "iteration": iteration,
            "goal_id": goal_id,
            "goal_text": goal_text,
            "timestamp": _ts(),
            "prompt_structure": {
                "system_prompt_chars": len(system_prompt),
                "user_message_chars": len(user_message),
                "total_chars": len(system_prompt) + len(user_message),
            },
            "context_inputs": {
                "tools_count": len(tool_names),
                "tools_available": tool_names,
                "history_entries_count": history_entries_count,
                "attached_artifacts_bytes": attached_bytes,
            },
            "rules_present": {
                "single_tool_per_response": "never pick more than one" in system_prompt.lower(),
                "no_narration": "Never narrate" in system_prompt,
                "no_hallucinate_tools": "hallucinate" in system_prompt.lower(),
                "no_artifact_handles_as_args": "art:" in system_prompt,
            },
            "output_format": {
                "answer_option_defined": "Option A" in system_prompt,
                "tool_call_json_format_defined": "FUNCTION_CALL" in system_prompt,
                "exclusive_choice_required": "EXACTLY ONE" in system_prompt,
            },
        }
        (d / f"{prefix}_pop.json").write_text(
            json.dumps(pop, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except Exception:
        pass  # never crash the agent due to logging
