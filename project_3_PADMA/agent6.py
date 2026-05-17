"""agent6.py — main loop wiring Memory, Perception, Decision, Action."""
from __future__ import annotations

import asyncio
import sys
import uuid
from collections.abc import Callable
from pathlib import Path

from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

import action
import artifacts
import decision
import memory
import perception
from schemas import Goal

load_dotenv()

MAX_ITERATIONS = 12
MCP_SERVER = Path(__file__).parent / "mcp_server.py"

_progress_cb: Callable[[str], None] | None = None


def _emit(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        # Windows cp1252 console can't encode chars like → ✓ — degrade gracefully.
        print(msg.encode("ascii", errors="replace").decode("ascii"))
    if _progress_cb:
        try:
            _progress_cb(msg)
        except Exception:
            pass


def _final_answer(history: list[dict]) -> str:
    for h in reversed(history):
        if h.get("role") == "answer":
            return h["text"]
    lines = [f"- {h.get('result_descriptor', '')}" for h in history if h.get("role") == "tool_result"]
    return "\n".join(lines) if lines else "Task completed."


async def run(query: str, on_progress: Callable[[str], None] | None = None) -> str:
    global _progress_cb
    _progress_cb = on_progress

    run_id = uuid.uuid4().hex[:8]
    _emit(f"[agent6] run_id={run_id}")
    _emit(f"[agent6] query: {query}")

    # Durable memory — runs in thread so it doesn't block the event loop
    try:
        mem_item = await asyncio.to_thread(memory.remember, query, "user_query", run_id)
        if mem_item:
            _emit(f"[memory] stored: {mem_item.descriptor}")
    except Exception as e:
        _emit(f"[memory] remember error: {e}")

    history: list[dict] = []
    prior_goals: list[Goal] = []
    _stall_tracker: dict[str, list[str]] = {}  # goal_id -> list of answers produced

    server_params = StdioServerParameters(
        command="python",
        args=[str(MCP_SERVER)],
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools_result = await session.list_tools()
            tools = tools_result.tools
            _emit(f"[mcp] {len(tools)} tools loaded: {[t.name for t in tools]}")

            for it in range(1, MAX_ITERATIONS + 1):
                _emit(f"--- iter {it} ---")

                # memory.read is pure-Python — no thread needed
                hits = memory.read(query, history)
                _emit(f"[memory.read] {len(hits)} hits")

                # history summary
                if history:
                    _emit(f"[agent6] history so far ({len(history)} entries):")
                    for h in history[-4:]:
                        role = h.get("role", "?")
                        if role == "answer":
                            _emit(f"  [answer] {h.get('text', '')[:150]}")
                        elif role == "tool_result":
                            _emit(f"  [tool:{h.get('tool')}] {h.get('result_descriptor', '')[:100]}")

                # perception calls the gateway — run in thread to free event loop
                _emit("[perception] calling LLM...")
                try:
                    obs = await asyncio.to_thread(
                        perception.observe, query, hits, history, prior_goals, run_id, _emit
                    )
                except Exception as e:
                    _emit(f"[perception] ERROR: {e}")
                    break

                prior_goals = obs.goals
                done_count = sum(1 for g in obs.goals if g.done)
                _emit(f"[perception] {len(obs.goals)} goals, {done_count} done")
                for g in obs.goals:
                    status = "✓" if g.done else "·"
                    _emit(f"  {status} {g.id}: {g.text[:70]}")

                if obs.all_done:
                    _emit("[perception] all goals done — breaking loop")
                    break

                goal = obs.next_unfinished()
                if goal is None:
                    break

                _emit(f"[decision] working on: {goal.text[:80]}")

                attached: list[bytes] = []
                if goal.attach_artifact_id and artifacts.exists(goal.attach_artifact_id):
                    attached = [artifacts.get(goal.attach_artifact_id)]
                    _emit(f"[decision] attached artifact {goal.attach_artifact_id} ({len(attached[0])} bytes)")

                # decision calls the gateway — run in thread
                _emit("[decision] calling LLM...")
                try:
                    out = await asyncio.to_thread(
                        decision.next_step, goal, attached, history, tools, _emit
                    )
                except Exception as e:
                    _emit(f"[decision] ERROR: {e}")
                    break

                if out.answer:
                    _emit(f"[decision] → answer ({len(out.answer)} chars)")
                    _emit(f"[decision] answer: {out.answer[:400]}")
                    history.append({"iter": it, "role": "answer", "text": out.answer, "goal_id": goal.id})

                    # Stall detection: same answer for same goal 2+ times → force done
                    bucket = _stall_tracker.setdefault(goal.id, [])
                    bucket.append(out.answer)
                    if len(bucket) >= 2 and len(set(bucket)) == 1:
                        _emit(f"[agent6] STALL on '{goal.id}': identical answer {len(bucket)}x — force-marking done, moving on")
                        prior_goals = [
                            Goal(
                                id=g.id, text=g.text,
                                done=True if g.id == goal.id else g.done,
                                attach_artifact_id=g.attach_artifact_id,
                            )
                            for g in prior_goals
                        ]
                    continue

                _emit(f"[action] → {out.tool_call.name}({list(out.tool_call.arguments.values())[:2]})")

                try:
                    descriptor, art_id = await action.execute(session, out.tool_call, run_id, goal.id)
                    await asyncio.to_thread(memory.record_outcome, out.tool_call, descriptor, art_id, run_id, goal.id)
                    _emit(f"[action] {descriptor}")
                    if art_id and artifacts.exists(art_id):
                        preview = artifacts.get(art_id).decode("utf-8", errors="replace")
                        _emit(f"[action] result preview ({len(preview)} chars): {preview[:600]}")
                except Exception as e:
                    _emit(f"[action] ERROR: {e}")
                    break

                history.append({
                    "iter": it,
                    "role": "tool_result",
                    "tool": out.tool_call.name,
                    "result_descriptor": descriptor,
                    "artifact_id": art_id,
                })

    final = _final_answer(history)
    _emit("=== FINAL ANSWER ===")
    _emit(final)
    return final


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: uv run python agent6.py \"your query here\"")
        sys.exit(1)
    query = " ".join(sys.argv[1:])
    asyncio.run(run(query))
