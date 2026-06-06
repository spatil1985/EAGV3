"""flow.py — Session 8 DAG-based agent orchestrator.

Replaces the single-loop agent6.py with a directed acyclic graph of skills.
The Planner generates the graph; the Executor runs every node whose predecessors
are complete; independent nodes run concurrently via asyncio.gather.

Usage:
    uv run python flow.py "your query"
    uv run python flow.py --resume <session_id>
"""
from __future__ import annotations

import asyncio
import json
import os
import sys
import time
import uuid
from collections.abc import Callable
from enum import Enum
from pathlib import Path
from typing import Any

import httpx
import networkx as nx
from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from pydantic import BaseModel, model_validator

import memory
import sandbox
from parsers import parse_coder_output, parse_critic_output, parse_planner_output
from persistence import SessionLoadError, SessionStore
from recovery import classify_failure
from skills import Skill, load_skills

load_dotenv()

GATEWAY_V8_URL = os.getenv("GATEWAY_V8_URL", "http://localhost:8108")
MCP_SERVER = Path(__file__).parent / "mcp_server.py"
MAX_NODES = 60
MAX_HOPS = 12   # tool-use hops per skill
MAX_RECOVERY_PER_NODE = 1

_progress_cb: Callable[[str], None] | None = None


def _emit(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", errors="replace").decode("ascii"))
    if _progress_cb:
        try:
            _progress_cb(msg)
        except Exception:
            pass


def _emit_graph(event_type: str, **data: Any) -> None:
    """Emit a structured graph event parseable by the UI."""
    payload = json.dumps({"event": event_type, **data}, default=str)
    _emit(f"GRAPH|{payload}")


# ── Pydantic contracts ──────────────────────────────────────────────────────

class NodeStatus(str, Enum):
    pending = "pending"
    running = "running"
    complete = "complete"
    skipped = "skipped"
    failed = "failed"


class AgentResult(BaseModel):
    output: str
    code: str | None = None       # populated by coder skill
    summary: str | None = None    # populated by coder skill


class NodeState(BaseModel):
    node_id: str
    skill: str
    label: str
    inputs: list[str]
    metadata: dict[str, Any] = {}
    status: NodeStatus = NodeStatus.pending
    result: AgentResult | None = None
    error: str | None = None


class NodeSpec(BaseModel):
    """One node as emitted by the Planner."""
    skill: str
    inputs: list[str]
    metadata: dict[str, Any] = {}

    @model_validator(mode="after")
    def check_label(self) -> "NodeSpec":
        if "label" not in self.metadata:
            self.metadata["label"] = f"node_{uuid.uuid4().hex[:6]}"
        return self


# ── Graph wrapper ───────────────────────────────────────────────────────────

class Graph:
    def __init__(self) -> None:
        self.g: nx.DiGraph = nx.DiGraph()
        self._label_to_id: dict[str, str] = {}  # label → node_id

    def add_node(self, state: NodeState) -> None:
        self.g.add_node(state.node_id, state=state)
        self._label_to_id[state.label] = state.node_id

    def add_edge(self, from_id: str, to_id: str) -> None:
        self.g.add_edge(from_id, to_id)

    def get_node(self, node_id: str) -> NodeState:
        return self.g.nodes[node_id]["state"]

    def set_node(self, state: NodeState) -> None:
        self.g.nodes[state.node_id]["state"] = state

    def resolve_label(self, label: str) -> str | None:
        return self._label_to_id.get(label)

    def ready_nodes(self) -> list[NodeState]:
        """Nodes whose every predecessor is complete or skipped."""
        ready = []
        for nid in self.g.nodes:
            state = self.get_node(nid)
            if state.status != NodeStatus.pending:
                continue
            preds = list(self.g.predecessors(nid))
            if all(
                self.get_node(p).status in (NodeStatus.complete, NodeStatus.skipped)
                for p in preds
            ):
                ready.append(state)
        return ready

    def is_all_done(self) -> bool:
        return all(
            self.get_node(n).status in (NodeStatus.complete, NodeStatus.skipped, NodeStatus.failed)
            for n in self.g.nodes
        )

    def all_nodes(self) -> list[NodeState]:
        return [self.get_node(n) for n in self.g.nodes]

    def formatter_output(self) -> str | None:
        for nid in self.g.nodes:
            state = self.get_node(nid)
            if state.skill == "formatter" and state.status == NodeStatus.complete:
                return state.result.output if state.result else None
        return None

    def node_count(self) -> int:
        return len(self.g.nodes)

    def extend_from_planner(
        self,
        planner_id: str,
        nodes_spec: list[dict],
        skills_cfg: dict[str, Skill],
        store: SessionStore,
        counter: list[int],
    ) -> list[NodeState]:
        """Add nodes from a Planner output. Returns new NodeState list."""
        spec_list = [NodeSpec(**n) for n in nodes_spec]
        new_states: list[NodeState] = []
        label_to_new_id: dict[str, str] = {}

        for spec in spec_list:
            counter[0] += 1
            nid = f"n_{counter[0]:03d}"
            label = spec.metadata.get("label", nid)
            state = NodeState(
                node_id=nid,
                skill=spec.skill,
                label=label,
                inputs=spec.inputs,
                metadata=spec.metadata,
            )
            self.add_node(state)
            label_to_new_id[label] = nid
            new_states.append(state)
            store.save_node(state)
            _emit_graph("node_added", id=nid, skill=spec.skill,
                        label=spec.metadata.get("label", nid), status="pending")

        # Rebuild label→id index with newly added nodes
        for label, nid in label_to_new_id.items():
            self._label_to_id[label] = nid

        # Wire edges
        for state in new_states:
            for inp in state.inputs:
                if inp.startswith("n:"):
                    src_label = inp[2:]
                    src_id = self._label_to_id.get(src_label)
                    if src_id and src_id != state.node_id:
                        skill_obj = skills_cfg.get(
                            self.get_node(src_id).skill if src_id in self.g else ""
                        )
                        # Auto-insert critic between distiller and its successors
                        if skill_obj and skill_obj.critic:
                            counter[0] += 1
                            critic_id = f"n_{counter[0]:03d}"
                            critic_label = f"critic_{src_label}_{state.label}"
                            critic_state = NodeState(
                                node_id=critic_id,
                                skill="critic",
                                label=critic_label,
                                inputs=[f"n:{src_label}"],
                                metadata={
                                    "label": critic_label,
                                    "question": state.metadata.get("question", ""),
                                    "target_node_id": state.node_id,
                                },
                            )
                            self.add_node(critic_state)
                            self._label_to_id[critic_label] = critic_id
                            self.add_edge(src_id, critic_id)
                            self.add_edge(critic_id, state.node_id)
                            store.save_node(critic_state)
                            _emit_graph("node_added", id=critic_id, skill="critic",
                                        label=critic_label, status="pending")
                            _emit_graph("edge_added", from_id=src_id, to_id=critic_id)
                            _emit_graph("edge_added", from_id=critic_id, to_id=state.node_id)
                        else:
                            self.add_edge(src_id, state.node_id)
                            _emit_graph("edge_added", from_id=src_id, to_id=state.node_id)
            # Planner itself is a predecessor for first-level nodes
            if not state.inputs or state.inputs == ["USER_QUERY"]:
                if planner_id in self.g and state.node_id != planner_id:
                    self.add_edge(planner_id, state.node_id)
                    _emit_graph("edge_added", from_id=planner_id, to_id=state.node_id)

        store.save_graph(self.g)
        return new_states


# ── Gateway client ──────────────────────────────────────────────────────────

def _call_gateway(
    messages: list[dict],
    skill: Skill,
    session_id: str,
    response_format: dict | None = None,
) -> str | dict:
    payload: dict[str, Any] = {
        "messages": messages,
        "agent": skill.name,
        "session": session_id,
        "temperature": skill.temperature,
        "max_tokens": skill.max_tokens,
    }
    if response_format:
        payload["response_format"] = response_format

    resp = httpx.post(
        f"{GATEWAY_V8_URL}/v1/chat",
        json=payload,
        timeout=90,
    )
    resp.raise_for_status()
    return resp.json().get("content", "")


def _parse_function_call(content: str | dict) -> dict | None:
    """Return {"name": ..., "arguments": ...} or None."""
    if isinstance(content, dict):
        if "FUNCTION_CALL" in content:
            return content["FUNCTION_CALL"]
        return None

    text = str(content).strip()
    if text.startswith("```"):
        nl = text.find("\n")
        if nl != -1:
            text = text[nl + 1:]
        if text.rstrip().endswith("```"):
            text = text.rstrip()[:-3].rstrip()
        text = text.strip()
    try:
        parsed = json.loads(text)
        if "FUNCTION_CALL" in parsed:
            return parsed["FUNCTION_CALL"]
    except (json.JSONDecodeError, TypeError):
        pass
    return None


def _to_str(content: str | dict) -> str:
    if isinstance(content, dict):
        return json.dumps(content)
    return str(content)


# ── Tool-use loop ───────────────────────────────────────────────────────────

async def _run_skill_with_tools(
    skill: Skill,
    system_prompt: str,
    user_msg: str,
    mcp_session: ClientSession,
    session_id: str,
) -> str:
    """Multi-turn tool-use loop for skills that have tools_allowed."""
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_msg},
    ]
    for hop in range(MAX_HOPS):
        content = await asyncio.to_thread(_call_gateway, messages, skill, session_id)
        fc = _parse_function_call(content)
        if fc is None:
            return _to_str(content)

        tool_name = fc.get("name", "")
        tool_args = fc.get("arguments", {})

        if tool_name not in skill.tools_allowed:
            _emit(f"[skill:{skill.name}] tool {tool_name} not in tools_allowed — treating as answer")
            return _to_str(content)

        _emit(f"[skill:{skill.name}] hop {hop+1}: calling {tool_name}({tool_args})")
        try:
            result = await mcp_session.call_tool(tool_name, tool_args)
            raw = ""
            if hasattr(result, "content"):
                for block in result.content:
                    if hasattr(block, "text"):
                        raw += block.text
            else:
                raw = str(result)
        except Exception as exc:
            raw = f"Tool error: {exc}"

        _emit(f"[skill:{skill.name}] tool result ({len(raw)} chars): {raw[:200]}")
        messages.append({"role": "model", "content": _to_str(content)})
        messages.append({"role": "user", "content": f"Tool result for {tool_name}:\n{raw}"})

    return _to_str(messages[-1].get("content", ""))


# ── Node execution ──────────────────────────────────────────────────────────

async def _run_one(
    state: NodeState,
    graph: Graph,
    query: str,
    memory_hits: list,
    mcp_session: ClientSession,
    skills_cfg: dict[str, Skill],
    session_id: str,
) -> AgentResult:
    """Execute one node and return its AgentResult."""
    skill = skills_cfg.get(state.skill)
    if skill is None:
        raise ValueError(f"Unknown skill: {state.skill}")

    # Special dispatch: sandbox_executor
    if state.skill == "sandbox_executor":
        return await _run_sandbox(state, graph)

    # Resolve inputs to text
    inputs_parts: list[str] = []
    for inp in state.inputs:
        if inp == "USER_QUERY":
            inputs_parts.append(f"USER QUERY: {query}")
        elif inp.startswith("n:"):
            label = inp[2:]
            src_id = graph.resolve_label(label)
            if src_id:
                src_state = graph.get_node(src_id)
                if src_state.result:
                    inputs_parts.append(
                        f"[{src_state.skill}:{label}]\n{src_state.result.output}"
                    )
        elif inp.startswith("art:"):
            try:
                import artifacts
                data = artifacts.get(inp)
                inputs_parts.append(
                    f"[artifact:{inp}]\n{data.decode('utf-8', errors='replace')[:3000]}"
                )
            except Exception:
                inputs_parts.append(f"[artifact:{inp}] (not found)")
        elif inp.startswith("FAILURE:") or "FAILURE" in inp:
            inputs_parts.append(inp)

    inputs_text = "\n\n".join(inputs_parts) if inputs_parts else query

    question = state.metadata.get("question", "")
    failure = state.metadata.get("FAILURE", "")

    system_prompt = skill.system_prompt()
    user_msg = skill.render_prompt(
        inputs_text=inputs_text,
        memory_hits=memory_hits if state.skill in ("planner", "retriever") else None,
        question=question or None,
        failure_report=failure or None,
    )

    _emit(f"[skill:{skill.name}] node={state.node_id} label={state.label}")

    if skill.tools_allowed:
        output = await _run_skill_with_tools(
            skill, system_prompt, user_msg, mcp_session, session_id
        )
    else:
        # Single gateway call
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_msg},
        ]
        raw = await asyncio.to_thread(_call_gateway, messages, skill, session_id)
        output = _to_str(raw)

    # Post-process coder output: parse {"code": ..., "summary": ...}
    if state.skill == "coder":
        summary, code, _ = parse_coder_output(output)
        return AgentResult(output=summary, code=code, summary=summary)

    # Critic: extract verdict
    if state.skill == "critic":
        out_str, rationale = parse_critic_output(output, inputs_text)
        return AgentResult(output=out_str, summary=rationale)

    return AgentResult(output=output.strip())


async def _run_sandbox(state: NodeState, graph: Graph) -> AgentResult:
    """Run coder's code field in sandbox subprocess."""
    # Find the coder predecessor
    code = ""
    for pred_id in graph.g.predecessors(state.node_id):
        pred = graph.get_node(pred_id)
        if pred.skill == "coder" and pred.result and pred.result.code:
            code = pred.result.code
            break

    if not code:
        return AgentResult(output="sandbox_executor: no code found from coder predecessor")

    _emit(f"[sandbox] running code ({len(code)} chars)")
    result = await asyncio.to_thread(sandbox.run_code, code)
    summary = (
        f"exit_code={result['exit_code']} "
        f"elapsed={result['elapsed_sec']}s\n"
        f"stdout: {result['stdout'][:500]}"
    )
    if result["stderr"]:
        summary += f"\nstderr: {result['stderr'][:200]}"
    _emit(f"[sandbox] {summary[:300]}")
    return AgentResult(output=summary)




# ── Executor ────────────────────────────────────────────────────────────────

class Executor:
    def __init__(
        self,
        session_id: str,
        store: SessionStore,
        skills_cfg: dict[str, Skill],
        mcp_session: ClientSession,
    ) -> None:
        self.sid = session_id
        self.store = store
        self.skills = skills_cfg
        self.mcp = mcp_session
        self._recovery_count: dict[str, int] = {}
        self._counter = [0]

    async def execute(self, graph: Graph, query: str, memory_hits: list) -> str:
        while not graph.is_all_done():
            if graph.node_count() >= MAX_NODES:
                _emit(f"[executor] MAX_NODES={MAX_NODES} reached — stopping")
                break

            ready = graph.ready_nodes()
            if not ready:
                pending = [n for n in graph.all_nodes() if n.status == NodeStatus.pending]
                if not pending:
                    break
                _emit(f"[executor] deadlock? {len(pending)} pending but 0 ready — stopping")
                break

            _emit(f"[executor] {len(ready)} ready node(s): {[n.node_id for n in ready]}")

            # Mark all ready nodes as running
            for state in ready:
                state.status = NodeStatus.running
                graph.set_node(state)
                self.store.save_node(state)
                _emit_graph("node_status", id=state.node_id, status="running")

            start_times = {s.node_id: time.monotonic() for s in ready}

            # Run concurrently
            results = await asyncio.gather(
                *[
                    _run_one(
                        state, graph, query, memory_hits,
                        self.mcp, self.skills, self.sid,
                    )
                    for state in ready
                ],
                return_exceptions=True,
            )

            for state, result in zip(ready, results):
                elapsed = round(time.monotonic() - start_times[state.node_id], 2)
                if isinstance(result, Exception):
                    await self._handle_failure(state, result, graph, query)
                else:
                    await self._handle_success(state, result, graph, query, elapsed)

            self.store.save_graph(graph.g)

        answer = graph.formatter_output()
        if answer:
            # Strip CRITIC_PASS prefix if formatter received critic output
            if answer.startswith("CRITIC_PASS\n"):
                answer = answer[len("CRITIC_PASS\n"):]
        return answer or _fallback_answer(graph)

    async def _handle_success(
        self,
        state: NodeState,
        result: AgentResult,
        graph: Graph,
        query: str,
        elapsed: float,
    ) -> None:
        state.status = NodeStatus.complete
        state.result = result
        graph.set_node(state)
        self.store.save_node(state)
        _emit(f"[executor] {state.node_id} ({state.skill}) complete in {elapsed}s")
        _emit(f"[executor] output: {result.output[:300]}")
        _emit_graph("node_status", id=state.node_id, status="complete",
                    elapsed=elapsed, preview=result.output[:120])

        # Planner: extend graph
        if state.skill == "planner":
            await self._extend_from_planner(state, graph, query)

        # Coder: auto-add sandbox_executor sibling
        elif state.skill == "coder":
            skill_obj = self.skills.get("coder")
            if skill_obj and "sandbox_executor" in skill_obj.internal_successors:
                await self._add_sandbox(state, graph)

        # Critic: handle verdict
        elif state.skill == "critic":
            await self._handle_critic(state, graph, query)

    async def _handle_failure(
        self, state: NodeState, exc: Exception, graph: Graph, query: str
    ) -> None:
        error_str = str(exc)
        kind = classify_failure(error_str)
        state.status = NodeStatus.failed
        state.error = error_str
        graph.set_node(state)
        self.store.save_node(state)
        _emit(f"[executor] {state.node_id} ({state.skill}) FAILED [{kind}]: {error_str[:200]}")
        _emit_graph("node_status", id=state.node_id, status="failed", preview=error_str[:120])

        if kind == "transient":
            _emit("[executor] transient error — surfacing to user, no recovery")
        elif kind == "validation_error":
            _emit("[executor] validation error — no recovery (fix the prompt)")
        else:
            # upstream_failure: one recovery planner
            if state.skill == "planner":
                _emit("[executor] planner itself failed — no planner-of-planner")
                return
            count = self._recovery_count.get(state.node_id, 0)
            if count >= MAX_RECOVERY_PER_NODE:
                _emit(f"[executor] recovery cap reached for {state.node_id}")
                return
            self._recovery_count[state.node_id] = count + 1
            self._counter[0] += 1
            rec_id = f"n_{self._counter[0]:03d}"
            rec_state = NodeState(
                node_id=rec_id,
                skill="planner",
                label=f"recovery_{state.node_id}",
                inputs=["USER_QUERY"],
                metadata={
                    "label": f"recovery_{state.node_id}",
                    "FAILURE": f"Node {state.node_id} ({state.skill}) failed: {error_str[:300]}",
                    "failed_node_id": state.node_id,
                },
            )
            graph.add_node(rec_state)
            self.store.save_node(rec_state)
            _emit(f"[executor] queued recovery planner {rec_id}")

    async def _extend_from_planner(
        self, planner_state: NodeState, graph: Graph, query: str
    ) -> None:
        try:
            rationale, nodes_spec = parse_planner_output(planner_state.result.output)
        except Exception as exc:
            _emit(f"[executor] planner output parse error: {exc}")
            return

        _emit(f"[executor] planner rationale: {rationale}")
        _emit(f"[executor] adding {len(nodes_spec)} nodes from planner")

        graph.extend_from_planner(
            planner_state.node_id, nodes_spec, self.skills, self.store, self._counter
        )
        self.store.save_graph(graph.g)

    async def _add_sandbox(self, coder_state: NodeState, graph: Graph) -> None:
        self._counter[0] += 1
        sb_id = f"n_{self._counter[0]:03d}"
        sb_label = f"sandbox_{coder_state.label}"
        sb_state = NodeState(
            node_id=sb_id,
            skill="sandbox_executor",
            label=sb_label,
            inputs=[f"n:{coder_state.label}"],
            metadata={"label": sb_label},
        )
        graph.add_node(sb_state)
        graph.add_edge(coder_state.node_id, sb_id)
        self.store.save_node(sb_state)
        _emit(f"[executor] auto-added sandbox_executor {sb_id}")
        _emit_graph("node_added", id=sb_id, skill="sandbox_executor",
                    label=sb_label, status="pending")
        _emit_graph("edge_added", from_id=coder_state.node_id, to_id=sb_id)

    async def _handle_critic(
        self, critic_state: NodeState, graph: Graph, query: str
    ) -> None:
        output = critic_state.result.output if critic_state.result else ""
        verdict = "pass" if output.startswith("CRITIC_PASS") else "fail"
        _emit(f"[executor] critic {critic_state.node_id} verdict: {verdict}")

        if verdict == "pass":
            return  # successors will run normally

        # Fail: skip successors, queue recovery planner
        target_node_id = critic_state.metadata.get("target_node_id")
        for child_id in list(graph.g.successors(critic_state.node_id)):
            child = graph.get_node(child_id)
            if child.status == NodeStatus.pending:
                child.status = NodeStatus.skipped
                graph.set_node(child)
                self.store.save_node(child)
                _emit(f"[executor] skipped {child_id} due to critic fail")

        cap_key = f"critic_{critic_state.node_id}"
        if self._recovery_count.get(cap_key, 0) >= MAX_RECOVERY_PER_NODE:
            _emit(f"[executor] critic recovery cap reached for {critic_state.node_id}")
            return
        self._recovery_count[cap_key] = self._recovery_count.get(cap_key, 0) + 1

        self._counter[0] += 1
        rec_id = f"n_{self._counter[0]:03d}"
        rationale = critic_state.result.summary or "Critic returned fail"
        rec_state = NodeState(
            node_id=rec_id,
            skill="planner",
            label=f"crit_recovery_{critic_state.node_id}",
            inputs=["USER_QUERY"],
            metadata={
                "label": f"crit_recovery_{critic_state.node_id}",
                "FAILURE": f"Critic failed on node {target_node_id}: {rationale}",
            },
        )
        graph.add_node(rec_state)
        self.store.save_node(rec_state)
        _emit(f"[executor] queued critic recovery planner {rec_id}")


def _fallback_answer(graph: Graph) -> str:
    # Fall back to the last complete non-planner node's output
    for state in reversed(graph.all_nodes()):
        if state.status == NodeStatus.complete and state.skill not in ("planner", "sandbox_executor"):
            return state.result.output if state.result else "Task completed."
    return "Task completed."


# ── Main entry points ───────────────────────────────────────────────────────

async def run(
    query: str,
    session_id: str | None = None,
    on_progress: Callable[[str], None] | None = None,
) -> str:
    global _progress_cb
    _progress_cb = on_progress

    sid = session_id or f"s8-{uuid.uuid4().hex[:8]}"
    _emit(f"[flow] session_id={sid}")
    _emit(f"[flow] query: {query}")

    store = SessionStore(sid)
    store.save_query(query)

    # Memory read at session start — shared across all skills
    hits = memory.read(query)
    _emit(f"[flow] memory hits: {len(hits)}")
    try:
        await asyncio.to_thread(memory.remember, query, "user_query", sid)
    except Exception as e:
        _emit(f"[flow] memory.remember error: {e}")

    skills_cfg = load_skills()
    graph = Graph()

    server_params = StdioServerParameters(command=sys.executable, args=[str(MCP_SERVER)])

    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as mcp_session:
            await mcp_session.initialize()
            tools = await mcp_session.list_tools()
            _emit(f"[flow] MCP tools: {len(tools.tools)}")

            executor = Executor(sid, store, skills_cfg, mcp_session)

            # Seed: initial Planner node
            executor._counter[0] = 1
            planner_state = NodeState(
                node_id="n_001",
                skill="planner",
                label="planner_0",
                inputs=["USER_QUERY"],
                metadata={"label": "planner_0"},
            )
            graph.add_node(planner_state)
            store.save_node(planner_state)
            store.save_graph(graph.g)
            _emit_graph("node_added", id="n_001", skill="planner",
                        label="planner_0", status="pending")

            answer = await executor.execute(graph, query, hits)

    _emit("=== FINAL ANSWER ===")
    _emit(answer)
    return answer


async def resume(session_id: str, on_progress: Callable[[str], None] | None = None) -> str:
    global _progress_cb
    _progress_cb = on_progress

    _emit(f"[flow] resuming session_id={session_id}")
    store = SessionStore(session_id)
    query = store.load_query()
    _emit(f"[flow] query: {query}")

    g_raw = store.load_graph()
    graph = Graph()
    graph.g = g_raw

    # Rebuild label→id index
    for nid in g_raw.nodes:
        state_data = g_raw.nodes[nid].get("state")
        if isinstance(state_data, dict) and state_data.get("_state_typed"):
            state_data = {k: v for k, v in state_data.items() if k != "_state_typed"}
            state = NodeState.model_validate(state_data)
            g_raw.nodes[nid]["state"] = state
        elif isinstance(state_data, NodeState):
            state = state_data
        else:
            continue
        graph._label_to_id[state.label] = nid
        # Reset running → pending
        if state.status == NodeStatus.running:
            state.status = NodeStatus.pending
            graph.set_node(state)
            store.save_node(state)

    hits = memory.read(query)
    skills_cfg = load_skills()

    # Find max node counter
    max_num = 0
    for nid in g_raw.nodes:
        try:
            max_num = max(max_num, int(nid.split("_")[1]))
        except (IndexError, ValueError):
            pass

    server_params = StdioServerParameters(command=sys.executable, args=[str(MCP_SERVER)])
    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as mcp_session:
            await mcp_session.initialize()
            executor = Executor(session_id, store, skills_cfg, mcp_session)
            executor._counter[0] = max_num
            answer = await executor.execute(graph, query, hits)

    _emit("=== FINAL ANSWER (resumed) ===")
    _emit(answer)
    return answer


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('Usage: python flow.py "your query"')
        print('       python flow.py --resume <session_id>')
        sys.exit(1)

    if sys.argv[1] == "--resume":
        if len(sys.argv) < 3:
            print("Usage: python flow.py --resume <session_id>")
            sys.exit(1)
        asyncio.run(resume(sys.argv[2]))
    else:
        q = " ".join(sys.argv[1:])
        asyncio.run(run(q))
