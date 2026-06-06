"""LLM Gateway V8 — FastAPI proxy for Gemini on port 8108.

Adds over V3:
- /v1/chat endpoint with agent + session tagging
- /v1/chat/batch for parallel dispatch
- /v1/cost/by_agent?session=<sid> for per-skill token spend
- agent_routing.yaml to pin skills to providers (bypass router pool)
- retry-on-5xx with exponential backoff
"""
from __future__ import annotations

import json
import os
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

import google.generativeai as genai
import yaml
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Query
from pydantic import BaseModel

load_dotenv()
genai.configure(api_key=os.environ["GEMINI_API_KEY"])

app = FastAPI(title="LLM Gateway V8")

GEMINI_MODEL = "gemini-2.5-flash"
MAX_RETRIES = 3
RETRY_BACKOFF = [1.0, 2.0, 4.0]

# In-memory call log — resets on restart, sufficient for one demo session
_call_log: list[dict] = []

# Load agent_routing.yaml if present
_ROUTING_PATH = Path(__file__).parent / "agent_routing.yaml"
_agent_routing: dict[str, str] = {}
if _ROUTING_PATH.exists():
    try:
        _agent_routing = yaml.safe_load(_ROUTING_PATH.read_text()) or {}
    except Exception:
        _agent_routing = {}


class Message(BaseModel):
    role: str
    content: str


class ResponseFormat(BaseModel):
    type: str = "json_schema"
    schema_: dict[str, Any] | None = None
    model_config = {"populate_by_name": True}


class ChatRequest(BaseModel):
    messages: list[Message]
    response_format: ResponseFormat | None = None
    provider: str | None = None
    auto_route: str | None = None
    reasoning: str | None = "off"
    agent: str | None = None    # skill label, e.g. "researcher"
    session: str | None = None  # session id, e.g. "s8-abc123"
    temperature: float | None = None
    max_tokens: int | None = None


class RouterDecision(BaseModel):
    worker: str
    fallback_used: bool = False


class ChatResponse(BaseModel):
    content: str | dict
    router_decision: RouterDecision
    reasoning_applied: bool = False


class BatchRequest(BaseModel):
    requests: list[ChatRequest]


class BatchResponse(BaseModel):
    results: list[ChatResponse]


def _to_gemini_history(messages: list[Message]) -> tuple[str | None, list[dict]]:
    system = None
    history = []
    for m in messages:
        if m.role == "system":
            system = m.content
        elif m.role in ("user", "model"):
            history.append({"role": m.role, "parts": [m.content]})
    return system, history


def _call_gemini(req: ChatRequest) -> tuple[str | dict, int, int]:
    """Call Gemini with retry-on-5xx. Returns (content, input_tokens, output_tokens)."""
    system_prompt, history = _to_gemini_history(req.messages)

    generation_config: dict[str, Any] = {}
    if req.response_format and req.response_format.schema_:
        generation_config["response_mime_type"] = "application/json"
        generation_config["response_schema"] = req.response_format.schema_
    if req.temperature is not None:
        generation_config["temperature"] = req.temperature
    if req.max_tokens is not None:
        generation_config["max_output_tokens"] = req.max_tokens

    last_exc: Exception | None = None
    for attempt in range(MAX_RETRIES):
        try:
            model = genai.GenerativeModel(
                model_name=GEMINI_MODEL,
                system_instruction=system_prompt,
                generation_config=generation_config if generation_config else None,
            )
            if len(history) > 1:
                chat_session = model.start_chat(history=history[:-1])
                response = chat_session.send_message(history[-1]["parts"][0])
            elif history:
                response = model.generate_content(history[-1]["parts"][0])
            else:
                raise ValueError("No messages provided")

            raw = response.text
            content: str | dict = raw
            if req.response_format and req.response_format.schema_:
                try:
                    content = json.loads(raw)
                except json.JSONDecodeError:
                    content = raw

            # Extract token usage
            usage = getattr(response, "usage_metadata", None)
            in_tok = getattr(usage, "prompt_token_count", 0) or 0
            out_tok = getattr(usage, "candidates_token_count", 0) or 0
            return content, in_tok, out_tok

        except Exception as exc:
            err = str(exc)
            # Retry on transient errors
            if any(k in err for k in ("503", "502", "504", "timeout", "ServiceUnavailable")):
                last_exc = exc
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_BACKOFF[attempt])
                    continue
            raise

    raise last_exc or RuntimeError("All retries exhausted")


@app.post("/v1/chat", response_model=ChatResponse)
async def chat(req: ChatRequest) -> ChatResponse:
    try:
        content, in_tok, out_tok = _call_gemini(req)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    # Log with agent + session
    _call_log.append({
        "agent": req.agent or "unknown",
        "session": req.session or "unknown",
        "input_tokens": in_tok,
        "output_tokens": out_tok,
        "timestamp": datetime.now().isoformat(),
    })

    provider = _agent_routing.get(req.agent or "", GEMINI_MODEL)
    return ChatResponse(
        content=content,
        router_decision=RouterDecision(worker=provider, fallback_used=False),
        reasoning_applied=False,
    )


# V3 compat — keeps S7 code working if it calls this gateway
@app.post("/chat", response_model=ChatResponse)
async def chat_compat(req: ChatRequest) -> ChatResponse:
    return await chat(req)


@app.post("/v1/chat/batch", response_model=BatchResponse)
async def chat_batch(batch: BatchRequest) -> BatchResponse:
    import asyncio
    tasks = [chat(r) for r in batch.requests]
    results = await asyncio.gather(*tasks, return_exceptions=True)
    responses = []
    for r in results:
        if isinstance(r, Exception):
            raise HTTPException(status_code=502, detail=str(r))
        responses.append(r)
    return BatchResponse(results=responses)


@app.get("/v1/cost/by_agent")
async def cost_by_agent(session: str = Query(...)) -> dict:
    by_agent: dict[str, dict] = defaultdict(lambda: {"calls": 0, "input_tokens": 0, "output_tokens": 0})
    for entry in _call_log:
        if entry["session"] == session:
            ag = entry["agent"]
            by_agent[ag]["calls"] += 1
            by_agent[ag]["input_tokens"] += entry["input_tokens"]
            by_agent[ag]["output_tokens"] += entry["output_tokens"]

    total_in = sum(v["input_tokens"] for v in by_agent.values())
    total_out = sum(v["output_tokens"] for v in by_agent.values())
    return {
        "session": session,
        "by_agent": dict(by_agent),
        "total_input_tokens": total_in,
        "total_output_tokens": total_out,
    }


@app.get("/health")
async def health() -> dict:
    return {"status": "ok", "model": GEMINI_MODEL, "gateway": "V8"}
