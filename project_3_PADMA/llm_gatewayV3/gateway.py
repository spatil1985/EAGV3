"""LLM Gateway V3 — FastAPI proxy for Gemini. Runs at http://localhost:8101."""
import json
import os
from typing import Any

import google.generativeai as genai
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

load_dotenv()
genai.configure(api_key=os.environ["GEMINI_API_KEY"])

app = FastAPI(title="LLM Gateway V3")

GEMINI_MODEL = "gemini-2.5-flash"


class Message(BaseModel):
    role: str   # "user" | "model" | "system"
    content: str


class ResponseFormat(BaseModel):
    type: str = "json_schema"
    schema_: dict[str, Any] | None = None

    model_config = {"populate_by_name": True}


class ChatRequest(BaseModel):
    messages: list[Message]
    response_format: ResponseFormat | None = None
    provider: str | None = None      # "g" forces Gemini (same here regardless)
    auto_route: str | None = None    # "memory" | "decision" | "perception"
    reasoning: str | None = "off"


class RouterDecision(BaseModel):
    worker: str
    fallback_used: bool = False


class ChatResponse(BaseModel):
    content: str | dict
    router_decision: RouterDecision
    reasoning_applied: bool = False


def _to_gemini_history(messages: list[Message]) -> tuple[str | None, list[dict]]:
    """Split system prompt from conversation history for Gemini."""
    system = None
    history = []
    for m in messages:
        if m.role == "system":
            system = m.content
        elif m.role in ("user", "model"):
            history.append({"role": m.role, "parts": [m.content]})
    return system, history


@app.post("/chat", response_model=ChatResponse)
async def chat(req: ChatRequest) -> ChatResponse:
    system_prompt, history = _to_gemini_history(req.messages)

    generation_config: dict[str, Any] = {}
    if req.response_format and req.response_format.schema_:
        generation_config["response_mime_type"] = "application/json"
        generation_config["response_schema"] = req.response_format.schema_

    try:
        model = genai.GenerativeModel(
            model_name=GEMINI_MODEL,
            system_instruction=system_prompt,
            generation_config=generation_config if generation_config else None,
        )
        # Last message is the current user turn; rest is history
        if len(history) > 1:
            chat_session = model.start_chat(history=history[:-1])
            last_content = history[-1]["parts"][0]
            response = chat_session.send_message(last_content)
        elif history:
            response = model.generate_content(history[-1]["parts"][0])
        else:
            raise HTTPException(status_code=400, detail="No messages provided")

        raw = response.text
        content: str | dict = raw

        if req.response_format and req.response_format.schema_:
            try:
                content = json.loads(raw)
            except json.JSONDecodeError:
                content = raw

        return ChatResponse(
            content=content,
            router_decision=RouterDecision(worker=GEMINI_MODEL, fallback_used=False),
            reasoning_applied=False,
        )

    except Exception as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc


@app.get("/health")
async def health() -> dict:
    return {"status": "ok", "model": GEMINI_MODEL}
