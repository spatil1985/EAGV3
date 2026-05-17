"""All Pydantic v2 contracts. No logic lives here."""
from __future__ import annotations

from datetime import datetime
from typing import Any, Literal

from pydantic import BaseModel, model_validator


class MemoryItem(BaseModel):
    id: str
    kind: Literal["fact", "preference", "tool_outcome", "scratchpad"]
    keywords: list[str]
    descriptor: str
    value: dict[str, Any]
    artifact_id: str | None = None
    source: str
    run_id: str
    goal_id: str | None = None
    confidence: float = 1.0
    created_at: datetime = datetime.now()


class Artifact(BaseModel):
    id: str                 # "art:<sha256-prefix>"
    content_type: str
    size_bytes: int
    source: str
    descriptor: str


class Goal(BaseModel):
    id: str
    text: str
    done: bool = False
    attach_artifact_id: str | None = None


class Observation(BaseModel):
    goals: list[Goal]

    @property
    def all_done(self) -> bool:
        return all(g.done for g in self.goals)

    def next_unfinished(self) -> Goal | None:
        for g in self.goals:
            if not g.done:
                return g
        return None


class ToolCall(BaseModel):
    name: str
    arguments: dict[str, Any] = {}


class DecisionOutput(BaseModel):
    answer: str | None = None
    tool_call: ToolCall | None = None

    @model_validator(mode="after")
    def exactly_one(self) -> "DecisionOutput":
        if (self.answer is None) == (self.tool_call is None):
            raise ValueError("DecisionOutput must have exactly one of answer or tool_call")
        return self


class GatewayMessage(BaseModel):
    role: str
    content: str


class GatewayRequest(BaseModel):
    messages: list[GatewayMessage]
    response_format: dict[str, Any] | None = None
    provider: str | None = None
    auto_route: str | None = None
    reasoning: str = "off"


class RouterDecision(BaseModel):
    worker: str
    fallback_used: bool = False


class GatewayResponse(BaseModel):
    content: str | dict
    router_decision: RouterDecision
    reasoning_applied: bool = False
