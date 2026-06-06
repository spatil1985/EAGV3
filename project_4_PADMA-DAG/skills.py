"""Skill catalogue loader. One Skill class parameterised by agent_config.yaml."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "agent_config.yaml"


@dataclass
class Skill:
    name: str
    prompt_path: Path
    tools_allowed: list[str]
    temperature: float
    max_tokens: int
    critic: bool = False
    internal_successors: list[str] = field(default_factory=list)

    def system_prompt(self) -> str:
        return self.prompt_path.read_text(encoding="utf-8")

    def render_prompt(
        self,
        inputs_text: str,
        memory_hits: list | None = None,
        question: str | None = None,
        failure_report: str | None = None,
    ) -> str:
        """Build the user-turn message for this skill."""
        parts: list[str] = []

        if memory_hits:
            lines = ["MEMORY HITS:"]
            for m in memory_hits:
                preview = ""
                val = getattr(m, "value", {})
                if isinstance(val, dict):
                    preview = val.get("chunk", val.get("raw", ""))
                    if hasattr(preview, "__str__"):
                        preview = str(preview)[:400]
                lines.append(f"  [{m.kind}] {m.descriptor} — {preview}")
            parts.append("\n".join(lines))

        if failure_report:
            parts.append(f"FAILURE:\n{failure_report}")

        if question:
            parts.append(f"QUESTION: {question}")

        parts.append(f"INPUTS:\n{inputs_text}")

        return "\n\n".join(parts)


def load_skills() -> dict[str, Skill]:
    """Load all skills from agent_config.yaml. Returns {name: Skill}."""
    raw = yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8"))
    skills: dict[str, Skill] = {}
    for name, cfg in raw.get("skills", {}).items():
        prompt_path = BASE_DIR / cfg["prompt"]
        skills[name] = Skill(
            name=name,
            prompt_path=prompt_path,
            tools_allowed=cfg.get("tools_allowed", []),
            temperature=float(cfg.get("temperature", 0.5)),
            max_tokens=int(cfg.get("max_tokens", 1024)),
            critic=bool(cfg.get("critic", False)),
            internal_successors=cfg.get("internal_successors", []),
        )
    return skills
