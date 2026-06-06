"""Action role — pure MCP dispatcher. No LLM calls."""
from __future__ import annotations

import json

import artifacts as artifact_store
from schemas import Artifact, ToolCall


async def execute(
    session,
    tool_call: ToolCall,
    run_id: str,
    goal_id: str | None = None,
) -> tuple[str, str | None]:
    """
    Dispatch tool_call via MCP session.
    Returns (descriptor, artifact_id | None).
    """
    result = await session.call_tool(tool_call.name, tool_call.arguments)

    # Extract text content from MCP result
    raw_text = ""
    if hasattr(result, "content"):
        for block in result.content:
            if hasattr(block, "text"):
                raw_text += block.text
    else:
        raw_text = str(result)

    # Guard against art: prefix hallucination
    if raw_text.startswith("art:"):
        raw_text = f"[result] {raw_text}"

    # Store in artifact store
    data = raw_text.encode("utf-8")
    artifact = artifact_store.put(
        data=data,
        content_type="text/plain",
        source=tool_call.name,
        descriptor=f"{tool_call.name}({list(tool_call.arguments.values())[:1]}): {raw_text[:60]}",
    )

    descriptor = f"{tool_call.name} result ({len(data)} bytes) → {artifact.id}"
    return descriptor, artifact.id
