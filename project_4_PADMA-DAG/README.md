# PADMA — Planner Agent with Durable Memory Architecture

**EAG V3 · Session 6 Assignment** — Agentic Basic Architecture

---

## Table of Contents

1. [What is PADMA?](#1-what-is-padma)
2. [Architecture Overview](#2-architecture-overview)
3. [The Four Cognitive Roles](#3-the-four-cognitive-roles)
4. [Supporting Components](#4-supporting-components)
5. [Project Structure](#5-project-structure)
6. [Pydantic Contracts (schemas.py)](#6-pydantic-contracts-schemaspy)
7. [The Control Flow](#7-the-control-flow)
8. [Prerequisites & Setup](#8-prerequisites--setup)
9. [Running the Services](#9-running-the-services)
10. [The Four Assignment Queries](#10-the-four-assignment-queries)
11. [The Web UI](#11-the-web-ui)
12. [LLM Gateway V3](#12-llm-gateway-v3)
13. [MCP Server & Tools](#13-mcp-server--tools)
14. [Memory System](#14-memory-system)
15. [Artifact Store](#15-artifact-store)
16. [Design Constraints](#16-design-constraints)
17. [Dependency Reference](#17-dependency-reference)

---

## 1. What is PADMA?

PADMA is a **fully typed, role-decomposed agentic loop** built from scratch in Python.
It decomposes a user query into goals, iterates over those goals using four specialised
cognitive roles, dispatches MCP tools, stores results as content-addressable artifacts,
and maintains durable memory across sessions.

Key properties:
- **No third-party agent frameworks** (no LangChain, LangGraph, CrewAI)
- **Pydantic v2 at every boundary** — every role input/output is a typed model
- **All LLM calls route through the Gateway** — roles never call providers directly
- **MCP stdio transport** for tool dispatch — tools are not re-implemented in the agent
- **Persistent memory** across runs via `state/memory.json`
- **Content-addressable artifact store** for raw bytes produced by tools

---

## 2. Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                          User / Web UI                               │
└──────────────────────────────┬──────────────────────────────────────┘
                               │  query (str)
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                          agent6.py                                   │
│                        Main Loop (≤12 iter)                          │
│                                                                      │
│   memory.remember()  ──► stores durable facts/preferences           │
│                                                                      │
│   ┌──────────────────────────────────────────────────────────────┐  │
│   │  for each iteration:                                          │  │
│   │   1. memory.read()       ──► keyword search over state/      │  │
│   │   2. perception.observe() ─► goal list (LLM · Gemini)        │  │
│   │   3. if all goals done → break                               │  │
│   │   4. decision.next_step() ─► answer OR tool_call (LLM)       │  │
│   │   5. action.execute()    ──► MCP dispatch → ArtifactStore    │  │
│   │   6. memory.record_outcome()                                  │  │
│   └──────────────────────────────────────────────────────────────┘  │
│                                                                      │
│   final_answer_from(history)  ──► str                               │
└─────────────────────────────────────────────────────────────────────┘
         │                │                        │
         ▼                ▼                        ▼
  LLM Gateway V3    MCP Server              state/
  localhost:8101    (stdio)                 ├── memory.json
  (Gemini Flash)    9 tools                 └── artifacts/
```

---

## 3. The Four Cognitive Roles

The architecture decomposes every agent turn into four responsibilities.
Each role has a single well-defined contract and is invoked in a fixed order.

### 3.1 Memory (`memory.py`)

| Property | Detail |
|---|---|
| **LLM call?** | Only for `remember()` classification and `relevant()` re-ranking |
| **Gateway route** | `auto_route="memory"` |
| **Persistence** | `state/memory.json` — survives across runs |

**Three public methods:**

```python
memory.remember(query, source, run_id)
# Classifies the query — stores a durable MemoryItem if it contains a
# fact ("My mom's birthday is 15 May") or preference. Uses the gateway
# to extract keywords, descriptor, and structured value.

memory.read(query, history, kinds, top_k)
# Pure-Python keyword overlap search over all stored items.
# No LLM call for basic recall — tokens from the query are intersected
# with item keywords. Returns top-k MemoryItems by overlap score.

memory.record_outcome(tool_call, result_text, artifact_id, run_id, goal_id)
# Writes a kind="tool_outcome" item after every Action dispatch.
# Captures tool name, arguments, result text, and the artifact handle.
```

**Memory item kinds:**

| Kind | Carries | Example |
|---|---|---|
| `fact` | A durable observed truth | `"John's office is HSR Layout, Bangalore"` |
| `preference` | A user-stated preference | `"User prefers morning meetings"` |
| `tool_outcome` | Record of one MCP dispatch | `fetch_url → art:3f9a2b...` |
| `scratchpad` | Intermediate working note | Planner state during current run |

---

### 3.2 Perception (`perception.py`)

| Property | Detail |
|---|---|
| **LLM call?** | Yes — every iteration |
| **Gateway route** | `provider="g"` (always Gemini) |
| **Input** | query + memory hits + history + prior goal list |
| **Output** | `Observation` (updated goal list with `done` flags) |

Perception is the **orchestrator**. It is the only role that sees the full
run state and decides what work remains.

**Four obligations enforced in the system prompt:**

1. **Decompose** — on iteration 1 (empty prior goals) break the query into 1–5 ordered goals
2. **Mark done** — if history already contains a satisfying action for a goal, set `done=true`
3. **Attach** — if a goal requires reading the raw bytes of a previously fetched resource,
   set `attach_artifact_id` to the `art:...` handle from history
4. **Preserve order** — never reorder or remove goals; only update flags

---

### 3.3 Decision (`decision.py`)

| Property | Detail |
|---|---|
| **LLM call?** | Yes — every iteration |
| **Gateway route** | `auto_route="decision"` |
| **Input** | goal + attached bytes + history + available tools |
| **Output** | `DecisionOutput` — exactly one of: `answer` or `tool_call` |

Decision makes **one choice per call**: answer in plain prose, or call exactly
one MCP tool. It is not allowed to narrate, plan, or pick multiple tools.

The system prompt enforces:
- If answer → respond in plain text (no JSON)
- If tool call → respond with `{"FUNCTION_CALL": {"name": "...", "arguments": {...}}}`
- Never pass `art:` handles as tool arguments
- Only use tools from the provided list

---

### 3.4 Action (`action.py`)

| Property | Detail |
|---|---|
| **LLM call?** | **No** — pure dispatcher |
| **Transport** | MCP stdio session |
| **Input** | `ToolCall` (name + arguments) |
| **Output** | `(descriptor: str, artifact_id: str \| None)` |

Action is the simplest role. It calls `session.call_tool()`, stores the raw
result bytes in the ArtifactStore, and returns a short descriptor and the
artifact handle. No reasoning happens here.

---

## 4. Supporting Components

### ArtifactStore (`artifacts.py`)

A **content-addressable file store** that lives parallel to Memory.

- Raw bytes from tools are stored under `state/artifacts/<sha256>.bin`
- Metadata (type, size, source, descriptor) is stored alongside as `<sha256>.json`
- IDs have the form `art:<first-16-chars-of-sha256>`
- Memory records the artifact handle; Decision receives raw bytes only when
  Perception explicitly attaches them via `attach_artifact_id`
- Deduplicates automatically — same bytes → same ID

```
state/artifacts/
├── 3f9a2b1c4d5e6f7a.bin   ← raw text from fetch_url
├── 3f9a2b1c4d5e6f7a.json  ← Artifact metadata
├── a1b2c3d4e5f6a7b8.bin
└── a1b2c3d4e5f6a7b8.json
```

### LLM Gateway V3 (`llm_gatewayV3/gateway.py`)

See [Section 12](#12-llm-gateway-v3) for full details.

### MCP Server (`mcp_server.py`)

See [Section 13](#13-mcp-server--tools) for full details.

---

## 5. Project Structure

```
project_3_PADMA/
│
├── agent6.py           ← Main loop — wires all 4 roles
├── schemas.py          ← ALL Pydantic v2 contracts (no logic)
├── memory.py           ← Memory role: read / write / remember
├── perception.py       ← Perception role: observe() → Observation
├── decision.py         ← Decision role: next_step() → DecisionOutput
├── action.py           ← Action role: execute() → (descriptor, artifact_id)
├── artifacts.py        ← ArtifactStore: put / get / exists
├── mcp_server.py       ← FastMCP server with 9 tools (stdio transport)
├── ui_server.py        ← FastAPI web UI with SSE streaming
│
├── start.ps1           ← Start gateway + UI, open browser
├── stop.ps1            ← Stop both services (by PID or port)
├── restart.ps1         ← stop.ps1 + start.ps1
│
├── llm_gatewayV3/
│   ├── gateway.py      ← FastAPI LLM proxy (Gemini Flash, port 8101)
│   ├── pyproject.toml
│   └── .env            ← GEMINI_API_KEY (gitignored)
│
├── pyproject.toml      ← Main project deps (uv-managed)
├── .env                ← API keys (gitignored)
├── .gitignore
│
├── state/              ← Runtime state (gitignored)
│   ├── memory.json     ← Durable memory store
│   └── artifacts/      ← Content-addressable byte store
│
└── logs/               ← Service logs (gitignored)
    ├── gateway.log
    ├── gateway_err.log
    ├── ui.log
    └── ui_err.log
```

---

## 6. Pydantic Contracts (`schemas.py`)

Every boundary between roles is a Pydantic v2 model. Raw dicts never cross
role lines. This enables validation, serialisation, and clear failure modes.

```python
class MemoryItem(BaseModel):
    id: str
    kind: Literal["fact", "preference", "tool_outcome", "scratchpad"]
    keywords: list[str]
    descriptor: str          # one short human-readable line
    value: dict[str, Any]    # structured payload
    artifact_id: str | None  # handle into ArtifactStore
    source: str
    run_id: str
    goal_id: str | None
    confidence: float
    created_at: datetime

class Artifact(BaseModel):
    id: str              # "art:<sha256-prefix>"
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
    # .all_done property
    # .next_unfinished() method

class ToolCall(BaseModel):
    name: str
    arguments: dict[str, Any] = {}

class DecisionOutput(BaseModel):
    answer: str | None = None
    tool_call: ToolCall | None = None
    # validator: exactly one must be set
```

---

## 7. The Control Flow

```
agent6.run(query)
│
├── memory.remember(query)          # classify + store durable fact/pref
│
├── async with mcp_session() as session:
│   ├── tools = load_tools(session)
│   │
│   └── for it in range(MAX_ITERATIONS):
│       │
│       ├── hits = memory.read(query, history)
│       │
│       ├── obs = perception.observe(
│       │         query, hits, history, prior_goals, run_id)
│       │   └── returns Observation (goal list + done flags)
│       │
│       ├── if obs.all_done → break
│       │
│       ├── goal = obs.next_unfinished()
│       │
│       ├── attached = []
│       │   if goal.attach_artifact_id:
│       │       attached = [artifacts.get(goal.attach_artifact_id)]
│       │
│       ├── out = decision.next_step(goal, attached, history, tools)
│       │   └── returns DecisionOutput (answer xor tool_call)
│       │
│       ├── if out.answer:
│       │       history.append({role: "answer", text: out.answer})
│       │       continue
│       │
│       └── (tool path)
│           ├── descriptor, art_id = action.execute(session, out.tool_call)
│           ├── memory.record_outcome(out.tool_call, descriptor, art_id)
│           └── history.append({role: "tool_result", ...})
│
└── return final_answer_from(history)
```

**Three key properties of this flow:**

1. **Memory is consulted at the start of every iteration** — not just once at
   the beginning. Facts written in iteration 2 can be recalled in iteration 3.

2. **Artifact attachment is Perception's decision** — Decision cannot request
   raw bytes directly. Perception decides which artifact to attach based on
   the goal and the history.

3. **The loop terminates when Perception marks all goals done** — not when
   Decision returns an answer. Perception re-evaluates every iteration.

---

## 8. Prerequisites & Setup

### Requirements

| Tool | Version | Purpose |
|---|---|---|
| Python | 3.11+ | Runtime |
| `uv` | Any | Dependency management |
| Gemini API key | — | All LLM calls |

### Install `uv`

```powershell
py -m pip install uv
```

Verify:
```powershell
py -m uv --version
```

### Get a Gemini API Key

1. Go to [Google AI Studio](https://aistudio.google.com/app/apikey)
2. Create an API key
3. Add to both `.env` files (see below)

### Configure `.env` files

**`project_3_PADMA/.env`** (for the agent + UI):
```
GEMINI_API_KEY=your_key_here
TAVILY_API_KEY=          # optional
GATEWAY_URL=http://localhost:8101
```

**`project_3_PADMA/llm_gatewayV3/.env`** (for the gateway):
```
GEMINI_API_KEY=your_key_here
```

### Install Dependencies

```powershell
# Gateway
cd llm_gatewayV3
py -m uv sync

# Agent + UI
cd ..
py -m uv sync
```

---

## 9. Running the Services

### Option A — PowerShell Scripts (recommended)

```powershell
cd C:\work\code\EAGV3\project_3_PADMA

.\start.ps1      # starts gateway + UI, waits for health check, opens browser
.\stop.ps1       # stops both services gracefully
.\restart.ps1    # stop then start in one command
```

`start.ps1` behaviour:
- Starts the gateway, polls `http://localhost:8101/health` for up to 15 seconds
- Starts the UI server only after the gateway is confirmed healthy
- Saves process IDs to `.pids` for clean shutdown
- Opens `http://localhost:8000` in your default browser
- Writes logs to `logs/`

### Option B — Manual (two terminals)

**Terminal 1 — Gateway:**
```powershell
cd llm_gatewayV3
py -m uv run uvicorn gateway:app --port 8101
```

**Terminal 2 — UI server:**
```powershell
cd project_3_PADMA
py -m uv run python ui_server.py
```

**Terminal 2 (CLI only, no UI):**
```powershell
py -m uv run python agent6.py "your query here"
```

---

## 10. The Four Assignment Queries

The assignment requires the agent to pass four specific queries that each exercise
a different architectural feature. Run each from a **clean `state/` directory**
(except Query C Run 2, which intentionally reuses state from Run 1).

> **Clean state:** delete `state/memory.json` and `state/artifacts/` before each query.

---

### Query A — Wikipedia Artifact Attach

```
Fetch the Claude Shannon Wikipedia page and tell me his birth date,
death date, and three key contributions to information theory.
```

**What it tests:** The artifact-attach path.

**Expected trace:**
```
iter 1: [decision] → fetch_url(Wikipedia URL)
        [action]   → art:3f9a... stored
iter 2: [perception] sets attach_artifact_id=art:3f9a... on extraction goal
        [decision] reads 8KB of Wikipedia text from attached bytes
        [decision] → answer (birth/death/contributions extracted)
iter 3: [perception] all goals done
```

**Key mechanism:** Perception detects that goal 2 ("extract birth date and
contributions") requires reading the fetched page. It sets `attach_artifact_id`
to the artifact from iteration 1. Decision receives the raw Wikipedia text
and extracts the required facts without an additional fetch.

---

### Query B — Tokyo Activities (Multi-Goal + Memory Carryover)

```
Find 3 family-friendly things to do in Tokyo this weekend.
Check Saturday's weather forecast and tell me which one is most appropriate.
```

**What it tests:** Multi-goal planning and intra-run memory carryover.

**Expected trace:**
```
iter 1: [perception] 3 goals created
        [decision] → web_search("family friendly Tokyo activities")
iter 2: [decision] → web_search("Tokyo Saturday weather forecast")
        [memory]   weather fact recorded as tool_outcome
iter 3: [memory.read] weather fact hit from iter 2
        [perception] attach_artifact_id set for synthesis goal
        [decision] reads weather + activities → answer
```

**Key mechanism:** The weather fact fetched in iteration 2 is stored in memory
as a `tool_outcome`. In iteration 3, `memory.read()` retrieves it, and Perception
passes it to Decision so the activity recommendation can be weather-appropriate.

---

### Query C — Mom's Birthday (Durable Memory Across Two Runs)

```
My mom's birthday is 15 May 2026. Remember that and give me a
calendar reminder for two weeks before on the day.
```

**Run 1** — from clean state:
```
[memory.remember] → stores kind="fact", keywords=["mom","birthday","May","2026"]
[decision] → create_file("state/reminders/mom_birthday_2026.txt", ...)
[answer]   reminder created for 1 May 2026
```

**Run 2** — same command, **do NOT delete `state/`**:
```
[memory.read] 1 hit: fact about mom's birthday
[perception] detects fact already known
[decision] → answer (recalls birthday from memory, no re-fetch needed)
```

**Key mechanism:** The durable-memory contract. `memory.remember()` runs at the
very start of the agent loop (before any iteration). On Run 2, the keyword
intersection of `{"mom", "birthday", "may", "2026"}` scores the stored fact
highly, so `memory.read()` surfaces it and Perception marks the recall goal done.

---

### Query D — Python Asyncio Research (Multi-Artifact Synthesis)

```
Search for Python asyncio best practices and give me a short
numbered list of the advice that the top 3 results agree on.
```

**What it tests:** Multi-artifact attachment and synthesis.

**Expected trace:**
```
iter 1: [decision] → web_search("Python asyncio best practices")
        [action]   → art:aa1b... (search results)
iter 2: [decision] → fetch_url(result 1 URL)  → art:bb2c...
iter 3: [decision] → fetch_url(result 2 URL)  → art:cc3d...
iter 4: [decision] → fetch_url(result 3 URL)  → art:dd4e...
iter 5: [perception] sets attach_artifact_id to most recent synthesis artifact
        [decision] reads all fetched content → numbered list answer
```

**Key mechanism:** Perception's force-attach safety net. After 3 fetch iterations,
Perception sets `attach_artifact_id` to the most recently stored artifact so
Decision can synthesise across the collected material in a single call.

---

## 11. The Web UI

The UI is a single-page application served by FastAPI at `http://localhost:8000`.

### Layout

```
┌─────────────────┬──────────────────────────────────────────────┐
│  SIDEBAR        │  TOPBAR  (Gateway status dot)                 │
│                 ├──────────────────────────────────────────────┤
│  Assignment     │                                               │
│  Queries        │  CHAT WINDOW                                  │
│                 │                                               │
│  [A] Wikipedia  │  User bubble (right-aligned, indigo)          │
│  [B] Tokyo      │                                               │
│  [C] Birthday   │  Agent bubble (left-aligned, dark):           │
│  [D] Asyncio    │  ▶ Show agent trace (collapsible log)         │
│                 │    [iter 1] [memory] [perception] ...         │
│                 │  Final answer text                            │
│                 │                                               │
│                 ├──────────────────────────────────────────────┤
│                 │  [ text input              ] [Send ▶]        │
└─────────────────┴──────────────────────────────────────────────┘
```

### Features

- **Sidebar query cards** — click any card to pre-fill the input with the
  assignment query, colour-coded by type
- **SSE streaming** — agent progress is streamed in real-time via
  Server-Sent Events; the UI never polls
- **Collapsible trace log** — each agent response shows a "Show agent trace"
  toggle revealing the full colour-coded iteration log:
  - Purple — `--- iter N ---`
  - Green — `[memory]` reads and writes
  - Blue — `[perception]` goal updates
  - Yellow — `[decision]` choices
  - Orange — `[action]` tool dispatches
- **Gateway health indicator** — green/red dot in the top-right, polled every 10s
- **Auto-grow textarea** — input box expands with content up to 5 lines
- `Enter` sends; `Shift+Enter` inserts a newline

### How streaming works

```
Browser                           ui_server.py              agent6.py
   │                                    │                        │
   │── POST /chat ──────────────────────►│                        │
   │                                    │── asyncio.create_task ─►│
   │◄── SSE stream ─────────────────────│                        │
   │                                    │◄── on_progress(msg) ───│
   │◄── data: {"type":"progress",...} ──│                        │
   │◄── data: {"type":"progress",...} ──│  (each _emit() call)   │
   │                                    │◄── final answer ────────│
   │◄── data: {"type":"answer",...} ────│                        │
   │◄── data: [DONE] ───────────────────│                        │
```

---

## 12. LLM Gateway V3

**File:** `llm_gatewayV3/gateway.py`
**Port:** `8101`
**Model:** `gemini-2.0-flash`

The gateway is a FastAPI proxy that routes all LLM calls from the four roles.
No role imports `google.generativeai` directly — they all call the gateway.

### API

```
POST http://localhost:8101/chat

Request body:
{
  "messages": [
    {"role": "system", "content": "..."},
    {"role": "user",   "content": "..."},
    {"role": "model",  "content": "..."}
  ],
  "response_format": {           // optional — triggers structured output
    "type": "json_schema",
    "schema_": { ... }           // JSON Schema dict
  },
  "provider": "g",               // optional — explicit Gemini (Perception uses this)
  "auto_route": "decision",      // optional label (memory / decision / perception)
  "reasoning": "off"             // optional
}

Response 200:
{
  "content": "..." | {...},      // string or dict if response_format given
  "router_decision": {
    "worker": "gemini-2.0-flash",
    "fallback_used": false
  },
  "reasoning_applied": false
}
```

### Routing logic

| Role | How it calls the gateway | Why |
|---|---|---|
| Perception | `provider="g"` | Pinned to Gemini for reliable structured output |
| Decision | `auto_route="decision"` | Router pool (currently Gemini Flash) |
| Memory | `auto_route="memory"` | Router pool |
| Action | — | No LLM call |

### Structured output

When `response_format.schema_` is provided, the gateway passes the schema
to Gemini as `response_mime_type="application/json"` + `response_schema`.
The response is automatically parsed from JSON to a Python dict before returning.

### Health check

```
GET http://localhost:8101/health
→ {"status": "ok", "model": "gemini-2.0-flash"}
```

---

## 13. MCP Server & Tools

**File:** `mcp_server.py`
**Transport:** stdio (launched as a subprocess by `agent6.py`)
**Framework:** FastMCP

The MCP server exposes 9 tools. Action dispatches to them via
`session.call_tool(name, arguments)` — no tool logic lives in the agent.

| Tool | Arguments | Returns | Notes |
|---|---|---|---|
| `web_search` | `query`, `max_results=5` | Ranked list of title + URL + snippet | DuckDuckGo |
| `fetch_url` | `url` | Page text (stripped HTML, max 8000 chars) | httpx + BeautifulSoup |
| `get_time` | `timezone="UTC"` | Formatted datetime string | zoneinfo |
| `currency_convert` | `amount`, `from_currency`, `to_currency` | Converted amount | Frankfurter API |
| `read_file` | `path` | File contents (UTF-8) | |
| `create_file` | `path`, `content` | Success message | Fails if file exists |
| `update_file` | `path`, `content` | Success message | Creates or overwrites |
| `edit_file` | `path`, `old_text`, `new_text` | Success message | First-occurrence replace |
| `list_dir` | `path="."` | DIR/FILE entries | |

**Important:** `fetch_url` uses `httpx` + `BeautifulSoup` (not `crawl4ai`)
to avoid a Windows build dependency on MSVC. Output is capped at 8000
characters to prevent context overflow.

---

## 14. Memory System

### Storage format

All memory items live in a single JSON array at `state/memory.json`.
The file is loaded on first read and written back after every mutation.
Deleting this file resets the agent's memory completely.

### Read algorithm

`memory.read()` uses **lowercase token intersection** — no LLM call, no
vector embeddings. This keeps each Perception call fast (no extra LLM
latency before the main Perception call).

```python
query_tokens = {"john", "birthday", "may", "2026"}  # stopwords removed
item_score   = |query_tokens ∩ item_keywords| / |query_tokens|
```

Items are ranked by score; top-k with score > 0 are returned.

### Durable-memory contract

`memory.remember()` runs **once per agent invocation**, before the loop
starts. The gateway classifies whether the query contains a durable fact
or preference. If yes, it extracts keywords, a descriptor, and a structured
value, and stores a `MemoryItem`. This item persists to disk immediately.

On the next run, `memory.read()` will find this item via keyword overlap
before the first Perception call, enabling zero-tool-call recall.

---

## 15. Artifact Store

### Purpose

Some tools return large payloads (a 250KB Wikipedia page, three web pages
for synthesis). Storing all of this in the LLM context window on every
iteration would be wasteful and expensive. The ArtifactStore sidesteps
this: tools write raw bytes to disk; the LLM only sees a short handle
(`art:3f9a2b1c...`) until Perception explicitly decides to attach the bytes.

### Deduplication

The store uses SHA-256 of the content as the file name. Fetching the same
URL twice produces the same artifact ID. Memory records the handle; no
duplicate bytes are stored.

### Architectural boundary

```
Memory  ← holds the string "art:3f9a2b1c..."  inside MemoryItem.artifact_id
Perception ← sets Goal.attach_artifact_id = "art:3f9a2b1c..." when needed
agent6     ← calls artifacts.get(id) → raw bytes
Decision   ← receives bytes in attached[] parameter
Action     ← calls artifacts.put() after every tool dispatch
```

Decision **cannot** request an artifact — only Perception can grant access
to artifact bytes by setting `attach_artifact_id` on a goal.

---

## 16. Design Constraints

These constraints are enforced throughout and must be preserved:

| Constraint | Why |
|---|---|
| `uv` for all dependency management | Reproducible environments; no manual venv |
| Pydantic v2 on every role boundary | Validation at compile time; clear failures |
| All LLM calls through Gateway | Single point of routing, logging, fallback |
| MCP stdio transport for tools | Action is a pure dispatcher; no tool re-implementation |
| No LangChain / LangGraph / CrewAI | Build the architecture from first principles |
| `state/` gitignored | Clean state between assignment attempts |
| Decision picks at most one tool | Prevents runaway fan-out; Perception controls sequencing |
| Perception re-runs every iteration | Goals can be updated based on what tools return |

---

## 17. Dependency Reference

### `llm_gatewayV3/pyproject.toml`

| Package | Purpose |
|---|---|
| `fastapi` | HTTP server framework |
| `uvicorn[standard]` | ASGI server |
| `google-generativeai` | Gemini SDK |
| `pydantic` | Request/response validation |
| `python-dotenv` | Load `.env` files |
| `httpx` | HTTP client |

### `pyproject.toml` (main agent)

| Package | Purpose |
|---|---|
| `fastmcp` | MCP server framework (mcp_server.py) + client (agent6.py) |
| `httpx` | Gateway calls from roles |
| `pydantic` | All Pydantic contracts |
| `python-dotenv` | Load `.env` files |
| `duckduckgo-search==5.3.1` | `web_search` tool (pinned to avoid lxml build dep) |
| `beautifulsoup4` | HTML stripping in `fetch_url` |
| `aiofiles` | Async file I/O |
| `fastapi` | UI server |
| `uvicorn[standard]` | ASGI server for UI |

---

## Acknowledgements

Built as part of the **EAG V3 (Engineering Agentic GPT V3)** programme,
Session 6 — Agentic Basic Architecture.

The architecture follows the role decomposition pattern described in the
session materials: Memory, Perception, Decision, Action — each a typed
service with a single responsibility, communicating only through Pydantic
models and the shared Gateway.
