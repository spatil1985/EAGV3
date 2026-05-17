"""UI server — FastAPI + SSE chat interface for the PADMA agent.

Start with: uv run python ui_server.py
Then open:  http://localhost:8000
Gateway must already be running at http://localhost:8101
"""
from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

import uvicorn
from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel

load_dotenv()

# Add project dir to path so agent6 can find its siblings
sys.path.insert(0, str(Path(__file__).parent))
import agent6  # noqa: E402

app = FastAPI(title="PADMA Agent UI")

# ── Queries shown in the sidebar ────────────────────────────────────────────
QUERIES = [
    {
        "label": "A · Wikipedia Artifact",
        "short": "Claude Shannon Wikipedia",
        "query": "Fetch the Claude Shannon Wikipedia page and tell me his birth date, death date, and three key contributions to information theory.",
        "color": "#6366f1",
    },
    {
        "label": "B · Multi-Goal + Memory",
        "short": "Tokyo weekend activities",
        "query": "Find 3 family-friendly things to do in Tokyo this weekend. Check Saturday's weather forecast and tell me which one is most appropriate.",
        "color": "#0ea5e9",
    },
    {
        "label": "C · Durable Memory",
        "short": "Mom's birthday reminder",
        "query": "My mom's birthday is 15 May 2026. Remember that and give me a calendar reminder for two weeks before on the day.",
        "color": "#10b981",
    },
    {
        "label": "D · Multi-Artifact Synthesis",
        "short": "Python asyncio research",
        "query": "Search for Python asyncio best practices and give me a short numbered list of the advice that the top 3 results agree on.",
        "color": "#f59e0b",
    },
]

# ── HTML page ────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PADMA Agent</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #0f172a;
    --sidebar-bg: #1e293b;
    --chat-bg: #0f172a;
    --bubble-user: #6366f1;
    --bubble-agent: #1e293b;
    --text: #e2e8f0;
    --text-muted: #94a3b8;
    --border: #334155;
    --input-bg: #1e293b;
    --log-bg: #0d1526;
    --radius: 12px;
    --font: 'Inter', system-ui, sans-serif;
  }

  body {
    font-family: var(--font);
    background: var(--bg);
    color: var(--text);
    height: 100vh;
    display: flex;
    overflow: hidden;
  }

  /* ── Sidebar ── */
  #sidebar {
    width: 280px;
    min-width: 280px;
    background: var(--sidebar-bg);
    border-right: 1px solid var(--border);
    display: flex;
    flex-direction: column;
    padding: 20px 16px;
    gap: 12px;
    overflow-y: auto;
  }

  #sidebar h2 {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--text-muted);
    padding: 0 4px 8px;
    border-bottom: 1px solid var(--border);
  }

  .query-card {
    border-radius: var(--radius);
    padding: 14px 16px;
    cursor: pointer;
    border: 1px solid transparent;
    transition: all 0.15s ease;
    position: relative;
    overflow: hidden;
  }

  .query-card::before {
    content: '';
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    border-radius: 3px 0 0 3px;
  }

  .query-card:hover {
    border-color: var(--border);
    background: rgba(255,255,255,0.04);
    transform: translateX(2px);
  }

  .query-card:active { transform: translateX(1px) scale(0.99); }

  .query-label {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.05em;
    margin-bottom: 6px;
    opacity: 0.85;
  }

  .query-short {
    font-size: 13px;
    font-weight: 500;
    color: var(--text);
    margin-bottom: 6px;
  }

  .query-preview {
    font-size: 11px;
    color: var(--text-muted);
    line-height: 1.5;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }

  #sidebar-footer {
    margin-top: auto;
    padding-top: 12px;
    border-top: 1px solid var(--border);
    font-size: 11px;
    color: var(--text-muted);
    line-height: 1.6;
  }

  /* ── Main chat area ── */
  #main {
    flex: 1;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  #topbar {
    height: 56px;
    display: flex;
    align-items: center;
    padding: 0 24px;
    border-bottom: 1px solid var(--border);
    gap: 12px;
    flex-shrink: 0;
  }

  #topbar .logo {
    width: 32px; height: 32px;
    background: linear-gradient(135deg, #6366f1, #0ea5e9);
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 14px;
  }

  #topbar h1 { font-size: 16px; font-weight: 700; }
  #topbar span { font-size: 12px; color: var(--text-muted); margin-left: 4px; }

  #gateway-status {
    margin-left: auto;
    display: flex; align-items: center; gap: 6px;
    font-size: 12px; color: var(--text-muted);
  }

  .status-dot {
    width: 8px; height: 8px;
    border-radius: 50%;
    background: #ef4444;
    transition: background 0.3s;
  }
  .status-dot.ok { background: #10b981; }

  /* ── Messages ── */
  #messages {
    flex: 1;
    overflow-y: auto;
    padding: 24px;
    display: flex;
    flex-direction: column;
    gap: 20px;
    scroll-behavior: smooth;
  }

  #messages::-webkit-scrollbar { width: 6px; }
  #messages::-webkit-scrollbar-track { background: transparent; }
  #messages::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

  .msg { display: flex; gap: 12px; max-width: 820px; }
  .msg.user { align-self: flex-end; flex-direction: row-reverse; }
  .msg.agent { align-self: flex-start; }

  .avatar {
    width: 32px; height: 32px; border-radius: 8px;
    flex-shrink: 0; display: flex; align-items: center;
    justify-content: center; font-size: 14px; font-weight: 700;
  }
  .msg.user .avatar { background: var(--bubble-user); }
  .msg.agent .avatar { background: #334155; }

  .bubble {
    padding: 12px 16px;
    border-radius: var(--radius);
    font-size: 14px;
    line-height: 1.65;
    max-width: 680px;
    word-break: break-word;
  }

  .msg.user .bubble {
    background: var(--bubble-user);
    border-bottom-right-radius: 4px;
  }

  .msg.agent .bubble {
    background: var(--bubble-agent);
    border: 1px solid var(--border);
    border-bottom-left-radius: 4px;
  }

  .bubble pre {
    background: var(--log-bg);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 12px;
    overflow-x: auto;
    font-size: 12px;
    line-height: 1.6;
    margin-top: 8px;
    font-family: 'Fira Code', 'Cascadia Code', monospace;
  }

  /* Progress log inside agent bubble */
  .log-toggle {
    display: flex; align-items: center; gap: 6px;
    font-size: 12px; color: var(--text-muted);
    cursor: pointer; user-select: none;
    margin-bottom: 8px;
    background: none; border: none; padding: 0;
    color: var(--text-muted);
  }
  .log-toggle:hover { color: var(--text); }
  .log-toggle svg { transition: transform 0.2s; }
  .log-toggle.open svg { transform: rotate(90deg); }

  .log-block {
    background: var(--log-bg);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px 12px;
    font-size: 11.5px;
    font-family: 'Fira Code', 'Cascadia Code', monospace;
    line-height: 1.7;
    max-height: 300px;
    overflow-y: auto;
    white-space: pre-wrap;
    margin-bottom: 10px;
    display: none;
  }
  .log-block.open { display: block; }
  .log-block::-webkit-scrollbar { width: 4px; }
  .log-block::-webkit-scrollbar-thumb { background: var(--border); }

  .log-line-iter { color: #6366f1; font-weight: 700; }
  .log-line-memory { color: #10b981; }
  .log-line-perception { color: #0ea5e9; }
  .log-line-decision { color: #f59e0b; }
  .log-line-action { color: #f97316; }
  .log-line-mcp { color: #a78bfa; }
  .log-line-answer { color: #34d399; font-weight: 700; }

  .answer-text {
    color: var(--text);
    line-height: 1.7;
    white-space: pre-wrap;
  }

  .thinking {
    display: flex; gap: 4px; align-items: center;
    padding: 4px 0;
  }
  .thinking span {
    width: 7px; height: 7px; border-radius: 50%;
    background: var(--text-muted);
    animation: blink 1.2s infinite ease-in-out;
  }
  .thinking span:nth-child(2) { animation-delay: 0.2s; }
  .thinking span:nth-child(3) { animation-delay: 0.4s; }
  @keyframes blink {
    0%, 80%, 100% { opacity: 0.2; transform: scale(0.8); }
    40% { opacity: 1; transform: scale(1); }
  }

  /* ── Input area ── */
  #input-area {
    padding: 16px 24px 20px;
    border-top: 1px solid var(--border);
    flex-shrink: 0;
  }

  #input-row {
    display: flex; gap: 10px; align-items: flex-end;
    background: var(--input-bg);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 10px 14px;
    transition: border-color 0.15s;
  }
  #input-row:focus-within { border-color: #6366f1; }

  #query-input {
    flex: 1; background: none; border: none; outline: none;
    color: var(--text); font-size: 14px; font-family: var(--font);
    resize: none; min-height: 22px; max-height: 120px;
    line-height: 1.5;
  }
  #query-input::placeholder { color: var(--text-muted); }

  #send-btn {
    width: 36px; height: 36px; border-radius: 9px;
    background: #6366f1; border: none; cursor: pointer;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0; transition: background 0.15s, opacity 0.15s;
  }
  #send-btn:hover { background: #818cf8; }
  #send-btn:disabled { opacity: 0.4; cursor: not-allowed; }
  #send-btn svg { color: white; }

  #hint {
    margin-top: 8px;
    font-size: 11px; color: var(--text-muted);
    text-align: center;
  }

  /* ── Empty state ── */
  #empty-state {
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    flex: 1; gap: 16px; color: var(--text-muted);
    text-align: center; padding: 40px;
  }
  #empty-state .big { font-size: 48px; }
  #empty-state h3 { font-size: 18px; color: var(--text); }
  #empty-state p { font-size: 14px; max-width: 360px; line-height: 1.6; }

  /* ── Scrollbar ── */
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

  /* ── Console panel ── */
  #console-panel {
    border-top: 2px solid var(--border);
    background: var(--log-bg);
    flex-shrink: 0;
    display: flex;
    flex-direction: column;
  }

  #console-header {
    display: flex;
    align-items: center;
    padding: 5px 16px;
    gap: 10px;
    background: #080f1e;
    border-bottom: 1px solid var(--border);
    flex-shrink: 0;
  }

  #console-header .console-title {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #6366f1;
  }

  #console-line-count {
    font-size: 11px;
    color: var(--text-muted);
    margin-right: auto;
  }

  .console-btn {
    background: none;
    border: 1px solid var(--border);
    border-radius: 4px;
    color: var(--text-muted);
    font-size: 11px;
    padding: 2px 8px;
    cursor: pointer;
    font-family: var(--font);
  }
  .console-btn:hover { color: var(--text); border-color: #4b5563; }

  #console-output {
    overflow-y: auto;
    padding: 8px 16px;
    font-family: 'Fira Code', 'Cascadia Code', monospace;
    font-size: 12px;
    line-height: 1.75;
    white-space: pre-wrap;
    height: 200px;
  }
  #console-output.hidden { display: none; }
  #console-output::-webkit-scrollbar { width: 4px; }
  #console-output::-webkit-scrollbar-thumb { background: var(--border); }

  .con-iter      { color: #6366f1; font-weight: 700; }
  .con-memory    { color: #10b981; }
  .con-perception{ color: #0ea5e9; }
  .con-decision  { color: #f59e0b; }
  .con-action    { color: #f97316; }
  .con-mcp       { color: #a78bfa; }
  .con-answer    { color: #34d399; font-weight: 700; }
  .con-error     { color: #f87171; font-weight: 700; }
  .con-separator { color: #334155; }
</style>
</head>
<body>

<!-- Sidebar -->
<div id="sidebar">
  <h2>Assignment Queries</h2>
  __QUERY_CARDS__
  <div id="sidebar-footer">
    EAG V3 · Session 6<br>PADMA Agent Architecture
  </div>
</div>

<!-- Main -->
<div id="main">
  <div id="topbar">
    <div class="logo">P</div>
    <h1>PADMA Agent</h1>
    <span>Session 6</span>
    <div id="gateway-status">
      <div class="status-dot" id="gw-dot"></div>
      <span id="gw-label">Gateway...</span>
    </div>
  </div>

  <div id="messages">
    <div id="empty-state">
      <div class="big">🤖</div>
      <h3>PADMA Agent Ready</h3>
      <p>Click a query from the sidebar to pre-fill it, or type your own question below.</p>
    </div>
  </div>

  <div id="input-area">
    <div id="input-row">
      <textarea id="query-input" rows="1" placeholder="Ask PADMA anything…"></textarea>
      <button id="send-btn" title="Send">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none"
             stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <line x1="22" y1="2" x2="11" y2="13"></line>
          <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
        </svg>
      </button>
    </div>
    <div id="hint">Enter to send &nbsp;·&nbsp; Shift+Enter for newline &nbsp;·&nbsp; Gateway must be running at :8101</div>
  </div>

  <!-- Console panel -->
  <div id="console-panel">
    <div id="console-header">
      <span class="console-title">◉ Console</span>
      <span id="console-line-count">0 lines</span>
      <button class="console-btn" onclick="consoleClear()">Clear</button>
      <button class="console-btn" id="console-toggle-btn" onclick="consoleToggle()">▼ Hide</button>
    </div>
    <div id="console-output"></div>
  </div>
</div>

<script>
const QUERIES = __QUERIES_JSON__;

// ── Auto-grow textarea ──────────────────────────────────────────────────────
const textarea = document.getElementById('query-input');
textarea.addEventListener('input', () => {
  textarea.style.height = 'auto';
  textarea.style.height = Math.min(textarea.scrollHeight, 120) + 'px';
});

// ── Enter key handling ──────────────────────────────────────────────────────
textarea.addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendQuery(); }
});
document.getElementById('send-btn').addEventListener('click', sendQuery);

// ── Sidebar card click ──────────────────────────────────────────────────────
function buildSidebarCards() {
  // Cards are server-rendered, just wire up clicks
  document.querySelectorAll('.query-card').forEach((card, i) => {
    card.addEventListener('click', () => {
      textarea.value = QUERIES[i].query;
      textarea.style.height = 'auto';
      textarea.style.height = Math.min(textarea.scrollHeight, 120) + 'px';
      textarea.focus();
      // Highlight card briefly
      card.style.background = 'rgba(99,102,241,0.12)';
      setTimeout(() => card.style.background = '', 400);
    });
  });
}
buildSidebarCards();

// ── Gateway health check ────────────────────────────────────────────────────
async function checkGateway() {
  try {
    const r = await fetch('/gateway-health');
    const ok = r.ok;
    document.getElementById('gw-dot').className = 'status-dot' + (ok ? ' ok' : '');
    document.getElementById('gw-label').textContent = ok ? 'Gateway online' : 'Gateway offline';
  } catch {
    document.getElementById('gw-label').textContent = 'Gateway offline';
  }
}
checkGateway();
setInterval(checkGateway, 10000);

// ── Messages ────────────────────────────────────────────────────────────────
const messagesEl = document.getElementById('messages');
const emptyState = document.getElementById('empty-state');
let running = false;

function scrollBottom() {
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function addUserMsg(text) {
  emptyState.style.display = 'none';
  const msg = document.createElement('div');
  msg.className = 'msg user';
  msg.innerHTML = `
    <div class="avatar">U</div>
    <div class="bubble">${escHtml(text)}</div>
  `;
  messagesEl.appendChild(msg);
  scrollBottom();
}

function addAgentThinking() {
  const msg = document.createElement('div');
  msg.className = 'msg agent';
  msg.id = 'agent-thinking';
  msg.innerHTML = `
    <div class="avatar">P</div>
    <div class="bubble">
      <div class="thinking"><span></span><span></span><span></span></div>
    </div>
  `;
  messagesEl.appendChild(msg);
  scrollBottom();
  return msg;
}

function colorLogLine(line) {
  if (line.startsWith('--- iter'))  return `<span class="log-line-iter">${escHtml(line)}</span>`;
  if (line.startsWith('[memory'))   return `<span class="log-line-memory">${escHtml(line)}</span>`;
  if (line.startsWith('[perception')) return `<span class="log-line-perception">${escHtml(line)}</span>`;
  if (line.startsWith('[decision')) return `<span class="log-line-decision">${escHtml(line)}</span>`;
  if (line.startsWith('[action'))   return `<span class="log-line-action">${escHtml(line)}</span>`;
  if (line.startsWith('[mcp'))      return `<span class="log-line-mcp">${escHtml(line)}</span>`;
  if (line.startsWith('=== FINAL')) return `<span class="log-line-answer">${escHtml(line)}</span>`;
  return escHtml(line);
}

function createAgentMsg(logLines, answer) {
  const logId = 'log-' + Date.now();
  const logHtml = logLines.map(colorLogLine).join('\n');
  const msg = document.createElement('div');
  msg.className = 'msg agent';
  msg.innerHTML = `
    <div class="avatar">P</div>
    <div class="bubble">
      <button class="log-toggle" onclick="toggleLog('${logId}', this)">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="9 18 15 12 9 6"></polyline>
        </svg>
        Show agent trace (${logLines.length} steps)
      </button>
      <div class="log-block" id="${logId}">${logHtml}</div>
      <div class="answer-text">${escHtml(answer)}</div>
    </div>
  `;
  return msg;
}

function toggleLog(id, btn) {
  const el = document.getElementById(id);
  const open = el.classList.toggle('open');
  btn.classList.toggle('open', open);
  btn.querySelector('span') || null;
  const countMatch = btn.textContent.match(/\d+ steps/);
  if (open) {
    btn.innerHTML = btn.innerHTML.replace('Show agent trace', 'Hide agent trace');
    el.scrollTop = el.scrollHeight;
  } else {
    btn.innerHTML = btn.innerHTML.replace('Hide agent trace', 'Show agent trace');
  }
  scrollBottom();
}

function escHtml(s) {
  return String(s)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}

// ── Console panel ──────────────────────────────────────────────────────────────
let consoleOpen = true;
let consoleLineCount = 0;
const consoleOutput = document.getElementById('console-output');
const consoleLineCountEl = document.getElementById('console-line-count');

function consoleToggle() {
  consoleOpen = !consoleOpen;
  consoleOutput.classList.toggle('hidden', !consoleOpen);
  document.getElementById('console-toggle-btn').textContent = consoleOpen ? '▼ Hide' : '▲ Show';
}

function consoleClear() {
  consoleOutput.innerHTML = '';
  consoleLineCount = 0;
  consoleLineCountEl.textContent = '0 lines';
}

function consoleAppend(text, cssClass) {
  const el = document.createElement('div');
  if (cssClass) el.className = cssClass;
  el.textContent = text;
  consoleOutput.appendChild(el);
  consoleLineCount++;
  consoleLineCountEl.textContent = consoleLineCount + ' lines';
  consoleOutput.scrollTop = consoleOutput.scrollHeight;
}

function consoleLog(line) {
  let cls = 'con-muted';
  if (line.startsWith('--- iter') || line.startsWith('[agent6]')) cls = 'con-iter';
  else if (line.startsWith('[memory'))      cls = 'con-memory';
  else if (line.startsWith('[perception'))  cls = 'con-perception';
  else if (line.startsWith('[decision'))    cls = 'con-decision';
  else if (line.startsWith('[action'))      cls = 'con-action';
  else if (line.startsWith('[mcp'))         cls = 'con-mcp';
  else if (line.startsWith('=== FINAL'))    cls = 'con-answer';
  consoleAppend(line, cls);
}

// ── Send query ───────────────────────────────────────────────────────────────
async function sendQuery() {
  const q = textarea.value.trim();
  if (!q || running) return;

  running = true;
  document.getElementById('send-btn').disabled = true;
  textarea.value = '';
  textarea.style.height = 'auto';

  addUserMsg(q);
  const thinkingMsg = addAgentThinking();

  // Console: separator + query header
  consoleAppend('─'.repeat(72), 'con-separator');
  consoleAppend('▶ ' + q, 'con-iter');

  const logLines = [];
  let answer = '';
  let inAnswer = false;

  try {
    const resp = await fetch('/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ query: q }),
    });

    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      const lines = buffer.split('\n');
      buffer = lines.pop();

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        const raw = line.slice(6);
        if (raw === '[DONE]') break;
        try {
          const evt = JSON.parse(raw);
          if (evt.type === 'progress') {
            logLines.push(evt.text);
            consoleLog(evt.text);
          } else if (evt.type === 'answer') {
            answer = evt.text;
            consoleAppend('=== ANSWER ===', 'con-answer');
            consoleAppend(evt.text, 'con-answer');
          } else if (evt.type === 'error') {
            answer = '⚠️ Error: ' + evt.text;
            consoleAppend('ERROR: ' + evt.text, 'con-error');
          }
        } catch {}
      }
    }
  } catch (err) {
    answer = '⚠️ Network error: ' + err.message;
  }

  // Replace thinking bubble with real message
  thinkingMsg.remove();
  if (answer) {
    const finalMsg = createAgentMsg(logLines, answer);
    messagesEl.appendChild(finalMsg);
  } else {
    // No explicit answer — show last log line
    const fallback = logLines[logLines.length - 1] || 'Done.';
    const finalMsg = createAgentMsg(logLines, fallback);
    messagesEl.appendChild(finalMsg);
  }

  scrollBottom();
  running = false;
  document.getElementById('send-btn').disabled = false;
  textarea.focus();
}
</script>
</body>
</html>
"""


def _build_html() -> str:
    # Build sidebar cards
    cards_html = ""
    for q in QUERIES:
        cards_html += f"""
  <div class="query-card" style="background: {q['color']}11; border-color: {q['color']}33;">
    <div class="query-label" style="color: {q['color']};">{q['label']}</div>
    <div class="query-short">{q['short']}</div>
    <div class="query-preview">{q['query']}</div>
    <style>.query-card:hover {{ border-color: {q['color']}66 !important; }}</style>
  </div>"""

    queries_json = json.dumps(QUERIES)
    return (
        HTML.replace("__QUERY_CARDS__", cards_html)
            .replace("__QUERIES_JSON__", queries_json)
    )


# ── Routes ───────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    return HTMLResponse(_build_html())


@app.get("/gateway-health")
async def gateway_health():
    import httpx
    try:
        r = httpx.get("http://localhost:8101/health", timeout=3)
        return {"ok": r.status_code == 200}
    except Exception:
        return {"ok": False}


class ChatRequest(BaseModel):
    query: str


def _unwrap_exc(e: BaseException) -> str:
    """Recursively unwrap ExceptionGroup/BaseExceptionGroup to the root cause."""
    if hasattr(e, "exceptions") and e.exceptions:
        return _unwrap_exc(e.exceptions[0])
    return f"{type(e).__name__}: {e}"


@app.post("/chat")
async def chat_stream(body: ChatRequest):
    queue: asyncio.Queue[dict] = asyncio.Queue()
    ui_loop = asyncio.get_running_loop()

    def on_progress(msg: str) -> None:
        # Thread-safe: posts from agent6's isolated loop back to the UI loop.
        ui_loop.call_soon_threadsafe(queue.put_nowait, {"type": "progress", "text": msg})

    def _run_agent_blocking() -> str:
        # Run agent6 in a brand-new event loop, fully isolated from the UI
        # server's loop.  This prevents the MCP stdio_client's anyio TaskGroup
        # from sharing the UI loop and causing nested ExceptionGroups.
        async def _inner() -> str:
            return await agent6.run(body.query, on_progress=on_progress)
        return asyncio.run(_inner())

    async def produce() -> None:
        try:
            result = await asyncio.to_thread(_run_agent_blocking)
            await queue.put({"type": "answer", "text": result})
        except BaseException as e:
            await queue.put({"type": "error", "text": _unwrap_exc(e)})
        finally:
            await queue.put(None)

    asyncio.create_task(produce())

    async def event_stream():
        while True:
            item = await queue.get()
            if item is None:
                yield "data: [DONE]\n\n"
                break
            yield f"data: {json.dumps(item, ensure_ascii=False)}\n\n"

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


if __name__ == "__main__":
    print("Starting PADMA UI at http://localhost:8000")
    print("Make sure the gateway is running: cd llm_gatewayV3 && uv run uvicorn gateway:app --port 8101")
    uvicorn.run(app, host="0.0.0.0", port=8000)
