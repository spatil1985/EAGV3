"""UI server — FastAPI + SSE chat interface for the PADMA S8 DAG agent.

Start with: python ui_server.py
Then open:  http://localhost:8000
Gateway V8 must be running at http://localhost:8108
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

sys.path.insert(0, str(Path(__file__).parent))
import flow  # noqa: E402

app = FastAPI(title="PADMA S8 DAG Agent UI")

# ── S8 base queries ──────────────────────────────────────────────────────────
QUERIES = [
    {
        "label": "Hello · Minimum DAG",
        "short": "Say hello",
        "query": "Say hello.",
        "color": "#6366f1",
    },
    {
        "label": "A · Wikipedia Artifact",
        "short": "Claude Shannon Wikipedia",
        "query": "Fetch https://en.wikipedia.org/wiki/Claude_Shannon and tell me his birth date, death date, and three key contributions to information theory.",
        "color": "#0ea5e9",
    },
    {
        "label": "I · Parallel Fan-out",
        "short": "London · Paris · Berlin populations",
        "query": "Find the populations of London, Paris, and Berlin and tell me which two are closest in size.",
        "color": "#10b981",
    },
    {
        "label": "J · Graceful Failure",
        "short": "Read nonexistent file",
        "query": "Read /nonexistent/path.txt and tell me what's in it.",
        "color": "#f59e0b",
    },
    {
        "label": "K · Parallel + Resume",
        "short": "Lagos · Cairo · Kinshasa growth",
        "query": "For Lagos, Cairo, and Kinshasa, find current populations and growth rates and tell me which is growing fastest.",
        "color": "#f97316",
    },
]

# ── HTML ─────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PADMA S8 — DAG Agent</title>
<script src="https://unpkg.com/vis-network/standalone/umd/vis-network.min.js"></script>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg: #0f172a; --sidebar-bg: #1e293b; --graph-bg: #0a1120;
  --bubble-user: #6366f1; --bubble-agent: #1e293b;
  --text: #e2e8f0; --text-muted: #94a3b8; --border: #334155;
  --input-bg: #1e293b; --log-bg: #0d1526; --radius: 12px;
  --font: 'Inter', system-ui, sans-serif;
}
body { font-family: var(--font); background: var(--bg); color: var(--text);
       height: 100vh; display: flex; overflow: hidden; }

/* Sidebar */
#sidebar { width: 252px; min-width: 252px; background: var(--sidebar-bg);
           border-right: 1px solid var(--border); display: flex;
           flex-direction: column; padding: 16px 12px; gap: 10px; overflow-y: auto; }
#sidebar h2 { font-size: 10px; font-weight: 700; letter-spacing: .1em;
              text-transform: uppercase; color: var(--text-muted);
              padding: 0 4px 8px; border-bottom: 1px solid var(--border); }
.query-card { border-radius: 10px; padding: 11px 13px; cursor: pointer;
              border: 1px solid transparent; transition: all .15s; }
.query-card:hover { border-color: var(--border); background: rgba(255,255,255,.04); transform: translateX(2px); }
.query-label { font-size: 10px; font-weight: 700; letter-spacing: .05em; margin-bottom: 3px; opacity: .85; }
.query-short { font-size: 12px; font-weight: 600; color: var(--text); margin-bottom: 3px; }
.query-preview { font-size: 11px; color: var(--text-muted); line-height: 1.45;
                 display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
#sidebar-footer { margin-top: auto; padding-top: 12px; border-top: 1px solid var(--border);
                  font-size: 11px; color: var(--text-muted); line-height: 1.6; }

/* Main */
#main { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-width: 0; }
#topbar { height: 50px; display: flex; align-items: center; padding: 0 18px;
          border-bottom: 1px solid var(--border); gap: 10px; flex-shrink: 0; }
#topbar .logo { width: 28px; height: 28px; background: linear-gradient(135deg,#6366f1,#0ea5e9);
                border-radius: 7px; display: flex; align-items: center; justify-content: center;
                font-weight: 800; font-size: 12px; }
#topbar h1 { font-size: 15px; font-weight: 700; }
.badge { font-size: 10px; background: #1e293b; border: 1px solid var(--border);
         border-radius: 6px; padding: 2px 7px; color: var(--text-muted); }
#gateway-status { margin-left: auto; display: flex; align-items: center; gap: 6px;
                  font-size: 11px; color: var(--text-muted); }
.status-dot { width: 7px; height: 7px; border-radius: 50%; background: #ef4444; transition: background .3s; }
.status-dot.ok { background: #10b981; }

/* Content row */
#content-row { flex: 1; display: flex; overflow: hidden; }

/* Chat column */
#chat-col { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-width: 0; }
#messages { flex: 1; overflow-y: auto; padding: 18px; display: flex;
            flex-direction: column; gap: 16px; scroll-behavior: smooth; }
#messages::-webkit-scrollbar { width: 5px; }
#messages::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
.msg { display: flex; gap: 10px; max-width: 760px; }
.msg.user { align-self: flex-end; flex-direction: row-reverse; }
.msg.agent { align-self: flex-start; }
.avatar { width: 28px; height: 28px; border-radius: 7px; flex-shrink: 0;
          display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; }
.msg.user .avatar { background: var(--bubble-user); }
.msg.agent .avatar { background: #334155; }
.bubble { padding: 10px 14px; border-radius: var(--radius); font-size: 13.5px;
          line-height: 1.65; max-width: 620px; word-break: break-word; }
.msg.user .bubble { background: var(--bubble-user); border-bottom-right-radius: 4px; }
.msg.agent .bubble { background: var(--bubble-agent); border: 1px solid var(--border); border-bottom-left-radius: 4px; }
.log-toggle { display: flex; align-items: center; gap: 5px; font-size: 11px;
              color: var(--text-muted); cursor: pointer; user-select: none;
              margin-bottom: 7px; background: none; border: none; padding: 0; font-family: var(--font); }
.log-toggle:hover { color: var(--text); }
.log-toggle svg { transition: transform .2s; }
.log-toggle.open svg { transform: rotate(90deg); }
.log-block { background: var(--log-bg); border: 1px solid var(--border); border-radius: 8px;
             padding: 9px 11px; font-size: 10.5px; font-family: 'Fira Code',monospace;
             line-height: 1.7; max-height: 240px; overflow-y: auto; white-space: pre-wrap;
             margin-bottom: 9px; display: none; }
.log-block.open { display: block; }
.answer-text { color: var(--text); line-height: 1.7; white-space: pre-wrap; }
.thinking { display: flex; gap: 4px; align-items: center; padding: 3px 0; }
.thinking span { width: 6px; height: 6px; border-radius: 50%; background: var(--text-muted);
                 animation: blink 1.2s infinite ease-in-out; }
.thinking span:nth-child(2) { animation-delay: .2s; }
.thinking span:nth-child(3) { animation-delay: .4s; }
@keyframes blink { 0%,80%,100%{opacity:.2;transform:scale(.8)} 40%{opacity:1;transform:scale(1)} }
.ll-flow     { color: #6366f1; font-weight: 700; }
.ll-executor { color: #a78bfa; }
.ll-skill    { color: #0ea5e9; }
.ll-sandbox  { color: #f59e0b; }
.ll-memory   { color: #10b981; }
.ll-answer   { color: #34d399; font-weight: 700; }

/* Input */
#input-area { padding: 12px 18px 16px; border-top: 1px solid var(--border); flex-shrink: 0; }
#input-row { display: flex; gap: 8px; align-items: flex-end; background: var(--input-bg);
             border: 1px solid var(--border); border-radius: 12px; padding: 8px 12px; transition: border-color .15s; }
#input-row:focus-within { border-color: #6366f1; }
#query-input { flex: 1; background: none; border: none; outline: none; color: var(--text);
               font-size: 13.5px; font-family: var(--font); resize: none; min-height: 20px; max-height: 90px; line-height: 1.5; }
#query-input::placeholder { color: var(--text-muted); }
#send-btn { width: 32px; height: 32px; border-radius: 8px; background: #6366f1; border: none;
            cursor: pointer; display: flex; align-items: center; justify-content: center;
            flex-shrink: 0; transition: background .15s; }
#send-btn:hover { background: #818cf8; }
#send-btn:disabled { opacity: .4; cursor: not-allowed; }
#hint { margin-top: 5px; font-size: 10.5px; color: var(--text-muted); text-align: center; }

/* Graph panel */
#graph-panel { width: 330px; min-width: 330px; background: var(--graph-bg);
               border-left: 1px solid var(--border); display: flex;
               flex-direction: column; overflow: hidden; }
#graph-header { padding: 9px 13px; border-bottom: 1px solid var(--border);
                display: flex; align-items: center; gap: 7px; flex-shrink: 0; background: #080f1e; }
.gh-title { font-size: 10px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: #a78bfa; }
#graph-stats { margin-left: auto; font-size: 10px; color: var(--text-muted); }
#graph-canvas { flex: 1; min-height: 0; }
#graph-legend { padding: 7px 11px; border-top: 1px solid var(--border); display: flex;
                flex-wrap: wrap; gap: 7px; flex-shrink: 0; background: #080f1e; }
.leg { display: flex; align-items: center; gap: 4px; font-size: 10px; color: var(--text-muted); }
.leg-dot { width: 8px; height: 8px; border-radius: 50%; }

/* Node detail popup */
#node-detail { display: none; position: fixed; bottom: 200px; right: 340px;
               width: 320px; max-height: 280px; background: #1e293b;
               border: 1px solid var(--border); border-radius: 10px; padding: 13px; z-index: 200; overflow-y: auto; }
#node-detail.show { display: block; }
#node-detail h4 { font-size: 12px; margin-bottom: 7px; color: var(--text); }
#node-detail pre { font-size: 10.5px; font-family: 'Fira Code',monospace; white-space: pre-wrap; color: var(--text-muted); }
#nd-close { float: right; background: none; border: none; color: var(--text-muted); cursor: pointer; font-size: 15px; }
#nd-close:hover { color: var(--text); }

/* Console */
#console-panel { border-top: 2px solid var(--border); background: var(--log-bg);
                 flex-shrink: 0; display: flex; flex-direction: column; }
#console-header { display: flex; align-items: center; padding: 5px 13px; gap: 8px;
                  background: #080f1e; border-bottom: 1px solid var(--border); flex-shrink: 0; }
.console-title { font-size: 10px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: #6366f1; }
#console-line-count { font-size: 10px; color: var(--text-muted); margin-right: auto; }
.console-btn { background: none; border: 1px solid var(--border); border-radius: 4px;
               color: var(--text-muted); font-size: 10px; padding: 2px 7px; cursor: pointer; font-family: var(--font); }
.console-btn:hover { color: var(--text); }
#console-output { overflow-y: auto; padding: 6px 13px; font-family: 'Fira Code',monospace;
                  font-size: 11px; line-height: 1.75; white-space: pre-wrap; height: 170px; }
#console-output.hidden { display: none; }
#console-output::-webkit-scrollbar { width: 4px; }
#console-output::-webkit-scrollbar-thumb { background: var(--border); }
.con-flow     { color: #6366f1; font-weight: 700; }
.con-executor { color: #a78bfa; }
.con-skill    { color: #0ea5e9; }
.con-sandbox  { color: #f59e0b; }
.con-memory   { color: #10b981; }
.con-answer   { color: #34d399; font-weight: 700; }
.con-error    { color: #f87171; font-weight: 700; }
.con-sep      { color: #334155; }
.con-graph    { color: #1e293b; font-size: 10px; }

/* Empty state */
#empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center;
               flex: 1; gap: 13px; color: var(--text-muted); text-align: center; padding: 32px; }
#empty-state .big { font-size: 42px; }
#empty-state h3 { font-size: 16px; color: var(--text); }
#empty-state p { font-size: 13px; max-width: 320px; line-height: 1.6; }
</style>
</head>
<body>

<div id="sidebar">
  <h2>Session 8 Queries</h2>
  __QUERY_CARDS__
  <div id="sidebar-footer">EAG V3 · Session 8<br>DAG Multi-Agent Orchestration</div>
</div>

<div id="main">
  <div id="topbar">
    <div class="logo">P</div>
    <h1>PADMA</h1>
    <span class="badge">S8 · DAG</span>
    <div id="gateway-status">
      <div class="status-dot" id="gw-dot"></div>
      <span id="gw-label">Gateway...</span>
    </div>
  </div>

  <div id="content-row">
    <div id="chat-col">
      <div id="messages">
        <div id="empty-state">
          <div class="big">🕸️</div>
          <h3>PADMA DAG Agent</h3>
          <p>Pick a query from the sidebar or type below. The DAG graph updates live as nodes execute.</p>
        </div>
      </div>
      <div id="input-area">
        <div id="input-row">
          <textarea id="query-input" rows="1" placeholder="Ask PADMA anything…"></textarea>
          <button id="send-btn" title="Send">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                 stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
              <line x1="22" y1="2" x2="11" y2="13"></line>
              <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
            </svg>
          </button>
        </div>
        <div id="hint">Enter to send · Shift+Enter for newline · Gateway V8 :8108</div>
      </div>
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

    <div id="graph-panel">
      <div id="graph-header">
        <span class="gh-title">⬡ DAG Graph</span>
        <span id="graph-stats">—</span>
        <button class="console-btn" onclick="resetGraph()" style="margin-left:6px">Clear</button>
      </div>
      <div id="graph-canvas"></div>
      <div id="graph-legend">
        <span class="leg"><span class="leg-dot" style="background:#374151;border:1px solid #6b7280"></span>pending</span>
        <span class="leg"><span class="leg-dot" style="background:#1d4ed8;border:2px solid #60a5fa"></span>running</span>
        <span class="leg"><span class="leg-dot" style="background:#065f46;border:1px solid #10b981"></span>complete</span>
        <span class="leg"><span class="leg-dot" style="background:#7f1d1d;border:1px solid #f87171"></span>failed</span>
        <span class="leg"><span class="leg-dot" style="background:#78350f;border:1px solid #f59e0b"></span>skipped</span>
      </div>
    </div>
  </div>
</div>

<div id="node-detail">
  <button id="nd-close" onclick="closeNodeDetail()">×</button>
  <h4 id="nd-title">Node Detail</h4>
  <pre id="nd-body"></pre>
</div>

<script>
const QUERIES = __QUERIES_JSON__;

// Textarea
const textarea = document.getElementById('query-input');
textarea.addEventListener('input', () => { textarea.style.height='auto'; textarea.style.height=Math.min(textarea.scrollHeight,90)+'px'; });
textarea.addEventListener('keydown', e => { if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();sendQuery();} });
document.getElementById('send-btn').addEventListener('click', sendQuery);

// Sidebar
document.querySelectorAll('.query-card').forEach((card,i) => {
  card.addEventListener('click', () => {
    textarea.value = QUERIES[i].query;
    textarea.style.height='auto'; textarea.style.height=Math.min(textarea.scrollHeight,90)+'px';
    textarea.focus();
    card.style.background='rgba(99,102,241,0.15)'; setTimeout(()=>card.style.background='',350);
  });
});

// Gateway health
async function checkGateway() {
  try {
    const r = await fetch('/gateway-health');
    const ok = r.ok;
    document.getElementById('gw-dot').className='status-dot'+(ok?' ok':'');
    document.getElementById('gw-label').textContent=ok?'Gateway online':'Gateway offline';
  } catch { document.getElementById('gw-label').textContent='Gateway offline'; }
}
checkGateway(); setInterval(checkGateway,10000);

// ── vis.js DAG ────────────────────────────────────────────────────────────
const ST = {
  pending:  {bg:'#1f2937',border:'#4b5563',font:'#9ca3af'},
  running:  {bg:'#1e3a8a',border:'#60a5fa',font:'#bfdbfe'},
  complete: {bg:'#064e3b',border:'#10b981',font:'#6ee7b7'},
  failed:   {bg:'#7f1d1d',border:'#f87171',font:'#fca5a5'},
  skipped:  {bg:'#78350f',border:'#f59e0b',font:'#fcd34d'},
};
const ICON = {planner:'🗺',researcher:'🔍',distiller:'⚗',summariser:'📝',
              critic:'🔬',formatter:'📄',coder:'💻',sandbox_executor:'⚙',
              retriever:'🗃',translator:'🌐',browser:'🌏'};

let nodesDS, edgesDS, network;
const nodeStore = {};

function initGraph() {
  nodesDS = new vis.DataSet();
  edgesDS = new vis.DataSet();
  network = new vis.Network(document.getElementById('graph-canvas'), {nodes:nodesDS,edges:edgesDS}, {
    layout: { hierarchical: {direction:'UD',sortMethod:'directed',levelSeparation:85,nodeSpacing:100} },
    physics: false,
    interaction: {hover:true,tooltipDelay:300,zoomView:true,dragView:true},
    nodes: {
      shape:'box', borderWidth:2, borderWidthSelected:3,
      font:{size:11,color:'#9ca3af',face:'Fira Code,monospace',multi:false},
      margin:{top:7,bottom:7,left:9,right:9},
      shadow:{enabled:true,color:'rgba(0,0,0,0.5)',size:6},
      widthConstraint:{maximum:130},
    },
    edges: {
      arrows:{to:{enabled:true,scaleFactor:0.55}},
      color:{color:'#334155',highlight:'#6366f1',hover:'#6366f1'},
      width:1.5, smooth:{type:'cubicBezier',forceDirection:'vertical'},
    },
  });
  network.on('click', p => { if(p.nodes.length) showNodeDetail(p.nodes[0]); });
}

function handleGraphUpdate(data) {
  if (data.event === 'node_added') {
    const s = ST.pending;
    const icon = ICON[data.skill]||'○';
    nodeStore[data.id] = {skill:data.skill, label:data.label, preview:'', status:'pending'};
    nodesDS.update({
      id: data.id,
      label: icon+' '+data.skill+'\n'+data.label,
      color: {background:s.bg, border:s.border, highlight:{background:s.bg,border:'#818cf8'}},
      font: {color:s.font},
      title: data.skill+' · '+data.label,
    });
    updateStats();
  } else if (data.event === 'edge_added') {
    const eid = data.from_id+'__'+data.to_id;
    if (!edgesDS.get(eid)) edgesDS.update({id:eid, from:data.from_id, to:data.to_id});
  } else if (data.event === 'node_status') {
    const s = ST[data.status]||ST.pending;
    const n = nodeStore[data.id]||{};
    n.status = data.status; if(data.preview) n.preview = data.preview;
    nodeStore[data.id] = n;
    const icon = ICON[n.skill]||'○';
    let lbl = icon+' '+n.skill+'\n'+(n.label||data.id);
    if(data.elapsed) lbl += '\n'+data.elapsed+'s';
    nodesDS.update({
      id: data.id, label: lbl,
      color:{background:s.bg,border:s.border,highlight:{background:s.bg,border:'#818cf8'}},
      font:{color:s.font},
      title: (n.skill||data.id)+' · '+(n.label||'')+'\n'+(data.preview||''),
      shadow: data.status==='running'
        ? {enabled:true,color:'rgba(96,165,250,0.7)',size:14}
        : {enabled:true,color:'rgba(0,0,0,0.5)',size:6},
    });
    updateStats();
  }
}

function updateStats() {
  const total = Object.keys(nodeStore).length;
  const done  = Object.values(nodeStore).filter(n=>n.status==='complete').length;
  const run   = Object.values(nodeStore).filter(n=>n.status==='running').length;
  let s = total+' nodes · '+done+' done';
  if(run) s += ' · '+run+' running ⟳';
  document.getElementById('graph-stats').textContent = s;
}

function resetGraph() {
  nodesDS.clear(); edgesDS.clear();
  Object.keys(nodeStore).forEach(k=>delete nodeStore[k]);
  document.getElementById('graph-stats').textContent='—';
}

function showNodeDetail(id) {
  const n = nodeStore[id]||{};
  document.getElementById('nd-title').textContent=(n.skill||id)+' · '+(n.label||id)+' ['+n.status+']';
  document.getElementById('nd-body').textContent=n.preview||(n.status==='pending'?'Waiting…':'No output yet');
  document.getElementById('node-detail').classList.add('show');
}
function closeNodeDetail() { document.getElementById('node-detail').classList.remove('show'); }

initGraph();

// ── Console ───────────────────────────────────────────────────────────────
let consoleOpen=true, cLines=0;
const consoleOut = document.getElementById('console-output');
const cCount = document.getElementById('console-line-count');

function consoleToggle() {
  consoleOpen=!consoleOpen;
  consoleOut.classList.toggle('hidden',!consoleOpen);
  document.getElementById('console-toggle-btn').textContent=consoleOpen?'▼ Hide':'▲ Show';
}
function consoleClear() { consoleOut.innerHTML=''; cLines=0; cCount.textContent='0 lines'; }

function consoleAppend(text, cls) {
  const el=document.createElement('div');
  if(cls) el.className=cls;
  el.textContent=text;
  consoleOut.appendChild(el); cLines++;
  cCount.textContent=cLines+' lines';
  consoleOut.scrollTop=consoleOut.scrollHeight;
}

function consoleLog(line) {
  let cls='';
  if(line.startsWith('[flow]'))         cls='con-flow';
  else if(line.startsWith('[executor]'))cls='con-executor';
  else if(line.startsWith('[skill:'))   cls='con-skill';
  else if(line.startsWith('[sandbox]')) cls='con-sandbox';
  else if(line.startsWith('[memory'))   cls='con-memory';
  else if(line.startsWith('[flow] memory')) cls='con-memory';
  else if(line.startsWith('=== FINAL'))cls='con-answer';
  else if(line.startsWith('GRAPH|'))   cls='con-graph';
  consoleAppend(line, cls);
}

// ── Messages ──────────────────────────────────────────────────────────────
const messagesEl=document.getElementById('messages');
const emptyState=document.getElementById('empty-state');
let running=false;
function scrollBottom(){messagesEl.scrollTop=messagesEl.scrollHeight;}

function addUserMsg(text){
  emptyState.style.display='none';
  const m=document.createElement('div'); m.className='msg user';
  m.innerHTML='<div class="avatar">U</div><div class="bubble">'+escHtml(text)+'</div>';
  messagesEl.appendChild(m); scrollBottom();
}

function addThinking(){
  const m=document.createElement('div'); m.className='msg agent'; m.id='agent-thinking';
  m.innerHTML='<div class="avatar">P</div><div class="bubble"><div class="thinking"><span></span><span></span><span></span></div></div>';
  messagesEl.appendChild(m); scrollBottom(); return m;
}

function colorLine(line){
  const e=escHtml(line);
  if(line.startsWith('[flow]'))      return '<span class="ll-flow">'+e+'</span>';
  if(line.startsWith('[executor]'))  return '<span class="ll-executor">'+e+'</span>';
  if(line.startsWith('[skill:'))     return '<span class="ll-skill">'+e+'</span>';
  if(line.startsWith('[sandbox]'))   return '<span class="ll-sandbox">'+e+'</span>';
  if(line.startsWith('[memory')||line.startsWith('[flow] memory')) return '<span class="ll-memory">'+e+'</span>';
  if(line.startsWith('=== FINAL'))   return '<span class="ll-answer">'+e+'</span>';
  return e;
}

function createAgentMsg(logLines, answer){
  const lid='log-'+Date.now();
  const msg=document.createElement('div'); msg.className='msg agent';
  msg.innerHTML=`<div class="avatar">P</div><div class="bubble">
    <button class="log-toggle" onclick="toggleLog('${lid}',this)">
      <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="9 18 15 12 9 6"></polyline></svg>
      Show trace (${logLines.length} lines)
    </button>
    <div class="log-block" id="${lid}">${logLines.map(colorLine).join('\n')}</div>
    <div class="answer-text">${escHtml(answer)}</div>
  </div>`;
  return msg;
}

function toggleLog(id,btn){
  const el=document.getElementById(id);
  const open=el.classList.toggle('open'); btn.classList.toggle('open',open);
  btn.innerHTML=btn.innerHTML.replace(open?'Show trace':'Hide trace',open?'Hide trace':'Show trace');
  scrollBottom();
}

function escHtml(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Send ──────────────────────────────────────────────────────────────────
async function sendQuery(){
  const q=textarea.value.trim(); if(!q||running) return;
  running=true; document.getElementById('send-btn').disabled=true;
  textarea.value=''; textarea.style.height='auto';

  resetGraph(); addUserMsg(q); const thinking=addThinking();
  consoleAppend('─'.repeat(60),'con-sep'); consoleAppend('▶ '+q,'con-flow');

  const logLines=[]; let answer='';

  try {
    const resp=await fetch('/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({query:q})});
    const reader=resp.body.getReader(); const dec=new TextDecoder(); let buf='';
    while(true){
      const {done,value}=await reader.read(); if(done) break;
      buf+=dec.decode(value,{stream:true});
      const lines=buf.split('\n'); buf=lines.pop();
      for(const line of lines){
        if(!line.startsWith('data: ')) continue;
        const raw=line.slice(6); if(raw==='[DONE]') break;
        try {
          const evt=JSON.parse(raw);
          if(evt.type==='progress'){logLines.push(evt.text);consoleLog(evt.text);}
          else if(evt.type==='graph_update'){handleGraphUpdate(evt.data);}
          else if(evt.type==='answer'){answer=evt.text;consoleAppend('=== ANSWER ===','con-answer');consoleAppend(evt.text,'con-answer');}
          else if(evt.type==='error'){answer='⚠️ '+evt.text;consoleAppend('ERROR: '+evt.text,'con-error');}
        } catch{}
      }
    }
  } catch(err){answer='⚠️ Network error: '+err.message;}

  thinking.remove();
  messagesEl.appendChild(createAgentMsg(logLines,answer||logLines[logLines.length-1]||'Done.'));
  scrollBottom(); running=false; document.getElementById('send-btn').disabled=false; textarea.focus();
}
</script>
</body>
</html>
"""


def _build_html() -> str:
    cards_html = ""
    for q in QUERIES:
        cards_html += (
            f'\n  <div class="query-card" '
            f'style="background:{q["color"]}11;border-color:{q["color"]}33;">'
            f'\n    <div class="query-label" style="color:{q["color"]};">{q["label"]}</div>'
            f'\n    <div class="query-short">{q["short"]}</div>'
            f'\n    <div class="query-preview">{q["query"]}</div>'
            f"\n  </div>"
        )
    return (
        HTML.replace("__QUERY_CARDS__", cards_html)
            .replace("__QUERIES_JSON__", json.dumps(QUERIES))
    )


# ── Routes ────────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    return HTMLResponse(_build_html())


@app.get("/gateway-health")
async def gateway_health():
    import httpx as _httpx
    try:
        r = _httpx.get("http://localhost:8108/health", timeout=3)
        return {"ok": r.status_code == 200}
    except Exception:
        return {"ok": False}


class ChatRequest(BaseModel):
    query: str


def _unwrap_exc(e: BaseException) -> str:
    if hasattr(e, "exceptions") and e.exceptions:
        return _unwrap_exc(e.exceptions[0])
    return f"{type(e).__name__}: {e}"


@app.post("/chat")
async def chat_stream(body: ChatRequest):
    queue: asyncio.Queue[dict] = asyncio.Queue()
    ui_loop = asyncio.get_running_loop()

    def on_progress(msg: str) -> None:
        if msg.startswith("GRAPH|"):
            try:
                data = json.loads(msg[6:])
                ui_loop.call_soon_threadsafe(
                    queue.put_nowait, {"type": "graph_update", "data": data}
                )
            except Exception:
                pass
        else:
            ui_loop.call_soon_threadsafe(
                queue.put_nowait, {"type": "progress", "text": msg}
            )

    def _run_flow_blocking() -> str:
        async def _inner() -> str:
            return await flow.run(body.query, on_progress=on_progress)
        return asyncio.run(_inner())

    async def produce() -> None:
        try:
            result = await asyncio.to_thread(_run_flow_blocking)
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
    print("Starting PADMA S8 UI at http://localhost:8000")
    print("Gateway V8 must be running at http://localhost:8108")
    uvicorn.run(app, host="0.0.0.0", port=8000)
