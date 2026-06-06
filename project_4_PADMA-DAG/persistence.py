"""SessionStore — atomic graph + per-node state persistence for Session 8."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import TYPE_CHECKING, Any

import networkx as nx

if TYPE_CHECKING:
    pass

SESSIONS_DIR = Path("state/sessions")


class SessionLoadError(Exception):
    pass


class SessionStore:
    def __init__(self, session_id: str) -> None:
        self.sid = session_id
        self.session_dir = SESSIONS_DIR / session_id
        self.nodes_dir = self.session_dir / "nodes"
        self.session_dir.mkdir(parents=True, exist_ok=True)
        self.nodes_dir.mkdir(parents=True, exist_ok=True)

    # ── Atomic write helper ────────────────────────────────────────────────

    def _atomic_write(self, path: Path, text: str) -> None:
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(text, encoding="utf-8")
        os.replace(tmp, path)

    # ── Query ──────────────────────────────────────────────────────────────

    def save_query(self, query: str) -> None:
        self._atomic_write(self.session_dir / "query.txt", query)

    def load_query(self) -> str:
        p = self.session_dir / "query.txt"
        if not p.exists():
            raise SessionLoadError(f"query.txt not found in {self.session_dir}")
        return p.read_text(encoding="utf-8")

    # ── Graph ──────────────────────────────────────────────────────────────

    def save_graph(self, g: nx.DiGraph) -> None:
        """Serialise the NetworkX DiGraph to graph.json via node_link_data."""
        data = nx.node_link_data(g)
        # Serialise any NodeState objects stored as node attributes
        for node in data.get("nodes", []):
            if "state" in node and hasattr(node["state"], "model_dump"):
                d = node["state"].model_dump(mode="json")
                d["_state_typed"] = True
                node["state"] = d
        self._atomic_write(
            self.session_dir / "graph.json",
            json.dumps(data, indent=2, default=str),
        )

    def load_graph(self) -> nx.DiGraph:
        """Revive a DiGraph from graph.json. Raises SessionLoadError on failure."""
        graph_path = self.session_dir / "graph.json"
        if not graph_path.exists():
            raise SessionLoadError(f"graph.json not found in {self.session_dir}")
        try:
            data = json.loads(graph_path.read_text(encoding="utf-8"))
            g = nx.node_link_graph(data, directed=True, multigraph=False)
            return g
        except Exception as exc:
            raise SessionLoadError(f"Failed to load graph: {exc}") from exc

    # ── Node state ─────────────────────────────────────────────────────────

    def save_node(self, state: Any) -> None:
        """Persist a NodeState to nodes/<node_id>.json atomically."""
        node_id = state.node_id
        path = self.nodes_dir / f"{node_id}.json"
        self._atomic_write(path, state.model_dump_json(indent=2))

    def load_node(self, node_id: str) -> dict:
        """Load a raw node state dict. Caller validates into NodeState."""
        path = self.nodes_dir / f"{node_id}.json"
        if not path.exists():
            raise SessionLoadError(f"Node file not found: {path}")
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            raise SessionLoadError(f"Failed to load node {node_id}: {exc}") from exc
