"""Content-addressable artifact store. No LLM. Pure file I/O."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

from schemas import Artifact

STORE_DIR = Path("state/artifacts")


def _ensure() -> None:
    STORE_DIR.mkdir(parents=True, exist_ok=True)


def put(data: bytes, content_type: str, source: str, descriptor: str) -> Artifact:
    _ensure()
    sha = hashlib.sha256(data).hexdigest()
    art_id = f"art:{sha[:16]}"
    bin_path = STORE_DIR / f"{sha[:16]}.bin"
    meta_path = STORE_DIR / f"{sha[:16]}.json"

    if not bin_path.exists():
        bin_path.write_bytes(data)

    artifact = Artifact(
        id=art_id,
        content_type=content_type,
        size_bytes=len(data),
        source=source,
        descriptor=descriptor,
    )
    meta_path.write_text(artifact.model_dump_json(indent=2))
    return artifact


def get(art_id: str) -> bytes:
    _ensure()
    prefix = art_id.removeprefix("art:")
    path = STORE_DIR / f"{prefix}.bin"
    if not path.exists():
        raise FileNotFoundError(f"Artifact {art_id} not found")
    return path.read_bytes()


def exists(art_id: str) -> bool:
    _ensure()
    prefix = art_id.removeprefix("art:")
    return (STORE_DIR / f"{prefix}.bin").exists()
