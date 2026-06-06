"""Sandbox executor — run Python code from Coder in an isolated subprocess.

Safety boundary: catches mistakes, not attacks.
- Scrubbed env (PATH, HOME, LANG, LC_ALL, LC_CTYPE only)
- Fresh temp directory (deleted on exit)
- Stdout/stderr capped at 1 MB each
- 30-second timeout
- Does NOT restrict network access or filesystem outside temp dir
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path

TIMEOUT_SEC = 30
MAX_OUTPUT_BYTES = 1024 * 1024  # 1 MB

_ALLOWED_ENV_VARS = {"PATH", "HOME", "LANG", "LC_ALL", "LC_CTYPE"}


def run_code(code: str) -> dict:
    """
    Execute Python code in a sandboxed subprocess.

    Returns:
        {
          "stdout": str,
          "stderr": str,
          "exit_code": int,
          "elapsed_sec": float,
          "timed_out": bool,
        }
    """
    import time

    # Scrub environment
    clean_env = {k: v for k, v in os.environ.items() if k in _ALLOWED_ENV_VARS}
    clean_env.setdefault("PATH", "/usr/bin:/bin")

    with tempfile.TemporaryDirectory(prefix="padma_sandbox_") as tmpdir:
        script_path = Path(tmpdir) / "script.py"
        script_path.write_text(code, encoding="utf-8")

        start = time.monotonic()
        timed_out = False
        try:
            proc = subprocess.run(
                [sys.executable, str(script_path)],
                capture_output=True,
                timeout=TIMEOUT_SEC,
                cwd=tmpdir,
                env=clean_env,
            )
            elapsed = time.monotonic() - start
            stdout = proc.stdout[:MAX_OUTPUT_BYTES].decode("utf-8", errors="replace")
            stderr = proc.stderr[:MAX_OUTPUT_BYTES].decode("utf-8", errors="replace")
            exit_code = proc.returncode
        except subprocess.TimeoutExpired:
            elapsed = time.monotonic() - start
            timed_out = True
            stdout = ""
            stderr = f"TimeoutExpired: script exceeded {TIMEOUT_SEC}s"
            exit_code = -1

    return {
        "stdout": stdout,
        "stderr": stderr,
        "exit_code": exit_code,
        "elapsed_sec": round(elapsed, 3),
        "timed_out": timed_out,
    }
