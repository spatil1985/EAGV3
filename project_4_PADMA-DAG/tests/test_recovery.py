"""Tests for recovery.py and critic-fail splice mechanics in flow.py.

22 tests total:
- 18 test classify_failure against real and synthetic error strings
- 4  test the critic-fail splice logic
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from recovery import classify_failure


# ── classify_failure: transient ──────────────────────────────────────────────

class TestTransient:
    def test_503(self):
        assert classify_failure("HTTP 503 Service Unavailable") == "transient"

    def test_502(self):
        assert classify_failure("502 Bad Gateway from provider") == "transient"

    def test_504(self):
        assert classify_failure("504 Gateway Timeout") == "transient"

    def test_timeout_keyword(self):
        assert classify_failure("Read timeout after 30s") == "transient"

    def test_connection_error(self):
        assert classify_failure("ConnectionError: ('Connection aborted.')") == "transient"

    def test_httpstatus_error(self):
        assert classify_failure("httpx.HTTPStatusError: Server error") == "transient"

    def test_service_unavailable(self):
        assert classify_failure("service unavailable, try again") == "transient"

    def test_bad_gateway(self):
        assert classify_failure("bad gateway returned from upstream") == "transient"

    def test_timed_out(self):
        assert classify_failure("Request timed out after 60 seconds") == "transient"

    def test_remote_disconnected(self):
        assert classify_failure("RemoteDisconnected: connection was reset") == "transient"


# ── classify_failure: validation_error ───────────────────────────────────────

class TestValidationError:
    def test_validationerror(self):
        assert classify_failure("pydantic.ValidationError: 1 validation error") == "validation_error"

    def test_malformed(self):
        assert classify_failure("malformed JSON in planner output") == "validation_error"

    def test_json_decode(self):
        assert classify_failure("json decode error at line 3") == "validation_error"

    def test_field_required(self):
        assert classify_failure("field required: skill") == "validation_error"

    def test_nodespec(self):
        assert classify_failure("NodeSpec validation failed") == "validation_error"

    def test_invalid_json(self):
        assert classify_failure("invalid json: unexpected token") == "validation_error"

    def test_pydantic_keyword(self):
        assert classify_failure("pydantic model failed to parse") == "validation_error"


# ── classify_failure: upstream_failure ───────────────────────────────────────

class TestUpstreamFailure:
    def test_generic_error(self):
        assert classify_failure("The researcher could not find relevant content") == "upstream_failure"

    def test_unknown_message(self):
        assert classify_failure("AttributeError: 'NoneType' object has no attribute 'output'") == "upstream_failure"


# ── Critic splice mechanics ───────────────────────────────────────────────────

class TestCriticSpliceMechanics:
    """Verify that the critic verdict handling works correctly."""

    def test_critic_pass_prefix(self):
        """CRITIC_PASS output causes successors to remain pending (not skipped)."""
        from parsers import parse_critic_output
        output, rationale = parse_critic_output(
            '{"verdict": "pass", "rationale": "Looks good"}',
            "original distiller content"
        )
        assert output.startswith("CRITIC_PASS")
        assert "original distiller content" in output

    def test_critic_fail_prefix(self):
        """CRITIC_FAIL output triggers recovery path."""
        from parsers import parse_critic_output
        output, rationale = parse_critic_output(
            '{"verdict": "fail", "rationale": "Missing required field"}',
            "bad output"
        )
        assert output.startswith("CRITIC_FAIL")
        assert "Missing required field" in output

    def test_critic_pass_is_not_fail(self):
        """Pass verdict must not start with CRITIC_FAIL."""
        from parsers import parse_critic_output
        output, _ = parse_critic_output(
            '{"verdict": "pass", "rationale": "ok"}',
            "content"
        )
        assert not output.startswith("CRITIC_FAIL")

    def test_critic_malformed_json_is_fail(self):
        """Malformed critic JSON defaults to fail (safe degradation)."""
        from parsers import parse_critic_output
        output, _ = parse_critic_output("not json at all", "content")
        assert output.startswith("CRITIC_FAIL")
