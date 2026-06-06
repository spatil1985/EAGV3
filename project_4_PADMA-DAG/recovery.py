"""Recovery classifier — maps a failure string to one of three labels.

Labels:
  transient        Gateway 5xx, network errors, timeouts — surface to user, no re-plan
  validation_error Malformed JSON / Pydantic validation failure — prompt bug, no re-plan
  upstream_failure Anything else — queue one Planner recovery node
"""
from __future__ import annotations

# Keywords matched against the lower-cased error string
_TRANSIENT_KEYWORDS = {
    "503", "502", "504",
    "timeout", "timed out",
    "connection", "connectionerror",
    "bad gateway", "gateway timeout",
    "service unavailable", "httpstatuserror",
    "read timeout", "connect timeout",
    "remotedisconnected", "chunkedencodingerror",
}

_VALIDATION_KEYWORDS = {
    "malformed", "validationerror", "validation error",
    "json decode", "jsondecoded", "invalid json",
    "pydantic", "field required", "value is not",
    "nodespec", "agentresult",
}


def classify_failure(error: str) -> str:
    """Return 'transient', 'validation_error', or 'upstream_failure'."""
    lower = error.lower()

    for kw in _TRANSIENT_KEYWORDS:
        if kw in lower:
            return "transient"

    for kw in _VALIDATION_KEYWORDS:
        if kw in lower:
            return "validation_error"

    return "upstream_failure"
