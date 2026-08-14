from __future__ import annotations

class ArtifactError(Exception):
    def __init__(self, reason: str, details: str = ""):
        super().__init__(details or reason)
        self.reason = reason
        self.details = details or reason


def refusal(reason: str, details: str = "") -> dict:
    return {"status": "refused", "reason": reason, "details": details or reason}
