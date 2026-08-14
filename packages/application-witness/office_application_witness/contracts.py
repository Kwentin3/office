from __future__ import annotations

from typing import Literal, TypedDict

ArtifactType = Literal["docx", "xlsx", "pptx"]
RefusalReason = Literal[
    "application_failure",
    "application_timeout",
    "application_unavailable",
    "stale_snapshot",
    "validation_failure",
]


class WitnessObservation(TypedDict):
    application: Literal["LibreOffice"]
    version: Literal["not_observed"]
    claim: Literal["libreoffice_private_clone_observed"]
    operation: Literal["pdf_render", "recalculation_roundtrip"]
    process_exit: Literal["pass"]
    output_validation: Literal["pass"]
    output_sha256: str
    output_bytes: int
    repair_dialog: Literal["not_observable_headless"]
    formula_recalculation: Literal["requested_not_semantically_verified", "not_applicable"]
    application_normalized_clone: bool | Literal["not_applicable"]
    process_isolation: Literal["trusted_executable_not_sandboxed"]


class WitnessSuccess(TypedDict):
    schema_version: Literal[1]
    status: Literal["ok"]
    artifact_type: ArtifactType
    source_sha256: str
    source_bytes: int
    source_unchanged: Literal[True]
    witness: WitnessObservation
    private_workspace_artifacts_retained: Literal[False]
    microsoft_office_equivalence: Literal["not_claimed"]
    latency_ms: float


class WitnessRefusal(TypedDict):
    schema_version: Literal[1]
    status: Literal["refused"]
    reason: RefusalReason
    details: str


WitnessResult = WitnessSuccess | WitnessRefusal
