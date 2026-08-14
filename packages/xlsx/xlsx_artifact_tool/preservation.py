from __future__ import annotations

from typing import Any, TypedDict

PACKAGE_CELL_OPERATIONS = frozenset({"set_cell_value", "set_cell_formula", "clear_cell", "set_cell_style"})
STRUCTURAL_OPERATIONS = frozenset({"append_rows", "reorder_rows"})
GLOBAL_BLOCKERS = frozenset(
    {
        "external_links",
        "external_relationships",
        "connections",
        "ole_objects",
        "activex",
        "workbook_protection",
        "sheet_protection",
        "macros",
        "signatures",
    }
)

# Executable tests are the authority for this table. "preserve" means the
# operation's package backend has an exact semantic oracle and byte equality
# outside its operation-derived member allowlist. "refuse_affected_sheet"
# means another worksheet is not collateral, but the target worksheet is not
# admitted until a feature-specific row oracle exists.
PRESERVATION_MATRIX: dict[str, dict[str, str]] = {
    operation: {
        "comments": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "data_validations": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "conditional_formatting": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "defined_names": "preserve",
        "merged_ranges": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "drawings": "preserve",
        "charts": "preserve",
        "tables": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "pivots": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "shared_formulas": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "array_formulas": "preserve" if operation in PACKAGE_CELL_OPERATIONS else "refuse_affected_sheet",
        "embedded_packages": "preserve",
    }
    for operation in (*sorted(PACKAGE_CELL_OPERATIONS), *sorted(STRUCTURAL_OPERATIONS))
}


class Admission(TypedDict):
    supported: bool
    unsupported: list[dict[str, Any]]


def admit_operations(inventory: dict[str, Any], operations: list[dict[str, Any]]) -> Admission:
    """Apply format-local global and exact worksheet/member preservation policy."""
    unsupported: list[dict[str, Any]] = []
    features = inventory.get("features", {})
    for feature in sorted(GLOBAL_BLOCKERS):
        if features.get(feature):
            unsupported.append({"feature": feature, "scope": "global"})
    if unsupported:
        return {"supported": False, "unsupported": unsupported}

    findings = inventory.get("findings", [])
    structural = [operation for operation in operations if operation.get("type") in STRUCTURAL_OPERATIONS]
    if structural and inventory.get("findings_truncated"):
        return {"supported": False, "unsupported": [{"feature": "inventory", "scope": "global"}]}

    for operation in operations:
        operation_type = operation.get("type")
        matrix = PRESERVATION_MATRIX.get(operation_type)
        if matrix is None:
            unsupported.append({"feature": "operation", "scope": "global", "operation": operation_type})
            continue
        target_sheet = operation.get("sheet")
        for finding in findings:
            feature = finding.get("feature")
            rule = matrix.get(feature)
            if rule in (None, "preserve"):
                continue
            # A member whose worksheet cannot be derived is not safe to treat as
            # unrelated. Location-aware admission is conservative on ambiguity.
            finding_sheet = finding.get("sheet")
            if rule == "refuse_affected_sheet" and (finding_sheet is None or finding_sheet == target_sheet):
                unsupported.append(
                    {
                        "feature": feature,
                        "scope": finding.get("scope", "member"),
                        "part": finding.get("part"),
                        **({"sheet": finding_sheet} if finding_sheet is not None else {}),
                        "operation": operation_type,
                    }
                )
    unsupported.sort(
        key=lambda item: (
            str(item.get("operation", "")),
            str(item.get("feature", "")),
            str(item.get("scope", "")),
            str(item.get("part", "")),
            str(item.get("sheet", "")),
        )
    )
    return {"supported": not unsupported, "unsupported": unsupported}
