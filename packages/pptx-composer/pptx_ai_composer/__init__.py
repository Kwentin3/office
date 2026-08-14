"""Contract-first building blocks for the bounded PPTX composer."""

from .compiler import CompileError, compile_deck
from .contracts import ContractError, validate_deck_spec
from .scene_contract import SceneContractError, validate_scene_spec

__all__ = [
    "CompileError",
    "ContractError",
    "SceneContractError",
    "compile_deck",
    "validate_deck_spec",
    "validate_scene_spec",
]
