"""Contract-first building blocks for the bounded PPTX composer."""

from .compiler import CompileError, compile_deck
from .contracts import ContractError, validate_deck_spec
from .review_contract import ReviewContractError, validate_review_packet
from .scene_contract import SceneContractError, validate_scene_spec

__all__ = [
    "CompileError",
    "ContractError",
    "ReviewContractError",
    "SceneContractError",
    "compile_deck",
    "validate_deck_spec",
    "validate_review_packet",
    "validate_scene_spec",
]
