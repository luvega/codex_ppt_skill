from __future__ import annotations

import importlib.util
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "validate_project_state.py"
SPEC = importlib.util.spec_from_file_location("validate_project_state", SCRIPT)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


VALID_PROFILE = {
    "mood": ["calm"],
    "tone": ["academic"],
    "formality": "high",
    "delivery_modes": ["speaker_led", "reading_first"],
    "surface_scheme": "light",
    "best_for": "科研汇报",
    "avoid_for": "娱乐展示",
}


def test_selection_profile_accepts_valid_contract() -> None:
    errors: list[str] = []
    MODULE.validate_selection_profile("style-a", VALID_PROFILE, errors)
    assert errors == []


def test_selection_profile_rejects_invalid_modes_and_empty_guidance() -> None:
    profile = dict(VALID_PROFILE)
    profile["delivery_modes"] = ["mixed"]
    profile["best_for"] = ""
    errors: list[str] = []

    MODULE.validate_selection_profile("style-a", profile, errors)

    assert any("delivery_modes" in error for error in errors)
    assert any("best_for" in error for error in errors)
