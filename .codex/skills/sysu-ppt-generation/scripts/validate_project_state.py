#!/usr/bin/env python
"""Validate the SYSU PPT generation project state without modifying files."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from typing import Any

from pptx import Presentation


ROOT = Path(__file__).resolve().parents[4]
STYLE_INDEX = ROOT / "templates" / "styles" / "style-index.json"
INVENTORY = ROOT / ".codex" / "skills" / "sysu-ppt-generation" / "references" / "template-inventory.json"

REQUIRED_STYLE_FIELDS = {
    "id",
    "name",
    "group",
    "source",
    "demo_pptx",
    "asset_manifest",
    "slide_size_inches",
    "fonts",
    "palette_or_colors",
    "use_case",
    "generation_status",
    "rules",
    "taste_profile",
    "layout_pattern_ids",
    "anti_patterns",
}

TASTE_MODES = {"preserve", "evolve", "selection"}
REQUIRED_TASTE_FIELDS = {"mode", "layout_variance", "visual_density", "visual_energy", "shape_system", "layout_repetition_limit"}
PATTERN_ROOT = ROOT / ".codex" / "skills" / "sysu-ppt-generation" / "references" / "layout-patterns"
REQUIRED_SELECTION_FIELDS = {
    "mood",
    "tone",
    "formality",
    "delivery_modes",
    "surface_scheme",
    "best_for",
    "avoid_for",
}
DELIVERY_MODES = {"speaker_led", "reading_first"}
FORMALITY_VALUES = {"low", "medium-low", "medium", "medium-high", "high"}
SURFACE_SCHEMES = {"light", "dark", "mixed"}
FRONTEND_SLIDES_COMMIT = "9906a34d640d2111f724544cbc50f7f130569ae1"
ADAPTATION_REFERENCE = ROOT / ".codex" / "skills" / "sysu-ppt-generation" / "references" / "frontend-slides-adaptation.md"
OUTPUT_CONTRACT = ROOT / ".codex" / "skills" / "sysu-ppt-generation" / "references" / "output-contract.md"
VISUAL_DISCOVERY_DIR = ROOT / "outputs" / "style-showcase" / "visual-discovery"

STALE_PATTERNS = [
    "/".join(["F:", "AI_PPT"]),
    "中山大学幻灯片模板-" + "橙",
    "中山大学幻灯片模板-" + "黑",
    "five color " + "variants",
    "strict-sysu-" + "medical-ai",
    "sysu-" + "medical-ai",
    "needs_" + "resync",
    "Medical " + "AI",
    "medical " + "AI",
    "医学" + "人工智能",
    "医学" + "五年制",
    "人工智能" + "导论",
]

TEXT_SUFFIXES = {
    ".md",
    ".json",
    ".py",
    ".ps1",
    ".yaml",
    ".yml",
    ".txt",
    ".toml",
}


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def add_path_error(errors: list[str], label: str, owner: str, path_value: str | None) -> None:
    if not path_value:
        errors.append(f"{owner}: missing {label}")
        return
    path = ROOT / path_value
    if not path.exists():
        errors.append(f"{owner}: {label} does not exist: {path_value}")


def validate_selection_profile(owner: str, profile: Any, errors: list[str]) -> None:
    if not isinstance(profile, dict):
        errors.append(f"{owner}: selection_profile must be an object")
        return
    missing = sorted(REQUIRED_SELECTION_FIELDS - set(profile))
    if missing:
        errors.append(f"{owner}: selection_profile missing fields: {', '.join(missing)}")
    for field in ("mood", "tone"):
        value = profile.get(field)
        if not isinstance(value, list) or not value or not all(isinstance(item, str) and item.strip() for item in value):
            errors.append(f"{owner}: selection_profile {field} must be a non-empty string list")
    if profile.get("formality") not in FORMALITY_VALUES:
        errors.append(f"{owner}: selection_profile formality is invalid: {profile.get('formality')}")
    delivery_modes = profile.get("delivery_modes")
    if not isinstance(delivery_modes, list) or not delivery_modes or not set(delivery_modes).issubset(DELIVERY_MODES):
        errors.append(f"{owner}: selection_profile delivery_modes must use speaker_led/reading_first")
    if profile.get("surface_scheme") not in SURFACE_SCHEMES:
        errors.append(f"{owner}: selection_profile surface_scheme is invalid: {profile.get('surface_scheme')}")
    for field in ("best_for", "avoid_for"):
        if not isinstance(profile.get(field), str) or not profile.get(field, "").strip():
            errors.append(f"{owner}: selection_profile {field} must be non-empty")


def validate_style_registry(errors: list[str], warnings: list[str]) -> dict[str, dict[str, Any]]:
    index = load_json(STYLE_INDEX)
    styles: dict[str, dict[str, Any]] = {}
    for entry in index.get("styles", []):
        sid = entry.get("id", "<missing-id>")
        style_spec = entry.get("style_spec")
        add_path_error(errors, "style_spec", sid, style_spec)
        for label in ("source", "demo_pptx", "asset_manifest"):
            add_path_error(errors, label, sid, entry.get(label))
        if entry.get("template_pptx"):
            add_path_error(errors, "template_pptx", sid, entry.get("template_pptx"))
        if not style_spec or not (ROOT / style_spec).exists():
            continue

        style = load_json(ROOT / style_spec)
        styles[sid] = style
        missing = sorted(REQUIRED_STYLE_FIELDS - set(style))
        if missing:
            errors.append(f"{sid}: style.json missing fields: {', '.join(missing)}")
        if style.get("id") != sid:
            errors.append(f"{sid}: style.json id mismatch: {style.get('id')}")
        if style.get("source") != entry.get("source"):
            errors.append(f"{sid}: style source differs from style-index source")
        if style.get("demo_pptx") != entry.get("demo_pptx"):
            errors.append(f"{sid}: style demo_pptx differs from style-index demo_pptx")
        if style.get("asset_manifest") != entry.get("asset_manifest"):
            errors.append(f"{sid}: style asset_manifest differs from style-index asset_manifest")
        if entry.get("generation_status") != style.get("generation_status"):
            errors.append(f"{sid}: generation_status differs between style-index and style.json")
        validate_selection_profile(sid, entry.get("selection_profile"), errors)
        profile = style.get("taste_profile", {})
        missing_profile = sorted(REQUIRED_TASTE_FIELDS - set(profile))
        if missing_profile:
            errors.append(f"{sid}: taste_profile missing fields: {', '.join(missing_profile)}")
        if profile.get("mode") not in TASTE_MODES:
            errors.append(f"{sid}: invalid taste mode: {profile.get('mode')}")
        for key in ("layout_variance", "visual_density", "visual_energy"):
            value = profile.get(key)
            if not isinstance(value, int) or not 1 <= value <= 10:
                errors.append(f"{sid}: {key} must be an integer from 1 to 10")
        for pattern_id in style.get("layout_pattern_ids", []):
            if not (PATTERN_ROOT / f"{pattern_id}.md").exists():
                errors.append(f"{sid}: missing layout pattern file: {pattern_id}.md")
        if not isinstance(style.get("anti_patterns"), list) or not style.get("anti_patterns"):
            errors.append(f"{sid}: anti_patterns must be a non-empty list")

    for label in ("style_discovery_showcase_pptx", "style_discovery_contact_sheet"):
        add_path_error(errors, label, "style-index", index.get(label))

    return styles


def validate_visual_discovery_contract(errors: list[str]) -> None:
    for path in (ADAPTATION_REFERENCE, OUTPUT_CONTRACT):
        if not path.exists():
            errors.append(f"missing Visual Discovery reference: {rel(path)}")
            return
    adaptation = ADAPTATION_REFERENCE.read_text(encoding="utf-8")
    if FRONTEND_SLIDES_COMMIT not in adaptation:
        errors.append("frontend-slides adaptation does not pin the required commit")
    contract = OUTPUT_CONTRACT.read_text(encoding="utf-8")
    for token in ("deck-brief.json", "asset-review.json", "style-selection.json", "Content Intake QA", "Style Discovery QA"):
        if token not in contract:
            errors.append(f"output contract missing Visual Discovery token: {token}")


def validate_visual_discovery_showcase(errors: list[str]) -> None:
    script = ROOT / ".codex" / "skills" / "sysu-ppt-generation" / "scripts" / "validate_style_discovery.py"
    result = subprocess.run(
        [
            sys.executable,
            str(script),
            str(VISUAL_DISCOVERY_DIR),
            "--brief",
            str(VISUAL_DISCOVERY_DIR / "deck-brief.json"),
            "--asset-review",
            str(VISUAL_DISCOVERY_DIR / "asset-review.json"),
            "--require-confirmed",
            "--require-powerpoint-rendered",
        ],
        cwd=ROOT,
        text=True,
        capture_output=True,
        encoding="utf-8",
    )
    if result.returncode:
        messages = [line.removeprefix("ERROR: ") for line in result.stdout.splitlines() if line.startswith("ERROR:")]
        errors.extend(f"Visual Discovery showcase: {message}" for message in (messages or ["validation failed"]))


def validate_inventory(errors: list[str]) -> None:
    inventory = load_json(INVENTORY)
    for item in inventory.get("templates", []):
        file_value = item.get("file", "")
        if not file_value.startswith("templates/source/"):
            errors.append(f"template inventory includes non-source PPTX: {file_value}")
        if not (ROOT / file_value).exists():
            errors.append(f"template inventory file does not exist: {file_value}")


def validate_stale_text(errors: list[str]) -> None:
    skip_parts = {".git", "__pycache__"}
    for path in ROOT.rglob("*"):
        if path.is_dir() or skip_parts.intersection(path.parts):
            continue
        if path.suffix.lower() not in TEXT_SUFFIXES:
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        for pattern in STALE_PATTERNS:
            if pattern in text:
                errors.append(f"{rel(path)}: stale text remains: {pattern}")


def pptx_size(path: Path) -> tuple[float, float, float]:
    prs = Presentation(str(path))
    width = round(prs.slide_width / 914400, 3)
    height = round(prs.slide_height / 914400, 3)
    aspect = round(width / height, 3) if height else 0.0
    return width, height, aspect


def validate_pptx_sizes(errors: list[str], warnings: list[str]) -> None:
    pptx_paths = sorted((ROOT / "outputs" / "style-showcase").rglob("*.pptx"))
    pptx_paths += sorted((ROOT / "templates" / "generated" / "beamer-inspired").glob("*.pptx"))
    for path in pptx_paths:
        width, height, aspect = pptx_size(path)
        if abs(aspect - 1.778) > 0.01:
            errors.append(f"{rel(path)}: not 16:9 aspect ({width} x {height})")
        expected_size = abs(width - 13.333) <= 0.02 and abs(height - 7.5) <= 0.02
        if not expected_size:
            errors.append(f"{rel(path)}: unexpected physical size {width} x {height}")


def main() -> int:
    errors: list[str] = []
    warnings: list[str] = []
    validate_style_registry(errors, warnings)
    validate_inventory(errors)
    validate_stale_text(errors)
    validate_pptx_sizes(errors, warnings)
    validate_visual_discovery_contract(errors)
    validate_visual_discovery_showcase(errors)

    for warning in warnings:
        print(f"WARN: {warning}")
    for error in errors:
        print(f"ERROR: {error}")

    if errors:
        print(f"Validation failed: {len(errors)} error(s), {len(warnings)} warning(s)")
        return 1
    print(f"Validation passed: 0 errors, {len(warnings)} warning(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
