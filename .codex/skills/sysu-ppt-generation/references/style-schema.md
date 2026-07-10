# Style Schema

This project uses `templates/styles/style-index.json` as the public style registry and `templates/styles/<style-id>/style.json` as the per-style generation contract.

## Selection Profile

Every entry in `style-index.json` includes a compact `selection_profile` so an
agent can shortlist styles without loading every full style spec:

| Field | Meaning |
|---|---|
| `mood` | Short descriptive signals such as calm, structured, or technical. |
| `tone` | Intended presentation voice and subject signals. |
| `formality` | `low`, `medium-low`, `medium`, `medium-high`, or `high`. |
| `delivery_modes` | Non-empty subset of `speaker_led` and `reading_first`. |
| `surface_scheme` | `light`, `dark`, or `mixed`; current SYSU styles are light. |
| `best_for` | Positive selection guidance. |
| `avoid_for` | Situations where the style should not be shortlisted. |

Read the index first. Open full `style.json` files only after a style is
selected, except that the preview generator may read shortlisted specs
internally to draw the three preview slides.

## Required Fields

Every `style.json` must include:

| Field | Meaning |
|---|---|
| `id` | Stable style identifier used in prompts, mappings, and output records. |
| `name` | Human-readable style name. |
| `group` | One of `strict-original`, `beamer-inspired`, or `beamer-candidates`. |
| `source` | Source PPTX that provides official assets, colors, and identity. |
| `demo_pptx` | Showcase deck used for human style review. |
| `asset_manifest` | Extracted source-asset manifest to use before external imagery. |
| `slide_size_inches` | Physical slide size for compatibility checks. |
| `fonts` | Primary, fallback, or observed font guidance. |
| `palette_or_colors` | Palette, extracted theme colors, or candidate palette variants. |
| `use_case` | Short statement of when the style should be selected. |
| `generation_status` | Current production status. |
| `rules` | Generation rules that downstream agents must apply. |
| `taste_profile` | PPT taste mode and three dial values, including shape and repetition rules. |
| `layout_pattern_ids` | Pattern IDs that this style can use for repeatable composition. |
| `anti_patterns` | Style-specific visual defaults that generation agents must avoid. |

Optional fields such as `template_pptx`, `contact_sheet`, `reference`, `assets`, and `status_reason` may be present when relevant.

## Recommended Optional Fields

Use these fields when a style is intended for repeatable deck generation or visual review:

| Field | Meaning |
|---|---|
| `visual_tokens` | PPT design-system roles for color, typography, spacing, and component reuse. |
| `qa_focus` | Visual and scientific checks that should be recorded in `qa-notes.md`. |
| `sample_page_types` | Page types demonstrated by the style showcase, such as long Chinese title, two-column lecture, evidence figure, table, flow, summary, references, or backup. |

These fields are additive. Older agents may ignore them, but newer agents should read them before creating custom shapes or deciding whether a style is suitable for a teaching or scientific-report task.

## Generation Status

Use these values:

| Status | Meaning |
|---|---|
| `ready` | Can be used as a production style when the user intent matches. |
| `style_selection_only` | Showcase/reference for selecting a direction; do not use as default production template. |

## Agent Reading Order

1. Read `templates/styles/style-index.json`.
2. Shortlist by `selection_profile`, `use_case`, `group`, and `generation_status`.
3. Open the selected `style.json`.
4. For `strict-original`, read `source`, `asset_manifest`, and `template-inventory.md` before mapping slides.
5. For `beamer-inspired`, start from `template_pptx` and use the listed source assets.
6. For `beamer-candidates`, treat the deck as visual selection material unless the user explicitly promotes it to a production direction.
7. Read `visual_tokens`, `sample_page_types`, and `qa_focus` when available, then copy the selected `style.json` into the output folder.
8. Read `taste_profile`, choose a `layout_pattern_id` for every slide, and record deck-level overrides in `outline.md`.

When style discovery is required, record all options and the final choice in
`style-selection.json`. A selected `style_selection_only` style needs
`approval_scope: deck_only`; this never changes the registry status.

## Taste Profile Contract

- `mode`: one of `preserve`, `evolve`, or `selection`.
- `layout_variance`, `visual_density`, `visual_energy`: integers from 1 through 10.
- `shape_system`: `source_defined`, `sharp`, or `subtle`.
- `layout_repetition_limit`: integer from 1 through 3; current active styles use 2.
