# Style Schema

This project uses `templates/styles/style-index.json` as the public style registry and `templates/styles/<style-id>/style.json` as the per-style generation contract.

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

Optional fields such as `template_pptx`, `contact_sheet`, `reference`, `assets`, and `status_reason` may be present when relevant.

## Generation Status

Use these values:

| Status | Meaning |
|---|---|
| `ready` | Can be used as a production style when the user intent matches. |
| `style_selection_only` | Showcase/reference for selecting a direction; do not use as default production template. |

## Agent Reading Order

1. Read `templates/styles/style-index.json`.
2. Choose a style by `use_case`, `group`, and `generation_status`.
3. Open the selected `style.json`.
4. For `strict-original`, read `source`, `asset_manifest`, and `template-inventory.md` before mapping slides.
5. For `beamer-inspired`, start from `template_pptx` and use the listed source assets.
6. For `beamer-candidates`, treat the deck as visual selection material unless the user explicitly promotes it to a production direction.
