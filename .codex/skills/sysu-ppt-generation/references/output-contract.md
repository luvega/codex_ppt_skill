# Deck Output Contract

Generated decks should be written under `outputs/<deck-slug>/` and must keep enough intermediate state for review and regeneration.

## Required Files

| File | Required content |
|---|---|
| `deck-brief.json` | Workflow mode, purpose, audience, duration, slide target, content readiness, delivery mode, brand fidelity, requested style, discovery requirement, and content confirmation. |
| `asset-review.json` | Supplied-asset review, or `status: none_provided` with an empty asset list. |
| `outline.md` | Slide-by-slide outline with action titles, content role, exhibit/image need, and selected template/layout. |
| `style.json` | Copy of the selected style spec at generation time. |
| `template-mapping.json` | Mapping from new slide numbers to source slide indices or generated-template layouts. |
| `working.pptx` | Editable intermediate deck after structural slide operations. |
| `replacements.json` | Text/image replacement inventory keyed by slide and shape. |
| `final.pptx` | Final user-facing deck. |
| `qa-notes.md` | Structural and visual QA notes, including any skipped rendering step. |

Conditional files:

| File | When required |
|---|---|
| `source-deck-extract.json` and `extracted-assets/` | `pptx_redesign` or `pptx_revision`. |
| `style-discovery/style-options.pptx` and three preview PNGs | Style discovery is required. |
| `style-discovery/contact-sheet.png` | Style discovery is required and PowerPoint rendering is available. |
| `style-discovery/style-selection.json` | Style discovery is required. |

`style-selection.json` records `render_method`, `approved_asset_count`, the
generator version, and SHA-256 hashes for the brief, asset review, style
registry, and selected style specs. Use `powerpoint` rendering for final visual
selection when Microsoft PowerPoint is available; a `pillow` fallback must be
recorded as a limitation in `qa-notes.md`.

## Deck Brief Contract

`workflow_mode` is one of `new_deck`, `pptx_redesign`, or `pptx_revision`.
`delivery_mode` is `speaker_led` or `reading_first`. `content_readiness` is
`complete`, `rough_notes`, or `topic_only`; `brand_fidelity` is `strict` or
`sysu_evolved`.

The brief also records the real preview content: non-empty `title`, plus
optional `subtitle`, `author`, and `date`. These values, not workflow labels or
style names, are the visible text used by style-discovery slides.

Discovery is required for `new_deck` and `pptx_redesign` when
`requested_style_id` is null. A `pptx_revision` preserves the existing visual
system by default. Confirm extracted content before setting
`content_confirmed` to true for a redesign.

## Asset Review Contract

Each asset record includes `path`, `kind`, `content_summary`, `usable`,
`reason`, `dominant_colors`, `aspect_ratio`, `source`, `planned_slide_roles`,
and `processing`. Never overwrite an original asset. Complete asset review
before the final slide outline so figures and text shape the deck together.
Use `status: none_provided` only with an empty list; otherwise use
`status: reviewed`. The preview generator inserts the first approved usable
image into all three options so visual comparison includes the same asset.

## Outline Requirements

Begin with:

```markdown
## Design Read

将本任务理解为：...

- Style: `style-id`
- Mode: `preserve|evolve|selection`
- Layout variance: `1-10`
- Visual density: `1-10`
- Visual energy: `1-10`
- Delivery mode: `speaker_led|reading_first`
```

Use this table shape:

```markdown
| New Slide | Action Title | Content Role | Layout Pattern ID | Exhibit/Image | Template File | Template Slide/Layout |
|---:|---|---|---|---|---|---|
```

Action titles should state the slide claim or teaching point. Avoid topic-only titles such as `Methods` unless the slide is a section divider.

## QA Requirements

`qa-notes.md` must record:

- Selected `style-id` and `generation_status`.
- Source PPTX or generated template used.
- Slide count and slide size.
- Placeholder coverage or intentionally cleared placeholders.
- Asset manifest used and major inserted assets.
- Rendering method, thumbnail/PDF review result, or the reason visual rendering was not run.
- Known limitations, especially if `generation_status` is not `ready`.

Use these fixed headings so later reviewers and agents can find the relevant evidence quickly:

```markdown
# QA Notes

## Structural QA

## Content Intake QA

## Style Discovery QA

## Visual QA

## Taste QA

## Scientific Figure QA

## Template Fidelity

## Known Limitations
```

`Visual QA` should cover title overflow, body text overflow, element overlap, image aspect ratio, chart readability, footer consistency, and off-brand colors. `Scientific Figure QA` should cover axis labels, units, legends/direct labels, source notes, sample sizes or statistics when relevant, and whether any figure was simplified for projection.

`Taste QA` should record hierarchy, rhythm, restraint, composition, brand fidelity, evidence clarity, pattern repetition, and any justified high-density or source-template exception.

`Content Intake QA` records asset review, existing-deck extraction, content
confirmation, and source-file immutability. `Style Discovery QA` records the
three preview roles, preview authenticity, final selection, and any
`deck_only` candidate approval.
