# Deck Output Contract

Generated decks should be written under `outputs/<deck-slug>/` and must keep enough intermediate state for review and regeneration.

## Required Files

| File | Required content |
|---|---|
| `outline.md` | Slide-by-slide outline with action titles, content role, exhibit/image need, and selected template/layout. |
| `style.json` | Copy of the selected style spec at generation time. |
| `template-mapping.json` | Mapping from new slide numbers to source slide indices or generated-template layouts. |
| `working.pptx` | Editable intermediate deck after structural slide operations. |
| `replacements.json` | Text/image replacement inventory keyed by slide and shape. |
| `final.pptx` | Final user-facing deck. |
| `qa-notes.md` | Structural and visual QA notes, including any skipped rendering step. |

## Outline Requirements

Use this table shape:

```markdown
| New Slide | Action Title | Content Role | Exhibit/Image | Template File | Template Slide/Layout |
|---:|---|---|---|---|---|
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
