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

Begin with:

```markdown
## Design Read

将本任务理解为：...

- Style: `style-id`
- Mode: `preserve|evolve|selection`
- Layout variance: `1-10`
- Visual density: `1-10`
- Visual energy: `1-10`
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

## Visual QA

## Taste QA

## Scientific Figure QA

## Template Fidelity

## Known Limitations
```

`Visual QA` should cover title overflow, body text overflow, element overlap, image aspect ratio, chart readability, footer consistency, and off-brand colors. `Scientific Figure QA` should cover axis labels, units, legends/direct labels, source notes, sample sizes or statistics when relevant, and whether any figure was simplified for projection.

`Taste QA` should record hierarchy, rhythm, restraint, composition, brand fidelity, evidence clarity, pattern repetition, and any justified high-density or source-template exception.
