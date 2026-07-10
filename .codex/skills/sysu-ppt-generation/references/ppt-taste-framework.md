# PPT Taste Framework

This framework adapts design-taste ideas to static 16:9 academic PowerPoint. SYSU source templates and the selected `style.json` remain authoritative.

## Design Read

Before mapping slides, write one Chinese sentence that identifies the audience, report type, projection context, content shape, formality, and required brand fidelity.

> 将本任务理解为：面向课题组教师与研究生的 15 分钟科研报告，在普通教室投影，以结果图为主，保持正式、克制并严格使用中山大学蓝色系统。

Record the sentence and the three dial values under `## Design Read` in `outline.md`.

Also record `delivery_mode` from `deck-brief.json`. Use `speaker_led` for live
explanation and `reading_first` for asynchronous or detailed review. When the
need is mixed, choose the closer mode: live delivery defaults to
`speaker_led`, while async circulation defaults to `reading_first`.

## Three Dials

| Dial | 1-3 | 4-6 | 7-8 | 9-10 |
|---|---|---|---|---|
| `layout_variance` | source-faithful, symmetric, predictable | varied academic layouts | editorial/experimental selection | not a default for SYSU decks |
| `visual_density` | keynote-like, one exhibit | normal teaching/research density | methods or backup density | prohibited in the main projected talk |
| `visual_energy` | calm institutional emphasis | teaching/research contrast | strong section or selection-page tension | not a default for official decks |

`visual_energy` describes static compositional tension, not animation.

## Modes

- `preserve`: keep source master, geometry, typography, colors, and shape language. Use for `strict-*`.
- `evolve`: keep SYSU identity while improving hierarchy, spacing, and pattern rotation. Use for `beamer-sysu-*`.
- `selection`: demonstrate a direction without making it a production default. Use for Beamer candidates.

Never change slide size, logo/wordmark, official palette, source font family, footer, or master identity silently.

## Deck-Level Locks

- One accent system, one shape/radius system, one line-weight system, and one shadow policy per deck.
- Do not repeat one layout pattern on three consecutive content slides.
- Main decks with eight or more slides should use at least four layout pattern families.
- Cards require semantic grouping. Prefer whitespace, alignment, rules, or a single visual region when elevation adds no meaning.
- Chinese titles break at semantic phrases. Do not add tracking to Chinese or solve overflow by shrinking body text.
- Real evidence outranks decoration: use research figures, tables, diagrams, screenshots, or relevant images before stock icons.

## Density Budgets

The budgets are defaults, not substitutes for visual review.

| Density | Main-slide guidance |
|---:|---|
| 1-3 | one message, one exhibit, normally no more than three supporting lines |
| 4-6 | one claim plus one exhibit or two coordinated content regions |
| 7-8 | methods, dense comparison, references, or backup only; label the exception in QA |
| 9-10 | split the slide before delivery |

## Delivery Modes

| Mode | Main-slide behavior |
|---|---|
| `speaker_led` | Density 3-5, one claim or teaching point, one dominant exhibit, normally one to three supporting points. |
| `reading_first` | Density 5-6, self-contained context, annotated figures, comparison tables, or concise explanatory copy. |

Density 7-8 remains limited to methods, references, or backup slides. Both
modes keep substantive text at 16 pt or larger and split content before
shrinking it.

## Anti-Patterns

- Three equal cards created only because the outline has three bullets.
- Nested cards, decorative pills, repeated micro-eyebrows, fake dashboards, ornamental section numbers, and stock-icon filler.
- Mixed corner radii, arbitrary accent switching, pure-black shadows, decorative gradients/glows, or unrelated image treatments.
- Fake screenshots, fake data, or decorative diagrams presented as evidence.
- Repeated topic-only titles; academic slides should use action titles or explicit teaching points.

Source-defined serif titles, centered covers, and distinctive official shapes are allowed when present in the selected source template.

## Reference Boundary

The framework is a concept-level adaptation of `leonxlnx/taste-skill` v2 at commit `b17742737e796305d829b3ad39eda3add0d79060` (MIT). It does not vendor the external skill or import its frontend stack, motion, dark-mode, or web-component rules.

The conditional visual-discovery, delivery-mode, asset-first outline, and PPTX
intake workflow is conceptually adapted from `zarazhangrui/frontend-slides` at
commit `9906a34d640d2111f724544cbc50f7f130569ae1` (MIT). PPTX remains the only
production output.
