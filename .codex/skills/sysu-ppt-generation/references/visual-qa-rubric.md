# Visual QA Rubric

Use this checklist for every generated SYSU deck and every refreshed showcase preview. Record the result in `qa-notes.md` or in the generation run notes.

## Required Review Items

| Area | Pass condition | Common fix |
|---|---|---|
| Title fit | Chinese and English titles fit without clipping, forced overlap, or unreadable shrinking. | Split the title intentionally, shorten the action title, or move detail to subtitle/notes. |
| Text readability | Body text is normally 16 pt or larger; teaching and report pages prefer 18-22 pt when space allows. | Split the slide or replace paragraphs with a figure/table plus one callout. |
| Element overlap | Text, logos, charts, images, footer, and page numbers do not collide. | Restore margins, reduce component count, or move overflow to backup. |
| Image aspect | Photos, microscopy, screenshots, diagrams, and logos are not stretched. Intentional crops are documented. | Use contain/crop logic with stable frames and preserve aspect ratio. |
| Chart readability | Axis labels, tick labels, legends, annotations, and units remain readable in 16:9 preview images. | Enlarge labels, direct-label series, simplify panels, or split the chart. |
| Citation proximity | Source notes and citations sit near the figure, table, or claim they support. | Move citation from global footer to the relevant visual/callout region. |
| Footer consistency | Footer text, page numbers, and wordmarks stay in the expected location and style. | Reuse the style footer component; avoid placing content in the footer band. |
| Brand color | Colors come from the selected style tokens; no orange/black/medical-AI leftovers or arbitrary palettes. | Replace colors with style tokens from `style.json`. |
| Taste profile | Layout, density, and energy match the selected `taste_profile`. | Split content, change the pattern, or reduce decorative emphasis. |
| Pattern rhythm | No pattern appears on three consecutive content slides; long decks use at least four families. | Remap one slide to a figure, comparison, flow, divider, or summary pattern. |
| Shape consistency | Radius, line weight, and shadow treatment form one coherent system. | Remove isolated pills, mixed radii, and black drop shadows. |
| Evidence integrity | Visual regions contain real evidence or relevant imagery, not stock-icon filler or fake dashboards. | Replace decoration with a real figure, table, diagram, or explicitly empty region. |

## Scientific Figure Checks

- One slide should emphasize one result or one comparison.
- Direct labels are preferred over legends for four or fewer series.
- Use neutral gray for context series and a single accent for the key result.
- Axis labels should normally be 18 pt or larger; tick labels should normally be 16 pt or larger.
- Preserve units, sample size, statistical annotation, and source note when needed for interpretation.
- For multi-panel journal figures, split panels across slides instead of shrinking the whole figure.

## Review Record

Use this compact ledger shape in `qa-notes.md`:

```markdown
| Slide | Check | Result | Fix or note |
|---:|---|---|---|
| 3 | Title fit | pass | Long Chinese title split across two lines. |
| 5 | Chart readability | warn | Tick labels readable in PNG preview; raw source figure not available. |
```

## Taste Pre-Flight

- Design Read and three dial values are recorded.
- Every slide has a `layout_pattern_id` in `template-mapping.json`.
- One accent, shape system, line hierarchy, and shadow policy are used consistently.
- Three equal cards appear only when the content is genuinely three independent peers.
- Chinese titles use semantic line breaks and zero letter spacing.
- Source-defined typography and geometry override generic taste rules in `preserve` mode.
