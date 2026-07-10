# Frontend Slides Adaptation

This project conceptually adapts selected workflow ideas from
`zarazhangrui/frontend-slides` at commit
`9906a34d640d2111f724544cbc50f7f130569ae1` under the MIT License.

## Methods Adopted

- Visual discovery: when a new deck or redesign has no requested style ID,
  show three real title-slide previews before generating the full deck.
- Progressive disclosure: shortlist from the compact `style-index.json`, then
  read the selected style's full `style.json` only after selection.
- Delivery modes: choose `speaker_led` or `reading_first`; split content rather
  than inventing a vague middle mode or shrinking text.
- Asset-first outlining: inspect supplied images and figures before finalizing
  the outline, so evidence and text determine slide structure together.
- PPTX intake confirmation: extract titles, text, notes, images, geometry, and
  crops from an existing deck, then confirm the extracted content before a
  redesign.
- Rendered QA: inspect exported slide images because structural checks alone
  cannot prove that elements do not overlap visually.

## SYSU Translation

The three preview roles are fixed:

| Role | Source |
|---|---|
| `safe` | Domain-matched `strict-*` style. |
| `structured` | Corresponding `beamer-sysu-*` production style. |
| `exploratory` | Best-matched registered candidate, or another `ready` style for high-stakes work. |

Every preview uses the real deck title, subtitle, author, date, and approved
logo or image. Option labels, style IDs, file paths, and workflow notes stay
outside the slide image. Selecting a `style_selection_only` candidate requires
an explicit `deck_only` approval in `style-selection.json`; it does not promote
the candidate globally.

## Delivery Mode Translation

- `speaker_led`: live delivery, density 3-5, one main idea, normally one to
  three supporting points, and more slides when necessary.
- `reading_first`: asynchronous or detailed review, density 5-6 on main slides,
  with annotated figures, comparison tables, and concise explanatory text.
- Density 7-8 remains limited to methods, references, or backup. Density 9-10
  is never used for a main projected slide.

## Excluded Scope

This project does not vendor the external repository or its template pack. It
does not add HTML output, CSS/JavaScript, React or Node.js, browser editing,
animation, responsive reflow, online fonts, Vercel deployment, or a second
presentation runtime. Final production output remains PPTX. SYSU source
templates, assets, fonts, colors, masters, `style-index.json`, `style.json`, and
local generation scripts retain authority.

## Source

- Repository: <https://github.com/zarazhangrui/frontend-slides>
- Core workflow: <https://github.com/zarazhangrui/frontend-slides/blob/9906a34d640d2111f724544cbc50f7f130569ae1/SKILL.md>
- Selection index: <https://github.com/zarazhangrui/frontend-slides/blob/9906a34d640d2111f724544cbc50f7f130569ae1/bold-template-pack/selection-index.json>
- License: <https://github.com/zarazhangrui/frontend-slides/blob/9906a34d640d2111f724544cbc50f7f130569ae1/LICENSE>
