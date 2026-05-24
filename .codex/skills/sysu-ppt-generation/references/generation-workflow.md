# Generation Workflow

## Output Structure

Create generated decks under `outputs/<deck-slug>/`:

```text
outputs/<deck-slug>/
  outline.md
  style.json
  template-mapping.json
  working.pptx
  replacements.json
  final.pptx
  qa-notes.md
```

Keep `templates/source/` immutable. Copy a template into the output folder before editing.

## Style Selection

Read `templates/styles/style-index.json` first. If the user does not specify a style, use this decision order:

| Content | Style ID |
|---|---|
| General SYSU academic/report deck | `strict-sysu-official-blue` |
| Biomedical, public health, life-science | `strict-sysu-official-green` |
| Official ceremonial or policy/reporting tone | `strict-sysu-official-red` |
| Medical AI teaching/courseware | `strict-sysu-medical-ai` |
| Technical ML paper explanation using SYSU branding | choose the closest `strict-*` style, then borrow only content-page rhythm from `ppt169_building_effective_agents` |
| User asks for Beamer-like academic PPTX structure | `beamer-sysu-blue`, `beamer-sysu-green`, or `beamer-sysu-red` |

Copy the chosen `style.json` into the output folder so the deck is reproducible.

The source template and extracted asset manifest are the authority. Use the style spec to locate the source PPTX, `asset_manifest`, contact sheet, and demo file. Do not reduce a strict style to a copied template or a palette swap. The removed `sysu-minimal-*` styles are not active generation targets.

For generated strict showcase decks, keep slide 1 as the unmodified first slide from the source PPTX. Subsequent explanatory pages may learn from `ppt169_building_effective_agents` page rhythm, but must use the original source-template background colors plus SYSU fonts, extracted SYSU marks/images/elements, and the selected official or medical-AI color system. Do not apply the external deck's dark background. Do not add a generic full-height left vertical stripe unless the selected source slide itself uses that exact motif.

For `beamer-sysu-*` styles, use the generated PPTX template under `templates/generated/beamer-inspired/`. These styles intentionally do not preserve source slide geometry; they adapt Beamer/Madrid academic frame structure to SYSU assets. Keep compact density, top frame title, footer page number, at most two colored blocks, and at least one substantive element per content slide.

## Outline First

Write `outline.md` before touching the PPTX. Use this table:

```markdown
| New Slide | Action Title | Content Role | Exhibit/Image | Template File | Template Slide/Layout |
|---:|---|---|---|---|---|
| 1 | ... | cover | ... | ... | slide 0 / layout 0 |
```

For research or teaching decks, use action titles. A weak title says `Methods`; a strong title says `Single-cell profiles reveal an exhausted T cell state enriched after treatment`.

## Template Mapping

Use zero-based slide indices from `template-inventory.md`.

Prefer mappings that reuse complete existing template slides when the target page resembles a sample slide. If a deck only needs normal placeholders, choose layouts instead:

| Content shape | Preferred official layout |
|---|---|
| Cover | `标题幻灯片` |
| Chapter divider | `节标题` |
| Single idea, body text, one figure/table | `标题和内容` |
| Two distinct concepts | `两栏内容` |
| Before/after or two options | `比较` |
| Image plus explanatory text | `图片与标题` or `内容与标题` |
| Sparse statement | `仅标题` |
| Custom chart or full figure | `空白`, but keep footer/slide number if required |

Do not force 3+ items into a two-column or comparison slide. Split across pages or use a custom grid on a blank layout.

## Editing Path

Use this order:

1. Copy the selected PPTX to `working.pptx`.
2. Create or reorder slides according to `template-mapping.json`.
3. Extract text shape inventory from `working.pptx`.
4. Write `replacements.json` with one entry per slide/shape that should contain text.
5. Replace text while preserving paragraph properties, font sizes, alignment, and theme colors where possible.
6. Insert images into picture placeholders or stable image frames.
7. Save `final.pptx`.

Use `python-pptx` for normal placeholder filling. If it cannot preserve a feature, unpack the PPTX and edit OOXML directly. Validate the PPTX by reopening it with `python-pptx` and rerunning `extract_pptx_template.py` on the output.

## Strict Original Editing Path

Use this path when the selected style ID begins with `strict-`:

1. Copy the source PPTX to `working.pptx`.
2. Read the selected style's `asset_manifest` and contact sheet. Pick校徽/wordmark variants, recurring marks, campus photos, cutouts, and decorative elements from extracted assets before searching elsewhere.
3. Render or thumbnail the source when tools are available, then choose source slide layouts before editing.
4. Complete slide structure first: duplicate, delete, reorder, and rename mapping entries.
5. Replace content only after structure is stable. Keep source text properties, bullets, placeholder positions, image masks, logos, footers, and numbering.
6. Preserve the source typography from `style.json`, including the extracted primary fonts such as `思源宋体 CN Heavy` or `思源黑体 CN Heavy` when listed.
7. Remove excess source groups when the new content has fewer items than the template.
8. Reopen the output with `python-pptx`; compare slide count, size, source style ID, placeholder coverage, font family coverage, and asset usage with `template-mapping.json`.
9. Run visual QA or explicitly record why visual rendering was not available.

## Visual QA

When render tools are available:

1. Convert final PPTX to PDF or images.
2. Build a thumbnail grid.
3. Check every slide for text cutoff, overlaps, wrong fonts, broken images, footer collisions, and off-brand colors.
4. Fix and rerender until the issue list is empty.

When render tools are not available, still perform structural QA and explicitly note that visual rendering was not run.
