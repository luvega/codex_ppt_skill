# Design Rules

## Project Structure

- Source templates: `templates/source/`
- Switchable style kits: `templates/styles/<style-id>/`
- Extracted style assets: `templates/assets/<style-id>/`
- Style showcase PPTX files: `outputs/style-showcase/`
- Template element showcase PPTX files: `outputs/style-showcase/template-elements/`
- Skill scripts and references: `.codex/skills/sysu-ppt-generation/`

Always choose a `style-id` before creating a deck. The style spec controls palette, font, density, and design tone; the source template controls official assets, masters, and placeholder geometry.

## Brand Constraints

- Canvas: 16:9 widescreen, 13.333 x 7.5 in.
- Official template family: five color variants under `templates/source/sysu-official/`.
- Medical AI source template: `templates/source/sysu-medical-ai/人工智能导论（医学五年制）简单模板.pptx`.
- Core fonts are template-specific. Preserve the selected style's extracted primary fonts, including `思源黑体 CN Medium`, `思源黑体 CN Heavy`, `思源宋体 CN Heavy`, `思源宋体 CN Medium`, `黑体`, `微软雅黑`, `Arial`, and `KaiTi` when they are listed in the style spec.
- The smaller Tang Feng template uses Arial plus Microsoft YaHei. Keep that pairing when starting from that file.
- Keep slide footers and numbers unless the selected source slide omits them.

## Style Switching

| Style ID | Palette role | Best use |
|---|---|---|
| `strict-sysu-official-blue` | original official blue, no simplification | exact official SYSU blue template reuse |
| `strict-sysu-official-green` | original official green, no simplification | exact official SYSU green template reuse |
| `strict-sysu-official-red` | original official red, no simplification | exact official SYSU red template reuse |
| `strict-sysu-medical-ai` | original medical AI courseware, no simplification | exact medical AI teaching template reuse |
| `beamer-sysu-blue` | official blue adapted to Beamer/Madrid frame structure | academic reports and technical seminars |
| `beamer-sysu-green` | official green adapted to Beamer/Madrid frame structure | biomedical/life-science talks |
| `beamer-sysu-red` | official red adapted to Beamer/Madrid frame structure | formal academic talks |

The earlier `sysu-minimal-*`, simplified `sysu-medical-ai`, and `lora-hu-academic` style kits have been removed from the active library. Use LoRA Hu and `ppt169_building_effective_agents` only as outside references when explicitly useful; do not expose them as SYSU style IDs. The `strict-*` styles are source-faithful modes backed by extracted fonts, colors, and media manifests.

The `beamer-sysu-*` styles are PPTX adaptations of Beamer academic layout discipline, not strict source-template copies. Use them when the desired result is a Beamer-like academic deck with SYSU identity: top title bar, footer frame number, theorem/example/alert blocks, columns, booktabs-like tables, algorithm frames, diagrams, references, and backup slides.

## Strict Original Rules

- For `strict-sysu-official-blue`, `strict-sysu-official-green`, `strict-sysu-official-red`, and `strict-sysu-medical-ai`, read both the source PPTX inventory and the style's `asset_manifest` before generating.
- Preserve source masters, slide size, theme, logos, footers, page numbers, placeholder geometry, and source typography.
- Do structural work first: duplicate, delete, or reorder slides before replacing content.
- Replace text and images inside existing placeholders whenever possible. Remove unused source elements entirely; do not leave empty labels, sample captions, or mismatched icon groups.
- Use varied original layouts. Avoid mapping every content section to a plain title-and-body slide when the source deck contains dividers, image layouts, comparison layouts, or callout pages.
- Use extracted assets deliberately: choose校徽/wordmark/logo variants, recurring icons, campus photos, building cutouts, and decorative elements from `templates/assets/<style-id>/asset-manifest.json`.
- Lingnan PDF is intentionally not an active strict style because this project does not currently contain an editable Lingnan PPTX source.
- Keep the generated strict showcase cover as the original source PPT first slide. Do not replace it with a generic generated title page.
- Do not add a generic full-height left vertical stripe unless that stripe exists in the selected source layout. This was a previous generated artifact, not a SYSU brand rule.
- For technical content pages, the allowed outside aesthetic reference is `ppt169_building_effective_agents` page rhythm: precise cards, hero visual first, concise interpretation second, and breathing/dense information hierarchy. Do not copy its dark background; keep the selected source template's original background color system and replace fonts, colors, and icons with SYSU extracted fonts, colors, and assets.
- Run a rendered or image-based QA pass when tools are available. If rendering tools are missing, reopen with `python-pptx`, check slide counts and dimensions, and note that visual rendering was limited.

## Color Rules

Use colors from the selected style spec. The blue style is the default for formal SYSU academic/report decks. Do not mix accent colors from multiple styles in the same deck.

For charts, use one primary template accent plus neutral grays. Highlight only the key result. Avoid red/green-only encodings for comparisons.

## Typography Rules

- Title: short, specific, and readable from the back of a room.
- Body text: prefer 20 pt or larger for presentations; do not go below 16 pt except citations or footer metadata.
- Chinese titles need more space than English titles; split lines intentionally.
- Keep one main font family per deck. Use size and weight for hierarchy.
- Do not use emoji as visual markers in official or academic decks.
- For `strict-sysu-medical-ai`, favor the source courseware's clinical clarity, teal/blue structures, computational diagrams, and restrained data emphasis.
- For `beamer-sysu-*`, keep Beamer-like compactness: slide titles at 14-18 pt, body around 10-13 pt, no oversized marketing hero type, and no more than two colored blocks per slide.

## Academic Content Rules

Use these when the deck is a seminar, defense, research talk, lab meeting, conference talk, grant briefing, or paper-to-slides conversion:

- Action titles, not topic labels.
- One claim or insight per slide.
- One main exhibit per results slide.
- Every figure, external claim, and data source needs an in-slide citation.
- End on a conclusions or takeaways slide, not a blank thank-you slide.
- Keep appendix material outside the main talk flow.

## Layout Rules

- Match content count to placeholder count.
- Use image layouts only when real images exist.
- Use quote layouts only for actual quotes with attribution.
- Keep alignment consistent with the template margins.
- Do not draw decorative cards inside existing template cards.
- Prefer existing placeholders over manual text boxes. If manual boxes are necessary, align to the same margins and baseline rhythm.

## Tables, Figures, and Images

- For dense figures, allocate most of the slide to the figure and use a short interpretive callout.
- Do not stack a chart below long bullets. Use a figure-left/text-right or figure-full layout.
- Preserve image aspect ratios unless intentionally cropping into a template frame.
- Keep captions close to the visual they describe.
- Use high-resolution images; avoid blurry screenshots.
