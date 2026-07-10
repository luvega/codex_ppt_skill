---
name: sysu-ppt-generation
description: Use when Codex is asked to generate, edit, analyze, or refresh Sun Yat-sen University / SYSU / 中山大学 / 岭南学院 PowerPoint decks in this project from the local PPTX templates, preserve the official university slide style, extract template elements, or map an outline to reusable template slides.
---

# SYSU PPT Generation

## Overview

Generate PowerPoint decks by starting from the local SYSU templates and style kits in this project, not from a blank deck. Preserve official typography, colors, masters, placeholders, and slide geometry unless the user explicitly asks for a redesign.

## First Moves

1. Work from the project root that contains `templates/source/`, `templates/styles/`, `outputs/`, and `.codex/`.
2. Read `references/template-inventory.md` or `references/template-inventory.json` before choosing template slides.
3. Read `templates/styles/style-index.json` before choosing a style. Use the user's requested style ID when provided.
4. If templates were added or changed, refresh the inventory:

```powershell
python .codex\skills\sysu-ppt-generation\scripts\extract_pptx_template.py --root . --out-dir .codex\skills\sysu-ppt-generation\references
```

5. If strict style specs or showcase decks need refreshing, run:

```powershell
python .codex\skills\sysu-ppt-generation\scripts\generate_strict_original_showcases.py
```

6. If Beamer-inspired SYSU templates need refreshing, run:

```powershell
python .codex\skills\sysu-ppt-generation\scripts\generate_beamer_inspired_templates.py
```

7. If Beamer candidate showcase decks need refreshing, run:

```powershell
python .codex\skills\sysu-ppt-generation\scripts\generate_beamer_candidate_showcases.py
```

8. If README preview images need refreshing and PowerPoint is available on Windows, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .codex\skills\sysu-ppt-generation\scripts\export_readme_previews.ps1
```

9. Read `references/generation-workflow.md` before creating a deck.
10. Read `references/style-schema.md` before interpreting `style.json` fields or `generation_status`.
11. Read `references/output-contract.md` before creating a new deck output folder.
12. Read `references/design-rules.md` when choosing colors, fonts, slide patterns, or academic presentation structure.
13. Read `references/visual-qa-rubric.md` before final visual review or when writing `qa-notes.md`.
14. Read `references/ppt-taste-framework.md` before choosing freeform layouts or changing an existing deck's visual language.
15. Read `references/layout-patterns/README.md`, then open only the selected pattern files before drawing custom shapes.
16. Read `references/reference-skill-notes.md` only when you need the rationale from the referenced third-party skills.
17. Read `references/frontend-slides-adaptation.md` when style discovery, delivery mode, asset-first outlining, or existing-PPTX redesign is involved.

## Style Choice

This project keeps source-derived strict style IDs, the earlier Beamer/Madrid adaptation, and a separate Beamer candidate showcase group for style selection. The earlier generated `sysu-minimal-*` showcase styles were removed because they were not faithful enough to the source templates.

| Style ID | Use when |
|---|---|
| `strict-sysu-official-blue` | Strict original official blue template, preserving source masters and source slide geometry |
| `strict-sysu-official-green` | Strict original official green template, preserving source masters and source slide geometry |
| `strict-sysu-official-red` | Strict original official red template, preserving source masters and source slide geometry |
| `beamer-sysu-blue` | Beamer/Madrid-inspired SYSU blue academic template with title bar, footline, blocks, tables, diagrams |
| `beamer-sysu-green` | Beamer/Madrid-inspired SYSU green academic template for biomedical/life-science talks |
| `beamer-sysu-red` | Beamer/Madrid-inspired SYSU red academic template for formal talks |
| `simpleplus-sysu-clean` | SimplePlus-inspired clean academic candidate; white canvas, light rules, restrained SYSU identity |
| `ustc-thu-sysu-institutional` | USTC/THU-inspired formal Chinese university candidate; stronger institutional header and navigation |
| `moloch-sysu-minimal` | Moloch/Metropolis-inspired minimal candidate; light background, large hierarchy, progress bar |
| `sleek-sysu-research` | Sleek-inspired research candidate for algorithms, code, methods, and technical data |
| `river-sysu-atelier` | River/Beamer Atelier-inspired structured teaching candidate with outline, section, and environment blocks |

Each style has a spec under `templates/styles/<style-id>/style.json` and a showcase PPTX under `outputs/style-showcase/`.

Use the ready `strict-*` styles for normal generation in this project. Treat the source PPTX plus `asset_manifest` as the style authority: preserve source layouts and typography, use extracted校徽/标志物/校区图片/绘图元素, and derive new slides from those assets rather than copying the source file as the final answer. Lingnan PDF is not an active strict style unless an editable PPTX source is added.

For strict showcase/demo decks, keep slide 1 as the original source PPT cover. Do not replace it with a synthetic generated cover, and do not add a generic full-height left vertical stripe unless that exact element exists in the selected source layout. `ppt169_building_effective_agents` may be used only as a content-page rhythm reference: precise cards, hero visual first, concise interpretation second, and clear information hierarchy. Do not copy its dark background, Anthropic palette, icons, or typography. Keep each strict style's source-template background colors, SYSU fonts, extracted SYSU assets, and official SYSU colors. Do not mix the LoRA Hu palette or LoRA-specific page grammar into SYSU strict styles unless the user explicitly asks for that hybrid.

Use `beamer-sysu-*` styles when the user wants a Beamer-like academic structure in PowerPoint. Borrow Beamer/Madrid structure: frame titles, footline frame numbers, theorem/example/alert blocks, figure/text slides, booktabs-like tables, algorithm frames, diagrams, references, and backup slides. Redraw for PPTX projection scale: slide titles around 20 pt, body text normally 16 pt or larger, fewer simultaneous elements, larger visual regions, and generous spacing. Keep SYSU colors, fonts, logos, and campus imagery; do not output LaTeX or copy another university/institute identity.

Use the `beamer-candidates` styles when the user is selecting a Beamer-derived direction before producing a final course/report template. Each candidate showcase is 16:9 and demonstrates blue, green, and red SYSU variants inside one PPTX. Treat public Beamer themes as aesthetic references only: translate layout grammar to PPTX, keep SYSU fonts/assets/colors, and do not copy another theme's branding or proprietary assets.

For all Beamer-derived PPTX styles, do not mechanically preserve LaTeX Beamer density. A Beamer PDF can tolerate 10 pt text because it is typeset as a document; a PPTX template for projection needs larger text, larger blocks, fewer cards, and more breathing room.

## Template Choice

Default to a style-kit plus template-first editing workflow. The local source templates are 16:9 and include official SYSU, Tang Feng/legacy, and Lingnan reference materials.

Use this starting point unless the user specifies otherwise:

| User intent | Starting template |
|---|---|
| General official SYSU deck, academic talk, report, defense | `templates/source/sysu-official/中山大学幻灯片模板-蓝.pptx` |
| Formal red-themed university communication | `templates/source/sysu-official/中山大学幻灯片模板-红.pptx` |
| Sustainability, biomedical, public health, ecology | `templates/source/sysu-official/中山大学幻灯片模板-绿.pptx` |
| Tang Feng / older SYSU look or when user points to that file | `templates/source/sysu-legacy/中山大学模板-唐峰素材 (33).pptx` |

Ask at most three clarifying questions when required: target audience, talk duration or slide count, and preferred style ID. If those are not blocking, choose `strict-sysu-official-blue` for academic/report decks and state it in the working notes.

## Creation Workflow

1. Create `deck-brief.json` with workflow mode, delivery mode, content readiness, brand fidelity, and style-discovery decision.
2. Create `asset-review.json`; inspect supplied figures and images before finalizing the outline. Use an empty manifest when no assets were provided.
3. For `pptx_redesign` or `pptx_revision`, run `extract_deck_content.py`, review `source-deck-extract.json`, and confirm content before editing.
4. If a new deck or redesign has no requested style ID, generate three real-content style previews and record the user's choice in `style-selection.json`. Otherwise use the requested or preserved style directly.
5. Open only the selected full `style.json` and copy it into the output folder.
6. Write a one-line Chinese Design Read, the selected `delivery_mode`, and the style's three taste dials in `outline.md`.
7. Save a slide-by-slide outline with action titles. For academic content, titles state findings or claims, not only topics.
8. Choose a `layout_pattern_id` for every slide, then choose template slide indices from `template-inventory.md`. Slide indices are zero-based.
9. Copy the selected source or generated PPTX to `working.pptx`. Never edit original templates.
10. Duplicate or delete slides to match `template-mapping.json`, then replace text and images while preserving placeholders and formatting.
11. Use `python-pptx` for straightforward replacement. Use OOXML edits only when needed to preserve masters, notes, comments, or complex formatting.
12. Keep every template placeholder filled meaningfully or intentionally cleared. Do not leave sample text.
13. Run structural extraction or audit when needed; compare slide count, layout use, font usage, colors, delivery mode, and content confirmation against the recorded contracts.
14. Render thumbnails at normal and downscaled review sizes and complete Content Intake, Style Discovery, Visual, Taste, Scientific Figure, and Template Fidelity QA. Record unavailable rendering explicitly.

When strict fidelity matters, follow the ai-lingnan `/pptx` pattern: analyze the template visually, map content to varied source layouts, complete structural slide changes first, then edit text/images and run a visual check. Also read the selected style's `asset_manifest` so font families such as `思源宋体 CN Heavy` and reusable media elements are available during generation. Use the Marp reference only for density and review discipline; this skill's PPTX template-first workflow remains the authority.

## Quality Bar

- Preserve 16:9 dimensions: 13.333 x 7.5 in.
- Keep SYSU theme fonts unless the user asks for another language/font strategy.
- Preserve extracted source fonts from the style spec, including `思源宋体 CN Heavy`, `思源黑体 CN Heavy`, `思源黑体 CN Medium`, `思源宋体 CN Medium`, `黑体`, or `微软雅黑` when those are primary in the selected template.
- Use colors from the selected style spec. Do not invent a new palette for official decks.
- Use extracted template assets before adding new visual material:校徽 variants, wordmarks, source icons, campus photos, building cutouts, decorative bars, and repeated diagram elements.
- Do not use a full-height left vertical stripe as a default decoration. Only keep that motif when it is visibly present in the chosen source slide.
- Prefer template layouts over hand-drawn freeform arrangements.
- Match content shape to layout shape: two items use two-column or comparison layouts; image slides need real images; quotes need real quotes.
- Do not use decorative icons, emoji, gradients, or stock-photo filler in academic or official decks.
- For academic decks, apply the ghost deck test: reading only slide titles should tell the argument.
- Keep source citations on slides when using borrowed figures, claims, or data.
- Record visual QA in `qa-notes.md`: title overflow, text readability, element overlap, image aspect ratio, chart readability, citation proximity, footer consistency, and brand color.
- Use `speaker_led` for live explanation and `reading_first` for asynchronous or detailed review. Split content rather than inventing a middle mode or shrinking below 16 pt.
- When style discovery is required, previews must use real deck content and must not show option labels, style IDs, file paths, or workflow notes inside the slide.
- For scientific figures, simplify journal-density figures for projection: split multi-panel images, enlarge axis labels, direct-label key series, and keep one highlighted result per slide.
- Do not repeat one layout pattern on three consecutive content slides. Main decks with eight or more slides should use at least four pattern families.
- Lock one accent, shape/radius system, line-weight system, and shadow policy per deck.
- Do not default to three equal cards, nested cards, decorative pills, repeated micro-labels, fake dashboards, stock-icon filler, or evidence-free diagrams.

## Local References

The third-party reference skills are stored only inside `.codex/reference-skills/` and are not global installs. Treat them as design/process references, not as project dependencies. Do not copy their generated branding or provenance text into final SYSU decks.
