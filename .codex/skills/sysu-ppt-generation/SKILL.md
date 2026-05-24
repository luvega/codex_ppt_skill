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
10. Read `references/design-rules.md` when choosing colors, fonts, slide patterns, or academic presentation structure.
11. Read `references/reference-skill-notes.md` only when you need the rationale from the referenced third-party skills.

## Style Choice

This project keeps source-derived strict style IDs, the earlier Beamer/Madrid adaptation, and a separate Beamer candidate showcase group for style selection. The earlier generated `sysu-minimal-*` showcase styles were removed because they were not faithful enough to the source templates.

| Style ID | Use when |
|---|---|
| `strict-sysu-official-blue` | Strict original official blue template, preserving source masters and source slide geometry |
| `strict-sysu-official-green` | Strict original official green template, preserving source masters and source slide geometry |
| `strict-sysu-official-red` | Strict original official red template, preserving source masters and source slide geometry |
| `strict-sysu-medical-ai` | Strict original medical AI courseware template |
| `beamer-sysu-blue` | Beamer/Madrid-inspired SYSU blue academic template with title bar, footline, blocks, tables, diagrams |
| `beamer-sysu-green` | Beamer/Madrid-inspired SYSU green academic template for biomedical/life-science talks |
| `beamer-sysu-red` | Beamer/Madrid-inspired SYSU red academic template for formal talks |
| `simpleplus-sysu-clean` | SimplePlus-inspired clean academic candidate; white canvas, light rules, restrained SYSU identity |
| `ustc-thu-sysu-institutional` | USTC/THU-inspired formal Chinese university candidate; stronger institutional header and navigation |
| `moloch-sysu-minimal` | Moloch/Metropolis-inspired minimal candidate; light background, large hierarchy, progress bar |
| `sleek-sysu-research` | Sleek-inspired research candidate for algorithms, code, methods, and technical data |
| `river-sysu-atelier` | River/Beamer Atelier-inspired structured teaching candidate with outline, section, and environment blocks |

Each style has a spec under `templates/styles/<style-id>/style.json` and a showcase PPTX under `outputs/style-showcase/`.

Use the `strict-*` styles for normal generation in this project. Treat the source PPTX plus `asset_manifest` as the style authority: preserve source layouts and typography, use extracted校徽/标志物/校区图片/绘图元素, and derive new slides from those assets rather than copying the source file as the final answer. Lingnan PDF is not an active strict style unless an editable PPTX source is added.

For strict showcase/demo decks, keep slide 1 as the original source PPT cover. Do not replace it with a synthetic generated cover, and do not add a generic full-height left vertical stripe unless that exact element exists in the selected source layout. `ppt169_building_effective_agents` may be used only as a content-page rhythm reference: precise cards, hero visual first, concise interpretation second, and clear information hierarchy. Do not copy its dark background, Anthropic palette, icons, or typography. Keep each strict style's source-template background colors, SYSU fonts, extracted SYSU assets, and SYSU official/medical-AI colors. Do not mix the LoRA Hu palette or LoRA-specific page grammar into SYSU strict styles unless the user explicitly asks for that hybrid.

Use `beamer-sysu-*` styles when the user wants a Beamer-like academic structure in PowerPoint. Borrow Beamer/Madrid structure: frame titles, footline frame numbers, theorem/example/alert blocks, figure/text slides, booktabs-like tables, algorithm frames, diagrams, references, and backup slides. Redraw for PPTX projection scale: slide titles around 20 pt, body text normally 16 pt or larger, fewer simultaneous elements, larger visual regions, and generous spacing. Keep SYSU colors, fonts, logos, and campus imagery; do not output LaTeX or copy another university/institute identity.

Use the `beamer-candidates` styles when the user is selecting a Beamer-derived direction before producing a final course/report template. Each candidate showcase is 16:9 and demonstrates blue, green, and red SYSU variants inside one PPTX. Treat public Beamer themes as aesthetic references only: translate layout grammar to PPTX, keep SYSU fonts/assets/colors, and do not copy another theme's branding or proprietary assets.

For all Beamer-derived PPTX styles, do not mechanically preserve LaTeX Beamer density. A Beamer PDF can tolerate 10 pt text because it is typeset as a document; a PPTX template for projection needs larger text, larger blocks, fewer cards, and more breathing room.

## Template Choice

Default to a style-kit plus template-first editing workflow. The local source templates are 16:9 and include official SYSU, Tang Feng/legacy, Lingnan reference, and medical AI teaching materials.

Use this starting point unless the user specifies otherwise:

| User intent | Starting template |
|---|---|
| General official SYSU deck, academic talk, report, defense | `templates/source/sysu-official/中山大学幻灯片模板-蓝.pptx` |
| Warm administrative or event tone | `templates/source/sysu-official/中山大学幻灯片模板-橙.pptx` |
| Formal red-themed university communication | `templates/source/sysu-official/中山大学幻灯片模板-红.pptx` |
| Sustainability, biomedical, public health, ecology | `templates/source/sysu-official/中山大学幻灯片模板-绿.pptx` |
| High-contrast closing, solemn review, monochrome visual tone | `templates/source/sysu-official/中山大学幻灯片模板-黑.pptx` |
| Tang Feng / older SYSU look or when user points to that file | `templates/source/sysu-legacy/中山大学模板-唐峰素材 (33).pptx` |
| Medical AI courseware | `templates/source/sysu-medical-ai/人工智能导论（医学五年制）简单模板.pptx` |

Ask at most three clarifying questions when required: target audience, talk duration or slide count, and preferred style ID. If those are not blocking, choose `strict-sysu-official-blue` for academic/report decks and state it in the working notes.

## Creation Workflow

1. Choose one style ID from `templates/styles/style-index.json`.
2. Save a slide-by-slide outline with action titles. For academic content, titles must state findings or claims, not just topics.
3. Choose template slide indices from `template-inventory.md`. Slide indices are zero-based.
4. Create an output folder under `outputs/<deck-slug>/` with `outline.md`, `style.json`, `template-mapping.json`, intermediate files, and the final `.pptx`.
5. Copy the chosen source PPTX into the output folder. Never edit the original templates.
6. Duplicate or delete slides to match `template-mapping.json`, then replace text and images while preserving placeholders and formatting.
7. Use `python-pptx` for straightforward text/image replacement. Use OOXML edits only when needed to preserve masters, notes, comments, or complex formatting.
8. Keep every template placeholder either filled with meaningful content or intentionally cleared. Do not leave sample text.
9. Run the extraction script on the final PPTX when structural verification is needed, and compare slide count, layout use, font usage, and colors against the template inventory.
10. If LibreOffice/Poppler are available, render thumbnails or PDF pages and visually check for cropped text, overlap, broken images, and off-brand colors. If they are not available, report that visual rendering was not run.

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

## Local References

The third-party reference skills are stored only inside `.codex/reference-skills/` and are not global installs. Treat them as design/process references, not as project dependencies. Do not copy their generated branding or provenance text into final SYSU decks.
