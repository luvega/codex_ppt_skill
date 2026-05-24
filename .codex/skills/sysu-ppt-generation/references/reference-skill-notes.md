# Reference Skill Notes

These repositories are stored under `.codex/reference-skills/` for project-local reference only.

| Reference | Local path | Commit | Reusable idea |
|---|---|---:|---|
| hugohe3/ppt-master | `.codex/reference-skills/ppt-master` | f477b63 | Use a serial pipeline, explicit design spec, locked visual parameters, per-page QA, and a final export gate. Useful when building a more ambitious SVG-to-PPTX pipeline. |
| op7418/guizang-ppt-skill | `.codex/reference-skills/guizang-ppt-skill` | 6bfa520 | Clarify style/audience/assets first, use a template as the class/source of truth, preflight class/layout availability, and require visual checking. |
| Gabberflast/academic-pptx-skill | `.codex/reference-skills/academic-pptx-skill` | f77d3a0 | For academic decks, privilege argument structure, action titles, one insight per slide, citations, and the ghost deck test. |
| tfriedel/claude-office-skills | `.codex/reference-skills/claude-office-skills` | d4241e4 | For PPTX work, inspect OOXML/theme/layouts, build a zero-based template inventory, map outline slides to template slides, replace text with validation, and visually QA thumbnails when possible. |
| ai-lingnan/ai-tools-installation recommended_skills | `.codex/reference-skills/ai-tools-installation/recommended_skills` | f76ba65 | For `/pptx`, analyze templates visually, use varied source layouts, do structural edits before content replacement, and run visual QA. For `marp-slides-creator`, keep one point per slide, strong hierarchy, limited density, staged reviews, and PNG checks. For `marp-export`, prefer explicit export commands and report missing rendering dependencies. |
| hugohe3/ppt-master building_effective_agents | `.codex/reference-skills/ppt-master/examples/ppt169_building_effective_agents` | f477b63 | Borrow only the aesthetic discipline: precise cards, hero diagram first, concise interpretation second, and breathing/dense page rhythm. Do not import its dark background, Anthropic colors, Helvetica typography, chunk icons, or generic vertical accents into SYSU decks. |
| Noi1r/beamer-skill | `.codex/reference-skills/beamer-skill/beamer` | b73d48a | Borrow Beamer/Madrid academic structure: 16:9, compact frame title bar, footline frame numbers, theorem/example blocks, two-column layouts, booktabs tables, algorithm frames, references, backup slides, density and overflow discipline. Do not emit LaTeX unless requested. |

## Local Adaptation

Use the SYSU skill as the controlling workflow. The reference skills inform process decisions but must not override the local template-first requirement.

Do not include third-party provenance text, repository branding, or hidden workflow notes in generated PPTX files. Mention references only in internal `qa-notes.md` when useful.

Use `ppt-master` style `design_spec.md` / `spec_lock.md` only if the user asks for a high-design custom pipeline. For ordinary SYSU decks, a concise `outline.md`, `template-mapping.json`, and `qa-notes.md` are enough.

Use ai-lingnan `/pptx` ideas for strict original PPTX work: inspect first, map to varied layouts, duplicate/delete/reorder slides before text replacement, and check for overflow or leftover placeholders. Use its Marp-related guidance only as review discipline; do not convert this SYSU PPTX workflow into a Marp-first workflow unless the user asks for Marp output.

Use `ppt169_building_effective_agents` as the technical-deck rhythm reference for content pages only. The SYSU source PPTX remains the cover, background, and brand authority; use SYSU fonts, extracted SYSU assets, and SYSU template colors. Do not carry over the LoRA Hu blue/orange paper palette, the effective-agents dark background, or the effective-agents Anthropic coral palette unless the user explicitly asks for those styles.

Use `Noi1r/beamer-skill` as a structural reference for Beamer-like academic PPTX templates only. Translate its LaTeX concepts into PowerPoint components: frame title bars, footlines, block styles, compact columns, centered tables, diagram spacing rules, references, and backup slides. Replace its default institute text and colors with SYSU assets and selected SYSU official blue/green/red palettes.
