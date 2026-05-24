# SYSU Codex PPT Skill

Version: `0.1`

Project-local Codex skill and template library for generating Sun Yat-sen University PowerPoint decks. The repository keeps reusable style specs, extracted SYSU visual assets, generated PPTX templates, and showcase decks together so future Codex runs can switch styles without starting from a blank deck.

## Preview

### Source-Derived SYSU Elements

<img src="docs/previews/strict-template-gallery.png" width="760" alt="SYSU strict template element gallery">

### Beamer Inspired SYSU Templates

These are PPTX-first academic templates: Beamer-like structure, SYSU colors and assets, larger projection-safe typography, and no LaTeX Beamer density.

| Blue | Green | Red |
|---|---|---|
| <img src="docs/previews/beamer-inspired-blue.png" width="300" alt="Beamer SYSU blue preview"> | <img src="docs/previews/beamer-inspired-green.png" width="300" alt="Beamer SYSU green preview"> | <img src="docs/previews/beamer-inspired-red.png" width="300" alt="Beamer SYSU red preview"> |

### Beamer Candidate Families

| SimplePlus-SYSU | USTC/THU-SYSU |
|---|---|
| <img src="docs/previews/candidate-simpleplus.png" width="420" alt="SimplePlus SYSU candidate preview"> | <img src="docs/previews/candidate-ustc-thu.png" width="420" alt="USTC THU SYSU candidate preview"> |

| Moloch-SYSU | Sleek-SYSU | River/Atelier-SYSU |
|---|---|---|
| <img src="docs/previews/candidate-moloch.png" width="300" alt="Moloch SYSU candidate preview"> | <img src="docs/previews/candidate-sleek.png" width="300" alt="Sleek SYSU candidate preview"> | <img src="docs/previews/candidate-river.png" width="300" alt="River Beamer Atelier SYSU candidate preview"> |

## What Is Included

- `.codex/skills/sysu-ppt-generation/`: the project-local Codex skill, references, and deterministic generation scripts.
- `templates/styles/`: switchable style specs and `style-index.json`.
- `templates/assets/`: extracted SYSU logos, wordmarks, campus imagery, and reusable media.
- `templates/generated/beamer-inspired/`: generated Beamer-inspired PPTX templates.
- `templates/generated/beamer-assets/`: generated high-contrast SYSU wordmark variants for light and dark headers.
- `outputs/style-showcase/`: generated showcase PPTX files for review.
- `docs/previews/`: PowerPoint-exported PNG previews used in this README.

## Source Template Boundary

Original reference PPTX/PDF files under `templates/source/` are local-only and intentionally ignored by Git. They are required for regenerating source-derived inventories and strict showcases, but are not uploaded to GitHub.

Tracked files include derived style metadata, extracted reusable assets, generated templates, generated showcase PPTX files, and README preview images.

## Active Style IDs

Strict source-derived styles:

- `strict-sysu-official-blue`
- `strict-sysu-official-green`
- `strict-sysu-official-red`
- `strict-sysu-medical-ai`

Generated Beamer-inspired PPTX templates:

- `beamer-sysu-blue`
- `beamer-sysu-green`
- `beamer-sysu-red`

Beamer-derived candidate directions:

- `simpleplus-sysu-clean`
- `ustc-thu-sysu-institutional`
- `moloch-sysu-minimal`
- `sleek-sysu-research`
- `river-sysu-atelier`

The earlier `sysu-minimal-*` style kits were removed. Strict styles preserve source-template identity. Beamer styles use public Beamer themes only as layout references and are redrawn for 16:9 PPTX projection: larger typography, larger logos, high-contrast SYSU wordmarks, fewer simultaneous elements, and more spacing.

## Usage

Open this repository in Codex and ask for a SYSU PowerPoint deck using one of the style IDs above. The project-local skill lives at:

```text
.codex/skills/sysu-ppt-generation/SKILL.md
```

For generation work, use the style index first:

```text
templates/styles/style-index.json
```

For Beamer-style PPTX output, start from the generated templates under:

```text
templates/generated/beamer-inspired/
```

## Regeneration

From the repository root:

```powershell
python .codex\skills\sysu-ppt-generation\scripts\extract_pptx_template.py --root . --out-dir .codex\skills\sysu-ppt-generation\references
python .codex\skills\sysu-ppt-generation\scripts\generate_strict_original_showcases.py
python .codex\skills\sysu-ppt-generation\scripts\generate_beamer_inspired_templates.py
python .codex\skills\sysu-ppt-generation\scripts\generate_beamer_candidate_showcases.py
powershell -NoProfile -ExecutionPolicy Bypass -File .codex\skills\sysu-ppt-generation\scripts\export_readme_previews.ps1
```

Preview export uses local Microsoft PowerPoint COM automation on Windows.

## Reference Skills

Third-party reference repositories are not vendored into this repo. Their local clone paths and commit IDs are recorded in:

```text
.codex/skills/sysu-ppt-generation/references/reference-skill-notes.md
```

## Git LFS

PPTX, PDF, image, EMF, and WDP assets are tracked with Git LFS because generated decks and extracted media are binary files.
