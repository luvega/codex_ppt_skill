# SYSU Codex PPT Skill

Project-local Codex skill and template library for generating Sun Yat-sen University PowerPoint decks.

## What Is Included

- `.codex/skills/sysu-ppt-generation/`: the Codex skill, references, and generation scripts.
- `templates/source/`: local SYSU source PPTX/PDF templates.
- `templates/assets/`: extracted logos, marks, campus imagery, and reusable media from the source templates.
- `templates/styles/`: style specs and `style-index.json`.
- `templates/generated/beamer-inspired/`: generated Beamer-inspired SYSU PPTX templates.
- `outputs/style-showcase/`: generated showcase PPTX files for style review.

## Active Style IDs

- `strict-sysu-official-blue`
- `strict-sysu-official-green`
- `strict-sysu-official-red`
- `strict-sysu-medical-ai`
- `beamer-sysu-blue`
- `beamer-sysu-green`
- `beamer-sysu-red`

The earlier `sysu-minimal-*` style kits were removed. Strict styles preserve source-template identity. Beamer-inspired styles adapt Beamer/Madrid academic frame structure to SYSU colors, fonts, logos, and campus imagery.

## Regeneration

From the repository root:

```powershell
python .codex\skills\sysu-ppt-generation\scripts\extract_pptx_template.py --root . --out-dir .codex\skills\sysu-ppt-generation\references
python .codex\skills\sysu-ppt-generation\scripts\generate_strict_original_showcases.py
python .codex\skills\sysu-ppt-generation\scripts\generate_beamer_inspired_templates.py
```

## Reference Skills

Third-party reference repositories are not vendored into this repo. Their local clone paths and commit IDs are recorded in `.codex/skills/sysu-ppt-generation/references/reference-skill-notes.md`.

## Git LFS

PPTX, PDF, image, EMF, and WDP assets are tracked with Git LFS because several source templates and generated showcase decks are large binary files.

