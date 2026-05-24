# Project Structure

This project keeps source templates, generated style kits, and generated decks separate.

```text
F:/AI_PPT/
  templates/
    source/
      sysu-official/        # original official SYSU PPTX templates
      sysu-legacy/          # older/Tang Feng SYSU PPTX
      sysu-medical-ai/      # medical AI course PPTX source
      lingnan-reference/    # PDF reference only; not an active strict style
    assets/
      <style-id>/           # extracted media, contact sheet, asset manifest
    styles/
      style-index.json      # style IDs for switching
      <style-id>/style.json # palette/font/use-case spec
      <style-id>/style.md
    generated/
      beamer-inspired/      # generated PPTX templates inspired by Beamer structure
      beamer-assets/        # generated high-contrast SYSU wordmark/logo variants
  docs/
    previews/               # exported PNG previews used by README.md
  outputs/
    style-showcase/         # generated demo PPTX for each style
      template-elements/    # generated demo PPTX built from extracted template assets
      beamer-inspired/      # Beamer-inspired SYSU blue/green/red demos
      beamer-candidates/    # SimplePlus, USTC/THU, Moloch, Sleek, River candidate demos
  .codex/
    skills/sysu-ppt-generation/
    reference-skills/
```

Current active style IDs:

- `strict-sysu-official-blue`
- `strict-sysu-official-green`
- `strict-sysu-official-red`
- `strict-sysu-medical-ai`
- `beamer-sysu-blue`
- `beamer-sysu-green`
- `beamer-sysu-red`
- `simpleplus-sysu-clean`
- `ustc-thu-sysu-institutional`
- `moloch-sysu-minimal`
- `sleek-sysu-research`
- `river-sysu-atelier`

The earlier generated `sysu-minimal-*`, simplified `sysu-medical-ai`, and `lora-hu-academic` style kits have been removed from the active library. The `strict-*` styles are for requests that must follow the original source templates and are backed by extracted fonts, colors, and media assets. The `beamer-sysu-*` styles are generated PowerPoint templates that adapt Beamer academic slide structure to SYSU assets and palettes, redrawn with PPTX-scale typography and spacing. The `beamer-candidates` styles are 16:9 selection decks that compare five Beamer-derived directions while showing blue, green, and red SYSU variants as full-size slides.

Generated preview decks:

- `outputs/style-showcase/template-elements/template-elements-gallery.pptx`
- `outputs/style-showcase/template-elements/strict-sysu-official-blue-elements-showcase.pptx`
- `outputs/style-showcase/template-elements/strict-sysu-official-green-elements-showcase.pptx`
- `outputs/style-showcase/template-elements/strict-sysu-official-red-elements-showcase.pptx`
- `outputs/style-showcase/template-elements/strict-sysu-medical-ai-elements-showcase.pptx`
- `outputs/style-showcase/beamer-inspired/beamer-sysu-blue-showcase.pptx`
- `outputs/style-showcase/beamer-inspired/beamer-sysu-green-showcase.pptx`
- `outputs/style-showcase/beamer-inspired/beamer-sysu-red-showcase.pptx`
- `outputs/style-showcase/beamer-candidates/simpleplus-sysu-clean-showcase.pptx`
- `outputs/style-showcase/beamer-candidates/ustc-thu-sysu-institutional-showcase.pptx`
- `outputs/style-showcase/beamer-candidates/moloch-sysu-minimal-showcase.pptx`
- `outputs/style-showcase/beamer-candidates/sleek-sysu-research-showcase.pptx`
- `outputs/style-showcase/beamer-candidates/river-sysu-atelier-showcase.pptx`

README preview images:

- `docs/previews/strict-template-gallery.png`
- `docs/previews/beamer-inspired-blue.png`
- `docs/previews/beamer-inspired-green.png`
- `docs/previews/beamer-inspired-red.png`
- `docs/previews/candidate-simpleplus.png`
- `docs/previews/candidate-ustc-thu.png`
- `docs/previews/candidate-moloch.png`
- `docs/previews/candidate-sleek.png`
- `docs/previews/candidate-river.png`

Strict official and medical AI showcase decks are not source PPTX copies. They keep the original source PPT cover as slide 1, then demonstrate extracted fonts, template colors,校徽/wordmark assets, campus imagery, recurring marks, and applied slide patterns. Generic full-height left vertical stripes are not part of the strict style unless the source slide itself contains that element.
