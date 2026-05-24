#!/usr/bin/env python
"""Generate Beamer-inspired SYSU PPTX templates.

These decks borrow Beamer/Madrid academic slide structure: clear top title bar,
frame numbers, theorem/example blocks, booktabs-like tables, columns, and
diagram-first content. They keep SYSU colors, fonts, logos, and imagery.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[4]
STYLE_ROOT = ROOT / "templates" / "styles"
ASSET_ROOT = ROOT / "templates" / "assets"
OUTPUT_ROOT = ROOT / "outputs" / "style-showcase" / "beamer-inspired"
TEMPLATE_ROOT = ROOT / "templates" / "generated" / "beamer-inspired"
REF_ROOT = ROOT / ".codex" / "reference-skills" / "beamer-skill" / "beamer"

SLIDE_W = 13.333
SLIDE_H = 7.5


STYLES: list[dict[str, Any]] = [
    {
        "id": "beamer-sysu-blue",
        "name": "Beamer SYSU Blue",
        "strict_source_id": "strict-sysu-official-blue",
        "source": "templates/source/sysu-official/中山大学幻灯片模板-蓝.pptx",
        "asset_manifest": "templates/assets/strict-sysu-official-blue/asset-manifest.json",
        "accent": "0B4F6C",
        "secondary": "3494BA",
        "emphasis": "029E73",
        "warning": "DE8F05",
        "bg": "FFFFFF",
        "surface": "F3F8FB",
        "surface2": "E6F1F6",
        "border": "D8E5EC",
        "text": "1D2733",
        "muted": "667085",
        "font": "思源黑体 CN Medium",
        "font_fallback": "Microsoft YaHei",
        "use_case": "Academic reports, seminars, and technical talks using SYSU official blue.",
    },
    {
        "id": "beamer-sysu-green",
        "name": "Beamer SYSU Green",
        "strict_source_id": "strict-sysu-official-green",
        "source": "templates/source/sysu-official/中山大学幻灯片模板-绿.pptx",
        "asset_manifest": "templates/assets/strict-sysu-official-green/asset-manifest.json",
        "accent": "2F6F4E",
        "secondary": "73A45D",
        "emphasis": "029E73",
        "warning": "DE8F05",
        "bg": "FFFFFF",
        "surface": "F4F8F1",
        "surface2": "EAF2E4",
        "border": "DCE8DC",
        "text": "1E2A22",
        "muted": "617067",
        "font": "思源宋体 CN Heavy",
        "font_fallback": "Microsoft YaHei",
        "use_case": "Biomedical, life-science, public-health, and sustainability talks using SYSU official green.",
    },
    {
        "id": "beamer-sysu-red",
        "name": "Beamer SYSU Red",
        "strict_source_id": "strict-sysu-official-red",
        "source": "templates/source/sysu-official/中山大学幻灯片模板-红.pptx",
        "asset_manifest": "templates/assets/strict-sysu-official-red/asset-manifest.json",
        "accent": "8F1D2C",
        "secondary": "C83A3A",
        "emphasis": "0173B2",
        "warning": "DE8F05",
        "bg": "FFFFFF",
        "surface": "FBF4F4",
        "surface2": "F3E5E5",
        "border": "ECDADA",
        "text": "2B1B1D",
        "muted": "735A5D",
        "font": "思源宋体 CN Medium",
        "font_fallback": "Microsoft YaHei",
        "use_case": "Formal academic reports, defenses, and official talks using SYSU red.",
    },
]


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def rgb(hex_value: str) -> RGBColor:
    value = hex_value.strip().lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def add_rect(slide, x, y, w, h, fill, line=None, radius=False):
    shape_type = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(0.8)
    else:
        shape.line.fill.background()
    return shape


def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size,
    color,
    font,
    *,
    bold=False,
    align=None,
    valign=None,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.text_frame.clear()
    box.text_frame.word_wrap = True
    if valign:
        box.text_frame.vertical_anchor = valign
    p = box.text_frame.paragraphs[0]
    if align:
        p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box


def add_picture_contain(slide, image_path: Path, x: float, y: float, w: float, h: float):
    with Image.open(image_path) as image:
        ratio = image.width / max(image.height, 1)
    box_ratio = w / h
    if ratio > box_ratio:
        pw = w
        ph = w / ratio
        px = x
        py = y + (h - ph) / 2
    else:
        ph = h
        pw = h * ratio
        px = x + (w - pw) / 2
        py = y
    return slide.shapes.add_picture(str(image_path), Inches(px), Inches(py), width=Inches(pw), height=Inches(ph))


def add_picture_crop(slide, image_path: Path, x: float, y: float, w: float, h: float):
    pic = slide.shapes.add_picture(str(image_path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    with Image.open(image_path) as image:
        img_ratio = image.width / max(image.height, 1)
    frame_ratio = w / h
    if img_ratio > frame_ratio:
        crop = 1 - frame_ratio / img_ratio
        pic.crop_left = crop / 2
        pic.crop_right = crop / 2
    elif img_ratio < frame_ratio:
        crop = 1 - img_ratio / frame_ratio
        pic.crop_top = crop / 2
        pic.crop_bottom = crop / 2
    return pic


def load_manifest(style: dict[str, Any]) -> dict[str, Any]:
    return json.loads((ROOT / style["asset_manifest"]).read_text(encoding="utf-8"))


def assets_by_category(manifest: dict[str, Any], category: str, limit: int) -> list[dict[str, Any]]:
    allowed = {".png", ".jpg", ".jpeg", ".gif", ".tiff"}
    rows = [
        item
        for item in manifest["assets"]
        if item["category"] == category
        and item.get("width")
        and item.get("height")
        and (ROOT / item["file"]).suffix.lower() in allowed
    ]
    rows.sort(key=lambda item: (len(item.get("slides", [])), item.get("width", 0) * item.get("height", 0)), reverse=True)
    return rows[:limit]


def add_frame_header(slide, style: dict[str, Any], title: str, index: int, section: str = "SYSU Beamer"):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, style["bg"])
    add_rect(slide, 0, 0, SLIDE_W, 0.58, style["accent"])
    add_text(slide, title, 0.48, 0.13, 9.0, 0.25, 14, "FFFFFF", style["font"], bold=True)
    add_text(slide, section, 10.0, 0.16, 2.45, 0.18, 8.5, "FFFFFF", style["font_fallback"], align=PP_ALIGN.RIGHT)
    add_rect(slide, 0, 7.08, SLIDE_W, 0.02, style["border"])
    add_text(slide, "Sun Yat-sen University", 0.48, 7.2, 2.8, 0.16, 7.5, style["muted"], "Arial")
    add_text(slide, f"{index:02d}", 12.08, 7.2, 0.5, 0.16, 7.5, style["muted"], "Arial", align=PP_ALIGN.RIGHT)


def add_logo(slide, logo: Path | None, x: float, y: float, w: float, h: float):
    if logo and logo.exists():
        add_picture_contain(slide, logo, x, y, w, h)


def add_block(slide, style: dict[str, Any], x, y, w, h, label, body, *, kind="normal"):
    header_color = {"normal": style["accent"], "example": style["emphasis"], "alert": style["warning"]}[kind]
    add_rect(slide, x, y, w, h, style["surface"], style["border"], radius=True)
    add_rect(slide, x, y, w, 0.34, header_color)
    add_text(slide, label, x + 0.16, y + 0.08, w - 0.32, 0.16, 8.5, "FFFFFF", style["font"], bold=True)
    add_text(slide, body, x + 0.2, y + 0.52, w - 0.4, h - 0.66, 12, style["text"], style["font_fallback"])


def add_arrow(slide, x1, y1, x2, y2, color):
    line = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    line.line.color.rgb = rgb(color)
    line.line.width = Pt(2.0)
    line.line.end_arrowhead = True
    return line


def slide_cover(prs, style, logo, hero):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, style["bg"])
    add_rect(slide, 0, 0, SLIDE_W, 0.74, style["accent"])
    if hero:
        add_picture_crop(slide, hero, 7.35, 1.0, 5.35, 4.4)
        add_rect(slide, 7.35, 5.32, 5.35, 0.06, style["secondary"])
    add_logo(slide, logo, 0.72, 1.02, 2.1, 0.72)
    add_text(slide, "Beamer-Inspired SYSU Template", 0.72, 2.16, 5.9, 0.64, 26, style["accent"], style["font"], bold=True)
    add_text(slide, "中山大学学术演示版式", 0.75, 2.88, 5.4, 0.38, 17, style["text"], style["font"], bold=True)
    add_text(slide, "10pt discipline · 16:9 academic frames · semantic colors · SYSU assets", 0.76, 3.6, 5.6, 0.28, 11, style["muted"], "Arial")
    add_rect(slide, 0.76, 4.3, 4.9, 0.02, style["border"])
    add_text(slide, style["name"], 0.76, 4.55, 4.4, 0.28, 12, style["secondary"], style["font"], bold=True)
    add_text(slide, "Presenter · Institute · Date", 0.76, 6.78, 4.1, 0.18, 8.5, style["muted"], "Arial")


def slide_agenda(prs, style, logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "Outline", 2, "Structure")
    add_logo(slide, logo, 11.25, 0.72, 1.15, 0.38)
    items = [
        ("01", "Motivation", "Why this topic matters now"),
        ("02", "Formal Setup", "Definitions, notation, assumptions"),
        ("03", "Core Construction", "Algorithm, diagram, theorem"),
        ("04", "Evidence", "Comparison table or experiment"),
        ("05", "Takeaways", "References, Q&A, backup slides"),
    ]
    for i, (num, title, desc) in enumerate(items):
        y = 1.38 + i * 0.84
        add_text(slide, num, 0.92, y, 0.62, 0.23, 13, style["secondary"], "Arial", bold=True)
        add_text(slide, title, 1.72, y - 0.02, 2.6, 0.25, 14, style["accent"], style["font"], bold=True)
        add_text(slide, desc, 4.45, y + 0.02, 5.9, 0.2, 10, style["text"], style["font_fallback"])
        add_rect(slide, 1.72, y + 0.45, 8.6, 0.01, style["border"])


def slide_theorem(prs, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "Definition and Theorem Blocks", 3, "Formalism")
    add_text(slide, "Motivation before formalism: state why the object matters, then introduce notation.", 0.78, 0.92, 9.5, 0.28, 12, style["text"], style["font_fallback"])
    add_block(
        slide,
        style,
        0.86,
        1.46,
        5.55,
        1.9,
        "Definition",
        "A model family F is efficient if evaluation cost grows sublinearly in the full parameter count.",
        kind="normal",
    )
    add_block(
        slide,
        style,
        6.92,
        1.46,
        5.55,
        1.9,
        "Theorem",
        "Under bounded-rank updates, adaptation can preserve the main representation while changing only the task interface.",
        kind="example",
    )
    add_text(slide, "Immediate consequences", 0.94, 4.08, 2.8, 0.28, 15, style["accent"], style["font"], bold=True)
    bullets = ["≤ 2 colored blocks per slide", "one display idea per block", "worked example within two slides"]
    for i, item in enumerate(bullets):
        y = 4.56 + i * 0.48
        add_rect(slide, 1.0, y + 0.08, 0.12, 0.12, style["secondary"])
        add_text(slide, item, 1.24, y, 5.0, 0.22, 11.5, style["text"], style["font_fallback"])
    add_block(slide, style, 7.15, 4.08, 4.8, 1.45, "Example", "Use a tiny worked example here rather than another abstract bullet list.", kind="alert")


def slide_columns(prs, style, hero):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "Columns: Visual First, Text Second", 4, "Layout")
    if hero:
        add_rect(slide, 0.78, 1.0, 5.65, 4.55, style["surface"], style["border"], radius=True)
        add_picture_crop(slide, hero, 0.98, 1.22, 5.25, 3.55)
        add_text(slide, "Source SYSU image or extracted figure", 1.0, 4.92, 4.7, 0.18, 8, style["muted"], "Arial")
    add_text(slide, "Claim", 7.05, 1.05, 2.2, 0.32, 17, style["accent"], style["font"], bold=True)
    add_text(slide, "The slide title states the argument; the figure supplies evidence; bullets only interpret the visual.", 7.05, 1.52, 4.7, 0.65, 14, style["text"], style["font_fallback"])
    for i, item in enumerate(["Top-align columns", "One figure, one takeaway", "No nested columns", "Caption stays near visual"]):
        y = 2.65 + i * 0.5
        add_rect(slide, 7.08, y + 0.09, 0.12, 0.12, style["secondary"])
        add_text(slide, item, 7.32, y, 3.9, 0.22, 11, style["text"], style["font_fallback"])
    add_rect(slide, 7.05, 5.25, 4.6, 0.04, style["secondary"])
    add_text(slide, "Beamer columns translated to PPT: 50/45 split with a fixed gutter.", 7.05, 5.45, 4.8, 0.3, 10, style["muted"], style["font_fallback"])


def slide_table(prs, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "Booktabs-Style Comparison", 5, "Evidence")
    add_text(slide, "Keep tables centered, sparse, and separated from the title bar.", 0.8, 0.9, 7.2, 0.24, 11.5, style["text"], style["font_fallback"])
    x, y = 1.1, 1.55
    col_w = [2.45, 2.35, 2.35, 2.35]
    headers = ["Method", "Cost", "Evidence", "Takeaway"]
    rows = [
        ["Baseline", "O(n)", "reference", "simple"],
        ["Improved", "O(r)", "diagram", "efficient"],
        ["SYSU template", "fixed", "assets", "brand-safe"],
        ["Beamer rhythm", "bounded", "blocks", "readable"],
    ]
    total_w = sum(col_w)
    add_rect(slide, x, y, total_w, 0.04, style["accent"])
    cursor = x
    for text, w in zip(headers, col_w):
        add_text(slide, text, cursor + 0.08, y + 0.16, w - 0.16, 0.2, 9.5, style["accent"], style["font"], bold=True)
        cursor += w
    add_rect(slide, x, y + 0.48, total_w, 0.02, style["border"])
    for r, row in enumerate(rows):
        yy = y + 0.58 + r * 0.55
        cursor = x
        for j, (text, w) in enumerate(zip(row, col_w)):
            color = style["secondary"] if text in {"efficient", "brand-safe"} else style["text"]
            add_text(slide, text, cursor + 0.08, yy + 0.14, w - 0.16, 0.18, 9.8, color, style["font_fallback"], bold=j == 0 or color == style["secondary"])
            cursor += w
        add_rect(slide, x, yy + 0.48, total_w, 0.01, style["border"])
    add_rect(slide, x, y + 0.58 + len(rows) * 0.55, total_w, 0.04, style["accent"])
    add_text(slide, "One highlighted cell or row is enough; avoid rainbow tables.", 1.12, 5.2, 7.5, 0.28, 10.5, style["muted"], style["font_fallback"])


def slide_algorithm(prs, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "Algorithm Frame", 6, "Procedure")
    add_text(slide, "Pseudocode stays under 10 lines and highlights one critical step.", 0.8, 0.9, 7.4, 0.24, 11.5, style["text"], style["font_fallback"])
    add_rect(slide, 0.92, 1.38, 5.9, 4.75, "FFFFFF", style["border"], radius=True)
    add_rect(slide, 0.92, 1.38, 5.9, 0.42, style["surface2"])
    add_text(slide, "Algorithm 1: Template-Aware Slide Build", 1.12, 1.51, 5.2, 0.16, 8.8, style["accent"], style["font"], bold=True)
    lines = [
        "Input: outline O, style S, asset manifest A",
        "1  choose source layout by content shape",
        "2  place SYSU logo, footer, frame number",
        "3  insert one main exhibit",
        "4  highlight only the key result",
        "5  validate overflow and contrast",
        "Output: presentation-ready frame",
    ]
    for i, line in enumerate(lines):
        y = 2.05 + i * 0.42
        color = style["secondary"] if i == 3 else style["text"]
        add_text(slide, line, 1.18, y, 5.1, 0.18, 9.8, color, "Consolas", bold=i == 3)
    add_block(slide, style, 7.35, 1.52, 4.45, 1.45, "Guardrail", "No overlays in PPT templates. Duplicate frames for progressive builds.", kind="alert")
    add_block(slide, style, 7.35, 3.4, 4.45, 1.45, "QA", "Every output should reopen, keep text in bounds, and show a real visual element.", kind="example")


def slide_diagram(prs, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "Diagram Frame", 7, "TikZ-to-PPT")
    labels = ["Question", "Model", "Evidence", "Conclusion"]
    xs = [0.98, 3.8, 6.62, 9.44]
    for i, (x, label) in enumerate(zip(xs, labels)):
        add_rect(slide, x, 2.15, 1.7, 0.85, style["surface"], style["border"], radius=True)
        add_text(slide, label, x + 0.15, 2.43, 1.4, 0.18, 10, style["accent"] if i != 2 else style["secondary"], style["font"], bold=True, align=PP_ALIGN.CENTER)
        if i < len(xs) - 1:
            add_arrow(slide, x + 1.78, 2.57, xs[i + 1] - 0.1, 2.57, style["secondary"])
    add_text(slide, "Diagram rules from Beamer/TikZ", 0.98, 4.05, 3.6, 0.3, 15, style["accent"], style["font"], bold=True)
    rules = ["labels never overlap", "arrow labels need clearance", "solid vs dashed has semantic meaning", "diagram fits below title bar"]
    for i, item in enumerate(rules):
        y = 4.5 + i * 0.42
        add_rect(slide, 1.02, y + 0.08, 0.12, 0.12, style["secondary"])
        add_text(slide, item, 1.25, y, 5.3, 0.2, 10.2, style["text"], style["font_fallback"])


def slide_refs(prs, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "References", 8, "Citations")
    refs = [
        "Noi1r/beamer-skill. Beamer workflow and academic slide quality rules.",
        "SYSU official PowerPoint templates. Source fonts, colors, marks, and imagery.",
        "Local extracted asset manifests. Reusable logos, campus photos, and decorative elements.",
    ]
    for i, ref in enumerate(refs):
        y = 1.35 + i * 0.72
        add_text(slide, f"[{i + 1}]", 0.92, y, 0.42, 0.2, 9.5, style["secondary"], "Arial", bold=True)
        add_text(slide, ref, 1.45, y, 9.6, 0.3, 10.5, style["text"], style["font_fallback"])


def slide_thanks(prs, style, logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, style["accent"])
    add_logo(slide, logo, 0.9, 0.95, 2.0, 0.68)
    add_text(slide, "Thank You", 0.9, 2.58, 4.6, 0.7, 33, "FFFFFF", "Arial", bold=True)
    add_text(slide, "Questions and discussion", 0.94, 3.48, 4.5, 0.32, 15, "FFFFFF", style["font_fallback"])
    add_text(slide, "Sun Yat-sen University", 0.94, 6.78, 3.8, 0.2, 9, "FFFFFF", "Arial")


def slide_backup(prs, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_frame_header(slide, style, "Backup: Extended Details", 10, "Appendix")
    add_text(slide, "Use backup slides for proof details, parameter choices, extended tables, or anticipated questions.", 0.84, 1.1, 8.5, 0.28, 12, style["text"], style["font_fallback"])
    add_block(slide, style, 0.92, 1.75, 3.4, 2.0, "Proof detail", "Place the full derivation here, linked from the main theorem slide.", kind="normal")
    add_block(slide, style, 4.92, 1.75, 3.4, 2.0, "Extended table", "Move dense results here instead of shrinking fonts in the main deck.", kind="example")
    add_block(slide, style, 8.92, 1.75, 3.4, 2.0, "Q&A", "Prepare likely objections and implementation details.", kind="alert")


def make_deck(style: dict[str, Any]) -> tuple[Path, Path]:
    manifest = load_manifest(style)
    logos = assets_by_category(manifest, "logo-wordmark-emblem", 3)
    photos = assets_by_category(manifest, "campus-photo-or-background", 4)
    cutouts = assets_by_category(manifest, "cutout-building-or-decorative", 2)
    logo = ROOT / logos[0]["file"] if logos else None
    hero_item = (photos + cutouts)[0] if (photos + cutouts) else None
    hero = ROOT / hero_item["file"] if hero_item else None

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    slide_cover(prs, style, logo, hero)
    slide_agenda(prs, style, logo)
    slide_theorem(prs, style)
    slide_columns(prs, style, hero)
    slide_table(prs, style)
    slide_algorithm(prs, style)
    slide_diagram(prs, style)
    slide_refs(prs, style)
    slide_thanks(prs, style, logo)
    slide_backup(prs, style)

    TEMPLATE_ROOT.mkdir(parents=True, exist_ok=True)
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    template_path = TEMPLATE_ROOT / f"{style['id']}-template.pptx"
    showcase_path = OUTPUT_ROOT / f"{style['id']}-showcase.pptx"
    prs.save(template_path)
    prs.save(showcase_path)
    return template_path, showcase_path


def write_style(style: dict[str, Any], template_path: Path, showcase_path: Path) -> dict[str, Any]:
    style_dir = STYLE_ROOT / style["id"]
    style_dir.mkdir(parents=True, exist_ok=True)
    spec = {
        "id": style["id"],
        "name": style["name"],
        "group": "beamer-inspired",
        "source": style["source"],
        "asset_manifest": style["asset_manifest"],
        "template_pptx": rel(template_path),
        "demo_pptx": rel(showcase_path),
        "reference": {
            "repo": "https://github.com/Noi1r/beamer-skill/tree/main/beamer",
            "local_path": ".codex/reference-skills/beamer-skill/beamer",
            "commit": "b73d48a",
        },
        "palette": {
            key: style[key]
            for key in ["accent", "secondary", "emphasis", "warning", "bg", "surface", "surface2", "border", "text", "muted"]
        },
        "fonts": {
            "primary": style["font"],
            "fallback": style["font_fallback"],
        },
        "rules": [
            "Use a Beamer/Madrid-like top title bar and frame-number footer.",
            "Keep 16:9 canvas and compact academic density.",
            "Use at most two colored blocks per slide.",
            "Use one substantive element per slide: equation, table, algorithm, diagram, or figure.",
            "Use SYSU extracted logos, campus imagery, and source template colors.",
            "Do not copy LaTeX output mechanically; this is a PPTX adaptation.",
        ],
        "use_case": style["use_case"],
    }
    (style_dir / "style.json").write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [
        f"# {style['name']}",
        "",
        f"- Style ID: `{style['id']}`",
        "- Group: `beamer-inspired`",
        f"- Source SYSU template: `{style['source']}`",
        f"- PPTX template: `{rel(template_path)}`",
        f"- Demo PPTX: `{rel(showcase_path)}`",
        f"- Reference: `Noi1r/beamer-skill` at `{spec['reference']['commit']}`",
        "",
        "## Use",
        "",
        style["use_case"],
        "",
        "This is a PPTX adaptation of Beamer academic layout discipline: top frame title, footline, theorem/example blocks, columns, booktabs-like tables, algorithm frames, diagrams, references, and backup slides.",
        "",
        "## Palette",
        "",
        "| Role | HEX |",
        "|---|---|",
    ]
    for key, value in spec["palette"].items():
        lines.append(f"| {key} | `#{value}` |")
    (style_dir / "style.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {
        "id": style["id"],
        "name": style["name"],
        "group": "beamer-inspired",
        "style_spec": rel(style_dir / "style.json"),
        "template_pptx": rel(template_path),
        "demo_pptx": rel(showcase_path),
        "asset_manifest": style["asset_manifest"],
        "source": style["source"],
    }


def update_style_index(entries: list[dict[str, Any]]) -> None:
    index_path = STYLE_ROOT / "style-index.json"
    payload = json.loads(index_path.read_text(encoding="utf-8")) if index_path.exists() else {"styles": []}
    new_ids = {entry["id"] for entry in entries}
    retained = [entry for entry in payload.get("styles", []) if entry.get("id") not in new_ids]
    payload["styles"] = retained + entries
    index_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    entries = []
    for style in STYLES:
        template_path, showcase_path = make_deck(style)
        entries.append(write_style(style, template_path, showcase_path))
        print(f"Wrote {template_path}")
        print(f"Wrote {showcase_path}")
    update_style_index(entries)
    print(f"Wrote {STYLE_ROOT / 'style-index.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
