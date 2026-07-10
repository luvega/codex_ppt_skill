#!/usr/bin/env python
"""Generate a SYSU PPT taste calibration showcase."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[4]
OUT = ROOT / "outputs" / "style-showcase" / "taste-calibration" / "taste-calibration-showcase.pptx"
W, H = 13.333, 7.5
BLUE, CYAN, GREEN, RED = "0B4F6C", "3494BA", "2F6F4E", "8F1D2C"
TEXT, MUTED, LINE, SURFACE = "1D2733", "667085", "D8E5EC", "F3F8FB"
FONT = "微软雅黑"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def rect(slide, x, y, w, h, fill, line=None, radius=False):
    kind = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius and getattr(shape, "adjustments", None):
        shape.adjustments[0] = 0.06
    shape.fill.solid(); shape.fill.fore_color.rgb = rgb(fill)
    if line:
        shape.line.color.rgb = rgb(line); shape.line.width = Pt(0.7)
    else:
        shape.line.fill.background()
    return shape


def text(slide, value, x, y, w, h, size, color=TEXT, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.text_frame.clear(); box.text_frame.word_wrap = True
    p = box.text_frame.paragraphs[0]; p.text = value; p.alignment = align
    p.font.name = FONT; p.font.size = Pt(size); p.font.bold = bold; p.font.color.rgb = rgb(color)
    return box


def base(prs, title, number):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, W, H, "FFFFFF")
    rect(slide, 0, 0, W, 0.78, BLUE)
    text(slide, title, 0.72, 0.2, 9.8, 0.32, 20, "FFFFFF", True)
    text(slide, f"PPT Taste Calibration  |  {number:02d}", 9.9, 0.24, 2.7, 0.2, 9.5, "FFFFFF", False, PP_ALIGN.RIGHT)
    text(slide, "示例内容仅用于版式校准，不作为研究证据。", 0.74, 7.08, 5.5, 0.18, 9.5, MUTED)
    return slide


def profile_slide(prs, title, dial, accent, mode, visual_fraction, notes):
    slide = base(prs, title, len(prs.slides) + 1)
    text(slide, "过程性反馈使概念迁移完成率由 42% 提高至 68%", 0.82, 1.08, 11.2, 0.55, 24, TEXT, True)
    chart_w = 10.7 * visual_fraction
    rect(slide, 0.86, 1.92, chart_w, 3.9, "FFFFFF", LINE)
    text(slide, "概念迁移完成率", 1.18, 2.2, 2.2, 0.24, 13, MUTED)
    rect(slide, 1.22, 4.35, 1.3, 0.72, "AAB4C0")
    rect(slide, 3.05, 3.28, 1.3, 1.79, accent)
    text(slide, "42%", 1.22, 4.0, 1.3, 0.22, 16, MUTED, True, PP_ALIGN.CENTER)
    text(slide, "68%", 3.05, 2.91, 1.3, 0.22, 16, accent, True, PP_ALIGN.CENTER)
    text(slide, "常规教学", 1.12, 5.23, 1.5, 0.22, 11, MUTED, False, PP_ALIGN.CENTER)
    text(slide, "过程反馈", 2.95, 5.23, 1.5, 0.22, 11, MUTED, False, PP_ALIGN.CENTER)
    side_x = 0.86 + chart_w + 0.36
    side_w = 11.55 - side_x
    if side_w > 2.1:
        text(slide, "Interpretation", side_x, 2.02, side_w, 0.26, 13, accent, True)
        text(slide, "突出一个比较；单位、来源与解释贴近图表。", side_x, 2.55, side_w, 0.8, 16, TEXT)
        text(slide, notes, side_x, 3.8, side_w, 0.9, 12.5, MUTED)
    v, d, e = dial
    text(slide, f"{mode}  |  variance {v}  density {d}  energy {e}", 0.9, 6.35, 7.2, 0.28, 13, accent, True)


def main() -> int:
    prs = Presentation(); prs.slide_width = Inches(W); prs.slide_height = Inches(H)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, W, H, "FFFFFF"); rect(slide, 0, 0, 0.18, H, BLUE)
    text(slide, "PPT Taste Layer", 0.9, 1.05, 7.5, 0.65, 32, BLUE, True)
    text(slide, "从合规模板到有判断的学术构图", 0.92, 1.88, 8.6, 0.48, 22, TEXT, True)
    text(slide, "Design Read → three dials → layout pattern → pre-flight", 0.94, 2.66, 8.7, 0.32, 15, MUTED)
    for i, (label, value, color) in enumerate([("Layout variance", "构图变化", CYAN), ("Visual density", "信息密度", GREEN), ("Visual energy", "视觉张力", RED)]):
        y = 3.45 + i * 0.78; rect(slide, 0.96, y, 0.18, 0.42, color)
        text(slide, label, 1.38, y - 0.02, 2.4, 0.24, 14, TEXT, True)
        text(slide, value, 3.78, y - 0.02, 2.2, 0.24, 14, MUTED)
    text(slide, "SYSU source templates remain the authority", 8.25, 5.95, 3.8, 0.28, 12, BLUE, True, PP_ALIGN.RIGHT)

    slide = base(prs, "三旋钮把审美判断变成可复查参数", 2)
    rows = [("Layout variance", "版式轮换、对称与非对称、视觉区域比例", "3 / 4-6 / 7-8"), ("Visual density", "每页文字、图表、组件与证据负载", "3-6 主讲 / 7-8 backup"), ("Visual energy", "对比、强调面积、图像尺度与章节张力", "2-4 正式 / 4-6 教学科研")]
    for i, (a, b, c) in enumerate(rows):
        y = 1.38 + i * 1.45; rect(slide, 0.86, y, 11.45, 1.0, SURFACE, LINE)
        text(slide, a, 1.14, y + 0.18, 2.4, 0.24, 16, BLUE, True)
        text(slide, b, 3.45, y + 0.18, 5.15, 0.38, 15, TEXT)
        text(slide, c, 8.82, y + 0.18, 2.9, 0.3, 13, MUTED, True, PP_ALIGN.RIGHT)
    text(slide, "旋钮不能覆盖源模板身份；它们只控制内容页的构图选择。", 0.9, 6.25, 10.8, 0.32, 16, TEXT, True)

    profile_slide(prs, "低变化：正式、可预测、源模板优先", (3, 5, 3), BLUE, "preserve", 0.64, "适合正式汇报和 strict 风格。")
    profile_slide(prs, "中等变化：科研证据与解释形成清晰节奏", (4, 5, 4), GREEN, "evolve", 0.70, "适合 Beamer 教学与科研报告。")
    profile_slide(prs, "较高变化：用于选型展示而非默认生产", (6, 4, 6), RED, "selection", 0.76, "强调视觉方向，但仍保持投影可读。")

    slide = base(prs, "Anti-slop：先删除无意义容器，再建立证据层级", 6)
    text(slide, "常见 AI 模板页", 0.9, 1.15, 4.8, 0.3, 17, RED, True)
    for i in range(3):
        rect(slide, 0.92 + i * 1.72, 1.78, 1.42, 2.3, "FBF4F4", "ECDADA", True)
        text(slide, f"要点 {i+1}", 1.1 + i * 1.72, 2.08, 1.05, 0.22, 13, RED, True, PP_ALIGN.CENTER)
        text(slide, "三个等宽卡片并不能自动形成信息层级。", 1.08 + i * 1.72, 2.68, 1.1, 0.9, 11, MUTED, False, PP_ALIGN.CENTER)
    text(slide, "修正版", 6.85, 1.15, 4.8, 0.3, 17, GREEN, True)
    rect(slide, 6.86, 1.78, 4.95, 2.3, "FFFFFF", LINE)
    text(slide, "68%", 7.15, 2.08, 1.6, 0.58, 32, GREEN, True)
    text(slide, "过程性反馈组的概念迁移完成率", 8.75, 2.15, 2.65, 0.48, 17, TEXT, True)
    text(slide, "一项结果成为视觉中心，解释和来源围绕证据组织。", 7.16, 3.15, 4.15, 0.42, 14, MUTED)
    text(slide, "卡片只在语义需要分组时使用；留白、对齐和比例本身就是设计。", 0.94, 5.15, 10.9, 0.5, 18, TEXT, True)

    OUT.parent.mkdir(parents=True, exist_ok=True); prs.save(OUT)
    print(f"Wrote {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
