# 中山大学 PPT 生成技能

版本：`0.1`

这是一个面向后续 Codex 使用的项目内 PowerPoint 生成技能库。仓库包含可切换的风格规范、从本地模板提取出的视觉资产、生成模板、展示 PPT 和 README 预览图，方便后续按不同场景选择风格并生成 16:9 演示文稿。

## 效果预览

### 原始模板要素

<img src="docs/previews/strict-template-gallery.png" width="760" alt="原始模板要素预览">

### Beamer 启发模板

这一组是面向 PPTX 重新绘制的学术演示模板：保留 Beamer 的标题栏、页脚、block、表格、算法、图示等结构，但使用更适合投影的字号、留白和中山大学配色与视觉元素。

| 蓝色 | 绿色 | 红色 |
|---|---|---|
| <img src="docs/previews/beamer-inspired-blue.png" width="300" alt="蓝色 Beamer 启发模板预览"> | <img src="docs/previews/beamer-inspired-green.png" width="300" alt="绿色 Beamer 启发模板预览"> | <img src="docs/previews/beamer-inspired-red.png" width="300" alt="红色 Beamer 启发模板预览"> |

### Beamer 候选系列

| SimplePlus 中山大学版 | USTC/THU Institutional 中山大学版 |
|---|---|
| <img src="docs/previews/candidate-simpleplus.png" width="420" alt="SimplePlus 候选模板预览"> | <img src="docs/previews/candidate-ustc-thu.png" width="420" alt="USTC THU 候选模板预览"> |

| Moloch Minimal 中山大学版 | Sleek Research 中山大学版 | River/Atelier Inspired 中山大学版 |
|---|---|---|
| <img src="docs/previews/candidate-moloch.png" width="300" alt="Moloch 候选模板预览"> | <img src="docs/previews/candidate-sleek.png" width="300" alt="Sleek 候选模板预览"> | <img src="docs/previews/candidate-river.png" width="300" alt="River Atelier 候选模板预览"> |

## 仓库内容

- `.codex/skills/`：项目内 Codex 技能、参考说明和确定性生成脚本。
- `templates/styles/`：可切换风格规范和 `style-index.json`。
- `templates/assets/`：从本地模板提取的校徽、标志、校区图片和可复用媒体资产。
- `templates/generated/beamer-inspired/`：生成后的 Beamer 启发 PPTX 模板。
- `templates/generated/beamer-assets/`：用于深浅底色的高对比度标志变体。
- `outputs/style-showcase/`：各风格的展示 PPT。
- `docs/previews/`：由 PowerPoint 导出的 README 预览图。

## 原始模板边界

`templates/source/` 下的原始 PPTX/PDF 参考模板只保留在本地，已通过 `.gitignore` 排除，不上传到 GitHub。

仓库中跟踪的是派生后的风格元数据、提取资产、生成模板、展示 PPT 和预览图。

## 风格选择

完整风格入口以 `templates/styles/style-index.json` 为准。当前包含三类：

- 严格原始模板风格：官方蓝、官方绿、官方红。
- Beamer 启发模板：蓝、绿、红三套可直接复用的 PPTX 模板。
- Beamer 候选方向：SimplePlus、USTC/THU Institutional、Moloch Minimal、Sleek Research、River/Atelier Inspired。

严格原始模板用于需要尽量贴近源模板的场景；Beamer 系列用于学术报告、课程讲授、方法介绍、算法说明和研究展示。Beamer 系列只借鉴公共 Beamer 主题的结构语言，不复制外部院校品牌、图形资产或专有视觉元素。

### 我该选哪个风格

| 场景 | 推荐入口 | 说明 |
|---|---|---|
| 通用课程汇报、学术报告、课题组分享 | `strict-sysu-official-blue` | 默认正式中山大学蓝色模板，适合大多数报告。 |
| 生命科学、公共卫生、生态、医学相关课程 | `strict-sysu-official-green` | 绿色系统更适合生命科学和公共健康语境。 |
| 正式会议、政策汇报、仪式性或偏行政场景 | `strict-sysu-official-red` | 红色系统更正式，但应控制文字密度。 |
| 课程讲授、方法课、结构化学术报告 | `beamer-sysu-blue` / `beamer-sysu-green` / `beamer-sysu-red` | 使用 Beamer 式标题栏、页脚、block、表格和算法页，但已按 PPTX 投影尺度重绘。 |
| 还在比较视觉方向 | `beamer-candidates` 系列 | 只作为选型展示，不作为默认正式生成模板。 |

## 从大纲到 PPT

正式生成 deck 时建议使用固定输出目录，便于复查和复现：

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

生成前先写 `outline.md` 和 `template-mapping.json`，再复制源模板或生成模板为 `working.pptx`。最终交付 `final.pptx`，并在 `qa-notes.md` 记录 slide count、尺寸、资产来源、渲染/缩略图检查结果和已知限制。

## 使用方式

在 Codex 中打开本仓库后，先读取：

```text
templates/styles/style-index.json
.codex/skills/sysu-ppt-generation/references/style-schema.md
.codex/skills/sysu-ppt-generation/references/output-contract.md
```

再按目标场景选择风格。若需要 Beamer 系列 PPTX，可优先从以下目录中的生成模板开始：

```text
templates/generated/beamer-inspired/
```

## 重新生成

在仓库根目录运行：

```powershell
$skill = ".codex\skills\<本项目技能目录>"
python "$skill\scripts\extract_pptx_template.py" --root . --out-dir "$skill\references"
python "$skill\scripts\generate_strict_original_showcases.py"
python "$skill\scripts\generate_beamer_inspired_templates.py"
python "$skill\scripts\generate_beamer_candidate_showcases.py"
python "$skill\scripts\validate_project_state.py"
powershell -NoProfile -ExecutionPolicy Bypass -File "$skill\scripts\export_readme_previews.ps1"
```

预览图导出依赖 Windows 本地 Microsoft PowerPoint COM 自动化。

## Beamer 系列参考来源

这些来源用于版式研究和结构借鉴，最终 PPTX 版本已按中山大学模板资产、中文字体、16:9 投影字号和 PowerPoint 编辑需求重新绘制。

| 方向 | 参考来源 |
|---|---|
| Beamer 基础结构 | [Beamer Themes Manual](https://www.beamer.plus/Themes.html)、[CTAN beamer](https://www.ctan.org/pkg/beamer) |
| 技能流程参考 | [Noi1r/beamer-skill](https://github.com/Noi1r/beamer-skill/tree/main/beamer) |
| SimplePlus | [pm25/SimplePlus-BeamerTheme](https://github.com/pm25/SimplePlus-BeamerTheme)、[Overleaf SimplePlus](https://www.overleaf.com/latex/templates/simpleplus-beamertheme/wfmfjhdcrdfx) |
| USTC/THU Institutional | [USTC Presentation/Beamer Template](https://www.overleaf.com/latex/templates/ustc-presentation-slash-beamer-template/rvpmgprgfhmr)、[ustctug/ustcbeamer](https://github.com/ustctug/ustcbeamer)、[CTAN thubeamer](https://ctan.org/tex-archive/macros/latex/contrib/beamer-contrib/themes/thubeamer) |
| Moloch Minimal | [Moloch 文档](https://moloch.ink/)、[jolars/moloch](https://github.com/jolars/moloch) |
| Sleek Research | [Overleaf Sleek Beamer Theme](https://cs.overleaf.com/latex/templates/sleek-beamer-theme/zzpczkprdbqs) |
| River/Atelier Inspired | [Beamer Atelier](https://beameratelier.com/)、[River demo PDF](https://beameratelier.com/assets/River.pdf) |

## 第三方参考技能

第三方参考仓库没有 vendored 到本仓库。其本地路径和提交版本记录在：

```text
.codex/skills/<本项目技能目录>/references/reference-skill-notes.md
```

## Git LFS

PPTX、PDF、图片、EMF 和 WDP 资产使用 Git LFS 管理，因为生成展示稿和提取媒体均为二进制文件。
