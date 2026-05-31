# 项目理解与内容梳理

日期：2026-05-27

## 总体定位

本项目是一个面向 Codex 后续生成 PowerPoint 的本地模板与风格能力库，不是普通课件成品仓库。它的核心目标是把中山大学相关 PPTX/PDF 源模板、提取出的视觉资产、可切换风格规范、确定性生成脚本、展示稿和 README 预览图组织成可复用的生成基础。

当前工作区位于 `E:\Codex_Projects\AI_PPT`，Git 分支为 `main`，检查时工作区干净。

## 内容分层

| 层级 | 目录 | 作用 | Git 边界 |
|---|---|---|---|
| 项目技能层 | `.codex/skills/sysu-ppt-generation/` | Codex 项目内 PPT 生成技能、脚本、生成流程、设计规则、模板 inventory | 跟踪 |
| 参考技能层 | `.codex/reference-skills/` | 第三方 PPT/Beamer/Office 技能克隆，仅作本地参考 | `.gitignore` 排除 |
| 原始模板层 | `templates/source/` | 本地原始 PPTX/PDF 模板，作为抽取和再生成源 | 除 README 外排除 |
| 提取资产层 | `templates/assets/<style-id>/` | 从源 PPTX 提取的校徽、wordmark、校区图片、装饰图、图标、contact sheet、asset manifest | 跟踪，二进制走 Git LFS |
| 风格规范层 | `templates/styles/<style-id>/` | 每个风格的 `style.json` 与 `style.md`，由 `style-index.json` 统一入口 | 跟踪 |
| 生成模板层 | `templates/generated/` | Beamer 启发 PPTX 模板和高对比度 wordmark 变体 | 跟踪，二进制走 Git LFS |
| 展示输出层 | `outputs/style-showcase/` | 各风格和模板元素展示 PPTX，用于选型和 QA | 跟踪，二进制走 Git LFS |
| 文档预览层 | `docs/previews/` | README 使用的 PowerPoint 导出 PNG 预览 | 跟踪，二进制走 Git LFS |

## 当前库存

### 顶层体量

| 目录 | 文件数 | 约占空间 |
|---|---:|---:|
| `templates/` | 438 | 405.01 MB |
| `outputs/` | 13 | 238.86 MB |
| `.codex/` | 714 | 43.44 MB |
| `docs/` | 11 | 4.19 MB |

全仓库当前 `git ls-files` 为 473 个文件。按扩展名看，主要体量来自 `.pptx`、`.png`、`.jpeg/.jpg`、`.wdp` 等二进制资产，已由 `.gitattributes` 配置 Git LFS。

### 本地源模板

`templates/source/` 当前实际存在：

- `sysu-official/中山大学幻灯片模板-蓝.pptx`
- `sysu-official/中山大学幻灯片模板-绿.pptx`
- `sysu-official/中山大学幻灯片模板-红.pptx`
- `sysu-legacy/中山大学模板-唐峰素材 (33).pptx`
- `lingnan-reference/参考slides.pdf`

缺失的官方源模板记录已从当前源模板 README 和模板 inventory 中移除。后续若重新加入对应源 PPTX，应先补回文件，再运行模板提取脚本刷新清单。

### 活跃风格

`templates/styles/style-index.json` 当前登记 12 个活跃风格：

| 分组 | 数量 | 风格 |
|---|---:|---|
| `strict-original` | 3 | `strict-sysu-official-blue`, `strict-sysu-official-green`, `strict-sysu-official-red` |
| `beamer-inspired` | 3 | `beamer-sysu-blue`, `beamer-sysu-green`, `beamer-sysu-red` |
| `beamer-candidates` | 5 | `simpleplus-sysu-clean`, `ustc-thu-sysu-institutional`, `moloch-sysu-minimal`, `sleek-sysu-research`, `river-sysu-atelier` |

`strict-*` 风格用于尽量贴近源模板的正式生成；`beamer-sysu-*` 是已生成的 PPTX 模板，适合学术报告和技术讲解；`beamer-candidates` 用于在正式定稿前比较五种 Beamer 派生方向。

### 提取资产

当前 `templates/assets/` 下有 3 个严格源模板资产包：

| 资产包 | 资产数 | 主要类别 |
|---|---:|---|
| `strict-sysu-official-blue` | 46 | logo/wordmark/emblem、校区/背景图、建筑/装饰 cutout、WDP |
| `strict-sysu-official-green` | 74 | 校区/背景图、logo/wordmark/emblem、小图、SVG、模板重复元素 |
| `strict-sysu-official-red` | 36 | logo/wordmark/emblem、WDP、小图、校区/背景图、装饰 cutout |

这些 asset manifest 是后续生成 PPT 时优先查找校徽、标识、校区照片和装饰元素的依据，不应跳过后直接使用外部素材。

## 生成链路

项目当前推荐流程是：

1. 从 `templates/styles/style-index.json` 选择 `style-id`。
2. 对严格原始模板风格，读取对应源 PPTX inventory、`style.json` 和 `asset-manifest.json`。
3. 先写 `outputs/<deck-slug>/outline.md`，使用 action title 描述每页主张。
4. 写 `template-mapping.json`，把内容页映射到源模板 slide/layout 或生成模板页面。
5. 复制源 PPTX 或生成模板为 `working.pptx`，不直接修改 `templates/source/`。
6. 先做结构编辑，再替换文本和图片。
7. 输出 `final.pptx`，并保留 `style.json`、`replacements.json`、`qa-notes.md` 等可追踪中间文件。
8. 能渲染时做图片/缩略图 QA；不能渲染时至少用 `python-pptx` 重开并检查 slide count、尺寸、字体和占位符覆盖。

核心脚本位于 `.codex/skills/sysu-ppt-generation/scripts/`：

- `extract_pptx_template.py`：扫描源 PPTX，生成模板 inventory。
- `generate_strict_original_showcases.py`：抽取严格源模板资产并生成元素展示稿。
- `generate_beamer_inspired_templates.py`：生成蓝/绿/红三套 Beamer 启发 PPTX 模板与 showcase。
- `generate_beamer_candidate_showcases.py`：生成五套 Beamer 候选 showcase。
- `export_readme_previews.ps1`：通过本机 PowerPoint COM 导出 README 预览图。

## 输出与预览

`outputs/style-showcase/` 当前包含三类展示稿：

- `template-elements/`：严格源模板元素展示，包括蓝、绿、红和总览。
- `beamer-inspired/`：蓝、绿、红三套 Beamer 启发模板展示。
- `beamer-candidates/`：SimplePlus、USTC/THU Institutional、Moloch Minimal、Sleek Research、River/Atelier 五套候选。

`docs/previews/` 中的 PNG 是 README 的视觉索引，适合快速浏览风格，但不能替代打开 PPTX 做布局、字体、溢出和图片质量检查。

## 维护边界

- `templates/source/**` 是本地原始资料层，除 README 外不上传。
- `.codex/reference-skills/` 是本地参考资料层，不作为项目依赖暴露。
- `templates/assets/`、`templates/styles/`、`templates/generated/`、`outputs/style-showcase/` 是可复用派生成果层，可以被 Git 跟踪。
- PPTX/PDF/图片/EMF/WDP 走 Git LFS，避免普通 Git 对二进制文件膨胀。
- 严格中山大学风格生成时，源模板和资产 manifest 是权威；第三方 Beamer/PPT 技能只提供流程和结构参考。

## 当前值得注意的问题

1. 后续真正生成新 deck 时，应在输出目录保存 outline、mapping、替换表和 QA 记录，而不是只给一个最终 PPTX。

## 建议的后续整理动作

1. 为未来生成任务新增一个标准输出模板目录说明，例如 `outputs/<deck-slug>/README.md` 或 `qa-notes.md` 模板。
2. 在生成前优先用 `style-index.json` 和对应 `asset-manifest.json` 做风格选择，不直接从 showcase 文件反向猜规则。
3. 下一轮视觉更新时补充中文真实课件页，并重新导出 README 预览图。
