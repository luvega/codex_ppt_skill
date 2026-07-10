# AI_PPT 完整使用说明

适用版本：`v0.3`

本项目是面向中山大学教学、科研报告和正式汇报的项目内 PPT 生成技能库。正式生成入口只有 `.codex/skills/sysu-ppt-generation/`；源模板、风格规范、派生模板、展示稿和质量检查都围绕这一入口组织。

## 1. 使用边界

- 源模板位于 `templates/source/`，只读使用，不直接编辑或上传。
- `templates/styles/style-index.json` 是风格注册表。
- `templates/styles/<style-id>/style.json` 是单个风格的生成契约。
- `templates/assets/<style-id>/` 保存从源模板提取的校徽、字体、颜色和媒体资产。
- `templates/generated/` 保存可直接复用的派生模板。
- `outputs/style-showcase/` 用于选型和视觉检查，不是用户最终课件。
- 正式生成结果写入 `outputs/<deck-slug>/`。

不得恢复已移除的非活动模板，也不得把候选风格自动当作生产模板。

## 2. 环境要求

必需环境：

- Windows PowerShell。
- Python 3.11 或更高版本。
- `python-pptx`、Pillow 等项目脚本依赖。
- 本地存在蓝、绿、红 SYSU 源 PPTX。

可选环境：

- Microsoft PowerPoint：用于 COM 导出 PNG 预览。
- Git LFS：用于克隆和提交 PPTX、PNG 等二进制资产。

快速确认：

```powershell
python --version
python -c "import pptx, PIL; print('PPT dependencies OK')"
git lfs version
```

## 3. 第一次进入项目

从项目根目录开始：

```powershell
Set-Location E:\Codex_Projects\AI_PPT
python .codex\skills\sysu-ppt-generation\scripts\validate_project_state.py
```

看到 `Validation passed: 0 errors` 后再开始生成。

按以下顺序读取规则：

1. `templates/styles/style-index.json`
2. `.codex/skills/sysu-ppt-generation/SKILL.md`
3. `.codex/skills/sysu-ppt-generation/references/style-schema.md`
4. `.codex/skills/sysu-ppt-generation/references/output-contract.md`
5. `.codex/skills/sysu-ppt-generation/references/ppt-taste-framework.md`
6. `.codex/skills/sysu-ppt-generation/references/layout-patterns/README.md`
7. `.codex/skills/sysu-ppt-generation/references/visual-qa-rubric.md`

## 4. 选择风格

| 场景 | 推荐风格 | 状态 |
|---|---|---|
| 通用课程、科研汇报、答辩 | `strict-sysu-official-blue` | `ready` |
| 生命科学、公共卫生、生态 | `strict-sysu-official-green` | `ready` |
| 正式会议、政策或行政汇报 | `strict-sysu-official-red` | `ready` |
| 结构化课程、方法介绍、技术报告 | `beamer-sysu-blue` | `ready` |
| 生物医学科研/教学报告 | `beamer-sysu-green` | `ready` |
| 正式结构化报告 | `beamer-sysu-red` | `ready` |
| 比较视觉方向 | 五个 Beamer candidate | `style_selection_only` |

选择原则：

- 用户指定 style ID 时优先使用指定风格。
- 未指定时，通用正式报告默认选择 `strict-sysu-official-blue`。
- strict 风格以源模板为权威，不随意重绘母版。
- Beamer 风格以生成模板为起点，允许在 SYSU 身份内改善内容页构图。
- candidate 只能用于选型；正式生成前必须明确提升为某个生产方向。

## 5. PPT Taste Layer

### 5.1 Design Read

生成前先写一句中文设计判断：

```text
将本任务理解为：面向课题组教师与研究生的 15 分钟科研报告，在普通教室投影，以结果图为主，保持正式、克制并严格使用中山大学蓝色系统。
```

这句话至少包含：

- 受众。
- 报告类型和时长。
- 投影或使用场景。
- 主要内容形态。
- 正式程度。
- 品牌保真要求。

### 5.2 三个旋钮

| 字段 | 控制内容 |
|---|---|
| `layout_variance` | 版式轮换、对称/非对称和视觉区域比例 |
| `visual_density` | 每页文字、图表、组件和证据负载 |
| `visual_energy` | 静态对比、强调面积、图片尺度和章节张力 |

不要把 `visual_energy` 理解为动画。正式主讲页的 `visual_density` 通常保持在 3-6；7-8 只用于方法或 backup；9-10 必须拆页。

### 5.3 Taste mode

- `preserve`：strict 风格。保持源模板母版、字体、配色和构图身份。
- `evolve`：Beamer 风格。在 SYSU 身份内改善节奏和层级。
- `selection`：候选展示，只用于比较视觉方向。

## 6. 从大纲到 PPT

### 6.1 准备输入

建议提供：

- 报告题目、受众、时长和预计页数。
- 中文或英文大纲。
- 原始图表、图片、表格和引用来源。
- 指定 style ID；未指定时由生成代理选择。
- 必须保留的校徽、课题组、学院或会议信息。

不得用装饰图替代缺失的研究证据。原始图表不可用时，在 QA 中明确记录限制。

### 6.2 创建输出目录

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

`deck-slug` 使用简短 ASCII 名称，例如 `crc-lab-meeting-202607`。

### 6.3 编写 outline.md

```markdown
# Outline

## Design Read

将本任务理解为：面向研究生课程的 20 分钟方法讲授，在普通教室投影，以流程和结果图为主，使用克制的 SYSU 蓝色 Beamer 结构。

- Style: `beamer-sysu-blue`
- Mode: `evolve`
- Layout variance: `4`
- Visual density: `5`
- Visual energy: `4`

| New Slide | Action Title | Content Role | Layout Pattern ID | Exhibit/Image | Template File | Template Slide/Layout |
|---:|---|---|---|---|---|---|
| 1 | 过程性反馈帮助学生更快完成概念迁移 | cover | `action-title-hero-visual` | 校园图/主题图 | beamer template | cover |
| 2 | 三类证据共同回答教学效果问题 | structure | `process-or-causal-flow` | 流程图 | beamer template | flow |
| 3 | 反馈组完成率由 42% 提高至 68% | result | `full-figure-evidence` | 结果图 | beamer template | figure |
| 4 | 提升主要发生在第二次课堂活动后 | interpretation | `figure-plus-interpretation` | 结果图 | beamer template | figure+text |
| 5 | 课程改进应聚焦三个可执行动作 | summary | `claim-summary` | 无 | beamer template | summary |
```

只读标题应能说明完整论证，这就是 ghost deck 检查。

### 6.4 编写 template-mapping.json

```json
{
  "style_id": "beamer-sysu-blue",
  "slides": [
    {
      "new_slide": 1,
      "content_role": "cover",
      "layout_pattern_id": "action-title-hero-visual",
      "template_slide": 1
    },
    {
      "new_slide": 2,
      "content_role": "structure",
      "layout_pattern_id": "process-or-causal-flow",
      "template_slide": 7
    },
    {
      "new_slide": 3,
      "content_role": "result",
      "layout_pattern_id": "full-figure-evidence",
      "template_slide": 5
    }
  ]
}
```

同一 pattern 不得连续出现三页。8 页以上主报告至少使用四类 pattern。

### 6.5 复制 style 和模板

将选定的 `templates/styles/<style-id>/style.json` 复制到输出目录。然后：

- strict：复制对应源 PPTX 为 `working.pptx`。
- Beamer：复制 `templates/generated/beamer-inspired/<style-id>-template.pptx` 为 `working.pptx`。

永远不要原地修改 `templates/source/`。

### 6.6 完成替换

`replacements.json` 应记录每个替换对象：

```json
{
  "slides": [
    {
      "slide": 3,
      "replacements": [
        {"shape": "title", "type": "text", "value": "反馈组完成率由 42% 提高至 68%"},
        {"shape": "result_figure", "type": "image", "source": "inputs/result-figure.png"}
      ]
    }
  ]
}
```

优先替换模板占位符；只有模板无法承载时才绘制自由形状。自由形状必须遵守选定 pattern 和 style tokens。

## 7. Layout Pattern 选择

| 内容 | Pattern |
|---|---|
| 章节过渡 | `section-divider` |
| 一个结论和一个大视觉 | `action-title-hero-visual` |
| 结果图本身就是核心证据 | `full-figure-evidence` |
| 图表需要简短解释 | `figure-plus-interpretation` |
| 概念/案例、问题/解释 | `two-column-teaching` |
| 方法、方案或证据比较 | `comparison-table` |
| 流程、机制、因果链 | `process-or-causal-flow` |
| 结论、建议和教学总结 | `claim-summary` |
| 文献、方法细节和补充结果 | `references-backup` |

详细内容预算和失败模式位于 `.codex/skills/sysu-ppt-generation/references/layout-patterns/`。

## 8. 视觉与科研图表要求

- 正文通常不小于 16 pt，教学主文本优先 18-22 pt。
- 每页只突出一个结果或比较。
- 图表轴标签通常不小于 18 pt，刻度通常不小于 16 pt。
- 四条以内数据系列优先直接标注，避免复杂 legend。
- 单位、样本量、统计标记和来源必须留在可读区域。
- 多面板论文图优先拆页，不整体缩小。
- 图片保持比例；任何裁剪、对比度处理或标注都写入 QA。
- 卡片只用于真实语义分组，不因“有三个要点”自动画三个卡片。
- 一个 deck 锁定强调色、圆角、线宽和阴影策略。

## 9. QA 与交付

`qa-notes.md` 使用固定结构：

```markdown
# QA Notes

## Structural QA

## Visual QA

## Taste QA

## Scientific Figure QA

## Template Fidelity

## Known Limitations
```

Taste QA 至少检查：

- Design Read 和旋钮是否与任务一致。
- 层级、节奏、克制、构图和证据清晰度。
- 是否连续三页重复同一 pattern。
- 是否出现无意义等宽卡片、嵌套卡片、胶囊标签或假仪表盘。
- 强调色、圆角、线宽和阴影是否统一。
- strict 模式是否保留源模板身份。

运行只读审计：

```powershell
python .codex\skills\sysu-ppt-generation\scripts\audit_deck_taste.py `
  outputs\<deck-slug>\final.pptx `
  --style outputs\<deck-slug>\style.json `
  --mapping outputs\<deck-slug>\template-mapping.json
```

脚本只报告问题，不自动修改 PPTX。`error` 必须修复；`warning` 必须在 `qa-notes.md` 中解释。

## 10. 维护与重新生成

完整刷新顺序：

```powershell
$skill = ".codex\skills\sysu-ppt-generation"

python "$skill\scripts\extract_pptx_template.py" --root . --out-dir "$skill\references"
python "$skill\scripts\generate_strict_original_showcases.py"
python "$skill\scripts\generate_beamer_inspired_templates.py"
python "$skill\scripts\generate_beamer_candidate_showcases.py"
python "$skill\scripts\generate_taste_calibration_showcase.py"
python "$skill\scripts\validate_project_state.py"
powershell -NoProfile -ExecutionPolicy Bypass -File "$skill\scripts\export_readme_previews.ps1"
```

顺序不能随意颠倒：生成器会更新 style metadata、style-index 和模板清单，预览必须在 PPTX 更新后重新导出。

### 新增或修改风格

1. 确认源 PPTX 位于 `templates/source/`。
2. 提取资产和模板清单。
3. 创建或更新生成器中的风格定义。
4. 确保 `style.json` 含全部必需字段和合法 `taste_profile`。
5. 生成 showcase 和预览图。
6. 运行 validator、JSON 校验、Taste audit 和人工截图检查。

不要直接手工修改派生 `style.json` 后跳过生成器，否则下一次刷新会丢失改动。

## 11. 常见问题

### Validator 报 style 字段缺失

重新运行对应生成脚本，不要只修改 style-index。所有活动 style 必须有 `taste_profile`、`layout_pattern_ids` 和非空 `anti_patterns`。

### 正文低于 16 pt

优先拆页、缩短文本或改用图表/流程。不要直接缩小字号。references 和 footer 可以更小，但必须可读并在 QA 中说明。

### strict 模板出现越界 warning

先确认是源模板已有装饰或出血元素。`preserve` 模式可以保留并记录 warning；新绘制元素越界仍必须修复。

### PowerPoint 预览导出失败

确认本机安装桌面版 PowerPoint，关闭占用相关 PPTX 的窗口，再重新运行 `export_readme_previews.ps1`。

### 中文字体显示异常

确认 style spec 中列出的思源字体或微软雅黑已安装。不要在同一 deck 中随意替换多套字体。

### Git 中 PPTX/PNG 显示为指针文件

运行：

```powershell
git lfs install
git lfs pull
```

## 12. 外部参考边界

PPT Taste Layer 概念性参考 `leonxlnx/taste-skill` v2，固定提交为 `b17742737e796305d829b3ad39eda3add0d79060`。本项目只提取 brief inference、三旋钮、anti-default、redesign mode、pattern library 和 pre-flight 方法，不安装该 skill，不引入 React、CSS、动效、暗色模式或在线服务依赖。

SYSU 源模板、style registry、style spec 和本地生成脚本始终拥有更高优先级。
