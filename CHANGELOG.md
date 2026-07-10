# 更新记录

## v0.4 - 2026-07-10

- 新增 Visual Discovery：未指定风格的新建/重设计任务先用真实内容生成 `safe`、`structured`、`exploratory` 三个封面方案。
- 为全部活动风格增加紧凑 `selection_profile`，生成代理先读 style index，选定后再读完整 `style.json`。
- 新增 `deck-brief.json`、`asset-review.json`、`style-selection.json` 契约，以及 `speaker_led` / `reading_first` 两种交付模式。
- 新增 `extract_deck_content.py`，只读提取旧 PPT 的页序、标题、正文、备注、图片、几何、裁剪和源文件 SHA-256。
- 新增 `generate_style_discovery_previews.py` 与 `validate_style_discovery.py`，生成/校验三页 PPTX、1600x900 PNG、contact sheet 和选型记录。
- 扩展 `audit_deck_taste.py --brief`、项目 validator、Content Intake QA 和 Style Discovery QA。
- 新增中文 Visual Discovery showcase 和 README 第一屏预览，完整使用说明覆盖新建、旧稿重设计和旧稿修订流程。
- 固定参考 `zarazhangrui/frontend-slides` 提交 `9906a34d640d2111f724544cbc50f7f130569ae1`；仅做 MIT 许可下的方法改写，不引入 HTML、CSS、动画、浏览器编辑、部署、在线字体或外部模板包。

## v0.3 - 2026-07-10

- 新增 PPT Taste Framework，以 `Design Read`、`layout_variance`、`visual_density` 和 `visual_energy` 显式约束构图判断。
- 新增九个独立 layout pattern，覆盖章节、证据图、两栏讲授、比较表、流程、总结和 references/backup。
- 为所有活动 `style.json` 增加 `taste_profile`、`layout_pattern_ids` 和 `anti_patterns`，保持 strict/Beamer/候选状态边界。
- 新增 Taste QA、anti-slop 排版规则和只读 `audit_deck_taste.py`。
- 新增 `taste-calibration-showcase.pptx` 与 README 预览图，用同一组科研内容展示不同旋钮和修正版式。
- 固定参考 `leonxlnx/taste-skill` 提交 `b17742737e796305d829b3ad39eda3add0d79060`；只做 MIT 许可下的概念性改写，不引入前端运行时依赖。
- 新增 `docs/usage-guide.md` 完整中文使用说明，覆盖首次使用、风格选择、Taste Layer、pattern 映射、输出契约、QA、维护和故障排查。

## v0.2 - 2026-06-03

- 重做 strict 官方蓝/绿/红教师选型总览，新增三色单独 README 预览图。
- 升级 Beamer 蓝/绿/红生成模板和 showcase，补充长中文标题、两栏讲授、科研结果图、表格、流程、总结和参考/backup 页型。
- 重做五套 Beamer 候选展示的基准内容，保留 `style_selection_only` 状态，不作为默认生产模板。
- 新增 PPT design tokens、科研图表规则、视觉 QA rubric，并扩展输出契约中的 `qa-notes.md` 章节要求。
- 更新 `style.json` 和 `style-index.json` 的 `visual_tokens`、`qa_focus`、`sample_page_types`、推荐用途和适用场景字段。
- 重新导出 README 预览图，并通过项目状态、JSON、禁用关键词、PPTX 16:9 尺寸和空白格式校验。

## v0.1 - 2026-05-24

- 新增项目内 Codex PPT 生成技能和风格索引。
- 新增严格原始模板风格展示：官方蓝、官方绿、官方红。
- 新增三套 16:9 Beamer 启发 PPTX 模板：蓝、绿、红。
- 新增五套 Beamer 候选展示：SimplePlus、USTC/THU Institutional、Moloch Minimal、Sleek Research、River/Atelier Inspired。
- 新增由 PowerPoint 导出的 README 预览图。
- 通过 `.gitignore` 保持原始 PPTX/PDF 参考模板仅在本地使用。
