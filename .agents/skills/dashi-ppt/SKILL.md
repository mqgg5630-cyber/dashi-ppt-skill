---
name: dashi-ppt
description: 使用 Dashi PPT 制作可编辑 HTML、PPTX 和 PDF 演示文稿。通过 dashi-ppt MCP 工具完成布局查询、组稿、校验、渲染和导出。
---

# Dashi PPT for Antigravity

本仓库已经包含 Antigravity 所需的 MCP 配置：`.agents/mcp_config.json`。连接后可直接调用 `dashi-ppt` 服务中的工具，不需要手写 shell 命令。

## 标准流程

1. 先明确标题、目标、受众、页数、主题（theme01-theme12）、语言和是否需要图片。
2. 首次使用调用 `prepare_dashi_ppt` 安装运行时依赖。
3. 用 `query_dashi_layouts` 选择候选版式，必要时用 `inspect_dashi_layout` 检查字段容量。
4. 用 `scaffold_dashi_presentation` 生成 workspace 内的 `output/<name>/goal.json`。如有逐页 brief，先写入 JSON 文件并通过 `contentBriefs` 传入。
5. 调用 `validate_dashi_goal`，再调用 `render_dashi_presentation` 输出 `output/<name>/ppt/index.html`。
6. 用户明确要求 PPTX 或 PDF 时调用 `export_dashi_presentation`，`deck` 填包含 `index.html` 的 `ppt/` 目录，`format` 填 `pptx` 或 `pdf`。

## 规则

- 所有输入和输出路径都必须是当前 workspace 内的相对路径。
- 不要复用旧 output 目录中的 goal.json；每次任务创建新目录。
- 默认生成 HTML；只有用户明确要求 PPTX、PowerPoint、可编辑 PPTX、PDF 时才导出对应文件。
- 文案跟随用户语言；不要保留与主题无关的模板默认文案。
- 页面选择遵循原始 `skills/dashi-ppt/SKILL.md` 的主题、容量、媒体和四方案约束。
- 需要图片时先确认素材来源；不要伪造媒体路径。导出 PPTX/PDF 需要本机 Chromium/Chrome/Edge。

## 可用工具

- `prepare_dashi_ppt`
- `query_dashi_layouts`
- `inspect_dashi_layout`
- `scaffold_dashi_presentation`
- `validate_dashi_goal`
- `render_dashi_presentation`
- `export_dashi_presentation`
