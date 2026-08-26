# Dashi PPT × Antigravity demo deck

这是使用仓库内置的 Antigravity Skill 与 Dashi PPT MCP 工作流生成的 10 页可编辑 PPTX 示例，主题为“Dashi PPT × Antigravity”。

## 文件

- `goal.json`：Dashi PPT schema v2 源文件，包含每页的 canonical copy、版式候选和 bespoke composition。
- `dashi-ppt-antigravity-mcp-demo.pptx`：已生成的原生可编辑 PPTX，文字、形状、线条和图表元素均可继续编辑。
- `generate-pptx.mjs`：从 `goal.json` 重新生成演示文件的脚本。

## 重新生成

在仓库根目录执行：

```bash
# 首次使用先安装内置 Runtime 依赖
npm --prefix skills/dashi-ppt/project install

node examples/antigravity-mcp-demo/generate-pptx.mjs \
  examples/antigravity-mcp-demo/goal.json \
  examples/antigravity-mcp-demo/dashi-ppt-antigravity-mcp-demo.pptx
```

如需生成 Dashi PPT 的完整 HTML 编辑器版本，可在 Antigravity 中调用 `dashi_ppt_render`，或执行：

```bash
bash skills/dashi-ppt/scripts/render_goal_deck.sh \
  examples/antigravity-mcp-demo/goal.json \
  output/antigravity-mcp-demo/ppt/index.html
```

PPTX 文件和本目录的源代码会随仓库一起同步；`output/` 目录仍作为本地临时渲染目录被 Git 忽略。
