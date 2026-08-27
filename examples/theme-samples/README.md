# Dashi PPT × Antigravity：12 个主题事例模板

本目录为仓库中的 `theme01`–`theme12` 各提供一份 8 页中文事例模板。12 份模板使用同一套 `briefs.json` 内容源，便于横向比较主题的色彩、字体、装饰和信息节奏差异。

每个主题目录都包含：

- `goal.json`：由 `dashi_ppt_scaffold` 生成的 Dashi PPT goal scaffold，保留该主题的 8 个版式选择。
- `goal.fill-plan.json`：对应的可填充字段与控件检查计划。
- `<theme>-example.pptx`：原生可编辑 PPTX；文字、形状、线条和图表绘制元素均不是截图，可在 PowerPoint 中继续编辑。

## 主题索引

| 主题 | 视觉系统 | 适合场景 | 示例文件 |
|---|---|---|---|
| `theme01` | 轻拟态风 | 产品介绍、企业汇报、方案说明 | [`theme01/theme01-example.pptx`](./theme01/theme01-example.pptx) |
| `theme02` | 炫光紫绿风 | AI、机器人、科技发布、创新项目 | [`theme02/theme02-example.pptx`](./theme02/theme02-example.pptx) |
| `theme03` | 深浅代码风 | 技术方案、架构、开发者内容 | [`theme03/theme03-example.pptx`](./theme03/theme03-example.pptx) |
| `theme04` | 玻璃糖果风 | 年轻化品牌、消费产品、创意提案 | [`theme04/theme04-example.pptx`](./theme04/theme04-example.pptx) |
| `theme05` | 色谱图表风 | 数据报告、市场分析、KPI 复盘 | [`theme05/theme05-example.pptx`](./theme05/theme05-example.pptx) |
| `theme06` | 深色图谱风 | 战略分析、产业报告、高密度数据 | [`theme06/theme06-example.pptx`](./theme06/theme06-example.pptx) |
| `theme07` | 冷白调研风 | 调研报告、白皮书、竞品与政策分析 | [`theme07/theme07-example.pptx`](./theme07/theme07-example.pptx) |
| `theme08` | 黑金实验风 | 高端发布、品牌提案、实验性概念 | [`theme08/theme08-example.pptx`](./theme08/theme08-example.pptx) |
| `theme09` | 深蓝杂志风 | 品牌故事、人物访谈、深度专题 | [`theme09/theme09-example.pptx`](./theme09/theme09-example.pptx) |
| `theme10` | 金色指数风 | 金融数据、投资报告、年度榜单 | [`theme10/theme10-example.pptx`](./theme10/theme10-example.pptx) |
| `theme11` | 高能增长风 | 增长复盘、商业计划、融资路演 | [`theme11/theme11-example.pptx`](./theme11/theme11-example.pptx) |
| `theme12` | 声波霓虹风 | 音乐娱乐、直播、潮流活动 | [`theme12/theme12-example.pptx`](./theme12/theme12-example.pptx) |

## 批量重新生成

从仓库根目录执行。第一条命令通过本地 MCP server 调用 `dashi_ppt_scaffold`，第二条命令将共享内容投影到 12 份原生 PPTX：

```bash
node examples/theme-samples/scaffold-all.mjs
node examples/theme-samples/generate-all.mjs
node examples/theme-samples/verify-all.mjs
```

`verify-all.mjs` 会逐一执行 goal spec 校验、PPTX ZIP 完整性检查，并确认每份演示稿包含 8 个 slide XML 与 8 个 presentation slide id。

只重新生成部分 PPTX：

```bash
node examples/theme-samples/generate-all.mjs theme03 theme08
```

修改文案时编辑 `briefs.json`，然后重新运行 `generate-all.mjs`。如果同时调整主题版式选择，先运行 `scaffold-all.mjs`。`SAMPLE_DATE=YYYYMMDD` 可用于固定批量 scaffold 的 workflow id：

```bash
SAMPLE_DATE=20260827 node examples/theme-samples/scaffold-all.mjs
```

`generate-pptx.mjs` 优先使用已安装的 PptxGenJS；新 clone 在无法访问 npm 时会自动使用同目录的 `pptxgenjs-lite.cjs`，直接写出 OOXML 原生文本框、形状和线条，因此样例的 PPTX 生成不依赖 Chrome。Dashi PPT 的完整 HTML 编辑器和正式 MCP 导出流程仍按 `.agents/skills/dashi-ppt/SKILL.md` 与 `mcp/server.mjs` 执行。
