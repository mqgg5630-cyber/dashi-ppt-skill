# Antigravity integration

本目录让 Google Antigravity / Antigravity CLI 直接发现本仓库的 Skill 和 MCP Server。

- `.agents/mcp_config.json`：workspace MCP 配置
- `.agents/skills/dashi-ppt/SKILL.md`：Antigravity Skill 入口
- `skills/dashi-ppt/mcp-server.mjs`：stdio MCP 适配器

在 Antigravity 中打开此仓库后，使用 Agent 面板的 MCP 管理入口重新加载配置即可。若使用全局配置，把 `mcpServers.dashi-ppt` 复制到 `~/.gemini/config/mcp_config.json`，并将 `cwd` 改为仓库绝对路径。
