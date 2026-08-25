# Dashi PPT MCP server

This directory contains the workspace MCP server that lets Google Antigravity call Dashi PPT directly.

The server uses the MCP stdio transport and has no extra runtime dependency. It resolves the renderer from the repository checkout, so a fresh clone only needs Node.js 20+; the first render installs the renderer's own dependencies under `skills/dashi-ppt/project/node_modules/`.

## Antigravity

When this repository is opened as a workspace, Antigravity discovers the checked-in files automatically:

- `.agents/skills/dashi-ppt/SKILL.md` — progressive-disclosure skill instructions
- `.agents/mcp_config.json` — workspace MCP server registration

The registered server is equivalent to:

```json
{
  "mcpServers": {
    "dashi-ppt": {
      "command": "node",
      "args": ["mcp/server.mjs"],
      "cwd": "."
    }
  }
}
```

If Antigravity is already open, reload the workspace or use its MCP manager to reload the configuration. The server is local and communicates over stdin/stdout; it does not upload deck content.

## Run or test manually

From the repository root:

```bash
node mcp/server.mjs
```

The process is an MCP stdio process, so it waits for JSON-RPC messages and should not be used as a normal HTTP server. Antigravity starts it automatically from `.agents/mcp_config.json`.

Available tools:

- `dashi_ppt_scaffold`
- `dashi_ppt_layout_query`
- `dashi_ppt_inspect_layout`
- `dashi_ppt_stage_media`
- `dashi_ppt_render`
- `dashi_ppt_preview`
- `dashi_ppt_export`
- `dashi_ppt_validate`

All generated deck paths must stay inside the checkout. Media source paths may be absolute local files so the staging tool can import user assets.
