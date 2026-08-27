#!/usr/bin/env node
/**
 * Rebuild the twelve theme goal.json files through the checked-in Dashi PPT
 * MCP server. The shared briefs source is intentionally kept separate so
 * one copy edit can be projected into every theme sample.
 *
 * Run from the repository root:
 *   node examples/theme-samples/scaffold-all.mjs
 *
 * Set SAMPLE_DATE=YYYYMMDD to make the workflow ids explicit for a rerun.
 */

import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '../..');
const sampleDate = String(process.env.SAMPLE_DATE || '20260827');
const themes = Array.from({ length: 12 }, (_, index) => `theme${String(index + 1).padStart(2, '0')}`);

const messages = [
  { jsonrpc: '2.0', id: 1, method: 'initialize', params: {
    protocolVersion: '2025-06-18', capabilities: {},
    clientInfo: { name: 'dashi-theme-samples', version: '1.0.0' },
  } },
  { jsonrpc: '2.0', method: 'notifications/initialized', params: {} },
];

themes.forEach((theme, index) => {
  messages.push({
    jsonrpc: '2.0',
    id: index + 2,
    method: 'tools/call',
    params: {
      name: 'dashi_ppt_scaffold',
      arguments: {
        title: `Dashi PPT 主题示例｜${theme}`,
        goal: `展示 ${theme} 主题下的 Dashi PPT 内容组织、版式选择与可编辑交付`,
        audience: '产品、技术与业务团队',
        owner: 'Dashi PPT',
        theme,
        pages: 8,
        output_path: `examples/theme-samples/${theme}/goal.json`,
        content_briefs_path: 'examples/theme-samples/briefs.json',
        layout_variants: 1,
        seed: `legacy-theme-sample-${theme}-2026`,
        workflow_run_id: `theme-sample-${theme}-${sampleDate}`,
        chunk_size: 0,
      },
    },
  });
});

const result = execFileSync(process.execPath, ['mcp/server.mjs'], {
  cwd: ROOT,
  input: `${messages.map((message) => JSON.stringify(message)).join('\n')}\n`,
  encoding: 'utf8',
  maxBuffer: 16 * 1024 * 1024,
});

const responses = result.trim().split(/\n+/).filter(Boolean).map((line) => JSON.parse(line));
for (const response of responses) {
  const toolResult = response.result;
  if (toolResult?.isError) {
    const detail = toolResult.content?.[0]?.text || `MCP request ${response.id} failed`;
    throw new Error(detail);
  }
}

// The MCP fill-plan writer records an absolute local path. Keep the checked-in
// sample artifacts cloneable while leaving the runtime behavior unchanged.
for (const theme of themes) {
  const planPath = path.join(HERE, theme, 'goal.fill-plan.json');
  const plan = JSON.parse(readFileSync(planPath, 'utf8'));
  plan.goal = path.relative(ROOT, path.join(HERE, theme, 'goal.json')).replaceAll(path.sep, '/');
  writeFileSync(planPath, `${JSON.stringify(plan, null, 2)}\n`);
}

console.log(`Scaffolded ${themes.length} theme goals through dashi_ppt_scaffold.`);
