#!/usr/bin/env node
/**
 * Dashi PPT MCP adapter for Antigravity (stdio transport).
 * No third-party MCP SDK is required; the server speaks the MCP JSON-RPC
 * line protocol so a fresh clone can be connected immediately.
 */
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync } from 'node:fs';
import path from 'node:path';
import readline from 'node:readline';

const SKILL_ROOT = path.resolve(import.meta.dirname);
const PROJECT_ROOT = path.join(SKILL_ROOT, 'project');
const WORKSPACE_ROOT = path.resolve(process.env.DASHI_PPT_WORKSPACE || process.cwd());
const SERVER_INFO = { name: 'dashi-ppt', version: '0.4.11' };

const tools = [
  {
    name: 'prepare_dashi_ppt',
    description: 'Install the Dashi PPT runtime dependencies in this checkout. Run once before rendering or exporting.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'query_dashi_layouts',
    description: 'Find suitable Dashi PPT layouts for a theme and content shape.',
    inputSchema: { type: 'object', properties: { theme: { type: 'string', description: 'theme01 through theme12' }, limit: { type: 'integer', default: 12 }, itemCount: { type: 'integer', default: 4 }, priority: { type: 'string', default: 'balanced' }, needsMedia: { type: 'boolean', default: false } }, required: ['theme'] },
  },
  {
    name: 'inspect_dashi_layout',
    description: 'Inspect the editable fields and capacity contract of one or more Dashi PPT layouts.',
    inputSchema: { type: 'object', properties: { layouts: { type: 'array', items: { type: 'string' } }, layout: { type: 'string' } } },
  },
  {
    name: 'scaffold_dashi_presentation',
    description: 'Create a schema v2 goal.json with three template variants and one bespoke variant per slide.',
    inputSchema: { type: 'object', properties: { title: { type: 'string' }, goal: { type: 'string' }, theme: { type: 'string' }, pages: { type: 'integer', default: 10 }, output: { type: 'string', description: 'Workspace-relative path ending in goal.json' }, contentBriefs: { type: 'string', description: 'Workspace-relative JSON file containing page briefs' }, language: { type: 'string', enum: ['zh', 'en'] } }, required: ['title', 'goal', 'theme', 'output'] },
  },
  {
    name: 'validate_dashi_goal',
    description: 'Validate a Dashi PPT goal.json before rendering.',
    inputSchema: { type: 'object', properties: { goal: { type: 'string' } }, required: ['goal'] },
  },
  {
    name: 'render_dashi_presentation',
    description: 'Render a goal.json into an editable HTML deck and start no server; use the returned index path with a local preview.',
    inputSchema: { type: 'object', properties: { goal: { type: 'string' }, output: { type: 'string', description: 'Workspace-relative path ending in ppt/index.html' } }, required: ['goal', 'output'] },
  },
  {
    name: 'export_dashi_presentation',
    description: 'Export a rendered Dashi PPT deck directory (the folder containing index.html) to editable PPTX or PDF.',
    inputSchema: { type: 'object', properties: { deck: { type: 'string', description: 'Workspace-relative ppt directory or index.html' }, output: { type: 'string' }, format: { type: 'string', enum: ['pptx', 'pdf'], default: 'pptx' } }, required: ['deck', 'output'] },
  },
];

function reply(id, result) { process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id, result }) + '\n'); }
function fail(id, message, code = -32000) { reply(id, { isError: true, content: [{ type: 'text', text: message }], _jsonrpcError: { code, message } }); }
function textResult(text, isError = false) { return { isError, content: [{ type: 'text', text: String(text) }] }; }
function workspacePath(value) {
  const resolved = path.resolve(WORKSPACE_ROOT, String(value));
  if (resolved !== WORKSPACE_ROOT && !resolved.startsWith(WORKSPACE_ROOT + path.sep)) throw new Error(`Path must stay inside workspace: ${value}`);
  return resolved;
}
function arg(value, fallback) { return value == null || value === '' ? fallback : String(value); }
function run(command, args, { cwd = PROJECT_ROOT, env = {} } = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, args, { cwd, env: { ...process.env, INIT_CWD: WORKSPACE_ROOT, ...env }, stdio: ['ignore', 'pipe', 'pipe'] });
    let out = ''; let err = '';
    child.stdout.on('data', (b) => { out += b; }); child.stderr.on('data', (b) => { err += b; });
    child.on('error', (e) => resolve({ code: 1, out, err: e.message }));
    child.on('close', (code) => resolve({ code: code ?? 1, out, err }));
  });
}
async function ensureRuntime() {
  if (existsSync(path.join(PROJECT_ROOT, 'node_modules', '.package-lock.json'))) return '';
  mkdirSync(PROJECT_ROOT, { recursive: true });
  const result = await run('npm', ['install', '--no-audit', '--no-fund']);
  if (result.code) throw new Error(result.err || result.out || 'npm install failed');
  return result.out;
}
async function callTool(name, a = {}) {
  if (name === 'prepare_dashi_ppt') return textResult(await ensureRuntime() || 'Dashi PPT runtime is ready.');
  if (name === 'query_dashi_layouts') {
    const r = await run(process.execPath, ['scripts/layout-query.mjs', '--theme', arg(a.theme), '--limit', arg(a.limit, 12), '--item-count', arg(a.itemCount, 4), '--priority', arg(a.priority, 'balanced'), ...(a.needsMedia ? ['--needs-media'] : [])]);
    return textResult(r.out || r.err, r.code !== 0);
  }
  if (name === 'inspect_dashi_layout') {
    const layouts = a.layouts || (a.layout ? [a.layout] : []);
    if (!layouts.length) throw new Error('Provide layout or layouts');
    const r = await run(process.execPath, ['scripts/inspect-layout.mjs', '--compact', ...layouts]);
    return textResult(r.out || r.err, r.code !== 0);
  }
  if (name === 'scaffold_dashi_presentation') {
    const output = workspacePath(a.output); const args = ['scripts/goal-scaffold.mjs', '--title', arg(a.title), '--goal', arg(a.goal), '--theme', arg(a.theme), '--pages', arg(a.pages, 10), '--layout-variants', '3', '--out', output];
    if (a.contentBriefs) args.push('--content-briefs', workspacePath(a.contentBriefs));
    if (a.language) args.push('--language', a.language);
    const r = await run(process.execPath, args); return textResult(r.out || r.err, r.code !== 0);
  }
  if (name === 'validate_dashi_goal') {
    const goal = workspacePath(a.goal); const r = await run(process.execPath, ['scripts/validate-goal-spec.mjs', goal]); return textResult(r.out || r.err || 'Goal is valid.', r.code !== 0);
  }
  if (name === 'render_dashi_presentation') {
    await ensureRuntime(); const goal = workspacePath(a.goal); const output = workspacePath(a.output); const r = await run(process.execPath, ['scripts/render-goal-deck.jsx', goal, output]); return textResult(r.out || r.err, r.code !== 0);
  }
  if (name === 'export_dashi_presentation') {
    await ensureRuntime(); const deck = workspacePath(a.deck); const output = workspacePath(a.output); const r = await run(process.execPath, ['scripts/export-pptx.mjs', deck, output, ...(a.format === 'pdf' ? ['--pdf'] : [])]); return textResult(r.out || r.err, r.code !== 0);
  }
  throw new Error(`Unknown tool: ${name}`);
}

const rl = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
rl.on('line', async (line) => {
  if (!line.trim()) return;
  let message; try { message = JSON.parse(line); } catch { return; }
  if (message.method === 'initialize') {
    reply(message.id, { protocolVersion: '2024-11-05', capabilities: { tools: {} }, serverInfo: SERVER_INFO });
  } else if (message.method === 'notifications/initialized') {
    // notification: no response
  } else if (message.method === 'tools/list') {
    reply(message.id, { tools });
  } else if (message.method === 'tools/call') {
    try { reply(message.id, await callTool(message.params?.name, message.params?.arguments || {})); }
    catch (e) { reply(message.id, textResult(e.message || e, true)); }
  } else if (message.id != null) {
    fail(message.id, `Unsupported MCP method: ${message.method}`, -32601);
  }
});
