#!/usr/bin/env node
/**
 * Dashi PPT MCP server for Google Antigravity and other MCP clients.
 *
 * This server intentionally has no runtime dependency on the MCP SDK.  It
 * implements the MCP stdio transport directly so a fresh git clone can be
 * connected by Antigravity without a second npm install step.  Node.js 20+ is
 * required by the bundled Dashi PPT runtime.
 */

import { createInterface } from 'node:readline';
import { spawnSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const SERVER_DIR = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_REPO_ROOT = path.resolve(SERVER_DIR, '..');
const REPO_ROOT = resolveRepoRoot();
const SKILL_ROOT = path.join(REPO_ROOT, 'skills', 'dashi-ppt');
const PROJECT_ROOT = path.join(SKILL_ROOT, 'project');
const VERSION = readRuntimeVersion();
const MAX_COMMAND_OUTPUT = 16_000;
const MAX_TOOL_RESULT = 120_000;
const COMMAND_TIMEOUT_MS = 30 * 60 * 1000;
const SUPPORTED_PROTOCOL_VERSIONS = [
  '2025-06-18',
  '2025-03-26',
  '2024-11-05',
];

const THEMES = [
  'theme01',
  'theme02',
  'theme03',
  'theme04',
  'theme05',
  'theme06',
  'theme07',
  'theme08',
  'theme09',
  'theme10',
  'theme11',
  'theme12',
];

const TOOLS = [
  {
    name: 'dashi_ppt_scaffold',
    description:
      'Create a Dashi PPT goal.json scaffold with theme-aware layouts and three template variants plus one bespoke variant. Use this before editing slide copy.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['title', 'goal', 'theme', 'pages'],
      properties: {
        title: { type: 'string', description: 'Presentation title.' },
        goal: { type: 'string', description: 'What the presentation must communicate or help the audience decide.' },
        audience: { type: 'string', description: 'Target audience.' },
        owner: { type: 'string', description: 'Team or person owning the deck.' },
        theme: { type: 'string', enum: THEMES, description: 'One of the twelve bundled Dashi PPT themes.' },
        pages: { type: 'integer', minimum: 1, maximum: 50, description: 'Number of logical slides.' },
        output_path: { type: 'string', description: 'Workspace-relative or absolute path for goal.json. Defaults to output/<slug>/goal.json.' },
        content_briefs_path: { type: 'string', description: 'Path to a JSON array of per-slide content briefs.' },
        content_briefs: {
          type: 'array',
          description: 'Optional per-slide brief objects. They are written to a temporary file for the scaffold command.',
          items: { type: 'object', additionalProperties: true },
        },
        roles: {
          type: 'array',
          description: 'Optional body roles, one for every non-cover/non-closing slide.',
          items: { type: 'string' },
        },
        layout_variants: { type: 'integer', enum: [1, 3], description: 'Use 3 for the current four-variant schema; defaults to 3.' },
        seed: { type: 'string', description: 'Optional deterministic layout selection seed.' },
        workflow_run_id: { type: 'string', description: 'Optional workflow correlation id.' },
        chunk_size: { type: 'integer', minimum: 0, maximum: 50, description: 'Optional brief chunk size for long decks.' },
        planned_images: { type: 'integer', minimum: 0, maximum: 50, description: 'Reserve this many image slots when the brief needs media.' },
        needs_visual: { type: 'boolean', description: 'Tell the layout selector that visual media is needed.' },
        image_gen: { type: 'boolean', description: 'Tell the layout selector that generated images are planned.' },
        provided_images: { type: 'integer', minimum: 0, maximum: 50, description: 'Number of supplied image assets.' },
        provided_media: { type: 'integer', minimum: 0, maximum: 50, description: 'Number of supplied media assets.' },
      },
    },
  },
  {
    name: 'dashi_ppt_layout_query',
    description:
      'Find compatible Dashi PPT layouts by theme, role, content capacity, and media requirements. Returns machine-readable JSON.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        theme: { type: 'string', enum: THEMES },
        role: { type: 'string', description: 'Page role such as cover, metrics, comparison, process, trend, image, or risks.' },
        keyword: { type: 'string' },
        limit: { type: 'integer', minimum: 1, maximum: 100 },
        seed: { type: 'string' },
        title_chars: { type: 'integer', minimum: 0 },
        summary_chars: { type: 'integer', minimum: 0 },
        takeaway_chars: { type: 'integer', minimum: 0 },
        item_count: { type: 'integer', minimum: 0 },
        min_item_count: { type: 'integer', minimum: 0 },
        numeric_item_count: { type: 'integer', minimum: 0 },
        value_item_count: { type: 'integer', minimum: 0 },
        raw_numeric_item_count: { type: 'integer', minimum: 0 },
        textual_value_item_count: { type: 'integer', minimum: 0 },
        nested_depth: { type: 'integer', minimum: 0 },
        requires_value: { type: 'boolean' },
        priority: { type: 'string' },
        needs_media: { type: 'boolean' },
        planned_images: { type: 'integer', minimum: 0 },
        provided_images: { type: 'integer', minimum: 0 },
        provided_media: { type: 'integer', minimum: 0 },
        image_gen: { type: 'boolean' },
        needs_visual: { type: 'boolean' },
        media_count: { type: 'integer', minimum: 0 },
        media_kind: { type: 'string' },
        require_initial_media: { type: 'boolean' },
      },
    },
  },
  {
    name: 'dashi_ppt_inspect_layout',
    description:
      'Inspect the writable copy fields, array shapes, media slots, and control contract for one or more concrete Dashi PPT layouts.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['layouts'],
      properties: {
        layouts: {
          type: 'array',
          minItems: 1,
          maxItems: 30,
          items: { type: 'string' },
        },
        compact: { type: 'boolean', description: 'Return compact inspection output.' },
      },
    },
  },
  {
    name: 'dashi_ppt_stage_media',
    description:
      'Copy user-provided images or videos into a rendered deck assets directory and return the relative paths to use in goal.json.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['deck_path', 'media_paths'],
      properties: {
        deck_path: { type: 'string', description: 'Deck output directory, its ppt directory, or the rendered index.html path.' },
        media_paths: {
          type: 'array',
          minItems: 1,
          maxItems: 50,
          items: { type: 'string' },
        },
      },
    },
  },
  {
    name: 'dashi_ppt_render',
    description:
      'Render a completed Dashi PPT goal.json into an editable HTML deck, run the built-in validation gates, and start its local preview server.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['goal_path'],
      properties: {
        goal_path: { type: 'string', description: 'Workspace-relative or absolute path to goal.json.' },
        html_path: { type: 'string', description: 'Output index.html path. Defaults to <goal directory>/ppt/index.html.' },
        preview_port: { type: 'integer', minimum: 1, maximum: 65535, description: 'Optional preferred preview port.' },
      },
    },
  },
  {
    name: 'dashi_ppt_preview',
    description:
      'Start or restart the Dashi PPT local preview server for a rendered deck and return its browser URLs and process state.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['deck_path'],
      properties: {
        deck_path: { type: 'string', description: 'Rendered deck directory, ppt directory, or index.html path.' },
        port: { type: 'integer', minimum: 1, maximum: 65535 },
      },
    },
  },
  {
    name: 'dashi_ppt_export',
    description:
      'Export a rendered Dashi PPT deck to editable PPTX or PDF using the bundled local export engine.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['deck_path', 'format'],
      properties: {
        deck_path: { type: 'string', description: 'Rendered deck directory, ppt directory, or index.html path.' },
        format: { type: 'string', enum: ['pptx', 'pdf'] },
        output_path: { type: 'string', description: 'Destination file. Defaults beside the deck under output/<deck>/<deck>.<format>.' },
        title: { type: 'string', description: 'Optional document title used by the export engine.' },
      },
    },
  },
  {
    name: 'dashi_ppt_validate',
    description:
      'Run Dashi PPT goal-spec and rendered-copy validation. Optionally run the browser-based four-variant quality check.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['goal_path'],
      properties: {
        goal_path: { type: 'string' },
        html_path: { type: 'string' },
        quality: { type: 'boolean', description: 'Also run the browser-based four-variant quality validator.' },
      },
    },
  },
];

main().catch((error) => {
  // Never put logs on stdout: stdout is the MCP JSON-RPC channel.
  console.error(`[dashi-ppt-mcp] fatal: ${error?.stack || error}`);
  process.exitCode = 1;
});

async function main() {
  assertRuntime();
  const input = createInterface({ input: process.stdin, crlfDelay: Infinity });

  for await (const line of input) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    await handleLine(trimmed);
  }
}

async function handleLine(line) {
  let message;
  try {
    message = JSON.parse(line);
  } catch {
    writeMessage({
      jsonrpc: '2.0',
      id: null,
      error: { code: -32700, message: 'Parse error: expected one JSON object per line.' },
    });
    return;
  }

  if (Array.isArray(message)) {
    const responses = [];
    for (const item of message) {
      const response = await dispatch(item);
      if (response) responses.push(response);
    }
    if (responses.length) writeMessage(responses);
    return;
  }

  const response = await dispatch(message);
  if (response) writeMessage(response);
}

async function dispatch(message) {
  if (!message || typeof message !== 'object' || message.jsonrpc !== '2.0') {
    return jsonRpcError(message?.id ?? null, -32600, 'Invalid Request');
  }

  const method = String(message.method || '');
  const hasId = Object.prototype.hasOwnProperty.call(message, 'id');
  const id = message.id ?? null;

  // MCP notifications do not receive a response.
  if (!hasId) {
    if (method === 'notifications/initialized' || method === 'notifications/cancelled' || method === 'notifications/progress') {
      return null;
    }
    return null;
  }

  try {
    switch (method) {
      case 'initialize':
        return jsonRpcResult(id, initializeResult(message.params || {}));
      case 'ping':
        return jsonRpcResult(id, {});
      case 'tools/list':
        return jsonRpcResult(id, { tools: TOOLS });
      case 'tools/call':
        return jsonRpcResult(id, await callTool(message.params || {}));
      case 'resources/list':
        return jsonRpcResult(id, { resources: [] });
      case 'prompts/list':
        return jsonRpcResult(id, { prompts: [] });
      default:
        return jsonRpcError(id, -32601, `Method not found: ${method}`);
    }
  } catch (error) {
    return jsonRpcResult(id, {
      content: [{ type: 'text', text: formatError(error) }],
      isError: true,
    });
  }
}

function initializeResult(params) {
  const requested = String(params.protocolVersion || '');
  const protocolVersion = SUPPORTED_PROTOCOL_VERSIONS.includes(requested)
    ? requested
    : SUPPORTED_PROTOCOL_VERSIONS[0];

  return {
    protocolVersion,
    capabilities: {
      tools: { listChanged: false },
      logging: {},
    },
    serverInfo: {
      name: 'dashi-ppt',
      version: VERSION,
    },
    instructions:
      'Use dashi_ppt_scaffold, then edit the returned goal.json, then call dashi_ppt_render. Use dashi_ppt_export only after render validation passes.',
  };
}

async function callTool(params) {
  const name = String(params.name || '');
  const args = params.arguments && typeof params.arguments === 'object' ? params.arguments : {};
  if (!TOOLS.some((tool) => tool.name === name)) {
    throw new Error(`Unknown tool: ${name}`);
  }

  const value = await invokeTool(name, args);
  const text = stringifyResult(value);
  return {
    content: [{ type: 'text', text }],
  };
}

async function invokeTool(name, args) {
  switch (name) {
    case 'dashi_ppt_scaffold':
      return scaffold(args);
    case 'dashi_ppt_layout_query':
      return layoutQuery(args);
    case 'dashi_ppt_inspect_layout':
      return inspectLayouts(args);
    case 'dashi_ppt_stage_media':
      return stageMedia(args);
    case 'dashi_ppt_render':
      return renderDeck(args);
    case 'dashi_ppt_preview':
      return startPreview(args);
    case 'dashi_ppt_export':
      return exportDeck(args);
    case 'dashi_ppt_validate':
      return validateDeck(args);
    default:
      throw new Error(`Unsupported tool: ${name}`);
  }
}

function scaffold(args) {
  const title = requiredString(args.title, 'title');
  const goal = requiredString(args.goal, 'goal');
  const theme = requiredString(args.theme, 'theme');
  const pages = integerInRange(args.pages, 'pages', 1, 50);
  if (!THEMES.includes(theme)) throw new Error(`theme must be one of: ${THEMES.join(', ')}`);

  const outputPath = resolveGoalOutputPath(args.output_path, title);
  mkdirSync(path.dirname(outputPath), { recursive: true });

  const commandArgs = [
    scriptPath('scripts/goal-scaffold.mjs'),
    '--title', title,
    '--goal', goal,
    '--theme', theme,
    '--pages', String(pages),
    '--out', outputPath,
  ];
  appendOptional(commandArgs, '--audience', args.audience);
  appendOptional(commandArgs, '--owner', args.owner);
  appendOptional(commandArgs, '--layout-variants', args.layout_variants);
  appendOptional(commandArgs, '--seed', args.seed);
  appendOptional(commandArgs, '--workflow-run-id', args.workflow_run_id);
  appendOptional(commandArgs, '--chunk-size', args.chunk_size);

  if (Array.isArray(args.roles) && args.roles.length) {
    appendOptional(commandArgs, '--roles', args.roles.map((role) => String(role)).join(','));
  }

  for (const [key, flag] of [
    ['planned_images', '--planned-images'],
    ['provided_images', '--provided-images'],
    ['provided_media', '--provided-media'],
  ]) {
    if (args[key] !== undefined) appendOptional(commandArgs, flag, args[key]);
  }
  for (const [key, flag] of [
    ['needs_visual', '--needs-visual'],
    ['image_gen', '--image-gen'],
  ]) {
    if (args[key] === true) commandArgs.push(flag);
  }

  let temporaryBriefs = null;
  try {
    if (args.content_briefs_path !== undefined) {
      const briefsPath = resolveWorkspaceFile(args.content_briefs_path, 'content_briefs_path');
      appendOptional(commandArgs, '--content-briefs', briefsPath);
    } else if (args.content_briefs !== undefined) {
      if (!Array.isArray(args.content_briefs)) throw new Error('content_briefs must be an array');
      temporaryBriefs = path.join(fsTempDir(), `dashi-ppt-briefs-${process.pid}-${Date.now()}.json`);
      writeFileSync(temporaryBriefs, `${JSON.stringify(args.content_briefs, null, 2)}\n`);
      appendOptional(commandArgs, '--content-briefs', temporaryBriefs);
    }

    const result = runNodeCommand(commandArgs, { timeoutMs: 10 * 60 * 1000 });
    return {
      ok: true,
      goal_path: outputPath,
      deck_directory: path.dirname(outputPath),
      theme,
      pages,
      stdout: tail(result.stdout),
      stderr: tail(result.stderr),
    };
  } finally {
    if (temporaryBriefs) rmSync(temporaryBriefs, { force: true });
  }
}

function layoutQuery(args) {
  const commandArgs = [scriptPath('scripts/layout-query.mjs')];
  const scalarFlags = [
    ['theme', '--theme'],
    ['role', '--role'],
    ['keyword', '--keyword'],
    ['limit', '--limit'],
    ['seed', '--seed'],
    ['title_chars', '--title-chars'],
    ['summary_chars', '--summary-chars'],
    ['takeaway_chars', '--takeaway-chars'],
    ['item_count', '--item-count'],
    ['min_item_count', '--min-item-count'],
    ['numeric_item_count', '--numeric-item-count'],
    ['value_item_count', '--value-item-count'],
    ['raw_numeric_item_count', '--raw-numeric-item-count'],
    ['textual_value_item_count', '--textual-value-item-count'],
    ['nested_depth', '--nested-depth'],
    ['priority', '--priority'],
    ['planned_images', '--planned-images'],
    ['provided_images', '--provided-images'],
    ['provided_media', '--provided-media'],
    ['media_count', '--media-count'],
    ['media_kind', '--media-kind'],
  ];
  for (const [key, flag] of scalarFlags) {
    if (args[key] !== undefined && args[key] !== null && args[key] !== '') appendOptional(commandArgs, flag, args[key]);
  }

  const booleanFlags = [
    ['requires_value', '--requires-value'],
    ['needs_media', '--needs-media'],
    ['image_gen', '--image-gen'],
    ['needs_visual', '--needs-visual'],
    ['require_initial_media', '--require-initial-media'],
  ];
  for (const [key, flag] of booleanFlags) if (args[key] === true) commandArgs.push(flag);

  const result = runNodeCommand(commandArgs, { timeoutMs: 60_000 });
  return parseJsonOutput(result.stdout, 'layout query');
}

function inspectLayouts(args) {
  if (!Array.isArray(args.layouts) || args.layouts.length < 1) {
    throw new Error('layouts must be a non-empty array');
  }
  if (args.layouts.length > 30) throw new Error('layouts may contain at most 30 entries');
  const commandArgs = [scriptPath('scripts/inspect-layout.mjs')];
  if (args.compact !== false) commandArgs.push('--compact');
  for (const layout of args.layouts) {
    const value = requiredString(layout, 'layout');
    commandArgs.push('--layout', value);
  }
  const result = runNodeCommand(commandArgs, { timeoutMs: 60_000 });
  return parseJsonOutput(result.stdout, 'layout inspection');
}

function stageMedia(args) {
  const deckPath = requiredString(args.deck_path, 'deck_path');
  if (!Array.isArray(args.media_paths) || args.media_paths.length < 1) {
    throw new Error('media_paths must be a non-empty array');
  }
  const deckDir = resolveDeckDir(deckPath);
  const sources = args.media_paths.map((item) => resolveExternalFile(item, 'media path'));
  const commandArgs = [scriptPath('scripts/stage-media.mjs'), deckDir, ...sources];
  const result = runNodeCommand(commandArgs, { timeoutMs: 10 * 60 * 1000 });
  return parseJsonOutput(result.stdout, 'media staging');
}

function renderDeck(args) {
  const goalPath = resolveWorkspaceFile(requiredString(args.goal_path, 'goal_path'), 'goal_path');
  const htmlPath = args.html_path === undefined
    ? path.join(path.dirname(goalPath), 'ppt', 'index.html')
    : resolveWorkspaceOutput(args.html_path, 'html_path');
  if (!htmlPath.toLowerCase().endsWith('.html')) throw new Error('html_path must point to an .html file');
  mkdirSync(path.dirname(htmlPath), { recursive: true });

  const script = process.platform === 'win32'
    ? 'powershell.exe'
    : 'bash';
  const scriptArgs = process.platform === 'win32'
    ? ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(SKILL_ROOT, 'scripts', 'render_goal_deck.ps1'), goalPath, htmlPath]
    : [path.join(SKILL_ROOT, 'scripts', 'render_goal_deck.sh'), goalPath, htmlPath];
  const env = {};
  if (args.preview_port !== undefined) env.DASHI_PPT_PREVIEW_PORT = String(integerInRange(args.preview_port, 'preview_port', 1, 65535));
  const result = runCommand(script, scriptArgs, {
    cwd: REPO_ROOT,
    env,
    timeoutMs: COMMAND_TIMEOUT_MS,
    label: 'render',
  });
  if (!existsSync(htmlPath)) throw new Error(`Render completed without creating ${displayPath(htmlPath)}`);

  const preview = readPreviewState(path.dirname(htmlPath));
  return {
    ok: true,
    goal_path: goalPath,
    html_path: htmlPath,
    preview,
    stdout: tail(result.stdout),
    stderr: tail(result.stderr),
  };
}

function startPreview(args) {
  const deckDir = resolveDeckDir(requiredString(args.deck_path, 'deck_path'));
  const commandArgs = [scriptPath('scripts/start-preview-server.mjs'), deckDir];
  if (args.port !== undefined) commandArgs.push(String(integerInRange(args.port, 'port', 1, 65535)));
  const result = runNodeCommand(commandArgs, {
    timeoutMs: 90_000,
    env: { DASHI_PPT_PREVIEW_HOST: '0.0.0.0' },
  });
  const preview = readPreviewState(deckDir);
  return {
    ok: true,
    deck_path: deckDir,
    preview,
    stdout: tail(result.stdout),
    stderr: tail(result.stderr),
  };
}

function exportDeck(args) {
  const deckDir = resolveDeckDir(requiredString(args.deck_path, 'deck_path'));
  const format = requiredString(args.format, 'format').toLowerCase();
  if (!['pptx', 'pdf'].includes(format)) throw new Error('format must be pptx or pdf');
  const outPath = args.output_path === undefined
    ? defaultExportPath(deckDir, format)
    : resolveWorkspaceOutput(args.output_path, 'output_path');
  const commandArgs = [scriptPath('scripts/export-pptx.mjs'), deckDir, outPath];
  if (format === 'pdf') commandArgs.push('--pdf');
  if (args.title !== undefined) appendOptional(commandArgs, '--title', args.title);
  const result = runNodeCommand(commandArgs, {
    timeoutMs: COMMAND_TIMEOUT_MS,
    env: { INIT_CWD: REPO_ROOT },
  });
  if (!existsSync(outPath)) throw new Error(`Export completed without creating ${displayPath(outPath)}`);
  return {
    ok: true,
    format,
    deck_path: deckDir,
    output_path: outPath,
    bytes: statSync(outPath).size,
    stdout: tail(result.stdout),
    stderr: tail(result.stderr),
  };
}

function validateDeck(args) {
  const goalPath = resolveWorkspaceFile(requiredString(args.goal_path, 'goal_path'), 'goal_path');
  const htmlPath = args.html_path === undefined
    ? null
    : resolveWorkspaceFile(args.html_path, 'html_path');
  const checks = [];
  checks.push(runValidation('goal-spec', [scriptPath('scripts/validate-goal-spec.mjs'), goalPath]));
  if (htmlPath) {
    checks.push(runValidation('swiss', [scriptPath('scripts/validate-swiss-deck.mjs'), htmlPath]));
    checks.push(runValidation('goal-copy', [scriptPath('scripts/validate-goal-copy.mjs'), goalPath, htmlPath]));
    if (args.quality === true) {
      checks.push(runValidation('four-variant-quality', [
        scriptPath('scripts/validate-four-variant-quality.mjs'),
        '--deck', htmlPath,
        '--goal', goalPath,
      ], 10 * 60 * 1000));
    }
  }
  const failed = checks.filter((check) => !check.ok);
  if (failed.length) {
    const error = new Error(`Dashi PPT validation failed: ${failed.map((check) => check.name).join(', ')}`);
    error.details = checks;
    throw error;
  }
  return { ok: true, goal_path: goalPath, html_path: htmlPath, checks };
}

function runValidation(name, commandArgs, timeoutMs = 120_000) {
  try {
    const result = runNodeCommand(commandArgs, { timeoutMs });
    return { name, ok: true, stdout: tail(result.stdout), stderr: tail(result.stderr) };
  } catch (error) {
    return { name, ok: false, error: formatError(error) };
  }
}

function runNodeCommand(args, options = {}) {
  return runCommand(process.execPath, args, {
    cwd: PROJECT_ROOT,
    ...options,
  });
}

function runCommand(command, args, options = {}) {
  const env = {
    ...process.env,
    INIT_CWD: REPO_ROOT,
    ...(options.env || {}),
  };
  const result = spawnSync(command, args, {
    cwd: options.cwd || REPO_ROOT,
    env,
    encoding: 'utf8',
    timeout: options.timeoutMs || COMMAND_TIMEOUT_MS,
    maxBuffer: 64 * 1024 * 1024,
    windowsHide: true,
  });

  if (result.error) {
    const label = options.label ? `${options.label}: ` : '';
    throw new Error(`${label}${result.error.message}`);
  }
  if (result.status !== 0) {
    const label = options.label ? `${options.label} ` : '';
    const signal = result.signal ? ` (signal ${result.signal})` : '';
    const commandOutput = [result.stderr, result.stdout]
      .filter((value) => String(value || '').trim())
      .join('\n');
    const error = new Error(`${label}command failed with exit code ${result.status}${signal}: ${command} ${args.join(' ')}\n${tail(commandOutput)}`);
    error.stdout = result.stdout || '';
    error.stderr = result.stderr || '';
    error.status = result.status;
    throw error;
  }
  return {
    stdout: result.stdout || '',
    stderr: result.stderr || '',
    status: result.status,
  };
}

function scriptPath(relative) {
  const full = path.join(PROJECT_ROOT, relative);
  if (!existsSync(full)) throw new Error(`Dashi PPT runtime file not found: ${displayPath(full)}`);
  return full;
}

function resolveRepoRoot() {
  const configured = process.env.DASHI_PPT_REPO_ROOT;
  const candidate = configured
    ? (path.isAbsolute(configured) ? path.resolve(configured) : path.resolve(process.cwd(), configured))
    : DEFAULT_REPO_ROOT;
  return path.resolve(candidate);
}

function assertRuntime() {
  for (const required of [
    path.join(SKILL_ROOT, 'SKILL.md'),
    path.join(PROJECT_ROOT, 'package.json'),
    path.join(PROJECT_ROOT, 'scripts', 'goal-scaffold.mjs'),
  ]) {
    if (!existsSync(required)) throw new Error(`Dashi PPT runtime is incomplete: ${displayPath(required)}`);
  }
}

function readRuntimeVersion() {
  try {
    return JSON.parse(readFileSync(path.join(PROJECT_ROOT, 'package.json'), 'utf8')).version || 'unknown';
  } catch {
    return 'unknown';
  }
}

function resolveWorkspaceFile(value, label) {
  const full = resolveWorkspacePath(value, label);
  if (!existsSync(full)) throw new Error(`${label} does not exist: ${displayPath(full)}`);
  if (!statSync(full).isFile()) throw new Error(`${label} is not a file: ${displayPath(full)}`);
  return full;
}

function resolveWorkspaceOutput(value, label) {
  const full = resolveWorkspacePath(value, label);
  if (path.basename(full) === '.git' || full.includes(`${path.sep}.git${path.sep}`)) {
    throw new Error(`${label} cannot be inside .git`);
  }
  return full;
}

function resolveWorkspacePath(value, label) {
  const raw = requiredString(value, label);
  const full = path.isAbsolute(raw) ? path.resolve(raw) : path.resolve(REPO_ROOT, raw);
  const relative = path.relative(REPO_ROOT, full);
  if (relative === '..' || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new Error(`${label} must be inside the cloned repository: ${raw}`);
  }
  if (relative === '.git' || relative.startsWith(`.git${path.sep}`)) {
    throw new Error(`${label} cannot be inside .git: ${raw}`);
  }
  return full;
}

function resolveExternalFile(value, label) {
  const raw = requiredString(value, label);
  const full = path.isAbsolute(raw) ? path.resolve(raw) : path.resolve(REPO_ROOT, raw);
  if (!existsSync(full)) throw new Error(`${label} does not exist: ${raw}`);
  if (!statSync(full).isFile()) throw new Error(`${label} is not a file: ${raw}`);
  return full;
}

function resolveGoalOutputPath(value, title) {
  const raw = value === undefined || value === null || value === ''
    ? path.join('output', `${slugify(title)}-${Date.now()}`, 'goal.json')
    : requiredString(value, 'output_path');
  const full = resolveWorkspaceOutput(raw, 'output_path');
  if (!full.toLowerCase().endsWith('.json')) throw new Error('output_path must point to a .json file');
  return full;
}

function resolveDeckDir(value) {
  const full = resolveWorkspacePath(value, 'deck_path');
  if (full.toLowerCase().endsWith('.html')) return path.dirname(full);
  if (path.basename(full).toLowerCase() === 'ppt') return full;
  if (existsSync(path.join(full, 'index.html'))) return full;
  return path.join(full, 'ppt');
}

function defaultExportPath(deckDir, format) {
  const parent = path.dirname(deckDir);
  const deckName = slugify(path.basename(parent) || path.basename(deckDir) || 'presentation');
  return resolveWorkspaceOutput(path.join(parent, `${deckName}.${format}`), 'output_path');
}

function readPreviewState(deckDir) {
  const statePath = path.join(deckDir, '.preview-server.json');
  if (!existsSync(statePath)) return null;
  try {
    return JSON.parse(readFileSync(statePath, 'utf8'));
  } catch {
    return { state_path: statePath, error: 'Preview state exists but is not valid JSON.' };
  }
}

function fsTempDir() {
  const temp = path.join(os.tmpdir(), 'dashi-ppt-mcp');
  mkdirSync(temp, { recursive: true });
  return temp;
}

function parseJsonOutput(output, label) {
  const text = String(output || '').trim();
  if (!text) throw new Error(`${label} returned no JSON output`);
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`${label} returned invalid JSON: ${error.message}\n${tail(text)}`);
  }
}

function appendOptional(args, flag, value) {
  if (value === undefined || value === null || value === '') return;
  args.push(flag, String(value));
}

function requiredString(value, label) {
  if (typeof value !== 'string' || !value.trim()) throw new Error(`${label} is required`);
  return value.trim();
}

function integerInRange(value, label, min, max) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < min || number > max) {
    throw new Error(`${label} must be an integer between ${min} and ${max}`);
  }
  return number;
}

function slugify(value) {
  const slug = String(value || '')
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[^\p{Letter}\p{Number}]+/gu, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80);
  return slug || 'presentation';
}

function displayPath(full) {
  const relative = path.relative(REPO_ROOT, full);
  return relative && !relative.startsWith('..') && !path.isAbsolute(relative)
    ? relative
    : full;
}

function tail(value, limit = MAX_COMMAND_OUTPUT) {
  const text = String(value || '');
  if (text.length <= limit) return text;
  return `[...output truncated...]\n${text.slice(-limit)}`;
}

function stringifyResult(value) {
  const json = JSON.stringify(value, null, 2);
  if (json.length <= MAX_TOOL_RESULT) return json;
  return `${json.slice(0, MAX_TOOL_RESULT)}\n... result truncated ...`;
}

function formatError(error) {
  const details = error?.details ? `\n${JSON.stringify(error.details, null, 2)}` : '';
  return `Dashi PPT MCP tool failed: ${error?.message || error}${details}`;
}

function jsonRpcResult(id, result) {
  return { jsonrpc: '2.0', id, result };
}

function jsonRpcError(id, code, message) {
  return { jsonrpc: '2.0', id, error: { code, message } };
}

function writeMessage(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}
