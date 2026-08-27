#!/usr/bin/env node
/**
 * Build a native, editable PowerPoint deck from a Dashi PPT goal scaffold.
 *
 * The Dashi runtime goal.json remains the source of truth for the selected
 * theme-aware layouts. This fallback exporter keeps the same canonical brief
 * content but draws every visible element as editable PowerPoint text, lines,
 * shapes, and chart primitives. It is intentionally dependency-light so a
 * cloned repository can regenerate the samples without Chrome.
 *
 * Usage:
 *   node generate-pptx.mjs --goal <goal.json> --briefs <briefs.json> --out <deck.pptx>
 *
 * The original two positional arguments are still supported:
 *   node generate-pptx.mjs [goal.json] [deck.pptx]
 */

import { createRequire } from 'node:module';
import { existsSync, mkdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
let PptxGenJS;
try {
  // Prefer the full library when a clone has installed the runtime packages.
  PptxGenJS = require('../../skills/dashi-ppt/project/node_modules/pptxgenjs');
} catch {
  // The checked-in lite writer keeps sample regeneration usable offline.
  PptxGenJS = require('./pptxgenjs-lite.cjs');
}

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(HERE, '../..');
const DEFAULT_GOAL = path.join(HERE, 'goal.json');
const DEFAULT_OUTPUT = path.join(HERE, 'dashi-ppt-antigravity-mcp-demo.pptx');
const DEFAULT_BRIEFS = path.join(REPO_ROOT, 'examples/theme-samples/briefs.json');

const THEME_PROFILES = {
  theme01: {
    name: '轻拟态风', style: 'soft',
    colors: { bg: 'EEF0F5', panel: 'FFFFFF', panelAccent: 'E7ECFF', line: 'D7DCE8', text: '2B2B30', muted: '7B7F8C', accent: '5B8DEF', cyan: '46B083', purple: '7A5AE0', warning: 'E0A23A' },
  },
  theme02: {
    name: '炫光紫绿风', style: 'neon',
    colors: { bg: '110B22', panel: '21183A', panelAccent: '2B274C', line: '4B3B72', text: 'FBF8FF', muted: 'B9A9D4', accent: 'C7FF47', cyan: '69E6FF', purple: 'D78BFF', warning: 'FFBC6B' },
  },
  theme03: {
    name: '深浅代码风', style: 'code',
    colors: { bg: '08131F', panel: '101F30', panelAccent: '123040', line: '2A4C5F', text: 'EAF7FF', muted: '8EADBE', accent: '7CFFB2', cyan: '7EC8FF', purple: 'C0A8FF', warning: 'FFCB66' },
  },
  theme04: {
    name: '玻璃糖果风', style: 'glass',
    colors: { bg: 'F7EFF4', panel: 'FFFFFF', panelAccent: 'FDE1EA', line: 'E7CFE0', text: '301E31', muted: '866F85', accent: 'E95078', cyan: '47B8E8', purple: '9A67D3', warning: 'F3A84F' },
  },
  theme05: {
    name: '色谱图表风', style: 'spectrum',
    colors: { bg: 'F4F7FB', panel: 'FFFFFF', panelAccent: 'EAF3FF', line: 'CED9E8', text: '15263D', muted: '6B7B91', accent: '3A7DFF', cyan: '21B6A8', purple: '7A5CE0', warning: 'F39A3D' },
  },
  theme06: {
    name: '深色图谱风', style: 'graph',
    colors: { bg: '0B0E17', panel: '151B29', panelAccent: '1D2538', line: '323C58', text: 'F4F6FF', muted: '9AA7C2', accent: 'FFCC4D', cyan: '4FE4FF', purple: '9A83FF', warning: 'FF7C8B' },
  },
  theme07: {
    name: '冷白调研风', style: 'research',
    colors: { bg: 'F7F9FC', panel: 'FFFFFF', panelAccent: 'EEF4FA', line: 'C9D6E3', text: '1B2C3D', muted: '647485', accent: '2563A8', cyan: '3B82A0', purple: '6977B8', warning: 'B87333' },
  },
  theme08: {
    name: '黑金实验风', style: 'gold',
    colors: { bg: '090909', panel: '151515', panelAccent: '241E13', line: '4E422E', text: 'F8F0DD', muted: 'AFA48F', accent: 'D8A94A', cyan: 'BCA66A', purple: 'E3C47D', warning: 'E56A3D' },
  },
  theme09: {
    name: '深蓝杂志风', style: 'editorial',
    colors: { bg: '0A1936', panel: '10254A', panelAccent: '173560', line: '34527C', text: 'F7F4EA', muted: 'A3B3CE', accent: 'FFB14E', cyan: '7ED7E7', purple: 'C7B1FF', warning: 'F4826B' },
  },
  theme10: {
    name: '金色指数风', style: 'index',
    colors: { bg: 'F2EEE4', panel: 'FFFFFF', panelAccent: 'E9DFC8', line: 'CFC2A7', text: '202B39', muted: '747B83', accent: 'A67520', cyan: '398A9C', purple: '7355A9', warning: 'D95A3A' },
  },
  theme11: {
    name: '高能增长风', style: 'growth',
    colors: { bg: '17110B', panel: '281A11', panelAccent: '3A2114', line: '6D3A1E', text: 'FFF4E9', muted: 'C3AFA0', accent: 'FF6B3D', cyan: '51D7C1', purple: 'FF8BCB', warning: 'FFD166' },
  },
  theme12: {
    name: '声波霓虹风', style: 'sound',
    colors: { bg: 'F3E2E6', panel: 'FFFFFF', panelAccent: 'FAD5E5', line: 'DABFCB', text: '211821', muted: '7D7176', accent: 'F15A29', cyan: '3BB6EC', purple: 'C44EE0', warning: 'BAF04F' },
  },
};

const positional = [];
const flags = {};
for (let i = 0; i < process.argv.slice(2).length; i += 1) {
  const token = process.argv.slice(2)[i];
  if (token.startsWith('--')) {
    const key = token.slice(2).replaceAll('-', '_');
    const next = process.argv.slice(2)[i + 1];
    if (next && !next.startsWith('--')) {
      flags[key] = next;
      i += 1;
    } else {
      flags[key] = true;
    }
  } else {
    positional.push(token);
  }
}

const goalPath = path.resolve(flags.goal || positional[0] || DEFAULT_GOAL);
const outputPath = path.resolve(flags.out || positional[1] || DEFAULT_OUTPUT);
const briefsPath = flags.briefs ? path.resolve(flags.briefs) : DEFAULT_BRIEFS;

if (!existsSync(goalPath)) throw new Error(`goal.json not found: ${goalPath}`);
const spec = JSON.parse(readFileSync(goalPath, 'utf8'));
const briefSource = existsSync(briefsPath) ? JSON.parse(readFileSync(briefsPath, 'utf8')) : [];
const rawSlides = Array.isArray(spec.slides) ? spec.slides : [];
if (!rawSlides.length) throw new Error('goal.json contains no slides');

const themeKey = String(flags.theme || spec.themePack || 'theme03');
const theme = THEME_PROFILES[themeKey];
if (!theme) throw new Error(`Unknown theme ${themeKey}; expected theme01 through theme12`);
const C = theme.colors;
const FONT = {
  head: 'Aptos Display',
  body: 'Aptos',
  mono: 'Aptos Mono',
};

// Single-variant scaffold goals intentionally keep canonical copy in the
// shared brief source. If a goal already has frozen content, it wins.
const slides = rawSlides.map((entry, index) => ({
  ...(entry || {}),
  content: entry?.content || briefSource[index]?.content || { presentation: {}, meta: {} },
}));

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Dashi PPT';
pptx.company = 'Dashi PPT';
pptx.subject = String(spec.goal || spec.title || 'Dashi PPT demo');
pptx.title = String(spec.title || `Dashi PPT · ${theme.name}`);
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: FONT.head,
  bodyFontFace: FONT.body,
  lang: 'zh-CN',
};
pptx.defineSlideMaster({
  title: 'DASHI_THEME',
  background: { color: C.bg },
  objects: [],
  slideNumber: { x: 12.35, y: 7.08, color: C.muted, fontFace: FONT.body, fontSize: 8 },
});

slides.forEach((entry, index) => renderSlide(entry, index));
mkdirSync(path.dirname(outputPath), { recursive: true });
await pptx.writeFile({ fileName: outputPath });
console.log(`Wrote editable PPTX: ${outputPath}`);

function renderSlide(entry, index) {
  const slide = pptx.addSlide('DASHI_THEME');
  decorateSlide(slide, index, entry);
  const presentation = entry?.content?.presentation || {};
  const meta = entry?.content?.meta || {};
  const items = Array.isArray(presentation.items) ? presentation.items : [];
  const title = String(presentation.title || presentation.titleShort || spec.title || 'Dashi PPT');
  const summary = String(presentation.summary || presentation.summaryShort || '');
  const takeaway = String(presentation.takeaway || '');
  const kicker = String(meta.panelTitle || 'DASHI PPT × ANTIGRAVITY');

  if (index === 0) {
    renderCover(slide, presentation, meta);
    return;
  }
  if (index === slides.length - 1) {
    renderClosing(slide, presentation, meta);
    return;
  }

  addHeader(slide, kicker, title, summary, index);
  switch (index) {
    case 1:
      renderOperatingModel(slide, items, takeaway);
      break;
    case 2:
      renderMetrics(slide, items, takeaway);
      break;
    case 3:
      renderWorkflow(slide, items, takeaway);
      break;
    case 4:
      renderComparison(slide, items, takeaway);
      break;
    case 5:
      renderMaturity(slide, items, presentation.chartData || [], takeaway);
      break;
    case 6:
      renderRepository(slide, items, takeaway);
      break;
    case 7:
      renderRisks(slide, items, takeaway);
      break;
    case 8:
      renderActions(slide, items, takeaway);
      break;
    default:
      renderGeneric(slide, items, takeaway);
  }
}

function decorateSlide(slide, index, entry) {
  // These are deliberately native primitives rather than screenshots. The
  // same brief therefore remains editable while each bundled theme gets its
  // own background treatment, motif, and visual rhythm.
  switch (theme.style) {
    case 'soft':
      addNativeShape(slide, 'ellipse', 10.55, -0.48, 3.2, 2.15, C.purple, 78, null);
      addNativeShape(slide, 'ellipse', -0.72, 5.88, 2.55, 2.1, C.cyan, 84, null);
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.line, 0.8);
      break;
    case 'neon':
      addNativeShape(slide, 'ellipse', 10.25, -0.62, 3.55, 2.7, C.purple, 72, C.purple, 1.2, 86);
      addNativeShape(slide, 'ellipse', -0.58, 5.95, 2.4, 1.65, C.accent, 82, C.accent, 1, 82);
      addLine(slide, 8.85, 0.05, 12.85, 1.18, C.cyan, 2, { transparency: 25 });
      addLine(slide, 9.3, 0.05, 12.85, 1.04, C.accent, 0.8, { transparency: 35 });
      break;
    case 'code':
      for (let x = 0.55; x < 13.2; x += 0.72) addLine(slide, x, 0.12, x, 6.95, C.line, 0.35, { transparency: 70 });
      for (let y = 0.48; y < 7.1; y += 0.52) addLine(slide, 0.38, y, 13.08, y, C.line, 0.35, { transparency: 72 });
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.accent, 1.1, { transparency: 25 });
      break;
    case 'glass':
      addNativeShape(slide, 'ellipse', 10.3, -0.38, 3.55, 2.3, C.accent, 82, C.accent, 1, 62);
      addNativeShape(slide, 'ellipse', -0.55, 5.7, 2.9, 2.25, C.cyan, 86, C.cyan, 1, 70);
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.line, 0.75);
      break;
    case 'spectrum':
      ['3A7DFF', '21B6A8', '7A5CE0', 'F39A3D', 'E95078'].forEach((color, i) => {
        addNativeShape(slide, 'rect', 0.72 + i * 2.36, 0.18, 2.12, 0.045, color, 0, null);
      });
      for (let x = 0.95; x < 13; x += 1.18) addLine(slide, x, 6.6, x, 6.86, C.line, 0.65);
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.line, 0.8);
      break;
    case 'graph':
      for (let x = 0.45; x < 13.2; x += 0.72) addLine(slide, x, 0.12, x, 6.95, C.line, 0.4, { transparency: 62 });
      for (let y = 0.52; y < 7.1; y += 0.52) addLine(slide, 0.38, y, 13.08, y, C.line, 0.4, { transparency: 64 });
      addLine(slide, 0.75, 6.92, 12.55, 6.92, C.accent, 1.4);
      addNativeShape(slide, 'ellipse', 11.92, 0.28, 0.16, 0.16, C.accent, 0, null);
      break;
    case 'research':
      addNativeShape(slide, 'rect', 0.42, 0.2, 0.05, 6.72, C.accent, 0, null);
      addLine(slide, 0.75, 6.92, 12.55, 6.92, C.line, 0.8);
      [1.08, 1.32, 1.56].forEach((y, i) => addNativeShape(slide, 'rect', 0.26, y, 0.12, 0.12, [C.accent, C.cyan, C.purple][i], 0, null));
      break;
    case 'gold':
      addLine(slide, 0.48, 0.35, 1.35, 0.35, C.accent, 1.1);
      addLine(slide, 0.48, 0.35, 0.48, 1.22, C.accent, 1.1);
      addLine(slide, 12.7, 6.58, 12.7, 7.02, C.accent, 1.1);
      addLine(slide, 11.85, 7.02, 12.7, 7.02, C.accent, 1.1);
      addNativeShape(slide, 'ellipse', 10.75, -0.52, 2.7, 2.7, C.accent, 100, C.accent, 1.2, 70);
      break;
    case 'editorial':
      addNativeShape(slide, 'rect', 0.28, 0, 0.18, 7.5, C.accent, 0, null);
      addNativeShape(slide, 'rect', 12.72, 0, 0.06, 7.5, C.cyan, 0, null);
      addLine(slide, 0.78, 6.92, 12.55, 6.92, C.line, 0.8);
      break;
    case 'index':
      addLine(slide, 0.72, 0.37, 12.55, 0.37, C.accent, 1.2);
      for (let x = 0.9; x < 12.5; x += 0.48) addLine(slide, x, 0.37, x, 0.49, C.line, 0.55);
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.line, 0.8);
      break;
    case 'growth':
      addLine(slide, 8.8, 0.02, 13.2, 1.32, C.accent, 5, { transparency: 40 });
      addLine(slide, 9.48, 0.02, 13.2, 1.12, C.purple, 1.2, { transparency: 25 });
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.accent, 1.4);
      break;
    case 'sound':
      for (let i = 0; i < 6; i += 1) {
        const x = 10.2 + i * 0.42;
        const h = 0.22 + (i % 3) * 0.12;
        addLine(slide, x, 0.72 - h, x, 0.72 + h, i % 2 ? C.accent : C.purple, 2.1, { transparency: 18 });
      }
      addNativeShape(slide, 'ellipse', -0.55, 5.82, 2.35, 1.72, C.cyan, 84, C.cyan, 1, 70);
      addNativeShape(slide, 'ellipse', 11.25, 5.95, 2.05, 1.45, C.purple, 88, C.purple, 1, 72);
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.line, 0.8);
      break;
    default:
      addLine(slide, 0.72, 6.92, 12.55, 6.92, C.line, 0.8);
  }

  const layout = entry?.layout || entry?.variants?.[0]?.layout || 'native-editable';
  addText(slide, `${themeKey} · ${layout}`, 0.78, 7.02, 6.5, 0.16, {
    fontFace: FONT.mono, fontSize: 7, color: C.muted, breakLine: false,
  });
  addText(slide, theme.name, 9.4, 7.02, 2.6, 0.16, {
    fontSize: 7, color: C.muted, align: 'right', breakLine: false,
  });
}

function addNativeShape(slide, type, x, y, w, h, fill, transparency = 0, line = null, lineWidth = 0.8, lineTransparency = 0) {
  const shapeType = type === 'ellipse' ? pptx.ShapeType.ellipse : pptx.ShapeType.rect;
  slide.addShape(shapeType, {
    x, y, w, h,
    fill: { color: fill, transparency },
    line: { color: line || fill, width: line ? lineWidth : 0, transparency: line ? lineTransparency : 100 },
  });
}

function renderCover(slide, presentation, meta) {
  addText(slide, `${themeKey.toUpperCase()}  /  ${theme.name}  /  ANTIGRAVITY EDITION`, 0.72, 0.62, 8.1, 0.28, {
    fontSize: 10, bold: true, color: C.accent, charSpacing: 1.05, breakLine: false,
  });
  addText(slide, String(presentation.title || spec.title), 0.72, 1.22, 7.05, 1.35, {
    fontFace: FONT.head, fontSize: 30, bold: true, color: C.text, breakLine: false,
  });
  addText(slide, String(presentation.summary || ''), 0.76, 2.88, 5.7, 0.72, {
    fontSize: 16, color: C.muted, breakLine: false,
  });
  addText(slide, String(presentation.takeaway || ''), 0.76, 4.02, 5.6, 0.72, {
    fontSize: 18, bold: true, color: C.accent, breakLine: false,
  });

  // Native shapes form an editable Skill → MCP → Runtime diagram.
  addText(slide, 'WORKFLOW STACK', 8.2, 0.92, 3.9, 0.25, {
    fontSize: 9, bold: true, color: C.muted, charSpacing: 1,
  });
  const stack = [
    ['01', 'Agent Skill', '方法与质量边界', C.purple],
    ['02', 'MCP Server', '工具连接与执行', C.cyan],
    ['03', 'Dashi Runtime', 'HTML / PPTX 产物', C.accent],
  ];
  stack.forEach(([number, label, detail, color], i) => {
    const y = 1.45 + i * 1.25;
    addBox(slide, 8.18, y, 4.25, 0.86, C.panel, C.line, 1);
    addText(slide, number, 8.42, y + 0.19, 0.4, 0.28, { fontSize: 10, bold: true, color });
    addText(slide, label, 8.94, y + 0.16, 2.45, 0.3, { fontSize: 16, bold: true, color: C.text });
    addText(slide, detail, 8.94, y + 0.5, 2.9, 0.2, { fontSize: 9, color: C.muted });
    if (i < stack.length - 1) {
      addLine(slide, 10.3, y + 0.88, 10.3, y + 1.2, color, 1.4, { endArrowType: 'triangle' });
    }
  });

  const metrics = [
    ['12', '视觉主题'],
    ['1,020', '版式页面'],
    ['8,576', '可调控件'],
  ];
  metrics.forEach(([value, label], i) => {
    const x = 0.76 + i * 2.0;
    addText(slide, value, x, 6.38, 1.25, 0.34, { fontFace: FONT.head, fontSize: 20, bold: true, color: C.text });
    addText(slide, label, x, 6.78, 1.4, 0.18, { fontSize: 9, color: C.muted });
  });
  addText(slide, String(meta.pageLabel || '2026'), 11.7, 6.78, 0.8, 0.18, { fontSize: 9, color: C.muted, align: 'right' });
}

function renderOperatingModel(slide, items, takeaway) {
  const cards = [
    ['SKILL', items[0], C.purple],
    ['MCP', items[1], C.cyan],
    ['OUTPUT', items[2], C.accent],
  ];
  cards.forEach(([label, item, color], i) => {
    const x = 0.75 + i * 4.18;
    addBox(slide, x, 2.05, 3.72, 3.1, C.panel, C.line, 1);
    addText(slide, `0${i + 1}`, x + 0.28, 2.32, 0.45, 0.28, { fontSize: 10, bold: true, color });
    addText(slide, label, x + 2.38, 2.34, 1.05, 0.22, { fontSize: 9, bold: true, color, align: 'right', charSpacing: 1 });
    addText(slide, String(item?.label || label), x + 0.28, 2.9, 3.05, 0.5, { fontFace: FONT.head, fontSize: 22, bold: true, color: C.text });
    addText(slide, String(item?.detail || ''), x + 0.28, 3.75, 3.05, 0.7, { fontSize: 13, color: C.muted, breakLine: false });
    addLine(slide, x + 0.28, 4.72, x + 3.28, 4.72, color, 2);
  });
  addText(slide, takeaway, 0.8, 5.72, 11.8, 0.52, { fontSize: 16, bold: true, color: C.accent, align: 'center' });
}

function renderMetrics(slide, items, takeaway) {
  items.slice(0, 4).forEach((item, i) => {
    const x = 0.75 + (i % 2) * 6.0;
    const y = 1.95 + Math.floor(i / 2) * 1.95;
    addBox(slide, x, y, 5.35, 1.45, i === 0 ? C.panelAccent : C.panel, C.line, 1);
    addText(slide, String(item?.displayValue || item?.value || ''), x + 0.3, y + 0.25, 1.55, 0.5, { fontFace: FONT.head, fontSize: 27, bold: true, color: i === 0 ? C.accent : C.text });
    addText(slide, String(item?.unit || ''), x + 1.92, y + 0.48, 0.7, 0.22, { fontSize: 10, color: C.muted });
    addText(slide, String(item?.label || ''), x + 2.75, y + 0.25, 2.2, 0.3, { fontSize: 15, bold: true, color: C.text, align: 'right' });
    addText(slide, String(item?.detail || ''), x + 2.0, y + 0.82, 3.0, 0.25, { fontSize: 9, color: C.muted, align: 'right' });
  });
  addText(slide, takeaway, 0.8, 6.05, 11.8, 0.42, { fontSize: 15, bold: true, color: C.accent, align: 'center' });
}

function renderWorkflow(slide, items, takeaway) {
  const list = items.slice(0, 5);
  const yLine = 3.3;
  addLine(slide, 1.05, yLine, 12.15, yLine, C.line, 2);
  list.forEach((item, i) => {
    const x = 0.78 + i * 2.45;
    addBox(slide, x + 0.25, yLine - 0.25, 0.5, 0.5, C.bg, C.accent, 2);
    addText(slide, String(i + 1).padStart(2, '0'), x + 0.25, yLine - 0.1, 0.5, 0.16, { fontSize: 9, bold: true, color: C.accent, align: 'center' });
    addText(slide, String(item?.label || ''), x, 4.02, 1.95, 0.28, { fontSize: 13, bold: true, color: C.text, align: 'center' });
    addText(slide, String(item?.detail || ''), x - 0.15, 4.52, 2.25, 0.62, { fontSize: 9, color: C.muted, align: 'center', breakLine: false });
  });
  addText(slide, takeaway, 0.8, 6.1, 11.8, 0.4, { fontSize: 15, bold: true, color: C.accent, align: 'center' });
}

function renderComparison(slide, items, takeaway) {
  addBox(slide, 0.75, 1.95, 5.7, 3.95, C.panel, C.line, 1);
  addBox(slide, 6.9, 1.95, 5.7, 3.95, C.panelAccent, C.accent, 1);
  addText(slide, '手工制作', 1.1, 2.28, 2.2, 0.35, { fontFace: FONT.head, fontSize: 22, bold: true, color: C.text });
  addText(slide, 'AGENT + MCP', 7.25, 2.28, 2.8, 0.35, { fontFace: FONT.head, fontSize: 22, bold: true, color: C.accent });
  const left = items[0] || {};
  const right = items[1] || {};
  addBullet(slide, 1.1, 3.0, String(left.label || '依赖个人经验'), String(left.detail || ''), C.muted);
  addBullet(slide, 1.1, 4.12, String(items[2]?.label || '质量后置'), String(items[2]?.detail || ''), C.muted);
  addBullet(slide, 7.25, 3.0, String(right.label || '工具化执行'), String(right.detail || ''), C.text, C.accent);
  addBullet(slide, 7.25, 4.12, String(items[3]?.label || '源文件可交接'), String(items[3]?.detail || ''), C.text, C.accent);
  addText(slide, takeaway, 0.8, 6.25, 11.8, 0.35, { fontSize: 15, bold: true, color: C.accent, align: 'center' });
}

function renderMaturity(slide, items, chartData, takeaway) {
  const data = chartData.length ? chartData : items;
  const points = data.slice(0, 4).map((item, i) => ({
    label: String(item?.label || ''),
    value: Number(item?.value) || (i + 1) * 25,
  }));
  addText(slide, 'AUTOMATION MATURITY', 0.82, 1.83, 3.2, 0.22, { fontSize: 9, bold: true, color: C.muted, charSpacing: 1 });
  addLine(slide, 1.05, 5.25, 8.55, 5.25, C.line, 1);
  addLine(slide, 1.05, 2.2, 1.05, 5.25, C.line, 1);
  [0, 25, 50, 75, 100].forEach((value) => {
    const y = 5.25 - (value / 100) * 3.05;
    addLine(slide, 1.05, y, 8.55, y, C.line, 0.5, { transparency: 45 });
    addText(slide, `${value}`, 0.65, y - 0.08, 0.3, 0.14, { fontSize: 8, color: C.muted, align: 'right' });
  });
  const coords = points.map((point, i) => ({
    x: 1.55 + i * 2.15,
    y: 5.25 - Math.max(0, Math.min(100, point.value)) / 100 * 3.05,
  }));
  coords.forEach((point, i) => {
    if (i) addLine(slide, coords[i - 1].x, coords[i - 1].y, point.x, point.y, C.accent, 2.5);
    addBox(slide, point.x - 0.09, point.y - 0.09, 0.18, 0.18, C.accent, C.accent, 2);
    addText(slide, `${points[i].value}%`, point.x - 0.35, point.y - 0.45, 0.7, 0.2, { fontSize: 10, bold: true, color: C.text, align: 'center' });
    addText(slide, points[i].label, point.x - 0.55, 5.52, 1.1, 0.25, { fontSize: 9, color: C.muted, align: 'center' });
  });
  addBox(slide, 9.15, 2.12, 3.35, 3.65, C.panel, C.line, 1);
  addText(slide, '关键判断', 9.5, 2.52, 1.5, 0.3, { fontSize: 16, bold: true, color: C.text });
  addText(slide, takeaway, 9.5, 3.35, 2.55, 1.1, { fontSize: 15, bold: true, color: C.accent, breakLine: false });
}

function renderRepository(slide, items, takeaway) {
  const layers = [
    ['01', '.agents/skills/dashi-ppt', items[0], C.purple],
    ['02', '.agents/mcp_config.json', items[1], C.cyan],
    ['03', 'mcp/server.mjs  +  skills/dashi-ppt', items[2] || items[3], C.accent],
  ];
  layers.forEach(([number, pathLabel, item, color], i) => {
    const y = 1.88 + i * 1.28;
    addBox(slide, 0.85, y, 11.65, 0.92, C.panel, C.line, 1);
    addText(slide, number, 1.15, y + 0.29, 0.5, 0.2, { fontSize: 10, bold: true, color });
    addText(slide, pathLabel, 1.85, y + 0.22, 4.55, 0.28, { fontFace: FONT.mono, fontSize: 14, bold: true, color: C.text });
    addText(slide, String(item?.detail || ''), 6.65, y + 0.28, 5.2, 0.22, { fontSize: 10, color: C.muted, align: 'right' });
  });
  addText(slide, takeaway, 0.8, 6.2, 11.8, 0.35, { fontSize: 15, bold: true, color: C.accent, align: 'center' });
}

function renderRisks(slide, items, takeaway) {
  const colors = [C.warning, C.purple, C.cyan, C.accent];
  items.slice(0, 4).forEach((item, i) => {
    const x = 0.75 + (i % 2) * 6.0;
    const y = 1.95 + Math.floor(i / 2) * 1.72;
    addBox(slide, x, y, 5.35, 1.22, C.panel, C.line, 1);
    addBox(slide, x, y, 0.1, 1.22, colors[i], colors[i], 0);
    addText(slide, String(item?.label || ''), x + 0.35, y + 0.22, 1.9, 0.28, { fontSize: 15, bold: true, color: C.text });
    addText(slide, String(item?.detail || ''), x + 0.35, y + 0.68, 4.55, 0.25, { fontSize: 9, color: C.muted, breakLine: false });
  });
  addText(slide, takeaway, 0.8, 6.12, 11.8, 0.38, { fontSize: 15, bold: true, color: C.accent, align: 'center' });
}

function renderActions(slide, items, takeaway) {
  items.slice(0, 4).forEach((item, i) => {
    const x = 0.8 + i * 3.05;
    addText(slide, String(i + 1).padStart(2, '0'), x, 2.05, 0.6, 0.28, { fontFace: FONT.head, fontSize: 18, bold: true, color: C.accent });
    addLine(slide, x, 2.7, x + 2.25, 2.7, i === 0 ? C.accent : C.line, 2);
    addText(slide, String(item?.label || ''), x, 3.12, 2.4, 0.55, { fontSize: 16, bold: true, color: C.text, breakLine: false });
    addText(slide, String(item?.detail || ''), x, 4.05, 2.35, 0.8, { fontSize: 10, color: C.muted, breakLine: false });
  });
  addBox(slide, 0.8, 5.55, 11.65, 0.76, C.panelAccent, C.accent, 1);
  addText(slide, takeaway, 1.15, 5.79, 10.95, 0.25, { fontSize: 14, bold: true, color: C.accent, align: 'center' });
}

function renderGeneric(slide, items, takeaway) {
  items.slice(0, 4).forEach((item, i) => {
    const x = 0.8 + (i % 2) * 6.0;
    const y = 1.95 + Math.floor(i / 2) * 1.65;
    addBox(slide, x, y, 5.3, 1.25, C.panel, C.line, 1);
    addText(slide, String(item?.label || ''), x + 0.3, y + 0.25, 2.2, 0.25, { fontSize: 15, bold: true, color: C.text });
    addText(slide, String(item?.detail || ''), x + 0.3, y + 0.7, 4.55, 0.25, { fontSize: 10, color: C.muted });
  });
  addText(slide, takeaway, 0.8, 6.1, 11.8, 0.4, { fontSize: 15, bold: true, color: C.accent, align: 'center' });
}

function renderClosing(slide, presentation, meta) {
  addText(slide, String(meta.panelTitle || 'DASHI PPT'), 0.75, 0.78, 4.2, 0.25, { fontSize: 10, bold: true, color: C.accent, charSpacing: 1 });
  addText(slide, String(presentation.title || '可调用 · 可编辑 · 可复用'), 0.75, 1.48, 8.8, 1.05, { fontFace: FONT.head, fontSize: 30, bold: true, color: C.text });
  addText(slide, String(presentation.summary || ''), 0.78, 2.88, 6.4, 0.38, { fontSize: 16, color: C.muted });
  addBox(slide, 0.78, 4.25, 8.45, 1.0, C.panel, C.line, 1);
  addText(slide, 'git clone -b arena/01a038e7-dashi-ppt-skill', 1.08, 4.51, 7.75, 0.22, { fontFace: FONT.mono, fontSize: 12, color: C.cyan });
  addText(slide, 'https://github.com/mqgg5630-cyber/dashi-ppt-skill.git', 1.08, 4.86, 7.75, 0.22, { fontFace: FONT.mono, fontSize: 10, color: C.muted });
  addText(slide, String(presentation.takeaway || ''), 0.8, 6.22, 8.4, 0.38, { fontSize: 15, bold: true, color: C.accent });
  addText(slide, 'MCP  /  SKILL  /  PPTX', 9.3, 5.15, 2.85, 0.3, { fontSize: 12, bold: true, color: C.purple, align: 'right', charSpacing: 1 });
}

function addHeader(slide, kicker, title, summary, index) {
  addText(slide, kicker.toUpperCase(), 0.75, 0.42, 5.5, 0.2, { fontSize: 9, bold: true, color: C.accent, charSpacing: 1 });
  addText(slide, title, 0.75, 0.78, 10.5, 0.58, { fontFace: FONT.head, fontSize: 25, bold: true, color: C.text, breakLine: false });
  addText(slide, summary, 0.78, 1.46, 9.4, 0.32, { fontSize: 11, color: C.muted, breakLine: false });
  addText(slide, String(index).padStart(2, '0'), 11.65, 0.43, 0.85, 0.2, { fontSize: 9, color: C.muted, align: 'right' });
  addLine(slide, 0.78, 1.88, 12.52, 1.88, C.line, 0.7);
}

function addBullet(slide, x, y, label, detail, color = C.text, bulletColor = C.muted) {
  addBox(slide, x, y + 0.02, 0.12, 0.12, bulletColor, bulletColor, 2);
  addText(slide, label, x + 0.28, y - 0.02, 4.5, 0.25, { fontSize: 13, bold: true, color });
  addText(slide, detail, x + 0.28, y + 0.36, 4.4, 0.24, { fontSize: 9, color: C.muted });
}

function addText(slide, text, x, y, w, h, options = {}) {
  slide.addText(String(text ?? ''), {
    x, y, w, h,
    margin: 0,
    fontFace: options.fontFace || FONT.body,
    fontSize: options.fontSize || 12,
    color: options.color || C.text,
    bold: options.bold || false,
    align: options.align || 'left',
    valign: options.valign || 'mid',
    breakLine: options.breakLine,
    fit: 'shrink',
    charSpacing: options.charSpacing,
    italic: options.italic,
    transparency: options.transparency,
  });
}

function addBox(slide, x, y, w, h, fill, line, radius = 0) {
  slide.addShape(radius ? pptx.ShapeType.roundRect : pptx.ShapeType.rect, {
    x, y, w, h,
    rectRadius: radius,
    fill: { color: fill },
    line: { color: line || fill, width: line ? 0.8 : 0, transparency: line ? 0 : 100 },
    radius,
  });
}

function addLine(slide, x1, y1, x2, y2, color, width = 1, options = {}) {
  slide.addShape(pptx.ShapeType.line, {
    x: x1, y: y1, w: x2 - x1, h: y2 - y1,
    line: { color, width, beginArrowType: options.beginArrowType, endArrowType: options.endArrowType, transparency: options.transparency || 0 },
  });
}
