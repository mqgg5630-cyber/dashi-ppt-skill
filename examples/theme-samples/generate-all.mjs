#!/usr/bin/env node
/**
 * Regenerate all twelve native editable sample decks from their goal.json
 * scaffolds and the shared briefs.json source.
 *
 * Run from the repository root:
 *   node examples/theme-samples/generate-all.mjs
 *
 * Pass theme keys to limit the run, for example:
 *   node examples/theme-samples/generate-all.mjs theme03 theme08
 */

import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '../..');
const GENERATOR = path.join(ROOT, 'examples/antigravity-mcp-demo/generate-pptx.mjs');
const BRIEFS = path.join(HERE, 'briefs.json');
const requested = process.argv.slice(2).filter((value) => !value.startsWith('--'));
const themes = requested.length
  ? requested
  : Array.from({ length: 12 }, (_, index) => `theme${String(index + 1).padStart(2, '0')}`);

for (const theme of themes) {
  if (!/^theme(?:0[1-9]|1[0-2])$/.test(theme)) {
    throw new Error(`Unknown theme ${theme}; use theme01 through theme12`);
  }
  const goal = path.join(HERE, theme, 'goal.json');
  const output = path.join(HERE, theme, `${theme}-example.pptx`);
  const result = spawnSync(process.execPath, [
    GENERATOR,
    '--goal', goal,
    '--briefs', BRIEFS,
    '--out', output,
  ], { cwd: ROOT, stdio: 'inherit' });
  if (result.status !== 0) process.exit(result.status || 1);
}

console.log(`Regenerated ${themes.length} theme sample deck${themes.length === 1 ? '' : 's'}.`);
