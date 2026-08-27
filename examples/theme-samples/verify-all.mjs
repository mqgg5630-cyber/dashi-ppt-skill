#!/usr/bin/env node
/** Verify every checked-in theme goal and native PPTX artifact. */

import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '../..');
const VALIDATOR = path.join(ROOT, 'skills/dashi-ppt/project/scripts/validate-goal-spec.mjs');
const themes = Array.from({ length: 12 }, (_, index) => `theme${String(index + 1).padStart(2, '0')}`);

for (const theme of themes) {
  const goal = path.join(HERE, theme, 'goal.json');
  const deck = path.join(HERE, theme, `${theme}-example.pptx`);
  if (!existsSync(goal) || !existsSync(deck)) throw new Error(`Missing artifact for ${theme}`);

  const goalResult = spawnSync(process.execPath, [VALIDATOR, goal], { cwd: ROOT, encoding: 'utf8' });
  if (goalResult.status !== 0) throw new Error(`${theme} goal validation failed:\n${goalResult.stdout}\n${goalResult.stderr}`);
  execFileSync('unzip', ['-t', deck], { stdio: 'ignore' });
  const names = execFileSync('unzip', ['-Z1', deck], { encoding: 'utf8' });
  const slideFiles = (names.match(/^ppt\/slides\/slide\d+\.xml$/gm) || []).length;
  const presentation = execFileSync('unzip', ['-p', deck, 'ppt/presentation.xml'], { encoding: 'utf8' });
  const slideIds = (presentation.match(/<p:sldId /g) || []).length;
  if (slideFiles !== 8 || slideIds !== 8) {
    throw new Error(`${theme} expected 8 slides, found files=${slideFiles}, ids=${slideIds}`);
  }
  console.log(`${theme}: goal ok · pptx zip ok · 8 slides`);
}
