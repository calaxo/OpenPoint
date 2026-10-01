// Croise la vue XML (specimen/timeline.json, issu de parse-timing.mjs) avec la vue
// COM de PowerPoint (specimen/com-dump.json, issu de dump-com.ps1).
// Hypothèse à valider : l'ordre document des noeuds à presetID dans la mainSeq XML
// == l'ordre de TimeLine.MainSequence côté COM. Si vrai, on en déduit la table
// presetClass:presetID:subtype -> msoAnimEffect (nom lisible de l'effet).

import { readFileSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const timeline = JSON.parse(readFileSync(join(ROOT, 'specimen', 'timeline.json'), 'utf8'));
const comDump = JSON.parse(readFileSync(join(ROOT, 'specimen', 'com-dump.json'), 'utf8').replace(/^﻿/, ''));

const mapping = new Map(); // preset -> Set(effectName)
let pairs = 0, mismatches = 0;

for (const sl of timeline) {
  if (!sl.animated) continue;
  const com = comDump.find((c) => c.slide === sl.slide);
  const xmlEffects = sl.clickSteps.flatMap((s) => s.effects);
  if (xmlEffects.length !== com.effectCount) {
    console.log(`slide ${sl.slide}: XML=${xmlEffects.length} vs COM=${com.effectCount} !!`);
    continue;
  }
  xmlEffects.forEach((xe, i) => {
    const ce = com.effects[i];
    pairs++;
    if (String(ce.shapeId) !== String(xe.spid)) {
      mismatches++;
      console.log(`slide ${sl.slide} #${i + 1}: spid XML=${xe.spid} vs COM shapeId=${ce.shapeId} (${ce.shape})`);
      return;
    }
    const name = ce.effectName?.replace('msoAnimEffect', '') ?? `#${ce.effectType}`;
    const key = xe.preset;
    if (!mapping.has(key)) mapping.set(key, new Map());
    const m = mapping.get(key);
    m.set(name, (m.get(name) ?? 0) + 1);
  });
}

console.log(`\n${pairs} paires XML<->COM, ${mismatches} désaccords de forme\n`);
console.log('presetClass:presetID:subtype -> effet PowerPoint (occurrences)');
const rows = [...mapping.entries()].sort();
for (const [preset, names] of rows) {
  const detail = [...names.entries()].map(([n, c]) => `${n} x${c}`).join(', ');
  console.log(`  ${preset.padEnd(14)} -> ${detail}`);
}

writeFileSync(
  join(ROOT, 'specimen', 'preset-mapping.json'),
  JSON.stringify(Object.fromEntries(rows.map(([k, v]) => [k, Object.fromEntries(v)])), null, 2)
);
