// Contrôle batch : rend chaque diapo de chaque .pptx d'un dossier, sans produire de HTML.
// Rapporte par deck : nb diapos, erreurs, diapos "maigres" (rendu suspect = peu d'éléments).
//
// Usage : node renderer/check.mjs "specimen/deck A"

import { renderSlide } from './pptx-render.mjs';
import { readFileSync, readdirSync, mkdirSync, existsSync, rmSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { unzipSync } from 'fflate';

const dir = process.argv[2] || 'specimen';
const files = readdirSync(dir).filter(f => f.toLowerCase().endsWith('.pptx') && !f.startsWith('~$'));
const WORK = join('renderer', '_work', '_check');

let totErr = 0, totSlides = 0;
for (const f of files) {
  let zip; try { zip = unzipSync(new Uint8Array(readFileSync(join(dir, f)))); }
  catch { console.log(`${f}\n  !! zip invalide`); continue; }
  if (existsSync(WORK)) rmSync(WORK, { recursive: true, force: true });
  for (const [p, data] of Object.entries(zip)) {
    if (p.endsWith('/')) continue;
    const out = join(WORK, p); mkdirSync(dirname(out), { recursive: true }); writeFileSync(out, Buffer.from(data));
  }
  const ppt = join(WORK, 'ppt');
  const nums = readdirSync(join(ppt, 'slides')).filter(x => /^slide\d+\.xml$/.test(x)).map(x => +x.match(/\d+/)[0]).sort((a,b)=>a-b);
  const errs = [], thin = [];
  for (const n of nums) {
    try {
      const { svg } = renderSlide(ppt, n, { embedImages: false });
      const shapes = (svg.match(/<path|<image|<foreignObject/g) || []).length;
      if (shapes < 3) thin.push(`${n}(${shapes})`);
    } catch (e) { errs.push(`${n}: ${e.message}`); }
  }
  totErr += errs.length; totSlides += nums.length;
  console.log(`${f}  (${nums.length} diapos)`);
  if (errs.length) console.log('  ERREURS: ' + errs.join(' | '));
  if (thin.length) console.log('  maigres: ' + thin.join(' '));
  if (!errs.length && !thin.length) console.log('  OK');
}
console.log(`\nTOTAL : ${totSlides} diapos, ${totErr} erreurs`);
if (existsSync(WORK)) rmSync(WORK, { recursive: true, force: true });
