// Contrôle batch : rend chaque diapo de chaque .pptx d'un dossier, sans produire de HTML.
// Rapporte par deck : nb diapos, erreurs, diapos "maigres" (rendu suspect = peu d'éléments).
//
// Usage : node renderer/check.mjs "specimen/deck A"

import { renderSlide } from './pptx-render.mjs';
import { readdirSync, existsSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { openDeck } from './lib/deck.mjs';

const dir = process.argv[2] || 'specimen';
const files = readdirSync(dir).filter(f => f.toLowerCase().endsWith('.pptx') && !f.startsWith('~$'));
const WORK = join('renderer', '_work', '_check');

let totErr = 0, totSlides = 0;
for (const f of files) {
  // décompression sûre (lib/deck.mjs) ; pas de conversion EMF : les images ne sont pas embarquées ici
  let deck; try { deck = openDeck(join(dir, f), { workDir: WORK, metafiles: false }); }
  catch (e) { console.log(`${f}\n  !! zip invalide (${e.message})`); continue; }
  const ppt = deck.pptDir, nums = deck.slides; // ordre de présentation
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
