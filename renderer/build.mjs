// ═══ [SORTIE · viewer DÉFILEMENT] ═══ docs/ARCHITECTURE.md (§3) · utilise pptx-render.mjs · état FINAL (pas d'animation)
// Rend un PPTX (ou le deck de démo) en un viewer HTML autonome, SANS PowerPoint.
//
// Usage :
//   node renderer/build.mjs                       -> deck de démo (specimen/extracted/ppt)
//   node renderer/build.mjs "C:\chemin\deck.pptx" -> n'importe quel .pptx (décompressé auto)
//   node renderer/build.mjs "C:\chemin\deck.pptx" 1 5 8   -> seulement ces diapos
//
// Sortie : renderer/viewer.html (+ dossier de travail renderer/_work/<nom>/)

import { renderSlide } from './pptx-render.mjs';
import { fontsCssFor } from './web-fonts.mjs'; // polices du LMS (woff2 + métriques PowerPoint)
import { writeFileSync, readFileSync, readdirSync, mkdirSync, existsSync, rmSync } from 'node:fs';
import { join, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const pptxArg = args.find(a => a.toLowerCase().endsWith('.pptx'));
const only = args.filter(a => /^\d+$/.test(a)).map(Number);

// --- résout le dossier ppt/ (décompresse le .pptx si fourni) ---
let PPT;
if (pptxArg) {
  const name = basename(pptxArg).replace(/\.pptx$/i, '').replace(/[^\w.-]/g, '_');
  const workDir = join(ROOT, 'renderer', '_work', name);
  if (existsSync(workDir)) rmSync(workDir, { recursive: true, force: true });
  const zip = unzipSync(new Uint8Array(readFileSync(pptxArg)));
  for (const [path, data] of Object.entries(zip)) {
    if (path.endsWith('/')) continue;
    const out = join(workDir, path);
    mkdirSync(dirname(out), { recursive: true });
    writeFileSync(out, Buffer.from(data));
  }
  PPT = join(workDir, 'ppt');
  console.log(`Décompressé : ${pptxArg}\n           -> ${workDir}`);
} else {
  PPT = join(ROOT, 'specimen', 'extracted', 'ppt');
}

if (!existsSync(join(PPT, 'slides'))) {
  console.error(`Introuvable : ${join(PPT, 'slides')} — le .pptx est-il valide ?`);
  process.exit(1);
}

const nums = readdirSync(join(PPT, 'slides'))
  .filter(f => /^slide\d+\.xml$/.test(f)).map(f => +f.match(/\d+/)[0]).sort((a, b) => a - b)
  .filter(n => !only.length || only.includes(n));

let cards = '';
for (const n of nums) {
  try {
    const { svg } = renderSlide(PPT, n, { embedImages: true });
    cards += `<div class="card"><div class="lbl">diapo ${n}</div>${svg}</div>`;
  } catch (e) {
    cards += `<div class="card err"><div class="lbl">diapo ${n} — ERREUR</div><pre>${e.stack}</pre></div>`;
    console.error(`slide ${n}:`, e.message);
  }
}

// Polices : exactement celles servies en production, limitées aux familles que les diapos citent (voir web-fonts.mjs).
const fonts = fontsCssFor(cards);
// PAS de rétrécissement de texte côté navigateur (comme en production) : PowerPoint ne recalcule pas
// l'autofit à l'affichage, il applique le fontScale ENREGISTRÉ dans le fichier — que le moteur a déjà
// reporté dans le SVG. Rétrécir ici faisait diverger nos captures du rendu attendu.
const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>PPTX renderer</title>
<style>${fonts}
body{margin:0;background:#333;font:13px system-ui}.card{margin:16px auto;width:min(90vw,1000px);background:#fff;box-shadow:0 4px 16px rgba(0,0,0,.4)}.lbl{background:#222;color:#ddd;padding:4px 10px}.err{background:#511}pre{color:#fbb;white-space:pre-wrap;padding:8px;font-size:11px}</style>
</head><body>${cards}
</body></html>`;
writeFileSync(join(ROOT, 'renderer', 'viewer.html'), html);
console.log(`OK -> renderer/viewer.html (${nums.length} diapos)`);
