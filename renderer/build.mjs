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

const ff = (w, style, file) => `@font-face{font-family:'Poppins';font-weight:${w};font-style:${style};font-display:block;src:url('fonts/${file}') format('truetype')}`;
const fonts = [ff(400,'normal','Poppins-Regular.ttf'), ff(500,'normal','Poppins-Medium.ttf'),
  ff(600,'normal','Poppins-SemiBold.ttf'), ff(700,'normal','Poppins-Bold.ttf'),
  ff(400,'italic','Poppins-Italic.ttf'), ff(700,'italic','Poppins-BoldItalic.ttf')].join('');
const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>PPTX renderer</title>
<style>${fonts}
body{margin:0;background:#333;font:13px system-ui}.card{margin:16px auto;width:min(90vw,1000px);background:#fff;box-shadow:0 4px 16px rgba(0,0,0,.4)}.lbl{background:#222;color:#ddd;padding:4px 10px}.err{background:#511}pre{color:#fbb;white-space:pre-wrap;padding:8px;font-size:11px}</style>
</head><body>${cards}
<script>
function fit(){
  // Réduit le texte qui déborde. normAutofit = boîte FIXE -> réduire beaucoup (cap 20). Sinon
  // (spAutoFit/aucun) = on ne compense que l'écart de métrique de police -> réduction MODÉRÉE
  // (cap 6, plancher ~0.74x) pour éviter le pied de page réduit à ~2px. Tolérance 4px = ignore
  // un simple débordement d'interligne.
  document.querySelectorAll('foreignObject > .af').forEach(div => {
    // norm : rétrécit en hauteur ET largeur. fixed : en largeur (compense le wrapping) ; en hauteur SEULEMENT
    // si la boîte peut contenir au moins une ligne (débordement multi-lignes). Boîte plus courte qu'une ligne
    // (insets, ex. chiffre 40pt d'une pastille) -> pas de rétrécissement, le texte déborde comme PowerPoint.
    const norm = div.getAttribute('data-af') === 'norm';
    const cap = norm ? 20 : 6;
    const p = div.querySelector('[data-par]');
    const oneLine = (p ? parseFloat(getComputedStyle(p).fontSize) : 16) * 1.2;
    const canH = norm || div.clientHeight >= oneLine - 2;
    let g = 0;
    while ((((canH && div.scrollHeight > div.clientHeight + 4) || div.scrollWidth > div.clientWidth + 4)) && g++ < cap) {
      div.querySelectorAll('span').forEach(s => { s.style.fontSize = (parseFloat(getComputedStyle(s).fontSize) * 0.95) + 'px'; });
    }
  });
}
(document.fonts && document.fonts.ready ? document.fonts.ready : Promise.resolve()).then(fit);
</script>
</body></html>`;
writeFileSync(join(ROOT, 'renderer', 'viewer.html'), html);
console.log(`OK -> renderer/viewer.html (${nums.length} diapos)`);
