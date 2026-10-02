// ═══ [SORTIE · viewer DÉFILEMENT] ═══ docs/ARCHITECTURE.md (§3) · utilise pptx-render.mjs · état FINAL (pas d'animation)
// Rend un PPTX en un viewer HTML autonome (un seul fichier), SANS PowerPoint.
//
// Usage :
//   node renderer/build.mjs "C:\chemin\deck.pptx"          -> renderer/viewer.html
//   node renderer/build.mjs "C:\chemin\deck.pptx" 1 5 8    -> seulement ces diapos (numéros de fichier slideN.xml)
//   node renderer/build.mjs "deck.pptx" --out=autre.html   -> autre fichier de sortie
//   node renderer/build.mjs                                -> deck de démo (specimen/extracted/ppt), si présent
//
// Pièces : lib/deck.mjs (décompression), lib/deck-data.mjs (rendu + images dé-dupliquées),
// lib/html-page.mjs (page), viewer/common.js + viewer/scroll-viewer.js + viewer/viewer.css (runtime).
import { writeFileSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { slideOrder } from './pptx-render.mjs';
import { fontsCssFor } from './web-fonts.mjs'; // polices du LMS (woff2 + métriques PowerPoint)
import { openDeck, hiddenSlides } from './lib/deck.mjs';
import { renderDeckData } from './lib/deck-data.mjs';
import { buildPage } from './lib/html-page.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const pptxArg = args.find((a) => a.toLowerCase().endsWith('.pptx'));
const only = args.filter((a) => /^\d+$/.test(a)).map(Number);
const outArg = args.find((a) => a.startsWith('--out='));
const OUT = outArg ? outArg.slice(6) : join(ROOT, 'renderer', 'viewer.html');

let deck;
try {
  if (pptxArg) {
    deck = openDeck(pptxArg, { workRoot: join(ROOT, 'renderer', '_work') });
    console.log(`Décompressé : ${pptxArg}\n           -> ${deck.workDir}`);
  } else {
    // deck de démo déjà décompressé (dossier local, non versionné)
    const pptDir = join(ROOT, 'specimen', 'extracted', 'ppt');
    if (!existsSync(join(pptDir, 'slides'))) throw new Error('aucun .pptx donné et pas de deck de démo (specimen/extracted/ppt)');
    const slides = slideOrder(pptDir);
    deck = { pptDir, name: 'specimen', slides, hidden: hiddenSlides(pptDir, slides) };
  }
} catch (e) {
  console.error(`Impossible d'ouvrir le deck : ${e.message}`);
  process.exit(1);
}

const { data, fontsSource } = renderDeckData(deck, { only });

// Polices : exactement celles servies en production, limitées aux familles que les diapos citent (voir web-fonts.mjs).
// PAS de rétrécissement de texte côté navigateur (comme en production) : PowerPoint ne recalcule pas
// l'autofit à l'affichage, il applique le fontScale ENREGISTRÉ dans le fichier — que le moteur a déjà
// reporté dans le SVG. Rétrécir ici ferait diverger le rendu de l'original PowerPoint.
const html = buildPage({
  title: deck.name,
  headCss: fontsCssFor(fontsSource),
  css: ['viewer.css'],
  scripts: ['common.js', 'scroll-viewer.js'],
  body: `<header id="bar"><span id="deck-name"></span><span id="count"></span>
<span class="nav"><button id="prev" title="Diapo précédente (k)" aria-label="Diapo précédente">‹</button>
<input id="jump" type="number" min="1" inputmode="numeric" aria-label="Aller à la diapo" title="Aller à la diapo (n° + Entrée)">
<button id="next" title="Diapo suivante (j)" aria-label="Diapo suivante">›</button></span></header>
<main id="deck"></main>`,
  data,
});
writeFileSync(OUT, html);
const mb = (Buffer.byteLength(html) / 1e6).toFixed(1);
console.log(`OK -> ${OUT} (${data.slides.length} diapos, ${data.images.length} images distinctes, ${mb} Mo)`);
