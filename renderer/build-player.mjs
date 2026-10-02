// ═══ [SORTIE · PRÉSENTATION animée] ═══ docs/ARCHITECTURE.md (§3) · utilise pptx-render.mjs + extractTiming / extractTransition
// Lecteur PPTX animé, autonome (un seul fichier HTML), SANS PowerPoint : rendu fidèle (SVG) + timeline rejouée.
//
// Usage :
//   node renderer/build-player.mjs "chemin/deck.pptx"                   -> renderer/player-out.html
//   node renderer/build-player.mjs "chemin/deck.pptx" --out=autre.html
// Commandes du lecteur : voir l'en-tête de renderer/viewer/player.js (clic / → / ← / f / r / #3…).
//
// Pièces : lib/deck.mjs (décompression + EMF), lib/deck-data.mjs (rendu + timing + images
// dé-dupliquées), lib/html-page.mjs (page), viewer/common.js + viewer/player.js + viewer/player.css.
import { writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { fontsCssFor } from './web-fonts.mjs'; // polices du LMS (woff2 + métriques PowerPoint)
import { openDeck } from './lib/deck.mjs';
import { renderDeckData } from './lib/deck-data.mjs';
import { buildPage } from './lib/html-page.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const pptxArg = args.find((a) => !a.startsWith('--'));
const outArg = args.find((a) => a.startsWith('--out='));
const OUT = outArg ? outArg.slice(6) : join(ROOT, 'renderer', 'player-out.html');
if (!pptxArg) { console.error('usage : node renderer/build-player.mjs <deck.pptx> [--out=fichier.html]'); process.exit(1); }

let deck;
try {
  deck = openDeck(pptxArg, { workRoot: join(ROOT, 'renderer', '_work') });
} catch (e) {
  console.error(`Impossible d'ouvrir le deck : ${e.message}`);
  process.exit(1);
}

const { data, fontsSource } = renderDeckData(deck, { timing: true });

// Polices : exactement celles servies en production, limitées aux familles citées par les diapos (voir web-fonts.mjs).
// PAS de rétrécissement de texte côté navigateur (comme en production) : PowerPoint ne recalcule pas
// l'autofit à l'affichage, il applique le fontScale ENREGISTRÉ dans le fichier — que le moteur a déjà
// reporté dans le SVG. Rétrécir ici ferait diverger le rendu de l'original PowerPoint.
const html = buildPage({
  title: deck.name,
  headCss: fontsCssFor(fontsSource),
  css: ['player.css'],
  scripts: ['common.js', 'player.js'],
  body: `<div id="stage"><div id="slide"></div></div>
<div id="ui"><button id="b-prev" title="Précédent (←)" aria-label="Précédent">‹</button>
<span id="counter"></span><button id="b-next" title="Suivant (→, clic)" aria-label="Suivant">›</button>
<span id="step"></span><button id="b-restart" title="Rejouer la diapo (r)" aria-label="Rejouer la diapo">↺</button>
<button id="b-full" title="Plein écran (f)" aria-label="Plein écran">⛶</button></div>
<div id="progress"><div></div></div><div id="jumpbox" hidden></div>`,
  data,
});
writeFileSync(OUT, html);
const animated = data.slides.filter((s) => s.steps && s.steps.length).length;
const mb = (Buffer.byteLength(html) / 1e6).toFixed(1);
console.log(`OK -> ${OUT} (${data.slides.length} diapos, ${animated} animées, ${deck.hidden.size} masquées, ${data.images.length} images distinctes, ${mb} Mo)`);
