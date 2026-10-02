// ═══ Données d'un deck pour les viewers (rendu des diapos + dé-duplication des images) ═══
// Partagé par build.mjs (défilement) et build-player.mjs (présentation).
//
// API :
//   renderDeckData(deck, { timing = false, log = console.error }) -> { data, fontsSource }
//     deck : résultat de openDeck (lib/deck.mjs)
//     data : objet sérialisé dans la page et lu par renderer/viewer/*.js :
//       { name, images: [dataURI…],
//         slides: [{ n, pos, hidden, w, h, svg | error, steps?, trans? }] }
//         n = numéro de fichier (slideN.xml), pos = position 1-based dans l'ordre de présentation,
//         w / h = taille en points (viewBox du SVG),
//         svg = SVG avec jetons d'image (lib/images.mjs), error = pile si renderSlide a levé,
//         steps / trans = extractTiming(...).steps / extractTransition(...) quand timing=true.
//     fontsSource : texte de tous les SVG, à passer à fontsCssFor (web-fonts.mjs).
import { renderSlide, extractTiming, extractTransition } from '../pptx-render.mjs';
import { createImageStore } from './images.mjs';

// taille de la diapo en POINTS, lue sur le viewBox du SVG (renderSlide renvoie w/h en EMU)
export function viewBoxSize(svg) {
  const m = String(svg).match(/viewBox="\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)\s*"/);
  return m ? { w: +m[1], h: +m[2] } : { w: 960, h: 540 };
}

export function renderDeckData(deck, { timing = false, only = [], log = console.error } = {}) {
  const store = createImageStore();
  const slides = [];
  let fontsSource = '';
  deck.slides.forEach((n, i) => {
    if (only.length && !only.includes(n)) return;
    const s = { n, pos: i + 1, hidden: deck.hidden.has(n) };
    try {
      const { svg } = renderSlide(deck.pptDir, n, { embedImages: true });
      fontsSource += svg.replace(/data:[^"')\s]+/g, ''); // les polices ne se cherchent pas dans le base64
      Object.assign(s, viewBoxSize(svg), { svg: store.intern(svg) });
      if (timing) {
        s.steps = extractTiming(deck.pptDir, n).steps;
        s.trans = extractTransition(deck.pptDir, n);
      }
    } catch (e) {
      s.error = String(e && e.stack || e);
      log(`diapo ${n} :`, e && e.message);
    }
    slides.push(s);
  });
  return { data: { name: deck.name, images: store.list, slides }, fontsSource };
}
