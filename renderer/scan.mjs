// Scanner de fonctionnalités DrawingML sur un lot de .pptx.
// Inventorie ce que les decks utilisent, en marquant ce que le moteur ne gère pas encore,
// pour prioriser les améliorations. N'affiche rien -> ne parse pas, juste des regex sur le XML.
//
// Usage : node renderer/scan.mjs "specimen/deck A"

import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';

const dir = process.argv[2] || 'specimen';
const files = readdirSync(dir).filter(f => f.toLowerCase().endsWith('.pptx'));

// presets gérés par presetPath() du moteur (sinon -> fallback rectangle)
const SUPPORTED_PRST = new Set(['rect','flowChartProcess','roundRect','ellipse','flowChartConnector',
  'triangle','rtTriangle','diamond','parallelogram','trapezoid','pentagon','hexagon','line','straightConnector1']);

const agg = {
  prst: {}, prstUnsup: {}, fonts: {}, fill: {}, effect: {}, blip: {}, bullet: {}, dash: {}, misc: {},
  slides: 0, decks: 0,
};
const add = (o, k, n = 1) => { o[k] = (o[k] || 0) + n; };
const count = (s, re) => (s.match(re) || []).length;

for (const f of files) {
  let zip;
  try { zip = unzipSync(new Uint8Array(readFileSync(join(dir, f)))); }
  catch (e) { console.error(`!! ${f} : unzip échoué (${e.message})`); continue; }
  agg.decks++;
  for (const [path, data] of Object.entries(zip)) {
    if (!/ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(path)) continue;
    const isSlide = path.includes('/slides/');
    if (isSlide) agg.slides++;
    const x = strFromU8(data);
    // géométries préréglées
    for (const m of x.matchAll(/<a:prstGeom prst="([^"]+)"/g)) { add(agg.prst, m[1]); if (!SUPPORTED_PRST.has(m[1])) add(agg.prstUnsup, m[1]); }
    // polices explicites (hors réfs de thème +mn/+mj)
    for (const m of x.matchAll(/typeface="([^"]+)"/g)) { if (!m[1].startsWith('+')) add(agg.fonts, m[1]); }
    // remplissages
    add(agg.fill, 'gradFill', count(x, /<a:gradFill/g));
    add(agg.fill, 'blipFill', count(x, /<a:blipFill|<p:blipFill/g));
    add(agg.fill, 'pattFill', count(x, /<a:pattFill/g));
    add(agg.fill, 'grad-path', count(x, /<a:path path="/g)); // dégradé radial/rectangulaire
    // effets
    add(agg.effect, 'outerShdw', count(x, /<a:outerShdw/g));
    add(agg.effect, 'innerShdw', count(x, /<a:innerShdw/g));
    add(agg.effect, 'glow', count(x, /<a:glow/g));
    add(agg.effect, 'reflection', count(x, /<a:reflection/g));
    add(agg.effect, 'softEdge', count(x, /<a:softEdge/g));
    // images
    add(agg.blip, 'svgBlip', count(x, /svgBlip/g));
    add(agg.blip, 'alphaModFix', count(x, /<a:alphaModFix/g));
    add(agg.blip, 'duotone', count(x, /<a:duotone/g));
    // puces & listes
    add(agg.bullet, 'buChar', count(x, /<a:buChar/g));
    add(agg.bullet, 'buAutoNum', count(x, /<a:buAutoNum/g));
    // traits
    add(agg.dash, 'prstDash', count(x, /<a:prstDash/g));
    // divers structurels
    add(agg.misc, 'graphicFrame(table/chart)', count(x, /<p:graphicFrame/g));
    add(agg.misc, 'grpSp', count(x, /<p:grpSp>/g));
    add(agg.misc, 'custGeom', count(x, /<a:custGeom/g));
    add(agg.misc, 'arcTo', count(x, /<a:arcTo/g));
    add(agg.misc, 'clrMapOvr-override', count(x, /overrideClrMapping/g));
    add(agg.misc, 'normAutofit', count(x, /<a:normAutofit/g));
    add(agg.misc, 'spAutoFit', count(x, /<a:spAutoFit/g));
  }
}

const show = (title, o, opt = {}) => {
  const entries = Object.entries(o).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
  if (!entries.length) return;
  console.log(`\n== ${title} ==`);
  for (const [k, v] of entries) console.log(`  ${String(v).padStart(6)}  ${k}${opt.mark && opt.mark(k) ? '   <-- NON GÉRÉ' : ''}`);
};

console.log(`Corpus : ${agg.decks} decks, ${agg.slides} diapos\n`);
show('Géométries préréglées NON GÉRÉES (fallback rectangle)', agg.prstUnsup);
show('Polices explicites (à embarquer pour fidélité)', agg.fonts);
show('Effets (ombres/lueur/reflet — non rendus)', agg.effect);
show('Remplissages', agg.fill);
show('Puces / numérotation auto (non rendues)', agg.bullet);
show('Traits pointillés', agg.dash);
show('Divers structurels', agg.misc);
show('Toutes géométries préréglées (référence)', agg.prst);
