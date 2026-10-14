// ═══ [DIAPO · PEINTURE] ═══ remplissages, dégradés, motifs, pointillés, pointes de flèche, paint().
import { resolveColorFlat, cssColor } from '../color.mjs';

// Ces fonctions renvoient une VALEUR d'attribut SVG (couleur CSS, 'none' ou 'url(#id)') et
// enregistrent au passage les définitions nécessaires (grads, patterns, lineMarkers) qui seront
// écrites dans <defs> à l'assemblage (render-slide.mjs : gradientDefs / patternDefs).
//
// fillFromSpPrFlat(rc, { 'p:spPr', … }, p:style plat) -> valeur fill. Priorité (§20.1.8, fills) :
//   remplissage EXPLICITE du spPr (noFill, grpFill, solid, grad, patt ; blipFill -> null, géré par
//   l'appelant) ; sinon référence de STYLE <p:style><a:fillRef idx> vers fillStyleLst du thème, la
//   couleur du fillRef remplaçant phClr ; sinon 'none'.
export function fillFromSpPrFlat(rc, spFlat, styleFlat) {
  const { ctx } = rc;
  const spPr = spFlat['p:spPr'];
  if (spPr?.['a:noFill'] !== undefined) return 'none';
  // <a:grpFill/> (§20.1.8.35) : la forme reprend le remplissage de son GROUPE parent (ex. régions
  // colorées « fusionnées » : le fill accent3/bg2 est sur le grpSpPr, les membres l'héritent).
  if (spPr?.['a:grpFill'] !== undefined) return ctx.grpFill ?? 'none';
  if (spPr?.['a:solidFill']) return cssColor(resolveColorFlat(spPr['a:solidFill'], ctx));
  if (spPr?.['a:gradFill']) return gradId(rc, spPr['a:gradFill']);
  if (spPr?.['a:pattFill']) return pattId(rc, spPr['a:pattFill']); // remplissage à MOTIF (hachures/points/grille)
  if (spPr?.['a:blipFill']) return null; // géré à part
  // style fillRef : idx (1-based) dans le fillStyleLst du thème, teinté par phClr
  const fr = styleFlat?.['a:fillRef'];
  if (fr) {
    const idx = +fr['@idx'];
    // idx=0 = AUCUN remplissage (la couleur du ref n'est PAS à utiliser) — ex. forme `arc`
    // ouverte (le 45° de deck-12 d.10) rendue en trait seul. Ne pas remplir.
    if (idx === 0) return 'none';
    const phName = fr['a:schemeClr']?.['@val'] ?? fr['a:srgbClr']?.['@val'];
    if (idx >= 1 && ctx.fillStyles[idx - 1]) { const p = paint(rc, ctx.fillStyles[idx - 1], phName); if (p) return p; }
    const c = resolveColorFlat(fr, ctx); if (c) return cssColor(c); // fallback : couleur du ref
  }
  return 'none';
}
// remplissage à MOTIF (a:pattFill prst + fgClr/bgClr) -> tuile <pattern> SVG. §20.1.8.24.
// Approxime par famille (points pct*, lignes horz/vert, diagonales, grille/croix) : recognizable
// sans mapper les 54 presets exacts. Absent des decks -> gain pur pour PowerPoint externes.
// rc.patterns : { id, fg, bg, prst } -> <pattern> (patternDefs)
// pattId(rc, a:pattFill plat) -> 'url(#…)' ; fgClr défaut noir, bgClr défaut blanc, prst défaut pct50
function pattId(rc, pf) {
  const { ctx, uid, patterns } = rc;
  const id = uid + 'p' + patterns.length;
  const fg = pf['a:fgClr'] ? cssColor(resolveColorFlat(pf['a:fgClr'], ctx)) : '#000';
  const bg = pf['a:bgClr'] ? cssColor(resolveColorFlat(pf['a:bgClr'], ctx)) : '#fff';
  patterns.push({ id, fg, bg, prst: pf['@prst'] || 'pct50' });
  return `url(#${id})`;
}
// rc.grads : { id, stops, ang, radial } -> <linearGradient>/<radialGradient> (gradientDefs)
// gradId(rc, a:gradFill plat) -> 'url(#…)'. a:gsLst/a:gs@pos en 1/1000 de % (/1000 -> %), couleur de
// chaque stop résolue MAINTENANT (donc avec le phClr courant : important pour les styles de thème).
// a:lin@ang (1/60000 °) -> radians ; a:path (circle/rect/shape) -> radial, rendu comme un cercle
// centré quel que soit le type (a:fillToRect ignoré). @scaled, @rotWithShape, a:tileRect ignorés.
export function gradId(rc, gradFlat) {
  const { ctx, uid, grads } = rc;
  const id = uid + 'g' + grads.length;
  const gs = gradFlat['a:gsLst']?.['a:gs'];
  // stops TRIÉS par position croissante : SVG l'exige (sinon Chrome clampe les offsets décroissants ->
  // dégradé faux, ex. fond de deck-12 d.35 rouge au lieu de vert). PowerPoint accepte l'ordre du document.
  const stops = (Array.isArray(gs)?gs:gs?[gs]:[]).map(s => ({ pos:(+s['@pos'])/1000, col: cssColor(resolveColorFlat(s, ctx)) }))
    .sort((a, b) => a.pos - b.pos);
  const ang = ((gradFlat['a:lin']?.['@ang']||0)/60000) * Math.PI/180;
  const radial = !!gradFlat['a:path']; // dégradé radial/rectangulaire
  grads.push({ id, stops, ang, radial });
  return `url(#${id})`;
}
// trait pointillé : prstDash -> stroke-dasharray (en points, relatif à l'épaisseur)
// (ST_PresetLineDashVal : longueurs tiret/espace exprimées en multiples de
// l'épaisseur, ex. dash = 4 plein / 3 vide, sysDot = 1 / 1). swp = épaisseur en pt. a:custDash non géré.
export function dashArray(val, swp) {
  if (!val || val === 'solid') return '';
  const s = swp || 1;
  return ({ dot:`${s} ${s*2}`, sysDot:`${s} ${s}`, dash:`${s*4} ${s*3}`, sysDash:`${s*3} ${s}`,
    lgDash:`${s*8} ${s*3}`, dashDot:`${s*4} ${s*3} ${s} ${s*3}`, lgDashDot:`${s*8} ${s*3} ${s} ${s*3}`,
    sysDashDot:`${s*3} ${s} ${s} ${s}`, lgDashDotDot:`${s*8} ${s*3} ${s} ${s*3} ${s} ${s*3}` }[val]) || '';
}
// Pointes de flèche -> marqueurs SVG. Clé = type + dimensions (points remplacés par « p » pour
// rester un identifiant valide) : deux traits de même type/taille partagent le même <marker>.
// pointes de flèche des lignes/connecteurs (marker SVG partagé par type ; couleur = trait via
// context-stroke ; orient=auto-start-reverse -> même marker pour début et fin).
// (rc.lineMarkers : Map clé -> <marker>)
// dimensions en multiples de l'épaisseur du trait (markerUnits=strokeWidth) selon @len/@w (sm/med/lg).
// PowerPoint « medium » ≈ 3-3,5× l'épaisseur (pas 8× — sinon pointes énormes, cf. flèches deck-12 d.28).
// Taille calibrée sur PowerPoint (export COM, traits de 1 / 2,25 / 4,5 pt) : pointe = 2 / 3 / 5 x
// l'épaisseur (sm / med / lg), avec une épaisseur PLANCHER de 2 pt (un trait de 1 pt a une pointe
// moyenne de 6 pt, pas 3) ; longueur = largeur ; le bout de la pointe tombe sur l'extrémité du trait.
// lineMarker(rc, type, nœud headEnd/tailEnd, épaisseur pt) -> id du <marker>. type (ST_LineEndType) :
// triangle, arrow, stealth, diamond, oval (arrow est dessiné comme triangle plein ; PowerPoint le
// dessine en chevron ouvert).
export function lineMarker(rc, type, end, swPt) {
  const { uid, lineMarkers } = rc;
  const SZ = { sm: 2, med: 3, lg: 5 };
  const k = swPt > 0 ? Math.max(1, 2 / swPt) : 1; // plancher 2 pt, en unités d'épaisseur de trait
  const w = (SZ[end?.['@len']] ?? 3) * k, h = (SZ[end?.['@w']] ?? 3) * k;   // len = longueur (X), w = largeur (Y)
  const key = (type + '_' + w + '_' + h).replace(/\./g, 'p');
  if (!lineMarkers.has(key)) {
    const cx = (w / 2).toFixed(2), cy = (h / 2).toFixed(2), W = w.toFixed(2), H = h.toFixed(2);
    let shape;
    if (type === 'oval') shape = `<ellipse cx="${cx}" cy="${cy}" rx="${(w/2*0.95).toFixed(2)}" ry="${(h/2*0.95).toFixed(2)}" fill="context-stroke"/>`;
    else if (type === 'diamond') shape = `<path d="M0 ${cy} L${cx} 0 L${W} ${cy} L${cx} ${H} z" fill="context-stroke"/>`;
    else if (type === 'stealth') shape = `<path d="M0 0 L${W} ${cy} L0 ${H} L${(w*0.32).toFixed(2)} ${cy} z" fill="context-stroke"/>`;
    else shape = `<path d="M0 ${(h*0.06).toFixed(2)} L${W} ${cy} L0 ${(h*0.94).toFixed(2)} z" fill="context-stroke"/>`; // triangle/arrow
    // refX = bout de la pointe (W) pour les flèches -> la pointe se termine sur l'extrémité du trait ;
    // centre pour oval/diamond (centrés sur l'extrémité). auto-start-reverse retourne le marqueur
    // en début de chemin : le même sert pour headEnd et tailEnd.
    lineMarkers.set(key, `<marker id="${uid}mk_${key}" markerWidth="${W}" markerHeight="${H}" refX="${type === 'oval' || type === 'diamond' ? cx : W}" refY="${cy}" orient="auto-start-reverse" markerUnits="strokeWidth">${shape}</marker>`);
  }
  return uid + 'mk_' + key;
}
// peinture générique d'un noeud de remplissage (solid/grad) avec substitution phClr
// phClr (« placeholder color ») : les styles du thème (fillStyleLst…) sont écrits avec
// <a:schemeClr val="phClr"/> ; la couleur réelle est fournie par la forme qui y fait référence
// (<a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef>). On la pose dans ctx.phClr le
// temps de résoudre le remplissage, puis on restaure. Ne gère que solid/grad (pas blip/patt).
export function paint(rc, fillFlat, phClrName) {
  const { ctx } = rc;
  if (!fillFlat) return null;
  const prev = ctx.phClr; ctx.phClr = phClrName || ctx.phClr;
  let out = null;
  if (fillFlat['a:solidFill']) out = cssColor(resolveColorFlat(fillFlat['a:solidFill'], ctx));
  else if (fillFlat['a:gradFill']) out = gradId(rc, fillFlat['a:gradFill']);
  ctx.phClr = prev;
  return out;
}
// dégradés définis -> <defs>
// Linéaire : vecteur (x1,y1)->(x2,y2) centré sur la boîte, d'angle ang, en unités objectBoundingBox
// (défaut SVG) -> l'angle est appliqué dans le carré unité de la boîte, donc « étiré » sur une forme
// non carrée (@scaled non distingué). 0° = gauche -> droite, 90° = haut -> bas (y vers le bas,
// comme DrawingML). Radial : centre 50 %, rayon 70 % fixes.
export function gradientDefs(grads) {
  return grads.map(g => {
  const stops = g.stops.map(s=>`<stop offset="${s.pos.toFixed(1)}%" stop-color="${s.col}"/>`).join('');
  if (g.radial) return `<radialGradient id="${g.id}" cx="50%" cy="50%" r="70%">${stops}</radialGradient>`;
  const dx = Math.cos(g.ang), dy = Math.sin(g.ang);
  return `<linearGradient id="${g.id}" x1="${((0.5-dx/2)*100).toFixed(1)}%" y1="${((0.5-dy/2)*100).toFixed(1)}%" x2="${((0.5+dx/2)*100).toFixed(1)}%" y2="${((0.5+dy/2)*100).toFixed(1)}%">${stops}</linearGradient>`;
}).join('');
}

// motifs -> tuiles <pattern> 8×8 (fond bgClr + géométrie fgClr selon la famille de preset)
// Tuile en userSpaceOnUse : 8 pt de côté quelle que soit la taille de la forme. Les regex sont
// testées dans l'ordre : le 1er motif reconnu gagne (ex. « dkUpDiag » -> diagonale montante).
export function patternDefs(patterns) {
  return patterns.map(p => {
  const s = 8, prst = p.prst; let geo = '';
  const line = (x1,y1,x2,y2) => `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${p.fg}" stroke-width="1"/>`;
  if (/^pct(\d+)/.test(prst)) { // points, densité ~ pourcentage
    // rayon ∝ √(pourcentage) (surface encrée ∝ densité), borné à [0,6 ; 3] dans une tuile de 8.
    // RegExp.$1 = groupe capturé par le test() juste au-dessus (API héritée mais fonctionnelle).
    const r = Math.min(3, Math.max(0.6, Math.sqrt(+RegExp.$1 / 100) * 3));
    geo = `<circle cx="4" cy="4" r="${r.toFixed(1)}" fill="${p.fg}"/>`;
  } else if (/(horz|Horz)$/.test(prst)) geo = line(0,4,s,4);
  else if (/(vert|Vert)$/.test(prst)) geo = line(4,0,4,s);
  else if (/(cross|Grid|trellis|weave|plaid|Check)/i.test(prst)) geo = line(0,4,s,4) + line(4,0,4,s);
  else if (/(DnDiag|dnDiag|diagBrick|shingle|wave|zigZag)/i.test(prst)) geo = line(0,0,s,s);
  else if (/(UpDiag|upDiag)/i.test(prst)) geo = line(0,s,s,0);
  else if (/(diagCross|Confetti|Dmnd|dotGrid|divot|sphere)/i.test(prst)) geo = line(0,0,s,s) + line(0,s,s,0);
  else geo = `<circle cx="4" cy="4" r="2" fill="${p.fg}"/>`; // défaut : semis de points
  return `<pattern id="${p.id}" patternUnits="userSpaceOnUse" width="${s}" height="${s}">`
    + `<rect width="${s}" height="${s}" fill="${p.bg}"/>${geo}</pattern>`;
}).join('');
}
