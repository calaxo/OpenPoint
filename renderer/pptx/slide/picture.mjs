// ═══ [DIAPO · IMAGE] ═══ renderPic : image (p:pic).
import { flatOf, at, first, deepFind, deepFindSuffix, esc } from '../xml.mjs';
import { u } from '../units.mjs';
import { presetPath, custGeomPath } from '../geometry.mjs';
import { imageHref, phInherit } from '../parts.mjs';
import { getXfrm, applyTf } from './walk.mjs';
import { effectFilter, imageFilter } from './effects.mjs';

// renderPic(rc, nœud ORDONNÉ, tf) -> <g data-spid transform clip filter><image …/></g>.
// Étapes : boîte (xfrm propre ou du placeholder) -> choix de la source (SVG > raster) -> forme de
// découpe (prstGeom / custGeom) -> mode d'ajustement (stretch / contain / cover) -> recadrage
// srcRect -> filtres couleur (imageFilter) et effets (effectFilter).
// Aussi appelée par renderGraphicFrame pour l'image de secours d'un objet OLE.
export function renderPic(rc, pic, tf) {
  const { curRels, pptDir, embed, uid, clips, spidPfx, phGeom } = rc;
  // placeholder image : géométrie (forme de découpe) héritée du layout
  const ph = deepFind(pic, 'p:ph');
  const { inh } = ph ? phInherit(phGeom, null, ph) : {};
  const spPrO = first(pic, 'p:spPr');
  let g = getXfrm(spPrO);
  if (!g && inh?.xfrm) { const xf=inh.xfrm, o=xf['a:off'], e=xf['a:ext']; if (o&&e) g={ x:+o['@x'],y:+o['@y'],cx:+e['@cx'],cy:+e['@cy'],rot:+(xf['@rot']||0) }; }
  if (!g) return '';
  const m = applyTf(g, tf);
  // priorité au SVG moderne (asvg:svgBlip), sinon le blip raster (png/jpeg/wdp)
  const blip = deepFind(pic, 'a:blip');
  const svgB = deepFindSuffix(pic, 'svgBlip');
  // ordre d'essai des rId ; le 1er dont imageHref aboutit gagne (un .wdp ou un EMF sans PNG
  // pré-converti échoue -> on passe au suivant)
  const candidates = [at(svgB)['@r:embed'], at(blip)['@r:embed'], at(blip)['@embed']].filter(Boolean);
  let href = null;
  for (const id of candidates) { const tgt = curRels[id]; if (tgt) { href = imageHref(pptDir, tgt, embed); if (href) break; } }
  if (!href) return ''; // image non exploitable (ex. .wdp seul) -> ne rien afficher
  // forme de découpe : géométrie propre du pic, sinon celle héritée du placeholder de layout
  const picFlat = flatOf(pic)['p:pic'];
  let geomFlat = picFlat['p:spPr'] || {};
  if (!geomFlat['a:prstGeom'] && !geomFlat['a:custGeom'] && inh?.spPr) geomFlat = inh.spPr;
  const prst = geomFlat['a:prstGeom']?.['@prst'];
  const custFlat = geomFlat['a:custGeom'];
  const custO = custFlat ? deepFind(pic, 'a:custGeom') : null;
  // <a:stretch><a:fillRect/> (mode par défaut d'une image) = ÉTIRER pour remplir la forme, sans
  // letterbox (preserveAspectRatio="none") — sinon des bandes transparentes laissent voir ce qui est
  // dessous (ex. TEMP d.19 : le badge sous l'image transparaissait dans la bande du haut). Sans
  // stretch (tile/rare) : contain.
  const hasStretch = picFlat['p:blipFill']?.['a:stretch'] !== undefined;
  let clipAttr = '', fit = hasStretch ? 'none' : 'xMidYMid meet';
  // forme de découpe NON rectangulaire -> clipPath (en coordonnées locales de l'image, comme le
  // chemin : le clip est posé sur le <g> déjà translaté/tourné)
  if (custO || (prst && prst !== 'rect')) {
    let d;
    if (custO) d = custGeomPath(custO, u(m.cx), u(m.cy));
    else { const adj={}; const gd=geomFlat['a:prstGeom']?.['a:avLst']?.['a:gd'];
      for (const x of (Array.isArray(gd)?gd:gd?[gd]:[])) adj[x['@name']]=x['@fmla'];
      d = presetPath(prst, u(m.cx), u(m.cy), adj); }
    const cid = uid + 'clip' + clips.length;
    clips.push(`<clipPath id="${cid}"><path d="${d}"/></clipPath>`);
    clipAttr = ` clip-path="url(#${cid})"`;
    fit = 'xMidYMid slice'; // forme non-rect : remplir + clipper (cover)
  }
  const rot = m.rot ? ` rotate(${(m.rot/60000).toFixed(2)} ${u(m.cx/2)} ${u(m.cy/2)})` : '';
  // miroir flipH/flipV (renderPic l'ignorait -> images de flèches courbes flip+rot dans le mauvais sens,
  // ex. VENT d.14). Appliqué APRÈS rot dans la chaîne = AVANT la rotation sur le contenu (SVG right-to-
  // left), comme PowerPoint (miroir en repère local puis rotation). Le clip suit le même repère.
  const flip = (m.flipH || m.flipV) ? ` translate(${u(m.flipH?m.cx:0)} ${u(m.flipV?m.cy:0)}) scale(${m.flipH?-1:1} ${m.flipV?-1:1})` : '';
  const imgF = imageFilter(rc, picFlat['p:blipFill']?.['a:blip']); // grayscl/duotone/alphaModFix
  const effF = effectFilter(rc, picFlat['p:spPr'], picFlat['p:style']); // glow/ombre/softEdge (+ effectRef)
  const bw = u(m.cx), bh = u(m.cy);
  // recadrage source a:srcRect (§20.1.8.55) : l/t/r/b = fraction à RETIRER de chaque bord
  // (millièmes de %). On agrandit l'image et on la décale pour que la région visible
  // remplisse la boîte, puis on clippe (mapping exact -> preserveAspectRatio="none").
  let ix = 0, iy = 0, iw = bw, ih = bh, par = fit;
  const sr = picFlat['p:blipFill']?.['a:srcRect'];
  // srcRect (§20.1.8.55) : l/t/r/b = fraction (millièmes de %) retirée de chaque bord de la SOURCE
  // avant remplissage. POSITIF = recadrage (zoom) ; NÉGATIF = marge hors-image (letterbox) → sert à
  // faire tenir une photo paysage dans une forme carrée SANS déformation (ex. photo d'intro : b=-54251
  // rend la région source ~carrée, l'image occupe le haut ~65 % de la boîte, non déformée). On mappe
  // la région source (agrandie+décalée) exactement sur la boîte -> par="none".
  if (sr && (sr['@l'] || sr['@t'] || sr['@r'] || sr['@b'])) {
    const l=(+sr['@l']||0)/100000, t=(+sr['@t']||0)/100000, r=(+sr['@r']||0)/100000, b=(+sr['@b']||0)/100000;
    // vw, vh = fraction de l'image source qui reste visible (bornée > 0 pour éviter /0)
    const vw = Math.max(1e-3, 1-l-r), vh = Math.max(1e-3, 1-t-b);
    // image entière dessinée à la taille boîte / fraction visible, décalée de la part rognée à
    // gauche / en haut : la région visible tombe exactement sur la boîte (0,0)-(bw,bh)
    iw = bw/vw; ih = bh/vh; ix = -l*iw; iy = -t*ih; par = 'none';
    // sans forme de découpe, un clip rectangulaire de la boîte masque ce qui dépasse
    if (!clipAttr) { const cid=uid+'clip'+clips.length;
      clips.push(`<clipPath id="${cid}"><rect x="0" y="0" width="${bw.toFixed(1)}" height="${bh.toFixed(1)}"/></clipPath>`);
      clipAttr = ` clip-path="url(#${cid})"`; }
  }
  const spid = at(deepFind(pic, 'p:cNvPr'))['@id'] || '';
  return `<g data-spid="${spidPfx}${esc(spid)}" transform="translate(${u(m.x)} ${u(m.y)})${rot}${flip}"${clipAttr}${effF}>`
    + `<image x="${ix.toFixed(1)}" y="${iy.toFixed(1)}" width="${iw.toFixed(1)}" height="${ih.toFixed(1)}" href="${href}" preserveAspectRatio="${par}"${imgF.attr}${imgF.opacity}/></g>`;
}
