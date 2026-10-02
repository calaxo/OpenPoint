// ═══ [DIAPO · FORME] ═══ renderSp : forme (p:sp) ou connecteur (p:cxnSp).
import { flatOf, at, first, deepFind, esc } from '../xml.mjs';
import { u } from '../units.mjs';
import { presetGeomParts, presetTextRect } from '../../preset-geom.mjs';
import { custGeomParts } from '../geometry.mjs';
import { imageHref, phInherit } from '../parts.mjs';
import { resolveColorFlat, cssColor } from '../color.mjs';
import { getXfrm, applyTf } from './walk.mjs';
import { fillFromSpPrFlat, gradId, dashArray, lineMarker } from './paint.mjs';
import { effectFilter } from './effects.mjs';
import { renderText } from './text.mjs';

// renderSp(rc, nœud ORDONNÉ, tf) -> fragment SVG :
//   <g data-spid transform="translate(x y) rotate(θ cx cy)" filter>   <- position + rotation
//     <g transform="translate(…) scale(±1 ±1)">                       <- miroir (flipH/flipV)
//       <path remplissage (+ trait)/> [<path trait seul/>] [image découpée]
//     </g>
//     <foreignObject texte/>                                         <- NON miroité (lisible)
//   </g>
// Étapes : parse plat -> héritage placeholder -> boîte (xfrm) -> géométrie -> fill -> trait ->
// image de remplissage -> effets -> texte. Toutes les coordonnées internes sont en pt, relatives au
// coin haut-gauche de la forme.
export function renderSp(rc, sp, tf) {
  const { ctx, curRels, pptDir, embed, uid, clips, spidPfx, phGeom, mphGeom } = rc;
  // version flat du même noeud pour lire fills/geom/txBody
  const flat = flatOf(sp);
  const spFlat = flat['p:sp'] || flat['p:cxnSp'] || flat;
  const spPrO = first(sp, 'p:spPr');
  // placeholder : héritage depuis le layout (géométrie + fill + styles de texte)
  const ph = deepFind(sp, 'p:ph');
  // inh = placeholder du LAYOUT, inhM = celui du MASTER (maillon entre le layout et p:txStyles)
  const { inh, inhM } = ph ? phInherit(phGeom, mphGeom, ph) : {};
  // boîte : a:xfrm de la forme, sinon celui du placeholder de layout (sans flip hérité) ; aucune
  // boîte -> forme non dessinée
  let g = getXfrm(spPrO);
  if (!g && inh?.xfrm) { const xf=inh.xfrm, off=xf['a:off'], ext=xf['a:ext']; if (off&&ext) g = { x:+off['@x'],y:+off['@y'],cx:+ext['@cx'],cy:+ext['@cy'],rot:+(xf['@rot']||0) }; }
  if (!g) return '';
  const m = applyTf(g, tf);
  let geomFlat = spFlat['p:spPr'] || {};
  // géométrie héritée du placeholder de layout si la diapo n'en définit pas
  if (!geomFlat['a:prstGeom'] && !geomFlat['a:custGeom'] && inh?.spPr) geomFlat = inh.spPr;
  const prst = geomFlat['a:prstGeom']?.['@prst'];
  // custGeom : il faut le nœud ORDONNÉ (ordre des commandes), cherché dans la forme elle-même. Si
  // la custGeom vient du LAYOUT (geomFlat = inh.spPr), deepFind(sp) ne la trouve pas -> custO null
  // -> branche preset avec prst undefined -> rectangle.
  const custO = geomFlat['a:custGeom'] ? deepFind(sp, 'a:custGeom') : null;
  // d = chemin de remplissage, dStroke = chemin de trait, textRect = rectangle de texte du preset
  // (EMU, relatif à la forme ; les formes custGeom n'en ont pas ici -> boîte entière)
  let d, dStroke = '', textRect = null;
  if (custO) { const p = custGeomParts(custO, u(m.cx), u(m.cy)); d = p.fill; dStroke = p.stroke; }
  else {
    // ajustements (a:avLst/a:gd name="adj" fmla="val 16667") -> { adj: 'val 16667' } ; les valeurs
    // absentes prennent le défaut du preset (géré par preset-geom.mjs)
    const adj = {}; const gd = geomFlat['a:prstGeom']?.['a:avLst']?.['a:gd'];
    for (const x of (Array.isArray(gd)?gd:gd?[gd]:[])) adj[x['@name']] = x['@fmla'];
    const p = presetGeomParts(prst || 'rect', u(m.cx), u(m.cy), adj);
    if (p) { d = p.fill; dStroke = p.stroke; } else { d = `M0 0 H${u(m.cx)} V${u(m.cy)} H0 Z`; dStroke = d; }
    textRect = presetTextRect(prst || 'rect', m.cx, m.cy, adj); // en EMU, comme les insets
  }
  let fill = fillFromSpPrFlat(rc, spFlat, spFlat['p:style']);
  // fill hérité du placeholder de layout (fond de carte, cercle de badge)
  // NB : fillFromSpPrFlat renvoie aussi 'none' pour un <a:noFill/> EXPLICITE de la diapo ; dans ce
  // cas le fill du layout est quand même repris (le test ne distingue pas les deux situations).
  if (fill === 'none' && inh) fill = fillFromSpPrFlat(rc, { 'p:spPr': inh.spPr, 'p:style': inh.style }, inh.style);
  // Trait (a:ln, §20.1.2.2.24) : celui du spPr retenu (diapo, ou layout si la géométrie en vient),
  // sinon celui du placeholder de layout. Deux cas :
  //   1. a:ln présent sans a:noFill -> épaisseur/couleur/pointillés/jointures/pointes explicites,
  //      les manques étant complétés par le style de ligne du thème (lnRef) ;
  //   2. pas de a:ln et forme NON remplie -> trait entièrement issu du thème (lnRef).
  const ln = geomFlat['a:ln'] ?? inh?.spPr?.['a:ln'];
  let stroke = 'none', sw = 0, dash = '', markers = '', joinCap = '';
  if (ln && ln['a:noFill'] === undefined) {
    // épaisseur : @w, sinon celle du style de ligne du thème désigné par <a:lnRef idx> (défaut 1 pt)
    const lnRefIdx = +(spFlat['p:style']?.['a:lnRef']?.['@idx'] || 0);
    sw = Number(ln['@w']) || Number(ctx.lnStyles?.[lnRefIdx - 1]?.['a:ln']?.['@w']) || 12700;
    // trait en DÉGRADÉ (cartes à bordure estompée) : sans ça, repli sur la couleur lnRef (souvent
    // un accent assombri à 15 %) -> bordure quasi invisible.
    stroke = ln['a:solidFill'] ? cssColor(resolveColorFlat(ln['a:solidFill'], ctx))
           : ln['a:gradFill'] ? gradId(rc, ln['a:gradFill'])
           : spFlat['p:style']?.['a:lnRef'] ? cssColor(resolveColorFlat(spFlat['p:style']['a:lnRef'], ctx)) : 'none';
    dash = dashArray(ln['a:prstDash']?.['@val'], u(sw));
    // jointure (§20.1.8.{35,42,49} round/bevel/miter) + extrémité (@cap rnd/sq/flat) du trait.
    // On n'émet QUE les valeurs NON-défaut SVG (défaut = miter/butt) : aucun bruit, blast radius =
    // uniquement les traits réellement à coins/bouts arrondis. Additif -> zéro régression ailleurs.
    const join = ln['a:round'] !== undefined ? 'round' : ln['a:bevel'] !== undefined ? 'bevel' : '';
    const cap = ln['@cap'] === 'rnd' ? 'round' : ln['@cap'] === 'sq' ? 'square' : '';
    if (join) joinCap += ` stroke-linejoin="${join}"`;
    if (cap) joinCap += ` stroke-linecap="${cap}"`;
    // pointes de flèche (§20.1.8.24/38) : headEnd = début du chemin (marker-start),
    // tailEnd = fin (marker-end). Marqueurs SVG partagés, couleur = trait (context-stroke).
    const endType = (e) => e && e['@type'] && e['@type'] !== 'none' ? e['@type'] : null;
    const he = endType(ln['a:headEnd']), te = endType(ln['a:tailEnd']);
    if (he) markers += ` marker-start="url(#${lineMarker(rc, he, ln['a:headEnd'], u(sw))})"`;
    if (te) markers += ` marker-end="url(#${lineMarker(rc, te, ln['a:tailEnd'], u(sw))})"`;
  } else if (!ln && fill === 'none') {
    // Pas de <a:ln> mais une ligne thémée via <p:style><a:lnRef idx>=1> (ex. le connecteur en angle =
    // carré d'angle droit de VENT d.26, invisible sinon). Scopé aux formes NON remplies (connecteurs/
    // lignes) pour ne pas ajouter de bordure parasite aux cartes/boîtes pleines (pas de régression).
    const lr = spFlat['p:style']?.['a:lnRef'];
    const idx = lr ? +lr['@idx'] : 0;
    if (lr && idx >= 1) {
      const themeLn = ctx.lnStyles[idx - 1]?.['a:ln'] || {};
      sw = Number(themeLn['@w']) || 12700;
      stroke = cssColor(resolveColorFlat(lr, ctx));
      dash = dashArray(themeLn['a:prstDash']?.['@val'], u(sw));
      const endType2 = (e) => e && e['@type'] && e['@type'] !== 'none' ? e['@type'] : null;
      const he = endType2(themeLn['a:headEnd']), te = endType2(themeLn['a:tailEnd']);
      if (he) markers += ` marker-start="url(#${lineMarker(rc, he, themeLn['a:headEnd'], u(sw))})"`;
      if (te) markers += ` marker-end="url(#${lineMarker(rc, te, themeLn['a:tailEnd'], u(sw))})"`;
    }
  }
  // forme remplie par une IMAGE (blipFill) : image découpée à la forme
  // (clip = chemin de remplissage d ; « slice » = couvrir la boîte en rognant, sans déformer ;
  // srcRect / stretch / tile de ce blipFill ne sont pas lus ici, contrairement à renderPic)
  let picLayer = '';
  const bf = spFlat['p:spPr']?.['a:blipFill'] || (fill === 'none' && inh?.spPr?.['a:blipFill']);
  if (bf?.['a:blip']) {
    const id = bf['a:blip']['@r:embed'] || bf['a:blip']['@embed'];
    const tgt = id && curRels[id];
    const href = tgt && imageHref(pptDir, tgt, embed);
    if (href) {
      const cid = uid + 'clip' + clips.length;
      clips.push(`<clipPath id="${cid}" clipPathUnits="userSpaceOnUse"><path d="${d}"/></clipPath>`);
      picLayer = `<g clip-path="url(#${cid})"><image x="0" y="0" width="${u(m.cx)}" height="${u(m.cy)}" href="${href}" preserveAspectRatio="xMidYMid slice"/></g>`;
      fill = 'none';
    }
  }
  const filt = effectFilter(rc, spFlat['p:spPr'], spFlat['p:style'] || inh?.style);
  // rotation autour du CENTRE de la forme (règle a:xfrm@rot : rotation autour du centre de la boîte,
  // sens horaire) ; miroir : translate(largeur) scale(-1 1) = retournement dans la boîte.
  // Le miroir est dans un <g> INTÉRIEUR qui ne contient pas le texte : le texte n'est jamais miroité
  // (conforme à PowerPoint pour flipH ; pour flipV, PowerPoint retourne le texte de 180°, pas ici).
  const rot = m.rot ? ` rotate(${(m.rot/60000).toFixed(2)} ${u(m.cx/2)} ${u(m.cy/2)})` : '';
  const sc = (m.flipH||m.flipV) ? ` translate(${u(m.flipH?m.cx:0)} ${u(m.flipV?m.cy:0)}) scale(${m.flipH?-1:1} ${m.flipV?-1:1})` : '';
  // phKind : title/body seulement pour un VRAI placeholder ; une zone de texte LIBRE (sans p:ph)
  // = 'none' -> n'hérite PAS du style body du master (sinon puces « • » et couleurs parasites) ;
  // elle s'appuie sur defaultTextStyle + couleur de thème.
  const phType = at(ph)['@type'];
  const phKind = !ph ? 'none' : (phType === 'title' || phType === 'ctrTitle') ? 'title' : 'body';
  // renderText reçoit w/h en EMU (m.cx, m.cy), le txBody plat ET ordonné, et les éléments hérités
  // du layout (lstStyle, bodyPr, style) pour la cascade de styles
  const txt = renderText(rc, spFlat['p:txBody'], m.cx, m.cy, phKind, inh?.lstStyle, first(sp, 'p:txBody'), inh?.bodyPr, spFlat['p:style'] || inh?.style, textRect, inhM?.lstStyle);
  const spid = at(deepFind(sp, 'p:cNvPr'))['@id'] || '';
  const strokeAttr = stroke !== 'none' ? `stroke="${stroke}" stroke-width="${u(sw)}"${dash?` stroke-dasharray="${dash}"`:''}${joinCap}${markers}` : '';
  // Tracé sur le même chemin que le remplissage (cas courant) OU chemin de tracé distinct (formes
  // arc/chord… : le remplissage est un secteur, le tracé n'est que la courbe -> pas de rayons parasites).
  // NB : si dStroke est VIDE (aucun sous-chemin à tracer, cf. evalGeomParts), sameGeom est vrai et le
  // trait est posé sur le chemin de remplissage.
  const sameGeom = !dStroke || dStroke === d;
  const paths = `<path d="${d}" fill="${fill}" ${sameGeom ? strokeAttr : ''}/>`
    + (!sameGeom && strokeAttr ? `<path d="${dStroke}" fill="none" ${strokeAttr}/>` : '');
  return `<g data-spid="${spidPfx}${esc(spid)}" transform="translate(${u(m.x)} ${u(m.y)})${rot}"${filt}>`
    + `<g transform="${sc.trim()||'translate(0 0)'}">`
    + paths
    + picLayer
    + `</g>` + txt + `</g>`;
}
