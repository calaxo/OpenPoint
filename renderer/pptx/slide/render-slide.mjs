// ═══ [PPTX · RENDU D'UNE DIAPO] ═══ renderSlide(pptDir, slideNo, { embedImages }) -> { svg, w, h }
//   pptDir  : dossier « ppt » du .pptx décompressé ; slideNo : N de slides/slideN.xml (≠ ordre
//             d'affichage : voir slideOrder() dans ../parts.mjs) ;
//   embedImages (défaut true) : images en data URI (sinon file:///) ;
//   w, h    : taille de diapo en EMU (le viewBox du svg, lui, est en points).
//
// ÉTAT DU RENDU : un objet `rc` (render context) créé ici et passé EN PREMIER ARGUMENT à toutes les
// fonctions de rendu des modules voisins (walk.mjs, shape.mjs, text.mjs…). Champs :
//   pptDir, slideNo, embed, uid (préfixe d'ID des <defs>, 's<N>_'), W, H (taille de diapo, EMU),
//   ctx      : contexte de thème (theme.mjs) + styles hérités ; MUTÉ pendant le rendu (phClr, grpFill),
//   phGeom / mphGeom : placeholders du layout / du master (parts.mjs),
//   body     : fragments SVG dans l'ordre de dessin (ordre z),
//   curRels  : rels du niveau en cours de parcours (master, layout ou diapo) pour résoudre les r:embed,
//   skipPh   : true pendant le parcours du master/layout (leurs placeholders ne sont pas dessinés),
//   spidPfx  : préfixe de data-spid ('m' master, 'l' layout, '' diapo),
//   grads, patterns, clips, filters, lineMarkers : définitions accumulées, écrites dans <defs> à la fin.
// renderSlide n'est pas réentrant (état de module dans fonts.mjs, ctx muté) : un appel à la fois.
import { readFileSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { P, PF, cachedParse, tagOf, first, serialize } from '../xml.mjs';
import { u } from '../units.mjs';
import { loadContext } from '../theme.mjs';
import { slideRels, relsOf, imageHref, placeholderGeom } from '../parts.mjs';
import { setMinorFont } from '../fonts.mjs';
import { loadMasterStyles, loadDefaultTextStyle } from '../text-styles.mjs';
import { walk, IDENT } from './walk.mjs';
import { paint, gradientDefs, patternDefs } from './paint.mjs';

export function renderSlide(pptDir, slideNo, opts = {}) {
  const embed = opts.embedImages !== false;
  // préfixe d'ID UNIQUE par diapo : clip/gradient/marqueur/filtre. Sans ça, plusieurs diapos dans UNE
  // page (viewer.html) partagent des IDs (clip0, g0…) qui collisionnent -> mauvais clip/dégradé (ex. image
  // circulaire rendue en petit rectangle dans le viewer mais OK dans le player qui isole les diapos).
  const uid = 's' + slideNo + '_';
  const rels = slideRels(pptDir, slideNo);
  // --- résout layout + masque + THÈME (le thème dépend du master : cf. loadContext) ---
  // norm : résout une cible de rels relativement au DOSSIER de la part source (« ../slideMasters/x »)
  const norm = (base, tgt) => join(dirname(base), tgt);
  // NB : le layout est joint à « slides/ » (et non via norm) car sa cible part de slides/ (« ../slideLayouts/… »)
  const layoutTgt = Object.values(rels).find(t => t && t.includes('slideLayout'));
  const layoutPath = layoutTgt ? join(pptDir, 'slides', layoutTgt) : null;
  const layoutRels = layoutPath ? relsOf(layoutPath) : {};
  const masterTgt = Object.values(layoutRels).find(t => t && t.includes('slideMaster'));
  const masterPath = (layoutPath && masterTgt) ? norm(layoutPath, masterTgt) : null;
  const masterRels = masterPath ? relsOf(masterPath) : {};
  const themeTgt = Object.values(masterRels).find(t => t && t.includes('/theme/'));
  const themePath = (masterPath && themeTgt) ? norm(masterPath, themeTgt) : null;
  const ctx = loadContext(pptDir, themePath);
  setMinorFont(ctx.minorFont); // pour la mesure des cellules de tableau (cellTextHeight)
  const { geom: phGeom, mgeom: mphGeom } = placeholderGeom(pptDir, slideNo, rels);

  // La diapo elle-même n'est PAS mise en cache (propre à cet appel) : parse ordonné direct.
  const slidePath = join(pptDir, 'slides', `slide${slideNo}.xml`);
  const ordered = P.parse(readFileSync(slidePath, 'utf8'));
  const sld = ordered.find((n) => tagOf(n) === 'p:sld');
  const cSld = first(sld, 'p:cSld');
  const spTree = first(cSld, 'p:spTree');
  const { W, H } = slideSize(pptDir);

  // clrMap + txStyles du master, puis defaultTextStyle de la présentation (bas de la cascade de texte)
  loadMasterStyles(ctx, masterPath);
  loadDefaultTextStyle(ctx, pptDir);

  const rc = {
    pptDir, slideNo, embed, uid, W, H, ctx, phGeom, mphGeom,
    body: [], curRels: rels, skipPh: false, spidPfx: '',
    grads: [], patterns: [], clips: [], filters: [], lineMarkers: new Map(),
  };

  // --- composition : fond unique, puis master (déco) -> layout (déco) -> diapo (tout) ---
  const masterPart = masterPath ? partOf(masterPath, 'p:sldMaster') : null;
  const layoutPart = layoutPath ? partOf(layoutPath, 'p:sldLayout') : null;
  // UN SEUL fond effectif (précédence diapo > layout > masque), dessiné DERRIÈRE toutes les formes.
  // Le fond n'est PAS cumulatif : un <p:bg> de layout/diapo REMPLACE celui du niveau au-dessus. Le
  // dessiner à chaque niveau repeindrait par-dessus les formes du niveau précédent (logos du masque
  // masqués par le fond du layout). curRels doit pointer le bon niveau (un fond peut être une image).
  rc.curRels = rels;        const slideBg  = bgRect(rc, cSld);
  rc.curRels = layoutRels;  const layoutBg = layoutPart ? bgRect(rc, layoutPart.cSld) : '';
  rc.curRels = masterRels;  const masterBg = masterPart ? bgRect(rc, masterPart.cSld) : '';
  rc.body.push(slideBg || layoutBg || masterBg || '');
  // master / layout : placeholders NON dessinés (gabarits), spid préfixés pour ne pas entrer en
  // collision avec ceux de la diapo (que cible la timeline d'animation)
  if (masterPart) { rc.curRels = masterRels; rc.skipPh = true; rc.spidPfx = 'm'; walk(rc, masterPart.spTree, IDENT); }
  if (layoutPart) { rc.curRels = layoutRels; rc.skipPh = true; rc.spidPfx = 'l'; walk(rc, layoutPart.spTree, IDENT); }
  // la diapo : placeholders DESSINÉS, spid bruts
  rc.curRels = rels; rc.skipPh = false; rc.spidPfx = '';
  walk(rc, spTree, IDENT);

  // Assemblage : viewBox en POINTS, largeur 100 % (le conteneur fixe l'échelle) ; <defs> regroupe
  // dégradés, motifs, filtres, clips et marqueurs ; un rectangle BLANC sous tout (repli si aucun fond
  // n'est défini) ; puis body dans l'ordre z. w/h renvoyés en EMU.
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${u(W)} ${u(H)}" width="100%" style="display:block">`
    + `<defs>${gradientDefs(rc.grads)}${patternDefs(rc.patterns)}${rc.filters.join('')}${rc.clips.join('')}${[...rc.lineMarkers.values()].join('')}</defs>`
    + `<rect width="${u(W)}" height="${u(H)}" fill="#ffffff"/>`
    + rc.body.join('') + `</svg>`;
  return { svg, w: W, h: H };
}

// slideSize(pptDir) -> { W, H } en EMU, lus dans presentation.xml (p:sldSz ; gère 16:9, 4:3…).
// Défaut = 16:9 « grand écran » de PowerPoint : 12 192 000 × 6 858 000 EMU = 960 × 540 pt.
export function slideSize(pptDir) {
  let W = 12192000, H = 6858000;
  try {
    const pres = cachedParse(join(pptDir, 'presentation.xml'), 'flat');
    const sz = pres['p:presentation']?.['p:sldSz'];
    if (sz) { W = +sz['@cx'] || W; H = +sz['@cy'] || H; }
  } catch {}
  return { W, H };
}

// partOf(chemin, balise racine) -> { spTree, cSld } d'un master / layout (arbre ORDONNÉ, en cache,
// donc en lecture seule), ou null si le fichier manque.
function partOf(xmlPath, rootTag) {
  if (!xmlPath || !existsSync(xmlPath)) return null;
  const o = cachedParse(xmlPath, 'ord');
  const root = o.find(n => tagOf(n) === rootTag);
  const cS = first(root, 'p:cSld');
  return { spTree: first(cS, 'p:spTree'), cSld: cS };
}

// bgRect(rc, cSld) -> fragment SVG du fond de CE niveau (diapo, layout ou master), '' s'il n'en
// définit pas. p:bg contient soit p:bgPr (remplissage explicite), soit p:bgRef (référence au thème).
// Dépend de rc.curRels (le fond image est un r:embed du niveau concerné) : l'appelant le positionne.
function bgRect(rc, cS) {
  const { ctx, curRels, pptDir, embed, W, H } = rc;
  const bg = cS && first(cS, 'p:bg');
  if (!bg) return '';
  const bgFlat = PF.parse(`<root xmlns:a="a" xmlns:p="p">${serialize(bg)}</root>`).root['p:bg'];
  // bgPr : fill direct (solid/grad/blip)
  const bgPr = bgFlat['p:bgPr'];
  if (bgPr) {
    if (bgPr['a:blipFill']) {
      const embedId = bgPr['a:blipFill']['a:blip']?.['@r:embed'] ?? bgPr['a:blipFill']['a:blip']?.['@embed'];
      const href = curRels[embedId] ? imageHref(pptDir, curRels[embedId], embed) : null;
      if (href) return `<image x="0" y="0" width="${u(W)}" height="${u(H)}" href="${href}" preserveAspectRatio="xMidYMid slice"/>`;
    }
    const p = paint(rc, bgPr);
    if (p) return `<rect width="${u(W)}" height="${u(H)}" fill="${p}"/>`;
  }
  // bgRef : idx>=1001 -> theme bgFillStyleLst[idx-1001], phClr = couleur donnée
  // (Convention d'index des style matrix refs : 0 = aucun, 1..999 = fillStyleLst, 1001+ =
  // bgFillStyleLst. Ici un idx < 1000 prend bgFillStyles[0] — approximation.)
  // Seuls solid/grad sont rendus via paint : un fond de thème en blipFill/pattFill est ignoré.
  const bgRef = bgFlat['p:bgRef'];
  if (bgRef) {
    const idx = +bgRef['@idx'];
    const phName = bgRef['a:schemeClr']?.['@val'] ?? bgRef['a:srgbClr']?.['@val'];
    const styleFill = ctx.bgFillStyles[idx >= 1000 ? idx - 1001 : 0];
    const p = paint(rc, styleFill, phName);
    if (p) return `<rect width="${u(W)}" height="${u(H)}" fill="${p}"/>`;
  }
  return '';
}
