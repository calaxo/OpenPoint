// ═══ [DIAPO · PARCOURS] ═══ walk() : parcours du spTree (ordre z), aiguillage par type, groupes.
// Repère : tf = transformation « repère enfant -> repère diapo » : x' = ox + (x - cx0) · sx (idem y).
import { flatOf, tagOf, at, kids, first, deepFind, esc } from '../xml.mjs';
import { u } from '../units.mjs';
import { resolveColorFlat, cssColor } from '../color.mjs';
import { gradId } from './paint.mjs';
import { renderSp } from './shape.mjs';
import { renderPic } from './picture.mjs';
import { renderGraphicFrame } from './frame.mjs';

// IDENT = identité (spTree de 1er niveau : les coordonnées sont déjà celles de la diapo). Les groupes
// (renderGroup) et SmartArt (renderDiagram) en fabriquent d'autres ; inGroup:true signale un membre de groupe.
export const IDENT = { ox:0, oy:0, sx:1, sy:1, cx0:0, cy0:0 };

// walk(rc, conteneur, tf) : rend chaque enfant d'un spTree / grpSp dans l'ordre du document (= ordre z).
// Les enfants non-formes (p:nvGrpSpPr, p:grpSpPr, extLst…) tombent dans aucun cas de walkChild.
export function walk(rc, node, tf) {
  if (!node) return;
  for (const child of kids(node)) walkChild(rc, child, tf);
}
// walkChild : aiguillage par type d'élément. Les fonctions de rendu RENVOIENT leur fragment (poussé
// dans body ici), sauf renderGroup et renderDiagram qui poussent elles-mêmes (elles doivent
// envelopper / insérer plusieurs fragments).
function walkChild(rc, child, tf) {
  const { skipPh, body } = rc;
  const t = tagOf(child);
  // MCE (Part 3) : une forme peut être enveloppée dans <mc:AlternateContent> (ex. formules OMML
  // sous mc:Choice Requires="a14"). Sans le gérer, la forme + son texte sont INVISIBLES.
  // On rend la branche Fallback (repli sûr) si présente, sinon Choice.
  if (t === 'mc:AlternateContent') {
    // On préfère mc:Choice (contenu AUTHORÉ : texte + formules ; on rend au moins le texte
    // normal, le math OMML est ignoré). Le Fallback des formules est souvent une forme VIDE
    // -> il donnerait une boîte sans texte. Repli sur Fallback si pas de Choice.
    const branch = first(child, 'mc:Choice') || first(child, 'mc:Fallback');
    if (branch) for (const c of kids(branch)) walkChild(rc, c, tf);
    return;
  }
  // forme masquée (p:cNvPr@hidden, volet Sélection de PowerPoint) : jamais dessinée. Pour un groupe,
  // deepFind trouve le cNvPr du GROUPE (nvGrpSpPr précède les membres) : le groupe entier est masqué.
  if (['1', 'true'].includes(at(deepFind(child, 'p:cNvPr'))['@hidden'])) return;
  // masque/layout : on NE rend PAS les placeholders (la diapo en hérite fill/style).
  if (t === 'p:sp') { if (skipPh && deepFind(child, 'p:ph')) return; body.push(renderSp(rc, child, tf)); }
  // image ESPACE RÉSERVÉ (<p:ph>) à l'intérieur d'un GROUPE : structure que l'interface PowerPoint
  // n'autorise pas ; PowerPoint en dessine les autres membres mais PAS l'image (vérifié par export
  // COM, 022_04 d.65) -> la rendre ajoutait une 2e copie décalée de l'illustration.
  else if (t === 'p:pic') { if (tf?.inGroup && deepFind(first(child, 'p:nvPicPr'), 'p:ph')) return; body.push(renderPic(rc, child, tf)); }
  else if (t === 'p:cxnSp') body.push(renderSp(rc, child, tf));
  else if (t === 'p:graphicFrame') body.push(renderGraphicFrame(rc, child, tf));
  else if (t === 'p:grpSp') renderGroup(rc, child, tf);
}
// getXfrm(spPr ORDONNÉ) -> { x, y, cx, cy (EMU), rot (1/60000 °), flipH, flipV } ou null si la forme
// n'a pas de a:xfrm complet (cas d'un placeholder qui hérite sa position du layout). §20.1.7.6
export function getXfrm(spPrOrdered) {
  const xf = first(spPrOrdered, 'a:xfrm');
  if (!xf) return null;
  const off = first(xf, 'a:off'), ext = first(xf, 'a:ext');
  if (!off || !ext) return null;
  return { x:+at(off)['@x'], y:+at(off)['@y'], cx:+at(ext)['@cx'], cy:+at(ext)['@cy'],
           rot:+(at(xf)['@rot']||0), flipH:at(xf)['@flipH']==='1', flipV:at(xf)['@flipV']==='1' };
}
// applyTf(boîte, tf) : passe une boîte du repère enfant (d'un groupe) au repère de la diapo. Seules
// position et taille sont transformées : rot / flip restent ceux de la forme (la rotation / le
// miroir d'un GROUPE est appliqué par un <g transform> englobant, voir renderGroup).
export function applyTf(g, tf) { return { x: tf.ox + (g.x - tf.cx0)*tf.sx, y: tf.oy + (g.y - tf.cy0)*tf.sy, cx: g.cx*tf.sx, cy: g.cy*tf.sy, rot:g.rot, flipH:g.flipH, flipV:g.flipV }; }

// renderGroup(rc, p:grpSp, tf) : un groupe DrawingML a son propre repère enfant. a:xfrm du grpSpPr
// donne où le groupe est posé (off/ext, dans le repère PARENT) et quel rectangle de son repère
// enfant cela représente (chOff/chExt). Donc, pour un membre : x_parent = off.x + (x - chOff.x) ·
// ext.cx / chExt.cx (d'où childTf). Les groupes imbriqués composent ces transformations.
// Pousse dans body un <g data-spid=id-du-groupe> contenant les fragments des membres.
function renderGroup(rc, grp, tf) {
  const { ctx, body, spidPfx } = rc;
  const gPr = first(grp, 'p:grpSpPr');
  const xf = gPr && first(gPr, 'a:xfrm');
  let childTf = tf, wrap = '';
  if (xf) {
    const off=first(xf,'a:off'),ext=first(xf,'a:ext'),chOff=first(xf,'a:chOff'),chExt=first(xf,'a:chExt');
    const g = applyTf({ x:+at(off)['@x'],y:+at(off)['@y'],cx:+at(ext)['@cx'],cy:+at(ext)['@cy'] }, tf);
    // chExt nul / absent -> sx/sy vaudraient Infinity ou NaN : on retombe sur 1 (aucune mise à
    // l'échelle), comme pour le mapping des SmartArt (cf. renderDiagram). Cas réel : un groupe
    // purement horizontal (chExt cy=0) contenant un connecteur droit -> sinon son y devient NaN.
    const cE = { cx:+at(chExt)['@cx'], cy:+at(chExt)['@cy'] };
    childTf = { ox:g.x, oy:g.y, cx0:+at(chOff)['@x'], cy0:+at(chOff)['@y'],
                sx: cE.cx ? g.cx/cE.cx : 1, sy: cE.cy ? g.cy/cE.cy : 1, inGroup: true };
    // rotation / miroir du GROUPE (les membres sont placés dans la boîte NON tournée, puis tout
    // le <g> tourne autour du centre du groupe). Sans ça, un groupe tourné rend à plat
    // (ex. manches à air inclinées de VENT d.10).
    const rot = +at(xf)['@rot'] || 0, fH = at(xf)['@flipH']==='1', fV = at(xf)['@flipV']==='1';
    const ccx = u(g.x + g.cx/2), ccy = u(g.y + g.cy/2), parts = [];
    if (rot) parts.push(`rotate(${(rot/60000).toFixed(3)} ${ccx.toFixed(1)} ${ccy.toFixed(1)})`);
    // miroir autour du centre : translate(2·cx) scale(-1) ⇔ x -> 2cx - x (les membres étant en
    // coordonnées absolues de diapo, le centre du groupe sert de pivot)
    if (fH || fV) parts.push(`translate(${(fH?2*ccx:0).toFixed(1)} ${(fV?2*ccy:0).toFixed(1)}) scale(${fH?-1:1} ${fV?-1:1})`);
    wrap = parts.join(' ');
  }
  // remplissage du GROUPE (hérité par les membres <a:grpFill/>). On le résout et on le propage
  // pendant le walk des enfants, puis on restaure (gère les groupes imbriqués).
  const prevGrpFill = ctx.grpFill;
  if (gPr) {
    const gf = flatOf(gPr)['p:grpSpPr'];
    if (gf?.['a:solidFill']) ctx.grpFill = cssColor(resolveColorFlat(gf['a:solidFill'], ctx));
    else if (gf?.['a:gradFill']) ctx.grpFill = gradId(rc, gf['a:gradFill']);
    else if (gf?.['a:noFill'] !== undefined) ctx.grpFill = 'none';
    // sinon (pas de fill explicite ou grpFill imbriqué) : on garde le fill du groupe parent
  }
  // enveloppe les membres dans un <g data-spid=groupId> : le groupe s'anime EN BLOC
  // (une animation ciblant le groupe s'applique à tous ses membres, comme PowerPoint).
  // technique : on note la longueur de body, on rend les membres (qui s'y ajoutent), puis on
  // retire ces fragments (splice) pour les réinsérer à l'intérieur du <g>.
  const start = body.length;
  walk(rc, grp, childTf);
  ctx.grpFill = prevGrpFill;
  const inner = body.splice(start).join('');
  const gid = at(deepFind(grp, 'p:cNvPr'))['@id'] || '';
  body.push(`<g data-spid="${spidPfx}${esc(gid)}"${wrap?` transform="${wrap}"`:''}>${inner}</g>`);
}
