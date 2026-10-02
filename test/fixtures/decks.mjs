// Decks d'exemple SYNTHÉTIQUES (OOXML écrit à la main, aucun fichier privé) pour tester le moteur
// (renderer/pptx-render.mjs) et les viewers (build.mjs / build-player.mjs).
//
// API :
//   buildDecks()          -> { [nomDeck]: { [cheminZip]: string | Uint8Array } }   (cartes de fichiers en mémoire)
//   writeDeck(files, dir) -> écrit l'arbre décompressé dans dir (dir/ppt/... = le « pptDir » du moteur)
//   zipDeck(files)        -> Uint8Array d'un .pptx (fflate.zipSync)
//
// Conventions : coordonnées des helpers en POINTS (1 pt = 12700 EMU), diapo 16:9 = 720 x 405 pt.
// Les ids de formes (cNvPr@id) sont ceux ciblés par la timeline (p:spTgt@spid).

import { zipSync, strToU8 } from 'fflate';
import { PNG } from 'pngjs';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';

// ───────────────────────── constantes et helpers de base ─────────────────────────
const SLIDE_W = 720, SLIDE_H = 405;
const E = (pt) => Math.round(pt * 12700);
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = (t) => `${NS_R}/${t}`;
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NSDECL = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const rels = (list) => XML + `<Relationships xmlns="${NS_PKG_REL}">`
  + list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type.startsWith('http') ? type : REL(type)}" Target="${target}"/>`).join('')
  + '</Relationships>';

// ───────────────────────── remplissages / traits ─────────────────────────
const clr = (c) => /^[0-9a-fA-F]{6}$/.test(c) ? `<a:srgbClr val="${c}"/>` : `<a:schemeClr val="${c}"/>`;
const solid = (c) => `<a:solidFill>${clr(c)}</a:solidFill>`;
const grad = (c1, c2, ang = 5400000) =>
  `<a:gradFill><a:gsLst><a:gs pos="0">${clr(c1)}</a:gs><a:gs pos="100000">${clr(c2)}</a:gs></a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
const patt = (prst, fg, bg) => `<a:pattFill prst="${prst}"><a:fgClr>${clr(fg)}</a:fgClr><a:bgClr>${clr(bg)}</a:bgClr></a:pattFill>`;
const ln = (c, wPt = 1, { dash, head, tail } = {}) =>
  `<a:ln w="${E(wPt)}">${solid(c)}${dash ? `<a:prstDash val="${dash}"/>` : ''}`
  + `${head ? `<a:headEnd type="${head}" w="med" len="med"/>` : ''}${tail ? `<a:tailEnd type="${tail}" w="med" len="med"/>` : ''}</a:ln>`;
const shadow = `<a:effectLst><a:outerShdw blurRad="63500" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`;

// ───────────────────────── texte ─────────────────────────
// run(texte, {b,i,u,sz,color,font}) ; para(runsOuTexte, {algn, bu, lvl, pPrInner, marL, indent})
const run = (t, o = {}) => {
  const at = `lang="fr-FR"${o.sz ? ` sz="${o.sz * 100}"` : ''}${o.b ? ' b="1"' : ''}${o.i ? ' i="1"' : ''}${o.u ? ' u="sng"' : ''} dirty="0"`;
  const inner = (o.color ? solid(o.color) : '') + (o.font ? `<a:latin typeface="${o.font}"/>` : '');
  return `<a:r><a:rPr ${at}>${inner}</a:rPr><a:t>${esc(t)}</a:t></a:r>`;
};
const para = (content, o = {}) => {
  const runs = Array.isArray(content) ? content.join('') : (content.startsWith('<') ? content : run(content, o.r));
  const at = `${o.algn ? ` algn="${o.algn}"` : ''}${o.lvl ? ` lvl="${o.lvl}"` : ''}${o.marL !== undefined ? ` marL="${o.marL}"` : ''}${o.indent !== undefined ? ` indent="${o.indent}"` : ''}`;
  return `<a:p><a:pPr${at}>${o.pPr ?? ''}</a:pPr>${runs}</a:p>`;
};
const txBody = (paras, bodyPr = '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>') =>
  `<p:txBody>${bodyPr}<a:lstStyle/>${paras.join('')}</p:txBody>`;

// ───────────────────────── formes ─────────────────────────
// sp({id,name,x,y,w,h,prst,geom,fill,line,effect,rot,flipH,flipV,paras,bodyPr,ph,txBox,hidden})
const xfrm = (o, tag = 'a:xfrm') =>
  `<${tag}${o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : ''}${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}>`
  + `<a:off x="${E(o.x)}" y="${E(o.y)}"/><a:ext cx="${E(o.w)}" cy="${E(o.h)}"/></${tag}>`;
const sp = (o) => {
  const geom = o.geom ?? `<a:prstGeom prst="${o.prst ?? 'rect'}"><a:avLst>${o.av ?? ''}</a:avLst></a:prstGeom>`;
  const ph = o.ph ? `<p:ph ${o.ph}/>` : '';
  const spPr = o.ph && o.x === undefined
    ? '<p:spPr/>'
    : `<p:spPr>${xfrm(o)}${geom}${o.fill ?? ''}${o.line ?? ''}${o.effect ?? ''}</p:spPr>`;
  const body = o.paras ? txBody(o.paras, o.bodyPr) : (o.ph ? '' : txBody([`<a:p><a:endParaRPr lang="fr-FR"/></a:p>`]));
  return `<p:sp><p:nvSpPr><p:cNvPr id="${o.id}" name="${esc(o.name ?? 'Forme ' + o.id)}"${o.hidden ? ' hidden="1"' : ''}/><p:cNvSpPr${o.txBox ? ' txBox="1"' : ''}/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>${spPr}${body}</p:sp>`;
};
const cxn = (o) => `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${o.id}" name="${esc(o.name ?? 'Connecteur ' + o.id)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>`
  + `<p:spPr>${xfrm(o)}<a:prstGeom prst="${o.prst ?? 'line'}"><a:avLst/></a:prstGeom>${o.line}</p:spPr></p:cxnSp>`;
// groupe : (x,y,w,h) = cadre sur la diapo, (cx,cy,cw,ch) = repère enfant (par défaut identique)
const grp = (o, children) => {
  const c = o.child ?? { x: o.x, y: o.y, w: o.w, h: o.h };
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${o.id}" name="${esc(o.name ?? 'Groupe ' + o.id)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
    + `<p:grpSpPr><a:xfrm${o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : ''}><a:off x="${E(o.x)}" y="${E(o.y)}"/><a:ext cx="${E(o.w)}" cy="${E(o.h)}"/>`
    + `<a:chOff x="${E(c.x)}" y="${E(c.y)}"/><a:chExt cx="${E(c.w)}" cy="${E(c.h)}"/></a:xfrm></p:grpSpPr>${children.join('')}</p:grpSp>`;
};
const pic = (o) => `<p:pic><p:nvPicPr><p:cNvPr id="${o.id}" name="${esc(o.name ?? 'Image ' + o.id)}" descr="image synthétique"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="${o.rId}"/>${o.srcRect ? `<a:srcRect ${o.srcRect}/>` : ''}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`
  + `<p:spPr>${xfrm(o)}<a:prstGeom prst="${o.prst ?? 'rect'}"><a:avLst/></a:prstGeom>${o.line ?? ''}</p:spPr></p:pic>`;
const graphicFrame = (o, dataXml) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${o.id}" name="${esc(o.name ?? 'Cadre ' + o.id)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
  + `${xfrm(o, 'p:xfrm')}<a:graphic>${dataXml}</a:graphic></p:graphicFrame>`;

// ───────────────────────── diapo ─────────────────────────
// slide({shapes, bg, transition, timing, layout}) -> { xml, layout }
// sldAttrs : attributs supplémentaires de <p:sld> (ex. ' showMasterSp="0"', ' show="0"')
const slide = ({ shapes, bg = '', transition = '', timing = '', layout = 1, clrMapOvr = '<a:masterClrMapping/>', rels: extra = [], name, sldAttrs = '' }) => ({
  layout, rels: extra, name,
  xml: XML + `<p:sld ${NSDECL}${sldAttrs}><p:cSld>${name ? '' : ''}${bg ? `<p:bg>${bg}</p:bg>` : ''}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
    + `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld>`
    + `<p:clrMapOvr>${clrMapOvr}</p:clrMapOvr>${transition}${timing}</p:sld>`,
});
// petit titre de diapo (zone texte) pour repérer visuellement les diapos
const label = (id, text, y = 8) => sp({ id, name: 'Étiquette', x: 20, y, w: 680, h: 30, txBox: true, fill: '<a:noFill/>',
  paras: [para(text, { r: { sz: 16, b: true, color: '404040' } })] });

// ───────────────────────── timing (p:timing) ─────────────────────────
// Un effet : fx(spid, kind, {par, delay, dur, node}). kind -> preset + comportements PowerPoint.
//   node = clickEffect (défaut pour le 1er d'un clic) | withEffect | afterEffect
const KINDS = {
  fade:      { id: 10, cls: 'entr', sub: 0, filter: 'fade' },
  wipe:      { id: 22, cls: 'entr', sub: 1, filter: 'wipe(down)' },
  zoom:      { id: 53, cls: 'entr', sub: 16, filter: 'fade', zoom: true },
  exitFade:  { id: 10, cls: 'exit', sub: 0, filter: 'fade' },
  grow:      { id: 6,  cls: 'emph', sub: 0, scale: true },
  motion:    { id: 0,  cls: 'path', sub: 0, path: 'M 0 0 L 0.25 0.1 E' },
  spin:      { id: 8,  cls: 'emph', sub: 0, rot: 21600000 },
  flyLeft:   { id: 2,  cls: 'entr', sub: 8, fly: true },
};
let _ctn = 0;
const ctnId = () => ++_ctn;
const tgt = (spid, par) => `<p:tgtEl><p:spTgt spid="${spid}">${par !== undefined ? `<p:txEl><p:pRg st="${par}" end="${par}"/></p:txEl>` : ''}</p:spTgt></p:tgtEl>`;
const bhvrSet = (spid, par, vis, dur = 1) =>
  `<p:set><p:cBhvr><p:cTn id="${ctnId()}" dur="${dur}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt(spid, par)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${vis}"/></p:to></p:set>`;
const bhvrEffect = (spid, par, transition, filter, dur) =>
  `<p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr><p:cTn id="${ctnId()}" dur="${dur}"/>${tgt(spid, par)}</p:cBhvr></p:animEffect>`;
const bhvrAnim = (spid, par, attr, from, to, dur) =>
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="${ctnId()}" dur="${dur}" fill="hold"/>${tgt(spid, par)}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr>`
  + `<p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
// -> { xml, end } (end = délai propre + durée, pour calculer le délai d'un groupe « après »)
const fx = (spid, kind, { par, delay = 0, dur = 500, node = 'clickEffect' } = {}) => {
  const k = KINDS[kind];
  let beh = '';
  if (k.cls === 'entr') {
    beh = bhvrSet(spid, par, 'visible');
    if (k.zoom) beh += bhvrAnim(spid, par, 'ppt_w', '0', '#ppt_w', dur) + bhvrAnim(spid, par, 'ppt_h', '0', '#ppt_h', dur);
    // « Venir de la gauche » : PowerPoint anime ppt_x depuis hors diapo (pas d'animEffect)
    if (k.fly) beh += bhvrAnim(spid, par, 'ppt_x', '0-#ppt_w/2', '#ppt_x', dur) + bhvrAnim(spid, par, 'ppt_y', '#ppt_y', '#ppt_y', dur);
    else beh += bhvrEffect(spid, par, 'in', k.filter, dur);
  } else if (k.cls === 'exit') {
    beh = bhvrEffect(spid, par, 'out', k.filter, dur) + bhvrSet(spid, par, 'hidden').replace('delay="0"', `delay="${dur}"`);
  } else if (k.scale) {
    beh = `<p:animScale><p:cBhvr><p:cTn id="${ctnId()}" dur="${dur}" fill="hold"/>${tgt(spid, par)}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`;
  } else if (k.path) {
    beh = `<p:animMotion origin="layout" path="${k.path}" pathEditMode="relative"><p:cBhvr><p:cTn id="${ctnId()}" dur="${dur}" fill="hold"/>${tgt(spid, par)}<p:rCtr x="0.125" y="0.05"/></p:cBhvr></p:animMotion>`;
  } else if (k.rot) {
    beh = `<p:animRot by="${k.rot}"><p:cBhvr><p:cTn id="${ctnId()}" dur="${dur}" fill="hold"/>${tgt(spid, par)}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
  }
  const xml = `<p:par><p:cTn id="${ctnId()}" presetID="${k.id}" presetClass="${k.cls}" presetSubtype="${k.sub}" fill="hold" grpId="0" nodeType="${node}">`
    + `<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${beh}</p:childTnLst></p:cTn></p:par>`;
  return { xml, end: delay + dur };
};
// timing(clicks, bldP) : clicks = [ [ [effets du groupe 1 (avec)], [groupe 2 (après)], ... ], ... ]
// Chaque groupe = un p:par intermédiaire dont le délai = fin du groupe précédent (comme PowerPoint).
// opts.autoFirst : la 1re étape démarre seule (1er effet « après le précédent ») ; PowerPoint ajoute alors
// une condition onBegin sur le conteneur de l'étape.
const timing = (clicks, bldList = '', opts = {}) => {
  _ctn = 2; // 1 = tmRoot, 2 = mainSeq
  const clickXml = clicks.map((groups, ci) => {
    const outer = ctnId();
    let t = 0;
    const inner = groups.map((effs) => {
      const gid = ctnId();
      const x = `<p:par><p:cTn id="${gid}" fill="hold"><p:stCondLst><p:cond delay="${t}"/></p:stCondLst><p:childTnLst>${effs.map((e) => e.xml).join('')}</p:childTnLst></p:cTn></p:par>`;
      t += Math.max(...effs.map((e) => e.end));
      return x;
    }).join('');
    return `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>${opts.autoFirst && ci === 0 ? '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : ''}</p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
  }).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>`
    + `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clickXml}</p:childTnLst></p:cTn>`
    + `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>`
    + `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`
    + `</p:childTnLst></p:cTn></p:par></p:tnLst>${bldList ? `<p:bldLst>${bldList}</p:bldLst>` : ''}</p:timing>`;
};

// transitions
const trFade = (dur) => `<p:transition spd="med"${dur ? ` xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="${dur}"` : ''}><p:fade/></p:transition>`;
const trPush = (dir) => `<p:transition spd="med"><p:push dir="${dir}"/></p:transition>`;
const trMorph = () => `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">`
  + `<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">`
  + `<p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1500"><p159:morph option="byObject"/></p:transition></mc:Choice>`
  + `<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`;

// ───────────────────────── thème / master / layouts ─────────────────────────
const themeXml = XML + `<a:theme xmlns:a="${NS_A}" name="Synthétique"><a:themeElements>`
  + `<a:clrScheme name="Synth"><a:dk1><a:srgbClr val="1F1F1F"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="22304A"/></a:dk2><a:lt2><a:srgbClr val="EEF1F6"/></a:lt2>`
  + `<a:accent1><a:srgbClr val="2E75B6"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="70AD47"/></a:accent3>`
  + `<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="7030A0"/></a:accent5><a:accent6><a:srgbClr val="C00000"/></a:accent6>`
  + `<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>`
  + `<a:fontScheme name="Synth"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`
  + `<a:fmtScheme name="Synth"><a:fillStyleLst>${solid('phClr')}${solid('phClr')}${solid('phClr')}</a:fillStyleLst>`
  + `<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>`
  + `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>`
  + `<a:bgFillStyleLst>${solid('phClr')}${solid('phClr')}${solid('phClr')}</a:bgFillStyleLst></a:fmtScheme>`
  + `</a:themeElements></a:theme>`.replace(/<a:solidFill><a:srgbClr val="phClr"\/>/g, '<a:solidFill><a:schemeClr val="phClr"/>');

const lvl = (n, sz, extra = '') => `<a:lvl${n}pPr marL="${(n - 1) * 342900 + 228600 * (n > 0 ? 1 : 0)}" indent="-228600" algn="l"><a:buFont typeface="Arial"/><a:buChar char="•"/>${extra}<a:defRPr sz="${sz}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl${n}pPr>`;
const masterXml = XML + `<p:sldMaster ${NSDECL}><p:cSld><p:bg><p:bgPr>${solid('bg1')}<a:effectLst/></p:bgPr></p:bg><p:spTree>`
  + `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`
  + sp({ id: 2, name: 'Titre maître', ph: 'type="title"', x: 36, y: 24, w: 648, h: 70, paras: [para('Titre')] })
  + sp({ id: 3, name: 'Corps maître', ph: 'type="body" idx="1"', x: 36, y: 110, w: 648, h: 260, paras: [para('Texte')] })
  + sp({ id: 4, name: 'Numéro maître', ph: 'type="sldNum" sz="quarter" idx="4"', x: 600, y: 375, w: 84, h: 22,
    paras: [`<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="fr-FR" sz="1000"/><a:t>‹#›</a:t></a:fld></a:p>`] })
  // décoration du master (non placeholder) : visible sur toutes les diapos sauf showMasterSp="0"
  + sp({ id: 5, name: 'Bandeau maître', x: 0, y: 400, w: 720, h: 5, fill: solid('accent1') })
  + `</p:spTree></p:cSld>`
  + `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`
  + `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>`
  + `<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200" b="1"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>`
  + `<p:bodyStyle>${lvl(1, 2000)}${lvl(2, 1800)}${lvl(3, 1600)}</p:bodyStyle>`
  + `<p:otherStyle><a:defPPr><a:defRPr lang="fr-FR"/></a:defPPr><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;

// layout 1 « Titre et contenu » (clair) ; layout 2 « Sombre » : clrMapOvr bg1->dk1, tx1->lt1 + fond bg1
const layoutXml = (name, dark) => XML + `<p:sldLayout ${NSDECL} type="obj" preserve="1"><p:cSld name="${name}">`
  + (dark ? `<p:bg><p:bgPr>${solid('bg1')}<a:effectLst/></p:bgPr></p:bg>` : '')
  + `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`
  + sp({ id: 2, name: 'Titre', ph: 'type="title"', x: 36, y: 24, w: 648, h: 70, paras: [para('Titre')] })
  + sp({ id: 3, name: 'Contenu', ph: 'idx="1"', x: 36, y: 110, w: 648, h: 250, paras: [para('Texte')] })
  + `</p:spTree></p:cSld>`
  + (dark ? `<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>`
    : `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`) + `</p:sldLayout>`;

// ───────────────────────── assemblage d'un deck ─────────────────────────
// order : numéros de diapos dans l'ordre de PRÉSENTATION (par défaut 1..n)
const assemble = (slides, { order, extraFiles = {}, extraTypes = '' } = {}) => {
  const n = slides.length;
  const ord = order ?? slides.map((_, i) => i + 1);
  const f = {};
  f['[Content_Types].xml'] = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
    + '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    + '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    + '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    + '<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    + '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    + '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
    + slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')
    + extraTypes + '</Types>';
  f['_rels/.rels'] = rels([['rId1', 'officeDocument', 'ppt/presentation.xml']]);
  f['ppt/presentation.xml'] = XML + `<p:presentation ${NSDECL} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>`
    + `<p:sldIdLst>${ord.map((no, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + no}"/>`).join('')}</p:sldIdLst>`
    + `<p:sldSz cx="${E(SLIDE_W)}" cy="${E(SLIDE_H)}"/><p:notesSz cx="6858000" cy="9144000"/>`
    + `<p:defaultTextStyle><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>`;
  f['ppt/_rels/presentation.xml.rels'] = rels([['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'], ['rId2', 'theme', 'theme/theme1.xml'], ['rId3', 'tableStyles', 'tableStyles.xml'],
    ...slides.map((_, i) => [`rId${11 + i}`, 'slide', `slides/slide${i + 1}.xml`])]);
  f['ppt/tableStyles.xml'] = XML + `<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
  f['ppt/theme/theme1.xml'] = themeXml;
  f['ppt/slideMasters/slideMaster1.xml'] = masterXml;
  f['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId2', 'slideLayout', '../slideLayouts/slideLayout2.xml'], ['rId3', 'theme', '../theme/theme1.xml']]);
  f['ppt/slideLayouts/slideLayout1.xml'] = layoutXml('Titre et contenu', false);
  f['ppt/slideLayouts/slideLayout2.xml'] = layoutXml('Sombre', true);
  for (const k of [1, 2]) f[`ppt/slideLayouts/_rels/slideLayout${k}.xml.rels`] = rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]);
  slides.forEach((s, i) => {
    f[`ppt/slides/slide${i + 1}.xml`] = s.xml;
    f[`ppt/slides/_rels/slide${i + 1}.xml.rels`] = rels([['rId1', 'slideLayout', `../slideLayouts/slideLayout${s.layout}.xml`], ...s.rels]);
  });
  void n;
  return Object.assign(f, extraFiles);
};

// ───────────────────────── PNG synthétique ─────────────────────────
function makePng(w = 64, h = 48) {
  const png = new PNG({ width: w, height: h });
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = (y * w + x) * 4;
    const chk = ((x >> 3) + (y >> 3)) & 1;
    png.data[i] = Math.round(255 * x / w); png.data[i + 1] = chk ? 200 : 90; png.data[i + 2] = Math.round(255 * y / h); png.data[i + 3] = 255;
  }
  return new Uint8Array(PNG.sync.write(png));
}

// ───────────────────────── deck « animations » ─────────────────────────
function animationsDeck() {
  const box = (id, name, x, color, y = 120) => sp({ id, name, x, y, w: 160, h: 120, fill: solid(color), prst: 'roundRect',
    paras: [para(name, { algn: 'ctr', r: { sz: 18, b: true, color: 'FFFFFF' } })] });
  const slides = [];
  // 1 : trois formes, trois clics (fade / wipe(down) / zoom)
  slides.push(slide({ shapes: [label(2, 'Trois clics : fondu, balayage, zoom'), box(3, 'Fondu', 40, '2E75B6'), box(4, 'Balayage', 280, 'ED7D31'), box(5, 'Zoom', 520, '70AD47')],
    timing: timing([[[fx(3, 'fade')]], [[fx(4, 'wipe')]], [[fx(5, 'zoom', { dur: 600 })]]]) }));
  // 2 : texte construit paragraphe par paragraphe
  slides.push(slide({ shapes: [label(2, 'Texte par paragraphe'),
    sp({ id: 3, name: 'Liste', x: 60, y: 80, w: 600, h: 200, txBox: true, fill: '<a:noFill/>',
      paras: [para('Premier point', { r: { sz: 24 } }), para('Deuxième point', { r: { sz: 24 } }), para('Troisième point', { r: { sz: 24 } })] })],
    timing: timing([[[fx(3, 'fade', { par: 0 })]], [[fx(3, 'fade', { par: 1 })]], [[fx(3, 'fade', { par: 2 })]]],
      '<p:bldP spid="3" grpId="0" build="p"/>') }));
  // 3 : sortie (forme 3) puis emphase (forme 4)
  slides.push(slide({ shapes: [label(2, 'Sortie puis emphase'), box(3, 'Sort', 100, 'C00000'), box(4, 'Grossit', 420, '7030A0')],
    timing: timing([[[fx(3, 'exitFade')]], [[fx(4, 'grow', { dur: 800 })]]]) }));
  // 4 : clic + avec le précédent (délai 250) + après le précédent + avec (délai 300)
  slides.push(slide({ shapes: [label(2, 'Avec / après le précédent'), box(3, 'A', 30, '2E75B6', 100), box(4, 'B', 200, 'ED7D31', 100), box(5, 'C', 370, '70AD47', 100), box(6, 'D', 540, 'FFC000', 100)],
    timing: timing([[
      [fx(3, 'fade'), fx(4, 'fade', { delay: 250, node: 'withEffect' })],
      [fx(5, 'fade', { dur: 400, node: 'afterEffect' }), fx(6, 'wipe', { delay: 300, node: 'withEffect' })],
    ]]) }));
  // 5 : trajectoire + rotation
  slides.push(slide({ shapes: [label(2, 'Trajectoire et rotation'), box(3, 'Trajet', 60, '2E75B6'), box(4, 'Tourne', 420, 'ED7D31')],
    timing: timing([[[fx(3, 'motion', { dur: 1500 })]], [[fx(4, 'spin', { dur: 1000 })]]]) }));
  // 6-8 : transitions
  slides.push(slide({ shapes: [label(2, 'Transition : fondu'), box(3, 'Fondu', 280, '2E75B6', 140)], transition: trFade(700) }));
  slides.push(slide({ shapes: [label(2, 'Transition : poussée vers le haut'), box(3, 'Push', 280, 'ED7D31', 140)], transition: trPush('u') }));
  slides.push(slide({ shapes: [label(2, 'Transition : morphing'), box(3, 'Morph', 100, '70AD47', 200)], transition: trMorph() }));
  // 9 : 1re étape automatique (« après le précédent » en tête de diapo), puis un clic
  slides.push(slide({ shapes: [label(2, 'Démarrage automatique puis clic'), box(3, 'Auto', 100, '2E75B6'), box(4, 'Clic', 420, 'ED7D31')],
    timing: timing([[[fx(3, 'fade', { node: 'afterEffect' })]], [[fx(4, 'fade')]]], '', { autoFirst: true }) }));
  // 10 : entrée « Venir de la gauche » (fly) puis emphase « Agrandir » (scale 150 %)
  slides.push(slide({ shapes: [label(2, 'Venir de la gauche, puis agrandir'), box(3, 'Arrive', 100, '7030A0'), box(4, 'Grandit', 420, '70AD47')],
    timing: timing([[[fx(3, 'flyLeft', { dur: 700 })]], [[fx(4, 'grow', { dur: 600 })]]]) }));
  return assemble(slides);
}

// ───────────────────────── deck « shapes » ─────────────────────────
function shapesDeck() {
  const slides = [];
  const cell = (id, prst, x, y, fill, extra = {}) => sp({ id, name: prst, prst, x, y, w: 120, h: 80, fill, line: ln('1F1F1F', 1.5), ...extra });
  // 1 : presets + custGeom + remplissages
  const cust = '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:cubicBezTo><a:pt x="70" y="60"/><a:pt x="30" y="60"/><a:pt x="0" y="100"/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom>';
  slides.push(slide({ shapes: [label(2, 'Formes prédéfinies et remplissages'),
    cell(3, 'rect', 30, 50, solid('2E75B6')), cell(4, 'roundRect', 170, 50, solid('ED7D31')), cell(5, 'ellipse', 310, 50, solid('70AD47')),
    cell(6, 'rightArrow', 450, 50, solid('FFC000')), cell(7, 'star5', 590, 50, solid('C00000')),
    sp({ id: 8, name: 'custGeom', x: 30, y: 160, w: 120, h: 100, geom: cust, fill: solid('7030A0'), line: ln('1F1F1F', 1) }),
    cell(9, 'rect', 170, 160, grad('2E75B6', 'FFFFFF')), cell(10, 'ellipse', 310, 160, grad('ED7D31', '7030A0', 0)),
    cell(11, 'rect', 450, 160, patt('dkUpDiag', '2E75B6', 'FFFFFF')), cell(12, 'roundRect', 590, 160, patt('pct25', 'C00000', 'FFFFFF')),
    sp({ id: 13, name: 'Ombre', x: 30, y: 290, w: 160, h: 80, prst: 'rect', fill: solid('FFFFFF'), line: ln('404040', 1), effect: shadow,
      paras: [para('Avec ombre', { algn: 'ctr', r: { sz: 16 } })] }),
  ] }));
  // 2 : lignes, flèches, tirets (+ une forme masquée hidden="1", + décor du master masqué showMasterSp="0")
  const lineRow = (id, y, dash, head, tail, color) => cxn({ id, name: `ligne ${dash ?? 'solid'}`, x: 60, y, w: 600, h: 0, line: ln(color, 3, { dash, head, tail }) });
  slides.push(slide({ shapes: [label(2, 'Lignes, flèches et tirets'),
    lineRow(3, 70, null, null, 'triangle', '2E75B6'), lineRow(4, 120, 'dash', 'oval', 'arrow', 'ED7D31'), lineRow(5, 170, 'sysDot', 'diamond', 'stealth', '70AD47'),
    lineRow(6, 220, 'dashDot', null, null, '7030A0'), lineRow(7, 270, 'lgDash', 'arrow', null, 'C00000'),
    cxn({ id: 8, name: 'diagonale', x: 60, y: 300, w: 200, h: 80, line: ln('404040', 2, { tail: 'triangle' }) }),
    sp({ id: 9, name: 'Masquée', x: 400, y: 300, w: 120, h: 60, fill: solid('FF00FF'), hidden: true, paras: [para('NE DOIT PAS APPARAÎTRE')] }),
  ], sldAttrs: ' showMasterSp="0"' }));
  // 3 : groupes imbriqués, forme tournée + miroir
  const inner = grp({ id: 5, name: 'Groupe interne', x: 40, y: 20, w: 140, h: 80, child: { x: 0, y: 0, w: 280, h: 160 } }, [
    sp({ id: 6, name: 'Interne A', x: 0, y: 0, w: 130, h: 160, fill: solid('2E75B6') }),
    sp({ id: 7, name: 'Interne B', x: 150, y: 0, w: 130, h: 160, prst: 'ellipse', fill: solid('ED7D31') }),
  ]);
  slides.push(slide({ shapes: [label(2, 'Groupes imbriqués, rotation, miroir'),
    grp({ id: 3, name: 'Groupe externe', x: 40, y: 70, w: 400, h: 240, child: { x: 0, y: 0, w: 200, h: 120 } }, [
      sp({ id: 4, name: 'Fond du groupe', x: 0, y: 0, w: 200, h: 120, fill: solid('EEF1F6'), line: ln('404040', 1) }), inner]),
    sp({ id: 8, name: 'Tournée', prst: 'rightArrow', x: 480, y: 90, w: 180, h: 70, rot: 30, fill: solid('70AD47') }),
    sp({ id: 9, name: 'Miroir', prst: 'rightArrow', x: 480, y: 200, w: 180, h: 70, flipH: true, fill: solid('C00000') }),
    sp({ id: 10, name: 'Tournée + miroir', prst: 'rightArrow', x: 480, y: 290, w: 180, h: 60, rot: 330, flipV: true, fill: solid('7030A0') }),
  ] }));
  return assemble(slides);
}

// ───────────────────────── deck « text » ─────────────────────────
function textDeck() {
  const slides = [];
  // slide1 (fichier) : placeholders hérités + puces — sera la 2e dans l'ordre de présentation
  slides.push(slide({ shapes: [
    sp({ id: 2, name: 'Titre 1', ph: 'type="title"', paras: [para('Titre hérité du gabarit')] }),
    sp({ id: 3, name: 'Contenu 2', ph: 'idx="1"', paras: [
      para('Puce de niveau 1'), para('Puce de niveau 2', { lvl: 1 }),
      para('Numéro un', { pPr: '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>', marL: 342900, indent: -342900 }),
      para('Numéro deux', { pPr: '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>', marL: 342900, indent: -342900 }),
    ] }),
  ] }));
  // slide2 (fichier) : runs, autofit, texte vertical, tabulations, numéro de diapo — 1re dans l'ordre de présentation
  slides.push(slide({ shapes: [
    sp({ id: 2, name: 'Titre 1', ph: 'type="title"', paras: [para('Runs et mise en forme')] }),
    sp({ id: 3, name: 'Runs', x: 36, y: 100, w: 440, h: 120, txBox: true, fill: '<a:noFill/>', paras: [para([
      run('Gras ', { b: 1, sz: 20 }), run('italique ', { i: 1, sz: 20 }), run('souligné ', { u: 1, sz: 20 }), run('rouge ', { color: 'C00000', sz: 20 }), run('grand', { sz: 32, color: '2E75B6' })])] }),
    sp({ id: 4, name: 'Autofit', x: 36, y: 230, w: 300, h: 70, fill: solid('EEF1F6'), line: ln('404040', 1),
      bodyPr: '<a:bodyPr wrap="square"><a:normAutofit fontScale="70000" lnSpcReduction="20000"/></a:bodyPr>',
      paras: [para('Texte réduit par normAutofit fontScale 70 % : beaucoup de mots pour remplir la boîte entière.', { r: { sz: 24 } })] }),
    sp({ id: 5, name: 'Vertical', x: 600, y: 100, w: 70, h: 220, fill: solid('FFF2CC'), line: ln('404040', 1),
      bodyPr: '<a:bodyPr vert="vert270" wrap="square"><a:noAutofit/></a:bodyPr>', paras: [para('Texte vertical', { algn: 'ctr', r: { sz: 18 } })] }),
    sp({ id: 6, name: 'Tabulations', x: 36, y: 320, w: 420, h: 40, txBox: true, fill: '<a:noFill/>',
      paras: [para('Nom\tQuantité\tPrix', { pPr: '<a:tabLst><a:tab pos="1905000" algn="l"/><a:tab pos="3810000" algn="l"/></a:tabLst>', r: { sz: 18 } })] }),
    sp({ id: 7, name: 'Numéro', x: 500, y: 330, w: 100, h: 30, txBox: true, fill: '<a:noFill/>',
      paras: [`<a:p><a:r><a:rPr lang="fr-FR" sz="1800"/><a:t>Diapo </a:t></a:r><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="fr-FR" sz="1800"/><a:t>‹#›</a:t></a:fld></a:p>`] }),
  ] }));
  // slide3 : thème sombre (layout 2 : clrMapOvr bg1->dk1)
  slides.push(slide({ layout: 2, bg: '<p:bgPr><a:solidFill><a:srgbClr val="1B1B2F"/></a:solidFill><a:effectLst/></p:bgPr>', shapes: [
    sp({ id: 2, name: 'Titre 1', ph: 'type="title"', paras: [para('Diapo sombre', { r: { color: 'FFFFFF' } })] }),
    sp({ id: 3, name: 'Contenu 2', ph: 'idx="1"', paras: [para('Texte clair sur fond sombre', { r: { color: 'E0E0E0' } })] }),
    sp({ id: 4, name: 'Carte', x: 36, y: 280, w: 300, h: 60, fill: solid('accent1'), paras: [para('Accent 1', { algn: 'ctr', r: { sz: 20, color: 'FFFFFF' } })] }),
  ] }));
  // slide4 : layout « Sombre » SANS couleur explicite : fond et texte viennent du clrMapOvr du layout
  slides.push(slide({ layout: 2, shapes: [
    sp({ id: 2, name: 'Titre 1', ph: 'type="title"', paras: [para('Couleurs héritées du layout sombre')] }),
    // titre idx=7 : absent du layout ET du master (clé exacte) -> hérite du titre du master PAR TYPE
    sp({ id: 3, name: 'Numéro 01', ph: 'type="title" idx="7"', x: 36, y: 200, w: 200, h: 60, paras: [para('01')] }),
  ] }));
  // ordre de présentation : slide2, slide1, slide3, slide4
  return assemble(slides, { order: [2, 1, 3, 4] });
}

// ───────────────────────── deck « tables-charts » ─────────────────────────
function chartXml(kind) {
  const cats = ['T1', 'T2', 'T3', 'T4'], vals = [12, 19, 7, 15];
  const strCache = `<c:strCache><c:ptCount val="4"/>${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
  const ser = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Feuil1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Ventes</c:v></c:pt></c:strCache></c:strRef></c:tx>`
    + `${kind === 'bar' ? '<c:spPr><a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>' : ''}`
    + `<c:cat><c:strRef><c:f>Feuil1!$A$2:$A$5</c:f>${strCache}</c:strRef></c:cat><c:val><c:numRef><c:f>Feuil1!$B$2:$B$5</c:f>${numCache}</c:numRef></c:val></c:ser>`;
  const plot = kind === 'bar'
    ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="80"/><c:axId val="111"/><c:axId val="222"/></c:barChart>`
      + `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>`
      + `<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>`
    : `<c:pieChart><c:varyColors val="1"/>${ser}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls><c:firstSliceAng val="0"/></c:pieChart>`;
  const title = kind === 'bar' ? 'Ventes par trimestre' : 'Répartition';
  return XML + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>`
    + `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="fr-FR" sz="1400"/><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    + `<c:plotArea><c:layout/>${plot}</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function tablesChartsDeck() {
  const slides = [];
  const tc = (t, attrs = '', tcPr = '') => `<a:tc${attrs}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="fr-FR" sz="1600"/><a:t>${esc(t)}</a:t></a:r></a:p></a:txBody><a:tcPr>${tcPr}</a:tcPr></a:tc>`;
  const tcMerged = `<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="fr-FR"/></a:p></a:txBody><a:tcPr/></a:tc>`;
  const tbl = `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>`
    + `<a:tblGrid><a:gridCol w="${E(180)}"/><a:gridCol w="${E(140)}"/><a:gridCol w="${E(140)}"/></a:tblGrid>`
    + `<a:tr h="${E(36)}">${tc('Produit')}${tc('Prix')}${tc('Stock')}</a:tr>`
    + `<a:tr h="${E(36)}">${tc('Stylo')}${tc('1,20')}${tc('300')}</a:tr>`
    + `<a:tr h="${E(36)}">${tc('Fusion sur deux colonnes', ' gridSpan="2"')}${tcMerged}${tc('42')}</a:tr>`
    + `<a:tr h="${E(36)}">${tc('Cahier')}${tc('3,50')}${tc('120')}</a:tr></a:tbl></a:graphicData>`;
  slides.push(slide({ shapes: [label(2, 'Tableau avec style, bandes et cellule fusionnée'), graphicFrame({ id: 3, name: 'Tableau', x: 60, y: 70, w: 460, h: 144 }, tbl)] }));
  const chartData = (rId) => `<a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rId}"/></a:graphicData>`;
  slides.push(slide({ shapes: [label(2, 'Histogramme et camembert'),
    graphicFrame({ id: 3, name: 'Histogramme', x: 20, y: 50, w: 340, h: 320 }, chartData('rId2')),
    graphicFrame({ id: 4, name: 'Camembert', x: 370, y: 50, w: 330, h: 320 }, chartData('rId3'))],
    rels: [['rId2', 'chart', '../charts/chart1.xml'], ['rId3', 'chart', '../charts/chart2.xml']] }));
  slides.push(slide({ shapes: [label(2, 'Images : entière, rognée, ellipse'),
    pic({ id: 3, name: 'Image entière', rId: 'rId2', x: 30, y: 70, w: 200, h: 150 }),
    pic({ id: 4, name: 'Image rognée', rId: 'rId2', x: 260, y: 70, w: 200, h: 150, srcRect: 'l="25000" t="25000" r="25000" b="25000"' }),
    pic({ id: 5, name: 'Image ellipse', rId: 'rId2', x: 490, y: 70, w: 200, h: 150, prst: 'ellipse', line: ln('2E75B6', 3) })],
    rels: [['rId2', 'image', '../media/image1.png']] }));
  return assemble(slides, {
    extraFiles: { 'ppt/charts/chart1.xml': chartXml('bar'), 'ppt/charts/chart2.xml': chartXml('pie'), 'ppt/media/image1.png': makePng() },
    extraTypes: ['1', '2'].map((n) => `<Override PartName="/ppt/charts/chart${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join(''),
  });
}

// ───────────────────────── API publique ─────────────────────────
export function buildDecks() {
  return { animations: animationsDeck(), shapes: shapesDeck(), text: textDeck(), 'tables-charts': tablesChartsDeck() };
}

export function writeDeck(files, dir) {
  for (const [p, data] of Object.entries(files)) {
    const out = join(dir, p);
    mkdirSync(dirname(out), { recursive: true });
    writeFileSync(out, typeof data === 'string' ? data : Buffer.from(data));
  }
}

export function zipDeck(files) {
  const entries = {};
  for (const [p, data] of Object.entries(files)) entries[p] = typeof data === 'string' ? strToU8(data) : data;
  return zipSync(entries);
}
