// ═══ [CŒUR PPTX] ═══ carte d'ensemble : docs/ARCHITECTURE.md (§4) · consommé par build.mjs + build-player.mjs
// Moteur de rendu PPTX -> SVG, en JS pur, SANS PowerPoint.
// Rend une diapo (formes vectorielles + texte + images) dans le repère EMU de la diapo.
// Le SVG en viewBox EMU = fidélité géométrique exacte (coins arrondis, custGeom, rotations).
//
// Couverture v1 : spTree ordonné (z-order), groupes + transforms, prstGeom courants + custGeom,
// fills solides/dégradés/image, lignes, texte multi-run via foreignObject, héritage thème
// (clrScheme + fmtScheme) et placeholders (géométrie + styles via layout/master).
//
// API : renderSlide(pptDir, slideNo, {embedImages}) -> { svg, w, h }

import { readFileSync, existsSync, statSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { XMLParser } from 'fast-xml-parser';
import { presetGeomD, evalGeom, presetGeomParts, evalGeomParts } from './preset-geom.mjs';
import { renderChart } from './chart-render.mjs';

// trimValues:false -> conserve les espaces en bord de <a:t> (sinon les mots se collent
// entre runs successifs, ex. "Marge " + "qu'on" -> "Margequ'on").
const P = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', preserveOrder: true, parseTagValue: false, parseAttributeValue: false, trimValues: false });
const PF = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', parseTagValue: false, parseAttributeValue: false, trimValues: false }); // flat, pour thème/rels

// ---- cache de parse pour fichiers STABLES d'un deck (thème/master/layout/rels/presentation/tableStyles) ----
// renderSlide est appelé par diapo, mais N diapos partagent le même master/layout/thème : sans cache on
// reparse ~64 Ko de XML par diapo (master+layout parsés 2× chacun). Clé = chemin + mtime + taille : si un
// dossier temporaire est réutilisé avec un contenu DIFFÉRENT (ex. batch check.mjs), la signature change et
// le cache est invalidé -> jamais de rendu périmé. Les fichiers PROPRES à la diapo (slideN.xml, ses rels)
// ne sont PAS mis en cache. Les arbres mis en cache sont traités en LECTURE SEULE (le code re-sérialise +
// reparse tout sous-arbre qu'il doit muter) -> partage sûr entre diapos (prouvé par hash SVG identique).
const _pc = new Map();
function _stamp(path) { const s = statSync(path); return s.mtimeMs + ':' + s.size; }
function cachedParse(path, kind /* 'txt' | 'flat' | 'ord' */) {
  const sig = _stamp(path);
  let e = _pc.get(path);
  if (!e || e.sig !== sig) { e = { sig }; _pc.set(path, e); }
  if (e.txt === undefined) e.txt = readFileSync(path, 'utf8');
  if (kind === 'txt') return e.txt;
  if (kind === 'flat') return (e.flat ??= PF.parse(e.txt));
  return (e.ord ??= P.parse(e.txt));
}

// ---- helpers arbre preserveOrder ----
const tagOf = (n) => Object.keys(n).find((k) => k !== ':@');
const at = (n) => n?.[':@'] ?? {};
const kids = (n) => { const v = n?.[tagOf(n)]; return Array.isArray(v) ? v : []; };
const byTag = (n, t) => kids(n).filter((c) => tagOf(c) === t);
const first = (n, t) => byTag(n, t)[0];
function deepFind(n, t) {
  if (!n || typeof n !== 'object') return null;
  if (tagOf(n) === t) return n;
  for (const c of kids(n)) { const r = deepFind(c, t); if (r) return r; }
  return null;
}
// recherche profonde d'une clé dans un objet PLAT (parse non-ordonné) -> 1re valeur trouvée
function deepFindFlat(obj, key) {
  if (!obj || typeof obj !== 'object') return null;
  if (obj[key] !== undefined) return obj[key];
  for (const k of Object.keys(obj)) { const r = deepFindFlat(obj[k], key); if (r) return r; }
  return null;
}
// trouve le 1er noeud dont le tag se termine par `suffix` (ignore le préfixe de namespace)
function deepFindSuffix(n, suffix) {
  if (!n || typeof n !== 'object') return null;
  const t = tagOf(n);
  if (t && t.endsWith(suffix)) return n;
  for (const c of kids(n)) { const r = deepFindSuffix(c, suffix); if (r) return r; }
  return null;
}
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// OMML (§22.1 Office Math) -> HTML dans le foreignObject. Noeud ORDONNÉ (preserveOrder).
// Gère fraction (m:f), parenthèses/délimiteurs (m:d), exposant/indice (m:sSup/m:sSub),
// racine (m:rad), n-aire (m:nary, somme/intégrale). Le math est en italique (convention).
function mathText(node) { // texte brut d'un noeud (m:t / #text)
  // NFKC : convertit les « Mathematical Alphanumeric Symbols » (𝜌 𝑓 𝒅 𝑷 𝒁 𝒍…, U+1D400+)
  // en lettres normales (ρ f d P Z l) — sinon Poppins ne les a pas et affiche des blancs/tofu.
  if (node['#text'] !== undefined) return esc(String(node['#text']).normalize('NFKC'));
  return kids(node).map(mathText).join('');
}
function renderMath(node) {
  if (node['#text'] !== undefined) return esc(String(node['#text']).normalize('NFKC'));
  const t = tagOf(node);
  if (/Pr$/.test(t)) return ''; // m:fPr, m:dPr, m:rPr, m:ctrlPr, m:naryPr... = propriétés
  switch (t) {
    case 'm:t': return `<span style="font-style:italic">${mathText(node)}</span>`;
    case 'm:f': {
      const num = renderMath(first(node, 'm:num') || {}), den = renderMath(first(node, 'm:den') || {});
      return `<span style="display:inline-flex;flex-direction:column;vertical-align:middle;text-align:center;line-height:1.05;font-size:.95em">`
        + `<span style="border-bottom:.06em solid;padding:0 .25em">${num}</span><span style="padding:0 .25em">${den}</span></span>`;
    }
    case 'm:d': {
      const dPr = first(node, 'm:dPr');
      const beg = at(first(dPr || {}, 'm:begChr'))['@m:val'] ?? '(', end = at(first(dPr || {}, 'm:endChr'))['@m:val'] ?? ')';
      return esc(beg) + byTag(node, 'm:e').map(renderMath).join('') + esc(end);
    }
    case 'm:sSup': return renderMath(first(node, 'm:e') || {}) + `<sup style="font-size:.7em">${renderMath(first(node, 'm:sup') || {})}</sup>`;
    case 'm:sSub': return renderMath(first(node, 'm:e') || {}) + `<sub style="font-size:.7em">${renderMath(first(node, 'm:sub') || {})}</sub>`;
    case 'm:rad': return `√<span style="border-top:.06em solid;padding:0 .15em">${renderMath(first(node, 'm:e') || {})}</span>`;
    case 'm:nary': {
      const op = at(first(first(node, 'm:naryPr') || {}, 'm:chr'))['@m:val'] ?? '∫';
      const sub = renderMath(first(node, 'm:sub') || {}), sup = renderMath(first(node, 'm:sup') || {});
      return `<span style="font-size:1.3em">${esc(op)}</span>${sub ? `<sub style="font-size:.7em">${sub}</sub>` : ''}${sup ? `<sup style="font-size:.7em">${sup}</sup>` : ''}${renderMath(first(node, 'm:e') || {})}`;
    }
    default: return kids(node).map(renderMath).join(''); // m:oMath, m:oMathPara, m:e, m:r, m:num, m:den...
  }
}

// puce d'un pPr (flat) : {none} | {char,font} | {num,font} | null (non défini -> hériter)
function bulletOf(pPr) {
  if (!pPr) return null;
  if (pPr['a:buNone'] !== undefined) return { none: true };
  // a:buClr = couleur PROPRE de la puce (souvent une couleur d'accent différente du texte). Nœud brut
  // conservé ; résolu au rendu (a besoin du thème/ctx). a:buSzPct/@val = taille de puce en % de la police.
  const clr = pPr['a:buClr'] || null;
  const szPct = pPr['a:buSzPct']?.['@val'] ? +pPr['a:buSzPct']['@val'] / 100000 : null;
  // a:buBlip = puce IMAGE (petit logo/icône). Nœud brut conservé ; résolu au rendu (curRels/pptDir).
  if (pPr['a:buBlip']) return { blip: pPr['a:buBlip']['a:blip'], szPct };
  if (pPr['a:buChar']) return { char: pPr['a:buChar']['@char'], font: pPr['a:buFont']?.['@typeface'], clr, szPct };
  if (pPr['a:buAutoNum']) return { num: pPr['a:buAutoNum']['@type'] || 'arabicPeriod', font: pPr['a:buFont']?.['@typeface'], clr, szPct };
  return null;
}
// police -> famille CSS. Substituts LIBRES métrique-compatibles (embarqués) pour les polices non libres
// courantes : garantit largeur/retour-à-la-ligne corrects MÊME sans la police installée (ex. LMS Linux).
// Ordre : police réelle d'abord (si installée), puis substitut embarqué, puis générique.
const SERIF_RE = /times|cambria|georgia|bernhard|garamond|antiqua|minion|book|schoolbook/i;
const SYMB_RE = /wingding|webding|symbol/i;
const FONT_SUBST = {
  'arial': "'Arial','Liberation Sans',sans-serif", 'helvetica': "'Helvetica','Liberation Sans',sans-serif",
  'arial narrow': "'Arial Narrow','Liberation Sans Narrow',sans-serif",
  'times new roman': "'Times New Roman','Liberation Serif',serif", 'times': "'Times','Liberation Serif',serif",
  'courier new': "'Courier New','Liberation Mono',monospace", 'courier': "'Courier','Liberation Mono',monospace",
  'calibri': "'Calibri','Carlito',sans-serif", 'cambria': "'Cambria','Caladea',serif",
};
function fontFamily(f) {
  if (!f) return 'sans-serif';
  if (SYMB_RE.test(f)) return `'${esc(f)}'`;
  const key = f.toLowerCase().trim();
  if (FONT_SUBST[key]) return FONT_SUBST[key];
  return SERIF_RE.test(f) ? `'${esc(f)}',serif` : `'${esc(f)}',sans-serif`;
}
// Polices SYMBOLE (Wingdings/Webdings/Symbol) : le glyphe n'existe qu'avec la police installée -> tofu
// sur Linux (cible LMS) et dans Chromium headless. On mappe les codepoints vers leur équivalent Unicode,
// rendus en police NORMALE = lisibles partout. Gère l'encodage zone privée PUA (0xF000+octet en run).
const WINGD_RE = /wingding|webding/i; // dingbats PURS (que des symboles) — décoratif
const SYMBOLFONT_RE = /^symbol$/i;      // Adobe « Symbol » = grec + flèches + math (PAS décoratif)
// Wingdings -> Unicode (puces/dingbats). Inconnu -> « • » (marqueur générique, jamais de tofu).
const WINGDINGS_MAP = {
  0x6C: '●', 0x6D: '○', 0x6E: '■', 0x6F: '□', 0x71: '❑', 0x75: '◆',
  0xA7: '▪', 0xA8: '▫', 0xB7: '•', 0xD8: '➢', 0xD9: '➢',
  0xFC: '✔', 0xFD: '✔', 0xFE: '☒', 0xFB: '✘',
};
// Adobe « Symbol » -> Unicode : grec (ASCII->grec) + flèches (0xAB-AF, 0xDB-DF) + opérateurs math.
// Inconnu -> on GARDE le caractère d'origine (symbole signifiant : ne jamais deviner à tort).
const SYMBOL_MAP = (() => {
  const m = {};
  const low = 'abgdezhqiklmnxoprstufcyw', lowG = 'αβγδεζηθικλμνξοπρστυφχψω';
  const up  = 'ABGDEZHQIKLMNXOPRSTUFCYW', upG  = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ';
  for (let i = 0; i < low.length; i++) { m[low.charCodeAt(i)] = lowG[i]; m[up.charCodeAt(i)] = upG[i]; }
  Object.assign(m, {
    0xAB:'↔',0xAC:'←',0xAD:'↑',0xAE:'→',0xAF:'↓', 0xDB:'⇔',0xDC:'⇐',0xDD:'⇑',0xDE:'⇒',0xDF:'⇓',
    0xA3:'≤',0xB3:'≥',0xB9:'≠',0xBB:'≈',0xBA:'≡',0xB1:'±',0xB4:'×',0xB8:'÷',0xB0:'°',0xB7:'·',
    0xD6:'√',0xA5:'∞',0xB6:'∂',0xF2:'∫',0x2D:'−',0xE5:'∑',0xCF:'∨',0xCE:'∈',
  });
  return m;
})();
function symFontKind(font) {
  return !font ? null : WINGD_RE.test(font) ? 'wing' : SYMBOLFONT_RE.test(font.trim()) ? 'sym' : null;
}
// mappe UN char (codepoint) d'une police symbole vers Unicode. wing: inconnu->'•' ; sym: inconnu->char.
function symToUnicode(ch, kind) {
  let cp = ch.codePointAt(0);
  if (cp >= 0xF000 && cp <= 0xF0FF) cp -= 0xF000; // PUA -> octet réel
  if (kind === 'wing') return WINGDINGS_MAP[cp] || '•';
  return SYMBOL_MAP[cp] ?? ch; // sym : garder l'original si non mappé
}
// -> {sym, symbolic} pour une PUCE caractère : symbolic=true si remappée (rendre en police normale)
function bulletSym(char, font) {
  const kind = symFontKind(font);
  if (char == null || !kind) return { sym: char, symbolic: false };
  return { sym: symToUnicode(char, kind), symbolic: true };
}
// remappe un TEXTE de run en police symbole -> Unicode (anti-tofu). `sym` = typeface de <a:sym> (prio).
function symbolText(raw, latinFont, symFont) {
  const kind = symFontKind(symFont) || symFontKind(latinFont);
  if (!raw || !kind) return null;
  return [...raw].map(c => /\s/.test(c) ? c : symToUnicode(c, kind)).join('');
}

// EMU -> points (1 pt = 12700 EMU). On rend en points pour des tailles de police normales
// (les foreignObject cassent avec des font-size en centaines de milliers de px).
const S = 1 / 12700;
const u = (v) => +(Number(v) * S).toFixed(2);

// ================= contexte document =================
// Cache du CŒUR de contexte (parties immuables dérivées du thème) : keyé par fichier de thème (+ sa
// signature mtime/taille). loadContext() renvoie une COPIE avec clrMap/phClr remis à null car ceux-là
// sont MUTÉS par diapo (clrMap du masque, phClr courant) et ne doivent jamais être partagés.
const _ctxCache = new Map();
export function loadContext(pptDir, themePath) {
  const themeFile = (themePath && existsSync(themePath)) ? themePath : join(pptDir, 'theme', 'theme1.xml');
  const sig = _stamp(themeFile);
  let c = _ctxCache.get(themeFile);
  if (!c || c.sig !== sig) { c = { sig, core: buildContext(pptDir, themeFile) }; _ctxCache.set(themeFile, c); }
  return { ...c.core, clrMap: null, phClr: null };
}
function buildContext(pptDir, themeFile) {
  // theme du MASTER de la diapo (un deck multi-master a plusieurs thèmes : ex. deck C a
  // theme1 lt2=C78243 orange ET theme2 lt2=D9CAB7 beige -> mauvais thème = mauvaise couleur de fond).
  const theme = cachedParse(themeFile, 'flat');
  const el = theme['a:theme']['a:themeElements'];
  const scheme = el['a:clrScheme'];
  const clr = {}; // slots BRUTS du thème (dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink)
  for (const [k, v] of Object.entries(scheme)) {
    if (!k.startsWith('a:')) continue;
    clr[k.slice(2)] = v['a:srgbClr']?.['@val'] ??
      { windowText: '000000', window: 'FFFFFF' }[v['a:sysClr']?.['@val']] ??
      v['a:sysClr']?.['@lastClr'] ?? '000000';
  }
  // listes de styles du fmtScheme, EN ORDRE (fillRef/bgRef indexent dedans) -> noeuds flat
  const themeO = cachedParse(themeFile, 'ord');
  const themeRoot = themeO.find(n => tagOf(n) === 'a:theme');
  const fmt = deepFind(themeRoot, 'a:fmtScheme');
  const orderedFlatList = (listTag) => {
    const lst = fmt && first(fmt, listTag);
    if (!lst) return [];
    return kids(lst).map(c => PF.parse(`<root xmlns:a="a" xmlns:p="p">${serialize(c)}</root>`).root);
  };
  // styles de tableaux (fills header/bandes), best-effort
  const tableStyles = {};
  try {
    const tsx = cachedParse(join(pptDir, 'tableStyles.xml'), 'flat');
    const list = tsx['a:tblStyleLst']?.['a:tblStyle'];
    for (const st of (Array.isArray(list) ? list : list ? [list] : [])) {
      const fillOf = (part) => st[part]?.['a:tcStyle']?.['a:fill'] || null;
      tableStyles[st['@styleId']] = {
        firstRow: fillOf('a:firstRow'), band1: fillOf('a:band1H'), band2: fillOf('a:band2H'),
      };
    }
  } catch {}
  return {
    pptDir,
    clr,
    clrMap: null, // rempli avec le clrMap du masque
    phClr: null,  // couleur de substitution courante (pour phClr)
    majorFont: el['a:fontScheme']?.['a:majorFont']?.['a:latin']?.['@typeface'] ?? 'Calibri',
    minorFont: el['a:fontScheme']?.['a:minorFont']?.['a:latin']?.['@typeface'] ?? 'Calibri',
    fillStyles: orderedFlatList('a:fillStyleLst'),
    lnStyles: orderedFlatList('a:lnStyleLst'),
    effectStyles: orderedFlatList('a:effectStyleLst'), // ombres/lueurs de thème indexées par a:effectRef
    bgFillStyles: orderedFlatList('a:bgFillStyleLst'),
    tableStyles,
  };
}

// ================= couleurs =================
function hslFromRgb(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h, s, l = (mx + mn) / 2;
  if (mx === mn) { h = s = 0; }
  else { const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
  return [h, s, l];
}
function rgbFromHsl(h, s, l) {
  let r, g, b;
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const hue = (t) => { t = (t + 1) % 1; return t < 1/6 ? p + (q - p) * 6 * t : t < 1/2 ? q : t < 2/3 ? p + (q - p) * (2/3 - t) * 6 : p; };
    r = hue(h + 1/3); g = hue(h); b = hue(h - 1/3); }
  return [Math.round(r*255), Math.round(g*255), Math.round(b*255)];
}
// résout un noeud couleur DrawingML flat -> {r,g,b,a}
function resolveColorFlat(node, ctx) {
  if (!node) return null;
  let hex = null, mods = node;
  const schemeHex = (nm) => {
    if (nm === 'phClr') nm = ctx.phClr || 'accent1';
    if (ctx.clrMap && ctx.clrMap[nm]) nm = ctx.clrMap[nm]; // bg1/tx1/... -> dk1/lt1/...
    return ctx.clr[nm] ?? '888888';
  };
  if (node['a:srgbClr']) { hex = node['a:srgbClr']['@val']; mods = node['a:srgbClr']; }
  else if (node['a:schemeClr']) { hex = schemeHex(node['a:schemeClr']['@val']); mods = node['a:schemeClr']; }
  else if (node['a:sysClr']) { hex = node['a:sysClr']['@lastClr'] ?? '000000'; mods = node['a:sysClr']; }
  else if (node['a:prstClr']) { hex = ({ black:'000000', white:'FFFFFF', red:'FF0000' })[node['a:prstClr']['@val']] ?? '888888'; mods = node['a:prstClr']; }
  else if (node['a:scrgbClr']) { const c = node['a:scrgbClr']; const to = (v)=>Math.round(parseInt(v)/100000*255); return { r:to(c['@r']), g:to(c['@g']), b:to(c['@b']), a:1 }; }
  else return null;
  let r = parseInt(hex.slice(0,2),16), g = parseInt(hex.slice(2,4),16), b = parseInt(hex.slice(4,6),16), a = 1;
  const num = (v)=>parseInt(v)/100000;
  let [h,s,l] = hslFromRgb(r,g,b);
  if (mods['a:lumMod']) l *= num(mods['a:lumMod']['@val']);
  if (mods['a:lumOff']) l += num(mods['a:lumOff']['@val']);
  l = Math.max(0, Math.min(1, l));
  if (mods['a:lumMod'] || mods['a:lumOff']) [r,g,b] = rgbFromHsl(h,s,l);
  if (mods['a:shade']) { const k = num(mods['a:shade']['@val']); r*=k; g*=k; b*=k; }
  if (mods['a:tint']) { const k = num(mods['a:tint']['@val']); r = r*k + 255*(1-k); g = g*k + 255*(1-k); b = b*k + 255*(1-k); }
  if (mods['a:alpha']) a = num(mods['a:alpha']['@val']);
  return { r:Math.round(Math.max(0,Math.min(255,r))), g:Math.round(Math.max(0,Math.min(255,g))), b:Math.round(Math.max(0,Math.min(255,b))), a };
}
const cssColor = (c) => c ? (c.a < 1 ? `rgba(${c.r},${c.g},${c.b},${c.a.toFixed(3)})` : `rgb(${c.r},${c.g},${c.b})`) : 'none';

// ================= géométrie preset -> chemin SVG =================
// Délègue au moteur générique preset-geom.mjs (187 formes officielles ECMA-376 §20.1.9,
// formules de guides évaluées). adjO = map { nomGuide: fmla } depuis l'avLst de la forme.
function presetPath(prst, w, h, adjO) {
  const d = presetGeomD(prst, w, h, adjO);
  return d ?? `M0 0 H${w} V${h} H0 Z`;   // nom inconnu -> rectangle
}

// custGeom (nœud ORDONNÉ preserveOrder) -> chemin SVG, avec évaluation des formules de guides.
// Construit un def { av, gd, paths:[{w,h,fill,stroke,cmds}] } et réutilise evalGeom.
const CUST_CMD = { 'a:moveTo':'M', 'a:lnTo':'L', 'a:arcTo':'A', 'a:cubicBezTo':'C', 'a:quadBezTo':'Q', 'a:close':'Z' };
function custGeomDef(custO) {
  const gdOf = (listTag) => {
    const lst = first(custO, listTag); if (!lst) return [];
    return byTag(lst, 'a:gd').map((g) => [at(g)['@name'], at(g)['@fmla']]);
  };
  const def = { av: gdOf('a:avLst'), gd: gdOf('a:gdLst'), paths: [] };
  const pathLst = first(custO, 'a:pathLst');
  for (const p of (pathLst ? byTag(pathLst, 'a:path') : [])) {
    const pa = at(p);
    const path = { w: +pa['@w'] || 0, h: +pa['@h'] || 0,
      fill: pa['@fill'] ?? 'norm', stroke: pa['@stroke'] !== 'false', cmds: [] };
    for (const seg of kids(p)) {
      const c = CUST_CMD[tagOf(seg)]; if (!c) continue;
      if (c === 'Z') { path.cmds.push(['Z']); continue; }
      if (c === 'A') { const a = at(seg); path.cmds.push(['A', a['@wR'], a['@hR'], a['@stAng'], a['@swAng']]); continue; }
      const pts = byTag(seg, 'a:pt').map((pt) => [at(pt)['@x'], at(pt)['@y']]);
      path.cmds.push([c, ...pts.flat()]);
    }
    def.paths.push(path);
  }
  return def;
}
function custGeomPath(custO, w, h) {
  return evalGeom(custGeomDef(custO), w, h);
}
function custGeomParts(custO, w, h) {
  return evalGeomParts(custGeomDef(custO), w, h);
}

// ================= rels & images =================
function slideRels(pptDir, slideNo) {
  const p = join(pptDir, 'slides', '_rels', `slide${slideNo}.xml.rels`);
  if (!existsSync(p)) return {};
  const r = cachedParse(p, 'flat');
  const out = {};
  const rl = r.Relationships?.Relationship;
  for (const x of (Array.isArray(rl) ? rl : rl ? [rl] : [])) out[x['@Id']] = x['@Target'];
  return out;
}
function relsOf(path) {
  const d = dirname(path), b = path.split(/[\\/]/).pop();
  const p = join(d, '_rels', b + '.rels');
  if (!existsSync(p)) return {};
  const r = cachedParse(p, 'flat');
  const out = {}; const rl = r.Relationships?.Relationship;
  for (const x of (Array.isArray(rl) ? rl : rl ? [rl] : [])) out[x['@Id']] = x['@Target'];
  return out;
}
const MIME = { png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif', bmp:'image/bmp', svg:'image/svg+xml', emf:'', wmf:'' };
function imageHref(pptDir, target, embed) {
  let abs = join(pptDir, target.replace(/^\.\.\//,''));
  if (!existsSync(abs)) return null;
  let ext = abs.split('.').pop().toLowerCase();
  // EMF/WMF (vecteur Windows, non affichable en <image>) : servir le PNG pré-converti « <fichier>.png »
  // (préprocessing convert-metafiles.ps1 via .NET/WIC). Beaucoup de diagrammes techniques sont en EMF.
  if (ext === 'emf' || ext === 'wmf') {
    const png = abs + '.png';
    if (existsSync(png)) { abs = png; ext = 'png'; } else return null;
  }
  if (!embed) return 'file:///' + abs.replace(/\\/g,'/');
  const mime = MIME[ext];
  if (!mime) return null;
  return `data:${mime};base64,` + readFileSync(abs).toString('base64');
}

// ================= placeholders (héritage layout/master) =================
// Une forme-texte de la diapo hérite géométrie + fill + styles de texte de SON placeholder
// de layout (identifié par type:idx). Le layout hérite lui-même du masque.
function placeholderGeom(pptDir, slideNo, rels) {
  const layoutTgt = Object.values(rels).find((t) => t && t.includes('slideLayout'));
  const geom = {};
  if (!layoutTgt) return { geom };
  const layoutPath = join(pptDir, 'slides', layoutTgt);
  if (!existsSync(layoutPath)) return { geom };
  const lx = cachedParse(layoutPath, 'flat');
  const tree = lx['p:sldLayout']?.['p:cSld']?.['p:spTree'];
  const arr = (x)=> x===undefined?[]:Array.isArray(x)?x:[x];
  for (const sp of arr(tree?.['p:sp'])) {
    const ph = sp['p:nvSpPr']?.['p:nvPr']?.['p:ph']; if (!ph) continue;
    const key = `${ph['@type']??'body'}:${ph['@idx']??'0'}`;
    // le layout a plusieurs placeholders "body:0" -> on garde le premier par clé unique,
    // mais les cartes ont des idx distincts (27,30,33...) donc pas de collision gênante
    if (!geom[key]) geom[key] = { xfrm: sp['p:spPr']?.['a:xfrm'], spPr: sp['p:spPr'], style: sp['p:style'],
      lstStyle: sp['p:txBody']?.['a:lstStyle'], bodyPr: sp['p:txBody']?.['a:bodyPr'] };
  }
  return { geom };
}

// ================= rendu d'une forme =================
function xfrmAttrs(xfrm) {
  if (!xfrm) return null;
  return {
    x: Number(xfrm['@x'] ?? at(deepFind(xfrm,'a:off'))['@x']),
  };
}

// version preserveOrder: on parse le spPr flat pour lire fills/geom facilement
export function renderSlide(pptDir, slideNo, opts = {}) {
  const embed = opts.embedImages !== false;
  // préfixe d'ID UNIQUE par diapo : clip/gradient/marqueur/filtre. Sans ça, plusieurs diapos dans UNE
  // page (viewer.html) partagent des IDs (clip0, g0…) qui collisionnent -> mauvais clip/dégradé (ex. image
  // circulaire rendue en petit rectangle dans le viewer mais OK dans le player qui isole les diapos).
  const uid = 's' + slideNo + '_';
  const rels = slideRels(pptDir, slideNo);
  // --- résout layout + masque + THÈME (le thème dépend du master : cf. loadContext) ---
  const norm = (base, tgt) => join(dirname(base), tgt);
  const layoutTgt = Object.values(rels).find(t => t && t.includes('slideLayout'));
  const layoutPath = layoutTgt ? join(pptDir, 'slides', layoutTgt) : null;
  const layoutRels = layoutPath ? relsOf(layoutPath) : {};
  const masterTgt = Object.values(layoutRels).find(t => t && t.includes('slideMaster'));
  const masterPath = (layoutPath && masterTgt) ? norm(layoutPath, masterTgt) : null;
  const masterRels = masterPath ? relsOf(masterPath) : {};
  const themeTgt = Object.values(masterRels).find(t => t && t.includes('/theme/'));
  const themePath = (masterPath && themeTgt) ? norm(masterPath, themeTgt) : null;
  const ctx = loadContext(pptDir, themePath);
  const { geom: phGeom } = placeholderGeom(pptDir, slideNo, rels);

  const slidePath = join(pptDir, 'slides', `slide${slideNo}.xml`);
  const ordered = P.parse(readFileSync(slidePath, 'utf8'));
  const sld = ordered.find((n) => tagOf(n) === 'p:sld');
  const cSld = first(sld, 'p:cSld');
  const spTree = first(cSld, 'p:spTree');
  // taille de diapo lue dans presentation.xml (gère 16:9, 4:3, etc.)
  let W = 12192000, H = 6858000;
  try {
    const pres = cachedParse(join(pptDir, 'presentation.xml'), 'flat');
    const sz = pres['p:presentation']?.['p:sldSz'];
    if (sz) { W = +sz['@cx'] || W; H = +sz['@cy'] || H; }
  } catch {}

  const body = [];
  const IDENT = { ox:0, oy:0, sx:1, sy:1, cx0:0, cy0:0 };
  let curRels = rels;    // rels du tree en cours (pour résoudre les images)
  let skipPh = false;    // dans masque/layout : ne pas rendre le texte des placeholders
  // préfixe de data-spid : masque/layout et diapo ont chacun leur espace d'id -> collisions (plusieurs
  // formes "data-spid=8"). La timeline d'animation ne cible QUE la diapo ; on préfixe donc les spid du
  // masque/layout ('m'/'l') pour que le player (querySelector par spid) trouve la BONNE forme de diapo.
  let spidPfx = '';

  // (layout/masque/thème déjà résolus en tête de fonction)
  // clrMap du masque : remappe bg1/tx1/bg2/tx2 vers dk1/lt1/... (crucial pour thème sombre)
  // + txStyles : couleurs/tailles par défaut des placeholders (titre/corps) hérités par les runs
  ctx.txStyles = {};
  if (masterPath && existsSync(masterPath)) {
    const mflat = cachedParse(masterPath, 'flat');
    const cm = mflat['p:sldMaster']?.['p:clrMap'];
    if (cm) { ctx.clrMap = {}; for (const [k, v] of Object.entries(cm)) if (k.startsWith('@')) ctx.clrMap[k.slice(1)] = v; }
    const ts = mflat['p:sldMaster']?.['p:txStyles'];
    if (ts) for (const [k, tag] of Object.entries({ title:'p:titleStyle', body:'p:bodyStyle', other:'p:otherStyle' })) {
      const st = ts[tag]; if (!st) continue; ctx.txStyles[k] = {};
      for (let lvl = 1; lvl <= 9; lvl++) {
        const pPr = st[`a:lvl${lvl}pPr`]; if (!pPr) continue;
        const d = pPr['a:defRPr'] || {};
        ctx.txStyles[k][lvl] = {
          sz: d['@sz'] ? +d['@sz']/100 : null, b: d['@b'] === '1',
          color: d['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null,
          font: d['a:latin']?.['@typeface'] || null,
          algn: pPr['@algn'] || null,
          bu: bulletOf(pPr), marL: pPr['@marL'] ? +pPr['@marL'] : null, ind: pPr['@indent'] ? +pPr['@indent'] : null,
          lnSpc: pPr['a:lnSpc'] || null,
          // espacement avant/après paragraphe hérité du master (§21.1.2.2.{10,12}) : les cartes/listes
          // du corpus tirent leur « gap entre paragraphes » d'ici (bodyStyle lvl1 spcBef 10pt), jamais
          // redéfini dans la diapo -> sans héritage, les paragraphes se collent (vu vs PowerPoint).
          spcBef: pPr['a:spcBef'] || null, spcAft: pPr['a:spcAft'] || null,
        };
      }
    }
  }
  // defaultTextStyle de presentation.xml : dernier recours pour le texte hors placeholder
  // (zones de texte libres) — §19.2.1.6. Chargé par niveau, même structure que txStyles.
  ctx.defaultTextStyle = {};
  try {
    const pf = cachedParse(join(pptDir, 'presentation.xml'), 'flat');
    const dts = pf['p:presentation']?.['p:defaultTextStyle'];
    if (dts) for (let lvl = 1; lvl <= 9; lvl++) {
      const pPr = dts[`a:lvl${lvl}pPr`]; if (!pPr) continue;
      const d = pPr['a:defRPr'] || {};
      ctx.defaultTextStyle[lvl] = {
        sz: d['@sz'] ? +d['@sz']/100 : null, b: d['@b'] === '1',
        color: d['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null,
        font: d['a:latin']?.['@typeface'] || null, algn: pPr['@algn'] || null,
      };
    }
  } catch {}

  function partOf(xmlPath, rootTag) {
    if (!xmlPath || !existsSync(xmlPath)) return null;
    const o = cachedParse(xmlPath, 'ord');
    const root = o.find(n => tagOf(n) === rootTag);
    const cS = first(root, 'p:cSld');
    return { spTree: first(cS, 'p:spTree'), cSld: cS };
  }
  function bgRect(cS) {
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
      const p = paint(bgPr);
      if (p) return `<rect width="${u(W)}" height="${u(H)}" fill="${p}"/>`;
    }
    // bgRef : idx>=1001 -> theme bgFillStyleLst[idx-1001], phClr = couleur donnée
    const bgRef = bgFlat['p:bgRef'];
    if (bgRef) {
      const idx = +bgRef['@idx'];
      const phName = bgRef['a:schemeClr']?.['@val'] ?? bgRef['a:srgbClr']?.['@val'];
      const styleFill = ctx.bgFillStyles[idx >= 1000 ? idx - 1001 : 0];
      const p = paint(styleFill, phName);
      if (p) return `<rect width="${u(W)}" height="${u(H)}" fill="${p}"/>`;
    }
    return '';
  }

  function walk(node, tf) {
    if (!node) return;
    for (const child of kids(node)) walkChild(child, tf);
  }
  function walkChild(child, tf) {
    const t = tagOf(child);
    // MCE (Part 3) : une forme peut être enveloppée dans <mc:AlternateContent> (ex. formules OMML
    // sous mc:Choice Requires="a14"). Sans le gérer, la forme + son texte sont INVISIBLES.
    // On rend la branche Fallback (repli sûr) si présente, sinon Choice.
    if (t === 'mc:AlternateContent') {
      // On préfère mc:Choice (contenu AUTHORÉ : texte + formules ; on rend au moins le texte
      // normal, le math OMML est ignoré). Le Fallback des formules est souvent une forme VIDE
      // -> il donnerait une boîte sans texte. Repli sur Fallback si pas de Choice.
      const branch = first(child, 'mc:Choice') || first(child, 'mc:Fallback');
      if (branch) for (const c of kids(branch)) walkChild(c, tf);
      return;
    }
    // masque/layout : on NE rend PAS les placeholders (la diapo en hérite fill/style).
    if (t === 'p:sp') { if (skipPh && deepFind(child, 'p:ph')) return; body.push(renderSp(child, tf)); }
    else if (t === 'p:pic') body.push(renderPic(child, tf));
    else if (t === 'p:cxnSp') body.push(renderSp(child, tf));
    else if (t === 'p:graphicFrame') body.push(renderTable(child, tf));
    else if (t === 'p:grpSp') renderGroup(child, tf);
  }
  function getXfrm(spPrOrdered) {
    const xf = first(spPrOrdered, 'a:xfrm');
    if (!xf) return null;
    const off = first(xf, 'a:off'), ext = first(xf, 'a:ext');
    if (!off || !ext) return null;
    return { x:+at(off)['@x'], y:+at(off)['@y'], cx:+at(ext)['@cx'], cy:+at(ext)['@cy'],
             rot:+(at(xf)['@rot']||0), flipH:at(xf)['@flipH']==='1', flipV:at(xf)['@flipV']==='1' };
  }
  function applyTf(g, tf) { return { x: tf.ox + (g.x - tf.cx0)*tf.sx, y: tf.oy + (g.y - tf.cy0)*tf.sy, cx: g.cx*tf.sx, cy: g.cy*tf.sy, rot:g.rot, flipH:g.flipH, flipV:g.flipV }; }

  function renderGroup(grp, tf) {
    const gPr = first(grp, 'p:grpSpPr');
    const xf = gPr && first(gPr, 'a:xfrm');
    let childTf = tf, wrap = '';
    if (xf) {
      const off=first(xf,'a:off'),ext=first(xf,'a:ext'),chOff=first(xf,'a:chOff'),chExt=first(xf,'a:chExt');
      const g = applyTf({ x:+at(off)['@x'],y:+at(off)['@y'],cx:+at(ext)['@cx'],cy:+at(ext)['@cy'] }, tf);
      const cE = { cx:+at(chExt)['@cx'], cy:+at(chExt)['@cy'] };
      childTf = { ox:g.x, oy:g.y, cx0:+at(chOff)['@x'], cy0:+at(chOff)['@y'], sx:g.cx/cE.cx, sy:g.cy/cE.cy };
      // rotation / miroir du GROUPE (les membres sont placés dans la boîte NON tournée, puis tout
      // le <g> tourne autour du centre du groupe). Sans ça, un groupe tourné rend à plat
      // (ex. manches à air inclinées de VENT d.10).
      const rot = +at(xf)['@rot'] || 0, fH = at(xf)['@flipH']==='1', fV = at(xf)['@flipV']==='1';
      const ccx = u(g.x + g.cx/2), ccy = u(g.y + g.cy/2), parts = [];
      if (rot) parts.push(`rotate(${(rot/60000).toFixed(3)} ${ccx.toFixed(1)} ${ccy.toFixed(1)})`);
      if (fH || fV) parts.push(`translate(${(fH?2*ccx:0).toFixed(1)} ${(fV?2*ccy:0).toFixed(1)}) scale(${fH?-1:1} ${fV?-1:1})`);
      wrap = parts.join(' ');
    }
    // remplissage du GROUPE (hérité par les membres <a:grpFill/>). On le résout et on le propage
    // pendant le walk des enfants, puis on restaure (gère les groupes imbriqués).
    const prevGrpFill = ctx.grpFill;
    if (gPr) {
      const gf = PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(gPr)}</root>`).root['p:grpSpPr'];
      if (gf?.['a:solidFill']) ctx.grpFill = cssColor(resolveColorFlat(gf['a:solidFill'], ctx));
      else if (gf?.['a:gradFill']) ctx.grpFill = gradId(gf['a:gradFill']);
      else if (gf?.['a:noFill'] !== undefined) ctx.grpFill = 'none';
      // sinon (pas de fill explicite ou grpFill imbriqué) : on garde le fill du groupe parent
    }
    // enveloppe les membres dans un <g data-spid=groupId> : le groupe s'anime EN BLOC
    // (une animation ciblant le groupe s'applique à tous ses membres, comme PowerPoint).
    const start = body.length;
    walk(grp, childTf);
    ctx.grpFill = prevGrpFill;
    const inner = body.splice(start).join('');
    const gid = at(deepFind(grp, 'p:cNvPr'))['@id'] || '';
    body.push(`<g data-spid="${spidPfx}${esc(gid)}"${wrap?` transform="${wrap}"`:''}>${inner}</g>`);
  }

  function fillFromSpPrFlat(spFlat, styleFlat) {
    const spPr = spFlat['p:spPr'];
    if (spPr?.['a:noFill'] !== undefined) return 'none';
    // <a:grpFill/> (§20.1.8.35) : la forme reprend le remplissage de son GROUPE parent (ex. régions
    // colorées « fusionnées » : le fill accent3/bg2 est sur le grpSpPr, les membres l'héritent).
    if (spPr?.['a:grpFill'] !== undefined) return ctx.grpFill ?? 'none';
    if (spPr?.['a:solidFill']) return cssColor(resolveColorFlat(spPr['a:solidFill'], ctx));
    if (spPr?.['a:gradFill']) return gradId(spPr['a:gradFill']);
    if (spPr?.['a:pattFill']) return pattId(spPr['a:pattFill']); // remplissage à MOTIF (hachures/points/grille)
    if (spPr?.['a:blipFill']) return null; // géré à part
    // style fillRef : idx (1-based) dans le fillStyleLst du thème, teinté par phClr
    const fr = styleFlat?.['a:fillRef'];
    if (fr) {
      const idx = +fr['@idx'];
      // idx=0 = AUCUN remplissage (la couleur du ref n'est PAS à utiliser) — ex. forme `arc`
      // ouverte (le 45° de VENT d.10) rendue en trait seul. Ne pas remplir.
      if (idx === 0) return 'none';
      const phName = fr['a:schemeClr']?.['@val'] ?? fr['a:srgbClr']?.['@val'];
      if (idx >= 1 && ctx.fillStyles[idx - 1]) { const p = paint(ctx.fillStyles[idx - 1], phName); if (p) return p; }
      const c = resolveColorFlat(fr, ctx); if (c) return cssColor(c); // fallback : couleur du ref
    }
    return 'none';
  }
  // remplissage à MOTIF (a:pattFill prst + fgClr/bgClr) -> tuile <pattern> SVG. §20.1.8.24.
  // Approxime par famille (points pct*, lignes horz/vert, diagonales, grille/croix) : recognizable
  // sans mapper les 54 presets exacts. Absent des decks du corpus -> gain pur pour PowerPoint externes.
  const patterns = [];
  function pattId(pf) {
    const id = uid + 'p' + patterns.length;
    const fg = pf['a:fgClr'] ? cssColor(resolveColorFlat(pf['a:fgClr'], ctx)) : '#000';
    const bg = pf['a:bgClr'] ? cssColor(resolveColorFlat(pf['a:bgClr'], ctx)) : '#fff';
    patterns.push({ id, fg, bg, prst: pf['@prst'] || 'pct50' });
    return `url(#${id})`;
  }
  const grads = [];
  function gradId(gradFlat) {
    const id = uid + 'g' + grads.length;
    const gs = gradFlat['a:gsLst']?.['a:gs'];
    // stops TRIÉS par position croissante : SVG l'exige (sinon Chrome clampe les offsets décroissants ->
    // dégradé faux, ex. sol de VENT d.35 rouge au lieu de vert). PowerPoint accepte l'ordre du document.
    const stops = (Array.isArray(gs)?gs:gs?[gs]:[]).map(s => ({ pos:(+s['@pos'])/1000, col: cssColor(resolveColorFlat(s, ctx)) }))
      .sort((a, b) => a.pos - b.pos);
    const ang = ((gradFlat['a:lin']?.['@ang']||0)/60000) * Math.PI/180;
    const radial = !!gradFlat['a:path']; // dégradé radial/rectangulaire
    grads.push({ id, stops, ang, radial });
    return `url(#${id})`;
  }
  // trait pointillé : prstDash -> stroke-dasharray (en points, relatif à l'épaisseur)
  function dashArray(val, swp) {
    if (!val || val === 'solid') return '';
    const s = swp || 1;
    return ({ dot:`${s} ${s*2}`, sysDot:`${s} ${s}`, dash:`${s*4} ${s*3}`, sysDash:`${s*3} ${s}`,
      lgDash:`${s*8} ${s*3}`, dashDot:`${s*4} ${s*3} ${s} ${s*3}`, lgDashDot:`${s*8} ${s*3} ${s} ${s*3}`,
      sysDashDot:`${s*3} ${s} ${s} ${s}`, lgDashDotDot:`${s*8} ${s*3} ${s} ${s*3} ${s} ${s*3}` }[val]) || '';
  }
  // clips d'images découpées en formes (cercle, roundRect, custGeom)
  const clips = [];
  // pointes de flèche des lignes/connecteurs (marker SVG partagé par type ; couleur = trait via
  // context-stroke ; orient=auto-start-reverse -> même marker pour début et fin).
  const lineMarkers = new Map();
  // dimensions en multiples de l'épaisseur du trait (markerUnits=strokeWidth) selon @len/@w (sm/med/lg).
  // PowerPoint « medium » ≈ 3-3,5× l'épaisseur (pas 8× — sinon pointes énormes, cf. flèches VENT d.28).
  function lineMarker(type, end) {
    const SZ = { sm: 2, med: 3.3, lg: 4.8 };
    const w = SZ[end?.['@len']] ?? 3.3, h = SZ[end?.['@w']] ?? 3.3;   // len = longueur (X), w = largeur (Y)
    const key = (type + '_' + w + '_' + h).replace(/\./g, 'p');
    if (!lineMarkers.has(key)) {
      const cx = (w / 2).toFixed(2), cy = (h / 2).toFixed(2), W = w.toFixed(2), H = h.toFixed(2);
      let shape;
      if (type === 'oval') shape = `<ellipse cx="${cx}" cy="${cy}" rx="${(w/2*0.95).toFixed(2)}" ry="${(h/2*0.95).toFixed(2)}" fill="context-stroke"/>`;
      else if (type === 'diamond') shape = `<path d="M0 ${cy} L${cx} 0 L${W} ${cy} L${cx} ${H} z" fill="context-stroke"/>`;
      else if (type === 'stealth') shape = `<path d="M0 0 L${W} ${cy} L0 ${H} L${(w*0.32).toFixed(2)} ${cy} z" fill="context-stroke"/>`;
      else shape = `<path d="M0 ${(h*0.06).toFixed(2)} L${W} ${cy} L0 ${(h*0.94).toFixed(2)} z" fill="context-stroke"/>`; // triangle/arrow
      lineMarkers.set(key, `<marker id="${uid}mk_${key}" markerWidth="${W}" markerHeight="${H}" refX="${(w*0.9).toFixed(2)}" refY="${cy}" orient="auto-start-reverse" markerUnits="strokeWidth">${shape}</marker>`);
    }
    return uid + 'mk_' + key;
  }
  // effets a:effectLst -> filtre SVG combiné (renvoie ' filter="url(#..)"'). §20.1.8
  // Gère : ombre portée (outerShdw), halo lumineux (glow), bords adoucis (softEdge).
  const filters = [];
  function shadowFilter(spPr, style) { return effectFilter(spPr, style); } // compat appel existant
  function effectFilter(spPr, style) {
    // effets EXPLICITES de la forme, sinon repli sur le style de thème via <p:style><a:effectRef idx>=1>
    // (§20.1.4.2.8) : l'ombre/lueur "thème" des formes stylées (SmartArt, cartes) vient de là.
    let fx = spPr?.['a:effectLst'];
    if (!fx) {
      const er = style?.['a:effectRef']; const eidx = er ? +er['@idx'] : 0;
      if (eidx >= 1) fx = ctx.effectStyles?.[eidx - 1]?.['a:effectStyle']?.['a:effectLst'];
    }
    if (!fx) return '';
    const sh = fx['a:outerShdw'], gl = fx['a:glow'], se = fx['a:softEdge'];
    if (!sh && !gl && !se) return '';
    const prims = []; const layers = []; let pad = 30;
    // bords adoucis : flou du canal alpha puis re-découpe -> source adoucie ("src")
    let src = 'SourceGraphic';
    if (se) { const r = u(+se['@rad'] || 0) / 2;
      prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${r.toFixed(1)}" result="seA"/>`
        + `<feComposite in="SourceGraphic" in2="seA" operator="in" result="src"/>`);
      src = 'src';
    }
    // ombre portée (derrière)
    if (sh) {
      const dist = u(+sh['@dist'] || 0), dir = (+sh['@dir'] || 0) / 60000 * Math.PI/180;
      const blur = u(+sh['@blurRad'] || 40000), c = resolveColorFlat(sh, ctx) || { r:0,g:0,b:0,a:0.4 };
      pad = Math.max(pad, 40 + blur);
      prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${(blur/2).toFixed(1)}" result="shB"/>`
        + `<feOffset in="shB" dx="${(dist*Math.cos(dir)).toFixed(1)}" dy="${(dist*Math.sin(dir)).toFixed(1)}" result="shO"/>`
        + `<feFlood flood-color="rgb(${c.r},${c.g},${c.b})" flood-opacity="${(c.a??0.4).toFixed(2)}"/>`
        + `<feComposite in2="shO" operator="in" result="shadow"/>`);
      layers.push('shadow');
    }
    // halo lumineux (derrière la source, devant l'ombre)
    if (gl) {
      const r = u(+gl['@rad'] || 0) / 2, c = resolveColorFlat(gl, ctx) || { r:255,g:255,b:0,a:0.6 };
      pad = Math.max(pad, 40 + r * 2);
      prims.push(`<feMorphology in="SourceAlpha" operator="dilate" radius="1" result="glD"/>`
        + `<feGaussianBlur in="glD" stdDeviation="${r.toFixed(1)}" result="glB"/>`
        + `<feFlood flood-color="rgb(${c.r},${c.g},${c.b})" flood-opacity="${(c.a??0.6).toFixed(2)}"/>`
        + `<feComposite in2="glB" operator="in" result="glow"/>`);
      layers.push('glow', 'glow'); // doublé = halo plus dense (comme PowerPoint)
    }
    layers.push(src);
    const merge = `<feMerge>${layers.map((n) => `<feMergeNode in="${n}"/>`).join('')}</feMerge>`;
    const m = Math.min(200, Math.max(30, Math.round(pad)));
    const id = uid + 'fx' + filters.length;
    filters.push(`<filter id="${id}" x="-${m}%" y="-${m}%" width="${100+2*m}%" height="${100+2*m}%">${prims.join('')}${merge}</filter>`);
    return ` filter="url(#${id})"`;
  }
  // effets couleur/alpha d'une image (enfants de a:blip / blipFill) -> filtre SVG. §20.1.8.{20,25,33,4}
  function imageFilter(blip) {
    if (!blip) return { attr: '', opacity: '' };
    const prims = []; let opacity = '';
    const amf = blip['a:alphaModFix']; // opacité globale (ST_PositivePercentage)
    if (amf?.['@amt']) opacity = ` opacity="${(+amf['@amt']/100000).toFixed(3)}"`;
    if (blip['a:grayscl'] !== undefined) prims.push('<feColorMatrix type="saturate" values="0"/>');
    const lum = blip['a:lum']; // §20.1.8.42 luminance : bright (additif) + contrast (autour du gris moyen)
    if (lum && (lum['@bright'] || lum['@contrast'])) {
      const br = (+lum['@bright'] || 0) / 100000, ct = (+lum['@contrast'] || 0) / 100000;
      const slope = 1 + ct, inter = br + 0.5 * (1 - slope);
      const fn = (c) => `<feFunc${c} type="linear" slope="${slope.toFixed(3)}" intercept="${inter.toFixed(3)}"/>`;
      prims.push(`<feComponentTransfer>${fn('R')}${fn('G')}${fn('B')}</feComponentTransfer>`);
    }
    const duo = blip['a:duotone']; // bichromie : 2 couleurs de a:duotone -> remap luminance
    if (duo) {
      const cols = [];
      for (const k of Object.keys(duo)) if (/Clr$/.test(k)) { const v = duo[k]; for (const c of (Array.isArray(v)?v:[v])) cols.push(resolveColorFlat({ [k]:c }, ctx)); }
      const [c0, c1] = cols.length >= 2 ? cols : [{r:0,g:0,b:0}, {r:255,g:255,b:255}];
      prims.push('<feColorMatrix type="matrix" values="0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0 0 0 1 0"/>'
        + `<feComponentTransfer><feFuncR type="table" tableValues="${(c0.r/255).toFixed(3)} ${(c1.r/255).toFixed(3)}"/>`
        + `<feFuncG type="table" tableValues="${(c0.g/255).toFixed(3)} ${(c1.g/255).toFixed(3)}"/>`
        + `<feFuncB type="table" tableValues="${(c0.b/255).toFixed(3)} ${(c1.b/255).toFixed(3)}"/></feComponentTransfer>`);
    }
    // corrections photo MS 2010 (extension a14 dans blip>extLst) : luminosité/contraste, saturation.
    // §MS-ODRAWXML a14:imgEffect. Valeurs en millièmes de % (ST_(Fixed)Percentage).
    // ⚠ Si un <a14:imgLayer r:embed=.wdp> est présent, ces corrections décrivent l'édition Office
    // appliquée à la HD photo ORIGINALE (.wdp) pour PRODUIRE l'image de base (rId2) que l'on rend déjà :
    // elles sont donc DÉJÀ cuites dans rId2. Les ré-appliquer écraserait la photo (ex. bright -40% →
    // soustraction de 102/canal → tout en noir). On ne les applique que si elles sont « directes ».
    const baked = !!deepFindFlat(blip, 'a14:imgLayer');
    const bc = baked ? null : deepFindFlat(blip, 'a14:brightnessContrast');
    if (bc && (bc['@bright'] || bc['@contrast'])) {
      const br = (+bc['@bright'] || 0) / 100000, ct = (+bc['@contrast'] || 0) / 100000;
      const slope = 1 + ct, inter = br + 0.5 * (1 - slope); // contraste autour du gris moyen + luminosité
      const fn = (c) => `<feFunc${c} type="linear" slope="${slope.toFixed(3)}" intercept="${inter.toFixed(3)}"/>`;
      prims.push(`<feComponentTransfer>${fn('R')}${fn('G')}${fn('B')}</feComponentTransfer>`);
    }
    const sat = baked ? null : deepFindFlat(blip, 'a14:saturation');
    if (sat && sat['@sat'] != null) prims.push(`<feColorMatrix type="saturate" values="${(+sat['@sat']/100000).toFixed(3)}"/>`);
    if (!prims.length) return { attr: '', opacity };
    const id = uid + 'img' + filters.length;
    // color-interpolation-filters="sRGB" : opérer en sRGB comme PowerPoint. Sinon les filtres SVG
    // tournent en linearRGB (défaut) → luminosité/contraste écrasent tout vers le noir.
    filters.push(`<filter id="${id}" color-interpolation-filters="sRGB">${prims.join('')}</filter>`);
    return { attr: ` filter="url(#${id})"`, opacity };
  }
  // peinture générique d'un noeud de remplissage (solid/grad) avec substitution phClr
  function paint(fillFlat, phClrName) {
    if (!fillFlat) return null;
    const prev = ctx.phClr; ctx.phClr = phClrName || ctx.phClr;
    let out = null;
    if (fillFlat['a:solidFill']) out = cssColor(resolveColorFlat(fillFlat['a:solidFill'], ctx));
    else if (fillFlat['a:gradFill']) out = gradId(fillFlat['a:gradFill']);
    ctx.phClr = prev;
    return out;
  }

  function renderSp(sp, tf, suppressText) {
    // version flat du même noeud pour lire fills/geom/txBody
    const flat = PF.parse(`<root xmlns:a="a" xmlns:p="p">${serialize(sp)}</root>`).root;
    const spFlat = flat['p:sp'] || flat['p:cxnSp'] || flat;
    const spPrO = first(sp, 'p:spPr') || first(sp, 'p:spPr');
    // placeholder : héritage depuis le layout (géométrie + fill + styles de texte)
    const ph = deepFind(sp, 'p:ph');
    const key = ph ? `${at(ph)['@type']??'body'}:${at(ph)['@idx']??'0'}` : null;
    const inh = key ? phGeom[key] : null;
    let g = getXfrm(spPrO);
    if (!g && inh?.xfrm) { const xf=inh.xfrm, off=xf['a:off'], ext=xf['a:ext']; if (off&&ext) g = { x:+off['@x'],y:+off['@y'],cx:+ext['@cx'],cy:+ext['@cy'],rot:+(xf['@rot']||0) }; }
    if (!g) return '';
    const m = applyTf(g, tf);
    let geomFlat = spFlat['p:spPr'] || {};
    // géométrie héritée du placeholder de layout si la diapo n'en définit pas
    if (!geomFlat['a:prstGeom'] && !geomFlat['a:custGeom'] && inh?.spPr) geomFlat = inh.spPr;
    const prst = geomFlat['a:prstGeom']?.['@prst'];
    const custO = geomFlat['a:custGeom'] ? deepFind(sp, 'a:custGeom') : null;
    let d, dStroke = '';
    if (custO) { const p = custGeomParts(custO, u(m.cx), u(m.cy)); d = p.fill; dStroke = p.stroke; }
    else {
      const adj = {}; const gd = geomFlat['a:prstGeom']?.['a:avLst']?.['a:gd'];
      for (const x of (Array.isArray(gd)?gd:gd?[gd]:[])) adj[x['@name']] = x['@fmla'];
      const p = presetGeomParts(prst || 'rect', u(m.cx), u(m.cy), adj);
      if (p) { d = p.fill; dStroke = p.stroke; } else { d = `M0 0 H${u(m.cx)} V${u(m.cy)} H0 Z`; dStroke = d; }
    }
    let fill = fillFromSpPrFlat(spFlat, spFlat['p:style']);
    // fill hérité du placeholder de layout (fond de carte, cercle de badge)
    if (fill === 'none' && inh) fill = fillFromSpPrFlat({ 'p:spPr': inh.spPr, 'p:style': inh.style }, inh.style);
    const ln = geomFlat['a:ln'] ?? inh?.spPr?.['a:ln'];
    let stroke = 'none', sw = 0, dash = '', markers = '', joinCap = '';
    if (ln && ln['a:noFill'] === undefined) {
      sw = Number(ln['@w']) || 12700;
      stroke = ln['a:solidFill'] ? cssColor(resolveColorFlat(ln['a:solidFill'], ctx))
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
      if (he) markers += ` marker-start="url(#${lineMarker(he, ln['a:headEnd'])})"`;
      if (te) markers += ` marker-end="url(#${lineMarker(te, ln['a:tailEnd'])})"`;
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
        if (he) markers += ` marker-start="url(#${lineMarker(he, themeLn['a:headEnd'])})"`;
        if (te) markers += ` marker-end="url(#${lineMarker(te, themeLn['a:tailEnd'])})"`;
      }
    }
    // forme remplie par une IMAGE (blipFill) : image découpée à la forme
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
    const filt = shadowFilter(spFlat['p:spPr'], spFlat['p:style'] || inh?.style);
    const rot = m.rot ? ` rotate(${(m.rot/60000).toFixed(2)} ${u(m.cx/2)} ${u(m.cy/2)})` : '';
    const sc = (m.flipH||m.flipV) ? ` translate(${u(m.flipH?m.cx:0)} ${u(m.flipV?m.cy:0)}) scale(${m.flipH?-1:1} ${m.flipV?-1:1})` : '';
    // phKind : title/body seulement pour un VRAI placeholder ; une zone de texte LIBRE (sans p:ph)
    // = 'none' -> n'hérite PAS du style body du master (sinon puces « • » et couleurs parasites) ;
    // elle s'appuie sur defaultTextStyle + couleur de thème.
    const phType = at(ph)['@type'];
    const phKind = !ph ? 'none' : (phType === 'title' || phType === 'ctrTitle') ? 'title' : 'body';
    const txt = suppressText ? '' : renderText(spFlat['p:txBody'], m.cx, m.cy, phKind, inh?.lstStyle, first(sp, 'p:txBody'), inh?.bodyPr, spFlat['p:style'] || inh?.style);
    const spid = at(deepFind(sp, 'p:cNvPr'))['@id'] || '';
    const strokeAttr = stroke !== 'none' ? `stroke="${stroke}" stroke-width="${u(sw)}"${dash?` stroke-dasharray="${dash}"`:''}${joinCap}${markers}` : '';
    // Tracé sur le même chemin que le remplissage (cas courant) OU chemin de tracé distinct (formes
    // arc/chord… : le remplissage est un secteur, le tracé n'est que la courbe -> pas de rayons parasites).
    const sameGeom = !dStroke || dStroke === d;
    const paths = `<path d="${d}" fill="${fill}" ${sameGeom ? strokeAttr : ''}/>`
      + (!sameGeom && strokeAttr ? `<path d="${dStroke}" fill="none" ${strokeAttr}/>` : '');
    return `<g data-spid="${spidPfx}${esc(spid)}" transform="translate(${u(m.x)} ${u(m.y)})${rot}"${filt}>`
      + `<g transform="${sc.trim()||'translate(0 0)'}">`
      + paths
      + picLayer
      + `</g>` + txt + `</g>`;
  }

  // graphique (p:graphicFrame > c:chart) : lit la part ppt/charts/chartN.xml (réf r:id) et la rend en SVG
  // via chart-render.mjs. Fond blanc (défaut PowerPoint) + palette = accents du thème.
  function renderChartFrame(chartRef, off, ext, tf) {
    const rId = chartRef['@r:id'] || chartRef['@r:embed'];
    const tgt = rId && curRels[rId];
    if (!tgt) return '';
    const chartPath = join(pptDir, tgt.replace(/^\.\.\//, ''));
    if (!existsSync(chartPath)) return '';
    let chartFlat; try { chartFlat = cachedParse(chartPath, 'flat'); } catch { return ''; }
    const g = applyTf({ x:+off['@x'], y:+off['@y'], cx:+ext['@cx'], cy:+ext['@cy'], rot:0 }, tf);
    const W2 = u(g.cx), H2 = u(g.cy);
    const palette = ['accent1','accent2','accent3','accent4','accent5','accent6'].map((a) => ctx.clr[a]).filter(Boolean);
    const color = (sp) => sp?.['a:solidFill'] ? cssColor(resolveColorFlat(sp['a:solidFill'], ctx)) : null;
    let inner; try { inner = renderChart(chartFlat, W2, H2, { palette, color }); } catch { return ''; }
    if (!inner) return '';
    // fond blanc du graphique (sauf si c:chartSpace définit noFill) puis contenu
    const cs = chartFlat['c:chartSpace'] || chartFlat;
    const bg = cs['c:spPr']?.['a:noFill'] !== undefined ? '' : `<rect width="${W2.toFixed(1)}" height="${H2.toFixed(1)}" fill="#ffffff"/>`;
    return `<g transform="translate(${u(g.x).toFixed(1)} ${u(g.y).toFixed(1)})">${bg}${inner}</g>`;
  }
  // SmartArt : la part de dessin de secours dsp:drawing (ppt/diagrams/drawingN.xml) contient un spTree
  // de vraies formes DrawingML, structurellement identique à un p:spTree mais préfixé « dsp: ». On
  // remplace « dsp: » -> « p: » et on réutilise le walk normal, avec un mapping repère-dessin -> cadre.
  function renderDiagram(relId, off, ext, tf) {
    const tgt = curRels[relId]; if (!tgt) return;
    const p = join(pptDir, tgt.replace(/^\.\.\//, ''));
    if (!existsSync(p)) return;
    let txt; try { txt = readFileSync(p, 'utf8'); } catch { return; }
    txt = txt.replace(/dsp:/g, 'p:'); // même structure que PresentationML (nvSpPr/spPr/txBody/style + a:*)
    let ord; try { ord = P.parse(txt); } catch { return; }
    const root = ord.find((n) => tagOf(n) === 'p:drawing');
    const spTree = root && deepFind(root, 'p:spTree'); if (!spTree) return;
    const frame = applyTf({ x:+off['@x'], y:+off['@y'], cx:+ext['@cx'], cy:+ext['@cy'], rot:0 }, tf);
    // mapping : chOff/chExt du grpSpPr du spTree (repère du dessin) -> cadre placé
    const gpr = first(spTree, 'p:grpSpPr'); const xf = gpr && first(gpr, 'a:xfrm');
    let dtf = { ox: frame.x, oy: frame.y, cx0: 0, cy0: 0, sx: 1, sy: 1 };
    if (xf) {
      const chOff = first(xf, 'a:chOff'), chExt = first(xf, 'a:chExt');
      const cx = +at(chExt)['@cx'] || frame.cx, cy = +at(chExt)['@cy'] || frame.cy;
      dtf = { ox: frame.x, oy: frame.y, cx0: +at(chOff)['@x'] || 0, cy0: +at(chOff)['@y'] || 0, sx: frame.cx / cx, sy: frame.cy / cy };
    }
    walk(spTree, dtf);
  }
  // tableau (p:graphicFrame > a:tbl) : grille + bordures + fusions + texte
  function renderTable(gf, tf) {
    const arr = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];
    const flat = PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(gf)}</root>`).root['p:graphicFrame'];
    if (!flat) return '';
    const xf = flat['p:xfrm']; const off = xf?.['a:off'], ext = xf?.['a:ext'];
    const gd = flat['a:graphic']?.['a:graphicData'];
    // un p:graphicFrame porte un TABLEAU (a:tbl) OU un GRAPHIQUE (c:chart, réf. r:id vers ppt/charts/…)
    const chartRef = gd?.['c:chart'] || gd?.['cx:chart'];
    if (off && ext && chartRef) return renderChartFrame(chartRef, off, ext, tf);
    // SmartArt (dgm:relIds) : rendu de secours dsp:drawing référencé par dsp:dataModelExt@relId
    const dgm = gd?.['dgm:relIds'];
    if (off && ext && dgm) {
      const exts = arr(dgm['a:extLst']?.['a:ext']);
      let relId; for (const e of exts) { const dm = e['dsp:dataModelExt']; if (dm?.['@relId']) { relId = dm['@relId']; break; } }
      if (relId) renderDiagram(relId, off, ext, tf); // pousse directement dans body
      return '';
    }
    const tbl = gd?.['a:tbl'];
    if (!off || !tbl) {
      // repli générique : un graphicFrame non-tableau/graphique/SmartArt (objet OLE incorporé, contrôle…)
      // embarque presque toujours une IMAGE DE SECOURS (p:pic). La rendre plutôt que laisser un vide.
      if (off && ext) { const picO = deepFind(gf, 'p:pic'); if (picO) return renderPic(picO, tf); }
      return '';
    }
    const g = applyTf({ x:+off['@x'], y:+off['@y'], cx:+ext['@cx'], cy:+ext['@cy'], rot:0 }, tf);
    const cols = arr(tbl['a:tblGrid']?.['a:gridCol']).map(c => +c['@w']);
    const rows = arr(tbl['a:tr']);
    // hauteurs de lignes : spécifiées sinon réparties sur la hauteur du cadre
    let hs = rows.map(r => +r['@h'] || 0);
    const sumH = hs.reduce((a,b)=>a+b,0), zero = hs.filter(h=>!h).length;
    if (zero) { const rem = Math.max(0, (+ext['@cy']) - sumH) / zero; hs = hs.map(h => h || rem); }
    const colX = [0]; cols.forEach((w,i)=>colX[i+1]=colX[i]+w);
    const rowY = [0]; hs.forEach((h,i)=>rowY[i+1]=rowY[i]+h);
    // style de tableau (fills header/bandes) — best-effort
    const styleId = tbl['a:tblPr']?.['a:tableStyleId'];
    const tp = tbl['a:tblPr'] || {};
    const ts = styleId && ctx.tableStyles?.[styleId];
    // bordure PAR DÉFAUT des styles « grille » INTÉGRÉS (non stockés dans le fichier) : {5940675A...} =
    // « No Style, Table Grid » = toutes bordures (couleur tx1, 1pt). Sinon les cellules à <a:tcPr/> vide
    // n'avaient aucune grille (ex. d.46). Utilisée quand la cellule ne définit pas sa propre bordure.
    const GRID_STYLES = new Set(['{5940675A-B579-460E-94D1-54222C63F5DA}', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}']);
    const dfltB = GRID_STYLES.has(styleId) ? { '@w': '12700', 'a:solidFill': { 'a:schemeClr': { '@val': 'tx1' } } } : null;
    let out = '';
    const sx = g.cx / (colX[cols.length] || 1), sy = g.cy / (rowY[rows.length] || 1); // ajuste à la taille réelle
    rows.forEach((tr, ri) => {
      arr(tr['a:tc']).forEach((tc, ci) => {
        if (tc['@hMerge'] === '1' || tc['@vMerge'] === '1') return; // cellule fusionnée (couverte)
        const span = +tc['@gridSpan'] || 1, rspan = +tc['@rowSpan'] || 1;
        const cx = g.x + colX[ci] * sx, cy = g.y + rowY[ri] * sy;
        const cw = ((colX[Math.min(ci+span, cols.length)] - colX[ci])) * sx;
        const ch = ((rowY[Math.min(ri+rspan, rows.length)] - rowY[ri])) * sy;
        const tcPr = tc['a:tcPr'] || {};
        // fill : explicite, sinon style (firstRow / bandes)
        let cellFill = null;
        if (tcPr['a:noFill'] !== undefined) cellFill = null;
        else if (tcPr['a:solidFill'] || tcPr['a:gradFill']) cellFill = paint(tcPr);
        else if (ts) {
          if (ri === 0 && tp['@firstRow'] !== '0' && ts.firstRow) cellFill = paint(ts.firstRow, ts.firstRowPh);
          else if (ts.band1 && (ri % 2 === 1)) cellFill = paint(ts.band1, ts.band1Ph);
          else if (ts.band2 && ri % 2 === 0 && ri > 0) cellFill = paint(ts.band2, ts.band2Ph);
        }
        if (cellFill && cellFill !== 'none')
          out += `<rect x="${u(cx)}" y="${u(cy)}" width="${u(cw)}" height="${u(ch)}" fill="${cellFill}"/>`;
        // bordures
        const border = (ln, x1, y1, x2, y2) => {
          if (!ln || ln['a:noFill'] !== undefined) return '';
          const w = +ln['@w'] || 12700;
          const col = ln['a:solidFill'] ? cssColor(resolveColorFlat(ln['a:solidFill'], ctx)) : '#888';
          return `<line x1="${u(x1)}" y1="${u(y1)}" x2="${u(x2)}" y2="${u(y2)}" stroke="${col}" stroke-width="${u(w)}"/>`;
        };
        out += border(tcPr['a:lnT'] || dfltB, cx, cy, cx+cw, cy);
        out += border(tcPr['a:lnB'] || dfltB, cx, cy+ch, cx+cw, cy+ch);
        out += border(tcPr['a:lnL'] || dfltB, cx, cy, cx, cy+ch);
        out += border(tcPr['a:lnR'] || dfltB, cx+cw, cy, cx+cw, cy+ch);
        // texte (ancrage vertical depuis tcPr @anchor)
        const tb = tc['a:txBody'];
        if (tb) {
          tb['a:bodyPr'] = tb['a:bodyPr'] || {};
          if (!tb['a:bodyPr']['@anchor'] && tcPr['@anchor']) tb['a:bodyPr']['@anchor'] = tcPr['@anchor'];
          // 'none' : une cellule de tableau N'hérite PAS du style « corps » du master (souvent gras/gros
          // pour ces decks) — sinon texte trop gras/grand qui déborde. Elle utilise le défaut neutre + le
          // rPr propre de chaque run (ex. b="1" sur la 1ʳᵉ colonne seulement).
          const txt = renderText(tb, cw, ch, 'none');
          if (txt) out += `<g transform="translate(${u(cx)} ${u(cy)})">${txt}</g>`;
        }
      });
    });
    const gid = flat['p:nvGraphicFramePr']?.['p:cNvPr']?.['@id'] || '';
    return `<g data-spid="${spidPfx}${esc(gid)}">${out}</g>`;
  }

  function renderPic(pic, tf) {
    // placeholder image : géométrie (forme de découpe) héritée du layout
    const ph = deepFind(pic, 'p:ph');
    const key = ph ? `${at(ph)['@type']??'body'}:${at(ph)['@idx']??'0'}` : null;
    const inh = key ? phGeom[key] : null;
    const spPrO = first(pic, 'p:spPr');
    let g = getXfrm(spPrO);
    if (!g && inh?.xfrm) { const xf=inh.xfrm, o=xf['a:off'], e=xf['a:ext']; if (o&&e) g={ x:+o['@x'],y:+o['@y'],cx:+e['@cx'],cy:+e['@cy'],rot:+(xf['@rot']||0) }; }
    if (!g) return '';
    const m = applyTf(g, tf);
    // priorité au SVG moderne (asvg:svgBlip), sinon le blip raster (png/jpeg/wdp)
    const blip = deepFind(pic, 'a:blip');
    const svgB = deepFindSuffix(pic, 'svgBlip');
    const candidates = [at(svgB)['@r:embed'], at(blip)['@r:embed'], at(blip)['@embed']].filter(Boolean);
    let href = null;
    for (const id of candidates) { const tgt = curRels[id]; if (tgt) { href = imageHref(pptDir, tgt, embed); if (href) break; } }
    if (!href) return ''; // image non exploitable (ex. .wdp seul) -> ne rien afficher
    // forme de découpe : géométrie propre du pic, sinon celle héritée du placeholder de layout
    const picFlat = PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(pic)}</root>`).root['p:pic'];
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
    const imgF = imageFilter(picFlat['p:blipFill']?.['a:blip']); // grayscl/duotone/alphaModFix
    const effF = effectFilter(picFlat['p:spPr'], picFlat['p:style']); // glow/ombre/softEdge (+ effectRef)
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
      const vw = Math.max(1e-3, 1-l-r), vh = Math.max(1e-3, 1-t-b);
      iw = bw/vw; ih = bh/vh; ix = -l*iw; iy = -t*ih; par = 'none';
      if (!clipAttr) { const cid=uid+'clip'+clips.length;
        clips.push(`<clipPath id="${cid}"><rect x="0" y="0" width="${bw.toFixed(1)}" height="${bh.toFixed(1)}"/></clipPath>`);
        clipAttr = ` clip-path="url(#${cid})"`; }
    }
    const spid = at(deepFind(pic, 'p:cNvPr'))['@id'] || '';
    return `<g data-spid="${spidPfx}${esc(spid)}" transform="translate(${u(m.x)} ${u(m.y)})${rot}${flip}"${clipAttr}${effF}>`
      + `<image x="${ix.toFixed(1)}" y="${iy.toFixed(1)}" width="${iw.toFixed(1)}" height="${ih.toFixed(1)}" href="${href}" preserveAspectRatio="${par}"${imgF.attr}${imgF.opacity}/></g>`;
  }

  function renderText(txBody, w, h, phKind, inhLst, txBodyO, inhBodyPr, shapeStyle) {
    if (!txBody) return '';
    const arr = (x)=> x===undefined?[]:Array.isArray(x)?x:[x];
    const bodyPr = txBody['a:bodyPr'] || {};
    // bodyPr hérité du placeholder de LAYOUT (insets + ancre) : une carte-placeholder avec
    // <a:bodyPr/> vide DOIT reprendre le tIns/anchor du layout (ex. tIns=288000 réserve la place
    // de la pastille). Priorité : valeur de la diapo → valeur héritée → défaut OOXML.
    const inhBP = inhBodyPr || {};
    const pick = (k, def) => bodyPr[k] !== undefined ? +bodyPr[k] : (inhBP[k] !== undefined ? +inhBP[k] : def);
    const anchor = bodyPr['@anchor'] ?? inhBP['@anchor']; // t/ctr/b
    const just = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
    const iL = pick('@lIns',91440), iR=pick('@rIns',91440), iT=pick('@tIns',45720), iB=pick('@bIns',45720);
    const styleLvls = ctx.txStyles?.[phKind] || {};
    // couleur de dernier recours = tx1 du thème (clair en thème sombre) — jamais #000 en dur
    // (sinon texte noir invisible sur fond sombre, ex. pied de page).
    const DEFCOL = cssColor(resolveColorFlat({ 'a:schemeClr': { '@val': 'tx1' } }, ctx)) || '#000';
    // autofit : PowerPoint réduit la police (et l'interligne) pour faire tenir le texte
    const naf = bodyPr['a:normAutofit'];
    const fontScale = naf && naf['@fontScale'] ? +naf['@fontScale']/100000 : 1;
    const lnRed = naf && naf['@lnSpcReduction'] ? +naf['@lnSpcReduction']/100000 : 0;
    const lineH = Math.max(0.8, 1.15 * (1 - lnRed));
    // +mn-lt/+mj-lt = polices du thème (mineure/majeure)
    const resolveFont = (tf) => !tf ? null
      : /^\+mn-/.test(tf) ? ctx.minorFont : /^\+mj-/.test(tf) ? ctx.majorFont : tf;
    // style par niveau, du + prioritaire au - : lstStyle INLINE de la forme -> lstStyle du
    // placeholder de LAYOUT hérité -> txStyles du MASQUE.
    const localLst = txBody['a:lstStyle'] || {};
    const inheritLst = inhLst || {};
    const lvlOf = (lst, lvl) => lst[`a:lvl${lvl}pPr`] || lst['a:lvl1pPr'];
    const colOf = (d) => d?.['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null;
    // dernier recours : defaultTextStyle de la présentation (texte hors placeholder)
    const dflt = ctx.defaultTextStyle || {};
    // style de FORME <p:style><a:fontRef idx="minor|major"><couleur/></a:fontRef> : donne au texte de la
    // forme sa couleur ET sa police de thème PAR DÉFAUT (souvent la SEULE source de couleur d'une forme
    // thémée non-placeholder, ex. bloc/callout SmartArt). Injecté en DERNIER repli (après lstStyle/master/
    // defaultTextStyle, avant DEFCOL) -> n'agit QUE là où le texte n'a aucune couleur/police définie.
    const fref = shapeStyle?.['a:fontRef'];
    const frefRgb = fref ? resolveColorFlat(fref, ctx) : null;
    const fontRefColor = frefRgb ? cssColor(frefRgb) : null;
    const fontRefFont = fref?.['@idx'] === 'major' ? ctx.majorFont : fref?.['@idx'] === 'minor' ? ctx.minorFont : null;
    const lvlStyle = (lvl) => {
      const loc = lvlOf(localLst, lvl), inh = lvlOf(inheritLst, lvl);
      const ld = loc?.['a:defRPr'], id = inh?.['a:defRPr'];
      const mas = styleLvls[lvl] || styleLvls[1] || {};
      const dt = dflt[lvl] || dflt[1] || {};
      return {
        sz: (ld?.['@sz'] ? +ld['@sz']/100 : null) ?? (id?.['@sz'] ? +id['@sz']/100 : null) ?? mas.sz ?? dt.sz,
        color: colOf(ld) ?? colOf(id) ?? mas.color ?? dt.color ?? fontRefColor,
        font: ld?.['a:latin']?.['@typeface'] || id?.['a:latin']?.['@typeface'] || mas.font || dt.font || fontRefFont,
        // gras : le niveau le PLUS SPÉCIFIQUE qui définit @b gagne (un b="0" du layout doit battre
        // le b="1" du master) — un simple || laisserait le master l'emporter à tort.
        b: ld?.['@b'] !== undefined ? ld['@b'] === '1'
         : id?.['@b'] !== undefined ? id['@b'] === '1'
         : mas.b,
        algn: loc?.['@algn'] || inh?.['@algn'] || mas.algn,
        bu: bulletOf(loc) ?? bulletOf(inh) ?? mas.bu,
        marL: (loc?.['@marL'] ? +loc['@marL'] : null) ?? (inh?.['@marL'] ? +inh['@marL'] : null) ?? mas.marL,
        ind: (loc?.['@indent'] ? +loc['@indent'] : null) ?? (inh?.['@indent'] ? +inh['@indent'] : null) ?? mas.ind,
        // interligne hérité (lstStyle inline → layout → master). Ex. cartes = spcPct 90 % : sans ça
        // j'utilise 1.15 → lignes trop espacées → dérive verticale vs PowerPoint (vu au comparateur).
        lnSpc: loc?.['a:lnSpc'] ?? inh?.['a:lnSpc'] ?? mas.lnSpc,
        // spcBef/spcAft résolus dans la même chaîne (inline lstStyle → layout → master)
        spcBef: loc?.['a:spcBef'] ?? inh?.['a:spcBef'] ?? mas.spcBef,
        spcAft: loc?.['a:spcAft'] ?? inh?.['a:spcAft'] ?? mas.spcAft,
      };
    };
    const counters = {};
    const paras = arr(txBody['a:p']);
    let html = '', hasText = false;
    for (let pi = 0; pi < paras.length; pi++) {
      const p = paras[pi];
      const pPr = p['a:pPr'] || {};
      const lvl = (+pPr['@lvl'] || 0) + 1;
      const def = lvlStyle(lvl);
      const algn = { ctr:'center', r:'right', just:'justify' }[pPr['@algn'] || def.algn] || 'left';
      const runs = arr(p['a:r']);
      if (!runs.length && p['a:fld']) runs.push(p['a:fld']);
      let line = '', firstColor = def.color || DEFCOL, firstSz = (def.sz || 18) * fontScale;
      // paragraphe contenant des FORMULES OMML (a14:m / m:oMath) : rendre en ORDRE le texte (a:r)
      // et le math (fractions, etc.) intercalés — le parse plat perd l'ordre a:r ↔ math.
      const paraO = txBodyO && (p['a14:m'] !== undefined || p['m:oMathPara'] !== undefined || p['m:oMath'] !== undefined)
        ? byTag(txBodyO, 'a:p')[pi] : null;
      if (paraO) {
        const fam = fontFamily(resolveFont(def.font) || ctx.minorFont), pc = def.color || DEFCOL;
        for (const ch of kids(paraO)) {
          const ct = tagOf(ch);
          if (ct === 'a:r') {
            const rPrO = first(ch, 'a:rPr');
            const a = at(rPrO || {});
            const sz = ((a['@sz'] ? +a['@sz']/100 : def.sz) || 18) * fontScale;
            const b = (a['@b'] === '1' || (a['@b'] === undefined && def.b)) ? 'font-weight:700;' : '';
            const it = a['@i'] === '1' ? 'font-style:italic;' : '';
            // couleur PROPRE du run (sinon défaut paragraphe) — le chemin OMML l'ignorait, d'où du texte
            // en couleur de fond invisible (ex. encart rouge d.15 : run tx1=clair rendu en bg1=sombre).
            const sfO = rPrO && first(rPrO, 'a:solidFill');
            const rc = sfO ? cssColor(resolveColorFlat(PF.parse(`<r xmlns:a="a">${serialize(sfO)}</r>`).r['a:solidFill'], ctx)) : pc;
            const tN = first(ch, 'a:t'); const txt = esc(tN ? kids(tN).map(c => c['#text'] ?? '').join('') : '');
            if (txt.trim()) { hasText = true; firstColor = rc; }
            line += `<span style="font-size:${sz.toFixed(2)}px;color:${rc};font-family:${fam};${b}${it}">${txt}</span>`;
          } else if (ct === 'a14:m' || ct === 'm:oMathPara' || ct === 'm:oMath') {
            line += `<span style="font-size:${((def.sz || 18) * fontScale).toFixed(2)}px;color:${pc};font-family:${fam}">${renderMath(ch)}</span>`;
            hasText = true;
          }
        }
      } else
      for (let ri = 0; ri < runs.length; ri++) {
        const r = runs[ri];
        const rPr = r['a:rPr'] || {};
        const sz = ((+rPr['@sz'] ? +rPr['@sz']/100 : def.sz) || 18) * fontScale; // points, autofit
        const col = rPr['a:solidFill'] ? cssColor(resolveColorFlat(rPr['a:solidFill'], ctx)) : (def.color || DEFCOL);
        const font = resolveFont(rPr['a:latin']?.['@typeface']) || resolveFont(def.font) || ctx.minorFont;
        const b = (rPr['@b']==='1' || (rPr['@b']===undefined && def.b))?'font-weight:700;':'';
        const i = rPr['@i']==='1'?'font-style:italic;':'';
        // décorations cumulables : soulignement + barré (§21.1.2.3.6 u / strike)
        const deco = [];
        if (rPr['@u'] && rPr['@u']!=='none') deco.push('underline');
        if (rPr['@strike'] && rPr['@strike']!=='noStrike') deco.push('line-through');
        const td = deco.length ? `text-decoration:${deco.join(' ')};` : '';
        // cap : MAJUSCULES auto (all) / petites capitales (small) — §20.1.10.65 ST_TextCapsType
        const cap = rPr['@cap']==='all' ? 'text-transform:uppercase;'
                  : rPr['@cap']==='small' ? 'font-variant:small-caps;' : '';
        // spc : espacement de caractères en centièmes de point (ST_TextPoint)
        const spc = rPr['@spc'] != null && +rPr['@spc'] ? `letter-spacing:${(+rPr['@spc']/100).toFixed(2)}px;` : '';
        // baseline : exposant/indice (ST_Percentage, millièmes de %) -> décalage vertical
        const bl = rPr['@baseline'] && +rPr['@baseline'] ? `vertical-align:${(+rPr['@baseline']/1000).toFixed(1)}%;` : '';
        // highlight : surlignage du texte (§21.1.2.3.4) -> fond coloré derrière le run (comme Word)
        const hl = rPr['a:highlight'] ? `background-color:${cssColor(resolveColorFlat(rPr['a:highlight'], ctx))};` : '';
        // contour de texte (rPr>a:ln, type WordArt/titre stylé §21.1.2.1.11) -> -webkit-text-stroke.
        // Additif : présent uniquement si le run déclare un a:ln avec remplissage -> aucun autre run touché.
        const tln = rPr['a:ln'];
        let ol = '';
        if (tln && tln['a:noFill'] === undefined && (tln['a:solidFill'] || tln['a:gradFill'])) {
          const oc = tln['a:solidFill'] ? cssColor(resolveColorFlat(tln['a:solidFill'], ctx)) : col;
          const ow = Math.max(0.3, u(+tln['@w'] || 9525) / 2); // EMU->px, moitié (le trait déborde des 2 côtés)
          ol = `-webkit-text-stroke:${ow.toFixed(2)}px ${oc};`;
        }
        // champ slidenum : afficher le VRAI numéro de diapo (sinon le texte-invite « ‹N°› »
        // du layout/master fuit — c'est une zone libre non-placeholder, donc non sautée).
        let raw = r['@type'] === 'slidenum' ? String(slideNo) : (r['a:t'] ?? '');
        // espaces de FIN de paragraphe (dernier run) : non rendus par PowerPoint → couper (sinon
        // ils élargissent la boîte et déclenchent une sur-réduction de l'autofit). Espaces INTERNES
        // multiples conservés (pre-wrap) pour le positionnement type « TURBULENCE   vu à ».
        if (ri === runs.length - 1) raw = raw.replace(/[ \t]+$/, '');
        // run en police SYMBOLE (Wingdings/Webdings/Symbol, souvent via <a:sym>) : remapper vers Unicode
        // + police normale (anti-tofu Linux/LMS). Sinon police telle quelle.
        const symMapped = symbolText(raw, font, rPr['a:sym']?.['@typeface']);
        const txt = esc(symMapped ?? raw);
        const famCss = symMapped != null ? "'Arial','Liberation Sans',sans-serif" : fontFamily(font);
        if (ri === 0) firstSz = sz; // taille réelle du 1er run -> interligne correct du <div>
        if (txt.trim()) { hasText = true; firstColor = col; }
        line += `<span style="font-size:${sz.toFixed(2)}px;color:${col};font-family:${famCss};${b}${i}${td}${cap}${spc}${bl}${hl}${ol}">${txt}</span>`;
      }
      // puce : pPr du paragraphe prioritaire, sinon héritée du niveau de style
      const bu = bulletOf(pPr) ?? def.bu;
      let bullet = '';
      if (bu && !bu.none && bu.blip && (line || '').trim()) {
        // puce IMAGE (a:buBlip) : petite icône alignée sur le texte
        const bid = bu.blip['@r:embed'] || bu.blip['@embed'];
        const bhref = bid && curRels[bid] ? imageHref(pptDir, curRels[bid], embed) : null;
        if (bhref) { const bsz = (bu.szPct ? firstSz * bu.szPct : firstSz); bullet = `<img src="${bhref}" style="height:${bsz.toFixed(1)}px;vertical-align:middle"/> `; }
      } else if (bu && !bu.none && (bu.char || bu.num) && (line || '').trim()) {
        let sym = bu.char, buFam;
        if (bu.num) { counters[lvl] = (counters[lvl] || 0) + 1; sym = counters[lvl] + (/paren/i.test(bu.num) ? ')' : '.'); buFam = fontFamily(resolveFont(bu.font)); }
        else {
          // puce caractère : mapper les puces en police SYMBOLE (Wingdings) vers Unicode (anti-tofu)
          const bs = bulletSym(bu.char, bu.font);
          sym = bs.sym;
          buFam = bs.symbolic ? "'Arial','Liberation Sans',sans-serif" : fontFamily(resolveFont(bu.font));
        }
        // couleur PROPRE de la puce (a:buClr) si présente, sinon couleur du 1er run ; taille via a:buSzPct
        const buCol = bu.clr ? cssColor(resolveColorFlat(bu.clr, ctx)) : firstColor;
        const buSz = bu.szPct ? firstSz * bu.szPct : firstSz;
        bullet = `<span style="font-size:${buSz.toFixed(2)}px;color:${buCol};font-family:${buFam}">${esc(sym)}</span> `;
      }
      // retrait : pPr du paragraphe prioritaire (marL/indent équilibrés du niveau de liste)
      const marLe = pPr['@marL'] != null ? +pPr['@marL'] : def.marL;
      const inde = pPr['@indent'] != null ? +pPr['@indent'] : def.ind;
      const marL = marLe != null ? u(marLe) : 0;
      // le retrait négatif (puce suspendue) ne doit jamais dépasser la marge -> pas de débordement gauche
      const ind = inde != null ? Math.max(u(inde), -marL) : 0;
      const pad = (marL || ind) ? `padding-left:${marL.toFixed(1)}px;text-indent:${ind.toFixed(1)}px;` : '';
      // interligne (lnSpc) + espacement avant/après (spcBef/spcAft) — §21.1.2.2.{5,7,12}
      // spcPct = millièmes de % (100000 = 100 %) ; spcPts = centièmes de point.
      // lnSpc du paragraphe, sinon hérité (lstStyle/layout/master via def.lnSpc). spcPct 100 % =
      // interligne SIMPLE ≈ 1.2× la police (pas 1.0) → facteur 1.2 pour coller à PowerPoint.
      // normAutofit @lnSpcReduction : PowerPoint réduit AUSSI l'interligne explicite du paragraphe
      // (pas seulement les boîtes sans lnSpc via le conteneur). Sans ça, un paragraphe à lnSpc 120 %
      // dans une boîte autofit à 20 % de réduction rend 1.44 au lieu de 0.96×1.2=1.152 → texte trop
      // haut → déborde → l'autofit navigateur sur-rétrécit → mauvais nb de lignes (ex. question d.6).
      const ls = pPr['a:lnSpc'] || def.lnSpc;
      const lnMul = 1 - lnRed; // 1 si pas de réduction -> aucun changement pour ces boîtes-là
      const lnSpcCss = ls?.['a:spcPct']?.['@val'] ? `line-height:${(+ls['a:spcPct']['@val']/100000*1.2*lnMul).toFixed(3)};`
                     : ls?.['a:spcPts']?.['@val'] ? `line-height:${(+ls['a:spcPts']['@val']/100*lnMul).toFixed(1)}px;` : '';
      const spcMk = (node, side) => !node ? ''
        : node['a:spcPts']?.['@val'] ? `margin-${side}:${(+node['a:spcPts']['@val']/100).toFixed(1)}px;`
        : node['a:spcPct']?.['@val'] ? `margin-${side}:${(+node['a:spcPct']['@val']/100000).toFixed(2)}em;` : '';
      // spcBef/spcAft : pPr de la diapo prioritaire, sinon HÉRITÉ (lstStyle/layout/master via def).
      // spcBef supprimé sur le 1er paragraphe : PowerPoint ne pousse pas le bloc vers le bas en haut
      // de boîte (l'espace-avant ne vaut qu'ENTRE paragraphes) — sinon décalage vertical du bloc entier.
      const spcBefNode = pi === 0 ? null : (pPr['a:spcBef'] ?? def.spcBef);
      const spcAftNode = pPr['a:spcAft'] ?? def.spcAft;
      const spcCss = spcMk(spcBefNode, 'top') + spcMk(spcAftNode, 'bottom');
      // data-par : index de paragraphe (pour le build par paragraphe bldP dans le player)
      // white-space:pre-wrap : préserve les espaces MULTIPLES (PowerPoint s'en sert pour positionner
      // du texte, ex. « TURBULENCE      vu à » où le « vu à » tombe entre deux badges) ; wrap conservé.
      // font-size sur le div = taille du texte -> l'interligne (unitless 1.15) se calcule dessus,
      // sinon il se base sur 16px (défaut) et gonfle la ligne (pied de page 9pt coupé dans sa boîte).
      html += `<div data-par="${pi}" style="font-size:${firstSz.toFixed(1)}px;white-space:pre-wrap;text-align:${algn};margin:0;${pad}${lnSpcCss}${spcCss}">${bullet}${line || '&#8203;'}</div>`;
    }
    if (!hasText) return '';
    // data-af : l'autofit navigateur ne doit réduire QUE les boîtes normAutofit (texte rétréci pour
    // tenir dans une boîte FIXE). spAutoFit = la boîte s'ajuste au texte → JAMAIS de réduction
    // (sinon pied de page « sample… » réduit à ~2px). Sans autofit = clip.
    const afMode = (naf !== undefined) ? 'norm' : 'fixed';
    // <a:noAutofit/> = « ne pas ajuster » : PowerPoint NE clippe PAS, le texte peut déborder la forme
    // (ex. libellés/pastilles : police 40pt dans une pastille de 40pt -> les capitales tiennent, mais
    // l'aire réduite par les insets clipperait). On étend alors le foreignObject à la boîte PLEINE, on
    // met les insets en padding et on laisse déborder (overflow visible).
    const noaf = bodyPr['a:noAutofit'] !== undefined;
    if (noaf) {
      return `<foreignObject x="0" y="0" width="${u(w)}" height="${u(h)}" style="overflow:visible">`
        + `<div xmlns="http://www.w3.org/1999/xhtml" class="af" data-af="fixed" style="box-sizing:border-box;width:100%;height:100%;padding:${u(iT)}px ${u(iR)}px ${u(iB)}px ${u(iL)}px;display:flex;flex-direction:column;justify-content:${just};line-height:${lineH.toFixed(3)};overflow:visible;">${html}</div>`
        + `</foreignObject>`;
    }
    // overflow VISIBLE : PowerPoint ne clippe pas le texte d'une forme au bord de la boîte. Avec
    // overflow:hidden, une ligne un peu plus haute que sa boîte (ex. chiffre 24pt dans une pastille dont
    // les insets réduisent la boîte à 25.7pt) était ROGNÉE en haut/bas. Le texte peut déborder (comme
    // PowerPoint) ; l'autofit navigateur (data-af=norm) rétrécit déjà les boîtes normAutofit trop pleines.
    return `<foreignObject x="${u(iL)}" y="${u(iT)}" width="${u(w-iL-iR)}" height="${u(h-iT-iB)}" style="overflow:visible">`
      + `<div xmlns="http://www.w3.org/1999/xhtml" class="af" data-af="${afMode}" style="width:100%;height:100%;display:flex;flex-direction:column;justify-content:${just};line-height:${lineH.toFixed(3)};overflow:visible;">${html}</div>`
      + `</foreignObject>`;
  }

  // --- composition : masque (bg + déco) -> layout (déco) -> diapo (tout) ---
  // exécutée ici, après TOUTES les définitions (évite la zone morte temporelle des const)
  const masterPart = masterPath ? partOf(masterPath, 'p:sldMaster') : null;
  const layoutPart = layoutPath ? partOf(layoutPath, 'p:sldLayout') : null;
  // UN SEUL fond effectif (précédence diapo > layout > masque), dessiné DERRIÈRE toutes les formes.
  // Le fond n'est PAS cumulatif : un <p:bg> de layout/diapo REMPLACE celui du niveau au-dessus. Le
  // dessiner à chaque niveau repeindrait par-dessus les formes du niveau précédent — ex. les logos du
  // masque (logos partenaires) masqués par le fond bleu du layout. curRels doit pointer le bon niveau
  // (un fond peut être une image r:embed).
  curRels = rels;                                  const slideBg  = bgRect(cSld);
  curRels = layoutRels;                            const layoutBg = layoutPart ? bgRect(layoutPart.cSld) : '';
  curRels = masterPath ? relsOf(masterPath) : {};  const masterBg = masterPart ? bgRect(masterPart.cSld) : '';
  body.push(slideBg || layoutBg || masterBg || '');
  if (masterPart) { curRels = relsOf(masterPath); skipPh = true; spidPfx = 'm'; walk(masterPart.spTree, IDENT); }
  if (layoutPart) { curRels = layoutRels; skipPh = true; spidPfx = 'l'; walk(layoutPart.spTree, IDENT); }
  curRels = rels; skipPh = false; spidPfx = '';
  walk(spTree, IDENT);

  // dégradés définis -> <defs>
  const defs = grads.map(g => {
    const stops = g.stops.map(s=>`<stop offset="${s.pos.toFixed(1)}%" stop-color="${s.col}"/>`).join('');
    if (g.radial) return `<radialGradient id="${g.id}" cx="50%" cy="50%" r="70%">${stops}</radialGradient>`;
    const dx = Math.cos(g.ang), dy = Math.sin(g.ang);
    return `<linearGradient id="${g.id}" x1="${((0.5-dx/2)*100).toFixed(1)}%" y1="${((0.5-dy/2)*100).toFixed(1)}%" x2="${((0.5+dx/2)*100).toFixed(1)}%" y2="${((0.5+dy/2)*100).toFixed(1)}%">${stops}</linearGradient>`;
  }).join('');

  // motifs -> tuiles <pattern> 8×8 (fond bgClr + géométrie fgClr selon la famille de preset)
  const pdefs = patterns.map(p => {
    const s = 8, prst = p.prst; let geo = '';
    const line = (x1,y1,x2,y2) => `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${p.fg}" stroke-width="1"/>`;
    if (/^pct(\d+)/.test(prst)) { // points, densité ~ pourcentage
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

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${u(W)} ${u(H)}" width="100%" style="display:block">`
    + `<defs>${defs}${pdefs}${filters.join('')}${clips.join('')}${[...lineMarkers.values()].join('')}</defs>`
    + `<rect width="${u(W)}" height="${u(H)}" fill="#ffffff"/>`
    + body.join('') + `</svg>`;
  return { svg, w: W, h: H };
}

// ================= extraction de la timeline d'animation (<p:timing>) =================
// Renvoie les "étapes de clic" : [{ effects:[{spid, cls, filter, transition, delay, dur}] }]
// cls = entr|exit|emph ; filter (fade/wipe(..)/circle(out)..) ; delay/dur en ms.
export function extractTiming(pptDir, slideNo) {
  const text = readFileSync(join(pptDir, 'slides', `slide${slideNo}.xml`), 'utf8');
  const m = text.match(/<p:timing>[\s\S]*<\/p:timing>/);
  if (!m) return { steps: [], animatedSpids: [] };
  const tree = P.parse(m[0]);
  const timing = tree.find((n) => tagOf(n) === 'p:timing');
  const seq = deepFind(timing, 'p:seq');
  if (!seq) return { steps: [], animatedSpids: [] };
  const seqCTn = first(seq, 'p:cTn');
  const clickPars = byTag(first(seqCTn, 'p:childTnLst'), 'p:par');

  const animatedSpids = new Set();
  const steps = clickPars.map((par) => {
    const effects = [];
    collectEffects(par, 0, effects);
    // entrée ciblant la forme entière (sans plage de paragraphes) -> cachée au départ.
    // Les entrées par paragraphe (bldP) sont gérées séparément par le player (data-par).
    for (const e of effects) if (e.cls === 'entr' && !e.par) animatedSpids.add(e.spid);
    return { effects };
  }).filter((s) => s.effects.length);
  return { steps, animatedSpids: [...animatedSpids] };

  function collectEffects(node, accDelay, acc) {
    if (tagOf(node) === 'p:par') {
      const cTn = first(node, 'p:cTn');
      if (cTn && at(cTn)['@presetID'] !== undefined) { acc.push(parseEffect(cTn, accDelay)); return; }
      if (cTn) {
        const cond = deepFind(first(cTn, 'p:stCondLst') ?? {}, 'p:cond');
        const d = cond && at(cond)['@delay'];
        if (d !== undefined && d !== 'indefinite') accDelay += Number(d) || 0;
      }
    }
    for (const c of kids(node)) if (typeof c === 'object') collectEffects(c, accDelay, acc);
  }
  function parseEffect(cTn, accDelay) {
    const a = at(cTn);
    const spTgt = deepFind(cTn, 'p:spTgt');
    const animEff = deepFind(cTn, 'p:animEffect');
    const pRg = spTgt ? deepFind(spTgt, 'p:pRg') : null; // plage de paragraphes (bldP)
    const own = deepFind(first(cTn, 'p:stCondLst') ?? {}, 'p:cond');
    const ownDelay = own ? Number(at(own)['@delay']) || 0 : 0;
    // durée : 1er cTn d'un behavior avec dur numérique
    let dur = 500;
    const durNode = [...allCTn(cTn)].map((c) => at(c)['@dur']).find((d) => d && d !== 'indefinite' && +d > 1);
    if (durNode) dur = +durNode;
    return {
      cls: a['@presetClass'], spid: spTgt ? at(spTgt)['@spid'] : null,
      par: pRg ? { from: +at(pRg)['@st'] || 0, to: +at(pRg)['@end'] || 0 } : null,
      filter: animEff ? at(animEff)['@filter'] : null,
      transition: animEff ? at(animEff)['@transition'] : null,
      motion: parseMotion(cTn),
      delay: accDelay + ownDelay, dur,
    };
  }
  // mouvement/rotation (§19.5.3 animMotion, §19.5.5 animRot) -> {kind, dx/dy (fractions) | deg, dur, rev}
  function behTiming(node) { const c = deepFind(node, 'p:cTn'); const a = at(c || {}); return { dur: +a['@dur'] || 1000, rev: a['@autoRev'] === '1' }; }
  function parseMotion(cTn) {
    const mm = deepFind(cTn, 'p:animMotion');
    if (mm) {
      const path = at(mm)['@path'] || '';
      const pts = (path.match(/-?\d*\.?\d+(?:[eE][-+]?\d+)?/g) || []).map(Number);
      if (pts.length >= 2) {
        const fx = pts[0], fy = pts[1], lx = pts[pts.length - 2], ly = pts[pts.length - 1];
        const t = behTiming(mm); return { kind: 'move', dx: lx - fx, dy: ly - fy, dur: t.dur, rev: t.rev };
      }
    }
    const mr = deepFind(cTn, 'p:animRot');
    if (mr) { const t = behTiming(mr); return { kind: 'rot', deg: (+at(mr)['@by'] || 0) / 60000, dur: t.dur, rev: t.rev }; }
    return null;
  }
  function* allCTn(node) {
    if (tagOf(node) === 'p:cTn') yield node;
    for (const c of kids(node)) if (typeof c === 'object') yield* allCTn(c);
  }
}

// sérialise un noeud preserveOrder -> XML (pour re-parser en flat)
function serialize(node) {
  const t = tagOf(node);
  const a = at(node);
  const attrs = Object.entries(a).map(([k,v])=>` ${k.slice(1)}="${esc(v)}"`).join('');
  const children = kids(node);
  if (typeof children === 'string') return `<${t}${attrs}>${esc(children)}</${t}>`;
  const inner = (children||[]).map(c => {
    if (c['#text'] !== undefined) return esc(c['#text']);
    return serialize(c);
  }).join('');
  return `<${t}${attrs}>${inner}</${t}>`;
}

// ─── EXPORTS POUR TESTS UNITAIRES (fonctions PURES uniquement — AUCUNE I/O) ───
// Bloc AJOUTÉ en fin de fichier : n'altère aucune ligne de logique existante.
// Il expose uniquement des helpers déterministes (sortie = fonction des entrées,
// sans fs/réseau/DOM/état global) pour test/pptx-render.test.js.
export {
  u, esc, serialize,                                   // conversion EMU->pt · échappement XML · (dé)sérialisation
  hslFromRgb, rgbFromHsl, cssColor, resolveColorFlat,  // couleurs : conversions, format CSS, résolution + clamp
  symFontKind, symToUnicode, bulletSym, symbolText,    // polices symbole (Wingdings/Webdings/Symbol) -> Unicode
  fontFamily, bulletOf,                                // police -> familles CSS · lecture d'une puce (pPr)
  tagOf, at, kids, byTag, first, deepFindFlat,          // helpers d'arbre (preserveOrder + objet plat)
};
