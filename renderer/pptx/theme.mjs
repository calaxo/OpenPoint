// ═══ [PPTX · THÈME] ═══ loadContext() : palette, polices, fmtScheme et styles de tableaux d'un thème.
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { PF, cachedParse, fileStamp as _stamp, tagOf, kids, first, deepFind, serialize } from './xml.mjs';
// ████ §I ─ STYLE DE TABLEAU INTÉGRÉ, UNITÉS ────────────────────────────────────────────────────
// Style de tableau par défaut de PowerPoint (copie exacte de sa définition, telle qu'écrite dans un
// tableStyles.xml) : repli quand le fichier y renvoie sans la stocker.
// Structure (§20.1.4.2 table styles) : wholeTbl (tout le tableau : texte dk1, bordures lt1 1 pt, fond
// accent1 teinté 20 %), band1H/band1V (bandes impaires : accent1 teinté 40 %), firstRow / lastRow /
// firstCol / lastCol (texte gras lt1 sur accent1 plein, filet 3 pt entre en-tête et corps).
// Parsé une fois en objet plat par buildContext, puis appliqué en couches par renderTable.
const BUILTIN_TBL_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
const BUILTIN_TBL_STYLE = '<a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Style moyen 2 - Accentuation 1"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right><a:top><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top><a:bottom><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom><a:insideH><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideH><a:insideV><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideV></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H><a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V><a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V><a:lastCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastCol><a:firstCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol><a:lastRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:top><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastRow><a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle>';
// ████ §J ─ CONTEXTE DOCUMENT (THÈME) ───────────────────────────────────────────────────────────
// Le « ctx » est l'objet transmis (par fermeture) à tout le rendu d'une diapo. Champs :
//   pptDir, clr (slots bruts du thème : dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink -> hex),
//   clrMap (alias bg1/tx1/bg2/tx2/accentN… -> slot, rempli par renderSlide depuis le master),
//   phClr (couleur de substitution courante, voir paint), majorFont / minorFont (polices du
//   fontScheme, script latin uniquement), fillStyles / lnStyles / effectStyles / bgFillStyles
//   (listes ordonnées du fmtScheme, indexées par les fillRef/lnRef/effectRef/bgRef des formes),
//   tableStyles (styleId -> style). Ajoutés ensuite par renderSlide : txStyles, defaultTextStyle,
//   grpFill.
// ================= contexte document =================
// Cache du CŒUR de contexte (parties immuables dérivées du thème) : keyé par fichier de thème (+ sa
// signature mtime/taille). loadContext() renvoie une COPIE avec clrMap/phClr remis à null car ceux-là
// sont MUTÉS par diapo (clrMap du masque, phClr courant) et ne doivent jamais être partagés.
const _ctxCache = new Map();
// loadContext(pptDir, themePath?) -> ctx neuf (copie superficielle du cœur en cache).
// themePath absent ou introuvable -> theme/theme1.xml (cas d'un deck mono-master).
// Attention : la copie est SUPERFICIELLE — fillStyles, tableStyles… sont partagés et ne doivent pas
// être modifiés ; seuls les champs de premier niveau (clrMap, phClr, txStyles…) peuvent l'être.
export function loadContext(pptDir, themePath) {
  const themeFile = (themePath && existsSync(themePath)) ? themePath : join(pptDir, 'theme', 'theme1.xml');
  const sig = _stamp(themeFile);
  let c = _ctxCache.get(themeFile);
  if (!c || c.sig !== sig) { c = { sig, core: buildContext(pptDir, themeFile) }; _ctxCache.set(themeFile, c); }
  return { ...c.core, clrMap: null, phClr: null };
}
// buildContext : construit le cœur immuable du ctx depuis le thème (a:theme/a:themeElements, §20.1.6)
// et tableStyles.xml. Appelé seulement en cas d'absence/péremption du cache.
function buildContext(pptDir, themeFile) {
  // theme du MASTER de la diapo (un deck multi-master a plusieurs thèmes : ex. deck C a
  // theme1 lt2=C78243 orange ET theme2 lt2=D9CAB7 beige -> mauvais thème = mauvaise couleur de fond).
  const theme = cachedParse(themeFile, 'flat');
  const el = theme['a:theme']['a:themeElements'];
  const scheme = el['a:clrScheme'];
  const clr = {}; // slots BRUTS du thème (dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink)
  // Chaque slot contient soit a:srgbClr@val (hex), soit a:sysClr (couleur système Windows : on prend
  // windowText -> noir, window -> blanc, sinon @lastClr = dernière valeur calculée par Office).
  for (const [k, v] of Object.entries(scheme)) {
    if (!k.startsWith('a:')) continue; // saute '@name' (attribut du clrScheme)
    clr[k.slice(2)] = v['a:srgbClr']?.['@val'] ??
      { windowText: '000000', window: 'FFFFFF' }[v['a:sysClr']?.['@val']] ??
      v['a:sysClr']?.['@lastClr'] ?? '000000';
  }
  // listes de styles du fmtScheme, EN ORDRE (fillRef/bgRef indexent dedans) -> noeuds flat
  const themeO = cachedParse(themeFile, 'ord');
  const themeRoot = themeO.find(n => tagOf(n) === 'a:theme');
  const fmt = deepFind(themeRoot, 'a:fmtScheme');
  // Pourquoi le détour par l'arbre ORDONNÉ : fillStyleLst contient p. ex. solidFill, gradFill,
  // gradFill — en plat, les deux gradFill seraient regroupés dans un tableau et l'index de position
  // (fillRef idx=2 -> 2e élément) serait perdu. On prend chaque enfant dans l'ordre, puis on le
  // re-parse SEUL en plat : chaque entrée de la liste = { 'a:solidFill': … } ou { 'a:gradFill': … }.
  // Les xmlns factices du wrapper <root> ne servent qu'à fournir une racine (fast-xml-parser ne
  // résout pas les namespaces).
  const orderedFlatList = (listTag) => {
    const lst = fmt && first(fmt, listTag);
    if (!lst) return [];
    return kids(lst).map(c => PF.parse(`<root xmlns:a="a" xmlns:p="p">${serialize(c)}</root>`).root);
  };
  // styles de tableaux : ceux du fichier, plus le style PAR DÉFAUT de PowerPoint (« Moyen 2 -
  // Accentuation 1 ») que le fichier omet souvent alors que ses tableaux y renvoient.
  const tableStyles = {};
  try { tableStyles[BUILTIN_TBL_STYLE_ID] = PF.parse(`<r xmlns:a="a">${BUILTIN_TBL_STYLE}</r>`).r['a:tblStyle']; } catch {}
  // tableStyles.xml du deck (facultatif : absent -> statSync lève -> catch silencieux). Un style du
  // fichier portant le même styleId que le style intégré le remplace.
  try {
    const tsx = cachedParse(join(pptDir, 'tableStyles.xml'), 'flat');
    const list = tsx['a:tblStyleLst']?.['a:tblStyle'];
    for (const st of (Array.isArray(list) ? list : list ? [list] : [])) {
      tableStyles[st['@styleId']] = st; // style complet (couches wholeTbl, bandes, 1re/dern. ligne/colonne)
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
