// ═══ [PPTX · POLICES] ═══ typeface -> pile CSS, métriques (ligne de base, chasses 1/8 pt), largeurs.
import { esc } from './xml.mjs';
import { familyMetrics, advanceTable } from '../animFonts.js';
// police -> famille CSS. Substituts LIBRES métrique-compatibles (embarqués) pour les polices non libres
// courantes : garantit largeur/retour-à-la-ligne corrects MÊME sans la police installée (ex. LMS Linux).
// Ordre : police réelle d'abord (si installée), puis substitut embarqué, puis générique.
// SERIF_RE : heuristique « cette police inconnue est-elle à empattements ? » -> générique serif.
// (« book » attrape aussi des sans-serif comme Franklin Gothic Book, mais celle-ci est dans FONT_SUBST,
// testé avant.)
const SERIF_RE = /times|cambria|georgia|bernhard|garamond|antiqua|minion|book|schoolbook/i;
// SYMB_RE : polices symboles -> pas de repli générique (le glyphe d'une autre police serait faux ; le
// texte de ces polices est de toute façon remappé vers Unicode, voir §H).
const SYMB_RE = /wingding|webding|symbol/i;
// Table typeface (minuscules) -> pile CSS « police réelle, substitut libre, générique ». Les
// substituts (Liberation, Carlito, Caladea, Gelasio…) sont conçus avec les MÊMES chasses que
// l'original : texte de même largeur, donc mêmes retours à la ligne. Ils sont embarqués côté web
// (animFonts.js / web-fonts.mjs).
const FONT_SUBST = {
  'arial': "'Arial','Liberation Sans',sans-serif", 'helvetica': "'Helvetica','Liberation Sans',sans-serif",
  'arial narrow': "'Arial Narrow','Liberation Sans Narrow',sans-serif",
  'times new roman': "'Times New Roman','Liberation Serif',serif", 'times': "'Times','Liberation Serif',serif",
  'courier new': "'Courier New','Liberation Mono',monospace", 'courier': "'Courier','Liberation Mono',monospace",
  'calibri': "'Calibri','Carlito',sans-serif", 'cambria': "'Cambria','Caladea',serif",
  // police PowerPoint d'abord (poste qui l'a installée), puis équivalent libre embarqué (animFonts.js)
  'calibri light': "'Calibri Light','Carlito',sans-serif",
  'cambria math': "'Cambria Math','Caladea','DejaVu Sans',serif",
  'consolas': "'Consolas','Liberation Mono',monospace",
  'georgia': "'Georgia','Gelasio',serif",
  'comic sans ms': "'Comic Sans MS','Comic Relief',cursive",
  'arial black': "'Arial Black','Archivo Black',sans-serif",
  'impact': "'Impact','Anton',sans-serif",
  'franklin gothic book': "'Franklin Gothic Book','Libre Franklin',sans-serif",
  'franklin gothic medium': "'Franklin Gothic Medium','Libre Franklin',sans-serif",
  'news gothic mt': "'News Gothic MT','News Cycle',sans-serif",
  'gill sans mt': "'Gill Sans MT','Cabin',sans-serif", 'gill sans': "'Gill Sans','Cabin',sans-serif",
  'bernhard modern roman': "'Bernhard Modern Roman','Sorts Mill Goudy',serif",
  'century schoolbook': "'Century Schoolbook','TeX Gyre Schola',serif",
  'palatino linotype': "'Palatino Linotype','TeX Gyre Pagella',serif",
  'palatino': "'Palatino','TeX Gyre Pagella',serif", 'book antiqua': "'Book Antiqua','TeX Gyre Pagella',serif",
  'corbel': "'Corbel','Source Sans 3',sans-serif",
  'aptos': "'Aptos','Source Sans 3',sans-serif", 'aptos display': "'Aptos Display','Source Sans 3',sans-serif",
};
// fontFamily(typeface) -> valeur CSS font-family (déjà échappée). Le typeface doit être RÉSOLU
// (pas « +mn-lt ») : c'est resolveFont, dans renderText, qui remplace les références de thème.
// Les noms entre quotes de la pile servent aussi de clés à metricsOf / advanceTable (§F) : le 1er
// nom pour lequel on a des métriques gagne (souvent le substitut embarqué).
function fontFamily(f) {
  if (!f) return 'sans-serif';
  if (SYMB_RE.test(f)) return `'${esc(f)}'`;
  const key = f.toLowerCase().trim();
  if (FONT_SUBST[key]) return FONT_SUBST[key];
  return SERIF_RE.test(f) ? `'${esc(f)}',serif` : `'${esc(f)}',sans-serif`;
}
// ████ §F ─ MÉTRIQUES DE POLICE : LIGNE DE BASE, CHASSES, LARGEURS ──────────────────────────────
// Le texte est mis en page par le navigateur (HTML dans foreignObject). Trois écarts systématiques
// avec PowerPoint sont corrigés ici, à partir des métriques des polices embarquées (animFonts.js) :
//   1. position VERTICALE de la ligne de base dans la ligne  -> baselineShift / applyShift / VA_TOK
//   2. largeur des mots (chasses arrondies au 1/8 pt)        -> advanceFix (letter-spacing)
//   3. mesures de largeur pour les tabulations / tableaux    -> pptTextWidth
// Notations : sz = taille de police (pt = px CSS ici), L = hauteur de ligne (px), A / D = ascendante /
// descendante de la police en fraction d'em (valeurs positives).
//
// Ligne de base façon PowerPoint. Calibrage (PowerPoint exporté vs navigateur, 4 polices, interlignes
// 80-200 %) : la ligne de base de chaque ligne tombe, depuis le haut de sa ligne de hauteur L, à
//   L > 1,2 x taille (interligne > 100 %) : 0,75 L (quelle que soit la police)
//   sinon                                  : L - min(0,25 L, descendante win)
// alors que le navigateur la met à L/2 + (ascendante - descendante)/2. L'écart (constant pour toutes les
// lignes du paragraphe) est rattrapé par un décalage VISUEL (position:relative) : la mise en page (pas
// de ligne, hauteur du bloc, ancrage) reste celle du navigateur, identique à PowerPoint.
let _fm = null; // cache paresseux de familyMetrics() : { 'Nom de famille': [A, D], … } ({} si échec)
// metricsOf(pile CSS) -> [A, D] de la 1re famille de la pile qui a des métriques, sinon null.
function metricsOf(famCss) {
  if (!_fm) { try { _fm = familyMetrics(); } catch { _fm = {}; } }
  const fam = (famCss.match(/'([^']+)'/g) || []).map((q) => q.slice(1, -1)).find((f) => _fm[f]);
  return fam ? _fm[fam] : null;
}
// Recalage de ligne de base appliqué aux RUNS (vertical-align), et non au paragraphe : un
// position:relative dans un foreignObject est dessiné par WebKit (Safari / iPad) par rapport à la PAGE
// et non au SVG -> tout le texte s'empilait en haut à gauche de la diapo sur iPad. Les runs étant en
// line-height 0, leur vertical-align déplace les glyphes sans changer la mise en page.
// Mécanisme en deux temps : chaque run est émis avec `vertical-align:calc(<exposant>em - @@VA@@px)`,
// car le décalage d dépend de la police du 1er run et de l'interligne du paragraphe, connus seulement
// après avoir lu tous les runs. applyShift remplace ensuite le jeton par la valeur de d. Soustraire d
// (vertical-align positif = vers le HAUT) abaisse les glyphes quand PowerPoint place la base plus bas.
const VA_TOK = ' - @@VA@@px)'; // fin de « vertical-align:calc(Xem - ?px) », remplacée par paragraphe
// applyShift(html, d) : substitue toutes les occurrences du jeton (split/join = replaceAll) par « - d px) ».
function applyShift(html, d) {
  return html.split(VA_TOK).join(` - ${(d || 0).toFixed(2)}px)`);
}
// baselineShift(pile CSS, sz, L) -> d (px) = position PowerPoint - position navigateur de la ligne de
// base, mesurée depuis le haut de la ligne. 0 si police sans métriques ou entrées invalides.
export function baselineShift(famCss, sz, L) {
  const m = metricsOf(famCss);
  if (!m || !(sz > 0) || !(L > 0)) return 0;
  const [A, D] = m;
  // modèle CSS (« half-leading ») : la boîte A+D est centrée dans L -> base à L/2 + (A-D)/2 · sz
  const css = L / 2 + (A - D) / 2 * sz;
  // modèle PowerPoint calibré (voir le commentaire au-dessus de metricsOf) ; +0,01 = tolérance
  // flottante pour que l'interligne simple (L = 1,2 sz exactement) tombe dans la 2e branche
  const ppt = L > 1.2 * sz + 0.01 ? 0.75 * L : L - Math.min(0.25 * L, D * sz);
  const d = ppt - css;
  // écart < 0,05 px ignoré : évite d'émettre un calc() pour un bruit d'arrondi
  return Math.abs(d) < 0.05 ? 0 : d;
}
// Arrondi des chasses façon PowerPoint : chaque glyphe avance d'un multiple de 1/8 pt (1/576 pouce,
// arrondi au plus proche) — calibré sur PowerPoint via COM : 0,014 % d'écart contre 0,14 % sans. Le
// navigateur, lui, cumule des chasses exactes : à 18 pt Poppins il compose 0,36 % plus large, assez
// pour renvoyer un mot à la ligne. -> écart moyen par caractère, ajouté au letter-spacing du run.
// Cache des tables de chasses : clé « famille|gras|italique » -> { a: {codepoint: chasse}, upm } ou null
const _advCache = new Map();
// advanceFix(pile CSS, gras, italique, sz, texte) -> correction MOYENNE par caractère (px) à ajouter
// au letter-spacing du run. letter-spacing s'ajoutant après chaque caractère, n × moyenne = somme des
// écarts d'arrondi : la largeur totale du run égale celle de PowerPoint (à la répartition près).
export function advanceFix(famCss, bold, italic, sz, text) {
  if (!text || !(sz > 0)) return 0;
  const fams = (famCss.match(/'([^']+)'/g) || []).map((q) => q.slice(1, -1)); // noms entre quotes de la pile
  let tbl = null;
  for (const f of fams) { // 1re famille de la pile qui possède une table
    const k = f + '|' + bold + '|' + italic;
    if (!_advCache.has(k)) { let t = null; try { t = advanceTable(f, bold, italic); } catch { /* sans table */ } _advCache.set(k, t); }
    if ((tbl = _advCache.get(k))) break;
  }
  if (!tbl) return 0;
  let delta = 0, n = 0;
  for (const ch of text) {
    const a = tbl.a[ch.codePointAt(0)];
    if (a == null) return 0; // glyphe hors table (repli de police) : ne pas deviner
    const w = a * sz / tbl.upm; // chasse exacte en pt (a en unités de fonte, upm = unités par em)
    delta += Math.round(w * 8) / 8 - w; n++; // arrondi au 1/8 pt le plus proche, moins la valeur exacte
  }
  return n ? delta / n : 0;
}
// Largeur (pt) d'un texte avec les chasses ARRONDIES de PowerPoint (1/8 pt) ; null si glyphe inconnu.
// Sert aux calculs que le navigateur ne sait pas faire lui-même : position des tabulations
// (renderText) et hauteur des lignes de tableau (cellTextHeight). N'utilise pas _advCache (appelle
// advanceTable à chaque fois) et ne tient pas compte du crénage ni de @spc.
function pptTextWidth(famCss, bold, italic, sz, text) {
  const fams = (famCss.match(/'([^']+)'/g) || []).map((q) => q.slice(1, -1));
  let tbl = null;
  for (const f of fams) { try { tbl = advanceTable(f, bold, italic); } catch { tbl = null; } if (tbl) break; }
  if (!tbl) return null;
  let w = 0;
  for (const ch of text) { const a = tbl.a[ch.codePointAt(0)]; if (a == null) return null; w += Math.round(a * sz / tbl.upm * 8) / 8; }
  return w;
}
// ████ §G ─ HAUTEUR DE TEXTE D'UNE CELLULE DE TABLEAU ───────────────────────────────────────────
// Dans un tableau, a:tr@h n'est qu'une hauteur MINIMALE : PowerPoint agrandit la ligne si le texte
// ne tient pas. Le navigateur ne peut pas faire ce calcul pour nous (chaque cellule
// est un foreignObject indépendant, de hauteur fixée), d'où cette mesure « à la PowerPoint » faite
// côté Node, appelée par renderTable.
//
// Hauteur (EMU) qu'occupe le texte d'une cellule de largeur cellW (EMU) : coupure mot à mot comme
// PowerPoint, 1,2 x taille par ligne (interligne simple), + marges haut/bas de la cellule.
// Entrée : tc = cellule en arbre PLAT ; styleBold = gras imposé par le style (1re ligne / 1re colonne).
// Sortie : EMU, ou null si la mesure est impossible (glyphe sans table de chasses, cellule vide) ->
// l'appelant garde alors @h tel quel.
// Simplifications : ignore lnSpc / spcBef / spcAft, les a:br (lus en plat, seuls les a:r comptent),
// les tabulations et l'autofit ; la police de thème est toujours la MINEURE (+mj- compris).
export function cellTextHeight(tc, cellW, styleBold) {
  const tcPr = tc['a:tcPr'] || {}, S = 1 / 12700; // S local = même valeur que la constante S de §I
  // largeur utile en pt : largeur de colonne(s) - marges gauche/droite (défaut 91440 EMU = 0,1")
  const avail = (cellW - (+(tcPr['@marL'] ?? 91440)) - (+(tcPr['@marR'] ?? 91440))) * S;
  const arrOf = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];
  let h = 0;
  for (const p of arrOf(tc['a:txBody']?.['a:p'])) {
    const runs = arrOf(p['a:r']);
    const endSz = +(p['a:endParaRPr']?.['@sz']) || 0;
    // paragraphe vide : une ligne à la taille de sa marque de fin (défaut 18 pt = @sz 1800)
    if (!runs.length) { h += 1.2 * ((endSz || 1800) / 100); continue; }
    let lineW = 0, lines = 1, maxSz = 0;
    for (const r of runs) {
      const rp = r['a:rPr'] || {}, sz = (+rp['@sz'] || endSz || 1800) / 100;
      const bold = rp['@b'] !== undefined ? rp['@b'] === '1' : styleBold;
      const tf = rp['a:latin']?.['@typeface'];
      const fam = fontFamily(!tf || /^\+m[nj]-/.test(tf) ? _ctxMinorFont : tf);
      maxSz = Math.max(maxSz, sz); // la plus grande taille du paragraphe fixe la hauteur de TOUTES ses lignes
      for (const word of String(r['a:t'] ?? '').split(/(?<= )/)) { // mots avec leur espace finale
        const w = pptTextWidth(fam, bold, rp['@i'] === '1', sz, word);
        if (w == null) return null;
        // règle de coupure PowerPoint : l'espace finale d'un mot ne compte pas pour savoir s'il tient
        const wNoSp = word.endsWith(' ') ? w - (pptTextWidth(fam, bold, false, sz, ' ') || 0) : w;
        // -0,02 pt : petite tolérance ; PowerPoint renvoie le mot dès que la ligne atteint la largeur
        // disponible à ~0,01 pt près (même constat que le margin-right 0,2 px de renderText)
        if (lineW > 0 && lineW + wNoSp > avail - 0.02) { lines++; lineW = w; } else lineW += w; // même règle que le texte
      }
    }
    h += lines * 1.2 * maxSz;
  }
  if (!h) return null;
  // retour en EMU + marges haut/bas (défaut 45720 EMU = 0,05")
  return h / S + (+(tcPr['@marT'] ?? 45720)) + (+(tcPr['@marB'] ?? 45720));
}
// État de MODULE (et non de ctx) parce que cellTextHeight est exportée et appelée sans ctx.
// Écrit par renderSlide à chaque appel ; lu par cellTextHeight.
let _ctxMinorFont = 'Calibri'; // police mineure du thème de la diapo en cours (renderSlide)
// setMinorFont(nom) : renderSlide renseigne la police mineure du thème avant de rendre une diapo.
export function setMinorFont(f) { _ctxMinorFont = f || 'Calibri'; }

export { fontFamily, metricsOf, VA_TOK, applyShift, pptTextWidth };
