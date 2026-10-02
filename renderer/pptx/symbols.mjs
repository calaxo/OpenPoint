// ═══ [PPTX · SYMBOLES] ═══ polices Wingdings / Webdings / Symbol -> Unicode (anti-tofu).
// ████ §H ─ POLICES SYMBOLES -> UNICODE ─────────────────────────────────────────────────────────
// Deux façons dont un .pptx utilise une police symbole :
//   (1) la police latin du run EST la police symbole (a:latin typeface="Wingdings") -> tout le texte
//       du run est à remapper (octets 0x20-0xFF interprétés dans le jeu de la police) ;
//   (2) le run porte <a:sym typeface="Symbol"> et contient des caractères de la zone privée
//       U+F000-F0FF (convention Windows pour les polices symboles : U+F000 + octet) -> seuls ces
//       caractères-là sont des symboles.
// Chaîne d'appel : bulletSym (puces) ; symbolParts / symbolText (texte des runs) -> symToUnicode.
// Polices SYMBOLE (Wingdings/Webdings/Symbol) : le glyphe n'existe qu'avec la police installée -> tofu
// sur Linux (cible LMS) et dans Chromium headless. On mappe les codepoints vers leur équivalent Unicode,
// rendus en police NORMALE = lisibles partout. Gère l'encodage zone privée PUA (0xF000+octet en run).
const WINGD_RE = /wingding|webding/i; // dingbats PURS (que des symboles) — décoratif
const SYMBOLFONT_RE = /^symbol$/i;      // Adobe « Symbol » = grec + flèches + math (PAS décoratif)
// Wingdings -> Unicode (puces/dingbats). Inconnu -> « • » (marqueur générique, jamais de tofu).
const WINGDINGS_MAP = {
  0x6C: '●', 0x6D: '○', 0x6E: '■', 0x6F: '□', 0x71: '❑', 0x75: '◆',
  0xA7: '▪', 0xA8: '▫', 0xB7: '•', 0xD8: '➢', 0xD9: '➢',
  // 0xFB-FE (table Wingdings) : ✗ croix, ✓ coche FINE (pas ✔ épaisse, cf. 040_02_01 d.57), ☒, ☑
  0xFB: '✗', 0xFC: '✓', 0xFD: '☒', 0xFE: '☑',
};
// Adobe « Symbol » -> Unicode : grec (ASCII->grec) + flèches (0xAB-AF, 0xDB-DF) + opérateurs math.
// Inconnu -> on GARDE le caractère d'origine (symbole signifiant : ne jamais deviner à tort).
const SYMBOL_MAP = (() => {
  const m = {};
  const low = 'abgdezhqiklmnxoprstufcyw', lowG = 'αβγδεζηθικλμνξοπρστυφχψω';
  const up  = 'ABGDEZHQIKLMNXOPRSTUFCYW', upG  = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ';
  // correspondance positionnelle : a->α, b->β, g->γ, d->δ, … (translittération de la police Symbol)
  for (let i = 0; i < low.length; i++) { m[low.charCodeAt(i)] = lowG[i]; m[up.charCodeAt(i)] = upG[i]; }
  // phi : la police Symbol DESSINE 0x66 « f » en ϕ (trait droit) et 0x6A « j » en φ (bouclé) -> afficher ce
  // que montre PowerPoint (l'encodage Adobe donne l'inverse, d'où des φ bouclés non conformes au support).
  m[0x66] = 'ϕ'; m[0x6A] = 'φ';
  Object.assign(m, {
    0xAB:'↔',0xAC:'←',0xAD:'↑',0xAE:'→',0xAF:'↓', 0xDB:'⇔',0xDC:'⇐',0xDD:'⇑',0xDE:'⇒',0xDF:'⇓',
    0xA3:'≤',0xB3:'≥',0xB9:'≠',0xBB:'≈',0xBA:'≡',0xB1:'±',0xB4:'×',0xB8:'÷',0xB0:'°',0xB7:'·',
    0x40:'≅',0xD6:'√',0xA5:'∞',0xB6:'∂',0xF2:'∫',0x2D:'−',0xE5:'∑',0xCF:'∨',0xCE:'∈',
  });
  return m;
})();
// symFontKind(typeface) -> 'wing' (Wingdings/Webdings), 'sym' (Adobe Symbol, nom EXACT « Symbol »
// pour ne pas attraper « Segoe UI Symbol ») ou null (police ordinaire).
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
// remappe un TEXTE de run en police symbole -> Unicode (anti-tofu). `symFont` = typeface de <a:sym>.
// Police latin symbole -> TOUT le texte est remappé. <a:sym> seul ne vaut que pour les caractères
// PUA U+F000-F0FF (c'est ainsi que PowerPoint l'applique) : le texte ordinaire d'un run qui porte
// <a:sym typeface="Symbol"> reste en police latin, sinon « The Earth » deviendrait « Τηε Εαρτη ».
const isPua = (c) => { const cp = c.codePointAt(0); return cp >= 0xF000 && cp <= 0xF0FF; };
// -> null (rien à remapper) ou segments [{ text, sym }] : `sym` = glyphe issu de la police symbole,
// à rendre dans SYM_FAMILY ; le reste garde la police latin du run (comme PowerPoint).
// Cas (1) : un seul segment { sym:true } avec tout le texte remappé (les blancs sont conservés tels
// quels). Cas (2) : segments alternés latin / symbole, consécutifs fusionnés.
export function symbolParts(raw, latinFont, symFont) {
  if (!raw) return null;
  const latinKind = symFontKind(latinFont);
  if (latinKind) return [{ text: [...raw].map(c => /\s/.test(c) ? c : symToUnicode(c, latinKind)).join(''), sym: true }];
  const kind = symFontKind(symFont);
  if (!kind || ![...raw].some(isPua)) return null;
  const parts = [];
  for (const c of raw) {
    const sym = isPua(c), text = sym ? symToUnicode(c, kind) : c, last = parts[parts.length - 1];
    if (last && last.sym === sym) last.text += text; else parts.push({ text, sym });
  }
  return parts;
}
// symbolText : même chose aplatie en une chaîne (null si rien à remapper). Export pour les outils /
// tests ; non utilisée dans ce fichier (renderText appelle symbolParts).
export function symbolText(raw, latinFont, symFont) {
  const parts = symbolParts(raw, latinFont, symFont);
  return parts && parts.map(p => p.text).join('');
}
// Police des glyphes remappés : la Symbol d'Adobe a un dessin Times -> Liberation Serif, puis
// DejaVu Sans (embarquée) pour ce que Liberation ne couvre pas (≅, ⇒, ✔, ➢…).
const SYM_FAMILY = "'Times New Roman','Liberation Serif','DejaVu Sans',serif";
export { symFontKind, symToUnicode, bulletSym, SYM_FAMILY };
