// Tests unitaires des fonctions PURES du moteur de rendu PPTX -> SVG.
//
// Portée : UNIQUEMENT des helpers déterministes (sortie = fonction des entrées,
// aucune I/O : pas de fs / réseau / DOM / puppeteer / état global). Les grosses
// fonctions à effets (renderSlide, loadContext, extractTiming) ne sont PAS testées
// ici : elles lisent des fichiers et dépendent du disque.
//
// La racine du dépôt est en `type: commonjs`, mais ce dossier `test/` porte son
// propre package.json `{"type":"module"}` : ce fichier .js est donc chargé en ESM
// et peut importer directement le module renderer `.mjs`.
//
// Lancement : `node --test test/`  (ou `node --test test/pptx-render.test.js`).

import test from 'node:test';
import assert from 'node:assert/strict';

import {
  u, esc, serialize,
  hslFromRgb, rgbFromHsl, cssColor, resolveColorFlat,
  symFontKind, symToUnicode, bulletSym, symbolText,
  fontFamily, bulletOf,
  tagOf, at, kids, byTag, first, deepFindFlat,
} from '../renderer/pptx-render.mjs';

// ───────────────────────── conversion géométrique ─────────────────────────
// u(v) convertit des EMU en points (1 pt = 12700 EMU), arrondi à 2 décimales.
test('u : EMU -> points (1 pt = 12700 EMU)', () => {
  assert.equal(u(12700), 1);            // exactement 1 point
  assert.equal(u(0), 0);
  assert.equal(u(6858000), 540);        // hauteur d'une diapo 16:9 (7,5 po -> 540 pt)
  assert.equal(u(190500), 15);          // 15 pt
  assert.equal(u(6350), 0.5);           // demi-point
  // arrondi à 2 décimales : 100 EMU = 0,007874... pt -> 0.01
  assert.equal(u(100), 0.01);
});

// ───────────────────────── échappement XML ─────────────────────────
// esc() protège &, <, >, " (mais PAS l'apostrophe) pour l'injection dans le SVG.
test('esc : échappe les 4 métacaractères XML, pas l\'apostrophe', () => {
  assert.equal(esc('a & b'), 'a &amp; b');
  assert.equal(esc('<x>'), '&lt;x&gt;');
  assert.equal(esc('"q"'), '&quot;q&quot;');
  assert.equal(esc("l'eau"), "l'eau");        // apostrophe conservée telle quelle
  assert.equal(esc('<a href="x">&'), '&lt;a href=&quot;x&quot;&gt;&amp;');
});
test('esc : null / undefined -> chaîne vide', () => {
  assert.equal(esc(null), '');
  assert.equal(esc(undefined), '');
  assert.equal(esc(0), '0');                  // 0 est rendu (String(0))
});

// ───────────────────────── couleurs : HSL <-> RGB ─────────────────────────
// hslFromRgb et rgbFromHsl doivent faire un aller-retour fidèle (à ±1 près sur 8 bits).
test('hslFromRgb : valeurs de référence', () => {
  assert.deepEqual(hslFromRgb(255, 0, 0), [0, 1, 0.5]);      // rouge pur
  assert.deepEqual(hslFromRgb(255, 255, 255), [0, 0, 1]);    // blanc (mx === mn)
  assert.deepEqual(hslFromRgb(0, 0, 0), [0, 0, 0]);          // noir
});
test('rgbFromHsl : valeurs de référence', () => {
  assert.deepEqual(rgbFromHsl(0, 1, 0.5), [255, 0, 0]);      // rouge pur
  assert.deepEqual(rgbFromHsl(0, 0, 1), [255, 255, 255]);    // blanc (s === 0)
  assert.deepEqual(rgbFromHsl(0, 0, 0), [0, 0, 0]);          // noir
});
test('HSL <-> RGB : aller-retour fidèle (±1)', () => {
  const cases = [
    [255, 0, 0], [0, 255, 0], [0, 0, 255],
    [128, 128, 128], [100, 150, 200], [17, 34, 51], [240, 12, 200],
  ];
  for (const [r, g, b] of cases) {
    const [h, s, l] = hslFromRgb(r, g, b);
    const [r2, g2, b2] = rgbFromHsl(h, s, l);
    assert.ok(Math.abs(r2 - r) <= 1, `R ${r}->${r2}`);
    assert.ok(Math.abs(g2 - g) <= 1, `G ${g}->${g2}`);
    assert.ok(Math.abs(b2 - b) <= 1, `B ${b}->${b2}`);
  }
});

// ───────────────────────── couleurs : format CSS ─────────────────────────
// cssColor() : rgb(...) si opaque, rgba(...) sinon, 'none' si null.
test('cssColor : rgb / rgba / none', () => {
  assert.equal(cssColor({ r: 1, g: 2, b: 3, a: 1 }), 'rgb(1,2,3)');
  assert.equal(cssColor({ r: 10, g: 20, b: 30, a: 0.5 }), 'rgba(10,20,30,0.500)');
  assert.equal(cssColor(null), 'none');
  assert.equal(cssColor(undefined), 'none');
});

// ───────────────────────── couleurs : résolution DrawingML ─────────────────────────
// resolveColorFlat() résout un nœud couleur (plat) en {r,g,b,a}, en appliquant
// les modificateurs (alpha/shade/tint/lumMod) et la substitution de thème.
const ctxBase = () => ({ clr: {}, clrMap: null, phClr: null });

test('resolveColorFlat : srgbClr direct', () => {
  const c = resolveColorFlat({ 'a:srgbClr': { '@val': 'FF0000' } }, ctxBase());
  assert.deepEqual(c, { r: 255, g: 0, b: 0, a: 1 });
});
test('resolveColorFlat : nœud absent -> null', () => {
  assert.equal(resolveColorFlat(null, ctxBase()), null);
  assert.equal(resolveColorFlat({ 'a:unknown': {} }, ctxBase()), null);
});
test('resolveColorFlat : alpha (opacité)', () => {
  const c = resolveColorFlat({ 'a:srgbClr': { '@val': 'FF0000', 'a:alpha': { '@val': '50000' } } }, ctxBase());
  assert.equal(c.a, 0.5);
});
test('resolveColorFlat : shade assombrit', () => {
  // shade 50% sur rouge pur -> composante R divisée par 2 (127.5 -> 128 arrondi)
  const c = resolveColorFlat({ 'a:srgbClr': { '@val': 'FF0000', 'a:shade': { '@val': '50000' } } }, ctxBase());
  assert.deepEqual(c, { r: 128, g: 0, b: 0, a: 1 });
});
test('resolveColorFlat : tint éclaircit', () => {
  // tint 50% : R = 255*0.5 + 255*0.5 = 255 ; G,B = 0*0.5 + 255*0.5 = 128
  const c = resolveColorFlat({ 'a:srgbClr': { '@val': 'FF0000', 'a:tint': { '@val': '50000' } } }, ctxBase());
  assert.deepEqual(c, { r: 255, g: 128, b: 128, a: 1 });
});
test('resolveColorFlat : lumMod assombrit un gris', () => {
  // gris 808080 (l≈0.502), lumMod 50% -> l≈0.251 -> ~64 par canal
  const c = resolveColorFlat({ 'a:srgbClr': { '@val': '808080', 'a:lumMod': { '@val': '50000' } } }, ctxBase());
  assert.deepEqual(c, { r: 64, g: 64, b: 64, a: 1 });
});
test('resolveColorFlat : schemeClr résolu via ctx.clr', () => {
  const ctx = ctxBase(); ctx.clr = { accent1: '00FF00' };
  const c = resolveColorFlat({ 'a:schemeClr': { '@val': 'accent1' } }, ctx);
  assert.deepEqual(c, { r: 0, g: 255, b: 0, a: 1 });
});
test('resolveColorFlat : phClr -> ctx.phClr, défaut accent1', () => {
  const ctx = ctxBase(); ctx.clr = { accent1: '112233', accent2: 'AABBCC' }; ctx.phClr = 'accent2';
  const c = resolveColorFlat({ 'a:schemeClr': { '@val': 'phClr' } }, ctx);
  assert.deepEqual(c, { r: 0xAA, g: 0xBB, b: 0xCC, a: 1 });
});
test('resolveColorFlat : clrMap remappe bg1/tx1 -> dk1/lt1', () => {
  const ctx = ctxBase(); ctx.clr = { dk1: '000000' }; ctx.clrMap = { tx1: 'dk1' };
  const c = resolveColorFlat({ 'a:schemeClr': { '@val': 'tx1' } }, ctx);
  assert.deepEqual(c, { r: 0, g: 0, b: 0, a: 1 });
});
test('resolveColorFlat : scrgbClr (composantes en millièmes de %)', () => {
  const c = resolveColorFlat({ 'a:scrgbClr': { '@r': '100000', '@g': '0', '@b': '0' } }, ctxBase());
  assert.deepEqual(c, { r: 255, g: 0, b: 0, a: 1 });
});

// ───────────────────────── polices symbole -> Unicode ─────────────────────────
test('symFontKind : classe la police symbole', () => {
  assert.equal(symFontKind('Wingdings'), 'wing');
  assert.equal(symFontKind('Webdings'), 'wing');
  assert.equal(symFontKind('Wingdings 2'), 'wing');   // sous-chaîne 'wingding'
  assert.equal(symFontKind('Symbol'), 'sym');         // Adobe Symbol exact
  assert.equal(symFontKind('Symbolic'), null);        // pas 'symbol' exact -> pas 'sym'
  assert.equal(symFontKind('Arial'), null);
  assert.equal(symFontKind(''), null);
  assert.equal(symFontKind(undefined), null);
});
test('symToUnicode wing : dingbat mappé, inconnu -> puce', () => {
  assert.equal(symToUnicode('l', 'wing'), '●');   // 0x6C
  assert.equal(symToUnicode('n', 'wing'), '■');   // 0x6E
  assert.equal(symToUnicode('z', 'wing'), '•');   // 0x7A non mappé -> puce générique
});
test('symToUnicode wing : décodage zone privée PUA (0xF000+octet)', () => {
  assert.equal(symToUnicode('', 'wing'), '●'); // 0xF06C -> 0x6C -> ●
});
test('symToUnicode sym : grec mappé, inconnu -> caractère d\'origine', () => {
  assert.equal(symToUnicode('a', 'sym'), 'α');
  assert.equal(symToUnicode('b', 'sym'), 'β');
  assert.equal(symToUnicode('#', 'sym'), '#');    // non mappé -> garde l'original
});
test('bulletSym : puce en police symbole -> {sym, symbolic}', () => {
  assert.deepEqual(bulletSym('l', 'Wingdings'), { sym: '●', symbolic: true });
  assert.deepEqual(bulletSym('•', 'Arial'), { sym: '•', symbolic: false });   // police normale
  assert.deepEqual(bulletSym(null, 'Wingdings'), { sym: null, symbolic: false }); // pas de char
});
test('symbolText : remappe un run entier, priorité à <a:sym>', () => {
  assert.equal(symbolText('a b', 'Symbol'), 'α β');          // espace conservé
  assert.equal(symbolText('l', 'Arial', 'Wingdings'), '●');  // symFont prioritaire sur latinFont
  assert.equal(symbolText('abc', 'Arial'), null);            // police non symbole -> null
  assert.equal(symbolText('', 'Symbol'), null);              // texte vide -> null
});

// ───────────────────────── police -> famille CSS ─────────────────────────
test('fontFamily : substituts libres et génériques', () => {
  assert.equal(fontFamily(undefined), 'sans-serif');
  assert.equal(fontFamily('Arial'), "'Arial','Liberation Sans',sans-serif");
  assert.equal(fontFamily('Times New Roman'), "'Times New Roman','Liberation Serif',serif");
  assert.equal(fontFamily('Georgia'), "'Georgia',serif");     // serif via SERIF_RE
  assert.equal(fontFamily('Verdana'), "'Verdana',sans-serif"); // défaut sans-serif
  assert.equal(fontFamily('Wingdings'), "'Wingdings'");        // police symbole : telle quelle
});

// ───────────────────────── lecture d'une puce (bulletOf) ─────────────────────────
test('bulletOf : pas de pPr -> null', () => {
  assert.equal(bulletOf(null), null);
  assert.equal(bulletOf({}), null);                       // aucun descripteur de puce
});
test('bulletOf : buNone', () => {
  assert.deepEqual(bulletOf({ 'a:buNone': '' }), { none: true });
});
test('bulletOf : puce caractère + police', () => {
  assert.deepEqual(
    bulletOf({ 'a:buChar': { '@char': '§' }, 'a:buFont': { '@typeface': 'Wingdings' } }),
    { char: '§', font: 'Wingdings', clr: null, szPct: null },
  );
});
test('bulletOf : puce numérotée', () => {
  assert.deepEqual(
    bulletOf({ 'a:buAutoNum': { '@type': 'arabicParenR' } }),
    { num: 'arabicParenR', font: undefined, clr: null, szPct: null },
  );
});
test('bulletOf : buSzPct converti en fraction', () => {
  const b = bulletOf({ 'a:buChar': { '@char': 'x' }, 'a:buSzPct': { '@val': '50000' } });
  assert.equal(b.szPct, 0.5);
  assert.equal(b.char, 'x');
});

// ───────────────────────── helpers d'arbre preserveOrder ─────────────────────────
// Un nœud preserveOrder = { '<tag>': [enfants...], ':@': { '@attr': val } }.
const sample = {
  'p:sp': [
    { 'p:nvSpPr': [], ':@': { '@id': '5' } },
    { 'a:t': [{ '#text': 'hello' }] },
    { 'a:t': [{ '#text': 'world' }] },
  ],
  ':@': { '@name': 'shape1' },
};
test('tagOf / at : tag et attributs d\'un nœud', () => {
  assert.equal(tagOf(sample), 'p:sp');
  assert.deepEqual(at(sample), { '@name': 'shape1' });
  assert.deepEqual(at({}), {});           // pas de ':@' -> {}
  assert.deepEqual(at(undefined), {});
});
test('kids : enfants (tableau) ou []', () => {
  assert.equal(kids(sample).length, 3);
  assert.deepEqual(kids({ 'a:t': 'texte' }), []); // valeur non-tableau -> []
  assert.deepEqual(kids(undefined), []);
});
test('byTag / first : filtre et premier par tag', () => {
  assert.equal(byTag(sample, 'a:t').length, 2);
  assert.equal(byTag(sample, 'p:nvSpPr').length, 1);
  assert.equal(byTag(sample, 'absent').length, 0);
  assert.deepEqual(at(first(sample, 'p:nvSpPr')), { '@id': '5' });
  assert.equal(first(sample, 'absent'), undefined);
});
test('deepFindFlat : 1re valeur trouvée en profondeur (objet plat)', () => {
  const obj = { 'p:sp': { 'p:spPr': { 'a:xfrm': { '@rot': '60000' } } } };
  assert.equal(deepFindFlat(obj, '@rot'), '60000');
  assert.equal(deepFindFlat({ x: 1 }, 'absent'), null);
  assert.equal(deepFindFlat(null, 'x'), null);
});

// ───────────────────────── sérialisation preserveOrder -> XML ─────────────────────────
test('serialize : reconstruit le XML avec attributs et texte échappés', () => {
  const node = { 'a:t': [{ '#text': 'hi' }], ':@': { '@x': '1' } };
  assert.equal(serialize(node), '<a:t x="1">hi</a:t>');
});
test('serialize : arbre imbriqué', () => {
  const node = { 'a:p': [{ 'a:r': [{ 'a:t': [{ '#text': 'X' }] }] }] };
  assert.equal(serialize(node), '<a:p><a:r><a:t>X</a:t></a:r></a:p>');
});
test('serialize : échappe les valeurs d\'attribut', () => {
  const node = { 'a:x': [], ':@': { '@v': 'a&b' } };
  assert.equal(serialize(node), '<a:x v="a&amp;b"></a:x>');
});
