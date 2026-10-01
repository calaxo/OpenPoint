// Tests des fonctions PURES de renderer/docx-build.mjs et renderer/chart-render.mjs.
//
// Pourquoi ces fonctions et pas d'autres ?
//   On ne teste QUE des fonctions pures : leur sortie ne dépend que de leurs entrées,
//   sans lecture disque, réseau, DOM, ni état global mutable. Les fonctions qui lisent
//   un .docx, parsent le XML global ou écrivent un fichier (renderDocx, loadStyles...)
//   sont volontairement exclues : elles ne sont pas des fonctions pures.
//
// Exécution : node --test test/docx-chart.test.js   (Node >= 22 détecte l'ESM dans un .js).
import { test } from 'node:test';
import assert from 'node:assert/strict';

import {
  tagOf, kids, at, mrg, esc as escDocx, child, childAttrs,
  onOff, twip, emuPx, bulletChar, fontFamily, resolveColor,
} from '../renderer/docx-build.mjs';

import {
  esc as escChart, arr, num, ptsOf, serOf, renderChart,
} from '../renderer/chart-render.mjs';

// ─────────────────────────────────────────────────────────────────────────────
// docx-build.mjs — helpers "preserveOrder"
//
// fast-xml-parser en mode preserveOrder représente un noeud ainsi :
//   { 'w:b': [ ...enfants... ], ':@': { '@w:val': 'true' } }
// tagOf renvoie la 1re clé qui n'est ni ':@' (attributs) ni '#text' (texte).
// ─────────────────────────────────────────────────────────────────────────────

test('tagOf renvoie le nom de balise et ignore ":@"/"#text"', () => {
  assert.equal(tagOf({ 'w:p': [], ':@': { '@x': '1' } }), 'w:p');
  // ordre des clés : ':@' d'abord ne doit pas masquer la vraie balise
  assert.equal(tagOf({ ':@': { '@x': '1' }, 'w:r': [] }), 'w:r');
  // aucun tag -> null
  assert.equal(tagOf({ ':@': {}, '#text': 'x' }), null);
});

test('kids renvoie le tableau d\'enfants (ou [])', () => {
  const enfants = [{ 'w:r': [] }];
  assert.equal(kids({ 'w:p': enfants }), enfants);
  // balise auto-fermante -> tableau vide, jamais undefined
  assert.deepEqual(kids({ 'w:b': [] }), []);
});

test('at renvoie les attributs, {} si absent ou noeud nul', () => {
  assert.deepEqual(at({ 'w:sz': [], ':@': { '@w:val': '24' } }), { '@w:val': '24' });
  assert.deepEqual(at({ 'w:b': [] }), {});
  assert.deepEqual(at(null), {});
});

test('child / childAttrs cherchent le 1er enfant du tag donné', () => {
  const pr = { 'w:rPr': [
    { 'w:b': [] },
    { 'w:sz': [], ':@': { '@w:val': '24' } },
  ] };
  assert.equal(tagOf(child(pr, 'w:sz')), 'w:sz');
  assert.equal(child(pr, 'w:absent'), undefined);
  assert.deepEqual(childAttrs(pr, 'w:sz'), { '@w:val': '24' });
  // enfant présent mais sans attributs -> {} ; enfant absent -> null
  assert.deepEqual(childAttrs(pr, 'w:b'), {});
  assert.equal(childAttrs(pr, 'w:absent'), null);
});

test('onOff interprète un toggle Word (<w:b/>, w:val=true/1/on/false, absent)', () => {
  const pr = { 'w:rPr': [
    { 'w:b': [] },                                     // présent, pas de val -> true
    { 'w:i': [], ':@': { '@w:val': 'false' } },        // val=false -> false
    { 'w:strike': [], ':@': { '@w:val': '1' } },       // val=1 -> true
    { 'w:caps': [], ':@': { '@w:val': 'on' } },        // val=on -> true
  ] };
  assert.equal(onOff(pr, 'w:b'), true);
  assert.equal(onOff(pr, 'w:i'), false);
  assert.equal(onOff(pr, 'w:strike'), true);
  assert.equal(onOff(pr, 'w:caps'), true);
  assert.equal(onOff(pr, 'w:absent'), null); // absent -> null (distinct de false)
});

// ─────────────────────────────────────────────────────────────────────────────
// docx-build.mjs — conversions d'unités et de style
// ─────────────────────────────────────────────────────────────────────────────

test('twip convertit les twips en px (1440 twips = 96 px)', () => {
  assert.equal(twip('1440'), 96);
  assert.equal(twip('15'), 1);
  assert.equal(twip(undefined), 0); // valeur manquante -> 0, pas NaN
  assert.equal(twip('0'), 0);
});

test('emuPx convertit les EMU en px (9525 EMU = 1 px)', () => {
  assert.equal(emuPx('9525'), 1);
  assert.equal(emuPx(0), 0);
  assert.equal(emuPx(undefined), 0);
});

test('mrg fusionne "le dernier défini gagne" en ignorant null/undefined', () => {
  assert.deepEqual(mrg({ a: 1, b: 2 }, { b: null, c: 3 }), { a: 1, b: 2, c: 3 });
  assert.deepEqual(mrg({ a: 1 }, { a: 5 }), { a: 5 });          // le dernier gagne
  assert.deepEqual(mrg({ a: 1 }, undefined, { d: 4 }), { a: 1, d: 4 }); // source nulle ignorée
  // une valeur undefined dans une source ne doit pas écraser une valeur définie
  assert.deepEqual(mrg({ a: 1 }, { a: undefined }), { a: 1 });
});

test('esc (docx) échappe &, < et >', () => {
  assert.equal(escDocx('a & b <c> "d"'), 'a &amp; b &lt;c&gt; "d"');
  assert.equal(escDocx(42), '42'); // coercition en chaîne
});

test('fontFamily mappe les polices connues et devine serif/sans', () => {
  assert.equal(fontFamily(undefined), 'inherit');
  assert.equal(fontFamily('Calibri'), "'Calibri','Carlito',sans-serif");   // substitut connu
  assert.equal(fontFamily('Times New Roman'), "'Times New Roman','Liberation Serif',serif");
  assert.equal(fontFamily('Garamond'), "'Garamond',serif");                // heuristique serif
  assert.equal(fontFamily('Foobar'), "'Foobar',sans-serif");               // défaut sans-serif
});

test('bulletChar mappe les puces symbole (Wingdings/PUA) vers Unicode', () => {
  assert.equal(bulletChar('', undefined), '•');                 // vide -> puce par défaut
  assert.equal(bulletChar(''), '●');                      // PUA 0xF06C -> 0x6C -> disque plein
  assert.equal(bulletChar(''), '•');                      // PUA 0xF0B7 -> 0xB7 -> puce
  assert.equal(bulletChar('x', 'Wingdings'), '•');              // police symbole, code inconnu -> défaut
  assert.equal(bulletChar('-', undefined), '-');                // caractère normal conservé tel quel
});

test('resolveColor résout couleur de thème puis valeur brute', () => {
  const theme = { accent1: 'ABCDEF', dk1: '000000' };
  assert.equal(resolveColor(null, 'accent1', theme), '#ABCDEF'); // themeColor prioritaire
  assert.equal(resolveColor(null, 'dark1', theme), '#000000');   // alias dark1 -> dk1
  assert.equal(resolveColor('FF0000', null, theme), '#FF0000');  // valeur srgb brute
  assert.equal(resolveColor('auto', null, theme), null);         // "auto" -> pas de couleur
  assert.equal(resolveColor(null, null, theme), null);           // rien -> null
});

// ─────────────────────────────────────────────────────────────────────────────
// chart-render.mjs — helpers purs
// ─────────────────────────────────────────────────────────────────────────────

test('esc (chart) échappe &,<,> et tolère null/undefined', () => {
  assert.equal(escChart('<a & b>'), '&lt;a &amp; b&gt;');
  assert.equal(escChart(null), '');       // null -> chaîne vide
  assert.equal(escChart(undefined), '');
});

test('arr normalise en tableau', () => {
  assert.deepEqual(arr(undefined), []);
  assert.deepEqual(arr([1, 2]), [1, 2]);
  assert.deepEqual(arr('x'), ['x']);      // scalaire -> tableau à 1 élément
  assert.deepEqual(arr(0), [0]);          // 0 n'est pas "undefined" -> [0]
});

test('num parse un nombre, 0 si non fini', () => {
  assert.equal(num('3.5'), 3.5);
  assert.equal(num('42'), 42);
  assert.equal(num('abc'), 0);
  assert.equal(num(undefined), 0);
  assert.equal(num('1e3'), 1000);
});

test('ptsOf extrait un cache en tableau dense indexé par @idx', () => {
  const ref = { 'c:numCache': {
    'c:ptCount': { '@val': '3' },
    'c:pt': [
      { '@idx': '0', 'c:v': '10' },
      { '@idx': '2', 'c:v': '30' }, // idx 1 absent -> trou (null)
    ],
  } };
  assert.deepEqual(ptsOf(ref), ['10', null, '30']);
  // pas de cache -> []
  assert.deepEqual(ptsOf(undefined), []);
  assert.deepEqual(ptsOf({}), []);
  // un seul point (objet non-tableau) est bien géré
  const one = { 'c:strCache': { 'c:ptCount': { '@val': '1' }, 'c:pt': { '@idx': '0', 'c:v': 'A' } } };
  assert.deepEqual(ptsOf(one), ['A']);
});

test('serOf construit une série {name,cats,vals,fill,pts}', () => {
  const strCache = (vals) => ({ 'c:strCache': {
    'c:ptCount': { '@val': String(vals.length) },
    'c:pt': vals.map((v, i) => ({ '@idx': String(i), 'c:v': v })),
  } });
  const numCache = (vals) => ({ 'c:numCache': {
    'c:ptCount': { '@val': String(vals.length) },
    'c:pt': vals.map((v, i) => ({ '@idx': String(i), 'c:v': String(v) })),
  } });
  const ser = {
    'c:tx': { 'c:strRef': strCache(['Série A']) },
    'c:cat': { 'c:strRef': strCache(['Jan', 'Fév']) },
    'c:val': { 'c:numRef': numCache([5, 7]) },
  };
  const s = serOf(ser, () => null); // color() renvoie null : aucune couleur explicite
  assert.equal(s.name, 'Série A');
  assert.deepEqual(s.cats, ['Jan', 'Fév']);
  assert.deepEqual(s.vals, [5, 7]);   // convertis en nombres
  assert.equal(s.fill, null);          // pas de c:spPr
  assert.deepEqual(s.pts, []);         // pas de c:dPt
});

test('serOf : nom via c:tx>c:v et couleur via color(spPr)', () => {
  const ser = {
    'c:tx': { 'c:v': 'Direct' },        // nom littéral (pas de strRef)
    'c:spPr': { marqueur: true },        // objet opaque passé tel quel à color()
    'c:val': { 'c:numRef': { 'c:numCache': { 'c:ptCount': { '@val': '1' }, 'c:pt': { '@idx': '0', 'c:v': '9' } } } },
  };
  // color() est injecté : la fonction reste PURE (sa sortie dépend de l'argument fourni)
  const s = serOf(ser, (spPr) => (spPr?.marqueur ? '#123456' : null));
  assert.equal(s.name, 'Direct');
  assert.deepEqual(s.vals, [9]);
  assert.equal(s.fill, '#123456');
});

// ─────────────────────────────────────────────────────────────────────────────
// chart-render.mjs — renderChart (pure : entrées -> chaîne SVG, aucun effet de bord)
// On vérifie la forme du SVG produit, pas le pixel exact.
// ─────────────────────────────────────────────────────────────────────────────

// fabrique un cache prêt pour la structure "à plat" attendue par renderChart
const numRef = (vals) => ({ 'c:numRef': { 'c:numCache': {
  'c:ptCount': { '@val': String(vals.length) },
  'c:pt': vals.map((v, i) => ({ '@idx': String(i), 'c:v': String(v) })),
} } });
const strRef = (vals) => ({ 'c:strRef': { 'c:strCache': {
  'c:ptCount': { '@val': String(vals.length) },
  'c:pt': vals.map((v, i) => ({ '@idx': String(i), 'c:v': v })),
} } });

test('renderChart : histogramme -> SVG avec des <rect>', () => {
  const flat = { 'c:chartSpace': { 'c:chart': { 'c:plotArea': {
    'c:barChart': {
      'c:barDir': { '@val': 'col' },
      'c:grouping': { '@val': 'clustered' },
      'c:ser': {
        'c:cat': strRef(['A', 'B', 'C']),
        'c:val': numRef([1, 2, 3]),
      },
    },
    'c:valAx': {},
  } } } };
  const svg = renderChart(flat, 300, 200);
  assert.equal(typeof svg, 'string');
  assert.match(svg, /<rect/);           // au moins une barre
  assert.match(svg, /<g transform=/);   // groupe de tracé
});

test('renderChart : camembert -> SVG avec des <path>', () => {
  const flat = { 'c:chartSpace': { 'c:chart': { 'c:plotArea': {
    'c:pieChart': {
      'c:ser': {
        'c:cat': strRef(['x', 'y']),
        'c:val': numRef([30, 70]),
      },
    },
  } } } };
  const svg = renderChart(flat, 200, 200);
  assert.match(svg, /<path/);
});

test('renderChart : plotArea absent -> chaîne vide', () => {
  assert.equal(renderChart({ 'c:chartSpace': { 'c:chart': {} } }, 100, 100), '');
  assert.equal(renderChart({}, 100, 100), '');
});
