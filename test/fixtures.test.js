// Tests d'intégration sur les decks synthétiques de test/fixtures/decks.mjs (aucun deck privé).
import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { unzipSync } from 'fflate';
import { buildDecks, writeDeck, zipDeck } from './fixtures/decks.mjs';
import { renderSlide, extractTiming, extractTransition, slideOrder } from '../renderer/pptx-render.mjs';

const decks = buildDecks();
const root = mkdtempSync(join(tmpdir(), 'openpoint-fx-'));
const dirs = {};
for (const [name, files] of Object.entries(decks)) { dirs[name] = join(root, name); writeDeck(files, dirs[name]); }
test.after(() => rmSync(root, { recursive: true, force: true }));
const ppt = (d) => join(dirs[d], 'ppt');
const nSlides = (d) => Object.keys(decks[d]).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p)).length;
const r = (d, n) => renderSlide(ppt(d), n);
const spids = (svg) => [...svg.matchAll(/data-spid="([^"]*)"/g)].map((m) => m[1]);
const plain = (svg) => svg.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');

test('zipDeck produit un .pptx décompressable avec [Content_Types].xml', () => {
  const z = unzipSync(zipDeck(decks.animations));
  assert.ok(z['[Content_Types].xml'] && z['ppt/presentation.xml'] && z['ppt/slides/slide1.xml']);
});

test('chaque diapo de chaque deck se rend en SVG avec viewBox 720x405', () => {
  for (const d of Object.keys(decks)) for (let n = 1; n <= nSlides(d); n++) {
    const { svg, w, h } = r(d, n);
    assert.match(svg, /^<svg /, `${d} d.${n}`);
    assert.match(svg, /viewBox="0 0 720 405"/, `${d} d.${n}`);
    assert.equal(w, 9144000); assert.equal(h, 5143500);
  }
});

// ───── animations ─────
test('animations d.1 : 3 clics fade / wipe(down) / zoom', () => {
  const t = extractTiming(ppt('animations'), 1);
  assert.equal(t.steps.length, 3);
  assert.deepEqual(t.steps.map((s) => s.effects[0].spid), ['3', '4', '5']);
  assert.deepEqual(t.steps.map((s) => s.effects[0].cls), ['entr', 'entr', 'entr']);
  assert.deepEqual(t.steps.map((s) => s.effects[0].filter), ['fade', 'wipe(down)', 'fade']);
  assert.equal(t.steps[2].effects[0].dur, 600);
  assert.deepEqual([...t.animatedSpids].sort(), ['3', '4', '5']);
  const { svg } = r('animations', 1);
  for (const id of ['3', '4', '5']) assert.ok(spids(svg).includes(id));
  assert.ok(plain(svg).includes('Balayage'));
});

test('animations d.2 : texte par paragraphe (data-par, plages pRg)', () => {
  const t = extractTiming(ppt('animations'), 2);
  assert.equal(t.steps.length, 3);
  assert.deepEqual(t.steps.map((s) => s.effects[0].par), [{ from: 0, to: 0 }, { from: 1, to: 1 }, { from: 2, to: 2 }]);
  assert.deepEqual(t.animatedSpids, []); // entrées par paragraphe : pas de forme entière cachée
  const { svg } = r('animations', 2);
  assert.equal((svg.match(/data-par="/g) || []).length, 4); // 1 étiquette + 3 paragraphes
  assert.ok(plain(svg).includes('Deuxième point'));
});

test('animations d.3 : sortie puis emphase', () => {
  const t = extractTiming(ppt('animations'), 3);
  assert.deepEqual(t.steps.map((s) => s.effects[0].cls), ['exit', 'emph']);
  assert.deepEqual(t.steps.map((s) => s.effects[0].spid), ['3', '4']);
  assert.equal(t.steps[0].effects[0].filter, 'fade');
  assert.deepEqual(t.animatedSpids, []);
});

test('animations d.4 : avec / après le précédent (délais cumulés)', () => {
  const t = extractTiming(ppt('animations'), 4);
  assert.equal(t.steps.length, 1);
  const e = t.steps[0].effects;
  assert.deepEqual(e.map((x) => x.spid), ['3', '4', '5', '6']);
  assert.deepEqual(e.map((x) => x.delay), [0, 250, 750, 1050]); // groupe « après » = fin du 1er groupe (250+500)
  assert.deepEqual(e.map((x) => x.dur), [500, 500, 400, 500]);
  assert.equal(e[3].filter, 'wipe(down)');
});
test('animations d.4 : @nodeType exposé (node) ; étape déclenchée au clic', () => {
  const t = extractTiming(ppt('animations'), 4);
  assert.deepEqual(t.steps[0].effects.map((x) => x.node), ['clickEffect', 'withEffect', 'afterEffect', 'withEffect']);
  assert.equal(t.steps[0].trigger, 'click');
});

test('animations d.9 : 1re étape automatique (afterPrev), puis un clic', () => {
  const t = extractTiming(ppt('animations'), 9);
  assert.deepEqual(t.steps.map((s) => s.trigger), ['afterPrev', 'click']);
  assert.deepEqual(t.animatedSpids, ['3', '4']);
});

test('animations d.5 : trajectoire (animMotion) et rotation (animRot)', () => {
  const t = extractTiming(ppt('animations'), 5);
  const [m, rot] = t.steps.map((s) => s.effects[0]);
  assert.equal(m.cls, 'path');
  assert.equal(m.motion.kind, 'move'); assert.equal(m.motion.dx, 0.25); assert.equal(m.motion.dy, 0.1); assert.equal(m.motion.dur, 1500);
  assert.equal(rot.cls, 'emph');
  assert.deepEqual({ k: rot.motion.kind, deg: rot.motion.deg, dur: rot.motion.dur }, { k: 'rot', deg: 360, dur: 1000 });
});

test('animations d.6-8 : transitions fade / push u / morph', () => {
  assert.deepEqual(extractTransition(ppt('animations'), 6), { type: 'fade', dur: 700, dir: null });
  assert.deepEqual(extractTransition(ppt('animations'), 7), { type: 'push', dur: 750, dir: 'u' });
  assert.deepEqual(extractTransition(ppt('animations'), 8), { type: 'morph', dur: 1500, dir: null });
  assert.equal(extractTransition(ppt('animations'), 1), null);
  assert.deepEqual(extractTiming(ppt('animations'), 6).steps, []);
});

// ───── shapes ─────
test('shapes d.1 : présets, custGeom, dégradés, motifs, ombre', () => {
  const { svg } = r('shapes', 1);
  for (let i = 3; i <= 13; i++) assert.ok(spids(svg).includes(String(i)), `spid ${i}`);
  assert.match(svg, /<linearGradient/);
  assert.match(svg, /<pattern /);
  assert.match(svg, /<filter/); // ombre
  assert.ok(plain(svg).includes('Avec ombre'));
});
test('shapes d.2 : lignes, pointillés, têtes de flèche', () => {
  const { svg } = r('shapes', 2);
  for (let i = 3; i <= 8; i++) assert.ok(spids(svg).includes(String(i)), `spid ${i}`);
  assert.match(svg, /stroke-dasharray/);
  assert.match(svg, /<marker/);
});
test('shapes d.2 : forme hidden="1" non dessinée ; showMasterSp="0" masque le décor du master', () => {
  const { svg } = r('shapes', 2);
  assert.ok(!plain(svg).includes('NE DOIT PAS'));
  assert.ok(!spids(svg).includes('9'));
  assert.ok(!spids(svg).includes('m5'));
  assert.ok(spids(r('shapes', 1).svg).includes('m5')); // diapo normale : décor du master présent
});
test('shapes d.3 : groupes imbriqués, rotation et miroir', () => {
  const { svg } = r('shapes', 3);
  for (const i of ['3', '4', '5', '6', '7', '8', '9', '10']) assert.ok(spids(svg).includes(i), `spid ${i}`);
  assert.match(svg, /rotate\(30/);
  assert.match(svg, /scale\(-1 1\)/);
});

// ───── text ─────
test('text d.1 : placeholders hérités du gabarit, puces', () => {
  const a = r('text', 1).svg;
  assert.ok(plain(a).includes('Titre hérité du gabarit'));
  assert.ok(plain(a).includes('Puce de niveau 2'));
  assert.equal((a.match(/data-par="/g) || []).length, 5); // titre + 4 paragraphes
  assert.match(a, /•/);
  assert.match(a, /1\./); assert.match(a, /2\./);
});
test('text d.2 : runs, autofit, texte vertical, tabulations', () => {
  const b = r('text', 2).svg;
  for (const t of ['Gras', 'italique', 'souligné', 'rouge', 'grand', 'Texte réduit', 'Texte vertical', 'Nom', 'Quantité', 'Prix']) assert.ok(plain(b).includes(t), t);
  assert.match(b, /font-weight:\s*(bold|700)/);
  assert.match(b, /italic/); assert.match(b, /underline/);
  assert.match(b, /vertical|writing-mode/);
  assert.match(b, /data-af="norm"/);
  assert.ok(plain(b).includes('Diapo 1')); // champ slidenum = position dans la présentation (slide2 est 1re)
});
test('text d.3 : diapo sombre, fond foncé', () => {
  const { svg } = r('text', 3);
  assert.ok(plain(svg).includes('Texte clair sur fond sombre'));
  assert.match(svg, /rgb[(]27,27,47[)]/); // fond sombre explicite de la diapo (p:bg)
});
test('text : la présentation liste slide2 avant slide1 (fichiers inchangés)', () => {
  const p = decks.text['ppt/presentation.xml'];
  assert.ok(p.indexOf('rId12') < p.indexOf('rId11'));
});
test('text : slideOrder suit p:sldIdLst (slide2, slide1, slide3, slide4)', () => {
  assert.deepEqual(slideOrder(ppt('text')), [2, 1, 3, 4]);
  assert.deepEqual(slideOrder(ppt('shapes')), [1, 2, 3]);
});
test('text d.4 : clrMapOvr du layout (bg1->dk1, tx1->lt1) appliqué au fond et au texte', () => {
  const { svg } = r('text', 4);
  assert.ok(plain(svg).includes('Couleurs héritées du layout sombre'));
  assert.match(svg, /<rect width="720" height="405" fill="rgb\(31,31,31\)"\/>/); // fond bg1 = dk1
  assert.match(svg, /color:rgb\(238,241,246\)/); // titre en tx2 (titleStyle du master) -> lt2 clair
});

// ───── tables-charts ─────
test('tables-charts d.1 : tableau (style, bandes, fusion)', () => {
  const { svg } = r('tables-charts', 1);
  assert.ok(spids(svg).includes('3'));
  for (const t of ['Produit', 'Stylo', 'Fusion sur deux colonnes', 'Cahier', '300']) assert.ok(plain(svg).includes(t), t);
  assert.match(svg, /rgb[(]46,117,182[)]/); // en-tête du style Medium Style 2 - Accent 1
});
test('tables-charts d.2 : histogramme et camembert', () => {
  const { svg } = r('tables-charts', 2);
  assert.ok(plain(svg).includes('Ventes par trimestre'));
  assert.ok(plain(svg).includes('Répartition'));
  for (const t of ['T1', 'T4', '19']) assert.ok(plain(svg).includes(t), t);
  assert.match(svg, /<path /); // parts du camembert
});
test('tables-charts d.3 : images entière / rognée / ellipse', () => {
  const { svg } = renderSlide(ppt('tables-charts'), 3, { embedImages: true });
  assert.equal((svg.match(/<image /g) || []).length, 3);
  assert.match(svg, /data:image\/png;base64,/);
  assert.equal((svg.match(/<clipPath /g) || []).length, 2); // rognage (rect) + ellipse
});

test('text d.4 : placeholder « titre idx=7 » sans équivalent exact -> style du titre du master (par type)', () => {
  const { svg } = r('text', 4);
  const i = svg.indexOf('>01<');
  const span = svg.slice(svg.lastIndexOf('<span', i), i);
  assert.match(span, /font-size:32\.00px/); // titleStyle du master : sz 3200
  assert.match(span, /font-weight:700/);    // b="1"
});

test('animations : métadonnées d\'effet (preset, zoom, scale, fly)', () => {
  const s1 = extractTiming(ppt('animations'), 1).steps;
  assert.equal(s1[2].effects[0].preset, 53);
  assert.equal(s1[2].effects[0].zoom, true);   // ppt_w / ppt_h depuis 0
  assert.equal(s1[0].effects[0].zoom, false);
  assert.deepEqual(extractTiming(ppt('animations'), 3).steps[1].effects[0].scale, { x: 1.5, y: 1.5 });
  const s10 = extractTiming(ppt('animations'), 10).steps;
  assert.equal(s10[0].effects[0].fly, 'l');
  assert.equal(s10[0].effects[0].subtype, 8);
  assert.deepEqual(s10[1].effects[0].scale, { x: 1.5, y: 1.5 });
});
