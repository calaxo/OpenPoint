// Tests du calage du texte sur PowerPoint : rectangle de texte des formes, ligne de base, chasses.
//
// POURQUOI : ces trois modèles viennent d'un CALIBRAGE contre PowerPoint (fichiers générés et
// exportés via COM, mesurés au pixel). Sans eux, le texte tombait jusqu'à 40 pt trop haut ou trop
// bas selon l'interligne, et des mots passaient à la ligne plus tôt que dans le support.
// Les valeurs attendues ci-dessous sont celles du modèle calibré : si un test casse, c'est que
// le rendu ne colle plus à PowerPoint.
import { test } from "node:test";
import assert from "node:assert/strict";
import { presetTextRect } from "../renderer/preset-geom.mjs";
import { baselineShift, advanceFix, cellTextHeight, mathTypeset } from "../renderer/pptx-render.mjs";

const POP = "'Poppins',sans-serif";

test("géométrie libre : stroke=\"0\" (booléen xsd) = tracé NON dessiné (arc ouvert, pas un secteur)", async () => {
  const { evalGeomParts } = await import("../renderer/preset-geom.mjs");
  // arc de 022_04 d.44 : path 1 = contour sans fond, path 2 = secteur rempli SANS contour
  const def = { av: [], gd: [], paths: [
    { w: 100, h: 100, fill: "none", stroke: true, cmds: [["M", "0", "0"], ["L", "100", "100"]] },
    { w: 100, h: 100, fill: "norm", stroke: !["false", "0"].includes("0"), cmds: [["M", "0", "0"], ["L", "100", "100"], ["L", "0", "100"], ["Z"]] },
  ] };
  const p = evalGeomParts(def, 100, 100);
  assert.doesNotMatch(p.stroke, /Z/); // le secteur fermé n'est pas tracé
});

test("rectangle de texte : rect = boîte entière, roundRect rétréci de 0,29 x rayon", () => {
  assert.deepEqual(presetTextRect("rect", 100, 50, {}), { l: 0, t: 0, r: 100, b: 50 });
  // pastille ronde (adj 50 000 : rayon = min(w,h)/2 = 250) -> il = 250 x 0,29289
  const r = presetTextRect("roundRect", 1000, 500, { adj: "val 50000" });
  assert.ok(Math.abs(r.l - 73.2225) < 0.01 && Math.abs(r.t - r.l) < 1e-9);
  assert.ok(Math.abs(r.r - (1000 - r.l)) < 1e-9 && Math.abs(r.b - (500 - r.l)) < 1e-9);
  assert.equal(presetTextRect("formeInconnue", 10, 10, {}), null);
});

test("ligne de base, interligne simple : L - min(0,25 L, descendante win)", () => {
  // Poppins 60 px, L = 1,2 x 60 = 72 : PowerPoint 54 ; navigateur 36 + (1,135 - 0,627)/2 x 60 = 51,24
  assert.ok(Math.abs(baselineShift(POP, 60, 72) - 2.76) < 0.02);
});

test("ligne de base, interligne > 100 % : 0,75 L quelle que soit la police", () => {
  // L = 1,5 x 72 = 108 : PowerPoint 81 ; navigateur 54 + 15,24
  assert.ok(Math.abs(baselineShift(POP, 60, 108) - 11.76) < 0.02);
});

test("ligne de base : police sans métriques connues -> aucun décalage", () => {
  assert.equal(baselineShift("'PoliceInconnue',sans-serif", 20, 24), 0);
  assert.equal(baselineShift(POP, 0, 24), 0);
});

test("chasses : PowerPoint compose Poppins 18 pt plus serré (arrondi au 1/8 pt)", () => {
  const txt = "where the axis of rotation is constrained to";
  const f = advanceFix(POP, false, false, 18, txt);
  assert.ok(f < 0, `correction négative attendue, obtenu ${f}`);
  // l'arrondi au 1/8 pt ne peut décaler un glyphe de plus d'1/16 pt
  assert.ok(Math.abs(f) <= 1 / 16);
  // total : ~0,36 % de moins sur ~381 px (calibré via COM : ratio PowerPoint/Chrome 0,9964)
  assert.ok(Math.abs(f * [...txt].length + 1.37) < 0.3);
});

test("chasses : caractère hors table, police inconnue ou texte vide -> 0", () => {
  assert.equal(advanceFix(POP, false, false, 18, "abc\u{1F600}"), 0);
  assert.equal(advanceFix("'PoliceInconnue'", false, false, 18, "abc"), 0);
  assert.equal(advanceFix(POP, false, false, 18, ""), 0);
});

test("formules OMML : lettres en italique, chiffres droits, opérateurs espacés, styles m:sty", () => {
  assert.equal(mathTypeset("x+2", null),
    '<span style="font-style:italic;">x</span><span style="padding:0 .22em">+</span>2');
  assert.equal(mathTypeset("P<500", "p"), 'P<span style="padding:0 .28em">&lt;</span>500');
  // alphanumériques mathématiques : 𝑷 (italique) et 𝟓 (gras) gardent leur style
  assert.equal(mathTypeset("\u{1D443}\u{1D7D3}", null),
    '<span style="font-style:italic;">P</span><span style="font-weight:700;">5</span>');
});

test("cellule de tableau : hauteur = lignes coupées comme PowerPoint x 1,2 x taille + marges", () => {
  const cell = (t) => ({ "a:tcPr": {}, "a:txBody": { "a:p": { "a:r": { "a:rPr": { "@sz": "1400", "a:latin": { "@typeface": "Arial" } }, "a:t": t } } } });
  const W = 2 * 914400; // 2 pouces
  // 1 ligne : 1,2 x 14 pt + marges haut/bas 2 x 3,6 pt = 24 pt
  const one = cellTextHeight(cell("Heredity"), W, false) / 12700;
  assert.ok(Math.abs(one - 24) < 0.01, `obtenu ${one}`);
  const long = cellTextHeight(cell("Insufficient physical activity and a sedentary lifestyle over many years"), W, false) / 12700;
  assert.ok(long >= 24 + 2 * 16.8 - 0.01, `au moins 3 lignes attendues, obtenu ${long}`);
  assert.equal(cellTextHeight(cell("x\u{1F600}"), W, false), null); // glyphe non mesurable -> null
});
