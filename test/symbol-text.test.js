// Tests du remappage des polices symbole dans le texte des runs PPTX — `symbolText`.
//
// POURQUOI : PowerPoint pose souvent <a:sym typeface="Symbol"> sur des runs de texte ordinaire
// (cas réel : cours 022_04 GYROSCOPIC INSTRUMENTS, pages 29-31, 44-46, 59-60). <a:sym> ne vaut que
// pour les caractères PUA U+F000-F0FF ; remapper tout le run affichait « Τηε Εαρτη » au lieu
// de « The Earth ». Seule une police LATIN Symbol justifie de remapper tout le texte.
import { test } from "node:test";
import assert from "node:assert/strict";
import { symbolText, symbolParts } from "../renderer/pptx-render.mjs";

test("a:sym Symbol + texte ordinaire → inchangé (null)", () => {
  assert.equal(symbolText("The Earth's rotational speed;", undefined, "Symbol"), null);
  assert.equal(symbolText("z", "+mn-lt", "Symbol"), null);
});

test("a:sym Symbol → seuls les caractères PUA sont remappés", () => {
  assert.equal(symbolText("\uF057", "Times New Roman", "Symbol"), "Ω");
  assert.equal(symbolText("\uF066", "+mn-lt", "Symbol"), "ϕ"); // Symbol dessine « f » en ϕ droit
  assert.equal(symbolText("\uF06A", "+mn-lt", "Symbol"), "φ");
  assert.equal(symbolText("\uF072 = 1", undefined, "Symbol"), "ρ = 1");
  assert.equal(symbolText("\uF06D", undefined, "Symbol"), "μ");
  assert.equal(symbolText("a \uF0DE b", undefined, "Symbol"), "a ⇒ b");
  assert.equal(symbolText("cos φ \uF040 1", "+mn-lt", "Symbol"), "cos φ ≅ 1");
});

test("police latin Symbol → remappage complet (ASCII compris)", () => {
  assert.equal(symbolText("abc", "Symbol", undefined), "αβχ");
  assert.equal(symbolText("-", "Symbol", "Symbol"), "−");
  assert.equal(symbolText("W \uF057", "Symbol", undefined), "Ω Ω");
});

test("a:sym Wingdings : même règle (texte ordinaire intact, PUA → puce)", () => {
  assert.equal(symbolText("Note", "Arial", "Wingdings"), null);
  assert.equal(symbolText("\uF0FC OK", "Arial", "Wingdings"), "✓ OK");
});

test("aucune police symbole / texte vide → null", () => {
  assert.equal(symbolText("abc", "Arial", undefined), null);
  assert.equal(symbolText("", "Symbol", "Symbol"), null);
});

test("texte mêlé : seuls les glyphes symbole forment un segment à part (police du run conservée)", () => {
  assert.deepEqual(symbolParts("\uF057z = \uF057 . Cos \uF066", "Poppins", "Symbol"), [
    { text: "Ω", sym: true }, { text: "z = ", sym: false }, { text: "Ω", sym: true },
    { text: " . Cos ", sym: false }, { text: "ϕ", sym: true },
  ]);
  assert.deepEqual(symbolParts("abc", "Symbol", undefined), [{ text: "αβχ", sym: true }]);
  assert.equal(symbolParts("The Earth", undefined, "Symbol"), null);
});
