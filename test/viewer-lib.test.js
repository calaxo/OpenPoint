// Tests des modules partagés des viewers : renderer/lib/{deck,images,html-page,deck-data}.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, rmSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { zipSync, strToU8 } from 'fflate';
import { isEntrySafe, openDeck, deckName } from '../renderer/lib/deck.mjs';
import { createImageStore, IMAGE_TOKEN } from '../renderer/lib/images.mjs';
import { jsonForScript } from '../renderer/lib/html-page.mjs';
import { viewBoxSize } from '../renderer/lib/deck-data.mjs';
import { buildDecks, zipDeck } from './fixtures/decks.mjs';

test('isEntrySafe : refuse absolu, .., lecteur, octet nul', () => {
  for (const ok of ['ppt/slides/slide1.xml', '[Content_Types].xml', 'a/b..c/d.xml']) assert.equal(isEntrySafe(ok), true, ok);
  for (const bad of ['/etc/passwd', '\\x.xml', 'C:/x.xml', 'c:x', '../x', 'ppt/../../x', 'ppt\\..\\x', 'a\0b', ''])
    assert.equal(isEntrySafe(bad), false, JSON.stringify(bad));
});

test('deckName : nom assaini', () => {
  assert.equal(deckName('C:\\x\\Mon deck (v2).pptx'), 'Mon_deck__v2_');
});

test('openDeck : ordre de présentation, diapos masquées, zip-slip rejeté', () => {
  const root = mkdtempSync(join(tmpdir(), 'op-deck-'));
  try {
    const files = { ...buildDecks().text };
    files['ppt/slides/slide3.xml'] = String(files['ppt/slides/slide3.xml']).replace(/<p:sld\b/, '<p:sld show="0"');
    const pptx = join(root, 'text.pptx');
    writeFileSync(pptx, zipDeck(files));
    const deck = openDeck(pptx, { workRoot: root, metafiles: false });
    assert.deepEqual(deck.slides.slice(0, 3), [2, 1, 3]);
    assert.deepEqual([...deck.hidden], [3]);
    assert.equal(deck.name, 'text');
    assert.ok(existsSync(join(deck.pptDir, 'presentation.xml')));

    const evil = join(root, 'evil.pptx');
    writeFileSync(evil, zipSync({ 'ppt/presentation.xml': strToU8('<x/>'), '../evil.txt': strToU8('x') }));
    assert.throws(() => openDeck(evil, { workRoot: root, metafiles: false }), /zip-slip/);
    assert.equal(existsSync(join(root, 'evil.txt')), false);
  } finally { rmSync(root, { recursive: true, force: true }); }
});

test('createImageStore : une data URI identique est stockée une seule fois', () => {
  const a = 'data:image/png;base64,iVBORw0KGgo=', b = 'data:image/jpeg;base64,/9j/4AAQ';
  const store = createImageStore();
  const s1 = store.intern(`<image href="${a}"/><image href="${b}"/>`);
  const s2 = store.intern(`<image href="${a}"/>`);
  assert.equal(s1, `<image href="${IMAGE_TOKEN}0"/><image href="${IMAGE_TOKEN}1"/>`);
  assert.equal(s2, `<image href="${IMAGE_TOKEN}0"/>`);
  assert.deepEqual(store.list, [a, b]);
});

test('jsonForScript : aucun « < » brut, données intactes', () => {
  const obj = { svg: '</script><!-- x -->' };
  const out = jsonForScript(obj);
  assert.equal(out.includes('<'), false);
  assert.deepEqual(JSON.parse(out), obj);
});

test('viewBoxSize : taille en points lue sur le viewBox', () => {
  assert.deepEqual(viewBoxSize('<svg viewBox="0 0 720 405" width="100%">'), { w: 720, h: 405 });
  assert.deepEqual(viewBoxSize('<svg>'), { w: 960, h: 540 });
});
