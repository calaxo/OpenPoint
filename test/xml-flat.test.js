// flatOf() (renderer/pptx/xml.mjs) convertit un nœud ordonné en arbre plat SANS passer par du texte XML.
// Il doit donner EXACTEMENT ce que donnait l'ancien aller-retour serialize() + PF.parse() — ordre des clés
// compris. Ce test le vérifie sur chaque nœud de chaque part XML des decks d'exemple : il casse si une
// mise à jour de fast-xml-parser change ses règles d'aplatissement.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildDecks } from './fixtures/decks.mjs';
import { P, PF, serialize, flatOf, tagOf, kids } from '../renderer/pptx/xml.mjs';

const roundTrip = (node) => PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(node)}</root>`).root[tagOf(node)];

test('flatOf ≡ serialize + PF.parse sur toutes les parts XML des decks d\'exemple', () => {
  let count = 0;
  for (const [deck, files] of Object.entries(buildDecks())) {
    for (const [path, content] of Object.entries(files)) {
      if (!/\.(xml|rels)$/.test(path)) continue;
      const xml = typeof content === 'string' ? content : new TextDecoder().decode(content);
      const visit = (node) => {
        const t = tagOf(node);
        if (t === '#text' || t.startsWith('?')) return;
        assert.deepEqual(JSON.stringify(flatOf(node)[t]), JSON.stringify(roundTrip(node)), `${deck}/${path} <${t}>`);
        count++;
        for (const c of kids(node)) visit(c);
      };
      for (const root of P.parse(xml)) visit(root);
    }
  }
  assert.ok(count > 1000, `${count} nœuds vérifiés`);
});

test('flatOf : cas limites (texte seul, vide, attributs + texte, balises répétées)', () => {
  const [n] = P.parse('<a:p><a:t>x</a:t><a:br/><a:r b="1"><a:t> y </a:t></a:r><a:r><a:t/></a:r><a:e k="v">t</a:e></a:p>');
  assert.deepEqual(flatOf(n)['a:p'], {
    'a:t': 'x', 'a:br': '', 'a:r': [{ 'a:t': ' y ', '@b': '1' }, { 'a:t': '' }], 'a:e': { '#text': 't', '@k': 'v' },
  });
  assert.equal(flatOf(n), flatOf(n)); // mis en cache : même objet (donc LECTURE SEULE)
});
