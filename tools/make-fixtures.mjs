// Écrit chaque deck d'exemple synthétique dans test/fixtures/out/<nom>.pptx (rien d'autre).
// Usage : node tools/make-fixtures.mjs
import { mkdirSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDecks, zipDeck } from '../test/fixtures/decks.mjs';

const OUT = join(dirname(fileURLToPath(import.meta.url)), '..', 'test', 'fixtures', 'out');
mkdirSync(OUT, { recursive: true });
for (const [name, files] of Object.entries(buildDecks())) {
  const p = join(OUT, `${name}.pptx`);
  writeFileSync(p, zipDeck(files));
  console.log(`OK ${p}`);
}
