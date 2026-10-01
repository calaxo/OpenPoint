// Balaye TOUT le corpus : pour chaque .pptx de specimen/, lance sweep.mjs avec le dossier de réfs
// correspondant (<refsBase>/<nameSanitizé>), agrège en un rapport global trié par divergence.
// Usage : node renderer/sweep-all.mjs <refsBase> <outDir>
import { readdirSync, existsSync, statSync, writeFileSync, mkdirSync, readFileSync } from 'node:fs';
import { join, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const [, , refsBase, outDir] = process.argv;
if (!refsBase || !outDir) { console.error('usage: sweep-all.mjs <refsBase> <outDir>'); process.exit(1); }
mkdirSync(outDir, { recursive: true });

function findPptx(dir) { const out = []; for (const e of readdirSync(dir)) { const p = join(dir, e); if (statSync(p).isDirectory()) out.push(...findPptx(p)); else if (/\.pptx$/i.test(e) && !/^~\$/.test(e)) out.push(p); } return out; }
const decks = findPptx(join(ROOT, 'specimen'));

const global = [];
for (const deck of decks) {
  const name = basename(deck).replace(/\.pptx$/i, '').replace(/[^\w.-]/g, '_');
  const refs = join(refsBase, name);
  if (!existsSync(refs)) { console.log(`-- ${name} : pas de réfs, ignoré`); continue; }
  const out = join(outDir, name + '.json');
  process.stdout.write(`\n### ${name}\n`);
  try {
    execFileSync(process.execPath, [join(ROOT, 'renderer', 'sweep.mjs'), deck, refs, out], { stdio: 'inherit', cwd: ROOT });
    const rep = JSON.parse(readFileSync(out, 'utf8'));
    for (const r of rep.results) if (!r.error && r.diffPct >= 8) global.push({ deck: name, ...r });
  } catch (e) { console.log(`  ERREUR sweep ${name}: ${e.message}`); }
}

global.sort((a, b) => b.diffPct - a.diffPct);
writeFileSync(join(outDir, '_GLOBAL.json'), JSON.stringify(global, null, 1));
console.log('\n\n========== TOP 40 ÉCARTS DU CORPUS (structurel si régions localisées) ==========');
for (const r of global.slice(0, 40)) {
  const reg = (r.regions || []).slice(0, 3).map(z => `(${z.x},${z.y}):${z.pct}%`).join(' ');
  console.log(`  ${r.deck} d.${r.n}: ${r.diffPct}% ${reg ? '| ' + reg : '| diffus'}`);
}
