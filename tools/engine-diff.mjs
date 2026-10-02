// ═══ [QUALITÉ · diff de moteur] ═══ Quelles diapos ma modification change-t-elle ?
// Compare le moteur PPTX du dossier de travail à celui d'un commit git (HEAD par défaut), diapo par
// diapo : SVG, timeline d'animation et transition. Une refonte doit donner « 0 différence » ; une
// correction doit changer EXACTEMENT les diapos visées (le « blast radius »).
//
// Usage :
//   node tools/engine-diff.mjs                       -> decks d'exemple (test/fixtures/out) + input/*.pptx
//   node tools/engine-diff.mjs deck.pptx autre.pptx  -> decks choisis
//   node tools/engine-diff.mjs --ref main            -> compare à un autre commit / une branche
//   node tools/engine-diff.mjs --html diff.html      -> planche avant / après des diapos modifiées
//                                                       (à ouvrir dans un navigateur, ou tools/screenshot.mjs)
// Code de sortie : 0 si aucune différence, 1 sinon.
//
// Le moteur de référence est extrait avec `git archive` dans renderer/_work/ref-<sha>/ (gardé en cache).
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync, mkdirSync, existsSync, readdirSync, statSync } from 'node:fs';
import { join, dirname, basename, resolve } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const WORK = join(ROOT, 'renderer', '_work');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); return i >= 0 ? args.splice(i, 2)[1] : null; };
const ref = opt('--ref') || 'HEAD';
const htmlOut = opt('--html');

// --- moteur de référence (commit `ref`) ---
const sha = execFileSync('git', ['rev-parse', '--short', ref], { cwd: ROOT, encoding: 'utf8' }).trim();
const refDir = join(WORK, 'ref-' + sha);
if (!existsSync(join(refDir, 'renderer', 'pptx-render.mjs'))) {
  const zip = unzipSync(new Uint8Array(execFileSync('git', ['archive', '--format=zip', sha, 'renderer'], { cwd: ROOT, maxBuffer: 1 << 30 })));
  for (const [p, data] of Object.entries(zip)) {
    if (p.endsWith('/')) continue;
    mkdirSync(dirname(join(refDir, p)), { recursive: true });
    writeFileSync(join(refDir, p), data);
  }
}
const OLD = await import(pathToFileURL(join(refDir, 'renderer', 'pptx-render.mjs')).href);
const NEW = await import(pathToFileURL(join(ROOT, 'renderer', 'pptx-render.mjs')).href);

// --- decks à comparer ---
const listPptx = (d) => existsSync(d) ? readdirSync(d).filter((f) => /\.pptx$/i.test(f) && !f.startsWith('~$')).map((f) => join(d, f)) : [];
let decks = args.map((a) => resolve(a));
if (!decks.length) {
  if (!listPptx(join(ROOT, 'test', 'fixtures', 'out')).length) execFileSync('node', [join(ROOT, 'tools', 'make-fixtures.mjs')], { stdio: 'ignore' });
  decks = [...listPptx(join(ROOT, 'test', 'fixtures', 'out')), ...listPptx(join(ROOT, 'input'))];
}

// décompresse un .pptx dans renderer/_work/_diff/<nom>/ (refait si le .pptx a changé depuis)
function unzipDeck(file) {
  const dir = join(WORK, '_diff', basename(file).replace(/[^\w.-]/g, '_'));
  const stamp = String(statSync(file).mtimeMs);
  if (!existsSync(join(dir, '.stamp')) || readFileSync(join(dir, '.stamp'), 'utf8') !== stamp) {
    for (const [p, data] of Object.entries(unzipSync(new Uint8Array(readFileSync(file))))) {
      if (p.endsWith('/') || p.includes('..')) continue;
      mkdirSync(dirname(join(dir, p)), { recursive: true });
      writeFileSync(join(dir, p), data);
    }
    writeFileSync(join(dir, '.stamp'), stamp);
  }
  return join(dir, 'ppt');
}

const run = (M, ppt, n) => {
  try { return JSON.stringify([M.renderSlide(ppt, n, { embedImages: false }), M.extractTiming(ppt, n), M.extractTransition(ppt, n)]); }
  catch (e) { return 'ERREUR ' + e.message; }
};
let total = 0; const changed = [];
for (const deck of decks) {
  const ppt = unzipDeck(deck);
  const nums = readdirSync(join(ppt, 'slides')).filter((f) => /^slide\d+\.xml$/.test(f)).map((f) => +f.match(/\d+/)[0]).sort((a, b) => a - b);
  for (const n of nums) {
    total++;
    const a = run(OLD, ppt, n), b = run(NEW, ppt, n);
    if (a === b) continue;
    let i = 0; while (i < a.length && a[i] === b[i]) i++;
    changed.push({ deck, ppt, n });
    console.log(`≠ ${basename(deck)} diapo ${n} (1er écart au caractère ${i})`);
    console.log(`    avant : …${a.slice(Math.max(0, i - 60), i + 90)}`);
    console.log(`    après : …${b.slice(Math.max(0, i - 60), i + 90)}`);
  }
}
console.log(`\n${total} diapos comparées à ${ref} (${sha}) : ${changed.length} modifiée(s)`);

if (htmlOut && changed.length) {
  const { fontsCssFor } = await import(pathToFileURL(join(ROOT, 'renderer', 'web-fonts.mjs')).href);
  let body = '';
  for (const { deck, ppt, n } of changed.slice(0, 60)) {
    // préfixes d'ID distincts : les deux SVG d'une même diapo cohabitent dans la page
    const a = OLD.renderSlide(ppt, n, { embedImages: true }).svg.replace(/\bs(\d+)_/g, 'a$1_');
    const b = NEW.renderSlide(ppt, n, { embedImages: true }).svg.replace(/\bs(\d+)_/g, 'b$1_');
    body += `<h3>${basename(deck)} — diapo ${n}</h3><div class="row"><div><i>avant (${sha})</i>${a}</div><div><i>après</i>${b}</div></div>`;
  }
  writeFileSync(htmlOut, `<!doctype html><html lang="fr"><meta charset="utf-8"><title>engine-diff</title><style>${fontsCssFor(body)}`
    + `body{margin:0;padding:8px;font:13px system-ui;background:#eee}h3{margin:16px 0 4px}.row{display:flex;gap:8px}.row>div{flex:1;background:#fff}</style>${body}</html>`);
  console.log(`planche avant/après : ${htmlOut}`);
}
process.exit(changed.length ? 1 : 0);
