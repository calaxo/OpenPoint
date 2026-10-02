// Test de bout en bout des DEUX viewers dans un vrai navigateur (puppeteer). Hors de `npm test`.
// Usage : npm run test:e2e        (ou node tools/e2e-viewers.mjs)
//
// 1. écrit les decks synthétiques (test/fixtures/decks.mjs) + une variante « text-hidden » (diapo
//    masquée) dans test/fixtures/out/ ;
// 2. construit viewer.html (build.mjs) et player-out.html (build-player.mjs) dans test/fixtures/out/e2e/ ;
// 3. vérifie dans Chrome : aucune erreur de page, nombre et ordre des diapos, étapes d'animation,
//    démarrage automatique, retour arrière à l'état final, navigation par #ancre, diapos masquées,
//    images dé-dupliquées affichées.
// Code de sortie 1 au premier échec (message « ÉCHEC : … »). Chrome : env CHROME_PATH, sinon Chrome
// installé (Windows), sinon celui de puppeteer.
// les fonctions passées à page.evaluate / $eval s'exécutent DANS le navigateur :
/* global document, window, location, getComputedStyle, Image */
import { launchBrowser } from './browser.mjs';
import { mkdirSync, writeFileSync, readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { execFileSync } from 'node:child_process';
import { buildDecks, zipDeck } from '../test/fixtures/decks.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const FIX = join(ROOT, 'test', 'fixtures', 'out');
const OUT = join(FIX, 'e2e');
mkdirSync(OUT, { recursive: true });

// ─── 1. decks ───
const decks = buildDecks();
for (const [name, files] of Object.entries(decks)) writeFileSync(join(FIX, `${name}.pptx`), zipDeck(files));
// variante : deck « text » dont le fichier slide1.xml (2e dans l'ordre de présentation) est masqué
const hiddenFiles = { ...decks.text };
hiddenFiles['ppt/slides/slide1.xml'] = String(hiddenFiles['ppt/slides/slide1.xml']).replace(/<p:sld\b/, '<p:sld show="0"');
writeFileSync(join(FIX, 'text-hidden.pptx'), zipDeck(hiddenFiles));

// ─── 2. builds ───
function build(script, deck, out) {
  execFileSync(process.execPath, [join(ROOT, 'renderer', script), join(FIX, `${deck}.pptx`), `--out=${join(OUT, out)}`], { cwd: ROOT, stdio: 'pipe' });
  return pathToFileURL(join(OUT, out)).href;
}
const url = {
  viewerText: build('build.mjs', 'text', 'viewer-text.html'),
  viewerHidden: build('build.mjs', 'text-hidden', 'viewer-text-hidden.html'),
  viewerTables: build('build.mjs', 'tables-charts', 'viewer-tables.html'),
  playerAnim: build('build-player.mjs', 'animations', 'player-animations.html'),
  playerHidden: build('build-player.mjs', 'text-hidden', 'player-text-hidden.html'),
};
const tablesHtml = readFileSync(join(OUT, 'viewer-tables.html'), 'utf8');

// ─── 3. navigateur ───
let failures = 0, passes = 0;
function check(cond, msg) {
  if (cond) { passes++; console.log('  ok  ', msg); } else { failures++; console.error('  ÉCHEC :', msg); }
}

const browser = await launchBrowser();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ouvre une page, compte ses erreurs (exception, console.error, requête échouée)
async function open(href, { width = 1280, height = 720, reduced = false } = {}) {
  const page = await browser.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') page.errors.push(m.text()); });
  page.on('requestfailed', (r) => page.errors.push('requête échouée ' + r.url().slice(0, 80)));
  await page.setViewport({ width, height });
  if (reduced) await page.emulateMediaFeatures([{ name: 'prefers-reduced-motion', value: 'reduce' }]);
  await page.goto(href, { waitUntil: 'load' });
  await sleep(150);
  return page;
}
// attend la fin de toutes les animations Web (celles en « fill: forwards » restent listées, mais finies)
async function settle(page) {
  await page.waitForFunction(() => document.getAnimations().every((a) => a.playState === 'finished' || a.playState === 'idle'), { timeout: 10000 });
}
// opacité effective (produit des ancêtres) de la forme data-spid dans le calque courant du player
const opacityOf = (page, spid) => page.evaluate((id) => {
  const layers = document.querySelectorAll('#slide .layer');
  let el = layers[layers.length - 1].querySelector('[data-spid="' + id + '"]');
  let o = 1;
  while (el && el.id !== 'slide') { o *= +getComputedStyle(el).opacity; el = el.parentElement; }
  return o;
}, spid);
const body = (page) => page.evaluate(() => ({ ...document.body.dataset }));

try {
  // ── viewer défilement : deck « text » (ordre de présentation slide2, slide1, slide3…) ──
  console.log('viewer défilement — text');
  {
    const page = await open(url.viewerText);
    const cards = await page.$$eval('.card', (cs) => cs.map((c) => ({ pos: c.dataset.pos, n: c.dataset.n, num: c.querySelector('.num').textContent })));
    const nSlides = Object.keys(decks.text).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p)).length;
    check(cards.length === nSlides, `${nSlides} cartes (${cards.length})`);
    check(cards.map((c) => c.num).join(',') === cards.map((_, i) => i + 1).join(','), 'étiquettes 1..n dans l\'ordre');
    check(cards.slice(0, 3).map((c) => c.n).join(',') === '2,1,3', `ordre de présentation slide2, slide1, slide3 (${cards.map((c) => c.n).join(',')})`);
    // insertion paresseuse : toutes les diapos finissent par être insérées en défilant jusqu'en bas
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await sleep(300);
    const filled = await page.$$eval('.card[data-filled] svg', (s) => s.length);
    check(filled === nSlides, `toutes les diapos insérées après défilement (${filled})`);
    // clavier j / k et ancre
    await page.evaluate(() => window.scrollTo(0, 0)); await sleep(100);
    await page.keyboard.press('j'); await sleep(150);
    check(await page.evaluate(() => location.hash) === '#2', 'j -> #2');
    await page.keyboard.press('k'); await sleep(150);
    check(await page.evaluate(() => location.hash) === '#1', 'k -> #1');
    check(page.errors.length === 0, 'aucune erreur de page ' + page.errors.join(' | '));
    await page.close();
  }
  {
    const page = await open(url.viewerText + '#3', { width: 390, height: 844 });
    const top = await page.$eval('#slide-3', (c) => c.getBoundingClientRect().top);
    const bar = await page.$eval('#bar', (b) => b.offsetHeight);
    check(top >= bar - 1 && top < bar + 40, `#3 : carte 3 juste sous la barre (top ${Math.round(top)})`);
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    check(overflow <= 0, `téléphone : pas de défilement horizontal (${overflow})`);
    await page.close();
  }
  console.log('viewer défilement — diapo masquée, images');
  {
    const page = await open(url.viewerHidden);
    const hidden = await page.$$eval('.card.is-hidden', (cs) => cs.map((c) => c.dataset.n + ':' + c.querySelector('.badge').textContent));
    check(hidden.join() === '1:masquée', `slide1.xml affichée grisée avec badge « masquée » (${hidden})`);
    await page.close();
  }
  {
    const page = await open(url.viewerTables);
    check(!/<image[^>]*href="data:/.test(tablesHtml) && /op-img:0/.test(tablesHtml), 'images remplacées par des jetons op-img dans la page');
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight)); await sleep(300);
    const imgs = await page.$$eval('.card image', (is) => is.map((i) => i.getAttribute('href') || i.getAttribute('xlink:href')));
    check(imgs.length > 0 && imgs.every((h) => /^blob:/.test(h)), `images résolues en URL blob: (${imgs.length})`);
    const ok = await page.evaluate(async () => {
      const hrefs = [...document.querySelectorAll('.card image')].map((i) => i.getAttribute('href'));
      const res = await Promise.all(hrefs.map((h) => new Promise((r) => { const im = new Image(); im.onload = () => r(im.naturalWidth > 0); im.onerror = () => r(false); im.src = h; })));
      return res.every(Boolean);
    });
    check(ok, 'chaque image se décode');
    check(page.errors.length === 0, 'aucune erreur de page ' + page.errors.join(' | '));
    await page.close();
  }

  // ── player : deck « animations » ──
  console.log('player — animations');
  {
    const page = await open(url.playerAnim);
    let b = await body(page);
    check(b.pos === '1' && b.steps === '3', `diapo 1, 3 clics (${b.pos}, ${b.steps})`);
    const ops0 = [await opacityOf(page, '3'), await opacityOf(page, '4'), await opacityOf(page, '5')];
    check(ops0.every((o) => o === 0), `3 formes cachées au départ (${ops0})`);
    for (let i = 0; i < 3; i++) { await page.mouse.click(640, 360); await sleep(50); }
    await settle(page);
    const ops3 = [await opacityOf(page, '3'), await opacityOf(page, '4'), await opacityOf(page, '5')];
    check(ops3.every((o) => o === 1), `3 formes visibles après 3 clics (${ops3})`);
    b = await body(page);
    check(b.pos === '1' && b.step === '3', 'toujours diapo 1 après 3 clics');
    // clic suivant -> diapo 2 ; retour -> diapo 1 dans son état FINAL
    await page.keyboard.press('ArrowRight'); await settle(page);
    check((await body(page)).pos === '2', '→ : diapo 2');
    await page.keyboard.press('ArrowLeft'); await settle(page);
    b = await body(page);
    const opsBack = [await opacityOf(page, '3'), await opacityOf(page, '4'), await opacityOf(page, '5')];
    check(b.pos === '1' && b.step === '3' && opsBack.every((o) => o === 1), `← : diapo 1 à l'état final (${b.pos}, clic ${b.step}, ${opsBack})`);
    // ← encore : étape précédente sur la même diapo
    await page.keyboard.press('ArrowLeft'); await settle(page);
    check((await body(page)).step === '2' && await opacityOf(page, '5') === 0, '← : étape précédente (forme 5 recachée)');
    // ancre : #9 -> démarrage automatique sans clic
    await page.evaluate(() => { location.hash = '#9'; });
    await sleep(100); await settle(page);
    b = await body(page);
    check(b.pos === '9', `#9 : diapo 9 (${b.pos})`);
    check(await opacityOf(page, '3') === 1 && await opacityOf(page, '4') === 0, 'diapo 9 : étape auto jouée seule, forme du clic encore cachée');
    check(b.step === '1' && b.steps === '2', `diapo 9 : 1/2 étapes (${b.step}/${b.steps})`);
    // la diapo 3 : sortie -> forme 3 cachée à la fin ; End / Home
    await page.keyboard.press('Home'); await settle(page);
    check((await body(page)).pos === '1', 'Début : diapo 1');
    await page.keyboard.press('End'); await settle(page);
    const lastPos = String(buildDecks().animations['ppt/presentation.xml'].match(/<p:sldId /g).length);
    check((await body(page)).pos === lastPos, `Fin : dernière diapo (${lastPos})`);
    // numéro + Entrée
    await page.keyboard.press('3'); await page.keyboard.press('Enter'); await settle(page);
    check((await body(page)).pos === '3', '3 + Entrée : diapo 3');
    await page.keyboard.press('ArrowRight'); await page.keyboard.press('ArrowRight'); await settle(page);
    check(await opacityOf(page, '3') === 0 && await opacityOf(page, '4') === 1, 'diapo 3 : forme sortie cachée, forme emphase visible');
    check(await page.evaluate(() => location.hash) === '#3', 'URL tenue à jour (#3)');
    // clics très rapides : pas d'erreur, état cohérent
    await page.keyboard.press('Home');
    for (let i = 0; i < 12; i++) await page.mouse.click(640, 360, { delay: 0 });
    await settle(page);
    b = await body(page);
    check(+b.pos >= 2 && page.errors.length === 0, `12 clics rapides : diapo ${b.pos}, aucune erreur`);
    // retour en arrière rapide jusqu'au début : état initial complet
    for (let i = 0; i < 30; i++) await page.keyboard.press('ArrowLeft');
    await settle(page);
    b = await body(page);
    check(b.pos === '1' && b.step === '0' && await opacityOf(page, '3') === 0, 'retour rapide jusqu\'à la diapo 1, étape 0');
    check(page.errors.length === 0, 'aucune erreur de page ' + page.errors.join(' | '));
    await page.close();
  }
  {
    const page = await open(url.playerAnim + '#3', { reduced: true });
    await page.keyboard.press('ArrowRight');
    const n = await page.evaluate(() => document.getAnimations().length);
    check(n === 0 && await opacityOf(page, '3') === 0, `prefers-reduced-motion : aucune animation, état immédiat (${n})`);
    await page.close();
  }
  {
    const page = await open(url.playerAnim, { width: 390, height: 844 });
    const r = await page.$eval('#slide', (s) => { const b = s.getBoundingClientRect(); return { w: b.width, h: b.height }; });
    check(Math.abs(r.w / r.h - 16 / 9) < 0.02 && r.w <= 390, `téléphone : diapo letterboxée au ratio 16:9 (${Math.round(r.w)}x${Math.round(r.h)})`);
    await page.close();
  }
  console.log('player — diapo masquée');
  {
    const page = await open(url.playerHidden);
    const seen = [];
    for (let i = 0; i < 6; i++) { seen.push((await body(page)).pos); await page.keyboard.press('ArrowRight'); await settle(page); }
    check(!seen.includes('2') && seen[0] === '1' && seen[1] === '3', `diapo masquée (position 2) sautée (${[...new Set(seen)]})`);
    check(page.errors.length === 0, 'aucune erreur de page ' + page.errors.join(' | '));
    await page.close();
  }
} finally {
  await browser.close();
}
console.log(`\n${passes} vérification(s) OK, ${failures} échec(s)`);
process.exit(failures ? 1 : 0);
