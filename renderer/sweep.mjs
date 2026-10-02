// Balayage automatique : compare CHAQUE diapo (mon rendu) à la réf PowerPoint (oracle) et sort un
// rapport trié par divergence, en localisant les écarts STRUCTURELS (grille), hors bruit de police.
// Usage : node renderer/sweep.mjs <deck.pptx> <refsDir> [outReport.json]
//   refsDir doit contenir slide-<N>.png (export COM PowerPoint, 1280x720).
import { readFileSync, writeFileSync, existsSync, mkdirSync, readdirSync } from 'node:fs';
import { join, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import puppeteer from 'puppeteer';
import { renderSlide } from './pptx-render.mjs';
import { convertMetafiles } from './metafiles.mjs';
import { fontsCssFor } from './web-fonts.mjs'; // polices du LMS (woff2 + métriques PowerPoint)

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const [, , pptxPath, refsDir, outPath] = process.argv;
if (!pptxPath || !refsDir) { console.error('usage: sweep.mjs <deck.pptx> <refsDir> [out.json]'); process.exit(1); }

// unzip -> _work
const name = basename(pptxPath).replace(/\.pptx$/i, '').replace(/[^\w.-]/g, '_');
const work = join(ROOT, 'renderer', '_work', name);
if (!existsSync(join(work, 'ppt', 'presentation.xml'))) {
  const zip = unzipSync(new Uint8Array(readFileSync(pptxPath)));
  for (const [p, d] of Object.entries(zip)) { if (p.endsWith('/')) continue; const o = join(work, p); mkdirSync(dirname(o), { recursive: true }); writeFileSync(o, Buffer.from(d)); }
  convertMetafiles(join(work, 'ppt', 'media'));
}
const PPT = join(work, 'ppt');

// Polices : celles servies en production (voir web-fonts.mjs), calculées par diapo plus bas.
// PAS de rétrécissement de texte côté navigateur (comme en production) : PowerPoint ne recalcule pas
// l'autofit à l'affichage, il applique le fontScale ENREGISTRÉ dans le fichier — que le moteur a déjà
// reporté dans le SVG. Rétrécir ici ferait diverger le rendu de l'original PowerPoint.

// slides ayant une réf
const refNums = readdirSync(refsDir).map(f => (f.match(/slide-(\d+)\.png$/) || [])[1]).filter(Boolean).map(Number).sort((a, b) => a - b);
const W = 1280, H = 720;

function diffOf(minePng, refPng) {
  const a = PNG.sync.read(minePng), b = PNG.sync.read(refPng);
  const w = Math.min(a.width, b.width, W), h = Math.min(a.height, b.height, H);
  const crop = (s) => { if (s.width === w && s.height === h) return s.data; const o = Buffer.alloc(w*h*4); for (let y=0;y<h;y++) for (let x=0;x<w;x++){const si=(s.width*y+x)<<2,di=(w*y+x)<<2;o[di]=s.data[si];o[di+1]=s.data[si+1];o[di+2]=s.data[si+2];o[di+3]=s.data[si+3];} return o; };
  const da = crop(a), db = crop(b), dd = Buffer.alloc(w*h*4);
  const n = pixelmatch(da, db, dd, w, h, { threshold: 0.12, includeAA: true });
  diffOf._last = { w, h, dd, da, db };
  // grille 16x9 -> zones structurelles
  const GX=16, GY=9, cellW=w/GX, cellH=h/GY, cells=new Array(GX*GY).fill(0);
  for (let y=0;y<h;y++) for (let x=0;x<w;x++){ const i=(w*y+x)<<2; if (dd[i+3]>0 && dd[i]>200 && dd[i+1]<200){ cells[((y/cellH)|0)*GX + ((x/cellW)|0)]++; } }
  const regions = cells.map((c,idx)=>({gx:idx%GX,gy:(idx/GX)|0,pct:+(100*c/(cellW*cellH)).toFixed(0),x:Math.round((idx%GX)*cellW),y:Math.round(((idx/GX)|0)*cellH)})).filter(r=>r.pct>25).sort((p,q)=>q.pct-p.pct);
  return { diffPct: +(100*n/(w*h)).toFixed(2), regions };
}

const browser = await puppeteer.launch({ headless: 'new' });
const page = await browser.newPage();
await page.setViewport({ width: W, height: H, deviceScaleFactor: 1 });
const results = [];
for (const n of refNums) {
  let svg;
  try { ({ svg } = renderSlide(PPT, n, { embedImages: true })); } catch (e) { results.push({ n, error: e.message }); continue; }
  svg = svg.replace('width="100%"', `width="${W}" height="${H}"`);
  const html = `<!doctype html><meta charset=utf8><style>${fontsCssFor(svg)}*{margin:0}html,body{background:#0c0d10}</style><body>${svg}</body>`;
  await page.setContent(html, { waitUntil: 'load' });
  try { await page.evaluate(() => globalThis.document.fonts.ready); } catch {}
  const buf = await page.screenshot({ clip: { x: 0, y: 0, width: W, height: H } });
  const ref = readFileSync(join(refsDir, `slide-${n}.png`));
  try {
    const d = diffOf(buf, ref); results.push({ n, ...d });
    if (d.diffPct >= 10 && outPath) { // dump mine + diff pour revue visuelle des écarts notables
      const dir = join(dirname(outPath), name + '_img'); mkdirSync(dir, { recursive: true });
      writeFileSync(join(dir, `slide-${n}-mine.png`), buf);
      const { w, h, dd } = diffOf._last; const p = new PNG({ width: w, height: h }); dd.copy(p.data); writeFileSync(join(dir, `slide-${n}-diff.png`), PNG.sync.write(p));
    }
  } catch (e) { results.push({ n, error: 'diff:' + e.message }); }
  process.stderr.write(`\r  diapo ${n}/${refNums.length}   `);
}
await browser.close();

results.sort((a, b) => (b.diffPct || 0) - (a.diffPct || 0));
const report = { deck: name, slides: refNums.length, results };
if (outPath) writeFileSync(outPath, JSON.stringify(report, null, 1));
console.log('\n=== ' + name + ' — diapos les plus divergentes (structurel probable si régions localisées) ===');
for (const r of results.slice(0, 15)) {
  if (r.error) { console.log(`  d.${r.n}: ERREUR ${r.error}`); continue; }
  const reg = (r.regions || []).slice(0, 4).map(z => `(${z.x},${z.y}):${z.pct}%`).join(' ');
  console.log(`  d.${r.n}: ${r.diffPct}% ${reg ? '| zones ' + reg : '| (diffus = bruit police)'}`);
}
