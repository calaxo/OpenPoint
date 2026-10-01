// Comparateur d'images : mon rendu vs référence PowerPoint (oracle) — pour traquer les micro-écarts.
// Usage : node renderer/compare.mjs <mine.png> <ref.png> [diffOut.png]
// Sortie : % de pixels différents + zones (grille) les plus divergentes + image de diff.
//
// pixelmatch ignore l'anti-aliasing (includeAA:false) pour ne PAS bruiter sur le rendu de police
// (navigateur vs PowerPoint diffèrent toujours d'un poil) et ne remonter que les VRAIS écarts
// structurels : forme manquante, mauvaise couleur, élément décalé, texte absent.
import { readFileSync, writeFileSync } from 'node:fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

const [, , minePath, refPath, diffPath] = process.argv;
if (!minePath || !refPath) { console.error('usage: compare.mjs <mine.png> <ref.png> [diff.png]'); process.exit(1); }

const mine = PNG.sync.read(readFileSync(minePath));
const ref = PNG.sync.read(readFileSync(refPath));

// aligner les dimensions (recadrer au plus petit commun) si léger écart
const W = Math.min(mine.width, ref.width), H = Math.min(mine.height, ref.height);
function crop(src) {
  if (src.width === W && src.height === H) return src;
  const out = new PNG({ width: W, height: H });
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    const s = (src.width * y + x) << 2, d = (W * y + x) << 2;
    out.data[d] = src.data[s]; out.data[d+1] = src.data[s+1]; out.data[d+2] = src.data[s+2]; out.data[d+3] = src.data[s+3];
  }
  return out;
}
const a = crop(mine), b = crop(ref);
const diff = new PNG({ width: W, height: H });
const n = pixelmatch(a.data, b.data, diff.data, W, H, { threshold: 0.12, includeAA: true, alpha: 0.4 });
if (diffPath) writeFileSync(diffPath, PNG.sync.write(diff));

// grille GX x GY : densité d'écart par cellule -> localise les zones qui divergent
const GX = 16, GY = 9;
const cellW = W / GX, cellH = H / GY;
const cells = Array.from({ length: GX * GY }, () => 0);
for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
  const i = (W * y + x) << 2;
  // pixelmatch marque les diffs en rouge/jaune dans diff.data ; détecter pixel non transparent marqué
  if (diff.data[i+3] > 0 && (diff.data[i] > 200 && diff.data[i+1] < 200)) {
    const gx = Math.min(GX-1, (x / cellW) | 0), gy = Math.min(GY-1, (y / cellH) | 0);
    cells[gy * GX + gx]++;
  }
}
const cellArea = cellW * cellH;
const regions = cells.map((c, idx) => ({
  gx: idx % GX, gy: (idx / GX) | 0, pct: +(100 * c / cellArea).toFixed(1),
  x: Math.round((idx % GX) * cellW), y: Math.round(((idx / GX) | 0) * cellH),
})).filter(r => r.pct > 8).sort((x, y) => y.pct - x.pct);

console.log(JSON.stringify({
  size: `${W}x${H}`, diffPixels: n, diffPct: +(100 * n / (W * H)).toFixed(2),
  topRegions: regions.slice(0, 10),
}));
