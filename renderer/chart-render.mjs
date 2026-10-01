// ═══ [GRAPHIQUES] ═══ docs/ARCHITECTURE.md · rend un c:chart (DrawingML Chart, ppt/charts/chartN.xml)
// en groupe SVG. Consommé par pptx-render.mjs quand un p:graphicFrame contient un graphique (et non
// un tableau). Couvre les familles courantes : colonnes/barres (groupées + empilées), lignes, aires,
// camembert/anneau. Les valeurs/catégories viennent des CACHES (c:numCache/c:strCache) écrits par
// PowerPoint dans le fichier -> pas besoin du classeur Excel source.
//
// renderChart(chartFlat, w, h, { palette, color }) -> chaîne SVG (repère 0..w × 0..h, points).
//   chartFlat : c:chartSpace parsé à PLAT (attributs préfixés '@').
//   palette   : couleurs d'accent du thème [hex,...] (série i -> palette[i % n]).
//   color(fillOrSpPr) : résout un <a:solidFill>/<c:spPr> -> couleur CSS, ou null si absent.

const esc = (s) => String(s ?? '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
const arr = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];
const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; };

// extrait les points d'un cache (c:strCache/c:numCache) en tableau dense indexé par c:pt/@idx
function ptsOf(ref) {
  const cache = ref?.['c:numCache'] || ref?.['c:strCache'] || ref?.['c:numLit'] || ref?.['c:strLit'];
  if (!cache) return [];
  const cnt = +cache['c:ptCount']?.['@val'] || 0;
  const out = new Array(cnt).fill(null);
  for (const p of arr(cache['c:pt'])) { const i = +p['@idx'] || 0; out[i] = p['c:v']; }
  return out;
}
// une série -> { name, cats:[..], vals:[num..], fill:hex|null, pts:[{fill}] }
function serOf(ser, color) {
  const name = ptsOf(ser['c:tx']?.['c:strRef'])[0] ?? ser['c:tx']?.['c:v'] ?? '';
  const cats = ptsOf(ser['c:cat']?.['c:strRef'] || ser['c:cat']?.['c:numRef']).map((v) => v ?? '');
  const vals = ptsOf(ser['c:val']?.['c:numRef']).map(num);
  const fill = ser['c:spPr'] ? color(ser['c:spPr']) : null;
  const pts = arr(ser['c:dPt']).map((d) => ({ idx: +d['c:idx']?.['@val'] || 0, fill: d['c:spPr'] ? color(d['c:spPr']) : null }));
  return { name, cats, vals, fill, pts };
}

export function renderChart(chartFlat, w, h, opts = {}) {
  const palette = opts.palette?.length ? opts.palette : ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  const color = opts.color || (() => null);
  const cs = chartFlat?.['c:chartSpace'] || chartFlat; // tolère racine chartSpace ou son contenu
  const plot = cs?.['c:chart']?.['c:plotArea'];
  if (!plot) return '';
  // couleur d'un élément : point explicite (c:dPt) -> série (c:spPr) -> palette de thème (hex sans #)
  const paint = (i, ser, ptIdx) => {
    const dp = ser?.pts?.find((p) => p.idx === ptIdx);
    if (dp?.fill) return dp.fill;
    if (ser?.fill) return ser.fill;
    return '#' + palette[i % palette.length];
  };
  // étiquettes de données (c:dLbls) : afficher valeur / pourcentage sur les éléments
  const dlblFlags = (node) => { const d = node?.['c:dLbls'] || {}; return { val: d['c:showVal']?.['@val'] === '1', pct: d['c:showPercent']?.['@val'] === '1' }; };
  const fmtNum = (v) => (Math.round(v * 100) / 100);
  // marges : place pour labels d'axe (bas) et échelle (gauche). Camembert n'en a pas besoin.
  const isPie = plot['c:pieChart'] || plot['c:doughnutChart'] || plot['c:pie3DChart'];
  const isLine = plot['c:lineChart'] || plot['c:line3DChart'];
  const isArea = plot['c:areaChart'] || plot['c:area3DChart'];
  const barC = plot['c:barChart'] || plot['c:bar3DChart'];
  // titre du graphique (c:title>c:tx>c:rich) — sauf si autoTitleDeleted
  const titleOf = () => {
    const t = cs['c:chart']?.['c:title'];
    if (!t || cs['c:chart']['c:autoTitleDeleted']?.['@val'] === '1') return '';
    const rich = t['c:tx']?.['c:rich']; if (!rich) return '';
    return arr(rich['a:p']).map((p) => arr(p['a:r']).map((r) => r['a:t']).join('')).join(' ').trim();
  };
  const title = titleOf();
  const hasLegend = !!cs['c:chart']?.['c:legend'];
  const legPos = cs['c:chart']?.['c:legend']?.['c:legendPos']?.['@val'] || 'r';
  // entrées de légende : séries (cartésien) ou catégories (camembert)
  const chartNode = isPie || barC || isLine || isArea || plot['c:scatterChart'] || plot['c:bubbleChart'] || plot['c:radarChart'];
  const sersAll = chartNode ? arr(chartNode['c:ser']).map((s) => serOf(s, color)) : [];
  let legend = [];
  if (hasLegend && sersAll.length) {
    if (isPie) { const s0 = sersAll[0]; const cats = s0.cats.length ? s0.cats : s0.vals.map((_, i) => 'Élément ' + (i + 1)); legend = cats.map((c, i) => ({ name: c, col: paint(i, s0, i) })); }
    else legend = sersAll.map((s, i) => ({ name: s.name || 'Série ' + (i + 1), col: paint(i, s, 0) }));
  }
  // réservation d'espace : titre en haut, légende à droite (r/l) ou en bas (t/b)
  const titleH = title ? 18 : 0;
  const legRight = (hasLegend && legend.length && (legPos === 'r' || legPos === 'l' || legPos === 'tr')) ? Math.min(140, 20 + Math.max(...legend.map((e) => e.name.length)) * 5.5) : 0;
  const legBottom = (hasLegend && legend.length && (legPos === 't' || legPos === 'b')) ? 16 : 0;
  const ox = 0, oy = titleH;
  const cw = Math.max(10, w - legRight), ch = Math.max(10, h - titleH - legBottom);

  const plotSVG = (w, h) => {
  let out = '';

  if (isPie) {
    const chart = isPie;
    const sers = arr(chart['c:ser']).map((s) => serOf(s, color));
    const s0 = sers[0]; if (!s0 || !s0.vals.length) return '';
    const total = s0.vals.reduce((a, b) => a + Math.abs(b), 0) || 1;
    const cx = w / 2, cy = h / 2, R = Math.min(w, h) / 2 * 0.92;
    const inner = plot['c:doughnutChart'] ? R * (Math.min(0.85, Math.max(0.1, (+plot['c:doughnutChart']['c:holeSize']?.['@val'] || 50) / 100))) : 0;
    const dl = dlblFlags(chart);
    let a0 = -Math.PI / 2; // 12 h
    s0.vals.forEach((v, i) => {
      const frac = Math.abs(v) / total; const a1 = a0 + frac * 2 * Math.PI;
      const x0 = cx + R * Math.cos(a0), y0 = cy + R * Math.sin(a0);
      const x1 = cx + R * Math.cos(a1), y1 = cy + R * Math.sin(a1);
      const large = frac > 0.5 ? 1 : 0;
      const col = paint(i, s0, i);
      if (inner > 0) {
        const ix0 = cx + inner * Math.cos(a0), iy0 = cy + inner * Math.sin(a0);
        const ix1 = cx + inner * Math.cos(a1), iy1 = cy + inner * Math.sin(a1);
        out += `<path d="M${x0.toFixed(1)} ${y0.toFixed(1)} A${R.toFixed(1)} ${R.toFixed(1)} 0 ${large} 1 ${x1.toFixed(1)} ${y1.toFixed(1)} L${ix1.toFixed(1)} ${iy1.toFixed(1)} A${inner.toFixed(1)} ${inner.toFixed(1)} 0 ${large} 0 ${ix0.toFixed(1)} ${iy0.toFixed(1)} Z" fill="${col}" stroke="#fff" stroke-width="1.5"/>`;
      } else {
        out += `<path d="M${cx.toFixed(1)} ${cy.toFixed(1)} L${x0.toFixed(1)} ${y0.toFixed(1)} A${R.toFixed(1)} ${R.toFixed(1)} 0 ${large} 1 ${x1.toFixed(1)} ${y1.toFixed(1)} Z" fill="${col}" stroke="#fff" stroke-width="1.5"/>`;
      }
      if ((dl.pct || dl.val) && frac > 0.03) { // étiquette au milieu de la part
        const am = a0 + frac * Math.PI, rl = inner > 0 ? (R + inner) / 2 : R * 0.62;
        const lx = cx + rl * Math.cos(am), ly = cy + rl * Math.sin(am);
        const txt = dl.pct ? Math.round(frac * 100) + '%' : String(fmtNum(v));
        out += `<text x="${lx.toFixed(1)}" y="${(ly + 3).toFixed(1)}" font-size="9" text-anchor="middle" fill="#fff" font-weight="bold">${esc(txt)}</text>`;
      }
      a0 = a1;
    });
    return out;
  }

  // nuage de points / bulles (c:scatterChart / c:bubbleChart) : axes X ET Y numériques
  const scat = plot['c:scatterChart'] || plot['c:bubbleChart'];
  if (scat) {
    const sers = arr(scat['c:ser']).map((s) => ({
      name: ptsOf(s['c:tx']?.['c:strRef'])[0] ?? '', fill: s['c:spPr'] ? color(s['c:spPr']) : null,
      xs: ptsOf(s['c:xVal']?.['c:numRef']).map(num), ys: ptsOf(s['c:yVal']?.['c:numRef']).map(num),
    }));
    if (!sers.length) return '';
    let xmn = Infinity, xmx = -Infinity, ymn = Infinity, ymx = -Infinity;
    sers.forEach((s) => s.xs.forEach((x, i) => { const y = s.ys[i]; if (x == null || y == null) return; xmn = Math.min(xmn, x); xmx = Math.max(xmx, x); ymn = Math.min(ymn, y); ymx = Math.max(ymx, y); }));
    if (!Number.isFinite(xmn)) return '';
    if (xmn === xmx) xmx = xmn + 1; if (ymn === ymx) ymx = ymn + 1;
    const mL = 34, mB = 20, mT = 8, mR = 8, pw = Math.max(1, w - mL - mR), ph = Math.max(1, h - mT - mB);
    const PX = (x) => mL + pw * (x - xmn) / (xmx - xmn), PY = (y) => mT + ph * (1 - (y - ymn) / (ymx - ymn));
    for (let t = 0; t <= 4; t++) { const y = mT + ph * t / 4, v = ymx - (ymx - ymn) * t / 4;
      out += `<line x1="${mL}" y1="${y.toFixed(1)}" x2="${(w - mR).toFixed(1)}" y2="${y.toFixed(1)}" stroke="#e0e0e0" stroke-width="0.5"/><text x="${(mL - 3).toFixed(1)}" y="${(y + 3).toFixed(1)}" font-size="9" text-anchor="end" fill="#666">${esc(fmtNum(v))}</text>`; }
    for (let t = 0; t <= 4; t++) { const x = mL + pw * t / 4, v = xmn + (xmx - xmn) * t / 4;
      out += `<text x="${x.toFixed(1)}" y="${(h - mB + 12).toFixed(1)}" font-size="9" text-anchor="middle" fill="#666">${esc(fmtNum(v))}</text>`; }
    sers.forEach((s, si) => { const col = paint(si, s, 0);
      for (let i = 0; i < s.xs.length; i++) { const x = s.xs[i], y = s.ys[i]; if (x == null || y == null) continue;
        out += `<circle cx="${PX(x).toFixed(1)}" cy="${PY(y).toFixed(1)}" r="3" fill="${col}"/>`; } });
    return out;
  }

  // radar (c:radarChart) : catégories réparties en cercle, chaque série = polygone
  const radar = plot['c:radarChart'];
  if (radar) {
    const sers = arr(radar['c:ser']).map((s) => serOf(s, color));
    if (!sers.length) return '';
    const cats = sers[0].cats.length ? sers[0].cats : sers[0].vals.map((_, i) => String(i + 1));
    const n = Math.max(3, cats.length);
    let vmax = 0; sers.forEach((s) => s.vals.forEach((v) => { vmax = Math.max(vmax, v); })); if (!vmax) vmax = 1;
    const cx = w / 2, cy = h / 2 + 4, R = Math.min(w, h) / 2 * 0.78;
    const ang = (i) => -Math.PI / 2 + 2 * Math.PI * i / n;   // catégorie i, départ en haut
    const PT = (i, r) => [cx + r * Math.cos(ang(i)), cy + r * Math.sin(ang(i))];
    for (let g = 1; g <= 4; g++) { const rr = R * g / 4; // anneaux concentriques
      let d = ''; for (let i = 0; i < n; i++) { const [x, y] = PT(i, rr); d += (i ? 'L' : 'M') + x.toFixed(1) + ' ' + y.toFixed(1); }
      out += `<path d="${d}Z" fill="none" stroke="#e0e0e0" stroke-width="0.5"/>`; }
    for (let i = 0; i < n; i++) { const [x, y] = PT(i, R); out += `<line x1="${cx.toFixed(1)}" y1="${cy.toFixed(1)}" x2="${x.toFixed(1)}" y2="${y.toFixed(1)}" stroke="#e0e0e0" stroke-width="0.5"/>`;
      const [lx, ly] = PT(i, R + 10); out += `<text x="${lx.toFixed(1)}" y="${(ly + 3).toFixed(1)}" font-size="8" text-anchor="middle" fill="#666">${esc(cats[i] ?? '')}</text>`; }
    sers.forEach((s, si) => { const col = paint(si, s, 0); let d = '';
      for (let i = 0; i < n; i++) { const [x, y] = PT(i, R * (s.vals[i] || 0) / vmax); d += (i ? 'L' : 'M') + x.toFixed(1) + ' ' + y.toFixed(1); }
      out += `<path d="${d}Z" fill="${col}" fill-opacity="0.25" stroke="${col}" stroke-width="2"/>`; });
    return out;
  }

  // colonnes / barres / lignes / aires : axes cartésiens
  const chart = barC || isLine || isArea;
  if (!chart) return '';
  const sers = arr(chart['c:ser']).map((s) => serOf(s, color));
  if (!sers.length) return '';
  const dl = dlblFlags(chart);
  const cats = sers[0].cats.length ? sers[0].cats : sers[0].vals.map((_, i) => String(i + 1));
  const nCat = Math.max(1, cats.length, ...sers.map((s) => s.vals.length));
  const stacked = /stacked/i.test(chart['c:grouping']?.['@val'] || '');
  const horiz = barC && (barC['c:barDir']?.['@val'] === 'bar');
  // domaine des valeurs (avec 0 et empilement éventuel)
  let vmax = 0, vmin = 0;
  if (stacked) for (let c = 0; c < nCat; c++) { let pos = 0, neg = 0; sers.forEach((s) => { const v = s.vals[c] || 0; if (v >= 0) pos += v; else neg += v; }); vmax = Math.max(vmax, pos); vmin = Math.min(vmin, neg); }
  else sers.forEach((s) => s.vals.forEach((v) => { vmax = Math.max(vmax, v); vmin = Math.min(vmin, v); }));
  // bornes d'axe EXPLICITES (c:valAx>c:scaling>c:min/c:max) prioritaires sur l'auto-échelle
  const valAx0 = arr(plot['c:valAx'])[0];
  const scl = valAx0?.['c:scaling'];
  if (scl?.['c:max']?.['@val'] != null) vmax = num(scl['c:max']['@val']);
  if (scl?.['c:min']?.['@val'] != null) vmin = num(scl['c:min']['@val']);
  if (vmax === vmin) vmax = vmin + 1;

  // BARRES HORIZONTALES (barDir="bar") : axe des catégories vertical, axe des valeurs horizontal
  if (horiz) {
    const mLh = 62, mBh = 20, mTh = 8, mRh = 12;
    const pwh = Math.max(1, w - mLh - mRh), phh = Math.max(1, h - mTh - mBh);
    const VX = (v) => mLh + pwh * (v - vmin) / (vmax - vmin);  // valeur -> x
    const CY = (i) => mTh + phh * (i + 0.5) / nCat;            // catégorie i -> y (centre)
    for (let t = 0; t <= 4; t++) { const v = vmin + (vmax - vmin) * t / 4, x = VX(v);
      out += `<line x1="${x.toFixed(1)}" y1="${mTh}" x2="${x.toFixed(1)}" y2="${(h - mBh).toFixed(1)}" stroke="#d0d0d0" stroke-width="0.5"/><text x="${x.toFixed(1)}" y="${(h - mBh + 12).toFixed(1)}" font-size="9" text-anchor="middle" fill="#666">${esc(fmtNum(v))}</text>`; }
    cats.forEach((c, i) => { if (i < nCat) out += `<text x="${(mLh - 4).toFixed(1)}" y="${(CY(i) + 3).toFixed(1)}" font-size="9" text-anchor="end" fill="#666">${esc(c)}</text>`; });
    const groupH = phh / nCat * 0.8, bh = stacked ? groupH : groupH / sers.length;
    const vary = (barC['c:varyColors']?.['@val'] === '1') && sers.length === 1;
    for (let c = 0; c < nCat; c++) {
      let posR = VX(0), negR = VX(0);
      sers.forEach((s, si) => {
        const v = s.vals[c] || 0; if (!v) return; const col = paint(vary ? c : si, s, c);
        const wdt = Math.abs(VX(v) - VX(0)); let x, y;
        if (stacked) { y = CY(c) - groupH / 2; if (v >= 0) { x = posR; posR += wdt; } else { negR -= wdt; x = negR; } }
        else { y = CY(c) - groupH / 2 + si * bh; x = Math.min(VX(0), VX(v)); }
        const bhEff = stacked ? bh : bh * 0.92;
        out += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${wdt.toFixed(1)}" height="${bhEff.toFixed(1)}" fill="${col}"/>`;
        if (dl.val && !stacked) out += `<text x="${(x + wdt + 3).toFixed(1)}" y="${(y + bhEff / 2 + 3).toFixed(1)}" font-size="8" fill="#555">${esc(fmtNum(v))}</text>`;
      });
    }
    out += `<line x1="${VX(0).toFixed(1)}" y1="${mTh}" x2="${VX(0).toFixed(1)}" y2="${(h - mBh).toFixed(1)}" stroke="#888" stroke-width="1"/>`;
    return out;
  }

  const mL = 34, mB = 22, mT = 8, mR = 8; // marges pour échelle + labels
  const pw = Math.max(1, w - mL - mR), ph = Math.max(1, h - mT - mB);
  const X = (i) => mL + (pw * (i + 0.5)) / nCat;             // centre de la catégorie i
  const Y = (v) => mT + ph * (1 - (v - vmin) / (vmax - vmin)); // valeur -> y
  // grille + axe des valeurs (5 graduations)
  const ticks = 4;
  for (let t = 0; t <= ticks; t++) {
    const v = vmin + (vmax - vmin) * t / ticks; const y = Y(v);
    out += `<line x1="${mL}" y1="${y.toFixed(1)}" x2="${(w - mR).toFixed(1)}" y2="${y.toFixed(1)}" stroke="#d0d0d0" stroke-width="0.5"/>`;
    out += `<text x="${(mL - 3).toFixed(1)}" y="${(y + 3).toFixed(1)}" font-size="9" text-anchor="end" fill="#666">${esc(Math.round(v * 100) / 100)}</text>`;
  }
  // axe des catégories
  cats.forEach((c, i) => { if (i < nCat) out += `<text x="${X(i).toFixed(1)}" y="${(h - mB + 12).toFixed(1)}" font-size="9" text-anchor="middle" fill="#666">${esc(c)}</text>`; });

  if (barC) {
    const groupW = pw / nCat * 0.8;               // largeur allouée par catégorie
    const bw = stacked ? groupW : groupW / sers.length; // largeur d'une barre
    // varyColors : série unique dont CHAQUE barre a une couleur différente (par catégorie)
    const vary = (barC['c:varyColors']?.['@val'] === '1') && sers.length === 1;
    for (let c = 0; c < nCat; c++) {
      let posTop = Y(0), negTop = Y(0);           // sommets courants des piles positive/négative
      sers.forEach((s, si) => {
        const v = s.vals[c] || 0; if (!v) return; const col = paint(vary ? c : si, s, c);
        const hgt = Math.abs(Y(v) - Y(0));         // hauteur du segment en px
        let x, y;
        if (stacked) { x = X(c) - groupW / 2; if (v >= 0) { posTop -= hgt; y = posTop; } else { y = negTop; negTop += hgt; } }
        else { x = X(c) - groupW / 2 + si * bw; y = Math.min(Y(0), Y(v)); }
        const bwEff = stacked ? bw : bw * 0.92;
        out += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bwEff.toFixed(1)}" height="${hgt.toFixed(1)}" fill="${col}"/>`;
        if (dl.val) out += `<text x="${(x + bwEff / 2).toFixed(1)}" y="${(stacked ? y + hgt / 2 + 3 : y - 3).toFixed(1)}" font-size="8" text-anchor="middle" fill="${stacked ? '#fff' : '#555'}">${esc(fmtNum(v))}</text>`;
      });
    }
  } else if (isArea) {
    sers.forEach((s, si) => {
      const col = paint(si, s, 0);
      let d = `M${X(0).toFixed(1)} ${Y(0).toFixed(1)}`;
      for (let c = 0; c < nCat; c++) d += ` L${X(c).toFixed(1)} ${Y(s.vals[c] || 0).toFixed(1)}`;
      d += ` L${X(nCat - 1).toFixed(1)} ${Y(0).toFixed(1)} Z`;
      out += `<path d="${d}" fill="${col}" fill-opacity="0.6" stroke="${col}" stroke-width="1"/>`;
    });
  } else { // lignes
    sers.forEach((s, si) => {
      const col = paint(si, s, 0);
      let d = '';
      for (let c = 0; c < nCat; c++) { const v = s.vals[c]; if (v == null) continue; d += (d ? ' L' : 'M') + X(c).toFixed(1) + ' ' + Y(v).toFixed(1); }
      out += `<path d="${d}" fill="none" stroke="${col}" stroke-width="2"/>`;
      for (let c = 0; c < nCat; c++) { const v = s.vals[c]; if (v == null) continue; out += `<circle cx="${X(c).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="2.5" fill="${col}"/>`; if (dl.val) out += `<text x="${X(c).toFixed(1)}" y="${(Y(v) - 5).toFixed(1)}" font-size="8" text-anchor="middle" fill="#555">${esc(fmtNum(v))}</text>`; }
    });
  }
  // axe zéro
  out += `<line x1="${mL}" y1="${Y(0).toFixed(1)}" x2="${(w - mR).toFixed(1)}" y2="${Y(0).toFixed(1)}" stroke="#888" stroke-width="1"/>`;
  return out;
  }; // fin plotSVG

  const body = plotSVG(cw, ch);
  if (!body) return '';
  let svg = `<g transform="translate(${ox} ${oy})">${body}</g>`;
  if (title) svg += `<text x="${(w / 2).toFixed(1)}" y="13" font-size="12" font-weight="bold" text-anchor="middle" fill="#333">${esc(title)}</text>`;
  if (legend.length) {
    if (legRight) { // colonne à droite
      const x = w - legRight + 6; let y = oy + Math.max(6, (ch - legend.length * 14) / 2) + 8;
      for (const e of legend) { svg += `<rect x="${x}" y="${(y - 8).toFixed(1)}" width="9" height="9" fill="${e.col}"/><text x="${x + 13}" y="${y.toFixed(1)}" font-size="9" fill="#444">${esc(e.name)}</text>`; y += 14; }
    } else if (legBottom) { // ligne en bas, centrée
      const gap = 14, sw = 9; const widths = legend.map((e) => sw + 4 + e.name.length * 5.2 + gap);
      let x = Math.max(4, (w - widths.reduce((a, b) => a + b, 0)) / 2); const y = h - 4;
      for (let i = 0; i < legend.length; i++) { const e = legend[i]; svg += `<rect x="${x.toFixed(1)}" y="${(y - 8).toFixed(1)}" width="${sw}" height="9" fill="${e.col}"/><text x="${(x + sw + 3).toFixed(1)}" y="${y.toFixed(1)}" font-size="9" fill="#444">${esc(e.name)}</text>`; x += widths[i]; }
    }
  }
  return svg;
}

// --- exports nommés pour les TESTS (fonctions PURES uniquement) ---
// Helpers sans effet de bord (pas de fs/réseau/DOM/état global). renderChart est déjà exporté.
export { esc, arr, num, ptsOf, serOf };
