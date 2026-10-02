// ═══ [DIAPO · CADRES] ═══ p:graphicFrame : graphiques, SmartArt, tableaux, objets OLE (image de secours).
import { readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { P, PF, cachedParse, tagOf, at, byTag, first, deepFind, esc, serialize } from '../xml.mjs';
import { u } from '../units.mjs';
import { resolveColorFlat, cssColor } from '../color.mjs';
import { cellTextHeight } from '../fonts.mjs';
import { renderChart } from '../../chart-render.mjs';
import { walk, applyTf } from './walk.mjs';
import { paint } from './paint.mjs';
import { renderPic } from './picture.mjs';
import { renderText } from './text.mjs';

const arr = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];

// Un p:graphicFrame est un cadre générique dont a:graphic/a:graphicData porte un contenu typé :
// tableau (a:tbl), graphique (c:chart ou cx:chart), SmartArt (dgm:relIds), objet OLE… Point
// d'entrée unique : renderGraphicFrame, qui aiguille vers renderTable / vers renderChartFrame / renderDiagram.
// Ces cadres n'ont ni rotation ni miroir ici (rot:0 forcé).
//
// graphique (p:graphicFrame > c:chart) : lit la part ppt/charts/chartN.xml (réf r:id) et la rend en SVG
// via chart-render.mjs. Fond blanc (défaut PowerPoint) + palette = accents du thème.
function renderChartFrame(rc, chartRef, off, ext, tf) {
  const { ctx, curRels, pptDir } = rc;
  const rId = chartRef['@r:id'] || chartRef['@r:embed'];
  const tgt = rId && curRels[rId];
  if (!tgt) return '';
  const chartPath = join(pptDir, tgt.replace(/^\.\.\//, ''));
  if (!existsSync(chartPath)) return '';
  let chartFlat; try { chartFlat = cachedParse(chartPath, 'flat'); } catch { return ''; }
  const g = applyTf({ x:+off['@x'], y:+off['@y'], cx:+ext['@cx'], cy:+ext['@cy'], rot:0 }, tf);
  const W2 = u(g.cx), H2 = u(g.cy);
  // palette = hex des 6 accents du thème (couleurs automatiques des séries, dans l'ordre) ; color =
  // rappel pour que chart-render.mjs résolve les solidFill (thème + clrMap) sans connaître ctx
  const palette = ['accent1','accent2','accent3','accent4','accent5','accent6'].map((a) => ctx.clr[a]).filter(Boolean);
  const color = (sp) => sp?.['a:solidFill'] ? cssColor(resolveColorFlat(sp['a:solidFill'], ctx)) : null;
  let inner; try { inner = renderChart(chartFlat, W2, H2, { palette, color }); } catch { return ''; }
  if (!inner) return '';
  // fond blanc du graphique (sauf si c:chartSpace définit noFill) puis contenu
  const cs = chartFlat['c:chartSpace'] || chartFlat;
  const bg = cs['c:spPr']?.['a:noFill'] !== undefined ? '' : `<rect width="${W2.toFixed(1)}" height="${H2.toFixed(1)}" fill="#ffffff"/>`;
  return `<g transform="translate(${u(g.x).toFixed(1)} ${u(g.y).toFixed(1)})">${bg}${inner}</g>`;
}
// SmartArt : la part de dessin de secours dsp:drawing (ppt/diagrams/drawingN.xml) contient un spTree
// de vraies formes DrawingML, structurellement identique à un p:spTree mais préfixé « dsp: ». On
// remplace « dsp: » -> « p: » et on réutilise le walk normal, avec un mapping repère-dessin -> cadre.
// Ne renvoie rien : pousse directement les formes dans body (via walk). Conséquence : les formes du
// SmartArt ne sont pas enveloppées dans un <g data-spid> du cadre (une animation ciblant le SmartArt
// entier ne trouve pas de cible) ; leurs data-spid sont ceux du dessin dsp (espace d'id distinct).
function renderDiagram(rc, relId, off, ext, tf) {
  const { curRels, pptDir } = rc;
  const tgt = curRels[relId]; if (!tgt) return;
  const p = join(pptDir, tgt.replace(/^\.\.\//, ''));
  if (!existsSync(p)) return;
  let txt; try { txt = readFileSync(p, 'utf8'); } catch { return; }
  txt = txt.replace(/dsp:/g, 'p:'); // même structure que PresentationML (nvSpPr/spPr/txBody/style + a:*)
  let ord; try { ord = P.parse(txt); } catch { return; }
  const root = ord.find((n) => tagOf(n) === 'p:drawing');
  const spTree = root && deepFind(root, 'p:spTree'); if (!spTree) return;
  const frame = applyTf({ x:+off['@x'], y:+off['@y'], cx:+ext['@cx'], cy:+ext['@cy'], rot:0 }, tf);
  // mapping : chOff/chExt du grpSpPr du spTree (repère du dessin) -> cadre placé
  const gpr = first(spTree, 'p:grpSpPr'); const xf = gpr && first(gpr, 'a:xfrm');
  let dtf = { ox: frame.x, oy: frame.y, cx0: 0, cy0: 0, sx: 1, sy: 1 };
  if (xf) {
    const chOff = first(xf, 'a:chOff'), chExt = first(xf, 'a:chExt');
    const cx = +at(chExt)['@cx'] || frame.cx, cy = +at(chExt)['@cy'] || frame.cy;
    dtf = { ox: frame.x, oy: frame.y, cx0: +at(chOff)['@x'] || 0, cy0: +at(chOff)['@y'] || 0, sx: frame.cx / cx, sy: frame.cy / cy };
  }
  walk(rc, spTree, dtf);
}
// renderGraphicFrame(rc, p:graphicFrame ORDONNÉ, tf) -> fragment SVG (ou '' si délégué / vide).
// Aiguillage selon le contenu de a:graphicData.
export function renderGraphicFrame(rc, gf, tf) {
  const flat = PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(gf)}</root>`).root['p:graphicFrame'];
  if (!flat) return '';
  const xf = flat['p:xfrm']; const off = xf?.['a:off'], ext = xf?.['a:ext'];
  const gd = flat['a:graphic']?.['a:graphicData'];
  // un p:graphicFrame porte un TABLEAU (a:tbl) OU un GRAPHIQUE (c:chart, réf. r:id vers ppt/charts/…)
  const chartRef = gd?.['c:chart'] || gd?.['cx:chart'];
  if (off && ext && chartRef) return renderChartFrame(rc, chartRef, off, ext, tf);
  // SmartArt (dgm:relIds) : rendu de secours dsp:drawing référencé par dsp:dataModelExt@relId
  const dgm = gd?.['dgm:relIds'];
  if (off && ext && dgm) {
    const exts = arr(dgm['a:extLst']?.['a:ext']);
    let relId; for (const e of exts) { const dm = e['dsp:dataModelExt']; if (dm?.['@relId']) { relId = dm['@relId']; break; } }
    if (relId) renderDiagram(rc, relId, off, ext, tf); // pousse directement dans body
    return '';
  }
  const tbl = gd?.['a:tbl'];
  if (!off || !tbl) {
    // repli générique : un graphicFrame non-tableau/graphique/SmartArt (objet OLE incorporé, contrôle…)
    // embarque presque toujours une IMAGE DE SECOURS (p:pic). La rendre plutôt que laisser un vide.
    if (off && ext) { const picO = deepFind(gf, 'p:pic'); if (picO) return renderPic(rc, picO, tf); }
    return '';
  }
  return renderTable(rc, gf, flat, tbl, off, ext, tf);
}

// tableau (p:graphicFrame > a:tbl) : grille + bordures + fusions + texte
// renderTable(rc, gf ordonné, gf plat, a:tbl plat, a:off, a:ext, tf) -> fragment SVG.
// Tableau (a:tbl, §21.1.3) : a:tblGrid/a:gridCol@w = largeurs de colonnes ; a:tr@h = hauteur
// MINIMALE de ligne ; a:tc = cellule (gridSpan / rowSpan = fusion ; hMerge / vMerge = cellule
// couverte par une fusion, à ne pas dessiner) ; a:tcPr = fond, bordures lnL/lnR/lnT/lnB, marges,
// ancrage. Algorithme :
//   1. hauteurs de lignes (agrandies selon le texte, cf. cellTextHeight), positions colX / rowY ;
//   2. pour chaque cellule visible : couches du style de tableau -> fond, puis bordures, puis texte
//      (renderText dans un repère translaté au coin de la cellule).
function renderTable(rc, gf, flat, tbl, off, ext, tf) {
  const { ctx, spidPfx } = rc;
  const g = applyTf({ x:+off['@x'], y:+off['@y'], cx:+ext['@cx'], cy:+ext['@cy'], rot:0 }, tf);
  const cols = arr(tbl['a:tblGrid']?.['a:gridCol']).map(c => +c['@w']);
  const rows = arr(tbl['a:tr']);
  // hauteurs de lignes : spécifiées sinon réparties sur la hauteur du cadre
  let hs = rows.map(r => +r['@h'] || 0);
  // PowerPoint AGRANDIT une ligne dont le texte ne tient pas (le @h est un minimum) ; le cadre (ext cy)
  // enregistre la hauteur réelle. Répartir l'écart uniformément donnait des lignes trop hautes et une
  // dernière ligne qui débordait. -> hauteur = max(@h, texte coupé mot à mot avec les chasses arrondies
  // de PowerPoint). Cellules fusionnées verticalement ignorées ; mesure impossible -> @h conservé.
  {
    const tp0 = tbl['a:tblPr'] || {}, on0 = (f) => tp0['@' + f] === '1' || tp0['@' + f] === 'true';
    const colW = (ci, span) => cols.slice(ci, ci + span).reduce((a, b) => a + b, 0);
    rows.forEach((tr, ri) => {
      arr(tr['a:tc']).forEach((tc, ci) => {
        if (tc['@hMerge'] === '1' || tc['@vMerge'] === '1' || (+tc['@rowSpan'] || 1) > 1) return;
        const styleBold = (on0('firstRow') && ri === 0) || (on0('firstCol') && ci === 0);
        const need = cellTextHeight(tc, colW(ci, +tc['@gridSpan'] || 1), styleBold);
        if (need && need > hs[ri]) hs[ri] = need;
      });
    });
    // reste jusqu'à la hauteur ENREGISTRÉE du cadre : aux seules lignes qui ont grandi (c'est leur texte
    // qui a poussé le cadre) — l'étaler sur toutes décalait chaque ligne d'1 px de plus que la précédente.
    const h0 = rows.map(r => +r['@h'] || 0), grown = hs.map((h, i) => h - h0[i]);
    const gsum = grown.reduce((a, b) => a + b, 0), rest = (+ext['@cy']) - hs.reduce((a, b) => a + b, 0);
    if (gsum > 0 && rest > 0 && h0.every(h => h > 0)) hs = hs.map((h, i) => h + rest * grown[i] / gsum);
  }
  // lignes sans @h : se partagent également la hauteur restante du cadre
  const sumH = hs.reduce((a,b)=>a+b,0), zero = hs.filter(h=>!h).length;
  if (zero) { const rem = Math.max(0, (+ext['@cy']) - sumH) / zero; hs = hs.map(h => h || rem); }
  // abscisses / ordonnées cumulées (EMU, depuis le coin du tableau) : colX[i] = bord gauche de la col. i
  const colX = [0]; cols.forEach((w,i)=>colX[i+1]=colX[i]+w);
  const rowY = [0]; hs.forEach((h,i)=>rowY[i+1]=rowY[i]+h);
  // style de tableau (fills header/bandes) — best-effort
  const styleId = tbl['a:tblPr']?.['a:tableStyleId'];
  const tp = tbl['a:tblPr'] || {};
  const ts = styleId && ctx.tableStyles?.[styleId];
  // drapeaux du tblPr (firstRow, lastRow, firstCol, lastCol, bandRow, bandCol) : quelles couches du
  // style s'appliquent. Absent = faux.
  const on = (flag) => tp['@' + flag] === '1' || tp['@' + flag] === 'true';
  // bordure PAR DÉFAUT des styles « grille » INTÉGRÉS (non stockés dans le fichier) : {5940675A...} =
  // « No Style, Table Grid » = toutes bordures (couleur tx1, 1pt). Sinon les cellules à <a:tcPr/> vide
  // n'avaient aucune grille (ex. d.46). Utilisée quand la cellule ne définit pas sa propre bordure.
  const GRID_STYLES = new Set(['{5940675A-B579-460E-94D1-54222C63F5DA}', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}']);
  const dfltB = GRID_STYLES.has(styleId) ? { '@w': '12700', 'a:solidFill': { 'a:schemeClr': { '@val': 'tx1' } } } : null;
  let out = '';
  // la grille est mise à l'échelle du cadre (a:ext) : si la somme des colonnes / lignes diffère de la
  // taille enregistrée du cadre, c'est le cadre qui fait foi
  const sx = g.cx / (colX[cols.length] || 1), sy = g.cy / (rowY[rows.length] || 1); // ajuste à la taille réelle
  // lignes/cellules du parse ORDONNÉ, alignées sur le parse plat : renderText en a besoin pour les
  // sauts de ligne <a:br/> et champs intercalés (perdus sinon dans les cellules).
  const tblO = deepFind(gf, 'a:tbl'), rowsO = tblO ? byTag(tblO, 'a:tr') : [];
  rows.forEach((tr, ri) => {
    const cellsO = rowsO[ri] ? byTag(rowsO[ri], 'a:tc') : [];
    arr(tr['a:tc']).forEach((tc, ci) => {
      if (tc['@hMerge'] === '1' || tc['@vMerge'] === '1') return; // cellule fusionnée (couverte)
      const span = +tc['@gridSpan'] || 1, rspan = +tc['@rowSpan'] || 1;
      const cx = g.x + colX[ci] * sx, cy = g.y + rowY[ri] * sy;
      const cw = ((colX[Math.min(ci+span, cols.length)] - colX[ci])) * sx;
      const ch = ((rowY[Math.min(ri+rspan, rows.length)] - rowY[ri])) * sy;
      const tcPr = tc['a:tcPr'] || {};
      // style de tableau en COUCHES (§20.1.4.2.*), de la plus faible à la plus forte : wholeTbl, bandes
      // verticales, bandes horizontales, dernière/première colonne, dernière/première ligne. Chacune
      // remplace fond, texte (tcTxStyle) et bordures qu'elle définit. Les bordures d'une couche sont
      // nommées par rapport à SA région (top/bottom/left/right/insideH/insideV).
      const lastR = ri + rspan >= rows.length, lastC = ci + span >= cols.length;
      const layers = [];
      if (ts) {
        const L = (name, kind) => { const n = ts['a:' + name]; if (n) layers.push({ n, kind }); };
        L('wholeTbl', 'whole');
        const bodyC = !(on('firstCol') && ci === 0) && !(on('lastCol') && lastC);
        const bodyR = !(on('firstRow') && ri === 0) && !(on('lastRow') && lastR);
        // bandes : la parité se compte À PARTIR de la 1re ligne/colonne de CORPS (l'en-tête ne compte
        // pas) ; les cellules d'en-tête / de total ne reçoivent pas de bande (bodyR / bodyC)
        if (on('bandCol') && bodyC) L(((ci - (on('firstCol') ? 1 : 0)) % 2 === 0) ? 'band1V' : 'band2V', 'col');
        if (on('bandRow') && bodyR) L(((ri - (on('firstRow') ? 1 : 0)) % 2 === 0) ? 'band1H' : 'band2H', 'row');
        if (on('lastCol') && lastC) L('lastCol', 'col');
        if (on('firstCol') && ci === 0) L('firstCol', 'col');
        if (on('lastRow') && lastR) L('lastRow', 'row');
        if (on('firstRow') && ri === 0) L('firstRow', 'row');
      }
      // parcours des couches de la plus faible à la plus forte : la dernière qui définit un fond gagne ;
      // les tcTxStyle se FUSIONNENT (attributs/couleur de la couche forte écrasent ceux de la faible)
      let cellFill = null, txStyle = null;
      for (const { n } of layers) {
        const f = n['a:tcStyle']?.['a:fill'];
        if (f) cellFill = f['a:noFill'] !== undefined ? null : paint(rc, f);
        if (n['a:tcTxStyle']) txStyle = { ...(txStyle || {}), ...n['a:tcTxStyle'] };
      }
      // fond PROPRE de la cellule prioritaire
      if (tcPr['a:noFill'] !== undefined) cellFill = null;
      else if (tcPr['a:solidFill'] || tcPr['a:gradFill']) cellFill = paint(rc, tcPr);
      if (cellFill && cellFill !== 'none')
        out += `<rect x="${u(cx)}" y="${u(cy)}" width="${u(cw)}" height="${u(ch)}" fill="${cellFill}"/>`;
      // bordure d'un bord : couches (clé selon la région) puis bordure propre (lnT/lnB/lnL/lnR)
      // Correspondance bord de cellule -> nom de bordure dans une couche :
      //   couche 'row' (bande H, 1re/dern. ligne) : top/bottom = bords de la ligne ; left/right hors
      //     bords du tableau = insideV ;
      //   couche 'col' : symétrique ;
      //   couche 'whole' : bords extérieurs = top/bottom/left/right, bords intérieurs = insideH/insideV.
      // Chaque cellule dessine ses 4 bords : un bord intérieur est donc tracé deux fois (par les deux
      // cellules voisines), sans conséquence visible pour des traits identiques.
      const edge = (side, own) => {
        let ln = dfltB;
        for (const { n, kind } of layers) {
          const bdr = n['a:tcStyle']?.['a:tcBdr']; if (!bdr) continue;
          const key = side === 'top' ? (kind === 'row' ? 'top' : (ri === 0 ? 'top' : 'insideH'))
            : side === 'bottom' ? (kind === 'row' ? 'bottom' : (lastR ? 'bottom' : 'insideH'))
            : side === 'left' ? (kind === 'col' ? 'left' : (ci === 0 ? 'left' : 'insideV'))
            : (kind === 'col' ? 'right' : (lastC ? 'right' : 'insideV'));
          if (bdr['a:' + key]?.['a:ln']) ln = bdr['a:' + key]['a:ln'];
        }
        return tcPr[own] || ln;
      };
      // border : une <line> par bord ; couleur non-solidFill -> gris #888 ; épaisseur défaut 1 pt.
      // Pointillés (prstDash) des bordures de cellule non gérés.
      const border = (ln, x1, y1, x2, y2) => {
        if (!ln || ln['a:noFill'] !== undefined) return '';
        const w = +ln['@w'] || 12700;
        const col = ln['a:solidFill'] ? cssColor(resolveColorFlat(ln['a:solidFill'], ctx)) : '#888';
        return `<line x1="${u(x1)}" y1="${u(y1)}" x2="${u(x2)}" y2="${u(y2)}" stroke="${col}" stroke-width="${u(w)}"/>`;
      };
      out += border(edge('top', 'a:lnT'), cx, cy, cx+cw, cy);
      out += border(edge('bottom', 'a:lnB'), cx, cy+ch, cx+cw, cy+ch);
      out += border(edge('left', 'a:lnL'), cx, cy, cx, cy+ch);
      out += border(edge('right', 'a:lnR'), cx+cw, cy, cx+cw, cy+ch);
      // texte du style (tcTxStyle : couleur + gras/italique + police de thème) -> lstStyle hérité
      // Le tcTxStyle (couleur, b/i en « on/off », fontRef major/minor) est TRADUIT en un faux lstStyle
      // { a:lvl1pPr: { a:defRPr: … } } passé à renderText comme style « hérité » : ainsi le rPr propre
      // de chaque run reste prioritaire, comme dans PowerPoint.
      let cellLst = null;
      if (txStyle) {
        const ck = Object.keys(txStyle).find(k => /^a:(schemeClr|srgbClr|prstClr|sysClr|scrgbClr|hslClr)$/.test(k));
        const d = {};
        if (ck) d['a:solidFill'] = { [ck]: txStyle[ck] };
        if (txStyle['@b']) d['@b'] = txStyle['@b'] === 'on' ? '1' : '0';
        if (txStyle['@i']) d['@i'] = txStyle['@i'] === 'on' ? '1' : '0';
        const fr = txStyle['a:fontRef']?.['@idx'];
        if (fr === 'major' || fr === 'minor') d['a:latin'] = { '@typeface': fr === 'major' ? '+mj-lt' : '+mn-lt' };
        cellLst = { 'a:lvl1pPr': { 'a:defRPr': d } };
      }
      // texte (ancrage vertical depuis tcPr @anchor)
      const tb = tc['a:txBody'];
      if (tb) {
        // mutation SANS risque : `flat` vient d'être re-parsé pour cet appel (pas un arbre en cache)
        tb['a:bodyPr'] = tb['a:bodyPr'] || {};
        if (!tb['a:bodyPr']['@anchor'] && tcPr['@anchor']) tb['a:bodyPr']['@anchor'] = tcPr['@anchor'];
        // marges de cellule (tcPr marL/R/T/B, défauts 0,1" / 0,05") -> insets du texte
        for (const [m, ins, dv] of [['@marL','@lIns',91440],['@marR','@rIns',91440],['@marT','@tIns',45720],['@marB','@bIns',45720]])
          if (tb['a:bodyPr'][ins] === undefined) tb['a:bodyPr'][ins] = tcPr[m] ?? dv;
        // 'none' : une cellule de tableau N'hérite PAS du style « corps » du master (souvent gras/gros
        // pour ces decks) — sinon texte trop gras/grand qui déborde. Elle utilise le défaut neutre + le
        // rPr propre de chaque run (ex. b="1" sur la 1ʳᵉ colonne seulement).
        const txt = renderText(rc, tb, cw, ch, 'none', cellLst, cellsO[ci] ? first(cellsO[ci], 'a:txBody') : null);
        if (txt) out += `<g transform="translate(${u(cx)} ${u(cy)})">${txt}</g>`;
      }
    });
  });
  const gid = flat['p:nvGraphicFramePr']?.['p:cNvPr']?.['@id'] || '';
  return `<g data-spid="${spidPfx}${esc(gid)}">${out}</g>`;
}
