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
//
// ─── COMMENT ÇA MARCHE (vue d'ensemble) ─────────────────────────────────────────────────────────
// PLACE DANS LE PIPELINE
//   pptx-render.mjs (renderChartFrame) rencontre un p:graphicFrame dont le a:graphicData contient un
//   <c:chart r:id="…"/>. Il suit la relation de la diapo vers ppt/charts/chartN.xml, parse ce XML en
//   mode « flat » (fast-xml-parser, attributs préfixés '@', parseTagValue:false -> TOUTES les valeurs
//   sont des chaînes), convertit la taille du cadre d'EMU en points (1 pt = 12700 EMU) et appelle
//   renderChart(chartFlat, W, H, { palette, color }). Il pose ensuite lui-même le fond blanc et le
//   translate() vers la position du cadre : ce module ne dessine QUE l'intérieur, dans 0..w × 0..h.
//
// FORMAT D'ENTRÉE (extrait de la structure DrawingML Chart, ECMA-376 Partie 1 §21.2)
//   c:chartSpace
//     c:chart
//       c:title > c:tx > c:rich > a:p > a:r > a:t      titre (texte riche)
//       c:autoTitleDeleted @val="1"                     titre explicitement supprimé
//       c:plotArea
//         c:barChart | c:lineChart | c:pieChart | …     UN nœud par famille de graphique
//           c:barDir @val = "col" (colonnes) | "bar" (barres horizontales)
//           c:grouping @val = clustered | stacked | percentStacked | standard
//           c:varyColors @val="1"                       une couleur par point (série unique)
//           c:dLbls > c:showVal / c:showPercent         étiquettes de données
//           c:ser (répété)                              séries
//             c:tx > c:strRef > c:strCache              nom de la série
//             c:cat > c:strRef|c:numRef > …Cache        catégories (axe X)
//             c:val > c:numRef > c:numCache             valeurs (axe Y)
//             c:xVal / c:yVal                           nuage de points (X et Y numériques)
//             c:spPr                                    remplissage de la série
//             c:dPt > c:idx + c:spPr                    surcharge de couleur d'UN point
//         c:valAx > c:scaling > c:min / c:max           bornes d'axe imposées
//       c:legend > c:legendPos @val = r | l | t | b | tr
//
//   Les caches (c:numCache / c:strCache) sont la copie que PowerPoint écrit des cellules Excel
//   référencées (c:f) : on ne lit jamais le classeur embarqué (ppt/embeddings/*.xlsx).
//
// SORTIE
//   Une chaîne SVG (sans <svg> englobant) : un <g transform="translate(0 titreH)"> contenant la zone
//   de tracé, suivi du titre et de la légende. Chaîne vide si rien d'exploitable (l'appelant n'affiche
//   alors rien plutôt qu'un cadre cassé).
//
// CE QUE CE MODULE NE FAIT PAS (approximations assumées)
//   - pas de mise en page automatique exacte de PowerPoint (c:layout / c:manualLayout ignorés) : les
//     marges sont des constantes en points, calibrées « à l'œil » ;
//   - pas de 3D (bar3DChart, pie3DChart… sont rendus à plat), pas d'axes secondaires, pas de
//     percentStacked réel (traité comme stacked : le regex /stacked/ matche aussi percentStacked) ;
//   - pas de formats de nombre (c:numFmt) : arrondi au centième ;
//   - police des textes = celle héritée du SVG parent (aucun font-family ici), tailles fixes 8-12 pt.
//
// DÉPENDANCES : aucune (module pur, sans fs/réseau/DOM) -> testable directement (test/docx-chart.test.js).
// ────────────────────────────────────────────────────────────────────────────────────────────────

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 1 — Petits utilitaires (échappement XML, normalisation de tableaux, nombres)
// ════════════════════════════════════════════════════════════════════════════════════════════════

// esc(s) : échappe &, <, > pour insérer du texte (titre, catégories, légende) dans un nœud <text>.
//   Les guillemets ne sont pas échappés : les valeurs échappées ne sont jamais placées dans un
//   attribut, seulement dans du contenu texte. null/undefined -> chaîne vide.
const esc = (s) => String(s ?? '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
// arr(x) : le parseur « flat » donne un OBJET quand un élément est unique et un TABLEAU quand il est
//   répété (ex. un seul <c:ser> -> objet, deux -> tableau). arr() ramène toujours à un tableau.
//   Seul undefined donne [] (null donnerait [null] — ne se produit pas avec ce parseur).
const arr = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];
// num(v) : chaîne XML -> nombre ; tout ce qui n'est pas un nombre fini (vide, "#N/A", undefined) vaut 0.
//   Conséquence : une cellule vide d'une série devient 0 (et non une interruption de courbe).
const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; };

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 2 — Lecture des données : caches de points et séries
// ════════════════════════════════════════════════════════════════════════════════════════════════

// extrait les points d'un cache (c:strCache/c:numCache) en tableau dense indexé par c:pt/@idx
// ptsOf(ref)
//   ref : un c:strRef / c:numRef (référence de cellules + cache) — ou directement un nœud qui porte
//         un littéral c:numLit / c:strLit (valeurs saisies sans classeur).
//   -> tableau de longueur c:ptCount ; case i = texte brut de <c:v> du point d'idx i, null si absent.
// Pourquoi « dense » : le cache peut SAUTER des indices (cellules vides non écrites). Indexer par
//   @idx garde l'alignement catégorie <-> valeur entre séries ; les trous restent null.
// Cas limites : pas de cache -> [] ; ptCount absent -> tableau vide mais les points sont quand même
//   posés par out[i] = … (le tableau s'agrandit) ; @idx absent -> 0.
function ptsOf(ref) {
  const cache = ref?.['c:numCache'] || ref?.['c:strCache'] || ref?.['c:numLit'] || ref?.['c:strLit'];
  if (!cache) return [];
  const cnt = +cache['c:ptCount']?.['@val'] || 0;          // nombre de points déclaré par PowerPoint
  const out = new Array(cnt).fill(null);
  for (const p of arr(cache['c:pt'])) { const i = +p['@idx'] || 0; out[i] = p['c:v']; } // <c:pt idx="i"><c:v>…</c:v></c:pt>
  return out;
}
// une série -> { name, cats:[..], vals:[num..], fill:hex|null, pts:[{fill}] }
// serOf(ser, color)
//   ser   : nœud c:ser parsé à plat.
//   color : résolveur fourni par l'appelant (spPr -> couleur CSS ou null).
//   -> { name   : nom de la série (1er point du strCache de c:tx, sinon texte littéral c:tx>c:v),
//        cats   : libellés de catégories (texte ou nombres en chaîne ; trous -> ''),
//        vals   : valeurs numériques (trous -> 0 via num()),
//        fill   : couleur de la série (c:spPr) ou null -> palette de thème,
//        pts    : surcharges par point [{ idx, fill }] issues des c:dPt (ex. une part de camembert
//                 recolorée à la main). }
//   Note : fill vaut aussi null quand spPr existe mais sans a:solidFill (dégradé, motif…) : le
//   résolveur renvoie null, et paint() retombe alors sur la palette.
function serOf(ser, color) {
  const name = ptsOf(ser['c:tx']?.['c:strRef'])[0] ?? ser['c:tx']?.['c:v'] ?? '';
  const cats = ptsOf(ser['c:cat']?.['c:strRef'] || ser['c:cat']?.['c:numRef']).map((v) => v ?? ''); // catégories texte OU numériques (dates, années)
  const vals = ptsOf(ser['c:val']?.['c:numRef']).map(num);
  const fill = ser['c:spPr'] ? color(ser['c:spPr']) : null;
  const pts = arr(ser['c:dPt']).map((d) => ({ idx: +d['c:idx']?.['@val'] || 0, fill: d['c:spPr'] ? color(d['c:spPr']) : null }));
  return { name, cats, vals, fill, pts };
}

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 3 — renderChart : point d'entrée public
// ════════════════════════════════════════════════════════════════════════════════════════════════
// Déroulé :
//   1. options + repli de palette, repérage de c:plotArea ;
//   2. détection de la famille de graphique (camembert, lignes, aires, barres, nuage, radar) ;
//   3. titre et légende -> réservation d'espace (titre en haut, légende à droite ou en bas) ;
//   4. plotSVG(cw, ch) dessine la zone de tracé dans le rectangle restant (branches par famille) ;
//   5. assemblage : zone de tracé décalée sous le titre, puis titre, puis légende.
// Unités : tout est en points (l'unité du repère SVG de la diapo), w/h compris.
// Couleurs : la palette arrive en hex SANS '#' (accents du thème) ; color() renvoie du CSS complet
//   (rgb(...)/rgba(...)) ; paint() normalise en ajoutant '#' au seul cas palette.

export function renderChart(chartFlat, w, h, opts = {}) {
  // Palette de repli = accents 1..6 du thème Office par défaut (bleu, orange, gris, or, bleu clair, vert).
  const palette = opts.palette?.length ? opts.palette : ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  const color = opts.color || (() => null);                // sans résolveur : tout vient de la palette
  const cs = chartFlat?.['c:chartSpace'] || chartFlat; // tolère racine chartSpace ou son contenu
  const plot = cs?.['c:chart']?.['c:plotArea'];
  if (!plot) return '';                                    // pas de zone de tracé : rien à dessiner
  // couleur d'un élément : point explicite (c:dPt) -> série (c:spPr) -> palette de thème (hex sans #)
  // paint(i, ser, ptIdx)
  //   i     : rang utilisé pour la palette (rang de SÉRIE en général ; rang de CATÉGORIE pour un
  //           camembert ou un graphique varyColors, où chaque point a sa couleur) ;
  //   ser   : série issue de serOf() (peut être un objet partiel : le nuage de points n'a pas de pts) ;
  //   ptIdx : index du point, pour chercher une surcharge c:dPt.
  //   Même ordre de priorité que PowerPoint : le plus spécifique gagne. La palette boucle (i % n),
  //   PowerPoint, lui, décline ensuite les accents en teintes plus claires/foncées : au-delà de
  //   6 séries les couleurs se répètent donc ici alors qu'elles varient dans PowerPoint.
  const paint = (i, ser, ptIdx) => {
    const dp = ser?.pts?.find((p) => p.idx === ptIdx);
    if (dp?.fill) return dp.fill;
    if (ser?.fill) return ser.fill;
    return '#' + palette[i % palette.length];
  };
  // étiquettes de données (c:dLbls) : afficher valeur / pourcentage sur les éléments
  //   Lu au niveau du GROUPE (c:barChart>c:dLbls…), pas par série (c:ser>c:dLbls ignoré).
  //   Seule la forme "1" est reconnue ("true", autorisé par xsd:boolean, ne l'est pas).
  const dlblFlags = (node) => { const d = node?.['c:dLbls'] || {}; return { val: d['c:showVal']?.['@val'] === '1', pct: d['c:showPercent']?.['@val'] === '1' }; };
  // fmtNum : arrondi au centième, sans zéros inutiles (12.5 et non 12.50). Aucun c:numFmt appliqué.
  const fmtNum = (v) => (Math.round(v * 100) / 100);
  // marges : place pour labels d'axe (bas) et échelle (gauche). Camembert n'en a pas besoin.
  // Détection de famille : chaque variable vaut le NŒUD du graphique (objet truthy) ou undefined.
  // Si plusieurs familles coexistent (graphique combiné barres + ligne), une seule est rendue,
  // selon l'ordre des tests de plotSVG (camembert > nuage > radar > barres > lignes > aires).
  const isPie = plot['c:pieChart'] || plot['c:doughnutChart'] || plot['c:pie3DChart'];
  const isLine = plot['c:lineChart'] || plot['c:line3DChart'];
  const isArea = plot['c:areaChart'] || plot['c:area3DChart'];
  const barC = plot['c:barChart'] || plot['c:bar3DChart'];
  // ── Titre et légende ────────────────────────────────────────────────────────────────────────
  // titre du graphique (c:title>c:tx>c:rich) — sauf si autoTitleDeleted
  // titleOf() -> texte brut : runs a:r concaténés par paragraphe, paragraphes joints par une espace.
  //   Limite : un c:title SANS c:tx (titre automatique = nom de la série unique dans PowerPoint)
  //   n'affiche rien ici ; un titre lié à une cellule (c:tx>c:strRef) non plus.
  const titleOf = () => {
    const t = cs['c:chart']?.['c:title'];
    if (!t || cs['c:chart']['c:autoTitleDeleted']?.['@val'] === '1') return '';
    const rich = t['c:tx']?.['c:rich']; if (!rich) return '';
    return arr(rich['a:p']).map((p) => arr(p['a:r']).map((r) => r['a:t']).join('')).join(' ').trim();
  };
  const title = titleOf();
  const hasLegend = !!cs['c:chart']?.['c:legend'];         // légende présente = élément c:legend existe
  const legPos = cs['c:chart']?.['c:legend']?.['c:legendPos']?.['@val'] || 'r'; // défaut du schéma : r (droite)
  // entrées de légende : séries (cartésien) ou catégories (camembert)
  // chartNode = premier nœud de graphique trouvé, uniquement pour lire ses séries (sersAll).
  const chartNode = isPie || barC || isLine || isArea || plot['c:scatterChart'] || plot['c:bubbleChart'] || plot['c:radarChart'];
  const sersAll = chartNode ? arr(chartNode['c:ser']).map((s) => serOf(s, color)) : [];
  // legend = [{ name, col }] : une entrée par catégorie (camembert : chaque part est un point de la
  //   série 0) ou par série (autres familles). Libellés de repli « Élément n » / « Série n » quand le
  //   fichier ne donne pas de nom. c:legendEntry (entrée masquée) n'est pas pris en compte.
  let legend = [];
  if (hasLegend && sersAll.length) {
    if (isPie) { const s0 = sersAll[0]; const cats = s0.cats.length ? s0.cats : s0.vals.map((_, i) => 'Élément ' + (i + 1)); legend = cats.map((c, i) => ({ name: c, col: paint(i, s0, i) })); }
    else legend = sersAll.map((s, i) => ({ name: s.name || 'Série ' + (i + 1), col: paint(i, s, 0) }));
  }
  // réservation d'espace : titre en haut, légende à droite (r/l) ou en bas (t/b)
  // ── Réservation d'espace (points) ───────────────────────────────────────────────────────────
  //   titleH    : 18 pt pour un titre de 12 pt (ligne de base posée à y = 13 plus bas).
  //   legRight  : largeur de la colonne de légende = 20 pt (pastille + marges) + 5,5 pt par caractère
  //               du libellé le plus long (chasse moyenne ESTIMÉE pour 9 pt), plafonnée à 140 pt.
  //               'l' (gauche) et 'tr' (haut-droite) sont aussi rangés à DROITE : simplification.
  //   legBottom : 16 pt pour une ligne de légende ; 't' (haut) est aussi rangé EN BAS : simplification.
  //   ox/oy     : origine de la zone de tracé (décalée sous le titre) ; cw/ch : sa taille (≥ 10 pt).
  const titleH = title ? 18 : 0;
  const legRight = (hasLegend && legend.length && (legPos === 'r' || legPos === 'l' || legPos === 'tr')) ? Math.min(140, 20 + Math.max(...legend.map((e) => e.name.length)) * 5.5) : 0;
  const legBottom = (hasLegend && legend.length && (legPos === 't' || legPos === 'b')) ? 16 : 0;
  const ox = 0, oy = titleH;
  const cw = Math.max(10, w - legRight), ch = Math.max(10, h - titleH - legBottom);

  // ══════════════════════════════════════════════════════════════════════════════════════════
  // plotSVG(w, h) : dessine la ZONE DE TRACÉ dans 0..w × 0..h (w/h masquent ici ceux de renderChart :
  //   ce sont cw/ch, l'espace restant après titre et légende). Renvoie le SVG ou '' si rien à tracer.
  //   Chaque famille est une branche qui fait « return » : la première qui correspond gagne.
  //   Repère SVG : y croît vers le BAS -> les valeurs positives vont vers les petits y.
  // ══════════════════════════════════════════════════════════════════════════════════════════
  const plotSVG = (w, h) => {
  let out = '';

  // ── Branche 1 : CAMEMBERT / ANNEAU (c:pieChart, c:doughnutChart, c:pie3DChart) ──────────────
  //   Seule la série 0 est dessinée (un camembert PowerPoint n'en montre qu'une ; un anneau à
  //   plusieurs séries, qui ferait des anneaux concentriques, est donc réduit au premier).
  //   Les valeurs négatives comptent en valeur absolue. Pas de c:firstSliceAng (rotation de départ)
  //   ni de c:explosion (part détachée) : la 1re part commence toujours à 12 h, sens horaire.
  if (isPie) {
    const chart = isPie;
    const sers = arr(chart['c:ser']).map((s) => serOf(s, color));
    const s0 = sers[0]; if (!s0 || !s0.vals.length) return '';
    const total = s0.vals.reduce((a, b) => a + Math.abs(b), 0) || 1; // || 1 : évite /0 si tout vaut 0
    const cx = w / 2, cy = h / 2, R = Math.min(w, h) / 2 * 0.92;      // rayon = 92 % du demi-côté (petite marge)
    // trou de l'anneau : c:holeSize en POURCENT du rayon (défaut 50), borné ici à [10 %, 85 %]
    const inner = plot['c:doughnutChart'] ? R * (Math.min(0.85, Math.max(0.1, (+plot['c:doughnutChart']['c:holeSize']?.['@val'] || 50) / 100))) : 0;
    const dl = dlblFlags(chart);
    let a0 = -Math.PI / 2; // 12 h
    s0.vals.forEach((v, i) => {
      // part i : de l'angle a0 à a1 (radians ; sens horaire à l'écran puisque y descend)
      const frac = Math.abs(v) / total; const a1 = a0 + frac * 2 * Math.PI;
      const x0 = cx + R * Math.cos(a0), y0 = cy + R * Math.sin(a0);
      const x1 = cx + R * Math.cos(a1), y1 = cy + R * Math.sin(a1);
      const large = frac > 0.5 ? 1 : 0;                   // drapeau large-arc SVG : arc > 180°
      const col = paint(i, s0, i);                        // une couleur par PART (rang de catégorie)
      // Limite : une part à 100 % (série à une seule valeur non nulle) donne un arc de départ = arrivée,
      // que SVG ne dessine pas (rien d'affiché pour ce cas).
      if (inner > 0) {
        // anneau : arc extérieur (sens horaire, sweep=1) puis arc intérieur en sens inverse (sweep=0)
        const ix0 = cx + inner * Math.cos(a0), iy0 = cy + inner * Math.sin(a0);
        const ix1 = cx + inner * Math.cos(a1), iy1 = cy + inner * Math.sin(a1);
        out += `<path d="M${x0.toFixed(1)} ${y0.toFixed(1)} A${R.toFixed(1)} ${R.toFixed(1)} 0 ${large} 1 ${x1.toFixed(1)} ${y1.toFixed(1)} L${ix1.toFixed(1)} ${iy1.toFixed(1)} A${inner.toFixed(1)} ${inner.toFixed(1)} 0 ${large} 0 ${ix0.toFixed(1)} ${iy0.toFixed(1)} Z" fill="${col}" stroke="#fff" stroke-width="1.5"/>`;
      } else {
        // part pleine : centre -> bord au départ -> arc -> fermeture vers le centre ; liseré blanc 1,5 pt
        out += `<path d="M${cx.toFixed(1)} ${cy.toFixed(1)} L${x0.toFixed(1)} ${y0.toFixed(1)} A${R.toFixed(1)} ${R.toFixed(1)} 0 ${large} 1 ${x1.toFixed(1)} ${y1.toFixed(1)} Z" fill="${col}" stroke="#fff" stroke-width="1.5"/>`;
      }
      if ((dl.pct || dl.val) && frac > 0.03) { // étiquette au milieu de la part
        // (parts < 3 % sans étiquette : le texte déborderait de la part)
        // am = angle médian de la part (a0 + moitié de l'ouverture frac·2π) ; rl = rayon de pose :
        //   milieu de l'épaisseur pour un anneau, 62 % du rayon pour un camembert plein.
        const am = a0 + frac * Math.PI, rl = inner > 0 ? (R + inner) / 2 : R * 0.62;
        const lx = cx + rl * Math.cos(am), ly = cy + rl * Math.sin(am);
        // pourcentage prioritaire sur la valeur si les deux sont demandés (PowerPoint afficherait les deux)
        const txt = dl.pct ? Math.round(frac * 100) + '%' : String(fmtNum(v));
        // ly + 3 : recentrage vertical approximatif d'un texte de 9 pt (y = ligne de base en SVG)
        out += `<text x="${lx.toFixed(1)}" y="${(ly + 3).toFixed(1)}" font-size="9" text-anchor="middle" fill="#fff" font-weight="bold">${esc(txt)}</text>`;
      }
      a0 = a1;                                            // la part suivante démarre où finit celle-ci
    });
    return out;
  }

  // ── Branche 2 : NUAGE DE POINTS / BULLES ────────────────────────────────────────────────────
  // nuage de points / bulles (c:scatterChart / c:bubbleChart) : axes X ET Y numériques
  //   Ici pas de catégories : chaque point a un X (c:xVal) et un Y (c:yVal). L'échelle couvre
  //   exactement [min, max] des données (sans arrondi « joli » ni inclusion de 0, contrairement aux
  //   barres). Les bulles sont dessinées comme de simples points (c:bubbleSize ignoré) et les
  //   lignes de liaison d'un scatterStyle « lineMarker » ne sont pas tracées.
  //   Note : num() convertit les trous en 0, donc le test « x == null » ne filtre que les indices
  //   au-delà de la longueur de ys (undefined).
  const scat = plot['c:scatterChart'] || plot['c:bubbleChart'];
  if (scat) {
    // séries simplifiées : pas de serOf() car les données sont dans c:xVal/c:yVal et non c:cat/c:val
    const sers = arr(scat['c:ser']).map((s) => ({
      name: ptsOf(s['c:tx']?.['c:strRef'])[0] ?? '', fill: s['c:spPr'] ? color(s['c:spPr']) : null,
      xs: ptsOf(s['c:xVal']?.['c:numRef']).map(num), ys: ptsOf(s['c:yVal']?.['c:numRef']).map(num),
    }));
    if (!sers.length) return '';
    // domaine des données sur toutes les séries (bornes du repère)
    let xmn = Infinity, xmx = -Infinity, ymn = Infinity, ymx = -Infinity;
    sers.forEach((s) => s.xs.forEach((x, i) => { const y = s.ys[i]; if (x == null || y == null) return; xmn = Math.min(xmn, x); xmx = Math.max(xmx, x); ymn = Math.min(ymn, y); ymx = Math.max(ymx, y); }));
    if (!Number.isFinite(xmn)) return '';                 // aucun point valide
    if (xmn === xmx) xmx = xmn + 1; if (ymn === ymx) ymx = ymn + 1; // domaine nul -> évite une division par 0
    // marges (pt) : gauche 34 pour les graduations Y, bas 20 pour les graduations X ; pw/ph = zone utile
    const mL = 34, mB = 20, mT = 8, mR = 8, pw = Math.max(1, w - mL - mR), ph = Math.max(1, h - mT - mB);
    // PX/PY : données -> coordonnées SVG (interpolation linéaire ; PY inversé car y descend)
    const PX = (x) => mL + pw * (x - xmn) / (xmx - xmn), PY = (y) => mT + ph * (1 - (y - ymn) / (ymx - ymn));
    // 5 graduations horizontales (t = 0..4, de ymx en haut à ymn en bas) + libellés à gauche
    for (let t = 0; t <= 4; t++) { const y = mT + ph * t / 4, v = ymx - (ymx - ymn) * t / 4;
      out += `<line x1="${mL}" y1="${y.toFixed(1)}" x2="${(w - mR).toFixed(1)}" y2="${y.toFixed(1)}" stroke="#e0e0e0" stroke-width="0.5"/><text x="${(mL - 3).toFixed(1)}" y="${(y + 3).toFixed(1)}" font-size="9" text-anchor="end" fill="#666">${esc(fmtNum(v))}</text>`; }
    // 5 libellés d'axe X (sans ligne de grille verticale)
    for (let t = 0; t <= 4; t++) { const x = mL + pw * t / 4, v = xmn + (xmx - xmn) * t / 4;
      out += `<text x="${x.toFixed(1)}" y="${(h - mB + 12).toFixed(1)}" font-size="9" text-anchor="middle" fill="#666">${esc(fmtNum(v))}</text>`; }
    // un disque de rayon 3 pt par point, couleur de la série (s.pts n'existe pas ici -> pas de c:dPt)
    sers.forEach((s, si) => { const col = paint(si, s, 0);
      for (let i = 0; i < s.xs.length; i++) { const x = s.xs[i], y = s.ys[i]; if (x == null || y == null) continue;
        out += `<circle cx="${PX(x).toFixed(1)}" cy="${PY(y).toFixed(1)}" r="3" fill="${col}"/>`; } });
    return out;
  }

  // ── Branche 3 : RADAR ───────────────────────────────────────────────────────────────────────
  // radar (c:radarChart) : catégories réparties en cercle, chaque série = polygone
  //   Échelle radiale : 0 au centre, vmax (plus grande valeur toutes séries confondues) au rayon R.
  //   Les valeurs négatives ne sont pas gérées (rayon négatif -> point de l'autre côté du centre).
  //   c:radarStyle (marker / filled) ignoré : toujours polygone semi-transparent (25 %) + contour.
  const radar = plot['c:radarChart'];
  if (radar) {
    const sers = arr(radar['c:ser']).map((s) => serOf(s, color));
    if (!sers.length) return '';
    const cats = sers[0].cats.length ? sers[0].cats : sers[0].vals.map((_, i) => String(i + 1)); // sans libellés : 1, 2, 3…
    const n = Math.max(3, cats.length);                  // au moins 3 axes pour former un polygone
    let vmax = 0; sers.forEach((s) => s.vals.forEach((v) => { vmax = Math.max(vmax, v); })); if (!vmax) vmax = 1;
    // cy + 4 : léger décalage vers le bas pour laisser la place au libellé de l'axe du haut ;
    // R = 78 % du demi-côté pour garder de la marge aux libellés posés à R + 10.
    const cx = w / 2, cy = h / 2 + 4, R = Math.min(w, h) / 2 * 0.78;
    const ang = (i) => -Math.PI / 2 + 2 * Math.PI * i / n;   // catégorie i, départ en haut
    const PT = (i, r) => [cx + r * Math.cos(ang(i)), cy + r * Math.sin(ang(i))]; // polaire -> cartésien
    // grille : 4 polygones concentriques à 25/50/75/100 % de R, puis les rayons et les libellés
    for (let g = 1; g <= 4; g++) { const rr = R * g / 4; // anneaux concentriques
      let d = ''; for (let i = 0; i < n; i++) { const [x, y] = PT(i, rr); d += (i ? 'L' : 'M') + x.toFixed(1) + ' ' + y.toFixed(1); }
      out += `<path d="${d}Z" fill="none" stroke="#e0e0e0" stroke-width="0.5"/>`; }
    for (let i = 0; i < n; i++) { const [x, y] = PT(i, R); out += `<line x1="${cx.toFixed(1)}" y1="${cy.toFixed(1)}" x2="${x.toFixed(1)}" y2="${y.toFixed(1)}" stroke="#e0e0e0" stroke-width="0.5"/>`;
      const [lx, ly] = PT(i, R + 10); out += `<text x="${lx.toFixed(1)}" y="${(ly + 3).toFixed(1)}" font-size="8" text-anchor="middle" fill="#666">${esc(cats[i] ?? '')}</text>`; }
    // un polygone par série : sommet i à la distance R · valeur / vmax sur l'axe de la catégorie i
    sers.forEach((s, si) => { const col = paint(si, s, 0); let d = '';
      for (let i = 0; i < n; i++) { const [x, y] = PT(i, R * (s.vals[i] || 0) / vmax); d += (i ? 'L' : 'M') + x.toFixed(1) + ' ' + y.toFixed(1); }
      out += `<path d="${d}Z" fill="${col}" fill-opacity="0.25" stroke="${col}" stroke-width="2"/>`; });
    return out;
  }

  // ── Branches 4 à 7 : familles CARTÉSIENNES (catégories × valeurs) ────────────────────────────
  // colonnes / barres / lignes / aires : axes cartésiens
  //   Préparation commune (séries, catégories, empilement, domaine de valeurs), puis :
  //     4. barres horizontales (barDir="bar") — repère tourné de 90° ;
  //     5. colonnes (barDir="col") ; 6. aires ; 7. lignes.
  const chart = barC || isLine || isArea;
  if (!chart) return '';                                  // famille non gérée (stock, surface…) : rien
  const sers = arr(chart['c:ser']).map((s) => serOf(s, color));
  if (!sers.length) return '';
  const dl = dlblFlags(chart);
  // libellés pris sur la série 0 (toutes les séries partagent les catégories dans PowerPoint)
  const cats = sers[0].cats.length ? sers[0].cats : sers[0].vals.map((_, i) => String(i + 1));
  // nombre de créneaux : le plus grand entre libellés et longueurs de séries (au moins 1 -> pas de /0)
  const nCat = Math.max(1, cats.length, ...sers.map((s) => s.vals.length));
  // c:grouping : "stacked" et "percentStacked" -> empilé (le pourcentage n'est PAS normalisé à 100 %)
  const stacked = /stacked/i.test(chart['c:grouping']?.['@val'] || '');
  const horiz = barC && (barC['c:barDir']?.['@val'] === 'bar'); // "col" (défaut) = colonnes verticales
  // domaine des valeurs (avec 0 et empilement éventuel)
  //   Le domaine inclut TOUJOURS 0 (vmin ≤ 0 ≤ vmax au départ) : les barres partent de l'axe zéro.
  //   Empilé : on somme séparément les positifs et les négatifs de chaque catégorie (comme PowerPoint,
  //   qui empile les négatifs sous l'axe) et on garde les extrêmes de ces sommes.
  //   Pas d'arrondi « joli » des bornes (PowerPoint prendrait 0..60 là où on prend 0..57).
  let vmax = 0, vmin = 0;
  if (stacked) for (let c = 0; c < nCat; c++) { let pos = 0, neg = 0; sers.forEach((s) => { const v = s.vals[c] || 0; if (v >= 0) pos += v; else neg += v; }); vmax = Math.max(vmax, pos); vmin = Math.min(vmin, neg); }
  else sers.forEach((s) => s.vals.forEach((v) => { vmax = Math.max(vmax, v); vmin = Math.min(vmin, v); }));
  // bornes d'axe EXPLICITES (c:valAx>c:scaling>c:min/c:max) prioritaires sur l'auto-échelle
  //   Seul le PREMIER c:valAx est lu (un axe secondaire éventuel est ignoré). c:scaling>c:orientation
  //   (axe inversé maxMin) et c:logBase (échelle log) ne sont pas pris en compte.
  const valAx0 = arr(plot['c:valAx'])[0];
  const scl = valAx0?.['c:scaling'];
  if (scl?.['c:max']?.['@val'] != null) vmax = num(scl['c:max']['@val']);
  if (scl?.['c:min']?.['@val'] != null) vmin = num(scl['c:min']['@val']);
  if (vmax === vmin) vmax = vmin + 1;                     // domaine nul (tout à 0) -> évite /0

  // ── Branche 4 : BARRES HORIZONTALES ─────────────────────────────────────────────────────────
  // BARRES HORIZONTALES (barDir="bar") : axe des catégories vertical, axe des valeurs horizontal
  //   Catégorie 0 en HAUT (CY croissant). Attention : PowerPoint, par défaut, place la catégorie 0
  //   en BAS pour un graphique à barres (axe des catégories orienté minMax de bas en haut) ; l'ordre
  //   vertical est donc inversé par rapport à PowerPoint sauf si l'axe y est « en ordre inverse ».
  if (horiz) {
    // marges (pt) : gauche 62 pour les libellés de catégories (plus longs que des nombres)
    const mLh = 62, mBh = 20, mTh = 8, mRh = 12;
    const pwh = Math.max(1, w - mLh - mRh), phh = Math.max(1, h - mTh - mBh);
    const VX = (v) => mLh + pwh * (v - vmin) / (vmax - vmin);  // valeur -> x
    const CY = (i) => mTh + phh * (i + 0.5) / nCat;            // catégorie i -> y (centre)
    // grille verticale + libellés de valeurs sous la zone (5 graduations régulières de vmin à vmax)
    for (let t = 0; t <= 4; t++) { const v = vmin + (vmax - vmin) * t / 4, x = VX(v);
      out += `<line x1="${x.toFixed(1)}" y1="${mTh}" x2="${x.toFixed(1)}" y2="${(h - mBh).toFixed(1)}" stroke="#d0d0d0" stroke-width="0.5"/><text x="${x.toFixed(1)}" y="${(h - mBh + 12).toFixed(1)}" font-size="9" text-anchor="middle" fill="#666">${esc(fmtNum(v))}</text>`; }
    // libellés de catégories, alignés à droite contre la zone de tracé
    cats.forEach((c, i) => { if (i < nCat) out += `<text x="${(mLh - 4).toFixed(1)}" y="${(CY(i) + 3).toFixed(1)}" font-size="9" text-anchor="end" fill="#666">${esc(c)}</text>`; });
    // groupH : 80 % du créneau de catégorie (les 20 % restants = espace entre groupes ; c:gapWidth,
    //   qui règle cet écart dans PowerPoint — défaut 150 % d'une barre —, n'est pas lu).
    // bh : épaisseur d'une barre = tout le groupe si empilé, sinon groupe partagé entre les séries.
    const groupH = phh / nCat * 0.8, bh = stacked ? groupH : groupH / sers.length;
    const vary = (barC['c:varyColors']?.['@val'] === '1') && sers.length === 1;
    for (let c = 0; c < nCat; c++) {
      let posR = VX(0), negR = VX(0);                     // extrémités courantes des piles positive (vers la droite) / négative (vers la gauche)
      sers.forEach((s, si) => {
        const v = s.vals[c] || 0; if (!v) return; const col = paint(vary ? c : si, s, c); // valeur nulle : pas de barre
        const wdt = Math.abs(VX(v) - VX(0)); let x, y;   // longueur de la barre en pt
        if (stacked) { y = CY(c) - groupH / 2; if (v >= 0) { x = posR; posR += wdt; } else { negR -= wdt; x = negR; } }
        else { y = CY(c) - groupH / 2 + si * bh; x = Math.min(VX(0), VX(v)); } // groupé : séries côte à côte, 0 en haut
        const bhEff = stacked ? bh : bh * 0.92;           // groupé : 8 % de jour entre barres voisines
        out += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${wdt.toFixed(1)}" height="${bhEff.toFixed(1)}" fill="${col}"/>`;
        // étiquette de valeur à droite de la barre (groupé seulement ; pour une valeur négative elle
        // tombe donc juste à droite de l'axe zéro, hors de la barre qui s'étend à gauche)
        if (dl.val && !stacked) out += `<text x="${(x + wdt + 3).toFixed(1)}" y="${(y + bhEff / 2 + 3).toFixed(1)}" font-size="8" fill="#555">${esc(fmtNum(v))}</text>`;
      });
    }
    // axe zéro vertical (trait plus foncé), dessiné après les barres pour rester visible
    out += `<line x1="${VX(0).toFixed(1)}" y1="${mTh}" x2="${VX(0).toFixed(1)}" y2="${(h - mBh).toFixed(1)}" stroke="#888" stroke-width="1"/>`;
    return out;
  }

  // ── Repère vertical commun aux branches 5 (colonnes), 6 (aires) et 7 (lignes) ───────────────
  //   Axe des catégories horizontal : chaque catégorie occupe un créneau de pw / nCat, repérée par
  //   son CENTRE (X(i)). Axe des valeurs vertical, linéaire entre vmin (bas) et vmax (haut).
  const mL = 34, mB = 22, mT = 8, mR = 8; // marges pour échelle + labels
  const pw = Math.max(1, w - mL - mR), ph = Math.max(1, h - mT - mB);
  const X = (i) => mL + (pw * (i + 0.5)) / nCat;             // centre de la catégorie i
  const Y = (v) => mT + ph * (1 - (v - vmin) / (vmax - vmin)); // valeur -> y
  // grille + axe des valeurs (5 graduations)
  //   ticks = 4 intervalles -> 5 lignes (t = 0..4). Espacement régulier du domaine, sans chercher un
  //   pas « rond » comme le fait c:majorUnit automatique de PowerPoint.
  const ticks = 4;
  for (let t = 0; t <= ticks; t++) {
    const v = vmin + (vmax - vmin) * t / ticks; const y = Y(v);
    out += `<line x1="${mL}" y1="${y.toFixed(1)}" x2="${(w - mR).toFixed(1)}" y2="${y.toFixed(1)}" stroke="#d0d0d0" stroke-width="0.5"/>`;
    out += `<text x="${(mL - 3).toFixed(1)}" y="${(y + 3).toFixed(1)}" font-size="9" text-anchor="end" fill="#666">${esc(Math.round(v * 100) / 100)}</text>`;
  }
  // axe des catégories
  //   Libellés centrés sous chaque créneau, sans rotation ni retour à la ligne : des libellés longs
  //   ou nombreux peuvent se chevaucher (PowerPoint les inclinerait ou en sauterait).
  cats.forEach((c, i) => { if (i < nCat) out += `<text x="${X(i).toFixed(1)}" y="${(h - mB + 12).toFixed(1)}" font-size="9" text-anchor="middle" fill="#666">${esc(c)}</text>`; });

  // ── Branche 5 : COLONNES (barDir="col"), groupées ou empilées ───────────────────────────────
  //   Même logique que les barres horizontales, axes permutés. Empilé : les segments positifs montent
  //   depuis l'axe zéro (posTop décroît en y), les négatifs descendent (negTop croît en y).
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
        const bwEff = stacked ? bw : bw * 0.92;           // groupé : 8 % de jour entre colonnes voisines
        out += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bwEff.toFixed(1)}" height="${hgt.toFixed(1)}" fill="${col}"/>`;
        // étiquette : au centre du segment en blanc si empilé, au-dessus de la colonne en gris sinon
        // (au-dessus même pour une valeur négative, donc côté axe zéro)
        if (dl.val) out += `<text x="${(x + bwEff / 2).toFixed(1)}" y="${(stacked ? y + hgt / 2 + 3 : y - 3).toFixed(1)}" font-size="8" text-anchor="middle" fill="${stacked ? '#fff' : '#555'}">${esc(fmtNum(v))}</text>`;
      });
    }
  } else if (isArea) {
    // ── Branche 6 : AIRES ─────────────────────────────────────────────────────────────────────
    //   Chaque série = polygone fermé : base (axe zéro) sous la 1re catégorie -> sommets des valeurs
    //   -> base sous la dernière catégorie. Opacité 60 % pour laisser voir les séries superposées.
    //   Pas d'empilement réel : même en grouping "stacked", les aires sont superposées depuis zéro
    //   (le domaine vmin/vmax, lui, a bien été calculé en empilé). Les aires démarrent au CENTRE
    //   du 1er créneau (PowerPoint les fait partir du bord quand l'axe « croise entre les graduations »
    //   est désactivé — non géré).
    sers.forEach((s, si) => {
      const col = paint(si, s, 0);
      let d = `M${X(0).toFixed(1)} ${Y(0).toFixed(1)}`;
      for (let c = 0; c < nCat; c++) d += ` L${X(c).toFixed(1)} ${Y(s.vals[c] || 0).toFixed(1)}`;
      d += ` L${X(nCat - 1).toFixed(1)} ${Y(0).toFixed(1)} Z`;
      out += `<path d="${d}" fill="${col}" fill-opacity="0.6" stroke="${col}" stroke-width="1"/>`;
    });
  } else { // lignes
    // ── Branche 7 : LIGNES ────────────────────────────────────────────────────────────────────
    //   Polyligne 2 pt + marqueur rond 2,5 pt à chaque point (c:marker ignoré : toujours un rond).
    //   « v == null » saute les indices au-delà de la longueur de la série (les trous internes sont
    //   déjà devenus 0 dans serOf, donc la courbe ne s'interrompt pas). Lignes empilées : non
    //   cumulées (chaque série est tracée à sa propre valeur). c:smooth (lissage) ignoré.
    sers.forEach((s, si) => {
      const col = paint(si, s, 0);
      let d = '';
      for (let c = 0; c < nCat; c++) { const v = s.vals[c]; if (v == null) continue; d += (d ? ' L' : 'M') + X(c).toFixed(1) + ' ' + Y(v).toFixed(1); }
      out += `<path d="${d}" fill="none" stroke="${col}" stroke-width="2"/>`;
      for (let c = 0; c < nCat; c++) { const v = s.vals[c]; if (v == null) continue; out += `<circle cx="${X(c).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="2.5" fill="${col}"/>`; if (dl.val) out += `<text x="${X(c).toFixed(1)}" y="${(Y(v) - 5).toFixed(1)}" font-size="8" text-anchor="middle" fill="#555">${esc(fmtNum(v))}</text>`; }
    });
  }
  // axe zéro
  //   Tracé en dernier (par-dessus barres et aires). Si 0 est hors du domaine imposé par c:min/c:max,
  //   la ligne sort de la zone de tracé (aucun écrêtage).
  out += `<line x1="${mL}" y1="${Y(0).toFixed(1)}" x2="${(w - mR).toFixed(1)}" y2="${Y(0).toFixed(1)}" stroke="#888" stroke-width="1"/>`;
  return out;
  }; // fin plotSVG

  // ════════════════════════════════════════════════════════════════════════════════════════════
  // Assemblage final : zone de tracé (dans cw × ch, décalée de oy sous le titre) + titre + légende.
  //   Si plotSVG ne produit rien, on renvoie '' même s'il y a un titre : pas de graphique « vide ».
  // ════════════════════════════════════════════════════════════════════════════════════════════
  const body = plotSVG(cw, ch);
  if (!body) return '';
  let svg = `<g transform="translate(${ox} ${oy})">${body}</g>`;
  // titre : 12 pt gras, centré sur TOUTE la largeur w (légende comprise), ligne de base à y = 13
  if (title) svg += `<text x="${(w / 2).toFixed(1)}" y="13" font-size="12" font-weight="bold" text-anchor="middle" fill="#333">${esc(title)}</text>`;
  if (legend.length) {
    if (legRight) { // colonne à droite
      // colonne centrée verticalement sur la zone de tracé ; une entrée tous les 14 pt ;
      // pastille 9×9 dont le bas est aligné sur la ligne de base du libellé (y - 8 .. y + 1)
      const x = w - legRight + 6; let y = oy + Math.max(6, (ch - legend.length * 14) / 2) + 8;
      for (const e of legend) { svg += `<rect x="${x}" y="${(y - 8).toFixed(1)}" width="9" height="9" fill="${e.col}"/><text x="${x + 13}" y="${y.toFixed(1)}" font-size="9" fill="#444">${esc(e.name)}</text>`; y += 14; }
    } else if (legBottom) { // ligne en bas, centrée
      // largeur ESTIMÉE de chaque entrée = pastille (9) + 4 + 5,2 pt par caractère + écart (14) ;
      // la ligne est centrée sur w (pas de retour à la ligne si elle déborde : x mini = 4)
      const gap = 14, sw = 9; const widths = legend.map((e) => sw + 4 + e.name.length * 5.2 + gap);
      let x = Math.max(4, (w - widths.reduce((a, b) => a + b, 0)) / 2); const y = h - 4;
      for (let i = 0; i < legend.length; i++) { const e = legend[i]; svg += `<rect x="${x.toFixed(1)}" y="${(y - 8).toFixed(1)}" width="${sw}" height="9" fill="${e.col}"/><text x="${(x + sw + 3).toFixed(1)}" y="${y.toFixed(1)}" font-size="9" fill="#444">${esc(e.name)}</text>`; x += widths[i]; }
    }
  }
  return svg;
}

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 4 — Exports pour les tests
// ════════════════════════════════════════════════════════════════════════════════════════════════
// --- exports nommés pour les TESTS (fonctions PURES uniquement) ---
// Helpers sans effet de bord (pas de fs/réseau/DOM/état global). renderChart est déjà exporté.
// (Le LMS les avait retirés en vendorisant le fichier ; ils sont rétablis ici pour test/docx-chart.test.js.)
export { esc, arr, num, ptsOf, serOf };
