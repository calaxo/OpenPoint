// ═══ [CŒUR PPTX] ═══ carte d'ensemble : docs/ARCHITECTURE.md (§4) · consommé par build.mjs + build-player.mjs
// Moteur de rendu PPTX -> SVG, en JS pur, SANS PowerPoint.
// Rend une diapo (formes vectorielles + texte + images) dans le repère EMU de la diapo.
// Le SVG en viewBox EMU = fidélité géométrique exacte (coins arrondis, custGeom, rotations).
//
// Couverture v1 : spTree ordonné (z-order), groupes + transforms, prstGeom courants + custGeom,
// fills solides/dégradés/image, lignes, texte multi-run via foreignObject, héritage thème
// (clrScheme + fmtScheme) et placeholders (géométrie + styles via layout/master).
//
// API : renderSlide(pptDir, slideNo, {embedImages}) -> { svg, w, h }

/* ═══════════════════════════════════════════════════════════════════════════════════════════════
 * GUIDE DE LECTURE (notes ajoutées pour qui découvre le fichier — aucune ne change le rendu)
 * ═══════════════════════════════════════════════════════════════════════════════════════════════
 *
 * 1. CE QUE FAIT CE FICHIER
 *    Il prend un .pptx DÉJÀ DÉCOMPRESSÉ (dossier `ppt/` : slides/, slideLayouts/, slideMasters/,
 *    theme/, media/, charts/, diagrams/, presentation.xml, tableStyles.xml) et produit, pour UNE
 *    diapo : une chaîne SVG autonome (renderSlide) + sa timeline d'animation (extractTiming).
 *    Aucun PowerPoint ni LibreOffice n'intervient : tout est relu depuis le XML OOXML
 *    (ECMA-376 : PresentationML §19, DrawingML §20, DrawingML-texte §21, Office Math §22).
 *    Consommateurs : build.mjs (diapos statiques) et build-player.mjs (lecteur animé), plus tout runtime
 *    d'animation hôte qui consomme le même SVG.
 *
 * 2. PIPELINE DE renderSlide(pptDir, slideNo)
 *    a) Résolution de la chaîne d'héritage par les fichiers .rels :
 *         slideN.xml --rels--> slideLayoutX.xml --rels--> slideMasterY.xml --rels--> themeZ.xml
 *       (un deck multi-master a PLUSIEURS thèmes : le bon est celui du master de la diapo).
 *    b) loadContext(thème) : palette clrScheme, polices majeure/mineure, listes de styles fmtScheme
 *       (fillStyleLst, lnStyleLst, effectStyleLst, bgFillStyleLst), styles de tableaux. Mis en cache.
 *    c) Lecture du master : p:clrMap (bg1/tx1/... -> dk1/lt1/...) et p:txStyles (titre/corps/autre,
 *       niveaux 1..9) ; puis presentation.xml : p:defaultTextStyle et la taille de diapo p:sldSz.
 *    d) placeholderGeom : index des placeholders du LAYOUT (clé "type:idx") -> xfrm, spPr, style,
 *       lstStyle, bodyPr dont héritent les formes <p:ph> de la diapo.
 *    e) Composition (en fin de renderSlide, après toutes les définitions internes) :
 *         UN fond (diapo > layout > master)  puis  formes du master (préfixe spid 'm', placeholders
 *         sautés)  puis  formes du layout ('l', idem)  puis  formes de la diapo (spid brut).
 *       Le parcours walk()/walkChild() suit l'ordre du spTree = ordre z (le 1er est dessous).
 *    f) Chaque forme : renderSp (p:sp / p:cxnSp), renderPic (p:pic), renderGroup (p:grpSp),
 *       renderTable (p:graphicFrame : tableau, graphique -> chart-render.mjs, SmartArt -> dsp:drawing,
 *       ou image de secours d'un objet OLE). Géométrie -> preset-geom.mjs ; remplissage/trait/effets
 *       -> <defs> (dégradés, motifs, filtres, clips, marqueurs) accumulés pendant le parcours.
 *    g) Texte : renderText produit un <foreignObject> contenant du HTML (un <div data-par> par
 *       paragraphe, un <span> par run). C'est le NAVIGATEUR qui fait la mise en page ; le code
 *       corrige ses écarts avec PowerPoint (interligne 1,2, ligne de base, chasses au 1/8 pt,
 *       tabulations, coupure de ligne) grâce aux métriques de polices d'animFonts.js.
 *    h) Assemblage : <svg viewBox="0 0 Wpt Hpt"><defs>…</defs><rect blanc/>…formes…</svg>.
 *
 * 3. PIPELINE DE extractTiming(pptDir, slideNo)
 *    Extrait <p:timing> de la diapo (regex), prend la 1re séquence <p:seq>, et considère chaque
 *    <p:par> enfant direct de sa childTnLst comme une ÉTAPE (un clic). Dans chaque étape, tout
 *    <p:cTn presetID=…> est un EFFET : cible (spid, plage de paragraphes), classe (entr/exit/emph…),
 *    filtre animEffect, mouvement/rotation, délai cumulé, durée. -> { steps, animatedSpids }.
 *
 * 4. PLAN DU FICHIER (numéros de ligne APPROXIMATIFS, après ajout de ces notes)
 *    §A   imports, parseurs XML .......................................... ~l. 140
 *    §B   cache de parse des fichiers stables ............................ ~l. 165
 *    §C   helpers d'arbre preserveOrder / objet plat, esc() .............. ~l. 195
 *    §D   Office Math (OMML) -> HTML ..................................... ~l. 250
 *    §E   puces, polices -> familles CSS ................................. ~l. 360
 *    §F   ligne de base, chasses 1/8 pt, largeur de texte PowerPoint ..... ~l. 435
 *    §G   hauteur de texte d'une cellule de tableau ...................... ~l. 530
 *    §H   polices symboles (Wingdings / Symbol) -> Unicode ............... ~l. 580
 *    §I   style de tableau intégré, unités EMU -> points ................. ~l. 665
 *    §J   contexte document (thème) ...................................... ~l. 685
 *    §K   couleurs DrawingML ............................................. ~l. 770
 *    §L   géométries preset / custGeom ................................... ~l. 840
 *    §M   relations (.rels) et images .................................... ~l. 900
 *    §N   placeholders du layout ......................................... ~l. 955
 *    §O   renderSlide : prélude (héritage, styles master) ................ ~l. 1000
 *    §O.1   fond, parcours, transformations, groupes ..................... ~l. 1115
 *    §O.2   remplissages, dégradés, motifs, pointillés, marqueurs ........ ~l. 1255
 *    §O.3   effets (ombre/halo/bords adoucis), filtres d'image, paint .... ~l. 1360
 *    §O.4   renderSp (formes) ............................................ ~l. 1510
 *    §O.5   graphiques, SmartArt, tableaux ............................... ~l. 1655
 *    §O.6   renderPic (images) ........................................... ~l. 1900
 *    §O.7   renderText (texte) ........................................... ~l. 1990
 *    §O.8   composition finale et assemblage SVG ......................... ~l. 2485
 *    §P   extractTiming (animations) ..................................... ~l. 2555
 *    §Q   serialize(), exports de test ................................... ~l. 2675
 *    (Chercher « ████ §X » pour sauter à une section.)
 *
 * 5. CONVENTIONS
 *    - UNITÉS. OOXML mesure en EMU (English Metric Unit) : 914 400 EMU = 1 pouce, 12 700 EMU = 1 pt,
 *      360 000 EMU = 1 cm. Le SVG produit est en POINTS : u(v) = v / 12700 arrondi à 2 décimales,
 *      et viewBox = "0 0 W/12700 H/12700". NB : l'en-tête d'origine ci-dessus parle de « repère
 *      EMU » ; en pratique tout passe par u() (cf. le commentaire de S/u : des font-size en
 *      centaines de milliers de px cassaient les foreignObject). Dans le HTML des foreignObject,
 *      « px » CSS = 1 unité du viewBox = 1 pt de diapo.
 *      Autres unités OOXML rencontrées : angles en 1/60 000 de degré (ST_Angle, d'où les /60000),
 *      pourcentages en 1/1000 de % (ST_Percentage : 100000 = 100 %, d'où les /100000), tailles de
 *      police en 1/100 de pt (@sz="1800" = 18 pt), spcPts et @spc en 1/100 de pt, positions de
 *      stops de dégradé en 1/1000 de % (/1000 -> % SVG).
 *    - DEUX FORMES D'ARBRE XML (fast-xml-parser) :
 *        * P  (preserveOrder:true) -> tableau de nœuds { 'a:tag': [enfants…], ':@': {attributs} },
 *          le texte étant { '#text': '…' }. Garde l'ORDRE des enfants (indispensable : z-order du
 *          spTree, alternance a:r / a:br / a:fld / formules dans un paragraphe, commandes custGeom).
 *          Lu avec les helpers tagOf/at/kids/byTag/first/deepFind.
 *        * PF (objet plat) -> { 'a:tag': { '@attr': …, 'a:enfant': {…} | [{…},{…}] } }. Beaucoup plus
 *          simple à lire (spPr['a:solidFill']?.['a:schemeClr']?.['@val']), mais PERD l'ordre entre
 *          balises de noms différents, et un enfant unique n'est pas un tableau (d'où les
 *          `Array.isArray(x) ? x : [x]` / arr() partout).
 *      Stratégie : parcours de structure en ordonné, puis pour chaque forme on RE-SÉRIALISE le nœud
 *      ordonné (serialize) et on le re-parse en plat (`PF.parse('<root …>' + serialize(n) + '</root>')`)
 *      pour lire propriétés et texte. Quand l'ordre compte (a:br, a:fld, OMML, tabulations), on
 *      revient au nœud ordonné correspondant (paramètre txBodyO, paraSeqO, paraO).
 *    - ATTRIBUTS LUS PAR LE RUNTIME D'ANIMATION (build-player.mjs, runtime hôte) :
 *        * data-spid="<id>" sur le <g> de chaque forme / groupe / tableau = p:cNvPr@id. Préfixé 'm'
 *          (master) ou 'l' (layout) pour éviter les collisions : la timeline ne cible que la diapo,
 *          donc le player fait querySelector('[data-spid="<spid brut>"]') et doit tomber sur la
 *          forme de la diapo. Les formes listées dans animatedSpids sont cachées au départ.
 *        * data-par="<n>" sur chaque <div> de paragraphe : cible des animations « par paragraphe »
 *          (bldP, p:pRg st..end dans la timeline).
 *        * class="af" data-af="norm|fixed" et data-wrap="none" sur le conteneur de texte : repères
 *          pour le player (mode d'ajustement de la boîte ; aucun rétrécissement côté navigateur).
 *    - IDENTIFIANTS SVG : tous préfixés par uid = 's<N>_' (clips, dégradés, motifs, filtres,
 *      marqueurs) pour que plusieurs diapos puissent cohabiter dans une même page (viewer.html).
 *
 * 6. PIÈGES CONNUS / À SAVOIR AVANT DE MODIFIER
 *    - Le rendu a été calé AU PIXEL sur des exports PowerPoint (via COM) : les constantes (1,2 ;
 *      0,75 L ; 1/8 pt ; 0,2 px ; tailles de pointes de flèche…) sont MESURÉES, pas théoriques. Les
 *      commentaires d'origine citent souvent la diapo de référence (« VENT d.10 », « 022_04 d.65 »…).
 *      Toute modification doit être revalidée sur le corpus (empreintes SVG / comparateur).
 *    - esc() échappe ' en &apos; (entité NOMMÉE) : fast-xml-parser décode &apos; au re-parse mais
 *      pas &#39; -> double échappement visible si on change ça (voir le commentaire de esc).
 *    - Les arbres renvoyés par cachedParse sont PARTAGÉS entre diapos : ne jamais les muter. Le code
 *      mute seulement des arbres plats fraîchement re-parsés (ex. tb['a:bodyPr'] dans renderTable).
 *    - État mutable : _ctxMinorFont (module) et ctx.clrMap / ctx.phClr / ctx.grpFill / ctx.txStyles
 *      (par appel) -> renderSlide n'est pas réentrant (pas d'appels concurrents entrelacés).
 *    - Dans renderSlide, les `const` (grads, patterns, clips, filters, lineMarkers) sont déclarés
 *      APRÈS certaines fonctions qui les utilisent : la composition est donc faite tout en bas, une
 *      fois tout initialisé (sinon zone morte temporelle / TDZ).
 *    - Les commentaires JS sont libres, mais JAMAIS dans une chaîne / un gabarit `…` : ils
 *      partiraient dans le SVG.
 *    - Non gérés (constat à la lecture du code) : p:sld@showMasterSp / cNvPr@hidden (les formes du
 *      master/layout sont toujours dessinées), @nodeType de la timeline (chaque p:par de la séquence
 *      = une étape de clic), custGeom HÉRITÉ du layout (repli rectangle), héritage de GÉOMÉTRIE de
 *      placeholder depuis le MASTER (seul le layout est indexé pour la géométrie — mais son lstStyle
 *      de texte, lui, est bien lu : maillon de la cascade avant p:txStyles).
 * ═══════════════════════════════════════════════════════════════════════════════════════════════ */

// ████ §A ─ IMPORTS ET PARSEURS XML ──────────────────────────────────────────────────────────────
// Dépendances :
//   - fast-xml-parser : seul parseur XML ; deux configurations (P ordonné, PF plat), voir le guide §5.
//   - preset-geom.mjs : moteur des 187 formes prédéfinies (prstGeom) + évaluateur de formules de
//     guides (custGeom). Renvoie des chemins SVG « d » et le rectangle de texte de chaque preset.
//   - chart-render.mjs : rendu SVG des graphiques (c:chart).
//   - animFonts.js : métriques des polices embarquées (ascendante/descendante via familyMetrics,
//     tables de chasses par glyphe via advanceTable) -> recalages ligne de base / largeur.
import { readFileSync, existsSync, statSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { XMLParser } from 'fast-xml-parser';
import { presetGeomD, evalGeom, presetGeomParts, evalGeomParts, presetTextRect } from './preset-geom.mjs';
import { renderChart } from './chart-render.mjs';
import { familyMetrics, advanceTable } from './animFonts.js';

// Options communes aux deux parseurs :
//   ignoreAttributes:false + attributeNamePrefix '@' -> les attributs sont lus comme n['@val'] ;
//   parseTagValue/parseAttributeValue:false -> TOUT reste chaîne (« 0100 » ne devient pas 100,
//   un hex « 1E10 » ne devient pas un nombre) : les conversions sont faites explicitement (+x).
// P  : preserveOrder -> tableau ordonné (structure, z-order, runs). Attributs sous la clé ':@'.
// PF : objet plat -> lecture directe des propriétés (thème, rels, spPr, txBody re-parsés).
// trimValues:false -> conserve les espaces en bord de <a:t> (sinon les mots se collent
// entre runs successifs, ex. "Marge " + "qu'on" -> "Margequ'on").
const P = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', preserveOrder: true, parseTagValue: false, parseAttributeValue: false, trimValues: false });
const PF = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', parseTagValue: false, parseAttributeValue: false, trimValues: false }); // flat, pour thème/rels

// ████ §B ─ CACHE DE PARSE ────────────────────────────────────────────────────────────────────────
// Rôle : ne parser qu'une fois les fichiers partagés par toutes les diapos d'un deck. Chaque entrée
// garde jusqu'à trois représentations paresseuses du même fichier : texte brut ('txt'), arbre plat
// ('flat', parseur PF) et arbre ordonné ('ord', parseur P).
// ---- cache de parse pour fichiers STABLES d'un deck (thème/master/layout/rels/presentation/tableStyles) ----
// renderSlide est appelé par diapo, mais N diapos partagent le même master/layout/thème : sans cache on
// reparse ~64 Ko de XML par diapo (master+layout parsés 2× chacun). Clé = chemin + mtime + taille : si un
// dossier temporaire est réutilisé avec un contenu DIFFÉRENT (ex. batch check.mjs), la signature change et
// le cache est invalidé -> jamais de rendu périmé. Les fichiers PROPRES à la diapo (slideN.xml, ses rels)
// ne sont PAS mis en cache. Les arbres mis en cache sont traités en LECTURE SEULE (le code re-sérialise +
// reparse tout sous-arbre qu'il doit muter) -> partage sûr entre diapos (prouvé par hash SVG identique).
const _pc = new Map(); // chemin -> { sig, txt?, flat?, ord? } ; jamais purgé (taille = nb de fichiers stables vus)
// Signature « mtime:taille » d'un fichier. statSync LÈVE si le fichier n'existe pas : les appelants
// vérifient existsSync avant, ou entourent d'un try (presentation.xml, tableStyles.xml).
function _stamp(path) { const s = statSync(path); return s.mtimeMs + ':' + s.size; }
// cachedParse(path, kind) -> texte ('txt'), arbre plat ('flat') ou arbre ordonné ('ord', défaut).
// Un stat par appel (peu coûteux) garantit l'invalidation si le fichier a changé sur disque.
// ??= : chaque représentation n'est calculée qu'à la première demande.
function cachedParse(path, kind /* 'txt' | 'flat' | 'ord' */) {
  const sig = _stamp(path);
  let e = _pc.get(path);
  // signature différente -> entrée neuve : on jette txt/flat/ord périmés d'un coup
  if (!e || e.sig !== sig) { e = { sig }; _pc.set(path, e); }
  if (e.txt === undefined) e.txt = readFileSync(path, 'utf8');
  if (kind === 'txt') return e.txt;
  if (kind === 'flat') return (e.flat ??= PF.parse(e.txt));
  return (e.ord ??= P.parse(e.txt));
}

// ████ §C ─ HELPERS D'ARBRE ET ÉCHAPPEMENT ──────────────────────────────────────────────────────
// Un nœud preserveOrder ressemble à : { 'a:off': [], ':@': { '@x': '0', '@y': '12700' } }
// ou, pour du texte : { '#text': 'Bonjour' }. La seule clé autre que ':@' est le nom de balise.
// Ces helpers sont tolérants : un nœud absent (undefined) donne {} / [] au lieu de lever, ce qui
// permet d'enchaîner first(first(x,'a'),'b') sans garde à chaque niveau.
// ---- helpers arbre preserveOrder ----
// tagOf(n) : nom de la balise ('a:sp', 'p:txBody'… ou '#text' pour un nœud texte)
const tagOf = (n) => Object.keys(n).find((k) => k !== ':@');
// at(n) : attributs du nœud ({ '@id': '3', … }) ; {} si aucun ou si n est absent
const at = (n) => n?.[':@'] ?? {};
// kids(n) : enfants ordonnés (tableau) ; [] pour un nœud texte ou absent
const kids = (n) => { const v = n?.[tagOf(n)]; return Array.isArray(v) ? v : []; };
// byTag(n, t) : enfants DIRECTS de balise t, dans l'ordre du document
const byTag = (n, t) => kids(n).filter((c) => tagOf(c) === t);
// first(n, t) : 1er enfant direct de balise t (undefined si aucun)
const first = (n, t) => byTag(n, t)[0];
// deepFind(n, t) : 1er nœud de balise t en profondeur d'abord (n lui-même compris). Attention :
// « le 1er » au sens du document — ex. deepFind(sp, 'p:cNvPr') trouve bien celui de la forme, mais
// deepFind(grp, 'p:cNvPr') sur un groupe trouve celui du GROUPE (nvGrpSpPr précède les membres).
function deepFind(n, t) {
  if (!n || typeof n !== 'object') return null;
  if (tagOf(n) === t) return n;
  for (const c of kids(n)) { const r = deepFind(c, t); if (r) return r; }
  return null;
}
// recherche profonde d'une clé dans un objet PLAT (parse non-ordonné) -> 1re valeur trouvée
// Note : `if (r)` ignore une valeur falsy trouvée en profondeur ('' ou 0) ; seul le niveau courant
// accepte toute valeur !== undefined. Sans conséquence pour les usages (nœuds objets a14:*).
function deepFindFlat(obj, key) {
  if (!obj || typeof obj !== 'object') return null;
  if (obj[key] !== undefined) return obj[key];
  for (const k of Object.keys(obj)) { const r = deepFindFlat(obj[k], key); if (r) return r; }
  return null;
}
// trouve le 1er noeud dont le tag se termine par `suffix` (ignore le préfixe de namespace)
// Utile quand le préfixe varie d'un fichier à l'autre (ex. asvg:svgBlip, préfixe choisi par l'auteur).
function deepFindSuffix(n, suffix) {
  if (!n || typeof n !== 'object') return null;
  const t = tagOf(n);
  if (t && t.endsWith(suffix)) return n;
  for (const c of kids(n)) { const r = deepFindSuffix(c, suffix); if (r) return r; }
  return null;
}
// Échappe aussi ' (XSS-01) : un typeface comme `x';color:red` sortait de la chaîne CSS 'font-family'
// et injectait des propriétés. On utilise l'entité NOMMÉE &apos; (et non la référence numérique &#39;) :
// fast-xml-parser NE décode PAS &#39; lors d'un re-parsing (voir PF.parse sur les fills), si bien qu'un
// texte re-sérialisé puis re-esc donnait &amp;#39; (double-échappement visible sur chaque apostrophe
// française — « l&#39;unité »). &apos; est décodée au round-trip et protège autant (apostrophe échappée).
// esc(s) : échappement XML/HTML des 5 caractères spéciaux. Utilisé PARTOUT où une valeur venant du
// fichier (texte, typeface, id, attribut re-sérialisé) entre dans le SVG/HTML produit : c'est la
// barrière anti-injection (un .pptx est une entrée non fiable). null/undefined -> ''.
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));

// ████ §D ─ OFFICE MATH (OMML) -> HTML ──────────────────────────────────────────────────────────
// Les équations de l'éditeur PowerPoint sont stockées en OMML (balises m:*), en général sous
// <mc:AlternateContent><mc:Choice Requires="a14"><… a14:m><m:oMathPara>…. Il n'y a pas de rendu
// MathML ici : on produit du HTML/CSS simple (inline-flex pour les fractions, sup/sub, bordures
// pour la barre de fraction et le radical) inséré dans le foreignObject du paragraphe.
// Chaîne d'appel : renderText (chemin « paraO ») -> renderMath(nœud) -> mathTypeset(texte, sty).
// OMML (§22.1 Office Math) -> HTML dans le foreignObject. Noeud ORDONNÉ (preserveOrder).
// Gère fraction (m:f), parenthèses/délimiteurs (m:d), exposant/indice (m:sSup/m:sSub),
// racine (m:rad), n-aire (m:nary, somme/intégrale). Le math est en italique (convention).
// mathText : non appelée dans ce fichier (remplacée par rawText + mathTypeset qui gardent le style
// gras/italique des alphanumériques mathématiques) ; conservée telle quelle.
function mathText(node) { // texte brut d'un noeud (m:t / #text)
  // NFKC : convertit les « Mathematical Alphanumeric Symbols » (𝜌 𝑓 𝒅 𝑷 𝒁 𝒍…, U+1D400+)
  // en lettres normales (ρ f d P Z l) — sinon Poppins ne les a pas et affiche des blancs/tofu.
  if (node['#text'] !== undefined) return esc(String(node['#text']).normalize('NFKC'));
  return kids(node).map(mathText).join('');
}
// texte brut d'un noeud, SANS normalisation (pour lire le style des alphanumériques mathématiques)
function rawText(node) {
  if (node['#text'] !== undefined) return String(node['#text']);
  return kids(node).map(rawText).join('');
}
// Typographie mathématique comme PowerPoint (OMML) : lettres en italique, chiffres et ponctuation droits,
// espace autour des relations (0,28 em) et des opérateurs binaires (0,22 em). Les « alphanumériques
// mathématiques » Unicode (𝑷 𝒅 𝟓…) donnent leur style (gras / italique) puis sont ramenés en lettres
// normales (NFKC) — la police n'a pas ces glyphes. sty : style imposé par m:sty (p, b, i, bi).
// MATH_REL : relations (espacées de 0,28 em de chaque côté) ; MATH_BIN : opérateurs binaires (0,22 em).
// Ces valeurs correspondent aux espaces « thick » (5/18 em) et « medium » (4/18 em) de la
// typographie mathématique classique (TeX).
const MATH_REL = new Set([...'=<>≤≥≈≠≅∼∝≡→←⇒⇔']), MATH_BIN = new Set([...'+−-×÷±·∓']);
// mathTypeset(text, sty) -> HTML. Découpe le texte en « sous-runs » de même style (gras/italique)
// pour limiter le nombre de <span>. Chaque opérateur est isolé dans son propre <span> espacé.
// sty (m:sty) : 'p' = droit, 'b' = gras droit, 'i' = italique, 'bi' = gras italique, null = règle
// par défaut OMML (lettres en italique, le reste droit).
export function mathTypeset(text, sty) {
  let out = '', runStyle = null, runText = '';
  // flush : émet le sous-run en cours (avec ou sans <span style>)
  const flush = () => { if (runText) out += runStyle ? `<span style="${runStyle}">${esc(runText)}</span>` : esc(runText); runText = ''; };
  // push : ajoute un caractère ; change de sous-run si le style diffère
  const push = (style, c) => { if (style !== runStyle) { flush(); runStyle = style; } runText += c; };
  for (const ch of text) { // for…of itère par point de code (les U+1D4xx sont des paires de substitution)
    const cp = ch.codePointAt(0);
    let bold = sty === 'b' || sty === 'bi', ital = null; // ital null = « non imposé par le glyphe »
    if (cp >= 0x1D400 && cp <= 0x1D7FF) { // bloc Mathematical Alphanumeric Symbols
      // Disposition Unicode du bloc (latin, 52 lettres = 0x34 par variante) :
      //   o 0x00-0x33 gras · 0x34-0x67 italique · 0x68-0x9B gras italique · au-delà : script,
      //   fraktur, double barre… (ni gras ni italique ici). Grec : 0x1D6A8-6E1 gras,
      //   0x1D6E2-71B italique, 0x1D71C-755 gras italique. Chiffres gras : 0x1D7CE-7D7.
      const o = cp - 0x1D400;
      if (o < 0x34 || (o >= 0x68 && o < 0x9C) || (cp >= 0x1D6A8 && cp <= 0x1D6E1) || (cp >= 0x1D71C && cp <= 0x1D755) || (cp >= 0x1D7CE && cp <= 0x1D7D7)) bold = true;
      ital = (o >= 0x34 && o < 0x9C) || (cp >= 0x1D6E2 && cp <= 0x1D755);
    }
    // NFKC ramène 𝑷 -> P, 𝜌 -> ρ, 𝟓 -> 5 (la police n'a pas les glyphes du bloc 1D400)
    const c = ch.normalize('NFKC');
    if (MATH_REL.has(c) || MATH_BIN.has(c)) { // opérateur : son propre bloc espacé
      flush(); runStyle = null;
      out += `<span style="padding:0 ${MATH_REL.has(c) ? '.28' : '.22'}em">${esc(c === '-' ? '−' : c)}</span>`; continue;
    }
    // italique : forcé droit par m:sty p/b ; sinon style porté par le glyphe ; sinon « lettre = italique »
    // (convention OMML : variables en italique, chiffres/ponctuation droits)
    const letter = /\p{L}/u.test(c);
    const it = sty === 'p' || sty === 'b' ? false : ital ?? letter;
    // style vide -> '' || null -> null = texte sans <span>, fusionnable avec ses voisins
    push((it ? 'font-style:italic;' : '') + (bold ? 'font-weight:700;' : '') || null, c);
  }
  flush();
  return out;
}
// renderMath(nœud OMML ordonné) -> HTML. Descente récursive sur la structure OMML :
//   m:r (run : m:rPr + m:t) · m:f (fraction m:num / m:den) · m:d (délimiteurs, m:dPr begChr/endChr,
//   défauts « ( » et « ) ») · m:sSup / m:sSub (base m:e + exposant/indice) · m:rad (racine, l'indice
//   de degré m:deg est ignoré) · m:nary (opérateur n-aire m:naryPr/m:chr, défaut ∫, bornes sub/sup).
// Toute autre balise (m:oMath, m:oMathPara, m:e, m:num…) est un simple conteneur : on concatène ses
// enfants. Les balises *Pr (propriétés) ne produisent rien.
// Non gérés (rendus « à plat » par le default) : matrices m:m, accents m:acc, barres m:bar, boîtes,
// m:sSubSup, limites m:limLow/m:limUpp, etc.
function renderMath(node) {
  if (node['#text'] !== undefined) return esc(String(node['#text']).normalize('NFKC'));
  const t = tagOf(node);
  if (/Pr$/.test(t)) return ''; // m:fPr, m:dPr, m:rPr, m:ctrlPr, m:naryPr... = propriétés
  switch (t) {
    case 'm:r': { // run mathématique : style m:sty (p droit, b gras, i italique, bi) puis texte typographié
      const rp = first(node, 'm:rPr');
      const sty = at(first(rp || {}, 'm:sty'))['@m:val'] ?? (first(rp || {}, 'm:nor') ? 'p' : null);
      return byTag(node, 'm:t').map((t) => mathTypeset(rawText(t), sty)).join('');
    }
    case 'm:t': return mathTypeset(rawText(node), null);
    // fraction : colonne flex num/den, barre = bordure basse du numérateur (0,06 em)
    case 'm:f': {
      const num = renderMath(first(node, 'm:num') || {}), den = renderMath(first(node, 'm:den') || {});
      return `<span style="display:inline-flex;flex-direction:column;vertical-align:middle;text-align:center;line-height:1.05;font-size:.95em">`
        + `<span style="border-bottom:.06em solid;padding:0 .25em">${num}</span><span style="padding:0 .25em">${den}</span></span>`;
    }
    case 'm:d': {
      const dPr = first(node, 'm:dPr');
      const beg = at(first(dPr || {}, 'm:begChr'))['@m:val'] ?? '(', end = at(first(dPr || {}, 'm:endChr'))['@m:val'] ?? ')';
      return esc(beg) + byTag(node, 'm:e').map(renderMath).join('') + esc(end);
    }
    case 'm:sSup': return renderMath(first(node, 'm:e') || {}) + `<sup style="font-size:.7em">${renderMath(first(node, 'm:sup') || {})}</sup>`;
    case 'm:sSub': return renderMath(first(node, 'm:e') || {}) + `<sub style="font-size:.7em">${renderMath(first(node, 'm:sub') || {})}</sub>`;
    // racine : signe √ + barre supérieure (border-top) sur le radicande m:e
    case 'm:rad': return `√<span style="border-top:.06em solid;padding:0 .15em">${renderMath(first(node, 'm:e') || {})}</span>`;
    case 'm:nary': {
      const op = at(first(first(node, 'm:naryPr') || {}, 'm:chr'))['@m:val'] ?? '∫';
      const sub = renderMath(first(node, 'm:sub') || {}), sup = renderMath(first(node, 'm:sup') || {});
      return `<span style="font-size:1.3em">${esc(op)}</span>${sub ? `<sub style="font-size:.7em">${sub}</sub>` : ''}${sup ? `<sup style="font-size:.7em">${sup}</sup>` : ''}${renderMath(first(node, 'm:e') || {})}`;
    }
    default: return kids(node).map(renderMath).join(''); // m:oMath, m:oMathPara, m:e, m:r, m:num, m:den...
  }
}

// ████ §E ─ PUCES ET POLICES -> FAMILLES CSS ────────────────────────────────────────────────────
// bulletOf lit la définition de puce d'un niveau de paragraphe (a:pPr / a:lvlNpPr, §21.1.2.4.x :
// buNone, buChar, buAutoNum, buBlip, buClr, buSzPct, buFont). fontFamily traduit un typeface
// OOXML en pile CSS avec substitut libre métrique-compatible (pour que la COUPURE DE LIGNE soit la
// même que sous Windows même sans la police d'origine).
//
// puce d'un pPr (flat) : {none} | {char,font} | {num,font} | null (non défini -> hériter)
// Le null est important : il laisse la chaîne d'héritage (lstStyle -> layout -> master) décider,
// alors que {none:true} COUPE l'héritage (a:buNone explicite).
// Non lus : a:buSzPts, a:buClrTx / a:buSzTx / a:buFontTx (qui signifient « comme le texte », ce
// qui est déjà le comportement par défaut ici) ; @startAt de buAutoNum (numérotation depuis 1).
function bulletOf(pPr) {
  if (!pPr) return null;
  if (pPr['a:buNone'] !== undefined) return { none: true };
  // a:buClr = couleur PROPRE de la puce (souvent une couleur d'accent différente du texte). Nœud brut
  // conservé ; résolu au rendu (a besoin du thème/ctx). a:buSzPct/@val = taille de puce en % de la police.
  const clr = pPr['a:buClr'] || null;
  const szPct = pPr['a:buSzPct']?.['@val'] ? +pPr['a:buSzPct']['@val'] / 100000 : null;
  // a:buBlip = puce IMAGE (petit logo/icône). Nœud brut conservé ; résolu au rendu (curRels/pptDir).
  if (pPr['a:buBlip']) return { blip: pPr['a:buBlip']['a:blip'], szPct };
  if (pPr['a:buChar']) return { char: pPr['a:buChar']['@char'], font: pPr['a:buFont']?.['@typeface'], clr, szPct };
  // numérotation auto : seul le suffixe est interprété au rendu (« ) » si le type contient « paren »,
  // sinon « . ») ; le compteur est toujours arabe (1, 2, 3…), même pour alphaLcPeriod / romanUcPeriod.
  if (pPr['a:buAutoNum']) return { num: pPr['a:buAutoNum']['@type'] || 'arabicPeriod', font: pPr['a:buFont']?.['@typeface'], clr, szPct };
  return null;
}
// police -> famille CSS. Substituts LIBRES métrique-compatibles (embarqués) pour les polices non libres
// courantes : garantit largeur/retour-à-la-ligne corrects MÊME sans la police installée (ex. LMS Linux).
// Ordre : police réelle d'abord (si installée), puis substitut embarqué, puis générique.
// SERIF_RE : heuristique « cette police inconnue est-elle à empattements ? » -> générique serif.
// (« book » attrape aussi des sans-serif comme Franklin Gothic Book, mais celle-ci est dans FONT_SUBST,
// testé avant.)
const SERIF_RE = /times|cambria|georgia|bernhard|garamond|antiqua|minion|book|schoolbook/i;
// SYMB_RE : polices symboles -> pas de repli générique (le glyphe d'une autre police serait faux ; le
// texte de ces polices est de toute façon remappé vers Unicode, voir §H).
const SYMB_RE = /wingding|webding|symbol/i;
// Table typeface (minuscules) -> pile CSS « police réelle, substitut libre, générique ». Les
// substituts (Liberation, Carlito, Caladea, Gelasio…) sont conçus avec les MÊMES chasses que
// l'original : texte de même largeur, donc mêmes retours à la ligne. Ils sont embarqués côté web
// (animFonts.js / web-fonts.mjs).
const FONT_SUBST = {
  'arial': "'Arial','Liberation Sans',sans-serif", 'helvetica': "'Helvetica','Liberation Sans',sans-serif",
  'arial narrow': "'Arial Narrow','Liberation Sans Narrow',sans-serif",
  'times new roman': "'Times New Roman','Liberation Serif',serif", 'times': "'Times','Liberation Serif',serif",
  'courier new': "'Courier New','Liberation Mono',monospace", 'courier': "'Courier','Liberation Mono',monospace",
  'calibri': "'Calibri','Carlito',sans-serif", 'cambria': "'Cambria','Caladea',serif",
  // police PowerPoint d'abord (poste qui l'a installée), puis équivalent libre embarqué (animFonts.js)
  'calibri light': "'Calibri Light','Carlito',sans-serif",
  'cambria math': "'Cambria Math','Caladea','DejaVu Sans',serif",
  'consolas': "'Consolas','Liberation Mono',monospace",
  'georgia': "'Georgia','Gelasio',serif",
  'comic sans ms': "'Comic Sans MS','Comic Relief',cursive",
  'arial black': "'Arial Black','Archivo Black',sans-serif",
  'impact': "'Impact','Anton',sans-serif",
  'franklin gothic book': "'Franklin Gothic Book','Libre Franklin',sans-serif",
  'franklin gothic medium': "'Franklin Gothic Medium','Libre Franklin',sans-serif",
  'news gothic mt': "'News Gothic MT','News Cycle',sans-serif",
  'gill sans mt': "'Gill Sans MT','Cabin',sans-serif", 'gill sans': "'Gill Sans','Cabin',sans-serif",
  'bernhard modern roman': "'Bernhard Modern Roman','Sorts Mill Goudy',serif",
  'century schoolbook': "'Century Schoolbook','TeX Gyre Schola',serif",
  'palatino linotype': "'Palatino Linotype','TeX Gyre Pagella',serif",
  'palatino': "'Palatino','TeX Gyre Pagella',serif", 'book antiqua': "'Book Antiqua','TeX Gyre Pagella',serif",
  'corbel': "'Corbel','Source Sans 3',sans-serif",
  'aptos': "'Aptos','Source Sans 3',sans-serif", 'aptos display': "'Aptos Display','Source Sans 3',sans-serif",
};
// fontFamily(typeface) -> valeur CSS font-family (déjà échappée). Le typeface doit être RÉSOLU
// (pas « +mn-lt ») : c'est resolveFont, dans renderText, qui remplace les références de thème.
// Les noms entre quotes de la pile servent aussi de clés à metricsOf / advanceTable (§F) : le 1er
// nom pour lequel on a des métriques gagne (souvent le substitut embarqué).
function fontFamily(f) {
  if (!f) return 'sans-serif';
  if (SYMB_RE.test(f)) return `'${esc(f)}'`;
  const key = f.toLowerCase().trim();
  if (FONT_SUBST[key]) return FONT_SUBST[key];
  return SERIF_RE.test(f) ? `'${esc(f)}',serif` : `'${esc(f)}',sans-serif`;
}
// ████ §F ─ MÉTRIQUES DE POLICE : LIGNE DE BASE, CHASSES, LARGEURS ──────────────────────────────
// Le texte est mis en page par le navigateur (HTML dans foreignObject). Trois écarts systématiques
// avec PowerPoint sont corrigés ici, à partir des métriques des polices embarquées (animFonts.js) :
//   1. position VERTICALE de la ligne de base dans la ligne  -> baselineShift / applyShift / VA_TOK
//   2. largeur des mots (chasses arrondies au 1/8 pt)        -> advanceFix (letter-spacing)
//   3. mesures de largeur pour les tabulations / tableaux    -> pptTextWidth
// Notations : sz = taille de police (pt = px CSS ici), L = hauteur de ligne (px), A / D = ascendante /
// descendante de la police en fraction d'em (valeurs positives).
//
// Ligne de base façon PowerPoint. Calibrage (PowerPoint exporté vs navigateur, 4 polices, interlignes
// 80-200 %) : la ligne de base de chaque ligne tombe, depuis le haut de sa ligne de hauteur L, à
//   L > 1,2 x taille (interligne > 100 %) : 0,75 L (quelle que soit la police)
//   sinon                                  : L - min(0,25 L, descendante win)
// alors que le navigateur la met à L/2 + (ascendante - descendante)/2. L'écart (constant pour toutes les
// lignes du paragraphe) est rattrapé par un décalage VISUEL (position:relative) : la mise en page (pas
// de ligne, hauteur du bloc, ancrage) reste celle du navigateur, identique à PowerPoint.
let _fm = null; // cache paresseux de familyMetrics() : { 'Nom de famille': [A, D], … } ({} si échec)
// metricsOf(pile CSS) -> [A, D] de la 1re famille de la pile qui a des métriques, sinon null.
function metricsOf(famCss) {
  if (!_fm) { try { _fm = familyMetrics(); } catch { _fm = {}; } }
  const fam = (famCss.match(/'([^']+)'/g) || []).map((q) => q.slice(1, -1)).find((f) => _fm[f]);
  return fam ? _fm[fam] : null;
}
// Recalage de ligne de base appliqué aux RUNS (vertical-align), et non au paragraphe : un
// position:relative dans un foreignObject est dessiné par WebKit (Safari / iPad) par rapport à la PAGE
// et non au SVG -> tout le texte s'empilait en haut à gauche de la diapo sur iPad. Les runs étant en
// line-height 0, leur vertical-align déplace les glyphes sans changer la mise en page.
// Mécanisme en deux temps : chaque run est émis avec `vertical-align:calc(<exposant>em - @@VA@@px)`,
// car le décalage d dépend de la police du 1er run et de l'interligne du paragraphe, connus seulement
// après avoir lu tous les runs. applyShift remplace ensuite le jeton par la valeur de d. Soustraire d
// (vertical-align positif = vers le HAUT) abaisse les glyphes quand PowerPoint place la base plus bas.
const VA_TOK = ' - @@VA@@px)'; // fin de « vertical-align:calc(Xem - ?px) », remplacée par paragraphe
// applyShift(html, d) : substitue toutes les occurrences du jeton (split/join = replaceAll) par « - d px) ».
function applyShift(html, d) {
  return html.split(VA_TOK).join(` - ${(d || 0).toFixed(2)}px)`);
}
// baselineShift(pile CSS, sz, L) -> d (px) = position PowerPoint - position navigateur de la ligne de
// base, mesurée depuis le haut de la ligne. 0 si police sans métriques ou entrées invalides.
export function baselineShift(famCss, sz, L) {
  const m = metricsOf(famCss);
  if (!m || !(sz > 0) || !(L > 0)) return 0;
  const [A, D] = m;
  // modèle CSS (« half-leading ») : la boîte A+D est centrée dans L -> base à L/2 + (A-D)/2 · sz
  const css = L / 2 + (A - D) / 2 * sz;
  // modèle PowerPoint calibré (voir le commentaire au-dessus de metricsOf) ; +0,01 = tolérance
  // flottante pour que l'interligne simple (L = 1,2 sz exactement) tombe dans la 2e branche
  const ppt = L > 1.2 * sz + 0.01 ? 0.75 * L : L - Math.min(0.25 * L, D * sz);
  const d = ppt - css;
  // écart < 0,05 px ignoré : évite d'émettre un calc() pour un bruit d'arrondi
  return Math.abs(d) < 0.05 ? 0 : d;
}
// Arrondi des chasses façon PowerPoint : chaque glyphe avance d'un multiple de 1/8 pt (1/576 pouce,
// arrondi au plus proche) — calibré sur PowerPoint via COM : 0,014 % d'écart contre 0,14 % sans. Le
// navigateur, lui, cumule des chasses exactes : à 18 pt Poppins il compose 0,36 % plus large, assez
// pour renvoyer un mot à la ligne. -> écart moyen par caractère, ajouté au letter-spacing du run.
// Cache des tables de chasses : clé « famille|gras|italique » -> { a: {codepoint: chasse}, upm } ou null
const _advCache = new Map();
// advanceFix(pile CSS, gras, italique, sz, texte) -> correction MOYENNE par caractère (px) à ajouter
// au letter-spacing du run. letter-spacing s'ajoutant après chaque caractère, n × moyenne = somme des
// écarts d'arrondi : la largeur totale du run égale celle de PowerPoint (à la répartition près).
export function advanceFix(famCss, bold, italic, sz, text) {
  if (!text || !(sz > 0)) return 0;
  const fams = (famCss.match(/'([^']+)'/g) || []).map((q) => q.slice(1, -1)); // noms entre quotes de la pile
  let tbl = null;
  for (const f of fams) { // 1re famille de la pile qui possède une table
    const k = f + '|' + bold + '|' + italic;
    if (!_advCache.has(k)) { let t = null; try { t = advanceTable(f, bold, italic); } catch { /* sans table */ } _advCache.set(k, t); }
    if ((tbl = _advCache.get(k))) break;
  }
  if (!tbl) return 0;
  let delta = 0, n = 0;
  for (const ch of text) {
    const a = tbl.a[ch.codePointAt(0)];
    if (a == null) return 0; // glyphe hors table (repli de police) : ne pas deviner
    const w = a * sz / tbl.upm; // chasse exacte en pt (a en unités de fonte, upm = unités par em)
    delta += Math.round(w * 8) / 8 - w; n++; // arrondi au 1/8 pt le plus proche, moins la valeur exacte
  }
  return n ? delta / n : 0;
}
// Largeur (pt) d'un texte avec les chasses ARRONDIES de PowerPoint (1/8 pt) ; null si glyphe inconnu.
// Sert aux calculs que le navigateur ne sait pas faire lui-même : position des tabulations
// (renderText) et hauteur des lignes de tableau (cellTextHeight). N'utilise pas _advCache (appelle
// advanceTable à chaque fois) et ne tient pas compte du crénage ni de @spc.
function pptTextWidth(famCss, bold, italic, sz, text) {
  const fams = (famCss.match(/'([^']+)'/g) || []).map((q) => q.slice(1, -1));
  let tbl = null;
  for (const f of fams) { try { tbl = advanceTable(f, bold, italic); } catch { tbl = null; } if (tbl) break; }
  if (!tbl) return null;
  let w = 0;
  for (const ch of text) { const a = tbl.a[ch.codePointAt(0)]; if (a == null) return null; w += Math.round(a * sz / tbl.upm * 8) / 8; }
  return w;
}
// ████ §G ─ HAUTEUR DE TEXTE D'UNE CELLULE DE TABLEAU ───────────────────────────────────────────
// Dans un tableau, a:tr@h n'est qu'une hauteur MINIMALE : PowerPoint agrandit la ligne si le texte
// ne tient pas. Le navigateur ne peut pas faire ce calcul pour nous (chaque cellule
// est un foreignObject indépendant, de hauteur fixée), d'où cette mesure « à la PowerPoint » faite
// côté Node, appelée par renderTable.
//
// Hauteur (EMU) qu'occupe le texte d'une cellule de largeur cellW (EMU) : coupure mot à mot comme
// PowerPoint, 1,2 x taille par ligne (interligne simple), + marges haut/bas de la cellule.
// Entrée : tc = cellule en arbre PLAT ; styleBold = gras imposé par le style (1re ligne / 1re colonne).
// Sortie : EMU, ou null si la mesure est impossible (glyphe sans table de chasses, cellule vide) ->
// l'appelant garde alors @h tel quel.
// Simplifications : ignore lnSpc / spcBef / spcAft, les a:br (lus en plat, seuls les a:r comptent),
// les tabulations et l'autofit ; la police de thème est toujours la MINEURE (+mj- compris).
export function cellTextHeight(tc, cellW, styleBold) {
  const tcPr = tc['a:tcPr'] || {}, S = 1 / 12700; // S local = même valeur que la constante S de §I
  // largeur utile en pt : largeur de colonne(s) - marges gauche/droite (défaut 91440 EMU = 0,1")
  const avail = (cellW - (+(tcPr['@marL'] ?? 91440)) - (+(tcPr['@marR'] ?? 91440))) * S;
  const arrOf = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];
  let h = 0;
  for (const p of arrOf(tc['a:txBody']?.['a:p'])) {
    const runs = arrOf(p['a:r']);
    const endSz = +(p['a:endParaRPr']?.['@sz']) || 0;
    // paragraphe vide : une ligne à la taille de sa marque de fin (défaut 18 pt = @sz 1800)
    if (!runs.length) { h += 1.2 * ((endSz || 1800) / 100); continue; }
    let lineW = 0, lines = 1, maxSz = 0;
    for (const r of runs) {
      const rp = r['a:rPr'] || {}, sz = (+rp['@sz'] || endSz || 1800) / 100;
      const bold = rp['@b'] !== undefined ? rp['@b'] === '1' : styleBold;
      const tf = rp['a:latin']?.['@typeface'];
      const fam = fontFamily(!tf || /^\+m[nj]-/.test(tf) ? _ctxMinorFont : tf);
      maxSz = Math.max(maxSz, sz); // la plus grande taille du paragraphe fixe la hauteur de TOUTES ses lignes
      for (const word of String(r['a:t'] ?? '').split(/(?<= )/)) { // mots avec leur espace finale
        const w = pptTextWidth(fam, bold, rp['@i'] === '1', sz, word);
        if (w == null) return null;
        // règle de coupure PowerPoint : l'espace finale d'un mot ne compte pas pour savoir s'il tient
        const wNoSp = word.endsWith(' ') ? w - (pptTextWidth(fam, bold, false, sz, ' ') || 0) : w;
        // -0,02 pt : petite tolérance ; PowerPoint renvoie le mot dès que la ligne atteint la largeur
        // disponible à ~0,01 pt près (même constat que le margin-right 0,2 px de renderText)
        if (lineW > 0 && lineW + wNoSp > avail - 0.02) { lines++; lineW = w; } else lineW += w; // même règle que le texte
      }
    }
    h += lines * 1.2 * maxSz;
  }
  if (!h) return null;
  // retour en EMU + marges haut/bas (défaut 45720 EMU = 0,05")
  return h / S + (+(tcPr['@marT'] ?? 45720)) + (+(tcPr['@marB'] ?? 45720));
}
// État de MODULE (et non de ctx) parce que cellTextHeight est exportée et appelée sans ctx.
// Écrit par renderSlide à chaque appel ; lu par cellTextHeight.
let _ctxMinorFont = 'Calibri'; // police mineure du thème de la diapo en cours (renderSlide)
// ████ §H ─ POLICES SYMBOLES -> UNICODE ─────────────────────────────────────────────────────────
// Deux façons dont un .pptx utilise une police symbole :
//   (1) la police latin du run EST la police symbole (a:latin typeface="Wingdings") -> tout le texte
//       du run est à remapper (octets 0x20-0xFF interprétés dans le jeu de la police) ;
//   (2) le run porte <a:sym typeface="Symbol"> et contient des caractères de la zone privée
//       U+F000-F0FF (convention Windows pour les polices symboles : U+F000 + octet) -> seuls ces
//       caractères-là sont des symboles.
// Chaîne d'appel : bulletSym (puces) ; symbolParts / symbolText (texte des runs) -> symToUnicode.
// Polices SYMBOLE (Wingdings/Webdings/Symbol) : le glyphe n'existe qu'avec la police installée -> tofu
// sur Linux (cible LMS) et dans Chromium headless. On mappe les codepoints vers leur équivalent Unicode,
// rendus en police NORMALE = lisibles partout. Gère l'encodage zone privée PUA (0xF000+octet en run).
const WINGD_RE = /wingding|webding/i; // dingbats PURS (que des symboles) — décoratif
const SYMBOLFONT_RE = /^symbol$/i;      // Adobe « Symbol » = grec + flèches + math (PAS décoratif)
// Wingdings -> Unicode (puces/dingbats). Inconnu -> « • » (marqueur générique, jamais de tofu).
const WINGDINGS_MAP = {
  0x6C: '●', 0x6D: '○', 0x6E: '■', 0x6F: '□', 0x71: '❑', 0x75: '◆',
  0xA7: '▪', 0xA8: '▫', 0xB7: '•', 0xD8: '➢', 0xD9: '➢',
  // 0xFB-FE (table Wingdings) : ✗ croix, ✓ coche FINE (pas ✔ épaisse, cf. 040_02_01 d.57), ☒, ☑
  0xFB: '✗', 0xFC: '✓', 0xFD: '☒', 0xFE: '☑',
};
// Adobe « Symbol » -> Unicode : grec (ASCII->grec) + flèches (0xAB-AF, 0xDB-DF) + opérateurs math.
// Inconnu -> on GARDE le caractère d'origine (symbole signifiant : ne jamais deviner à tort).
const SYMBOL_MAP = (() => {
  const m = {};
  const low = 'abgdezhqiklmnxoprstufcyw', lowG = 'αβγδεζηθικλμνξοπρστυφχψω';
  const up  = 'ABGDEZHQIKLMNXOPRSTUFCYW', upG  = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ';
  // correspondance positionnelle : a->α, b->β, g->γ, d->δ, … (translittération de la police Symbol)
  for (let i = 0; i < low.length; i++) { m[low.charCodeAt(i)] = lowG[i]; m[up.charCodeAt(i)] = upG[i]; }
  // phi : la police Symbol DESSINE 0x66 « f » en ϕ (trait droit) et 0x6A « j » en φ (bouclé) -> afficher ce
  // que montre PowerPoint (l'encodage Adobe donne l'inverse, d'où des φ bouclés non conformes au support).
  m[0x66] = 'ϕ'; m[0x6A] = 'φ';
  Object.assign(m, {
    0xAB:'↔',0xAC:'←',0xAD:'↑',0xAE:'→',0xAF:'↓', 0xDB:'⇔',0xDC:'⇐',0xDD:'⇑',0xDE:'⇒',0xDF:'⇓',
    0xA3:'≤',0xB3:'≥',0xB9:'≠',0xBB:'≈',0xBA:'≡',0xB1:'±',0xB4:'×',0xB8:'÷',0xB0:'°',0xB7:'·',
    0x40:'≅',0xD6:'√',0xA5:'∞',0xB6:'∂',0xF2:'∫',0x2D:'−',0xE5:'∑',0xCF:'∨',0xCE:'∈',
  });
  return m;
})();
// symFontKind(typeface) -> 'wing' (Wingdings/Webdings), 'sym' (Adobe Symbol, nom EXACT « Symbol »
// pour ne pas attraper « Segoe UI Symbol ») ou null (police ordinaire).
function symFontKind(font) {
  return !font ? null : WINGD_RE.test(font) ? 'wing' : SYMBOLFONT_RE.test(font.trim()) ? 'sym' : null;
}
// mappe UN char (codepoint) d'une police symbole vers Unicode. wing: inconnu->'•' ; sym: inconnu->char.
function symToUnicode(ch, kind) {
  let cp = ch.codePointAt(0);
  if (cp >= 0xF000 && cp <= 0xF0FF) cp -= 0xF000; // PUA -> octet réel
  if (kind === 'wing') return WINGDINGS_MAP[cp] || '•';
  return SYMBOL_MAP[cp] ?? ch; // sym : garder l'original si non mappé
}
// -> {sym, symbolic} pour une PUCE caractère : symbolic=true si remappée (rendre en police normale)
function bulletSym(char, font) {
  const kind = symFontKind(font);
  if (char == null || !kind) return { sym: char, symbolic: false };
  return { sym: symToUnicode(char, kind), symbolic: true };
}
// remappe un TEXTE de run en police symbole -> Unicode (anti-tofu). `symFont` = typeface de <a:sym>.
// Police latin symbole -> TOUT le texte est remappé. <a:sym> seul ne vaut que pour les caractères
// PUA U+F000-F0FF (c'est ainsi que PowerPoint l'applique) : le texte ordinaire d'un run qui porte
// <a:sym typeface="Symbol"> reste en police latin, sinon « The Earth » deviendrait « Τηε Εαρτη ».
const isPua = (c) => { const cp = c.codePointAt(0); return cp >= 0xF000 && cp <= 0xF0FF; };
// -> null (rien à remapper) ou segments [{ text, sym }] : `sym` = glyphe issu de la police symbole,
// à rendre dans SYM_FAMILY ; le reste garde la police latin du run (comme PowerPoint).
// Cas (1) : un seul segment { sym:true } avec tout le texte remappé (les blancs sont conservés tels
// quels). Cas (2) : segments alternés latin / symbole, consécutifs fusionnés.
export function symbolParts(raw, latinFont, symFont) {
  if (!raw) return null;
  const latinKind = symFontKind(latinFont);
  if (latinKind) return [{ text: [...raw].map(c => /\s/.test(c) ? c : symToUnicode(c, latinKind)).join(''), sym: true }];
  const kind = symFontKind(symFont);
  if (!kind || ![...raw].some(isPua)) return null;
  const parts = [];
  for (const c of raw) {
    const sym = isPua(c), text = sym ? symToUnicode(c, kind) : c, last = parts[parts.length - 1];
    if (last && last.sym === sym) last.text += text; else parts.push({ text, sym });
  }
  return parts;
}
// symbolText : même chose aplatie en une chaîne (null si rien à remapper). Export pour les outils /
// tests ; non utilisée dans ce fichier (renderText appelle symbolParts).
export function symbolText(raw, latinFont, symFont) {
  const parts = symbolParts(raw, latinFont, symFont);
  return parts && parts.map(p => p.text).join('');
}
// Police des glyphes remappés : la Symbol d'Adobe a un dessin Times -> Liberation Serif, puis
// DejaVu Sans (embarquée) pour ce que Liberation ne couvre pas (≅, ⇒, ✔, ➢…).
const SYM_FAMILY = "'Times New Roman','Liberation Serif','DejaVu Sans',serif";

// ████ §I ─ STYLE DE TABLEAU INTÉGRÉ, UNITÉS ────────────────────────────────────────────────────
// Style de tableau par défaut de PowerPoint (copie exacte de sa définition, telle qu'écrite dans un
// tableStyles.xml) : repli quand le fichier y renvoie sans la stocker.
// Structure (§20.1.4.2 table styles) : wholeTbl (tout le tableau : texte dk1, bordures lt1 1 pt, fond
// accent1 teinté 20 %), band1H/band1V (bandes impaires : accent1 teinté 40 %), firstRow / lastRow /
// firstCol / lastCol (texte gras lt1 sur accent1 plein, filet 3 pt entre en-tête et corps).
// Parsé une fois en objet plat par buildContext, puis appliqué en couches par renderTable.
const BUILTIN_TBL_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
const BUILTIN_TBL_STYLE = '<a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Style moyen 2 - Accentuation 1"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right><a:top><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top><a:bottom><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom><a:insideH><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideH><a:insideV><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideV></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H><a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V><a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V><a:lastCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastCol><a:firstCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol><a:lastRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:top><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastRow><a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle>';

// EMU -> points (1 pt = 12700 EMU). On rend en points pour des tailles de police normales
// (les foreignObject cassent avec des font-size en centaines de milliers de px).
// S = facteur EMU -> pt. u(v) = valeur en pt arrondie au 1/100 (toFixed(2) puis + pour revenir à un
// nombre : pas de « 12.50 » dans la sortie, mais « 12.5 »). Cet arrondi est appliqué PARTOUT où une
// coordonnée sort dans le SVG : le changer modifierait tous les rendus (empreintes). Note : u()
// renvoie un nombre ; plusieurs appels font ensuite .toFixed(1) dessus pour un format plus court.
const S = 1 / 12700;
const u = (v) => +(Number(v) * S).toFixed(2);

// ████ §J ─ CONTEXTE DOCUMENT (THÈME) ───────────────────────────────────────────────────────────
// Le « ctx » est l'objet transmis (par fermeture) à tout le rendu d'une diapo. Champs :
//   pptDir, clr (slots bruts du thème : dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink -> hex),
//   clrMap (alias bg1/tx1/bg2/tx2/accentN… -> slot, rempli par renderSlide depuis le master),
//   phClr (couleur de substitution courante, voir paint), majorFont / minorFont (polices du
//   fontScheme, script latin uniquement), fillStyles / lnStyles / effectStyles / bgFillStyles
//   (listes ordonnées du fmtScheme, indexées par les fillRef/lnRef/effectRef/bgRef des formes),
//   tableStyles (styleId -> style). Ajoutés ensuite par renderSlide : txStyles, defaultTextStyle,
//   grpFill.
// ================= contexte document =================
// Cache du CŒUR de contexte (parties immuables dérivées du thème) : keyé par fichier de thème (+ sa
// signature mtime/taille). loadContext() renvoie une COPIE avec clrMap/phClr remis à null car ceux-là
// sont MUTÉS par diapo (clrMap du masque, phClr courant) et ne doivent jamais être partagés.
const _ctxCache = new Map();
// loadContext(pptDir, themePath?) -> ctx neuf (copie superficielle du cœur en cache).
// themePath absent ou introuvable -> theme/theme1.xml (cas d'un deck mono-master).
// Attention : la copie est SUPERFICIELLE — fillStyles, tableStyles… sont partagés et ne doivent pas
// être modifiés ; seuls les champs de premier niveau (clrMap, phClr, txStyles…) peuvent l'être.
export function loadContext(pptDir, themePath) {
  const themeFile = (themePath && existsSync(themePath)) ? themePath : join(pptDir, 'theme', 'theme1.xml');
  const sig = _stamp(themeFile);
  let c = _ctxCache.get(themeFile);
  if (!c || c.sig !== sig) { c = { sig, core: buildContext(pptDir, themeFile) }; _ctxCache.set(themeFile, c); }
  return { ...c.core, clrMap: null, phClr: null };
}
// buildContext : construit le cœur immuable du ctx depuis le thème (a:theme/a:themeElements, §20.1.6)
// et tableStyles.xml. Appelé seulement en cas d'absence/péremption du cache.
function buildContext(pptDir, themeFile) {
  // theme du MASTER de la diapo (un deck multi-master a plusieurs thèmes : ex. deck C a
  // theme1 lt2=C78243 orange ET theme2 lt2=D9CAB7 beige -> mauvais thème = mauvaise couleur de fond).
  const theme = cachedParse(themeFile, 'flat');
  const el = theme['a:theme']['a:themeElements'];
  const scheme = el['a:clrScheme'];
  const clr = {}; // slots BRUTS du thème (dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink)
  // Chaque slot contient soit a:srgbClr@val (hex), soit a:sysClr (couleur système Windows : on prend
  // windowText -> noir, window -> blanc, sinon @lastClr = dernière valeur calculée par Office).
  for (const [k, v] of Object.entries(scheme)) {
    if (!k.startsWith('a:')) continue; // saute '@name' (attribut du clrScheme)
    clr[k.slice(2)] = v['a:srgbClr']?.['@val'] ??
      { windowText: '000000', window: 'FFFFFF' }[v['a:sysClr']?.['@val']] ??
      v['a:sysClr']?.['@lastClr'] ?? '000000';
  }
  // listes de styles du fmtScheme, EN ORDRE (fillRef/bgRef indexent dedans) -> noeuds flat
  const themeO = cachedParse(themeFile, 'ord');
  const themeRoot = themeO.find(n => tagOf(n) === 'a:theme');
  const fmt = deepFind(themeRoot, 'a:fmtScheme');
  // Pourquoi le détour par l'arbre ORDONNÉ : fillStyleLst contient p. ex. solidFill, gradFill,
  // gradFill — en plat, les deux gradFill seraient regroupés dans un tableau et l'index de position
  // (fillRef idx=2 -> 2e élément) serait perdu. On prend chaque enfant dans l'ordre, puis on le
  // re-parse SEUL en plat : chaque entrée de la liste = { 'a:solidFill': … } ou { 'a:gradFill': … }.
  // Les xmlns factices du wrapper <root> ne servent qu'à fournir une racine (fast-xml-parser ne
  // résout pas les namespaces).
  const orderedFlatList = (listTag) => {
    const lst = fmt && first(fmt, listTag);
    if (!lst) return [];
    return kids(lst).map(c => PF.parse(`<root xmlns:a="a" xmlns:p="p">${serialize(c)}</root>`).root);
  };
  // styles de tableaux : ceux du fichier, plus le style PAR DÉFAUT de PowerPoint (« Moyen 2 -
  // Accentuation 1 ») que le fichier omet souvent alors que ses tableaux y renvoient.
  const tableStyles = {};
  try { tableStyles[BUILTIN_TBL_STYLE_ID] = PF.parse(`<r xmlns:a="a">${BUILTIN_TBL_STYLE}</r>`).r['a:tblStyle']; } catch {}
  // tableStyles.xml du deck (facultatif : absent -> statSync lève -> catch silencieux). Un style du
  // fichier portant le même styleId que le style intégré le remplace.
  try {
    const tsx = cachedParse(join(pptDir, 'tableStyles.xml'), 'flat');
    const list = tsx['a:tblStyleLst']?.['a:tblStyle'];
    for (const st of (Array.isArray(list) ? list : list ? [list] : [])) {
      tableStyles[st['@styleId']] = st; // style complet (couches wholeTbl, bandes, 1re/dern. ligne/colonne)
    }
  } catch {}
  return {
    pptDir,
    clr,
    clrMap: null, // rempli avec le clrMap du masque
    phClr: null,  // couleur de substitution courante (pour phClr)
    majorFont: el['a:fontScheme']?.['a:majorFont']?.['a:latin']?.['@typeface'] ?? 'Calibri',
    minorFont: el['a:fontScheme']?.['a:minorFont']?.['a:latin']?.['@typeface'] ?? 'Calibri',
    fillStyles: orderedFlatList('a:fillStyleLst'),
    lnStyles: orderedFlatList('a:lnStyleLst'),
    effectStyles: orderedFlatList('a:effectStyleLst'), // ombres/lueurs de thème indexées par a:effectRef
    bgFillStyles: orderedFlatList('a:bgFillStyleLst'),
    tableStyles,
  };
}

// ████ §K ─ COULEURS DRAWINGML ──────────────────────────────────────────────────────────────────
// Une couleur DrawingML (§20.1.2.3) = une couleur de base (srgbClr, schemeClr, sysClr, prstClr,
// scrgbClr ; hslClr n'est pas géré ici) + une liste de TRANSFORMATIONS enfants (lumMod, lumOff,
// shade, tint, alpha…). Ex. « accent1 assombri 25 % » = <a:schemeClr val="accent1"><a:lumMod
// val="75000"/></a:schemeClr>.
// ================= couleurs =================
// Conversions HSL <-> RVB classiques (h, s, l dans [0,1]) ; utilisées pour lumMod / lumOff qui
// agissent sur la LUMINANCE HSL.
function hslFromRgb(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h, s, l = (mx + mn) / 2;
  if (mx === mn) { h = s = 0; }
  else { const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
  return [h, s, l];
}
function rgbFromHsl(h, s, l) {
  let r, g, b;
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const hue = (t) => { t = (t + 1) % 1; return t < 1/6 ? p + (q - p) * 6 * t : t < 1/2 ? q : t < 2/3 ? p + (q - p) * (2/3 - t) * 6 : p; };
    r = hue(h + 1/3); g = hue(h); b = hue(h - 1/3); }
  return [Math.round(r*255), Math.round(g*255), Math.round(b*255)];
}
// résout un noeud couleur DrawingML flat -> {r,g,b,a}
// Entrée : un nœud PARENT d'une couleur (ex. le contenu de a:solidFill, un a:gs, un a:fillRef, un
// a:outerShdw) — la fonction cherche l'enfant srgbClr / schemeClr / … Sortie : r,g,b entiers 0-255,
// a dans [0,1] ; null si aucun enfant couleur (l'appelant choisit alors son repli).
// Ordre d'application des transformations (fixe, indépendant de l'ordre du XML — le parse plat ne le
// connaît pas) : lumMod, puis lumOff (sur la luminance HSL), puis shade (vers le noir), puis tint
// (vers le blanc), puis alpha. Modèle SIMPLIFIÉ : shade/tint sont appliqués linéairement en sRVB.
// Transformations ignorées : satMod, satOff, hueMod, hueOff, comp, inv, gray, gamma, red*/green*/…
function resolveColorFlat(node, ctx) {
  if (!node) return null;
  let hex = null, mods = node;
  // schemeHex(nom) : phClr -> couleur de substitution courante (défaut accent1) ; alias du clrMap
  // (bg1 -> lt1 …) ; puis slot du thème. Slot inconnu -> gris 888888 (visible, jamais transparent).
  const schemeHex = (nm) => {
    if (nm === 'phClr') nm = ctx.phClr || 'accent1';
    if (ctx.clrMap && ctx.clrMap[nm]) nm = ctx.clrMap[nm]; // bg1/tx1/... -> dk1/lt1/...
    return ctx.clr[nm] ?? '888888';
  };
  if (node['a:srgbClr']) { hex = node['a:srgbClr']['@val']; mods = node['a:srgbClr']; }
  else if (node['a:schemeClr']) { hex = schemeHex(node['a:schemeClr']['@val']); mods = node['a:schemeClr']; }
  else if (node['a:sysClr']) { hex = node['a:sysClr']['@lastClr'] ?? '000000'; mods = node['a:sysClr']; }
  else if (node['a:prstClr']) { hex = ({ black:'000000', white:'FFFFFF', red:'FF0000' })[node['a:prstClr']['@val']] ?? '888888'; mods = node['a:prstClr']; }
  // prstClr : seuls black / white / red sont connus (le reste -> gris) ; les autres noms de
  // ST_PresetColorVal ne sont pas tabulés.
  // scrgbClr : composantes en % (0-100000), traitées comme du sRVB linéairement (pas de conversion
  // gamma) ; retour immédiat, SANS appliquer les transformations enfants.
  else if (node['a:scrgbClr']) { const c = node['a:scrgbClr']; const to = (v)=>Math.round(parseInt(v)/100000*255); return { r:to(c['@r']), g:to(c['@g']), b:to(c['@b']), a:1 }; }
  else return null;
  let r = parseInt(hex.slice(0,2),16), g = parseInt(hex.slice(2,4),16), b = parseInt(hex.slice(4,6),16), a = 1;
  const num = (v)=>parseInt(v)/100000; // ST_Percentage : 100000 = 100 %
  let [h,s,l] = hslFromRgb(r,g,b);
  if (mods['a:lumMod']) l *= num(mods['a:lumMod']['@val']);
  if (mods['a:lumOff']) l += num(mods['a:lumOff']['@val']);
  l = Math.max(0, Math.min(1, l));
  // repasse en RVB uniquement si la luminance a changé (évite un aller-retour HSL -> arrondis)
  if (mods['a:lumMod'] || mods['a:lumOff']) [r,g,b] = rgbFromHsl(h,s,l);
  // shade k : c × k (k = 1 -> inchangé, 0 -> noir) ; tint k : c × k + 255 × (1 - k) (0 -> blanc)
  if (mods['a:shade']) { const k = num(mods['a:shade']['@val']); r*=k; g*=k; b*=k; }
  if (mods['a:tint']) { const k = num(mods['a:tint']['@val']); r = r*k + 255*(1-k); g = g*k + 255*(1-k); b = b*k + 255*(1-k); }
  if (mods['a:alpha']) a = num(mods['a:alpha']['@val']);
  return { r:Math.round(Math.max(0,Math.min(255,r))), g:Math.round(Math.max(0,Math.min(255,g))), b:Math.round(Math.max(0,Math.min(255,b))), a };
}
// cssColor({r,g,b,a}) -> 'rgb(…)' ou 'rgba(…)' si transparente ; null -> 'none' (pratique pour fill=)
const cssColor = (c) => c ? (c.a < 1 ? `rgba(${c.r},${c.g},${c.b},${c.a.toFixed(3)})` : `rgb(${c.r},${c.g},${c.b})`) : 'none';

// ████ §L ─ GÉOMÉTRIES (prstGeom / custGeom) -> CHEMINS SVG ─────────────────────────────────────
// Une forme a soit une géométrie PRÉDÉFINIE (a:prstGeom prst="roundRect" + a:avLst d'ajustements,
// §20.1.9.18), soit une géométrie LIBRE (a:custGeom, §20.1.9.8) : guides (a:gdLst, formules
// « */ w 1 2 »…), chemins (a:pathLst > a:path w h fill stroke) et commandes moveTo / lnTo / arcTo /
// cubicBezTo / quadBezTo / close. Les deux sont évaluées par preset-geom.mjs à la taille réelle de la
// forme (w, h en POINTS ici) ; le résultat est un chemin « d » en coordonnées locales (0,0 = coin
// haut-gauche de la forme).
// Variante *Parts : sépare le chemin de REMPLISSAGE (sous-chemins fill != none) et le chemin de TRAIT
// (sous-chemins stroke != false) — ex. arc : secteur rempli, mais seule la courbe est tracée.
// Les ombrages 3D (fill darken*/lighten*) sont exclus du remplissage. La variante simple
// (presetGeomD / evalGeom) ne renvoie que la silhouette remplissable.
// ================= géométrie preset -> chemin SVG =================
// Délègue au moteur générique preset-geom.mjs (187 formes officielles ECMA-376 §20.1.9,
// formules de guides évaluées). adjO = map { nomGuide: fmla } depuis l'avLst de la forme.
// Utilisée par renderPic (forme de découpe) ; renderSp passe par presetGeomParts.
function presetPath(prst, w, h, adjO) {
  const d = presetGeomD(prst, w, h, adjO);
  return d ?? `M0 0 H${w} V${h} H0 Z`;   // nom inconnu -> rectangle
}

// custGeom (nœud ORDONNÉ preserveOrder) -> chemin SVG, avec évaluation des formules de guides.
// Construit un def { av, gd, paths:[{w,h,fill,stroke,cmds}] } et réutilise evalGeom.
// Balise DrawingML -> lettre de commande interne (proche de SVG, mais 'A' n'a PAS la sémantique SVG :
// arcTo = rayons wR/hR + angle de départ stAng + balayage swAng, en 1/60000 de degré ; c'est evalGeom
// qui le convertit en arc SVG).
const CUST_CMD = { 'a:moveTo':'M', 'a:lnTo':'L', 'a:arcTo':'A', 'a:cubicBezTo':'C', 'a:quadBezTo':'Q', 'a:close':'Z' };
// custGeomDef(nœud custGeom ORDONNÉ) -> définition brute (formules et coordonnées NON évaluées,
// encore sous forme de chaînes : une coordonnée peut être un nom de guide).
// L'arbre ordonné est obligatoire : l'ordre des commandes de chemin est le dessin lui-même.
// path.w / path.h : repère propre du chemin (ses coordonnées sont mises à l'échelle w×h de la forme) ;
// 0 = « pas de repère propre » (interprétation laissée à evalGeom).
function custGeomDef(custO) {
  const gdOf = (listTag) => {
    const lst = first(custO, listTag); if (!lst) return [];
    return byTag(lst, 'a:gd').map((g) => [at(g)['@name'], at(g)['@fmla']]);
  };
  const def = { av: gdOf('a:avLst'), gd: gdOf('a:gdLst'), paths: [] };
  const pathLst = first(custO, 'a:pathLst');
  for (const p of (pathLst ? byTag(pathLst, 'a:path') : [])) {
    const pa = at(p);
    const path = { w: +pa['@w'] || 0, h: +pa['@h'] || 0,
      fill: pa['@fill'] ?? 'norm', stroke: !['false', '0'].includes(pa['@stroke']), cmds: [] }; // xsd:boolean : 0/false
    for (const seg of kids(p)) {
      const c = CUST_CMD[tagOf(seg)]; if (!c) continue;
      if (c === 'Z') { path.cmds.push(['Z']); continue; }
      if (c === 'A') { const a = at(seg); path.cmds.push(['A', a['@wR'], a['@hR'], a['@stAng'], a['@swAng']]); continue; }
      // moveTo/lnTo : 1 point ; quadBezTo : 2 ; cubicBezTo : 3 -> aplatis [c, x1, y1, x2, y2, …]
      const pts = byTag(seg, 'a:pt').map((pt) => [at(pt)['@x'], at(pt)['@y']]);
      path.cmds.push([c, ...pts.flat()]);
    }
    def.paths.push(path);
  }
  return def;
}
// custGeomPath -> un seul « d » (tous sous-chemins) : forme de découpe des images (renderPic)
function custGeomPath(custO, w, h) {
  return evalGeom(custGeomDef(custO), w, h);
}
// custGeomParts -> { fill, stroke } séparés : formes (renderSp)
function custGeomParts(custO, w, h) {
  return evalGeomParts(custGeomDef(custO), w, h);
}

// ████ §M ─ RELATIONS (.rels) ET IMAGES ─────────────────────────────────────────────────────────
// Chaque part OOXML référence les autres (layout, master, thème, images, graphiques…) par un
// identifiant r:id / r:embed (« rId3 ») résolu dans son fichier _rels/<part>.rels (OPC, ECMA-376
// Part 2). Les cibles sont RELATIVES au dossier de la part (souvent « ../media/image1.png »).
// Ces fonctions renvoient une simple table { rId: Target } (les @Type ne sont pas lus : on reconnaît
// la cible à son chemin, ex. includes('slideLayout')). Les relations externes (TargetMode="External")
// ne sont pas distinguées.
// ================= rels & images =================
// slideRels(pptDir, n) -> rels de slides/slideN.xml ({} si pas de .rels)
function slideRels(pptDir, slideNo) {
  const p = join(pptDir, 'slides', '_rels', `slide${slideNo}.xml.rels`);
  if (!existsSync(p)) return {};
  const r = cachedParse(p, 'flat');
  const out = {};
  const rl = r.Relationships?.Relationship;
  for (const x of (Array.isArray(rl) ? rl : rl ? [rl] : [])) out[x['@Id']] = x['@Target'];
  return out;
}
// relsOf(chemin absolu d'une part) -> ses rels, par la convention OPC <dossier>/_rels/<nom>.rels
function relsOf(path) {
  const d = dirname(path), b = path.split(/[\\/]/).pop();
  const p = join(d, '_rels', b + '.rels');
  if (!existsSync(p)) return {};
  const r = cachedParse(p, 'flat');
  const out = {}; const rl = r.Relationships?.Relationship;
  for (const x of (Array.isArray(rl) ? rl : rl ? [rl] : [])) out[x['@Id']] = x['@Target'];
  return out;
}
// Extensions servables dans <image href="data:…">. Valeur '' (emf/wmf) ou extension absente (tif,
// wdp…) -> non embarquable -> imageHref renvoie null (image non dessinée).
const MIME = { webp:'image/webp', png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif', bmp:'image/bmp', svg:'image/svg+xml', emf:'', wmf:'' };
// imageHref(pptDir, Target de rels, embed) -> URL de l'image ou null.
//   embed=true (défaut de renderSlide) : data URI base64 (SVG autonome, lourd) ;
//   embed=false : URL file:/// absolue (aperçus locaux uniquement).
// La cible est résolue depuis ppt/ en retirant UN « ../ » de tête : correct pour les cibles des
// slides / layouts / masters (qui sont à un niveau sous ppt/), pas pour une cible sans « ../ »
// relative à un autre dossier.
function imageHref(pptDir, target, embed) {
  let abs = join(pptDir, target.replace(/^\.\.\//,''));
  if (!existsSync(abs)) return null;
  let ext = abs.split('.').pop().toLowerCase();
  // EMF/WMF (vecteur Windows, non affichable en <image>) : servir le PNG pré-converti « <fichier>.png »
  // (préprocessing convert-metafiles.ps1 via .NET/WIC). Beaucoup de diagrammes techniques sont en EMF.
  if (ext === 'emf' || ext === 'wmf') {
    const png = abs + '.png';
    if (existsSync(png)) { abs = png; ext = 'png'; } else return null;
  }
  if (!embed) return 'file:///' + abs.replace(/\\/g,'/');
  // version WebP allégée écrite à côté par mediaOptimize.js (si elle valait le coup)
  if (existsSync(abs + '.webp')) { abs += '.webp'; ext = 'webp'; }
  const mime = MIME[ext];
  if (!mime) return null;
  return `data:${mime};base64,` + readFileSync(abs).toString('base64');
}

// ████ §N ─ PLACEHOLDERS (HÉRITAGE DEPUIS LE LAYOUT) ────────────────────────────────────────────
// Règle PresentationML (p:ph, §19.3.1.36) : une forme <p:sp><p:nvSpPr><p:nvPr><p:ph type idx/> de la
// diapo est l'INSTANCE d'un placeholder du layout (lui-même instance de celui du master). Tout ce que
// la diapo ne redéfinit pas (position/taille, géométrie, fill, ligne, styles de texte, bodyPr) vient
// du niveau au-dessus. Ici, l'appariement se fait par clé exacte "type:idx" (@type absent -> 'body',
// @idx absent -> '0', des DEUX côtés, donc cohérent). NB : la spec donne 'obj' comme type par défaut
// et PowerPoint apparie surtout par idx ; la clé exacte suffit pour le corpus.
// ================= placeholders (héritage layout/master) =================
// Une forme-texte de la diapo hérite géométrie + fill + styles de texte de SON placeholder
// de layout (identifié par type:idx). Le layout hérite lui-même du masque.
// placeholderGeom(pptDir, n, rels de la diapo) -> { geom, mgeom } : geom = { "type:idx": { xfrm, spPr,
// style, lstStyle, bodyPr } } lu du LAYOUT ; mgeom = { "type:idx": { lstStyle, bodyPr } } lu du MASTER
// (STYLE DE TEXTE seulement, pour la cascade : c'est le master qui porte le plus souvent la police/la
// taille du titre et du corps). Tout en arbre PLAT, mis en cache. Seuls les p:sp de 1er niveau du spTree
// sont indexés (pas les p:pic ni les placeholders dans des groupes). La GÉOMÉTRIE n'est pas reprise du
// master : un placeholder de layout sans xfrm ne reçoit pas celui du master.
function placeholderGeom(pptDir, slideNo, rels) {
  const layoutTgt = Object.values(rels).find((t) => t && t.includes('slideLayout'));
  const geom = {};
  if (!layoutTgt) return { geom };
  const layoutPath = join(pptDir, 'slides', layoutTgt);
  if (!existsSync(layoutPath)) return { geom };
  const lx = cachedParse(layoutPath, 'flat');
  const tree = lx['p:sldLayout']?.['p:cSld']?.['p:spTree'];
  const arr = (x)=> x===undefined?[]:Array.isArray(x)?x:[x];
  for (const sp of arr(tree?.['p:sp'])) {
    const ph = sp['p:nvSpPr']?.['p:nvPr']?.['p:ph']; if (!ph) continue;
    const key = `${ph['@type']??'body'}:${ph['@idx']??'0'}`;
    // le layout a plusieurs placeholders "body:0" -> on garde le premier par clé unique,
    // mais les cartes ont des idx distincts (27,30,33...) donc pas de collision gênante
    if (!geom[key]) geom[key] = { xfrm: sp['p:spPr']?.['a:xfrm'], spPr: sp['p:spPr'], style: sp['p:style'],
      lstStyle: sp['p:txBody']?.['a:lstStyle'], bodyPr: sp['p:txBody']?.['a:bodyPr'] };
  }
  // MAILLON MANQUANT de la cascade : les placeholders du MASTER (même clé « type:idx »). C'est le master
  // qui porte le plus souvent la POLICE et la TAILLE du titre / du corps (le layout ne les redéfinit pas).
  // Sans ce maillon, un titre héritait de p:txStyles (titleStyle = défaut du fichier) : police et taille
  // fausses (vu : titre « Arial 14 » au lieu du « Space Grotesk 30 » du master). Le master est joint au
  // layout (« ../slideMasters/slideMasterN.xml »). On n'indexe QUE le style de texte (lstStyle), pas la
  // géométrie (cf. commentaire ci-dessus).
  const mgeom = {};
  const mTgt = Object.values(relsOf(layoutPath)).find((t) => t && t.includes('slideMaster'));
  if (mTgt) {
    const mPath = join(dirname(layoutPath), mTgt);
    if (existsSync(mPath)) {
      const mt = cachedParse(mPath, 'flat')['p:sldMaster']?.['p:cSld']?.['p:spTree'];
      for (const sp of arr(mt?.['p:sp'])) {
        const ph = sp['p:nvSpPr']?.['p:nvPr']?.['p:ph']; if (!ph) continue;
        const key = `${ph['@type']??'body'}:${ph['@idx']??'0'}`;
        if (!mgeom[key]) mgeom[key] = { lstStyle: sp['p:txBody']?.['a:lstStyle'], bodyPr: sp['p:txBody']?.['a:bodyPr'] };
      }
    }
  }
  return { geom, mgeom };
}

// ================= rendu d'une forme =================
// xfrmAttrs : VESTIGE non appelé (ne renvoie que x). La lecture réelle des a:xfrm se fait dans
// getXfrm (renderSlide). Conservée telle quelle.
function xfrmAttrs(xfrm) {
  if (!xfrm) return null;
  return {
    x: Number(xfrm['@x'] ?? at(deepFind(xfrm,'a:off'))['@x']),
  };
}

// ████ §O ─ renderSlide : RENDU D'UNE DIAPO ─────────────────────────────────────────────────────
// renderSlide(pptDir, slideNo, { embedImages }) -> { svg, w, h }
//   pptDir  : dossier « ppt » du .pptx décompressé ; slideNo : N de slides/slideN.xml (≠ ordre
//             d'affichage de presentation.xml, qui n'est pas lu ici) ;
//   embedImages (défaut true) : images en data URI (sinon file:///) ;
//   w, h    : taille de diapo en EMU (le viewBox du svg, lui, est en points).
// Organisation : une GRANDE fonction dont les sous-fonctions (walk, renderSp, renderText…) partagent
// par fermeture l'état de la diapo : ctx, body (liste des fragments SVG dans l'ordre z), curRels,
// skipPh, spidPfx, et les collections de <defs> (grads, patterns, clips, filters, lineMarkers).
// Les sous-sections §O.1 à §O.8 suivent l'ordre du code ; la composition réelle est à la fin (§O.8).
// version preserveOrder: on parse le spPr flat pour lire fills/geom facilement
export function renderSlide(pptDir, slideNo, opts = {}) {
  const embed = opts.embedImages !== false;
  // préfixe d'ID UNIQUE par diapo : clip/gradient/marqueur/filtre. Sans ça, plusieurs diapos dans UNE
  // page (viewer.html) partagent des IDs (clip0, g0…) qui collisionnent -> mauvais clip/dégradé (ex. image
  // circulaire rendue en petit rectangle dans le viewer mais OK dans le player qui isole les diapos).
  const uid = 's' + slideNo + '_';
  const rels = slideRels(pptDir, slideNo);
  // --- résout layout + masque + THÈME (le thème dépend du master : cf. loadContext) ---
  // norm : résout une cible de rels relativement au DOSSIER de la part source (« ../slideMasters/x »)
  const norm = (base, tgt) => join(dirname(base), tgt);
  // NB : le layout est joint à « slides/ » (et non via norm) car sa cible part de slides/ (« ../slideLayouts/… »)
  const layoutTgt = Object.values(rels).find(t => t && t.includes('slideLayout'));
  const layoutPath = layoutTgt ? join(pptDir, 'slides', layoutTgt) : null;
  const layoutRels = layoutPath ? relsOf(layoutPath) : {};
  const masterTgt = Object.values(layoutRels).find(t => t && t.includes('slideMaster'));
  const masterPath = (layoutPath && masterTgt) ? norm(layoutPath, masterTgt) : null;
  const masterRels = masterPath ? relsOf(masterPath) : {};
  const themeTgt = Object.values(masterRels).find(t => t && t.includes('/theme/'));
  const themePath = (masterPath && themeTgt) ? norm(masterPath, themeTgt) : null;
  const ctx = loadContext(pptDir, themePath);
  _ctxMinorFont = ctx.minorFont || 'Calibri'; // pour la mesure des cellules de tableau (cellTextHeight)
  const { geom: phGeom, mgeom: mphGeom } = placeholderGeom(pptDir, slideNo, rels);

  // La diapo elle-même n'est PAS mise en cache (propre à cet appel) : parse ordonné direct.
  const slidePath = join(pptDir, 'slides', `slide${slideNo}.xml`);
  const ordered = P.parse(readFileSync(slidePath, 'utf8'));
  const sld = ordered.find((n) => tagOf(n) === 'p:sld');
  const cSld = first(sld, 'p:cSld');
  const spTree = first(cSld, 'p:spTree');
  // taille de diapo lue dans presentation.xml (gère 16:9, 4:3, etc.)
  // défaut = 16:9 « grand écran » de PowerPoint : 12 192 000 × 6 858 000 EMU = 13,333" × 7,5" = 960 × 540 pt
  let W = 12192000, H = 6858000;
  try {
    const pres = cachedParse(join(pptDir, 'presentation.xml'), 'flat');
    const sz = pres['p:presentation']?.['p:sldSz'];
    if (sz) { W = +sz['@cx'] || W; H = +sz['@cy'] || H; }
  } catch {}

  // body : fragments SVG dans l'ordre de dessin (fond, master, layout, diapo). renderGroup y retire
  // (splice) les fragments de ses membres pour les envelopper dans un <g>.
  const body = [];
  // Transformation « repère enfant -> repère diapo » (tf) passée à walk : x' = ox + (x - cx0) · sx,
  // idem en y. IDENT = identité (spTree de 1er niveau : les coordonnées sont déjà celles de la diapo).
  // Les groupes (renderGroup) et SmartArt (renderDiagram) en fabriquent d'autres ; inGroup:true
  // signale un membre de groupe.
  const IDENT = { ox:0, oy:0, sx:1, sy:1, cx0:0, cy0:0 };
  let curRels = rels;    // rels du tree en cours (pour résoudre les images)
  let skipPh = false;    // dans masque/layout : ne pas rendre le texte des placeholders
  // préfixe de data-spid : masque/layout et diapo ont chacun leur espace d'id -> collisions (plusieurs
  // formes "data-spid=8"). La timeline d'animation ne cible QUE la diapo ; on préfixe donc les spid du
  // masque/layout ('m'/'l') pour que le player (querySelector par spid) trouve la BONNE forme de diapo.
  let spidPfx = '';

  // (layout/masque/thème déjà résolus en tête de fonction)
  // clrMap du masque : remappe bg1/tx1/bg2/tx2 vers dk1/lt1/... (crucial pour thème sombre)
  // + txStyles : couleurs/tailles par défaut des placeholders (titre/corps) hérités par les runs
  // Structure produite : ctx.txStyles[title|body|other][niveau 1..9] = { sz (pt), b, color (CSS déjà
  // résolue — donc avec le clrMap du master), font (typeface brut, éventuellement « +mj-lt »), algn,
  // bu, marL, ind (EMU), lnSpc, spcBef, spcAft (nœuds plats bruts) } — lu depuis p:sldMaster/p:txStyles.
  ctx.txStyles = {};
  if (masterPath && existsSync(masterPath)) {
    const mflat = cachedParse(masterPath, 'flat');
    const cm = mflat['p:sldMaster']?.['p:clrMap'];
    // p:clrMap bg1="lt1" tx1="dk1" … -> { bg1:'lt1', tx1:'dk1', … }. NB : un p:clrMapOvr du layout
    // ou de la diapo (surcharge locale) n'est pas lu : c'est toujours la table du master qui s'applique.
    if (cm) { ctx.clrMap = {}; for (const [k, v] of Object.entries(cm)) if (k.startsWith('@')) ctx.clrMap[k.slice(1)] = v; }
    const ts = mflat['p:sldMaster']?.['p:txStyles'];
    if (ts) for (const [k, tag] of Object.entries({ title:'p:titleStyle', body:'p:bodyStyle', other:'p:otherStyle' })) {
      const st = ts[tag]; if (!st) continue; ctx.txStyles[k] = {};
      for (let lvl = 1; lvl <= 9; lvl++) {
        const pPr = st[`a:lvl${lvl}pPr`]; if (!pPr) continue;
        const d = pPr['a:defRPr'] || {};
        ctx.txStyles[k][lvl] = {
          sz: d['@sz'] ? +d['@sz']/100 : null, b: d['@b'] === '1',
          color: d['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null,
          font: d['a:latin']?.['@typeface'] || null,
          algn: pPr['@algn'] || null,
          bu: bulletOf(pPr), marL: pPr['@marL'] ? +pPr['@marL'] : null, ind: pPr['@indent'] ? +pPr['@indent'] : null,
          lnSpc: pPr['a:lnSpc'] || null,
          // espacement avant/après paragraphe hérité du master (§21.1.2.2.{10,12}) : les cartes/listes
          // du corpus tirent leur « gap entre paragraphes » d'ici (bodyStyle lvl1 spcBef 10pt), jamais
          // redéfini dans la diapo -> sans héritage, les paragraphes se collent (vu vs PowerPoint).
          spcBef: pPr['a:spcBef'] || null, spcAft: pPr['a:spcAft'] || null,
        };
      }
    }
  }
  // defaultTextStyle de presentation.xml : dernier recours pour le texte hors placeholder
  // (zones de texte libres) — §19.2.1.6. Chargé par niveau, même structure que txStyles.
  ctx.defaultTextStyle = {};
  try {
    const pf = cachedParse(join(pptDir, 'presentation.xml'), 'flat');
    const dts = pf['p:presentation']?.['p:defaultTextStyle'];
    if (dts) for (let lvl = 1; lvl <= 9; lvl++) {
      const pPr = dts[`a:lvl${lvl}pPr`]; if (!pPr) continue;
      const d = pPr['a:defRPr'] || {};
      ctx.defaultTextStyle[lvl] = {
        sz: d['@sz'] ? +d['@sz']/100 : null, b: d['@b'] === '1',
        color: d['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null,
        font: d['a:latin']?.['@typeface'] || null, algn: pPr['@algn'] || null,
      };
    }
  } catch {}

  // ── §O.1 ─ FOND, PARCOURS DU spTree, TRANSFORMATIONS, GROUPES ─────────────────────────────────
  // partOf(chemin, balise racine) -> { spTree, cSld } d'un master / layout (arbre ORDONNÉ, en cache,
  // donc en lecture seule), ou null si le fichier manque.
  function partOf(xmlPath, rootTag) {
    if (!xmlPath || !existsSync(xmlPath)) return null;
    const o = cachedParse(xmlPath, 'ord');
    const root = o.find(n => tagOf(n) === rootTag);
    const cS = first(root, 'p:cSld');
    return { spTree: first(cS, 'p:spTree'), cSld: cS };
  }
  // bgRect(cSld) -> fragment SVG du fond de CE niveau (diapo, layout ou master), '' s'il n'en définit
  // pas. p:bg contient soit p:bgPr (remplissage explicite), soit p:bgRef (référence au thème).
  // Dépend de curRels (le fond image est un r:embed du niveau concerné) : l'appelant le positionne.
  function bgRect(cS) {
    const bg = cS && first(cS, 'p:bg');
    if (!bg) return '';
    const bgFlat = PF.parse(`<root xmlns:a="a" xmlns:p="p">${serialize(bg)}</root>`).root['p:bg'];
    // bgPr : fill direct (solid/grad/blip)
    const bgPr = bgFlat['p:bgPr'];
    if (bgPr) {
      if (bgPr['a:blipFill']) {
        const embedId = bgPr['a:blipFill']['a:blip']?.['@r:embed'] ?? bgPr['a:blipFill']['a:blip']?.['@embed'];
        const href = curRels[embedId] ? imageHref(pptDir, curRels[embedId], embed) : null;
        if (href) return `<image x="0" y="0" width="${u(W)}" height="${u(H)}" href="${href}" preserveAspectRatio="xMidYMid slice"/>`;
      }
      const p = paint(bgPr);
      if (p) return `<rect width="${u(W)}" height="${u(H)}" fill="${p}"/>`;
    }
    // bgRef : idx>=1001 -> theme bgFillStyleLst[idx-1001], phClr = couleur donnée
    // (Convention d'index des style matrix refs : 0 = aucun, 1..999 = fillStyleLst, 1001+ =
    // bgFillStyleLst. Ici un idx < 1000 prend bgFillStyles[0] — approximation.)
    // Seuls solid/grad sont rendus via paint : un fond de thème en blipFill/pattFill est ignoré.
    const bgRef = bgFlat['p:bgRef'];
    if (bgRef) {
      const idx = +bgRef['@idx'];
      const phName = bgRef['a:schemeClr']?.['@val'] ?? bgRef['a:srgbClr']?.['@val'];
      const styleFill = ctx.bgFillStyles[idx >= 1000 ? idx - 1001 : 0];
      const p = paint(styleFill, phName);
      if (p) return `<rect width="${u(W)}" height="${u(H)}" fill="${p}"/>`;
    }
    return '';
  }

  // walk(conteneur, tf) : rend chaque enfant d'un spTree / grpSp dans l'ordre du document (= ordre z).
  // Les enfants non-formes (p:nvGrpSpPr, p:grpSpPr, extLst…) tombent dans aucun cas de walkChild.
  function walk(node, tf) {
    if (!node) return;
    for (const child of kids(node)) walkChild(child, tf);
  }
  // walkChild : aiguillage par type d'élément. Les fonctions de rendu RENVOIENT leur fragment (poussé
  // dans body ici), sauf renderGroup et renderDiagram qui poussent elles-mêmes (elles doivent
  // envelopper / insérer plusieurs fragments).
  function walkChild(child, tf) {
    const t = tagOf(child);
    // MCE (Part 3) : une forme peut être enveloppée dans <mc:AlternateContent> (ex. formules OMML
    // sous mc:Choice Requires="a14"). Sans le gérer, la forme + son texte sont INVISIBLES.
    // On rend la branche Fallback (repli sûr) si présente, sinon Choice.
    if (t === 'mc:AlternateContent') {
      // On préfère mc:Choice (contenu AUTHORÉ : texte + formules ; on rend au moins le texte
      // normal, le math OMML est ignoré). Le Fallback des formules est souvent une forme VIDE
      // -> il donnerait une boîte sans texte. Repli sur Fallback si pas de Choice.
      const branch = first(child, 'mc:Choice') || first(child, 'mc:Fallback');
      if (branch) for (const c of kids(branch)) walkChild(c, tf);
      return;
    }
    // masque/layout : on NE rend PAS les placeholders (la diapo en hérite fill/style).
    if (t === 'p:sp') { if (skipPh && deepFind(child, 'p:ph')) return; body.push(renderSp(child, tf)); }
    // image ESPACE RÉSERVÉ (<p:ph>) à l'intérieur d'un GROUPE : structure que l'interface PowerPoint
    // n'autorise pas ; PowerPoint en dessine les autres membres mais PAS l'image (vérifié par export
    // COM, 022_04 d.65) -> la rendre ajoutait une 2e copie décalée de l'illustration.
    else if (t === 'p:pic') { if (tf?.inGroup && deepFind(first(child, 'p:nvPicPr'), 'p:ph')) return; body.push(renderPic(child, tf)); }
    else if (t === 'p:cxnSp') body.push(renderSp(child, tf));
    else if (t === 'p:graphicFrame') body.push(renderTable(child, tf));
    else if (t === 'p:grpSp') renderGroup(child, tf);
  }
  // getXfrm(spPr ORDONNÉ) -> { x, y, cx, cy (EMU), rot (1/60000 °), flipH, flipV } ou null si la forme
  // n'a pas de a:xfrm complet (cas d'un placeholder qui hérite sa position du layout). §20.1.7.6
  function getXfrm(spPrOrdered) {
    const xf = first(spPrOrdered, 'a:xfrm');
    if (!xf) return null;
    const off = first(xf, 'a:off'), ext = first(xf, 'a:ext');
    if (!off || !ext) return null;
    return { x:+at(off)['@x'], y:+at(off)['@y'], cx:+at(ext)['@cx'], cy:+at(ext)['@cy'],
             rot:+(at(xf)['@rot']||0), flipH:at(xf)['@flipH']==='1', flipV:at(xf)['@flipV']==='1' };
  }
  // applyTf(boîte, tf) : passe une boîte du repère enfant (d'un groupe) au repère de la diapo. Seules
  // position et taille sont transformées : rot / flip restent ceux de la forme (la rotation / le
  // miroir d'un GROUPE est appliqué par un <g transform> englobant, voir renderGroup).
  function applyTf(g, tf) { return { x: tf.ox + (g.x - tf.cx0)*tf.sx, y: tf.oy + (g.y - tf.cy0)*tf.sy, cx: g.cx*tf.sx, cy: g.cy*tf.sy, rot:g.rot, flipH:g.flipH, flipV:g.flipV }; }

  // renderGroup(p:grpSp, tf) : un groupe DrawingML a son propre repère enfant. a:xfrm du grpSpPr
  // donne où le groupe est posé (off/ext, dans le repère PARENT) et quel rectangle de son repère
  // enfant cela représente (chOff/chExt). Donc, pour un membre : x_parent = off.x + (x - chOff.x) ·
  // ext.cx / chExt.cx (d'où childTf). Les groupes imbriqués composent ces transformations.
  // Pousse dans body un <g data-spid=id-du-groupe> contenant les fragments des membres.
  function renderGroup(grp, tf) {
    const gPr = first(grp, 'p:grpSpPr');
    const xf = gPr && first(gPr, 'a:xfrm');
    let childTf = tf, wrap = '';
    if (xf) {
      const off=first(xf,'a:off'),ext=first(xf,'a:ext'),chOff=first(xf,'a:chOff'),chExt=first(xf,'a:chExt');
      const g = applyTf({ x:+at(off)['@x'],y:+at(off)['@y'],cx:+at(ext)['@cx'],cy:+at(ext)['@cy'] }, tf);
      // chExt nul / absent -> sx/sy vaudraient Infinity ou NaN : on retombe sur 1 (aucune mise à
      // l'échelle), comme pour le mapping des SmartArt (cf. renderDiagram). Cas réel : un groupe
      // purement horizontal (chExt cy=0) contenant un connecteur droit -> sinon son y devient NaN.
      const cE = { cx:+at(chExt)['@cx'], cy:+at(chExt)['@cy'] };
      childTf = { ox:g.x, oy:g.y, cx0:+at(chOff)['@x'], cy0:+at(chOff)['@y'],
                  sx: cE.cx ? g.cx/cE.cx : 1, sy: cE.cy ? g.cy/cE.cy : 1, inGroup: true };
      // rotation / miroir du GROUPE (les membres sont placés dans la boîte NON tournée, puis tout
      // le <g> tourne autour du centre du groupe). Sans ça, un groupe tourné rend à plat
      // (ex. manches à air inclinées de VENT d.10).
      const rot = +at(xf)['@rot'] || 0, fH = at(xf)['@flipH']==='1', fV = at(xf)['@flipV']==='1';
      const ccx = u(g.x + g.cx/2), ccy = u(g.y + g.cy/2), parts = [];
      if (rot) parts.push(`rotate(${(rot/60000).toFixed(3)} ${ccx.toFixed(1)} ${ccy.toFixed(1)})`);
      // miroir autour du centre : translate(2·cx) scale(-1) ⇔ x -> 2cx - x (les membres étant en
      // coordonnées absolues de diapo, le centre du groupe sert de pivot)
      if (fH || fV) parts.push(`translate(${(fH?2*ccx:0).toFixed(1)} ${(fV?2*ccy:0).toFixed(1)}) scale(${fH?-1:1} ${fV?-1:1})`);
      wrap = parts.join(' ');
    }
    // remplissage du GROUPE (hérité par les membres <a:grpFill/>). On le résout et on le propage
    // pendant le walk des enfants, puis on restaure (gère les groupes imbriqués).
    const prevGrpFill = ctx.grpFill;
    if (gPr) {
      const gf = PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(gPr)}</root>`).root['p:grpSpPr'];
      if (gf?.['a:solidFill']) ctx.grpFill = cssColor(resolveColorFlat(gf['a:solidFill'], ctx));
      else if (gf?.['a:gradFill']) ctx.grpFill = gradId(gf['a:gradFill']);
      else if (gf?.['a:noFill'] !== undefined) ctx.grpFill = 'none';
      // sinon (pas de fill explicite ou grpFill imbriqué) : on garde le fill du groupe parent
    }
    // enveloppe les membres dans un <g data-spid=groupId> : le groupe s'anime EN BLOC
    // (une animation ciblant le groupe s'applique à tous ses membres, comme PowerPoint).
    // technique : on note la longueur de body, on rend les membres (qui s'y ajoutent), puis on
    // retire ces fragments (splice) pour les réinsérer à l'intérieur du <g>.
    const start = body.length;
    walk(grp, childTf);
    ctx.grpFill = prevGrpFill;
    const inner = body.splice(start).join('');
    const gid = at(deepFind(grp, 'p:cNvPr'))['@id'] || '';
    body.push(`<g data-spid="${spidPfx}${esc(gid)}"${wrap?` transform="${wrap}"`:''}>${inner}</g>`);
  }

  // ── §O.2 ─ REMPLISSAGES, DÉGRADÉS, MOTIFS, POINTILLÉS, POINTES DE FLÈCHE ──────────────────────
  // Ces fonctions renvoient une VALEUR d'attribut SVG (couleur CSS, 'none' ou 'url(#id)') et
  // enregistrent au passage les définitions nécessaires (grads, patterns, lineMarkers) qui seront
  // écrites dans <defs> à l'assemblage (§O.8).
  //
  // fillFromSpPrFlat({ 'p:spPr', … }, p:style plat) -> valeur fill. Priorité (§20.1.8, fills) :
  //   remplissage EXPLICITE du spPr (noFill, grpFill, solid, grad, patt ; blipFill -> null, géré par
  //   l'appelant) ; sinon référence de STYLE <p:style><a:fillRef idx> vers fillStyleLst du thème, la
  //   couleur du fillRef remplaçant phClr ; sinon 'none'.
  function fillFromSpPrFlat(spFlat, styleFlat) {
    const spPr = spFlat['p:spPr'];
    if (spPr?.['a:noFill'] !== undefined) return 'none';
    // <a:grpFill/> (§20.1.8.35) : la forme reprend le remplissage de son GROUPE parent (ex. régions
    // colorées « fusionnées » : le fill accent3/bg2 est sur le grpSpPr, les membres l'héritent).
    if (spPr?.['a:grpFill'] !== undefined) return ctx.grpFill ?? 'none';
    if (spPr?.['a:solidFill']) return cssColor(resolveColorFlat(spPr['a:solidFill'], ctx));
    if (spPr?.['a:gradFill']) return gradId(spPr['a:gradFill']);
    if (spPr?.['a:pattFill']) return pattId(spPr['a:pattFill']); // remplissage à MOTIF (hachures/points/grille)
    if (spPr?.['a:blipFill']) return null; // géré à part
    // style fillRef : idx (1-based) dans le fillStyleLst du thème, teinté par phClr
    const fr = styleFlat?.['a:fillRef'];
    if (fr) {
      const idx = +fr['@idx'];
      // idx=0 = AUCUN remplissage (la couleur du ref n'est PAS à utiliser) — ex. forme `arc`
      // ouverte (le 45° de VENT d.10) rendue en trait seul. Ne pas remplir.
      if (idx === 0) return 'none';
      const phName = fr['a:schemeClr']?.['@val'] ?? fr['a:srgbClr']?.['@val'];
      if (idx >= 1 && ctx.fillStyles[idx - 1]) { const p = paint(ctx.fillStyles[idx - 1], phName); if (p) return p; }
      const c = resolveColorFlat(fr, ctx); if (c) return cssColor(c); // fallback : couleur du ref
    }
    return 'none';
  }
  // remplissage à MOTIF (a:pattFill prst + fgClr/bgClr) -> tuile <pattern> SVG. §20.1.8.24.
  // Approxime par famille (points pct*, lignes horz/vert, diagonales, grille/croix) : recognizable
  // sans mapper les 54 presets exacts. Absent des decks -> gain pur pour PowerPoint externes.
  const patterns = []; // { id, fg, bg, prst } -> <pattern> en §O.8
  // pattId(a:pattFill plat) -> 'url(#…)' ; fgClr défaut noir, bgClr défaut blanc, prst défaut pct50
  function pattId(pf) {
    const id = uid + 'p' + patterns.length;
    const fg = pf['a:fgClr'] ? cssColor(resolveColorFlat(pf['a:fgClr'], ctx)) : '#000';
    const bg = pf['a:bgClr'] ? cssColor(resolveColorFlat(pf['a:bgClr'], ctx)) : '#fff';
    patterns.push({ id, fg, bg, prst: pf['@prst'] || 'pct50' });
    return `url(#${id})`;
  }
  const grads = []; // { id, stops, ang, radial } -> <linearGradient>/<radialGradient> en §O.8
  // gradId(a:gradFill plat) -> 'url(#…)'. a:gsLst/a:gs@pos en 1/1000 de % (/1000 -> %), couleur de
  // chaque stop résolue MAINTENANT (donc avec le phClr courant : important pour les styles de thème).
  // a:lin@ang (1/60000 °) -> radians ; a:path (circle/rect/shape) -> radial, rendu comme un cercle
  // centré quel que soit le type (a:fillToRect ignoré). @scaled, @rotWithShape, a:tileRect ignorés.
  function gradId(gradFlat) {
    const id = uid + 'g' + grads.length;
    const gs = gradFlat['a:gsLst']?.['a:gs'];
    // stops TRIÉS par position croissante : SVG l'exige (sinon Chrome clampe les offsets décroissants ->
    // dégradé faux, ex. sol de VENT d.35 rouge au lieu de vert). PowerPoint accepte l'ordre du document.
    const stops = (Array.isArray(gs)?gs:gs?[gs]:[]).map(s => ({ pos:(+s['@pos'])/1000, col: cssColor(resolveColorFlat(s, ctx)) }))
      .sort((a, b) => a.pos - b.pos);
    const ang = ((gradFlat['a:lin']?.['@ang']||0)/60000) * Math.PI/180;
    const radial = !!gradFlat['a:path']; // dégradé radial/rectangulaire
    grads.push({ id, stops, ang, radial });
    return `url(#${id})`;
  }
  // trait pointillé : prstDash -> stroke-dasharray (en points, relatif à l'épaisseur)
  // (ST_PresetLineDashVal : longueurs tiret/espace exprimées en multiples de
  // l'épaisseur, ex. dash = 4 plein / 3 vide, sysDot = 1 / 1). swp = épaisseur en pt. a:custDash non géré.
  function dashArray(val, swp) {
    if (!val || val === 'solid') return '';
    const s = swp || 1;
    return ({ dot:`${s} ${s*2}`, sysDot:`${s} ${s}`, dash:`${s*4} ${s*3}`, sysDash:`${s*3} ${s}`,
      lgDash:`${s*8} ${s*3}`, dashDot:`${s*4} ${s*3} ${s} ${s*3}`, lgDashDot:`${s*8} ${s*3} ${s} ${s*3}`,
      sysDashDot:`${s*3} ${s} ${s} ${s}`, lgDashDotDot:`${s*8} ${s*3} ${s} ${s*3} ${s} ${s*3}` }[val]) || '';
  }
  // clips d'images découpées en formes (cercle, roundRect, custGeom)
  const clips = []; // chaînes <clipPath …> prêtes à écrire dans <defs>
  // Pointes de flèche -> marqueurs SVG. Clé = type + dimensions (points remplacés par « p » pour
  // rester un identifiant valide) : deux traits de même type/taille partagent le même <marker>.
  // pointes de flèche des lignes/connecteurs (marker SVG partagé par type ; couleur = trait via
  // context-stroke ; orient=auto-start-reverse -> même marker pour début et fin).
  const lineMarkers = new Map();
  // dimensions en multiples de l'épaisseur du trait (markerUnits=strokeWidth) selon @len/@w (sm/med/lg).
  // PowerPoint « medium » ≈ 3-3,5× l'épaisseur (pas 8× — sinon pointes énormes, cf. flèches VENT d.28).
  // Taille calibrée sur PowerPoint (export COM, traits de 1 / 2,25 / 4,5 pt) : pointe = 2 / 3 / 5 x
  // l'épaisseur (sm / med / lg), avec une épaisseur PLANCHER de 2 pt (un trait de 1 pt a une pointe
  // moyenne de 6 pt, pas 3) ; longueur = largeur ; le bout de la pointe tombe sur l'extrémité du trait.
  // lineMarker(type, nœud headEnd/tailEnd, épaisseur pt) -> id du <marker>. type (ST_LineEndType) :
  // triangle, arrow, stealth, diamond, oval (arrow est dessiné comme triangle plein ; PowerPoint le
  // dessine en chevron ouvert).
  function lineMarker(type, end, swPt) {
    const SZ = { sm: 2, med: 3, lg: 5 };
    const k = swPt > 0 ? Math.max(1, 2 / swPt) : 1; // plancher 2 pt, en unités d'épaisseur de trait
    const w = (SZ[end?.['@len']] ?? 3) * k, h = (SZ[end?.['@w']] ?? 3) * k;   // len = longueur (X), w = largeur (Y)
    const key = (type + '_' + w + '_' + h).replace(/\./g, 'p');
    if (!lineMarkers.has(key)) {
      const cx = (w / 2).toFixed(2), cy = (h / 2).toFixed(2), W = w.toFixed(2), H = h.toFixed(2);
      let shape;
      if (type === 'oval') shape = `<ellipse cx="${cx}" cy="${cy}" rx="${(w/2*0.95).toFixed(2)}" ry="${(h/2*0.95).toFixed(2)}" fill="context-stroke"/>`;
      else if (type === 'diamond') shape = `<path d="M0 ${cy} L${cx} 0 L${W} ${cy} L${cx} ${H} z" fill="context-stroke"/>`;
      else if (type === 'stealth') shape = `<path d="M0 0 L${W} ${cy} L0 ${H} L${(w*0.32).toFixed(2)} ${cy} z" fill="context-stroke"/>`;
      else shape = `<path d="M0 ${(h*0.06).toFixed(2)} L${W} ${cy} L0 ${(h*0.94).toFixed(2)} z" fill="context-stroke"/>`; // triangle/arrow
      // refX = bout de la pointe (W) pour les flèches -> la pointe se termine sur l'extrémité du trait ;
      // centre pour oval/diamond (centrés sur l'extrémité). auto-start-reverse retourne le marqueur
      // en début de chemin : le même sert pour headEnd et tailEnd.
      lineMarkers.set(key, `<marker id="${uid}mk_${key}" markerWidth="${W}" markerHeight="${H}" refX="${type === 'oval' || type === 'diamond' ? cx : W}" refY="${cy}" orient="auto-start-reverse" markerUnits="strokeWidth">${shape}</marker>`);
    }
    return uid + 'mk_' + key;
  }
  // effets a:effectLst -> filtre SVG combiné (renvoie ' filter="url(#..)"'). §20.1.8
  // Gère : ombre portée (outerShdw), halo lumineux (glow), bords adoucis (softEdge).
  // ── §O.3 ─ EFFETS (OMBRE, HALO, BORDS ADOUCIS), FILTRES D'IMAGE, paint() ────────────────────
  // Principe : chaque effet DrawingML est traduit en primitives de filtre SVG (feGaussianBlur,
  // feOffset, feFlood, feComposite, feMorphology, feColorMatrix, feComponentTransfer) puis empilé par
  // feMerge (de l'arrière vers l'avant : ombre, halo ×2, forme). Un <filter> par forme concernée.
  const filters = []; // chaînes <filter …> (effets de forme ET filtres d'image) -> <defs>
  function shadowFilter(spPr, style) { return effectFilter(spPr, style); } // compat appel existant
  // effectFilter(spPr plat, p:style plat) -> ' filter="url(#…)"' ou ''.
  // Autres effets de a:effectLst non gérés : innerShdw, prstShdw, reflection, blur, fillOverlay ;
  // a:effectDag et scene3d/sp3d non plus.
  function effectFilter(spPr, style) {
    // effets EXPLICITES de la forme, sinon repli sur le style de thème via <p:style><a:effectRef idx>=1>
    // (§20.1.4.2.8) : l'ombre/lueur "thème" des formes stylées (SmartArt, cartes) vient de là.
    let fx = spPr?.['a:effectLst'];
    if (!fx) {
      const er = style?.['a:effectRef']; const eidx = er ? +er['@idx'] : 0;
      if (eidx >= 1) fx = ctx.effectStyles?.[eidx - 1]?.['a:effectStyle']?.['a:effectLst'];
    }
    if (!fx) return '';
    const sh = fx['a:outerShdw'], gl = fx['a:glow'], se = fx['a:softEdge'];
    if (!sh && !gl && !se) return '';
    // pad = débord de la zone de filtre en % de la boîte de la forme (SVG coupe tout ce qui sort de la
    // région du filtre : il faut la place pour le flou et le décalage de l'ombre)
    const prims = []; const layers = []; let pad = 30;
    // bords adoucis : flou du canal alpha puis re-découpe -> source adoucie ("src")
    let src = 'SourceGraphic';
    // @rad (EMU) -> pt, /2 : le rayon d'effet OOXML est traité comme ≈ 2 écarts-types de flou gaussien
    // (même conversion pour blurRad et glow ci-dessous)
    if (se) { const r = u(+se['@rad'] || 0) / 2;
      prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${r.toFixed(1)}" result="seA"/>`
        + `<feComposite in="SourceGraphic" in2="seA" operator="in" result="src"/>`);
      src = 'src';
    }
    // ombre portée (derrière)
    if (sh) {
      // dist (EMU) + dir (1/60000 °, 0 = vers la droite, sens horaire car y vers le bas) -> dx, dy.
      // blurRad défaut 40000 EMU ≈ 3,15 pt ; couleur absente -> noir 40 %. sx/sy/kx/ky/algn ignorés.
      const dist = u(+sh['@dist'] || 0), dir = (+sh['@dir'] || 0) / 60000 * Math.PI/180;
      const blur = u(+sh['@blurRad'] || 40000), c = resolveColorFlat(sh, ctx) || { r:0,g:0,b:0,a:0.4 };
      pad = Math.max(pad, 40 + blur);
      prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${(blur/2).toFixed(1)}" result="shB"/>`
        + `<feOffset in="shB" dx="${(dist*Math.cos(dir)).toFixed(1)}" dy="${(dist*Math.sin(dir)).toFixed(1)}" result="shO"/>`
        + `<feFlood flood-color="rgb(${c.r},${c.g},${c.b})" flood-opacity="${(c.a??0.4).toFixed(2)}"/>`
        + `<feComposite in2="shO" operator="in" result="shadow"/>`);
      layers.push('shadow');
    }
    // halo lumineux (derrière la source, devant l'ombre)
    if (gl) {
      // halo = silhouette légèrement dilatée (1 unité) puis floutée, colorée ; défaut jaune 60 %
      const r = u(+gl['@rad'] || 0) / 2, c = resolveColorFlat(gl, ctx) || { r:255,g:255,b:0,a:0.6 };
      pad = Math.max(pad, 40 + r * 2);
      prims.push(`<feMorphology in="SourceAlpha" operator="dilate" radius="1" result="glD"/>`
        + `<feGaussianBlur in="glD" stdDeviation="${r.toFixed(1)}" result="glB"/>`
        + `<feFlood flood-color="rgb(${c.r},${c.g},${c.b})" flood-opacity="${(c.a??0.6).toFixed(2)}"/>`
        + `<feComposite in2="glB" operator="in" result="glow"/>`);
      layers.push('glow', 'glow'); // doublé = halo plus dense (comme PowerPoint)
    }
    layers.push(src);
    const merge = `<feMerge>${layers.map((n) => `<feMergeNode in="${n}"/>`).join('')}</feMerge>`;
    // marge de la région de filtre bornée à [30 %, 200 %]. NB : pad mélange des pt (blur) et des % ;
    // c'est une heuristique, pas une conversion exacte.
    const m = Math.min(200, Math.max(30, Math.round(pad)));
    const id = uid + 'fx' + filters.length;
    filters.push(`<filter id="${id}" x="-${m}%" y="-${m}%" width="${100+2*m}%" height="${100+2*m}%">${prims.join('')}${merge}</filter>`);
    return ` filter="url(#${id})"`;
  }
  // effets couleur/alpha d'une image (enfants de a:blip / blipFill) -> filtre SVG. §20.1.8.{20,25,33,4}
  // imageFilter(a:blip plat) -> { attr: ' filter="url(#…)"' | '', opacity: ' opacity="…"' | '' }.
  // Les primitives s'enchaînent dans l'ordre du code (chaque primitive prend la sortie de la
  // précédente) : grayscl, lum, biLevel, duotone, puis corrections a14.
  function imageFilter(blip) {
    if (!blip) return { attr: '', opacity: '' };
    const prims = []; let opacity = '';
    const amf = blip['a:alphaModFix']; // opacité globale (ST_PositivePercentage)
    if (amf?.['@amt']) opacity = ` opacity="${(+amf['@amt']/100000).toFixed(3)}"`;
    if (blip['a:grayscl'] !== undefined) prims.push('<feColorMatrix type="saturate" values="0"/>');
    const lum = blip['a:lum']; // §20.1.8.42 luminance : bright (additif) + contrast (autour du gris moyen)
    if (lum && (lum['@bright'] || lum['@contrast'])) {
      // c' = c · (1 + contraste) + luminosité + 0,5 · (1 - pente) : le contraste pivote autour de 0,5
      // (gris moyen reste gris moyen), la luminosité décale le tout.
      const br = (+lum['@bright'] || 0) / 100000, ct = (+lum['@contrast'] || 0) / 100000;
      const slope = 1 + ct, inter = br + 0.5 * (1 - slope);
      const fn = (c) => `<feFunc${c} type="linear" slope="${slope.toFixed(3)}" intercept="${inter.toFixed(3)}"/>`;
      prims.push(`<feComponentTransfer>${fn('R')}${fn('G')}${fn('B')}</feComponentTransfer>`);
    }
    // noir et blanc à seuil (§20.1.8.11 biLevel) : luminance >= thresh -> blanc, sinon noir. Schémas
    // scannés : sans ça le fond gris du scan restait gris au lieu du blanc de PowerPoint.
    const bil = blip['a:biLevel'];
    if (bil) {
      // seuil en 1/1000 de % (défaut 50 %) ; table discrète de 100 marches (0 sous le seuil, 1 au-dessus)
      // appliquée à la luminance Rec. 709 (0,2126 R + 0,7152 G + 0,0722 B) recopiée sur les 3 canaux
      const t = Math.min(0.99, Math.max(0.01, (+bil['@thresh'] || 50000) / 100000)), n = 100;
      const tbl = Array.from({ length: n }, (_, i) => (i / n >= t ? 1 : 0)).join(' ');
      prims.push('<feColorMatrix type="matrix" values="0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0"/>'
        + `<feComponentTransfer><feFuncR type="discrete" tableValues="${tbl}"/><feFuncG type="discrete" tableValues="${tbl}"/><feFuncB type="discrete" tableValues="${tbl}"/></feComponentTransfer>`);
    }
    const duo = blip['a:duotone']; // bichromie : 2 couleurs de a:duotone -> remap luminance
    if (duo) {
      // Les deux couleurs sont prises dans l'ordre des CLÉS du parse plat : si elles sont de types
      // différents (ex. prstClr puis srgbClr), l'ordre du document n'est pas garanti. Chaque couleur est
      // re-emballée ({ [k]: c }) pour que resolveColorFlat lise aussi ses transformations enfants.
      // Moins de 2 couleurs -> repli noir -> blanc.
      const cols = [];
      for (const k of Object.keys(duo)) if (/Clr$/.test(k)) { const v = duo[k]; for (const c of (Array.isArray(v)?v:[v])) cols.push(resolveColorFlat({ [k]:c }, ctx)); }
      const [c0, c1] = cols.length >= 2 ? cols : [{r:0,g:0,b:0}, {r:255,g:255,b:255}];
      prims.push('<feColorMatrix type="matrix" values="0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0 0 0 1 0"/>'
        + `<feComponentTransfer><feFuncR type="table" tableValues="${(c0.r/255).toFixed(3)} ${(c1.r/255).toFixed(3)}"/>`
        + `<feFuncG type="table" tableValues="${(c0.g/255).toFixed(3)} ${(c1.g/255).toFixed(3)}"/>`
        + `<feFuncB type="table" tableValues="${(c0.b/255).toFixed(3)} ${(c1.b/255).toFixed(3)}"/></feComponentTransfer>`);
    }
    // corrections photo MS 2010 (extension a14 dans blip>extLst) : luminosité/contraste, saturation.
    // §MS-ODRAWXML a14:imgEffect. Valeurs en millièmes de % (ST_(Fixed)Percentage).
    // ⚠ Si un <a14:imgLayer r:embed=.wdp> est présent, ces corrections décrivent l'édition Office
    // appliquée à la HD photo ORIGINALE (.wdp) pour PRODUIRE l'image de base (rId2) que l'on rend déjà :
    // elles sont donc DÉJÀ cuites dans rId2. Les ré-appliquer écraserait la photo (ex. bright -40% →
    // soustraction de 102/canal → tout en noir). On ne les applique que si elles sont « directes ».
    const baked = !!deepFindFlat(blip, 'a14:imgLayer');
    const bc = baked ? null : deepFindFlat(blip, 'a14:brightnessContrast');
    if (bc && (bc['@bright'] || bc['@contrast'])) {
      const br = (+bc['@bright'] || 0) / 100000, ct = (+bc['@contrast'] || 0) / 100000;
      const slope = 1 + ct, inter = br + 0.5 * (1 - slope); // contraste autour du gris moyen + luminosité
      const fn = (c) => `<feFunc${c} type="linear" slope="${slope.toFixed(3)}" intercept="${inter.toFixed(3)}"/>`;
      prims.push(`<feComponentTransfer>${fn('R')}${fn('G')}${fn('B')}</feComponentTransfer>`);
    }
    const sat = baked ? null : deepFindFlat(blip, 'a14:saturation');
    if (sat && sat['@sat'] != null) prims.push(`<feColorMatrix type="saturate" values="${(+sat['@sat']/100000).toFixed(3)}"/>`);
    if (!prims.length) return { attr: '', opacity };
    const id = uid + 'img' + filters.length;
    // color-interpolation-filters="sRGB" : opérer en sRGB comme PowerPoint. Sinon les filtres SVG
    // tournent en linearRGB (défaut) → luminosité/contraste écrasent tout vers le noir.
    filters.push(`<filter id="${id}" color-interpolation-filters="sRGB">${prims.join('')}</filter>`);
    return { attr: ` filter="url(#${id})"`, opacity };
  }
  // peinture générique d'un noeud de remplissage (solid/grad) avec substitution phClr
  // phClr (« placeholder color ») : les styles du thème (fillStyleLst…) sont écrits avec
  // <a:schemeClr val="phClr"/> ; la couleur réelle est fournie par la forme qui y fait référence
  // (<a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef>). On la pose dans ctx.phClr le
  // temps de résoudre le remplissage, puis on restaure. Ne gère que solid/grad (pas blip/patt).
  function paint(fillFlat, phClrName) {
    if (!fillFlat) return null;
    const prev = ctx.phClr; ctx.phClr = phClrName || ctx.phClr;
    let out = null;
    if (fillFlat['a:solidFill']) out = cssColor(resolveColorFlat(fillFlat['a:solidFill'], ctx));
    else if (fillFlat['a:gradFill']) out = gradId(fillFlat['a:gradFill']);
    ctx.phClr = prev;
    return out;
  }

  // ── §O.4 ─ renderSp : FORME (p:sp) OU CONNECTEUR (p:cxnSp) ───────────────────────────────────
  // renderSp(nœud ORDONNÉ, tf, suppressText?) -> fragment SVG :
  //   <g data-spid transform="translate(x y) rotate(θ cx cy)" filter>   <- position + rotation
  //     <g transform="translate(…) scale(±1 ±1)">                       <- miroir (flipH/flipV)
  //       <path remplissage (+ trait)/> [<path trait seul/>] [image découpée]
  //     </g>
  //     <foreignObject texte/>                                         <- NON miroité (lisible)
  //   </g>
  // Étapes : parse plat -> héritage placeholder -> boîte (xfrm) -> géométrie -> fill -> trait ->
  // image de remplissage -> effets -> texte. Toutes les coordonnées internes sont en pt, relatives au
  // coin haut-gauche de la forme. suppressText : jamais passé par les appels actuels.
  function renderSp(sp, tf, suppressText) {
    // version flat du même noeud pour lire fills/geom/txBody
    const flat = PF.parse(`<root xmlns:a="a" xmlns:p="p">${serialize(sp)}</root>`).root;
    const spFlat = flat['p:sp'] || flat['p:cxnSp'] || flat;
    const spPrO = first(sp, 'p:spPr') || first(sp, 'p:spPr'); // (les deux membres sont identiques : redondance sans effet)
    // placeholder : héritage depuis le layout (géométrie + fill + styles de texte)
    const ph = deepFind(sp, 'p:ph');
    const key = ph ? `${at(ph)['@type']??'body'}:${at(ph)['@idx']??'0'}` : null;
    const inh = key ? phGeom[key] : null;
    // lstStyle du placeholder de MASTER (maillon entre le layout et p:txStyles du master)
    const inhM = key ? mphGeom?.[key] : null;
    // boîte : a:xfrm de la forme, sinon celui du placeholder de layout (sans flip hérité) ; aucune
    // boîte -> forme non dessinée
    let g = getXfrm(spPrO);
    if (!g && inh?.xfrm) { const xf=inh.xfrm, off=xf['a:off'], ext=xf['a:ext']; if (off&&ext) g = { x:+off['@x'],y:+off['@y'],cx:+ext['@cx'],cy:+ext['@cy'],rot:+(xf['@rot']||0) }; }
    if (!g) return '';
    const m = applyTf(g, tf);
    let geomFlat = spFlat['p:spPr'] || {};
    // géométrie héritée du placeholder de layout si la diapo n'en définit pas
    if (!geomFlat['a:prstGeom'] && !geomFlat['a:custGeom'] && inh?.spPr) geomFlat = inh.spPr;
    const prst = geomFlat['a:prstGeom']?.['@prst'];
    // custGeom : il faut le nœud ORDONNÉ (ordre des commandes), cherché dans la forme elle-même. Si
    // la custGeom vient du LAYOUT (geomFlat = inh.spPr), deepFind(sp) ne la trouve pas -> custO null
    // -> branche preset avec prst undefined -> rectangle.
    const custO = geomFlat['a:custGeom'] ? deepFind(sp, 'a:custGeom') : null;
    // d = chemin de remplissage, dStroke = chemin de trait, textRect = rectangle de texte du preset
    // (EMU, relatif à la forme ; les formes custGeom n'en ont pas ici -> boîte entière)
    let d, dStroke = '', textRect = null;
    if (custO) { const p = custGeomParts(custO, u(m.cx), u(m.cy)); d = p.fill; dStroke = p.stroke; }
    else {
      // ajustements (a:avLst/a:gd name="adj" fmla="val 16667") -> { adj: 'val 16667' } ; les valeurs
      // absentes prennent le défaut du preset (géré par preset-geom.mjs)
      const adj = {}; const gd = geomFlat['a:prstGeom']?.['a:avLst']?.['a:gd'];
      for (const x of (Array.isArray(gd)?gd:gd?[gd]:[])) adj[x['@name']] = x['@fmla'];
      const p = presetGeomParts(prst || 'rect', u(m.cx), u(m.cy), adj);
      if (p) { d = p.fill; dStroke = p.stroke; } else { d = `M0 0 H${u(m.cx)} V${u(m.cy)} H0 Z`; dStroke = d; }
      textRect = presetTextRect(prst || 'rect', m.cx, m.cy, adj); // en EMU, comme les insets
    }
    let fill = fillFromSpPrFlat(spFlat, spFlat['p:style']);
    // fill hérité du placeholder de layout (fond de carte, cercle de badge)
    // NB : fillFromSpPrFlat renvoie aussi 'none' pour un <a:noFill/> EXPLICITE de la diapo ; dans ce
    // cas le fill du layout est quand même repris (le test ne distingue pas les deux situations).
    if (fill === 'none' && inh) fill = fillFromSpPrFlat({ 'p:spPr': inh.spPr, 'p:style': inh.style }, inh.style);
    // Trait (a:ln, §20.1.2.2.24) : celui du spPr retenu (diapo, ou layout si la géométrie en vient),
    // sinon celui du placeholder de layout. Deux cas :
    //   1. a:ln présent sans a:noFill -> épaisseur/couleur/pointillés/jointures/pointes explicites,
    //      les manques étant complétés par le style de ligne du thème (lnRef) ;
    //   2. pas de a:ln et forme NON remplie -> trait entièrement issu du thème (lnRef).
    const ln = geomFlat['a:ln'] ?? inh?.spPr?.['a:ln'];
    let stroke = 'none', sw = 0, dash = '', markers = '', joinCap = '';
    if (ln && ln['a:noFill'] === undefined) {
      // épaisseur : @w, sinon celle du style de ligne du thème désigné par <a:lnRef idx> (défaut 1 pt)
      const lnRefIdx = +(spFlat['p:style']?.['a:lnRef']?.['@idx'] || 0);
      sw = Number(ln['@w']) || Number(ctx.lnStyles?.[lnRefIdx - 1]?.['a:ln']?.['@w']) || 12700;
      // trait en DÉGRADÉ (cartes à bordure estompée) : sans ça, repli sur la couleur lnRef (souvent
      // un accent assombri à 15 %) -> bordure quasi invisible.
      stroke = ln['a:solidFill'] ? cssColor(resolveColorFlat(ln['a:solidFill'], ctx))
             : ln['a:gradFill'] ? gradId(ln['a:gradFill'])
             : spFlat['p:style']?.['a:lnRef'] ? cssColor(resolveColorFlat(spFlat['p:style']['a:lnRef'], ctx)) : 'none';
      dash = dashArray(ln['a:prstDash']?.['@val'], u(sw));
      // jointure (§20.1.8.{35,42,49} round/bevel/miter) + extrémité (@cap rnd/sq/flat) du trait.
      // On n'émet QUE les valeurs NON-défaut SVG (défaut = miter/butt) : aucun bruit, blast radius =
      // uniquement les traits réellement à coins/bouts arrondis. Additif -> zéro régression ailleurs.
      const join = ln['a:round'] !== undefined ? 'round' : ln['a:bevel'] !== undefined ? 'bevel' : '';
      const cap = ln['@cap'] === 'rnd' ? 'round' : ln['@cap'] === 'sq' ? 'square' : '';
      if (join) joinCap += ` stroke-linejoin="${join}"`;
      if (cap) joinCap += ` stroke-linecap="${cap}"`;
      // pointes de flèche (§20.1.8.24/38) : headEnd = début du chemin (marker-start),
      // tailEnd = fin (marker-end). Marqueurs SVG partagés, couleur = trait (context-stroke).
      const endType = (e) => e && e['@type'] && e['@type'] !== 'none' ? e['@type'] : null;
      const he = endType(ln['a:headEnd']), te = endType(ln['a:tailEnd']);
      if (he) markers += ` marker-start="url(#${lineMarker(he, ln['a:headEnd'], u(sw))})"`;
      if (te) markers += ` marker-end="url(#${lineMarker(te, ln['a:tailEnd'], u(sw))})"`;
    } else if (!ln && fill === 'none') {
      // Pas de <a:ln> mais une ligne thémée via <p:style><a:lnRef idx>=1> (ex. le connecteur en angle =
      // carré d'angle droit de VENT d.26, invisible sinon). Scopé aux formes NON remplies (connecteurs/
      // lignes) pour ne pas ajouter de bordure parasite aux cartes/boîtes pleines (pas de régression).
      const lr = spFlat['p:style']?.['a:lnRef'];
      const idx = lr ? +lr['@idx'] : 0;
      if (lr && idx >= 1) {
        const themeLn = ctx.lnStyles[idx - 1]?.['a:ln'] || {};
        sw = Number(themeLn['@w']) || 12700;
        stroke = cssColor(resolveColorFlat(lr, ctx));
        dash = dashArray(themeLn['a:prstDash']?.['@val'], u(sw));
        const endType2 = (e) => e && e['@type'] && e['@type'] !== 'none' ? e['@type'] : null;
        const he = endType2(themeLn['a:headEnd']), te = endType2(themeLn['a:tailEnd']);
        if (he) markers += ` marker-start="url(#${lineMarker(he, themeLn['a:headEnd'], u(sw))})"`;
        if (te) markers += ` marker-end="url(#${lineMarker(te, themeLn['a:tailEnd'], u(sw))})"`;
      }
    }
    // forme remplie par une IMAGE (blipFill) : image découpée à la forme
    // (clip = chemin de remplissage d ; « slice » = couvrir la boîte en rognant, sans déformer ;
    // srcRect / stretch / tile de ce blipFill ne sont pas lus ici, contrairement à renderPic)
    let picLayer = '';
    const bf = spFlat['p:spPr']?.['a:blipFill'] || (fill === 'none' && inh?.spPr?.['a:blipFill']);
    if (bf?.['a:blip']) {
      const id = bf['a:blip']['@r:embed'] || bf['a:blip']['@embed'];
      const tgt = id && curRels[id];
      const href = tgt && imageHref(pptDir, tgt, embed);
      if (href) {
        const cid = uid + 'clip' + clips.length;
        clips.push(`<clipPath id="${cid}" clipPathUnits="userSpaceOnUse"><path d="${d}"/></clipPath>`);
        picLayer = `<g clip-path="url(#${cid})"><image x="0" y="0" width="${u(m.cx)}" height="${u(m.cy)}" href="${href}" preserveAspectRatio="xMidYMid slice"/></g>`;
        fill = 'none';
      }
    }
    const filt = shadowFilter(spFlat['p:spPr'], spFlat['p:style'] || inh?.style);
    // rotation autour du CENTRE de la forme (règle a:xfrm@rot : rotation autour du centre de la boîte,
    // sens horaire) ; miroir : translate(largeur) scale(-1 1) = retournement dans la boîte.
    // Le miroir est dans un <g> INTÉRIEUR qui ne contient pas le texte : le texte n'est jamais miroité
    // (conforme à PowerPoint pour flipH ; pour flipV, PowerPoint retourne le texte de 180°, pas ici).
    const rot = m.rot ? ` rotate(${(m.rot/60000).toFixed(2)} ${u(m.cx/2)} ${u(m.cy/2)})` : '';
    const sc = (m.flipH||m.flipV) ? ` translate(${u(m.flipH?m.cx:0)} ${u(m.flipV?m.cy:0)}) scale(${m.flipH?-1:1} ${m.flipV?-1:1})` : '';
    // phKind : title/body seulement pour un VRAI placeholder ; une zone de texte LIBRE (sans p:ph)
    // = 'none' -> n'hérite PAS du style body du master (sinon puces « • » et couleurs parasites) ;
    // elle s'appuie sur defaultTextStyle + couleur de thème.
    const phType = at(ph)['@type'];
    const phKind = !ph ? 'none' : (phType === 'title' || phType === 'ctrTitle') ? 'title' : 'body';
    // renderText reçoit w/h en EMU (m.cx, m.cy), le txBody plat ET ordonné, et les éléments hérités
    // du layout (lstStyle, bodyPr, style) pour la cascade de styles
    const txt = suppressText ? '' : renderText(spFlat['p:txBody'], m.cx, m.cy, phKind, inh?.lstStyle, first(sp, 'p:txBody'), inh?.bodyPr, spFlat['p:style'] || inh?.style, textRect, inhM?.lstStyle);
    const spid = at(deepFind(sp, 'p:cNvPr'))['@id'] || '';
    const strokeAttr = stroke !== 'none' ? `stroke="${stroke}" stroke-width="${u(sw)}"${dash?` stroke-dasharray="${dash}"`:''}${joinCap}${markers}` : '';
    // Tracé sur le même chemin que le remplissage (cas courant) OU chemin de tracé distinct (formes
    // arc/chord… : le remplissage est un secteur, le tracé n'est que la courbe -> pas de rayons parasites).
    // NB : si dStroke est VIDE (aucun sous-chemin à tracer, cf. evalGeomParts), sameGeom est vrai et le
    // trait est posé sur le chemin de remplissage.
    const sameGeom = !dStroke || dStroke === d;
    const paths = `<path d="${d}" fill="${fill}" ${sameGeom ? strokeAttr : ''}/>`
      + (!sameGeom && strokeAttr ? `<path d="${dStroke}" fill="none" ${strokeAttr}/>` : '');
    return `<g data-spid="${spidPfx}${esc(spid)}" transform="translate(${u(m.x)} ${u(m.y)})${rot}"${filt}>`
      + `<g transform="${sc.trim()||'translate(0 0)'}">`
      + paths
      + picLayer
      + `</g>` + txt + `</g>`;
  }

  // ── §O.5 ─ GRAPHIQUES, SmartArt, TABLEAUX (p:graphicFrame) ───────────────────────────────────
  // Un p:graphicFrame est un cadre générique dont a:graphic/a:graphicData porte un contenu typé :
  // tableau (a:tbl), graphique (c:chart ou cx:chart), SmartArt (dgm:relIds), objet OLE… Point
  // d'entrée unique : renderTable, qui aiguille vers renderChartFrame / renderDiagram.
  // Ces cadres n'ont ni rotation ni miroir ici (rot:0 forcé).
  //
  // graphique (p:graphicFrame > c:chart) : lit la part ppt/charts/chartN.xml (réf r:id) et la rend en SVG
  // via chart-render.mjs. Fond blanc (défaut PowerPoint) + palette = accents du thème.
  function renderChartFrame(chartRef, off, ext, tf) {
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
  function renderDiagram(relId, off, ext, tf) {
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
    walk(spTree, dtf);
  }
  // tableau (p:graphicFrame > a:tbl) : grille + bordures + fusions + texte
  // renderTable(p:graphicFrame ORDONNÉ, tf) -> fragment SVG (ou '' si délégué / vide).
  // Tableau (a:tbl, §21.1.3) : a:tblGrid/a:gridCol@w = largeurs de colonnes ; a:tr@h = hauteur
  // MINIMALE de ligne ; a:tc = cellule (gridSpan / rowSpan = fusion ; hMerge / vMerge = cellule
  // couverte par une fusion, à ne pas dessiner) ; a:tcPr = fond, bordures lnL/lnR/lnT/lnB, marges,
  // ancrage. Algorithme :
  //   1. hauteurs de lignes (agrandies selon le texte, cf. cellTextHeight), positions colX / rowY ;
  //   2. pour chaque cellule visible : couches du style de tableau -> fond, puis bordures, puis texte
  //      (renderText dans un repère translaté au coin de la cellule).
  function renderTable(gf, tf) {
    const arr = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];
    const flat = PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(gf)}</root>`).root['p:graphicFrame'];
    if (!flat) return '';
    const xf = flat['p:xfrm']; const off = xf?.['a:off'], ext = xf?.['a:ext'];
    const gd = flat['a:graphic']?.['a:graphicData'];
    // un p:graphicFrame porte un TABLEAU (a:tbl) OU un GRAPHIQUE (c:chart, réf. r:id vers ppt/charts/…)
    const chartRef = gd?.['c:chart'] || gd?.['cx:chart'];
    if (off && ext && chartRef) return renderChartFrame(chartRef, off, ext, tf);
    // SmartArt (dgm:relIds) : rendu de secours dsp:drawing référencé par dsp:dataModelExt@relId
    const dgm = gd?.['dgm:relIds'];
    if (off && ext && dgm) {
      const exts = arr(dgm['a:extLst']?.['a:ext']);
      let relId; for (const e of exts) { const dm = e['dsp:dataModelExt']; if (dm?.['@relId']) { relId = dm['@relId']; break; } }
      if (relId) renderDiagram(relId, off, ext, tf); // pousse directement dans body
      return '';
    }
    const tbl = gd?.['a:tbl'];
    if (!off || !tbl) {
      // repli générique : un graphicFrame non-tableau/graphique/SmartArt (objet OLE incorporé, contrôle…)
      // embarque presque toujours une IMAGE DE SECOURS (p:pic). La rendre plutôt que laisser un vide.
      if (off && ext) { const picO = deepFind(gf, 'p:pic'); if (picO) return renderPic(picO, tf); }
      return '';
    }
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
          if (f) cellFill = f['a:noFill'] !== undefined ? null : paint(f);
          if (n['a:tcTxStyle']) txStyle = { ...(txStyle || {}), ...n['a:tcTxStyle'] };
        }
        // fond PROPRE de la cellule prioritaire
        if (tcPr['a:noFill'] !== undefined) cellFill = null;
        else if (tcPr['a:solidFill'] || tcPr['a:gradFill']) cellFill = paint(tcPr);
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
          const txt = renderText(tb, cw, ch, 'none', cellLst, cellsO[ci] ? first(cellsO[ci], 'a:txBody') : null);
          if (txt) out += `<g transform="translate(${u(cx)} ${u(cy)})">${txt}</g>`;
        }
      });
    });
    const gid = flat['p:nvGraphicFramePr']?.['p:cNvPr']?.['@id'] || '';
    return `<g data-spid="${spidPfx}${esc(gid)}">${out}</g>`;
  }

  // ── §O.6 ─ renderPic : IMAGE (p:pic) ──────────────────────────────────────────────────────────
  // renderPic(nœud ORDONNÉ, tf) -> <g data-spid transform clip filter><image …/></g>.
  // Étapes : boîte (xfrm propre ou du placeholder) -> choix de la source (SVG > raster) -> forme de
  // découpe (prstGeom / custGeom) -> mode d'ajustement (stretch / contain / cover) -> recadrage
  // srcRect -> filtres couleur (imageFilter) et effets (effectFilter).
  // Aussi appelée par renderTable pour l'image de secours d'un objet OLE.
  function renderPic(pic, tf) {
    // placeholder image : géométrie (forme de découpe) héritée du layout
    const ph = deepFind(pic, 'p:ph');
    const key = ph ? `${at(ph)['@type']??'body'}:${at(ph)['@idx']??'0'}` : null;
    const inh = key ? phGeom[key] : null;
    const spPrO = first(pic, 'p:spPr');
    let g = getXfrm(spPrO);
    if (!g && inh?.xfrm) { const xf=inh.xfrm, o=xf['a:off'], e=xf['a:ext']; if (o&&e) g={ x:+o['@x'],y:+o['@y'],cx:+e['@cx'],cy:+e['@cy'],rot:+(xf['@rot']||0) }; }
    if (!g) return '';
    const m = applyTf(g, tf);
    // priorité au SVG moderne (asvg:svgBlip), sinon le blip raster (png/jpeg/wdp)
    const blip = deepFind(pic, 'a:blip');
    const svgB = deepFindSuffix(pic, 'svgBlip');
    // ordre d'essai des rId ; le 1er dont imageHref aboutit gagne (un .wdp ou un EMF sans PNG
    // pré-converti échoue -> on passe au suivant)
    const candidates = [at(svgB)['@r:embed'], at(blip)['@r:embed'], at(blip)['@embed']].filter(Boolean);
    let href = null;
    for (const id of candidates) { const tgt = curRels[id]; if (tgt) { href = imageHref(pptDir, tgt, embed); if (href) break; } }
    if (!href) return ''; // image non exploitable (ex. .wdp seul) -> ne rien afficher
    // forme de découpe : géométrie propre du pic, sinon celle héritée du placeholder de layout
    const picFlat = PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(pic)}</root>`).root['p:pic'];
    let geomFlat = picFlat['p:spPr'] || {};
    if (!geomFlat['a:prstGeom'] && !geomFlat['a:custGeom'] && inh?.spPr) geomFlat = inh.spPr;
    const prst = geomFlat['a:prstGeom']?.['@prst'];
    const custFlat = geomFlat['a:custGeom'];
    const custO = custFlat ? deepFind(pic, 'a:custGeom') : null;
    // <a:stretch><a:fillRect/> (mode par défaut d'une image) = ÉTIRER pour remplir la forme, sans
    // letterbox (preserveAspectRatio="none") — sinon des bandes transparentes laissent voir ce qui est
    // dessous (ex. TEMP d.19 : le badge sous l'image transparaissait dans la bande du haut). Sans
    // stretch (tile/rare) : contain.
    const hasStretch = picFlat['p:blipFill']?.['a:stretch'] !== undefined;
    let clipAttr = '', fit = hasStretch ? 'none' : 'xMidYMid meet';
    // forme de découpe NON rectangulaire -> clipPath (en coordonnées locales de l'image, comme le
    // chemin : le clip est posé sur le <g> déjà translaté/tourné)
    if (custO || (prst && prst !== 'rect')) {
      let d;
      if (custO) d = custGeomPath(custO, u(m.cx), u(m.cy));
      else { const adj={}; const gd=geomFlat['a:prstGeom']?.['a:avLst']?.['a:gd'];
        for (const x of (Array.isArray(gd)?gd:gd?[gd]:[])) adj[x['@name']]=x['@fmla'];
        d = presetPath(prst, u(m.cx), u(m.cy), adj); }
      const cid = uid + 'clip' + clips.length;
      clips.push(`<clipPath id="${cid}"><path d="${d}"/></clipPath>`);
      clipAttr = ` clip-path="url(#${cid})"`;
      fit = 'xMidYMid slice'; // forme non-rect : remplir + clipper (cover)
    }
    const rot = m.rot ? ` rotate(${(m.rot/60000).toFixed(2)} ${u(m.cx/2)} ${u(m.cy/2)})` : '';
    // miroir flipH/flipV (renderPic l'ignorait -> images de flèches courbes flip+rot dans le mauvais sens,
    // ex. VENT d.14). Appliqué APRÈS rot dans la chaîne = AVANT la rotation sur le contenu (SVG right-to-
    // left), comme PowerPoint (miroir en repère local puis rotation). Le clip suit le même repère.
    const flip = (m.flipH || m.flipV) ? ` translate(${u(m.flipH?m.cx:0)} ${u(m.flipV?m.cy:0)}) scale(${m.flipH?-1:1} ${m.flipV?-1:1})` : '';
    const imgF = imageFilter(picFlat['p:blipFill']?.['a:blip']); // grayscl/duotone/alphaModFix
    const effF = effectFilter(picFlat['p:spPr'], picFlat['p:style']); // glow/ombre/softEdge (+ effectRef)
    const bw = u(m.cx), bh = u(m.cy);
    // recadrage source a:srcRect (§20.1.8.55) : l/t/r/b = fraction à RETIRER de chaque bord
    // (millièmes de %). On agrandit l'image et on la décale pour que la région visible
    // remplisse la boîte, puis on clippe (mapping exact -> preserveAspectRatio="none").
    let ix = 0, iy = 0, iw = bw, ih = bh, par = fit;
    const sr = picFlat['p:blipFill']?.['a:srcRect'];
    // srcRect (§20.1.8.55) : l/t/r/b = fraction (millièmes de %) retirée de chaque bord de la SOURCE
    // avant remplissage. POSITIF = recadrage (zoom) ; NÉGATIF = marge hors-image (letterbox) → sert à
    // faire tenir une photo paysage dans une forme carrée SANS déformation (ex. photo d'intro : b=-54251
    // rend la région source ~carrée, l'image occupe le haut ~65 % de la boîte, non déformée). On mappe
    // la région source (agrandie+décalée) exactement sur la boîte -> par="none".
    if (sr && (sr['@l'] || sr['@t'] || sr['@r'] || sr['@b'])) {
      const l=(+sr['@l']||0)/100000, t=(+sr['@t']||0)/100000, r=(+sr['@r']||0)/100000, b=(+sr['@b']||0)/100000;
      // vw, vh = fraction de l'image source qui reste visible (bornée > 0 pour éviter /0)
      const vw = Math.max(1e-3, 1-l-r), vh = Math.max(1e-3, 1-t-b);
      // image entière dessinée à la taille boîte / fraction visible, décalée de la part rognée à
      // gauche / en haut : la région visible tombe exactement sur la boîte (0,0)-(bw,bh)
      iw = bw/vw; ih = bh/vh; ix = -l*iw; iy = -t*ih; par = 'none';
      // sans forme de découpe, un clip rectangulaire de la boîte masque ce qui dépasse
      if (!clipAttr) { const cid=uid+'clip'+clips.length;
        clips.push(`<clipPath id="${cid}"><rect x="0" y="0" width="${bw.toFixed(1)}" height="${bh.toFixed(1)}"/></clipPath>`);
        clipAttr = ` clip-path="url(#${cid})"`; }
    }
    const spid = at(deepFind(pic, 'p:cNvPr'))['@id'] || '';
    return `<g data-spid="${spidPfx}${esc(spid)}" transform="translate(${u(m.x)} ${u(m.y)})${rot}${flip}"${clipAttr}${effF}>`
      + `<image x="${ix.toFixed(1)}" y="${iy.toFixed(1)}" width="${iw.toFixed(1)}" height="${ih.toFixed(1)}" href="${href}" preserveAspectRatio="${par}"${imgF.attr}${imgF.opacity}/></g>`;
  }

  // ── §O.7 ─ renderText : TEXTE (a:txBody -> foreignObject HTML) ──────────────────────────────
  // renderText(txBody plat, w, h (EMU), phKind, inhLst, txBodyO, inhBodyPr, shapeStyle, textRect)
  //   txBody     : a:txBody / p:txBody en arbre PLAT (propriétés) ;
  //   txBodyO    : le même en arbre ORDONNÉ (ordre a:r / a:br / a:fld / formules), peut être null ;
  //   phKind     : 'title' | 'body' | 'none' -> quel p:txStyles du master s'applique ('none' = aucun) ;
  //   inhLst / inhBodyPr : lstStyle / bodyPr hérités (placeholder de layout ou style de tableau) ;
  //   shapeStyle : p:style de la forme (fontRef = police/couleur de thème par défaut) ;
  //   textRect   : rectangle de texte du preset (EMU) ou null.
  // Sortie : '' si aucun texte visible, sinon <foreignObject> (repère local de la forme, en pt)
  // contenant <div class="af"> (flex colonne = ancrage vertical) > <div data-par> par paragraphe >
  // <span> par run.
  //
  // CASCADE DES PROPRIÉTÉS D'UN RUN (du plus fort au plus faible), reproduisant PowerPoint :
  //   a:rPr du run  >  a:pPr du paragraphe (algn, marL, indent, puces, lnSpc…)
  //   >  a:lstStyle de la forme (niveau lvl)  >  lstStyle du placeholder de layout (ou style de tableau)
  //   >  lstStyle du placeholder de MASTER  >  p:txStyles du master (titleStyle / bodyStyle selon phKind)
  //   >  p:defaultTextStyle de
  //   presentation.xml  >  p:style/a:fontRef de la forme  >  couleur tx1 / taille 18 pt / police mineure.
  // Le niveau lvl d'un paragraphe = a:pPr@lvl + 1 (lvl 0..8 -> lvl1pPr..lvl9pPr).
  //
  // CORRECTIONS « À LA POWERPOINT » (le navigateur fait la mise en page, on corrige ses écarts) :
  //   interligne simple 1,2 × taille ; ligne de base (baselineShift, via VA_TOK) ; chasses arrondies au
  //   1/8 pt (advanceFix -> letter-spacing) ; runs en line-height:0 ; tabulations calculées en pt ;
  //   margin-right 0,2 px (règle de coupure) ; spcBef du 1er et spcAft du dernier paragraphe ignorés ;
  //   « tailCut » sous la dernière ligne ; espaces de fin de paragraphe supprimés.
  //
  // Autofit : a:normAutofit@fontScale / @lnSpcReduction (valeurs CALCULÉES et ENREGISTRÉES par
  // PowerPoint) sont appliquées telles quelles ; a:spAutoFit (forme ajustée au texte) n'a pas de
  // traitement particulier ; a:noAutofit -> débordement libre.
  function renderText(txBody, w, h, phKind, inhLst, txBodyO, inhBodyPr, shapeStyle, textRect, mstLst) {
    if (!txBody) return '';
    const arr = (x)=> x===undefined?[]:Array.isArray(x)?x:[x];
    const bodyPr = txBody['a:bodyPr'] || {};
    // texte VERTICAL (§21.1.2.1.1 @vert) : vert = tourné de 90° (lecture de haut en bas), vert270 = -90°
    // (de bas en haut). On compose dans un cadre aux dimensions permutées puis on le fait pivoter autour
    // du centre. Combiné à la rotation de la forme : rot 90° + vert270 = texte à l'endroit (040_02_01 d.32,
    // rendu tête en bas avant).
    const vert = bodyPr['@vert'] ?? inhBodyPr?.['@vert'];
    if (vert === 'vert' || vert === 'vert270' || vert === 'eaVert') {
      // appel RÉCURSIF avec w/h permutés et vert forcé à 'horz' (copies superficielles : l'arbre
      // d'origine n'est pas modifié). textRect n'est pas repassé (il est exprimé dans le repère non tourné).
      // eaVert (texte vertical est-asiatique, caractères empilés) est approximé par une rotation de 90°.
      const flat = { ...txBody, 'a:bodyPr': { ...bodyPr, '@vert': 'horz' } };
      const inner = renderText(flat, h, w, phKind, inhLst, txBodyO, { ...(inhBodyPr || {}), '@vert': 'horz' }, shapeStyle, null, mstLst);
      if (!inner) return '';
      const ang = vert === 'vert270' ? 270 : 90;
      return `<g transform="translate(${u(w/2)} ${u(h/2)}) rotate(${ang}) translate(${u(-h/2)} ${u(-w/2)})">${inner}</g>`;
    }
    // ── Boîte de texte : insets, ancrage, retour à la ligne ──
    // bodyPr hérité du placeholder de LAYOUT (insets + ancre) : une carte-placeholder avec
    // <a:bodyPr/> vide DOIT reprendre le tIns/anchor du layout (ex. tIns=288000 réserve la place
    // de la pastille). Priorité : valeur de la diapo → valeur héritée → défaut OOXML.
    const inhBP = inhBodyPr || {};
    const pick = (k, def) => bodyPr[k] !== undefined ? +bodyPr[k] : (inhBP[k] !== undefined ? +inhBP[k] : def);
    const anchor = bodyPr['@anchor'] ?? inhBP['@anchor']; // t/ctr/b
    // ancrage vertical (a:bodyPr@anchor t / ctr / b ; « just » et « dist » -> haut) -> justify-content
    // du conteneur flex en colonne
    const just = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
    // insets (marges intérieures) en EMU ; défauts OOXML : 0,1" (91440) à gauche/droite, 0,05" (45720)
    // en haut/bas
    let iL = pick('@lIns',91440), iR=pick('@rIns',91440), iT=pick('@tIns',45720), iB=pick('@bIns',45720);
    // insets plus grands que la forme (pastilles « 1 » de 19pt avec 11pt de marge de chaque côté) :
    // PowerPoint garde le CENTRE de la zone de texte. Sans ça la zone se colle à l'inset gauche/haut
    // (texte décalé en bas à droite), voire prend une largeur négative (foreignObject non rendu).
    // rectangle de texte de la forme (roundRect, ellipse…) : les insets s'appliquent À L'INTÉRIEUR
    // textRect = {l,t,r,b} en EMU depuis le coin de la forme ; r/b sont des COORDONNÉES, d'où w - r
    if (textRect) { iL += textRect.l; iT += textRect.t; iR += w - textRect.r; iB += h - textRect.b; }
    // insets trop grands : on ramène la zone à une largeur/hauteur nulle placée au milieu des insets
    if (iL + iR > w) { const c = (iL + w - iR) / 2; iL = c; iR = w - c; }
    if (iT + iB > h) { const c = (iT + h - iB) / 2; iT = c; iB = h - c; }
    // anchorCtr : le BLOC de texte est centré horizontalement (l'algn joue à l'intérieur du bloc).
    const ctrCss = (bodyPr['@anchorCtr'] ?? inhBP['@anchorCtr']) === '1' ? 'align-items:center;' : '';
    // wrap="none" : aucune coupure de ligne (seuls les sauts explicites comptent).
    const ws = (bodyPr['@wrap'] ?? inhBP['@wrap']) === 'none' ? 'pre' : 'pre-wrap';
    const wsAttr = ws === 'pre' ? ' data-wrap="none"' : '';
    // styles du master pour ce type de placeholder (phKind 'none' -> {} : aucun ; 'other' n'est jamais
    // choisi par les appelants actuels)
    const styleLvls = ctx.txStyles?.[phKind] || {};
    // couleur de dernier recours = tx1 du thème (clair en thème sombre) — jamais #000 en dur
    // (sinon texte noir invisible sur fond sombre, ex. pied de page).
    const DEFCOL = cssColor(resolveColorFlat({ 'a:schemeClr': { '@val': 'tx1' } }, ctx)) || '#000';
    // autofit : PowerPoint réduit la police (et l'interligne) pour faire tenir le texte
    // Seul le bodyPr PROPRE de la forme est lu ici (pas l'hérité). fontScale 62500 -> 0,625 ;
    // lnSpcReduction 20000 -> 0,2 (interligne réduit de 20 points de %)
    const naf = bodyPr['a:normAutofit'];
    const fontScale = naf && naf['@fontScale'] ? +naf['@fontScale']/100000 : 1;
    const lnRed = naf && naf['@lnSpcReduction'] ? +naf['@lnSpcReduction']/100000 : 0;
    // interligne SIMPLE = 1,2 x la taille pour toutes les polices (calibré contre PowerPoint : 72 pt de
    // pas à 60 pt en Poppins/Arial/Calibri/Times) ; 1,15 serrait chaque boîte de 4 %.
    // interligne du conteneur (sans unité, relatif au font-size) : 1,2, réduit par l'autofit, plancher 0,8
    const lineH = Math.max(0.8, 1.2 * (1 - lnRed));
    // +mn-lt/+mj-lt = polices du thème (mineure/majeure)
    const resolveFont = (tf) => !tf ? null
      : /^\+mn-/.test(tf) ? ctx.minorFont : /^\+mj-/.test(tf) ? ctx.majorFont : tf;
    // style par niveau, du + prioritaire au - : lstStyle INLINE de la forme -> lstStyle du
    // placeholder de LAYOUT hérité -> lstStyle du placeholder de MASTER -> txStyles du MASQUE.
    const localLst = txBody['a:lstStyle'] || {};
    const inheritLst = inhLst || {};
    // lstStyle du placeholder de MASTER : maillon de la cascade entre le layout (inheritLst) et
    // p:txStyles du master (styleLvls). Vide si la forme n'est pas un placeholder.
    const mstLstS = mstLst || {};
    // niveau absent -> repli sur le niveau 1 de la même liste (choix de ce code : évite de perdre le
    // style d'un lstStyle qui ne définit que lvl1pPr)
    const lvlOf = (lst, lvl) => lst[`a:lvl${lvl}pPr`] || lst['a:lvl1pPr'];
    const colOf = (d) => d?.['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null;
    // dernier recours : defaultTextStyle de la présentation (texte hors placeholder)
    const dflt = ctx.defaultTextStyle || {};
    // style de FORME <p:style><a:fontRef idx="minor|major"><couleur/></a:fontRef> : donne au texte de la
    // forme sa couleur ET sa police de thème PAR DÉFAUT (souvent la SEULE source de couleur d'une forme
    // thémée non-placeholder, ex. bloc/callout SmartArt). Injecté en DERNIER repli (après lstStyle/master/
    // defaultTextStyle, avant DEFCOL) -> n'agit QUE là où le texte n'a aucune couleur/police définie.
    const fref = shapeStyle?.['a:fontRef'];
    const frefRgb = fref ? resolveColorFlat(fref, ctx) : null;
    const fontRefColor = frefRgb ? cssColor(frefRgb) : null;
    const fontRefFont = fref?.['@idx'] === 'major' ? ctx.majorFont : fref?.['@idx'] === 'minor' ? ctx.minorFont : null;
    // lvlStyle(lvl) -> propriétés PAR DÉFAUT d'un paragraphe de ce niveau, fusionnées selon la cascade
    // (loc = lstStyle de la forme, inh = hérité, mas = master, dt = defaultTextStyle). Les couleurs
    // sont déjà en CSS ; sz en pt ; marL/ind/marR en EMU ; lnSpc/spcBef/spcAft = nœuds bruts.
    // NB : mas.marR n'existe pas (non lu depuis le master) et defaultTextStyle ne fournit que
    // sz/b/color/font/algn.
    const lvlStyle = (lvl) => {
      const loc = lvlOf(localLst, lvl), inh = lvlOf(inheritLst, lvl);
      // mp = niveau du placeholder de MASTER ; s'intercale entre le layout (inh) et p:txStyles (mas)
      const mp = lvlOf(mstLstS, lvl);
      const ld = loc?.['a:defRPr'], id = inh?.['a:defRPr'], md = mp?.['a:defRPr'];
      const mas = styleLvls[lvl] || styleLvls[1] || {};
      const dt = dflt[lvl] || dflt[1] || {};
      return {
        sz: (ld?.['@sz'] ? +ld['@sz']/100 : null) ?? (id?.['@sz'] ? +id['@sz']/100 : null) ?? (md?.['@sz'] ? +md['@sz']/100 : null) ?? mas.sz ?? dt.sz,
        color: colOf(ld) ?? colOf(id) ?? colOf(md) ?? mas.color ?? dt.color ?? fontRefColor,
        font: ld?.['a:latin']?.['@typeface'] || id?.['a:latin']?.['@typeface'] || md?.['a:latin']?.['@typeface'] || mas.font || dt.font || fontRefFont,
        // gras : le niveau le PLUS SPÉCIFIQUE qui définit @b gagne (un b="0" du layout doit battre
        // le b="1" du master) — un simple || laisserait le master l'emporter à tort.
        b: ld?.['@b'] !== undefined ? ld['@b'] === '1'
         : id?.['@b'] !== undefined ? id['@b'] === '1'
         : md?.['@b'] !== undefined ? md['@b'] === '1'
         : mas.b,
        algn: loc?.['@algn'] || inh?.['@algn'] || mp?.['@algn'] || mas.algn,
        bu: bulletOf(loc) ?? bulletOf(inh) ?? bulletOf(mp) ?? mas.bu,
        marL: (loc?.['@marL'] ? +loc['@marL'] : null) ?? (inh?.['@marL'] ? +inh['@marL'] : null) ?? (mp?.['@marL'] ? +mp['@marL'] : null) ?? mas.marL,
        ind: (loc?.['@indent'] ? +loc['@indent'] : null) ?? (inh?.['@indent'] ? +inh['@indent'] : null) ?? (mp?.['@indent'] ? +mp['@indent'] : null) ?? mas.ind,
        marR: (loc?.['@marR'] ? +loc['@marR'] : null) ?? (inh?.['@marR'] ? +inh['@marR'] : null) ?? (mp?.['@marR'] ? +mp['@marR'] : null) ?? mas.marR,
        // interligne hérité (lstStyle inline → layout → master). Ex. cartes = spcPct 90 % : sans ça
        // j'utilise 1.15 → lignes trop espacées → dérive verticale vs PowerPoint (vu au comparateur).
        lnSpc: loc?.['a:lnSpc'] ?? inh?.['a:lnSpc'] ?? mp?.['a:lnSpc'] ?? mas.lnSpc,
        // spcBef/spcAft résolus dans la même chaîne (inline lstStyle → layout → master → master placeholder → txStyles)
        spcBef: loc?.['a:spcBef'] ?? inh?.['a:spcBef'] ?? mp?.['a:spcBef'] ?? mas.spcBef,
        spcAft: loc?.['a:spcAft'] ?? inh?.['a:spcAft'] ?? mp?.['a:spcAft'] ?? mas.spcAft,
      };
    };
    // ── Boucle des paragraphes ──
    // counters[lvl] : compteur de numérotation automatique par niveau (jamais remis à zéro dans la
    // boîte, même entre listes séparées par des paragraphes non numérotés)
    const counters = {};
    const paras = arr(txBody['a:p']);
    // hasText : au moins un caractère visible dans la boîte (sinon aucun foreignObject n'est émis)
    let html = '', hasText = false;
    for (let pi = 0; pi < paras.length; pi++) {
      const p = paras[pi];
      const pPr = p['a:pPr'] || {};
      const lvl = (+pPr['@lvl'] || 0) + 1;
      const def = lvlStyle(lvl);
      // Hauteur de ligne CSS du paragraphe = lnSpc du paragraphe, sinon hérité (lot/master via def.lnSpc),
      // sinon lineH du conteneur. Sert de line-height au <div> ET à chacun de ses runs (cf. lh0 plus bas).
      // Sans ça (ancien « line-height:0 » sur les runs), la boîte de ligne était nulle : un paragraphe qui
      // passe à la ligne voyait ses lignes, et les paragraphes d'une même zone de texte se superposaient
      // tous sur la même ligne de base (bug « texte qui se chevauche »).
      const lnSpcNode = pPr['a:lnSpc'] || def.lnSpc;
      const paraLH = lnSpcNode?.['a:spcPct']?.['@val'] ? (+lnSpcNode['a:spcPct']['@val']/100000*1.2*(1-lnRed)).toFixed(3)
                   : lnSpcNode?.['a:spcPts']?.['@val'] ? (+lnSpcNode['a:spcPts']['@val']/100*(1-lnRed)).toFixed(1)+'px'
                   : lineH.toFixed(3);
      // ST_TextAlignType : l / ctr / r / just (dist, justLow, thaiDist -> 'left' ici)
      const algn = { ctr:'center', r:'right', just:'justify' }[pPr['@algn'] || def.algn] || 'left';
      let runs = arr(p['a:r']);
      // sauts de ligne <a:br/> et champs <a:fld> intercalés entre les runs : le parse plat range chaque
      // balise dans son tableau et perd l'ORDRE (6 lignes séparées par des a:br rendues bout à bout).
      // On le reprend du parse ordonné : les tableaux plats gardent l'ordre relatif de chaque balise.
      const paraSeqO = txBodyO && (p['a:br'] !== undefined || (p['a:fld'] !== undefined && runs.length))
        ? byTag(txBodyO, 'a:p')[pi] : null;
      // Reconstruction : on parcourt les balises du paragraphe ORDONNÉ et on reprend, dans cet ordre,
      // le k-ième a:r / a:fld du parse plat (même ordre relatif). a:br devient un marqueur { __br }.
      // L'index pi suppose que txBodyO a exactement les mêmes a:p que le txBody plat.
      if (paraSeqO) {
        const flds = arr(p['a:fld']); let r = 0, f = 0;
        runs = kids(paraSeqO).map(tagOf).flatMap(t =>
          t === 'a:r' ? [runs[r++]] : t === 'a:fld' ? [flds[f++]] : t === 'a:br' ? [{ __br: true }] : [])
          .filter(Boolean);
      }
      // paragraphe fait d'un seul champ (a:fld sans a:r) : le champ est rendu comme un run (son a:t =
      // valeur mise en cache par PowerPoint ; slidenum est recalculé plus bas)
      if (!runs.length && p['a:fld']) runs.push(p['a:fld']);
      // line = HTML des runs ; firstColor / firstSz / firstFam = couleur, taille, police du 1er run
      // réel (servent à la puce, à l'interligne du <div> et au calcul de ligne de base)
      let line = '', firstColor = def.color || DEFCOL, firstSz = (def.sz || 18) * fontScale;
      // TABULATIONS de la 1re ligne, calculées comme PowerPoint : le texte va au prochain taquet après la
      // position courante — taquet personnalisé (a:tabLst, mesuré depuis le bord de la zone de texte), sinon
      // marL s'il est devant (retrait suspendu : « •<tab>Texte » tapé à la main), sinon multiple de defTabSz.
      // Le CSS n'a pas de taquet personnalisé -> tabulation remplacée par un blanc de la largeur exacte,
      // mesurée avec les chasses de PowerPoint (062_02 d.85 : texte 55 pt trop loin avec tab-size seul).
      // Toutes les grandeurs de tabulation sont en pt (u()). Un paragraphe à PUCE automatique ne suit
      // pas la position x (tabX = null) : la largeur de la puce n'est pas mesurée -> tabulations CSS.
      const bu0 = bulletOf(pPr) ?? def.bu;
      const autoBu0 = bu0 && !bu0.none && (bu0.char || bu0.num || bu0.blip);
      const marL0 = u((pPr['@marL'] != null ? +pPr['@marL'] : def.marL) || 0);
      const ind0 = Math.max(u(pPr['@indent'] != null ? +pPr['@indent'] : (def.ind || 0)), -marL0);
      // defTabSz : intervalle des taquets par défaut, 914400 EMU = 1" = 72 pt
      const defTab0 = u(+(pPr['@defTabSz'] || 914400)) || 72;
      const stops0 = arr(pPr['a:tabLst']?.['a:tab']).map(t => u(+t['@pos'] || 0)).filter(x => x > 0).sort((x, y) => x - y);
      // x courant (depuis le bord de la zone de texte) ; null = mesure impossible -> tabulations CSS
      let tabX = autoBu0 ? null : marL0 + ind0;
      // bord droit de la ligne (depuis le bord de la zone de texte) : permet de suivre les RETOURS À LA
      // LIGNE mot à mot (chasses et règle de coupure de PowerPoint) pour placer aussi les tabulations des
      // lignes suivantes. Une ligne suivante repart à marL.
      const marR0 = u((pPr['@marR'] != null ? +pPr['@marR'] : def.marR) || 0);
      const lineRight = u(w - iL - iR) - marR0;
      // nextStop(x) : 1er taquet personnalisé strictement après x ; sinon marL si x est avant ; sinon
      // le multiple suivant de defTabSz (1e-6 évite qu'un x pile sur un multiple reste sur place).
      // Les types de taquet (@algn l/ctr/r/dec) ne sont pas distingués : tous traités comme gauche.
      const nextStop = (x) => stops0.find(t => t > x + 0.01) ?? (x < marL0 - 0.01 ? marL0 : (Math.floor(x / defTab0 + 1e-6) + 1) * defTab0);
      // police du « strut » CSS de la ligne = celle du 1er run : sinon le <div> prend la police par défaut
      // du navigateur, dont les métriques décalent la boîte et grandissent chaque ligne (texte ré-autofité).
      let firstFam = fontFamily(resolveFont(def.font) || ctx.minorFont), firstSet = false, paraInk = false;
      // paragraphe VIDE : sa hauteur vient de la marque de fin (endParaRPr sz/latin), pas du style par
      // défaut — sinon chaque ligne vide 18 pt au lieu de 16 pt décale tout le texte qui suit.
      const endRPr = p['a:endParaRPr'];
      if (!runs.length && endRPr) {
        if (+endRPr['@sz']) firstSz = +endRPr['@sz'] / 100 * fontScale;
        const ef = resolveFont(endRPr['a:latin']?.['@typeface']);
        if (ef) firstFam = fontFamily(ef);
      }
      // paragraphe contenant des FORMULES OMML (a14:m / m:oMath) : rendre en ORDRE le texte (a:r)
      // et le math (fractions, etc.) intercalés — le parse plat perd l'ordre a:r ↔ math.
      const paraO = txBodyO && (p['a14:m'] !== undefined || p['m:oMathPara'] !== undefined || p['m:oMath'] !== undefined)
        ? byTag(txBodyO, 'a:p')[pi] : null;
      // Chemin SIMPLIFIÉ (paragraphe avec formules) : pas de tabulations, de décorations, de puce
      // calculée par run, ni d'advanceFix ; les runs n'ont pas de vertical-align (le jeton VA_TOK
      // n'y figure pas). La puce éventuelle est tout de même ajoutée plus bas.
      if (paraO) {
        const fam = fontFamily(resolveFont(def.font) || ctx.minorFont), pc = def.color || DEFCOL;
        for (const ch of kids(paraO)) {
          const ct = tagOf(ch);
          if (ct === 'a:r') {
            const rPrO = first(ch, 'a:rPr');
            const a = at(rPrO || {});
            const sz = ((a['@sz'] ? +a['@sz']/100 : def.sz) || 18) * fontScale;
            const b = (a['@b'] === '1' || (a['@b'] === undefined && def.b)) ? 'font-weight:700;' : '';
            const it = a['@i'] === '1' ? 'font-style:italic;' : '';
            // couleur PROPRE du run (sinon défaut paragraphe) — le chemin OMML l'ignorait, d'où du texte
            // en couleur de fond invisible (ex. encart rouge d.15 : run tx1=clair rendu en bg1=sombre).
            const sfO = rPrO && first(rPrO, 'a:solidFill');
            const rc = sfO ? cssColor(resolveColorFlat(PF.parse(`<r xmlns:a="a">${serialize(sfO)}</r>`).r['a:solidFill'], ctx)) : pc;
            const tN = first(ch, 'a:t'); const txt = esc(tN ? kids(tN).map(c => c['#text'] ?? '').join('') : '');
            if (txt.trim()) { hasText = true; firstColor = rc; }
            line += `<span style="font-size:${sz.toFixed(2)}px;color:${rc};font-family:${fam};${b}${it}">${txt}</span>`;
          } else if (ct === 'a:br') {
            line += '<br/>';
          } else if (ct === 'a14:m' || ct === 'm:oMathPara' || ct === 'm:oMath') {
            // formules OMML : PowerPoint les compose en Cambria Math (pas dans la police du paragraphe)
            line += `<span style="font-size:${((def.sz || 18) * fontScale).toFixed(2)}px;color:${pc};font-family:${fontFamily('Cambria Math')}">${renderMath(ch)}</span>`;
            hasText = true;
          }
        }
      } else
      // Chemin NORMAL : un <span> par run (ou par champ / saut de ligne), propriétés résolues run par run
      for (let ri = 0; ri < runs.length; ri++) {
        const r = runs[ri];
        if (r.__br) { line += '<br/>'; if (tabX != null) tabX = marL0; continue; } // nouvelle ligne : repart à marL
        const rPr = r['a:rPr'] || {};
        const sz = ((+rPr['@sz'] ? +rPr['@sz']/100 : def.sz) || 18) * fontScale; // points, autofit
        // couleur : seul a:solidFill est lu sur le run (gradFill de texte -> couleur par défaut)
        const col = rPr['a:solidFill'] ? cssColor(resolveColorFlat(rPr['a:solidFill'], ctx)) : (def.color || DEFCOL);
        // police : seul le script latin (a:latin) est lu ; a:ea / a:cs ignorés
        const font = resolveFont(rPr['a:latin']?.['@typeface']) || resolveFont(def.font) || ctx.minorFont;
        const b = (rPr['@b']==='1' || (rPr['@b']===undefined && def.b))?'font-weight:700;':'';
        const i = rPr['@i']==='1'?'font-style:italic;':'';
        // décorations cumulables : soulignement + barré (§21.1.2.3.6 u / strike)
        const deco = [];
        if (rPr['@u'] && rPr['@u']!=='none') deco.push('underline');
        if (rPr['@strike'] && rPr['@strike']!=='noStrike') deco.push('line-through');
        const td = deco.length ? `text-decoration:${deco.join(' ')};` : '';
        // cap : MAJUSCULES auto (all) / petites capitales (small) — §20.1.10.65 ST_TextCapsType
        const cap = rPr['@cap']==='all' ? 'text-transform:uppercase;'
                  : rPr['@cap']==='small' ? 'font-variant:small-caps;' : '';
        // spc : espacement de caractères en centièmes de point (ST_TextPoint)
        // (1 px CSS = 1 pt ici, voir guide §5 ; l'autofit fontScale n'est pas appliqué à @spc)
        const spcPx = rPr['@spc'] != null && +rPr['@spc'] ? +rPr['@spc'] / 100 : 0;
        // baseline : exposant/indice (ST_Percentage, millièmes de %) -> décalage vertical
        // (en em = % de la taille du run : un % CSS se rapporterait à line-height, nul pour les runs lh0)
        // ex. baseline="30000" = exposant à +30 % ; "-25000" = indice. La taille du run n'est PAS
        // réduite ici : il garde son @sz.
        const blEm = rPr['@baseline'] && +rPr['@baseline'] ? +rPr['@baseline'] / 100000 : 0;
        // décalage vertical du run = exposant/indice (em) - recalage de ligne de base (px, VA_TOK remplacé
        // en fin de paragraphe par baselineShift, connu seulement après tous les runs)
        const bl = `vertical-align:calc(${blEm.toFixed(3)}em${VA_TOK};`;
        // highlight : surlignage du texte (§21.1.2.3.4) -> fond coloré derrière le run (comme Word)
        const hl = rPr['a:highlight'] ? `background-color:${cssColor(resolveColorFlat(rPr['a:highlight'], ctx))};` : '';
        // contour de texte (rPr>a:ln, type WordArt/titre stylé §21.1.2.1.11) -> -webkit-text-stroke.
        // Additif : présent uniquement si le run déclare un a:ln avec remplissage -> aucun autre run touché.
        const tln = rPr['a:ln'];
        let ol = '';
        if (tln && tln['a:noFill'] === undefined && (tln['a:solidFill'] || tln['a:gradFill'])) {
          const oc = tln['a:solidFill'] ? cssColor(resolveColorFlat(tln['a:solidFill'], ctx)) : col;
          const ow = Math.max(0.3, u(+tln['@w'] || 9525) / 2); // EMU->px, moitié (le trait déborde des 2 côtés)
          ol = `-webkit-text-stroke:${ow.toFixed(2)}px ${oc};`;
        }
        // champ slidenum : afficher le VRAI numéro de diapo (sinon le texte-invite « ‹N°› »
        // du layout/master fuit — c'est une zone libre non-placeholder, donc non sautée).
        let raw = r['@type'] === 'slidenum' ? String(slideNo) : (r['a:t'] ?? '');
        // espaces de FIN de paragraphe (dernier run) : non rendus par PowerPoint → couper (sinon
        // ils élargissent la boîte et déclenchent une sur-réduction de l'autofit). Espaces INTERNES
        // multiples conservés (pre-wrap) pour le positionnement type « TURBULENCE   vu à ».
        if (ri === runs.length - 1) raw = raw.replace(/[ \t]+$/, '');
        // run en police SYMBOLE (Wingdings/Webdings/Symbol, souvent via <a:sym>) : glyphes remappés vers
        // Unicode (anti-tofu Linux/LMS) dans SYM_FAMILY ; le texte ordinaire du run garde sa police.
        const parts = symbolParts(raw, font, rPr['a:sym']?.['@typeface']);
        const allSym = parts?.every(p => p.sym);
        // txt = HTML du run : texte échappé ; si remappage partiel, les glyphes symboles sont dans des
        // <span> en SYM_FAMILY ; si remappage total, c'est le run entier qui prend SYM_FAMILY (famCss)
        let txt = !parts ? esc(raw) : allSym ? esc(parts.map(p => p.text).join(''))
          : parts.map(p => p.sym ? `<span style="font-family:${SYM_FAMILY}">${esc(p.text)}</span>` : esc(p.text)).join('');
        const famCss = allSym ? SYM_FAMILY : fontFamily(font);
        // Suivi de la position x (tabX) et remplacement des tabulations par des blancs de largeur exacte.
        // tabX est mis à jour pour TOUT run mesurable (même sans tabulation), car la position d'une
        // tabulation dépend de tout ce qui la précède sur la ligne ; il simule aussi les retours à la
        // ligne automatiques (même règle de coupure que cellTextHeight). Dès qu'une mesure échoue (glyphe
        // inconnu, run symbole), tabX passe à null pour le reste du paragraphe -> tabulations laissées au
        // CSS (tab-size).
        if (tabX != null && !parts) {
          // découpe aux tabulations ; chaque segment avance tabX de sa largeur mesurée
          const segs = raw.split('	'), out = [];
          for (let k = 0; k < segs.length && tabX != null; k++) {
            // avance mot par mot (espace finale comprise) ; un mot qui ne tient plus ouvre une ligne à marL
            let okSeg = true;
            for (const word of segs[k].split(/(?<= )/)) {
              if (!word) continue;
              const ww = pptTextWidth(famCss, !!b, !!i, sz, word);
              const wn = word.endsWith(' ') ? pptTextWidth(famCss, !!b, !!i, sz, word.trimEnd()) : ww;
              if (ww == null || wn == null) { okSeg = false; break; }
              // retour à la ligne simulé : le mot (sans son espace) dépasse le bord droit -> il ouvre une
              // nouvelle ligne à marL (sauf en wrap="none", et jamais pour le 1er mot d'une ligne)
              if (tabX > marL0 + 0.01 && tabX + wn > lineRight - 0.02 && ws !== 'pre') tabX = marL0 + ww;
              else tabX += ww;
            }
            if (!okSeg) { tabX = null; break; }
            out.push(esc(segs[k]));
            if (k < segs.length - 1) {
              // tabulation -> <span inline-block> vide de la largeur jusqu'au prochain taquet
              const gap = nextStop(tabX) - tabX;
              out.push(`<span style="display:inline-block;width:${gap.toFixed(2)}px"></span>`); tabX += gap;
            }
          }
          // txt n'est remplacé que s'il y avait au moins une tabulation ET que toutes les mesures ont
          // réussi ; sinon le texte brut (avec son caractère TAB, rendu par tab-size CSS) est gardé
          if (tabX != null && segs.length > 1) txt = out.join('');
        } else if (parts) tabX = null;
        if (!firstSet) { firstSz = sz; firstFam = famCss; firstSet = true; } // 1er run réel -> interligne correct du <div>
        if (txt.trim()) { hasText = true; firstColor = col; }
        // espacement : spc du run + correction d'arrondi des chasses façon PowerPoint (advanceFix)
        const ls = spcPx + (allSym ? 0 : advanceFix(famCss, !!b, !!i, sz, rPr['@cap'] === 'all' ? raw.toUpperCase() : raw));
        const spc = Math.abs(ls) >= 0.001 ? `letter-spacing:${ls.toFixed(3)}px;` : '';
        // line-height du paragraphe (paraLH) pour un run pas plus grand que le 1er : sa boîte garde la
        // hauteur d'UNE ligne sans jamais dépasser celle du paragraphe. En CSS, des polices mêlées (Poppins
        // + symboles en Times) s'alignaient sur la ligne de base et leur union dépassait L (34 px au lieu de
        // 24) ; un line-height UNITAIRE (paraLH) donne 1,2 x la taille de CHAQUE run indépendamment de ses
        // métriques -> plus de dépassement. Un run PLUS grand agrandit la ligne, comme chez PowerPoint.
        // (Avant : « line-height:0 » ; il annulait la boîte de ligne et faisait s'effondrer le <div> à 0 ->
        // lignes retournées à la ligne ET paragraphes superposés au même y.)
        const lh0 = sz <= firstSz + 0.01 ? `line-height:${paraLH};` : '';
        if (raw.trim()) paraInk = true; // paragraphe « encré » (voir le &#8203; en fin de <div>)
        // Ordre des déclarations CSS du run : taille, line-height 0, couleur, police, gras, italique,
        // décoration, casse, letter-spacing, vertical-align (avec jeton VA_TOK), surlignage, contour.
        line += `<span style="font-size:${sz.toFixed(2)}px;${lh0}color:${col};font-family:${famCss};${b}${i}${td}${cap}${spc}${bl}${hl}${ol}">${txt}</span>`;
      }
      // ── Puce, retraits, espacements, émission du <div> du paragraphe ──
      // Une puce n'est émise que si le paragraphe a du texte visible (PowerPoint n'affiche pas la puce
      // d'un paragraphe vide). Taille = taille du 1er run × buSzPct ; couleur = buClr ou 1er run.
      // puce : pPr du paragraphe prioritaire, sinon héritée du niveau de style
      const bu = bulletOf(pPr) ?? def.bu;
      let bullet = '', bulletCore = '';
      if (bu && !bu.none && bu.blip && (line || '').trim()) {
        // puce IMAGE (a:buBlip) : petite icône alignée sur le texte
        const bid = bu.blip['@r:embed'] || bu.blip['@embed'];
        const bhref = bid && curRels[bid] ? imageHref(pptDir, curRels[bid], embed) : null;
        if (bhref) { const bsz = (bu.szPct ? firstSz * bu.szPct : firstSz); bullet = `<img src="${bhref}" style="height:${bsz.toFixed(1)}px;vertical-align:middle"/> `; }
      } else if (bu && !bu.none && (bu.char || bu.num) && (line || '').trim()) {
        let sym = bu.char, buFam;
        if (bu.num) { counters[lvl] = (counters[lvl] || 0) + 1; sym = counters[lvl] + (/paren/i.test(bu.num) ? ')' : '.'); buFam = bu.font ? fontFamily(resolveFont(bu.font)) : firstFam; }
        else {
          // puce caractère : mapper les puces en police SYMBOLE (Wingdings) vers Unicode (anti-tofu)
          const bs = bulletSym(bu.char, bu.font);
          sym = bs.sym;
          buFam = bs.symbolic ? "'Arial','Liberation Sans','DejaVu Sans',sans-serif" : bu.font ? fontFamily(resolveFont(bu.font)) : firstFam; // sans buFont (buFontTx) = police du texte
        }
        // couleur PROPRE de la puce (a:buClr) si présente, sinon couleur du 1er run ; taille via a:buSzPct
        const buCol = bu.clr ? cssColor(resolveColorFlat(bu.clr, ctx)) : firstColor;
        const buSz = bu.szPct ? firstSz * bu.szPct : firstSz;
        // la puce porte aussi le jeton VA_TOK : elle suit le même recalage de ligne de base que le texte
        bulletCore = `<span style="font-size:${buSz.toFixed(2)}px;line-height:0;vertical-align:calc(0em${VA_TOK};color:${buCol};font-family:${buFam}">${esc(sym)}</span>`;
      }
      // retrait : pPr du paragraphe prioritaire (marL/indent équilibrés du niveau de liste)
      const marLe = pPr['@marL'] != null ? +pPr['@marL'] : def.marL;
      const inde = pPr['@indent'] != null ? +pPr['@indent'] : def.ind;
      const marL = marLe != null ? u(marLe) : 0;
      // le retrait négatif (puce suspendue) ne doit jamais dépasser la marge -> pas de débordement gauche
      const indRaw = inde != null ? u(inde) : 0;
      const ind = Math.max(indRaw, -marL);
      // largeur de la puce suspendue = |indent| même quand marL + indent < 0 est borné à 0 : PowerPoint
      // pose la puce au bord et le texte |indent| plus loin (indent -36 pt sans marL : texte à 36 pt,
      // lignes suivantes au bord — vérifié sur l'export, 022_04 d.96).
      const hang = indRaw < 0 ? -indRaw : 0;
      // marge DROITE du paragraphe (@marR, §21.1.2.2.7) : ignorée avant -> texte coupé plus large que
      // dans PowerPoint (062_02 d.89 : marR 39,85 pt, lignes trop longues d'un ou deux mots).
      const marRe = pPr['@marR'] != null ? +pPr['@marR'] : def.marR;
      const marR = marRe ? u(marRe) : 0;
      // Modèle PowerPoint -> CSS : marL (marge gauche de TOUTES les lignes) = padding-left ; indent
      // (décalage de la 1re ligne, négatif pour une puce suspendue) = text-indent.
      const pad = ((marL || ind) ? `padding-left:${marL.toFixed(1)}px;text-indent:${ind.toFixed(1)}px;` : '')
        + (marR ? `padding-right:${marR.toFixed(1)}px;` : '');
      // puce SUSPENDUE (indent < 0) : elle occupe tout le retrait et le texte démarre à marL, comme le
      // taquet de PowerPoint (sinon « - The… » collé au tiret au lieu d'être aligné sous le texte).
      if (bulletCore) bullet = hang > 0
        ? `<span style="display:inline-block;min-width:${hang.toFixed(1)}px;text-indent:0">${bulletCore}</span>`
        : bulletCore + '\u00a0'; // puce non suspendue : suivie d'une espace insécable (NBSP)
      // interligne (lnSpc) + espacement avant/après (spcBef/spcAft) — §21.1.2.2.{5,7,12}
      // spcPct = millièmes de % (100000 = 100 %) ; spcPts = centièmes de point.
      // lnSpc du paragraphe, sinon hérité (lstStyle/layout/master via def.lnSpc). spcPct 100 % =
      // interligne SIMPLE ≈ 1.2× la police (pas 1.0) → facteur 1.2 pour coller à PowerPoint.
      // normAutofit @lnSpcReduction : PowerPoint réduit AUSSI l'interligne explicite du paragraphe
      // (pas seulement les boîtes sans lnSpc via le conteneur). Sans ça, un paragraphe à lnSpc 120 %
      // dans une boîte autofit à 20 % de réduction rend 1.44 au lieu de 0.96×1.2=1.152 → texte trop
      // haut → déborde → l'autofit navigateur sur-rétrécit → mauvais nb de lignes (ex. question d.6).
      const ls = pPr['a:lnSpc'] || def.lnSpc;
      const lnMul = 1 - lnRed; // 1 si pas de réduction -> aucun changement pour ces boîtes-là
      const lnSpcCss = ls?.['a:spcPct']?.['@val'] ? `line-height:${(+ls['a:spcPct']['@val']/100000*1.2*lnMul).toFixed(3)};`
                     : ls?.['a:spcPts']?.['@val'] ? `line-height:${(+ls['a:spcPts']['@val']/100*lnMul).toFixed(1)}px;` : '';
      // hauteur de ligne effective (px) -> décalage de ligne de base façon PowerPoint (baselineShift)
      const Lpx = ls?.['a:spcPct']?.['@val'] ? +ls['a:spcPct']['@val'] / 100000 * 1.2 * lnMul * firstSz
                : ls?.['a:spcPts']?.['@val'] ? +ls['a:spcPts']['@val'] / 100 * lnMul
                : lineH * firstSz;
      // un seul décalage par paragraphe, calculé sur le 1er run (police + taille) : les runs d'autres
      // tailles reçoivent le même décalage en px
      const blShift = baselineShift(firstFam, firstSz, Lpx);
      // margin-right 0,2 px : PowerPoint renvoie un mot à la ligne dès que la ligne (sans son espace finale)
      // atteint la largeur disponible, à ~0,01 pt près (mesuré via COM : 188,38 pt renvoyé pour 188,391
      // disponibles) ; Chrome le gardait (0,06 px ne suffisait pas : il arrondit ses positions au 1/64 px).
      // Borne haute : une ligne à 1,2 px sous la limite reste gardée chez PowerPoint (022_04 d.18).
      // taquets par défaut (defTabSz, 72 pt) : sans ça, tab = 8 espaces CSS -> « ⇒ » moins en retrait
      // (def.defTab n'est jamais renseigné par lvlStyle : en pratique pPr@defTabSz ou 72 pt)
      const tabPx = u(+(pPr['@defTabSz'] || def.defTab || 914400));
      // spcMk(nœud spcBef/spcAft, 'top'|'bottom') -> marge CSS : spcPts en 1/100 pt -> px ;
      // spcPct en 1/1000 de % converti en em (donc relatif à la taille du texte du <div>).
      const spcMk = (node, side) => !node ? ''
        : node['a:spcPts']?.['@val'] ? `margin-${side}:${(+node['a:spcPts']['@val']/100).toFixed(1)}px;`
        : node['a:spcPct']?.['@val'] ? `margin-${side}:${(+node['a:spcPct']['@val']/100000).toFixed(2)}em;` : '';
      // spcBef/spcAft : pPr de la diapo prioritaire, sinon HÉRITÉ (lstStyle/layout/master via def).
      // spcBef supprimé sur le 1er paragraphe : PowerPoint ne pousse pas le bloc vers le bas en haut
      // de boîte (l'espace-avant ne vaut qu'ENTRE paragraphes) — sinon décalage vertical du bloc entier.
      const spcBefNode = pi === 0 ? null : (pPr['a:spcBef'] ?? def.spcBef);
      // de même, l'espace-APRÈS du dernier paragraphe ne compte pas (calibré via COM : spcAft 12 pt sur le
      // dernier paragraphe = aucun effet en ancrage haut/milieu/bas) ; l'appliquer remontait de moitié
      // le texte centré des formes (040_02_01 d.32 : titres de cartes 3 pt trop haut).
      const spcAftNode = pi === paras.length - 1 ? null : (pPr['a:spcAft'] ?? def.spcAft);
      // interligne > 100 % : PowerPoint ajoute l'espace en plus AU-DESSUS de chaque ligne ; sous la dernière
      // ligne du bloc ne reste que la descente naturelle min(0,3 x taille, descendante) au lieu de 0,25 L
      // -> bloc plus court d'autant (calibré ; compte pour l'ancrage milieu/bas).
      const fmD = (metricsOf(firstFam) || [0, 0.25])[1];
      const tailCut = pi === paras.length - 1 && Lpx > 1.2 * firstSz + 0.01
        ? Math.max(0, 0.25 * Lpx - Math.min(0.3 * firstSz, fmD * firstSz)) : 0;
      // tailCut s'exprime comme une marge basse NÉGATIVE (combinée à un éventuel spcAft existant)
      let aftCss = spcMk(spcAftNode, 'bottom');
      if (tailCut > 0.05) {
        const aft = aftCss.match(/margin-bottom:([^;]+);/);
        aftCss = `margin-bottom:calc(${aft ? aft[1] : '0px'} - ${tailCut.toFixed(2)}px);`;
      }
      const spcCss = spcMk(spcBefNode, 'top') + aftCss;
      // data-par : index de paragraphe (pour le build par paragraphe bldP dans le player)
      // white-space:pre-wrap : préserve les espaces MULTIPLES (PowerPoint s'en sert pour positionner
      // du texte, ex. « TURBULENCE      vu à » où le « vu à » tombe entre deux badges) ; wrap conservé.
      // font-size sur le div = taille du texte -> l'interligne (unitless 1.15) se calcule dessus,
      // sinon il se base sur 16px (défaut) et gonfle la ligne (pied de page 9pt coupé dans sa boîte).
      // (L'interligne du conteneur vaut aujourd'hui lineH = 1,2, cf. plus haut.)
      // En fin de <div> : applyShift remplace le jeton VA_TOK de tous les runs + puce par blShift ;
      // &#8203; (espace de largeur nulle) donne une hauteur de ligne à un paragraphe sans encre
      // (vide ou seulement des espaces), sinon le <div> s'effondrerait à 0.
      html += `<div data-par="${pi}" style="font-size:${firstSz.toFixed(1)}px;font-family:${firstFam};white-space:${ws};${ws === 'pre' ? '' : 'overflow-wrap:break-word;'}tab-size:${tabPx.toFixed(1)}px;text-align:${algn};margin:0;${ws === 'pre' ? '' : 'margin-right:0.2px;'}${pad}${lnSpcCss}${spcCss}">${applyShift(bullet + line, blShift)}${paraInk ? '' : '&#8203;'}</div>`;
    }
    if (!hasText) return '';
    // data-af : mode d'ajustement de la boîte (norm = normAutofit, fixed sinon), repère pour le player.
    // Aucune réduction côté navigateur : le fontScale / lnSpcReduction que PowerPoint a ENREGISTRÉS sont
    // déjà appliqués ci-dessus, et il ne rétrécit jamais au rendu (slideFrame.js).
    const afMode = (naf !== undefined) ? 'norm' : 'fixed';
    // <a:noAutofit/> = « ne pas ajuster » : PowerPoint NE clippe PAS, le texte peut déborder la forme
    // (ex. libellés/pastilles : police 40pt dans une pastille de 40pt -> les capitales tiennent, mais
    // l'aire réduite par les insets clipperait). On étend alors le foreignObject à la boîte PLEINE, on
    // met les insets en padding et on laisse déborder (overflow visible).
    // Deux variantes de sortie :
    //   noAutofit : foreignObject = boîte ENTIÈRE, insets en padding (box-sizing:border-box) ;
    //   sinon     : foreignObject = boîte MOINS les insets (x = iL, y = iT).
    // Dans les deux cas, <div class="af"> est un flex colonne : justify-content = ancrage vertical,
    // align-items:center si anchorCtr.
    const noaf = bodyPr['a:noAutofit'] !== undefined;
    if (noaf) {
      return `<foreignObject x="0" y="0" width="${u(w)}" height="${u(h)}" style="overflow:visible">`
        + `<div xmlns="http://www.w3.org/1999/xhtml" class="af" data-af="fixed"${wsAttr} style="box-sizing:border-box;width:100%;height:100%;padding:${u(iT)}px ${u(iR)}px ${u(iB)}px ${u(iL)}px;display:flex;flex-direction:column;justify-content:${just};${ctrCss}line-height:${lineH.toFixed(3)};overflow:visible;">${html}</div>`
        + `</foreignObject>`;
    }
    // overflow VISIBLE : PowerPoint ne clippe pas le texte d'une forme au bord de la boîte. Avec
    // overflow:hidden, une ligne un peu plus haute que sa boîte (ex. chiffre 24pt dans une pastille dont
    // les insets réduisent la boîte à 25.7pt) était ROGNÉE en haut/bas. Le texte peut déborder (comme
    // PowerPoint).
    return `<foreignObject x="${u(iL)}" y="${u(iT)}" width="${u(w-iL-iR)}" height="${u(h-iT-iB)}" style="overflow:visible">`
      + `<div xmlns="http://www.w3.org/1999/xhtml" class="af" data-af="${afMode}"${wsAttr} style="width:100%;height:100%;display:flex;flex-direction:column;justify-content:${just};${ctrCss}line-height:${lineH.toFixed(3)};overflow:visible;">${html}</div>`
      + `</foreignObject>`;
  }

  // ── §O.8 ─ COMPOSITION FINALE ET ASSEMBLAGE DU SVG ────────────────────────────────────────────
  // Ordre de dessin (du fond vers l'avant) : fond unique, formes du master, formes du layout, formes
  // de la diapo. Pendant chaque passe, trois variables d'état de la fermeture sont positionnées :
  //   curRels (rels du niveau, pour résoudre ses r:embed), skipPh (placeholders du master/layout NON
  //   dessinés : ce sont des gabarits, leur contenu apparaît via les formes de la diapo) et spidPfx.
  // --- composition : masque (bg + déco) -> layout (déco) -> diapo (tout) ---
  // exécutée ici, après TOUTES les définitions (évite la zone morte temporelle des const)
  const masterPart = masterPath ? partOf(masterPath, 'p:sldMaster') : null;
  const layoutPart = layoutPath ? partOf(layoutPath, 'p:sldLayout') : null;
  // UN SEUL fond effectif (précédence diapo > layout > masque), dessiné DERRIÈRE toutes les formes.
  // Le fond n'est PAS cumulatif : un <p:bg> de layout/diapo REMPLACE celui du niveau au-dessus. Le
  // dessiner à chaque niveau repeindrait par-dessus les formes du niveau précédent — ex. les logos du
  // masque (logos partenaires) masqués par le fond bleu du layout. curRels doit pointer le bon niveau
  // (un fond peut être une image r:embed).
  // les trois fonds sont calculés (chacun avec les rels de son niveau), le 1er non vide gagne
  curRels = rels;                                  const slideBg  = bgRect(cSld);
  curRels = layoutRels;                            const layoutBg = layoutPart ? bgRect(layoutPart.cSld) : '';
  curRels = masterPath ? relsOf(masterPath) : {};  const masterBg = masterPart ? bgRect(masterPart.cSld) : '';
  body.push(slideBg || layoutBg || masterBg || '');
  if (masterPart) { curRels = relsOf(masterPath); skipPh = true; spidPfx = 'm'; walk(masterPart.spTree, IDENT); }
  if (layoutPart) { curRels = layoutRels; skipPh = true; spidPfx = 'l'; walk(layoutPart.spTree, IDENT); }
  // la diapo : placeholders DESSINÉS, spid bruts (ceux que cible la timeline d'animation)
  curRels = rels; skipPh = false; spidPfx = '';
  walk(spTree, IDENT);

  // dégradés définis -> <defs>
  // Linéaire : vecteur (x1,y1)->(x2,y2) centré sur la boîte, d'angle ang, en unités objectBoundingBox
  // (défaut SVG) -> l'angle est appliqué dans le carré unité de la boîte, donc « étiré » sur une forme
  // non carrée (@scaled non distingué). 0° = gauche -> droite, 90° = haut -> bas (y vers le bas,
  // comme DrawingML). Radial : centre 50 %, rayon 70 % fixes.
  const defs = grads.map(g => {
    const stops = g.stops.map(s=>`<stop offset="${s.pos.toFixed(1)}%" stop-color="${s.col}"/>`).join('');
    if (g.radial) return `<radialGradient id="${g.id}" cx="50%" cy="50%" r="70%">${stops}</radialGradient>`;
    const dx = Math.cos(g.ang), dy = Math.sin(g.ang);
    return `<linearGradient id="${g.id}" x1="${((0.5-dx/2)*100).toFixed(1)}%" y1="${((0.5-dy/2)*100).toFixed(1)}%" x2="${((0.5+dx/2)*100).toFixed(1)}%" y2="${((0.5+dy/2)*100).toFixed(1)}%">${stops}</linearGradient>`;
  }).join('');

  // motifs -> tuiles <pattern> 8×8 (fond bgClr + géométrie fgClr selon la famille de preset)
  // Tuile en userSpaceOnUse : 8 pt de côté quelle que soit la taille de la forme. Les regex sont
  // testées dans l'ordre : le 1er motif reconnu gagne (ex. « dkUpDiag » -> diagonale montante).
  const pdefs = patterns.map(p => {
    const s = 8, prst = p.prst; let geo = '';
    const line = (x1,y1,x2,y2) => `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${p.fg}" stroke-width="1"/>`;
    if (/^pct(\d+)/.test(prst)) { // points, densité ~ pourcentage
      // rayon ∝ √(pourcentage) (surface encrée ∝ densité), borné à [0,6 ; 3] dans une tuile de 8.
      // RegExp.$1 = groupe capturé par le test() juste au-dessus (API héritée mais fonctionnelle).
      const r = Math.min(3, Math.max(0.6, Math.sqrt(+RegExp.$1 / 100) * 3));
      geo = `<circle cx="4" cy="4" r="${r.toFixed(1)}" fill="${p.fg}"/>`;
    } else if (/(horz|Horz)$/.test(prst)) geo = line(0,4,s,4);
    else if (/(vert|Vert)$/.test(prst)) geo = line(4,0,4,s);
    else if (/(cross|Grid|trellis|weave|plaid|Check)/i.test(prst)) geo = line(0,4,s,4) + line(4,0,4,s);
    else if (/(DnDiag|dnDiag|diagBrick|shingle|wave|zigZag)/i.test(prst)) geo = line(0,0,s,s);
    else if (/(UpDiag|upDiag)/i.test(prst)) geo = line(0,s,s,0);
    else if (/(diagCross|Confetti|Dmnd|dotGrid|divot|sphere)/i.test(prst)) geo = line(0,0,s,s) + line(0,s,s,0);
    else geo = `<circle cx="4" cy="4" r="2" fill="${p.fg}"/>`; // défaut : semis de points
    return `<pattern id="${p.id}" patternUnits="userSpaceOnUse" width="${s}" height="${s}">`
      + `<rect width="${s}" height="${s}" fill="${p.bg}"/>${geo}</pattern>`;
  }).join('');

  // Assemblage : viewBox en POINTS, largeur 100 % (le conteneur fixe l'échelle) ; <defs> regroupe
  // dégradés, motifs, filtres, clips et marqueurs ; un rectangle BLANC sous tout (repli si aucun fond
  // n'est défini) ; puis body dans l'ordre z. w/h renvoyés en EMU.
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${u(W)} ${u(H)}" width="100%" style="display:block">`
    + `<defs>${defs}${pdefs}${filters.join('')}${clips.join('')}${[...lineMarkers.values()].join('')}</defs>`
    + `<rect width="${u(W)}" height="${u(H)}" fill="#ffffff"/>`
    + body.join('') + `</svg>`;
  return { svg, w: W, h: H };
}

// ████ §P ─ extractTiming : TIMELINE D'ANIMATION ────────────────────────────────────────────────
// Structure PresentationML d'une timeline (§19.5, p:timing) — arbre de conteneurs temporels :
//   p:timing > p:tnLst > p:par (racine, nodeType="tmRoot") > p:cTn > p:childTnLst
//     > p:seq (séquence principale, nodeType="mainSeq") > p:cTn > p:childTnLst
//       > p:par  <- UNE ÉTAPE (un clic)            [c'est le niveau lu par clickPars]
//         > p:cTn > p:childTnLst > p:par (groupe « avec / après le précédent », stCondLst@delay)
//           > p:par > p:cTn presetID presetClass presetSubtype nodeType  <- UN EFFET
//             > p:childTnLst > comportements : p:set, p:animEffect (filter="fade", "wipe(down)"…),
//               p:anim, p:animMotion (path), p:animRot (by), p:animScale… ; cible p:tgtEl/p:spTgt@spid
//               (+ p:txEl/p:pRg st..end pour une animation par paragraphe).
// Les attributs p:bldP (build par paragraphe) de p:bldLst ne sont pas lus : la plage vient de p:pRg.
//
// ================= extraction de la timeline d'animation (<p:timing>) =================
// Renvoie les "étapes de clic" : [{ effects:[{spid, cls, filter, transition, delay, dur}] }]
// cls = entr|exit|emph ; filter (fade/wipe(..)/circle(out)..) ; delay/dur en ms.
// Forme exacte d'un effet : { cls, spid, par: {from,to}|null, filter, transition, motion, delay, dur }
//   cls   : @presetClass (entr, exit, emph, path, verb, mediacall) ;
//   motion: { kind:'move', dx, dy, dur, rev } | { kind:'rot', deg, dur, rev } | null.
// animatedSpids : spids des formes ENTIÈRES qui ont une entrée -> à cacher avant leur 1re apparition.
// Seule la 1re p:seq trouvée est lue (en pratique la mainSeq) : les séquences interactives (déclencheur
// au clic sur une forme) sont ignorées.
export function extractTiming(pptDir, slideNo) {
  const text = readFileSync(join(pptDir, 'slides', `slide${slideNo}.xml`), 'utf8');
  // extraction par regex (évite de reparser toute la diapo). Suppose une balise <p:timing> SANS
  // attribut ; le glouton [\s\S]* va jusqu'au DERNIER </p:timing> (normalement il n'y en a qu'un).
  const m = text.match(/<p:timing>[\s\S]*<\/p:timing>/);
  if (!m) return { steps: [], animatedSpids: [] };
  const tree = P.parse(m[0]);
  const timing = tree.find((n) => tagOf(n) === 'p:timing');
  const seq = deepFind(timing, 'p:seq');
  if (!seq) return { steps: [], animatedSpids: [] };
  const seqCTn = first(seq, 'p:cTn');
  // chaque p:par enfant direct de la séquence = une étape. @nodeType n'est pas consulté : une étape
  // qui démarre toute seule (« après la précédente » en tête de séquence) est quand même traitée
  // comme un clic.
  const clickPars = byTag(first(seqCTn, 'p:childTnLst'), 'p:par');

  const animatedSpids = new Set();
  const steps = clickPars.map((par) => {
    const effects = [];
    collectEffects(par, 0, effects);
    // entrée ciblant la forme entière (sans plage de paragraphes) -> cachée au départ.
    // Les entrées par paragraphe (bldP) sont gérées séparément par le player (data-par).
    for (const e of effects) if (e.cls === 'entr' && !e.par) animatedSpids.add(e.spid);
    return { effects };
  }).filter((s) => s.effects.length); // étapes sans effet reconnu retirées
  return { steps, animatedSpids: [...animatedSpids] };

  // (déclarations de fonctions hissées : utilisables avant leur définition, après le return)
  // collectEffects(nœud, délai cumulé, accumulateur) : descente récursive. Un p:cTn portant
  // @presetID est un EFFET (feuille : on ne descend pas plus) ; sinon, un p:par conteneur ajoute son
  // délai de départ (stCondLst/cond@delay, en ms) au délai de ses descendants. « indefinite » = attend
  // un clic : ne compte pas. Les délais de cTn qui ne sont PAS dans un p:par (ex. p:seq) sont ignorés.
  function collectEffects(node, accDelay, acc) {
    if (tagOf(node) === 'p:par') {
      const cTn = first(node, 'p:cTn');
      if (cTn && at(cTn)['@presetID'] !== undefined) { acc.push(parseEffect(cTn, accDelay)); return; }
      if (cTn) {
        const cond = deepFind(first(cTn, 'p:stCondLst') ?? {}, 'p:cond');
        const d = cond && at(cond)['@delay'];
        if (d !== undefined && d !== 'indefinite') accDelay += Number(d) || 0;
      }
    }
    for (const c of kids(node)) if (typeof c === 'object') collectEffects(c, accDelay, acc);
  }
  // parseEffect(cTn de l'effet, délai hérité) -> objet effet (voir en-tête de §P).
  // spid / animEffect / pRg : 1ers trouvés en profondeur (un effet multi-cibles n'est lu que pour sa
  // 1re cible). delay = délai hérité + délai propre du cTn.
  function parseEffect(cTn, accDelay) {
    const a = at(cTn);
    const spTgt = deepFind(cTn, 'p:spTgt');
    const animEff = deepFind(cTn, 'p:animEffect');
    const pRg = spTgt ? deepFind(spTgt, 'p:pRg') : null; // plage de paragraphes (bldP)
    const own = deepFind(first(cTn, 'p:stCondLst') ?? {}, 'p:cond');
    const ownDelay = own ? Number(at(own)['@delay']) || 0 : 0;
    // durée : 1er cTn d'un behavior avec dur numérique
    // défaut 500 ms ; « > 1 » écarte les p:set de durée 1 ms (bascule de visibilité instantanée que
    // PowerPoint place en tête des effets d'entrée) pour prendre la durée du vrai comportement
    let dur = 500;
    const durNode = [...allCTn(cTn)].map((c) => at(c)['@dur']).find((d) => d && d !== 'indefinite' && +d > 1);
    if (durNode) dur = +durNode;
    return {
      cls: a['@presetClass'], spid: spTgt ? at(spTgt)['@spid'] : null,
      par: pRg ? { from: +at(pRg)['@st'] || 0, to: +at(pRg)['@end'] || 0 } : null,
      filter: animEff ? at(animEff)['@filter'] : null,
      transition: animEff ? at(animEff)['@transition'] : null,
      motion: parseMotion(cTn),
      delay: accDelay + ownDelay, dur,
    };
  }
  // mouvement/rotation (§19.5.3 animMotion, §19.5.5 animRot) -> {kind, dx/dy (fractions) | deg, dur, rev}
  // behTiming(comportement) -> { dur (ms, défaut 1000), rev (@autoRev : aller-retour) } lus sur son
  // p:cBhvr/p:cTn
  function behTiming(node) { const c = deepFind(node, 'p:cTn'); const a = at(c || {}); return { dur: +a['@dur'] || 1000, rev: a['@autoRev'] === '1' }; }
  // parseMotion : animMotion@path est un chemin type SVG (« M 0 0 L 0.25 0 E ») dont les coordonnées
  // sont des FRACTIONS de la taille de diapo. On n'en garde que le déplacement net (dernier point -
  // premier point) : les trajectoires courbes sont rendues en ligne droite. animRot@by en 1/60000 °.
  // Un effet qui n'a ni l'un ni l'autre -> null.
  function parseMotion(cTn) {
    const mm = deepFind(cTn, 'p:animMotion');
    if (mm) {
      const path = at(mm)['@path'] || '';
      // tous les nombres du chemin (y compris notation 1e-3), dans l'ordre ; les lettres de commande
      // sont ignorées -> pts = [x0, y0, …, xn, yn]
      const pts = (path.match(/-?\d*\.?\d+(?:[eE][-+]?\d+)?/g) || []).map(Number);
      if (pts.length >= 2) {
        const fx = pts[0], fy = pts[1], lx = pts[pts.length - 2], ly = pts[pts.length - 1];
        const t = behTiming(mm); return { kind: 'move', dx: lx - fx, dy: ly - fy, dur: t.dur, rev: t.rev };
      }
    }
    const mr = deepFind(cTn, 'p:animRot');
    if (mr) { const t = behTiming(mr); return { kind: 'rot', deg: (+at(mr)['@by'] || 0) / 60000, dur: t.dur, rev: t.rev }; }
    return null;
  }
  // allCTn : générateur de tous les p:cTn du sous-arbre (le nœud lui-même compris), ordre du document
  function* allCTn(node) {
    if (tagOf(node) === 'p:cTn') yield node;
    for (const c of kids(node)) if (typeof c === 'object') yield* allCTn(c);
  }
}

// ████ §P2 ─ extractTransition : TRANSITION DE DIAPO (<p:transition>) ──────────────────────────
// Beaucoup de decks (dont les gabarits Slidesgo/Google Slides) n'ont AUCUNE animation d'objet
// (<p:timing>) mais une TRANSITION par diapo : <mc:AlternateContent><mc:Choice Requires="p159">
// <p:transition p14:dur="2000"><p159:morph option="byObject"/></p:transition></mc:Choice>
// <mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>.
// On lit la variante moderne (morph) et on retombe sur le fade ; @p14:dur (ms) = durée. Le player
// (build-player.mjs) rejoue cette transition au changement de diapo. Renvoie null si la diapo n'en a pas.
export function extractTransition(pptDir, slideNo) {
  const text = readFileSync(join(pptDir, 'slides', `slide${slideNo}.xml`), 'utf8');
  // 1re <p:transition> = la variante « Choice » (morph) quand il y a un <mc:AlternateContent>, sinon
  // la transition simple. Le glouton non-greedy s'arrête à la 1re fermeture.
  const m = text.match(/<p:transition\b[^>]*>([\s\S]*?)<\/p:transition>/);
  if (!m) return null;
  const opener = m[0].slice(0, m[0].indexOf('>') + 1);
  const dur = +(opener.match(/p14:dur="(\d+)"/)?.[1] || 0); // durée en ms (p14:dur), sinon défaut
  const body = m[1];
  // type : morph si présent, sinon fade, sinon le 1er effet reconnu ; dir pour les poussées/balayages
  const type = /<p159:morph\b/.test(body) ? 'morph'
             : /<p:fade\b/.test(body) ? 'fade'
             : (body.match(/<p:(push|wipe|cover|split|reveal|zoom|dissolve|cut)\b/)?.[1]) || 'fade';
  const dir = body.match(/<p:(?:push|wipe|cover|split)[^>]*\bdir="([^"]+)"/)?.[1] || null;
  return { type, dur: dur || 2000, dir };
}

// ████ §Q ─ SÉRIALISATION ET EXPORTS ─────────────────────────────────────────────────────────────
// sérialise un noeud preserveOrder -> XML (pour re-parser en flat)
// Pièce centrale de la stratégie « parcours ordonné, lecture à plat » (guide §5) : utilisée avant
// chaque PF.parse d'un sous-arbre. Attributs et texte sont ré-échappés par esc() (fast-xml-parser
// les avait décodés) ; les balises gardent leur préfixe (a:, p:…) tel quel. Pas de balises
// auto-fermantes : <a:off x="0"/> devient <a:off x="0"></a:off> (équivalent au re-parse).
function serialize(node) {
  const t = tagOf(node);
  const a = at(node);
  const attrs = Object.entries(a).map(([k,v])=>` ${k.slice(1)}="${esc(v)}"`).join(''); // k.slice(1) retire le '@'
  const children = kids(node);
  // kids() renvoie toujours un tableau : cette branche « chaîne » ne se déclenche jamais (le texte
  // arrive comme enfant { '#text' } traité juste en dessous) ; conservée par prudence.
  if (typeof children === 'string') return `<${t}${attrs}>${esc(children)}</${t}>`;
  const inner = (children||[]).map(c => {
    if (c['#text'] !== undefined) return esc(c['#text']);
    return serialize(c);
  }).join('');
  return `<${t}${attrs}>${inner}</${t}>`;
}

// Exports du module : renderSlide, loadContext, extractTiming (API) ; mathTypeset, baselineShift,
// advanceFix, cellTextHeight, symbolParts, symbolText (exports directs, testés par
// test/text-layout.test.js et test/symbol-text.test.js) ; plus la liste ci-dessous.
// ─── Helpers purs exportés pour les tests unitaires (test/pptx-render.test.js) ───
// Aucun effet sur le rendu : on expose simplement des fonctions déjà définies plus haut.
export {
  u, esc, serialize,                                   // conversion EMU->pt · échappement XML · (dé)sérialisation
  hslFromRgb, rgbFromHsl, cssColor, resolveColorFlat,  // couleurs : conversions, format CSS, résolution + clamp
  symFontKind, symToUnicode, bulletSym,                // polices symbole (Wingdings/Webdings/Symbol) -> Unicode (symbolText : export direct plus haut)
  fontFamily, bulletOf,                                // police -> familles CSS · lecture d'une puce (pPr)
  tagOf, at, kids, byTag, first, deepFindFlat,          // helpers d'arbre (preserveOrder + objet plat)
};
