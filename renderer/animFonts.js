// Bloc @font-face (polices embarquées en data: URI) pour l'iframe des diapos animées.
// L'iframe est en ORIGINE OPAQUE (pas de réseau) : elle ne peut pas fetcher une route de polices,
// donc les polices doivent être inlinées. Identique pour tous les cours -> construit UNE fois (cache
// mémoire), servi par GET /api/anim/fonts.css ; le front le récupère une fois et l'injecte dans chaque
// srcDoc. Les familles correspondent à ce que le moteur émet (fontFamily de pptx-render.mjs) ; Arial /
// Times New Roman sont couverts par les substituts métrique-compatibles Liberation via le stack CSS.
//
// ─── COMMENT ÇA MARCHE (vue d'ensemble) ─────────────────────────────────────────────────────────
// TROIS SERVICES, UNE MÊME SOURCE (le dossier fonts/ et ses deux JSON)
//   1. getFontsCss() / preloadFontsCss() -> le CSS @font-face complet (~4 Mo, polices en base64).
//      Consommateurs : la route du LMS GET /api/anim/fonts.css ; dans ce dépôt, web-fonts.mjs (outils
//      viewer / player / sweep), qui n'en garde que les familles utiles.
//   2. familyMetrics() -> { famille: [ascent, descent] } : métriques « win » utilisées par
//      pptx-render.mjs (baselineShift) pour placer la ligne de base comme PowerPoint.
//   3. advanceTable(famille, gras, italique) -> chasses par caractère : utilisées par pptx-render.mjs
//      (advanceFix, pptTextWidth) pour reproduire l'arrondi des chasses de PowerPoint et donc ses
//      coupures de ligne.
//
// POURQUOI DES POLICES DE SUBSTITUTION
//   Les postes clients (et le serveur) n'ont pas forcément Calibri, Cambria… (licence Microsoft).
//   On embarque des polices LIBRES « métrique-compatibles » (mêmes chasses glyphe par glyphe : Carlito
//   pour Calibri, Liberation pour Arial/Times/Courier, Caladea pour Cambria…) ou, à défaut, proches.
//   Le moteur écrit une pile CSS « police d'origine, substitut, générique » : si la police d'origine est
//   installée, le navigateur la prend ; sinon il tombe sur le substitut embarqué.
//
// FICHIERS LUS (générés par fonts/build-web-fonts.py, à partir des TTF/OTF sources)
//   fonts/*.woff2        polices réduites aux plages Unicode utiles (latin étendu, grec, flèches,
//                        maths, puces…) : sans ce sous-ensemble le CSS pèserait des dizaines de Mo ;
//   fonts/metrics.json   { "Fichier.woff2": [usWinAscent / unitsPerEm, usWinDescent / unitsPerEm] }
//                        (table OS/2, arrondi à 4 décimales) ;
//   fonts/advances.json  { "Fichier.woff2": { upm: unitsPerEm, a: { "codepoint": chasse } } }
//                        (chasse = avance horizontale hmtx en unités de police, clé = code décimal).
//
// FORMAT DU CSS PRODUIT (une règle par ligne)
//   @font-face{font-family:'Poppins';font-weight:400;font-style:normal;font-display:swap;
//     ascent-override:113.50%;descent-override:62.70%;line-gap-override:0%;
//     src:url(data:font/woff2;base64,…) format('woff2')}
//   Les alias locaux (LOCAL_ALIASES) viennent EN PREMIER, puis les FACES embarquées.
//
// PIÈGES
//   - Les chemins sont relatifs à CE fichier (import.meta.url) : fonts/ doit rester à côté.
//   - Tout est mis en cache mémoire au premier appel : modifier un woff2 exige de relancer le process.
//   - Un fichier absent est ignoré sans erreur (face manquante = repli CSS sur la suivante de la pile).
// ────────────────────────────────────────────────────────────────────────────────────────────────
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Dossier des polices : renderer/fonts/, résolu depuis l'emplacement de ce module (pas depuis le cwd).
const FONTS_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "fonts");

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 1 — Catalogue des polices embarquées et des alias locaux
// ════════════════════════════════════════════════════════════════════════════════════════════════

// [famille, poids, style, fichier]. WOFF2 réduits aux plages utiles (voir fonts/build-web-fonts.py).
// Poids « 100 900 » = police variable. Chaque famille substitue une police PowerPoint (FONT_SUBST de
// pptx-render.mjs) : la police d'origine reste prioritaire si le poste l'a installée.
//   famille : nom CSS déclaré (doit correspondre EXACTEMENT au nom écrit par le moteur dans font-family) ;
//   poids   : nombre (face statique) ou chaîne « min max » (police variable : une seule face couvre la
//             plage, le navigateur interpole l'axe wght) ;
//   style   : "normal" | "italic" ; une famille sans face italique sera penchée artificiellement par le
//             navigateur (faux italique) ;
//   fichier : nom dans fonts/, clé aussi de metrics.json et advances.json.
//   L'ORDRE compte pour familyMetrics() (1re face « normal » de la famille) et, à score égal, pour
//   advanceTable() (1re face trouvée).
const FACES = [
  ["Poppins", 400, "normal", "Poppins-Regular.woff2"],
  ["Poppins", 500, "normal", "Poppins-Medium.woff2"],
  ["Poppins", 600, "normal", "Poppins-SemiBold.woff2"],
  ["Poppins", 700, "normal", "Poppins-Bold.woff2"],
  ["Poppins", 400, "italic", "Poppins-Italic.woff2"],
  ["Poppins", 700, "italic", "Poppins-BoldItalic.woff2"],
  // Arial / Helvetica
  ["Liberation Sans", 400, "normal", "LiberationSans-Regular.woff2"],
  ["Liberation Sans", 700, "normal", "LiberationSans-Bold.woff2"],
  ["Liberation Sans", 400, "italic", "LiberationSans-Italic.woff2"],
  ["Liberation Sans", 700, "italic", "LiberationSans-BoldItalic.woff2"],
  // Times New Roman
  ["Liberation Serif", 400, "normal", "LiberationSerif-Regular.woff2"],
  ["Liberation Serif", 700, "normal", "LiberationSerif-Bold.woff2"],
  ["Liberation Serif", 400, "italic", "LiberationSerif-Italic.woff2"],
  ["Liberation Serif", 700, "italic", "LiberationSerif-BoldItalic.woff2"],
  // Courier New / Consolas
  ["Liberation Mono", 400, "normal", "LiberationMono-Regular.woff2"],
  ["Liberation Mono", 700, "normal", "LiberationMono-Bold.woff2"],
  // Calibri / Calibri Light
  ["Carlito", 400, "normal", "Carlito-Regular.woff2"],
  ["Carlito", 700, "normal", "Carlito-Bold.woff2"],
  ["Carlito", 400, "italic", "Carlito-Italic.woff2"],
  ["Carlito", 700, "italic", "Carlito-BoldItalic.woff2"],
  // Cambria
  ["Caladea", 400, "normal", "Caladea-Regular.woff2"],
  ["Caladea", 700, "normal", "Caladea-Bold.woff2"],
  // Georgia (mêmes métriques)
  ["Gelasio", "400 700", "normal", "Gelasio-wght.woff2"],
  ["Gelasio", "400 700", "italic", "Gelasio-Italic-wght.woff2"],
  // Comic Sans MS (mêmes métriques)
  ["Comic Relief", 400, "normal", "ComicRelief-Regular.woff2"],
  ["Comic Relief", 700, "normal", "ComicRelief-Bold.woff2"],
  // Arial Black, Impact, Franklin Gothic (proches, métriques différentes)
  ["Archivo Black", 400, "normal", "ArchivoBlack-Regular.woff2"],
  ["Anton", 400, "normal", "Anton-Regular.woff2"],
  ["Libre Franklin", "100 900", "normal", "LibreFranklin-wght.woff2"],
  ["Libre Franklin", "100 900", "italic", "LibreFranklin-Italic-wght.woff2"],
  // News Gothic MT (dérivée), Gill Sans MT (approchante)
  ["News Cycle", 400, "normal", "NewsCycle-Regular.woff2"],
  ["News Cycle", 700, "normal", "NewsCycle-Bold.woff2"],
  ["Cabin", "400 700", "normal", "Cabin-wdth-wght.woff2"],
  ["Cabin", "400 700", "italic", "Cabin-Italic-wdth-wght.woff2"],
  // Bernhard Modern (approchante), Corbel / Aptos (approchante)
  ["Sorts Mill Goudy", 400, "normal", "SortsMillGoudy-Regular.woff2"],
  ["Sorts Mill Goudy", 400, "italic", "SortsMillGoudy-Italic.woff2"],
  ["Source Sans 3", "200 900", "normal", "SourceSans3-wght.woff2"],
  ["Source Sans 3", "200 900", "italic", "SourceSans3-Italic-wght.woff2"],
  // Century Schoolbook, Palatino / Book Antiqua (mêmes métriques)
  ["TeX Gyre Schola", 400, "normal", "texgyreschola-regular.woff2"],
  ["TeX Gyre Schola", 700, "normal", "texgyreschola-bold.woff2"],
  ["TeX Gyre Schola", 400, "italic", "texgyreschola-italic.woff2"],
  ["TeX Gyre Schola", 700, "italic", "texgyreschola-bolditalic.woff2"],
  ["TeX Gyre Pagella", 400, "normal", "texgyrepagella-regular.woff2"],
  ["TeX Gyre Pagella", 700, "normal", "texgyrepagella-bold.woff2"],
  ["TeX Gyre Pagella", 400, "italic", "texgyrepagella-italic.woff2"],
  ["TeX Gyre Pagella", 700, "italic", "texgyrepagella-bolditalic.woff2"],
  // Cairo / Space Grotesk : polices de très nombreux modèles Slidesgo (Google Fonts, OFL). Le deck les
  // référence par leur NOM EXACT (aucune ne remplace une police propriétaire) : sans face de ce nom, le
  // navigateur retombait sur son sans-serif générique -> glyphes différents ET chasses plus larges, donc
  // retours à la ligne erronés. On embarque donc la famille sous le nom qu'écrit le moteur.
  // « Space Grotesk Medium » est une FAMILLE distincte du fichier (instance wght=500) : le deck l'écrit
  // telle quelle, et la graisse 400 demandée retombe dessus (seule face de la famille).
  // Instances statiques générées depuis les variables Google (voir fonts/build-google-fonts.py).
  ["Cairo", 400, "normal", "Cairo-Regular.woff2"],
  ["Cairo", 700, "normal", "Cairo-Bold.woff2"],
  ["Space Grotesk", 400, "normal", "SpaceGrotesk-Regular.woff2"],
  ["Space Grotesk Medium", 500, "normal", "SpaceGrotesk-Medium.woff2"],
  // glyphes des polices Symbol / Wingdings remappés (≅ ⇒ ✔ ➢…) et secours maths
  ["DejaVu Sans", 400, "normal", "DejaVuSans.woff2"],
  ["DejaVu Sans", 700, "normal", "DejaVuSans-Bold.woff2"],
];

// Polices PowerPoint INSTALLÉES sur le poste dont le navigateur lirait mal les métriques (Calibri :
// hhea != win). Même famille, source local() seule : absente -> la face échoue et la pile CSS passe à
// Carlito. [famille, poids, style, nom local, fichier dont reprendre les métriques (compatible)]
// Mécanisme : on REDÉCLARE la famille 'Calibri' avec src:local('Calibri …') + les overrides de Carlito.
//   Une règle @font-face prend le pas sur la police système de même nom : le navigateur utilise alors
//   le fichier Calibri installé, mais avec l'ascent/descent imposé (celui de Carlito, identique aux
//   métriques win de Calibri puisque Carlito est métrique-compatible).
//   Nom local : nom complet (full name) ou PostScript de la face, tel que local() le recherche.
//   Calibri Light réutilise les métriques de Carlito Regular (pas de Carlito Light embarqué).
const LOCAL_ALIASES = [
  ["Calibri", 400, "normal", "Calibri", "Carlito-Regular.woff2"],
  ["Calibri", 700, "normal", "Calibri Bold", "Carlito-Bold.woff2"],
  ["Calibri", 400, "italic", "Calibri Italic", "Carlito-Italic.woff2"],
  ["Calibri", 700, "italic", "Calibri Bold Italic", "Carlito-BoldItalic.woff2"],
  ["Calibri Light", 400, "normal", "Calibri Light", "Carlito-Regular.woff2"],
  ["Calibri Light", 400, "italic", "Calibri Light Italic", "Carlito-Italic.woff2"],
];

// Métriques « win » (usWinAscent/Descent en fraction d'em) produites par fonts/build-web-fonts.py.
// PowerPoint cale ses lignes sur elles ; le navigateur prend hhea/typo, qui diffèrent pour Poppins
// (1050/350 contre 1135/627) : le texte d'une pastille tombait ~3 pt trop bas. Les imposer via
// ascent/descent-override aligne le placement vertical ; l'interligne (line-height explicite) ne change pas.
// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 2 — Métriques verticales « win » -> descripteurs CSS *-override
// ════════════════════════════════════════════════════════════════════════════════════════════════
// Contexte : une police TrueType/OpenType porte TROIS jeux de métriques verticales (hhea, OS/2 typo,
//   OS/2 win). Windows/GDI — donc PowerPoint — prend usWinAscent/usWinDescent ; les navigateurs
//   prennent en général hhea (ou typo si USE_TYPO_METRICS). Quand ils diffèrent, la ligne de base
//   n'est pas au même endroit. Les descripteurs CSS ascent-override / descent-override imposent au
//   navigateur la hauteur d'ascendante/descendante voulue, en POURCENTAGE de la taille de police.
// METRICS : { fichier -> [ascent, descent] } en fraction d'em, lu de façon SYNCHRONE à l'import
//   (petit fichier). Absent ou illisible -> {} : pas d'override, rendu navigateur par défaut.
let METRICS = {};
try { METRICS = JSON.parse(fs.readFileSync(path.join(FONTS_DIR, "metrics.json"), "utf8")); } catch { /* sans : défauts navigateur */ }
// overrides(file) -> fragment CSS à insérer dans @font-face, ou "" si métriques inconnues.
//   Exemple Poppins : [1.135, 0.627] -> ascent-override:113.50%;descent-override:62.70%;
//   line-gap-override:0% : l'interlignage additionnel (lineGap) est neutralisé ; la hauteur de ligne
//   effective vient du line-height explicite écrit par le moteur.
const overrides = (file) => {
  const m = METRICS[file];
  return m ? `ascent-override:${(m[0] * 100).toFixed(2)}%;descent-override:${(m[1] * 100).toFixed(2)}%;line-gap-override:0%;` : "";
};

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 3 — Construction du CSS @font-face (asynchrone au démarrage, synchrone en repli)
// ════════════════════════════════════════════════════════════════════════════════════════════════

// cached : CSS complet une fois construit (null tant qu'aucun appel n'a abouti). Partagé par
//   preloadFontsCss() et getFontsCss() : le premier des deux qui termine le remplit.
let cached = null;

// faceCss(...) -> une règle @font-face pour une police EMBARQUÉE (src = data: URI base64).
//   font-display:swap : le texte s'affiche tout de suite avec la police de repli puis bascule ; en
//   pratique la data: URI est décodée quasi immédiatement (pas de réseau).
//   Le nom de famille est entre apostrophes : aucun nom de FACES n'en contient.
const faceCss = (family, weight, style, file, b64) =>
  `@font-face{font-family:'${family}';font-weight:${weight};font-style:${style};font-display:swap;${overrides(file)}` +
  `src:url(data:font/woff2;base64,${b64}) format('woff2')}`;
// aliasCss() -> tableau de règles @font-face « local() seul » (voir LOCAL_ALIASES), métriques
//   empruntées au fichier compatible. Si la police n'est pas installée, la règle échoue et le
//   navigateur passe à la famille suivante de la pile (Carlito).
const aliasCss = () => LOCAL_ALIASES.map(([family, weight, style, local, file]) =>
  `@font-face{font-family:'${family}';font-weight:${weight};font-style:${style};font-display:swap;${overrides(file)}` +
  `src:local('${local}')}`);

// SEC-15 : préchargement ASYNC (appelé au démarrage) — évite le readFileSync synchrone de ~4 Mo qui
// gèlerait l'event loop au 1er appel de /anim/fonts.css. Idempotent.
// preloadFontsCss() -> Promise<string> (le CSS complet).
//   Lit les woff2 un par un (await séquentiel), les encode en base64 (+33 % de taille) et assemble.
//   Deux appels concurrents avant la fin feraient le travail deux fois (pas de promesse partagée),
//   sans autre conséquence : le résultat est identique.
export async function preloadFontsCss() {
  if (cached != null) return cached;                     // déjà construit : réponse immédiate
  const parts = aliasCss();                              // alias locaux d'abord
  for (const [family, weight, style, file] of FACES) {
    const p = path.join(FONTS_DIR, file);
    try {
      const b64 = (await fsp.readFile(p)).toString("base64"); // Buffer -> base64 pour la data: URI
      parts.push(faceCss(family, weight, style, file, b64));
    } catch { /* police absente : ignorée */ }
  }
  cached = parts.join("\n");                             // une règle par ligne (web-fonts.mjs les redécoupe)
  return cached;
}

// Accès synchrone (route). Normalement déjà en cache via preloadFontsCss() ; sinon repli sync (rare).
// getFontsCss() -> string. Même résultat que preloadFontsCss(), en lecture bloquante.
//   Dans ce dépôt, c'est la voie normale : les outils (web-fonts.mjs) sont des scripts en ligne de
//   commande où bloquer quelques centaines de ms au premier appel n'a pas d'importance.
export function getFontsCss() {
  if (cached != null) return cached;
  const parts = aliasCss();
  for (const [family, weight, style, file] of FACES) {
    const p = path.join(FONTS_DIR, file);
    if (!fs.existsSync(p)) continue;                     // police absente : ignorée (comme en async)
    parts.push(faceCss(family, weight, style, file, fs.readFileSync(p).toString("base64")));
  }
  cached = parts.join("\n");
  return cached;
}

// { famille CSS -> [ascent, descent] } (métriques « win », en em) des polices embarquées et des alias
// locaux : ce que le navigateur utilise réellement (ascent/descent-override). Sert au moteur pour caler
// la ligne de base comme PowerPoint (pptx-render.mjs, baselineShift).
// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 4 — Données pour le moteur de mise en page (pptx-render.mjs)
// ════════════════════════════════════════════════════════════════════════════════════════════════
// familyMetrics() -> { "Poppins": [1.135, 0.627], "Carlito": [0.9521, 0.2686], "Calibri": [...], … }
//   Une seule paire par FAMILLE : celle de la 1re face droite (style "normal") trouvée dans FACES —
//   les graisses d'une même famille ont en pratique les mêmes métriques win. Les alias locaux
//   (Calibri, Calibri Light) sont ajoutés ensuite avec les métriques de leur fichier compatible.
//   Une famille sans face « normal » (aucune ici) ou sans entrée dans metrics.json est absente.
//   Recalculé à chaque appel (pas de cache ici) : pptx-render.mjs le mémorise lui-même (_fm).
export function familyMetrics() {
  const out = {};
  for (const [family, , style, file] of FACES) if (!out[family] && style === "normal" && METRICS[file]) out[family] = METRICS[file];
  for (const [family, , , , file] of LOCAL_ALIASES) if (!out[family] && METRICS[file]) out[family] = METRICS[file];
  return out;
}

// Chasses (unités de police) de la face embarquée la plus proche de (famille, gras, italique) :
// { upm, a: { codepoint: chasse } } ou null. Sert au moteur à reproduire l'arrondi des chasses de
// PowerPoint (1/8 pt par glyphe, calibré : 0,014 % d'erreur) -> mêmes coupures de ligne.
// advanceTable(family, bold, italic)
//   family : nom de famille EMBARQUÉE (FACES) — "Calibri" (alias local) renvoie null : l'appelant
//            parcourt la pile CSS et trouve ensuite "Carlito", aux chasses identiques ;
//   bold / italic : booléens du run de texte.
//   -> { upm, a } (référence partagée, NE PAS MODIFIER) ou null si aucune face de cette famille n'a
//      de table.
//   Usage côté moteur : largeur d'un glyphe en pt = a[codepoint] · taille / upm, puis arrondi au
//   1/8 pt comme PowerPoint.
//   Choix de la face : score = |graisse − graisse voulue| + 1000 si le style (droit/italique) diffère.
//   Le style prime donc toujours sur la graisse ; à défaut de face exacte on prend la plus proche
//   (ex. Poppins gras italique -> BoldItalic ; Liberation Mono italique -> Regular, pénalité 1000).
//   Graisse voulue : 700 si gras, 400 sinon (l'API ne reçoit qu'un booléen : pas de 500/600).
// ADV : advances.json (~450 Ko), chargé paresseusement au 1er appel puis gardé en mémoire ;
//   illisible -> {} et toutes les réponses valent null (le moteur n'applique alors aucune correction).
let ADV = null;
export function advanceTable(family, bold, italic) {
  if (!ADV) { try { ADV = JSON.parse(fs.readFileSync(path.join(FONTS_DIR, "advances.json"), "utf8")); } catch { ADV = {}; } }
  const want = bold ? 700 : 400, style = italic ? "italic" : "normal";
  let best = null, score = Infinity;
  for (const [fam, weight, st, file] of FACES) {
    if (fam !== family || !ADV[file]) continue;
    const w = typeof weight === "number" ? weight : want; // police variable : couvre la graisse voulue
    // (limite : pour une police variable, les chasses stockées sont celles de l'instance par défaut
    // du fichier ; elles peuvent varier avec la graisse réellement affichée)
    const s = Math.abs(w - want) + (st === style ? 0 : 1000);
    if (s < score) { score = s; best = ADV[file]; }     // strictement meilleur : à égalité, la 1re face gagne
  }
  return best;
}
