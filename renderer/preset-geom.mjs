// Moteur de géométrie DrawingML générique (ECMA-376 §20.1.9) — SANS PowerPoint.
// Évalue les FORMULES DE GUIDES (17 opérateurs) et construit le chemin SVG pour :
//   - les 187 formes preset (données bakées : preset-shapes.json, cf. gen-preset-shapes.mjs)
//   - les custGeom (mêmes formules + pathLst ordonné)
// Remplace les ~30 approximations manuelles par les définitions officielles exactes.
//
// Modèle : def = { av:[[name,fmla]], gd:[[name,fmla]], paths:[{w,h,fill,stroke,cmds}] }
//   cmds : ["M",x,y] ["L",x,y] ["C",x1,y1,x2,y2,x3,y3] ["Q",x1,y1,x2,y2]
//          ["A",wR,hR,stAng,swAng] ["Z"]   (valeurs = nom de guide OU nombre, en chaîne)
//
// ─── COMMENT ÇA MARCHE (vue d'ensemble) ─────────────────────────────────────────────────────────
// LE PROBLÈME
//   Une forme PowerPoint n'est pas stockée comme un dessin : le fichier dit seulement
//   <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>.
//   Le dessin réel est défini une fois pour toutes par la norme, dans presetShapeDefinitions.xml
//   (annexe de ECMA-376 Partie 1) : pour chaque forme, des guides (variables calculées), un chemin et
//   un rectangle de texte. gen-preset-shapes.mjs convertit ce XML en preset-shapes.json ; ce module
//   l'interprète.
//
// PLACE DANS LE PIPELINE
//   pptx-render.mjs appelle, pour chaque p:sp :
//     - presetGeomParts(prst, Wpt, Hpt, adj)  -> chemins SVG (remplissage + contour) EN POINTS ;
//     - presetTextRect(prst, Wemu, Hemu, adj) -> rectangle de texte EN EMU (même unité que les
//       insets lIns/tIns… du bodyPr, auxquels il est ensuite combiné) ;
//     - evalGeom / evalGeomParts(custGeomDef(...)) pour une a:custGeom (forme libre), dont le def est
//       construit à la volée depuis le XML de la diapo, au même format que preset-shapes.json.
//   Le calcul est INDÉPENDANT DE L'UNITÉ : W et H peuvent être en points, EMU ou pixels, le résultat
//   est dans la même unité (sauf les littéraux de chemin mis à l'échelle via path.w/path.h, voir emitPath).
//
// FORMAT D'UNE DÉFINITION (une entrée de preset-shapes.json)
//   {
//     av:    [[nom, formule], ...]   guides AJUSTABLES (avLst) avec leur valeur par défaut ; la forme
//                                    de la diapo peut en surcharger certains (poignées jaunes) ;
//     gd:    [[nom, formule], ...]   guides CALCULÉS (gdLst), évalués dans l'ordre, chacun pouvant
//                                    utiliser les précédents ;
//     rect:  [l, t, r, b]            rectangle de texte (noms de guides ou nombres) — optionnel ;
//     paths: [{ w, h, fill, stroke, cmds }]
//            w/h    : taille du repère propre au chemin (0 = même repère que la forme) ;
//            fill   : 'norm' | 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess' ;
//            stroke : false si le chemin ne doit pas être tracé ;
//            cmds   : commandes (voir plus haut), valeurs = nom de guide ou nombre en chaîne.
//   }
//
// UNITÉS DE LA NORME (à connaître pour lire les formules)
//   - angles : 60000es de degré (5400000 = 90°, 21600000 = 360°) ;
//   - valeurs d'ajustement : en général 100000es d'une dimension de référence (« val 50000 » =
//     50 %), converties par une formule du type « */ ss a 100000 » (ss = petit côté) ;
//   - longueurs : dans l'unité de W/H (repère de la forme, origine en haut à gauche, y vers le bas).
//
// EXEMPLE COMPLET (roundRect, adj = 16667 par défaut, W = 200, H = 100)
//   ss = 100 ; a = pin 0 adj 50000 = 16667 ; x1 = ss·a/100000 ≈ 16,67 (rayon d'angle) ;
//   il = x1·29289/100000 ≈ 4,88 (29289/100000 = 1 − cos 45° : point du quart de cercle à 45°) ;
//   rect = [il, il, r−il, b−il] -> le texte est rentré de ≈ 4,88 de chaque côté.
//
// DÉPENDANCES : node:fs (lecture synchrone de preset-shapes.json au chargement du module).
// PIÈGE : PRESETS est un objet PARTAGÉ et mutable (on y ajoute upArrow ci-dessous) ; ne pas le
//   modifier ailleurs sans le savoir.
// ────────────────────────────────────────────────────────────────────────────────────────────────

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 1 — Chargement des définitions officielles (+ correctif upArrow)
// ════════════════════════════════════════════════════════════════════════════════════════════════
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dir = dirname(fileURLToPath(import.meta.url));
// PRESETS : { nomDeForme (ST_ShapeType, ex. "roundRect") -> définition }, lu UNE fois à l'import.
//   Exporté pour les outils (couverture, tests) qui veulent énumérer les formes connues.
export const PRESETS = JSON.parse(readFileSync(join(__dir, 'preset-shapes.json'), 'utf8'));

// COMBLE une OMISSION du XML ECMA : `upArrow` est absent de presetShapeDefinitions.xml
// (il contient upDownArrow EN DOUBLE et upArrowCallout, mais pas upArrow — bug de la spec).
// Sans lui, une flèche vers le haut tombait en rectangle (barre). On le définit comme le
// MIROIR VERTICAL de downArrow (apex en haut au lieu d'en bas). Cf. deck-12 d.20 (vecteurs PGF).
// Lecture des guides :
//   adj1 = épaisseur du FÛT en 100000es de la largeur ; adj2 = hauteur de la POINTE en 100000es du
//   petit côté (ss), bornée par maxAdj2 = 100000·h/ss pour que la pointe ne dépasse pas la hauteur.
//   dy1 = hauteur de la pointe ; dx1 = demi-épaisseur du fût (w·a1/200000 = (w·a1/100000)/2) ;
//   x1/x2 = bords gauche/droit du fût autour du centre hc.
// Chemin : base de la pointe à gauche (l, dy1) -> fût jusqu'en bas (b) -> retour à dy1 à droite ->
//   bord droit (r, dy1) -> apex (hc, t) -> fermeture. Pas de rect : le texte occupe la boîte entière.
if (!PRESETS.upArrow) PRESETS.upArrow = {
  av: [['adj1', 'val 50000'], ['adj2', 'val 50000']],
  gd: [['maxAdj2', '*/ 100000 h ss'], ['a1', 'pin 0 adj1 100000'], ['a2', 'pin 0 adj2 maxAdj2'],
       ['dy1', '*/ ss a2 100000'], ['dx1', '*/ w a1 200000'], ['x1', '+- hc 0 dx1'], ['x2', '+- hc dx1 0']],
  paths: [{ w: 0, h: 0, fill: 'norm', stroke: true, cmds: [
    ['M', 'l', 'dy1'], ['L', 'x1', 'dy1'], ['L', 'x1', 'b'], ['L', 'x2', 'b'],
    ['L', 'x2', 'dy1'], ['L', 'r', 'dy1'], ['L', 'hc', 't'], ['Z']] }],
};

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 2 — Guides intégrés (variables prédéfinies disponibles dans toute formule)
// ════════════════════════════════════════════════════════════════════════════════════════════════

// constantes d'angle (60000es de degré) — §20.1.10 + guides intégrés
//   cd = « circle degrees » = cercle complet = 360° = 21600000 ; cd2 = 180°, cd4 = 90°, cd8 = 45°,
//   3cd4 = 270°, 3cd8 = 135°, 5cd8 = 225°, 7cd8 = 315°. '0' permet d'écrire un angle nul par nom.
const ANG = { '0':0, cd8:2700000, cd4:5400000, cd2:10800000, cd:21600000,
  '3cd4':16200000, '3cd8':8100000, '5cd8':13500000, '7cd8':18900000 };

// variables intégrées (guides prédéfinis) calculées depuis w/h réels — §20.1.10.56
// baseVars(W, H) -> table { nom -> nombre } des guides intégrés, point de départ de toute évaluation.
//   l/t/r/b : bords gauche/haut/droit/bas (l = t = 0 : repère local de la forme) ;
//   w/h : largeur/hauteur ; hc/vc : centre horizontal/vertical ;
//   ss/ls : petit côté / grand côté (« short side » / « long side ») — base des ajustements
//           proportionnels qui doivent rester ronds quel que soit l'étirement ;
//   wdN/hdN : largeur/hauteur divisée par N ; ssdN : petit côté divisé par N.
//   Le code génère quelques noms de plus que la liste de la norme (ex. hd10) : sans effet si inutilisés.
function baseVars(W, H) {
  const v = { l:0, t:0, r:W, b:H, w:W, h:H, hc:W/2, vc:H/2,
    ss:Math.min(W,H), ls:Math.max(W,H) };
  for (const n of [2,3,4,5,6,8,10,12]) { v['wd'+n]=W/n; v['hd'+n]=H/n; } // wd2 = W/2, hd4 = H/4…
  for (const n of [2,4,6,8,16,32]) v['ssd'+n]=v.ss/n;                   // ssd2 = ss/2…
  Object.assign(v, ANG);           // 3cd4, cd2, … accessibles par nom dans les formules
  return v;
}

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 3 — Évaluation des formules de guides
// ════════════════════════════════════════════════════════════════════════════════════════════════

// isNum(t) : vrai si t est un nombre JS, ou une chaîne qui COMMENCE comme un nombre littéral
//   ("12", "-5400000", ".5", "-.5"). Les noms de guides commencent par une lettre… sauf les
//   constantes d'angle "3cd4", "3cd8", "5cd8", "7cd8" qui commencent par un chiffre : isNum les
//   prend pour des nombres ! Voilà pourquoi tok() cherche D'ABORD dans vars, et pourquoi A()
//   (emitPath) teste ANG avant isNum.
const isNum = (t) => typeof t === 'number' || /^-?\.?\d/.test(t);
// tok(t, vars) : résout un opérande de formule ou de chemin -> nombre.
//   Ordre : nombre JS tel quel -> nom de guide connu (intégré, av ou gd déjà évalué) -> littéral
//   numérique -> 0. Le repli à 0 rend le moteur tolérant aux fichiers qui citent un guide inexistant
//   (ou un guide défini PLUS LOIN dans gdLst : l'évaluation étant séquentielle, il vaut encore 0).
function tok(t, vars) {
  if (typeof t === 'number') return t;
  if (t in vars) return vars[t];
  if (isNum(t)) return parseFloat(t);
  return 0;                         // guide inconnu -> 0 (robuste)
}
// rad(a) : angle en 60000es de degré -> radians (÷ 60000 -> degrés, × π/180 -> radians).
const rad = (sixtyK) => sixtyK / 60000 * Math.PI / 180;

// évalue une formule de guide (§20.1.9.11) — 17 opérateurs
// evalFmla(fmla, vars) -> nombre
//   fmla : chaîne « opérateur arg1 [arg2 [arg3]] » séparée par des espaces (notation préfixée de la
//          norme, attribut @fmla de <a:gd>), ex. "*/ ss adj 100000", "pin 0 adj 50000".
//          Un nombre JS est aussi accepté (String() -> "25000" -> opérateur inconnu -> 0 : attention,
//          une surcharge d'ajustement doit donc être une FORMULE "val 25000", pas un nombre nu).
//   vars : table des guides déjà connus.
//   Les arguments x, y, z sont notés a(), b(), c() et résolus paresseusement (tok) au besoin.
//   Table de la norme reproduite (x, y, z = arguments 1, 2, 3) :
//     val x          -> x                         (constante ou copie d'un guide)
//     */  x y z      -> x · y / z                 (« multiplier-diviser », l'opérateur le plus courant)
//     +-  x y z      -> x + y − z
//     +/  x y z      -> (x + y) / z
//     ?:  x y z      -> x > 0 ? y : z             (« if-else » : strictement positif)
//     abs x          -> |x|
//     sqrt x         -> √x
//     max x y / min x y
//     mod x y z      -> √(x² + y² + z²)           (norme d'un vecteur, PAS le modulo)
//     pin x y z      -> y borné à [x, z]          (« épingle » : clamp)
//     sin x y        -> x · sin(y)                (y en 60000es de degré)
//     cos x y        -> x · cos(y)
//     tan x y        -> x · tan(y)
//     at2 x y        -> atan2(y, x) en 60000es de degré (résultat dans ]−180°, 180°])
//     cat2 x y z     -> x · cos(atan2(z, y))      (« cosine arctan »)
//     sat2 x y z     -> x · sin(atan2(z, y))      (« sine arctan »)
//   Opérateur inconnu -> 0 (tolérance plutôt qu'exception).
function evalFmla(fmla, vars) {
  const p = String(fmla).trim().split(/\s+/);            // ["*/", "ss", "adj", "100000"]
  const op = p[0];
  const a = () => tok(p[1], vars), b = () => tok(p[2], vars), c = () => tok(p[3], vars);
  switch (op) {
    case 'val': return a();
    case '*/':  return a() * b() / c();
    case '+-':  return a() + b() - c();
    case '+/':  return (a() + b()) / c();
    case '?:':  return a() > 0 ? b() : c();
    case 'abs': return Math.abs(a());
    case 'sqrt':return Math.sqrt(a());
    case 'max': return Math.max(a(), b());
    case 'min': return Math.min(a(), b());
    case 'mod': return Math.hypot(a(), b(), c());              // sqrt(x²+y²+z²)
    case 'pin': { const x=a(), y=b(), z=c(); return y < x ? x : (y > z ? z : y); } // si x > z, x l'emporte
    case 'sin': return a() * Math.sin(rad(b()));
    case 'cos': return a() * Math.cos(rad(b()));
    case 'tan': return a() * Math.tan(rad(b()));
    case 'at2': return Math.atan2(b(), a()) * 180 / Math.PI * 60000;   // arctan(y/x) -> 60000es °
    case 'cat2':return a() * Math.cos(Math.atan2(c(), b()));           // x·cos(arctan(z/y))
    case 'sat2':return a() * Math.sin(Math.atan2(c(), b()));           // x·sin(arctan(z/y))
    default: return 0;
  }
}

// construit la table de guides : builtins -> avLst (avec overrides) -> gdLst séquentiel
// buildVars(def, W, H, adj) -> table complète { nom -> nombre }
//   adj : { nomGuide: formule } venu de <a:prstGeom><a:avLst><a:gd name fmla/> de la DIAPO
//         (poignées d'ajustement déplacées par l'auteur). Seuls les noms présents dans def.av sont
//         pris en compte ; un nom inconnu de la forme est ignoré.
//   Ordre imposé par la norme : guides intégrés, puis avLst (valeur de la diapo si fournie, sinon
//   défaut de la définition), puis gdLst DANS L'ORDRE (un guide ne voit que ceux définis avant lui).
//   Un guide peut écraser un nom intégré (ex. un gd nommé « w ») : c'est le comportement voulu.
function buildVars(def, W, H, adj) {
  const vars = baseVars(W, H);
  for (const [name, fmla] of (def.av || [])) {
    vars[name] = evalFmla((adj && adj[name] != null) ? adj[name] : fmla, vars); // surcharge diapo prioritaire
  }
  for (const [name, fmla] of (def.gd || [])) vars[name] = evalFmla(fmla, vars);
  return vars;
}

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 4 — Traduction des chemins DrawingML en données de chemin SVG (attribut d)
// ════════════════════════════════════════════════════════════════════════════════════════════════

// f(n) : arrondi au centième pour des chemins compacts ; convertit −0 en 0 (sinon « -0 » dans le SVG).
const f = (n) => { const r = Math.round(n * 100) / 100; return Object.is(r, -0) ? 0 : r; };

// un <path> -> segment de chemin SVG. Coord = guide (espace-forme, non mis à l'échelle)
// ou nombre littéral (mis à l'échelle par w/h du path si présents).
// emitPath(path, vars, W, H) -> chaîne « M… L… A… Z » (se termine par une espace)
//   Repère du chemin (a:path @w @h) : un chemin peut déclarer son propre repère, ex. w=21600 h=21600
//   dans une custGeom ; un littéral x = 10800 vaut alors W/2. sx/sy font cette conversion.
//   Les GUIDES, eux, sont déjà calculés à partir de W/H (repère de la forme) : pas de mise à l'échelle.
//   Correspondance des commandes (§20.1.9 : moveTo, lnTo, cubicBezTo, quadBezTo, arcTo, close) :
//     M x y -> M ; L x y -> L ; C (3 points) -> C ; Q (2 points) -> Q ; A wR hR stAng swAng -> A SVG
//     (conversion ci-dessous) ; Z -> Z.
//   (cx, cy) = point courant, nécessaire pour arcTo qui part de là ; (sX, sY) = début du sous-chemin,
//   où Z ramène le point courant.
function emitPath(path, vars, W, H) {
  const sx = path.w ? W / path.w : 1, sy = path.h ? H / path.h : 1; // w/h = 0 -> repère de la forme
  // X/Y : littéral -> mis à l'échelle ; nom de guide (ou nombre JS) -> valeur telle quelle
  const X = (t) => (typeof t !== 'number' && isNum(t)) ? parseFloat(t) * sx : tok(t, vars);
  const Y = (t) => (typeof t !== 'number' && isNum(t)) ? parseFloat(t) * sy : tok(t, vars);
  // A : angle -> radians. Test ANG AVANT isNum : "3cd4" ressemble à un nombre (voir isNum).
  const A = (t) => (t in ANG) ? rad(ANG[t]) : (isNum(t) ? rad(parseFloat(t)) : rad(tok(t, vars)));
  let d = '', cx = 0, cy = 0, sX = 0, sY = 0;
  for (const cmd of path.cmds) {
    const k = cmd[0];
    if (k === 'M') { cx = X(cmd[1]); cy = Y(cmd[2]); sX = cx; sY = cy; d += `M${f(cx)} ${f(cy)} `; }
    else if (k === 'L') { cx = X(cmd[1]); cy = Y(cmd[2]); d += `L${f(cx)} ${f(cy)} `; }
    else if (k === 'C') { const x1=X(cmd[1]),y1=Y(cmd[2]),x2=X(cmd[3]),y2=Y(cmd[4]),x3=X(cmd[5]),y3=Y(cmd[6]);
      d += `C${f(x1)} ${f(y1)} ${f(x2)} ${f(y2)} ${f(x3)} ${f(y3)} `; cx = x3; cy = y3; }
    else if (k === 'Q') { const x1=X(cmd[1]),y1=Y(cmd[2]),x2=X(cmd[3]),y2=Y(cmd[4]);
      d += `Q${f(x1)} ${f(y1)} ${f(x2)} ${f(y2)} `; cx = x2; cy = y2; }
    else if (k === 'A') {
      // arc elliptique : centre calé pour que le point de départ = position courante (§20.1.9.4)
      // DrawingML décrit un arc de façon RELATIVE (rayons wR/hR, angle de départ stAng, balayage
      // swAng, départ = point courant) alors que SVG veut une extrémité ABSOLUE + drapeaux. Donc :
      //   1. retrouver le centre (ecx, ecy) de l'ellipse : point courant − (wR·cos st, hR·sin st) ;
      //   2. calculer l'extrémité à l'angle st + sw ;
      //   3. large-arc = balayage > 180° ; sweep = 1 si balayage positif. Angles positifs = sens
      //      horaire à l'écran (y vers le bas), ce qui est aussi le sens « 1 » de SVG.
      // Le point de l'ellipse est pris sous forme PARAMÉTRIQUE (wR·cos, hR·sin). Pour un cercle
      // (wR = hR) c'est l'angle géométrique exact.
      const wR = X(cmd[1]), hR = Y(cmd[2]), st = A(cmd[3]), sw = A(cmd[4]);
      const ecx = cx - wR * Math.cos(st), ecy = cy - hR * Math.sin(st);
      const end = st + sw, ex = ecx + wR * Math.cos(end), ey = ecy + hR * Math.sin(end);
      const large = Math.abs(sw) > Math.PI ? 1 : 0, sweep = sw >= 0 ? 1 : 0;
      if (Math.abs(sw) >= 2 * Math.PI) {
        // cercle complet : 2 demi-arcs (SVG ne dessine pas 360° d'un coup)
        // (un arc SVG dont l'extrémité = le départ est ignoré) : on passe par le point diamétralement
        // opposé (st + π). Prévu pour |sw| = 360° exactement (cas des formes preset) ; au-delà,
        // le second demi-arc vise l'extrémité st + sw, sans refaire de tour supplémentaire.
        const mx = ecx + wR * Math.cos(st + Math.PI), my = ecy + hR * Math.sin(st + Math.PI);
        d += `A${f(wR)} ${f(hR)} 0 1 ${sweep} ${f(mx)} ${f(my)} A${f(wR)} ${f(hR)} 0 1 ${sweep} ${f(ex)} ${f(ey)} `;
      } else {
        d += `A${f(wR)} ${f(hR)} 0 ${large} ${sweep} ${f(ex)} ${f(ey)} `;
      }
      cx = ex; cy = ey;                                    // l'extrémité devient le point courant
    }
    else if (k === 'Z') { d += 'Z '; cx = sX; cy = sY; }   // fermeture : retour au début du sous-chemin
    // commande inconnue : ignorée silencieusement
  }
  return d;
}

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 5 — API publique : géométrie d'une définition / d'un preset par nom
// ════════════════════════════════════════════════════════════════════════════════════════════════

// def -> chemin SVG (silhouette : paths remplis + défaut ; ignore fill='none'/ombrage darken/lighten)
// evalGeom(def, W, H, adj) -> UNE chaîne d, tous chemins remplissables concaténés.
//   Chemins écartés :
//     fill='none'                  -> purement décoratifs au trait (ex. plis d'un parchemin) ;
//     fill='darken*'/'lighten*'    -> doublons d'une facette à assombrir/éclaircir (effet 3D de
//                                     bevel, cube, can…) : concaténés, ils referaient la même zone.
//   Repli 1 : si TOUS les chemins sont écartés, on les concatène quand même (sinon rien à voir).
//   Repli 2 : définition absente/vide ou chemin résultant vide -> rectangle 0,0 → W,H.
//   Version « silhouette » : pour séparer remplissage et contour, voir evalGeomParts.
export function evalGeom(def, W, H, adj) {
  if (!def || !def.paths || !def.paths.length) return `M0 0 H${f(W)} V${f(H)} H0 Z`;
  const vars = buildVars(def, W, H, adj);
  let d = '';
  for (const path of def.paths) {
    if (path.fill === 'none') continue;                       // détail au trait (v2 : non tracé)
    if (/^(darken|lighten)/.test(path.fill)) continue;        // ombrage 3D -> ignoré (rendu plat)
    d += emitPath(path, vars, W, H);
  }
  if (!d) d = def.paths.map((p) => emitPath(p, vars, W, H)).join(''); // tout en 'none' -> tracer quand même
  return d.trim() || `M0 0 H${f(W)} V${f(H)} H0 Z`;
}

// def -> { fill, stroke } : chemins SÉPARÉS pour le remplissage et le tracé, en honorant les attributs
// par sous-path (a:path @fill / @stroke). Certaines formes (arc, chord…) ont un path « secteur » à
// remplir SANS tracé (stroke=false) et un path « courbe » à tracer SANS remplissage (fill=none) :
// tout fusionner traçait les rayons du secteur (ex. 45° en camembert au lieu d'un arc fin).
// evalGeomParts(def, W, H, adj) -> { fill: d, stroke: d }
//   fill   : chemins à REMPLIR (mêmes exclusions qu'evalGeom : 'none', 'darken*', 'lighten*') ;
//   stroke : chemins à TRACER (tous sauf stroke=false) — les chemins d'ombrage darken/lighten sont
//            donc bien tracés s'ils ont un contour, fidèle à PowerPoint.
//   Exemple « arc » (voir preset-shapes.json) : chemin 1 = secteur (fill norm, stroke false),
//   chemin 2 = courbe seule (fill none, stroke true) -> fill = secteur, stroke = arc fin.
//   Repli : fill vide -> tous les chemins ; stroke peut rester VIDE (l'appelant ne trace alors rien).
//   Les guides sont évalués UNE fois pour tous les chemins (buildVars partagé).
export function evalGeomParts(def, W, H, adj) {
  if (!def || !def.paths || !def.paths.length) { const d = `M0 0 H${f(W)} V${f(H)} H0 Z`; return { fill: d, stroke: d }; }
  const vars = buildVars(def, W, H, adj);
  let fillD = '', strokeD = '';
  for (const path of def.paths) {
    const seg = emitPath(path, vars, W, H);
    const fillable = path.fill !== 'none' && !/^(darken|lighten)/.test(path.fill);
    const strokable = path.stroke !== false;              // stroke absent = true (défaut de la norme)
    if (fillable) fillD += seg;
    if (strokable) strokeD += seg;
  }
  if (!fillD) fillD = def.paths.map((p) => emitPath(p, vars, W, H)).join('');
  return { fill: fillD.trim() || `M0 0 H${f(W)} V${f(H)} H0 Z`, stroke: strokeD.trim() };
}

// preset par nom (ST_ShapeType). adj = map { nomGuide: fmla|nombre } depuis l'avLst de la forme.
// presetGeomD(name, W, H, adj) -> d (silhouette) ou null si la forme n'est pas dans PRESETS.
//   (Côté « nombre » : voir la note d'evalFmla — un nombre nu y est évalué à 0 ; pptx-render.mjs
//   transmet bien la chaîne @fmla, ex. "val 25000".)
//   null plutôt qu'un rectangle : l'appelant (presetPath de pptx-render.mjs) choisit son repli.
export function presetGeomD(name, W, H, adj) {
  const def = PRESETS[name];
  if (!def) return null;                                      // inconnu -> laisser l'appelant décider
  return evalGeom(def, W, H, adj);
}

// presetGeomParts(name, W, H, adj) -> { fill, stroke } (voir evalGeomParts) ou null si forme inconnue.
//   C'est l'entrée utilisée par pptx-render.mjs pour chaque p:sp à géométrie prédéfinie (W/H en pt).
export function presetGeomParts(name, W, H, adj) {
  const def = PRESETS[name];
  if (!def) return null;
  return evalGeomParts(def, W, H, adj);
}

// Rectangle de TEXTE d'une forme prédéfinie (<a:rect l t r b/> de presetShapeDefinitions) : PowerPoint
// y pose le texte, PUIS applique les insets du bodyPr. roundRect : rétréci de 0,29 x le rayon (≈ 6 pt
// sur une carte, ≈ 10 pt dans une pastille ronde) — l'ignorer décalait le texte vers le bord.
// -> { l, t, r, b } dans l'unité de W/H, ou null (forme inconnue / sans rect = boîte entière).
// Règle reproduite (ECMA-376 §20.1.9.22, <a:rect>) : le texte d'une forme se place dans le rectangle
//   défini par 4 guides (bords gauche, haut, droit, bas) exprimés dans le repère de la forme, et non
//   dans la boîte englobante. Ex. ellipse : carré inscrit à 45° (il = hc − wd2·cos 45°, soit ≈ 14,6 %
//   de marge de chaque côté) ; rightArrow : bande verticale du fût (y1..y2), de l jusqu'à x2 ; roundRect : rentré de il = rayon·(1 − cos 45°).
//   l, t, r, b sont des COORDONNÉES (pas des marges) : largeur utile = r − l, hauteur = b − t.
//   pptx-render.mjs l'appelle en EMU (m.cx/m.cy) pour le combiner directement aux insets du bodyPr.
//   Les rect non numériques (NaN / Infinity, ex. division par zéro sur une forme de taille nulle)
//   donnent null -> l'appelant revient à la boîte entière.
export function presetTextRect(name, W, H, adj) {
  const def = PRESETS[name];
  if (!def || !def.rect) return null;
  const vars = buildVars(def, W, H, adj);              // mêmes guides que pour le dessin
  const [l, t, r, b] = def.rect.map((k) => tok(k, vars)); // noms de guides ("il") ou intégrés ("l", "r")
  return [l, t, r, b].every(Number.isFinite) ? { l, t, r, b } : null;
}
