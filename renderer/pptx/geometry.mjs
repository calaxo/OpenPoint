// ═══ [PPTX · GÉOMÉTRIE] ═══ prstGeom / custGeom -> chemins SVG (délègue à preset-geom.mjs).
import { presetGeomD, evalGeom, evalGeomParts } from '../preset-geom.mjs';
import { tagOf, at, kids, byTag, first } from './xml.mjs';
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
export { presetPath, custGeomPath, custGeomParts };
