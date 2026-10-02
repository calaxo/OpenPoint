// ═══ [CŒUR PPTX] ═══ carte d'ensemble : docs/ARCHITECTURE.md (§4) · consommé par build.mjs + build-player.mjs
// Moteur de rendu PPTX -> SVG, en JS pur, SANS PowerPoint.
//
// Ce fichier n'est plus qu'une FAÇADE : il ré-exporte l'API publique. Le code vit dans renderer/pptx/.
//
// API :
//   renderSlide(pptDir, slideNo, { embedImages }) -> { svg, w, h }      (pptx/slide/render-slide.mjs)
//   extractTiming(pptDir, slideNo) -> { steps, animatedSpids }           (pptx/timing.mjs)
//   extractTransition(pptDir, slideNo) -> { type, dur, dir } | null      (pptx/timing.mjs)
//   slideOrder(pptDir) -> [n…] numéros de slideN.xml dans l'ordre de présentation (pptx/parts.mjs)
//   loadContext(pptDir, themePath) -> contexte de thème                  (pptx/theme.mjs)
// `pptDir` = dossier `ppt/` d'un .pptx DÉJÀ décompressé (voir lib/deck.mjs pour décompresser).
//
/* ═══════════════════════════════════════════════════════════════════════════════════════════════
 * GUIDE DE LECTURE
 * ═══════════════════════════════════════════════════════════════════════════════════════════════
 *
 * 1. CARTE DES MODULES (renderer/pptx/) — les flèches = « importe »
 *
 *    xml.mjs ........ parseurs P (ordonné) / PF (plat), cache de parse, helpers d'arbre, esc, serialize
 *    units.mjs ...... u() : EMU -> points
 *    color.mjs ...... couleur DrawingML -> {r,g,b,a} -> CSS
 *    theme.mjs ...... loadContext() : palette, polices, fmtScheme, styles de tableaux
 *    parts.mjs ...... .rels, images (data URI), placeholders du layout/master, ordre des diapos
 *    text-styles.mjs  puces (bulletOf), p:txStyles du master, p:defaultTextStyle
 *    fonts.mjs ...... police -> pile CSS, ligne de base, chasses 1/8 pt, hauteur de cellule
 *    symbols.mjs .... Wingdings / Symbol -> Unicode
 *    math.mjs ....... OMML -> HTML
 *    geometry.mjs ... prstGeom / custGeom -> chemins SVG (via ../preset-geom.mjs)
 *    timing.mjs ..... animations (<p:timing>) et transitions (<p:transition>)
 *    slide/
 *      render-slide.mjs  renderSlide : héritage diapo -> layout -> master -> thème, fond, composition, <defs>
 *      walk.mjs ........ parcours du spTree (ordre z), groupes, transformations de repère
 *      shape.mjs ....... p:sp / p:cxnSp (géométrie, fill, trait, flèches, image de remplissage)
 *      picture.mjs ..... p:pic (découpe, srcRect, filtres d'image)
 *      frame.mjs ....... p:graphicFrame : tableau, graphique (../chart-render.mjs), SmartArt, OLE
 *      text.mjs ........ texte -> <foreignObject> HTML (cascade de styles, puces, tabulations…)
 *      paint.mjs ....... remplissages, dégradés, motifs, pointillés, pointes de flèche
 *      effects.mjs ..... ombre / halo / bords adoucis, filtres d'image
 *
 * 2. PIPELINE DE renderSlide(pptDir, slideNo)
 *    a) Chaîne d'héritage par les .rels : slideN.xml -> slideLayoutX.xml -> slideMasterY.xml -> themeZ.xml
 *       (un deck multi-master a PLUSIEURS thèmes : le bon est celui du master de la diapo).
 *    b) loadContext(thème) ; clrMap + txStyles du master ; defaultTextStyle de presentation.xml.
 *    c) placeholderGeom : placeholders du layout (clé "type:idx") dont héritent les <p:ph> de la diapo.
 *    d) Composition : UN fond (diapo > layout > master), puis formes du master (spid préfixé 'm',
 *       placeholders sautés), du layout ('l', idem), puis de la diapo (spid brut).
 *    e) Chaque forme passe par walk -> renderSp / renderPic / renderGroup / renderGraphicFrame.
 *       Les <defs> (dégradés, motifs, filtres, clips, marqueurs) sont accumulés dans `rc` pendant
 *       le parcours puis écrits en tête du SVG.
 *    f) Sortie : <svg viewBox="0 0 Wpt Hpt"><defs>…</defs><rect blanc/>…formes…</svg>.
 *    Toutes les fonctions de slide/ reçoivent `rc` (état du rendu de la diapo) en 1er argument ;
 *    ses champs sont documentés en tête de slide/render-slide.mjs.
 *
 * 3. CONVENTIONS
 *    - UNITÉS. OOXML mesure en EMU : 914 400 EMU = 1 pouce, 12 700 EMU = 1 pt. Le SVG est en POINTS
 *      (u(v) = v / 12700 arrondi au 1/100). Dans le HTML des foreignObject, 1 px CSS = 1 pt de diapo.
 *      Angles en 1/60 000 de degré ; pourcentages en 1/1000 de % (100000 = 100 %) ; tailles de police
 *      en 1/100 de pt (@sz="1800" = 18 pt).
 *    - DEUX FORMES D'ARBRE XML (fast-xml-parser) :
 *        * P  (preserveOrder) -> [{ 'a:tag': [enfants…], ':@': {attributs} }] : garde l'ORDRE (z-order,
 *          runs/a:br, commandes custGeom). Lu avec tagOf/at/kids/byTag/first/deepFind.
 *        * PF (plat) -> { 'a:tag': { '@attr': …, 'a:enfant': {…} | [{…}] } } : lecture directe, mais
 *          perd l'ordre entre balises différentes, et un enfant unique n'est pas un tableau.
 *      Stratégie : parcours en ordonné, puis pour chaque forme on re-sérialise le nœud et on le
 *      re-parse en plat pour lire ses propriétés. Quand l'ordre compte, on revient au nœud ordonné.
 *    - ATTRIBUTS LUS PAR LES VIEWERS (contrat avec renderer/viewer/) :
 *        * data-spid="<id>" sur le <g> de chaque forme / groupe / tableau (= p:cNvPr@id ; préfixé 'm'
 *          ou 'l' pour le master / layout). Les animations ciblent les spid BRUTS de la diapo.
 *        * data-par="<n>" sur chaque <div> de paragraphe (animations par paragraphe).
 *        * class="af" data-af="norm|fixed" data-wrap="none" sur le conteneur de texte.
 *    - IDENTIFIANTS SVG préfixés 's<N>_' : plusieurs diapos peuvent cohabiter dans une même page.
 *
 * 4. PIÈGES CONNUS / AVANT DE MODIFIER
 *    - Le rendu a été calé AU PIXEL sur des exports PowerPoint : les constantes (1,2 ; 0,75 L ; 1/8 pt ;
 *      0,2 px…) sont MESURÉES. Toute modification se revalide sur un corpus (npm test, qa.mjs).
 *    - esc() échappe ' en &apos; (entité NOMMÉE) : fast-xml-parser ne décode pas &#39; au re-parse.
 *    - Les arbres renvoyés par cachedParse sont PARTAGÉS entre diapos : ne JAMAIS les muter.
 *    - renderSlide n'est pas réentrant (ctx muté pendant le rendu, police mineure dans fonts.mjs).
 *    - Jamais de commentaire JS dans une chaîne / un gabarit `…` : il partirait dans le SVG.
 *    - Non gérés : custGeom HÉRITÉ du layout (repli rectangle),
 *      géométrie de placeholder héritée du MASTER (seul son style de texte est lu).
 * ═══════════════════════════════════════════════════════════════════════════════════════════════ */

export { renderSlide, slideSize } from './pptx/slide/render-slide.mjs';
export { extractTiming, extractTransition } from './pptx/timing.mjs';
export { loadContext } from './pptx/theme.mjs';
export { slideOrder } from './pptx/parts.mjs';

// Helpers purs exportés pour les tests unitaires (test/*.test.js) et les outils.
export { u } from './pptx/units.mjs';
export { esc, serialize, tagOf, at, kids, byTag, first, deepFindFlat } from './pptx/xml.mjs';
export { hslFromRgb, rgbFromHsl, cssColor, resolveColorFlat } from './pptx/color.mjs';
export { symFontKind, symToUnicode, bulletSym, symbolParts, symbolText } from './pptx/symbols.mjs';
export { fontFamily, baselineShift, advanceFix, cellTextHeight } from './pptx/fonts.mjs';
export { bulletOf } from './pptx/text-styles.mjs';
export { mathTypeset } from './pptx/math.mjs';
