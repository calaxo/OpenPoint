// ═══ Polices web des outils (viewer, player, sweep) — MÊMES polices que la plateforme hôte ═══
//
// POURQUOI ce fichier
//   Les outils de ce dépôt (build.mjs, build-player.mjs, sweep.mjs) affichaient les diapos avec les
//   anciennes polices TTF (Poppins + substituts), SANS le calage de métriques appliqué en production.
//   Résultat : une comparaison pixel (sweep) mesurait un rendu qui n'est jamais celui servi.
//   On passe désormais par animFonts.js, le module de polices de l'hôte, pour obtenir EXACTEMENT le même CSS.
//
// CE QUE FAIT animFonts.getFontsCss()
//   - lit chaque fichier .woff2 de renderer/fonts/ (liste FACES dans animFonts.js) ;
//   - l'inline en data: URI (l'hôte l'injecte dans une iframe à origine opaque, qui ne peut rien
//     télécharger : il faut que les polices soient DANS le document) ;
//   - ajoute ascent-override / descent-override / line-gap-override à partir de fonts/metrics.json :
//     ce sont les métriques « win » (usWinAscent/usWinDescent) que PowerPoint utilise pour caler ses
//     lignes, alors que le navigateur prendrait sinon hhea/typo — d'où un texte décalé de ~3 pt.
//
// CE QUE FAIT fontsCssFor(texte)
//   Le CSS complet pèse ~4 Mo. On ne garde que les @font-face dont la famille apparaît dans le texte
//   passé (le moteur écrit toujours la pile complète dans font-family, secours compris). C'est la
//   même règle que facesFor() dans l'intégration hôte.
//
// UTILISATION TYPIQUE (build.mjs, build-player.mjs, sweep.mjs)
//   const css = fontsCssFor(svgOuHtmlDeLaDiapo);   // puis <style>${css}</style> dans la page produite
//
// PIÈGES
//   - Filtrage par SOUS-CHAÎNE : « Cabin » serait retenu par un texte contenant « Cabinet » ; c'est
//     sans danger (au pire une police de trop), jamais une police manquante.
//   - La famille 'Calibri' des alias locaux est retenue si le texte cite « Calibri » — et aussi pour
//     « Calibri Light », qui contient « Calibri » : même logique, au pire une règle de trop.
//   - Premier appel bloquant (lecture synchrone de tous les woff2 via getFontsCss) : sans importance
//     pour des scripts en ligne de commande.

// ═══ Cache des règles, découpées une fois ═══
import { getFontsCss } from './animFonts.js';

// faces : [{ css: "@font-face{…}", family: "Poppins" }], construit au premier appel puis réutilisé.
let faces = null;

/** CSS @font-face (polices embarquées + métriques PowerPoint) restreint aux familles citées par `text`. */
// fontsCssFor(text) -> chaîne CSS (règles séparées par des retours à la ligne), éventuellement vide.
//   text : n'importe quel texte où le moteur a écrit ses font-family (SVG, HTML) ; null/undefined -> ''.
export function fontsCssFor(text) {
  if (!faces) {
    // Découpe du CSS complet en règles : /@font-face\{[^}]*\}/ suffit car une règle @font-face ne
    // contient aucune accolade interne (la data: URI base64 n'utilise que A-Z a-z 0-9 + / =).
    faces = (getFontsCss().match(/@font-face\{[^}]*\}/g) || []).map((css) => ({
      css,
      family: (css.match(/font-family:'([^']+)'/) || [])[1], // nom entre apostrophes, écrit par animFonts.faceCss / aliasCss
    }));
  }
  const src = String(text || '');
  // règle sans famille lisible (ne devrait pas arriver) : gardée par prudence
  return faces.filter((f) => !f.family || src.includes(f.family)).map((f) => f.css).join('\n');
}
