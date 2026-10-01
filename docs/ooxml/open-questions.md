# Questions ouvertes — à éclaircir avec les docs à venir

À cocher/répondre au fur et à mesure. Format : [ ] question — pourquoi ça compte — où chercher.

## Texte / héritage
- [ ] Ordre EXACT de résolution d'un `rPr` manquant : run → lstStyle forme → lstStyle placeholder
      layout → lstStyle placeholder master → txStyles master → defaultTextStyle présentation →
      thème. Y a-t-il des priorités par propriété ? (§21.1 lstStyle / §19 txStyles)
- [ ] `defaultTextStyle` de `presentation.xml` : quand s'applique-t-il vraiment ? (dernier recours)
- [ ] `lnSpc` (spcPct vs spcPts), `spcBef`/`spcAft` : conversion exacte en interligne/espacement.
- [ ] `kern`, `spc` (letter-spacing), `baseline` (exposant/indice), `cap="all"` : rendu exact.
- [ ] Autofit : algorithme EXACT de calcul du fontScale quand non stocké (comment PowerPoint le
      calcule). Croiser avec `docs/dynamique-runtime.md` (hooker le calcul dans PPCORE ?).

## Couleur
- [ ] Formules exactes de lumMod/lumOff/shade/tint (espace de couleur : sRGB ? linéaire ? HSL ?).
      Nos approximations HSL suffisent-elles au pixel ? (§20.1.2.3.x color transforms)
- [ ] `clrMapOvr` : sémantique précise (overrideClrMapping vs masterClrMapping).

## Géométrie
- [ ] `custGeom` : évaluer les **formules de guides** (`gdLst` : `*/ +- +/ ?: abs sqrt sin cos
      tan at2 cat2 sat2 mod pin val min max`). Nécessaire pour les formes ajustables exactes.
- [ ] `arcTo` (wR hR stAng swAng) → conversion en arc SVG.
- [ ] Table des **presetGeom** (ST_ShapeType, ~180 formes) : où trouver leurs définitions
      officielles (presetShapeDefinitions) pour les générer toutes ? (LibreOffice `oox` en a une copie)

## Remplissages / effets
- [ ] `blipFill` : `srcRect` (recadrage %), `tile` (motif), `stretch/fillRect`, `duotone`,
      `alphaModFix` (opacité), `a14:imgEffect` — sémantique et rendu.
- [ ] `gradFill path` : foci exacts (rect/circle/shape), `tileRect`, `flip`.
- [ ] Effets : `outerShdw` params exacts (sx/sy/kx/ky rotation), `glow`, `reflection`, `softEdge`.

## Animation (le gros morceau après)
- [ ] Sémantique complète de `anim` (keyframes `tav` + `fmla`) : rejouer mouvement/échelle
      exactement (voir langage de formules, `docs/moteur-animation.md` §6 et LibreOffice
      `slideshow/source/engine/smilfunctionparser.cxx`).
- [ ] `animMotion` (path type SVG), `animRot`, `animScale`, `animClr` : rendu.
- [ ] Cible = **groupe** : propagation aux membres (comment PowerPoint résout spTgt sur grpSp).
- [ ] `bldP` (build par paragraphe/niveau) : apparition ligne par ligne.
- [ ] `iterate` (animation lettre/mot par lettre), easing (`accel`/`decel`, `spd`).
- [ ] Déclencheurs interactifs (`excl`, `interactiveSeq`, `cond evt=onClick tgt=spTgt`).

## Tables / charts
- [ ] `tableStyles.xml` : résolution complète (firstCol/lastCol/lastRow, bordures de style,
      `tcTxStyle` couleur/gras du texte selon la zone).
- [ ] Charts (`c:chart`) : structure chart1.xml + xlsx embarqué → rendre barres/lignes/camemberts.

## Polices
- [ ] Polices embarquées `ppt/fonts/*.fntdata` : format (TTF obfusqué ?), extraction, @font-face.
- [ ] Substituts métriques libres (Liberation Sans/Serif ↔ Arial/Times) : où les prendre (LibreOffice).

## Divers
- [ ] `mc:AlternateContent` : règle de choix Choice/Fallback (quelle branche PowerPoint prend).
- [ ] `p:transition` : catalogue des transitions et leur rendu.
- [ ] VML (Part 4, legacy) : présent dans de vieux fichiers ? rendu nécessaire ?
