# Concepts & pièges OOXML (déjà appris par le reverse)

Les règles transverses qui reviennent partout. À garder en tête en lisant la doc.

## Unités
- **EMU** (English Metric Unit) : 914400 EMU/pouce, **12700 EMU/point**, 360000 EMU/cm.
  On rend en **points** (viewBox 960×540) : les `foreignObject` cassent avec des font-size en
  centaines de milliers de px. `u = emu/12700`.
- Angles : **60000e de degré** (rot, gradient ang). Tailles de police : **centièmes de point**
  (`sz="1800"` = 18 pt). Couleurs de dégradé pos : **millièmes de %** (`pos="50000"` = 50 %).
- Keyframes anim `tm` : **millièmes de %** de la durée (0..100000).

## Couleur = le plus piégeux
- `clrScheme` du thème = 12 slots **bruts** : dk1 lt1 dk2 lt2 accent1..6 hlink folHlink.
- **`clrMap` du masque** remappe bg1/tx1/bg2/tx2 → dk/lt. Les decks du corpus ont un **thème sombre**
  (`bg1="dk1"`), donc `schemeClr="bg1"` = sombre et `"tx1"` = clair. **L'ignorer inverse tout.**
- **`clrMapOvr`** (par diapo/layout) peut re-inverser localement — pas encore géré.
- **`phClr`** = « placeholder color » : dans un style de thème (fillStyleLst/bgFillStyleLst), les
  `schemeClr val="phClr"` sont remplacés par la couleur passée par le `fillRef`/`bgRef` qui
  référence ce style (via son `idx`). C'est indirect mais central.
- Modificateurs appliqués dans l'ordre : lumMod (×L), lumOff (+L) en HSL, puis shade/tint (vers
  noir/blanc), alpha. Approximations HSL acceptables.

## Références indirectes (idx)
- `bgRef idx="1001"` → `theme.fmtScheme.bgFillStyleLst[idx-1001]` (1001 = 1er).
- `fillRef idx="N"` → `fillStyleLst[N-1]` (1-based ; 0 = pas de fill). idem lnRef/effectRef.
- `+mn-lt`/`+mj-lt` = police **mineure/majeure** du fontScheme (résoudre, sinon glyphes faux :
  un « 5 » peut sortir « € » avec une police de repli).

## Héritage (le cœur de la fidélité)
- **Composition** d'une diapo (z-order, fond→avant) : bg masque → formes masque → bg layout →
  formes layout → bg diapo → formes diapo. **Ne PAS rendre les placeholders du masque/layout**
  (leur texte = invite « Cliquez pour modifier »).
- **Placeholder** : une forme de la diapo (`p:ph type:idx`) hérite de son homologue du **layout**
  (même type:idx), qui hérite du **master**. Hérite : **géométrie** (xfrm + prstGeom, ex. le
  cercle de découpe d'une image), **remplissage** (carte blanche, cercle de badge), **styles de
  texte** (`lstStyle`). Sans ça : cartes/cercles manquants, texte de mauvaise couleur.
- **Style de run**, du + prioritaire au - : `rPr` du run → `lstStyle` INLINE de la forme →
  `lstStyle` du placeholder de LAYOUT hérité → `lstStyle` du placeholder de MASTER (police/taille
  du titre et du corps) → `p:txStyles` du MASQUE (title/body/other, par niveau lvl1..9) →
  `defaultTextStyle` de la présentation.
- **Spécificité des booléens (`b`/`i`) ≠ « OU »** : un niveau plus spécifique qui définit `@b="0"`
  doit BATTRE un `@b="1"` hérité (ex. layout titre `b="0"` sur master titre `b="1"` → titre NON gras).
  Un simple `a || b || c` laisse le master l'emporter à tort → prendre le 1er niveau qui DÉFINIT `@b`.
- **Zone de texte LIBRE (sans `p:ph`) ≠ placeholder body** : elle ne doit PAS hériter du `bodyStyle`
  du masque (sinon **puces « • » parasites** et couleurs de corps sur des pieds de page / légendes).
  → styleLvls vide, on s'appuie sur le `lstStyle` inline + `defaultTextStyle`. Un vrai placeholder
  (avec `p:ph`, même sans type = idx seul) reste `body`.
- **Couleur texte de dernier recours = `tx1` du thème** (clair en thème sombre), JAMAIS `#000` en dur
  (sinon texte noir invisible sur fond sombre, ex. pied de page « sample » / n° de diapo).
- **Fidélité de largeur = POLICE embarquée** : sans Poppins en `@font-face`, le repli système est plus
  large → le texte déborde (fausse impression de « mauvais placement »). Toujours tester AVEC Poppins.
- **Espaces multiples préservés (`white-space:pre-wrap`)** : PowerPoint positionne parfois du texte
  avec des SUITES d'espaces (ex. « TURBULENCE⎵⎵…⎵vu à » où « vu à » doit tomber entre deux badges
  superposés). HTML fusionne les espaces consécutifs → le texte se décale. `pre-wrap` les garde tout
  en conservant le retour à la ligne. (Complète `trimValues:false` du parseur.)
- **Retrait/puces** : lire `marL`/`indent` du **pPr du paragraphe** (pas que du lstStyle), sinon
  puces à retrait suspendu déséquilibrées → débordent à gauche et se font couper (overflow hidden).
  Garde : `text-indent >= -marL`.

## Texte : autofit
- `normAutofit` : PowerPoint **réduit la police** (fontScale) et l'interligne (lnSpcReduction) pour
  faire tenir le texte. Le fontScale est **parfois stocké**, parfois **calculé à l'affichage**
  (absent du XML) → prévoir un **autofit navigateur** (mesurer scrollHeight/Width, réduire).
- Piège parseur : désactiver `parseTagValue` (sinon « 6 » → nombre → crash) **et** `trimValues`
  (sinon espaces en bord de `<a:t>` coupés → « Marge »+« qu'on » = « Margequ'on »).

## Images
- Placeholder image = `<p:pic>` avec `p:ph type="pic"` ; sa **forme de découpe** (cercle…) vient
  du placeholder de layout. Clipper l'image à la forme (`preserveAspectRatio="slice"`).
- **svgBlip** (extension 2016) : `a:blip` sans `@r:embed` direct, le SVG est dans
  `extLst > ext > asvg:svgBlip @r:embed`. Prioriser le SVG (vectoriel net).
- `wdp` (HD Photo) non affichable navigateur → prendre le fallback ; `emf/wmf` → à rastériser.

## Animation (voir docs/moteur-animation.md, reverse complet)
- Arbre **SMIL** : `mainSeq` → 1 `par` gelé (`delay="indefinite"`) par **clic** → effets →
  *behaviors* (les vraies primitives). Un « effet » (Fade, Fly…) = recette compilée en behaviors.
- « Avec/Après le précédent » = **délais absolus pré-calculés** depuis le clic (pas de chaînage).
- Propriétés animables `ppt_x/y` = **centre** de la forme, en fractions de diapo (0..1).

## Extensions & compat (MCE, Part 3)
- `mc:AlternateContent > mc:Choice/mc:Fallback` : PowerPoint met souvent une version « moderne »
  (Choice) + un repli (Fallback). Bien choisir la branche (on prend à plat = imparfait).
- Beaucoup de `extLst`/`ext uri="{GUID}"` sont des **métadonnées d'édition inutiles au rendu**
  (`a16:creationId`, `p14:modId`). D'autres portent du vrai contenu (svgBlip, hlinkClr, imgEffect).

## PowerPoint ≠ spec
- PowerPoint n'écrit qu'un **sous-ensemble** du format + des **extensions non spécifiées**. Donc :
  toujours croiser la doc (spécifié) avec un **vrai .pptx** (ce que PowerPoint écrit) et le rendu
  **LibreOffice** (ce que ça doit donner). MS-OI29500 documente les écarts d'implémentation.

## Filtres SVG : espace colorimétrique
- **`color-interpolation-filters="sRGB"` OBLIGATOIRE** sur les filtres qui font de la luminosité/
  contraste/saturation (a14:brightnessContrast, grayscl, duotone). Défaut SVG = **linearRGB** →
  les valeurs sont gamma-décodées, un intercept négatif écrase toute l'image en noir. PowerPoint
  opère en sRGB. (Bug trouvé : fond photo assombri rendu tout noir, deck-12 d.28.)

## Groupes
- **Rotation/miroir du GROUPE** (`p:grpSpPr>a:xfrm @rot/@flipH/@flipV`) : à appliquer au `<g>`
  englobant, autour du centre du groupe. Les membres sont placés dans la boîte NON tournée puis
  tout le groupe tourne. Oublié = groupes tournés rendus à plat (formes inclinées, deck-12 d.10).

## Oracle de fidélité
- **PowerPoint COM** (`Presentation.Slides(i).Export png`) = référence PIXEL-PERFECT pour comparer.
  LibreOffice→PDF→SVG = fidélité proche mais PAS pixel-perfect (mêmes chevauchements que le moteur
  maison sur les boîtes spAutoFit + badges superposés). Cf. [[pptx-html-libreoffice-route]].
