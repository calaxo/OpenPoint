# Lire un PPTX entier sans PowerPoint

Objectif concret pour le LMS : afficher fidèlement un `.pptx` **sans** PowerPoint. Ce document
décrit le moteur de rendu maison (`renderer/pptx-render.mjs`) et ce qu'il faut comprendre pour
« lire » un PPTX. Il s'appuie sur le reverse du moteur d'animation (voir `docs/`) mais traite
ici le **rendu statique**, qui est la vraie difficulté (les animations, elles, sont un arbre
simple déjà élucidé).

## 1. Un PPTX, c'est quoi à lire

Un zip OPC. L'essentiel pour rendre une diapo :

```
ppt/presentation.xml         -> taille de diapo (sldSz, EMU), liste des diapos
ppt/slides/slideN.xml        -> le contenu : p:cSld/p:spTree (formes), p:bg, p:timing (anim)
ppt/slides/_rels/slideN.xml.rels -> résout les r:id (images, layout)
ppt/slideLayouts/slideLayoutK.xml + _rels -> le layout (déco + placeholders)
ppt/slideMasters/slideMasterM.xml + _rels -> le masque (fond, déco, clrMap, txStyles)
ppt/theme/theme1.xml         -> clrScheme (couleurs), fontScheme, fmtScheme (styles fill/line)
ppt/media/*                  -> images
```

**Ordre de composition d'une diapo** (z-order, du fond vers l'avant) :
`fond du masque → formes du masque → fond du layout → formes du layout → fond de la diapo →
formes de la diapo`. Les formes de masque/layout qui sont des **placeholders** ne sont PAS
rendues (ce ne sont que des gabarits de style/position ; les rendre affiche le texte fantôme
« Cliquez pour modifier… » — le piège classique). On garde leur déco (logos, filets) et le fond.

## 2. Les 6 sous-systèmes à implémenter

### 2.1 Unités
Tout est en **EMU** (914400/pouce, 12700/point). On rend en **points** (viewBox 960×540) :
les foreignObject HTML cassent avec des `font-size` en centaines de milliers de px. Angles en
60000e de degré, tailles de police en centièmes de point, keyframes en millièmes de %.

### 2.2 Couleurs (le plus piégeux)
- `clrScheme` du thème = 12 slots bruts (`dk1 lt1 dk2 lt2 accent1..6 hlink folHlink`).
- **`clrMap` du masque** remappe `bg1/tx1/bg2/tx2 → dk1/lt1/…`. **Déterminant** : le deck du corpus
  a un thème *sombre* (`bg1="dk1"`), donc un `schemeClr="bg1"` = sombre, `"tx1"` = clair.
  L'ignorer donne des couleurs inversées et un fond blanc au lieu de sombre.
- Modificateurs à appliquer : `lumMod/lumOff` (en HSL), `shade`, `tint`, `alpha`, `satMod`.
- **`phClr`** = couleur de substitution : un style de thème (`fillStyleLst`, `bgFillStyleLst`)
  contient `schemeClr="phClr"` remplacé par la couleur passée par le `fillRef`/`bgRef`.

### 2.3 Remplissages
- Direct : `solidFill`, `gradFill` (→ `linearGradient` SVG), `blipFill` (image), `noFill`.
- **Par référence de style** : beaucoup de formes n'ont pas de fill explicite mais un
  `<p:style><a:fillRef idx="N"><a:schemeClr val="accent1"/>` → on va chercher le N-ième fill
  du `fillStyleLst` du thème, teinté par `phClr=accent1`. Idem `bgRef idx="1001"` → fond =
  `bgFillStyleLst[idx-1001]` teinté par sa couleur (c'est ainsi qu'est défini le fond du masque).

### 2.4 Géométrie
- `custGeom` : chemin DrawingML (`moveTo/lnTo/cubicBezTo/quadBezTo/close`) → path SVG (général,
  gère tout tracé). Les arcs (`arcTo`) restent à faire.
- `prstGeom` : ~187 formes prédéfinies. On implémente les courantes (rect, roundRect avec `adj`,
  ellipse, triangle, diamant, parallélogramme, hexagone, ligne…) + fallback rectangle.
- Transform : `off/ext`, `rot` (60000e°), `flipH/flipV`, et pour les groupes le repère enfant
  (`chOff/chExt`) → mise à l'échelle des enfants.

### 2.5 Texte
- `txBody > bodyPr` (ancrage vertical, insets, wrap) `> a:p > a:r > a:rPr + a:t`.
- Rendu en **`<foreignObject>` HTML** (le wrap et les polices sont gérés par le navigateur ;
  faire le retour à la ligne en SVG pur demanderait de mesurer le texte).
- **Polices du thème** : un run avec `typeface="+mn-lt"` / `"+mj-lt"` = police mineure/majeure
  du thème (à résoudre, sinon fallback → glyphes faux, ex. « 5 » rendu « € »). Les polices
  réelles (Poppins) doivent être **embarquées** (`@font-face`) pour la fidélité métrique.
- **Héritage de style de run** (du + prioritaire au -) : `rPr` du run → `lstStyle` INLINE de la
  forme → `lstStyle` du **placeholder de LAYOUT** hérité (couleur/taille/police par niveau) →
  `lstStyle` du **placeholder de MASTER** (c'est LUI qui porte souvent la police / la taille du
  titre et du corps) → `p:txStyles` du **MASQUE** (`titleStyle/bodyStyle/otherStyle`) →
  `defaultTextStyle` de la présentation. Le niveau layout est indispensable (sans lui, le texte des
  cartes sort clair au lieu de sombre) ; le niveau master aussi (sans lui, un titre hérite d'un
  `titleStyle` générique — vu : « Arial 14 » au lieu du « Space Grotesk 30 » du master).
- **Autofit** (`<a:normAutofit>`) : PowerPoint réduit la police (et l'interligne via
  `lnSpcReduction`) pour faire tenir le texte. Le `fontScale` stocké est appliqué ; quand il est
  absent (PowerPoint calcule à l'affichage), un **autofit navigateur** (mesure + réduction) prend
  le relais — sinon le texte déborde et est coupé.

### 2.5bis Héritage de placeholder (déterminant)
Une forme-texte de la diapo est souvent un **placeholder** (`p:ph type=.. idx=..`) qui hérite de
son homologue dans le **layout** (même `type:idx`), lui-même du masque. On hérite : la
**géométrie** (si la diapo n'a pas d'`xfrm`), le **remplissage** (fond de carte, cercle de badge)
et les **styles de texte** (`lstStyle`). C'est ainsi que le template « cartes objectifs » marche :
le layout définit une carte `roundRect fill=bg2 txcol=tx1` (et un badge `fill=tx1 txcol=bg1`), la
diapo n'apporte que le texte. **Ne pas** rendre les placeholders du masque/layout eux-mêmes (leur
texte est l'invite « Cliquez pour modifier »).

### 2.6 Images
`p:pic > blipFill > a:blip @r:embed` → `_rels` → `media/*`. Embarquées en `data:` (autonome).
`emf/wmf` (vectoriel Windows) non supportés par `<image>` → à rastériser en amont.

## 3. État & fidélité

Validé visuellement sur le deck sample (28 diapos) contre un rendu de référence LibreOffice
(`renderer/ref/`, `soffice --convert-to pdf` — LibreOffice n'est pas PowerPoint). Rendu très
proche : fonds sombres du thème, swooshes/logos du masque, badges bleus (accent1),
coins arrondis exacts, texte hérité clair, objectifs, définitions, pieds de page.

Corrigé depuis v1 : résolution des polices de thème `+mn-lt/+mj-lt`, embarquement de Poppins,
autofit (fontScale + navigateur), héritage complet de placeholder (géométrie/fill/texte depuis le
layout), clrMap sombre, phClr, fillRef/bgRef. Les captures diapos 1/2/5/6 correspondent au réel.

**Limites connues restantes** (par ordre d'impact) :
- Pas d'`arcTo` dans custGeom ni la plupart des `prstGeom` exotiques (fallback rectangle).
- Léger chevauchement badge/texte sur certaines cartes (inset gauche non appliqué au pixel près).
- Puces (`buChar`/`buAutoNum`) non rendues ; tabulations approximées.
- Dégradés linéaires seulement (pas radiaux/path) ; fond `blipFill` en motif non géré.
- `emf/wmf` (images vectorielles Windows) ignorées → à rastériser en amont.
- clrMapOvr (override de mapping par diapo) non géré (absent de ce deck).

## 4. Reproduire

```
npm install                      # fast-xml-parser
node renderer/build.mjs          # -> renderer/viewer.html (les 28 diapos)
node renderer/build.mjs 1 5 8    # quelques diapos
# référence : renderer/ref/sample.pdf (LibreOffice)
```

## 5. Prochaines étapes vers un lecteur LMS

1. **Unzip côté navigateur** (fflate/JSZip) : déposer un `.pptx` → rendu sans serveur ni PowerPoint.
2. **Rebrancher la timeline d'animation** (déjà parsée, `tools/parse-timing.mjs`) par-dessus le
   SVG : chaque forme a un `id` (spid), on rejoue les behaviors au clic (POC `poc/player.html`).
3. Combler les limites v1 selon les decks réels (arcTo, prstGeom manquants, polices embarquées).
