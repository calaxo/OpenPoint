# Matrice de couverture OOXML → moteur

Statut : ✅ implémenté · 🟡 partiel · ❌ à faire · ➖ hors scope/inutile au rendu.
Code = fonction dans `renderer/pptx-render.mjs` (sauf indication). Réf = § ECMA-376 (voir spec-map).

Validé : 548 diapos (deck A + deck B), 0 erreur. Ce tableau dit où creuser ensuite.

## Package / OPC (Part 2)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| zip + `_rels/*.rels` | ✅ | `slideRels`, `relsOf` | résolution r:id → cible |
| `[Content_Types].xml` | ➖ | — | pas nécessaire (on déduit par extension) |
| `presentation.xml` sldSz | ✅ | `renderSlide` | taille diapo (16:9/4:3) |
| `presentation.xml` sldIdLst / ordre | 🟡 | `build*.mjs` | on prend l'ordre des fichiers slideN, pas sldIdLst |
| `defaultTextStyle` | ✅ | `loadContext`/`lvlStyle` | chargé par niveau ; repli final (sz/color/font) pour texte hors placeholder |

## PresentationML — structure (§19)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| composition master→layout→slide | ✅ | `renderSlide` | z-order correct, bg de chaque niveau |
| `p:cSld/p:spTree` ordonné | ✅ | `walk` | z-order = ordre document (preserveOrder) |
| `p:bg` (bgPr solid/grad/blip) | ✅ | `bgRect` | |
| `p:bg` (bgRef idx→theme) | ✅ | `bgRect` | bgFillStyleLst[idx-1001] + phClr |
| `p:clrMap` (masque) | ✅ | `renderSlide` | remap bg1/tx1/bg2/tx2→dk/lt (thème sombre !) |
| `p:clrMapOvr` (override diapo/layout) | ❌ | — | override par diapo (absent du corpus jusqu'ici) |
| `p:txStyles` (title/body/other) | ✅ | `loadContext` | défauts hérités par niveau |
| `p:ph` (placeholder type/idx) | ✅ | `renderSp/renderPic` | héritage géométrie+fill+lstStyle depuis layout |
| héritage placeholder multi-niveau (layout→master) | 🟡 | `placeholderGeom` | on lit le layout ; le master en 2e recours = partiel |
| `p:transition` (transitions diapo) | ❌ | — | fondu/pousser entre diapos |
| notes (`notesSlide`) | ➖ | — | pas affiché (contenu formateur) |

## Formes & conteneurs (§19 / §20)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| `p:sp` | ✅ | `renderSp` | |
| `p:pic` | ✅ | `renderPic` | + clip à la forme + svgBlip |
| `p:cxnSp` (connecteurs) | 🟡 | `renderSp` | rendu comme forme ; routage bentConnector approx |
| `p:grpSp` (+ chOff/chExt) | ✅ | `renderGroup` | transform enfants |
| `p:graphicFrame > a:tbl` | ✅ | `renderTable` | tableaux |
| `p:graphicFrame > c:chart` | ❌ | — | graphiques (aucun dans le corpus) |
| `p:graphicFrame > dgm` (SmartArt) | ❌ | — | non rencontré |
| `a:xfrm` off/ext/rot/flipH/flipV | ✅ | `getXfrm/applyTf` | |

## Géométrie (§20.1.2, §20.1.9) — moteur générique `preset-geom.mjs`
| Élément | Statut | Code | Notes |
|---|---|---|---|
| `custGeom` moveTo/lnTo/cubicBez/quadBez/close | ✅ | `custGeomDef`+`evalGeom` | nœud ORDONNÉ (deepFind) → ordre exact |
| `custGeom` **arcTo** | ✅ | `emitPath` | arc elliptique centré sur position courante (§20.1.9.4) |
| `custGeom` gdLst/avLst **formules** | ✅ | `evalFmla` | 17 opérateurs `*/ +- +/ ?: abs sqrt max min mod pin sin cos tan at2 cat2 sat2 val` |
| `prstGeom` **187 formes** (ST_ShapeType complet) | ✅ | `presetGeomD` | définitions officielles ECMA-376 bakées (`preset-shapes.json`) |
| — dont flèches courbes/uturn/bent, callouts, cloud, étoiles, rubans, blockArc, actionButtons | ✅ | `evalGeom` | validé visuellement (deck-10 d.28 uturnArrow, d.34) |
| guides intégrés (l/t/r/b/hc/vc/ss/ls/wdN/hdN/ssdN + angles cd2/3cd4…) | ✅ | `baseVars` | §20.1.10.56 |
| détail au trait `fill="none"` (glyphes actionButtons) + ombrage 3D `darken/lighten` | 🟡 | `evalGeom` | silhouette OK, détail interne non tracé (rare dans le corpus) |

## Couleur (§20.1.4)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| srgbClr/schemeClr/sysClr/prstClr/scrgbClr | ✅ | `resolveColorFlat` | hslClr ❌ |
| phClr (substitution) | ✅ | `paint` | fillRef/bgRef |
| mods lumMod/lumOff | ✅ (approx HSL) | `resolveColorFlat` | |
| mods shade/tint/alpha | ✅ (approx) | `resolveColorFlat` | |
| mods satMod/hueMod/gray/comp/inv… | ❌ | — | rares |

## Remplissages (§20.1.4)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| solidFill | ✅ | `fillFromSpPrFlat` | |
| gradFill linéaire (lin ang) | ✅ | `gradId` | |
| gradFill radial/rect (path) | 🟡 | `gradId` | rendu radial générique (pas les foci exacts) |
| blipFill (image dans forme) | ✅ | `renderSp`/`renderPic` | clip à la forme |
| blipFill srcRect (recadrage) | ✅ | `renderPic` | l/t/r/b millièmes % → image agrandie+décalée+clippée ; validé d.5 deck-08 |
| blipFill tile (motif répété) | ❌ | — | traité comme stretch |
| pattFill (hachures) | ❌ | — | |
| fillRef (style thème idx) | ✅ | `fillFromSpPrFlat` | fillStyleLst[idx-1] |

## Ligne (§20.1.4 / §20.1.8)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| ln w + solidFill / lnRef | ✅ | `renderSp` | |
| prstDash (pointillés) | ✅ | `dashArray` | dash/dot/lgDash… |
| ln cap/join/cmpd (double) | ❌ | — | |
| headEnd/tailEnd (flèches de trait) | ❌ | — | connecteurs fléchés |
| gradient/pattern sur ligne | ❌ | — | |

## Effets (§20.1.4)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| outerShdw | ✅ | `effectFilter` | flou+offset+flood mergé |
| glow | ✅ | `effectFilter` | halo coloré (morpho+blur+flood), doublé ; validé d.67 deck-04 |
| softEdge | ✅ | `effectFilter` | flou alpha + re-découpe |
| innerShdw / reflection | ❌ | — | reflection = copie miroir dégradée (rare : 1 dans le corpus) |
| effets d'image : grayscl / duotone / alphaModFix | ✅ | `imageFilter` | feColorMatrix + componentTransfer ; opacité |
| a14:imgEffect brightnessContrast / saturation | ✅ | `imageFilter`+`deepFindFlat` | feComponentTransfer linéaire + saturate ; validé deck-12 d.10 (12 images) |
| a14:imgEffect colorTemperature / sharpenSoften | ❌ | — | température couleur / netteté (approx. complexe) |
| a:clrChange (transparence couleur) | ❌ | — | |
| 3D (sp3d/scene3d/bevel) | ❌ | — | rare |

## Texte (§21.1)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| txBody > p > r > t (multi-run) | ✅ | `renderText` | foreignObject |
| bodyPr anchor (vertical) | ✅ | `renderText` | t/ctr/b |
| bodyPr insets (lIns…) | ✅ | `renderText` | |
| bodyPr **normAutofit** fontScale/lnSpcReduction | ✅ | `renderText` | + autofit navigateur si non stocké |
| bodyPr spAutoFit / noAutofit | 🟡 | — | la forme ne grandit pas ; autofit navigateur compense |
| bodyPr wrap / cols (colonnes) / rot texte | ❌ | — | texte vertical, multi-colonnes |
| pPr algn (h) | ✅ | `renderText` | |
| pPr marL/indent (retrait) | ✅ | `renderText` | pPr prioritaire + garde anti-débordement |
| pPr lvl (niveau) | ✅ | `renderText` | |
| pPr lnSpc / spcBef / spcAft | ✅ | `renderText` | spcPct (millièmes %) + spcPts (centièmes pt) → line-height / margin |
| pPr tabLst (tabulations) | ❌ | — | |
| puces buChar/buAutoNum/buNone/buFont | ✅ | `renderText`/`bulletOf` | |
| puces buClr/buSzPct (couleur/taille puce) | 🟡 | — | couleur=1er run, taille=run |
| rPr sz/b/i/u | ✅ | `renderText` | u simple |
| rPr strike/spc/baseline/cap | ✅ | `renderText` | barré (line-through), espacement (centièmes pt), exposant/indice (millièmes %), MAJUSCULES/petites caps ; validé d.16 « 40ᵉᵐᵉ » |
| rPr kern | 🟡 | — | crénage auto peu visible à l'échelle diapo |
| rPr solidFill (couleur run) | ✅ | `renderText` | |
| rPr gradient/pattern text fill | ❌ | — | texte en dégradé |
| rPr latin/ea/cs + réfs `+mn-lt/+mj-lt` | ✅ | `resolveFont`/`fontFamily` | repli serif/sans |
| rPr ln (contour de texte) / effectLst | ❌ | — | texte détouré/ombré |
| rPr hlinkClick + ahyp:hlinkClr | ❌ | — | liens + couleur de lien |
| fld (numéro de diapo/date) | 🟡 | `renderText` | rendu comme run, valeur brute |
| br (saut de ligne) | 🟡 | — | traité par séparation de paragraphes |

## Thème (§20.1)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| clrScheme (12 slots) | ✅ | `loadContext` | |
| fontScheme major/minor (latin) | ✅ | `loadContext` | ea/cs ❌ |
| fmtScheme fillStyleLst/bgFillStyleLst | ✅ | `loadContext` | ordonné |
| fmtScheme lnStyleLst/effectStyleLst | 🟡 | `loadContext` (ln chargé, non utilisé) | effets de style thème ❌ |

## Tableaux (§21.2)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| tblGrid/gridCol, tr h, tc | ✅ | `renderTable` | |
| tcPr bordures lnL/R/T/B | ✅ | `renderTable` | |
| tcPr fill / anchor / marges | 🟡 | `renderTable` | fill+anchor ✓ ; marges = défaut renderText |
| gridSpan/rowSpan/hMerge/vMerge | ✅ | `renderTable` | |
| tableStyles.xml (header/bandes) | ✅ | `loadContext`/`renderTable` | firstRow/band1H/band2H |
| tableStyles firstCol/lastCol/lastRow | ❌ | — | |

## Animation / timing (§19.5)  — voir aussi docs/moteur-animation.md
| Élément | Statut | Code | Notes |
|---|---|---|---|
| tnLst/par/seq, cTn, cond (clics) | ✅ | `extractTiming` | étapes de clic |
| set (visibility) | ✅ | player | cache/montre |
| animEffect (fade/wipe/circle/box…) | ✅ | player CSS | entrées |
| animMotion (path) | ✅ | `parseMotion`+player | translation le long du chemin (SMIL additive) ; validé deck-05 d.10 |
| animRot | ✅ | `parseMotion`+player | rotation centrée (getBBox) ; validé deck-12 d.44 (spin 360°) |
| anim (ppt_x/y/w/h keyframes/fmla) | 🟡 | — | motion/scale par keyframes fines non rejouées (animMotion couvre le mvt explicite) |
| animScale / animClr | ❌ | — | scale centré non exprimable en 1 transform additif ; changement de couleur |
| exit / emph | 🟡 | player | exit=hide ; emph via animMotion/animRot ✅, pulse/grow ❌ |
| delay/dur (timing) | ✅ | `extractTiming`/player | |
| cible = **groupe** (spid de grpSp) | ✅ | `renderGroup` | groupe enveloppé dans `<g data-spid>` → s'anime en bloc |
| cible = **image / tableau** (data-spid sur pic/graphicFrame) | ✅ | `renderPic`/`renderTable` | 99,3 % des cibles animées résolues (vs ~90 % avant) |
| excl / interactiveSeq (déclencheurs) | ❌ | — | |
| bldLst/bldP (build par paragraphe) | ✅ | `renderText`(data-par) + player | spTgt>txEl>pRg → paragraphes révélés un par un ; validé sample d.26 |

## Extensions (Part 3 / MCE)
| Élément | Statut | Code | Notes |
|---|---|---|---|
| asvg:svgBlip (image SVG) | ✅ | `renderPic` | prioritaire sur raster |
| mc:AlternateContent | 🟡 | (parser à plat prend une branche) | pas de choix Choice/Fallback explicite |
| ahyp:hlinkClr / a14:imgEffect | ❌ | — | |
| a16:creationId / p14:modId / extLst d'édition | ➖ | — | inutile au rendu |

## Média
| Type | Statut | Notes |
|---|---|---|
| png/jpg/gif/bmp | ✅ | data URI |
| svg (svgBlip) | ✅ | |
| emf/wmf (vectoriel Windows) | ❌ | à rastériser en amont |
| wdp (HD Photo) | ➖ | non affichable navigateur ; on prend le fallback |
| audio/vidéo | ❌ | |

## Polices
| Sujet | Statut | Notes |
|---|---|---|
| Poppins embarqué | ✅ | `build*.mjs` @font-face |
| Arial/Times/… | 🟡 | repli serif/sans ; dépend du client |
| polices embarquées `ppt/fonts/*` (fntdata) | ❌ | à extraire + @font-face |
