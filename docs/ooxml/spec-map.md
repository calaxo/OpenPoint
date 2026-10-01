# Carte de la spec ECMA-376 / ISO-IEC 29500

Pour situer rapidement un excerpt de doc. La spec est en **4 parties** (ISO) / 5 (ECMA 1re éd.).
La quasi-totalité de ce qui nous concerne est dans **Part 1** (Fundamentals & Markup Language
Reference, ~5560 p.).

## Parties

| Partie | Contenu | Pour nous |
|---|---|---|
| Part 1 | Fondamentaux + **référence de tous les langages** (Wordprocessing/Spreadsheet/**Presentation**ML, **DrawingML**, Shared ML) | ★ l'essentiel |
| Part 2 | **Open Packaging Conventions** (OPC) : zip, `_rels`, `[Content_Types].xml`, signatures | structure du .pptx |
| Part 3 | Markup Compatibility & Extensibility (**MCE** : `mc:AlternateContent`, `extLst`) | extensions (svgBlip, hlinkClr…) |
| Part 4 | Transitional Migration (legacy, **VML**, différences vs ECMA 1re éd.) | vieux fichiers, VML |

## Part 1 — sections clés (numérotation ISO 29500-1)

| § | Sujet | Éléments phares |
|---|---|---|
| §13 | Shared MLs | — |
| **§19** | **PresentationML** | `presentation`, `sld`, `sldLayout`, `sldMaster`, `cSld`, `spTree`, `p:sp/pic/grpSp/graphicFrame/cxnSp`, `p:ph`, `p:bg`, `p:clrMap`/`clrMapOvr`, `p:txStyles`, `p:timing`, `p:transition`, notes |
| §19.3 | Presentation & slides | `p:presentation` (sldSz, sldIdLst, notesSz, defaultTextStyle), `p:sld/sldLayout/sldMaster` |
| §19.5 | **Animation (timing)** | `tnLst`, `par/seq/excl`, `cTn`, `cond/stCondLst`, `cBhvr`, `animEffect/anim/animMotion/animScale/animRot/animClr/set/cmd`, `tgtEl/spTgt`, `tav/tavLst` (fmla), `bldLst/bldP` |
| **§20** | **DrawingML — Framework** | `a:xfrm`, `a:off/ext`, `graphic/graphicData` |
| **§20.1** | **DrawingML — Main** | formes, styles, couleurs, effets, texte, géométries (le gros pour nous) |
| §20.1.2 | Geometry | `prstGeom` (+ liste des presets ST_ShapeType), `custGeom` (`pathLst`, `moveTo/lnTo/cubicBezTo/quadBezTo/arcTo/close`, `gdLst/avLst`) |
| §20.1.3 | Coordinate systems & transforms | EMU, rot, flip |
| §20.1.4 | Styles & color | `solidFill/gradFill/blipFill/pattFill/noFill/grpFill`, `srgbClr/schemeClr/sysClr/prstClr/scrgbClr/hslClr` + modificateurs, `ln`, `effectLst` (`outerShdw/innerShdw/glow/reflection/softEdge/…`), `effectStyleLst` |
| §20.1.5 | 3D | `sp3d`, `scene3d`, `bevel`, `camera`, `lightRig` |
| §20.1.10 | Simple types (ST_*) | ST_ShapeType (presets), ST_LineCap, ST_PresetLineDashVal, unités |
| **§21.1** | **DrawingML — Text** | `txBody`, `bodyPr` (anchor, insets, `normAutofit/spAutoFit/noAutofit`, wrap, cols, spcCol), `lstStyle`, `p`, `pPr` (algn, marL, indent, lvl, `lnSpc`, `spcBef/spcAft`, `buChar/buAutoNum/buNone/buFont/buSzPct/buClr`, `tabLst`), `r`, `rPr` (sz, b, i, u, strike, kern, spc, baseline, cap, `latin/ea/cs`, `solidFill`, `hlinkClick`, `highlight`, `ln`, `effectLst`), `fld`, `br`, `defRPr`, `endParaRPr` |
| **§21.2** | **DrawingML — Tables** | `a:tbl`, `tblPr` (+ `tableStyleId`, firstRow/band…), `tblGrid/gridCol`, `tr`, `tc`, `tcPr` (`lnL/R/T/B`, fill, marges, anchor, `gridSpan/rowSpan/hMerge/vMerge`) ; styles dans `tableStyles.xml` |
| §21.2.x | Table styles | `a:tblStyleLst > a:tblStyle` (wholeTbl, band1H/2H, firstRow/lastRow, firstCol/lastCol → `tcStyle` fill/border + `tcTxStyle`) |
| §21.3 | Charts | `c:chart` (chart1.xml + xlsx embarqué) — **non implémenté** |
| §21.4 | Diagrams (SmartArt) | `dgm:` — non rencontré |

## Fichiers dans un .pptx (OPC, Part 2)

```
[Content_Types].xml            types MIME des parties
_rels/.rels                    relation racine -> presentation.xml
ppt/presentation.xml (+ .rels) sldSz, sldIdLst -> slides, masters, notesMaster, tableStyles, theme
ppt/slides/slideN.xml (+ _rels) contenu diapo (spTree, bg, timing, transition)
ppt/slideLayouts/slideLayoutK.xml (+ _rels)
ppt/slideMasters/slideMasterM.xml (+ _rels)  clrMap, txStyles, bg
ppt/theme/themeT.xml           clrScheme, fontScheme, fmtScheme
ppt/tableStyles.xml            styles de tableaux (par GUID)
ppt/notesSlides/, notesMasters/, presProps.xml, viewProps.xml
ppt/media/*                    images/vidéos/son
ppt/embeddings/*               xlsx des charts, OLE
ppt/fonts/*                    polices embarquées (fntdata) — non géré
```

## Extensions Microsoft (Part 3 / MCE) rencontrées

Namespaces d'extension via `extLst > ext uri="{GUID}"` (souvent dans `mc:AlternateContent`) :
- `asvg:svgBlip` (2016/SVG) — image SVG vectorielle (fallback raster dans `a:blip`). ✅ géré.
- `ahyp:hlinkClr` (2018/hyperlinkcolor) — couleur de lien par run. ❌.
- `a14:` (2010/main) — imgProps, effets images (`a14:imgLayer`, `a14:imgEffect`). ❌.
- `a16:creationId`, `p14:modId` — ids d'édition, **inutiles au rendu**.
