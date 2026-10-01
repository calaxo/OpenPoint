# Notes de recherche web — internes de l'animation PowerPoint (sources)

Synthèse d'une recherche documentaire (2026-07-22) croisant spec primaire (ECMA/ISO 29500,
MS-OI29500, MS-OE376), code open-source (LibreOffice, ONLYOFFICE) et docs Microsoft. Sert de
contrepoint aux findings reverse (voir `docs/`). Les points marqués ⚠️ sont non confirmés par
source primaire.

## Modèle de timing (spec)
- Racine `<p:timing>` placée **après `<p:transition>`, avant `<p:extLst>`**.
- 3 conteneurs de timing : `par` (parallèle), `seq` (séquentiel), `excl` (exclusif — un seul
  enfant actif ; absent de notre corpus).
- `presetClass` = 6 valeurs : `entr | exit | emph | path | verb | mediacall`.
  (Dans le binaire PPCORE, l'enum interne est ordonné 0=entr 1=exit 2=emph 3=path 4=verb
  5=mediacall — voir `docs/binaire-ppcore.md` §3 ; ordre d'implémentation, pas celui du schéma.)
- `nodeType` : clickEffect | withEffect | afterEffect | mainSeq | interactiveSeq | clickPar |
  withGroup | afterGroup | tmRoot. Déclencheur→nodeType : on-click→clickEffect,
  with-previous→withEffect, after-previous→afterEffect.
- `restart` : always | whenNotActive | never. `fill` : remove | freeze | hold | transition.
- Sur `cBhvr` : `additive` (défaut `base`), `accumulate` (`none`), `xfrmType` (`pt`), `override`,
  `rctx` (PPT|IE). Sur `anim` : `calcmode` (discrete|lin|fmla), `valueType` (str|num|clr).
- Sources : datypic (miroir XSD ISO 29500) http://www.datypic.com/sc/ooxml/s-pml-animationInfo.xsd.html ,
  http://www.datypic.com/sc/ooxml/a-presetClass-1.html ,
  MS-OI29500 §19.5.22 cBhvr https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/b9a20a78-6401-4a8b-8ca9-b31ead409132

## ⚠️ Table presetID : PAS de table officielle publique
La spec déclare `presetID` comme un simple entier sans publier la correspondance vers les noms.
Plages communautaires (confiance moyenne) : 1–53 entrance/exit (le flag exit distingue), 54–82
emphasis, 83–85 media, 86–149 motion path (~149 effets). Amorce de tuples (ppt-master) :
appear=1, fly=2, blinds=3, box=4, checkerboard=5, circle=6, diamond=8, dissolve=9, fade=10,
peek=12, split=16, wipe=22, zoom=23, expand=50.
Source : https://github.com/hugohe3/ppt-master/blob/main/skills/ppt-master/scripts/docs/pptx-animations.md
(Notre table est vérifiée empiriquement : voir `docs/synthese-globale.md`.)

## Le langage de formules (le cœur pour rejouer)
Deux grammaires distinctes, à ne PAS mélanger :
1. **Valeurs `by/from/to`** des behaviors — grammaire « 2d_point » (MS-OI29500 §19.5.22) :
   nombre `"0.5 0.5"` ou formule **entre parenthèses** `"(#ppt_x-0.1, #ppt_w)"`.
2. **Attribut `fmla` de `<tav>`** (keyframes interpolés) — le vrai mini-langage :
   - opérateurs `+ - * / ^ % (mod)`, constantes `pi e`, comparaisons `== >= <= != !`,
     sélection `abs min max ?(if)`, trigo `sin cos tan asin acos atan`, log `ln`,
     étendu : `ceil floor exp sqrt rand deg rad sinh cosh tanh`.
   - `atan2` **n'existe pas** dans le langage d'animation (il existe en géométrie de forme via
     `at2`, système différent, EMU).
   - variables `#` (propriétés hôte, ex. `#ppt_x`), entrées `$`, cible = l'`attrName` courant.
   - calcMode limité à linear/discrete pour les formules.
   - Repère : fractions de diapo (0..1), x/y = centre de la forme.
   - Sources : C-Rex `tav` https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tav_topic_ID0EUO6HB.html ,
     MS-OE376 §4.6.79 https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/981b17ff-5594-42cf-ad8d-7cb39e653afa ,
     VBA `AnimationPoint.Formula` https://learn.microsoft.com/en-us/office/vba/api/powerpoint.animationpoint.formula
- ⚠️ `#ppt_c` (couleur) / `#ppt_r` (rotation) : convention connue, pas de source primaire les
  énumérant. `#ppt_x/y/w/h` sont confirmés par la spec.

## Réimplémentations open-source (références exécutables)
- **LibreOffice Impress** (la plus complète) :
  - Import : `oox/source/ppt/timenodelistcontext.cxx` (parse par/seq/cTn/behaviors),
    `animationspersist.cxx`, `timenode.cxx`. Miroir : https://github.com/LibreOffice/core
  - Exécution (moteur SMIL) : module `slideshow/`. **Parseur de formules** =
    `slideshow/source/engine/smilfunctionparser.cxx` (boost::spirit), éval via
    `AnimationFunction`. C'est l'équivalent open-source du langage `#ppt_*`.
  - Doc historique modèle SMIL : https://wiki.openoffice.org/wiki/Impress/API/Animations
- **ONLYOFFICE sdkjs** (JS complet) : classes `CAnim/CAnimClr/CAnimMotion/CAnimRot/CAnimScale/
  CSet`, conteneurs par/seq/excl. https://deepwiki.com/ONLYOFFICE/sdkjs/5.2-animation-and-transitions
- **Apache POI XSLF** : DOM typé du timing, pas de moteur ni de mapping preset.
- **python-pptx** : support animation quasi nul (pas d'API `<p:timing>`).

## Architecture binaire (public)
- `PPCORE.DLL` = cœur applicatif (modèle doc + moteur animation/diaporama). Cible RE correcte.
  https://strontic.github.io/xcyclopedia/library/PPCORE.DLL-5D114585672C22FC717F1E25700AD304.html
- `mso*.dll` = Office partagé (OfficeArt/DrawingML, UI). Le rendu des formes vient de là.
- ⚠️ `ocpptview.dll` : rôle non documenté publiquement. (Notre reverse le montre = moteur de
  **rendu** : il a les propriétés `ppt_*` mais aucun vocabulaire de séquence — `docs/binaire-ppcore.md` §1.)
- Rendu : pile Direct2D/DirectComposition/DirectWrite (générique Windows ; chemin exact de
  PPCORE non documenté). https://learn.microsoft.com/en-us/windows/win32/directcomp/architecture-and-components
- Origine des noms `ppt_*` : refonte « Custom Animation » de **PowerPoint 2002/XP**, modèle
  dérivé de **HTML+TIME** (extension SMIL de Microsoft/W3C) — d'où par/seq/excl et les formules.
  https://en.wikipedia.org/wiki/HTML%2BTIME , https://en.wikipedia.org/wiki/PowerPoint_animation
- Ressource RE historique (Shyam Pillai) : http://skp.mvps.org/ppttimeline1.htm ⚠️ cert HTTPS expiré.

## Symboles & instrumentation
- **PDB Office = NON publics** (msdl.microsoft.com ne sert que Windows). → RE à froid, pas de
  types ni noms internes. Nos noms de fonctions sont déduits.
  https://learn.microsoft.com/en-us/windows-hardware/drivers/debugger/microsoft-public-symbols
- **Frida** : pas de write-up PPCORE public, mais méthode standard OK (validée ici).
- **WinDbg + Time Travel Debugging** : approche recommandée pour le chemin d'exécution (trace
  rejouable, requêtes LINQ). https://learn.microsoft.com/en-us/windows-hardware/drivers/debuggercmds/time-travel-debugging-overview
