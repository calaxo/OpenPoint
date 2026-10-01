# Journal de session — reprise instantanée

## 2026-07-24 (suite) — Formules OMML + pied de page coupé
- `21e175e` **rendu math OMML** : `renderMath()` (fraction m:f empilée, parenthèses m:d, exposant/indice
  m:sSup/sSub, racine m:rad, n-aire m:nary), intercalé avec le texte via le parse ORDONNÉ (txBodyO).
  **NFKC** pour convertir les Mathematical Alphanumeric (𝜌𝑓𝒅𝑷 → ρ f d P, absents de Poppins).
  Piège : nœud `#text` a `tagOf='#text'` (≠ undefined) → tester `#text` en premier. VENT d.21 formules OK.
- `92253db` **interligne** : `line-height:1.15` unitless se calculait contre 16px (défaut du `.af`) au lieu
  de la taille réelle → ligne trop haute → pied de page 9pt COUPÉ dans sa boîte. Fix : `font-size` sur le
  `<div>` paragraphe = taille du 1er run. Corrige la coupe ; cartes d.5 OK.
- Reste (à faire ensuite, demandé par l'utilisateur) : pastille qui mord la 1ʳᵉ lettre + agencement Poppins.

## 2026-07-24 — Fixes fidélité (option A retenue), ancrés sur oracle PowerPoint
- `769060c` **MCE `mc:AlternateContent`** (GROS fix générique) : `walk()`/`walkChild()` descend
  maintenant dans la branche `mc:Choice` (sinon `Fallback`). Avant, toute forme enveloppée
  (boîtes de FORMULES OMML sous `mc:Choice Requires="a14"`, présent 1-5×/deck dans TOUT le corpus)
  était INVISIBLE → « il manque carrément du texte ». VENT d.21 : 2 boîtes de formules réapparues.
  **Reste : le math OMML lui-même** (`<a14:m><m:oMath>` fractions Vg=1/ρf·dP/dl) non rendu → « Vg = »
  suivi d'un blanc. TODO : rendu des fractions (num/den empilés) ; un linéaire serait illisible.
- `bfd8850` **fillRef idx=0 = aucun remplissage** : forme `arc` ouverte rendue en trait seul
  (le 45° de VENT d.10 était un camembert bleu plein). Générique.
- `6969c28` **ajout `upArrow`** : ABSENT du XML ECMA (presetShapeDefinitions.xml a upDownArrow en
  DOUBLE + upArrowCallout mais PAS upArrow — bug de la spec !). Les flèches vers le haut tombaient
  en rectangle (barre) : VENT d.20 vecteurs PGF, d.14 AIR CHAUD. Défini en miroir de downArrow.
- **Constat important** : la plupart des écarts signalés par l'utilisateur (3ᵉ image cassée d.5,
  herbe noire d.10/14) étaient des **captures PRÉ-fix** — sur rendu FRAIS ils sont corrects.
- **Reste (métriques Poppins, non corrigé — risque de régression)** : pastille de carte qui
  chevauche la 1ʳᵉ lettre (d.5 « Définir » → « efinir »). Géométrie : texte de carte à x=751076,
  cercle jusqu'à x=1077044 → le « D » est sous le cercle ; PowerPoint le fait tenir de justesse
  (police plus étroite). Plafond FOSS connu. NE PAS hacker sans validation.
- Méthode : chaque fix comparé à `scratchpad/ref/slide-N.png` (export COM PowerPoint) + rendu maison
  `scratchpad/rs.mjs`, puis `node renderer/check.mjs` (548 diapos 0 erreur).

## 2026-07-23 (soir) — RÉCAP pour demain (utilisateur parti, PC peut s'éteindre)

### Fait aujourd'hui (tout committé, `OpenPoint`)
Fixes rendu moteur maison, validés contre **référence PowerPoint pixel-perfect** (nouvel oracle :
`scratchpad/ppt-export.ps1` exporte chaque diapo via COM ; helper de rendu maison `scratchpad/rs.mjs`) :
- `5e55810` rotation de GROUPE (rot/flip) → manches à air inclinées.
- `1cffe6c` `color-interpolation-filters="sRGB"` → images assombries plus écrasées en noir (fond nuageux).
- `7652dd1` champ `slidenum` = vrai n° (fini « ‹N°› ») + autofit navigateur seulement pertinent.
- `871d475` régression Sommaire corrigée : autofit toutes boîtes cap 20(normAutofit)/6(autres) +
  coupe des espaces de fin. **⚠ ces réglages autofit sont délicats — re-vérifier Sommaire+pied de page
  ensemble avant de retoucher.**
- bldP, a14:imgEffect, animMotion/animRot, srcRect, defaultTextStyle, etc. (voir entrées plus bas).

### Reverse Direct2D (demande explicite) — CONCLU : cul-de-sac pour le texte
`docs/reverse-d2d.md` + `binary/capture_d2d.py` (frida). Prouvé : PowerPoint rend le texte en
contours→FillGeometry ou atlas→DrawImage, **jamais DrawGlyphRun** → pas de display-list texte
exploitable. **Seul pixel-perfect = laisser PowerPoint rendre (COM `Slide.Export`/`ExportAsFixedFormat`).**

### Plan demain (par priorité)
1. **Décider l'architecture** avec l'utilisateur : (a) moteur maison JS + oracle (portable + animé, ~LibreOffice
   sur le statique) OU (b) COM Export pixel-perfect + couche anim maison par-dessus. C'est LE choix.
2. Si (a) : balayage oracle d'un deck complet (exporter toutes les diapos PowerPoint + rendre les miennes,
   comparer, punch-list). Corriger les écarts GÉNÉRIQUES seulement (jamais de hack par diapo).
3. Écarts restants signalés (à confirmer sur rendu FRAIS, beaucoup étaient des captures pré-fix) :
   couleurs de certaines images, arrondis ponctuels. Le vert des barbules d.12 = `srgbClr 51B980`
   EXPLICITE dans le fichier → PAS un bug.
4. Ne PAS re-tenter le reverse D2D pour le texte (cul-de-sac établi).

---


But : reprendre le travail sans rien redemander, même après coupure brutale (crédits/compact).
Règle : entrée AJOUTÉE EN TÊTE à chaque reprise + à chaque cible finie. La vérité détaillée reste
dans `coverage.md` (matrice) ; ici = le fil « où j'en suis / prochaine étape ».

Sauvegarde : **commit git après chaque cible**. Note mémoire à chaque jalon.

---

## 2026-07-22 — ✅ FIDÉLITÉ DIAPO TITRE (retour utilisateur) — commit `07b1dec`

Diagnostic sur VENT s1 (comparaison vraie diapo vs rendu). 3 bugs corrigés :
1. **Titre trop gras** : le layout titre a `defRPr b="0"` qui doit battre le master `b="1"`, mais mon
   `id?.b==='1' || mas.b` tombait sur le master → gras à tort. Fix = 1er niveau qui DÉFINIT `@b` gagne.
2. **Puces « • » parasites + texte pied de page invisible** : les zones de texte LIBRES (sans `p:ph` :
   « sample », « pied de page », n° de diapo) héritaient du `bodyStyle` du masque (puces + couleur). Fix =
   `phKind='none'` pour les zones libres (pas de style body) + couleur de dernier recours = `tx1` du
   thème (clair) au lieu de `#000`.
3. **« Placement bizarre »** : en grande partie mon HARNAIS de test qui n'embarquait pas Poppins →
   repli plus large → débordement. Avec Poppins embarqué, placement fidèle. (Piège noté dans concepts.md.)

Validé visuel : VENT/TEMPERATURE/NEURO s1 (titres légers, pastilles, pieds de page propres) +
PRECIPITATIONS d.5 (cartes blanches, texte foncé NON régressé) + TEMPERATURE d.28 (puces de contenu
légitimes conservées). 548 diapos 0 erreur. Reste mineur : « TURBULENCE vu à » (texte à grand espace
interne passant sous un badge).

---

## 2026-07-22 — ✅ CIBLES 7-8-9 FINIES : bldP + corrections photo + rejeu mouvement

- **7. bldP** (commit `4287f55`) : `renderText` émet `data-par` par paragraphe ; `extractTiming` lit
  `spTgt>txEl>pRg` ; le player cache/révèle au niveau paragraphe. Validé sample d.26 (build progressif).
- **8. a14:imgEffect** (commit `8d51280`) : `deepFindFlat` sur le blip → `brightnessContrast` (feComponentTransfer
  linéaire lum/contraste) + `saturation` (saturate). Validé VENT d.10 (12 images corrigées).
- **9. Rejeu mouvement** (commit `c4b3001`) : `parseMotion` (animMotion path → translation, animRot → rotation) ;
  player = `animateTransform` SMIL **additive** (se compose avec le transform de base, pas de clobber) ;
  rotation centrée via `getBBox`. Validé CIRCU_GEN d.10 (translation) + VENT d.44 (spin 360°, anémomètre).

Reste ❌ (mineur) : `animScale` (scale centré pas exprimable en 1 transform additif), `animClr`, `anim`
keyframes fines ppt_*, `colorTemperature`/`sharpenSoften`, reflection/innerShdw, déclencheurs interactifs.
Corpus 348+200 diapos 0 erreur à chaque commit.

---

## 2026-07-22 — ✅ CIBLE 6 FINIE : animation cible-groupe + data-spid images/tableaux

Committée `187e9f2`. `renderGroup` enveloppe ses membres dans `<g data-spid=groupId>` (le groupe
s'anime en bloc, comme PowerPoint) ; `renderPic`/`renderTable` émettent enfin `data-spid`.
**Résultat mesuré : 99,3 % des cibles d'animation résolues** (560/564 sur 3 decks ; avant ~90 % —
images/groupes orphelins). Les 4 restants = **bldP** (build par paragraphe : l'anim cible une plage
de paragraphes, pas la forme entière) → prochaine cible anim si besoin. Player rebuild OK (34 animées).
Wrapper `<g>` inerte au rendu statique (0 régression, 548 diapos 0 erreur).

### Reste (aucun bloquant), par intérêt décroissant
- **bldP** (build par paragraphe) : découper le texte en `<div data-spid-par>` révélés un par un.
- **anim keyframes** (`anim` ppt_x/y/w/h + fmla), `animMotion` (path), `animScale/Rot/Clr` : rejeu fin
  du mouvement (le player fait fade/wipe/zoom CSS ; le mouvement précis n'est pas rejoué).
- **a14:imgEffect** (brightnessContrast/saturation/colorTemperature) : corrections photo MS 2010.
- reflection/innerShdw (rares), détail au trait `fill="none"` des presets (quasi absent du corpus).

---

## 2026-07-22 — ✅ CIBLES 4 & 5 FINIES : recadrage image + defaultTextStyle

**Cible 4 (srcRect)** committée `926d173` : recadrage source `a:srcRect` (l/t/r/b millièmes %) —
image agrandie (/vw,/vh) + décalée + clippée. **Découverte : AUCUN `c:chart` dans le corpus** (les
10 graphicFrames = tableaux déjà rendus) → charts abandonné. srcRect au contraire massif (~125).
Validé visuel d.5 PRECIPITATIONS (4 cartes « Objectifs » recadrées correctement).

**Cible 5 (defaultTextStyle)** committée `565d351` : chargé par niveau depuis presentation.xml,
ajouté en **repli final** de `lvlStyle` (sz/color/font) — comble le texte hors placeholder au lieu de
`#000`/`#fff` en dur. `clrMapOvr` vérifié NON utilisé dans le corpus (pas la cause du bug couleur).

**Passe de non-régression visuelle** : 7 diapos sur 4 decks (TEMPERATURE d.28/34, CIRCU_GEN d.16,
BASES_PHY d.67, PRECIPITATIONS d.5, ALTIMETRIE d.19) → toutes fidèles, 0 régression. Corpus 548
diapos 0 erreur à chaque commit.

### État & prochaines cibles possibles (aucune bloquante)
Fidélité statique très forte. Restent, par ordre d'intérêt :
- **Animation avancée** (pour le lecteur) : `anim` keyframes (mouvement/échelle), `animMotion` (path),
  `bldP` (build par paragraphe), cible = **groupe** (propager aux membres). Réf pml.xsd + moteur-animation.md.
- **a14:imgEffect** (brightnessContrast/saturation/colorTemperature) : corrections photo MS 2010
  (extension, dans blip>extLst>AlternateContent). Impact moyen.
- **reflection / innerShdw** : rares (1 reflection dans le corpus).
- Détail au trait des presets `fill="none"` (glyphes actionButtons) — quasi absent du corpus.

---

## 2026-07-22 — ✅ CIBLES 2 & 3 FINIES : texte avancé + effets photo

**Cible 2 (texte)** committée `6a406a6` : rPr `cap`(all/small)/`baseline`(exposant-indice)/
`spc`(espacement)/`strike`(barré) + pPr `lnSpc`/`spcBef`/`spcAft` (unités spec : spcPct millièmes %,
spcPts centièmes pt). Validé visuel d.16 CIRCU_GEN (« 40ᵉᵐᵉ » en exposant), titres MAJUSCULES.

**Cible 3 (effets)** : `effectFilter` (ombre+glow+softEdge → filtre SVG mergé) sur formes ET images ;
`imageFilter` (grayscl/duotone/alphaModFix) sur les blips. Validé visuel d.67 BASES_PHY (glow bords).
Reste ❌ : reflection (rare), innerShdw, a14:imgEffect (corrections MS 2010), a:clrChange.

Corpus 548 diapos, 0 erreur à chaque étape. Prochaine cible (4) au choix : **charts** (`c:chart`,
dml-chart.xsd — 10 graphiques non rendus) OU **animation avancée** (anim keyframes/animMotion/
bldP/cible groupe, pml.xsd + moteur-animation.md) OU **résolution rPr defaultTextStyle/clrMapOvr**.

---

## 2026-07-22 — ✅ CIBLE 1 FINIE : moteur de géométrie générique (187 presets + custGeom)

Fait, testé (615 diapos 0 erreur), validé visuellement, **committé**.
- `renderer/gen-preset-shapes.mjs` → `renderer/preset-shapes.json` (187 formes bakées, COMMITÉ).
  Piège : racine XML mal orthographiée `presetShapeDefiniton​s` (coquille ECMA).
- `renderer/preset-geom.mjs` : évaluateur des 17 opérateurs de formules + guides intégrés +
  builder de path (arcTo → arc SVG centré sur position courante). Exporte `presetGeomD`, `evalGeom`.
- `pptx-render.mjs` : `presetPath` délègue ; `custGeomPath` réécrit (nœud ordonné + formules).
  Sites d'appel (renderSp, renderPic) passent le custGeom ORDONNÉ via `deepFind`.
- Sémantique confirmée dans la spec (ISO/ECMA §20.1.9.11 opérateurs, §20.1.9.4 arcTo,
  §20.1.10.56 guides). `at2/cat2/sat2` via `Math.atan2` (comme LibreOffice).

Prochaine cible (2) → **résolution rPr / héritage texte** : ordre exact de résolution des
propriétés de run manquantes (voir open-questions « Texte / héritage »), croiser dml-main.xsd
+ §21.1 lstStyle + §19 txStyles + MS-OI29500. Puis lnSpc/spcBef/spcAft, kern/spc/baseline/cap.

---

## 2026-07-22 — Docs reçues, démarrage ingestion

Docs déposées par l'utilisateur dans `docs/ooxml/specs/` :
- `ECMA-376-1..4` (extraits + zips) — Part 1 = Fundamentals + Markup Reference (le gros PDF).
- `ECMA-376-1/.../OfficeOpenXML-DrawingMLGeometries.zip` → extrait dans `_extract/geom/` :
  **`presetShapeDefinitions.xml`** (les ~180 formes preset AVEC leurs formules de guides) +
  `presetTextWarpDefinitions.xml`.
- `.../OfficeOpenXML-XMLSchema-Strict.zip` → extrait dans `_extract/xsd/` : **`dml-main.xsd`**
  (DrawingML), **`pml.xsd`** (PresentationML : timing/anim), sml/wml, shared-*.
- `iso et microsfdt/` : `ISO_IEC_29500-1_2016(en).pdf` + **`[MS-OI29500].pdf`** (écarts MS).

Chemins extraits (NON commités, droits ISO) :
- Formes : `docs/ooxml/specs/ECMA-376-1_5th_edition_december_2016/_extract/geom/presetShapeDefinitions.xml`
- XSD   : `docs/ooxml/specs/ECMA-376-1_5th_edition_december_2016/_extract/xsd/`

### Plan d'attaque (par impact visuel, depuis coverage.md)
1. **presetGeom + custGeom guide-formulas** — `presetShapeDefinitions.xml` donne TOUT :
   générer un évaluateur de formules (`*/ +- +/ ?: abs sqrt sin cos tan at2 pin mod min max val`)
   + convertir chaque preset en path. Débloque flèches courbes/uturn/callouts/arcTo d'un coup.
2. **Résolution rPr / héritage texte** — croiser dml-main.xsd + §21.1 + MS-OI29500.
3. **Effets photo** (glow/reflection/softEdge/duotone) — §20.1.8.
4. **Charts** (`c:chart`) — dml-chart.xsd.
5. **Animation avancée** (anim keyframes/animMotion/bldP/groupe) — pml.xsd + moteur-animation.md.

### Prochaine étape (EN COURS)
→ Lire `presetShapeDefinitions.xml` (structure d'un `<...>` : avLst/gdLst/cxnLst/rect/pathLst),
  puis écrire un module générique `renderer/preset-geom.mjs` (évaluateur de guides + builder de path)
  et le brancher dans `pptx-render.mjs` (`presetPath`/`custGeomPath`). Vérifier `node renderer/check.mjs`.
- `a8378f9` **héritage bodyPr du layout** : carte-placeholder avec `<a:bodyPr/>` vide reprend tIns/lIns/anchor du layout (tIns=288000 réserve la place de la pastille). Corrige « pastille mord la 1ʳᵉ lettre » (Définir→efinir) sur les 4 cartes d.5. NON lié à Poppins (intuition user correcte).

## Balayage automatique de TOUT le corpus (comparateur oracle) — 2026-07-24
Objectif user : « un algo qui compare deux images à chaque fois » + « fait tous les fichiers, le max ».

### Outillage (commité)
- `renderer/compare.mjs` — diff pixelmatch mine↔réf, grille 16×9 pour localiser les écarts structurels.
- `renderer/sweep.mjs` — pour un deck : unzip → rendu headless (puppeteer, Poppins base64 @font-face +
  autofit) → screenshot 1280×720 → `diffOf()` vs réf PowerPoint (export COM) → rapport trié.
  Dumpe `mine`/`diff` PNG par diapo si diffPct≥10 (revue visuelle).
- `renderer/sweep-all.mjs` — balaye les 13 decks, agrège `_GLOBAL.json`, sort le TOP 40 du corpus.
- Réfs oracle : export COM PowerPoint de TOUTES les diapos (≈615) via `scratchpad/export-all-refs.ps1`.

### Résultat : les plus gros écarts sont DOMINÉS par les images photo (plafond moteur, pas des bugs)
Le TOP du corpus = diapos avec grandes photos (cockpit d'intro ×12 decks, ciel/nuages, anatomie NEURO).
Le texte/formes/couleurs sont FIDÈLES. Deux causes structurelles génériques identifiées :

1. **Photo d'intro « Introduction » (d.4, présente sur ~12 decks) trop sombre** (luma mesurée cercle
   MINE 11.6 → 17.4 après fix, RÉF 83.6). CAUSE RACINE = **effet artistique Office** :
   `<a:blipFill><a:blip><a:grayscl/> + a14:imgProps>imgLayer r:embed(.wdp)>imgEffect>a14:brightnessContrast
   bright="-40000"</...> <a:srcRect r="-683" b="-54251"/> <a:stretch><a:fillRect/>`.
   → PowerPoint rend la **HD photo `.wdp` (JPEG-XR)** pré-traitée (crop+grayscale+expo baked-in), PAS le
   JPEG source. Le `.wdp` n'est pas décodable en navigateur → **plafond du pipeline propriétaire Office**
   pour les images à effet artistique. (Ne PAS appliquer le `brightnessContrast -40%` : assombrirait
   encore ; PowerPoint compense via le calque `.wdp`.)
   FIX partiel commité : **srcRect NÉGATIF = ajustement de remplissage, PAS un recadrage** → on l'ignore
   et on garde le cover (`slice`) au lieu d'écraser l'image dans un coin (améliore le cadrage, luma +6).

2. **Logos coin haut-droit (logos partenaires, pics du master) masqués** par un rectangle pleine-diapo
   dessiné APRÈS eux (ordre de calques master→layout→slide). `<image>` présent+positionné+valide dans le
   SVG (rendu OK en isolation) mais recouvert. NON corrigé : la sémantique exacte du z-order bg/formes
   inter-niveaux n'est pas confirmée (risque de régression sur toutes les diapos). CANDIDAT à traiter avec
   un oracle dédié.

### Non-régression : `node renderer/check.mjs` → deck A = 348 diapos / deck B = 200 diapos, **0 erreur**.

## Fix LOGOS (ordre de calques du fond) — 2026-07-24
Bug : le `<p:bg>` du LAYOUT (bleu) était peint APRÈS les formes du MASQUE → recouvrait les logos
master (logos partenaires, `Image 20/22`). Cause : `renderSlide` dessinait un `bgRect()` à CHAQUE
niveau (masque, layout, diapo), en interleaving avec les formes. Or le fond n'est PAS cumulatif.
Fix : **un seul fond effectif** (précédence diapo > layout > masque) dessiné DERRIÈRE toutes les formes ;
puis formes masque → layout → diapo. `curRels` pointé sur le bon niveau (un fond peut être une image).
Vérifié : logos VISIBLES sur d.4 ; check.mjs deck A=348 / deck B=200, 0 erreur.

## Fix LUMINOSITÉ photo (corrections a14 déjà cuites) — 2026-07-24
La photo d'intro d.4 rendait quasi-noire (luma 17 vs réf 83). CAUSE : `imageFilter` appliquait un
`a14:brightnessContrast bright="-40000"` → `intercept=-0.4` = soustraction de 102/canal → tout <102 en
noir. Or ce `brightnessContrast` est niché sous `a:blip>extLst>a14:imgProps>a14:imgLayer r:embed(.wdp)` :
il décrit l'édition Office appliquée à la HD photo ORIGINALE (`.wdp`) pour PRODUIRE l'image de base rId2
(déjà éditée, celle qu'on rend). Donc DÉJÀ cuit → ne pas ré-appliquer.
FIX : si `a14:imgLayer` présent (`baked`), ignorer les corrections a14 (brightnessContrast/saturation).
RÉSULTAT : luma 17 → 96,8 (réf 83,6) ; d.4 30,76% → 22,33% ; sortie du top des divergences.
(Preuve : `.wdp` décodé via .NET WIC = photo couleur HD lumineuse identique à rId2 → rId2 est bien la
version d'affichage déjà éditée.) check.mjs deck A=348 / deck B=200, 0 erreur.

## Fix CADRAGE photo (srcRect négatif) — 2026-07-24 (corrige un fix erroné antérieur)
Le fix "srcRect négatif ignoré" (commit b946668) était une MAUVAISE correction : je pensais que le
srcRect négatif causait la noirceur, mais c'était le brightnessContrast (corrigé à part). Or le srcRect
`b=-54251` sert à faire tenir la photo PAYSAGE (1232×805) dans la forme CARRÉE (929²) SANS déformation :
il rend la région source ~carrée (letterbox bas), l'image occupe le haut ~65 % de la boîte, non déformée.
L'ignorer forçait un zoom cover (trop zoomé). REVERT au mapping d'origine (négatif compris) → cadrage
fidèle. RÉSULTAT combiné (logos + luminosité + cadrage) : **d.4 diff 30,76% → 1,18 %** (quasi pixel-
parfait). check.mjs deck A=348, 0 erreur.

## Fix formes à paths fill/stroke distincts (arc 45°) — 2026-07-24
Bug : le 45° (prst="arc", fillRef=0 donc fill=none) s'affichait en « camembert » (secteur avec 2 rayons)
au lieu d'un arc fin. Cause : `evalGeom` ne renvoyait qu'UNE silhouette (paths remplissables) et
`renderSp` la traçait entièrement. Or le preset `arc` a 2 paths ECMA : path1 = secteur (fill=norm,
stroke=false), path2 = courbe (fill=none, stroke=true). Fusionner puis tracer dessine les rayons.
FIX générique : `evalGeomParts`/`presetGeomParts`/`custGeomParts` renvoient { fill, stroke } séparés
(en honorant a:path @fill/@stroke) ; `renderSp` trace le path de tracé à part quand il diffère du
remplissage (sinon comportement inchangé = cas courant, pas de régression). check.mjs 348+200, 0 erreur.

## Fix texte noAutofit rogné (boutons de réponse d.6) — 2026-07-24
Les pastilles de réponse (police 40pt, boîte 39,6pt, `<a:noAutofit/>`) affichaient le texte rogné en
haut. Cause : le foreignObject faisait la taille RÉDUITE par les insets (tIns+bIns) + overflow:hidden →
clip. Or noAutofit = « ne pas ajuster » : PowerPoint NE clippe PAS, le texte déborde librement.
FIX : pour `<a:noAutofit/>`, foreignObject = boîte PLEINE, insets en padding, overflow:visible (centrage
sur la boîte entière -> les capitales tiennent). normAutofit/spAutoFit/défaut inchangés. check.mjs 348+200, 0 err.

## Fix THÈME par-master (fond orange au lieu de beige) — 2026-07-24 (fort impact)
Les diapos du deck C (via slideMaster2) avaient un fond ORANGE au lieu de BEIGE.
Cause : `loadContext` chargeait `theme1.xml` EN DUR. Or un deck multi-master a plusieurs thèmes :
master1→theme1 (lt2=C78243 orange), master2→theme2 (lt2=D9CAB7 beige). Rectangle de fond = schemeClr
tx2 → clrMap tx2=lt2 → theme1 donnait l'orange au lieu du beige de theme2.
FIX : `renderSlide` résout la chaîne diapo→layout→master→rels→theme AVANT `loadContext`, qui prend
désormais le thème du master de la diapo. theme1 (VENT/PRESSION) inchangé. check.mjs 348+200, 0 erreur.

## Fix images EMF/WMF manquantes + a:lum — 2026-07-24
THERMO d.35 : grand graphique e-T ABSENT (fond noir). Cause : « Picture 2 » = image EMF (image38.emf),
format vecteur Windows non affichable en <image> -> imageHref renvoyait null.
FIX : préprocessing `convert-metafiles.ps1` (.NET/WIC, x2) convertit chaque EMF/WMF en « <fichier>.png » ;
`imageHref` sert ce PNG. Intégré à sweep.mjs + build-player.mjs (helper `metafiles.mjs`, via unzip ;
Windows only, no-op ailleurs). + Support `<a:lum bright/contrast>` (§20.1.8.42) dans imageFilter (le
graphe a lum -20%/-20% -> teinte grise). RÉSULTAT : graphe rendu net + teinte correcte. check.mjs 348+200, 0 err.
(Corpus : seulement 5 EMF/WMF au total, mais impact visuel fort sur les diapos concernées.)

## Fix <a:grpFill/> (régions colorées manquantes) — 2026-07-24
THERMO d.55 : régions orange/bleue « Instabilité/Stabilité absolue » absentes (seule la verte, à
solidFill explicite, s'affichait). Cause : ces membres de groupe ont `<a:grpFill/>` (remplissage hérité
du GROUPE) — le fill (accent3/bg2) est sur le grpSpPr. Mon moteur ne gérait pas grpFill -> fill=none.
FIX : renderGroup résout le fill du grpSpPr et le propage (ctx.grpFill, restauré après, gère l'imbrication) ;
fillFromSpPrFlat renvoie ctx.grpFill pour `<a:grpFill/>`. check.mjs 348+200, 0 err.

## Fix texte des tableaux trop gras/gros — 2026-07-24
Tables (SANTE d.38, PRECIP d.19…) : texte tout en gras et trop grand -> débordait/rognait les cellules.
Cause : renderText des cellules appelé avec phKind='body' -> héritait le style « corps » du master
(gras/gros sur ces decks). FIX : phKind='none' (défaut neutre + rPr propre de chaque run, ex. b="1" sur
la 1ʳᵉ colonne seulement). check.mjs 348+200, 0 err.

## Fix pointes de flèches des lignes/connecteurs (headEnd/tailEnd) — 2026-07-24
7 diapos VENT (d.22,23,28,30,31,32,45) avaient des flèches SANS pointe : le moteur ignorait
`<a:ln><a:headEnd/><a:tailEnd type="arrow"/>`. FIX : marqueurs SVG partagés par type (triangle/arrow/
stealth/diamond/oval), couleur = trait via `context-stroke`, `orient="auto-start-reverse"` (même marker
début/fin), `markerUnits="strokeWidth"`. Appliqués au strokeAttr. check.mjs 348+200, 0 err.

## Revue slide-par-slide VENT (48 diapos) — 2026-07-24
Demande user : vérifier chaque diapo VENT. Montages mien|PowerPoint des 48. Résultat : ~40/48 fidèles.
Écarts trouvés + statut :
- **Pointes de flèches (d.22,23,28,30,31,32,45)** -> CORRIGÉ (commit 9a67495, marqueurs SVG).
- d.27 labels « à 1000 m/500/200 » : présents (blanc, bien placés) mais + petits que PowerPoint =
  **mise à l'échelle de police par le GROUPE** (PowerPoint scale les fonts d'un groupe par ext/chExt ;
  mon moteur scale positions/tailles mais PAS la police). NON corrigé (risque de régression large).
- d.15 / d.21 : lignes de formule (math OMML dans encarts) incomplètes/rognées. NON corrigé (math).
- d.46 : bordures de tableau absentes = style de tableau intégré « Table Grid » {5940675A-...} non stocké
  dans le deck (bordures via tableStyle, pas via tcPr). NON corrigé.
- d.47 : « écoulement laminaire » = WordArt/texte à effet (gros orange) rendu en texte plat. NON corrigé.
LIMITES CONNUES restantes (par ordre d'impact) : 1) police non scalée par groupe ; 2) bordures de
tableStyle ; 3) math OMML dans encarts ; 4) WordArt/text effects.
