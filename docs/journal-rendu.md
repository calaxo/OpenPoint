# Journal du moteur de rendu (Office → Web)

> Chronologie des travaux sur le **rendu** (distinct de `ooxml/progress-log.md` qui suit le reverse
> d'animation). Voir aussi [`ARCHITECTURE.md`](ARCHITECTURE.md) et [`police-fidelite-suite.md`](police-fidelite-suite.md).

## 2026-07 — Fidélité PPTX (retours utilisateur, chaque correctif protégé + committé)

Méthode : oracle = PowerPoint COM ; protection = hash SVG des ~700 diapos avant/après + triptyques
+ `check.mjs` (doit rester 0 erreur). Tag `pptx-stable` = point de retour.

- **Flèches d.14** : `renderPic` ignorait flipH/flipV → images flip+rot sans miroir. Fix miroir.
- **Interligne question d.6** : `lnSpcReduction` de normAutofit non appliqué au `lnSpc` explicite →
  texte trop haut → sur-rétréci par l'autofit navigateur. Fix `line-height × (1 − lnRed)`.
- **Espacement paragraphes** : héritage `spcBef/spcAft` du master (bodyStyle) manquant.
- **Polices symbole (Wingdings/Symbol)** → tofu hors Windows/LMS. Mappage codepoints → Unicode
  (dingbats + grec/flèches/math), lecture de `<a:sym>`, gestion PUA 0xF000.
- **Clipping** : `overflow:hidden` rognait le bas des caractères → `overflow:visible`.
- **Autofit** : ne plus rétrécir en hauteur une boîte plus courte qu'une ligne (chiffres de pastilles).
- **Collision d'IDs SVG entre diapos** (viewer multi-diapos) : IDs `clip/grad/mk_/fx/img` locaux à chaque
  diapo → collision quand N diapos dans une page → mauvais clip (image circulaire en petit rectangle dans
  le viewer, OK dans le player). Fix : préfixe `s<slideNo>_`. (commit b8c9444)

Corpus validé sans erreur : `deck A` (348) · `deck B` (200) · `deck C` (519, source différente).

## 2026-07 — Nouveau moteur DOCX (`docx-build.mjs`, autonome)

- PoC WordprocessingML → HTML en flux : paragraphes/runs (gras/italique/couleur/taille/police/exposant),
  styles + héritage `basedOn`, tableaux, images (data URI), thème.
- **Suivi de modifications** : `w:del`/`w:delText` (texte supprimé) rendu barré.
- **En-têtes/pieds** : `header/footer`, champs `PAGE`/`NUMPAGES` → « Page X sur Y » réel (paginateur).
- **mc:AlternateContent**, zones de texte (`w:txbxContent`), images PUA/.wdp ignorées.
- **Pagination CÔTE CLIENT** : mesure des hauteurs → pages A4 uniformes (fin des « énormes pages »).
- **Fidélité texte** : interligne (`w:spacing line/lineRule`), espace avant/après + autospacing,
  indentation (left/right/firstLine/hanging), tailles par défaut, **titres** (héritage du style de
  paragraphe), **puces réelles** (Wingdings→Unicode), **TOC** (points de suite + n°),
  **tableaux** (largeurs de colonnes `tblGrid`, bordures).
- QA : `docx-check.mjs` (10 docs, 0 crash).
- **Limite dure** (rendu en flux) : collage de page de garde flottant + recoloration d'image →
  `office-to-pdf.mjs` (LibreOffice) pour le pixel-perfect.

## 2026-07 — Couverture de tags + performance (mesurée, protégée par hash)

Objectif : rendre le moteur prêt pour le plus de PowerPoint différents possible, SANS casser, en
optimisant le temps. Outil `coverage.mjs` : compte les tags XML lus par le moteur → **99,4 % compris**
(RENDU 76,8 % · ANIMATION 5,3 % · MÉTADONNÉES 17,3 % · trous 0,6 %).

- **Puces colorées** (`a:buClr`) + taille (`a:buSzPct`) : la puce a sa couleur/taille PROPRE (souvent une
  couleur d'accent ≠ texte), plus toujours la couleur du 1er run. No-op strict si absents.
- **Surlignage de texte** (`a:highlight`) → fond coloré derrière le run (comme Word). Additif.
- **Jointures/bouts de trait** (`a:round`/`a:bevel` → `stroke-linejoin` ; `@cap` → `stroke-linecap`) :
  on n'émet QUE les valeurs non-défaut SVG → corrige les pics « miter » sur les lignes de courant des
  schémas (ex. CHAP.07 VENT d.40), sans bruit ailleurs.
- **Cache de parse** (thème/master/layout/rels/presentation) keyé chemin+mtime+taille : `renderSlide`
  reparse ~64 Ko/diapo alors que N diapos partagent master/layout/thème. **−30 % deck seul, −37 % batch.**
  Transparence PROUVÉE : hash SVG identique sur 864 diapos avant/après (`102c8c115664f302`).
- **Trous restants (0,6 %, différés)** : `a:effectRef`/`a:fontRef` (styles de thème — risque sur le chemin
  cœur du texte, corpus déjà fidèle via héritage placeholder), `a:fillRect`, OMML (`m:*`), bordures
  diagonales de cellule, `asvg:svgBlip`. Tous à faible impact visuel.

Corpus validé sans erreur à CHAQUE changement : `deck A` (627) · `deck B` (200) · `deck C` (519).

## 2026-07-29 (nuit) — Poussée compatibilité MAX (session autonome, chaque pas protégé)

Mandat : « augmente au max la reconnaissance de tags pour que le max de PowerPoint soit compatible,
peu importe forme/couleur/police ; travaille en solo, sauvegarde souvent (le PC peut s'éteindre) ».
Filet de sécurité : nouvel outil **`qa.mjs`** (hash SVG par diapo des 1346 diapos de référence +
diff vs baseline) → chaque modif affiche EXACTEMENT les diapos changées + 0 erreur obligatoire.

- **`a:fontRef`** (style de forme) : couleur + police de thème par défaut du texte d'une forme (souvent
  la seule source de couleur d'un bloc SmartArt/callout externe). Repli en fin de chaîne. 0 diapo corpus.
- **`a:effectRef`** : ombre/lueur de thème (effectStyleLst ajouté au ctx) quand pas d'effectLst explicite.
  2 diapos (Decision Making #13/#22, ombres subtiles), vérifié.
- **`a:pattFill`** : remplissage à motif → tuile `<pattern>` SVG (points/lignes/diagonales/grille par
  famille de preset). Rendait 'none' avant. 2 diapos (CHAP.13 précipitations RA/SN enfin visibles).
- **Contour de texte** (`rPr>a:ln`) → `-webkit-text-stroke`. 9 diapos (labels A/D sur cartes satellite
  enfin lisibles). Additif pur.
- **GRAPHIQUES** (`c:chart`) : nouveau module **`chart-render.mjs`** (colonnes/barres groupées+empilées,
  lignes, aires, camembert/anneau ; axes+grille+labels ; **titre `c:title` + légende `c:legend`** droite/bas ;
  couleurs séries/points sinon palette thème). Lit les caches c:numCache/c:strCache (pas besoin d'Excel).
  Branché dans renderTable (dispatch tableau/graphique via r:id). 0 diapo du corpus (aucun graphique),
  vérifié bout-en-bout sur mini-deck (groupé/empilé/camembert, titres, légendes).
- **SmartArt** (`dgm:relIds`) : rendu via le dessin de secours **`dsp:drawing`** (ppt/diagrams/drawingN.xml).
  Astuce clé : ce dessin est un spTree structurellement identique à PresentationML mais préfixé `dsp:` ->
  on remplace `dsp:`->`p:` et on réutilise le walk normal, avec un mapping chOff/chExt -> cadre. Vérifié
  sur un SmartArt process synthétique (3 boîtes positionnées/centrées). 0 diapo corpus.
- **Graphiques — détails** : titre + légende (droite/bas) + étiquettes (c:dLbls showVal/showPercent) +
  nuage de points/bulles (c:scatterChart) + c:varyColors (barres colorées par catégorie) + bornes d'axe
  explicites (c:valAx>c:scaling>c:min/max). Tous vérifiés sur mini-charts.
- **Puces-image** (`a:buBlip`) : petite icône en puce (résolue via curRels). Additif.
- **Image de secours OLE** : un graphicFrame inconnu (objet incorporé Excel/PDF, contrôle) rend son
  p:pic de secours (deepFind -> renderPic) au lieu d'un vide. Additif.
- **Mini-decks synthétiques** : pour tester les features ABSENTES du corpus (graphiques, SmartArt), écrire
  un ppt/ minimal (thème+master+layout+slide + part chart/drawing) et appeler renderSlide -> vérif visuelle.
  C'est la méthode pour valider une feature invérifiable autrement, au lieu de l'ajouter à l'aveugle.

Bilan nuit : ~15 commits, chacun `qa.mjs` = 1346 diapos 0 erreur + hash contrôlé. Le moteur couvre
désormais toutes les grandes familles de contenu PowerPoint (formes/fills/texte/effets/tableaux/images/
fonds + GRAPHIQUES + SmartArt + OLE). Restent volontairement non faits (gros/risqués/invérifiables) :
WordArt courbé (prstTxWarp), 3D (bevel/rotation), bordures diagonales de cellule, reflection, radar.
- **REJETÉ — repli police `a:ea`/`a:cs`** : tenté (latin absent → ea/cs) mais 759 diapos changées =
  RÉGRESSION (réordonnait la priorité d'héritage ; +mn-ea → minorFont écrasait l'Arial hérité).
  Reverté. LEÇON : ne JAMAIS réordonner la priorité d'héritage ; un repli n'est sûr qu'en TOUT DERNIER.

## 2026-07 — Aiguilleur + voie fidèle + documentation

- `render.mjs` : un seul point d'entrée (`.pptx`/`.docx`).
- `office-to-pdf.mjs` : rendu **fidèle** via LibreOffice (docx + pptx), pour les mises en page complexes.
- **Docs** : `ARCHITECTURE.md` + 4 schémas SVG (`diagrams/`), `renderer/README.md`, bannières de rôle,
  renvoi depuis le README racine. Objectif : reprise facile (humain ou machine).
