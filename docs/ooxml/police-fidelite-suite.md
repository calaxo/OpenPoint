# Fidélité des polices — suite (2026-07-24, session autonome)

## Poppins (decks météo 050 : VENT, PRESSION…) = OK, RÉSOLU
Vérifié quantitativement sur d.5 (cartes Objectif) : contenu, retours à la ligne, LARGEUR (296 vs 297 px),
INTERLIGNE (même étendue verticale 376->582), graisse (400 corps / 700 gras), lettres = IDENTIQUES à
PowerPoint. Police byte-identique (cf. fonts-fidelity-research.md). Symptôme historique « pastille qui mord
la 1ʳᵉ lettre » déjà corrigé (héritage bodyPr). Résidu = plafond moteur (DirectWrite vs Blink), négligeable.
=> Poppins n'est PAS le problème.

## LE VRAI problème : polices NON embarquées (decks du corpus)
Le corpus utilise massivement des polices que je n'embarque PAS :
Arial (23197 occurrences !), Times New Roman (4781), Calibri, Century Schoolbook, Bernhard Modern Roman,
Comic Sans MS, News Gothic MT, Microsoft YaHei, Cambria Math, Wingdings/2/3, MS PGothic...
Je ne charge QUE Poppins (renderer/fonts/). Donc tout texte non-Poppins tombe en repli (sans-serif/serif
du navigateur) -> mauvaises MÉTRIQUES -> largeur/retour-à-la-ligne/positionnement faux = « texte mal affiché ».
- Sur Windows (mes rendus de test) : Arial/Times/Calibri sont installés -> OK par chance.
- Dans le LMS (Docker/Linux) : absents -> repli -> cassé.

## Plan
1. Embarquer des polices métrique-compatibles LIBRES (comme LibreOffice) :
   - Arial -> Liberation Sans (ou Arimo) ; Times New Roman -> Liberation Serif (Tinos) ;
     Calibri -> Carlito ; Cambria -> Caladea ; Courier -> Liberation Mono (Cousine).
   - Wingdings/Comic Sans/Century/Bernhard : plus rares, substituts approximatifs.
2. @font-face pour chacune + map typeface->family dans le moteur (fontCss).
3. Vérifier largeur/wrapping sur decks 040.

## Journal
- (init) Poppins validé OK ; identifié le vrai problème = polices non embarquées.

## FAIT (2026-07-24)
- Embarqué substituts métrique-compatibles LIBRES dans renderer/fonts/ : LiberationSans (Arial),
  LiberationSerif (Times), LiberationMono (Courier), Carlito (Calibri), Caladea (Cambria) — R/B/I/BI.
- fontFamily() : map Arial->'Arial','Liberation Sans',sans-serif ; idem Times/Courier/Calibri/Cambria
  (police réelle d'abord si installée, substitut embarqué ensuite, générique enfin).
- build-player.mjs : @font-face (url) pour ces familles. (sweep garde Poppins base64 : Windows a les vraies polices.)
- Fix bonus : font des PUCES passait par fontFamily SANS resolveFont -> '+mj-lt' littéral (repli). Corrigé.
- Portée : 46 diapos (surtout deck C) utilisent a:latin non-Poppins (Arial 1750, Times 335...).
  Rendu désormais correct MÊME sans la police installée (LMS Linux, ou polices rares Century/Bernhard/Comic Sans).
Note : Poppins (décks météo) était déjà OK ; ceci corrige les AUTRES polices.

## Bordures de tableau + contrôle final (2026-07-24)
- Bordures grille par défaut pour styles intégrés {5940675A...} (Table Grid) : d.46 VENT a sa grille.
- BALAYAGE DE CONTRÔLE (après data-spid/polices/bordures) : 0 régression. Corpus : 167 diapos >=8%,
  9 >15%, 1 >20% (PRECIP d.19 table), 0 >25%, moyenne 10.2% (= plancher de bruit). Stable.

## Revue « éléments manquants ? » (2026-07-24) — RIEN NE MANQUE
Inspection des diapos encore divergentes NON traitées (zones localisées = candidats à un élément absent) :
- sample d.26 (15.7%) : diagramme conv/div COMPLET (lobes, flèches, isobares, texte). Écart = retour à la
  ligne + forme de l'ellipse rouge. Rien ne manque.
- CIRCU d.6 (16.7%) : quiz COMPLET (question, 4 props, boutons, encart rouge). Écart = retour à la ligne. RAS.
- BASES d.43 (15.6%) : COMPLET (texte, 2 cartes+puces, tableau Facteurs à risque AVEC bordures per-cellule
  tcPr, encart vert). Écart = wrapping + nuances. RAS.
- NEURO d.59 (11.7%) : COMPLET (cartes Utricule/Saccule+puces, schéma cellules ciliées+flèches). Écart =
  texture raster du schéma. RAS.
CONCLUSION : aucun élément manquant dans le corpus. Les 10-16% résiduels = plafond moteur (retour à la
ligne Blink vs DirectWrite, texture raster des schémas, sous-pixel). Le contenu est COMPLET et fidèle.

## Fix couleur texte dans les encarts à formule OMML (d.15/21) — 2026-07-24
Le chemin OMML (paragraphes avec fractions) émettait color = couleur PAR DÉFAUT du paragraphe, en
ignorant la couleur PROPRE de chaque run. Sur l'encart rouge d.15, le run tx1 (clair) était rendu en
bg1 (sombre) -> texte quasi invisible sur fond rouge. FIX : résoudre la couleur du run (a:rPr>solidFill)
comme le chemin normal. Texte désormais blanc/lisible. (Débordement bas résiduel = hauteur de boîte,
non traité.) check.mjs 348+200, 0 err.

## Fix ligne via lnRef (carré d'angle droit d.26) — 2026-07-24
Le carré d'angle droit (bentConnector3, cxnSp) était invisible : sa ligne vient de <p:style><a:lnRef
idx=2 accent1> et NON d'un <a:ln>. Mon code ne résolvait le stroke que si <a:ln> existait. FIX : repli
sur lnRef (idx>=1) quand pas de <a:ln>, SCOPÉ aux formes non remplies (fill=none : connecteurs/lignes)
pour ne pas ajouter de bordure aux cartes pleines (0 régression). Largeur = thème lnStyleLst[idx-1],
couleur = lnRef. Vérifié : carré vert visible d.26 ; d.24/d.26 améliorées ; check.mjs 348+200 0 err ;
aucune nouvelle diapo en tête des divergences.

## Fix pointes de flèches TROP GROSSES (flèches bizarres VENT d.28) — 2026-07-24
Mon fix pointes (markerUnits=strokeWidth, markerWidth=8) donnait des pointes ~8× l'épaisseur du trait =
énormes (gros triangles pleins) vs PowerPoint ~3-3,5× (type="arrow" w/len="med"). lineMarker dimensionne
désormais selon @len/@w (sm=2/med=3.3/lg=4.8 × épaisseur). Pointes petites/proportionnées comme PowerPoint
(d.28, d.30, d.32...). check.mjs 348+200, 0 err.

## Investigation « Poppins décalé » (2026-07-24)
Méthode : overlay mien(vert)/réf(magenta) sur VENT d.5 (Objectif, 4 cartes).
CONSTAT : cartes 1-3 (texte 18pt UNIFORME, pas de lnSpc -> line-height 1.15 unitless) = ALIGNÉES (blanc).
Carte 4 « Reconnaître » (a un lnSpc explicite -> line-height fixe en px ; chemin noAutofit) = DÉRIVE
verticale croissante vers le bas (mien plus bas que réf) -> différence de PAS DE LIGNE cumulée.
=> Le décalage Poppins vient des boîtes à INTERLIGNE EXPLICITE (spcPts/spcPct) : mon `line-height` (px ou
facteur ×1.2) ne colle pas exactement à l'« interligne exact » de PowerPoint (baseline-to-baseline vs
boîte de ligne CSS + leading). Écart ~qq % par ligne, cumulatif sur les textes longs.
NON CORRIGÉ (volontaire) : changer le calcul de line-height impacte TOUT le texte = risque de régression
élevé sur un corpus déjà à 10% de moyenne. À traiter avec mesure fiable + validation large (prochaine passe).
Cartes/텍스tres sans lnSpc explicite (la majorité) sont déjà fidèles.

## À TRAITER (chantier dédié) : orientation flèches freeform + flip + animMotion (VENT d.14)
User : "la moitié des flèches dans le mauvais sens" sur d.14. Investigation : index 14 = 0 marqueur SVG
-> les pointes ne viennent PAS des line-markers récents. Ce sont des FREEFORM (custGeom) avec la pointe
DANS la géométrie, + shapes avec flipH/flipV (#9 "Forme libre 45" flipH=1), + balle verte (Ellipse 47)
à off(451,146)=haut-centre alors que PowerPoint la montre en bas-droite => elle bouge par animMotion.
=> Nœud complexe (custGeom + flip + animation). NON corrigé volontairement (risque régression sur flip/
rotation qui marche ailleurs). À traiter en passe dédiée : vérifier l'application flipH/flipV aux custGeom
+ l'orientation des pointes baked-in + l'état animMotion de la balle. Comparer chaque flèche mine/PowerPoint.

## Fix flipH/flipV des IMAGES (flèches courbes dans le mauvais sens, VENT d.14) — 2026-07-24
renderPic appliquait la rotation mais IGNORAIT flipH/flipV (renderSp le faisait, pas renderPic). Les
images flèches-courbes flip+rot (VENT d.14 "Graphique 26/31/39/41", avions ATMOS d.21...) étaient
tournées mais pas miroitées -> mauvais sens / fontaine asymétrique. FIX : ajout du miroir dans renderPic
(translate+scale, appliqué AVANT rot sur le contenu comme renderSp / PowerPoint).
PROTECTION (hash SVG des 576 diapos avant/après) : EXACTEMENT 10 diapos changées = les 10 à image flip
(ATMOS_STD 17/20/21/23/29/32, VENT 14/30/31, THERMO 58) ; 566 byte-identiques = 0 régression prouvée.
Vérifié améliorées : d.14 (courbes blanches symétriques), ATMOS d.21 (avions bon sens). check.mjs 348+200 0 err.

## Interligne : réduction normAutofit @lnSpcReduction appliquée AUSSI au lnSpc explicite — 2026-07-28
SYMPTÔME (question d.6 VENT + boîtes de cartes) : du texte gras s'enroulait sur MOINS de lignes que
PowerPoint (2 au lieu de 3) et paraissait sur-rétréci. MÉCANISME : la boîte a
<a:normAutofit lnSpcReduction="20000"/> ET le paragraphe a un <a:lnSpc><a:spcPct val="120000"/>.
Mon rendu du line-height du <div> = 120% x 1.2 = 1.44 SANS appliquer les 20% de réduction (le conteneur
lineH l'appliquait, mais le lnSpcCss du div l'ÉCRASAIT). PowerPoint = 120% x (1-0.20) x 1.2 = 1.152.
Conséquence : texte trop haut -> déborde la boîte -> l'autofit NAVIGATEUR sur-rétrécit la police ->
mauvais nb de lignes / texte minuscule (ex. boîte droite VENT d.35 illisible).
FIX (pptx-render.mjs ~L1067) : lnMul = 1 - lnRed appliqué au line-height explicite (spcPct ET spcPts).
lnRed=0 hors autofit -> AUCUN changement pour ces boîtes.
PROTECTION (hash SVG 576 diapos avant/après) : 48 diapos changées (celles à lnSpc explicite + réduction
autofit), 528 identiques. Vérifié en triptyque AVANT/APRÈS/PowerPoint : d.6 VENT (question 3 lignes OK),
TEMPERATURE d.21 (boîte milieu 8 lignes = PowerPoint), VENT d.35 (boîte droite plus sur-rétrécie),
SYST_NEURO d.27, SANTE_HYG d.20 -> tous améliorés ou neutres, 0 régression. check.mjs 348+200 0 err.
RESTE (hors périmètre, pré-existant) : gap entre paragraphes (spcBef/spcAft) parfois plus faible que
PowerPoint ; écart de largeur sous-pixel (~1 mot) sur certains gras.

## Espacement entre paragraphes : héritage de spcBef/spcAft (master bodyStyle) — 2026-07-28
SYMPTÔME : cartes/listes (SYST_NEURO d.27, TEMPERATURE d.21...) avec paragraphes/puces COLLÉS, alors que
PowerPoint laisse un gap. CAUSE : le gap vient du <p:bodyStyle> du MASTER (lvl1 spcBef=1000=10pt, lvl2/3=
500=5pt), jamais redéfini dans la diapo. ctx.txStyles et lvlStyle ne renvoyaient PAS spcBef/spcAft, et
spcMk ne lisait que le pPr -> aucun gap.
FIX : (1) capturer spcBef/spcAft dans ctx.txStyles ; (2) lvlStyle les résout inline->layout->master ;
(3) spcCss = pPr sinon def hérité. spcBef SUPPRIMÉ sur le 1er paragraphe (pi===0) : PowerPoint ne décale
pas le bloc en haut de boîte (l'espace-avant ne vaut qu'ENTRE paragraphes) -> sinon décalage vertical.
PAS de réduction lnSpcReduction sur spcBef/spcAft (elle ne touche que l'interligne).
PORTÉE : 470/576 diapos changées (quasi tout le corps de texte hérite du master) -> gros changement.
VALIDATION OBJECTIVE (quant.mjs diff% vs réf PowerPoint, avant/après) : SYST_NEURO d.27 8.49->7.71,
TEMPERATURE d.21 9.40->8.96, VENT d.35 6.25->6.09, VENT d.6 14.17->14.09, SANTE_HYG d.20 12.33=,
THERMO d.55 6.45=, ATMOS_ALT d.21 4.43->4.45 (bruit) -> TOUTES améliorées ou neutres, 0 dégradée.
Triptyques AVANT/APRÈS/PowerPoint : d.27/d.21/d.20 gaps désormais alignés sur PowerPoint. check.mjs
348+200 0 err. RESTE (pré-existant) : détection de puces sur lignes de continuation (SANTE_HYG d.20),
largeur sous-pixel des gras.

## LIMITATION connue : puces sur retours-ligne LITTÉRAUX (\n) dans <a:t> — 2026-07-28 (NON corrigé, volontaire)
SANTE_HYG d.20 (boîte verte) : un <a:r> contient du texte COLLÉ avec des LF (0x0A) internes
("Faites...\nÉvitez...\nNe pas brosser..."). PowerPoint traite chaque segment séparé par \n comme un
NOUVEAU paragraphe bulleté (une puce ▪ par segment). Mon rendu (white-space:pre-wrap) fait le saut de
ligne mais SANS répéter la puce -> 1 puce au lieu de 3. Corpus : 45 LF dans 12 fichiers (aucun VT/CR).
NON corrigé : le \n est au milieu de runs <span> ; découper proprement en paragraphes bulletés exige de
refactorer le rendu de paragraphe (fraîchement stabilisé sur interligne+espacement) pour un gain purement
cosmétique sur du texte collé -> ratio risque/valeur défavorable, mandat zéro-régression. À reprendre si
demandé, en découpant AU NIVEAU DES RUNS (répartir le texte sur \n en sous-lignes portant la même pPr,
même puce, sans spcBef entre elles).

## Polices SYMBOLE (Wingdings/Webdings/Symbol) -> Unicode : anti-tofu cross-plateforme — 2026-07-28
SYMPTÔME (trouvé en testant les vieux decks « deck A », CHAP.20 d.5) : puces et symboles en
police Wingdings/Symbol s'affichaient en TOFU (□) sur Chromium headless — et donc sur le LMS Linux (ces
polices y sont absentes). Ex. « ⇐ Centre rédacteur » (flèche double gauche liant deux colonnes) rendu □.
CAUSES : (1) fontFamily rend `'Wingdings'`/`'Symbol'` SANS repli -> glyphe absent = tofu ; (2) la police
symbole est souvent déclarée dans <a:sym typeface="Symbol"> (pas <a:latin>) -> non lue ; (3) chars stockés
en zone privée PUA 0xF000+octet.
FIX : mapper les codepoints vers leur équivalent UNICODE, rendus en police NORMALE (Arial) -> lisibles
partout. Deux tables : WINGDINGS_MAP (dingbats : ●■□▪✔➢… inconnu -> « • ») et SYMBOL_MAP (Adobe Symbol :
ASCII->grec αβγδ…ΩΔΣ + flèches ↔←↑→↓ ⇔⇐⇑⇒⇓ + math ≤≥≠≈±×÷°√∞∂∫ ; inconnu -> on GARDE le char, jamais
deviner à tort). Appliqué aux PUCES (buChar+buFont) ET aux RUNS de texte (via <a:sym> prioritaire sur
<a:latin>), PUA-aware. « Symbol » (grec/math) distingué de Wingdings (décoratif) pour ne pas transformer
un λ en puce.
PROTECTION (hash SVG 703 diapos, corpus + decks dépliés) : 125 changées (celles à puces/runs
symbole). Quant diff% vs PowerPoint avant/après : SYST_NEURO d.13 9.83->9.68, d.20 10.61->10.58,
SANTE_HYG d.17 7.02->6.84 (améliorées) ; NUAGES d.36 9.06->9.23 (le ≤ « 1000 m ≤ visibilité » s'affiche
enfin = amélioration, pixels ≠ glyphe PowerPoint). 0 tofu restant, 0 texte mutilé (VENT d.21 formule
ρ/Ω intacte, grec préservé). check.mjs 348+200 0 err.
NB : les 22 vieux decks V2 « deck A » (627 diapos) rendent SANS erreur ; robustesse confirmée.
Seule diapo « maigre » = CHAP.01 d.13 (vidéo mp4, rendue via son poster -> normal).

## Clipping du texte : overflow:hidden -> visible (bas des caractères rogné) — 2026-07-28
SYMPTÔME (repéré par l'utilisateur, pastilles Sommaire VENT d.2) : le BAS des chiffres 0-4 était coupé,
alors que 5-7 non. CAUSE : les pastilles 0-4 sont des placeholders -> chemin de rendu NORMAL avec
overflow:hidden ; la ligne (chiffre 24px, line-height 1.15 = 27.6px) est plus haute que la boîte réduite
par les insets (foreignObject h=25.7px) -> le glyphe centré déborde de ~1px en haut/bas et se fait ROGNER.
Les pastilles 5-7 (bodyPr avec <a:noAutofit> -> chemin overflow:visible) n'étaient pas rognées.
FIX : chemin normal passe en overflow:visible (foreignObject + div). PowerPoint ne clippe JAMAIS le texte
d'une forme au bord de la boîte ; l'autofit navigateur (data-af=norm) rétrécit déjà les boîtes normAutofit
trop pleines, donc pas de débordement parasite.
PROTECTION : overflow est dans le markup -> les 824 diapos changent (inutile en hash). Quant diff% vs
PowerPoint avant/après sur 8 diapos variées (Sommaire, dense, formules, autres decks) : IDENTIQUE
(VENT d.2 3.38->3.39 = les pixels de chiffres récupérés ; tout le reste inchangé) -> 0 débordement
parasite, 0 régression. check.mjs 348+200 0 err.

## Autofit navigateur : ne pas rétrécir en hauteur une boîte plus courte qu'une ligne — 2026-07-28
SYMPTÔME (repéré par l'utilisateur, pastilles Objectifs VENT d.5) : les chiffres 1-2-3 rendus à 34.3px
(=40×0.95³) alors que le 4 (et le vrai PowerPoint) sont à 40px. CAUSE : 1-2-3 sont des placeholders ->
chemin de rendu NORMAL, boîte réduite par les GROS insets du layout ; le chiffre 40pt (1 ligne) dépasse
en hauteur -> mon autofit navigateur (cap 6 sur boîtes fixed) le RÉTRÉCISSAIT de 3 crans. Le 4 est une
forme libre <a:noAutofit> -> boîte pleine -> pas rétréci. PowerPoint ne rétrécit JAMAIS une boîte
non-autofit ; le texte déborde simplement.
FIX (autofit dans build.mjs + build-player.mjs + sweep.mjs) : les boîtes non-autofit (fixed) ne
rétrécissent en HAUTEUR que si elles peuvent contenir au moins une ligne (canH = clientHeight >= 1 ligne) ;
si la boîte est plus COURTE qu'une ligne (insets qui la mangent -> pastille) on ne rétrécit pas, le texte
déborde (overflow:visible déjà en place). Largeur : inchangé (compense le wrapping). oneLine lue sur la
police du <div data-par> interne (PAS du conteneur .af qui est à 16px par défaut -> piège).
PROTECTION : autofit = script runtime (pas dans le SVG). Quant diff% vs PowerPoint ANCIEN vs FINAL sur 15
diapos variées : d.5 8.2->8.15 (chiffres corrigés), les 14 autres IDENTIQUES (dont sample d.26 13.9 et
THERMO d.55 6.45 qui régressaient avec une règle trop brutale « fixed=largeur seule » -> l'affinage
« plus courte qu'une ligne » les préserve). 0 régression. Chiffres d.5 mesurés : 1/2/3/4 tous à 40px.
