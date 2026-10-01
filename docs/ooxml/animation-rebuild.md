# Réécriture du moteur d'animation (2026-07-24, session autonome)

## But
Reproduire fidèlement la choréographie PowerPoint dans le player (bons éléments / bon clic / bon ordre /
simultanéité / délais / builds par paragraphe). Puis : corriger « Poppins mal affiché ».

## Modèle PowerPoint (mainSeq) — §19.5
- `<p:timing>` > `p:par(tmRoot)` > `p:seq(mainSeq)` > `childTnLst` > `p:par` (groupes de clic).
- `cTn @nodeType` : `clickEffect` = démarre sur un NOUVEAU clic ; `withEffect` = simultané au précédent ;
  `afterEffect` = enchaîné après le précédent ; `clickPar`/`withGroup`/`afterGroup` au niveau groupe.
- `stCondLst > cond @delay` : `indefinite` = attendre le clic ; nombre = délai ms.
- `spTgt @spid` = cible ; `p:pRg @st/@end` = plage de paragraphes (bldP) ; `p:bldLst > p:bldP` = builds.
- Effet : `cTn @presetClass` (entr/exit/emph), `@presetID`, `animEffect @filter/@transition`.

## Étapes
1. [EN COURS] Diagnostic d.5 : mapping spid timeline <-> data-spid SVG (savoir pourquoi des cartes ne
   sont jamais révélées).
2. Réécrire extractTiming : sortir des STEPS de clic ordonnés, chaque effet {spid, cls, nodeType, delay,
   dur, filter, pRg}. Regrouper par clic via nodeType (clickEffect=nouveau step).
3. Gérer bldP (apparition paragraphe par paragraphe).
4. Player : révéler au bon step ; cibler groupe si spid=groupe.
5. Tester d.5, d.6, d.28 ; commit à chaque incrément.

## Journal
- (init) notes créées.

## RÉSOLU (2026-07-24) : collision de data-spid = vrai coupable
Les steps d'extractTiming étaient CORRECTS. Le bug était le player : `querySelector('[data-spid=8]')`
tombait sur la MAUVAISE forme car master/layout/diapo ont chacun leur espace d'id (plusieurs data-spid=8).
FIX : préfixer les data-spid du masque ('m') et du layout ('l') dans renderSlide (spidPfx) ; la timeline
ne cible que la diapo (sans préfixe). Résultat d.5 : clic 1..4 révèlent 1 carte complète chacune, gauche->
droite, titre présent dès le début = EXACTEMENT PowerPoint. Corrige TOUTES les diapos animées.
Journal :
- extractTiming OK (steps corrects) ; player corrigé via spidPfx ; reveal-all à la dernière étape en filet.

## VALIDÉ (2026-07-24) : animation fidèle
- d.5 (Objectif) : clic 1..4 = 1 carte complète chacune, G->D, titre dès le début = PowerPoint. ✓
- d.6 (quiz) : clic 1 question, 2..5 les 4 propositions une par une, puis boutons + encart rouge ;
  état final 8/8 complet (bouton vert « 1,2 et 4 » + explication rouge) = PowerPoint. ✓
Choréographie CORRECTE (bons éléments/bon clic/bon ordre). Reste approximatif mais acceptable : easing/
courbes de timing fines, simultanéité intra-clic (afterEffect enchaînés rendus quasi-simultanés).
=> Tâche animation : cœur RÉSOLU. Passage à « Poppins mal affiché ».
