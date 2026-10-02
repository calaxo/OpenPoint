# `renderer/` — carte du dossier

> Moteur de rendu Office → Web. **Vue d'ensemble illustrée : [`../docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md).**
> Ici : le rôle d'un coup d'œil de chaque fichier.

## Les 2 moteurs (le cœur)

| Fichier | Rôle |
|---|---|
| **`pptx-render.mjs`** | Cœur **PPTX** → SVG. `renderSlide()`, `loadContext()`, `extractTiming()`. Le plus gros fichier ; toute modif se re-teste avec `check.mjs` **et `qa.mjs`**. |
| **`chart-render.mjs`** | Rend un **graphique** (`c:chart`) → SVG (colonnes/barres/lignes/aires/camembert/nuage + titre/légende/étiquettes). Appelé par `pptx-render.mjs` (dispatch `renderTable`). |
| **`docx-build.mjs`** | Cœur **DOCX** → HTML. `renderDocx()`. **Autonome** : ne dépend pas du PPTX. |

## Les sorties (quoi lancer)

| Fichier | Produit | Mode |
|---|---|---|
| `render.mjs` | — | **Aiguilleur** : `.pptx`→build.mjs, `.docx`→docx-build.mjs |
| `build.mjs` | `viewer.html` | PPTX en **défilement** (toutes les diapos, état final, masquées grisées) |
| `build-player.mjs` | `player-out.html` | PPTX en **présentation animée** (clic = étape suivante, diapos masquées sautées) |
| `office-to-pdf.mjs` | `*.pdf` | Rendu **fidèle** via LibreOffice (docx **et** pptx) |

## Les viewers (enveloppe HTML autour du moteur)

`build.mjs` et `build-player.mjs` ne font qu'assembler des pièces partagées ; le résultat reste **un seul
fichier HTML autonome** (ouvrable en `file://`, sans serveur).

| Fichier | Rôle |
|---|---|
| `lib/deck.mjs` | `openDeck(pptx)` : décompression **sûre** (anti zip-slip), EMF/WMF → PNG (Windows), ordre de présentation (`slides`), diapos masquées (`hidden`). Utilisé aussi par `check` / `qa` / `sweep`. |
| `lib/deck-data.mjs` | `renderDeckData(deck)` : rend chaque diapo (+ timing / transition pour le player) → objet JSON de la page. |
| `lib/images.mjs` | Dé-duplication : chaque `data:` URI est stockée **une fois** ; le SVG porte un jeton `op-img:N`. |
| `lib/html-page.mjs` | `buildPage()` : inline les fichiers de `viewer/` + les données (`<script type="application/json" id="op-data">`). |
| `viewer/common.js` | Runtime commun (global `OP`) : lecture des données, jetons → URL `blob:`, `#ancre`. |
| `viewer/scroll-viewer.js` + `viewer.css` | Viewer **défilement** : insertion paresseuse des SVG, barre d'outils, j/k, `#3`, diapos masquées grisées. |
| `viewer/player.js` + `player.css` | **Présentation** : étapes de clic, effets (Web Animations), transitions, clavier/tactile, `#3`. Le modèle est expliqué en tête du fichier. |

Les fichiers de `viewer/` sont des **scripts navigateur classiques** (pas de module, pas de bundler), lintés
par `npm run lint`. Test de bout en bout dans Chrome : `npm run test:e2e` (`tools/e2e-viewers.mjs`).

## Qualité / oracle

| Fichier | Rôle |
|---|---|
| `check.mjs` | QA PPTX : rend chaque diapo d'un dossier, compte erreurs + diapos « maigres ». **Doit rester 0 erreur.** |
| `qa.mjs` | QA non-régression : hash SVG par diapo des 3 packs de référence (1346 diapos) + **diff vs baseline** = blast radius exact d'une modif. `--save` fige la baseline. Filet du mandat zéro-régression. |
| `coverage.mjs` | Mesure quels tags XML le moteur LIT (RENDU/ANIMATION/MÉTA/AUTRE en %). |
| `docx-check.mjs` | QA DOCX : rend chaque `.docx`, crash + métriques (texte/img/tbl). |
| `sweep.mjs` / `sweep-all.mjs` | Diff **pixel** vs référence PowerPoint (oracle COM), rapport trié. |
| `compare.mjs` | Diff de deux images, localise les écarts (grille 16×9). |
| `scan.mjs` | Inventaire des fonctionnalités DrawingML utilisées par un corpus. |

## Support (utilisé par le moteur PPTX)

> Ces fichiers proviennent de l'intégration hôte (plateforme d'apprentissage) et sont réutilisés ici tels quels.

| Fichier | Rôle |
|---|---|
| `animFonts.js` | Polices web embarquées (data: URI), métriques « win » PowerPoint (`fonts/metrics.json`) et chasses au 1/8 pt (`fonts/advances.json`). Lu par le moteur **et** par les outils. |
| `web-fonts.mjs` | Pour les outils : le CSS des polices du LMS, restreint aux familles citées par les diapos. |
| `pptxNotes.mjs` | Extrait les notes du présentateur d'un `.pptx`. |
| `preset-geom.mjs` | Évaluateur de géométrie DrawingML (187 formes + guides/formules). |
| `preset-shapes.json` | Formes prédéfinies « bakées » (généré par `gen-preset-shapes.mjs`). |
| `metafiles.mjs` + `convert-metafiles.ps1` | EMF/WMF → PNG (.NET/WIC, Windows). |
| `gen-preset-shapes.mjs` | Outil ponctuel : régénère `preset-shapes.json` depuis la spec. |
| `fonts/` | Polices embarquées : `.woff2` (18 familles + substituts libres) + `metrics.json` / `advances.json` / `LICENCES.md` / `build-web-fonts.py` ; anciens `.ttf` conservés. |

## Artefacts (générés, non versionnés — cf. `.gitignore`)

`viewer.html` · `player-out.html` · `docx-out.html` · `*.pdf` · `_work/` (décompressions).

## Rappels

- **Couche hôte** (file de rendu, cache, livraison, iframe sandbox, runtime d'animation, front) : hors de ce dépôt — ce dépôt produit les pages Web que la plateforme d'apprentissage intègre.
- **Plus d'autofit dans le navigateur** : `build`/`build-player`/`sweep` affichent le SVG tel quel, comme en sortie finale.
- **Ne pas casser le PPTX** : `pptx-render.mjs` est le cœur. Après modif → `check.mjs` sur un corpus de
  référence local = **0 erreur**.
- Tag git `pptx-stable` = point de retour connu-bon.
- Les deux moteurs sont **indépendants** : DOCX ne peut pas casser PPTX (et inversement).
