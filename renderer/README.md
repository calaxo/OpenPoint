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
| `build.mjs` | `viewer.html` | PPTX en **défilement** (toutes les diapos, état final) |
| `build-player.mjs` | `player-out.html` | PPTX en **présentation animée** (clic = étape suivante) |
| `office-to-pdf.mjs` | `*.pdf` | Rendu **fidèle** via LibreOffice (docx **et** pptx) |

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

| Fichier | Rôle |
|---|---|
| `preset-geom.mjs` | Évaluateur de géométrie DrawingML (187 formes + guides/formules). |
| `preset-shapes.json` | Formes prédéfinies « bakées » (généré par `gen-preset-shapes.mjs`). |
| `metafiles.mjs` + `convert-metafiles.ps1` | EMF/WMF → PNG (.NET/WIC, Windows). |
| `gen-preset-shapes.mjs` | Outil ponctuel : régénère `preset-shapes.json` depuis la spec. |
| `fonts/` | Polices embarquées (Poppins + substituts libres métrique-compatibles). |

## Artefacts (générés, non versionnés — cf. `.gitignore`)

`viewer.html` · `player-out.html` · `docx-out.html` · `*.pdf` · `_work/` (décompressions).

## Rappels

- **Ne pas casser le PPTX** : `pptx-render.mjs` est le cœur. Après modif → `check.mjs` sur
  `specimen/deck A` (348), `specimen/deck B` (200), `specimen/deck C…` (519) = **0 erreur**.
- Tag git `pptx-stable` = point de retour connu-bon.
- Les deux moteurs sont **indépendants** : DOCX ne peut pas casser PPTX (et inversement).
