# Guide pour contribuer (humains et assistants IA)

Ce fichier dit **où est quoi**, **ce qu'il ne faut pas casser** et **comment prouver qu'on n'a rien
cassé**. Lis-le en entier avant de modifier le moteur. Les commentaires du code sont en français.

## 1. En une phrase

`renderer/` convertit un `.pptx` en SVG (une image vectorielle par diapo) **sans PowerPoint**, puis
deux viewers emballent ces SVG dans une page HTML autonome :

| Viewer | Commande | Sortie | Usage |
|---|---|---|---|
| **Défilement** | `node renderer/build.mjs deck.pptx` | `renderer/viewer.html` | toutes les diapos empilées, état final |
| **Présentation** | `node renderer/build-player.mjs deck.pptx` | `renderer/player-out.html` | une diapo à la fois, animations au clic |

Les deux viewers comptent autant l'un que l'autre : toute modification du moteur se vérifie dans les deux.

## 2. Carte du moteur PPTX

`renderer/pptx-render.mjs` n'est qu'une **façade** (API publique). Le code est dans `renderer/pptx/` :

```
pptx/xml.mjs            parseurs XML, cache, helpers d'arbre (tagOf/at/kids/first/deepFind), flatOf, esc
pptx/units.mjs          u() : EMU -> points
pptx/color.mjs          couleurs DrawingML -> CSS        pptx/preset-colors.mjs  noms prstClr
pptx/theme.mjs          loadContext() : thème (palette, polices, styles)
pptx/parts.mjs          .rels, images, placeholders hérités, slideOrder()
pptx/text-styles.mjs    puces, styles de texte du master / de la présentation
pptx/fonts.mjs          polices -> CSS, métriques (ligne de base, chasses)
pptx/symbols.mjs        Wingdings/Symbol -> Unicode       pptx/math.mjs  équations OMML
pptx/geometry.mjs       formes prédéfinies / custGeom -> chemins SVG
pptx/timing.mjs         animations (extractTiming) et transitions (extractTransition)
pptx/slide/render-slide.mjs   renderSlide : héritage diapo -> layout -> master -> thème, composition
pptx/slide/walk.mjs     parcours des formes (ordre z), groupes
pptx/slide/shape.mjs    formes   picture.mjs images   frame.mjs tableaux/graphiques/SmartArt
pptx/slide/text.mjs     texte (HTML dans <foreignObject>)
pptx/slide/paint.mjs    remplissages, dégradés, motifs, flèches   effects.mjs ombres/halos/filtres
```

Toutes les fonctions de `pptx/slide/` reçoivent `rc` (l'état du rendu de la diapo) en 1er argument ; ses
champs sont documentés en tête de `render-slide.mjs`. Le guide détaillé (unités, deux formes d'arbre
XML, pièges) est en tête de `renderer/pptx-render.mjs`. Vue d'ensemble : `docs/ARCHITECTURE.md`.

## 3. Règles à ne pas enfreindre

1. **Ne jamais muter** un arbre XML renvoyé par `cachedParse` ou `flatOf` : ils sont partagés entre
   diapos. Copier avant de modifier (`{ ...obj }`).
2. **Contrat SVG avec les viewers** : `data-spid` sur chaque forme (préfixe `m`/`l` pour master/layout),
   `data-par` sur chaque paragraphe, IDs de `<defs>` préfixés `s<N>_`. Les changer casse les animations.
3. Les constantes de mise en page (1,2 ; 0,75 L ; 1/8 pt ; 0,2 px…) sont **mesurées sur PowerPoint**.
   Ne pas les « arrondir » ou « simplifier ».
4. Un `.pptx` est une **entrée non fiable** : tout texte ou attribut qui sort dans le SVG passe par `esc()`.
5. Pas de commentaire à l'intérieur d'une chaîne ou d'un gabarit `` `…` `` : il partirait dans le SVG.
6. Les decks réels (`input/`, `specimen/`) ne se commitent **jamais**.

## 4. Protocole de vérification (obligatoire avant chaque commit)

```bash
npm test            # tests unitaires + intégration sur les decks d'exemple (doit rester vert)
npm run lint        # ESLint : variables non définies, imports manquants (0 erreur)
npm run diff        # quelles diapos ta modification change-t-elle, vs HEAD ?
npm run test:e2e    # les deux viewers dans Chrome (clics, navigation, animations)
```

- **Refactor** (réorganiser sans changer le comportement) : `npm run diff` doit afficher
  `0 modifiée(s)`. Sinon, ce n'est pas un refactor : trouve l'écart.
- **Correction / fonctionnalité** : `npm run diff` doit lister **exactement** les diapos visées.
  Regarde-les : `node tools/engine-diff.mjs --html renderer/_work/diff.html` produit une planche
  avant/après, à capturer avec `node tools/screenshot.mjs renderer/_work/diff.html renderer/_work/diff shotFull:all`.
- Ajoute un cas dans `test/fixtures/decks.mjs` (decks d'exemple générés par du code) et un test dans
  `test/fixtures.test.js` pour toute correction.

## 5. Decks d'exemple

`test/fixtures/decks.mjs` fabrique 4 petits decks en OOXML écrit à la main (`animations`, `shapes`,
`text`, `tables-charts`). `npm run fixtures` les écrit dans `test/fixtures/out/*.pptx` pour les ouvrir
dans les viewers. Ajouter une diapo = quelques lignes avec les helpers `sp()`, `para()`, `fx()`, `timing()`.

## 6. Outils

| Outil | Rôle |
|---|---|
| `tools/engine-diff.mjs` | compare le moteur courant à un commit (`--ref`), diapo par diapo |
| `tools/screenshot.mjs` | capture une page HTML dans Chrome (clics, touches, éléments) |
| `tools/make-fixtures.mjs` | écrit les decks d'exemple |
| `renderer/check.mjs`, `renderer/qa.mjs` | contrôles batch sur un dossier de decks réels |

## 7. Licence des contributions

Le code est sous **AGPL-3.0-or-later** (`LICENSE`). En proposant une contribution (commit, pull request,
patch), vous acceptez qu'elle soit publiée sous cette licence **et** que le mainteneur du projet puisse
aussi la distribuer sous une licence commerciale (double licence). Les polices de `renderer/fonts/`
gardent leur propre licence (`renderer/fonts/LICENCES.md`) : n'y ajouter que des polices libres, avec
leur texte de licence dans `renderer/fonts/licenses/`.
