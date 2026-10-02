# OpenPoint — PowerPoint (.pptx) → Web, en JavaScript pur

Convertit une présentation `.pptx` en pages Web autonomes **sans PowerPoint ni LibreOffice** : chaque
diapo devient un SVG fidèle (formes, texte, images, tableaux, graphiques, thèmes), et les animations
et transitions sont rejouées. Un moteur Word (`.docx` → HTML) est aussi inclus.

## Démarrage rapide

```bash
npm install
node renderer/build.mjs deck.pptx          # viewer en défilement  -> renderer/viewer.html
node renderer/build-player.mjs deck.pptx   # mode présentation     -> renderer/player-out.html
```

Chaque sortie est **un seul fichier HTML** qui s'ouvre sans serveur (images dédupliquées, polices
embarquées). Pas de deck sous la main ? `npm run fixtures` génère des decks d'exemple dans
`test/fixtures/out/`.

- **Viewer en défilement** : toutes les diapos empilées, barre d'outils, saut direct (`#12`), clavier
  (j/k, flèches), diapos masquées signalées, chargement progressif.
- **Mode présentation** : une diapo à la fois, animations au clic (fondu, balayage, zoom, venir de,
  trajectoire, rotation, agrandissement, par paragraphe…), étapes « après le précédent » automatiques,
  transitions de diapo, clavier / tactile / plein écran, lien direct `#3`, contrôles qui s'effacent.

## Contribuer

**Lire [AGENTS.md](AGENTS.md)** (humains comme assistants IA) : carte du code, règles à ne pas enfreindre
et protocole de vérification (`npm test`, `npm run lint`, `npm run diff`, `npm run test:e2e`).
Vue d'ensemble illustrée : [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) · carte de `renderer/` :
[renderer/README.md](renderer/README.md).

## Origine : le reverse du moteur d'animation PowerPoint

Étude née d'un besoin concret : impossible de sortir proprement des decks PowerPoint
riches en animations. Avant de retenter quoi que ce soit, ce repo répond à la question préalable :
**comment fonctionne le moteur d'animation de PowerPoint, exactement ?**

## Base de connaissances OOXML (plan de travail)

Pour comprendre le format à fond et savoir quoi implémenter ensuite :
**[docs/ooxml/](docs/ooxml/)** — carte de la spec, **matrice de couverture** (élément → statut →
code), concepts/pièges, questions ouvertes. Les docs techniques se déposent dans
`docs/ooxml/specs/` (local, non commité).

Historique du moteur de rendu : [docs/lecteur-pptx.md](docs/lecteur-pptx.md).

---

**Le reverse du moteur d'animation — commence par la [synthèse globale](docs/synthese-globale.md).** Quatre niveaux :
- **[docs/moteur-animation.md](docs/moteur-animation.md)** — le *format* `<p:timing>` et le modèle COM (boîte noire).
- **[docs/binaire-ppcore.md](docs/binaire-ppcore.md)** — le *binaire* : quel module implémente le
  moteur (`PPCORE.DLL`), désassemblage x86-64 (tables d'enums, compilateur de presets, sérialiseurs).
- **[docs/dynamique-runtime.md](docs/dynamique-runtime.md)** — le moteur *en train de tourner* :
  hooks frida sur PowerPoint vivant, compilateur preset→behaviors capturé en direct.
- **[notes/recherche-web.md](notes/recherche-web.md)** — spec ECMA-376, langage de formules, réimplémentations OSS (sourcé).

## TL;DR

- Les animations d'une diapo sont un arbre **SMIL** dans `ppt/slides/slideN.xml` (`<p:timing>`) :
  séquence principale → un nœud gelé (`delay="indefinite"`) par clic → effets → *behaviors*.
- Les « effets » (Fondu, Balayage, …) ne sont pas des primitives : ce sont des recettes compilées
  en **8 behaviors** (`set`, `animEffect`, `anim` à keyframes, `animScale`, …). Implémenter les
  behaviors suffit à tout rejouer — 4 ont couvert 100 % du deck étudié.
- « Avec/Après le précédent » n'existe pas à l'exécution : tout est pré-calculé en **délais
  absolus depuis le clic**.
- Validation : la vue XML et la vue COM de PowerPoint ont été alignées **280/280 effets, 0 écart**,
  puis la timeline a été **ré-exécutée** dans un lecteur HTML maison.

## Contenu

| Chemin | Rôle |
|---|---|
| `docs/moteur-animation.md` | Reverse du format `<p:timing>` + modèle COM |
| `docs/binaire-ppcore.md` | Reverse du binaire PPCORE.DLL (désassemblage) |
| `docs/dynamique-runtime.md` | Reverse dynamique (frida sur PowerPoint vivant) |
| `docs/synthese-globale.md` | Synthèse des 4 niveaux + table presetID vérifiée |
| `notes/recherche-web.md` | Recherche documentaire sourcée (spec, OSS) |
| `binary/scan-strings.mjs` `dump-region.mjs` | Localiser le vocabulaire / lire les tables de données |
| `binary/xref.py` `disasm.py` `pe-info.py` | Xrefs + désassemblage x86-64 (capstone/pefile) |
| `binary/frida_probe.py` `dyn_hooks.py` `dyn_profile.py` | Instrumentation runtime (frida + COM) |
| `specimen/` | Deck(s) d'étude local — **non versionné** (cf. `.gitignore`) |
| `tools/dump-com.ps1` | Vue PowerPoint du moteur via COM → `specimen/com-dump.json` |
| `tools/parse-timing.mjs` | Parseur `<p:timing>` → `specimen/timeline.json` |
| `tools/crosscheck.mjs` | Validation croisée XML↔COM + table preset↔effet |
| `tools/export-poc.mjs` | Géométrie + timeline → `poc/poc-data.js` |
| `poc/player.html` | Lecteur : rejoue les animations au clic (ouvrir en direct dans un navigateur) |

## Usage

```bash
npm install
node tools/parse-timing.mjs 8          # timeline JSON de la diapo 8
node tools/crosscheck.mjs              # re-valider le mapping
node tools/export-poc.mjs && start poc/player.html
```

`tools/dump-com.ps1` nécessite PowerPoint installé (Windows).
