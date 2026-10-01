# Base de connaissances OOXML — travail préparatoire

Objectif : comprendre à fond le format `.pptx` (OOXML / ISO-IEC 29500 / ECMA-376) pour rendre les
diapos **exactement** comme PowerPoint, sans PowerPoint. Cette base est le plan de travail : elle
survit aux compactions de contexte (ce sont des fichiers) et sert de carte pour croiser la **doc**
avec **notre code** (`renderer/pptx-render.mjs`) et avec **le binaire PowerPoint** (`PPCORE.DLL`,
voir `docs/binaire-ppcore.md`).

## Fichiers de cette base

- **[spec-map.md](spec-map.md)** — carte de la spec ECMA-376 : quelle partie / quel §
  couvre quoi. Pour situer vite un excerpt de doc.
- **[coverage.md](coverage.md)** — **LE document central** : matrice élément OOXML → statut dans
  notre moteur (✅/🟡/❌) + fonction du code + § de spec + notes. C'est là qu'on voit ce qui
  reste à faire et où.
- **[concepts.md](concepts.md)** — les pièges/concepts transverses déjà appris (EMU, clrMap,
  phClr, héritage placeholder, autofit, modèle de timing, extensions MC…).
- **[open-questions.md](open-questions.md)** — ce que je ne comprends pas encore et que les docs
  à venir devront éclaircir. À cocher au fur et à mesure.
- **`specs/`** — déposer ici les docs techniques (PDF/XML/XSD/texte). **Non commité** (droits :
  ISO 29500 payant ; ECMA-376 libre). Voir workflow ci-dessous.

## Workflow quand de nouveaux docs arrivent (après le compact)

1. **Déposer** les fichiers dans `docs/ooxml/specs/` (ils restent locaux, gitignorés).
2. **Lire par cible** : ne pas lire les 5000 pages linéairement. Partir de `coverage.md`, prendre
   une ligne 🟡/❌ prioritaire, ouvrir le § correspondant (via `spec-map.md`), lire, implémenter.
3. **Croiser 3 sources** pour chaque point : (a) la doc = ce qui est *spécifié* ; (b) un vrai
   `.pptx` du corpus = ce que PowerPoint *écrit réellement* (souvent un sous-ensemble + des
   extensions non spécifiées) ; (c) le rendu de référence LibreOffice = ce que ça *doit donner*.
4. **Implémenter dans le code** de façon générique (jamais par diapo), puis
   `node renderer/check.mjs <dossier>` (0 erreur attendu) + vérif visuelle.
5. **Mettre à jour `coverage.md`** (statut) et `concepts.md` (nouveau piège) et cocher
   `open-questions.md`.

## Sources de référence (libres)

- **ECMA-376** (téléchargeable gratuitement, techniquement équivalent à ISO 29500) — la spec.
- **datypic.com/sc/ooxml/** — miroir navigable des schémas XSD (élément par élément).
- **c-rex.net** — autre miroir de la spec, pratique pour les §.
- **MS-OI29500 / MS-OE376** (Microsoft Open Specs) — les *écarts d'implémentation* PowerPoint vs
  standard (crucial : PowerPoint ne suit pas la spec à la lettre).
- **Code LibreOffice** `oox/source/ppt/` (import) et `slideshow/` (exécution animation) —
  implémentation de référence lisible.
- Notre reverse binaire : `docs/binaire-ppcore.md`, `docs/dynamique-runtime.md`.

## État actuel (résumé)

Rendu statique fidèle validé sur **548 diapos / 2 UV** (0 erreur). Lecteur animé (entrées au clic)
fonctionnel. Voir `coverage.md` pour le détail élément par élément et ce qui reste.
