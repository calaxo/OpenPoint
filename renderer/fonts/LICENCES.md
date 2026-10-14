# Polices embarquées dans le lecteur de diapos

Toutes libres de redistribution, sous LEUR propre licence (indépendante de la licence AGPL du code,
voir `../../LICENSE`). Textes complets dans [`licenses/`](licenses/). Elles remplacent les polices PowerPoint (propriétaires Microsoft /
Monotype, **non redistribuables**) quand le poste client ne les a pas : la police d'origine
reste prioritaire dans la pile CSS (`FONT_SUBST`, `renderer/pptx/fonts.mjs`).

| Fichiers | Remplace | Licence | Source |
|---|---|---|---|
| Poppins-* | Poppins (thème) | OFL 1.1 | Google Fonts |
| LiberationSans/Serif/Mono-* | Arial, Times New Roman, Courier New, Consolas | OFL 1.1 | Liberation Fonts |
| Carlito-* | Calibri (+ Light) | OFL 1.1 | Google (ChromeOS) |
| Caladea-* | Cambria | Apache 2.0 | Google (ChromeOS) |
| Gelasio-* | Georgia (mêmes métriques) | OFL 1.1 | Google Fonts |
| ComicRelief-* | Comic Sans MS (mêmes métriques) | OFL 1.1 | Google Fonts |
| ArchivoBlack-*, Anton-* | Arial Black, Impact (approchantes) | OFL 1.1 | Google Fonts |
| LibreFranklin-* | Franklin Gothic | OFL 1.1 | Google Fonts |
| NewsCycle-* | News Gothic MT | OFL 1.1 | Google Fonts |
| Cabin-* | Gill Sans MT (approchante) | OFL 1.1 | Google Fonts |
| SortsMillGoudy-* | Bernhard Modern Roman (approchante) | OFL 1.1 | Google Fonts |
| SourceSans3-* | Corbel, Aptos (approchantes) | OFL 1.1 | Google Fonts |
| Cairo-* | — (police réelle du modèle Slidesgo) | OFL 1.1 | Google Fonts |
| SpaceGrotesk-*, SpaceGrotesk-Medium | — (police réelle du modèle ; « Medium » = instance wght 500) | OFL 1.1 | Google Fonts |
| texgyreschola-*, texgyrepagella-* | Century Schoolbook, Palatino / Book Antiqua | GUST Font License | CTAN (TeX Gyre) |
| DejaVuSans-* (symboles seuls) | glyphes Symbol / Wingdings remappés | Bitstream Vera + domaine public | DejaVu |

Textes des licences : [OFL 1.1](licenses/OFL-1.1.txt) · [Apache 2.0](licenses/Apache-2.0.txt) ·
[GUST](licenses/GUST-FONT-LICENSE.txt) · [Bitstream Vera / DejaVu](licenses/DejaVu-LICENSE.txt).

## Mentions de copyright (lues dans les fichiers de police)

- Anton : Copyright 2020 The Anton Project Authors (https://github.com/googlefonts/AntonFont)
- Archivo Black : Copyright 2017 The Archivo Black Project Authors (https://github.com/Omnibus-Type/ArchivoBlack)
- Cabin : Copyright 2018 The Cabin Project Authors (https://github.com/impallari/Cabin)
- Cairo : Copyright 2009 The Cairo Project Authors (https://github.com/Gue3bara/Cairo)
- Caladea : Copyright (c) 2012 Huerta Tipografia
- Carlito : Copyright (c) 2010-2013 by tyPoland Lukasz Dziedzic, with Reserved Font Name "Carlito"
- Comic Relief : Copyright 2013 The Comic Relief Project Authors (https://github.com/loudifier/Comic-Relief)
- DejaVu Sans : Copyright (c) 2003 by Bitstream, Inc. ; Copyright (c) 2006 by Tavmjong Bah ; DejaVu changes are in public domain
- Gelasio : Copyright 2022 The Gelasio Project Authors (https://github.com/SorkinType/Gelasio)
- Liberation Sans / Serif / Mono : Digitized data copyright (c) 2010 Google Corporation ; Copyright (c) 2012 Red Hat, Inc.
- Libre Franklin : Copyright 2020 The Libre Franklin Project Authors (https://github.com/googlefonts/Libre-Franklin)
- News Cycle : Copyright (c) 2010-2011, Nathan Willis, with Reserved Font Name "News Cycle"
- Poppins : Copyright 2020 The Poppins Project Authors (https://github.com/itfoundry/Poppins)
- Sorts Mill Goudy : Copyright (c) 2010, Barry Schwartz
- Source Sans 3 : © 2023 Adobe (http://www.adobe.com/), with Reserved Font Name "Source"
- Space Grotesk : Copyright 2020 The Space Grotesk Project Authors (https://github.com/floriankarsten/space-grotesk)
- TeX Gyre Pagella / Schola : Copyright 2006-2018 B. Jackowski, J.M. Nowacki et al. (on behalf of TeX users groups)

Les `.woff2` sont réduits aux plages utiles : `python build-web-fonts.py <dossier des TTF/OTF sources>`
(nécessite `pip install fonttools brotli`).

## Fichiers générés à côté des polices

- `metrics.json` : ascendante / descendante « win » de chaque police (celles qu'utilise PowerPoint),
  imposées au navigateur (`ascent-override` / `descent-override`) et utilisées par le moteur pour
  caler la ligne de base comme PowerPoint.
- `advances.json` : chasse de chaque caractère, pour reproduire l'arrondi des chasses de PowerPoint
  (1/8 pt par glyphe) et donc ses coupures de ligne. Lu côté serveur uniquement.

Les deux sont régénérés par `build-web-fonts.py`.
