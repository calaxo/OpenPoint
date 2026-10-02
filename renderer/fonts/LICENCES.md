# Polices embarquées dans le lecteur de diapos

Toutes libres de redistribution. Elles remplacent les polices PowerPoint (propriétaires Microsoft /
Monotype, **non redistribuables**) quand le poste client ne les a pas : la police d'origine
reste prioritaire dans la pile CSS (`FONT_SUBST`, `renderer/pptx-render.mjs`).

| Fichiers | Remplace | Licence | Source |
|---|---|---|---|
| Poppins-* | Poppins (thème) | OFL 1.1 | Google Fonts |
| LiberationSans/Serif/Mono-* | Arial, Times New Roman, Courier New, Consolas | OFL 1.1 | Liberation Fonts |
| Carlito-*, Caladea-* | Calibri (+ Light), Cambria | OFL 1.1 | Google (ChromeOS) |
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

Les `.woff2` sont réduits aux plages utiles : `python build-web-fonts.py <dossier des TTF/OTF sources>`
(nécessite `pip install fonttools brotli`).

## Fichiers générés à côté des polices

- `metrics.json` : ascendante / descendante « win » de chaque police (celles qu'utilise PowerPoint),
  imposées au navigateur (`ascent-override` / `descent-override`) et utilisées par le moteur pour
  caler la ligne de base comme PowerPoint.
- `advances.json` : chasse de chaque caractère, pour reproduire l'arrondi des chasses de PowerPoint
  (1/8 pt par glyphe) et donc ses coupures de ligne. Lu côté serveur uniquement.

Les deux sont régénérés par `build-web-fonts.py`.
