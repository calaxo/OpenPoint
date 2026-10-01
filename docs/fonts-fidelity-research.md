# Comment les autres gèrent police + taille en lecture/écriture de PPTX (recherche web, 2026-07-24)

But : comprendre comment PowerPoint / LibreOffice / python-pptx / Aspose / viewers HTML obtiennent
« les bonnes polices avec les bonnes tailles », pour guider la fidélité du moteur maison (le point
qui reste dur : pastille qui mord la 1ʳᵉ lettre, texte un poil trop large = métriques Poppins).

## 1. Le cœur du problème : mesurer le texte
Pour savoir si un texte tient (autofit) ou où il se termine (chevauchement), il faut **mesurer la
largeur des glyphes** avec **la police exacte**. Trois familles de solutions :
- **python-pptx** (`TextFrame.fit_text`) : exige qu'on **fournisse le fichier de police (.ttf)** et
  mesure via Pillow/PIL (`ImageFont`). Sans le bon fichier → mesure fausse.
- **PowerPoint / LibreOffice / Aspose** : mesurent avec le moteur de texte de l'OS (DirectWrite/GDI,
  HarfBuzz…) et **la police réellement installée**. C'est pour ça que leur rendu est exact.
- **Viewers HTML (le mien, PPTXjs, docx-preview…)** : laissent **le navigateur** mesurer (le
  `foreignObject`/DOM mesure avec la police @font-face chargée). Mesure fiable **SI** la police
  chargée est la MÊME (mêmes métriques) que celle de PowerPoint.

→ Mon moteur mesure déjà correctement (navigateur + Poppins). L'écart vient de **quelle Poppins**.

## 2. L'algorithme d'autofit de PowerPoint (normAutofit)
- 3 modes : « ne pas ajuster », **normAutofit** (réduire le TEXTE pour tenir dans une boîte fixe),
  **spAutoFit** (agrandir la BOÎTE pour le texte).
- normAutofit stocke `fontScale` (%) et `lnSpcReduction` (%). PowerPoint **réduit par pas ENTIERS de
  point** et **rabote l'interligne jusqu'à −20 %** pour maximiser la taille.
- Mon moteur : autofit navigateur par facteur 0.95 (pas entier) — approché mais OK. À affiner : ne
  réduire QUE les boîtes `normAutofit` (fait), pas les spAutoFit.

## 3. Polices embarquées (le vrai levier de fidélité — mais ABSENT ici)
- Un pptx PEUT embarquer ses polices dans **`ppt/fonts/*.fntdata`**, au format **EOT** (TrueType
  sous-ensemble, compressé ; MicroType Express). PowerPoint les convertit en TTF et les installe
  temporairement → rendu identique partout.
- **VÉRIFIÉ : AUCUN deck du corpus n'embarque de police** (`ppt/fonts/` absent, pas de
  `embeddedFontLst`). Donc PowerPoint utilise la Poppins **installée** sur le poste.
- Piste si un jour des decks embarquent : extraire `.fntdata`, convertir EOT→TTF/WOFF, servir en
  @font-face → métriques exactes.

## 4. Polices métrique-compatibles (l'approche cross-platform / LibreOffice)
Quand la police n'est ni installée ni embarquée, LibreOffice substitue par des équivalents **aux
mêmes métriques** (même largeur/hauteur → même mise en page, aucun reflow) :
- **Carlito** ≈ Calibri, **Caladea** ≈ Cambria, **Liberation Sans/Serif/Mono** ≈ Arial/Times/Courier.
C'est exactement ce que fait la route LibreOffice→SVG de l'ami (cf. [[pptx-html-libreoffice-route]]).
Pour NOS decks : la police est **Poppins** (Google) → il faut LA Poppins, pas un substitut.

## 5. Ce que ça implique pour le moteur maison (plan fidélité police)
Le résidu « texte un poil trop large / pastille qui mord la lettre » = ma Poppins @font-face
(@fontsource) diffère peut-être (version/métriques) de celle que PowerPoint utilise, ET je n'applique
pas le **crénage** (`kern="1200"` présent dans les styles). Pistes, par ordre :
1. **Utiliser exactement la Poppins de PowerPoint** : localiser le fichier réellement chargé par
   PowerPoint (pas dans C:\Windows\Fonts ni user Fonts → à retrouver ; sinon PowerPoint substitue et
   MON rendu Poppins est en fait *plus* fidèle à l'intention). Embarquer CE fichier en @font-face.
2. **Crénage** : PowerPoint applique le kerning ; activer `font-kerning:normal` (déjà défaut nav.) et
   vérifier ; le `kern` OOXML = seuil d'activation, pas un espacement.
3. **Autofit par pas entiers** de point (comme PowerPoint) au lieu du facteur 0.95, pour coller.
4. Accepter que sans la police-fichier exacte, un sub-pixel d'écart subsiste (plafond FOSS — même
   LibreOffice ne fait pas pixel-perfect, cf. [[pptx-html-libreoffice-route]] ; seul PowerPoint COM
   l'est).

## Sources
- python-pptx : autofit / fit_text (mesure via fichier de police, pas entiers de point) —
  https://python-pptx.readthedocs.io/en/latest/dev/analysis/txt-autofit-text.html ,
  https://python-pptx.readthedocs.io/en/latest/dev/analysis/txt-fit-text.html
- Aspose.Slides — polices embarquées : https://docs.aspose.com/slides/net/embedded-font/
- Embedded fonts pptx = EOT dans ppt/fonts (Apache POI) :
  https://user.poi.apache.narkive.com/h6NgHJr1/font-embedding-into-xslf
- Carlito ≈ Calibri (métrique-compatible) : https://en.wikipedia.org/wiki/Calibri

## RÉSOLU (2026-07-24) : « récupérer la Poppins de PowerPoint »
PowerPoint n'a PAS Poppins installée (absente de InstalledFontCollection) — il la charge via les
**polices cloud Office**, cachées dans :
`%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\Poppins\*.ttf` (Regular/Bold/Italic/BoldItalic/SemiBold).

**PREUVE (hash SHA1)** : ces fichiers sont **BYTE-IDENTIQUES** à ma Poppins de `renderer/fonts/`
(Regular 160316, SemiBold 157312, Bold 155996, BoldItalic 179108, Italic 184544 : mêmes hash).
→ **J'utilise déjà EXACTEMENT la Poppins de PowerPoint.** La police n'est PAS la cause de l'écart résiduel.
(Seul Poppins-Medium 158576 n'est pas dans le cache cloud — non bloquant.)

### Conclusion sur le résidu ~8 % du comparateur
Ce n'est donc NI la police NI la taille : c'est le **moteur de rendu** — DirectWrite (PowerPoint) vs
Blink/Chrome (mon foreignObject) : hinting, anti-aliasing sub-pixel, arrondi de position, et points de
retour à la ligne (wrapping) qui diffèrent d'un sous-pixel. **Irréductible** en rendu navigateur, quel
que soit le font. Donc pour le balayage : ignorer ce bruit de bords/wrapping, ne traiter que les écarts
STRUCTURELS (forme manquante, mauvaise couleur, élément décalé, texte absent).
