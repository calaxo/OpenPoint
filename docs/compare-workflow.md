# Comparateur d'images : traquer les micro-écarts vs PowerPoint (oracle)

Outil demandé pour être précis sur les formes/texte. Compare MON rendu à la référence pixel-perfect
de PowerPoint et LOCALISE les écarts.

## Pipeline
1. **Réf PowerPoint** (oracle, pixel-perfect) : `powershell -File scratchpad/ppt-export.ps1 <pptx> <outDir> <slideN>`
   → `outDir/slide-N.png` (1280×720, export COM).
2. **Mon rendu en PNG** : `node scratchpad/rs.mjs <deckWork> <slideN> <name>` (SVG figé 1280×720 haut-gauche,
   Poppins embarquée + autofit) → ouvrir `scratchpad/<name>.html` dans chrome-devtools → `take_screenshot(filePath=mine.png)`.
3. **Diff** : `node renderer/compare.mjs <mine.png> <ref.png> [diff.png]`
   → JSON `{diffPct, topRegions:[{gx,gy,pct,x,y}]}` (grille 16×9) + image de diff (rouge = écart).

## Comment lire le résultat
- **Écart STRUCTUREL** (à corriger) = une région localisée à fort % (forme manquante, mauvaise
  couleur, élément décalé) OU du texte DÉDOUBLÉ dans la diff (lignes désalignées → interligne/position).
- **Bruit de POLICE** (plafond FOSS, non corrigeable sans la police exacte) = rouge fin sur les BORDS
  des lettres réparti partout + léger dédoublement en bas des longs paragraphes (wrapping ma Poppins
  ≈ mais ≠ celle de PowerPoint). `compare.mjs` a `includeAA` pour capter, mais ce bruit est attendu.

## Trouvé/corrigé avec (commit 3c42cab)
- **Interligne lstStyle non hérité** : les cartes ont `<a:lnSpc spcPct=90%>` dans le lstStyle du
  layout ; je ne l'héritais pas → lignes trop espacées → dédoublement. Fix : `def.lnSpc` hérité
  (inline→layout→master), facteur 1.2 (spcPct 100% = interligne simple ≈ 1.2× police). Cartes d.5 1-3 alignées.
- Résidu d.5 carte 4 : `lnSpc spcPts=2200` (22pt) OK, mais wrapping différent (largeur Poppins) →
  plafond métrique.

## Prochaines cibles (balayer avec le comparateur)
Screenshoter + comparer chaque diapo d'un deck, trier par `diffPct` et regarder les `topRegions`
localisés (hors bruit de police) → corriger le structurel générique uniquement.
