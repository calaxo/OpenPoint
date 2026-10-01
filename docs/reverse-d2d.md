# Reverse du rendu Direct2D de PowerPoint (2026-07-23, en cours)

But : capturer la display-list EXACTE de PowerPoint (positions/tailles/glyphes) pour avoir sa mise
en page pixel-perfect, que le moteur maison approxime mal. Menée à la demande explicite de l'utilisateur
(« pousser le hook Direct2D »).

## Faits établis (empiriques, frida injecté dans POWERPNT.EXE)

1. **Le contenu des diapos est rendu en Direct2D (`d2d1.dll`), PAS en GDI.**
   - `ocpptview.dll` importe `ExtTextOutW/TextOutW/BitBlt/StretchBlt/CreateDIBSection` (GDI) MAIS
     hooker `ExtTextOutW` ne capture QUE le pied de page (Poppins) + l'UI (Segoe) — à (0,0), position
     réelle dans les BitBlt. Titre/badges/formes = ABSENTS de GDI. (script `binary/capture_gdi.py`)
   - Donc GDI = chrome only ; le contenu passe par D2D.

2. **Mécanisme de hook vtable D2D — VALIDÉ** (`binary/capture_d2d.py`) :
   - spawn PowerPoint sous frida ; d2d1 pas chargé au spawn → poll `Process.findModuleByName('d2d1.dll')`
     puis hook l'export `D2D1CreateFactory` (arg[3] = ppIFactory).
   - Depuis la factory, hook les créateurs de render target (indices vtable ABI stables) :
     **Factory** : 13 CreateWicBitmapRenderTarget, 14 CreateHwndRenderTarget, 15 CreateDxgiSurfaceRenderTarget,
     16 CreateDCRenderTarget, 17 CreateDevice (Factory1). ppOut = dernier arg.
     **Device** (Factory1::CreateDevice) : slot 4 = CreateDeviceContext.
   - Sur chaque render target / device context, hook les slots **ID2D1RenderTarget** :
     15 DrawLine, 16 DrawRectangle, 17 FillRectangle, 20 DrawEllipse, 21 FillEllipse, 22 DrawGeometry,
     23 FillGeometry, 25 FillOpacityMask, 26 DrawBitmap, 27 DrawText, 28 DrawTextLayout,
     **29 DrawGlyphRun** ; **DeviceContext** : 83 DrawImage, 84 DrawGdiMetafile.
   - Dédup par (vtable,slot). Tous les hooks s'installent bien (factory Wic/Hwnd/Dxgi/DC + Device).

3. **DÉCOUVERTE CLÉ — comment PowerPoint dessine** : sur ~12 s de rendu de la diapo 1 :
   - **`DrawImage` = 404 appels**. `DrawGlyphRun` = 0, `DrawText`/`DrawTextLayout` = 0,
     `FillGeometry`/`DrawGeometry` = 0, `FillRectangle` = 0.
   - → PowerPoint **ne dessine pas les primitives directement**. Il **compose via `DrawImage`**
     (ID2D1DeviceContext::DrawImage dessine un ID2D1Image = bitmap OU **ID2D1CommandList** OU sortie
     d'effet). Le vrai contenu (glyphes/géométries) est soit dans des **command lists** enregistrées,
     soit **pré-rastérisé en bitmaps** (atlas), puis blitté par DrawImage.

## VERDICT (2026-07-23, après instrumentation complète)

Compte des primitives sur le rendu de la diapo 1 (slots RT **et** DeviceContext hookés, dont
DrawGlyphRun_DC @82 qui porterait le TEXTE via DWRITE_GLYPH_RUN_DESCRIPTION) :
- `DrawImage` = **404**, `FillGeometry` = **123**, `DrawBitmap_DC` = 1.
- `DrawGlyphRun` = **0** (slot 29 RT ET slot 82 DeviceContext), `DrawText`/`DrawTextLayout` = 0,
  `CreateCommandList` = 0.

**Conclusion dure et définitive : PowerPoint ne rend PAS le texte via DrawGlyphRun/DrawText.**
Le texte est soit converti en **contours de glyphes → `FillGeometry`** (chemins de Bézier, aucune
sémantique « chaîne/police/taille »), soit **pré-rastérisé (atlas) → composité par `DrawImage`**.
Donc l'interception D2D ne fournit **PAS** une display-list « texte à (x,y) taille N » exploitable —
seulement des paths de glyphes anonymes ou des tuiles bitmap. Reconstituer la mise en page texte à
partir de ça = non réaliste (il faudrait re-OCR / re-associer les contours aux caractères).

→ **La voie « reverse D2D pour extraire la mise en page » est un cul-de-sac pour le TEXTE.**
Ce qui reste capturable : les **géométries de formes** (FillGeometry, 123 — badges/pastilles/flèches,
via un ID2D1GeometrySink maison) et la **composition** (DrawImage). Mais ça n'aide pas la fidélité
du texte, qui est le vrai problème.

**Le seul moyen d'avoir le rendu texte pixel-perfect de PowerPoint = laisser PowerPoint rendre**
(COM `Slide.Export` PNG / `ExportAsFixedFormat` PDF vectoriel). C'est établi, testé, et c'est
la recommandation finale pour un rendu parfait ; le moteur maison JS reste l'option portable + animée.

## Prochaine étape (si on insiste sur D2D — faible ROI)

Si le contenu est dans des **ID2D1CommandList** : les intercepter et rejouer via un
**ID2D1CommandSink** maison → ses callbacks (DrawGlyphRun, FillGeometry, DrawBitmap…) reçoivent TOUS
les primitives enregistrés avec coords exactes.
- Hook `ID2D1DeviceContext::CreateCommandList` (slot DeviceContext, à trouver — après les slots RT ;
  ID2D1DeviceContext ajoute ses méthodes après le slot ~57 de ID2D1RenderTarget). Récupérer le CL.
- Implémenter un `ID2D1CommandSink` en frida : allouer une vtable de ~28 `NativeCallback`, chaque
  méthode logguant ses args (DrawGlyphRun donne baselineOrigin + DWRITE_GLYPH_RUN{fontEmSize@+8,
  glyphCount@+12, glyphIndices@+16, glyphAdvances@+24}).
- Appeler `commandList->Stream(sink)` (slot ID2D1CommandList) sur les CL capturés (à la fermeture /
  avant DrawImage) → énumère les commandes.
- Alternative si atlas bitmap (pas de CommandList) : DrawImage donne targetOffset + ID2D1Image ;
  capturer les offsets = positions des tuiles, mais pas la sémantique texte → route command-list
  préférable. Vérifier d'abord si des `CreateCommandList` sont appelés.

## DWRITE_GLYPH_RUN (offsets x64, pour quand on capturera DrawGlyphRun)
fontFace@+0, fontEmSize(float)@+8, glyphCount(u32)@+12, glyphIndices(u16*)@+16,
glyphAdvances(float*)@+24, glyphOffsets@+32, isSideways@+40, bidiLevel@+44.
DrawGlyphRun(x64) : this=rcx, baselineOrigin(D2D1_POINT_2F, 2 floats packés)=rdx, glyphRun*=r8, brush=r9.

## Réalité stratégique (honnête)
- Extraire un moteur portable via ce reverse = lourd (GPU, command lists, sink COM à implémenter).
- Le rendu pixel-perfect DISPONIBLE tout de suite : **PowerPoint COM `Slide.Export` (PNG)** ou
  `ExportAsFixedFormat` (PDF vectoriel) — c'est PowerPoint qui rend. (`binary/capture_gdi.py` l'a
  produit en passant : `scratchpad/pp_export.png`.)
- Outillage : frida 17.16.4 (API : `Process.getModuleByName(x).getExportByName(y)`, plus de
  `Module.getExportByName`), comtypes, capstone/pefile. Scripts : `binary/capture_gdi.py`,
  `binary/capture_d2d.py`.
