# Analyse dynamique : instrumenter PowerPoint vivant (frida)

Complète [binaire-ppcore.md](binaire-ppcore.md). Là, analyse statique du binaire. Ici, on
**exécute** PowerPoint et on observe son moteur d'animation en direct avec [frida](https://frida.re).

## 1. L'injection dans Office fonctionne

Contrairement à ce qu'on pourrait craindre (mitigations Office), **frida s'injecte sans
problème** dans `POWERPNT.EXE` (`binary/frida_probe.py`). Modules chargés au runtime :

```
POWERPNT.EXE   base 0x7ff610660000
ppcore.dll     base 0x7ffd69d30000  size 0x2219000   <- l'interpréteur de timing
mso.dll        base 0x7ffdaed20000  size 0x254f000   <- OfficeArt / services partagés
mso20/30/40/50/98win32client.dll                     <- UI, graphismes Office
ocpptview.dll  (chargé à la demande)                 <- rendu
```

Mapping adresse statique ↔ runtime : `runtime = ppcore.base + (VA_statique − 0x180000000)`.
Toutes les VA de [binaire-ppcore.md](binaire-ppcore.md) sont donc hookables en direct.

## 2. Le compilateur preset → behaviors, capturé en direct

Montage (`binary/dyn_hooks.py`) : PowerPoint piloté en **COM** (`comtypes`) ouvre le deck ;
frida hooke `AppendBehavior` (VA `0x1806408b0`, arg3 = `wchar*` nom de propriété) ; puis on
appelle `slide.TimeLine.MainSequence.AddEffect(shape, presetID)` pour chaque effet et on
enregistre les behaviors que le moteur fabrique. Résultat brut :

| Preset appliqué (COM) | presetID | behaviors « propriété » construits (observés en direct) |
|---|---|---|
| Appear | 1 | `style.visibility` |
| Fly | 2 | `style.visibility`, `ppt_x`, `ppt_y` |
| Fade | 10 | `style.visibility` |
| Wipe | 22 | `style.visibility` |
| Circle | 6 | `style.visibility` |
| FadedZoom | 53 | `style.visibility`, `ppt_w`, `ppt_h`, `ppt_x`, `ppt_y` |
| Ascend | 42 | `style.visibility`, `ppt_x`, `ppt_y` |
| Zoom | 23 | `style.visibility`, `ppt_w`, `ppt_h` |

Ceci **prouve dynamiquement** la thèse du reverse de format : un preset n'est pas une
primitive, il se déplie en behaviors. Fly = déplacement (x,y) ; Zoom = échelle (w,h) ;
FadedZoom = échelle + recentrage ; tous précédés du `set` de visibilité.

Deux fonctions-recettes apparaissent dans les backtraces :
- **`ppcore VA ~0x1806405b2`** : la recette de **base** (toujours appelée) — pose le
  behavior `style.visibility`.
- **`ppcore VA ~0x1807151fc`** : la recette de **géométrie** — appelée uniquement pour les
  effets à mouvement/échelle, ajoute les behaviors `ppt_x/y/w/h`.

Constat fin : Fade/Wipe/Circle ne passent **que** par la recette de base via `AppendBehavior`.
Leur effet visuel (le fondu, le balayage) n'est donc **pas** un behavior de propriété : c'est un
`animEffect` (filtre), construit par un autre chemin (§3).

## 3. Le chemin des filtres (`animEffect`), en statique

Les filtres n'utilisent pas `AppendBehavior` (qui prend un nom de propriété). En cherchant les
xrefs des chaînes de filtres uniques (`binary/xref.py "circle(out)" "wipe(left)" ...`), tout
converge vers une grande fonction autour de `0x180719xxx`. Extrait (site `circle`) :

```asm
lea  rax, [rip + ...]   ; -> "circle(out)"
lea  rdx, [rip + ...]   ; -> "circle(in)"
cmp  ebx, 0x10          ; ebx = flag de direction dérivé du presetSubtype
cmovne rdx, rax         ; choisit la variante in/out
jmp  0x18071e23c        ; queue commune : émet l'animEffect(filtre choisi)
...
lea  rax, [rip + ...]   ; -> "checkerboard(down)"
lea  rdx, [rip + ...]   ; -> "checkerboard(across)"
cmove rdx, rax
```

Donc **le filtre et sa direction sont calculés depuis `presetSubtype` à la création de l'effet**,
puis gravés en chaîne (`filter="circle(out)"`). C'est exactement ce qu'on lit ensuite dans le
XML. La queue commune `0x18071e2xx` est le constructeur générique `emit animEffect(filter)`.

## 4. Méthode (réutilisable)

```
python -m pip install frida frida-tools comtypes capstone pefile
python binary/frida_probe.py                 # valider l'injection + bases des modules
python binary/dyn_hooks.py                    # COM pilote + hooks -> recettes de presets en direct
```

`dyn_hooks.py` travaille sur une **copie temporaire** du deck (jamais le specimen) et tue le
process en fin de run.

## 5. Ce qui reste (le chemin d'exécution image par image)

On a élucidé **l'écriture** des animations (authoring : preset → behaviors → XML). L'étape non
encore instrumentée est **la lecture pendant un diaporama** : l'évaluateur de formules `#ppt_*`
et l'application des valeurs interpolées frame par frame. Pistes concrètes (cf.
[notes/recherche-web.md](../notes/recherche-web.md)) :
- Ces fonctions ne se déclenchent **pas** à l'authoring ni au chargement (le XML contient déjà
  les behaviors dépliés) — il faut lancer un vrai diaporama (`pres.SlideShowSettings.Run()`) et
  tracer.
- Vu l'absence de symboles PPCORE, l'approche recommandée est **WinDbg + Time Travel Debugging**
  sur `POWERPNT.EXE`, ou hooker en aval les API **Direct2D/DirectComposition** (qui, elles, ont
  des symboles publics) pour lire les matrices de transformation effectivement appliquées.
- L'implémentation de référence lisible du langage de formules est
  `slideshow/source/engine/smilfunctionparser.cxx` de LibreOffice (boost::spirit).
