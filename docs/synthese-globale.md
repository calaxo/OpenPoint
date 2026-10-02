# Synthèse globale — comment fonctionne l'animation dans PowerPoint

Quatre points de vue croisés, du plus externe au plus interne. Chaque affirmation est adossée à
une preuve (fichier, désassemblage, hook runtime, ou spec).

| Niveau | Ce qu'on voit | Doc | Preuve |
|---|---|---|---|
| Format | l'arbre SMIL `<p:timing>` d'un `.pptx` | [moteur-animation.md](moteur-animation.md) | parseur + 280 effets |
| COM | ce que PowerPoint dit de ses effets | [moteur-animation.md](moteur-animation.md) §7 | dump COM, 280/280 alignés |
| Binaire (statique) | le code de PPCORE.DLL | [binaire-ppcore.md](binaire-ppcore.md) | désassemblage capstone |
| Binaire (dynamique) | le moteur en train de tourner | [dynamique-runtime.md](dynamique-runtime.md) | hooks frida live |
| Spec / OSS | la doc et les réimplémentations | [../notes/recherche-web.md](../notes/recherche-web.md) | ECMA-376, LibreOffice |

## Le modèle en une page

1. **Un `.pptx` est un zip.** Les animations d'une diapo vivent dans `ppt/slides/slideN.xml`,
   élément `<p:timing>` — un arbre déclaratif copié sur **SMIL 2.0** (héritage HTML+TIME de
   PowerPoint 2002, d'où les noms `ppt_*`). Aucun code, aucun binaire dans le fichier.

2. **L'interpréteur est `PPCORE.DLL`** (prouvé : c'est le seul module qui contient tout le
   vocabulaire de séquence en dur). `ocpptview.dll` est le **moteur de rendu** (il connaît les
   propriétés `ppt_*` mais rien du séquençage). Le rendu des formes s'appuie sur `mso*.dll`
   (OfficeArt) et la pile Direct2D/DirectComposition.

3. **La hiérarchie temporelle** (identique dans le fichier, le COM et l'enum interne du binaire) :
   `tmRoot → mainSeq (p:seq) → un p:par gelé par CLIC → groupes de synchro → effets → behaviors`.
   Un clic est encodé `delay="indefinite"` libéré par `nextCondLst`.

4. **Un « effet » n'est pas une primitive.** C'est une **recette compilée en behaviors** — les 8
   seules primitives du moteur (`set`, `animEffect`, `anim`, `animScale`, `animMotion`,
   `animRot`, `animClr`, `cmd`). Prouvé aux deux bouts :
   - *statiquement* : `ppt_x` est référencé par 158 sites de code appelant `AppendBehavior`
     (entrées de 16 o) ;
   - *dynamiquement* : en ajoutant un effet en COM, on voit le moteur construire la liste
     (Fly → `visibility+ppt_x+ppt_y`, Zoom → `visibility+ppt_w+ppt_h`, etc.).
   Le sous-type/direction est calculé depuis `presetSubtype` (ex. `circle(out)` vs `circle(in)`
   via un `cmov`) et gravé.

5. **« Avec/Après le précédent » n'existe pas à l'exécution** : tout est pré-calculé en délais
   absolus (ms) depuis le clic. Rejouer = un `setTimeout(délai)` par effet (prouvé par le POC).

6. **Le langage de keyframes** (`fmla` de `<p:tav>`) est un vrai mini-langage :
   `+ - * / ^ %`, `sin cos tan asin acos atan`, `ln exp sqrt`, `abs min max ?`, `pi e`, variables
   `#ppt_x/y/w/h` (fractions de diapo, x/y = centre de forme). `atan2` n'y est pas. Implémentation
   de référence lisible : `smilfunctionparser.cxx` (LibreOffice). Le binaire expose même un
   modèle interne plus riche que le format (`#transform.scale_ctr_y`).

## Table presetID → effet, vérifiée empiriquement

La spec ne publie **aucune** table `presetID → nom`. Celle-ci est établie par mesure : croisement
XML↔COM sur 280 effets réels (`tools/crosscheck.mjs`) **et** ajout d'effets en direct via COM
(`binary/dyn_hooks.py`). `EffectType` du COM == `presetID` du XML.

| presetClass | presetID | Effet (msoAnimEffect) | Behaviors typiques |
|---|---|---|---|
| entr | 1 | Appear | visibility |
| entr | 2 | Fly | visibility, ppt_x, ppt_y |
| entr | 4 | Box | visibility, animEffect box(out) |
| entr | 6 | Circle | visibility, animEffect circle(out) |
| entr | 10 | Fade | visibility, animEffect fade |
| entr | 16 | Split | visibility, animEffect |
| entr | 22 | Wipe | visibility, animEffect wipe(dir) |
| entr | 23 | Zoom | visibility, ppt_w, ppt_h |
| entr | 34 | LightSpeed | visibility + composite |
| entr | 42 | Ascend | visibility, ppt_x, ppt_y |
| entr | 53 | FadedZoom | visibility, ppt_w, ppt_h, ppt_x, ppt_y |
| emph | 26 | FlashBulb | animScale (by ~105 %, autoRev) |
| exit | 1 | Appear (sortie) | visibility→hidden |
| exit | 10 | Fade (sortie) | animEffect fade out |

Le sous-type encode la direction en bitmask (wipe : 1=↑ 2=↓ 4=← 8=→ ; fly : 2=bas 8=gauche ;
32 = « out » radial).

## Ce qui est solidement établi vs. ouvert

**Établi (avec preuve) :**
- Localisation du moteur (PPCORE) et du rendu (ocpptview) — dépasse la doc publique.
- Modèle temporel, primitives behaviors, presets = recettes — vérifié statique + dynamique.
- Enum interne `presetClass`, table de tags, langage d'expressions, tailles de structures.
- Chaîne d'authoring complète : UI/COM → recette → behaviors → sérialisation XML.

**Ouvert (méthode identifiée) :**
- Le chemin d'**exécution image par image** pendant un diaporama : évaluateur de formules `#ppt_*`
  et application des valeurs interpolées. Non déclenché à l'authoring/chargement. À capturer par
  WinDbg+Time Travel Debugging sur un diaporama, ou hooks Direct2D/DirectComposition en aval
  (symboles publics). Réf. lisible : `smilfunctionparser.cxx` de LibreOffice.
- `#ppt_c`/`#ppt_r` (couleur/rotation) : conventions non confirmées par source primaire ;
  `animMotion`/`animRot`/`animClr` absents du corpus (sémantique connue, non rejouée).

## Fichiers du repo

- Format : `tools/*.mjs` (parseurs d'animation + lecteur POC)
- Binaire statique : `binary/scan-strings.mjs`, `binary/dump-region.mjs`, `binary/xref.py`,
  `binary/disasm.py`, `binary/pe-info.py`
- Binaire dynamique : `binary/frida_probe.py`, `binary/dyn_hooks.py`
