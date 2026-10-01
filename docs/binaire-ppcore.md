# Reverse du binaire : le moteur d'animation *dans* PowerPoint (PPCORE.DLL)

Ce document complète [moteur-animation.md](moteur-animation.md). Là, on avait reversé le
**format** (`<p:timing>`) et le **modèle COM** — de l'analyse boîte noire. Ici on ouvre le
**code machine de PowerPoint lui-même** : quel module natif implémente le moteur, et à quoi
ressemblent ses fonctions.

- **Cible** : `C:\Program Files\Microsoft Office\root\Office16\PPCORE.DLL`
- **Version** : PowerPoint 16.0.20131.20154 (x64), 34,1 Mo, build du 2026-07-20
- **Méthode** : analyse statique. PE parsé avec `pefile`, désassemblage x86-64 avec `capstone`.
  Pas de décompilation vers du C, pas de code source Microsoft — on lit les octets du binaire
  livré et on les interprète. `ImageBase` sur disque = `0x180000000` (l'ASLR décale à
  l'exécution, sans effet sur l'analyse statique). Outils dans `binary/`.
- **Cadre légal** (UE) : observation/test et étude de format = art. 5.3 dir. 2009/24/CE ;
  l'analyse statique ici sert la compréhension pour interopérabilité (rejouer les animations
  dans le LMS). On ne redistribue aucun octet de Microsoft — seulement des adresses et des
  constats.

## 1. Quel module ? PPCORE.DLL, prouvé par le vocabulaire

En cherchant le vocabulaire du moteur (relevé dans le format) **en chaînes littérales** dans les
binaires Office (`binary/scan-strings.mjs`) :

| Token | PPCORE.DLL | ocpptview.dll |
|---|---|---|
| `ppt_x` / `ppt_y` (propriétés animables) | 51 / 51 | 51 / 51 |
| `style.visibility`, `xshear`, `fillcolor` | ✓ | ✓ |
| `animEffect` | 4 | **0** |
| `clickEffect` / `withEffect` / `afterEffect` | ✓ | **0** |
| `mainSeq`, `tmRoot`, `presetClass`, `presetID`, `nodeType` | ✓ | **0** |

Conclusion : **`ocpptview.dll` est le moteur de rendu** (il connaît les propriétés `ppt_*` à
appliquer aux formes, mais rien du séquençage), et **`PPCORE.DLL` est l'interpréteur de la
timeline** (il possède tout le vocabulaire de séquence). Le reste de l'étude porte sur PPCORE.

## 2. Les tables de données du moteur, compilées dans la DLL

### 2.1 Table des tags du modèle temporel (offset fichier `0x1820480`)

Une suite contiguë de chaînes UTF-16 = l'énumération interne exacte du modèle que le format
expose (`binary/dump-region.mjs`) :

```
presetClass  presetSubtype  nodeType  presetID  afterEffect  indefinite  afterGroup
clickEffect  clickPar  interactiveSeq  mainSeq  tmRoot  withEffect  withGroup
always  never  whenNotActive  bldLvl  cTn  stCondLst  childTnLst ...
```

C'est le dictionnaire du sérialiseur/désérialiseur XML de `<p:timing>`.

### 2.2 Le langage d'expressions de keyframes, en dur (offsets `0x1993e00`, `0x1b11d00`)

Les gabarits que le moteur émet pour les trajectoires/tailles — le « langage » derrière les
`<p:tav>` que j'avais observés dans le fichier :

```
"#ppt_x; #ppt_x"      "#ppt_y; #ppt_y"        "#transform.scale_ctr_y"   ← propriété interne
"0; #ppt_w"           "0-#ppt_h/2;#ppt_y"     "ppt_y;0-ppt_h/2"          ← entrée depuis le haut
"1+#ppt_w/2;#ppt_x"   "1+#ppt_h/2;#ppt_y"     "0-#ppt_w/2;#ppt_x"        ← entrées gauche/droite/bas
"blinds(horizontal)"  "randombar(vertical)"   "opacity: "
```

`#transform.scale_ctr_y` révèle que le modèle **interne** est plus riche que ce que le format
sérialise (un système de transform centré, non exposé tel quel en XML).

## 3. Le code : enum `presetClass` → chaîne (jump table)

Site de référence à `"presetClass"` (`binary/xref.py`) → juste après, la fonction
`0x1802d6b7c` :

```asm
sub  rsp, 0x28
cmp  dword ptr [rcx], 5        ; l'enum ne vaut que 0..5  → 6 classes
ja   default
movsxd rdx, dword ptr [rcx]
lea  rax, [rip + 0x15a2769]    ; base d'une table de 6 pointeurs (VA 0x1818792F8)
mov  rdx, qword ptr [rax + rdx*8]
jmp  <writeString>            ; tail-call : écrit la chaîne obtenue
```

Lecture de la table de pointeurs (`binary/disasm.py ptrtab 0x1818792F8 6`) — **l'ordre réel de
l'enum interne de PowerPoint**, extrait du binaire :

| valeur | chaîne |
|---|---|
| 0 | `entr` |
| 1 | `exit` |
| 2 | `emph` |
| 3 | `path` |
| 4 | `verb` |
| 5 | `mediacall` |

(À noter : cet ordre interne diffère de l'ordre de déclaration du schéma ECMA-376 — c'est bien
la table de *l'implémentation*, pas de la spec.) La fonction sœur pour `nodeType`
(`0x1802d69dc`, appelée avant d'écrire le tag `"nodeType"`) suit le même motif de jump table.

## 4. Le code : le « compilateur » preset → behaviors

C'est le point clé de l'étude de format (« un effet n'est pas une primitive, il se compile en
behaviors »). On le retrouve **en code** :

- `"ppt_x"` est référencé par **158 instructions `lea`** distinctes dans `.text`, `"ppt_y"` par
  157 (`binary/xref.py`). Ce ne sont pas des données passives : ce sont 158 sites de code qui
  construisent des behaviors sur ces propriétés — la matérialisation des recettes de presets.

- Exemple, site `0x180640a71`, une recette de type mouvement :

```asm
lea  r8, [rip + ...]   ; -> "ppt_x"
mov  rdx, rsi
mov  rcx, rbx
call 0x1806408b0       ; AppendBehavior(builder, animInfo, "ppt_x")
lea  r8, [rip + ...]   ; -> "ppt_y"
mov  rcx, rbx
call 0x1806408b0       ; AppendBehavior(builder, animInfo, "ppt_y")
```

  → « ajoute un behavior sur X, puis un sur Y » : un preset directionnel qui se déplie en deux
  behaviors d'interpolation.

- Le helper `0x1806408b0` est un **appender de liste de behaviors** :

```asm
mov  ebx, dword ptr [rbp + 8]   ; count
shl  rbx, 4                     ; * 16  (taille d'une entrée = 16 octets)
add  rbx, qword ptr [rbp]       ; base + index  → &nouvelle entrée
call ...                        ; construit l'entrée
inc  dword ptr [rbp + 8]        ; count++
mov  rcx, rbx                   ; puis initialise l'entrée avec le nom de propriété (rdi)
```

  Structure classique `std::vector`-like d'entrées de 16 octets. C'est là que les 258 entrées,
  38 zooms fondus, etc. du deck étaient « fabriqués » à l'enregistrement.

## 5. Le code : sérialisation de `<p:cTn>`

Les tags du modèle sont écrits par une famille de petites fonctions au même moule (ex.
`0x1802d4c34` pour `afterEffect`, `0x1802d6b58` pour `presetClass`, `0x1802d69b7` pour
`nodeType`) :

```asm
mov  rdi, qword ptr [r8 + 0xNN]  ; champ optionnel de l'objet cTn
test rdi, rdi
je   skip                        ; absent -> on n'écrit pas l'attribut
... call <enum->string>          ; convertit la valeur
lea  rcx, [rip + ...]            ; -> "afterEffect" / "presetClass" / "nodeType"
call qword ptr [rip + ...]       ; writeAttribute(name, value)  (via table d'imports/vtable)
skip: ret
```

Chaque offset (`[r8+0xc8]`, `[r8+0x40]`, …) est un champ de la structure `cTn` en mémoire —
la disposition interne de l'objet nœud-temporel. C'est l'écriture du XML qu'on parse dans
`tools/parse-timing.mjs`, vue de l'autre côté.

## 6. Ce que ça apporte de plus que le reverse de format

- **Preuve d'implémentation** : le modèle déduit du fichier (presets = recettes de behaviors,
  vocabulaire de séquence, propriétés `ppt_*`) n'est pas une interprétation — il est câblé tel
  quel dans PPCORE.DLL (tables de tags, jump tables d'enums, 158 sites de construction).
- **Détails invisibles côté format** : l'ordre réel de l'enum `presetClass`, l'existence d'un
  modèle de transform interne (`#transform.scale_ctr_y`), la taille des entrées de behavior
  (16 o), la séparation nette interpréteur (PPCORE) / rendu (ocpptview).
- **Cartographie** pour aller plus loin : adresses des fonctions de (dé)sérialisation et du
  compilateur de presets, si un jour on veut suivre le flux jusqu'au calcul des frames.

## 7. Limites

- Analyse statique : on lit le code, on ne l'exécute pas sous debugger. Les cibles indirectes
  (`call qword ptr [rip+...]` = imports/vtables) ne sont pas toutes résolues.
- Binaire optimisé sans symboles publics complets pour Office : les noms de fonctions
  (`AppendBehavior`, etc.) sont **nos** étiquettes déduites du comportement, pas les symboles
  Microsoft.
- Version-dépendant : les adresses ci-dessus valent pour le build 16.0.20131.20154. La méthode
  (chercher le vocabulaire, remonter aux xrefs) reste valable sur toute version.

## 8. Reproduire

```
python -m pip install capstone pefile
node   binary/scan-strings.mjs "<...>\PPCORE.DLL"          # localiser le vocabulaire
node   binary/dump-region.mjs  "<...>\PPCORE.DLL" 0x1820480 0x260   # tables de données
python binary/xref.py          "<...>\PPCORE.DLL"          # xrefs + désassemblage des sites
python binary/disasm.py        "<...>\PPCORE.DLL" code   0x1802d6b7c 40
python binary/disasm.py        "<...>\PPCORE.DLL" ptrtab 0x1818792F8 6
```
