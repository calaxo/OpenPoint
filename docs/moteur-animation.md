# Le moteur d'animation PowerPoint, reverse-engineeré

Étude menée le 2026-07-22 sur le deck `sample` (28 diapos, 21 animées, 280 effets),
en croisant trois sources :

1. le **XML brut** du PPTX (`ppt/slides/slideN.xml`, élément `<p:timing>`) ;
2. l'**object model COM** de PowerPoint 16 (la façon dont PowerPoint lui-même décrit ses effets) ;
3. la **ré-exécution** dans un lecteur maison (`poc/player.html`) pour valider la sémantique.

Résultat central : **le croisement XML ↔ COM aligne 280/280 effets sans un seul désaccord**, ce qui
valide toutes les règles décrites ici.

---

## 1. Où vivent les animations

Un `.pptx` est un zip. Pour chaque diapo, `ppt/slides/slideN.xml` contient :

- `<p:cSld><p:spTree>` — les formes (géométrie DrawingML, texte, images) ;
- `<p:timing>` — **tout** le moteur d'animation de la diapo ;
- `<p:transition>` — la transition de diapo (mécanisme séparé, pas traité ici).

Il n'y a **rien d'autre** : pas de code, pas de binaire. Le moteur d'exécution de PowerPoint est un
interpréteur de cet arbre XML, dont le modèle est copié sur **SMIL 2.0** (le standard W3C
d'animation déclarative). C'est documenté formellement dans **ECMA-376 Part 1, §19.5**
(PresentationML — Animation) — le vocabulaire (`par`, `seq`, `dur`, `fill`, `restart`,
`indefinite`…) vient directement de SMIL.

## 2. L'arbre temporel : la hiérarchie complète

Chaque nœud temporel est un `<p:cTn>` (*common time node*) avec un cycle de vie propre
(délai de départ, durée, comportement de fin). Les conteneurs sont `<p:par>` (enfants en
parallèle) et `<p:seq>` (enfants en séquence). La hiérarchie réelle, constante sur tout le
corpus :

```
<p:timing>
 └ <p:tnLst>
    └ <p:par>  cTn id=1  dur="indefinite" restart="never" nodeType="tmRoot"
       └ <p:seq> concurrent="1" nextAc="seek"     ← LA séquence principale ("mainSeq")
          ├ cTn id=2 dur="indefinite" nodeType="mainSeq"
          │  ├ <p:par> cTn delay="indefinite"     ← 1 nœud par CLIC (étape de construction)
          │  │  └ <p:par> cTn delay="0"           ← groupe de synchro (regroupe with/after)
          │  │     └ <p:par> cTn presetID=.. presetClass=.. nodeType="clickEffect"  ← l'EFFET
          │  │        ├ <p:set>        (behavior)
          │  │        ├ <p:animEffect> (behavior)
          │  │        └ ...            (behaviors)
          │  ├ <p:par> cTn delay="indefinite"     ← clic suivant
          │  └ ...
          ├ <p:prevCondLst> cond evt="onPrev" tgt=sldTgt   ← navigation arrière
          └ <p:nextCondLst> cond evt="onNext" tgt=sldTgt   ← navigation avant
```

Points clés :

- **`delay="indefinite"`** = « n'avance pas tout seul, attends un événement ». C'est comme ça
  qu'un clic est encodé : chaque étape de clic est un `par` gelé, et le `nextCondLst`
  (`evt="onNext"`) de la séquence principale libère l'étape suivante. `nextAc="seek"` signifie
  qu'un clic pendant une animation **saute à la fin** de l'étape en cours au lieu d'empiler.
- **`fill="hold"`** = à la fin de l'animation, la valeur finale **reste appliquée** (sinon la
  propriété reviendrait à sa valeur de base — comportement SMIL par défaut).
- Il peut exister d'autres `<p:seq>` frères de la mainSeq : les **séquences interactives**
  (déclencheur = clic sur une forme précise, `cond evt="onClick" tgt=spTgt`). Absentes de ce
  corpus.

## 3. « Au clic / Avec le précédent / Après le précédent »

C'est l'enseignement le plus utile pour ré-implémenter :

- `nodeType="clickEffect|withEffect|afterEffect"` n'est qu'une **étiquette pour l'UI** du volet
  d'animation. Le moteur n'en a pas besoin.
- La vérité d'exécution est **structurelle** : tout ce qui appartient au même `par` de clic
  démarre au même clic, et les décalages sont encodés en **délais absolus depuis le clic**,
  répartis sur les `cTn` imbriqués. Un « après le précédent » de 500 ms après un effet de
  1000 ms est encodé `delay="1500"` (cumul des délais des groupes parents + le sien).
  **Il n'y a pas de chaînage « à la fin de X » dans les faits** — PowerPoint pré-calcule tout en
  offsets numériques à l'enregistrement (des conditions d'événement existent dans le schéma,
  mais ce deck réel n'en contient aucune à ce niveau).
- Conséquence pratique : pour rejouer, il suffit d'un `setTimeout(délai_cumulé)` par effet.
  C'est exactement ce que fait `poc/player.html`.

## 4. Les behaviors : les vraies primitives du moteur

Un « effet » (Fondu, Balayage…) n'est **pas** une primitive. Chaque effet est un petit paquet de
*behaviors*, les seules choses que le moteur sait exécuter. Le schéma en définit 8 ; occurrences
dans le corpus :

| Behavior | Rôle | Corpus |
|---|---|---|
| `p:set` | pose une valeur discrète (quasi toujours `style.visibility`) | 274 |
| `p:animEffect` | transition raster/clip sur la forme (`filter="fade"`, `wipe(left)`, `circle(out)`…) | 231 |
| `p:anim` | interpolation générique d'une propriété via keyframes (`p:tavLst`) | 103 |
| `p:animScale` | homothétie (`p:by x/y` en millièmes de %) | 9 |
| `p:animMotion` | déplacement le long d'un chemin (syntaxe type SVG path) | 0 |
| `p:animRot` | rotation | 0 |
| `p:animClr` | interpolation de couleur | 0 |
| `p:cmd` | commande (son, media, `onstopaudio`…) | 0 |

Chaque behavior porte un `<p:cBhvr>` commun : son propre `cTn` (durée `dur` en **ms**, délai,
`autoRev`, `repeatCount`), sa cible `<p:tgtEl><p:spTgt spid="18"/>` et la liste des propriétés
attaquées `<p:attrNameLst>`.

### 4.1 `p:set` — l'ossature de tout

Toute entrée commence par :
```xml
<p:set>
  <p:cBhvr><p:cTn dur="1" fill="hold"/> … </p:cBhvr>
  <p:to><p:strVal val="visible"/></p:to>
</p:set>
```
`dur="1"` (1 ms) + `fill="hold"` = un « échelon » qui rend la forme visible et le reste. Les
sorties font l'inverse (`hidden`) en fin d'effet.

### 4.2 `p:animEffect` — les transitions visuelles

`transition="in|out"` + `filter` : `fade`, `wipe(left|right|up|down)`, `circle(out)`, `box(out)`,
`barn(outHorizontal)`, `split`, etc. Le moteur applique un masque/filtre progressif sur la durée.
En HTML, `opacity` et `clip-path` suffisent à reproduire tous ceux du corpus.

### 4.3 `p:anim` — keyframes + expressions

```xml
<p:anim calcmode="lin" valueType="num">
  <p:cBhvr><p:cTn dur="1000" fill="hold"/> … <p:attrName>ppt_y</p:attrName></p:cBhvr>
  <p:tavLst>
    <p:tav tm="0">     <p:val><p:strVal val="#ppt_y+.1"/></p:val></p:tav>
    <p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav>
  </p:tavLst>
</p:anim>
```
- `tm` va de 0 à **100000** = position dans la durée en millièmes de pour cent.
- Les valeurs sont des **expressions** : `#ppt_y` = valeur de base de la propriété (préfixe `#`),
  `+.1` = décalage. C'est ainsi qu'« Ascendre » = « pars de ta position +10 % de diapo et
  remonte ». `calcmode="lin|discrete|fmla"` ; un attribut `fmla` par keyframe permet des
  formules complètes (sin, cos…) — non utilisé dans ce corpus.
- `valueType="num|clr|str"`.

### 4.4 Le système de propriétés animables

Espace de noms hérité du moteur HTML de PowerPoint 2002 (d'où les noms) :

| Propriété | Sens | Unité |
|---|---|---|
| `style.visibility` | visible/hidden | discret |
| `ppt_x`, `ppt_y` | position du **centre** de la forme | fraction de la largeur/hauteur de diapo (0..1) |
| `ppt_w`, `ppt_h` | taille | fraction de diapo |
| `ppt_r` / `style.rotation` | rotation | degrés |
| `xshear` | cisaillement | — |
| `style.opacity`, `fillcolor`, `stroke.color`… | style | — |

Le repère fractionnaire est ce qui rend les animations indépendantes de la résolution.

## 5. Les « presets » : des recettes, pas des primitives

`presetClass` (`entr`/`emph`/`exit`/`path`/`mediacall`), `presetID` (numéro d'effet dans l'UI) et
`presetSubtype` (direction/variante) ne servent **qu'à ré-afficher le bon nom dans le volet
d'animation**. À l'enregistrement, PowerPoint « compile » le preset en behaviors ; à l'exécution,
seuls les behaviors comptent. Table validée empiriquement (280 paires XML↔COM, 0 conflit) :

| `presetClass:ID:subtype` | Effet (enum COM `msoAnimEffect`) | Occurrences |
|---|---|---|
| entr:1:0 | Appear | 42 |
| entr:2:2 / entr:2:8 | Fly (depuis bas / gauche) | 2 |
| entr:4:32 | Box (out) | 2 |
| entr:6:32 | Circle (out) | 19 |
| entr:10:0 | Fade | 109 |
| entr:16:42 | Split | 1 |
| entr:22:1/2/4/8 | Wipe (↑ / ↓ / ← / →) | 35 |
| entr:23:288 | Zoom | 1 |
| entr:34:0 | LightSpeed | 3 |
| entr:42:0 | Ascend | 6 |
| entr:53:16 | FadedZoom (zoom fondu) | 38 |
| emph:26:0 | FlashBulb (= `animScale` by 105 % + `autoRev`) | 6 |
| exit:1:0 | Appear (sortie) | 1 |
| exit:10:0 | Fade (sortie) | 15 |

Le subtype encode la direction en bitmask (wipe : 1=bas→haut… ; fly : 2=bas, 8=gauche ;
32=« out » pour les effets radiaux).

**Conséquence majeure** : pour restituer fidèlement un deck, il est inutile d'implémenter
« 50 effets ». Il faut implémenter **les 8 behaviors** (et en pratique 4 ont suffi pour 100 %
de ce deck : `set`, `animEffect`, `anim`, `animScale`).

## 6. Règles du runtime découvertes en chemin

Validées par la ré-exécution dans `poc/player.html` :

1. **Visibilité initiale** : une forme dont le premier effet de la séquence est une **entrée**
   démarre cachée (le XML ne le dit nulle part — c'est le moteur qui le déduit). Sans cette
   règle, tout est visible avant le premier clic.
2. **Cible = groupe** : `spid` peut désigner un `p:grpSp` ; l'animation s'applique alors à
   **toutes les formes membres**.
3. **Placeholders** : une forme `p:ph` sans `<a:xfrm>` hérite sa géométrie du slideLayout
   (voire du master) — même piège que celui documenté sur le POC les suites bureautiques du LMS.
4. **`p:bldLst`** (après `tnLst`) décrit la « construction » : `<p:bldP spid grpId animBg="1"/>`
   par forme animée ; `build="p"` déclencherait un build **par paragraphe** (la cible devient
   `<p:txEl><p:pRg st end/>`). Corpus : builds par forme entière uniquement.
5. Durées/délais en **ms**, keyframes en **millièmes de %**, tailles/positions en **fractions
   de diapo**, angles en **60000e de degré** (dans la géométrie), scale en **millièmes de %**.

## 7. La même chose vue de l'object model COM

`Presentation.Slides(n).TimeLine` :

- `.MainSequence` — collection d'`Effect` **dans l'ordre document du XML** (validé 280/280) ;
- `.InteractiveSequences` — les séquences à déclencheur.

Chaque `Effect` expose `EffectType` (enum `msoAnimEffect` — la vue « preset »), `.Shape`,
`.Timing` (`TriggerType` = OnPageClick/WithPrevious/AfterPrevious, `TriggerDelayTime`,
`Duration`, `AutoReverse`, `RepeatCount`), `.EffectInformation` et `.Behaviors` — la vue
« behaviors », en miroir exact du XML (`msoAnimTypeSet=8`, `Filter=7`, `Property=5`,
`Scale=3`…). Scripts : `tools/dump-com.ps1`.

À noter côté export : `Presentation.CreateVideo()` (COM) rend les animations en MP4 avec un
timing de clic paramétrable (`DefaultSlideDuration`) — c'est la seule sortie « fidèle pixel »
native sans réimplémenter le moteur.

## 8. Ce que ça implique pour « sortir les animations proprement »

Trois stratégies, maintenant qu'on sait ce qu'il y a dedans :

1. **Ré-exécution web** (la voie prouvée par ce POC) : parser `<p:timing>` (≈300 lignes de
   Node), rendre le `spTree`, implémenter les behaviors. La *timeline* est la partie facile et
   100 % fiable ; la difficulté réelle est le **rendu des formes** (custGeom, texte riche,
   héritages layout/master) — c'est là que le POC reste approximatif et c'est là que les suites bureautiques
   apportait sa valeur.
2. **Rendu hybride** : fond de diapo rendu en image (LibreOffice/PowerPoint → PNG par forme ou
   par calque), timeline rejouée par-dessus en HTML. Contourne le rendu des formes en gardant
   les animations natives.
3. **Vidéo** : `CreateVideo` COM (nécessite un Windows + Office). Fidélité parfaite, mais on perd
   l'interactivité du clic.

## 9. Limites de cette étude

- Corpus sans `animMotion` (chemins), `animRot`, `animClr`, builds par paragraphe, triggers
  interactifs, `iterate` (animation lettre par lettre) ni `repeatCount` — sémantique connue via
  ECMA-376 mais non validée par ré-exécution ici.
- Les easing (`accel`/`decel`, `spd`) existent dans le schéma, non rencontrés.
- Le lecteur POC approxime le rendu (pas de custGeom, texte simplifié, couleurs de thème
  partielles) : il prouve le **moteur temporel**, pas le rendu.

## 10. Reproduire l'étude

```
npm install
powershell -File tools/dump-com.ps1      # vue COM (nécessite PowerPoint installé)
node tools/parse-timing.mjs              # XML -> specimen/timeline.json
node tools/crosscheck.mjs                # validation croisée + table des presets
node tools/export-poc.mjs                # données du lecteur
start poc/player.html                    # rejouer les diapos au clic
```

Références : ECMA-376 Part 1 §19.5 (PresentationML Animation), SMIL 2.0 (W3C),
documentation VBA `TimeLine`/`Effect`/`AnimationBehavior`.
