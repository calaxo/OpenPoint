// ═══ [DIAPO · TEXTE] ═══ renderText : a:txBody -> <foreignObject> HTML (un <div data-par> par paragraphe).
import { PF, at, kids, byTag, first, tagOf, esc, serialize } from '../xml.mjs';
import { u } from '../units.mjs';
import { resolveColorFlat, cssColor } from '../color.mjs';
import { fontFamily, metricsOf, VA_TOK, applyShift, baselineShift, advanceFix, pptTextWidth } from '../fonts.mjs';
import { symbolParts, bulletSym, SYM_FAMILY } from '../symbols.mjs';
import { bulletOf } from '../text-styles.mjs';
import { renderMath } from '../math.mjs';
import { imageHref } from '../parts.mjs';

const arr = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];

// renderText(rc, txBody plat, w, h (EMU), phKind, inhLst, txBodyO, inhBodyPr, shapeStyle, textRect)
//   txBody     : a:txBody / p:txBody en arbre PLAT (propriétés) ;
//   txBodyO    : le même en arbre ORDONNÉ (ordre a:r / a:br / a:fld / formules), peut être null ;
//   phKind     : 'title' | 'body' | 'none' -> quel p:txStyles du master s'applique ('none' = aucun) ;
//   inhLst / inhBodyPr : lstStyle / bodyPr hérités (placeholder de layout ou style de tableau) ;
//   shapeStyle : p:style de la forme (fontRef = police/couleur de thème par défaut) ;
//   textRect   : rectangle de texte du preset (EMU) ou null.
// Sortie : '' si aucun texte visible, sinon <foreignObject> (repère local de la forme, en pt)
// contenant <div class="af"> (flex colonne = ancrage vertical) > <div data-par> par paragraphe >
// <span> par run.
//
// CASCADE DES PROPRIÉTÉS D'UN RUN (du plus fort au plus faible), reproduisant PowerPoint :
//   a:rPr du run  >  a:pPr du paragraphe (algn, marL, indent, puces, lnSpc…)
//   >  a:lstStyle de la forme (niveau lvl)  >  lstStyle du placeholder de layout (ou style de tableau)
//   >  lstStyle du placeholder de MASTER  >  p:txStyles du master (titleStyle / bodyStyle selon phKind)
//   >  p:defaultTextStyle de
//   presentation.xml  >  p:style/a:fontRef de la forme  >  couleur tx1 / taille 18 pt / police mineure.
// Le niveau lvl d'un paragraphe = a:pPr@lvl + 1 (lvl 0..8 -> lvl1pPr..lvl9pPr).
//
// CORRECTIONS « À LA POWERPOINT » (le navigateur fait la mise en page, on corrige ses écarts) :
//   interligne simple 1,2 × taille ; ligne de base (baselineShift, via VA_TOK) ; chasses arrondies au
//   1/8 pt (advanceFix -> letter-spacing) ; runs en line-height:0 ; tabulations calculées en pt ;
//   margin-right 0,2 px (règle de coupure) ; spcBef du 1er et spcAft du dernier paragraphe ignorés ;
//   « tailCut » sous la dernière ligne ; espaces de fin de paragraphe supprimés.
//
// Autofit : a:normAutofit@fontScale / @lnSpcReduction (valeurs CALCULÉES et ENREGISTRÉES par
// PowerPoint) sont appliquées telles quelles ; a:spAutoFit (forme ajustée au texte) n'a pas de
// traitement particulier ; a:noAutofit -> débordement libre.
export function renderText(rc, txBody, w, h, phKind, inhLst, txBodyO, inhBodyPr, shapeStyle, textRect, mstLst) {
  if (!txBody) return '';
  const { ctx, curRels, pptDir, embed, slideNum } = rc;
  const bodyPr = txBody['a:bodyPr'] || {};
  // texte VERTICAL (§21.1.2.1.1 @vert) : vert = tourné de 90° (lecture de haut en bas), vert270 = -90°
  // (de bas en haut). On compose dans un cadre aux dimensions permutées puis on le fait pivoter autour
  // du centre. Combiné à la rotation de la forme : rot 90° + vert270 = texte à l'endroit (040_02_01 d.32,
  // rendu tête en bas avant).
  const vert = bodyPr['@vert'] ?? inhBodyPr?.['@vert'];
  if (vert === 'vert' || vert === 'vert270' || vert === 'eaVert') {
    // appel RÉCURSIF avec w/h permutés et vert forcé à 'horz' (copies superficielles : l'arbre
    // d'origine n'est pas modifié). textRect n'est pas repassé (il est exprimé dans le repère non tourné).
    // eaVert (texte vertical est-asiatique, caractères empilés) est approximé par une rotation de 90°.
    const flat = { ...txBody, 'a:bodyPr': { ...bodyPr, '@vert': 'horz' } };
    const inner = renderText(rc, flat, h, w, phKind, inhLst, txBodyO, { ...(inhBodyPr || {}), '@vert': 'horz' }, shapeStyle, null, mstLst);
    if (!inner) return '';
    const ang = vert === 'vert270' ? 270 : 90;
    return `<g transform="translate(${u(w/2)} ${u(h/2)}) rotate(${ang}) translate(${u(-h/2)} ${u(-w/2)})">${inner}</g>`;
  }
  // ── Boîte de texte : insets, ancrage, retour à la ligne ──
  // bodyPr hérité du placeholder de LAYOUT (insets + ancre) : une carte-placeholder avec
  // <a:bodyPr/> vide DOIT reprendre le tIns/anchor du layout (ex. tIns=288000 réserve la place
  // de la pastille). Priorité : valeur de la diapo → valeur héritée → défaut OOXML.
  const inhBP = inhBodyPr || {};
  const pick = (k, def) => bodyPr[k] !== undefined ? +bodyPr[k] : (inhBP[k] !== undefined ? +inhBP[k] : def);
  const anchor = bodyPr['@anchor'] ?? inhBP['@anchor']; // t/ctr/b
  // ancrage vertical (a:bodyPr@anchor t / ctr / b ; « just » et « dist » -> haut) -> justify-content
  // du conteneur flex en colonne
  const just = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
  // insets (marges intérieures) en EMU ; défauts OOXML : 0,1" (91440) à gauche/droite, 0,05" (45720)
  // en haut/bas
  let iL = pick('@lIns',91440), iR=pick('@rIns',91440), iT=pick('@tIns',45720), iB=pick('@bIns',45720);
  // insets plus grands que la forme (pastilles « 1 » de 19pt avec 11pt de marge de chaque côté) :
  // PowerPoint garde le CENTRE de la zone de texte. Sans ça la zone se colle à l'inset gauche/haut
  // (texte décalé en bas à droite), voire prend une largeur négative (foreignObject non rendu).
  // rectangle de texte de la forme (roundRect, ellipse…) : les insets s'appliquent À L'INTÉRIEUR
  // textRect = {l,t,r,b} en EMU depuis le coin de la forme ; r/b sont des COORDONNÉES, d'où w - r
  if (textRect) { iL += textRect.l; iT += textRect.t; iR += w - textRect.r; iB += h - textRect.b; }
  // insets trop grands : on ramène la zone à une largeur/hauteur nulle placée au milieu des insets
  if (iL + iR > w) { const c = (iL + w - iR) / 2; iL = c; iR = w - c; }
  if (iT + iB > h) { const c = (iT + h - iB) / 2; iT = c; iB = h - c; }
  // anchorCtr : le BLOC de texte est centré horizontalement (l'algn joue à l'intérieur du bloc).
  const ctrCss = (bodyPr['@anchorCtr'] ?? inhBP['@anchorCtr']) === '1' ? 'align-items:center;' : '';
  // wrap="none" : aucune coupure de ligne (seuls les sauts explicites comptent).
  const ws = (bodyPr['@wrap'] ?? inhBP['@wrap']) === 'none' ? 'pre' : 'pre-wrap';
  const wsAttr = ws === 'pre' ? ' data-wrap="none"' : '';
  // styles du master pour ce type de placeholder (phKind 'none' -> {} : aucun ; 'other' n'est jamais
  // choisi par les appelants actuels)
  const styleLvls = ctx.txStyles?.[phKind] || {};
  // couleur de dernier recours = tx1 du thème (clair en thème sombre) — jamais #000 en dur
  // (sinon texte noir invisible sur fond sombre, ex. pied de page).
  const DEFCOL = cssColor(resolveColorFlat({ 'a:schemeClr': { '@val': 'tx1' } }, ctx)) || '#000';
  // autofit : PowerPoint réduit la police (et l'interligne) pour faire tenir le texte
  // Seul le bodyPr PROPRE de la forme est lu ici (pas l'hérité). fontScale 62500 -> 0,625 ;
  // lnSpcReduction 20000 -> 0,2 (interligne réduit de 20 points de %)
  const naf = bodyPr['a:normAutofit'];
  const fontScale = naf && naf['@fontScale'] ? +naf['@fontScale']/100000 : 1;
  const lnRed = naf && naf['@lnSpcReduction'] ? +naf['@lnSpcReduction']/100000 : 0;
  // interligne SIMPLE = 1,2 x la taille pour toutes les polices (calibré contre PowerPoint : 72 pt de
  // pas à 60 pt en Poppins/Arial/Calibri/Times) ; 1,15 serrait chaque boîte de 4 %.
  // interligne du conteneur (sans unité, relatif au font-size) : 1,2, réduit par l'autofit, plancher 0,8
  const lineH = Math.max(0.8, 1.2 * (1 - lnRed));
  // +mn-lt/+mj-lt = polices du thème (mineure/majeure)
  const resolveFont = (tf) => !tf ? null
    : /^\+mn-/.test(tf) ? ctx.minorFont : /^\+mj-/.test(tf) ? ctx.majorFont : tf;
  // style par niveau, du + prioritaire au - : lstStyle INLINE de la forme -> lstStyle du
  // placeholder de LAYOUT hérité -> lstStyle du placeholder de MASTER -> txStyles du MASQUE.
  const localLst = txBody['a:lstStyle'] || {};
  const inheritLst = inhLst || {};
  // lstStyle du placeholder de MASTER : maillon de la cascade entre le layout (inheritLst) et
  // p:txStyles du master (styleLvls). Vide si la forme n'est pas un placeholder.
  const mstLstS = mstLst || {};
  // niveau absent -> repli sur le niveau 1 de la même liste (choix de ce code : évite de perdre le
  // style d'un lstStyle qui ne définit que lvl1pPr)
  const lvlOf = (lst, lvl) => lst[`a:lvl${lvl}pPr`] || lst['a:lvl1pPr'];
  const colOf = (d) => d?.['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null;
  // dernier recours : defaultTextStyle de la présentation (texte hors placeholder)
  const dflt = ctx.defaultTextStyle || {};
  // style de FORME <p:style><a:fontRef idx="minor|major"><couleur/></a:fontRef> : donne au texte de la
  // forme sa couleur ET sa police de thème PAR DÉFAUT (souvent la SEULE source de couleur d'une forme
  // thémée non-placeholder, ex. bloc/callout SmartArt). Injecté en DERNIER repli (après lstStyle/master/
  // defaultTextStyle, avant DEFCOL) -> n'agit QUE là où le texte n'a aucune couleur/police définie.
  const fref = shapeStyle?.['a:fontRef'];
  const frefRgb = fref ? resolveColorFlat(fref, ctx) : null;
  const fontRefColor = frefRgb ? cssColor(frefRgb) : null;
  const fontRefFont = fref?.['@idx'] === 'major' ? ctx.majorFont : fref?.['@idx'] === 'minor' ? ctx.minorFont : null;
  // lvlStyle(lvl) -> propriétés PAR DÉFAUT d'un paragraphe de ce niveau, fusionnées selon la cascade
  // (loc = lstStyle de la forme, inh = hérité, mas = master, dt = defaultTextStyle). Les couleurs
  // sont déjà en CSS ; sz en pt ; marL/ind/marR en EMU ; lnSpc/spcBef/spcAft = nœuds bruts.
  // NB : mas.marR n'existe pas (non lu depuis le master) et defaultTextStyle ne fournit que
  // sz/b/color/font/algn.
  const lvlStyle = (lvl) => {
    const loc = lvlOf(localLst, lvl), inh = lvlOf(inheritLst, lvl);
    // mp = niveau du placeholder de MASTER ; s'intercale entre le layout (inh) et p:txStyles (mas)
    const mp = lvlOf(mstLstS, lvl);
    const ld = loc?.['a:defRPr'], id = inh?.['a:defRPr'], md = mp?.['a:defRPr'];
    const mas = styleLvls[lvl] || styleLvls[1] || {};
    const dt = dflt[lvl] || dflt[1] || {};
    return {
      sz: (ld?.['@sz'] ? +ld['@sz']/100 : null) ?? (id?.['@sz'] ? +id['@sz']/100 : null) ?? (md?.['@sz'] ? +md['@sz']/100 : null) ?? mas.sz ?? dt.sz,
      color: colOf(ld) ?? colOf(id) ?? colOf(md) ?? mas.color ?? dt.color ?? fontRefColor,
      font: ld?.['a:latin']?.['@typeface'] || id?.['a:latin']?.['@typeface'] || md?.['a:latin']?.['@typeface'] || mas.font || dt.font || fontRefFont,
      // gras : le niveau le PLUS SPÉCIFIQUE qui définit @b gagne (un b="0" du layout doit battre
      // le b="1" du master) — un simple || laisserait le master l'emporter à tort.
      b: ld?.['@b'] !== undefined ? ld['@b'] === '1'
       : id?.['@b'] !== undefined ? id['@b'] === '1'
       : md?.['@b'] !== undefined ? md['@b'] === '1'
       : mas.b,
      algn: loc?.['@algn'] || inh?.['@algn'] || mp?.['@algn'] || mas.algn,
      bu: bulletOf(loc) ?? bulletOf(inh) ?? bulletOf(mp) ?? mas.bu,
      marL: (loc?.['@marL'] ? +loc['@marL'] : null) ?? (inh?.['@marL'] ? +inh['@marL'] : null) ?? (mp?.['@marL'] ? +mp['@marL'] : null) ?? mas.marL,
      ind: (loc?.['@indent'] ? +loc['@indent'] : null) ?? (inh?.['@indent'] ? +inh['@indent'] : null) ?? (mp?.['@indent'] ? +mp['@indent'] : null) ?? mas.ind,
      marR: (loc?.['@marR'] ? +loc['@marR'] : null) ?? (inh?.['@marR'] ? +inh['@marR'] : null) ?? (mp?.['@marR'] ? +mp['@marR'] : null) ?? mas.marR,
      // interligne hérité (lstStyle inline → layout → master). Ex. cartes = spcPct 90 % : sans ça
      // j'utilise 1.15 → lignes trop espacées → dérive verticale vs PowerPoint (vu au comparateur).
      lnSpc: loc?.['a:lnSpc'] ?? inh?.['a:lnSpc'] ?? mp?.['a:lnSpc'] ?? mas.lnSpc,
      // spcBef/spcAft résolus dans la même chaîne (inline lstStyle → layout → master → master placeholder → txStyles)
      spcBef: loc?.['a:spcBef'] ?? inh?.['a:spcBef'] ?? mp?.['a:spcBef'] ?? mas.spcBef,
      spcAft: loc?.['a:spcAft'] ?? inh?.['a:spcAft'] ?? mp?.['a:spcAft'] ?? mas.spcAft,
    };
  };
  // ── Boucle des paragraphes ──
  // counters[lvl] : compteur de numérotation automatique par niveau (jamais remis à zéro dans la
  // boîte, même entre listes séparées par des paragraphes non numérotés)
  const counters = {};
  const paras = arr(txBody['a:p']);
  // hasText : au moins un caractère visible dans la boîte (sinon aucun foreignObject n'est émis)
  let html = '', hasText = false;
  for (let pi = 0; pi < paras.length; pi++) {
    const p = paras[pi];
    const pPr = p['a:pPr'] || {};
    const lvl = (+pPr['@lvl'] || 0) + 1;
    const def = lvlStyle(lvl);
    // Hauteur de ligne CSS du paragraphe = lnSpc du paragraphe, sinon hérité (lot/master via def.lnSpc),
    // sinon lineH du conteneur. Sert de line-height au <div> ET à chacun de ses runs (cf. lh0 plus bas).
    // Sans ça (ancien « line-height:0 » sur les runs), la boîte de ligne était nulle : un paragraphe qui
    // passe à la ligne voyait ses lignes, et les paragraphes d'une même zone de texte se superposaient
    // tous sur la même ligne de base (bug « texte qui se chevauche »).
    const lnSpcNode = pPr['a:lnSpc'] || def.lnSpc;
    const paraLH = lnSpcNode?.['a:spcPct']?.['@val'] ? (+lnSpcNode['a:spcPct']['@val']/100000*1.2*(1-lnRed)).toFixed(3)
                 : lnSpcNode?.['a:spcPts']?.['@val'] ? (+lnSpcNode['a:spcPts']['@val']/100*(1-lnRed)).toFixed(1)+'px'
                 : lineH.toFixed(3);
    // ST_TextAlignType : l / ctr / r / just (dist, justLow, thaiDist -> 'left' ici)
    const algn = { ctr:'center', r:'right', just:'justify' }[pPr['@algn'] || def.algn] || 'left';
    let runs = arr(p['a:r']);
    // sauts de ligne <a:br/> et champs <a:fld> intercalés entre les runs : le parse plat range chaque
    // balise dans son tableau et perd l'ORDRE (6 lignes séparées par des a:br rendues bout à bout).
    // On le reprend du parse ordonné : les tableaux plats gardent l'ordre relatif de chaque balise.
    const paraSeqO = txBodyO && (p['a:br'] !== undefined || (p['a:fld'] !== undefined && runs.length))
      ? byTag(txBodyO, 'a:p')[pi] : null;
    // Reconstruction : on parcourt les balises du paragraphe ORDONNÉ et on reprend, dans cet ordre,
    // le k-ième a:r / a:fld du parse plat (même ordre relatif). a:br devient un marqueur { __br }.
    // L'index pi suppose que txBodyO a exactement les mêmes a:p que le txBody plat.
    if (paraSeqO) {
      const flds = arr(p['a:fld']); let r = 0, f = 0;
      runs = kids(paraSeqO).map(tagOf).flatMap(t =>
        t === 'a:r' ? [runs[r++]] : t === 'a:fld' ? [flds[f++]] : t === 'a:br' ? [{ __br: true }] : [])
        .filter(Boolean);
    }
    // paragraphe fait d'un seul champ (a:fld sans a:r) : le champ est rendu comme un run (son a:t =
    // valeur mise en cache par PowerPoint ; slidenum est recalculé plus bas)
    if (!runs.length && p['a:fld']) runs.push(p['a:fld']);
    // line = HTML des runs ; firstColor / firstSz / firstFam = couleur, taille, police du 1er run
    // réel (servent à la puce, à l'interligne du <div> et au calcul de ligne de base)
    let line = '', firstColor = def.color || DEFCOL, firstSz = (def.sz || 18) * fontScale;
    // TABULATIONS de la 1re ligne, calculées comme PowerPoint : le texte va au prochain taquet après la
    // position courante — taquet personnalisé (a:tabLst, mesuré depuis le bord de la zone de texte), sinon
    // marL s'il est devant (retrait suspendu : « •<tab>Texte » tapé à la main), sinon multiple de defTabSz.
    // Le CSS n'a pas de taquet personnalisé -> tabulation remplacée par un blanc de la largeur exacte,
    // mesurée avec les chasses de PowerPoint (062_02 d.85 : texte 55 pt trop loin avec tab-size seul).
    // Toutes les grandeurs de tabulation sont en pt (u()). Un paragraphe à PUCE automatique ne suit
    // pas la position x (tabX = null) : la largeur de la puce n'est pas mesurée -> tabulations CSS.
    const bu0 = bulletOf(pPr) ?? def.bu;
    const autoBu0 = bu0 && !bu0.none && (bu0.char || bu0.num || bu0.blip);
    const marL0 = u((pPr['@marL'] != null ? +pPr['@marL'] : def.marL) || 0);
    const ind0 = Math.max(u(pPr['@indent'] != null ? +pPr['@indent'] : (def.ind || 0)), -marL0);
    // defTabSz : intervalle des taquets par défaut, 914400 EMU = 1" = 72 pt
    const defTab0 = u(+(pPr['@defTabSz'] || 914400)) || 72;
    const stops0 = arr(pPr['a:tabLst']?.['a:tab']).map(t => u(+t['@pos'] || 0)).filter(x => x > 0).sort((x, y) => x - y);
    // x courant (depuis le bord de la zone de texte) ; null = mesure impossible -> tabulations CSS
    let tabX = autoBu0 ? null : marL0 + ind0;
    // bord droit de la ligne (depuis le bord de la zone de texte) : permet de suivre les RETOURS À LA
    // LIGNE mot à mot (chasses et règle de coupure de PowerPoint) pour placer aussi les tabulations des
    // lignes suivantes. Une ligne suivante repart à marL.
    const marR0 = u((pPr['@marR'] != null ? +pPr['@marR'] : def.marR) || 0);
    const lineRight = u(w - iL - iR) - marR0;
    // nextStop(x) : 1er taquet personnalisé strictement après x ; sinon marL si x est avant ; sinon
    // le multiple suivant de defTabSz (1e-6 évite qu'un x pile sur un multiple reste sur place).
    // Les types de taquet (@algn l/ctr/r/dec) ne sont pas distingués : tous traités comme gauche.
    const nextStop = (x) => stops0.find(t => t > x + 0.01) ?? (x < marL0 - 0.01 ? marL0 : (Math.floor(x / defTab0 + 1e-6) + 1) * defTab0);
    // police du « strut » CSS de la ligne = celle du 1er run : sinon le <div> prend la police par défaut
    // du navigateur, dont les métriques décalent la boîte et grandissent chaque ligne (texte ré-autofité).
    let firstFam = fontFamily(resolveFont(def.font) || ctx.minorFont), firstSet = false, paraInk = false;
    // paragraphe VIDE : sa hauteur vient de la marque de fin (endParaRPr sz/latin), pas du style par
    // défaut — sinon chaque ligne vide 18 pt au lieu de 16 pt décale tout le texte qui suit.
    const endRPr = p['a:endParaRPr'];
    if (!runs.length && endRPr) {
      if (+endRPr['@sz']) firstSz = +endRPr['@sz'] / 100 * fontScale;
      const ef = resolveFont(endRPr['a:latin']?.['@typeface']);
      if (ef) firstFam = fontFamily(ef);
    }
    // paragraphe contenant des FORMULES OMML (a14:m / m:oMath) : rendre en ORDRE le texte (a:r)
    // et le math (fractions, etc.) intercalés — le parse plat perd l'ordre a:r ↔ math.
    const paraO = txBodyO && (p['a14:m'] !== undefined || p['m:oMathPara'] !== undefined || p['m:oMath'] !== undefined)
      ? byTag(txBodyO, 'a:p')[pi] : null;
    // Chemin SIMPLIFIÉ (paragraphe avec formules) : pas de tabulations, de décorations, de puce
    // calculée par run, ni d'advanceFix ; les runs n'ont pas de vertical-align (le jeton VA_TOK
    // n'y figure pas). La puce éventuelle est tout de même ajoutée plus bas.
    if (paraO) {
      const fam = fontFamily(resolveFont(def.font) || ctx.minorFont), pc = def.color || DEFCOL;
      for (const ch of kids(paraO)) {
        const ct = tagOf(ch);
        if (ct === 'a:r') {
          const rPrO = first(ch, 'a:rPr');
          const a = at(rPrO || {});
          const sz = ((a['@sz'] ? +a['@sz']/100 : def.sz) || 18) * fontScale;
          const b = (a['@b'] === '1' || (a['@b'] === undefined && def.b)) ? 'font-weight:700;' : '';
          const it = a['@i'] === '1' ? 'font-style:italic;' : '';
          // couleur PROPRE du run (sinon défaut paragraphe) — le chemin OMML l'ignorait, d'où du texte
          // en couleur de fond invisible (ex. encart rouge d.15 : run tx1=clair rendu en bg1=sombre).
          const sfO = rPrO && first(rPrO, 'a:solidFill');
          const rc = sfO ? cssColor(resolveColorFlat(PF.parse(`<r xmlns:a="a">${serialize(sfO)}</r>`).r['a:solidFill'], ctx)) : pc;
          const tN = first(ch, 'a:t'); const txt = esc(tN ? kids(tN).map(c => c['#text'] ?? '').join('') : '');
          if (txt.trim()) { hasText = true; firstColor = rc; }
          line += `<span style="font-size:${sz.toFixed(2)}px;color:${rc};font-family:${fam};${b}${it}">${txt}</span>`;
        } else if (ct === 'a:br') {
          line += '<br/>';
        } else if (ct === 'a14:m' || ct === 'm:oMathPara' || ct === 'm:oMath') {
          // formules OMML : PowerPoint les compose en Cambria Math (pas dans la police du paragraphe)
          line += `<span style="font-size:${((def.sz || 18) * fontScale).toFixed(2)}px;color:${pc};font-family:${fontFamily('Cambria Math')}">${renderMath(ch)}</span>`;
          hasText = true;
        }
      }
    } else
    // Chemin NORMAL : un <span> par run (ou par champ / saut de ligne), propriétés résolues run par run
    for (let ri = 0; ri < runs.length; ri++) {
      const r = runs[ri];
      if (r.__br) { line += '<br/>'; if (tabX != null) tabX = marL0; continue; } // nouvelle ligne : repart à marL
      const rPr = r['a:rPr'] || {};
      const sz = ((+rPr['@sz'] ? +rPr['@sz']/100 : def.sz) || 18) * fontScale; // points, autofit
      // couleur : seul a:solidFill est lu sur le run (gradFill de texte -> couleur par défaut)
      const col = rPr['a:solidFill'] ? cssColor(resolveColorFlat(rPr['a:solidFill'], ctx)) : (def.color || DEFCOL);
      // police : seul le script latin (a:latin) est lu ; a:ea / a:cs ignorés
      const font = resolveFont(rPr['a:latin']?.['@typeface']) || resolveFont(def.font) || ctx.minorFont;
      const b = (rPr['@b']==='1' || (rPr['@b']===undefined && def.b))?'font-weight:700;':'';
      const i = rPr['@i']==='1'?'font-style:italic;':'';
      // décorations cumulables : soulignement + barré (§21.1.2.3.6 u / strike)
      const deco = [];
      if (rPr['@u'] && rPr['@u']!=='none') deco.push('underline');
      if (rPr['@strike'] && rPr['@strike']!=='noStrike') deco.push('line-through');
      const td = deco.length ? `text-decoration:${deco.join(' ')};` : '';
      // cap : MAJUSCULES auto (all) / petites capitales (small) — §20.1.10.65 ST_TextCapsType
      const cap = rPr['@cap']==='all' ? 'text-transform:uppercase;'
                : rPr['@cap']==='small' ? 'font-variant:small-caps;' : '';
      // spc : espacement de caractères en centièmes de point (ST_TextPoint)
      // (1 px CSS = 1 pt ici, voir guide §5 ; l'autofit fontScale n'est pas appliqué à @spc)
      const spcPx = rPr['@spc'] != null && +rPr['@spc'] ? +rPr['@spc'] / 100 : 0;
      // baseline : exposant/indice (ST_Percentage, millièmes de %) -> décalage vertical
      // (en em = % de la taille du run : un % CSS se rapporterait à line-height, nul pour les runs lh0)
      // ex. baseline="30000" = exposant à +30 % ; "-25000" = indice. La taille du run n'est PAS
      // réduite ici : il garde son @sz.
      const blEm = rPr['@baseline'] && +rPr['@baseline'] ? +rPr['@baseline'] / 100000 : 0;
      // décalage vertical du run = exposant/indice (em) - recalage de ligne de base (px, VA_TOK remplacé
      // en fin de paragraphe par baselineShift, connu seulement après tous les runs)
      const bl = `vertical-align:calc(${blEm.toFixed(3)}em${VA_TOK};`;
      // highlight : surlignage du texte (§21.1.2.3.4) -> fond coloré derrière le run (comme Word)
      const hl = rPr['a:highlight'] ? `background-color:${cssColor(resolveColorFlat(rPr['a:highlight'], ctx))};` : '';
      // contour de texte (rPr>a:ln, type WordArt/titre stylé §21.1.2.1.11) -> -webkit-text-stroke.
      // Additif : présent uniquement si le run déclare un a:ln avec remplissage -> aucun autre run touché.
      const tln = rPr['a:ln'];
      let ol = '';
      if (tln && tln['a:noFill'] === undefined && (tln['a:solidFill'] || tln['a:gradFill'])) {
        const oc = tln['a:solidFill'] ? cssColor(resolveColorFlat(tln['a:solidFill'], ctx)) : col;
        const ow = Math.max(0.3, u(+tln['@w'] || 9525) / 2); // EMU->px, moitié (le trait déborde des 2 côtés)
        ol = `-webkit-text-stroke:${ow.toFixed(2)}px ${oc};`;
      }
      // champ slidenum : afficher le VRAI numéro de diapo (sinon le texte-invite « ‹N°› »
      // du layout/master fuit — c'est une zone libre non-placeholder, donc non sautée).
      let raw = r['@type'] === 'slidenum' ? String(slideNum) : (r['a:t'] ?? '');
      // espaces de FIN de paragraphe (dernier run) : non rendus par PowerPoint → couper (sinon
      // ils élargissent la boîte et déclenchent une sur-réduction de l'autofit). Espaces INTERNES
      // multiples conservés (pre-wrap) pour le positionnement type « TURBULENCE   vu à ».
      if (ri === runs.length - 1) raw = raw.replace(/[ \t]+$/, '');
      // run en police SYMBOLE (Wingdings/Webdings/Symbol, souvent via <a:sym>) : glyphes remappés vers
      // Unicode (anti-tofu Linux/LMS) dans SYM_FAMILY ; le texte ordinaire du run garde sa police.
      const parts = symbolParts(raw, font, rPr['a:sym']?.['@typeface']);
      const allSym = parts?.every(p => p.sym);
      // txt = HTML du run : texte échappé ; si remappage partiel, les glyphes symboles sont dans des
      // <span> en SYM_FAMILY ; si remappage total, c'est le run entier qui prend SYM_FAMILY (famCss)
      let txt = !parts ? esc(raw) : allSym ? esc(parts.map(p => p.text).join(''))
        : parts.map(p => p.sym ? `<span style="font-family:${SYM_FAMILY}">${esc(p.text)}</span>` : esc(p.text)).join('');
      const famCss = allSym ? SYM_FAMILY : fontFamily(font);
      // Suivi de la position x (tabX) et remplacement des tabulations par des blancs de largeur exacte.
      // tabX est mis à jour pour TOUT run mesurable (même sans tabulation), car la position d'une
      // tabulation dépend de tout ce qui la précède sur la ligne ; il simule aussi les retours à la
      // ligne automatiques (même règle de coupure que cellTextHeight). Dès qu'une mesure échoue (glyphe
      // inconnu, run symbole), tabX passe à null pour le reste du paragraphe -> tabulations laissées au
      // CSS (tab-size).
      if (tabX != null && !parts) {
        // découpe aux tabulations ; chaque segment avance tabX de sa largeur mesurée
        const segs = raw.split('	'), out = [];
        for (let k = 0; k < segs.length && tabX != null; k++) {
          // avance mot par mot (espace finale comprise) ; un mot qui ne tient plus ouvre une ligne à marL
          let okSeg = true;
          for (const word of segs[k].split(/(?<= )/)) {
            if (!word) continue;
            const ww = pptTextWidth(famCss, !!b, !!i, sz, word);
            const wn = word.endsWith(' ') ? pptTextWidth(famCss, !!b, !!i, sz, word.trimEnd()) : ww;
            if (ww == null || wn == null) { okSeg = false; break; }
            // retour à la ligne simulé : le mot (sans son espace) dépasse le bord droit -> il ouvre une
            // nouvelle ligne à marL (sauf en wrap="none", et jamais pour le 1er mot d'une ligne)
            if (tabX > marL0 + 0.01 && tabX + wn > lineRight - 0.02 && ws !== 'pre') tabX = marL0 + ww;
            else tabX += ww;
          }
          if (!okSeg) { tabX = null; break; }
          out.push(esc(segs[k]));
          if (k < segs.length - 1) {
            // tabulation -> <span inline-block> vide de la largeur jusqu'au prochain taquet
            const gap = nextStop(tabX) - tabX;
            out.push(`<span style="display:inline-block;width:${gap.toFixed(2)}px"></span>`); tabX += gap;
          }
        }
        // txt n'est remplacé que s'il y avait au moins une tabulation ET que toutes les mesures ont
        // réussi ; sinon le texte brut (avec son caractère TAB, rendu par tab-size CSS) est gardé
        if (tabX != null && segs.length > 1) txt = out.join('');
      } else if (parts) tabX = null;
      if (!firstSet) { firstSz = sz; firstFam = famCss; firstSet = true; } // 1er run réel -> interligne correct du <div>
      if (txt.trim()) { hasText = true; firstColor = col; }
      // espacement : spc du run + correction d'arrondi des chasses façon PowerPoint (advanceFix)
      const ls = spcPx + (allSym ? 0 : advanceFix(famCss, !!b, !!i, sz, rPr['@cap'] === 'all' ? raw.toUpperCase() : raw));
      const spc = Math.abs(ls) >= 0.001 ? `letter-spacing:${ls.toFixed(3)}px;` : '';
      // line-height du paragraphe (paraLH) pour un run pas plus grand que le 1er : sa boîte garde la
      // hauteur d'UNE ligne sans jamais dépasser celle du paragraphe. En CSS, des polices mêlées (Poppins
      // + symboles en Times) s'alignaient sur la ligne de base et leur union dépassait L (34 px au lieu de
      // 24) ; un line-height UNITAIRE (paraLH) donne 1,2 x la taille de CHAQUE run indépendamment de ses
      // métriques -> plus de dépassement. Un run PLUS grand agrandit la ligne, comme chez PowerPoint.
      // (Avant : « line-height:0 » ; il annulait la boîte de ligne et faisait s'effondrer le <div> à 0 ->
      // lignes retournées à la ligne ET paragraphes superposés au même y.)
      const lh0 = sz <= firstSz + 0.01 ? `line-height:${paraLH};` : '';
      if (raw.trim()) paraInk = true; // paragraphe « encré » (voir le &#8203; en fin de <div>)
      // Ordre des déclarations CSS du run : taille, line-height 0, couleur, police, gras, italique,
      // décoration, casse, letter-spacing, vertical-align (avec jeton VA_TOK), surlignage, contour.
      line += `<span style="font-size:${sz.toFixed(2)}px;${lh0}color:${col};font-family:${famCss};${b}${i}${td}${cap}${spc}${bl}${hl}${ol}">${txt}</span>`;
    }
    // ── Puce, retraits, espacements, émission du <div> du paragraphe ──
    // Une puce n'est émise que si le paragraphe a du texte visible (PowerPoint n'affiche pas la puce
    // d'un paragraphe vide). Taille = taille du 1er run × buSzPct ; couleur = buClr ou 1er run.
    // puce : pPr du paragraphe prioritaire, sinon héritée du niveau de style
    const bu = bulletOf(pPr) ?? def.bu;
    let bullet = '', bulletCore = '';
    if (bu && !bu.none && bu.blip && (line || '').trim()) {
      // puce IMAGE (a:buBlip) : petite icône alignée sur le texte
      const bid = bu.blip['@r:embed'] || bu.blip['@embed'];
      const bhref = bid && curRels[bid] ? imageHref(pptDir, curRels[bid], embed) : null;
      if (bhref) { const bsz = (bu.szPct ? firstSz * bu.szPct : firstSz); bullet = `<img src="${bhref}" style="height:${bsz.toFixed(1)}px;vertical-align:middle"/> `; }
    } else if (bu && !bu.none && (bu.char || bu.num) && (line || '').trim()) {
      let sym = bu.char, buFam;
      if (bu.num) { counters[lvl] = (counters[lvl] || 0) + 1; sym = counters[lvl] + (/paren/i.test(bu.num) ? ')' : '.'); buFam = bu.font ? fontFamily(resolveFont(bu.font)) : firstFam; }
      else {
        // puce caractère : mapper les puces en police SYMBOLE (Wingdings) vers Unicode (anti-tofu)
        const bs = bulletSym(bu.char, bu.font);
        sym = bs.sym;
        buFam = bs.symbolic ? "'Arial','Liberation Sans','DejaVu Sans',sans-serif" : bu.font ? fontFamily(resolveFont(bu.font)) : firstFam; // sans buFont (buFontTx) = police du texte
      }
      // couleur PROPRE de la puce (a:buClr) si présente, sinon couleur du 1er run ; taille via a:buSzPct
      const buCol = bu.clr ? cssColor(resolveColorFlat(bu.clr, ctx)) : firstColor;
      const buSz = bu.szPct ? firstSz * bu.szPct : firstSz;
      // la puce porte aussi le jeton VA_TOK : elle suit le même recalage de ligne de base que le texte
      bulletCore = `<span style="font-size:${buSz.toFixed(2)}px;line-height:0;vertical-align:calc(0em${VA_TOK};color:${buCol};font-family:${buFam}">${esc(sym)}</span>`;
    }
    // retrait : pPr du paragraphe prioritaire (marL/indent équilibrés du niveau de liste)
    const marLe = pPr['@marL'] != null ? +pPr['@marL'] : def.marL;
    const inde = pPr['@indent'] != null ? +pPr['@indent'] : def.ind;
    const marL = marLe != null ? u(marLe) : 0;
    // le retrait négatif (puce suspendue) ne doit jamais dépasser la marge -> pas de débordement gauche
    const indRaw = inde != null ? u(inde) : 0;
    const ind = Math.max(indRaw, -marL);
    // largeur de la puce suspendue = |indent| même quand marL + indent < 0 est borné à 0 : PowerPoint
    // pose la puce au bord et le texte |indent| plus loin (indent -36 pt sans marL : texte à 36 pt,
    // lignes suivantes au bord — vérifié sur l'export, 022_04 d.96).
    const hang = indRaw < 0 ? -indRaw : 0;
    // marge DROITE du paragraphe (@marR, §21.1.2.2.7) : ignorée avant -> texte coupé plus large que
    // dans PowerPoint (062_02 d.89 : marR 39,85 pt, lignes trop longues d'un ou deux mots).
    const marRe = pPr['@marR'] != null ? +pPr['@marR'] : def.marR;
    const marR = marRe ? u(marRe) : 0;
    // Modèle PowerPoint -> CSS : marL (marge gauche de TOUTES les lignes) = padding-left ; indent
    // (décalage de la 1re ligne, négatif pour une puce suspendue) = text-indent.
    const pad = ((marL || ind) ? `padding-left:${marL.toFixed(1)}px;text-indent:${ind.toFixed(1)}px;` : '')
      + (marR ? `padding-right:${marR.toFixed(1)}px;` : '');
    // puce SUSPENDUE (indent < 0) : elle occupe tout le retrait et le texte démarre à marL, comme le
    // taquet de PowerPoint (sinon « - The… » collé au tiret au lieu d'être aligné sous le texte).
    if (bulletCore) bullet = hang > 0
      ? `<span style="display:inline-block;min-width:${hang.toFixed(1)}px;text-indent:0">${bulletCore}</span>`
      : bulletCore + '\u00a0'; // puce non suspendue : suivie d'une espace insécable (NBSP)
    // interligne (lnSpc) + espacement avant/après (spcBef/spcAft) — §21.1.2.2.{5,7,12}
    // spcPct = millièmes de % (100000 = 100 %) ; spcPts = centièmes de point.
    // lnSpc du paragraphe, sinon hérité (lstStyle/layout/master via def.lnSpc). spcPct 100 % =
    // interligne SIMPLE ≈ 1.2× la police (pas 1.0) → facteur 1.2 pour coller à PowerPoint.
    // normAutofit @lnSpcReduction : PowerPoint réduit AUSSI l'interligne explicite du paragraphe
    // (pas seulement les boîtes sans lnSpc via le conteneur). Sans ça, un paragraphe à lnSpc 120 %
    // dans une boîte autofit à 20 % de réduction rend 1.44 au lieu de 0.96×1.2=1.152 → texte trop
    // haut → déborde → l'autofit navigateur sur-rétrécit → mauvais nb de lignes (ex. question d.6).
    const ls = pPr['a:lnSpc'] || def.lnSpc;
    const lnMul = 1 - lnRed; // 1 si pas de réduction -> aucun changement pour ces boîtes-là
    const lnSpcCss = ls?.['a:spcPct']?.['@val'] ? `line-height:${(+ls['a:spcPct']['@val']/100000*1.2*lnMul).toFixed(3)};`
                   : ls?.['a:spcPts']?.['@val'] ? `line-height:${(+ls['a:spcPts']['@val']/100*lnMul).toFixed(1)}px;` : '';
    // hauteur de ligne effective (px) -> décalage de ligne de base façon PowerPoint (baselineShift)
    const Lpx = ls?.['a:spcPct']?.['@val'] ? +ls['a:spcPct']['@val'] / 100000 * 1.2 * lnMul * firstSz
              : ls?.['a:spcPts']?.['@val'] ? +ls['a:spcPts']['@val'] / 100 * lnMul
              : lineH * firstSz;
    // un seul décalage par paragraphe, calculé sur le 1er run (police + taille) : les runs d'autres
    // tailles reçoivent le même décalage en px
    const blShift = baselineShift(firstFam, firstSz, Lpx);
    // margin-right 0,2 px : PowerPoint renvoie un mot à la ligne dès que la ligne (sans son espace finale)
    // atteint la largeur disponible, à ~0,01 pt près (mesuré via COM : 188,38 pt renvoyé pour 188,391
    // disponibles) ; Chrome le gardait (0,06 px ne suffisait pas : il arrondit ses positions au 1/64 px).
    // Borne haute : une ligne à 1,2 px sous la limite reste gardée chez PowerPoint (022_04 d.18).
    // taquets par défaut (defTabSz, 72 pt) : sans ça, tab = 8 espaces CSS -> « ⇒ » moins en retrait
    // (def.defTab n'est jamais renseigné par lvlStyle : en pratique pPr@defTabSz ou 72 pt)
    const tabPx = u(+(pPr['@defTabSz'] || def.defTab || 914400));
    // spcMk(nœud spcBef/spcAft, 'top'|'bottom') -> marge CSS : spcPts en 1/100 pt -> px ;
    // spcPct en 1/1000 de % converti en em (donc relatif à la taille du texte du <div>).
    const spcMk = (node, side) => !node ? ''
      : node['a:spcPts']?.['@val'] ? `margin-${side}:${(+node['a:spcPts']['@val']/100).toFixed(1)}px;`
      : node['a:spcPct']?.['@val'] ? `margin-${side}:${(+node['a:spcPct']['@val']/100000).toFixed(2)}em;` : '';
    // spcBef/spcAft : pPr de la diapo prioritaire, sinon HÉRITÉ (lstStyle/layout/master via def).
    // spcBef supprimé sur le 1er paragraphe : PowerPoint ne pousse pas le bloc vers le bas en haut
    // de boîte (l'espace-avant ne vaut qu'ENTRE paragraphes) — sinon décalage vertical du bloc entier.
    const spcBefNode = pi === 0 ? null : (pPr['a:spcBef'] ?? def.spcBef);
    // de même, l'espace-APRÈS du dernier paragraphe ne compte pas (calibré via COM : spcAft 12 pt sur le
    // dernier paragraphe = aucun effet en ancrage haut/milieu/bas) ; l'appliquer remontait de moitié
    // le texte centré des formes (040_02_01 d.32 : titres de cartes 3 pt trop haut).
    const spcAftNode = pi === paras.length - 1 ? null : (pPr['a:spcAft'] ?? def.spcAft);
    // interligne > 100 % : PowerPoint ajoute l'espace en plus AU-DESSUS de chaque ligne ; sous la dernière
    // ligne du bloc ne reste que la descente naturelle min(0,3 x taille, descendante) au lieu de 0,25 L
    // -> bloc plus court d'autant (calibré ; compte pour l'ancrage milieu/bas).
    const fmD = (metricsOf(firstFam) || [0, 0.25])[1];
    const tailCut = pi === paras.length - 1 && Lpx > 1.2 * firstSz + 0.01
      ? Math.max(0, 0.25 * Lpx - Math.min(0.3 * firstSz, fmD * firstSz)) : 0;
    // tailCut s'exprime comme une marge basse NÉGATIVE (combinée à un éventuel spcAft existant)
    let aftCss = spcMk(spcAftNode, 'bottom');
    if (tailCut > 0.05) {
      const aft = aftCss.match(/margin-bottom:([^;]+);/);
      aftCss = `margin-bottom:calc(${aft ? aft[1] : '0px'} - ${tailCut.toFixed(2)}px);`;
    }
    const spcCss = spcMk(spcBefNode, 'top') + aftCss;
    // data-par : index de paragraphe (pour le build par paragraphe bldP dans le player)
    // white-space:pre-wrap : préserve les espaces MULTIPLES (PowerPoint s'en sert pour positionner
    // du texte, ex. « TURBULENCE      vu à » où le « vu à » tombe entre deux badges) ; wrap conservé.
    // font-size sur le div = taille du texte -> l'interligne (unitless 1.15) se calcule dessus,
    // sinon il se base sur 16px (défaut) et gonfle la ligne (pied de page 9pt coupé dans sa boîte).
    // (L'interligne du conteneur vaut aujourd'hui lineH = 1,2, cf. plus haut.)
    // En fin de <div> : applyShift remplace le jeton VA_TOK de tous les runs + puce par blShift ;
    // &#8203; (espace de largeur nulle) donne une hauteur de ligne à un paragraphe sans encre
    // (vide ou seulement des espaces), sinon le <div> s'effondrerait à 0.
    html += `<div data-par="${pi}" style="font-size:${firstSz.toFixed(1)}px;font-family:${firstFam};white-space:${ws};${ws === 'pre' ? '' : 'overflow-wrap:break-word;'}tab-size:${tabPx.toFixed(1)}px;text-align:${algn};margin:0;${ws === 'pre' ? '' : 'margin-right:0.2px;'}${pad}${lnSpcCss}${spcCss}">${applyShift(bullet + line, blShift)}${paraInk ? '' : '&#8203;'}</div>`;
  }
  if (!hasText) return '';
  // data-af : mode d'ajustement de la boîte (norm = normAutofit, fixed sinon), repère pour le player.
  // Aucune réduction côté navigateur : le fontScale / lnSpcReduction que PowerPoint a ENREGISTRÉS sont
  // déjà appliqués ci-dessus, et il ne rétrécit jamais au rendu (slideFrame.js).
  const afMode = (naf !== undefined) ? 'norm' : 'fixed';
  // <a:noAutofit/> = « ne pas ajuster » : PowerPoint NE clippe PAS, le texte peut déborder la forme
  // (ex. libellés/pastilles : police 40pt dans une pastille de 40pt -> les capitales tiennent, mais
  // l'aire réduite par les insets clipperait). On étend alors le foreignObject à la boîte PLEINE, on
  // met les insets en padding et on laisse déborder (overflow visible).
  // Deux variantes de sortie :
  //   noAutofit : foreignObject = boîte ENTIÈRE, insets en padding (box-sizing:border-box) ;
  //   sinon     : foreignObject = boîte MOINS les insets (x = iL, y = iT).
  // Dans les deux cas, <div class="af"> est un flex colonne : justify-content = ancrage vertical,
  // align-items:center si anchorCtr.
  const noaf = bodyPr['a:noAutofit'] !== undefined;
  if (noaf) {
    return `<foreignObject x="0" y="0" width="${u(w)}" height="${u(h)}" style="overflow:visible">`
      + `<div xmlns="http://www.w3.org/1999/xhtml" class="af" data-af="fixed"${wsAttr} style="box-sizing:border-box;width:100%;height:100%;padding:${u(iT)}px ${u(iR)}px ${u(iB)}px ${u(iL)}px;display:flex;flex-direction:column;justify-content:${just};${ctrCss}line-height:${lineH.toFixed(3)};overflow:visible;">${html}</div>`
      + `</foreignObject>`;
  }
  // overflow VISIBLE : PowerPoint ne clippe pas le texte d'une forme au bord de la boîte. Avec
  // overflow:hidden, une ligne un peu plus haute que sa boîte (ex. chiffre 24pt dans une pastille dont
  // les insets réduisent la boîte à 25.7pt) était ROGNÉE en haut/bas. Le texte peut déborder (comme
  // PowerPoint).
  return `<foreignObject x="${u(iL)}" y="${u(iT)}" width="${u(w-iL-iR)}" height="${u(h-iT-iB)}" style="overflow:visible">`
    + `<div xmlns="http://www.w3.org/1999/xhtml" class="af" data-af="${afMode}"${wsAttr} style="width:100%;height:100%;display:flex;flex-direction:column;justify-content:${just};${ctrCss}line-height:${lineH.toFixed(3)};overflow:visible;">${html}</div>`
    + `</foreignObject>`;
}
