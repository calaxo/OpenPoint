// ═══ [PPTX · STYLES DE TEXTE HÉRITÉS] ═══ puces (bulletOf), p:txStyles du master, p:defaultTextStyle.
// Ces styles forment le BAS de la cascade de texte (voir slide/text.mjs) ; ils sont posés sur ctx par
// renderSlide avant le parcours des formes.
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { cachedParse } from './xml.mjs';
import { resolveColorFlat, cssColor } from './color.mjs';

// ████ §E ─ PUCES ET POLICES -> FAMILLES CSS ────────────────────────────────────────────────────
// bulletOf lit la définition de puce d'un niveau de paragraphe (a:pPr / a:lvlNpPr, §21.1.2.4.x :
// buNone, buChar, buAutoNum, buBlip, buClr, buSzPct, buFont). fontFamily traduit un typeface
// OOXML en pile CSS avec substitut libre métrique-compatible (pour que la COUPURE DE LIGNE soit la
// même que sous Windows même sans la police d'origine).
//
// puce d'un pPr (flat) : {none} | {char,font} | {num,font} | null (non défini -> hériter)
// Le null est important : il laisse la chaîne d'héritage (lstStyle -> layout -> master) décider,
// alors que {none:true} COUPE l'héritage (a:buNone explicite).
// Non lus : a:buSzPts, a:buClrTx / a:buSzTx / a:buFontTx (qui signifient « comme le texte », ce
// qui est déjà le comportement par défaut ici) ; @startAt de buAutoNum (numérotation depuis 1).
function bulletOf(pPr) {
  if (!pPr) return null;
  if (pPr['a:buNone'] !== undefined) return { none: true };
  // a:buClr = couleur PROPRE de la puce (souvent une couleur d'accent différente du texte). Nœud brut
  // conservé ; résolu au rendu (a besoin du thème/ctx). a:buSzPct/@val = taille de puce en % de la police.
  const clr = pPr['a:buClr'] || null;
  const szPct = pPr['a:buSzPct']?.['@val'] ? +pPr['a:buSzPct']['@val'] / 100000 : null;
  // a:buBlip = puce IMAGE (petit logo/icône). Nœud brut conservé ; résolu au rendu (curRels/pptDir).
  if (pPr['a:buBlip']) return { blip: pPr['a:buBlip']['a:blip'], szPct };
  if (pPr['a:buChar']) return { char: pPr['a:buChar']['@char'], font: pPr['a:buFont']?.['@typeface'], clr, szPct };
  // numérotation auto : seul le suffixe est interprété au rendu (« ) » si le type contient « paren »,
  // sinon « . ») ; le compteur est toujours arabe (1, 2, 3…), même pour alphaLcPeriod / romanUcPeriod.
  if (pPr['a:buAutoNum']) return { num: pPr['a:buAutoNum']['@type'] || 'arabicPeriod', font: pPr['a:buFont']?.['@typeface'], clr, szPct };
  return null;
}

// loadMasterStyles(ctx, masterPath) : pose ctx.clrMap et ctx.txStyles depuis le master de la diapo.
export function loadMasterStyles(ctx, masterPath) {
  // clrMap du masque : remappe bg1/tx1/bg2/tx2 vers dk1/lt1/... (crucial pour thème sombre)
  // + txStyles : couleurs/tailles par défaut des placeholders (titre/corps) hérités par les runs
  // Structure produite : ctx.txStyles[title|body|other][niveau 1..9] = { sz (pt), b, color (CSS déjà
  // résolue — donc avec le clrMap du master), font (typeface brut, éventuellement « +mj-lt »), algn,
  // bu, marL, ind (EMU), lnSpc, spcBef, spcAft (nœuds plats bruts) } — lu depuis p:sldMaster/p:txStyles.
  ctx.txStyles = {};
  if (masterPath && existsSync(masterPath)) {
    const mflat = cachedParse(masterPath, 'flat');
    const cm = mflat['p:sldMaster']?.['p:clrMap'];
    // p:clrMap bg1="lt1" tx1="dk1" … -> { bg1:'lt1', tx1:'dk1', … }. NB : un p:clrMapOvr du layout
    // ou de la diapo (surcharge locale) n'est pas lu : c'est toujours la table du master qui s'applique.
    if (cm) { ctx.clrMap = {}; for (const [k, v] of Object.entries(cm)) if (k.startsWith('@')) ctx.clrMap[k.slice(1)] = v; }
    const ts = mflat['p:sldMaster']?.['p:txStyles'];
    if (ts) for (const [k, tag] of Object.entries({ title:'p:titleStyle', body:'p:bodyStyle', other:'p:otherStyle' })) {
      const st = ts[tag]; if (!st) continue; ctx.txStyles[k] = {};
      for (let lvl = 1; lvl <= 9; lvl++) {
        const pPr = st[`a:lvl${lvl}pPr`]; if (!pPr) continue;
        const d = pPr['a:defRPr'] || {};
        ctx.txStyles[k][lvl] = {
          sz: d['@sz'] ? +d['@sz']/100 : null, b: d['@b'] === '1',
          color: d['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null,
          font: d['a:latin']?.['@typeface'] || null,
          algn: pPr['@algn'] || null,
          bu: bulletOf(pPr), marL: pPr['@marL'] ? +pPr['@marL'] : null, ind: pPr['@indent'] ? +pPr['@indent'] : null,
          lnSpc: pPr['a:lnSpc'] || null,
          // espacement avant/après paragraphe hérité du master (§21.1.2.2.{10,12}) : les cartes/listes
          // du corpus tirent leur « gap entre paragraphes » d'ici (bodyStyle lvl1 spcBef 10pt), jamais
          // redéfini dans la diapo -> sans héritage, les paragraphes se collent (vu vs PowerPoint).
          spcBef: pPr['a:spcBef'] || null, spcAft: pPr['a:spcAft'] || null,
        };
      }
    }
  }
}

// loadDefaultTextStyle(ctx, pptDir) : pose ctx.defaultTextStyle depuis presentation.xml.
export function loadDefaultTextStyle(ctx, pptDir) {
  // defaultTextStyle de presentation.xml : dernier recours pour le texte hors placeholder
  // (zones de texte libres) — §19.2.1.6. Chargé par niveau, même structure que txStyles.
  ctx.defaultTextStyle = {};
  try {
    const pf = cachedParse(join(pptDir, 'presentation.xml'), 'flat');
    const dts = pf['p:presentation']?.['p:defaultTextStyle'];
    if (dts) for (let lvl = 1; lvl <= 9; lvl++) {
      const pPr = dts[`a:lvl${lvl}pPr`]; if (!pPr) continue;
      const d = pPr['a:defRPr'] || {};
      ctx.defaultTextStyle[lvl] = {
        sz: d['@sz'] ? +d['@sz']/100 : null, b: d['@b'] === '1',
        color: d['a:solidFill'] ? cssColor(resolveColorFlat(d['a:solidFill'], ctx)) : null,
        font: d['a:latin']?.['@typeface'] || null, algn: pPr['@algn'] || null,
      };
    }
  } catch {}
}

export { bulletOf };
