// ═══ [PPTX · PARTS OPC] ═══ relations (.rels), images (data URI), placeholders hérités du layout/master.
import { readFileSync, existsSync, readdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { cachedParse } from './xml.mjs';
// ████ §M ─ RELATIONS (.rels) ET IMAGES ─────────────────────────────────────────────────────────
// Chaque part OOXML référence les autres (layout, master, thème, images, graphiques…) par un
// identifiant r:id / r:embed (« rId3 ») résolu dans son fichier _rels/<part>.rels (OPC, ECMA-376
// Part 2). Les cibles sont RELATIVES au dossier de la part (souvent « ../media/image1.png »).
// Ces fonctions renvoient une simple table { rId: Target } (les @Type ne sont pas lus : on reconnaît
// la cible à son chemin, ex. includes('slideLayout')). Les relations externes (TargetMode="External")
// ne sont pas distinguées.
// ================= rels & images =================
// slideRels(pptDir, n) -> rels de slides/slideN.xml ({} si pas de .rels)
function slideRels(pptDir, slideNo) {
  const p = join(pptDir, 'slides', '_rels', `slide${slideNo}.xml.rels`);
  if (!existsSync(p)) return {};
  const r = cachedParse(p, 'flat');
  const out = {};
  const rl = r.Relationships?.Relationship;
  for (const x of (Array.isArray(rl) ? rl : rl ? [rl] : [])) out[x['@Id']] = x['@Target'];
  return out;
}
// relsOf(chemin absolu d'une part) -> ses rels, par la convention OPC <dossier>/_rels/<nom>.rels
function relsOf(path) {
  const d = dirname(path), b = path.split(/[\\/]/).pop();
  const p = join(d, '_rels', b + '.rels');
  if (!existsSync(p)) return {};
  const r = cachedParse(p, 'flat');
  const out = {}; const rl = r.Relationships?.Relationship;
  for (const x of (Array.isArray(rl) ? rl : rl ? [rl] : [])) out[x['@Id']] = x['@Target'];
  return out;
}
// Extensions servables dans <image href="data:…">. Valeur '' (emf/wmf) ou extension absente (tif,
// wdp…) -> non embarquable -> imageHref renvoie null (image non dessinée).
const MIME = { webp:'image/webp', png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif', bmp:'image/bmp', svg:'image/svg+xml', emf:'', wmf:'' };
// imageHref(pptDir, Target de rels, embed) -> URL de l'image ou null.
//   embed=true (défaut de renderSlide) : data URI base64 (SVG autonome, lourd) ;
//   embed=false : URL file:/// absolue (aperçus locaux uniquement).
// La cible est résolue depuis ppt/ en retirant UN « ../ » de tête : correct pour les cibles des
// slides / layouts / masters (qui sont à un niveau sous ppt/), pas pour une cible sans « ../ »
// relative à un autre dossier.
function imageHref(pptDir, target, embed) {
  let abs = join(pptDir, target.replace(/^\.\.\//,''));
  if (!existsSync(abs)) return null;
  let ext = abs.split('.').pop().toLowerCase();
  // EMF/WMF (vecteur Windows, non affichable en <image>) : servir le PNG pré-converti « <fichier>.png »
  // (préprocessing convert-metafiles.ps1 via .NET/WIC). Beaucoup de diagrammes techniques sont en EMF.
  if (ext === 'emf' || ext === 'wmf') {
    const png = abs + '.png';
    if (existsSync(png)) { abs = png; ext = 'png'; } else return null;
  }
  if (!embed) return 'file:///' + abs.replace(/\\/g,'/');
  // version WebP allégée écrite à côté par mediaOptimize.js (si elle valait le coup)
  if (existsSync(abs + '.webp')) { abs += '.webp'; ext = 'webp'; }
  const mime = MIME[ext];
  if (!mime) return null;
  return `data:${mime};base64,` + readFileSync(abs).toString('base64');
}

// ████ §N ─ PLACEHOLDERS (HÉRITAGE DEPUIS LE LAYOUT) ────────────────────────────────────────────
// Règle PresentationML (p:ph, §19.3.1.36) : une forme <p:sp><p:nvSpPr><p:nvPr><p:ph type idx/> de la
// diapo est l'INSTANCE d'un placeholder du layout (lui-même instance de celui du master). Tout ce que
// la diapo ne redéfinit pas (position/taille, géométrie, fill, ligne, styles de texte, bodyPr) vient
// du niveau au-dessus. Ici, l'appariement se fait par clé exacte "type:idx" (@type absent -> 'body',
// @idx absent -> '0', des DEUX côtés, donc cohérent). NB : la spec donne 'obj' comme type par défaut
// et PowerPoint apparie surtout par idx ; la clé exacte suffit pour le corpus.
// ================= placeholders (héritage layout/master) =================
// Une forme-texte de la diapo hérite géométrie + fill + styles de texte de SON placeholder
// de layout (identifié par type:idx). Le layout hérite lui-même du masque.
// placeholderGeom(pptDir, n, rels de la diapo) -> { geom, mgeom } : geom = { "type:idx": { xfrm, spPr,
// style, lstStyle, bodyPr } } lu du LAYOUT ; mgeom = { "type:idx": { lstStyle, bodyPr } } lu du MASTER
// (STYLE DE TEXTE seulement, pour la cascade : c'est le master qui porte le plus souvent la police/la
// taille du titre et du corps). Tout en arbre PLAT, mis en cache. Seuls les p:sp de 1er niveau du spTree
// sont indexés (pas les p:pic ni les placeholders dans des groupes). La GÉOMÉTRIE n'est pas reprise du
// master : un placeholder de layout sans xfrm ne reçoit pas celui du master.
function placeholderGeom(pptDir, slideNo, rels) {
  const layoutTgt = Object.values(rels).find((t) => t && t.includes('slideLayout'));
  const geom = {};
  if (!layoutTgt) return { geom };
  const layoutPath = join(pptDir, 'slides', layoutTgt);
  if (!existsSync(layoutPath)) return { geom };
  const lx = cachedParse(layoutPath, 'flat');
  const tree = lx['p:sldLayout']?.['p:cSld']?.['p:spTree'];
  const arr = (x)=> x===undefined?[]:Array.isArray(x)?x:[x];
  for (const sp of arr(tree?.['p:sp'])) {
    const ph = sp['p:nvSpPr']?.['p:nvPr']?.['p:ph']; if (!ph) continue;
    const key = `${ph['@type']??'body'}:${ph['@idx']??'0'}`;
    // le layout a plusieurs placeholders "body:0" -> on garde le premier par clé unique,
    // mais les cartes ont des idx distincts (27,30,33...) donc pas de collision gênante
    if (!geom[key]) geom[key] = { xfrm: sp['p:spPr']?.['a:xfrm'], spPr: sp['p:spPr'], style: sp['p:style'],
      lstStyle: sp['p:txBody']?.['a:lstStyle'], bodyPr: sp['p:txBody']?.['a:bodyPr'] };
  }
  // MAILLON MANQUANT de la cascade : les placeholders du MASTER (même clé « type:idx »). C'est le master
  // qui porte le plus souvent la POLICE et la TAILLE du titre / du corps (le layout ne les redéfinit pas).
  // Sans ce maillon, un titre héritait de p:txStyles (titleStyle = défaut du fichier) : police et taille
  // fausses (vu : titre « Arial 14 » au lieu du « Space Grotesk 30 » du master). Le master est joint au
  // layout (« ../slideMasters/slideMasterN.xml »). On n'indexe QUE le style de texte (lstStyle), pas la
  // géométrie (cf. commentaire ci-dessus).
  const mgeom = {};
  const mTgt = Object.values(relsOf(layoutPath)).find((t) => t && t.includes('slideMaster'));
  if (mTgt) {
    const mPath = join(dirname(layoutPath), mTgt);
    if (existsSync(mPath)) {
      const mt = cachedParse(mPath, 'flat')['p:sldMaster']?.['p:cSld']?.['p:spTree'];
      for (const sp of arr(mt?.['p:sp'])) {
        const ph = sp['p:nvSpPr']?.['p:nvPr']?.['p:ph']; if (!ph) continue;
        const key = `${ph['@type']??'body'}:${ph['@idx']??'0'}`;
        if (!mgeom[key]) mgeom[key] = { lstStyle: sp['p:txBody']?.['a:lstStyle'], bodyPr: sp['p:txBody']?.['a:bodyPr'] };
      }
    }
  }
  return { geom, mgeom };
}
// slideOrder(pptDir) -> numéros N des slides/slideN.xml dans l'ORDRE DE PRÉSENTATION : celui de
// p:presentation/p:sldIdLst (chaque p:sldId@r:id renvoie, via presentation.xml.rels, à une part de
// diapo). Le numéro de FICHIER n'est pas l'ordre d'affichage : une diapo déplacée dans PowerPoint garde
// son fichier. Les diapos masquées (p:sld@show="0") restent listées : c'est au viewer de décider.
// Repli (presentation.xml absent ou illisible) : tous les slideN.xml triés par N.
export function slideOrder(pptDir) {
  const dir = join(pptDir, 'slides');
  const onDisk = existsSync(dir)
    ? readdirSync(dir).filter((f) => /^slide\d+\.xml$/.test(f)).map((f) => +f.match(/\d+/)[0]).sort((a, b) => a - b)
    : [];
  try {
    const pres = cachedParse(join(pptDir, 'presentation.xml'), 'flat');
    const ids = pres['p:presentation']?.['p:sldIdLst']?.['p:sldId'];
    const rels = relsOf(join(pptDir, 'presentation.xml'));
    const order = (Array.isArray(ids) ? ids : ids ? [ids] : [])
      .map((s) => rels[s['@r:id']]?.match(/slides\/slide(\d+)\.xml$/)?.[1])
      .filter(Boolean).map(Number).filter((n) => onDisk.includes(n));
    if (order.length) return order;
  } catch {}
  return onDisk;
}

export { slideRels, relsOf, imageHref, placeholderGeom };
