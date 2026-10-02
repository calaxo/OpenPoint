// ═══ [DIAPO · EFFETS] ═══ ombre / halo / bords adoucis (effectFilter) et filtres d'image (imageFilter).
// Gère : ombre portée (outerShdw), halo lumineux (glow), bords adoucis (softEdge). §20.1.8
import { u } from '../units.mjs';
import { deepFindFlat } from '../xml.mjs';
import { resolveColorFlat } from '../color.mjs';

// Principe : chaque effet DrawingML est traduit en primitives de filtre SVG (feGaussianBlur,
// feOffset, feFlood, feComposite, feMorphology, feColorMatrix, feComponentTransfer) puis empilé par
// feMerge (de l'arrière vers l'avant : ombre, halo ×2, forme). Un <filter> par forme concernée.
// rc.filters : chaînes <filter …> (effets de forme ET filtres d'image) -> <defs>
// effectFilter(rc, spPr plat, p:style plat) -> ' filter="url(#…)"' ou ''.
// Autres effets de a:effectLst non gérés : innerShdw, prstShdw, reflection, blur, fillOverlay ;
// a:effectDag et scene3d/sp3d non plus.
export function effectFilter(rc, spPr, style) {
  const { ctx, uid, filters } = rc;
  // effets EXPLICITES de la forme, sinon repli sur le style de thème via <p:style><a:effectRef idx>=1>
  // (§20.1.4.2.8) : l'ombre/lueur "thème" des formes stylées (SmartArt, cartes) vient de là.
  let fx = spPr?.['a:effectLst'];
  if (!fx) {
    const er = style?.['a:effectRef']; const eidx = er ? +er['@idx'] : 0;
    if (eidx >= 1) fx = ctx.effectStyles?.[eidx - 1]?.['a:effectStyle']?.['a:effectLst'];
  }
  if (!fx) return '';
  const sh = fx['a:outerShdw'], gl = fx['a:glow'], se = fx['a:softEdge'];
  if (!sh && !gl && !se) return '';
  // pad = débord de la zone de filtre en % de la boîte de la forme (SVG coupe tout ce qui sort de la
  // région du filtre : il faut la place pour le flou et le décalage de l'ombre)
  const prims = []; const layers = []; let pad = 30;
  // bords adoucis : flou du canal alpha puis re-découpe -> source adoucie ("src")
  let src = 'SourceGraphic';
  // @rad (EMU) -> pt, /2 : le rayon d'effet OOXML est traité comme ≈ 2 écarts-types de flou gaussien
  // (même conversion pour blurRad et glow ci-dessous)
  if (se) { const r = u(+se['@rad'] || 0) / 2;
    prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${r.toFixed(1)}" result="seA"/>`
      + `<feComposite in="SourceGraphic" in2="seA" operator="in" result="src"/>`);
    src = 'src';
  }
  // ombre portée (derrière)
  if (sh) {
    // dist (EMU) + dir (1/60000 °, 0 = vers la droite, sens horaire car y vers le bas) -> dx, dy.
    // blurRad défaut 40000 EMU ≈ 3,15 pt ; couleur absente -> noir 40 %. sx/sy/kx/ky/algn ignorés.
    const dist = u(+sh['@dist'] || 0), dir = (+sh['@dir'] || 0) / 60000 * Math.PI/180;
    const blur = u(+sh['@blurRad'] || 40000), c = resolveColorFlat(sh, ctx) || { r:0,g:0,b:0,a:0.4 };
    pad = Math.max(pad, 40 + blur);
    prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${(blur/2).toFixed(1)}" result="shB"/>`
      + `<feOffset in="shB" dx="${(dist*Math.cos(dir)).toFixed(1)}" dy="${(dist*Math.sin(dir)).toFixed(1)}" result="shO"/>`
      + `<feFlood flood-color="rgb(${c.r},${c.g},${c.b})" flood-opacity="${(c.a??0.4).toFixed(2)}"/>`
      + `<feComposite in2="shO" operator="in" result="shadow"/>`);
    layers.push('shadow');
  }
  // halo lumineux (derrière la source, devant l'ombre)
  if (gl) {
    // halo = silhouette légèrement dilatée (1 unité) puis floutée, colorée ; défaut jaune 60 %
    const r = u(+gl['@rad'] || 0) / 2, c = resolveColorFlat(gl, ctx) || { r:255,g:255,b:0,a:0.6 };
    pad = Math.max(pad, 40 + r * 2);
    prims.push(`<feMorphology in="SourceAlpha" operator="dilate" radius="1" result="glD"/>`
      + `<feGaussianBlur in="glD" stdDeviation="${r.toFixed(1)}" result="glB"/>`
      + `<feFlood flood-color="rgb(${c.r},${c.g},${c.b})" flood-opacity="${(c.a??0.6).toFixed(2)}"/>`
      + `<feComposite in2="glB" operator="in" result="glow"/>`);
    layers.push('glow', 'glow'); // doublé = halo plus dense (comme PowerPoint)
  }
  layers.push(src);
  const merge = `<feMerge>${layers.map((n) => `<feMergeNode in="${n}"/>`).join('')}</feMerge>`;
  // marge de la région de filtre bornée à [30 %, 200 %]. NB : pad mélange des pt (blur) et des % ;
  // c'est une heuristique, pas une conversion exacte.
  const m = Math.min(200, Math.max(30, Math.round(pad)));
  const id = uid + 'fx' + filters.length;
  filters.push(`<filter id="${id}" x="-${m}%" y="-${m}%" width="${100+2*m}%" height="${100+2*m}%">${prims.join('')}${merge}</filter>`);
  return ` filter="url(#${id})"`;
}
// effets couleur/alpha d'une image (enfants de a:blip / blipFill) -> filtre SVG. §20.1.8.{20,25,33,4}
// imageFilter(rc, a:blip plat) -> { attr: ' filter="url(#…)"' | '', opacity: ' opacity="…"' | '' }.
// Les primitives s'enchaînent dans l'ordre du code (chaque primitive prend la sortie de la
// précédente) : grayscl, lum, biLevel, duotone, puis corrections a14.
export function imageFilter(rc, blip) {
  const { ctx, uid, filters } = rc;
  if (!blip) return { attr: '', opacity: '' };
  const prims = []; let opacity = '';
  const amf = blip['a:alphaModFix']; // opacité globale (ST_PositivePercentage)
  if (amf?.['@amt']) opacity = ` opacity="${(+amf['@amt']/100000).toFixed(3)}"`;
  if (blip['a:grayscl'] !== undefined) prims.push('<feColorMatrix type="saturate" values="0"/>');
  const lum = blip['a:lum']; // §20.1.8.42 luminance : bright (additif) + contrast (autour du gris moyen)
  if (lum && (lum['@bright'] || lum['@contrast'])) {
    // c' = c · (1 + contraste) + luminosité + 0,5 · (1 - pente) : le contraste pivote autour de 0,5
    // (gris moyen reste gris moyen), la luminosité décale le tout.
    const br = (+lum['@bright'] || 0) / 100000, ct = (+lum['@contrast'] || 0) / 100000;
    const slope = 1 + ct, inter = br + 0.5 * (1 - slope);
    const fn = (c) => `<feFunc${c} type="linear" slope="${slope.toFixed(3)}" intercept="${inter.toFixed(3)}"/>`;
    prims.push(`<feComponentTransfer>${fn('R')}${fn('G')}${fn('B')}</feComponentTransfer>`);
  }
  // noir et blanc à seuil (§20.1.8.11 biLevel) : luminance >= thresh -> blanc, sinon noir. Schémas
  // scannés : sans ça le fond gris du scan restait gris au lieu du blanc de PowerPoint.
  const bil = blip['a:biLevel'];
  if (bil) {
    // seuil en 1/1000 de % (défaut 50 %) ; table discrète de 100 marches (0 sous le seuil, 1 au-dessus)
    // appliquée à la luminance Rec. 709 (0,2126 R + 0,7152 G + 0,0722 B) recopiée sur les 3 canaux
    const t = Math.min(0.99, Math.max(0.01, (+bil['@thresh'] || 50000) / 100000)), n = 100;
    const tbl = Array.from({ length: n }, (_, i) => (i / n >= t ? 1 : 0)).join(' ');
    prims.push('<feColorMatrix type="matrix" values="0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0"/>'
      + `<feComponentTransfer><feFuncR type="discrete" tableValues="${tbl}"/><feFuncG type="discrete" tableValues="${tbl}"/><feFuncB type="discrete" tableValues="${tbl}"/></feComponentTransfer>`);
  }
  const duo = blip['a:duotone']; // bichromie : 2 couleurs de a:duotone -> remap luminance
  if (duo) {
    // Les deux couleurs sont prises dans l'ordre des CLÉS du parse plat : si elles sont de types
    // différents (ex. prstClr puis srgbClr), l'ordre du document n'est pas garanti. Chaque couleur est
    // re-emballée ({ [k]: c }) pour que resolveColorFlat lise aussi ses transformations enfants.
    // Moins de 2 couleurs -> repli noir -> blanc.
    const cols = [];
    for (const k of Object.keys(duo)) if (/Clr$/.test(k)) { const v = duo[k]; for (const c of (Array.isArray(v)?v:[v])) cols.push(resolveColorFlat({ [k]:c }, ctx)); }
    const [c0, c1] = cols.length >= 2 ? cols : [{r:0,g:0,b:0}, {r:255,g:255,b:255}];
    prims.push('<feColorMatrix type="matrix" values="0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0 0 0 1 0"/>'
      + `<feComponentTransfer><feFuncR type="table" tableValues="${(c0.r/255).toFixed(3)} ${(c1.r/255).toFixed(3)}"/>`
      + `<feFuncG type="table" tableValues="${(c0.g/255).toFixed(3)} ${(c1.g/255).toFixed(3)}"/>`
      + `<feFuncB type="table" tableValues="${(c0.b/255).toFixed(3)} ${(c1.b/255).toFixed(3)}"/></feComponentTransfer>`);
  }
  // corrections photo MS 2010 (extension a14 dans blip>extLst) : luminosité/contraste, saturation.
  // §MS-ODRAWXML a14:imgEffect. Valeurs en millièmes de % (ST_(Fixed)Percentage).
  // ⚠ Si un <a14:imgLayer r:embed=.wdp> est présent, ces corrections décrivent l'édition Office
  // appliquée à la HD photo ORIGINALE (.wdp) pour PRODUIRE l'image de base (rId2) que l'on rend déjà :
  // elles sont donc DÉJÀ cuites dans rId2. Les ré-appliquer écraserait la photo (ex. bright -40% →
  // soustraction de 102/canal → tout en noir). On ne les applique que si elles sont « directes ».
  const baked = !!deepFindFlat(blip, 'a14:imgLayer');
  const bc = baked ? null : deepFindFlat(blip, 'a14:brightnessContrast');
  if (bc && (bc['@bright'] || bc['@contrast'])) {
    const br = (+bc['@bright'] || 0) / 100000, ct = (+bc['@contrast'] || 0) / 100000;
    const slope = 1 + ct, inter = br + 0.5 * (1 - slope); // contraste autour du gris moyen + luminosité
    const fn = (c) => `<feFunc${c} type="linear" slope="${slope.toFixed(3)}" intercept="${inter.toFixed(3)}"/>`;
    prims.push(`<feComponentTransfer>${fn('R')}${fn('G')}${fn('B')}</feComponentTransfer>`);
  }
  const sat = baked ? null : deepFindFlat(blip, 'a14:saturation');
  if (sat && sat['@sat'] != null) prims.push(`<feColorMatrix type="saturate" values="${(+sat['@sat']/100000).toFixed(3)}"/>`);
  if (!prims.length) return { attr: '', opacity };
  const id = uid + 'img' + filters.length;
  // color-interpolation-filters="sRGB" : opérer en sRGB comme PowerPoint. Sinon les filtres SVG
  // tournent en linearRGB (défaut) → luminosité/contraste écrasent tout vers le noir.
  filters.push(`<filter id="${id}" color-interpolation-filters="sRGB">${prims.join('')}</filter>`);
  return { attr: ` filter="url(#${id})"`, opacity };
}
