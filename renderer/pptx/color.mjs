// ═══ [PPTX · COULEURS] ═══ couleur DrawingML (srgbClr, schemeClr… + lumMod/shade/tint/alpha) -> CSS.
import { presetHex } from './preset-colors.mjs';

// Une couleur DrawingML (§20.1.2.3) = une couleur de base (srgbClr, schemeClr, sysClr, prstClr,
// scrgbClr, hslClr) + une liste de TRANSFORMATIONS enfants (lumMod, lumOff,
// shade, tint, alpha…). Ex. « accent1 assombri 25 % » = <a:schemeClr val="accent1"><a:lumMod
// val="75000"/></a:schemeClr>.
// ================= couleurs =================
// Conversions HSL <-> RVB classiques (h, s, l dans [0,1]) ; utilisées pour lumMod / lumOff qui
// agissent sur la LUMINANCE HSL.
function hslFromRgb(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h, s, l = (mx + mn) / 2;
  if (mx === mn) { h = s = 0; }
  else { const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
  return [h, s, l];
}
function rgbFromHsl(h, s, l) {
  let r, g, b;
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const hue = (t) => { t = (t + 1) % 1; return t < 1/6 ? p + (q - p) * 6 * t : t < 1/2 ? q : t < 2/3 ? p + (q - p) * (2/3 - t) * 6 : p; };
    r = hue(h + 1/3); g = hue(h); b = hue(h - 1/3); }
  return [Math.round(r*255), Math.round(g*255), Math.round(b*255)];
}
// résout un noeud couleur DrawingML flat -> {r,g,b,a}
// Entrée : un nœud PARENT d'une couleur (ex. le contenu de a:solidFill, un a:gs, un a:fillRef, un
// a:outerShdw) — la fonction cherche l'enfant srgbClr / schemeClr / … Sortie : r,g,b entiers 0-255,
// a dans [0,1] ; null si aucun enfant couleur (l'appelant choisit alors son repli).
// Ordre d'application des transformations (fixe, indépendant de l'ordre du XML — le parse plat ne le
// connaît pas) : lumMod, puis lumOff (sur la luminance HSL), puis shade (vers le noir), puis tint
// (vers le blanc), puis alpha. Modèle SIMPLIFIÉ : shade/tint sont appliqués linéairement en sRVB.
// Transformations ignorées : satMod, satOff, hueMod, hueOff, comp, inv, gray, gamma, red*/green*/…
function resolveColorFlat(node, ctx) {
  if (!node) return null;
  let hex = null, mods = node;
  // schemeHex(nom) : phClr -> couleur de substitution courante (défaut accent1) ; alias du clrMap
  // (bg1 -> lt1 …) ; puis slot du thème. Slot inconnu -> gris 888888 (visible, jamais transparent).
  const schemeHex = (nm) => {
    if (nm === 'phClr') nm = ctx.phClr || 'accent1';
    if (ctx.clrMap && ctx.clrMap[nm]) nm = ctx.clrMap[nm]; // bg1/tx1/... -> dk1/lt1/...
    return ctx.clr[nm] ?? '888888';
  };
  if (node['a:srgbClr']) { hex = node['a:srgbClr']['@val']; mods = node['a:srgbClr']; }
  else if (node['a:schemeClr']) { hex = schemeHex(node['a:schemeClr']['@val']); mods = node['a:schemeClr']; }
  else if (node['a:sysClr']) { hex = node['a:sysClr']['@lastClr'] ?? '000000'; mods = node['a:sysClr']; }
  // prstClr : couleur nommée (ST_PresetColorVal, table CSS/X11 de preset-colors.mjs) ; inconnue -> gris
  else if (node['a:prstClr']) { hex = presetHex(node['a:prstClr']['@val']) ?? '888888'; mods = node['a:prstClr']; }
  // hslClr : teinte en 1/60000 de degré, saturation et luminance en 1/1000 de % (§20.1.2.3.13)
  else if (node['a:hslClr']) {
    const c = node['a:hslClr'];
    const rgb = rgbFromHsl(((+c['@hue'] || 0) / 60000 / 360) % 1, Math.min(1, (+c['@sat'] || 0) / 100000), Math.min(1, (+c['@lum'] || 0) / 100000));
    hex = rgb.map((v) => v.toString(16).padStart(2, '0')).join(''); mods = c;
  }
  // scrgbClr : composantes en % (0-100000), traitées comme du sRVB linéairement (pas de conversion
  // gamma) ; retour immédiat, SANS appliquer les transformations enfants.
  else if (node['a:scrgbClr']) { const c = node['a:scrgbClr']; const to = (v)=>Math.round(parseInt(v)/100000*255); return { r:to(c['@r']), g:to(c['@g']), b:to(c['@b']), a:1 }; }
  else return null;
  let r = parseInt(hex.slice(0,2),16), g = parseInt(hex.slice(2,4),16), b = parseInt(hex.slice(4,6),16), a = 1;
  const num = (v)=>parseInt(v)/100000; // ST_Percentage : 100000 = 100 %
  let [h,s,l] = hslFromRgb(r,g,b);
  if (mods['a:lumMod']) l *= num(mods['a:lumMod']['@val']);
  if (mods['a:lumOff']) l += num(mods['a:lumOff']['@val']);
  l = Math.max(0, Math.min(1, l));
  // repasse en RVB uniquement si la luminance a changé (évite un aller-retour HSL -> arrondis)
  if (mods['a:lumMod'] || mods['a:lumOff']) [r,g,b] = rgbFromHsl(h,s,l);
  // shade k : c × k (k = 1 -> inchangé, 0 -> noir) ; tint k : c × k + 255 × (1 - k) (0 -> blanc)
  if (mods['a:shade']) { const k = num(mods['a:shade']['@val']); r*=k; g*=k; b*=k; }
  if (mods['a:tint']) { const k = num(mods['a:tint']['@val']); r = r*k + 255*(1-k); g = g*k + 255*(1-k); b = b*k + 255*(1-k); }
  if (mods['a:alpha']) a = num(mods['a:alpha']['@val']);
  return { r:Math.round(Math.max(0,Math.min(255,r))), g:Math.round(Math.max(0,Math.min(255,g))), b:Math.round(Math.max(0,Math.min(255,b))), a };
}
// cssColor({r,g,b,a}) -> 'rgb(…)' ou 'rgba(…)' si transparente ; null -> 'none' (pratique pour fill=)
const cssColor = (c) => c ? (c.a < 1 ? `rgba(${c.r},${c.g},${c.b},${c.a.toFixed(3)})` : `rgb(${c.r},${c.g},${c.b})`) : 'none';
export { hslFromRgb, rgbFromHsl, resolveColorFlat, cssColor };
