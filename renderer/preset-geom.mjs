// Moteur de géométrie DrawingML générique (ECMA-376 §20.1.9) — SANS PowerPoint.
// Évalue les FORMULES DE GUIDES (17 opérateurs) et construit le chemin SVG pour :
//   - les 187 formes preset (données bakées : preset-shapes.json, cf. gen-preset-shapes.mjs)
//   - les custGeom (mêmes formules + pathLst ordonné)
// Remplace les ~30 approximations manuelles par les définitions officielles exactes.
//
// Modèle : def = { av:[[name,fmla]], gd:[[name,fmla]], paths:[{w,h,fill,stroke,cmds}] }
//   cmds : ["M",x,y] ["L",x,y] ["C",x1,y1,x2,y2,x3,y3] ["Q",x1,y1,x2,y2]
//          ["A",wR,hR,stAng,swAng] ["Z"]   (valeurs = nom de guide OU nombre, en chaîne)

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dir = dirname(fileURLToPath(import.meta.url));
export const PRESETS = JSON.parse(readFileSync(join(__dir, 'preset-shapes.json'), 'utf8'));

// COMBLE une OMISSION du XML ECMA : `upArrow` est absent de presetShapeDefinitions.xml
// (il contient upDownArrow EN DOUBLE et upArrowCallout, mais pas upArrow — bug de la spec).
// Sans lui, une flèche vers le haut tombait en rectangle (barre). On le définit comme le
// MIROIR VERTICAL de downArrow (apex en haut au lieu d'en bas). Cf. VENT d.20 (vecteurs PGF).
if (!PRESETS.upArrow) PRESETS.upArrow = {
  av: [['adj1', 'val 50000'], ['adj2', 'val 50000']],
  gd: [['maxAdj2', '*/ 100000 h ss'], ['a1', 'pin 0 adj1 100000'], ['a2', 'pin 0 adj2 maxAdj2'],
       ['dy1', '*/ ss a2 100000'], ['dx1', '*/ w a1 200000'], ['x1', '+- hc 0 dx1'], ['x2', '+- hc dx1 0']],
  paths: [{ w: 0, h: 0, fill: 'norm', stroke: true, cmds: [
    ['M', 'l', 'dy1'], ['L', 'x1', 'dy1'], ['L', 'x1', 'b'], ['L', 'x2', 'b'],
    ['L', 'x2', 'dy1'], ['L', 'r', 'dy1'], ['L', 'hc', 't'], ['Z']] }],
};

// constantes d'angle (60000es de degré) — §20.1.10 + guides intégrés
const ANG = { '0':0, cd8:2700000, cd4:5400000, cd2:10800000, cd:21600000,
  '3cd4':16200000, '3cd8':8100000, '5cd8':13500000, '7cd8':18900000 };

// variables intégrées (guides prédéfinis) calculées depuis w/h réels — §20.1.10.56
function baseVars(W, H) {
  const v = { l:0, t:0, r:W, b:H, w:W, h:H, hc:W/2, vc:H/2,
    ss:Math.min(W,H), ls:Math.max(W,H) };
  for (const n of [2,3,4,5,6,8,10,12]) { v['wd'+n]=W/n; v['hd'+n]=H/n; }
  for (const n of [2,4,6,8,16,32]) v['ssd'+n]=v.ss/n;
  Object.assign(v, ANG);           // 3cd4, cd2, … accessibles par nom dans les formules
  return v;
}

const isNum = (t) => typeof t === 'number' || /^-?\.?\d/.test(t);
function tok(t, vars) {
  if (typeof t === 'number') return t;
  if (t in vars) return vars[t];
  if (isNum(t)) return parseFloat(t);
  return 0;                         // guide inconnu -> 0 (robuste)
}
const rad = (sixtyK) => sixtyK / 60000 * Math.PI / 180;

// évalue une formule de guide (§20.1.9.11) — 17 opérateurs
function evalFmla(fmla, vars) {
  const p = String(fmla).trim().split(/\s+/);
  const op = p[0];
  const a = () => tok(p[1], vars), b = () => tok(p[2], vars), c = () => tok(p[3], vars);
  switch (op) {
    case 'val': return a();
    case '*/':  return a() * b() / c();
    case '+-':  return a() + b() - c();
    case '+/':  return (a() + b()) / c();
    case '?:':  return a() > 0 ? b() : c();
    case 'abs': return Math.abs(a());
    case 'sqrt':return Math.sqrt(a());
    case 'max': return Math.max(a(), b());
    case 'min': return Math.min(a(), b());
    case 'mod': return Math.hypot(a(), b(), c());              // sqrt(x²+y²+z²)
    case 'pin': { const x=a(), y=b(), z=c(); return y < x ? x : (y > z ? z : y); }
    case 'sin': return a() * Math.sin(rad(b()));
    case 'cos': return a() * Math.cos(rad(b()));
    case 'tan': return a() * Math.tan(rad(b()));
    case 'at2': return Math.atan2(b(), a()) * 180 / Math.PI * 60000;   // arctan(y/x) -> 60000es °
    case 'cat2':return a() * Math.cos(Math.atan2(c(), b()));           // x·cos(arctan(z/y))
    case 'sat2':return a() * Math.sin(Math.atan2(c(), b()));           // x·sin(arctan(z/y))
    default: return 0;
  }
}

// construit la table de guides : builtins -> avLst (avec overrides) -> gdLst séquentiel
function buildVars(def, W, H, adj) {
  const vars = baseVars(W, H);
  for (const [name, fmla] of (def.av || [])) {
    vars[name] = evalFmla((adj && adj[name] != null) ? adj[name] : fmla, vars);
  }
  for (const [name, fmla] of (def.gd || [])) vars[name] = evalFmla(fmla, vars);
  return vars;
}

const f = (n) => { const r = Math.round(n * 100) / 100; return Object.is(r, -0) ? 0 : r; };

// un <path> -> segment de chemin SVG. Coord = guide (espace-forme, non mis à l'échelle)
// ou nombre littéral (mis à l'échelle par w/h du path si présents).
function emitPath(path, vars, W, H) {
  const sx = path.w ? W / path.w : 1, sy = path.h ? H / path.h : 1;
  const X = (t) => (typeof t !== 'number' && isNum(t)) ? parseFloat(t) * sx : tok(t, vars);
  const Y = (t) => (typeof t !== 'number' && isNum(t)) ? parseFloat(t) * sy : tok(t, vars);
  const A = (t) => (t in ANG) ? rad(ANG[t]) : (isNum(t) ? rad(parseFloat(t)) : rad(tok(t, vars)));
  let d = '', cx = 0, cy = 0, sX = 0, sY = 0;
  for (const cmd of path.cmds) {
    const k = cmd[0];
    if (k === 'M') { cx = X(cmd[1]); cy = Y(cmd[2]); sX = cx; sY = cy; d += `M${f(cx)} ${f(cy)} `; }
    else if (k === 'L') { cx = X(cmd[1]); cy = Y(cmd[2]); d += `L${f(cx)} ${f(cy)} `; }
    else if (k === 'C') { const x1=X(cmd[1]),y1=Y(cmd[2]),x2=X(cmd[3]),y2=Y(cmd[4]),x3=X(cmd[5]),y3=Y(cmd[6]);
      d += `C${f(x1)} ${f(y1)} ${f(x2)} ${f(y2)} ${f(x3)} ${f(y3)} `; cx = x3; cy = y3; }
    else if (k === 'Q') { const x1=X(cmd[1]),y1=Y(cmd[2]),x2=X(cmd[3]),y2=Y(cmd[4]);
      d += `Q${f(x1)} ${f(y1)} ${f(x2)} ${f(y2)} `; cx = x2; cy = y2; }
    else if (k === 'A') {
      // arc elliptique : centre calé pour que le point de départ = position courante (§20.1.9.4)
      const wR = X(cmd[1]), hR = Y(cmd[2]), st = A(cmd[3]), sw = A(cmd[4]);
      const ecx = cx - wR * Math.cos(st), ecy = cy - hR * Math.sin(st);
      const end = st + sw, ex = ecx + wR * Math.cos(end), ey = ecy + hR * Math.sin(end);
      const large = Math.abs(sw) > Math.PI ? 1 : 0, sweep = sw >= 0 ? 1 : 0;
      if (Math.abs(sw) >= 2 * Math.PI) {
        // cercle complet : 2 demi-arcs (SVG ne dessine pas 360° d'un coup)
        const mx = ecx + wR * Math.cos(st + Math.PI), my = ecy + hR * Math.sin(st + Math.PI);
        d += `A${f(wR)} ${f(hR)} 0 1 ${sweep} ${f(mx)} ${f(my)} A${f(wR)} ${f(hR)} 0 1 ${sweep} ${f(ex)} ${f(ey)} `;
      } else {
        d += `A${f(wR)} ${f(hR)} 0 ${large} ${sweep} ${f(ex)} ${f(ey)} `;
      }
      cx = ex; cy = ey;
    }
    else if (k === 'Z') { d += 'Z '; cx = sX; cy = sY; }
  }
  return d;
}

// def -> chemin SVG (silhouette : paths remplis + défaut ; ignore fill='none'/ombrage darken/lighten)
export function evalGeom(def, W, H, adj) {
  if (!def || !def.paths || !def.paths.length) return `M0 0 H${f(W)} V${f(H)} H0 Z`;
  const vars = buildVars(def, W, H, adj);
  let d = '';
  for (const path of def.paths) {
    if (path.fill === 'none') continue;                       // détail au trait (v2 : non tracé)
    if (/^(darken|lighten)/.test(path.fill)) continue;        // ombrage 3D -> ignoré (rendu plat)
    d += emitPath(path, vars, W, H);
  }
  if (!d) d = def.paths.map((p) => emitPath(p, vars, W, H)).join(''); // tout en 'none' -> tracer quand même
  return d.trim() || `M0 0 H${f(W)} V${f(H)} H0 Z`;
}

// def -> { fill, stroke } : chemins SÉPARÉS pour le remplissage et le tracé, en honorant les attributs
// par sous-path (a:path @fill / @stroke). Certaines formes (arc, chord…) ont un path « secteur » à
// remplir SANS tracé (stroke=false) et un path « courbe » à tracer SANS remplissage (fill=none) :
// tout fusionner traçait les rayons du secteur (ex. 45° en camembert au lieu d'un arc fin).
export function evalGeomParts(def, W, H, adj) {
  if (!def || !def.paths || !def.paths.length) { const d = `M0 0 H${f(W)} V${f(H)} H0 Z`; return { fill: d, stroke: d }; }
  const vars = buildVars(def, W, H, adj);
  let fillD = '', strokeD = '';
  for (const path of def.paths) {
    const seg = emitPath(path, vars, W, H);
    const fillable = path.fill !== 'none' && !/^(darken|lighten)/.test(path.fill);
    const strokable = path.stroke !== false;
    if (fillable) fillD += seg;
    if (strokable) strokeD += seg;
  }
  if (!fillD) fillD = def.paths.map((p) => emitPath(p, vars, W, H)).join('');
  return { fill: fillD.trim() || `M0 0 H${f(W)} V${f(H)} H0 Z`, stroke: strokeD.trim() };
}

// preset par nom (ST_ShapeType). adj = map { nomGuide: fmla|nombre } depuis l'avLst de la forme.
export function presetGeomD(name, W, H, adj) {
  const def = PRESETS[name];
  if (!def) return null;                                      // inconnu -> laisser l'appelant décider
  return evalGeom(def, W, H, adj);
}

export function presetGeomParts(name, W, H, adj) {
  const def = PRESETS[name];
  if (!def) return null;
  return evalGeomParts(def, W, H, adj);
}
