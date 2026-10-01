// Exporte, pour chaque diapo animée, les données nécessaires au POC de lecteur :
// géométrie des formes (spTree, coordonnées EMU -> fractions de diapo), couleur,
// texte, images, + la timeline parsée (parse-timing.mjs).
// Sortie : poc/poc-data.json
//
// Usage : node tools/export-poc.mjs

import { readFileSync, writeFileSync, readdirSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { XMLParser } from 'fast-xml-parser';
import { parseSlide } from './parse-timing.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const PPT = join(ROOT, 'specimen', 'extracted', 'ppt');

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@' });

// taille de diapo (EMU)
const presXml = parser.parse(readFileSync(join(PPT, 'presentation.xml'), 'utf8'));
const sldSz = presXml['p:presentation']['p:sldSz'];
const SLIDE_W = Number(sldSz['@cx']);
const SLIDE_H = Number(sldSz['@cy']);

// couleurs du thème (schemeClr -> hex), best-effort
function themeColors() {
  const theme = parser.parse(readFileSync(join(PPT, 'theme', 'theme1.xml'), 'utf8'));
  const scheme = theme['a:theme']['a:themeElements']['a:clrScheme'];
  const out = {};
  for (const [k, v] of Object.entries(scheme)) {
    if (!k.startsWith('a:')) continue;
    const name = k.slice(2);
    const hex = v['a:srgbClr']?.['@val'] ?? { windowText: '000000', window: 'FFFFFF' }[v['a:sysClr']?.['@val']] ?? null;
    if (hex) out[name] = '#' + hex;
  }
  // alias PowerPoint : tx1/bg1 pointent sur dk1/lt1
  out.tx1 ??= out.dk1; out.bg1 ??= out.lt1; out.tx2 ??= out.dk2; out.bg2 ??= out.lt2;
  return out;
}
const THEME = themeColors();

const arr = (x) => (x === undefined ? [] : Array.isArray(x) ? x : [x]);

function solidColor(props) {
  const fill = props?.['a:solidFill'];
  if (!fill) return null;
  if (fill['a:srgbClr']) return '#' + fill['a:srgbClr']['@val'];
  if (fill['a:schemeClr']) {
    const base = THEME[fill['a:schemeClr']['@val']] ?? '#999999';
    const mod = fill['a:schemeClr'];
    // lumMod/lumOff approximés en opacité de mélange — POC seulement
    return base;
  }
  return null;
}

function textOf(sp) {
  const paras = arr(sp['p:txBody']?.['a:p']);
  return paras
    .map((p) => arr(p['a:r']).map((r) => r['a:t'] ?? '').join(''))
    .join('\n')
    .trim();
}

// géométrie des placeholders : héritée du slideLayout (clé "type:idx")
function layoutPlaceholders(slideNo, rels) {
  const layoutRel = Object.values(rels).find((t) => t.includes('slideLayout'));
  if (!layoutRel) return {};
  const layoutPath = join(PPT, 'slides', layoutRel);
  if (!existsSync(layoutPath)) return {};
  const lx = parser.parse(readFileSync(layoutPath, 'utf8'));
  const tree = lx['p:sldLayout']?.['p:cSld']?.['p:spTree'];
  const out = {};
  for (const sp of arr(tree?.['p:sp'])) {
    const ph = sp['p:nvSpPr']?.['p:nvPr']?.['p:ph'];
    const xfrm = sp['p:spPr']?.['a:xfrm'];
    if (!ph || !xfrm?.['a:off']) continue;
    out[`${ph['@type'] ?? 'body'}:${ph['@idx'] ?? '0'}`] = xfrm;
  }
  return out;
}

function collectShapes(node, rels, xform, out, parentGroups = [], phGeom = {}) {
  // xform : mapping du repère courant -> fractions de diapo
  const mapPt = (x, y) => ({ x: xform.ox + (x - xform.cx0) * xform.sx, y: xform.oy + (y - xform.cy0) * xform.sy });

  for (const sp of arr(node['p:sp']).concat(arr(node['p:pic']), arr(node['p:graphicFrame']), arr(node['p:cxnSp']))) {
    const nv = sp['p:nvSpPr'] ?? sp['p:nvPicPr'] ?? sp['p:nvGraphicFramePr'] ?? sp['p:nvCxnSpPr'];
    const cNvPr = nv?.['p:cNvPr'];
    if (!cNvPr) continue;
    let xfrm = sp['p:spPr']?.['a:xfrm'] ?? sp['p:xfrm'];
    if (!xfrm?.['a:off']) {
      // placeholder sans géométrie propre -> héritage layout (piège connu du POC les suites bureautiques)
      const ph = nv?.['p:nvPr']?.['p:ph'];
      if (ph) xfrm = phGeom[`${ph['@type'] ?? 'body'}:${ph['@idx'] ?? '0'}`];
    }
    if (!xfrm?.['a:off'] || !xfrm?.['a:ext']) continue;
    const off = { x: Number(xfrm['a:off']['@x']), y: Number(xfrm['a:off']['@y']) };
    const ext = { cx: Number(xfrm['a:ext']['@cx']), cy: Number(xfrm['a:ext']['@cy']) };
    const p0 = mapPt(off.x, off.y);
    const shape = {
      id: String(cNvPr['@id']),
      name: cNvPr['@name'],
      x: p0.x, y: p0.y,
      w: ext.cx * xform.sx, h: ext.cy * xform.sy,
      rot: xfrm['@rot'] ? Number(xfrm['@rot']) / 60000 : 0,
      geom: sp['p:spPr']?.['a:prstGeom']?.['@prst'] ?? (sp['p:spPr']?.['a:custGeom'] ? 'custom' : 'rect'),
      groups: parentGroups,
      fill: solidColor(sp['p:spPr']) ?? (sp['p:pic'] ? null : solidColor(sp['p:style'] ? { 'a:solidFill': { 'a:schemeClr': sp['p:style']['a:fillRef']?.['a:schemeClr'] } } : null)),
      text: sp['p:txBody'] ? textOf(sp) : '',
    };
    const embed = sp['p:blipFill']?.['a:blip']?.['@r:embed'];
    if (embed && rels[embed]) shape.img = rels[embed].replace('../media/', 'media/');
    out.push(shape);
  }

  for (const grp of arr(node['p:grpSp'])) {
    const gx = grp['p:grpSpPr']?.['a:xfrm'];
    const cNvPr = grp['p:nvGrpSpPr']?.['p:cNvPr'];
    if (gx) {
      const off = { x: Number(gx['a:off']['@x']), y: Number(gx['a:off']['@y']) };
      const ext = { cx: Number(gx['a:ext']['@cx']), cy: Number(gx['a:ext']['@cy']) };
      const chOff = { x: Number(gx['a:chOff']?.['@x'] ?? off.x), y: Number(gx['a:chOff']?.['@y'] ?? off.y) };
      const chExt = { cx: Number(gx['a:chExt']?.['@cx'] ?? ext.cx), cy: Number(gx['a:chExt']?.['@cy'] ?? ext.cy) };
      const p0 = { x: xform.ox + (off.x - xform.cx0) * xform.sx, y: xform.oy + (off.y - xform.cy0) * xform.sy };
      const child = {
        ox: p0.x, oy: p0.y,
        cx0: chOff.x, cy0: chOff.y,
        sx: (ext.cx / chExt.cx) * xform.sx,
        sy: (ext.cy / chExt.cy) * xform.sy,
      };
      // le groupe lui-même est ciblable par une animation -> on l'exporte aussi
      if (cNvPr) {
        collectShapesGroupBox(cNvPr, p0, ext, xform, out);
      }
      collectShapes(grp, rels, child, out, cNvPr ? [...parentGroups, String(cNvPr['@id'])] : parentGroups, phGeom);
    } else {
      collectShapes(grp, rels, xform, out, parentGroups, phGeom);
    }
  }
}

function collectShapesGroupBox(cNvPr, p0, ext, xform, out) {
  out.push({
    id: String(cNvPr['@id']),
    name: cNvPr['@name'],
    x: p0.x, y: p0.y,
    w: ext.cx * xform.sx, h: ext.cy * xform.sy,
    rot: 0, geom: 'group', fill: null, text: '', group: true,
  });
}

function slideRels(n) {
  const p = join(PPT, 'slides', '_rels', `slide${n}.xml.rels`);
  if (!existsSync(p)) return {};
  const rels = parser.parse(readFileSync(p, 'utf8'));
  const out = {};
  for (const r of arr(rels.Relationships?.Relationship)) out[r['@Id']] = r['@Target'];
  return out;
}

const nums = readdirSync(join(PPT, 'slides'))
  .filter((f) => /^slide\d+\.xml$/.test(f))
  .map((f) => Number(f.match(/\d+/)[0]))
  .sort((a, b) => a - b);

const slides = [];
for (const n of nums) {
  const timeline = parseSlide(n);
  if (!timeline.animated) continue;
  const xml = parser.parse(readFileSync(join(PPT, 'slides', `slide${n}.xml`), 'utf8'));
  const spTree = xml['p:sld']['p:cSld']['p:spTree'];
  const shapes = [];
  const rootXform = { ox: 0, oy: 0, cx0: 0, cy0: 0, sx: 1 / SLIDE_W, sy: 1 / SLIDE_H };
  const rels = slideRels(n);
  collectShapes(spTree, rels, rootXform, shapes, [], layoutPlaceholders(n, rels));
  slides.push({ slide: n, shapes, clickSteps: timeline.clickSteps });
}

const out = { slideW: SLIDE_W, slideH: SLIDE_H, ratio: SLIDE_W / SLIDE_H, slides };
writeFileSync(join(ROOT, 'poc', 'poc-data.json'), JSON.stringify(out));
// variante JS pour ouvrir player.html directement en file:// (fetch interdit)
writeFileSync(join(ROOT, 'poc', 'poc-data.js'), 'window.POC_DATA = ' + JSON.stringify(out) + ';');
console.log(`OK -> poc/poc-data.json + poc-data.js (${slides.length} diapos, ${slides.reduce((s, x) => s + x.shapes.length, 0)} formes)`);
