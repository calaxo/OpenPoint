// Générateur one-shot : presetShapeDefinitions.xml (ECMA-376 Part 1) -> renderer/preset-shapes.json
// Le JSON produit est auto-contenu et COMMITÉ (données d'interop, comme python-pptx / LibreOffice oox).
// Les 187 formes preset avec leurs formules de guides -> rendues génériquement par preset-geom.mjs.
//
// Usage : node renderer/gen-preset-shapes.mjs
// (nécessite le XML local dans docs/ooxml/specs/.../_extract/geom/, non commité, droits ECMA)

import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { XMLParser } from 'fast-xml-parser';

const __dir = dirname(fileURLToPath(import.meta.url));
const SRC = join(__dir, '..', 'docs', 'ooxml', 'specs',
  'ECMA-376-1_5th_edition_december_2016', '_extract', 'geom', 'presetShapeDefinitions.xml');
const OUT = join(__dir, 'preset-shapes.json');

if (!existsSync(SRC)) {
  console.error('XML source introuvable :', SRC);
  console.error('Déposer OfficeOpenXML-DrawingMLGeometries.zip et l\'extraire (voir progress-log.md).');
  process.exit(1);
}

const P = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', preserveOrder: true,
  parseTagValue: false, parseAttributeValue: false, trimValues: false, removeNSPrefix: true });

const tree = P.parse(readFileSync(SRC, 'utf8'));
const tagOf = (n) => Object.keys(n).find((k) => k !== ':@');
const at = (n) => n?.[':@'] ?? {};
const kids = (n) => { const v = n?.[tagOf(n)]; return Array.isArray(v) ? v : []; };
const byTag = (n, t) => kids(n).filter((c) => tagOf(c) === t);
const first = (n, t) => byTag(n, t)[0];

// racine <presetShapeDefinitons> (coquille dans le XML ECMA d'origine : « Definitons ») : 1 enfant/forme
const root = tree.find((n) => /ShapeDefinit/.test(tagOf(n) || ''));
const shapes = {};

const gdOf = (node, listTag) => {
  const lst = first(node, listTag); if (!lst) return [];
  return byTag(lst, 'gd').map((g) => [at(g)['@name'], at(g)['@fmla']]);
};

const CMD = { moveTo: 'M', lnTo: 'L', arcTo: 'A', cubicBezTo: 'C', quadBezTo: 'Q', close: 'Z' };

for (const sh of kids(root)) {
  const name = tagOf(sh); if (!name) continue;
  const def = { av: gdOf(sh, 'avLst'), gd: gdOf(sh, 'gdLst') };
  const rect = first(sh, 'rect');
  if (rect) { const r = at(rect); def.rect = [r['@l'], r['@t'], r['@r'], r['@b']]; }
  const pathLst = first(sh, 'pathLst');
  def.paths = [];
  for (const p of (pathLst ? byTag(pathLst, 'path') : [])) {
    const pa = at(p);
    const path = {
      w: pa['@w'] ? +pa['@w'] : 0, h: pa['@h'] ? +pa['@h'] : 0,
      fill: pa['@fill'] ?? 'norm',            // norm|none|lighten|lightenLess|darken|darkenLess
      stroke: pa['@stroke'] !== 'false',
      cmds: [],
    };
    for (const seg of kids(p)) {
      const t = tagOf(seg), c = CMD[t]; if (!c) continue;
      if (c === 'Z') { path.cmds.push(['Z']); continue; }
      if (c === 'A') { const a = at(seg); path.cmds.push(['A', a['@wR'], a['@hR'], a['@stAng'], a['@swAng']]); continue; }
      const pts = byTag(seg, 'pt').map((pt) => [at(pt)['@x'], at(pt)['@y']]);
      path.cmds.push([c, ...pts.flat()]);
    }
    def.paths.push(path);
  }
  shapes[name] = def;
}

writeFileSync(OUT, JSON.stringify(shapes));
console.log(`OK : ${Object.keys(shapes).length} formes -> ${OUT} (${(readFileSync(OUT).length/1024).toFixed(0)} Ko)`);
