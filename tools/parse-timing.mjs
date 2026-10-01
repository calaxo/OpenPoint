// Parseur de l'arbre d'animation PPTX (<p:timing>) -> timeline JSON lisible.
// Usage :
//   node tools/parse-timing.mjs                 -> toutes les diapos, écrit specimen/timeline.json
//   node tools/parse-timing.mjs 8               -> diapo 8 seule, affichée sur stdout
//
// Modèle reverse-engineeré (voir docs/moteur-animation.md) :
//   p:timing > p:tnLst > p:par cTn[tmRoot]
//     > p:seq cTn[mainSeq]            <- la séquence principale (1 par diapo)
//       > p:par cTn (delay=indefinite) <- 1 par CLIC (le "click step")
//         > p:par cTn (delay=0)         <- groupe de synchro (with/after regroupés)
//           > p:par cTn[clickEffect|withEffect|afterEffect] presetID/presetClass
//             > behaviors : p:set, p:animEffect, p:anim, p:animMotion, p:animRot, p:animScale, p:cmd

import { readFileSync, writeFileSync, readdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { XMLParser } from 'fast-xml-parser';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const SLIDES_DIR = join(ROOT, 'specimen', 'extracted', 'ppt', 'slides');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@',
  preserveOrder: false, // l'ordre des tags différents est perdu, mais on retraverse en preserveOrder pour la séquence
});
const orderedParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@',
  preserveOrder: true,
});

// ---------- helpers sur l'arbre preserveOrder de fast-xml-parser ----------
// Un noeud preserveOrder = { tagName: [children...], ':@': {attrs} }
const tagOf = (n) => Object.keys(n).find((k) => k !== ':@');
const attrs = (n) => n[':@'] ?? {};
const kids = (n) => n[tagOf(n)] ?? [];
const childrenByTag = (n, tag) => kids(n).filter((c) => tagOf(c) === tag);
const firstChild = (n, tag) => childrenByTag(n, tag)[0];

function findFirst(node, tag) {
  if (tagOf(node) === tag) return node;
  for (const c of kids(node)) {
    if (typeof c !== 'object') continue;
    const r = findFirst(c, tag);
    if (r) return r;
  }
  return null;
}

// ---------- carte spid -> nom de forme (spTree de la diapo) ----------
function shapeMap(slideXmlText) {
  const map = {};
  // les formes sont p:sp / p:pic / p:grpSp / p:graphicFrame / p:cxnSp ; toutes portent p:nvXxPr > p:cNvPr @id @name
  for (const m of slideXmlText.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)) {
    map[m[1]] = m[2];
  }
  return map;
}

// ---------- extraction d'un behavior (p:set, p:anim, p:animEffect, ...) ----------
function parseBehavior(node) {
  const tag = tagOf(node); // ex: p:animEffect
  const a = attrs(node);
  const out = { kind: tag.replace('p:', '') };

  // attributs propres au behavior
  for (const k of ['@transition', '@filter', '@calcmode', '@valueType', '@by', '@from', '@to']) {
    if (a[k] !== undefined) out[k.slice(1)] = a[k];
  }

  const cBhvr = firstChild(node, 'p:cBhvr');
  if (cBhvr) {
    const bCTn = firstChild(cBhvr, 'p:cTn');
    if (bCTn) {
      const ba = attrs(bCTn);
      if (ba['@dur'] !== undefined) out.dur = ba['@dur'];
      if (ba['@repeatCount'] !== undefined) out.repeat = ba['@repeatCount'];
      if (ba['@autoRev'] !== undefined) out.autoRev = ba['@autoRev'];
      if (ba['@fill'] !== undefined) out.fill = ba['@fill'];
      const cond = findFirst(bCTn, 'p:cond');
      if (cond && attrs(cond)['@delay'] !== undefined) out.delay = attrs(cond)['@delay'];
    }
    const tgt = findFirst(cBhvr, 'p:spTgt');
    if (tgt) {
      out.spid = attrs(tgt)['@spid'];
      const txEl = firstChild(tgt, 'p:txEl');
      if (txEl) {
        const pRg = findFirst(txEl, 'p:pRg');
        if (pRg) out.paragraphRange = [attrs(pRg)['@st'], attrs(pRg)['@end']];
      }
    }
    const attrNames = findFirst(cBhvr, 'p:attrNameLst');
    if (attrNames) out.attrs = childrenByTag(attrNames, 'p:attrName').map((n) => kids(n)[0]?.['#text']);
  }

  // valeurs to/from/by structurées (p:to > p:strVal/p:fltVal/p:clrVal)
  for (const dir of ['p:from', 'p:to', 'p:by']) {
    const v = firstChild(node, dir);
    if (v) {
      const val = kids(v)[0];
      if (val) {
        const va = attrs(val);
        out[dir.replace('p:', '')] = va['@val'] ?? (va['@x'] !== undefined ? { x: va['@x'], y: va['@y'] } : tagOf(val));
      }
    }
  }

  // keyframes (p:tavLst > p:tav @tm > p:val)
  const tavLst = findFirst(node, 'p:tavLst');
  if (tavLst) {
    out.keyframes = childrenByTag(tavLst, 'p:tav').map((tav) => {
      const kf = { tm: attrs(tav)['@tm'] };
      if (attrs(tav)['@fmla']) kf.fmla = attrs(tav)['@fmla'];
      const val = firstChild(tav, 'p:val');
      if (val) {
        const inner = kids(val)[0];
        if (inner) kf.val = attrs(inner)['@val'];
      }
      return kf;
    });
  }

  // animMotion : chemin
  if (a['@path']) out.path = a['@path'];
  return out;
}

// ---------- extraction d'un noeud "effet" (cTn avec presetID) ----------
function parseEffect(parNode, shapes, accDelay) {
  const cTn = firstChild(parNode, 'p:cTn');
  const a = attrs(cTn);
  const behaviors = childrenByTag(firstChild(cTn, 'p:childTnLst') ?? { x: [] }, undefined);
  const behaviorNodes = kids(firstChild(cTn, 'p:childTnLst') ?? { none: [] }).filter((c) =>
    ['p:set', 'p:anim', 'p:animEffect', 'p:animMotion', 'p:animRot', 'p:animScale', 'p:cmd', 'p:audio', 'p:video'].includes(tagOf(c))
  );
  const parsed = behaviorNodes.map(parseBehavior);
  const spid = parsed.find((b) => b.spid)?.spid;
  const stCond = firstChild(cTn, 'p:stCondLst');
  const ownDelay = stCond ? Number(attrs(firstChild(stCond, 'p:cond') ?? {})['@delay']) || 0 : 0;
  return {
    nodeType: a['@nodeType'],
    preset: `${a['@presetClass']}:${a['@presetID']}:${a['@presetSubtype'] ?? 0}`,
    spid,
    shape: shapes[spid] ?? null,
    // délai effectif depuis le clic = somme des délais des groupes parents + le sien
    delay: accDelay + ownDelay,
    behaviors: parsed,
  };
}

// Descend récursivement et collecte tous les cTn porteurs de presetID, dans l'ordre document.
// accDelay = somme des délais numériques des cTn parents (les groupes de synchro) —
// c'est ainsi que PowerPoint encode les "après le précédent" : un délai absolu depuis le clic.
function collectEffects(node, shapes, acc, accDelay = 0) {
  if (tagOf(node) === 'p:par') {
    const cTn = firstChild(node, 'p:cTn');
    if (cTn && attrs(cTn)['@presetID'] !== undefined) {
      acc.push(parseEffect(node, shapes, accDelay));
      return;
    }
    if (cTn) {
      const stCond = firstChild(cTn, 'p:stCondLst');
      const d = stCond ? attrs(firstChild(stCond, 'p:cond') ?? {})['@delay'] : undefined;
      if (d !== undefined && d !== 'indefinite') accDelay += Number(d) || 0;
    }
  }
  for (const c of kids(node)) if (typeof c === 'object') collectEffects(c, shapes, acc, accDelay);
}

export function parseSlide(slideNo) {
  const text = readFileSync(join(SLIDES_DIR, `slide${slideNo}.xml`), 'utf8');
  const shapes = shapeMap(text);
  const m = text.match(/<p:timing>[\s\S]*<\/p:timing>/);
  if (!m) return { slide: slideNo, animated: false, clickSteps: [] };

  const tree = orderedParser.parse(m[0]);
  const timing = tree.find((n) => tagOf(n) === 'p:timing');
  const mainSeq = findFirst(timing, 'p:seq'); // la séquence principale
  const seqCTn = firstChild(mainSeq, 'p:cTn');
  const clickPars = childrenByTag(firstChild(seqCTn, 'p:childTnLst'), 'p:par');

  const clickSteps = clickPars.map((par, i) => {
    const effects = [];
    collectEffects(par, shapes, effects);
    return { step: i + 1, effects };
  });

  // liste de build (quelles formes sont construites / animBg)
  const builds = [...m[0].matchAll(/<p:bldP spid="(\d+)"[^>]*>/g)].map((b) => b[0]);
  return { slide: slideNo, animated: true, clickSteps, buildCount: builds.length };
}

// ---------- main ----------
// Tolérant à l'absence d'échantillon local : sans dossier de diapos, l'outil ne fait rien
// (les fonctions exportées restent chargées et testables par les tests unitaires).
const arg = process.argv[2];
try {
  if (arg) {
    console.log(JSON.stringify(parseSlide(Number(arg)), null, 2));
  } else {
    const nums = readdirSync(SLIDES_DIR)
      .filter((f) => /^slide\d+\.xml$/.test(f))
      .map((f) => Number(f.match(/\d+/)[0]))
      .sort((a, b) => a - b);
    const all = nums.map(parseSlide);
    const out = join(ROOT, 'specimen', 'timeline.json');
    writeFileSync(out, JSON.stringify(all, null, 2));
    const total = all.reduce((s, sl) => s + sl.clickSteps.reduce((c, st) => c + st.effects.length, 0), 0);
    console.log(`OK -> ${out} (${all.filter((s) => s.animated).length} diapos animées, ${total} effets)`);
  }
} catch { /* pas d'échantillon local : rien à faire */ }

// ---------- exports nommés pour les tests unitaires ----------
// Bloc AJOUTÉ en fin de fichier (aucune ligne de logique existante modifiée).
// N'exporte que des fonctions PURES : leur sortie ne dépend que de leurs
// arguments (pas de fs/réseau/DOM/état global). `parseSlide` reste exclu ici
// car il lit le disque (déjà exporté plus haut de toute façon).
export {
  tagOf,
  attrs,
  kids,
  childrenByTag,
  firstChild,
  findFirst,
  shapeMap,
  parseBehavior,
  parseEffect,
  collectEffects,
};
