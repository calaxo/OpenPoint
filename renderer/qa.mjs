// ═══ [QUALITÉ · non-régression] ═══ docs/ARCHITECTURE.md (§6)
// Rend TOUT le corpus PPTX (les 3 packs de référence), compte les erreurs, calcule un hash SVG par
// diapo et un hash global, et DIFFE contre une baseline sauvegardée -> mesure exacte du "blast radius"
// (combien de diapos changent) après une modif du moteur. C'est le filet de sécurité du mandat
// zéro-régression : on note un hash, on modifie, on re-mesure ; si des diapos changent ce doit être
// EXACTEMENT celles voulues, et les erreurs doivent rester à 0.
//
// Usage :
//   node renderer/qa.mjs                 -> rend le corpus, affiche erreurs + hash global + diff vs baseline
//   node renderer/qa.mjs --save          -> idem, puis ÉCRIT la baseline (renderer/.qa-baseline.json)
//   node renderer/qa.mjs --list 40       -> affiche jusqu'à 40 diapos modifiées (défaut 25)
//   node renderer/qa.mjs "dossier"       -> corpus personnalisé (dossier de .pptx), répétable
import { renderSlide } from './pptx-render.mjs';
import { readFileSync, readdirSync, existsSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { openDeck } from './lib/deck.mjs';
import { createHash } from 'node:crypto';

const md5 = (s) => createHash('md5').update(s).digest('hex');
const args = process.argv.slice(2);
const save = args.includes('--save');
const listN = args.includes('--list') ? +args[args.indexOf('--list') + 1] || 25 : 25;
const dirArgs = args.filter((a, i) => !a.startsWith('--') && args[i - 1] !== '--list');
// corpus par défaut = les 3 packs de référence (design/sources différents = bon échantillon)
const CORPORA = dirArgs.length ? dirArgs : [
  'specimen/deck A',
  'specimen/deck B',
  'specimen/deck C',
];
const WORK = join('renderer', '_work', '_qa');
const BASELINE = join('renderer', '.qa-baseline.json');

const hashes = {}; // "pack/fichier#n" -> md5(svg)
let totErr = 0, totSlides = 0; const errList = [];
for (const dir of CORPORA) {
  if (!existsSync(dir)) { console.log(`(absent) ${dir}`); continue; }
  const files = readdirSync(dir).filter((f) => f.toLowerCase().endsWith('.pptx') && !f.startsWith('~$'));
  for (const f of files) {
    // décompression sûre (lib/deck.mjs), sans conversion EMF (hash stables, images non embarquées)
    let ppt; try { ppt = openDeck(join(dir, f), { workDir: WORK, metafiles: false }).pptDir; }
    catch (e) { console.log(`  !! deck illisible : ${f} (${e.message})`); continue; }
    // clés de hash par NUMÉRO DE FICHIER (tous les slideN.xml, tri numérique) : baseline inchangée
    const nums = readdirSync(join(ppt, 'slides')).filter((x) => /^slide\d+\.xml$/.test(x)).map((x) => +x.match(/\d+/)[0]).sort((a, b) => a - b);
    for (const n of nums) {
      totSlides++;
      const key = `${f}#${n}`;
      try { hashes[key] = md5(JSON.stringify(renderSlide(ppt, n, { embedImages: false }))); }
      catch (e) { hashes[key] = 'ERR'; totErr++; errList.push(`${key}: ${e.message}`); }
    }
  }
}
if (existsSync(WORK)) rmSync(WORK, { recursive: true, force: true });

const keys = Object.keys(hashes).sort();
const global = md5(keys.map((k) => k + '=' + hashes[k]).join('\n'));
console.log(`\n=== QA : ${totSlides} diapos, ${totErr} erreur(s) ===`);
if (errList.length) console.log('ERREURS:\n  ' + errList.slice(0, 30).join('\n  '));
console.log(`hash global : ${global}`);

if (existsSync(BASELINE)) {
  const base = JSON.parse(readFileSync(BASELINE, 'utf8'));
  const changed = [], added = [], removed = [];
  for (const k of keys) { if (!(k in base)) added.push(k); else if (base[k] !== hashes[k]) changed.push(k); }
  for (const k of Object.keys(base)) if (!(k in hashes)) removed.push(k);
  if (!changed.length && !added.length && !removed.length) console.log('vs baseline : IDENTIQUE (0 diapo modifiée)');
  else {
    console.log(`vs baseline : ${changed.length} modifiée(s), ${added.length} ajout, ${removed.length} retrait`);
    const byPack = {}; for (const k of changed) { const p = k.split('#')[0]; byPack[p] = (byPack[p] || 0) + 1; }
    for (const [p, c] of Object.entries(byPack).sort((a, b) => b[1] - a[1])) console.log(`  ${String(c).padStart(4)}  ${p}`);
    if (changed.length) console.log('  diapos : ' + changed.slice(0, listN).join(' ') + (changed.length > listN ? ' …' : ''));
  }
} else console.log('(pas de baseline — lance avec --save pour en créer une)');

if (save) { writeFileSync(BASELINE, JSON.stringify(hashes)); console.log(`baseline écrite : ${BASELINE} (${keys.length} diapos)`); }
