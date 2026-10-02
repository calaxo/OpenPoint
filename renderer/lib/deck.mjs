// ═══ Ouverture d'un .pptx pour les outils (viewers, check, qa, sweep) ═══
// Un seul endroit pour : décompresser un .pptx SANS risque (garde anti zip-slip), convertir les
// EMF/WMF (Windows), lire l'ordre de présentation et les diapos masquées.
//
// API :
//   openDeck(pptxPath, { workRoot, workDir, metafiles = true, reuse = false })
//     -> { pptDir, workDir, name, slides, hidden }
//        pptDir  : dossier ppt/ décompressé (argument attendu par renderSlide & co)
//        workDir : dossier de travail (= workDir, sinon <workRoot>/<name>) ; vidé avant décompression
//        name    : nom du fichier sans .pptx, assaini ([^\w.-] -> _)
//        slides  : numéros N de slideN.xml dans l'ORDRE DE PRÉSENTATION (slideOrder du moteur)
//        hidden  : Set des numéros de diapos masquées (p:sld@show="0")
//     options : metafiles=false -> pas de conversion EMF/WMF (check/qa n'embarquent pas d'images) ;
//               reuse=true -> si workDir contient déjà ppt/presentation.xml, on ne redécompresse pas.
//   isEntrySafe(nomDEntrée) -> true si l'entrée ZIP peut être écrite sous le dossier de travail.
//   hiddenSlides(pptDir, nums) -> Set des diapos masquées parmi nums.
//   deckName(pptxPath) -> nom assaini (même règle que openDeck).
import { readFileSync, writeFileSync, mkdirSync, existsSync, rmSync } from 'node:fs';
import { join, dirname, basename, resolve, sep } from 'node:path';
import { unzipSync } from 'fflate';
import { slideOrder } from '../pptx-render.mjs';
import { convertMetafiles } from '../metafiles.mjs';

// Nom de travail d'un deck : « Mon deck (v2).pptx » -> « Mon_deck__v2_ »
export function deckName(pptxPath) {
  return basename(pptxPath).replace(/\.pptx$/i, '').replace(/[^\w.-]/g, '_');
}

// Garde anti zip-slip : une entrée ZIP ne doit JAMAIS sortir du dossier de travail.
// Refusé : chemin absolu (/x, \x), lettre de lecteur (C:), segment « .. », octet nul.
export function isEntrySafe(entry) {
  if (!entry || entry.includes('\0')) return false;
  const p = entry.replace(/\\/g, '/');
  if (p.startsWith('/')) return false;
  if (/^[a-zA-Z]:/.test(p)) return false;
  return !p.split('/').includes('..');
}

// Diapos masquées : attribut show="0" sur la balise racine <p:sld …> (lu par regex, sans parser tout).
export function hiddenSlides(pptDir, nums) {
  const out = new Set();
  for (const n of nums) {
    const file = join(pptDir, 'slides', `slide${n}.xml`);
    if (!existsSync(file)) continue;
    const head = readFileSync(file, 'utf8').slice(0, 4000);
    const open = head.match(/<p:sld\b[^>]*>/);
    if (open && /\sshow="(0|false)"/.test(open[0])) out.add(n);
  }
  return out;
}

// Décompresse le .pptx dans workDir. Lève une erreur si une entrée est dangereuse ou si le
// fichier n'est pas un ZIP : l'appelant décide (check/qa affichent « zip invalide » et continuent).
function unzipTo(pptxPath, workDir) {
  const zip = unzipSync(new Uint8Array(readFileSync(pptxPath)));
  const root = resolve(workDir);
  for (const [entry, data] of Object.entries(zip)) {
    if (entry.endsWith('/')) continue; // dossier
    if (!isEntrySafe(entry)) throw new Error(`entrée ZIP refusée (zip-slip) : ${entry}`);
    const out = resolve(root, entry);
    // double contrôle : le chemin résolu reste sous root
    if (out !== root && !out.startsWith(root + sep)) throw new Error(`entrée ZIP hors du dossier : ${entry}`);
    mkdirSync(dirname(out), { recursive: true });
    writeFileSync(out, Buffer.from(data));
  }
}

export function openDeck(pptxPath, { workRoot, workDir, metafiles = true, reuse = false } = {}) {
  const name = deckName(pptxPath);
  const dir = workDir || join(workRoot || join('renderer', '_work'), name);
  const pptDir = join(dir, 'ppt');
  const already = reuse && existsSync(join(pptDir, 'presentation.xml'));
  if (!already) {
    if (existsSync(dir)) rmSync(dir, { recursive: true, force: true });
    unzipTo(pptxPath, dir);
    if (metafiles) convertMetafiles(join(pptDir, 'media')); // EMF/WMF -> PNG (Windows seulement)
  }
  if (!existsSync(join(pptDir, 'slides'))) throw new Error(`pas de dossier ppt/slides dans ${pptxPath}`);
  const slides = slideOrder(pptDir);
  return { pptDir, workDir: dir, name, slides, hidden: hiddenSlides(pptDir, slides) };
}
