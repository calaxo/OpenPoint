// ═══ [QUALITÉ · couverture] ═══ docs/ARCHITECTURE.md (§6) · mesure quels tags XML le moteur PPTX LIT
// Compare les tags présents dans les diapos (un dossier, ou tous les decks de _work) aux tags RÉFÉRENCÉS
// par pptx-render.mjs. Catégorise : RENDU (produit du visuel) · ANIMATION (lu par extractTiming) ·
// MÉTADONNÉES (sans effet visuel, ignoré exprès) · AUTRE (vrais trous potentiels). Sort les %.
//
// Usage :
//   node renderer/coverage.mjs                      -> agrège TOUS les decks de renderer/_work
//   node renderer/coverage.mjs <dir/ppt/slides>     -> un dossier de slides précis
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const R = dirname(fileURLToPath(import.meta.url));
const src = readFileSync(join(R, 'pptx-render.mjs'), 'utf8');

// tags du sous-arbre ANIMATION <p:timing> (lus par extractTiming / build-player, pas par le rendu)
const ANIM = new Set(['p:timing', 'p:tnLst', 'p:par', 'p:seq', 'p:cTn', 'p:stCondLst', 'p:endCondLst',
  'p:cond', 'p:childTnLst', 'p:cBhvr', 'p:tgtEl', 'p:spTgt', 'p:set', 'p:to', 'p:from', 'p:by', 'p:anim',
  'p:animEffect', 'p:animScale', 'p:animRot', 'p:animClr', 'p:animMotion', 'p:cmd', 'p:tav', 'p:tavLst',
  'p:val', 'p:attrName', 'p:attrNameLst', 'p:strVal', 'p:bldP', 'p:bldLst', 'p:bldSub', 'p:tmAbs',
  'p:tmpct', 'p:rCtr', 'p:progress', 'p:wheel', 'p:mainSeq', 'p:interactiveSeq', 'p:bg', 'p:excl',
  'p:subTnLst', 'p:iterate', 'p:fltVal', 'p:boolVal', 'p:intVal', 'p:clrVal', 'p:rgb',
  // cibles/conditions de la timeline (sous-arbre <p:timing>) : lus par le player, pas par le rendu
  'p:txEl', 'p:sldTgt', 'p:nextCondLst', 'p:prevCondLst', 'p:tn', 'p:charRg', 'p:pRg']);
// tags MÉTADONNÉES / non-visuels (ignorés à dessein). Inclut les tags « suit le texte » (a:*Tx :
// buClrTx/buSzTx/buFontTx/uLnTx/uFillTx) = comportement PAR DÉFAUT du moteur (la puce/soulignement
// hérite de la couleur/taille/police du run) -> rien à coder, correctement rendu. Idem masterClrMapping
// (= « utiliser le clrMap du master », ce qu'on fait) et les définitions de taquets a:tab/a:tabLst
// (aucun caractère tabulation réel dans le corpus -> sans effet de mise en page).
const META = /extLst|:ext$|creationId|Locks|cNvSpPr|cNvGrpSpPr|cNvCxnSpPr|cNvPicPr|nvGraphicFramePr|:cs$|:ea$|:sym$|endParaRPr|:cxn$|:pos$|:hlinkClick|:hlinkHover|contentPart|nvGrpSpPr|nvCxnSpPr|nvPicPr|nvSpPr|nvPr$|:audioFile|:videoFile|:snd$|:wavAudioFile|a14:hidden|a14:useLocalDpi|a14:cameraTool|:normAutofit|:spAutoFit|:noAutofit|prstTxWarp|:sldLayoutId|:sldMasterId|:notesMasterId|:sldSz|:notesSz|:defaultTextStyle|:embeddedFontLst|:custDataLst|:tags$|:cNvPr|:buClrTx$|:buSzTx$|:buFontTx$|:uLnTx$|:uFillTx$|:tab$|:tabLst$|:defPPr$|:masterClrMapping$|:clrMapOvr$|:ahLst$|:cxnLst$|:miter$/i;

// tags accédés DYNAMIQUEMENT dans le code (chaîne construite, ex. `a:lvl${lvl}pPr`) -> non détectables
// par recherche littérale ; on les marque comme lus explicitement.
const DYN = /^a:lvl[1-9]pPr$/;
const rendered = (t) => DYN.test(t) || src.includes("'" + t + "'") || src.includes('"' + t + '"');

// dossiers de diapos à analyser
const arg = process.argv[2];
let dirs = [];
if (arg) dirs = [arg];
else {
  const work = join(R, '_work');
  if (existsSync(work)) for (const d of readdirSync(work)) {
    const s = join(work, d, 'ppt', 'slides');
    if (existsSync(s) && d !== '_check') dirs.push(s);
  }
}
if (!dirs.length) { console.error('Aucun dossier de diapos (décompresse un .pptx via build.mjs, ou passe un chemin).'); process.exit(1); }

const count = {}; let nSlides = 0;
for (const dir of dirs) {
  for (const f of readdirSync(dir).filter((f) => /^slide\d+\.xml$/.test(f))) {
    nSlides++;
    const x = readFileSync(join(dir, f), 'utf8');
    for (const m of x.matchAll(/<(\w+:[\w]+)[ >\/]/g)) count[m[1]] = (count[m[1]] || 0) + 1;
  }
}

const cat = { rendu: 0, anim: 0, meta: 0, autre: 0 }; const autres = [];
let tot = 0;
for (const [t, n] of Object.entries(count)) {
  tot += n;
  if (rendered(t)) cat.rendu += n;
  else if (ANIM.has(t)) cat.anim += n;
  else if (META.test(t)) cat.meta += n;
  else { cat.autre += n; autres.push([t, n]); }
}
const pc = (n) => (100 * n / tot).toFixed(1) + '%';
console.log(`=== COUVERTURE (${dirs.length} deck(s), ${nSlides} diapos, ${tot} occurrences de tags) ===\n`);
console.log('  RENDU      (lu par pptx-render.mjs)        : ' + String(cat.rendu).padStart(6) + '  ' + pc(cat.rendu));
console.log('  ANIMATION  (lu par extractTiming/player)   : ' + String(cat.anim).padStart(6) + '  ' + pc(cat.anim));
console.log('  MÉTADONNÉES(sans effet visuel, ignoré)     : ' + String(cat.meta).padStart(6) + '  ' + pc(cat.meta));
console.log('  AUTRE      (trous potentiels)              : ' + String(cat.autre).padStart(6) + '  ' + pc(cat.autre));
console.log('\n  COMPRIS (rendu+anim+méta) : ' + pc(cat.rendu + cat.anim + cat.meta) + '   |   trous : ' + pc(cat.autre));
console.log('\n=== "AUTRE" (trous potentiels) par occurrences ===');
for (const [t, n] of autres.sort((a, b) => b[1] - a[1]).slice(0, 30)) console.log('  ' + String(n).padStart(5) + '  ' + t);
