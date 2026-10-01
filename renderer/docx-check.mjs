// Contrôle batch DOCX : rend chaque .docx d'un dossier, rapporte crashes + métriques (texte/img/tbl).
// Usage : node renderer/docx-check.mjs "specimen"
import { renderDocx } from './docx-build.mjs';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync } from 'fflate';

const dir = process.argv[2] || 'specimen';
const files = readdirSync(dir).filter((f) => f.toLowerCase().endsWith('.docx') && !f.startsWith('~$'));
let errs = 0;
for (const f of files) {
  const p = join(dir, f);
  // nb de w:t/w:delText dans la source (attendu) pour repérer le texte perdu
  let srcTxt = 0, srcTbl = 0, srcImg = 0;
  try {
    const zip = unzipSync(new Uint8Array(readFileSync(p)));
    const d = Buffer.from(zip['word/document.xml']).toString('utf8');
    srcTxt = (d.match(/<w:t[ >]/g) || []).length + (d.match(/<w:delText[ >]/g) || []).length;
    srcTbl = (d.match(/<w:tbl>/g) || []).length;
    srcImg = (d.match(/r:embed=/g) || []).length;
  } catch (e) { console.log(`${f}\n  !! zip: ${e.message}`); errs++; continue; }
  try {
    const t0 = Date.now();
    const { html } = renderDocx(p);
    const txt = (html.replace(/<[^>]+>/g, '').replace(/&[^;]+;/g, ' ').match(/\S/g) || []).length;
    const imgs = (html.match(/<img/g) || []).length;
    const tbls = (html.match(/<table/g) || []).length;
    const ms = Date.now() - t0;
    const warn = [];
    if (srcTxt > 20 && txt < srcTxt * 0.3) warn.push(`TEXTE FAIBLE (${txt} rendu / ~${srcTxt} runs src)`);
    if (srcTbl && !tbls) warn.push(`TABLEAUX absents (${srcTbl} src)`);
    console.log(`${f}  (${(readFileSync(p).length/1024|0)}Ko, ${ms}ms)  texte=${txt} img=${imgs}/${srcImg} tbl=${tbls}/${srcTbl}${warn.length ? '\n  ⚠ ' + warn.join(' | ') : ''}`);
  } catch (e) { console.log(`${f}\n  !! ERREUR: ${e.message}`); errs++; }
}
console.log(`\nTOTAL : ${files.length} docx, ${errs} erreurs`);
