// Point d'entrée UNIFIÉ Office -> HTML. Détecte le type et dispatche vers le bon moteur.
//   .pptx -> build.mjs      (PresentationML : diapos, formes positionnées, animations)  -> viewer.html
//   .docx -> docx-build.mjs (WordprocessingML : texte en flux)                            -> docx-out.html
// Usage : node renderer/render.mjs "C:\chemin\fichier.pptx|docx" [args build...]
import { execFileSync } from 'node:child_process';
import { dirname, join, extname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const file = process.argv[2];
if (!file) { console.error('usage: node renderer/render.mjs <fichier.pptx|.docx> [options]'); process.exit(1); }
const ext = extname(file).toLowerCase();
const rest = process.argv.slice(3);

const engines = {
  '.pptx': { script: 'build.mjs', out: 'viewer.html' },
  '.docx': { script: 'docx-build.mjs', out: 'docx-out.html' },
};
const e = engines[ext];
if (!e) { console.error(`Format non supporté : ${ext}. Attendu : .pptx ou .docx`); process.exit(1); }

console.log(`[render] ${basename(file)} -> moteur ${ext === '.pptx' ? 'PPTX (PresentationML)' : 'DOCX (WordprocessingML)'}`);
execFileSync('node', [join(HERE, e.script), file, ...rest], { stdio: 'inherit' });
console.log(`[render] -> renderer/${e.out}`);
