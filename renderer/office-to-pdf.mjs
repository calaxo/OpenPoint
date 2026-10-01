// Rendu FIDÈLE d'un document Office via LibreOffice (vrai moteur de mise en page) -> PDF.
// Pour les mises en page complexes (pages de garde, pagination exacte, TOC stylée) que le rendu HTML
// en flux (docx-build.mjs) ne peut pas reproduire. LibreOffice gère .docx ET .pptx.
//   Rendu HTML en flux (éditable/LMS léger)  : docx-build.mjs  / build.mjs (pptx)
//   Rendu PDF FIDÈLE (affichage pixel-perfect): office-to-pdf.mjs  <-- ici
// Le PPTX ANIMÉ reste le domaine du moteur JS maison (build-player.mjs) — LibreOffice fige les animations.
//
// Usage : node renderer/office-to-pdf.mjs "C:\chemin\fichier.docx|pptx" [dossierSortie]
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { dirname, join, basename, extname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
// emplacements usuels de LibreOffice (Windows). Adapter si besoin.
const CANDIDATES = [
  'C:/Program Files/LibreOffice/program/soffice.exe',
  'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
  'soffice',
];
const SOFFICE = CANDIDATES.find((p) => p === 'soffice' || existsSync(p));

const file = process.argv[2];
const outDir = process.argv[3] || join(ROOT, 'renderer');
if (!file) { console.error('usage: node renderer/office-to-pdf.mjs <fichier.docx|.pptx> [outDir]'); process.exit(1); }
if (!SOFFICE) { console.error('LibreOffice introuvable (soffice.exe). Installe-le ou ajuste CANDIDATES.'); process.exit(1); }
const ext = extname(file).toLowerCase();
if (ext !== '.docx' && ext !== '.pptx') { console.error(`Format non supporté : ${ext}`); process.exit(1); }

// profil dédié (évite les conflits avec une instance LibreOffice ouverte)
const profile = join(outDir, '.lo-profile');
try {
  execFileSync(SOFFICE, [
    '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', outDir,
    `-env:UserInstallation=file:///${profile.replace(/\\/g, '/')}`,
    file,
  ], { stdio: 'inherit', timeout: 120000 });
} catch (e) { /* soffice émet un avertissement bénin sur stderr ; le PDF est quand même produit */ }
const pdf = join(outDir, basename(file).replace(/\.(docx|pptx)$/i, '.pdf'));
if (existsSync(pdf)) console.log(`OK -> ${pdf}`);
else { console.error('Échec : PDF non produit (une instance LibreOffice est-elle déjà ouverte ?)'); process.exit(1); }
