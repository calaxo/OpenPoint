// ═══ Assemblage d'une page HTML AUTONOME (un seul fichier) pour les viewers ═══
// Les scripts et feuilles de style du runtime vivent dans renderer/viewer/ (vrais fichiers, lintés) ;
// ce module les lit et les INLINE au moment du build, avec les données du deck en JSON.
//
// API :
//   buildPage({ title, css: ['viewer.css'], scripts: ['common.js', 'scroll-viewer.js'],
//               headCss: '<css additionnel, ex. polices>', body: '<html du body>', data: {…} }) -> string
//   - css / scripts : noms de fichiers relatifs à renderer/viewer/, inlinés dans cet ordre ;
//   - data : objet sérialisé dans <script type="application/json" id="op-data"> (lu par OP.readData()).
//   jsonForScript(obj) -> JSON sûr à placer dans une balise <script> (aucun « < » brut).
import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

export const VIEWER_DIR = join(dirname(fileURLToPath(import.meta.url)), '..', 'viewer');

// « < » -> \u003c : empêche toute fermeture prématurée de </script> ou <!-- dans les données.
// (JSON.parse relit \u003c comme « < » : les données ne changent pas.)
export function jsonForScript(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c');
}

export function escHtml(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

const readViewerFile = (f) => readFileSync(join(VIEWER_DIR, f), 'utf8');

export function buildPage({ title, css = [], scripts = [], headCss = '', body = '', data = {} }) {
  const styles = css.map(readViewerFile).join('\n');
  // un script inliné ne doit pas contenir « </script » (les fichiers du runtime n'en ont pas)
  const code = scripts.map((f) => `<script>\n${readViewerFile(f)}\n</script>`).join('\n');
  return `<!DOCTYPE html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>${escHtml(title)}</title>
<style>${headCss}</style>
<style>${styles}</style>
</head><body>
${body}
<script type="application/json" id="op-data">${jsonForScript(data)}</script>
${code}
</body></html>
`;
}
