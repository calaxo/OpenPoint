// Capture d'écran d'un fichier HTML (viewers) avec puppeteer.
// Usage : node tools/screenshot.mjs <fichier.html[#ancre]> <préfixeSortie> [actions...] [--size=1280x720]
// Actions (exécutées dans l'ordre) :
//   wait:500            attendre 500 ms
//   key:ArrowRight      appuyer sur une touche
//   click               cliquer au centre de la page   (click:<selecteur> : cliquer sur un élément)
//   shot:nom            capture de la fenêtre -> <préfixe>-nom.png
//   shotFull:nom        capture de toute la page
//   shotEl:<sel>:nom    capture d'un élément (le sélecteur peut contenir ':' sauf le dernier segment)
//   scroll:<sel>        faire défiler jusqu'à l'élément
// Sans action : une seule capture <préfixe>.png. Chrome : env CHROME_PATH, sinon Chrome installé
// (C:/Program Files/Google/Chrome/Application/chrome.exe), sinon celui de puppeteer.
// Code de sortie 1 si une erreur de page (exception JS, console.error, requête échouée) est survenue.
import { launchBrowser } from './browser.mjs';
import { mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { pathToFileURL } from 'node:url';

const argv = process.argv.slice(2);
const sizeArg = argv.find((a) => a.startsWith('--size='));
const [width, height] = (sizeArg ? sizeArg.slice(7) : '1280x720').split('x').map(Number);
const pos = argv.filter((a) => !a.startsWith('--size='));
const [file, prefix, ...actions] = pos;
if (!file || !prefix) { console.error('usage: node tools/screenshot.mjs <fichier.html> <préfixeSortie> [actions...] [--size=WxH]'); process.exit(2); }

mkdirSync(dirname(resolve(prefix)), { recursive: true });

const browser = await launchBrowser();
let errors = 0;
try {
  const page = await browser.newPage();
  await page.setViewport({ width, height });
  page.on('pageerror', (e) => { errors++; console.error('[pageerror]', e.message); });
  page.on('console', (m) => { if (m.type() === 'error') { errors++; console.error('[console.error]', m.text()); } });
  page.on('requestfailed', (r) => { errors++; console.error('[requestfailed]', r.url().slice(0, 120)); });
  // « fichier.html#3 » : l'ancre est ajoutée à l'URL file:// (deep-link des viewers)
  const hashAt = file.indexOf('#');
  const url = pathToFileURL(resolve(hashAt < 0 ? file : file.slice(0, hashAt))).href + (hashAt < 0 ? '' : file.slice(hashAt));
  await page.goto(url, { waitUntil: 'load' });
  const list = actions.length ? actions : ['shot:'];
  for (const a of list) {
    const [cmd, ...rest] = a.split(':');
    const arg = rest.join(':');
    if (cmd === 'wait') await new Promise((r) => setTimeout(r, +arg || 0));
    else if (cmd === 'key') await page.keyboard.press(arg);
    else if (cmd === 'click') { if (arg) await page.click(arg); else await page.mouse.click(width / 2, height / 2); }
    else if (cmd === 'scroll') await page.$eval(arg, (el) => el.scrollIntoView());
    else if (cmd === 'shot' || cmd === 'shotFull') {
      const out = resolve(arg ? `${prefix}-${arg}.png` : `${prefix}.png`);
      await page.screenshot({ path: out, fullPage: cmd === 'shotFull' }); console.log('shot', out);
    } else if (cmd === 'shotEl') {
      const i = arg.lastIndexOf(':'); const sel = i < 0 ? arg : arg.slice(0, i), name = i < 0 ? 'el' : arg.slice(i + 1);
      const el = await page.$(sel); if (!el) { errors++; console.error('[shotEl] introuvable :', sel); continue; }
      const out = resolve(`${prefix}-${name}.png`); await el.screenshot({ path: out }); console.log('shot', out);
    } else { console.error('action inconnue :', a); errors++; }
  }
} finally { await browser.close(); }
process.exit(errors ? 1 : 0);
