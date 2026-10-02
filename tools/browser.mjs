// Lance Chrome (puppeteer) pour les outils : screenshot, e2e-viewers, sweep.
// Ordre : variable d'environnement CHROME_PATH, puis Chrome installé (Windows / macOS / Linux), puis le
// Chrome téléchargé par puppeteer (s'il existe).
import { existsSync } from 'node:fs';
import puppeteer from 'puppeteer';

const KNOWN = [
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  '/usr/bin/google-chrome',
  '/usr/bin/chromium',
];

export const chromePath = () => process.env.CHROME_PATH || KNOWN.find((p) => existsSync(p));

export function launchBrowser(extraArgs = []) {
  return puppeteer.launch({ executablePath: chromePath(), headless: true, args: ['--no-sandbox', '--allow-file-access-from-files', ...extraArgs] });
}
