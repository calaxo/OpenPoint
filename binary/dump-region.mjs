// Dump toutes les chaînes UTF-16LE (et ASCII) lisibles dans une plage d'offsets d'un binaire.
// Sert à reconstruire les tables de données du moteur (recettes de presets, langage
// d'expressions de keyframes) telles qu'elles sont compilées dans PPCORE.DLL.
//
// Usage : node binary/dump-region.mjs <fichier> <offsetHex> <longueur>

import { readFileSync } from 'node:fs';

const [file, offHex, lenStr] = process.argv.slice(2);
const start = parseInt(offHex, 16);
const len = parseInt(lenStr ?? '0x400', lenStr?.startsWith('0x') ? 16 : 10);
const buf = readFileSync(file).subarray(start, start + len);

// extrait les runs UTF-16LE imprimables (>=2 chars) et ASCII (>=3)
const out = [];
let i = 0;
while (i < buf.length - 1) {
  // tente UTF-16LE
  let j = i, s = '';
  while (j < buf.length - 1) {
    const lo = buf[j], hi = buf[j + 1];
    if (hi === 0 && lo >= 0x20 && lo < 0x7f) { s += String.fromCharCode(lo); j += 2; }
    else break;
  }
  if (s.length >= 2) { out.push(['0x' + (start + i).toString(16), 'w', s]); i = j; continue; }
  // tente ASCII
  j = i; s = '';
  while (j < buf.length && buf[j] >= 0x20 && buf[j] < 0x7f) { s += String.fromCharCode(buf[j]); j++; }
  if (s.length >= 3) { out.push(['0x' + (start + i).toString(16), 'a', s]); i = j; continue; }
  i++;
}
for (const [o, enc, s] of out) console.log(o.padEnd(12), enc, JSON.stringify(s));
