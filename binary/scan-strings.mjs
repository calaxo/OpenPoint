// Extrait les chaînes (ASCII + UTF-16LE) d'un binaire et cherche le vocabulaire du
// moteur d'animation. But : localiser DANS QUEL module natif vit l'interpréteur <p:timing>.
// Si les tokens du format (ppt_x, animEffect, mainSeq...) sont des littéraux du binaire,
// c'est que le code qui les compare/dispatch est là.
//
// Usage : node binary/scan-strings.mjs <chemin.dll> [token1 token2 ...]

import { readFileSync } from 'node:fs';

const file = process.argv[2];
const tokens = process.argv.slice(3).length
  ? process.argv.slice(3)
  : ['ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'ppt_r', 'xshear', 'style.visibility',
     'style.rotation', 'style.opacity', 'animEffect', 'clickEffect', 'withEffect',
     'afterEffect', 'mainSeq', 'presetClass', 'presetID', 'fillcolor', 'ppt_c',
     'onstopaudio', 'onprev', 'onnext', 'tmRoot'];

const buf = readFileSync(file);

// index -> offset(s) pour chaque token, en ascii et utf16le
function findAll(hay, needleBytes) {
  const offs = [];
  let i = 0;
  outer: while ((i = indexOfBytes(hay, needleBytes, i)) !== -1) {
    offs.push(i);
    i += needleBytes.length;
    if (offs.length > 50) break;
  }
  return offs;
}
function indexOfBytes(hay, needle, from) {
  for (let i = from; i <= hay.length - needle.length; i++) {
    let ok = true;
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) { ok = false; break; }
    if (ok) return i;
  }
  return -1;
}
const asciiBytes = (s) => Buffer.from(s, 'latin1');
const utf16Bytes = (s) => Buffer.from(s, 'utf16le');

console.log(`Fichier : ${file}  (${(buf.length / 1024 / 1024).toFixed(1)} Mo)\n`);
console.log('token'.padEnd(20), 'ascii', 'utf16', 'premiers offsets (hex)');
for (const t of tokens) {
  const a = findAll(buf, asciiBytes(t));
  const w = findAll(buf, utf16Bytes(t));
  const off = (a[0] ?? w[0]);
  const src = a.length ? a : w;
  console.log(
    t.padEnd(20),
    String(a.length).padStart(5),
    String(w.length).padStart(5),
    src.slice(0, 4).map((o) => '0x' + o.toString(16)).join(' ')
  );
}
