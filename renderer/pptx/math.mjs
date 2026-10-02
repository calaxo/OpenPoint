// ═══ [PPTX · OFFICE MATH] ═══ OMML (m:*) -> HTML simple pour le foreignObject du texte.
import { esc, tagOf, at, kids, byTag, first } from './xml.mjs';
// ████ §D ─ OFFICE MATH (OMML) -> HTML ──────────────────────────────────────────────────────────
// Les équations de l'éditeur PowerPoint sont stockées en OMML (balises m:*), en général sous
// <mc:AlternateContent><mc:Choice Requires="a14"><… a14:m><m:oMathPara>…. Il n'y a pas de rendu
// MathML ici : on produit du HTML/CSS simple (inline-flex pour les fractions, sup/sub, bordures
// pour la barre de fraction et le radical) inséré dans le foreignObject du paragraphe.
// Chaîne d'appel : renderText (chemin « paraO ») -> renderMath(nœud) -> mathTypeset(texte, sty).
// OMML (§22.1 Office Math) -> HTML dans le foreignObject. Noeud ORDONNÉ (preserveOrder).
// Gère fraction (m:f), parenthèses/délimiteurs (m:d), exposant/indice (m:sSup/m:sSub),
// racine (m:rad), n-aire (m:nary, somme/intégrale). Le math est en italique (convention).
// texte brut d'un noeud, SANS normalisation (pour lire le style des alphanumériques mathématiques)
function rawText(node) {
  if (node['#text'] !== undefined) return String(node['#text']);
  return kids(node).map(rawText).join('');
}
// Typographie mathématique comme PowerPoint (OMML) : lettres en italique, chiffres et ponctuation droits,
// espace autour des relations (0,28 em) et des opérateurs binaires (0,22 em). Les « alphanumériques
// mathématiques » Unicode (𝑷 𝒅 𝟓…) donnent leur style (gras / italique) puis sont ramenés en lettres
// normales (NFKC) — la police n'a pas ces glyphes. sty : style imposé par m:sty (p, b, i, bi).
// MATH_REL : relations (espacées de 0,28 em de chaque côté) ; MATH_BIN : opérateurs binaires (0,22 em).
// Ces valeurs correspondent aux espaces « thick » (5/18 em) et « medium » (4/18 em) de la
// typographie mathématique classique (TeX).
const MATH_REL = new Set([...'=<>≤≥≈≠≅∼∝≡→←⇒⇔']), MATH_BIN = new Set([...'+−-×÷±·∓']);
// mathTypeset(text, sty) -> HTML. Découpe le texte en « sous-runs » de même style (gras/italique)
// pour limiter le nombre de <span>. Chaque opérateur est isolé dans son propre <span> espacé.
// sty (m:sty) : 'p' = droit, 'b' = gras droit, 'i' = italique, 'bi' = gras italique, null = règle
// par défaut OMML (lettres en italique, le reste droit).
export function mathTypeset(text, sty) {
  let out = '', runStyle = null, runText = '';
  // flush : émet le sous-run en cours (avec ou sans <span style>)
  const flush = () => { if (runText) out += runStyle ? `<span style="${runStyle}">${esc(runText)}</span>` : esc(runText); runText = ''; };
  // push : ajoute un caractère ; change de sous-run si le style diffère
  const push = (style, c) => { if (style !== runStyle) { flush(); runStyle = style; } runText += c; };
  for (const ch of text) { // for…of itère par point de code (les U+1D4xx sont des paires de substitution)
    const cp = ch.codePointAt(0);
    let bold = sty === 'b' || sty === 'bi', ital = null; // ital null = « non imposé par le glyphe »
    if (cp >= 0x1D400 && cp <= 0x1D7FF) { // bloc Mathematical Alphanumeric Symbols
      // Disposition Unicode du bloc (latin, 52 lettres = 0x34 par variante) :
      //   o 0x00-0x33 gras · 0x34-0x67 italique · 0x68-0x9B gras italique · au-delà : script,
      //   fraktur, double barre… (ni gras ni italique ici). Grec : 0x1D6A8-6E1 gras,
      //   0x1D6E2-71B italique, 0x1D71C-755 gras italique. Chiffres gras : 0x1D7CE-7D7.
      const o = cp - 0x1D400;
      if (o < 0x34 || (o >= 0x68 && o < 0x9C) || (cp >= 0x1D6A8 && cp <= 0x1D6E1) || (cp >= 0x1D71C && cp <= 0x1D755) || (cp >= 0x1D7CE && cp <= 0x1D7D7)) bold = true;
      ital = (o >= 0x34 && o < 0x9C) || (cp >= 0x1D6E2 && cp <= 0x1D755);
    }
    // NFKC ramène 𝑷 -> P, 𝜌 -> ρ, 𝟓 -> 5 (la police n'a pas les glyphes du bloc 1D400)
    const c = ch.normalize('NFKC');
    if (MATH_REL.has(c) || MATH_BIN.has(c)) { // opérateur : son propre bloc espacé
      flush(); runStyle = null;
      out += `<span style="padding:0 ${MATH_REL.has(c) ? '.28' : '.22'}em">${esc(c === '-' ? '−' : c)}</span>`; continue;
    }
    // italique : forcé droit par m:sty p/b ; sinon style porté par le glyphe ; sinon « lettre = italique »
    // (convention OMML : variables en italique, chiffres/ponctuation droits)
    const letter = /\p{L}/u.test(c);
    const it = sty === 'p' || sty === 'b' ? false : ital ?? letter;
    // style vide -> '' || null -> null = texte sans <span>, fusionnable avec ses voisins
    push((it ? 'font-style:italic;' : '') + (bold ? 'font-weight:700;' : '') || null, c);
  }
  flush();
  return out;
}
// renderMath(nœud OMML ordonné) -> HTML. Descente récursive sur la structure OMML :
//   m:r (run : m:rPr + m:t) · m:f (fraction m:num / m:den) · m:d (délimiteurs, m:dPr begChr/endChr,
//   défauts « ( » et « ) ») · m:sSup / m:sSub (base m:e + exposant/indice) · m:rad (racine, l'indice
//   de degré m:deg est ignoré) · m:nary (opérateur n-aire m:naryPr/m:chr, défaut ∫, bornes sub/sup).
// Toute autre balise (m:oMath, m:oMathPara, m:e, m:num…) est un simple conteneur : on concatène ses
// enfants. Les balises *Pr (propriétés) ne produisent rien.
// Non gérés (rendus « à plat » par le default) : matrices m:m, accents m:acc, barres m:bar, boîtes,
// m:sSubSup, limites m:limLow/m:limUpp, etc.
function renderMath(node) {
  if (node['#text'] !== undefined) return esc(String(node['#text']).normalize('NFKC'));
  const t = tagOf(node);
  if (/Pr$/.test(t)) return ''; // m:fPr, m:dPr, m:rPr, m:ctrlPr, m:naryPr... = propriétés
  switch (t) {
    case 'm:r': { // run mathématique : style m:sty (p droit, b gras, i italique, bi) puis texte typographié
      const rp = first(node, 'm:rPr');
      const sty = at(first(rp || {}, 'm:sty'))['@m:val'] ?? (first(rp || {}, 'm:nor') ? 'p' : null);
      return byTag(node, 'm:t').map((t) => mathTypeset(rawText(t), sty)).join('');
    }
    case 'm:t': return mathTypeset(rawText(node), null);
    // fraction : colonne flex num/den, barre = bordure basse du numérateur (0,06 em)
    case 'm:f': {
      const num = renderMath(first(node, 'm:num') || {}), den = renderMath(first(node, 'm:den') || {});
      return `<span style="display:inline-flex;flex-direction:column;vertical-align:middle;text-align:center;line-height:1.05;font-size:.95em">`
        + `<span style="border-bottom:.06em solid;padding:0 .25em">${num}</span><span style="padding:0 .25em">${den}</span></span>`;
    }
    case 'm:d': {
      const dPr = first(node, 'm:dPr');
      const beg = at(first(dPr || {}, 'm:begChr'))['@m:val'] ?? '(', end = at(first(dPr || {}, 'm:endChr'))['@m:val'] ?? ')';
      return esc(beg) + byTag(node, 'm:e').map(renderMath).join('') + esc(end);
    }
    case 'm:sSup': return renderMath(first(node, 'm:e') || {}) + `<sup style="font-size:.7em">${renderMath(first(node, 'm:sup') || {})}</sup>`;
    case 'm:sSub': return renderMath(first(node, 'm:e') || {}) + `<sub style="font-size:.7em">${renderMath(first(node, 'm:sub') || {})}</sub>`;
    // racine : signe √ + barre supérieure (border-top) sur le radicande m:e
    case 'm:rad': return `√<span style="border-top:.06em solid;padding:0 .15em">${renderMath(first(node, 'm:e') || {})}</span>`;
    case 'm:nary': {
      const op = at(first(first(node, 'm:naryPr') || {}, 'm:chr'))['@m:val'] ?? '∫';
      const sub = renderMath(first(node, 'm:sub') || {}), sup = renderMath(first(node, 'm:sup') || {});
      return `<span style="font-size:1.3em">${esc(op)}</span>${sub ? `<sub style="font-size:.7em">${sub}</sub>` : ''}${sup ? `<sup style="font-size:.7em">${sup}</sup>` : ''}${renderMath(first(node, 'm:e') || {})}`;
    }
    default: return kids(node).map(renderMath).join(''); // m:oMath, m:oMathPara, m:e, m:r, m:num, m:den...
  }
}
export { renderMath };
