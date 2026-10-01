// ═══ [CŒUR DOCX] ═══ docs/ARCHITECTURE.md (§5) · AUTONOME (aucun lien avec pptx-render.mjs) · pagination client
// Moteur DOCX (WordprocessingML) -> HTML en flux continu. PoC : texte/styles/listes/tableaux/images.
// NE TOUCHE PAS à pptx-render.mjs (moteur PPTX). Réutilise fflate + fast-xml-parser.
//
// Usage : node renderer/docx-build.mjs "C:\chemin\doc.docx"  -> renderer/docx-out.html
//
// Modèle Word (vs PPTX) : word/document.xml <w:body> -> flux de <w:p> (paragraphes) / <w:tbl> (tableaux)
// / <w:sectPr> (page). Pas de positionnement X/Y : le texte COULE. Styles dans word/styles.xml,
// listes dans word/numbering.xml, images dans word/media via word/_rels/document.xml.rels.
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, basename, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';
import { XMLParser } from 'fast-xml-parser';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const P = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', preserveOrder: true, parseTagValue: false, parseAttributeValue: false, trimValues: false });
const PF = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', parseTagValue: false, parseAttributeValue: false, trimValues: false });

// --- helpers preserveOrder ---
const tagOf = (n) => { for (const k in n) if (k !== ':@' && k !== '#text') return k; return null; };
const kids = (n) => n[tagOf(n)] || [];
const at = (n) => (n && n[':@']) || {};
const arr = (x) => x === undefined ? [] : Array.isArray(x) ? x : [x];
// fusion « le dernier défini gagne » en ignorant les valeurs null/undefined (héritage docDefaults->style->inline)
const mrg = (...objs) => { const o = {}; for (const x of objs) if (x) for (const k in x) if (x[k] != null) o[k] = x[k]; return o; };
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// cherche le 1er enfant de tag donné
const child = (n, t) => kids(n).find((c) => tagOf(c) === t);
const childAttrs = (n, t) => { const c = child(n, t); return c ? at(c) : null; };
// valeur d'un enfant "toggle/val" : <w:b/> ou <w:b w:val="true"/> ; absent => null
function onOff(pr, t) {
  const c = child(pr, t); if (!c) return null;
  const v = at(c)['@w:val'];
  return (v === undefined || v === 'true' || v === '1' || v === 'on') ? true : false;
}
const twip = (v) => (+v || 0) / 15;   // twips -> px (1440/pouce, 96px/pouce)
const emuPx = (v) => (+v || 0) / 9525; // EMU -> px

// --- polices : substituts métrique-compatibles (comme le moteur PPTX) ---
const FONT_SUBST = {
  'calibri': "'Calibri','Carlito',sans-serif", 'cambria': "'Cambria','Caladea',serif",
  'arial': "'Arial','Liberation Sans',sans-serif", 'times new roman': "'Times New Roman','Liberation Serif',serif",
  'courier new': "'Courier New','Liberation Mono',monospace", 'verdana': "'Verdana',sans-serif",
  'georgia': "'Georgia',serif", 'tahoma': "'Tahoma',sans-serif",
};
const SERIF_RE = /times|cambria|georgia|garamond|minion|book|antiqua/i;
// puces en police SYMBOLE (Wingdings/Webdings/Symbol, souvent en zone privée PUA 0xF000+) -> Unicode
// (rendu en police normale = lisible partout, comme le moteur PPTX). Inconnu -> « • ».
const SYMB_RE = /wingding|webding|symbol/i;
const WINGDINGS_BULLET = { 0x6C: '●', 0x6D: '○', 0x6E: '■', 0x6F: '□', 0x71: '❑', 0xA7: '▪', 0xA8: '▫', 0xB7: '•', 0xD8: '➢', 0xD9: '➢', 0xFC: '✔', 0xFD: '✔' };
function bulletChar(ch, font) {
  if (!ch) return '•';
  let cp = ch.codePointAt(0);
  const pua = cp >= 0xF000 && cp <= 0xF0FF;
  if (pua) cp -= 0xF000;
  if (pua || SYMB_RE.test(font || '')) return WINGDINGS_BULLET[cp] || '•';
  return ch; // déjà un caractère normal (•, -, o…)
}
function fontFamily(f) {
  if (!f) return 'inherit';
  const k = f.toLowerCase().trim();
  if (FONT_SUBST[k]) return FONT_SUBST[k];
  return SERIF_RE.test(f) ? `'${esc(f)}',serif` : `'${esc(f)}',sans-serif`;
}

// --- couleur de thème (word/theme/theme1.xml) : accent1..6, dk1/lt1... ---
function loadTheme(files) {
  const raw = files['word/theme/theme1.xml']; if (!raw) return {};
  try {
    const t = PF.parse(raw)['a:theme']['a:themeElements']['a:clrScheme'];
    const out = {}; const map = { 'a:dk1': 'dk1', 'a:lt1': 'lt1', 'a:dk2': 'dk2', 'a:lt2': 'lt2', 'a:accent1': 'accent1', 'a:accent2': 'accent2', 'a:accent3': 'accent3', 'a:accent4': 'accent4', 'a:accent5': 'accent5', 'a:accent6': 'accent6', 'a:hlink': 'hlink' };
    for (const [k, name] of Object.entries(map)) {
      const c = t[k]; if (!c) continue;
      out[name] = c['a:srgbClr']?.['@val'] ?? c['a:sysClr']?.['@lastClr'] ?? null;
    }
    return out;
  } catch { return {}; }
}
function resolveColor(val, themeColor, theme) {
  if (themeColor) { const m = { dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hyperlink: 'hlink' }; const h = theme[m[themeColor] || themeColor]; if (h) return '#' + h; }
  if (val && val !== 'auto') return '#' + val;
  return null;
}

// --- styles.xml : styleId -> { rPr, pPr, basedOn, name, outline } ---
function loadStyles(files) {
  const raw = files['word/styles.xml']; const styles = {}; let docDefaults = { rPr: {}, pPr: {} };
  if (!raw) return { styles, docDefaults };
  const tree = P.parse(raw);
  const root = tree.find((n) => tagOf(n) === 'w:styles'); if (!root) return { styles, docDefaults };
  for (const s of kids(root)) {
    const t = tagOf(s);
    if (t === 'w:docDefaults') {
      const rpd = child(child(s, 'w:rPrDefault') || {}, 'w:rPr');
      const ppd = child(child(s, 'w:pPrDefault') || {}, 'w:pPr');
      docDefaults = { rPr: rpd ? parseRPr(rpd, {}) : {}, pPr: ppd ? parsePPr(ppd) : {} };
    } else if (t === 'w:style') {
      const a = at(s); const id = a['@w:styleId']; if (!id) continue;
      const rPr = child(s, 'w:rPr'), pPr = child(s, 'w:pPr');
      styles[id] = {
        type: a['@w:type'],
        basedOn: childAttrs(s, 'w:basedOn')?.['@w:val'] || null,
        name: childAttrs(s, 'w:name')?.['@w:val'] || id,
        rPrRaw: rPr || null, pPrRaw: pPr || null,
      };
    }
  }
  return { styles, docDefaults };
}

// parse rPr -> objet propriétés (résolu plus tard avec héritage)
function parseRPr(rPr, theme) {
  if (!rPr) return {};
  const o = {};
  const b = onOff(rPr, 'w:b'); if (b !== null) o.bold = b;
  const i = onOff(rPr, 'w:i'); if (i !== null) o.italic = i;
  const st = onOff(rPr, 'w:strike'); if (st !== null) o.strike = st;
  const u = childAttrs(rPr, 'w:u'); if (u) o.underline = u['@w:val'] !== 'none';
  const sz = childAttrs(rPr, 'w:sz'); if (sz) o.sizePt = (+sz['@w:val']) / 2; // demi-points
  const col = childAttrs(rPr, 'w:color'); if (col) o.color = resolveColor(col['@w:val'], col['@w:themeColor'], theme);
  const hl = childAttrs(rPr, 'w:highlight'); if (hl) o.highlight = hl['@w:val'];
  const rf = childAttrs(rPr, 'w:rFonts'); if (rf) o.font = rf['@w:ascii'] || rf['@w:hAnsi'] || rf['@w:cs'];
  const va = childAttrs(rPr, 'w:vertAlign'); if (va) o.vertAlign = va['@w:val'];
  const caps = onOff(rPr, 'w:caps'); if (caps) o.caps = true;
  return o;
}
function parsePPr(pPr) {
  if (!pPr) return {};
  const o = {};
  const jc = childAttrs(pPr, 'w:jc'); if (jc) o.align = jc['@w:val'];
  o.style = childAttrs(pPr, 'w:pStyle')?.['@w:val'] || null;
  const numPr = child(pPr, 'w:numPr');
  if (numPr) o.num = { ilvl: +(childAttrs(numPr, 'w:ilvl')?.['@w:val'] || 0), numId: childAttrs(numPr, 'w:numId')?.['@w:val'] };
  const sp = childAttrs(pPr, 'w:spacing');
  if (sp) o.spacing = { before: sp['@w:before'], after: sp['@w:after'], line: sp['@w:line'], lineRule: sp['@w:lineRule'], beforeAuto: sp['@w:beforeAutospacing'], afterAuto: sp['@w:afterAutospacing'] };
  const ind = childAttrs(pPr, 'w:ind');
  if (ind) o.ind = { left: ind['@w:left'] ?? ind['@w:start'], right: ind['@w:right'] ?? ind['@w:end'], firstLine: ind['@w:firstLine'], hanging: ind['@w:hanging'] };
  o.outline = childAttrs(pPr, 'w:outlineLvl')?.['@w:val'];
  const rPr = child(pPr, 'w:rPr'); if (rPr) o.rPrRaw = rPr; // rPr de fin de paragraphe (défaut runs)
  return o;
}
// résout la chaîne d'héritage de style (basedOn) -> rPr/pPr cumulés
function resolveStyle(id, styles, theme, seen = new Set()) {
  if (!id || !styles[id] || seen.has(id)) return { rPr: {}, pPr: {} };
  seen.add(id);
  const s = styles[id];
  const base = resolveStyle(s.basedOn, styles, theme, seen);
  return {
    rPr: { ...base.rPr, ...parseRPr(s.rPrRaw, theme) },
    pPr: { ...base.pPr, ...parsePPr(s.pPrRaw) },
    name: s.name,
  };
}

// --- numbering.xml : (numId, ilvl) -> format (bullet/decimal) ---
function loadNumbering(files) {
  const raw = files['word/numbering.xml']; const abstract = {}, num = {};
  const dflt = { fmt: 'bullet', ch: '•' };
  if (!raw) return { lvl: () => dflt };
  const tree = P.parse(raw); const root = tree.find((n) => tagOf(n) === 'w:numbering'); if (!root) return { lvl: () => dflt };
  for (const n of kids(root)) {
    const t = tagOf(n);
    if (t === 'w:abstractNum') {
      const id = at(n)['@w:abstractNumId']; abstract[id] = {};
      for (const lvl of kids(n)) if (tagOf(lvl) === 'w:lvl') {
        const il = at(lvl)['@w:ilvl'];
        const fmt = childAttrs(lvl, 'w:numFmt')?.['@w:val'] || 'bullet';
        const ch = childAttrs(lvl, 'w:lvlText')?.['@w:val'] || '•';           // caractère de puce (ou motif %1.)
        const buFont = child(lvl, 'w:rPr') && childAttrs(child(lvl, 'w:rPr'), 'w:rFonts')?.['@w:ascii']; // Wingdings/Symbol ?
        abstract[id][il] = { fmt, ch, buFont };
      }
    } else if (t === 'w:num') {
      const id = at(n)['@w:numId']; num[id] = childAttrs(n, 'w:abstractNumId')?.['@w:val'];
    }
  }
  return { lvl: (numId, ilvl) => { const aid = num[numId]; return (abstract[aid] && (abstract[aid][ilvl] || abstract[aid]['0'])) || dflt; } };
}

// ================= rendu =================
export function renderDocx(docxPath) {
  const files = {};
  const zip = unzipSync(new Uint8Array(readFileSync(docxPath)));
  for (const [p, d] of Object.entries(zip)) if (!p.endsWith('/')) files[p] = Buffer.from(d).toString(p.endsWith('.xml') || p.endsWith('.rels') ? 'utf8' : 'binary');
  const theme = loadTheme(files);
  const { styles, docDefaults } = loadStyles(files);
  const numbering = loadNumbering(files);

  // rels document -> images
  const rels = {};
  const relsRaw = files['word/_rels/document.xml.rels'];
  if (relsRaw) { const r = PF.parse(relsRaw)['Relationships']; for (const rel of arr(r?.['Relationship'])) rels[rel['@Id']] = rel['@Target']; }
  // curRels = rels de la partie en cours (document, ou header/footer qui ont LEURS propres rels pour
  // leurs images). Swappé pendant le rendu d'un header/footer puis restauré.
  let curRels = rels;
  const loadRelsFor = (partPath) => { const rp = partPath.replace(/([^/]+)$/, '_rels/$1.rels'); const raw = files[rp]; const m = {}; if (raw) { const r = PF.parse(raw)['Relationships']; for (const rel of arr(r?.['Relationship'])) m[rel['@Id']] = rel['@Target']; } return m; };
  function imageData(embedId) {
    const tgt = curRels[embedId]; if (!tgt) return null;
    const path = 'word/' + tgt.replace(/^\.\.\//, '').replace(/^\.\//, '');
    const zbin = zip[path]; if (!zbin) return null;
    const ext = extname(path).slice(1).toLowerCase();
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', emf: '', wmf: '' }[ext];
    if (!mime) return null; // emf/wmf non affichables direct
    return `data:${mime};base64,${Buffer.from(zbin).toString('base64')}`;
  }

  const doc = P.parse(files['word/document.xml']);
  const docRoot = doc.find((n) => tagOf(n) === 'w:document');
  const body = child(docRoot, 'w:body');

  // page (dernier sectPr du body)
  let pageW = 794, pageH = 1123, padT = 96, padR = 96, padB = 96, padL = 96; // A4 @96dpi défaut
  const sect = kids(body).find((n) => tagOf(n) === 'w:sectPr');
  if (sect) {
    const pg = childAttrs(sect, 'w:pgSz'); if (pg) { pageW = twip(pg['@w:w']) || pageW; pageH = twip(pg['@w:h']) || pageH; }
    const mg = childAttrs(sect, 'w:pgMar'); if (mg) { padT = twip(mg['@w:top']); padR = twip(mg['@w:right']); padB = twip(mg['@w:bottom']); padL = twip(mg['@w:left']); }
  }

  const counters = {}; // compteurs de listes numérotées par (numId+ilvl)
  // PAGINATION : le doc n'a (souvent) pas de saut de page explicite, mais Word cache la frontière de
  // chaque page dans <w:lastRenderedPageBreak/>. On découpe le flux là-dessus en pages A4 séparées.
  // pendingBreak est levé par renderRun quand il croise un saut ; on ferme la page après le bloc courant.
  let pendingBreak = false;
  let paraHasAbs = false; // levé par renderDrawing quand un flottant est positionné RELATIF au paragraphe
  let curParaRPr = docDefaults.rPr; // props de run héritées du paragraphe courant (docDefaults + style para)
  // flux plat de blocs + marqueurs de saut de page (lastRenderedPageBreak/w:br page). La PAGINATION
  // réelle est faite CÔTÉ CLIENT (mesure des hauteurs -> pages A4 uniformes), en respectant ces marqueurs.
  const flow = [];
  for (const node of kids(body)) {
    if (tagOf(node) === 'w:sectPr') continue; // props de section finale, pas du contenu
    pendingBreak = false;
    const h = renderBlock(node);
    if (h) flow.push(h);
    if (pendingBreak) flow.push('<div class="pgbreak"></div>');
  }

  // en-tête (logos etc.) : le sectPr référence word/headerN.xml (qui a SES propres rels pour ses images).
  // On le rend une fois en haut de la page (flux continu). Priorité type default > first > premier.
  function renderPart(rootTag, partPath) {
    const raw = files[partPath]; if (!raw) return '';
    const root = P.parse(raw).find((n) => tagOf(n) === rootTag); if (!root) return '';
    const saved = curRels; curRels = loadRelsFor(partPath);
    const parts = [];
    for (const node of kids(root)) parts.push(renderBlock(node));
    curRels = saved;
    return parts.join('');
  }
  // bloc générique : paragraphe / tableau / w:sdt (structured tag : TOC, encarts...). Le w:sdt au niveau
  // BLOC était ignoré -> la Table des matières (un w:sdt) ne s'affichait pas. On descend dans sdtContent.
  function renderBlock(node) {
    const t = tagOf(node);
    if (t === 'w:p') return renderP(node);
    if (t === 'w:tbl') return renderTbl(node);
    if (t === 'w:sdt') { const c = child(node, 'w:sdtContent'); return c ? kids(c).map(renderBlock).join('') : ''; }
    return '';
  }
  // les headerReference vivent dans N'IMPORTE quel sectPr (souvent un saut de SECTION, pas le sectPr
  // final du body). On les collecte dans TOUT le document, on essaie default > first > any, et on garde
  // le 1er en-tête qui a un contenu réel (texte ou image) -> évite les en-têtes vides.
  // helper commun en-tête/pied : collecte les références (dans TOUS les sectPr), essaie default>first>
  // even>any, renvoie le HTML du 1er part NON VIDE (texte ou image).
  function pickPart(refTag, rootTag, cls) {
    const refs = [];
    (function walk(n) { for (const c of arr(n)) { if (tagOf(c) === refTag) refs.push(at(c)); const kk = c[tagOf(c)]; if (Array.isArray(kk)) walk(kk); } })(doc);
    const ordered = [...refs.filter((a) => a['@w:type'] === 'default'), ...refs.filter((a) => a['@w:type'] === 'first'), ...refs.filter((a) => a['@w:type'] === 'even'), ...refs];
    const tried = new Set();
    for (const ref of ordered) {
      const tgt = rels[ref['@r:id']]; if (!tgt || tried.has(tgt)) continue; tried.add(tgt);
      const h = renderPart(rootTag, 'word/' + tgt.replace(/^\.\//, ''));
      if (h.replace(/<[^>]+>/g, '').trim() || /<img/.test(h)) return `<div class="${cls}">${h}</div>`;
    }
    return '';
  }
  const headerHtml = pickPart('w:headerReference', 'w:hdr', 'hdr');
  const footerHtml = pickPart('w:footerReference', 'w:ftr', 'ftr');

  function runsHtml(container) {
    let html = '';
    let fld = null; // suivi de champ (fldChar/instrText) pour PAGE/NUMPAGES -> placeholders
    for (const c of kids(container)) {
      const t = tagOf(c);
      if (t === 'w:r') {
        // détecte le début/code/valeur/fin de champ dans ce run
        for (const rc of kids(c)) {
          const rt = tagOf(rc);
          if (rt === 'w:fldChar') { const ft = at(rc)['@w:fldCharType']; if (ft === 'begin') fld = { code: '', val: false, emitted: false }; else if (ft === 'separate' && fld) fld.val = true; else if (ft === 'end') fld = null; }
          else if (rt === 'w:instrText' && fld && !fld.val) fld.code += kids(rc).map((k) => k['#text'] ?? '').join('');
        }
        // champ PAGE / NUMPAGES : émettre UN placeholder (rempli par le paginateur) et SAUTER tous ses
        // runs de valeur (le run 'separate' ET le run cache -> sinon doublé « 1616 »).
        if (fld && fld.val) {
          const isNum = /\bNUMPAGES\b/.test(fld.code), isPage = !isNum && /\bPAGE\b/.test(fld.code);
          if (isNum || isPage) { if (!fld.emitted) { html += isNum ? '<span class="pagecount">1</span>' : '<span class="pageno">1</span>'; fld.emitted = true; } continue; }
        }
        html += renderRun(c);
      }
      else if (t === 'w:hyperlink') html += `<a class="lnk">${runsHtml(c)}</a>`;
      // w:del/w:ins = suivi de modifications : suppressions (texte dans w:delText) et insertions.
      // On DESCEND dedans (sinon tout le contenu « suivi » est perdu, ex. pavé rose fiche SPC).
      else if (t === 'w:ins' || t === 'w:del' || t === 'w:smartTag' || t === 'w:sdt') html += runsHtml(t === 'w:sdt' ? (child(c, 'w:sdtContent') || c) : c);
      // mc:AlternateContent (Choice moderne / Fallback ancien) : formes, zones de texte, dessins.
      // On prend la Choice (contenu moderne) et on la traite (sinon perdu, ex. logos d'en-tête).
      else if (t === 'mc:AlternateContent') { const ch = child(c, 'mc:Choice') || child(c, 'mc:Fallback'); if (ch) html += runsHtml(ch); }
    }
    return html;
  }

  function renderRun(r) {
    const rPr = child(r, 'w:rPr');
    // héritage : docDefaults -> STYLE DU PARAGRAPHE (curParaRPr, ex. taille/gras d'un titre) -> style de
    // run -> rPr inline. Sans le style du paragraphe, les runs d'un titre restaient en 11pt.
    let props = { ...curParaRPr };
    const rStyle = childAttrs(rPr || {}, 'w:rStyle')?.['@w:val'];
    if (rStyle) props = { ...props, ...resolveStyle(rStyle, styles, theme).rPr };
    props = { ...props, ...parseRPr(rPr, theme) };
    let inner = '', deleted = false;
    for (const c of kids(r)) {
      const t = tagOf(c);
      if (t === 'w:t') { const txt = kids(c).map((k) => k['#text'] ?? '').join(''); inner += esc(txt); }
      // w:delText = texte SUPPRIMÉ (suivi de modifications) : on l'affiche (comme Word en mode markup)
      // et on le marque barré. Sans ça, tout le contenu « suivi » disparaît (pavé rose fiche SPC).
      else if (t === 'w:delText') { const txt = kids(c).map((k) => k['#text'] ?? '').join(''); inner += esc(txt); deleted = true; }
      else if (t === 'w:tab') inner += '<span style="display:inline-block;width:36px"></span>';
      else if (t === 'w:br') { if (at(c)['@w:type'] === 'page') pendingBreak = true; else inner += '<br>'; }
      else if (t === 'w:lastRenderedPageBreak') pendingBreak = true; // frontière de page cachée par Word
      else if (t === 'w:cr') inner += '<br>';
      else if (t === 'w:noBreakHyphen') inner += '-';
      else if (t === 'w:sym') { const a = at(c); const code = parseInt(a['@w:char'], 16); if (!isNaN(code)) inner += '&#' + ((code >= 0xF000 && code <= 0xF0FF) ? code - 0xF000 + 0x2022 : code) + ';'; }
      else if (t === 'w:drawing' || t === 'w:pict') inner += renderDrawing(c);
      // dessins/zones de texte enveloppés dans mc:AlternateContent (ex. logos d'en-tête)
      else if (t === 'mc:AlternateContent') { const ch = child(c, 'mc:Choice') || child(c, 'mc:Fallback'); if (ch) for (const cc of kids(ch)) { const ct = tagOf(cc); if (ct === 'w:drawing' || ct === 'w:pict') inner += renderDrawing(cc); } }
    }
    if (!inner) return '';
    const st = [];
    if (props.bold) st.push('font-weight:700');
    if (props.italic) st.push('font-style:italic');
    const deco = []; if (props.underline) deco.push('underline'); if (props.strike || deleted) deco.push('line-through');
    if (deco.length) st.push('text-decoration:' + deco.join(' '));
    if (props.sizePt) st.push('font-size:' + props.sizePt + 'pt');
    if (props.color) st.push('color:' + props.color);
    if (props.font) st.push('font-family:' + fontFamily(props.font));
    if (props.highlight && props.highlight !== 'none') st.push('background:' + (props.highlight === 'yellow' ? '#ff0' : props.highlight));
    if (props.caps) st.push('text-transform:uppercase');
    if (props.vertAlign === 'superscript') st.push('vertical-align:super;font-size:.7em');
    if (props.vertAlign === 'subscript') st.push('vertical-align:sub;font-size:.7em');
    return `<span style="${st.join(';')}">${inner}</span>`;
  }

  // positionnement d'un dessin ANCRÉ (wp:anchor flottant). On positionne en ABSOLU seulement si la
  // référence est la PAGE ou la MARGE (typique des pages de garde/fonds) -> reconstruit le collage.
  // Les flottants relatifs au texte (column/paragraph/line/character) restent EN FLUX (moins risqué).
  function anchorPos(node) {
    const anc = child(node, 'wp:anchor'); if (!anc) return null;
    const posOff = (P) => { const o = P && child(P, 'wp:posOffset'); return o ? +kids(o).map((k) => k['#text'] ?? '').join('') : null; };
    const H = child(anc, 'wp:positionH'), V = child(anc, 'wp:positionV');
    const relH = (H && at(H)['@relativeFrom']) || '', relV = (V && at(V)['@relativeFrom']) || '';
    const x = posOff(H), y = posOff(V); if (x == null || y == null) return null;
    const ext = child(anc, 'wp:extent'); const a = ext ? at(ext) : {};
    const pageH = relH === 'page', marginH = /margin/i.test(relH);
    const pageV = relV === 'page', marginV = /margin/i.test(relV);
    // page/margin sur les DEUX axes -> absolu dans la PAGE. Sinon (column/paragraph/line/character) ->
    // absolu dans le PARAGRAPHE conteneur (offset depuis le coin haut-gauche du paragraphe) = modèle Word.
    const toPage = (pageH || marginH) && (pageV || marginV);
    return {
      toPage,
      left: toPage ? emuPx(x) + (marginH ? padL : 0) : emuPx(x) + (marginH ? padL : 0),
      top: toPage ? emuPx(y) + (marginV ? padT : 0) : emuPx(y),
      cx: +a['@cx'] || 0, cy: +a['@cy'] || 0,
      behind: at(anc)['@behindDoc'] === '1',
    };
  }

  function renderDrawing(node) {
    // ZONE DE TEXTE (wps:txbx / v:textbox -> w:txbxContent) : contenu (texte ET images internes).
    let txbx = null; const findTx = (n) => { if (tagOf(n) === 'w:txbxContent') { txbx = n; return; } for (const c of kids(n)) if (!txbx) findTx(c); };
    findTx(node);
    let content = '', cx = 0;
    if (txbx) {
      let inner = '';
      for (const c of kids(txbx)) inner += renderBlock(c);
      content = inner ? `<span class="ibx">${inner}</span>` : '';
    } else {
      // IMAGE. premier embed = le VRAI blip ; suivants = a14:imgLayer .wdp (JPEG-XR) IGNORÉS.
      let embed = null;
      const walk = (n) => { const a = at(n); if (a['@r:embed'] && !embed) embed = a['@r:embed']; if (a['@cx'] && !cx) cx = +a['@cx']; for (const c of kids(n)) walk(c); };
      walk(node);
      const src = embed && imageData(embed);
      if (src) { const w = cx ? ` width="${Math.round(emuPx(cx))}"` : ''; content = `<img src="${src}"${w} style="max-width:100%">`; }
    }
    if (!content) return '';
    // dessin ANCRÉ (flottant) -> positionnement ABSOLU. toPage : relatif à la page (.page). Sinon :
    // relatif au PARAGRAPHE conteneur (renderP le passe en position:relative via le flag paraHasAbs).
    // NB : les dessins ANCRÉS (flottants, wp:anchor) sont rendus EN FLUX (inline), PAS en absolu.
    // Le positionnement absolu (collage de page de garde) est incompatible avec la pagination en flux
    // (les éléments absolus ont 0 hauteur -> le paginateur les disperse). Un collage flottant fidèle
    // exige un vrai moteur de layout (LibreOffice). Ici : images visibles, en flux, paginées proprement.
    return content;
  }

  function renderP(p) {
    const pPr = child(p, 'w:pPr');
    const pp = parsePPr(pPr);
    const styleId = pp.style;
    const st = styleId ? resolveStyle(styleId, styles, theme) : { rPr: {}, pPr: {}, name: '' };
    const align = pp.align || st.pPr.align;
    const cssAlign = { center: 'center', right: 'right', both: 'justify', justify: 'justify' }[align] || 'left';
    // titre ?
    const outline = pp.outline ?? st.pPr.outline;
    const isHeading = /heading|titre/i.test(st.name || '') || outline !== undefined;
    const hLvl = outline !== undefined ? Math.min(6, +outline + 1) : (/(\d)/.exec(st.name || '')?.[1] || 2);
    // TABLE DES MATIÈRES (styles TM1/TM2/TM3 FR ou TOC1.. EN) : texte + points de suite + n° à droite.
    // Le <w:tab leader=dot right> sépare le libellé du numéro (champ PAGEREF). On rend en flex.
    const tocM = /^(?:TM|TOC)(\d)/i.exec(styleId || '');
    if (tocM) {
      let before = '', after = '', seenTab = false;
      (function walk(n) { for (const c of kids(n)) { const t = tagOf(c);
        if (t === 'w:pPr') continue; // NE PAS descendre : <w:tabs><w:tab/> y définit le taquet (pas un tab)
        if (t === 'w:tab') seenTab = true;
        else if (t === 'w:t') { const s = kids(c).map((k) => k['#text'] ?? '').join(''); if (seenTab) after += s; else before += s; }
        else if (t === 'w:instrText' || t === 'w:fldChar' || t === 'w:delText') { /* code de champ : ignoré */ }
        else if (kids(c).length) walk(c);
      } })(p);
      before = before.trim(); after = after.trim();
      if (!before && !after) return '';
      const lvl = +tocM[1] || 1;
      if (!after) return `<p class="toc tm${lvl}"><span class="tx">${esc(before)}</span></p>`;
      return `<p class="toc tm${lvl}"><span class="tx">${esc(before)}</span><span class="dots"></span><span class="pg">${esc(after)}</span></p>`;
    }
    // props de run héritées du STYLE du paragraphe (taille/gras/couleur d'un titre, etc.)
    const savedRPr = curParaRPr; curParaRPr = mrg(docDefaults.rPr, st.rPr);
    const savedAbs = paraHasAbs; paraHasAbs = false;
    const inner = runsHtml(p);
    const hasAbs = paraHasAbs; paraHasAbs = savedAbs; // restaure (gère l'imbrication cellules/txbx)
    curParaRPr = savedRPr;
    const styleAttr = [`text-align:${cssAlign}`];
    if (hasAbs) styleAttr.push('position:relative'); // ancre les flottants relatifs au paragraphe
    // INTERLIGNE + ESPACEMENT + INDENTATION : docDefaults -> style -> inline (le + spécifique gagne).
    const sp = mrg(docDefaults.pPr.spacing, st.pPr.spacing, pp.spacing);
    const ind = mrg(docDefaults.pPr.ind, st.pPr.ind, pp.ind);
    // interligne : lineRule auto -> line/240 (240=simple) ; exact/atLeast -> line/20 en pt
    if (sp.line) {
      if (sp.lineRule === 'exact' || sp.lineRule === 'atLeast') styleAttr.push(`line-height:${(+sp.line / 20).toFixed(1)}pt`);
      else styleAttr.push(`line-height:${(+sp.line / 240).toFixed(3)}`);
    }
    // espace avant/après (twips->px) ; autospacing -> ~14px (comme Word pour le texte « HTML »)
    const before = sp.beforeAuto === '1' || sp.beforeAuto === 'true' ? 14 : (sp.before != null ? twip(sp.before) : null);
    const after = sp.afterAuto === '1' || sp.afterAuto === 'true' ? 14 : (sp.after != null ? twip(sp.after) : null);
    if (before != null) styleAttr.push(`margin-top:${before.toFixed(1)}px`);
    if (after != null) styleAttr.push(`margin-bottom:${after.toFixed(1)}px`);
    // indentation : gauche/droite + 1re ligne (firstLine positif / hanging négatif)
    if (ind.left != null) styleAttr.push(`padding-left:${twip(ind.left).toFixed(1)}px`);
    if (ind.right != null) styleAttr.push(`padding-right:${twip(ind.right).toFixed(1)}px`);
    if (ind.firstLine != null) styleAttr.push(`text-indent:${twip(ind.firstLine).toFixed(1)}px`);
    else if (ind.hanging != null) styleAttr.push(`text-indent:${(-twip(ind.hanging)).toFixed(1)}px`);
    // liste ?
    const num = pp.num || st.pPr.num;
    if (num && num.numId) {
      const L = numbering.lvl(num.numId, num.ilvl);
      const key = num.numId + ':' + num.ilvl;
      let marker;
      if (L.fmt === 'bullet') {
        // vraie PUCE : caractère du lvlText, mappé si police symbole/PUA (Wingdings ➢▪, Symbol...)
        marker = bulletChar(L.ch || '•', L.buFont);
      } else {
        counters[key] = (counters[key] || 0) + 1;
        marker = counters[key] + (/\)/.test(L.ch || '') ? ')' : '.'); // numéroté (décimal/romain approx)
      }
      if (ind.left == null) styleAttr.push(`padding-left:${24 + num.ilvl * 24}px`); // indent si non déjà défini
      return `<p class="li" style="${styleAttr.join(';')}"><span class="mk">${marker}</span>${inner || '&#8203;'}</p>`;
    }
    if (isHeading && inner) return `<h${hLvl} style="${styleAttr.join(';')}">${inner}</h${hLvl}>`;
    return `<p style="${styleAttr.join(';')}">${inner || '&#8203;'}</p>`;
  }

  function renderTbl(tbl) {
    // largeurs de colonnes (tblGrid > gridCol, en twips) -> <colgroup> pour respecter les proportions Word.
    const tblPr = child(tbl, 'w:tblPr');
    const grid = child(tbl, 'w:tblGrid');
    let colgroup = '';
    if (grid) {
      const cols = kids(grid).filter((c) => tagOf(c) === 'w:gridCol').map((c) => +at(c)['@w:w'] || 0);
      const totalW = cols.reduce((a, b) => a + b, 0);
      if (totalW > 0) colgroup = '<colgroup>' + cols.map((w) => `<col style="width:${(100 * w / totalW).toFixed(2)}%">`).join('') + '</colgroup>';
    }
    // largeur de table : tblW type dxa -> px ; pct -> %. Sinon auto (contenu).
    const tw = childAttrs(tblPr || {}, 'w:tblW');
    let twStyle = '';
    if (tw && tw['@w:type'] === 'dxa' && +tw['@w:w'] > 0) twStyle = `width:${twip(tw['@w:w']).toFixed(0)}px`;
    else if (tw && tw['@w:type'] === 'pct' && +tw['@w:w'] > 0) twStyle = `width:${(+tw['@w:w'] / 50).toFixed(0)}%`;
    // bordures : si tblBorders top=none/nil -> table SANS bordure (tableau de mise en page). Sinon 1px.
    const tb = child(tblPr || {}, 'w:tblBorders');
    const topB = tb && childAttrs(tb, 'w:top')?.['@w:val'];
    const noBorder = topB === 'none' || topB === 'nil';
    const cls = 'doctbl' + (noBorder ? ' nb' : '');
    let rows = '';
    for (const tr of kids(tbl)) {
      if (tagOf(tr) !== 'w:tr') continue;
      let cells = '';
      for (const tc of kids(tr)) {
        if (tagOf(tc) !== 'w:tc') continue;
        const tcPr = child(tc, 'w:tcPr');
        const gs = childAttrs(tcPr || {}, 'w:gridSpan')?.['@w:val'];
        // remplissage de cellule (w:shd @w:fill) -> fond coloré (ex. pavé rose des fiches SPC)
        const shd = childAttrs(tcPr || {}, 'w:shd');
        const fill = shd && shd['@w:fill'] && shd['@w:fill'] !== 'auto' ? shd['@w:fill'] : null;
        const bg = fill ? ` style="background:#${fill}"` : '';
        let cellHtml = '';
        for (const c of kids(tc)) cellHtml += renderBlock(c);
        cells += `<td${gs ? ` colspan="${gs}"` : ''}${bg}>${cellHtml}</td>`;
      }
      rows += `<tr>${cells}</tr>`;
    }
    return `<table class="${cls}"${twStyle ? ` style="${twStyle}"` : ''}>${colgroup}${rows}</table>`;
  }

  const css = `
  :root{color-scheme:light}
  body{margin:0;background:#525659;font-family:'Calibri','Carlito',sans-serif;color:#111}
  .page{position:relative;display:flex;flex-direction:column;width:${pageW}px;min-height:${pageH}px;margin:24px auto;background:#fff;box-shadow:0 3px 12px rgba(0,0,0,.5);padding:${padT}px ${padR}px ${padB}px ${padL}px;box-sizing:border-box;overflow:hidden}
  .pgbody{flex:0 1 auto}
  .ftr{margin-top:auto;padding-top:6px;font-size:9pt;color:#555}
  .floatabs,.floatpara{position:absolute}
  .floatabs img,.floatpara img{width:100%;height:100%;display:block;max-width:none}
  .page p{margin:0;line-height:1.15}
  .page h1,.page h2,.page h3,.page h4,.page h5,.page h6{margin:0;font-weight:700}
  .page .li{position:relative}
  .page .li .mk{display:inline-block;min-width:20px;margin-left:-20px}
  a.lnk{color:#0563c1;text-decoration:underline}
  .toc{display:flex;align-items:baseline;margin:2px 0;color:#000}
  .toc .tx{white-space:nowrap}
  .toc .dots{flex:1;border-bottom:1.5px dotted #444;margin:0 4px;position:relative;top:-3px;min-width:12px}
  .toc .pg{white-space:nowrap;color:#000}
  .toc.tm2{margin-left:22px}
  .toc.tm3{margin-left:44px}
  .toc.tm4{margin-left:66px}
  table.doctbl{border-collapse:collapse;margin:8px 0;width:100%;table-layout:fixed}
  table.doctbl td{border:1px solid #999;padding:4px 6px;vertical-align:top;overflow-wrap:break-word}
  table.doctbl.nb td{border:none}
  img{height:auto}
  .hdr{opacity:.85;border-bottom:1px solid #ddd;margin-bottom:16px;padding-bottom:8px}
  .ibx{display:inline-block;vertical-align:top}
  .ibx p{margin:0}
  `;
  // PAGINATION CLIENT : on émet le flux caché + un gabarit d'en-tête, et un script découpe en pages A4
  // uniformes en MESURANT les hauteurs réelles (ce qu'un rendu serveur ne peut pas faire), tout en
  // respectant les sauts de page Word (.pgbreak). -> plus d'« énormes pages A4 » de hauteur variable.
  const paginator = `
  (function(){
    var CH=${Math.round(pageH)}; // hauteur A4 cible (px)
    var src=document.getElementById('src'), pagesEl=document.getElementById('pages');
    var hdr=document.getElementById('hdrtpl'); var hdrHtml=hdr?hdr.innerHTML:'';
    var ftr=document.getElementById('ftrtpl'); var ftrHtml=ftr?ftr.innerHTML:'';
    var pageList=[];
    function newPage(){var pg=document.createElement('div');pg.className='page';if(hdrHtml)pg.insertAdjacentHTML('beforeend',hdrHtml);var body=document.createElement('div');body.className='pgbody';pg.appendChild(body);if(ftrHtml)pg.insertAdjacentHTML('beforeend',ftrHtml);pagesEl.appendChild(pg);pageList.push(pg);return body;}
    function run(){
      var blocks=[].slice.call(src.children); var body=newPage();
      for(var i=0;i<blocks.length;i++){
        var bl=blocks[i];
        if(bl.classList&&bl.classList.contains('pgbreak')){ body=newPage(); continue; }
        body.appendChild(bl);
        // déborde l'A4 -> nouvelle page (sauf si le bloc est seul et trop grand : on le laisse déborder)
        if(body.parentNode.scrollHeight>CH+2 && body.children.length>1){ body=newPage(); body.appendChild(bl); }
      }
      src.parentNode.removeChild(src);
      // retire une éventuelle dernière page vide
      var last=pageList[pageList.length-1]; if(last){var lb=last.querySelector('.pgbody'); if(lb&&!lb.children.length){last.parentNode.removeChild(last);pageList.pop();}}
      // numéros de page : PAGE -> index, NUMPAGES -> total
      var total=pageList.length;
      for(var p=0;p<pageList.length;p++){
        var no=p+1;
        var pn=pageList[p].querySelectorAll('.pageno'); for(var k=0;k<pn.length;k++)pn[k].textContent=no;
        var pc=pageList[p].querySelectorAll('.pagecount'); for(var k=0;k<pc.length;k++)pc[k].textContent=total;
      }
    }
    if(document.fonts&&document.fonts.ready){document.fonts.ready.then(run);}else{run();}
  })();`;
  const bodyHtml = `<div id="hdrtpl" style="display:none">${headerHtml}</div><div id="ftrtpl" style="display:none">${footerHtml}</div><div id="src" style="display:none">${flow.join('') || ''}</div><div id="pages"></div><script>${paginator}</script>`;
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${esc(basename(docxPath))}</title><style>${css}</style></head><body>${bodyHtml}</body></html>`;
  return { html, pageW };
}

// --- CLI --- (uniquement si appelé avec un .docx en argument ; sinon importé comme module)
const docxArg = process.argv[2];
if (docxArg && docxArg.toLowerCase().endsWith('.docx')) {
  const { html } = renderDocx(docxArg);
  const outPath = join(ROOT, 'renderer', 'docx-out.html');
  writeFileSync(outPath, html);
  console.log(`OK -> renderer/docx-out.html  (${basename(docxArg)})`);
}

// --- exports nommés pour les TESTS (fonctions PURES uniquement) ---
// N'expose que des helpers sans effet de bord (sortie = fonction des entrées).
// AUCUNE ligne de logique ci-dessus n'est modifiée par ce bloc.
export { tagOf, kids, at, mrg, esc, child, childAttrs, onOff, twip, emuPx, bulletChar, fontFamily, resolveColor };
