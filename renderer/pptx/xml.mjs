// ═══ [PPTX · XML] ═══ parseurs fast-xml-parser, cache de parse, helpers d'arbre, esc(), serialize().
// Module FEUILLE : n'importe rien du moteur. Voir pptx-render.mjs (guide §5) pour les deux formes d'arbre.
import { readFileSync, statSync } from 'node:fs';
import { XMLParser } from 'fast-xml-parser';
// Options communes aux deux parseurs :
//   ignoreAttributes:false + attributeNamePrefix '@' -> les attributs sont lus comme n['@val'] ;
//   parseTagValue/parseAttributeValue:false -> TOUT reste chaîne (« 0100 » ne devient pas 100,
//   un hex « 1E10 » ne devient pas un nombre) : les conversions sont faites explicitement (+x).
// P  : preserveOrder -> tableau ordonné (structure, z-order, runs). Attributs sous la clé ':@'.
// PF : objet plat -> lecture directe des propriétés (thème, rels, spPr, txBody re-parsés).
// trimValues:false -> conserve les espaces en bord de <a:t> (sinon les mots se collent
// entre runs successifs, ex. "Marge " + "qu'on" -> "Margequ'on").
const P = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', preserveOrder: true, parseTagValue: false, parseAttributeValue: false, trimValues: false });
const PF = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@', parseTagValue: false, parseAttributeValue: false, trimValues: false }); // flat, pour thème/rels

// ████ §B ─ CACHE DE PARSE ────────────────────────────────────────────────────────────────────────
// Rôle : ne parser qu'une fois les fichiers partagés par toutes les diapos d'un deck. Chaque entrée
// garde jusqu'à trois représentations paresseuses du même fichier : texte brut ('txt'), arbre plat
// ('flat', parseur PF) et arbre ordonné ('ord', parseur P).
// ---- cache de parse pour fichiers STABLES d'un deck (thème/master/layout/rels/presentation/tableStyles) ----
// renderSlide est appelé par diapo, mais N diapos partagent le même master/layout/thème : sans cache on
// reparse ~64 Ko de XML par diapo (master+layout parsés 2× chacun). Clé = chemin + mtime + taille : si un
// dossier temporaire est réutilisé avec un contenu DIFFÉRENT (ex. batch check.mjs), la signature change et
// le cache est invalidé -> jamais de rendu périmé. Les fichiers PROPRES à la diapo (slideN.xml, ses rels)
// ne sont PAS mis en cache. Les arbres mis en cache sont traités en LECTURE SEULE (le code re-sérialise +
// reparse tout sous-arbre qu'il doit muter) -> partage sûr entre diapos (prouvé par hash SVG identique).
const _pc = new Map(); // chemin -> { sig, txt?, flat?, ord? } ; jamais purgé (taille = nb de fichiers stables vus)
// Signature « mtime:taille » d'un fichier. statSync LÈVE si le fichier n'existe pas : les appelants
// vérifient existsSync avant, ou entourent d'un try (presentation.xml, tableStyles.xml).
function _stamp(path) { const s = statSync(path); return s.mtimeMs + ':' + s.size; }
// cachedParse(path, kind) -> texte ('txt'), arbre plat ('flat') ou arbre ordonné ('ord', défaut).
// Un stat par appel (peu coûteux) garantit l'invalidation si le fichier a changé sur disque.
// ??= : chaque représentation n'est calculée qu'à la première demande.
function cachedParse(path, kind /* 'txt' | 'flat' | 'ord' */) {
  const sig = _stamp(path);
  let e = _pc.get(path);
  // signature différente -> entrée neuve : on jette txt/flat/ord périmés d'un coup
  if (!e || e.sig !== sig) { e = { sig }; _pc.set(path, e); }
  if (e.txt === undefined) e.txt = readFileSync(path, 'utf8');
  if (kind === 'txt') return e.txt;
  if (kind === 'flat') return (e.flat ??= PF.parse(e.txt));
  return (e.ord ??= P.parse(e.txt));
}

// ████ §C ─ HELPERS D'ARBRE ET ÉCHAPPEMENT ──────────────────────────────────────────────────────
// Un nœud preserveOrder ressemble à : { 'a:off': [], ':@': { '@x': '0', '@y': '12700' } }
// ou, pour du texte : { '#text': 'Bonjour' }. La seule clé autre que ':@' est le nom de balise.
// Ces helpers sont tolérants : un nœud absent (undefined) donne {} / [] au lieu de lever, ce qui
// permet d'enchaîner first(first(x,'a'),'b') sans garde à chaque niveau.
// ---- helpers arbre preserveOrder ----
// tagOf(n) : nom de la balise ('a:sp', 'p:txBody'… ou '#text' pour un nœud texte)
const tagOf = (n) => Object.keys(n).find((k) => k !== ':@');
// at(n) : attributs du nœud ({ '@id': '3', … }) ; {} si aucun ou si n est absent
const at = (n) => n?.[':@'] ?? {};
// kids(n) : enfants ordonnés (tableau) ; [] pour un nœud texte ou absent
const kids = (n) => { const v = n?.[tagOf(n)]; return Array.isArray(v) ? v : []; };
// byTag(n, t) : enfants DIRECTS de balise t, dans l'ordre du document
const byTag = (n, t) => kids(n).filter((c) => tagOf(c) === t);
// first(n, t) : 1er enfant direct de balise t (undefined si aucun)
const first = (n, t) => byTag(n, t)[0];
// deepFind(n, t) : 1er nœud de balise t en profondeur d'abord (n lui-même compris). Attention :
// « le 1er » au sens du document — ex. deepFind(sp, 'p:cNvPr') trouve bien celui de la forme, mais
// deepFind(grp, 'p:cNvPr') sur un groupe trouve celui du GROUPE (nvGrpSpPr précède les membres).
function deepFind(n, t) {
  if (!n || typeof n !== 'object') return null;
  if (tagOf(n) === t) return n;
  for (const c of kids(n)) { const r = deepFind(c, t); if (r) return r; }
  return null;
}
// recherche profonde d'une clé dans un objet PLAT (parse non-ordonné) -> 1re valeur trouvée
// Note : `if (r)` ignore une valeur falsy trouvée en profondeur ('' ou 0) ; seul le niveau courant
// accepte toute valeur !== undefined. Sans conséquence pour les usages (nœuds objets a14:*).
function deepFindFlat(obj, key) {
  if (!obj || typeof obj !== 'object') return null;
  if (obj[key] !== undefined) return obj[key];
  for (const k of Object.keys(obj)) { const r = deepFindFlat(obj[k], key); if (r) return r; }
  return null;
}
// trouve le 1er noeud dont le tag se termine par `suffix` (ignore le préfixe de namespace)
// Utile quand le préfixe varie d'un fichier à l'autre (ex. asvg:svgBlip, préfixe choisi par l'auteur).
function deepFindSuffix(n, suffix) {
  if (!n || typeof n !== 'object') return null;
  const t = tagOf(n);
  if (t && t.endsWith(suffix)) return n;
  for (const c of kids(n)) { const r = deepFindSuffix(c, suffix); if (r) return r; }
  return null;
}
// Échappe aussi ' (XSS-01) : un typeface comme `x';color:red` sortait de la chaîne CSS 'font-family'
// et injectait des propriétés. On utilise l'entité NOMMÉE &apos; (et non la référence numérique &#39;) :
// fast-xml-parser NE décode PAS &#39; lors d'un re-parsing (voir PF.parse sur les fills), si bien qu'un
// texte re-sérialisé puis re-esc donnait &amp;#39; (double-échappement visible sur chaque apostrophe
// française — « l&#39;unité »). &apos; est décodée au round-trip et protège autant (apostrophe échappée).
// esc(s) : échappement XML/HTML des 5 caractères spéciaux. Utilisé PARTOUT où une valeur venant du
// fichier (texte, typeface, id, attribut re-sérialisé) entre dans le SVG/HTML produit : c'est la
// barrière anti-injection (un .pptx est une entrée non fiable). null/undefined -> ''.
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
// ████ §Q ─ SÉRIALISATION ET EXPORTS ─────────────────────────────────────────────────────────────
// sérialise un noeud preserveOrder -> XML (pour re-parser en flat)
// Pièce centrale de la stratégie « parcours ordonné, lecture à plat » (guide §5) : utilisée avant
// chaque PF.parse d'un sous-arbre. Attributs et texte sont ré-échappés par esc() (fast-xml-parser
// les avait décodés) ; les balises gardent leur préfixe (a:, p:…) tel quel. Pas de balises
// auto-fermantes : <a:off x="0"/> devient <a:off x="0"></a:off> (équivalent au re-parse).
function serialize(node) {
  const t = tagOf(node);
  const a = at(node);
  const attrs = Object.entries(a).map(([k,v])=>` ${k.slice(1)}="${esc(v)}"`).join(''); // k.slice(1) retire le '@'
  const children = kids(node);
  // kids() renvoie toujours un tableau : cette branche « chaîne » ne se déclenche jamais (le texte
  // arrive comme enfant { '#text' } traité juste en dessous) ; conservée par prudence.
  if (typeof children === 'string') return `<${t}${attrs}>${esc(children)}</${t}>`;
  const inner = (children||[]).map(c => {
    if (c['#text'] !== undefined) return esc(c['#text']);
    return serialize(c);
  }).join('');
  return `<${t}${attrs}>${inner}</${t}>`;
}

// parseFlat(nœud ordonné) : re-sérialise un sous-arbre ordonné puis le re-parse en arbre PLAT.
// Le wrapper <root xmlns…> ne sert qu'à fournir une racine (fast-xml-parser ne résout pas les namespaces).
export const parseFlat = (node) => PF.parse(`<root xmlns:a="a" xmlns:p="p" xmlns:r="r">${serialize(node)}</root>`).root;

export { P, PF, cachedParse, _stamp as fileStamp, tagOf, at, kids, byTag, first, deepFind, deepFindFlat, deepFindSuffix, esc, serialize };
