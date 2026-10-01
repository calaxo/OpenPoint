// Extraction des NOTES DU PRÉSENTATEUR d'un .pptx (le volet « Notes » de PowerPoint).
//
// But (LMS) : à l'import d'un cours PPTX, récupérer le texte des notes de chaque
// diapo pour le proposer aux instructeurs comme référence (voir models/slideNotes.js).
// On ne touche PAS au fichier source : lecture seule du ZIP, en mémoire.
//
// Mapping diapo -> note : on suit les relations OOXML dans l'ORDRE DE PRÉSENTATION
// (ppt/presentation.xml > sldIdLst), puis pour chaque diapo sa relation `notesSlide`
// (ppt/slides/_rels/slideN.xml.rels). C'est le même ordre que le PDF produit par
// LibreOffice (pdftoppm) -> la page N du lecteur correspond bien à la note N.
//
// Robuste par conception : toute diapo sans note, tout .pptx mal formé -> tableau
// vide / entrées manquantes ignorées (best-effort, jamais d'exception propagée
// au flux d'import). Réutilise fflate + fast-xml-parser (déjà dépendances du back).
//
// ─── COMMENT ÇA MARCHE (vue d'ensemble) ─────────────────────────────────────────────────────────
// RAPPEL : un .pptx est une archive ZIP de parties XML reliées par des fichiers de relations (.rels).
//   Chemin suivi par ce module, pour une présentation de N diapos :
//
//   ppt/presentation.xml                   <p:sldIdLst><p:sldId id="256" r:id="rId2"/>…  (ORDRE)
//     └─ ppt/_rels/presentation.xml.rels   rId2 -> "slides/slide3.xml"   (le n° de fichier ≠ rang !)
//          └─ ppt/slides/slide3.xml
//               └─ ppt/slides/_rels/slide3.xml.rels
//                    Type=".../relationships/notesSlide" Target="../notesSlides/notesSlide7.xml"
//                      └─ ppt/notesSlides/notesSlide7.xml   <p:notes><p:cSld><p:spTree><p:sp>…
//
//   Le numéro dans le nom de fichier (slide3, notesSlide7) n'a AUCUN sens d'ordre : seul sldIdLst
//   donne le rang de la diapo. C'est pour ça qu'on ne se contente pas de lister ppt/notesSlides/.
//
// CONTENU D'UNE DIAPO DE NOTES (p:notes)
//   Plusieurs p:sp « placeholders » : la vignette de la diapo (sldImg), le corps des notes (body),
//   et éventuellement numéro, date, en-tête, pied de page. On garde tout SAUF les types listés dans
//   SKIP_PH ; en pratique il reste le placeholder body (type="body"), ou des zones de texte libres.
//
// ENTRÉE / SORTIE
//   extractPptxNotes("/chemin/cours.pptx") -> [{ page: 1, body: "texte…" }, { page: 3, body: … }]
//   page = rang 1-based dans l'ordre de présentation ; les diapos sans note sont ABSENTES du tableau
//   (pas d'entrée vide) ; body est tronqué à 8000 caractères.
//
// LIMITES CONNUES (comportement du code tel quel)
//   - Texte brut uniquement : gras, puces, niveaux de liste et liens sont perdus ; les sauts de ligne
//     internes à un paragraphe (<a:br/>) sont ignorés (les deux morceaux sont collés).
//   - Les champs (a:fld) d'un paragraphe sont ajoutés APRÈS ses runs, pas à leur position réelle.
//   - Le parseur est construit avec les options par défaut de fast-xml-parser pour le reste
//     (parseTagValue et trimValues actifs) : un run dont le a:t est purement numérique (« 42 »)
//     devient un nombre JS et n'est PAS repris par shapeText (qui n'accepte que chaîne / #text) ;
//     les espaces en bord de run sont rognés (« Bonjour » + « le monde » -> « Bonjourle monde »
//     si l'espace était porté par le 2e run). Vérifié avec fast-xml-parser du dépôt.
//   - Une diapo masquée (p:sld @show="0") est comptée comme les autres dans la numérotation.
//   - Lecture SYNCHRONE du fichier entier en mémoire (readFileSync + unzipSync) : bloque l'event loop
//     le temps de la décompression — acceptable pour un import ponctuel.
// ────────────────────────────────────────────────────────────────────────────────────────────────
import { readFileSync } from "node:fs";
import { unzipSync } from "fflate";
import { XMLParser } from "fast-xml-parser";

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 1 — Parseur XML et petits utilitaires
// ════════════════════════════════════════════════════════════════════════════════════════════════
// Un seul parseur réutilisé pour toutes les parties (presentation.xml, .rels, notesSlideN.xml).
//   ignoreAttributes:false + prefix "@_" : les attributs deviennent des clés "@_type", "@_r:id"…
//   removeNSPrefix:false : les préfixes d'espace de noms restent dans les clés ("p:sp", "a:t") —
//     on suppose donc les préfixes standard p: / a: / r: écrits par PowerPoint.
//   isArray: () => false : aucun élément n'est forcé en tableau. Un élément unique donne un objet,
//     un élément répété donne un tableau -> asArray() uniformise à la lecture.
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: false,
  // On garde les tableaux dès qu'il y a répétition (spTree, a:p, a:r…).
  isArray: () => false,
});

// asArray(v) : undefined/null -> [], tableau -> tel quel, objet unique -> [objet].
const asArray = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]);

// Placeholders à IGNORER dans une diapo de notes (ne font pas partie du texte
// rédigé par l'auteur) : numéro de diapo, date, en-tête, pied de page, et la
// vignette de la diapo elle-même.
// (Valeurs de ST_PlaceholderType, attribut p:ph/@type. Un shape sans p:ph — zone de texte ajoutée
// à la main — a un type null et est donc conservé.)
const SKIP_PH = new Set(["sldNum", "dt", "hdr", "ftr", "sldImg"]);

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 2 — Extraction du texte d'une diapo de notes
// ════════════════════════════════════════════════════════════════════════════════════════════════

// Texte d'un shape de notes : concatène les runs (a:r > a:t) et champs (a:fld > a:t)
// paragraphe par paragraphe. Un saut de ligne sépare les paragraphes.
// shapeText(sp) -> chaîne (éventuellement vide).
//   Structure lue : p:sp > p:txBody > a:p* > (a:r > a:t | a:fld > a:t).
//   Un paragraphe vide (<a:p/>, ligne blanche dans PowerPoint) donne une ligne vide : les lignes
//   blanches volontaires sont conservées, mais 3 sauts ou plus sont ramenés à 2 (une ligne blanche).
function shapeText(sp) {
  const txBody = sp?.["p:txBody"];
  if (!txBody) return "";                                  // shape sans texte (image, trait…)
  const paras = asArray(txBody["a:p"]);
  const lines = [];
  for (const p of paras) {
    let line = "";
    for (const r of asArray(p["a:r"])) {
      const t = r?.["a:t"];
      if (typeof t === "string") line += t;                // cas courant : <a:t>texte</a:t> (un a:t numérique, parsé en nombre, n'entre pas ici)
      // cas où a:t porte des attributs : le parseur donne { "#text": "…", "@_…": … }
      else if (t != null && typeof t === "object" && "#text" in t) line += String(t["#text"]);
    }
    // Champs (rare dans les notes utiles) : ignorés s'ils portent le numéro de diapo,
    // sinon on prend leur texte.
    for (const f of asArray(p["a:fld"])) {
      if (f?.["@_type"] === "slidenum") continue;          // numéro de diapo auto : sans intérêt pour l'instructeur
      const t = f?.["a:t"];
      if (typeof t === "string") line += t;
    }
    lines.push(line);
  }
  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim(); // trim : retire les lignes vides en tête/fin
}

// Type de placeholder d'un shape (ou null).
//   Chemin : p:sp > p:nvSpPr > p:nvPr > p:ph/@type. Attention : un p:ph SANS @type (type par défaut
//   "obj" selon la norme) renvoie aussi null ici — sans conséquence, seul le test SKIP_PH l'utilise.
function phType(sp) {
  return sp?.["p:nvSpPr"]?.["p:nvPr"]?.["p:ph"]?.["@_type"] ?? null;
}

// Toutes les notes rédigées d'une diapo de notes -> chaîne (ou "").
// notesTextFromNotesSlide(xml) : xml = contenu texte de ppt/notesSlides/notesSlideN.xml.
//   Parcourt les p:sp de premier niveau du spTree (les shapes GROUPÉS, p:grpSp, ne sont pas visités),
//   écarte les placeholders techniques, et joint les textes restants par une ligne blanche.
//   Peut lever une exception si le XML est invalide : c'est extractPptxNotes qui l'appelle… sans
//   try/catch autour (voir la remarque plus bas).
function notesTextFromNotesSlide(xml) {
  const doc = parser.parse(xml);
  const spTree = doc?.["p:notes"]?.["p:cSld"]?.["p:spTree"];
  if (!spTree) return "";
  const parts = [];
  for (const sp of asArray(spTree["p:sp"])) {             // ordre du document = ordre d'empilement
    if (SKIP_PH.has(phType(sp))) continue;
    const t = shapeText(sp);
    if (t) parts.push(t);
  }
  return parts.join("\n\n").trim();
}

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 3 — Résolution des chemins de relations OPC
// ════════════════════════════════════════════════════════════════════════════════════════════════

// Résout une cible de relation relative au dossier de la partie source.
// resolveRel(baseDir, target) -> chemin normalisé dans le ZIP (sans "/" initial), ex.
//   resolveRel("ppt/slides", "../notesSlides/notesSlide1.xml") -> "ppt/notesSlides/notesSlide1.xml".
//   Règle OPC (Open Packaging Conventions, ECMA-376 Partie 2) : un Target relatif se résout par
//   rapport au DOSSIER de la partie qui porte le .rels (pas par rapport au dossier _rels).
//   Normalisation manuelle de « . » et « .. » (pas de path.posix pour rester neutre vis-à-vis de l'OS).
//   Limite : un Target ABSOLU ("/ppt/notesSlides/…", autorisé par OPC mais rare) serait concaténé
//   à baseDir au lieu de repartir de la racine. Un TargetMode="External" n'est pas filtré ici.
function resolveRel(baseDir, target) {
  // target peut être "../notesSlides/notesSlide1.xml" ; baseDir "ppt/slides".
  const parts = (baseDir + "/" + target).split("/");
  const out = [];
  for (const p of parts) {
    if (p === "" || p === ".") continue;                   // segments vides (//) ou « . » : ignorés
    if (p === "..") out.pop();                             // remonte d'un dossier
    else out.push(p);
  }
  return out.join("/");
}

// ════════════════════════════════════════════════════════════════════════════════════════════════
// SECTION 4 — Point d'entrée : parcours présentation -> diapos -> notes
// ════════════════════════════════════════════════════════════════════════════════════════════════

/**
 * extractPptxNotes(pptxPath) -> [{ page, body }] (pages avec note non vide, triées).
 * Best-effort : renvoie [] si le fichier n'est pas un pptx exploitable.
 */
// Étapes :
//   1. lire et décompresser tout le ZIP en mémoire (objet { "chemin/dans/zip": Uint8Array }) ;
//   2. table rId -> cible depuis presentation.xml.rels ;
//   3. liste ordonnée des diapos depuis presentation.xml (sldIdLst) ;
//   4. pour chaque diapo : son .rels -> relation notesSlide -> XML des notes -> texte.
// Robustesse réelle : seules la lecture/décompression (étape 1) est sous try/catch. Un XML invalide
//   plus loin ferait lever parser.parse() ; fast-xml-parser étant tolérant, c'est rare, mais
//   l'appelant côté LMS a intérêt à entourer l'appel d'un try/catch.
export function extractPptxNotes(pptxPath) {
  let zip;
  try {
    zip = unzipSync(new Uint8Array(readFileSync(pptxPath))); // fichier absent / pas un ZIP -> []
  } catch {
    return [];
  }
  // text(name) : partie du ZIP décodée en UTF-8 (encodage des parties XML Office), ou null si absente.
  //   Les noms de ZIP sont sensibles à la casse et sans "/" initial.
  const text = (name) => (zip[name] ? Buffer.from(zip[name]).toString("utf8") : null);

  const presXml = text("ppt/presentation.xml");
  const presRels = text("ppt/_rels/presentation.xml.rels");
  if (!presXml || !presRels) return [];                    // pas une présentation PowerPoint (ex. .docx renommé)

  // rId -> cible (relatif à ppt/)
  //   Contient TOUTES les relations de la présentation (masques, thème, diapos…) : on n'y pioche
  //   ensuite que les rId cités par sldIdLst, donc pas besoin de filtrer par Type ici.
  const relMap = new Map();
  for (const r of asArray(parser.parse(presRels)?.Relationships?.Relationship)) {
    if (r?.["@_Id"] && r?.["@_Target"]) relMap.set(r["@_Id"], r["@_Target"]);
  }

  // Ordre de présentation : sldIdLst -> rId -> slideN.xml
  const slideFiles = [];
  for (const sld of asArray(parser.parse(presXml)?.["p:presentation"]?.["p:sldIdLst"]?.["p:sldId"])) {
    const rid = sld?.["@_r:id"];                          // attribut r:id (espace de noms relationships)
    // rId introuvable : diapo sautée ET les suivantes remontent d'un rang (décalage des pages)
    const target = rid && relMap.get(rid);
    if (target) slideFiles.push(resolveRel("ppt", target)); // ex: ppt/slides/slide1.xml
  }
  if (!slideFiles.length) return [];

  const notes = [];
  slideFiles.forEach((slidePath, i) => {
    const page = i + 1;                                    // rang 1-based = n° de page du PDF du lecteur
    // (dans ce forEach, « return » = passer à la diapo suivante)
    // rels de la diapo : ppt/slides/_rels/slideN.xml.rels
    const dir = slidePath.slice(0, slidePath.lastIndexOf("/")); // ppt/slides
    const base = slidePath.slice(slidePath.lastIndexOf("/") + 1); // slideN.xml
    const relsXml = text(`${dir}/_rels/${base}.rels`);
    if (!relsXml) return;                                  // diapo sans aucune relation : pas de notes
    let notesTarget = null;
    // Type complet : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide".
    // endsWith() accepte aussi la variante « Strict » (purl.oclc.org/ooxml/...) du même type.
    for (const r of asArray(parser.parse(relsXml)?.Relationships?.Relationship)) {
      if (typeof r?.["@_Type"] === "string" && r["@_Type"].endsWith("/notesSlide")) {
        notesTarget = resolveRel(dir, r["@_Target"]);
        break;
      }
    }
    if (!notesTarget) return;                              // diapo sans page de notes
    const notesXml = text(notesTarget);
    if (!notesXml) return;                                 // relation cassée (partie absente du ZIP)
    const body = notesTextFromNotesSlide(notesXml);
    // notes vides ignorées ; plafond de 8000 caractères (unités UTF-16) pour borner le stockage côté LMS
    if (body) notes.push({ page, body: body.slice(0, 8000) });
  });

  return notes;
}
