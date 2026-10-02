// ═══ [PPTX · ANIMATIONS] ═══ extractTiming (<p:timing> -> étapes de clic) et extractTransition (<p:transition>).
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { P, tagOf, at, kids, byTag, first, deepFind } from './xml.mjs';

// ████ §P ─ extractTiming : TIMELINE D'ANIMATION ────────────────────────────────────────────────
// Structure PresentationML d'une timeline (§19.5, p:timing) — arbre de conteneurs temporels :
//   p:timing > p:tnLst > p:par (racine, nodeType="tmRoot") > p:cTn > p:childTnLst
//     > p:seq (séquence principale, nodeType="mainSeq") > p:cTn > p:childTnLst
//       > p:par  <- UNE ÉTAPE (un clic)            [c'est le niveau lu par clickPars]
//         > p:cTn > p:childTnLst > p:par (groupe « avec / après le précédent », stCondLst@delay)
//           > p:par > p:cTn presetID presetClass presetSubtype nodeType  <- UN EFFET
//             > p:childTnLst > comportements : p:set, p:animEffect (filter="fade", "wipe(down)"…),
//               p:anim, p:animMotion (path), p:animRot (by), p:animScale… ; cible p:tgtEl/p:spTgt@spid
//               (+ p:txEl/p:pRg st..end pour une animation par paragraphe).
// Les attributs p:bldP (build par paragraphe) de p:bldLst ne sont pas lus : la plage vient de p:pRg.
//
// ================= extraction de la timeline d'animation (<p:timing>) =================
// Renvoie les "étapes de clic" : [{ effects:[{spid, cls, filter, transition, delay, dur}] }]
// cls = entr|exit|emph ; filter (fade/wipe(..)/circle(out)..) ; delay/dur en ms.
// Forme exacte d'un effet : { cls, spid, par: {from,to}|null, filter, transition, motion, delay, dur }
//   cls   : @presetClass (entr, exit, emph, path, verb, mediacall) ;
//   motion: { kind:'move', dx, dy, dur, rev } | { kind:'rot', deg, dur, rev } | null.
// animatedSpids : spids des formes ENTIÈRES qui ont une entrée -> à cacher avant leur 1re apparition.
// Seule la 1re p:seq trouvée est lue (en pratique la mainSeq) : les séquences interactives (déclencheur
// au clic sur une forme) sont ignorées.
export function extractTiming(pptDir, slideNo) {
  const text = readFileSync(join(pptDir, 'slides', `slide${slideNo}.xml`), 'utf8');
  // extraction par regex (évite de reparser toute la diapo). Suppose une balise <p:timing> SANS
  // attribut ; le glouton [\s\S]* va jusqu'au DERNIER </p:timing> (normalement il n'y en a qu'un).
  const m = text.match(/<p:timing>[\s\S]*<\/p:timing>/);
  if (!m) return { steps: [], animatedSpids: [] };
  const tree = P.parse(m[0]);
  const timing = tree.find((n) => tagOf(n) === 'p:timing');
  const seq = deepFind(timing, 'p:seq');
  if (!seq) return { steps: [], animatedSpids: [] };
  const seqCTn = first(seq, 'p:cTn');
  // chaque p:par enfant direct de la séquence = une étape. @nodeType n'est pas consulté : une étape
  // qui démarre toute seule (« après la précédente » en tête de séquence) est quand même traitée
  // comme un clic.
  const clickPars = byTag(first(seqCTn, 'p:childTnLst'), 'p:par');

  const animatedSpids = new Set();
  const steps = clickPars.map((par) => {
    const effects = [];
    collectEffects(par, 0, effects);
    // entrée ciblant la forme entière (sans plage de paragraphes) -> cachée au départ.
    // Les entrées par paragraphe (bldP) sont gérées séparément par le player (data-par).
    for (const e of effects) if (e.cls === 'entr' && !e.par) animatedSpids.add(e.spid);
    return { effects };
  }).filter((s) => s.effects.length); // étapes sans effet reconnu retirées
  return { steps, animatedSpids: [...animatedSpids] };

  // (déclarations de fonctions hissées : utilisables avant leur définition, après le return)
  // collectEffects(nœud, délai cumulé, accumulateur) : descente récursive. Un p:cTn portant
  // @presetID est un EFFET (feuille : on ne descend pas plus) ; sinon, un p:par conteneur ajoute son
  // délai de départ (stCondLst/cond@delay, en ms) au délai de ses descendants. « indefinite » = attend
  // un clic : ne compte pas. Les délais de cTn qui ne sont PAS dans un p:par (ex. p:seq) sont ignorés.
  function collectEffects(node, accDelay, acc) {
    if (tagOf(node) === 'p:par') {
      const cTn = first(node, 'p:cTn');
      if (cTn && at(cTn)['@presetID'] !== undefined) { acc.push(parseEffect(cTn, accDelay)); return; }
      if (cTn) {
        const cond = deepFind(first(cTn, 'p:stCondLst') ?? {}, 'p:cond');
        const d = cond && at(cond)['@delay'];
        if (d !== undefined && d !== 'indefinite') accDelay += Number(d) || 0;
      }
    }
    for (const c of kids(node)) if (typeof c === 'object') collectEffects(c, accDelay, acc);
  }
  // parseEffect(cTn de l'effet, délai hérité) -> objet effet (voir en-tête de §P).
  // spid / animEffect / pRg : 1ers trouvés en profondeur (un effet multi-cibles n'est lu que pour sa
  // 1re cible). delay = délai hérité + délai propre du cTn.
  function parseEffect(cTn, accDelay) {
    const a = at(cTn);
    const spTgt = deepFind(cTn, 'p:spTgt');
    const animEff = deepFind(cTn, 'p:animEffect');
    const pRg = spTgt ? deepFind(spTgt, 'p:pRg') : null; // plage de paragraphes (bldP)
    const own = deepFind(first(cTn, 'p:stCondLst') ?? {}, 'p:cond');
    const ownDelay = own ? Number(at(own)['@delay']) || 0 : 0;
    // durée : 1er cTn d'un behavior avec dur numérique
    // défaut 500 ms ; « > 1 » écarte les p:set de durée 1 ms (bascule de visibilité instantanée que
    // PowerPoint place en tête des effets d'entrée) pour prendre la durée du vrai comportement
    let dur = 500;
    const durNode = [...allCTn(cTn)].map((c) => at(c)['@dur']).find((d) => d && d !== 'indefinite' && +d > 1);
    if (durNode) dur = +durNode;
    return {
      cls: a['@presetClass'], spid: spTgt ? at(spTgt)['@spid'] : null,
      par: pRg ? { from: +at(pRg)['@st'] || 0, to: +at(pRg)['@end'] || 0 } : null,
      filter: animEff ? at(animEff)['@filter'] : null,
      transition: animEff ? at(animEff)['@transition'] : null,
      motion: parseMotion(cTn),
      delay: accDelay + ownDelay, dur,
    };
  }
  // mouvement/rotation (§19.5.3 animMotion, §19.5.5 animRot) -> {kind, dx/dy (fractions) | deg, dur, rev}
  // behTiming(comportement) -> { dur (ms, défaut 1000), rev (@autoRev : aller-retour) } lus sur son
  // p:cBhvr/p:cTn
  function behTiming(node) { const c = deepFind(node, 'p:cTn'); const a = at(c || {}); return { dur: +a['@dur'] || 1000, rev: a['@autoRev'] === '1' }; }
  // parseMotion : animMotion@path est un chemin type SVG (« M 0 0 L 0.25 0 E ») dont les coordonnées
  // sont des FRACTIONS de la taille de diapo. On n'en garde que le déplacement net (dernier point -
  // premier point) : les trajectoires courbes sont rendues en ligne droite. animRot@by en 1/60000 °.
  // Un effet qui n'a ni l'un ni l'autre -> null.
  function parseMotion(cTn) {
    const mm = deepFind(cTn, 'p:animMotion');
    if (mm) {
      const path = at(mm)['@path'] || '';
      // tous les nombres du chemin (y compris notation 1e-3), dans l'ordre ; les lettres de commande
      // sont ignorées -> pts = [x0, y0, …, xn, yn]
      const pts = (path.match(/-?\d*\.?\d+(?:[eE][-+]?\d+)?/g) || []).map(Number);
      if (pts.length >= 2) {
        const fx = pts[0], fy = pts[1], lx = pts[pts.length - 2], ly = pts[pts.length - 1];
        const t = behTiming(mm); return { kind: 'move', dx: lx - fx, dy: ly - fy, dur: t.dur, rev: t.rev };
      }
    }
    const mr = deepFind(cTn, 'p:animRot');
    if (mr) { const t = behTiming(mr); return { kind: 'rot', deg: (+at(mr)['@by'] || 0) / 60000, dur: t.dur, rev: t.rev }; }
    return null;
  }
  // allCTn : générateur de tous les p:cTn du sous-arbre (le nœud lui-même compris), ordre du document
  function* allCTn(node) {
    if (tagOf(node) === 'p:cTn') yield node;
    for (const c of kids(node)) if (typeof c === 'object') yield* allCTn(c);
  }
}

// ████ §P2 ─ extractTransition : TRANSITION DE DIAPO (<p:transition>) ──────────────────────────
// Beaucoup de decks (dont les gabarits Slidesgo/Google Slides) n'ont AUCUNE animation d'objet
// (<p:timing>) mais une TRANSITION par diapo : <mc:AlternateContent><mc:Choice Requires="p159">
// <p:transition p14:dur="2000"><p159:morph option="byObject"/></p:transition></mc:Choice>
// <mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>.
// On lit la variante moderne (morph) et on retombe sur le fade ; @p14:dur (ms) = durée. Le player
// (build-player.mjs) rejoue cette transition au changement de diapo. Renvoie null si la diapo n'en a pas.
export function extractTransition(pptDir, slideNo) {
  const text = readFileSync(join(pptDir, 'slides', `slide${slideNo}.xml`), 'utf8');
  // 1re <p:transition> = la variante « Choice » (morph) quand il y a un <mc:AlternateContent>, sinon
  // la transition simple. Le glouton non-greedy s'arrête à la 1re fermeture.
  const m = text.match(/<p:transition\b[^>]*>([\s\S]*?)<\/p:transition>/);
  if (!m) return null;
  const opener = m[0].slice(0, m[0].indexOf('>') + 1);
  const dur = +(opener.match(/p14:dur="(\d+)"/)?.[1] || 0); // durée en ms (p14:dur), sinon défaut
  const body = m[1];
  // type : morph si présent, sinon fade, sinon le 1er effet reconnu ; dir pour les poussées/balayages
  const type = /<p159:morph\b/.test(body) ? 'morph'
             : /<p:fade\b/.test(body) ? 'fade'
             : (body.match(/<p:(push|wipe|cover|split|reveal|zoom|dissolve|cut)\b/)?.[1]) || 'fade';
  const dir = body.match(/<p:(?:push|wipe|cover|split)[^>]*\bdir="([^"]+)"/)?.[1] || null;
  return { type, dur: dur || 2000, dir };
}
