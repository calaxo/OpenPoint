// ═══ Viewer PRÉSENTATION (player-out.html) : une diapo à la fois, animations et transitions rejouées ═══
// Inliné par renderer/build-player.mjs (via lib/html-page.mjs), après common.js (global OP).
//
// MODÈLE (à lire avant de modifier)
//   1. Les étapes de extractTiming sont regroupées en « clics » (groups) : une étape 'click' ouvre un
//      groupe ; une étape 'withPrev' / 'afterPrev' s'ajoute au groupe précédent (elle part avec / après
//      lui). Si la 1re étape n'est pas 'click', le groupe 0 est AUTOMATIQUE : joué dès l'entrée.
//   2. L'état visible d'une diapo après p groupes joués se CALCULE (stateAt(p)) : visibilité,
//      translation, rotation, échelle de chaque cible. Il est appliqué d'un coup au DOM (paint).
//   3. Jouer un groupe = appliquer tout de suite l'état FINAL du groupe, puis lancer des animations
//      Web Animations (element.animate) qui partent de l'état précédent (délais compris). Aucune
//      minuterie : annuler les animations (finishRunning) donne instantanément l'état final, d'où
//      un clic rapide ou un changement de diapo qui ne laisse jamais d'effet « en retard ».
//   4. Fin de diapo : tout ce qui n'a pas d'effet de sortie est visible (stateAt(groups.length)).
//
// CIBLES : forme entière = <g data-spid> enveloppé dans un <g class="op-t"> (le transform CSS du
// wrapper ne remplace pas l'attribut transform de la forme) ; paragraphe = <div data-par> du texte.
// Clé d'état : 's<spid>' (forme) ou 'p<spid>:<n°>' (paragraphe).
//
// CLAVIER : → ↓ Espace PageDown Entrée = suivant · ← ↑ PageUp Retour arrière = précédent
//           Début / Fin · chiffres + Entrée = aller à la diapo · f = plein écran · r = rejouer la diapo
//           Échap = annule la saisie / quitte le plein écran.
// SOURIS / TACTILE : clic ou tap = suivant (tap sur le tiers gauche = précédent) ; balayage
//   horizontal = suivant / précédent ; bouger la souris ou toucher affiche les contrôles.
// URL : #3 = diapo 3 (numéro dans l'ordre de présentation, comme le viewer défilement), tenue à jour.
// Diapos masquées (p:sld@show="0") sautées, comme le diaporama PowerPoint.
/* global OP */
'use strict';

(function () {
  const SVGNS = 'http://www.w3.org/2000/svg';
  const data = OP.readData();
  const all = data.slides;
  const shownAll = all.filter((s) => !s.hidden);
  const shown = shownAll.length ? shownAll : all; // tout masqué : on montre tout plutôt que rien

  const slideEl = document.getElementById('slide');
  const ui = {
    counter: document.getElementById('counter'),
    step: document.getElementById('step'),
    bar: document.querySelector('#progress > div'),
    jump: document.getElementById('jumpbox'),
  };

  // état de navigation
  let cur = -1;      // index dans shown
  let pos = 0;       // nombre de groupes joués sur la diapo courante
  let groups = [];   // groupes de la diapo courante (buildGroups)
  let layer = null;  // calque (div) de la diapo courante dans #slide
  let targets = null; // clé -> élément DOM de la diapo courante
  let initial = null; // clé -> état initial

  // ─── 1. regroupement des étapes ───
  function effectDur(e) {
    const m = e.motion;
    const md = m ? (m.dur || 1000) * (m.rev ? 2 : 1) : 0;
    return Math.max(e.dur || 500, md);
  }
  function buildGroups(slide) {
    const out = [];
    let prevStart = 0, prevEnd = 0;
    for (const st of slide.steps || []) {
      let start = 0;
      if (st.trigger === 'click' || !out.length) out.push({ auto: st.trigger !== 'click', items: [], end: 0 });
      else start = st.trigger === 'withPrev' ? prevStart : prevEnd;
      const g = out[out.length - 1];
      let end = start;
      for (const e of st.effects) {
        if (!e.spid) continue; // son, média… : rien à afficher
        const at = start + (e.delay || 0);
        g.items.push({ e, at });
        end = Math.max(end, at + effectDur(e));
      }
      g.end = Math.max(g.end, end);
      prevStart = start; prevEnd = end;
    }
    for (const g of out) g.items.sort((a, b) => a.at - b.at);
    return out;
  }

  // ─── 2. cibles et état ───
  function keysOf(e) {
    if (!e.par) return ['s' + e.spid];
    const keys = [];
    for (let k = e.par.from; k <= e.par.to; k++) keys.push('p' + e.spid + ':' + k);
    return keys;
  }
  function wrapperOf(g) {
    const parent = g.parentNode;
    if (parent && parent.classList && parent.classList.contains('op-t')) return parent;
    const w = document.createElementNS(SVGNS, 'g');
    w.setAttribute('class', 'op-t');
    parent.insertBefore(w, g);
    w.appendChild(g);
    return w;
  }
  // clé -> élément, pour toutes les cibles de la diapo (les clés sans élément sont ignorées)
  function collectTargets(root) {
    const map = new Map();
    for (const g of groups) for (const { e } of g.items) {
      const host = root.querySelector('[data-spid="' + CSS.escape(String(e.spid)) + '"]');
      if (!host) continue;
      if (!e.par) { map.set('s' + e.spid, wrapperOf(host)); continue; }
      host.querySelectorAll('[data-par]').forEach((d) => {
        const k = +d.getAttribute('data-par');
        if (k >= e.par.from && k <= e.par.to) map.set('p' + e.spid + ':' + k, d);
      });
    }
    return map;
  }
  const blank = () => ({ vis: true, exited: false, tx: 0, ty: 0, rot: 0, sx: 1, sy: 1 });
  // état initial : une cible dont le 1er effet de visibilité est une ENTRÉE est cachée au départ
  function initialState() {
    const st = new Map(), decided = new Set();
    for (const key of targets.keys()) st.set(key, blank());
    for (const g of groups) for (const { e } of g.items) for (const key of keysOf(e)) {
      if (!st.has(key) || decided.has(key)) continue;
      if (e.cls === 'entr') { st.get(key).vis = false; decided.add(key); } else if (e.cls === 'exit') decided.add(key);
    }
    return st;
  }
  const copyState = (m) => new Map([...m].map(([k, v]) => [k, Object.assign({}, v)]));
  // effet -> nouvel état d'une cible (les effets « aller-retour » (rev) ne changent pas l'état final)
  function applyTo(s, e, slide) {
    if (e.cls === 'entr') s.vis = true;
    if (e.cls === 'exit') { s.vis = false; s.exited = true; }
    const m = e.motion;
    if (m && !m.rev && m.kind === 'move') { s.tx += m.dx * slide.w; s.ty += m.dy * slide.h; }
    if (m && !m.rev && m.kind === 'rot') s.rot += m.deg;
    if (e.scale) { s.sx *= e.scale.x; s.sy *= e.scale.y; } // animScale (si le moteur fournit e.scale)
  }
  function stateAt(p) {
    const slide = shown[cur];
    const st = copyState(initial);
    for (let i = 0; i < p && i < groups.length; i++) {
      for (const { e } of groups[i].items) for (const key of keysOf(e)) if (st.has(key)) applyTo(st.get(key), e, slide);
    }
    // fin de diapo : rien ne reste caché sauf ce qui est sorti (le lien étape -> cible peut être imparfait)
    if (p >= groups.length) for (const s of st.values()) if (!s.exited) s.vis = true;
    return st;
  }
  const tf = (s) => 'translate(' + s.tx + 'px, ' + s.ty + 'px) rotate(' + s.rot + 'deg) scale(' + s.sx + ', ' + s.sy + ')';
  const isIdentity = (s) => !s.tx && !s.ty && !s.rot && s.sx === 1 && s.sy === 1;
  function paint(st) {
    for (const [key, el] of targets) {
      const s = st.get(key);
      el.classList.toggle('op-hidden', !s.vis);
      el.style.transform = isIdentity(s) ? '' : tf(s);
    }
  }

  // ─── 3. animations des effets ───
  // filtre PowerPoint (animEffect@filter) -> [clip-path de départ, clip-path d'arrivée] ou null (= fondu)
  function clipFrames(filter) {
    const f = String(filter || '');
    if (/^wipe\(left/.test(f)) return ['inset(0 100% 0 0)', 'inset(0 0% 0 0)'];
    if (/^wipe\(right/.test(f)) return ['inset(0 0 0 100%)', 'inset(0 0 0 0%)'];
    if (/^wipe\(up/.test(f)) return ['inset(100% 0 0 0)', 'inset(0% 0 0 0)'];
    if (/^wipe\(down/.test(f)) return ['inset(0 0 100% 0)', 'inset(0 0 0% 0)'];
    if (/^circle/.test(f)) return ['circle(0% at 50% 50%)', 'circle(75% at 50% 50%)'];
    if (/^diamond/.test(f)) return ['polygon(50% 50%, 50% 50%, 50% 50%, 50% 50%)', 'polygon(50% -50%, 150% 50%, 50% 150%, -50% 50%)'];
    if (/^barn\(.*[Vv]ertical/.test(f)) return ['inset(0 50% 0 50%)', 'inset(0 0% 0 0%)'];
    if (/^barn/.test(f)) return ['inset(50% 0 50% 0)', 'inset(0% 0 0% 0)'];
    if (/^box|^split/.test(f)) return ['inset(50% 50% 50% 50%)', 'inset(0% 0% 0% 0%)'];
    return null;
  }
  const isZoom = (e) => !!e.zoom || /zoom/i.test(String(e.filter || ''));

  // keyframes d'un effet pour une cible, de l'état `a` vers l'état `b`
  function framesFor(e, a, b) {
    const ta = tf(a), tb = tf(b);
    if (e.cls === 'entr' || e.cls === 'exit') {
      const clip = clipFrames(e.filter);
      let k0, k1;
      if (clip) { k0 = { clipPath: clip[0], opacity: 1 }; k1 = { clipPath: clip[1], opacity: 1 }; }
      else if (isZoom(e)) { k0 = { opacity: 0, transform: ta + ' scale(0.3)' }; k1 = { opacity: 1, transform: ta + ' scale(1)' }; }
      else { k0 = { opacity: 0 }; k1 = { opacity: 1 }; }
      return e.cls === 'entr' ? [k0, k1] : [k1, k0];
    }
    const m = e.motion;
    if (m && m.rev) { // aller-retour : l'état ne change pas, on passe par le point extrême
      const mid = Object.assign({}, a);
      if (m.kind === 'move') { mid.tx += m.dx * shown[cur].w; mid.ty += m.dy * shown[cur].h; } else mid.rot += m.deg;
      return [{ transform: ta }, { transform: tf(mid) }, { transform: ta }];
    }
    if (ta !== tb) return [{ transform: ta }, { transform: tb }];
    // emphase sans information exploitable : petite pulsation, sans changement d'état
    if (e.cls === 'emph') return [{ transform: ta }, { transform: ta + ' scale(1.08)', offset: 0.5 }, { transform: ta }];
    return null;
  }

  function animate(el, frames, delay, dur, first) {
    if (OP.reducedMotion() || !frames || !el.animate) return;
    // 1re animation de la cible dans le groupe : 'both' (tient l'état de départ pendant son délai) ;
    // suivantes : 'forwards' (la précédente tient l'état pendant leur délai).
    el.animate(frames, { delay, duration: Math.max(1, dur), easing: 'cubic-bezier(.4,0,.2,1)', fill: first ? 'both' : 'forwards' });
  }

  // joue le groupe i : état final appliqué tout de suite, animations depuis l'état d'avant
  function playGroup(i, offset) {
    const slide = shown[cur];
    const before = stateAt(i), after = stateAt(i + 1);
    paint(after);
    const run = copyState(before), seen = new Set();
    for (const { e, at } of groups[i].items) {
      for (const key of keysOf(e)) {
        const el = targets.get(key);
        if (!el) continue;
        const a = Object.assign({}, run.get(key));
        applyTo(run.get(key), e, slide);
        const dur = e.motion ? (e.motion.dur || 1000) * (e.motion.rev ? 2 : 1) : (e.dur || 500);
        animate(el, framesFor(e, a, run.get(key)), offset + at, dur, !seen.has(key));
        seen.add(key);
      }
    }
    // cibles révélées par la règle de fin de diapo sans effet propre : court fondu à la fin
    for (const [key, el] of targets) {
      if (!seen.has(key) && !before.get(key).vis && after.get(key).vis) animate(el, [{ opacity: 0 }, { opacity: 1 }], offset + groups[i].end, 200, true);
    }
  }

  // annule toute animation en cours (l'état de base est déjà l'état final) et retire les anciens calques
  function finishRunning() {
    for (const a of slideEl.getAnimations ? slideEl.getAnimations({ subtree: true }) : []) a.cancel();
    for (const l of [...slideEl.children]) if (l !== layer) l.remove();
  }

  // ─── 4. transition de diapo (p:transition) ───
  // Renvoie sa durée (ms) : les effets automatiques de la nouvelle diapo partent après.
  function playTransition(newL, oldL, t) {
    if (!oldL) return 0;
    const done = () => { if (oldL.parentNode && oldL !== layer) oldL.remove(); };
    if (!t || t.type === 'cut' || OP.reducedMotion() || !newL.animate) { done(); return 0; }
    const dur = Math.max(150, Math.min(t.dur || 700, 3000));
    const opt = { duration: dur, easing: 'ease-in-out', fill: 'both' };
    const d = t.dir || 'l';
    // vecteur de déplacement « vers » la direction (l = vers la gauche : arrive par la droite)
    const vec = { l: ['100%', '0'], r: ['-100%', '0'], u: ['0', '100%'], d: ['0', '-100%'] }[d[0]] || ['100%', '0'];
    const neg = (v) => (v === '0' ? '0' : v.startsWith('-') ? v.slice(1) : '-' + v);
    let anim;
    if (t.type === 'push' || t.type === 'cover') {
      anim = newL.animate([{ transform: 'translate(' + vec[0] + ',' + vec[1] + ')' }, { transform: 'none' }], opt);
      if (t.type === 'push') oldL.animate([{ transform: 'none' }, { transform: 'translate(' + neg(vec[0]) + ',' + neg(vec[1]) + ')' }], opt);
    } else if (t.type === 'reveal') {
      oldL.style.zIndex = '1';
      anim = oldL.animate([{ transform: 'none' }, { transform: 'translate(' + neg(vec[0]) + ',' + neg(vec[1]) + ')' }], opt);
    } else if (t.type === 'wipe') {
      const from = { l: 'inset(0 0 0 100%)', r: 'inset(0 100% 0 0)', u: 'inset(100% 0 0 0)', d: 'inset(0 0 100% 0)' }[d[0]] || 'inset(0 0 0 100%)';
      anim = newL.animate([{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }], opt);
    } else if (t.type === 'split') {
      const from = /^h/.test(String(t.dir || '')) ? 'inset(50% 0 50% 0)' : 'inset(0 50% 0 50%)';
      anim = newL.animate([{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }], opt);
    } else if (t.type === 'zoom') {
      anim = newL.animate([{ opacity: 0, transform: 'scale(.3)' }, { opacity: 1, transform: 'none' }], opt);
    } else if (t.type === 'morph') { // approximation : fondu enchaîné + léger zoom
      anim = newL.animate([{ opacity: 0, transform: 'scale(1.03)' }, { opacity: 1, transform: 'none' }], opt);
    } else { // fade, dissolve et types inconnus : fondu enchaîné
      anim = newL.animate([{ opacity: 0 }, { opacity: 1 }], opt);
    }
    anim.addEventListener('finish', done);
    return dur;
  }

  // ─── 5. navigation ───
  function makeLayer(slide) {
    const l = document.createElement('div');
    l.className = 'layer';
    if (slide.error) l.innerHTML = '<div class="err"><b>Diapo ' + slide.pos + ' : erreur de rendu</b><pre>' + OP.esc(slide.error) + '</pre></div>';
    else l.innerHTML = OP.resolveImages(slide.svg);
    return l;
  }

  // entre dans la diapo i ; atEnd = l'afficher dans son état FINAL (retour arrière), transition = la jouer
  function enter(i, opts) {
    const o = Object.assign({ atEnd: false, transition: true }, opts);
    i = Math.max(0, Math.min(shown.length - 1, i));
    finishRunning();
    const slide = shown[i];
    const old = layer;
    cur = i;
    layer = makeLayer(slide);
    slideEl.style.setProperty('--ratio', String(OP.slideRatio(slide)));
    slideEl.appendChild(layer);
    groups = slide.error ? [] : buildGroups(slide);
    targets = collectTargets(layer);
    initial = initialState();
    pos = o.atEnd ? groups.length : 0;
    paint(stateAt(pos));
    let offset = 0;
    if (o.transition) offset = playTransition(layer, old, slide.trans);
    else if (old) old.remove();
    if (!o.atEnd && groups.length && groups[0].auto) { playGroup(0, offset); pos = 1; }
    preloadNext();
    OP.setHashSlide(slide.pos);
    updateUi();
  }

  function next() {
    if (cur < 0) return;
    if (pos < groups.length) { finishRunning(); playGroup(pos, 0); pos++; updateUi(); }
    else if (cur < shown.length - 1) enter(cur + 1);
    else finishRunning(); // dernière diapo terminée : on reste
  }
  function prev() {
    if (cur < 0) return;
    finishRunning();
    const minPos = groups.length && groups[0].auto ? 1 : 0;
    if (pos > minPos) { pos--; paint(stateAt(pos)); updateUi(); }
    else if (cur > 0) enter(cur - 1, { atEnd: true, transition: false });
  }
  function restart() { enter(cur, { transition: false }); }

  // position de présentation (1-based, masquées comprises) -> index dans shown (la 1re visible à partir de là)
  function indexOfPos(p) {
    for (let i = 0; i < shown.length; i++) if (shown[i].pos >= p) return i;
    return shown.length - 1;
  }

  // décode à l'avance les images de la diapo suivante (évite une transition sur une image vide)
  function preloadNext() {
    const s = shown[cur + 1];
    if (!s || s.error) return;
    const html = OP.resolveImages(s.svg);
    const re = /href="(blob:[^"]+)"/g;
    let m;
    while ((m = re.exec(html))) { const img = new Image(); img.src = m[1]; }
  }

  function updateUi() {
    const s = shown[cur];
    ui.counter.textContent = s.pos + ' / ' + all.length;
    ui.step.textContent = groups.length ? 'clic ' + pos + '/' + groups.length : '';
    ui.bar.style.width = ((cur + 1) / shown.length * 100) + '%';
    document.body.dataset.pos = String(s.pos);   // lus par tools/e2e-viewers.mjs
    document.body.dataset.step = String(pos);
    document.body.dataset.steps = String(groups.length);
  }

  // ─── 6. contrôles auto-masqués ───
  let uiTimer = 0;
  function showUi() {
    document.body.classList.add('show-ui');
    clearTimeout(uiTimer);
    uiTimer = setTimeout(() => document.body.classList.remove('show-ui'), 2500);
  }
  document.addEventListener('mousemove', showUi, { passive: true });

  // saisie d'un numéro de diapo (chiffres puis Entrée)
  let typed = '', typedTimer = 0;
  function setTyped(v) {
    typed = v;
    ui.jump.textContent = typed ? 'aller à la diapo ' + typed : '';
    ui.jump.hidden = !typed;
    clearTimeout(typedTimer);
    if (typed) typedTimer = setTimeout(() => setTyped(''), 3000);
  }

  function toggleFullscreen() {
    if (document.fullscreenElement) document.exitFullscreen();
    else if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => {});
  }

  document.addEventListener('keydown', (ev) => {
    if (ev.ctrlKey || ev.metaKey || ev.altKey) return;
    const k = ev.key;
    if (/^[0-9]$/.test(k)) { setTyped((typed + k).slice(0, 4)); return; }
    if (k === 'Enter' && typed) { const p = +typed; setTyped(''); enter(indexOfPos(p)); ev.preventDefault(); return; }
    if (k === 'Escape') { if (typed) setTyped(''); else if (document.fullscreenElement) document.exitFullscreen(); return; }
    if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter'].includes(k)) next();
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(k)) prev();
    else if (k === 'Home') enter(0);
    else if (k === 'End') enter(shown.length - 1);
    else if (k === 'f' || k === 'F') toggleFullscreen();
    else if (k === 'r' || k === 'R') restart();
    else return; // touche non gérée : comportement normal du navigateur
    ev.preventDefault();
  });

  // tactile : balayage horizontal = suivant / précédent ; tap = suivant (tiers gauche = précédent)
  let touch = null, lastTouchEnd = 0, swiped = false;
  document.addEventListener('touchstart', (ev) => {
    const t = ev.changedTouches[0];
    touch = { x: t.clientX, y: t.clientY };
    showUi();
  }, { passive: true });
  document.addEventListener('touchend', (ev) => {
    if (!touch) return;
    const t = ev.changedTouches[0];
    const dx = t.clientX - touch.x, dy = t.clientY - touch.y;
    touch = null;
    lastTouchEnd = Date.now();
    swiped = Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy) * 1.5;
    if (swiped) { if (dx < 0) next(); else prev(); }
  }, { passive: true });

  document.getElementById('stage').addEventListener('click', (ev) => {
    const fromTouch = Date.now() - lastTouchEnd < 700;
    if (fromTouch && swiped) return; // le balayage a déjà navigué
    if (fromTouch && ev.clientX < window.innerWidth / 3) prev(); else next();
  });

  // boutons des contrôles (ne propagent pas le clic à la scène). blur() : sinon Espace / Entrée
  // « cliqueraient » encore le bouton ET déclencheraient le raccourci clavier (double avance).
  const btn = (id, fn) => document.getElementById(id).addEventListener('click', (ev) => {
    ev.stopPropagation(); ev.currentTarget.blur(); fn(); showUi();
  });
  btn('b-prev', prev);
  btn('b-next', next);
  btn('b-restart', restart);
  btn('b-full', toggleFullscreen);

  window.addEventListener('hashchange', () => {
    const p = OP.hashSlide();
    if (p !== null && p !== shown[cur].pos) enter(indexOfPos(p), { transition: false });
  });

  const start = OP.hashSlide();
  enter(start === null ? 0 : indexOfPos(start), { transition: false });
})();
