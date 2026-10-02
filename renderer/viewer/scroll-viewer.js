// ═══ Viewer DÉFILEMENT (viewer.html) : toutes les diapos empilées, état final, sans animation ═══
// Inliné par renderer/build.mjs (via lib/html-page.mjs), après common.js (global OP).
//
// Comportement :
//   - une carte par diapo, dans l'ORDRE DE PRÉSENTATION ; étiquette = position (1, 2, 3…) + fichier ;
//   - diapos masquées (p:sld@show="0") AFFICHÉES, grisées, avec un badge « masquée » : ce viewer sert à
//     relire tout le contenu du fichier ; le player, lui, les saute (comme le diaporama PowerPoint) ;
//   - insertion PARESSEUSE : le SVG d'une diapo n'est inséré qu'à l'approche de la zone visible
//     (IntersectionObserver) ; la carte a déjà sa taille finale (aspect-ratio) -> pas de saut de page ;
//   - barre d'outils : nom du deck, nombre de diapos, aller à (n° + Entrée), précédente / suivante ;
//   - clavier : j / → / PageDown = diapo suivante, k / ← / PageUp = précédente ;
//   - URL #3 = aller à la diapo 3 ; l'URL suit le défilement.
/* global OP */
'use strict';

(function () {
  const data = OP.readData();
  const slides = data.slides;
  const deckEl = document.getElementById('deck');
  const countEl = document.getElementById('count');
  const jumpEl = document.getElementById('jump');
  const cards = []; // index i -> élément .card (même ordre que slides)

  document.getElementById('deck-name').textContent = data.name;
  const nHidden = slides.filter(function (s) { return s.hidden; }).length;
  countEl.textContent = slides.length + ' diapos' + (nHidden ? ' · ' + nHidden + ' masquée' + (nHidden > 1 ? 's' : '') : '');
  jumpEl.max = String(slides.length);

  // ─── construction des cartes (vides : le SVG arrive plus tard) ───
  slides.forEach(function (s) {
    const card = document.createElement('section');
    card.className = 'card' + (s.hidden ? ' is-hidden' : '') + (s.error ? ' is-error' : '');
    card.id = 'slide-' + s.pos;
    card.setAttribute('data-pos', String(s.pos));
    card.setAttribute('data-n', String(s.n));
    card.innerHTML = '<div class="lbl"><span class="num">' + s.pos + '</span>'
      + '<span class="file">slide' + s.n + '.xml</span>'
      + (s.hidden ? '<span class="badge">masquée</span>' : '')
      + (s.error ? '<span class="badge err">erreur de rendu</span>' : '') + '</div>'
      + '<div class="frame"></div>';
    const frame = card.querySelector('.frame');
    if (s.error) frame.innerHTML = '<pre>' + OP.esc(s.error) + '</pre>';
    else frame.style.aspectRatio = s.w + ' / ' + s.h;
    deckEl.appendChild(card);
    cards.push(card);
  });

  // ─── insertion paresseuse du SVG ───
  function fill(i) {
    const card = cards[i], s = slides[i];
    if (!card || s.error || card.hasAttribute('data-filled')) return;
    card.setAttribute('data-filled', '');
    card.querySelector('.frame').innerHTML = OP.resolveImages(s.svg);
  }
  if ('IntersectionObserver' in window) {
    const io = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (en.isIntersecting) { fill(cards.indexOf(en.target)); io.unobserve(en.target); }
      });
    }, { rootMargin: '150% 0px' }); // ~1,5 écran d'avance au-dessus et en dessous
    cards.forEach(function (c) { io.observe(c); });
  } else {
    cards.forEach(function (_, i) { fill(i); });
  }

  // ─── navigation ───
  function barHeight() { return document.getElementById('bar').offsetHeight; }

  // index de la diapo « courante » = la dernière dont le haut a passé le bas de la barre
  function currentIndex() {
    const limit = barHeight() + 24;
    let cur = 0;
    for (let i = 0; i < cards.length; i++) {
      if (cards[i].getBoundingClientRect().top <= limit) cur = i; else break;
    }
    return cur;
  }

  function goTo(i) {
    i = Math.max(0, Math.min(cards.length - 1, i));
    fill(i);
    const top = cards[i].getBoundingClientRect().top + window.scrollY - barHeight() - 8;
    window.scrollTo(0, Math.max(0, top));
    OP.setHashSlide(slides[i].pos);
    updateCurrent(i);
  }

  function updateCurrent(i) {
    cards.forEach(function (c, k) { c.classList.toggle('is-current', k === i); });
    if (document.activeElement !== jumpEl) jumpEl.value = String(slides[i].pos);
  }

  // l'URL et l'indicateur suivent le défilement (au plus une mise à jour par image)
  let ticking = false;
  window.addEventListener('scroll', function () {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(function () {
      ticking = false;
      const i = currentIndex();
      OP.setHashSlide(slides[i].pos);
      updateCurrent(i);
    });
  }, { passive: true });

  function indexOfPos(pos) {
    for (let i = 0; i < slides.length; i++) if (slides[i].pos === pos) return i;
    return -1;
  }

  jumpEl.addEventListener('keydown', function (ev) {
    if (ev.key === 'Enter') { const i = indexOfPos(+jumpEl.value); if (i >= 0) goTo(i); jumpEl.blur(); }
    ev.stopPropagation(); // les chiffres / j / k tapés ici ne naviguent pas
  });
  document.getElementById('prev').addEventListener('click', function () { goTo(currentIndex() - 1); });
  document.getElementById('next').addEventListener('click', function () { goTo(currentIndex() + 1); });

  document.addEventListener('keydown', function (ev) {
    if (ev.ctrlKey || ev.metaKey || ev.altKey) return;
    const k = ev.key;
    if (k === 'j' || k === 'ArrowRight' || k === 'PageDown') { ev.preventDefault(); goTo(currentIndex() + 1); }
    else if (k === 'k' || k === 'ArrowLeft' || k === 'PageUp') { ev.preventDefault(); goTo(currentIndex() - 1); }
  });

  function fromHash() {
    const pos = OP.hashSlide();
    const i = pos === null ? -1 : indexOfPos(pos);
    if (i >= 0) goTo(i);
  }
  window.addEventListener('hashchange', fromHash);
  if (OP.hashSlide() !== null) fromHash(); else updateCurrent(0);
})();
