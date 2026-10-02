// ═══ Runtime commun aux deux viewers (scroll-viewer.js, player.js) ═══
// Script navigateur classique (pas de module, pas de bundler) : inliné dans la page par
// renderer/lib/html-page.mjs. Expose un seul global : OP.
//
//   OP.readData()           -> objet JSON de <script id="op-data"> (voir lib/deck-data.mjs)
//   OP.resolveImages(svg)   -> SVG dont les jetons « op-img:N » sont remplacés par des URL blob:
//                              (chaque image n'est décodée qu'une fois, à la 1re utilisation)
//   OP.slideRatio(slide)    -> largeur / hauteur de la diapo (repli 16/9)
//   OP.hashSlide()          -> numéro de diapo lu dans l'URL (#3 -> 3) ou null
//   OP.setHashSlide(pos)    -> écrit #pos dans l'URL sans ajouter d'entrée d'historique
//   OP.esc(texte)           -> texte échappé pour innerHTML
//   OP.reducedMotion()      -> true si l'utilisateur demande moins d'animations
/* exported OP */
'use strict';

var OP = (function () {
  var IMAGE_TOKEN = /op-img:(\d+)/g; // même jeton que renderer/lib/images.mjs
  var data = null;
  var urls = []; // index -> URL blob: (ou data: URI en repli)

  function readData() {
    if (!data) data = JSON.parse(document.getElementById('op-data').textContent);
    return data;
  }

  // data:<mime>;base64,<…> -> URL blob: (plus court à insérer, décodé une seule fois)
  function imageUrl(i) {
    if (urls[i]) return urls[i];
    var uri = readData().images[i] || '';
    var url = uri;
    try {
      var comma = uri.indexOf(',');
      var mime = uri.slice(5, uri.indexOf(';'));
      var bin = atob(uri.slice(comma + 1));
      var bytes = new Uint8Array(bin.length);
      for (var k = 0; k < bin.length; k++) bytes[k] = bin.charCodeAt(k);
      url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    } catch (e) { url = uri; } // repli : la data: URI elle-même
    urls[i] = url;
    return url;
  }

  function resolveImages(svg) {
    return String(svg || '').replace(IMAGE_TOKEN, function (_, i) { return imageUrl(+i); });
  }

  function slideRatio(slide) {
    return slide && slide.w && slide.h ? slide.w / slide.h : 16 / 9;
  }

  function hashSlide() {
    var m = /^#(\d+)$/.exec(location.hash || '');
    return m ? +m[1] : null;
  }

  function setHashSlide(pos) {
    var h = '#' + pos;
    if (location.hash !== h) history.replaceState(null, '', h);
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }

  function reducedMotion() {
    return !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);
  }

  return { readData: readData, resolveImages: resolveImages, slideRatio: slideRatio,
    hashSlide: hashSlide, setHashSlide: setHashSlide, esc: esc, reducedMotion: reducedMotion };
})();
