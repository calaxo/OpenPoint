// ═══ Dé-duplication des images embarquées (data: URI) des SVG de diapos ═══
// POURQUOI : le moteur embarque chaque image en base64 dans le SVG de CHAQUE diapo qui l'utilise
// (images du master / layout comprises). Un deck de 66 diapos produisait ~55 Mo de HTML, surtout des
// copies de la même image.
//
// COMMENT : chaque data: URI distincte est rangée UNE fois dans une liste ; dans le SVG elle est
// remplacée par un jeton court « op-img:<index> ». Le runtime (renderer/viewer/common.js,
// OP.resolveImages) remplace les jetons par des URL blob: avant d'insérer le SVG dans la page.
//
// API :
//   const store = createImageStore();
//   const svg2 = store.intern(svg);   // SVG avec jetons à la place des data: URI
//   store.list                         // tableau des data: URI distinctes (index = n° du jeton)
//   IMAGE_TOKEN                        // préfixe du jeton ('op-img:'), partagé avec common.js
//
// Le moteur n'émet que des data: URI base64 (pptx/parts.mjs imageHref) : la regex ne vise que ce cas.

export const IMAGE_TOKEN = 'op-img:';
const DATA_URI = /data:[\w.+-]+\/[\w.+-]+(?:;[\w.+-]+=[\w.+-]+)*;base64,[A-Za-z0-9+/=]+/g;

export function createImageStore() {
  const list = [];
  const index = new Map(); // data URI -> index
  function intern(svg) {
    return String(svg).replace(DATA_URI, (uri) => {
      let i = index.get(uri);
      if (i === undefined) { i = list.length; list.push(uri); index.set(uri, i); }
      return IMAGE_TOKEN + i;
    });
  }
  return { intern, list };
}
