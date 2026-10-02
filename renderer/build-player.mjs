// ═══ [SORTIE · PRÉSENTATION animée] ═══ docs/ARCHITECTURE.md (§3) · utilise pptx-render.mjs + extractTiming (clic)
// Lecteur PPTX animé, autonome, SANS PowerPoint : rendu fidèle (SVG) + timeline rejouée.
// Usage : node renderer/build-player.mjs "specimen/deck A/deck.pptx"
// Sortie : renderer/player-out.html  (clic = animation suivante / diapo suivante ; ← → r f)

import { renderSlide, extractTiming, extractTransition } from './pptx-render.mjs';
import { readFileSync, writeFileSync, readdirSync, mkdirSync, existsSync, rmSync } from 'node:fs';
import { join, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';
import { convertMetafiles } from './metafiles.mjs';
import { fontsCssFor } from './web-fonts.mjs'; // polices du LMS (woff2 + métriques PowerPoint)

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptxArg = process.argv[2];
if (!pptxArg) { console.error('donne un .pptx'); process.exit(1); }
const name = basename(pptxArg).replace(/\.pptx$/i, '').replace(/[^\w.-]/g, '_');
const workDir = join(ROOT, 'renderer', '_work', name);
if (existsSync(workDir)) rmSync(workDir, { recursive: true, force: true });
const zip = unzipSync(new Uint8Array(readFileSync(pptxArg)));
for (const [p, data] of Object.entries(zip)) { if (p.endsWith('/')) continue; const o = join(workDir, p); mkdirSync(dirname(o), { recursive: true }); writeFileSync(o, Buffer.from(data)); }
const PPT = join(workDir, 'ppt');
convertMetafiles(join(PPT, 'media')); // EMF/WMF -> PNG (diagrammes vectoriels)

const nums = readdirSync(join(PPT, 'slides')).filter(f => /^slide\d+\.xml$/.test(f)).map(f => +f.match(/\d+/)[0]).sort((a, b) => a - b);
const slides = nums.map(n => {
  const { svg } = renderSlide(PPT, n, { embedImages: true });
  const { steps, animatedSpids } = extractTiming(PPT, n);
  // transition de DIAPO (p:transition) : jouée au changement de diapo (morph/fade/push…)
  const trans = extractTransition(PPT, n);
  return { n, svg, steps, animatedSpids, trans };
});

// Polices : exactement celles servies en production, limitées aux familles citées par les diapos (voir web-fonts.mjs).
const fonts = fontsCssFor(slides.map((s) => s.svg).join(''));
// PAS de rétrécissement de texte côté navigateur (comme en production) : PowerPoint ne recalcule pas
// l'autofit à l'affichage, il applique le fontScale ENREGISTRÉ dans le fichier — que le moteur a déjà
// reporté dans le SVG. Rétrécir ici faisait diverger nos captures du rendu attendu.

const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><title>${esc(name)}</title>
<style>${fonts}
*{box-sizing:border-box}html,body{margin:0;height:100%;background:#0c0d10;overflow:hidden;font:14px system-ui}
#stage{position:absolute;inset:0;display:flex;align-items:center;justify-content:center}
#slide{width:min(100vw,177.7vh);aspect-ratio:16/9;position:relative;box-shadow:0 10px 40px #000}
#slide svg{position:absolute;inset:0;width:100%;height:100%}
#slide .anim-hidden{opacity:0}
#hud{position:absolute;left:12px;bottom:10px;color:#8a94a6;font-size:12px;user-select:none}
#hint{position:absolute;right:12px;bottom:10px;color:#667;font-size:12px}
.eff{transition-timing-function:cubic-bezier(.4,0,.2,1)}
</style></head><body>
<div id="stage"><div id="slide"></div></div>
<div id="hud"></div><div id="hint">clic / → : suivant &nbsp; ← : diapo préc. &nbsp; r : rejouer &nbsp; f : plein écran</div>
<script>
const SLIDES = ${JSON.stringify(slides)};
const slideEl = document.getElementById('slide'), hud = document.getElementById('hud');
let si = 0, step = 0;

// éléments cibles d'un effet : paragraphes (bldP, data-par dans la plage) ou forme entière
function targetsOf(e){
  const host = slideEl.querySelector('[data-spid="'+CSS.escape(e.spid)+'"]');
  if (!host) return [];
  if (e.par) return [...host.querySelectorAll('[data-par]')]
    .filter(d => { const k = +d.getAttribute('data-par'); return k >= e.par.from && k <= e.par.to; });
  return [host];
}
// Transition de DIAPO (p:transition) rejouée au changement de diapo. Le deck d'exemple (gabarit Slidesgo)
// n'a AUCUNE animation d'objet (<p:timing>) mais une transition morph (repli fade) sur CHAQUE diapo : sans ceci,
// avancer d'une diapo ne montrerait aucune animation. Approximation : fondu + léger zoom pour morph,
// fondu pour fade, glissement pour push/wipe (selon @dir). Le conteneur #slide est animé en CSS.
function playSlideTransition(t){
  slideEl.style.transition = 'none';
  if (!t){ slideEl.style.opacity = '1'; slideEl.style.transform = ''; return; }
  const dur = Math.max(120, t.dur || 800);
  let from = { opacity: '0', transform: '' };
  if (t.type === 'morph') from = { opacity: '0', transform: 'scale(1.04)' };
  else if (t.type === 'push' || t.type === 'cover' || t.type === 'wipe' || t.type === 'reveal' || t.type === 'split') {
    const p = (t.dir === 'l' || t.dir === 'u') ? 8 : -8;
    from = { opacity: '0', transform: (t.dir === 'u' || t.dir === 'd') ? 'translateY(' + p + '%)' : 'translateX(' + p + '%)' };
  }
  Object.assign(slideEl.style, from);
  slideEl.getBoundingClientRect(); // reflow -> la transition part de l'état initial
  requestAnimationFrame(() => {
    slideEl.style.transition = 'opacity ' + dur + 'ms ease-out, transform ' + dur + 'ms ease-out';
    slideEl.style.opacity = '1'; slideEl.style.transform = 'none';
  });
}
function loadSlide(i){
  si = (i + SLIDES.length) % SLIDES.length; step = 0;
  const s = SLIDES[si];
  slideEl.innerHTML = s.svg;
  // cache les formes entières qui ont une entrée (apparaissent au clic)
  const hidden = new Set(s.animatedSpids);
  slideEl.querySelectorAll('[data-spid]').forEach(g => { if (hidden.has(g.getAttribute('data-spid'))) g.classList.add('anim-hidden'); });
  // cache les paragraphes qui ont une entrée par paragraphe (bldP)
  for (const st of s.steps) for (const e of st.effects)
    if (e.cls === 'entr' && e.par) targetsOf(e).forEach(el => el.classList.add('anim-hidden'));
  playSlideTransition(s.trans);
  updateHud();
}
function updateHud(){ const s=SLIDES[si]; hud.textContent = 'diapo '+s.n+'   ('+(si+1)+'/'+SLIDES.length+')'+(s.steps.length?'   clic '+step+'/'+s.steps.length:''); }

// mouvement/rotation : animateTransform SMIL additif (se compose avec le transform de base,
// contrairement à une transform CSS qui l'écraserait). Rotation centrée via getBBox.
function animateMotion(g, e){
  g.classList.remove('anim-hidden');
  const m = e.motion, NS = 'http://www.w3.org/2000/svg';
  const svg = slideEl.querySelector('svg'), vb = svg.viewBox.baseVal;
  let bb; try { bb = g.getBBox(); } catch(_) { bb = {x:0,y:0,width:0,height:0}; }
  const cx = (bb.x+bb.width/2).toFixed(1), cy = (bb.y+bb.height/2).toFixed(1);
  const at = document.createElementNS(NS, 'animateTransform');
  at.setAttribute('attributeName','transform'); at.setAttribute('additive','sum');
  at.setAttribute('dur', Math.max(0.1,(m.dur||1000)/1000)+'s'); at.setAttribute('begin','indefinite'); at.setAttribute('fill','freeze');
  if (m.kind==='move'){ at.setAttribute('type','translate');
    const dx=(m.dx*vb.width).toFixed(1), dy=(m.dy*vb.height).toFixed(1);
    at.setAttribute('values', m.rev ? '0 0;'+dx+' '+dy+';0 0' : '0 0;'+dx+' '+dy);
  } else if (m.kind==='rot'){ at.setAttribute('type','rotate');
    const p=' '+cx+' '+cy;
    at.setAttribute('values', m.rev ? '0'+p+';'+m.deg+p+';0'+p : '0'+p+';'+m.deg+p);
  }
  if (m.rev) at.removeAttribute('fill');
  g.appendChild(at); at.beginElement();
}
function applyEffect(g, e){
  if (e.motion){ animateMotion(g, e); return; }
  g.classList.add('eff');
  const dur = Math.max(120, e.dur||500);
  if (e.cls === 'exit'){ g.style.transition='opacity '+dur+'ms'; g.style.opacity='0'; return; }
  // état initial selon le filtre, puis transition vers l'état final
  const f = e.filter||'fade';
  const start = {opacity:'0'}, end = {opacity:'1', transform:'', clipPath:''};
  let props = 'opacity '+dur+'ms';
  if (/wipe\\(left/.test(f)) { start.clipPath='inset(0 100% 0 0)'; end.clipPath='inset(0 0 0 0)'; props='clip-path '+dur+'ms, opacity '+Math.min(200,dur)+'ms'; }
  else if (/wipe\\(right/.test(f)) { start.clipPath='inset(0 0 0 100%)'; end.clipPath='inset(0 0 0 0)'; props='clip-path '+dur+'ms, opacity '+Math.min(200,dur)+'ms'; }
  else if (/wipe\\(up/.test(f)) { start.clipPath='inset(100% 0 0 0)'; end.clipPath='inset(0 0 0 0)'; props='clip-path '+dur+'ms, opacity '+Math.min(200,dur)+'ms'; }
  else if (/wipe\\(down/.test(f)) { start.clipPath='inset(0 0 100% 0)'; end.clipPath='inset(0 0 0 0)'; props='clip-path '+dur+'ms, opacity '+Math.min(200,dur)+'ms'; }
  else if (/circle/.test(f)) { start.clipPath='circle(0% at 50% 50%)'; end.clipPath='circle(75% at 50% 50%)'; props='clip-path '+dur+'ms, opacity '+Math.min(200,dur)+'ms'; }
  else if (/box|split/.test(f)) { start.clipPath='inset(50%)'; end.clipPath='inset(0)'; props='clip-path '+dur+'ms, opacity '+Math.min(200,dur)+'ms'; }
  else if (/zoom/i.test(f)||e.zoom) { start.transform='scale(.3)'; start.opacity='0'; props='transform '+dur+'ms, opacity '+dur+'ms'; }
  Object.assign(g.style, {transition:'none'}, start);
  g.getBoundingClientRect(); // reflow
  requestAnimationFrame(()=>{ g.style.transition=props; g.classList.remove('anim-hidden'); Object.assign(g.style, end); });
}

function next(){
  const s = SLIDES[si];
  if (step < s.steps.length){
    for (const e of s.steps[step].effects){
      for (const g of targetsOf(e)) setTimeout(()=>applyEffect(g,e), e.delay||0);
    }
    step++; updateHud();
    // à la DERNIÈRE étape : révéler tout élément resté caché (jamais ciblé par une étape). L'état final
    // doit être complet comme PowerPoint — le mapping étape->cible peut être imparfait, mais rien ne doit
    // rester invisible à la fin (ex. carte « Expliquer » + pastilles de l'Objectif VENT d.5).
    if (step === s.steps.length)
      slideEl.querySelectorAll('.anim-hidden').forEach(g => { g.style.transition='opacity 200ms'; g.classList.remove('anim-hidden'); g.style.opacity='1'; });
  } else loadSlide(si+1);
}
addEventListener('click', next);
addEventListener('keydown', ev=>{
  if (ev.key==='ArrowRight'||ev.key===' ') next();
  else if (ev.key==='ArrowLeft') loadSlide(si-1);
  else if (ev.key==='r') loadSlide(si);
  else if (ev.key==='f') (document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen());
});
loadSlide(0);
</script></body></html>`;

writeFileSync(join(ROOT, 'renderer', 'player-out.html'), html);
function esc(s){return String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));}
console.log(`OK -> renderer/player-out.html (${slides.length} diapos, ${slides.filter(s=>s.steps.length).length} animées)`);
