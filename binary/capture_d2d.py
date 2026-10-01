# Capture la display-list Direct2D de PowerPoint : spawn PowerPoint sous frida, hooke
# D2D1CreateFactory -> les createurs de render-target du vtable factory -> DrawGlyphRun (slot 29)
# de chaque render target. DrawGlyphRun donne baselineOrigin (position EXACTE) + fontEmSize (taille
# EXACTE) : la mise en page GPU que PowerPoint calcule (autofit/metriques), inaccessible via GDI.
#
# Usage : python binary/capture_d2d.py
import frida, sys, time, json

POWERPNT = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
DECK = r"specimen/sample.pptx"
OUT = r"pp_d2d.json"

JS = r"""
// vtable helper : hooke le slot `idx` de l'objet COM `obj` une seule fois par vtable
const hookedVT = new Set();
function hookSlot(obj, idx, name, cb){
  try{
    const vt = obj.readPointer();
    const key = vt.toString()+':'+idx;
    if (hookedVT.has(key)) return; hookedVT.add(key);
    const fn = vt.add(idx*Process.pointerSize).readPointer();
    Interceptor.attach(fn, cb);
    send({t:'log', m:'hook '+name+' @vt['+idx+'] '+fn});
  }catch(e){ send({t:'log', m:'hookSlot err '+name+' '+e}); }
}

function bitsToFloat(u){ const b=Memory.alloc(4); b.writeU32(u>>>0); return b.readFloat(); }
// RECO : compter les primitives de dessin appelées, pour identifier COMMENT PowerPoint rend.
// Slots ID2D1RenderTarget : 15 DrawLine 16 DrawRect 17 FillRect 20 DrawEllipse 21 FillEllipse
// 22 DrawGeometry 23 FillGeometry 24 FillMesh 25 FillOpacityMask 26 DrawBitmap 27 DrawText
// 28 DrawTextLayout 29 DrawGlyphRun ; DeviceContext : 83 DrawImage 84 DrawGdiMetafile.
// slots RenderTarget (0-56) + DeviceContext (57+). DeviceContext a SON PROPRE DrawGlyphRun @82
// (signature avec DWRITE_GLYPH_RUN_DESCRIPTION -> contient LE TEXTE). 67=CreateCommandList,
// 83=DrawImage, 85=DrawBitmap(DC).
const SLOTS = {15:'DrawLine',17:'FillRect',22:'DrawGeometry',23:'FillGeometry',26:'DrawBitmap',
  27:'DrawText',28:'DrawTextLayout',29:'DrawGlyphRun_RT',63:'CreateEffect',67:'CreateCommandList',
  82:'DrawGlyphRun_DC',83:'DrawImage',84:'DrawGdiMetafile',85:'DrawBitmap_DC'};
const counts = {};
function capGlyph(a, hasDesc){    // baselineOrigin=rdx packé, glyphRun=r8, [desc=r9 si DC]
  try{
    const packed=a[1];
    const fx=bitsToFloat(packed.and(0xffffffff).toNumber()), fy=bitsToFloat(packed.shr(32).and(0xffffffff).toNumber());
    const gr=a[2], em=gr.add(8).readFloat(), n=gr.add(12).readU32();
    let s=null;
    if (hasDesc){ const d=a[3]; if(!d.isNull()){ const strp=d.add(8).readPointer(), len=d.add(16).readU32();
      if(!strp.isNull() && len>0 && len<4096) s=strp.readUtf16String(len); } }
    send({t:'glyph', x:Math.round(fx*10)/10, y:Math.round(fy*10)/10, em:Math.round(em*100)/100, n, s});
  }catch(e){ send({t:'log', m:'capGlyph '+e}); }
}
function hookDrawGlyphRun(rt, label){    // instrumente tous les slots de SLOTS
  for (const idx in SLOTS){
    const nm = SLOTS[idx];
    hookSlot(rt, parseInt(idx), nm+'('+label+')', { onEnter(a){
      counts[nm] = (counts[nm]||0)+1;
      if (nm==='DrawGlyphRun_DC') capGlyph(a, true);
      else if (nm==='DrawGlyphRun_RT') capGlyph(a, false);
    }});
  }
}
function reportCounts(){ send({t:'counts', c:counts}); }
recv('report', reportCounts);

// createurs de render target sur le vtable ID2D1Factory (indices ABI stables) :
// 13 CreateWicBitmapRenderTarget, 14 CreateHwndRenderTarget, 15 CreateDxgiSurfaceRenderTarget,
// 16 CreateDCRenderTarget, 17 CreateDevice (Factory1). ppOut = dernier argument.
function hookFactory(fac){
  const rtCreators = [[13,3,'Wic'],[14,3,'Hwnd'],[15,3,'Dxgi'],[16,2,'DC']];
  for (const [idx, outArg, nm] of rtCreators){
    hookSlot(fac, idx, 'Factory::Create'+nm, { onEnter(a){ this.pp=a[outArg]; }, onLeave(r){
      try{ const rt = this.pp.readPointer(); if(!rt.isNull()) hookDrawGlyphRun(rt, nm); }catch(e){}
    }});
  }
  // CreateDevice -> Device::CreateDeviceContext(slot 4) -> DeviceContext (RT-derive)
  hookSlot(fac, 17, 'Factory1::CreateDevice', { onEnter(a){ this.pp=a[2]; }, onLeave(r){
    try{ const dev = this.pp.readPointer(); if(dev.isNull()) return;
      hookSlot(dev, 4, 'Device::CreateDeviceContext', { onEnter(a){ this.pp=a[2]; }, onLeave(r){
        try{ const dc = this.pp.readPointer(); if(!dc.isNull()) hookDrawGlyphRun(dc, 'DevCtx'); }catch(e){}
      }});
    }catch(e){}
  }});
}

function installOnD2D(d2d){
  const cf = d2d.findExportByName('D2D1CreateFactory');
  if (!cf){ send({t:'log', m:'d2d1 charge mais pas d\'export D2D1CreateFactory ?'}); return; }
  Interceptor.attach(cf, { onEnter(a){ this.pp = a[3]; }, onLeave(r){    // ppIFactory = arg 3
    try{ const fac = this.pp.readPointer(); if(!fac.isNull()){ send({t:'log', m:'D2D1CreateFactory -> '+fac}); hookFactory(fac); } }catch(e){}
  }});
  send({t:'log', m:'D2D1CreateFactory hooke ('+d2d.base+')'});
}
// d2d1.dll pas encore charge au spawn -> on attend son chargement
let done=false;
const iv = setInterval(()=>{ if(done) return; const m=Process.findModuleByName('d2d1.dll'); if(m){ done=true; clearInterval(iv); installOnD2D(m); } }, 8);
send({t:'log', m:'attente chargement d2d1.dll...'});
"""

glyphs=[]; counts={}
def on_message(msg, data):
    p = msg.get('payload')
    if not p: print("[frida]", msg); return
    if p.get('t')=='glyph': glyphs.append(p)
    elif p.get('t')=='log': print("[frida]", p['m'])
    elif p.get('t')=='counts': counts.update(p['c'])

print("spawn PowerPoint sous frida...")
dev = frida.get_local_device()
pid = dev.spawn([POWERPNT, DECK])
session = dev.attach(pid)
sc = session.create_script(JS); sc.on('message', on_message); sc.load()
dev.resume(pid)
print("PowerPoint lance, rendu de la diapo 1 en cours...")
time.sleep(12)   # laisser PowerPoint ouvrir + rendre la diapo
sc.post({'type':'report'}); time.sleep(0.5)

print("\n=== primitives de dessin D2D appelées (reco) ===")
for k,v in sorted(counts.items(), key=lambda x:-x[1]): print(f"  {k:<16} {v}")
print(f"\n=== {len(glyphs)} DrawGlyphRun capturés ===")
for gph in glyphs[:60]:
    txt = ("'"+gph['s'][:40]+"'") if gph.get('s') else f"(n={gph['n']} glyphes)"
    print(f"  pos=({gph['x']:>7},{gph['y']:>7})  em={gph['em']:>6}  {txt}")
json.dump(glyphs, open(OUT,'w',encoding='utf-8'), ensure_ascii=False)
print(f"-> {OUT} ({len(glyphs)} runs)")

try: session.detach()
except: pass
import subprocess; subprocess.run(["taskkill","/F","/PID",str(pid)], capture_output=True)
print("done")
