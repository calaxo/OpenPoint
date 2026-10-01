# Capture la DISPLAY LIST de PowerPoint : hooke gdi32!ExtTextOutW/TextOutW pendant que PowerPoint
# rend une diapo (via Slide.Export -> rendu dans un DIB GDI, cf. ocpptview.dll importe
# ExtTextOutW+CreateDIBSection+BitBlt). Donne la mise en page EXACTE (position, texte, police,
# taille, couleur) que PowerPoint calcule — la vérité qu'un moteur maison approxime mal.
#
# Usage : python binary/capture_gdi.py <slideNo>
import frida, comtypes.client, time, os, sys, subprocess, json

DECK = r"specimen/sample.pptx"
SLIDE = int(sys.argv[1]) if len(sys.argv) > 1 else 1
OUT = r"pp_displaylist.json"
PNG = r"pp_export.png"

JS = r"""
const g = Process.getModuleByName('gdi32.dll');  // frida 17 : plus de Module.getExportByName
const GetCurrentObject = new NativeFunction(g.getExportByName('GetCurrentObject'),'pointer',['pointer','uint']);
const GetObjectW       = new NativeFunction(g.getExportByName('GetObjectW'),'int',['pointer','int','pointer']);
const GetTextColor     = new NativeFunction(g.getExportByName('GetTextColor'),'uint',['pointer']);
const GetWorldTransform= new NativeFunction(g.getExportByName('GetWorldTransform'),'int',['pointer','pointer']);
const GetCurrentPositionEx=new NativeFunction(g.getExportByName('GetCurrentPositionEx'),'int',['pointer','pointer']);
const GetTextAlign     = new NativeFunction(g.getExportByName('GetTextAlign'),'uint',['pointer']);

function fontOf(hdc){
  try{
    const buf = Memory.alloc(92); const hf = GetCurrentObject(hdc, 6); // OBJ_FONT
    if (hf.isNull()) return {};
    GetObjectW(hf, 92, buf);
    return { h: buf.readS32(), w: buf.add(16).readS32(), ital: buf.add(20).readU8(),
             face: buf.add(28).readUtf16String() };
  }catch(e){ return {}; }
}
function xformOf(hdc){
  try{ const x = Memory.alloc(24);
    if (GetWorldTransform(hdc, x)) return { sx:x.readFloat(), sy:x.add(12).readFloat(), dx:x.add(16).readFloat(), dy:x.add(20).readFloat() };
  }catch(e){}
  return { sx:1, sy:1, dx:0, dy:0 };
}
function hook(name){
  const p = g.findExportByName(name); if (!p) return;
  Interceptor.attach(p, { onEnter(a){
    // ExtTextOutW(hdc,x,y,opt,rect,str,c,dx) ; TextOutW(hdc,x,y,str,c)
    const isExt = name === 'ExtTextOutW';
    const hdc=a[0], x=a[1].toInt32(), y=a[2].toInt32();
    const strp = isExt ? a[5] : a[3], cnt = (isExt ? a[6] : a[4]).toInt32();
    if (cnt<=0 || cnt>4096) return;
    let s=''; try{ s = strp.readUtf16String(cnt); }catch(e){ return; }
    if (!s || !s.trim()) return;
    const f = fontOf(hdc), t = xformOf(hdc), col = GetTextColor(hdc)>>>0;
    // position courante (si TA_UPDATECP, ExtTextOut ignore x,y et part de la position courante)
    let cpx=0, cpy=0; try{ const pt=Memory.alloc(8); if(GetCurrentPositionEx(hdc,pt)){cpx=pt.readS32();cpy=pt.add(4).readS32();} }catch(e){}
    const ta = GetTextAlign(hdc)>>>0;
    send({t:'txt', x, y, cpx, cpy, ta, s, face:f.face, fh:f.h, fw:f.w, ital:f.ital, col,
          sx:t.sx, sy:t.sy, dx:t.dx, dy:t.dy});
  }});
}
hook('ExtTextOutW'); hook('TextOutW');
send({t:'ready'});
"""

items = []
def on_message(msg, data):
    p = msg.get('payload')
    if not p: print("[frida]", msg); return
    if p.get('t') == 'txt': items.append(p)
    elif p.get('t') == 'ready': print("[frida] hook ExtTextOutW/TextOutW armé")

print("PowerPoint COM...")
app = comtypes.client.CreateObject("PowerPoint.Application")
app.Visible = 1
pres = app.Presentations.Open(DECK, ReadOnly=1, WithWindow=1)
time.sleep(2)
pid = int(subprocess.run(["tasklist","/FI","IMAGENAME eq POWERPNT.EXE","/FO","CSV","/NH"],
      capture_output=True,text=True).stdout.split(',')[1].strip('"'))
print("PID", pid)
session = frida.get_local_device().attach(pid)
sc = session.create_script(JS); sc.on('message', on_message); sc.load()
time.sleep(0.5)

items.clear()
print(f"Export diapo {SLIDE} (force le rendu GDI)...")
try:
    pres.Slides(SLIDE).Export(PNG, "PNG", 1280, 720)
except Exception as ex:
    print("Export err:", ex)
time.sleep(1.5)

print(f"\n=== {len(items)} appels texte capturés ===")
for it in items[:40]:
    dh = round(abs(it.get('fh',0))*it.get('sy',1),1)
    px = it['x'] if it['x'] else it.get('cpx',0)
    py = it['y'] if it['y'] else it.get('cpy',0)
    print(f"  pos=({px:>6},{py:>6}) ta={it.get('ta',0):#06x} h~{dh}px {it.get('face','?'):<14} '{it['s'][:38]}'")
json.dump(items, open(OUT,'w',encoding='utf-8'), ensure_ascii=False, indent=1)
print(f"\n-> {OUT} ({len(items)} entrées)")

try: pres.Close()
except: pass
try: app.Quit()
except: pass
time.sleep(1)
subprocess.run(["taskkill","/F","/PID",str(pid)], capture_output=True)
print("done")
