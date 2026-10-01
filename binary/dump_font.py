# Dumpe la police RÉELLEMENT utilisée par PowerPoint pour le texte "Poppins" d'une diapo.
# Hooke gdi32!ExtTextOutW pendant le rendu ; sur le 1er texte dont la police demandée = Poppins,
# lit la face RÉELLE (GetTextFaceW, après mapping GDI) et dumpe les octets du fichier de police
# (GetFontData) -> dit si PowerPoint a la vraie Poppins ou substitue, et récupère le fichier exact.
import frida, comtypes.client, time, os, sys, subprocess, base64

DECK = r"specimen/sample.pptx"
OUT = r"pp_poppins.ttf"

JS = r"""
const g = Process.getModuleByName('gdi32.dll');
const GetCurrentObject = new NativeFunction(g.getExportByName('GetCurrentObject'),'pointer',['pointer','uint']);
const GetObjectW       = new NativeFunction(g.getExportByName('GetObjectW'),'int',['pointer','int','pointer']);
const GetTextFaceW     = new NativeFunction(g.getExportByName('GetTextFaceW'),'int',['pointer','int','pointer']);
const GetFontData      = new NativeFunction(g.getExportByName('GetFontData'),'uint',['pointer','uint','uint','pointer','uint']);
const GDI_ERROR = 0xFFFFFFFF;
let done = false;

function hook(name){
  const p = g.findExportByName(name); if (!p) return;
  Interceptor.attach(p, { onEnter(a){
    if (done) return;
    const hdc = a[0];
    // police DEMANDÉE (LOGFONTW.lfFaceName @+28)
    let reqFace = '';
    try { const buf = Memory.alloc(92); const hf = GetCurrentObject(hdc, 6); if(!hf.isNull()){ GetObjectW(hf,92,buf); reqFace = buf.add(28).readUtf16String(); } } catch(e){}
    if (!/poppins/i.test(reqFace || '')) return;
    // police RÉELLE mappée par GDI (GetTextFaceW)
    let realFace = '';
    try { const fb = Memory.alloc(128); const n = GetTextFaceW(hdc, 64, fb); if(n>0) realFace = fb.readUtf16String(); } catch(e){}
    // dump du fichier de police (GetFontData, table entière dwTable=0)
    try {
      const size = GetFontData(hdc, 0, 0, ptr(0), 0);
      if (size !== GDI_ERROR && size > 0 && size < 20*1024*1024) {
        const buf = Memory.alloc(size);
        const got = GetFontData(hdc, 0, 0, buf, size);
        if (got !== GDI_ERROR && got > 0) {
          done = true;
          send({t:'font', reqFace, realFace, size: got}, buf.readByteArray(got));
          return;
        }
      }
      send({t:'info', reqFace, realFace, size});
    } catch(e){ send({t:'err', m:''+e}); }
  }});
}
hook('ExtTextOutW'); hook('TextOutW');
send({t:'ready'});
"""

got = {}
def on_message(msg, data):
    p = msg.get('payload')
    if not p: print("[frida]", msg); return
    if p.get('t') == 'font':
        print(f"[frida] police demandée='{p['reqFace']}' | RÉELLE mappée='{p['realFace']}' | taille={p['size']} o")
        if data: open(OUT, 'wb').write(data); print(f"-> dump écrit : {OUT} ({len(data)} o)")
        got.update(p)
    elif p.get('t') == 'info':
        print(f"[frida] demandée='{p['reqFace']}' réelle='{p['realFace']}' taille GetFontData={p['size']} (pas de dump)")
    elif p.get('t') == 'ready': print("[frida] hook armé")
    else: print("[frida]", p)

print("PowerPoint COM...")
app = comtypes.client.CreateObject("PowerPoint.Application"); app.Visible = 1
pres = app.Presentations.Open(DECK, ReadOnly=1, WithWindow=1); time.sleep(2)
pid = int(subprocess.run(["tasklist","/FI","IMAGENAME eq POWERPNT.EXE","/FO","CSV","/NH"], capture_output=True, text=True).stdout.split(',')[1].strip('"'))
print("PID", pid)
session = frida.get_local_device().attach(pid)
sc = session.create_script(JS); sc.on('message', on_message); sc.load(); time.sleep(0.5)
# forcer le rendu de plusieurs diapos (footer "Poppins" y est) pour déclencher le hook
for n in (1, 5, 12):
    try: pres.Slides(n).Export(OUT + f".s{n}.png", "PNG", 640, 360)
    except Exception as ex: print("export", n, "err", ex)
    time.sleep(0.8)
time.sleep(1)
try: pres.Close()
except: pass
try: app.Quit()
except: pass
time.sleep(1); subprocess.run(["taskkill","/F","/PID",str(pid)], capture_output=True)
print("done", "| dumped" if got else "| RIEN dumpé (police pas passée par GDI ExtTextOut ?)")
