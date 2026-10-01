# Profileur différentiel : échantillonne le PC des threads "running" de PowerPoint pendant qu'un
# diaporama JOUE des animations, vs à l'arrêt. Les buckets qui n'apparaissent qu'en animation
# = le code du chemin d'exécution (tick, éval de formules, application des valeurs, rendu).
#
# Usage : python binary/dyn_profile.py
import frida, comtypes.client, time, os, shutil, subprocess
from collections import Counter

SPECIMEN = r"specimen\sample.pptx"
TMP = os.path.join(os.environ.get("TEMP", r"C:\tmp"), "pp_prof.pptx")
ANIM_SLIDES = [5, 6, 7, 15, 16, 20]   # diapos riches en animations

JS = r"""
let sampling = false;
let hist = {};
let modbounds = [];
function refreshMods() {
  modbounds = Process.enumerateModules()
    .filter(m => /ppcore|ocpptview|mso|d2d1|dcomp|dwrite|d3d|POWERPNT/i.test(m.name))
    .map(m => ({name: m.name, lo: m.base, hi: m.base.add(m.size), base: m.base}));
}
function bucket(pc) {
  for (const m of modbounds) {
    if (pc.compare(m.lo) >= 0 && pc.compare(m.hi) < 0) {
      const off = pc.sub(m.base);
      // arrondi 0x40 ; pour ppcore on reporte la VA statique
      const rounded = off.and(ptr('0xffffffffffffffc0'));
      if (/ppcore/i.test(m.name))
        return 'ppcore VA 0x' + rounded.add(ptr('0x180000000')).toString(16);
      return m.name + '+0x' + rounded.toString(16);
    }
  }
  return null;
}
function tick() {
  if (!sampling) return;
  const ths = Process.enumerateThreads();
  for (const t of ths) {
    if (t.state !== 'running') continue;
    try {
      const k = bucket(t.context.pc);
      if (k) hist[k] = (hist[k] || 0) + 1;
    } catch (e) {}
  }
  setTimeout(tick, 1);
}
rpc.exports = {
  start: function () { refreshMods(); hist = {}; sampling = true; tick(); },
  stop:  function () { sampling = false; return hist; }
};
"""

def merge(dst, h):
    for k, v in h.items():
        dst[k] += v

before = {int(l.split(',')[1].strip('"')) for l in subprocess.run(
    ["tasklist","/FI","IMAGENAME eq POWERPNT.EXE","/FO","CSV","/NH"],
    capture_output=True,text=True).stdout.splitlines() if 'POWERPNT' in l}
shutil.copy(SPECIMEN, TMP)

print("Demarrage PowerPoint...")
app = comtypes.client.CreateObject("PowerPoint.Application")
app.Visible = 1
pres = app.Presentations.Open(TMP, ReadOnly=0, WithWindow=1)
time.sleep(2)
after = {int(l.split(',')[1].strip('"')) for l in subprocess.run(
    ["tasklist","/FI","IMAGENAME eq POWERPNT.EXE","/FO","CSV","/NH"],
    capture_output=True,text=True).stdout.splitlines() if 'POWERPNT' in l}
pid = next(iter(after - before)) if (after - before) else next(iter(after))
print("PID", pid)

session = frida.get_local_device().attach(pid)
script = session.create_script(JS)
script.on('message', lambda m, d: print("[frida]", m.get('payload', m)))
script.load()

# --- baseline : diaporama a l'arret (idle) ---
print("Baseline idle (3s)...")
idle = Counter()
script.exports_sync.start()
time.sleep(3)
merge(idle, script.exports_sync.stop())

# --- animation : lancer le diaporama et avancer dans les diapos animees ---
print("Diaporama + animations...")
anim = Counter()
try:
    pres.SlideShowSettings.Run()
    time.sleep(1.5)
    view = pres.SlideShowWindow.View
    script.exports_sync.start()
    for s in ANIM_SLIDES:
        try:
            view.GotoSlide(s)
            time.sleep(0.4)
            for _ in range(8):      # déclenche les étapes d'animation de la diapo
                view.Next()
                time.sleep(0.35)
        except Exception as e:
            print("  slide", s, "err", e)
    merge(anim, script.exports_sync.stop())
    try: view.Exit()
    except Exception: pass
except Exception as e:
    print("slideshow err", e)

# --- diff ---
def norm(c):
    tot = sum(c.values()) or 1
    return {k: v / tot for k, v in c.items()}
ni, na = norm(idle), norm(anim)
diff = []
for k in set(na) | set(ni):
    d = na.get(k, 0) - ni.get(k, 0)
    diff.append((d, na.get(k, 0), ni.get(k, 0), k))
diff.sort(reverse=True)

print(f"\n=== échantillons : idle={sum(idle.values())} anim={sum(anim.values())} ===")
print("\n-- buckets SUR-représentés en animation (part_anim - part_idle) --")
for d, a, i, k in diff[:25]:
    if d <= 0: break
    print(f"  +{d*100:5.2f}%  anim {a*100:5.2f}%  idle {i*100:5.2f}%  {k}")

print("\n-- répartition par module en animation --")
modc = Counter()
for k, v in anim.items():
    mod = k.split(' ')[0] if k.startswith('ppcore') else k.split('+')[0]
    modc[mod] += v
tot = sum(modc.values()) or 1
for mod, v in modc.most_common(12):
    print(f"  {v*100/tot:5.1f}%  {mod}")

# cleanup
try: pres.Close()
except Exception: pass
try: app.Quit()
except Exception: pass
time.sleep(1)
subprocess.run(["taskkill","/F","/PID",str(pid)], capture_output=True)
try: os.remove(TMP)
except Exception: pass
print("\ndone")
