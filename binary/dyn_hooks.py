# Analyse DYNAMIQUE : pilote PowerPoint en COM, hooke PPCORE.DLL avec frida, et ajoute des
# effets d'animation en direct pour observer le compilateur preset->behaviors du vrai moteur.
#
# Ce qu'on hooke :
#   AppendBehavior  @ VA 0x1806408b0  (arg3 r8 = wchar* nom de propriete)  -> quel behavior est cree
#   + backtrace     -> quelle fonction-recette appelle, pour chaque preset
#
# Usage : python binary/dyn_hooks.py
import frida, comtypes.client, time, os, shutil, subprocess

VA_APPEND = 0x1806408b0
IMAGE_BASE = 0x180000000
SPECIMEN = r"specimen\sample.pptx"
TMP = os.path.join(os.environ.get("TEMP", r"C:\tmp"), "pp_dyn_probe.pptx")

# presets à tester : (nom, msoAnimEffect == presetID)
PRESETS = [("Appear", 1), ("Fly", 2), ("Fade", 10), ("Wipe", 22),
           ("Circle", 6), ("FadedZoom", 53), ("Ascend", 42), ("Zoom", 23)]

JS = r"""
const IMAGE_BASE = ptr('0x180000000');
const ppcore = Process.getModuleByName('PPCORE.DLL');
const VA_APPEND = ptr('%VA%');
const rt = ppcore.base.add(VA_APPEND.sub(IMAGE_BASE));

function toVA(addr) {  // adresse runtime -> VA statique (pour re-desassembler)
  if (addr.compare(ppcore.base) >= 0 && addr.compare(ppcore.base.add(ppcore.size)) < 0)
    return 'ppcore+0x' + addr.sub(ppcore.base).add(IMAGE_BASE).sub(IMAGE_BASE).toString(16)
           + ' (VA 0x' + addr.sub(ppcore.base).add(IMAGE_BASE).toString(16) + ')';
  const m = Process.findModuleByAddress(addr);
  return m ? m.name + '+0x' + addr.sub(m.base).toString(16) : addr.toString();
}

Interceptor.attach(rt, {
  onEnter(args) {
    let prop = '?';
    try { prop = args[2].readUtf16String(); } catch (e) {}
    const bt = Thread.backtrace(this.context, Backtracer.ACCURATE).slice(0, 3).map(toVA);
    send({t: 'beh', prop: prop, caller: bt[0]});
  }
});
send({t: 'ready', base: ppcore.base.toString()});
""".replace("%VA%", hex(VA_APPEND))

events = []
def on_message(msg, data):
    p = msg.get('payload')
    if not p:
        print("[frida]", msg); return
    if p.get('t') == 'beh':
        events.append(p)
    elif p.get('t') == 'ready':
        print("[frida] hooks armes, ppcore @", p['base'])

# --- 0. process PowerPoint existants (pour trouver le nouveau) ---
def ppids():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True).stdout
    return {int(l.split(',')[1].strip('"')) for l in out.splitlines() if 'POWERPNT' in l}

before = ppids()
shutil.copy(SPECIMEN, TMP)

# --- 1. lancer PowerPoint en COM ---
print("Demarrage PowerPoint (COM)...")
app = comtypes.client.CreateObject("PowerPoint.Application")
app.Visible = 1
try: app.DisplayAlerts = 1  # ppAlertsNone
except Exception: pass
pres = app.Presentations.Open(TMP, ReadOnly=0, WithWindow=1)
time.sleep(2)

new = ppids() - before
pid = next(iter(new)) if new else next(iter(ppids()))
print("PID PowerPoint =", pid)

# --- 2. frida attach + hooks ---
session = frida.get_local_device().attach(pid)
script = session.create_script(JS)
script.on('message', on_message)
script.load()
time.sleep(0.5)

# --- 3. scenario : ajouter chaque preset et observer les behaviors construits ---
slide = pres.Slides(1)
shp = slide.Shapes(1)
print(f"Cible : diapo 1, forme '{shp.Name}'\n")
results = {}
for name, eid in PRESETS:
    events.clear()
    try:
        eff = slide.TimeLine.MainSequence.AddEffect(shp, eid)  # EffectId = msoAnimEffect
        time.sleep(0.4)
        props = [e['prop'] for e in events]
        callers = sorted({e['caller'] for e in events})
        results[name] = (props, callers)
        print(f"[{name:<10} id={eid:<3}] behaviors: {props}")
        for c in callers: print(f"              recette: {c}")
        eff.Delete()
    except Exception as ex:
        print(f"[{name:<10} id={eid:<3}] ERREUR {ex}")
    time.sleep(0.2)

# --- 4. cleanup ---
try: pres.Close()
except Exception: pass
try: app.Quit()
except Exception: pass
time.sleep(1)
subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
try: os.remove(TMP)
except Exception: pass
print("\ndone")
