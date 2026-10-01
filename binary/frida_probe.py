# Test d'injection frida dans PowerPoint. Spawne POWERPNT.EXE avec le deck, attend le
# chargement de PPCORE.DLL, énumère les modules, imprime la base de PPCORE, puis tue le process.
# But : valider qu'on PEUT instrumenter Office en runtime avant d'écrire les hooks sérieux.

import frida, sys, time, subprocess

EXE = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
PPTX = sys.argv[1] if len(sys.argv) > 1 else \
    r"specimen\sample.pptx"
WAIT = int(sys.argv[2]) if len(sys.argv) > 2 else 12

JS = r"""
rpc.exports = {
  mods: function () {
    return Process.enumerateModules()
      .filter(m => /ppcore|ocpptview|mso|POWERPNT/i.test(m.name))
      .map(m => ({name: m.name, base: m.base.toString(), size: m.size, path: m.path}));
  },
  hasppcore: function () {
    try { const m = Process.getModuleByName('PPCORE.DLL'); return m.base.toString(); }
    catch (e) { return null; }
  }
};
send({tag: 'ready'});
"""

def on_message(msg, data):
    print("[js]", msg.get('payload', msg))

device = frida.get_local_device()
print("Spawning PowerPoint...")
pid = device.spawn([EXE, PPTX])
print("pid", pid)
session = device.attach(pid)
script = session.create_script(JS)
script.on('message', on_message)
script.load()
device.resume(pid)

base = None
for i in range(WAIT):
    time.sleep(1)
    try:
        base = script.exports_sync.hasppcore()
    except Exception as e:
        base = None
    if base:
        print(f"PPCORE.DLL base @ {base} (apres {i+1}s)")
        break

if base:
    for m in script.exports_sync.mods():
        print(f"  {m['name']:<16} base {m['base']} size 0x{m['size']:x}")
else:
    print("PPCORE pas encore charge / injection echouee")

try:
    session.detach()
except Exception:
    pass
subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print("done, process tue")
