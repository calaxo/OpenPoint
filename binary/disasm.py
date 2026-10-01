# Désassembleur ciblé + lecteur de tables de pointeurs pour PPCORE.DLL.
# Usage :
#   python binary/disasm.py <dll> code 0x1802d6b7c 40         -> 40 instructions depuis la VA
#   python binary/disasm.py <dll> ptrtab 0x1802d... 16        -> 16 qwords (table de pointeurs) résolus en chaînes

import sys, pefile
from capstone import Cs, CS_ARCH_X86, CS_MODE_64, CS_OP_MEM
from capstone.x86 import X86_REG_RIP

dll, mode, va_s = sys.argv[1], sys.argv[2], sys.argv[3]
count = int(sys.argv[4]) if len(sys.argv) > 4 else 32
va = int(va_s, 16)

pe = pefile.PE(dll, fast_load=True)
image_base = pe.OPTIONAL_HEADER.ImageBase
data = pe.__data__
sect = [(s.Name.rstrip(b'\x00').decode('latin1'), s.VirtualAddress, s.Misc_VirtualSize,
         s.PointerToRawData, s.SizeOfRawData) for s in pe.sections]

def rva_to_off(rva):
    for _, v, vs, praw, sraw in sect:
        if v <= rva < v + max(vs, sraw):
            o = praw + (rva - v)
            if o < praw + sraw: return o
    return None
def va_to_off(x): return rva_to_off(x - image_base)
def off_to_rva(off):
    for _, v, vs, praw, sraw in sect:
        if praw <= off < praw + sraw: return v + (off - praw)
    return None

# lit une chaîne (utf16 ou ascii) à une VA
def read_str(x):
    off = va_to_off(x)
    if off is None: return None
    # utf16 ?
    s = ''
    j = off
    while j < len(data) - 1 and data[j+1] == 0 and 0x20 <= data[j] < 0x7f:
        s += chr(data[j]); j += 2
    if len(s) >= 2: return ('w', s)
    s = ''; j = off
    while j < len(data) and 0x20 <= data[j] < 0x7f:
        s += chr(data[j]); j += 1
    if len(s) >= 3: return ('a', s)
    return None

# vocabulaire large pour annoter les lea rip-relatifs
VOCAB = ['presetClass','presetSubtype','nodeType','presetID','afterEffect','clickEffect',
         'withEffect','mainSeq','tmRoot','interactiveSeq','clickPar','withGroup','afterGroup',
         'indefinite','ppt_x','ppt_y','ppt_w','ppt_h','ppt_r','ppt_c','xshear','style.visibility',
         'style.rotation','style.opacity','fillcolor','entr','emph','exit','path','mediacall','verb',
         'animEffect','set','anim','animMotion','animRot','animScale','animClr','cmd','cBhvr']
tok_va = {}
for t in VOCAB:
    b = t.encode('utf-16le'); i = 0
    while True:
        i = data.find(b, i)
        if i < 0: break
        # frontière : précédé et suivi d'un non-alnum (évite ppt_x dans ppt_xxx)
        rva = off_to_rva(i)
        if rva is not None: tok_va[image_base + rva] = t
        i += len(b)

if mode == 'ptrtab':
    off = va_to_off(va)
    print(f"table @ 0x{va:x}")
    for k in range(count):
        q = int.from_bytes(data[off + k*8: off + k*8 + 8], 'little')
        s = read_str(q) if q else None
        extra = f'  "{s[1]}"' if s else ''
        print(f"  +0x{k*8:02x}  0x{q:016x}{extra}")
    sys.exit()

# mode code
md = Cs(CS_ARCH_X86, CS_MODE_64); md.detail = True
off = va_to_off(va)
chunk = data[off: off + count * 15]
n = 0
for insn in md.disasm(chunk, va):
    ann = ''
    for op in insn.operands:
        if op.type == CS_OP_MEM and op.mem.base == X86_REG_RIP:
            tgt = insn.address + insn.size + op.mem.disp
            if tgt in tok_va: ann = f'   ; -> "{tok_va[tgt]}"'
            else:
                s = read_str(tgt)
                if s: ann = f'   ; -> "{s[1]}"'
    print(f"0x{insn.address:x}: {insn.mnemonic:<8} {insn.op_str}{ann}")
    n += 1
    if n >= count: break
