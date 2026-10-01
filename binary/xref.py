# Cross-reference + désassemblage ciblé de PPCORE.DLL.
# Trouve, pour chaque token (chaîne du vocabulaire du moteur), les sites de CODE machine
# qui le référencent, puis désassemble une fenêtre autour — c.-à-d. le vrai code de
# PowerPoint qui manipule ces concepts.
#
# Deux types de références détectés :
#   1. lea reg,[rip+disp]  vers la chaîne  (usage direct)
#   2. pointeur qword == VA de la chaîne    (entrée de table serialiseur), + lea vers cette table
#
# Usage : python binary/xref.py <dll> [token ...]

import sys, pefile
from capstone import Cs, CS_ARCH_X86, CS_MODE_64, CS_OP_MEM
from capstone.x86 import X86_REG_RIP

dll = sys.argv[1]
tokens = sys.argv[2:] or ['presetClass','clickEffect','withEffect','afterEffect','mainSeq',
                          'tmRoot','presetID','nodeType','ppt_x','ppt_y','style.visibility']

pe = pefile.PE(dll, fast_load=True)
image_base = pe.OPTIONAL_HEADER.ImageBase
data = pe.__data__  # bytes du fichier entier

# sections
sect = []
for s in pe.sections:
    name = s.Name.rstrip(b'\x00').decode('latin1')
    sect.append((name, s.VirtualAddress, s.Misc_VirtualSize,
                 s.PointerToRawData, s.SizeOfRawData))

def off_to_rva(off):
    for name, va, vs, praw, sraw in sect:
        if praw <= off < praw + sraw:
            return va + (off - praw)
    return None

def rva_to_off(rva):
    for name, va, vs, praw, sraw in sect:
        if va <= rva < va + max(vs, sraw):
            o = praw + (rva - va)
            if o < praw + sraw:
                return o
    return None

text = next(s for s in sect if s[0] == '.text')
text_off, text_size, text_rva = text[3], text[4], text[1]
text_bytes = data[text_off:text_off + text_size]
text_va = image_base + text_rva

def find_utf16(tok):
    b = tok.encode('utf-16le')
    offs, i = [], 0
    while True:
        i = data.find(b, i)
        if i < 0: break
        offs.append(i); i += len(b)
    return offs

# VA de chaque token
tok_va = {}
for t in tokens:
    for off in find_utf16(t):
        rva = off_to_rva(off)
        if rva is not None:
            tok_va.setdefault(t, []).append(image_base + rva)

all_va = {va: t for t, vas in tok_va.items() for va in vas}

# --- 1. scan direct lea rip-relative dans .text ---
LEA_MODRM = {0x05:'rax',0x0d:'rcx',0x15:'rdx',0x1d:'rbx',0x25:'rsp',0x2d:'rbp',0x35:'rsi',0x3d:'rdi'}
lea_xrefs = {}  # token -> [code_va]
for i in range(len(text_bytes) - 7):
    b0 = text_bytes[i]
    if b0 not in (0x48, 0x4c): continue
    if text_bytes[i+1] != 0x8d: continue
    modrm = text_bytes[i+2]
    if modrm not in LEA_MODRM: continue
    disp = int.from_bytes(text_bytes[i+3:i+7], 'little', signed=True)
    insn_end_va = text_va + i + 7
    target = insn_end_va + disp
    if target in all_va:
        lea_xrefs.setdefault(all_va[target], []).append(text_va + i)

# --- 2. pointeurs qword vers les chaînes (tables du sérialiseur) ---
ptr_hits = {}  # token -> [va_of_pointer]
for t, vas in tok_va.items():
    for va in vas:
        pat = va.to_bytes(8, 'little')
        i = 0
        while True:
            i = data.find(pat, i)
            if i < 0: break
            prva = off_to_rva(i)
            if prva is not None:
                ptr_hits.setdefault(t, []).append(image_base + prva)
            i += 8

print(f"ImageBase 0x{image_base:x}  .text VA 0x{text_va:x} taille 0x{text_size:x}\n")
print("token".ljust(18), "occ", "lea-xref", "ptr-table")
for t in tokens:
    print(t.ljust(18),
          str(len(tok_va.get(t, []))).rjust(3),
          str(len(lea_xrefs.get(t, []))).rjust(8),
          str(len(ptr_hits.get(t, []))).rjust(9),
          " ".join(hex(x) for x in lea_xrefs.get(t, [])[:3]))

# --- désassemblage autour des premiers lea-xref ---
md = Cs(CS_ARCH_X86, CS_MODE_64)
md.detail = True

def annotate(insn):
    if insn.id == 0: return ""
    for op in insn.operands:
        if op.type == CS_OP_MEM and op.mem.base == X86_REG_RIP:
            tgt = insn.address + insn.size + op.mem.disp
            if tgt in all_va:
                return f"   ; -> \"{all_va[tgt]}\""
    return ""

def disasm_window(code_va, before=0x30, after=0x30):
    start = code_va - before
    off = rva_to_off(start - image_base)
    chunk = data[off:off + before + after]
    print(f"\n--- disasm @ 0x{code_va:x} ---")
    for insn in md.disasm(chunk, start):
        mark = ">>" if insn.address == code_va else "  "
        print(f"{mark} 0x{insn.address:x}: {insn.mnemonic:<7} {insn.op_str}{annotate(insn)}")

# montre le contexte pour les tokens de dispatch de séquence les plus parlants
for t in ['clickEffect','withEffect','afterEffect','mainSeq','presetClass']:
    for va in lea_xrefs.get(t, [])[:2]:
        disasm_window(va)
