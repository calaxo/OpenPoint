# Cartographie PE d'un module : sections, imports (par DLL), exports, quelques stats.
# Sert à cartographier la surface d'API de PPCORE.DLL (de quoi dépend-il, qu'expose-t-il).
# Usage : python binary/pe-info.py <dll> [--imports] [--exports] [--sections]

import sys, pefile
dll = sys.argv[1]
flags = set(sys.argv[2:])
show_all = not (flags & {'--imports', '--exports', '--sections'})

pe = pefile.PE(dll)
print(f"{dll}")
print(f"ImageBase 0x{pe.OPTIONAL_HEADER.ImageBase:x}  Machine 0x{pe.FILE_HEADER.Machine:x}  "
      f"Sections {pe.FILE_HEADER.NumberOfSections}")

if show_all or '--sections' in flags:
    print("\n== sections ==")
    for s in pe.sections:
        n = s.Name.rstrip(b'\x00').decode('latin1')
        print(f"  {n:<10} VA 0x{s.VirtualAddress:08x}  vsize 0x{s.Misc_VirtualSize:08x}  "
              f"raw 0x{s.SizeOfRawData:08x}")

if show_all or '--imports' in flags:
    print("\n== imports (par DLL, nb de fonctions) ==")
    if hasattr(pe, 'DIRECTORY_ENTRY_IMPORT'):
        rows = []
        for entry in pe.DIRECTORY_ENTRY_IMPORT:
            name = entry.dll.decode('latin1')
            rows.append((len(entry.imports), name))
        for cnt, name in sorted(rows, reverse=True):
            print(f"  {cnt:4}  {name}")

if show_all or '--exports' in flags:
    print("\n== exports ==")
    if hasattr(pe, 'DIRECTORY_ENTRY_EXPORT'):
        exps = pe.DIRECTORY_ENTRY_EXPORT.symbols
        print(f"  total : {len(exps)}")
        for e in exps[:40]:
            nm = e.name.decode('latin1') if e.name else f'#ord{e.ordinal}'
            print(f"    0x{e.address:08x}  {nm}")
    else:
        print("  (aucun export nommé)")
