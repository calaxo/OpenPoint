# Régénère les .woff2 de BACK/fonts à partir des polices sources (TTF/OTF libres).
# Usage : pip install fonttools brotli ; python build-web-fonts.py <dossier_sources>
# Le sous-ensemble garde ce qu'un support de cours FR/EN peut contenir (latin étendu, grec,
# ponctuation, flèches, opérateurs math, formes géométriques, puces) : sans lui, le CSS des
# polices (inliné dans CHAQUE diapo) pèserait plusieurs dizaines de Mo.
import sys, pathlib
from fontTools import subset
from fontTools.ttLib import TTFont

TEXT = ("U+0000-024F,U+02B0-036F,U+0370-03FF,U+1E00-1EFF,U+2000-214F,U+2150-22FF,"
        "U+2300-23FF,U+2460-24FF,U+25A0-27BF,U+FB00-FB06,U+FFFD")
SYMBOLS = "U+0020,U+00A0,U+00B0-00B7,U+00D7,U+00F7,U+0370-03FF,U+2000-206F,U+2190-23FF,U+25A0-27BF"

# fichier source -> (fichier woff2, plages)
FILES = {
  "Poppins-Regular.ttf": TEXT, "Poppins-Medium.ttf": TEXT, "Poppins-SemiBold.ttf": TEXT,
  "Poppins-Bold.ttf": TEXT, "Poppins-Italic.ttf": TEXT, "Poppins-BoldItalic.ttf": TEXT,
  "LiberationSans-Regular.ttf": TEXT, "LiberationSans-Bold.ttf": TEXT,
  "LiberationSans-Italic.ttf": TEXT, "LiberationSans-BoldItalic.ttf": TEXT,
  "LiberationSerif-Regular.ttf": TEXT, "LiberationSerif-Bold.ttf": TEXT,
  "LiberationSerif-Italic.ttf": TEXT, "LiberationSerif-BoldItalic.ttf": TEXT,
  "LiberationMono-Regular.ttf": TEXT, "LiberationMono-Bold.ttf": TEXT,
  "Carlito-Regular.ttf": TEXT, "Carlito-Bold.ttf": TEXT, "Carlito-Italic.ttf": TEXT, "Carlito-BoldItalic.ttf": TEXT,
  "Caladea-Regular.ttf": TEXT, "Caladea-Bold.ttf": TEXT,
  "Gelasio[wght].ttf": TEXT, "Gelasio-Italic[wght].ttf": TEXT,
  "ComicRelief-Regular.ttf": TEXT, "ComicRelief-Bold.ttf": TEXT,
  "ArchivoBlack-Regular.ttf": TEXT, "Anton-Regular.ttf": TEXT,
  "LibreFranklin[wght].ttf": TEXT, "LibreFranklin-Italic[wght].ttf": TEXT,
  "NewsCycle-Regular.ttf": TEXT, "NewsCycle-Bold.ttf": TEXT,
  "Cabin[wdth,wght].ttf": TEXT, "Cabin-Italic[wdth,wght].ttf": TEXT,
  "SortsMillGoudy-Regular.ttf": TEXT, "SortsMillGoudy-Italic.ttf": TEXT,
  "SourceSans3[wght].ttf": TEXT, "SourceSans3-Italic[wght].ttf": TEXT,
  "texgyreschola-regular.otf": TEXT, "texgyreschola-bold.otf": TEXT,
  "texgyreschola-italic.otf": TEXT, "texgyreschola-bolditalic.otf": TEXT,
  "texgyrepagella-regular.otf": TEXT, "texgyrepagella-bold.otf": TEXT,
  "texgyrepagella-italic.otf": TEXT, "texgyrepagella-bolditalic.otf": TEXT,
  "DejaVuSans.ttf": SYMBOLS, "DejaVuSans-Bold.ttf": SYMBOLS,
}

src, out = pathlib.Path(sys.argv[1]), pathlib.Path(__file__).parent
metrics, advances = {}, {}
for name, ranges in FILES.items():
  opts = subset.Options(); opts.flavor = "woff2"; opts.layout_features = ["*"]; opts.name_IDs = ["*"]
  opts.notdef_outline = True; opts.glyph_names = False
  font = TTFont(src / name)
  s = subset.Subsetter(opts); s.populate(unicodes=subset.parse_unicodes(ranges)); s.subset(font)
  dst = out / (pathlib.Path(name).stem.replace("[", "-").replace("]", "").replace(",", "-") + ".woff2")
  font.flavor = "woff2"; font.save(dst)
  # métriques « win » (celles de PowerPoint/GDI), en fraction d'em : animFonts.js les impose au
  # navigateur (ascent/descent-override), qui sinon cale la ligne sur hhea/typo -> texte décalé.
  o, upm = font["OS/2"], font["head"].unitsPerEm
  metrics[dst.name] = [round(o.usWinAscent / upm, 4), round(o.usWinDescent / upm, 4)]
  # chasses par caractère (unités de la police) : PowerPoint arrondit chaque chasse au 1/8 pt ; le moteur
  # en déduit une correction d'espacement par run (pptx-render.mjs, advanceFix) -> mêmes coupures de ligne.
  cmap, hmtx = font.getBestCmap(), font["hmtx"].metrics
  advances[dst.name] = {"upm": upm, "a": {str(cp): hmtx[g][0] for cp, g in sorted(cmap.items()) if g in hmtx}}
  print(f"{dst.name:40s} {dst.stat().st_size // 1024:5d} Ko")
import json
(out / "metrics.json").write_text(json.dumps(metrics, indent=1, sort_keys=True) + "\n", encoding="utf-8")
(out / "advances.json").write_text(json.dumps(advances, separators=(",", ":"), sort_keys=True), encoding="utf-8")
