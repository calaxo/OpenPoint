# Régénère Cairo-*.woff2 / SpaceGrotesk-*.woff2 depuis les polices VARIABLES Google Fonts (OFL).
# Usage : pip install fonttools brotli ; python build-google-fonts.py
#
# Cairo et Space Grotesk sont les polices de très nombreux modèles Slidesgo : le lecteur les embarque
# sous leur nom EXACT (FACES dans animFonts.js) pour retrouver les glyphes ET les coupures de ligne de
# PowerPoint. On part des variables Google, on fige l'instance voulue (instancer) puis on sous-ensemble
# comme build-web-fonts.py. « Space Grotesk Medium » est une famille distincte = instance wght 500.
#
# Les fichiers variables téléchargés sont mis en cache dans _vf_cache/ (non versionné). Les .woff2
# produits et les métriques sont écrits à côté de ce script et fusionnés dans metrics.json/advances.json
# (les entrées des autres polices sont préservées).
import json, pathlib, urllib.request
from fontTools.ttLib import TTFont
from fontTools.varLib import instancer
from fontTools import subset

HERE = pathlib.Path(__file__).parent
CACHE = HERE / "_vf_cache"

# Mêmes plages que build-web-fonts.py (latin étendu, grec, ponctuation, maths, formes, puces)
TEXT = ("U+0000-024F,U+02B0-036F,U+0370-03FF,U+1E00-1EFF,U+2000-214F,U+2150-22FF,"
        "U+2300-23FF,U+2460-24FF,U+25A0-27BF,U+FB00-FB06,U+FFFD")

SOURCES = {
    "Cairo-VF.ttf": "https://raw.githubusercontent.com/google/fonts/main/ofl/cairo/Cairo%5Bslnt,wght%5D.ttf",
    "SpaceGrotesk-VF.ttf": "https://raw.githubusercontent.com/google/fonts/main/ofl/spacegrotesk/SpaceGrotesk%5Bwght%5D.ttf",
}
# variable source -> (axe figé, fichier woff2 de sortie). Cairo a aussi un axe slnt (penché) : on fige 0.
JOBS = [
    ("Cairo-VF.ttf", {"wght": 400, "slnt": 0}, "Cairo-Regular.woff2"),
    ("Cairo-VF.ttf", {"wght": 700, "slnt": 0}, "Cairo-Bold.woff2"),
    ("SpaceGrotesk-VF.ttf", {"wght": 400}, "SpaceGrotesk-Regular.woff2"),
    ("SpaceGrotesk-VF.ttf", {"wght": 500}, "SpaceGrotesk-Medium.woff2"),
]

CACHE.mkdir(exist_ok=True)
for name, url in SOURCES.items():
    dst = CACHE / name
    if not dst.exists():
        print(f"Téléchargement {name} …")
        urllib.request.urlretrieve(url, dst)

metrics = json.loads((HERE / "metrics.json").read_text(encoding="utf-8"))
advances = json.loads((HERE / "advances.json").read_text(encoding="utf-8"))

for src, loc, outname in JOBS:
    f = instancer.instantiateVariableFont(TTFont(CACHE / src), loc)
    opts = subset.Options(); opts.flavor = "woff2"; opts.layout_features = ["*"]; opts.name_IDs = ["*"]
    opts.notdef_outline = True; opts.glyph_names = False
    s = subset.Subsetter(opts); s.populate(unicodes=subset.parse_unicodes(TEXT)); s.subset(f)
    dst = HERE / outname
    f.flavor = "woff2"; f.save(dst)
    o, upm = f["OS/2"], f["head"].unitsPerEm
    metrics[outname] = [round(o.usWinAscent / upm, 4), round(o.usWinDescent / upm, 4)]
    cmap, hmtx = f.getBestCmap(), f["hmtx"].metrics
    advances[outname] = {"upm": upm, "a": {str(cp): hmtx[g][0] for cp, g in sorted(cmap.items()) if g in hmtx}}
    print(f"{outname:30s} {dst.stat().st_size // 1024:5d} Ko  metrics={metrics[outname]}")

(HERE / "metrics.json").write_text(json.dumps(metrics, indent=1, sort_keys=True) + "\n", encoding="utf-8")
(HERE / "advances.json").write_text(json.dumps(advances, separators=(",", ":"), sort_keys=True), encoding="utf-8")
print(f"metrics.json : {len(metrics)} entrées | advances.json : {len(advances)} entrées")
