#!/usr/bin/env python3
"""Portuguese (pt-MZ) versions of the images whose words are drawn into the pixels.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_pt_images.py

The English tools draw legends and labels into the images, and a Portuguese page must not show
English inside its pictures. This tool reuses make_diagrams.py unchanged: it swaps each drawn string
for its Portuguese text (and measures the Portuguese text, so labels fit) and writes the results
next to the English masters with a -pt suffix:

  site/images/diagrams/corridor-pt.png, area-ring-pt.png

The sample pipeline's register views are drawn in both languages by make_samples.py itself (-pt files); the
Sentinel-2 route views carry no words. The English masters are not touched.
"""
import re
import sys
from pathlib import Path

from PIL import ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
import make_diagrams  # noqa: E402

PT = {
    # schematics (make_diagrams)
    "SCHEMATIC": "ESQUEMA",
    "Encroachment density per 500 m": "Densidade de ocupação por troço de 500 m",
    "Tailings": "Barragem de",
    "facility": "rejeitados",
    "Pit": "Cava",
    "New track since the last survey": "Nova picada desde o último levantamento",
    "distance to": "distância",
    "the boundary": "ao limite",
    "Your boundary": "O seu limite",
    "Ring around it": "Faixa envolvente",
    "Zone supplied by your engineers": "Zona definida pelos seus engenheiros",
}
RATINGS = {"High": "Alta", "Medium": "Média", "Low": "Baixa"}


def pt(s):
    if not isinstance(s, str):
        return s
    if s in PT:
        return PT[s]
    m = re.fullmatch(r"(High|Medium|Low) · (\d+)", s)
    if m:
        return f"{RATINGS[m.group(1)]} · {m.group(2)}"
    m = re.fullmatch(r"km (\d+)\.(\d)", s)                  # decimal comma
    if m:
        return f"km {m.group(1)},{m.group(2)}"
    return s


_text, _textlength = ImageDraw.ImageDraw.text, ImageDraw.ImageDraw.textlength
ImageDraw.ImageDraw.text = lambda self, xy, text, *a, **k: _text(self, xy, pt(text), *a, **k)
ImageDraw.ImageDraw.textlength = lambda self, text, *a, **k: _textlength(self, pt(text), *a, **k)

_save = make_diagrams.Canvas.save
make_diagrams.Canvas.save = lambda self, path: _save(self, path.with_name(path.stem + "-pt" + path.suffix))


if __name__ == "__main__":
    make_diagrams.corridor()
    make_diagrams.area_ring()
