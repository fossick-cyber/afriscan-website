#!/usr/bin/env python3
"""Portuguese (pt-MZ) versions of the images whose words are drawn into the pixels.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_pt_images.py

The English tools draw legends, credits and labels into the images, and a Portuguese page must not
show English inside its pictures. This tool reuses make_samples.py and make_diagrams.py unchanged:
it swaps each drawn string for its Portuguese text (and measures the Portuguese text, so legend
boxes fit) and writes the results next to the English masters with a -pt suffix:

  site/images/samples/t9-km5-6-pt.jpg, t9-rating-{high,medium,low}-pt.jpg
  site/images/diagrams/corridor-pt.png, area-ring-pt.png

The English masters are not touched. Read-only against /opt/favhousecheck, like make_samples.py.
"""
import re
import sys
from pathlib import Path

from PIL import ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent))
import make_diagrams  # noqa: E402
import make_samples  # noqa: E402

PT = {
    # T-9 sample legend and credit (make_samples.draw_view)
    "Route (as supplied)": "Traçado (tal como fornecido)",
    "50 m buffer": "Faixa de 50 m",
    "100 m buffer": "Faixa de 100 m",
    "Reviewer mark within 50 m": "Marcação do revisor até 50 m",
    "Reviewer mark 50–100 m": "Marcação do revisor a 50–100 m",
    "Reviewer mark beyond 100 m": "Marcação do revisor além de 100 m",
    "Imagery © Google · marks and buffers: AfriScan": "Imagens © Google · marcações e faixas: AfriScan",
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


def samples():
    line = make_samples.load_route()
    marks = make_samples.load_marks(line)
    make_samples.draw_view(["0080", "0081", "0103", "0104"], line, marks, "t9-km5-6-pt.jpg",
                           title="Traçado T-9 · km 5,0–6,3")
    for name, chunks, chain in (("t9-rating-high-pt.jpg", ["0080", "0081", "0103", "0104"], 5750),
                                ("t9-rating-medium-pt.jpg", ["0017", "0018", "0019", "0040", "0041", "0042"], 8990),
                                ("t9-rating-low-pt.jpg", ["0015", "0016", "0038", "0039"], 8000)):
        make_samples.draw_view(chunks, line, marks, name, labels=False, legend=False, center=chain)
    print("wrote site/images/samples/t9-*-pt.jpg")


if __name__ == "__main__":
    make_diagrams.corridor()
    make_diagrams.area_ring()
    samples()
