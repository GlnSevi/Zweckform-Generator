"""Build zweckform_formats.json from the LibreOffice and gLabels label catalogues.

    python tools/build_formats.py

Sources (tools/sources/):
  libreoffice-labels.xml          LibreOffice core, extras/source/labels/labels.xml
  glabels-avery-iso-templates.xml gLabels-qt, templates/avery-iso-templates.xml

Only A4 sheets (portrait or landscape) are taken; continuous/endless labels and
Letter/A5 sheets are skipped because the app prints A4 only. The hand-checked
formats in free_labels.CURATED win over catalogue values with the same code.
"""

import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from free_labels import CURATED, LabelFormat  # noqa: E402

SOURCES = ROOT / "tools" / "sources"
# German Zweckform names first: they win when a code occurs in several lists.
LO_MANUFACTURERS = ["Avery Zweckform", "Avery A4", "Avery A4/Asia"]
ROUND_WORDS = re.compile(r"\b(rund|round|kreis|circle|cd|dvd)\b|Ø", re.I)


def page(width, height):
    if (round(width), round(height)) == (210, 297):
        return False
    if (round(width), round(height)) == (297, 210):
        return True
    return None


def mm(value):
    """Parse gLabels lengths such as '63.5mm', '2.5in' or '72pt'."""
    match = re.fullmatch(r"([\d.]+)\s*(mm|cm|in|pt)?", value.strip())
    number, unit = float(match[1]), match[2] or "pt"
    return number * {"mm": 1, "cm": 10, "in": 25.4, "pt": 25.4 / 72}[unit]


def split_name(text):
    text = " ".join(text.split())
    code, _, name = text.partition(" ")
    return code, name


def libreoffice():
    root = ET.parse(SOURCES / "libreoffice-labels.xml").getroot()
    for manufacturer in LO_MANUFACTURERS:
        node = next(m for m in root.iter("manufacturer") if m.get("name") == manufacturer)
        for label in node.iter("label"):
            kind, *values = label.find("measure").text.split(";")
            if kind != "S":
                continue
            # 1/100 mm: pitch x/y, width, height, left, top, cols, rows, paper w/h
            pitch_x, pitch_y, width, height, left, top, _, _, pw, ph = (
                int(v) / 100 for v in values)
            cols, rows = int(values[6]), int(values[7])
            landscape = page(pw, ph)
            if landscape is None:
                continue
            code, name = split_name(label.find("name").text)
            yield dict(code=code, name=name, width=width, height=height,
                       columns=cols, rows=rows, left=left, top=top,
                       gap_x=round(pitch_x - width, 3) if cols > 1 else 0,
                       gap_y=round(pitch_y - height, 3) if rows > 1 else 0,
                       landscape=landscape,
                       shape="round" if ROUND_WORDS.search(name) and abs(width - height) < .5 else "rect",
                       source=f"LibreOffice ({manufacturer})")


def glabels():
    root = ET.parse(SOURCES / "glabels-avery-iso-templates.xml").getroot()
    templates = {t.get("part"): t for t in root.iter("Template")}
    for part, template in templates.items():
        base = templates.get(template.get("equiv"), template)
        if base.get("size") != "A4":
            continue
        shapes = [c for c in base if c.tag.startswith("Label-")]
        layouts = shapes[0].findall("Layout") if len(shapes) == 1 else []
        if len(layouts) != 1:
            continue  # mixed sheets (several label sizes) do not fit one grid
        shape, layout = shapes[0], layouts[0]
        if shape.tag == "Label-rectangle":
            width, height, kind = mm(shape.get("width")), mm(shape.get("height")), "rect"
        elif shape.tag in ("Label-round", "Label-cd"):
            width = height = 2 * mm(shape.get("radius")); kind = "round"
        elif shape.tag == "Label-ellipse":
            width, height, kind = mm(shape.get("width")), mm(shape.get("height")), "round"
        else:
            continue
        cols, rows = int(layout.get("nx")), int(layout.get("ny"))
        yield dict(code=part, name=base.get("_description") or "", width=width, height=height,
                   columns=cols, rows=rows, left=mm(layout.get("x0")), top=mm(layout.get("y0")),
                   gap_x=round(mm(layout.get("dx")) - width, 3) if cols > 1 else 0,
                   gap_y=round(mm(layout.get("dy")) - height, 3) if rows > 1 else 0,
                   landscape=False, shape=kind, source="gLabels (Avery A4)")


def geometry(entry):
    return tuple(round(entry[k], 1) for k in ("width", "height", "left", "top", "gap_x", "gap_y")) + (
        entry["columns"], entry["rows"], entry["landscape"])


def main():
    formats, seen_codes, code_source, skipped = [], {}, {}, []
    curated = {f.code for f in CURATED}
    for fmt in CURATED:
        seen_codes[fmt.code] = [geometry(fmt.as_dict())]
    for entry in [*libreoffice(), *glabels()]:
        for key in ("width", "height", "left", "top", "gap_x", "gap_y"):
            entry[key] = round(entry[key], 2)
        try:
            LabelFormat(**{k: v for k, v in entry.items() if k != "source"}).validate()
        except ValueError as exc:
            skipped.append(f"{entry['code']} {entry['name']}: {exc}")
            continue
        geo = geometry(entry)
        known = seen_codes.setdefault(entry["code"], [])
        source = code_source.setdefault(entry["code"], entry["source"])
        if entry["code"] in curated or geo in known or source != entry["source"]:
            continue  # the first (German) source wins for an article number
        if known:
            # Same article number, different sheet (e.g. spine + face): keep both.
            entry["code"] = f"{entry['code']}-{len(known) + 1}"
        known.append(geo)
        formats.append(entry)
    formats.sort(key=lambda e: [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", e["code"])])
    out = ROOT / "zweckform_formats.json"
    out.write_text(json.dumps(formats, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(CURATED)} geprüfte + {len(formats)} Katalog-Formate -> {out.name}")
    print(f"{len(skipped)} übersprungen (passen nicht auf A4):")
    for line in skipped:
        print("  ", line)


if __name__ == "__main__":
    main()
