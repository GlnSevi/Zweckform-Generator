# -*- coding: utf-8 -*-
"""
Lager-Etiketten Generator – Avery Zweckform L4761-100
A4-Bogen, 4 Etiketten je 192 x 61 mm.

GUI (tkinter) mit Vorschau und Direktdruck über Windows GDI (pywin32),
damit die Positionen millimetergenau auf dem Etikettenbogen landen.

Zweiter Reiter „Regalschilder (3D-Druck)“: große Schilder (Standard
300 × 300 mm) zur Regal-Markierung als STL-Datei für den 3D-Drucker.
Die Zahl wird als Vertiefung in die Platte eingelassen, damit sie sich
nach dem Druck sauber ausmalen bzw. per Filamentwechsel einfärben lässt.
"""

import ctypes
import json
import math
import os
import re
import struct
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

try:
    import win32con
    import win32gui
    import win32print
    import win32ui
except ImportError:
    win32con = win32gui = win32print = win32ui = None

try:
    import manifold3d as m3d
    import numpy as np
except ImportError:
    m3d = np = None

APP_ID = "Kunzer.LagerLabel.Etiketten"
APP_TITLE = "Lager-Etiketten & Regalschilder – Kunzer"
SETTINGS_FILE = os.path.join(os.environ.get("APPDATA", "."), "LagerLabel", "settings.json")

# ---------------------------------------------------------------- Bogen-Maße (mm)
PAGE_W, PAGE_H = 210.0, 297.0
LABEL_W, LABEL_H = 192.0, 61.0
MARGIN_LEFT = (PAGE_W - LABEL_W) / 2          # 9 mm
MARGIN_TOP = (PAGE_H - 4 * LABEL_H) / 2       # 26.5 mm
PAD_X = 10.0                                  # Innenabstand links/rechts im Etikett
PER_SHEET = 4

BC_MODULE = 0.33        # Barcode-Modulbreite mm
BC_HEIGHT = 10.0        # Barcode-Balkenhöhe mm (flach, damit der Text mittig bleibt)
BC_BOTTOM = 3.0         # Abstand Barcode zur Etiketten-Unterkante mm
BORDER_INSET = 1.5      # Umrandung: Abstand zur Etikettenkante mm
BORDER_LINE = 0.4       # Umrandung: Linienstärke mm

# ---------------------------------------------------------------- Code 128 (Code B)
C128 = [
    "212222","222122","222221","121223","121322","131222","122213","122312","132212","221213",
    "221312","231212","112232","122132","122231","113222","123122","123221","223211","221132",
    "221231","213212","223112","312131","311222","321122","321221","312212","322112","322211",
    "212123","212321","232121","111323","131123","131321","112313","132113","132311","211313",
    "231113","231311","112133","112331","132131","113123","113321","133121","313121","211331",
    "231131","213113","213311","213131","311123","311321","331121","312113","312311","332111",
    "314111","221411","431111","111224","111422","121124","121421","141122","141221","112214",
    "112412","122114","122411","142112","142211","241211","221114","413111","241112","134111",
    "111242","121142","121241","114212","124112","124211","411212","421112","421211","212141",
    "214121","412121","111143","111341","131141","114113","114311","411113","411311","113141",
    "114131","311141","411131","211412","211214","211232","2331112",
]


def code128_bars(text):
    """Balken als Liste (start_modul, breite_module); None bei ungültigen Zeichen."""
    codes = [104]
    for ch in text:
        v = ord(ch) - 32
        if v < 0 or v > 94:
            return None, 0
        codes.append(v)
    checksum = 104
    for i, v in enumerate(codes[1:], start=1):
        checksum += v * i
    codes += [checksum % 103, 106]

    bars, x = [], 10  # 10 Module Ruhezone links
    for c in codes:
        for i, w in enumerate(C128[c]):
            w = int(w)
            if i % 2 == 0:
                bars.append((x, w))
            x += w
    return bars, x + 10  # + Ruhezone rechts


# ---------------------------------------------------------------- Daten aufbauen
def parse_int(value, default=0):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def make_range(a, b):
    a, b = parse_int(a), parse_int(b)
    step = 1 if a <= b else -1
    return list(range(a, b + step, step))


def build_freitext_items(opts):
    """Freitext-Etiketten: je Zeile ein Text; „Text | 5“ überschreibt die Anzahl."""
    default = max(1, min(999, parse_int(opts.get("ftcount"), 1)))
    items = []
    for line in str(opts.get("freitext", "")).splitlines():
        line = line.strip()
        if not line:
            continue
        text, count = line, default
        if "|" in line:
            left, right = line.rsplit("|", 1)
            n = parse_int(right, 0)
            if n > 0 and left.strip():
                text, count = left.strip(), min(999, n)
        items.extend([text] * count)
    return items


def build_items(opts):
    if opts.get("mode") == "freitext":
        return build_freitext_items(opts)
    letters = [s.strip().upper() for s in opts["letters"].split(",") if s.strip()]
    r1 = make_range(opts["n1from"], opts["n1to"])
    r2 = make_range(opts["n2from"], opts["n2to"])
    pad = max(1, parse_int(opts["pad"], 2))

    items = []
    if opts["order"] == "pair":
        for n1 in r1:
            for n2 in r2:
                for letter in letters:
                    items.append((letter, str(n1).zfill(pad), str(n2).zfill(pad)))
    else:
        for letter in letters:
            for n1 in r1:
                for n2 in r2:
                    items.append((letter, str(n1).zfill(pad), str(n2).zfill(pad)))
    return items


def build_sheets(items, startpos):
    slots = [None] * (max(1, min(4, parse_int(startpos, 1))) - 1) + list(items)
    sheets = []
    for i in range(0, max(len(slots), 1), PER_SHEET):
        page = slots[i:i + PER_SHEET]
        page += [None] * (PER_SHEET - len(page))
        sheets.append(page)
    return sheets


def barcode_letters(opts):
    """Buchstaben, die einen Barcode bekommen (leer = alle)."""
    return {s.strip().upper() for s in str(opts.get("bcletters", "")).split(",") if s.strip()}


def label_layout(opts, letter):
    """Geometrie eines Etiketts in mm: (barcode_ja, barcode_oberkante, text_mittellinie).

    Der Text steht IMMER mittig (gleiche Position mit und ohne Barcode);
    der flache Barcode sitzt darunter am unteren Rand.
    """
    use_bc = bool(opts["barcode"])
    if use_bc:
        only = barcode_letters(opts)
        use_bc = not only or letter.upper() in only
    bc_top = (LABEL_H - BC_BOTTOM - BC_HEIGHT) if use_bc else None
    return use_bc, bc_top, LABEL_H / 2


# ---------------------------------------------------------------- Regalschilder (STL / 3D-Druck)
SIGN_FONT_NAME = "Arial"
SIGN_FONT_UNITS = 1024      # interne Auflösung der Schrift-Umrisse (Font-Einheiten)
SIGN_BORDER_INSET = 12.0    # mm: Abstand des eingefrästen Rahmens zur Plattenkante
SIGN_BORDER_LINE = 6.0      # mm: Breite des eingefrästen Rahmens
SIGN_TEXT_GAP = 10.0        # mm: Mindestabstand der Zahl zu Rahmen bzw. Kante
SIGN_MIN_FLOOR = 0.4        # mm: Restboden, damit die Vertiefung nie durchbricht
SIGN_HOLE_D = 6.0           # mm: Durchmesser der Aufhänge-Löcher oben
SIGN_CORNER_SEGS = 12       # Liniensegmente je 90°-Eckenrundung

_sign_contour_cache = {}


def parse_float(value, default=0.0):
    """Zahl mit Punkt ODER Komma (deutsche Eingabe) lesen."""
    try:
        return float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError):
        return default


def fmt(value):
    """Zahl fürs Anzeigen: kompakt und mit Komma (deutsch)."""
    return f"{value:g}".replace(".", ",")


def rounded_rect(x0, y0, x1, y1, radius, segs=SIGN_CORNER_SEGS):
    """Rechteck-Umriss (x0,y0)–(x1,y1) mit abgerundeten Ecken, gegen den Uhrzeigersinn."""
    r = max(0.0, min(radius, (x1 - x0) / 2, (y1 - y0) / 2))
    if r <= 1e-9:
        return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    pts = []
    corners = [((x0 + r, y0 + r), 180.0), ((x1 - r, y0 + r), 270.0),
               ((x1 - r, y1 - r), 0.0), ((x0 + r, y1 - r), 90.0)]
    for (mx, my), start in corners:
        for i in range(segs + 1):
            ang = math.radians(start + 90.0 * i / segs)
            pts.append((mx + r * math.cos(ang), my + r * math.sin(ang)))
    return pts


def circle_poly(cx, cy, r, segs=32):
    return [(cx + r * math.cos(2 * math.pi * i / segs),
             cy + r * math.sin(2 * math.pi * i / segs)) for i in range(segs)]


def sign_hole_contours(o):
    """Zwei Durchgangslöcher in den oberen Ecken (zum Anschrauben/Aufhängen).

    Die Löcher sitzen auf der Ecken-Diagonale mittig im Randstreifen vor dem
    Rahmen – so weichen sie Abrundung und eingefrästem Rahmen automatisch aus.
    """
    if not o["holes"]:
        return []
    w, h = o["width"], o["height"]
    rc = min(o["radius"], w / 2, h / 2)
    band = SIGN_BORDER_INSET / 2
    m = max(band, rc - (rc - band) / math.sqrt(2))
    return [circle_poly(m, h - m, SIGN_HOLE_D / 2),
            circle_poly(w - m, h - m, SIGN_HOLE_D / 2)]


def sign_text_contours(text):
    """Umriss-Polygone des Textes (Arial fett) in Font-Einheiten, y nach unten.

    Nutzt den GDI-Pfad (BeginPath/ExtTextOut/FlattenPath/GetPath), dadurch
    sind keine zusätzlichen Bibliotheken nötig – pywin32 ist ohnehin dabei.
    """
    key = text
    if key in _sign_contour_cache:
        return _sign_contour_cache[key]

    hdc_screen = win32gui.GetDC(0)
    hdc = win32gui.CreateCompatibleDC(hdc_screen)
    lf = win32gui.LOGFONT()
    lf.lfFaceName = SIGN_FONT_NAME
    lf.lfHeight = -SIGN_FONT_UNITS
    lf.lfWeight = 700
    hfont = win32gui.CreateFontIndirect(lf)
    old_font = win32gui.SelectObject(hdc, hfont)
    try:
        win32gui.SetBkMode(hdc, win32con.TRANSPARENT)
        win32gui.BeginPath(hdc)
        win32gui.ExtTextOut(hdc, 0, 0, 0, None, text)
        win32gui.EndPath(hdc)
        win32gui.FlattenPath(hdc)     # Kurven -> Liniensegmente
        points, types = win32gui.GetPath(hdc)
    finally:
        win32gui.SelectObject(hdc, old_font)
        win32gui.DeleteObject(hfont)
        win32gui.DeleteDC(hdc)
        win32gui.ReleaseDC(0, hdc_screen)

    contours, current = [], []
    for (x, y), t in zip(points, types):
        if t & ~win32con.PT_CLOSEFIGURE == win32con.PT_MOVETO:
            if len(current) >= 3:
                contours.append(current)
            current = [(float(x), float(y))]
        else:
            current.append((float(x), float(y)))
    if len(current) >= 3:
        contours.append(current)

    cleaned = []
    for poly in contours:
        if len(poly) > 1 and poly[0] == poly[-1]:
            poly = poly[:-1]
        # doppelte Nachbarpunkte entfernen
        slim = [p for i, p in enumerate(poly) if p != poly[i - 1]]
        if len(slim) >= 3:
            cleaned.append(slim)
    _sign_contour_cache[key] = cleaned
    return cleaned


def _contour_bounds(contours):
    xs = [x for poly in contours for x, _ in poly]
    ys = [y for poly in contours for _, y in poly]
    return min(xs), min(ys), max(xs), max(ys)


def sign_digit_ref_height():
    """Höhe der Ziffer „0“ in Font-Einheiten – Bezug für „Schrifthöhe (mm)“.

    So sind die Ziffern auf allen Schildern gleich groß, egal ob 1 oder 88.
    """
    x0, y0, x1, y1 = _contour_bounds(sign_text_contours("0"))
    return max(1.0, y1 - y0)


def sign_engrave_contours(text, o):
    """Alle einzufräsenden Umrisse in Platten-Millimetern (x rechts, y nach oben)."""
    w, h = o["width"], o["height"]
    contours = []

    if o["border"]:
        a = SIGN_BORDER_INSET
        b = SIGN_BORDER_INSET + SIGN_BORDER_LINE
        if w > 2 * (b + 5) and h > 2 * (b + 5):
            rc = min(o["radius"], w / 2, h / 2)
            contours.append(rounded_rect(a, a, w - a, h - a, rc - a))
            contours.append(rounded_rect(b, b, w - b, h - b, rc - b))

    glyphs = sign_text_contours(text)
    if glyphs:
        scale = o["fontmm"] / sign_digit_ref_height()
        x0, y0, x1, y1 = _contour_bounds(glyphs)
        text_w, text_h = (x1 - x0) * scale, (y1 - y0) * scale

        inset = SIGN_TEXT_GAP + (SIGN_BORDER_INSET + SIGN_BORDER_LINE if o["border"] else 0)
        avail_w = max(10.0, w - 2 * inset)
        avail_h = max(10.0, h - 2 * inset)
        if text_w > avail_w or text_h > avail_h:      # zu breit -> passend verkleinern
            scale *= min(avail_w / text_w, avail_h / text_h)
            text_w, text_h = (x1 - x0) * scale, (y1 - y0) * scale

        # zentrieren; GDI zählt y nach unten, die Platte nach oben -> spiegeln
        off_x = (w - text_w) / 2 - x0 * scale
        off_y = (h - text_h) / 2 + y1 * scale
        for poly in glyphs:
            contours.append([(x * scale + off_x, off_y - y * scale) for x, y in poly])
    return contours


def build_sign_mesh(text, o):
    """Komplettes Dreiecksnetz eines Schildes (Platte + eingefräste Zahl).

    Wie in der Modula-Karten-App: Platte, Löcher und Vertiefung werden als
    Volumenkörper mit echten Booleschen Operationen (manifold3d) verrechnet.
    Das Netz ist dadurch garantiert wasserdicht und alle Konturen geschlossen –
    wichtig, damit der Slicer die Vertiefung sauber zum Einfärben druckt.
    """
    w, h, t = o["width"], o["height"], o["thick"]
    depth = min(o["depth"], t - SIGN_MIN_FLOOR)

    def cross_section(contours):
        return m3d.CrossSection([np.asarray(c, dtype=np.float64) for c in contours],
                                m3d.FillRule.EvenOdd)

    plate = cross_section([rounded_rect(0.0, 0.0, w, h, o["radius"])]
                          + sign_hole_contours(o)).extrude(t)

    engrave = sign_engrave_contours(text, o) if depth > 0 else []
    if engrave:
        cutter = cross_section(engrave).extrude(depth + 1.0)   # ragt oben 1 mm heraus
        plate = plate - cutter.translate((0.0, 0.0, t - depth))

    if plate.is_empty() or plate.status() != m3d.Error.NoError:
        raise RuntimeError("Boolesche Verrechnung ergab ein ungültiges Netz.")

    mesh = plate.to_mesh()
    verts = np.asarray(mesh.vert_properties, dtype=np.float64)[:, :3]
    tris = []
    for i0, i1, i2 in np.asarray(mesh.tri_verts):
        tris.append((tuple(verts[i0]), tuple(verts[i1]), tuple(verts[i2])))
    return tris


def write_binary_stl(path, tris):
    """Binäre STL-Datei schreiben (kompakt, von jedem Slicer lesbar)."""

    def normal(a, b, c):
        ux, uy, uz = b[0] - a[0], b[1] - a[1], b[2] - a[2]
        vx, vy, vz = c[0] - a[0], c[1] - a[1], c[2] - a[2]
        nx, ny, nz = uy * vz - uz * vy, uz * vx - ux * vz, ux * vy - uy * vx
        length = (nx * nx + ny * ny + nz * nz) ** 0.5
        if length < 1e-12:
            return None
        return nx / length, ny / length, nz / length

    with open(path, "wb") as f:
        f.write(b"Kunzer Regalschild (Lager-Etiketten App)".ljust(80, b" "))
        f.write(struct.pack("<I", 0))          # Platzhalter, wird unten korrigiert
        count = 0
        for a, b, c in tris:
            n = normal(a, b, c)
            if n is None:                      # entartete Dreiecke überspringen
                continue
            f.write(struct.pack("<12fH", *n, *a, *b, *c, 0))
            count += 1
        f.seek(80)
        f.write(struct.pack("<I", count))
    return count


def point_in_polygon(x, y, poly):
    inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]
        if (y1 > y) != (y2 > y):
            if x < x1 + (x2 - x1) * (y - y1) / (y2 - y1):
                inside = not inside
    return inside


def contour_depths(contours):
    """Verschachtelungstiefe jedes Umrisses (gerade = Vertiefung, ungerade = Insel)."""
    depths = []
    for i, poly in enumerate(contours):
        x, y = poly[0]
        d = 0
        for j, other in enumerate(contours):
            if j != i and point_in_polygon(x, y, other):
                d += 1
        depths.append(d)
    return depths


# ---------------------------------------------------------------- Druck (GDI)
def print_sheets(printer_name, sheets, opts, devmode=None):
    if devmode is not None:
        # DEVMODE (z. B. gewähltes Papierfach) anwenden
        hdc = win32ui.CreateDCFromHandle(
            win32gui.CreateDC("WINSPOOL", printer_name, devmode))
    else:
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)

    dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
    dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)
    phys_x = hdc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
    phys_y = hdc.GetDeviceCaps(win32con.PHYSICALOFFSETY)

    off_x = float(opts["offx"] or 0)
    off_y = float(opts["offy"] or 0)

    def px(mm):  # X: mm (ab Papierkante) -> Gerätepixel
        return int(round((mm + off_x) / 25.4 * dpi_x)) - phys_x

    def py(mm):
        return int(round((mm + off_y) / 25.4 * dpi_y)) - phys_y

    def font(mm_height, bold=True, name="Arial"):
        return win32ui.CreateFont({
            "name": name,
            "height": -int(round(mm_height / 25.4 * dpi_y)),
            "weight": 700 if bold else 400,
        })

    big = font(float(opts["fontsize"] or 36))
    tmp_fonts = []                                             # verkleinerte Freitext-Schriften
    cutline = max(1, int(round(0.25 / 25.4 * dpi_x)))          # Schnittrahmen (Test)
    border = max(1, int(round(BORDER_LINE / 25.4 * dpi_x)))    # gedruckte Umrandung

    def frame(l, t, r, b, thickness):
        for rect in ((l, t, r, t + thickness), (l, b - thickness, r, b),
                     (l, t, l + thickness, b), (r - thickness, t, r, b)):
            hdc.FillSolidRect(rect, 0)

    def draw_barcode(bars, total, y0):
        if not bars:
            return
        bc_w = total * BC_MODULE
        bc_x = MARGIN_LEFT + (LABEL_W - bc_w) / 2              # mittig
        for start, width in bars:
            hdc.FillSolidRect((px(bc_x + start * BC_MODULE), py(y0),
                               px(bc_x + (start + width) * BC_MODULE),
                               py(y0 + BC_HEIGHT)), 0)

    hdc.StartDoc("Lager-Etiketten")
    try:
        for page in sheets:
            hdc.StartPage()
            hdc.SetBkMode(win32con.TRANSPARENT)
            for pos, item in enumerate(page):
                top = MARGIN_TOP + pos * LABEL_H
                if opts["frames"]:
                    frame(px(MARGIN_LEFT), py(top),
                          px(MARGIN_LEFT + LABEL_W), py(top + LABEL_H), cutline)
                if item is None:
                    continue

                if opts.get("border"):
                    frame(px(MARGIN_LEFT + BORDER_INSET), py(top + BORDER_INSET),
                          px(MARGIN_LEFT + LABEL_W - BORDER_INSET),
                          py(top + LABEL_H - BORDER_INSET), border)

                area_l = MARGIN_LEFT + PAD_X
                area_r = MARGIN_LEFT + LABEL_W - PAD_X

                if isinstance(item, str):          # Freitext, mittig auf dem Etikett
                    bars, total = (code128_bars(item) if opts.get("ftbarcode")
                                   else (None, 0))
                    hdc.SelectObject(big)
                    tw, th = hdc.GetTextExtent(item)
                    avail = px(area_r) - px(area_l)
                    if 0 < avail < tw:             # zu breit -> passend verkleinern
                        small = font(float(opts["fontsize"] or 36) * avail / tw)
                        tmp_fonts.append(small)
                        hdc.SelectObject(small)
                        tw, th = hdc.GetTextExtent(item)
                    ty = py(top + LABEL_H / 2) - th // 2
                    hdc.TextOut(px((area_l + area_r) / 2) - tw // 2, ty, item)
                    draw_barcode(bars, total, top + LABEL_H - BC_BOTTOM - BC_HEIGHT)
                    continue

                letter, n1, n2 = item
                use_bc, bc_top, text_cy = label_layout(opts, letter)

                hdc.SelectObject(big)
                _, th = hdc.GetTextExtent("09")
                ty = py(top + text_cy) - th // 2

                hdc.TextOut(px(area_l), ty, letter)                       # links
                tw, _ = hdc.GetTextExtent(n1)                             # mittig
                hdc.TextOut(px((area_l + area_r) / 2) - tw // 2, ty, n1)
                tw, _ = hdc.GetTextExtent(n2)                             # rechts
                hdc.TextOut(px(area_r) - tw, ty, n2)

                if use_bc:
                    bars, total = code128_bars(letter + n1 + n2)
                    draw_barcode(bars, total, top + bc_top)
            hdc.EndPage()
        hdc.EndDoc()
    except Exception:
        hdc.AbortDoc()
        raise
    finally:
        hdc.DeleteDC()


# ---------------------------------------------------------------- GUI
class App(tk.Tk):
    SCALE = 1.8  # Vorschau: Pixel pro mm

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(False, False)
        self.page_index = 0
        self.sheets = [[None] * PER_SHEET]
        self.devmodes = {}      # je Drucker: DEVMODE aus dem Einstellungs-Dialog
        self.tray_fields = {}   # je Drucker: gemerkte Fach-/Papier-Einstellungen
        self.sign_index = 0
        self.sign_texts = []
        self.freitext_saved = ""

        self.vars = {
            "mode":     tk.StringVar(value="serie"),
            "ftcount":  tk.StringVar(value="1"),
            "ftbarcode": tk.BooleanVar(value=False),
            "letters":  tk.StringVar(value="K,A"),
            "n1from":   tk.StringVar(value="1"),
            "n1to":     tk.StringVar(value="20"),
            "n2from":   tk.StringVar(value="9"),
            "n2to":     tk.StringVar(value="9"),
            "order":    tk.StringVar(value="pair"),
            "startpos": tk.StringVar(value="1"),
            "pad":      tk.StringVar(value="2"),
            "fontsize": tk.StringVar(value="36"),
            "offx":     tk.StringVar(value="0"),
            "offy":     tk.StringVar(value="0"),
            "barcode":  tk.BooleanVar(value=True),
            "bcletters": tk.StringVar(value="K"),
            "border":   tk.BooleanVar(value=True),
            "frames":   tk.BooleanVar(value=False),
            "printer":  tk.StringVar(value=""),
        }
        self.svars = {                       # Reiter „Regalschilder (3D-Druck)“
            "s_prefix": tk.StringVar(value=""),
            "s_from":   tk.StringVar(value="1"),
            "s_to":     tk.StringVar(value="1"),
            "s_pad":    tk.StringVar(value="2"),
            "s_fontmm": tk.StringVar(value="160"),
            "s_width":  tk.StringVar(value="300"),
            "s_height": tk.StringVar(value="300"),
            "s_thick":  tk.StringVar(value="4"),
            "s_depth":  tk.StringVar(value="0,6"),
            "s_radius": tk.StringVar(value="20"),
            "s_border": tk.BooleanVar(value=True),
            "s_holes":  tk.BooleanVar(value=True),
        }
        self.load_settings()
        self.build_ui()
        for var in self.vars.values():
            var.trace_add("write", lambda *_: self.refresh())
        for var in self.svars.values():
            var.trace_add("write", lambda *_: self.refresh_signs())
        self.refresh()
        self.refresh_signs()

    # ---------------- Einstellungen
    def opts(self):
        o = {k: (v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get()))
             for k, v in self.vars.items()}
        o["freitext"] = (self.freitext_text.get("1.0", "end")
                         if hasattr(self, "freitext_text") else self.freitext_saved)
        return o

    def load_settings(self):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.tray_fields = data.get("trays", {})
            self.freitext_saved = str(data.get("freitext", ""))
            for k, v in data.items():
                if k in self.vars:
                    self.vars[k].set(v)
                elif k in self.svars:
                    self.svars[k].set(v)
        except Exception:
            pass

    def save_settings(self):
        try:
            os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
            data = self.opts()
            for k, v in self.svars.items():
                data[k] = bool(v.get()) if isinstance(v, tk.BooleanVar) else v.get()
            data["trays"] = self.tray_fields
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------------- Oberfläche
    def build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=4, pady=4)
        tab_labels = ttk.Frame(self.notebook)
        tab_signs = ttk.Frame(self.notebook)
        self.notebook.add(tab_labels, text="  Etiketten (Papier)  ")
        self.notebook.add(tab_signs, text="  Regalschilder (3D-Druck)  ")
        self.build_label_tab(tab_labels)
        self.build_sign_tab(tab_signs)

    # ---------------- Reiter 1: Etiketten (Papier)
    def build_label_tab(self, parent):
        frame = ttk.Frame(parent, padding=12)
        frame.grid(row=0, column=0, sticky="ns")

        def group(row, text, widget_fn):
            ttk.Label(frame, text=text).grid(row=row, column=0, sticky="w", pady=(8, 1))
            widget_fn(row + 1)

        r = 0
        mode_row = ttk.Frame(frame)
        mode_row.grid(row=r, column=0, sticky="w"); r += 1
        ttk.Radiobutton(mode_row, text="Serie (Buchstabe + Zahlen)", value="serie",
                        variable=self.vars["mode"]).pack(side="left")
        ttk.Radiobutton(mode_row, text="Freitext", value="freitext",
                        variable=self.vars["mode"]).pack(side="left", padx=(14, 0))

        # Serie und Freitext teilen sich dieselbe Zeile – es ist immer nur einer sichtbar
        self.series_frame = ttk.Frame(frame)
        self.series_frame.grid(row=r, column=0, sticky="w")
        self.freitext_frame = ttk.Frame(frame)
        self.freitext_frame.grid(row=r, column=0, sticky="w"); r += 1
        self.build_series_box(self.series_frame)
        self.build_freitext_box(self.freitext_frame)
        self.update_mode_frames()

        def small_row(row, items):
            box = ttk.Frame(frame)
            box.grid(row=row, column=0, sticky="w")
            for label, key, width, values in items:
                ttk.Label(box, text=label).pack(side="left", padx=(0, 3))
                if values:
                    cb = ttk.Combobox(box, state="readonly", width=width,
                                      textvariable=self.vars[key], values=values)
                    cb.pack(side="left", padx=(0, 10))
                else:
                    ttk.Entry(box, textvariable=self.vars[key], width=width).pack(side="left", padx=(0, 10))

        group(r, "Start auf Bogen bei Etikett / führende Nullen:", lambda row: small_row(row, [
            ("Etikett", "startpos", 3, ["1", "2", "3", "4"]),
            ("Stellen", "pad", 3, ["1", "2", "3"])])); r += 2

        group(r, "Schriftgröße (mm) / Versatz X / Y (mm):", lambda row: small_row(row, [
            ("Schrift", "fontsize", 4, None),
            ("X", "offx", 4, None),
            ("Y", "offy", 4, None)])); r += 2

        ttk.Checkbutton(frame, text="Umrandung auf Etikett drucken",
                        variable=self.vars["border"]).grid(row=r, column=0, sticky="w"); r += 1
        ttk.Checkbutton(frame, text="Schnittrahmen andrucken (Test auf Normalpapier)",
                        variable=self.vars["frames"]).grid(row=r, column=0, sticky="w"); r += 1

        ttk.Label(frame, text="Drucker:").grid(row=r, column=0, sticky="w", pady=(10, 1)); r += 1
        printers = self.printer_list()
        self.printer_box = ttk.Combobox(frame, state="readonly", width=34,
                                        textvariable=self.vars["printer"], values=printers)
        if printers and self.vars["printer"].get() not in printers:
            self.vars["printer"].set(self.default_printer(printers))
        self.printer_box.grid(row=r, column=0, sticky="w"); r += 1
        ttk.Button(frame, text="Druckereinstellungen… (Papierfach usw.)",
                   command=self.open_printer_settings).grid(row=r, column=0, sticky="w", pady=(4, 0)); r += 1

        self.info = ttk.Label(frame, text="", wraplength=260, foreground="#444")
        self.info.grid(row=r, column=0, sticky="w", pady=(10, 6)); r += 1

        btns = ttk.Frame(frame)
        btns.grid(row=r, column=0, sticky="w", pady=(4, 0)); r += 1
        ttk.Button(btns, text="🖨 Drucken", command=self.do_print).pack(side="left")
        ttk.Button(btns, text="Testbogen (1. Seite m. Rahmen)",
                   command=self.do_test_print).pack(side="left", padx=(8, 0))

        # ---------------- Vorschau
        right = ttk.Frame(parent, padding=(0, 12, 12, 12))
        right.grid(row=0, column=1, sticky="n")
        nav = ttk.Frame(right)
        nav.pack()
        ttk.Button(nav, text="◀", width=3, command=lambda: self.turn_page(-1)).pack(side="left")
        self.page_label = ttk.Label(nav, text="Seite 1/1", width=12, anchor="center")
        self.page_label.pack(side="left")
        ttk.Button(nav, text="▶", width=3, command=lambda: self.turn_page(1)).pack(side="left")

        s = self.SCALE
        self.canvas = tk.Canvas(right, width=int(PAGE_W * s), height=int(PAGE_H * s),
                                bg="white", highlightthickness=1, highlightbackground="#999")
        self.canvas.pack(pady=(6, 0))

    def build_series_box(self, box):
        r = 0
        ttk.Label(box, text="Buchstabe(n), mehrere mit Komma:").grid(row=r, column=0, sticky="w", pady=(8, 1)); r += 1
        ttk.Entry(box, textvariable=self.vars["letters"], width=14).grid(row=r, column=0, sticky="w"); r += 1

        def range_row(key, label):
            nonlocal r
            ttk.Label(box, text=label).grid(row=r, column=0, sticky="w", pady=(8, 1)); r += 1
            row = ttk.Frame(box)
            row.grid(row=r, column=0, sticky="w"); r += 1
            ttk.Spinbox(row, from_=0, to=999, textvariable=self.vars[key + "from"], width=5).pack(side="left")
            ttk.Label(row, text=" – ").pack(side="left")
            ttk.Spinbox(row, from_=0, to=999, textvariable=self.vars[key + "to"], width=5).pack(side="left")

        range_row("n1", "1. Zahl (z. B. Feld) von – bis:")
        range_row("n2", "2. Zahl (z. B. Ebene) von – bis:")

        ttk.Label(box, text="Reihenfolge:").grid(row=r, column=0, sticky="w", pady=(8, 1)); r += 1
        self.order_box = ttk.Combobox(box, state="readonly", width=34, values=[
            "Buchstaben zusammen (K 01, A 01, K 02 …)",
            "Nacheinander (erst alle K, dann alle A)"])
        self.order_box.current(0 if self.vars["order"].get() == "pair" else 1)
        self.order_box.grid(row=r, column=0, sticky="w"); r += 1
        self.order_box.bind("<<ComboboxSelected>>", lambda e: self.vars["order"].set(
            "pair" if self.order_box.current() == 0 else "seq"))

        bc_row = ttk.Frame(box)
        bc_row.grid(row=r, column=0, sticky="w", pady=(10, 0)); r += 1
        ttk.Checkbutton(bc_row, text="Barcode (Code 128) darunter, nur bei:",
                        variable=self.vars["barcode"]).pack(side="left")
        ttk.Entry(bc_row, textvariable=self.vars["bcletters"], width=6).pack(side="left", padx=(4, 0))
        ttk.Label(box, text="(leer = Barcode auf allen Etiketten)",
                  foreground="#777").grid(row=r, column=0, sticky="w", padx=(20, 0)); r += 1

    def build_freitext_box(self, box):
        r = 0
        ttk.Label(box, text="Texte – je Zeile ein Etikett:").grid(row=r, column=0, sticky="w", pady=(8, 1)); r += 1
        self.freitext_text = tk.Text(box, width=36, height=6, undo=True)
        self.freitext_text.grid(row=r, column=0, sticky="w"); r += 1
        if self.freitext_saved.strip():
            self.freitext_text.insert("1.0", self.freitext_saved.rstrip("\n"))
        self.freitext_text.edit_modified(False)
        self.freitext_text.bind("<<Modified>>", self.on_freitext_change)

        row = ttk.Frame(box)
        row.grid(row=r, column=0, sticky="w", pady=(8, 0)); r += 1
        ttk.Label(row, text="Anzahl je Text:").pack(side="left", padx=(0, 4))
        ttk.Spinbox(row, from_=1, to=999, textvariable=self.vars["ftcount"], width=5).pack(side="left")

        ttk.Label(box, text="Abweichende Anzahl je Zeile mit senkrechtem Strich,\n"
                            "z. B.:  Reserviert Werkstatt | 8",
                  foreground="#777", justify="left").grid(row=r, column=0, sticky="w", pady=(2, 0)); r += 1

        ttk.Checkbutton(box, text="Barcode (Code 128) mit dem Text darunter",
                        variable=self.vars["ftbarcode"]).grid(row=r, column=0, sticky="w", pady=(8, 0)); r += 1

    def update_mode_frames(self):
        if self.vars["mode"].get() == "freitext":
            self.series_frame.grid_remove()
            self.freitext_frame.grid()
        else:
            self.freitext_frame.grid_remove()
            self.series_frame.grid()

    def on_freitext_change(self, _event=None):
        if self.freitext_text.edit_modified():
            self.freitext_text.edit_modified(False)
            self.refresh()

    # ---------------- Reiter 2: Regalschilder (3D-Druck)
    SIGN_PREVIEW = 470  # Vorschau-Kantenlänge in Pixeln

    def build_sign_tab(self, parent):
        form = ttk.Frame(parent, padding=12)
        form.grid(row=0, column=0, sticky="ns")
        r = 0

        ttk.Label(form, text="Große Schilder zum Markieren der Regale – als STL-Datei "
                             "für den 3D-Drucker. Die Zahl wird vertieft eingelassen "
                             "und lässt sich so sauber einfärben.",
                  wraplength=280, foreground="#444").grid(row=r, column=0, sticky="w"); r += 1

        ttk.Label(form, text="Buchstabe davor (optional):").grid(row=r, column=0, sticky="w", pady=(10, 1)); r += 1
        ttk.Entry(form, textvariable=self.svars["s_prefix"], width=8).grid(row=r, column=0, sticky="w"); r += 1

        ttk.Label(form, text="Zahl von – bis (je Zahl ein Schild):").grid(row=r, column=0, sticky="w", pady=(8, 1)); r += 1
        box = ttk.Frame(form)
        box.grid(row=r, column=0, sticky="w"); r += 1
        ttk.Spinbox(box, from_=0, to=999, textvariable=self.svars["s_from"], width=5).pack(side="left")
        ttk.Label(box, text=" – ").pack(side="left")
        ttk.Spinbox(box, from_=0, to=999, textvariable=self.svars["s_to"], width=5).pack(side="left")

        def pair_row(label, items, pady=(8, 1)):
            nonlocal r
            ttk.Label(form, text=label).grid(row=r, column=0, sticky="w", pady=pady); r += 1
            row = ttk.Frame(form)
            row.grid(row=r, column=0, sticky="w"); r += 1
            for text, key, width, values in items:
                ttk.Label(row, text=text).pack(side="left", padx=(0, 3))
                if values:
                    ttk.Combobox(row, state="readonly", width=width, values=values,
                                 textvariable=self.svars[key]).pack(side="left", padx=(0, 10))
                else:
                    ttk.Entry(row, textvariable=self.svars[key], width=width).pack(side="left", padx=(0, 10))

        pair_row("Führende Nullen / Schrifthöhe:", [
            ("Stellen", "s_pad", 3, ["1", "2", "3"]),
            ("Schrift (mm)", "s_fontmm", 5, None)])
        pair_row("Platte Breite × Höhe / Ecken-Radius (mm):", [
            ("Breite", "s_width", 5, None),
            ("Höhe", "s_height", 5, None),
            ("Radius", "s_radius", 4, None)])
        pair_row("Dicke / Vertiefung der Zahl (mm):", [
            ("Dicke", "s_thick", 5, None),
            ("Vertiefung", "s_depth", 5, None)])

        ttk.Label(form, text="Tipp: 0,6 mm Vertiefung = 3 Druckschichten à 0,2 mm – "
                             "ideal zum Ausmalen oder für einen Filament-Farbwechsel. "
                             "Weniger als eine Schichthöhe (ca. 0,2 mm) ist im Druck "
                             "nicht sichtbar.",
                  wraplength=280, foreground="#777").grid(row=r, column=0, sticky="w", pady=(4, 0)); r += 1

        ttk.Checkbutton(form, text="Rahmen mit einfräsen (wie auf den Etiketten)",
                        variable=self.svars["s_border"]).grid(row=r, column=0, sticky="w", pady=(8, 0)); r += 1
        ttk.Checkbutton(form, text=f"Lochung oben: 2 Löcher (Ø {fmt(SIGN_HOLE_D)} mm) in den Ecken",
                        variable=self.svars["s_holes"]).grid(row=r, column=0, sticky="w", pady=(4, 0)); r += 1

        self.sign_info = ttk.Label(form, text="", wraplength=280, foreground="#444")
        self.sign_info.grid(row=r, column=0, sticky="w", pady=(10, 6)); r += 1

        ttk.Button(form, text="💾 STL-Datei(en) erstellen…",
                   command=self.do_export_stl).grid(row=r, column=0, sticky="w"); r += 1

        # ---------------- Vorschau
        right = ttk.Frame(parent, padding=(0, 12, 12, 12))
        right.grid(row=0, column=1, sticky="n")
        nav = ttk.Frame(right)
        nav.pack()
        ttk.Button(nav, text="◀", width=3, command=lambda: self.turn_sign(-1)).pack(side="left")
        self.sign_page_label = ttk.Label(nav, text="Schild 1/1", width=14, anchor="center")
        self.sign_page_label.pack(side="left")
        ttk.Button(nav, text="▶", width=3, command=lambda: self.turn_sign(1)).pack(side="left")

        self.sign_canvas = tk.Canvas(right, width=self.SIGN_PREVIEW, height=self.SIGN_PREVIEW,
                                     bg="white", highlightthickness=1, highlightbackground="#999")
        self.sign_canvas.pack(pady=(6, 0))
        self.sign_caption = ttk.Label(right, text="", foreground="#444")
        self.sign_caption.pack(pady=(4, 0))

    def sign_opts(self):
        g = self.svars
        return {
            "prefix": g["s_prefix"].get().strip().upper(),
            "nfrom":  g["s_from"].get(),
            "nto":    g["s_to"].get(),
            "pad":    max(1, parse_int(g["s_pad"].get(), 2)),
            "fontmm": max(10.0, parse_float(g["s_fontmm"].get(), 160.0)),
            "width":  max(40.0, parse_float(g["s_width"].get(), 300.0)),
            "height": max(40.0, parse_float(g["s_height"].get(), 300.0)),
            "thick":  max(1.0, parse_float(g["s_thick"].get(), 4.0)),
            "depth":  max(0.0, parse_float(g["s_depth"].get(), 0.6)),
            "radius": max(0.0, parse_float(g["s_radius"].get(), 20.0)),
            "border": bool(g["s_border"].get()),
            "holes":  bool(g["s_holes"].get()),
        }

    def refresh_signs(self):
        o = self.sign_opts()
        self.sign_texts = [o["prefix"] + str(n).zfill(o["pad"])
                           for n in make_range(o["nfrom"], o["nto"])]
        self.sign_index = max(0, min(self.sign_index, len(self.sign_texts) - 1))

        depth = min(o["depth"], o["thick"] - SIGN_MIN_FLOOR)
        lines = [f"{len(self.sign_texts)} Schild(er) à "
                 f"{fmt(o['width'])} × {fmt(o['height'])} mm – je eine STL-Datei."]
        if o["depth"] > depth:
            lines.append(f"Hinweis: Vertiefung wird auf {fmt(depth)} mm begrenzt, "
                         f"damit ein Restboden von {fmt(SIGN_MIN_FLOOR)} mm bleibt.")
        elif 0 < o["depth"] < 0.2:
            lines.append("⚠ Vertiefung unter 0,2 mm ist beim 3D-Druck praktisch "
                         "unsichtbar (Schichthöhe). Empfehlung: 0,6 mm.")
        self.sign_info.config(text="\n".join(lines))
        self.draw_sign_preview()
        self.save_settings()

    def turn_sign(self, step):
        self.sign_index = max(0, min(len(self.sign_texts) - 1, self.sign_index + step))
        self.draw_sign_preview()

    def draw_sign_preview(self):
        c = self.sign_canvas
        c.delete("all")
        o = self.sign_opts()
        total = max(1, len(self.sign_texts))
        self.sign_page_label.config(text=f"Schild {self.sign_index + 1}/{total}")
        if not self.sign_texts:
            self.sign_caption.config(text="")
            return
        text = self.sign_texts[self.sign_index]

        size = self.SIGN_PREVIEW
        s = min((size - 30) / o["width"], (size - 30) / o["height"])
        ox = (size - o["width"] * s) / 2
        oy = (size - o["height"] * s) / 2
        plate = "#ffcc00"   # Signalgelb, Schrift/Rahmen schwarz

        def cx(x):
            return ox + x * s

        def cy(y):  # Platte zählt y nach oben, der Canvas nach unten
            return oy + (o["height"] - y) * s

        def poly_pts(poly):
            pts = []
            for x, y in poly:
                pts += [cx(x), cy(y)]
            return pts

        c.create_polygon(poly_pts(rounded_rect(0.0, 0.0, o["width"], o["height"], o["radius"])),
                         fill=plate, outline="#808080")
        for hole in sign_hole_contours(o):
            c.create_polygon(poly_pts(hole), fill="white", outline="#808080")
        if win32gui is None:
            c.create_text(size / 2, size / 2, fill="#888", justify="center",
                          text="pywin32 fehlt – keine Vorschau möglich.\n"
                               "Bitte einmal ausführen:  pip install pywin32")
            self.sign_caption.config(text="")
            return

        contours = sign_engrave_contours(text, o)
        for d, poly in sorted(zip(contour_depths(contours), contours),
                              key=lambda pair: pair[0]):
            c.create_polygon(poly_pts(poly), fill="black" if d % 2 == 0 else plate, outline="")

        depth = min(o["depth"], o["thick"] - SIGN_MIN_FLOOR)
        extra = f" · 2 Löcher Ø {fmt(SIGN_HOLE_D)} mm" if o["holes"] else ""
        self.sign_caption.config(
            text=f"Regalschild_{text}.stl · {fmt(o['width'])} × {fmt(o['height'])} × "
                 f"{fmt(o['thick'])} mm · Zahl {fmt(depth)} mm vertieft{extra}")

    def do_export_stl(self):
        if win32gui is None:
            messagebox.showerror(APP_TITLE, "pywin32 ist nicht installiert –\n"
                                            "bitte einmal ausführen:  pip install pywin32")
            return
        if m3d is None:
            messagebox.showerror(APP_TITLE, "manifold3d ist nicht installiert –\n"
                                            "bitte einmal ausführen:  pip install manifold3d numpy")
            return
        o = self.sign_opts()
        texts = list(self.sign_texts)
        if not texts:
            messagebox.showerror(APP_TITLE, "Keine Schilder – bitte den Zahlenbereich prüfen.")
            return

        def filename(text):
            clean = re.sub(r"[^A-Za-z0-9_-]+", "", text) or "Schild"
            return f"Regalschild_{clean}.stl"

        if len(texts) == 1:
            path = filedialog.asksaveasfilename(
                title="STL-Datei speichern", defaultextension=".stl",
                filetypes=[("STL-Datei", "*.stl")], initialfile=filename(texts[0]))
            if not path:
                return
            targets = [(texts[0], path)]
        else:
            folder = filedialog.askdirectory(
                title=f"Ordner für {len(texts)} STL-Dateien wählen")
            if not folder:
                return
            targets = [(t, os.path.join(folder, filename(t))) for t in texts]

        self.config(cursor="watch")
        self.update_idletasks()
        try:
            for text, path in targets:
                write_binary_stl(path, build_sign_mesh(text, o))
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"STL konnte nicht erstellt werden:\n{exc}")
            return
        finally:
            self.config(cursor="")

        depth = min(o["depth"], o["thick"] - SIGN_MIN_FLOOR)
        where = targets[0][1] if len(targets) == 1 else os.path.dirname(targets[0][1])
        messagebox.showinfo(
            APP_TITLE,
            f"{len(targets)} STL-Datei(en) erstellt:\n{where}\n\n"
            "Zum Drucken einfach im Slicer (z. B. OrcaSlicer) öffnen.\n"
            f"Tipp: Filament-Farbwechsel bei {fmt(o['thick'] - depth)} mm Höhe – "
            "dann bekommt die Deckfläche eine andere Farbe und die vertiefte "
            "Zahl bleibt in der Grundfarbe stehen.")

    # ---------------- Drucker
    @staticmethod
    def printer_list():
        if win32print is None:
            return []
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return [p[2] for p in win32print.EnumPrinters(flags)]

    @staticmethod
    def default_printer(printers):
        try:
            d = win32print.GetDefaultPrinter()
            return d if d in printers else printers[0]
        except Exception:
            return printers[0] if printers else ""

    # ---------------- Druckereinstellungen (DEVMODE / Papierfach)
    @staticmethod
    def fresh_devmode(handle, name):
        dm = win32print.GetPrinter(handle, 2).get("pDevMode")
        if dm is None:  # manche Netzwerkdrucker liefern hier nichts
            win32print.DocumentProperties(0, handle, name, None, None, 0)
            dm = win32print.GetPrinter(handle, 2).get("pDevMode")
        return dm

    def open_printer_settings(self):
        name = self.vars["printer"].get()
        if win32print is None or not name:
            messagebox.showerror(APP_TITLE, "Kein Drucker ausgewählt.")
            return
        try:
            handle = win32print.OpenPrinter(name)
            try:
                dm = self.devmodes.get(name) or self.fresh_devmode(handle, name)
                if dm is None:
                    raise RuntimeError("Der Treiber liefert keine Einstellungen (DEVMODE).")
                flags = win32con.DM_IN_BUFFER | win32con.DM_IN_PROMPT | win32con.DM_OUT_BUFFER
                result = win32print.DocumentProperties(
                    self.winfo_id(), handle, name, dm, dm, flags)
                if result == 1:  # IDOK
                    self.devmodes[name] = dm
                    self.tray_fields[name] = {
                        "DefaultSource": dm.DefaultSource,
                        "PaperSize": dm.PaperSize,
                        "Orientation": dm.Orientation,
                    }
                    self.save_settings()
            finally:
                win32print.ClosePrinter(handle)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Druckereinstellungen konnten nicht geöffnet werden:\n{exc}")

    def devmode_for_print(self, name):
        """DEVMODE für den Druck: aus Dialog dieser Sitzung oder gemerktem Papierfach."""
        if name in self.devmodes:
            return self.devmodes[name]
        saved = self.tray_fields.get(name)
        if not saved or win32print is None:
            return None
        try:
            handle = win32print.OpenPrinter(name)
            try:
                dm = self.fresh_devmode(handle, name)
                if dm is None:
                    return None
                dm.DefaultSource = int(saved["DefaultSource"])
                dm.PaperSize = int(saved["PaperSize"])
                dm.Orientation = int(saved["Orientation"])
                dm.Fields |= (win32con.DM_DEFAULTSOURCE | win32con.DM_PAPERSIZE
                              | win32con.DM_ORIENTATION)
                win32print.DocumentProperties(
                    0, handle, name, dm, dm,
                    win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
                self.devmodes[name] = dm
                return dm
            finally:
                win32print.ClosePrinter(handle)
        except Exception:
            return None

    # ---------------- Aktualisieren
    def refresh(self):
        opts = self.opts()
        if hasattr(self, "series_frame"):
            self.update_mode_frames()
        items = build_items(opts)
        self.sheets = build_sheets(items, opts["startpos"])
        self.page_index = min(self.page_index, len(self.sheets) - 1)

        rest = (parse_int(opts["startpos"], 1) - 1 + len(items)) % PER_SHEET
        note = f" – letzter Bogen: {PER_SHEET - rest} Etikett(en) frei." if rest else ""
        if opts.get("mode") == "freitext" and opts.get("ftbarcode"):
            bad = [t for t in dict.fromkeys(items) if code128_bars(t)[0] is None]
            if bad:
                note += ("\n⚠ Kein Barcode möglich (Umlaut/Sonderzeichen): "
                         + ", ".join(bad[:4]) + ("…" if len(bad) > 4 else ""))
        self.info.config(text=f"{len(items)} Etiketten auf {len(self.sheets)} Bogen/Bögen{note}")
        self.draw_preview()
        self.save_settings()

    def turn_page(self, step):
        self.page_index = max(0, min(len(self.sheets) - 1, self.page_index + step))
        self.draw_preview()

    def draw_preview(self):
        s = self.SCALE
        c = self.canvas
        c.delete("all")
        opts = self.opts()
        self.page_label.config(text=f"Seite {self.page_index + 1}/{len(self.sheets)}")

        font_mm = parse_int(opts["fontsize"], 36)
        big_px = max(6, int(font_mm * s))
        big = ("Arial", -big_px, "bold")
        measure = tkfont.Font(family="Arial", size=-big_px, weight="bold")

        def draw_bars(bars, total, y0_mm):
            if not bars:
                return
            bc_w = total * BC_MODULE * s
            bc_x = (MARGIN_LEFT + LABEL_W / 2) * s - bc_w / 2
            y0 = y0_mm * s
            for start, width in bars:
                c.create_rectangle(bc_x + start * BC_MODULE * s, y0,
                                   bc_x + (start + width) * BC_MODULE * s,
                                   y0 + BC_HEIGHT * s, fill="black", width=0)

        for pos, item in enumerate(self.sheets[self.page_index]):
            top = MARGIN_TOP + pos * LABEL_H
            l, t = MARGIN_LEFT * s, top * s
            rr, b = (MARGIN_LEFT + LABEL_W) * s, (top + LABEL_H) * s
            c.create_rectangle(l, t, rr, b, outline="#bbb", dash=(3, 3))
            if item is None:
                continue
            if opts.get("border"):
                c.create_rectangle(l + BORDER_INSET * s, t + BORDER_INSET * s,
                                   rr - BORDER_INSET * s, b - BORDER_INSET * s,
                                   outline="black", width=max(1, int(BORDER_LINE * s * 2)))

            area_l = (MARGIN_LEFT + PAD_X) * s
            area_r = (MARGIN_LEFT + LABEL_W - PAD_X) * s

            if isinstance(item, str):          # Freitext, mittig auf dem Etikett
                fnt = big
                tw = measure.measure(item)
                avail = area_r - area_l
                if 0 < avail < tw:             # zu breit -> passend verkleinern
                    fnt = ("Arial", -max(6, int(big_px * avail / tw)), "bold")
                c.create_text((area_l + area_r) / 2, (top + LABEL_H / 2) * s,
                              text=item, font=fnt, anchor="center")
                if opts.get("ftbarcode"):
                    bars, total = code128_bars(item)
                    draw_bars(bars, total, top + LABEL_H - BC_BOTTOM - BC_HEIGHT)
                continue

            letter, n1, n2 = item
            use_bc, bc_top, text_cy = label_layout(opts, letter)
            cy = (top + text_cy) * s

            c.create_text(area_l, cy, text=letter, font=big, anchor="w")
            c.create_text((area_l + area_r) / 2, cy, text=n1, font=big, anchor="center")
            c.create_text(area_r, cy, text=n2, font=big, anchor="e")

            if use_bc:
                bars, total = code128_bars(letter + n1 + n2)
                draw_bars(bars, total, top + bc_top)

    # ---------------- Drucken
    def do_print(self, sheets=None, force_frames=False):
        if win32ui is None:
            messagebox.showerror(APP_TITLE, "pywin32 ist nicht installiert –\n"
                                            "bitte einmal ausführen:  pip install pywin32")
            return
        printer = self.vars["printer"].get()
        if not printer:
            messagebox.showerror(APP_TITLE, "Kein Drucker ausgewählt.")
            return
        opts = self.opts()
        if force_frames:
            opts["frames"] = True
        use_sheets = sheets if sheets is not None else self.sheets
        try:
            print_sheets(printer, use_sheets, opts, devmode=self.devmode_for_print(printer))
            messagebox.showinfo(APP_TITLE, f"{len(use_sheets)} Bogen/Bögen an\n„{printer}“ gesendet.")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Druckfehler:\n{exc}")

    def do_test_print(self):
        self.do_print(sheets=self.sheets[:1], force_frames=True)


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def main():
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except Exception:
            pass
    app = App()
    try:
        app.iconbitmap(resource_path("icon.ico"))
    except Exception:
        pass
    app.mainloop()


if __name__ == "__main__":
    main()
