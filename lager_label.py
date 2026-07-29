# -*- coding: utf-8 -*-
"""
Lager-Etiketten Generator – Avery Zweckform L4761-100
A4-Bogen, 4 Etiketten je 192 x 61 mm.

GUI (tkinter) mit Vorschau und Direktdruck über Windows GDI (pywin32),
damit die Positionen millimetergenau auf dem Etikettenbogen landen.
"""

import ctypes
import json
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import win32con
    import win32gui
    import win32print
    import win32ui
except ImportError:
    win32con = win32gui = win32print = win32ui = None

APP_ID = "Kunzer.LagerLabel.Etiketten"
APP_TITLE = "Lager-Etiketten – Zweckform L4761 (192 × 61 mm)"
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


def build_items(opts):
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
    cutline = max(1, int(round(0.25 / 25.4 * dpi_x)))          # Schnittrahmen (Test)
    border = max(1, int(round(BORDER_LINE / 25.4 * dpi_x)))    # gedruckte Umrandung

    def frame(l, t, r, b, thickness):
        for rect in ((l, t, r, t + thickness), (l, b - thickness, r, b),
                     (l, t, l + thickness, b), (r - thickness, t, r, b)):
            hdc.FillSolidRect(rect, 0)

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

                letter, n1, n2 = item
                if opts.get("border"):
                    frame(px(MARGIN_LEFT + BORDER_INSET), py(top + BORDER_INSET),
                          px(MARGIN_LEFT + LABEL_W - BORDER_INSET),
                          py(top + LABEL_H - BORDER_INSET), border)

                area_l = MARGIN_LEFT + PAD_X
                area_r = MARGIN_LEFT + LABEL_W - PAD_X
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
                    bc_w = total * BC_MODULE
                    bc_x = MARGIN_LEFT + (LABEL_W - bc_w) / 2             # mittig darunter
                    y0 = top + bc_top
                    for start, width in bars:
                        hdc.FillSolidRect((px(bc_x + start * BC_MODULE), py(y0),
                                           px(bc_x + (start + width) * BC_MODULE),
                                           py(y0 + BC_HEIGHT)), 0)
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

        self.vars = {
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
        self.load_settings()
        self.build_ui()
        for var in self.vars.values():
            var.trace_add("write", lambda *_: self.refresh())
        self.refresh()

    # ---------------- Einstellungen
    def opts(self):
        o = {k: (v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get()))
             for k, v in self.vars.items()}
        return o

    def load_settings(self):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.tray_fields = data.get("trays", {})
            for k, v in data.items():
                if k in self.vars:
                    self.vars[k].set(v)
        except Exception:
            pass

    def save_settings(self):
        try:
            os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
            data = self.opts()
            data["trays"] = self.tray_fields
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------------- Oberfläche
    def build_ui(self):
        frame = ttk.Frame(self, padding=12)
        frame.grid(row=0, column=0, sticky="ns")

        def group(row, text, widget_fn):
            ttk.Label(frame, text=text).grid(row=row, column=0, sticky="w", pady=(8, 1))
            widget_fn(row + 1)

        r = 0
        ttk.Label(frame, text="Buchstabe(n), mehrere mit Komma:").grid(row=r, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.vars["letters"], width=14).grid(row=r + 1, column=0, sticky="w")
        r += 2

        def range_row(row, key):
            box = ttk.Frame(frame)
            box.grid(row=row, column=0, sticky="w")
            ttk.Spinbox(box, from_=0, to=999, textvariable=self.vars[key + "from"], width=5).pack(side="left")
            ttk.Label(box, text=" – ").pack(side="left")
            ttk.Spinbox(box, from_=0, to=999, textvariable=self.vars[key + "to"], width=5).pack(side="left")

        group(r, "1. Zahl (z. B. Feld) von – bis:", lambda row: range_row(row, "n1")); r += 2
        group(r, "2. Zahl (z. B. Ebene) von – bis:", lambda row: range_row(row, "n2")); r += 2

        ttk.Label(frame, text="Reihenfolge:").grid(row=r, column=0, sticky="w", pady=(8, 1)); r += 1
        self.order_box = ttk.Combobox(frame, state="readonly", width=34, values=[
            "Buchstaben zusammen (K 01, A 01, K 02 …)",
            "Nacheinander (erst alle K, dann alle A)"])
        self.order_box.current(0 if self.vars["order"].get() == "pair" else 1)
        self.order_box.grid(row=r, column=0, sticky="w"); r += 1
        self.order_box.bind("<<ComboboxSelected>>", lambda e: self.vars["order"].set(
            "pair" if self.order_box.current() == 0 else "seq"))

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

        bc_row = ttk.Frame(frame)
        bc_row.grid(row=r, column=0, sticky="w", pady=(10, 0)); r += 1
        ttk.Checkbutton(bc_row, text="Barcode (Code 128) darunter, nur bei:",
                        variable=self.vars["barcode"]).pack(side="left")
        ttk.Entry(bc_row, textvariable=self.vars["bcletters"], width=6).pack(side="left", padx=(4, 0))
        ttk.Label(frame, text="(leer = Barcode auf allen Etiketten)",
                  foreground="#777").grid(row=r, column=0, sticky="w", padx=(20, 0)); r += 1
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
        right = ttk.Frame(self, padding=(0, 12, 12, 12))
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
        items = build_items(opts)
        self.sheets = build_sheets(items, opts["startpos"])
        self.page_index = min(self.page_index, len(self.sheets) - 1)

        rest = (parse_int(opts["startpos"], 1) - 1 + len(items)) % PER_SHEET
        note = f" – letzter Bogen: {PER_SHEET - rest} Etikett(en) frei." if rest else ""
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
        big = ("Arial", -max(6, int(font_mm * s)), "bold")

        for pos, item in enumerate(self.sheets[self.page_index]):
            top = MARGIN_TOP + pos * LABEL_H
            l, t = MARGIN_LEFT * s, top * s
            rr, b = (MARGIN_LEFT + LABEL_W) * s, (top + LABEL_H) * s
            c.create_rectangle(l, t, rr, b, outline="#bbb", dash=(3, 3))
            if item is None:
                continue
            letter, n1, n2 = item
            if opts.get("border"):
                c.create_rectangle(l + BORDER_INSET * s, t + BORDER_INSET * s,
                                   rr - BORDER_INSET * s, b - BORDER_INSET * s,
                                   outline="black", width=max(1, int(BORDER_LINE * s * 2)))

            area_l = (MARGIN_LEFT + PAD_X) * s
            area_r = (MARGIN_LEFT + LABEL_W - PAD_X) * s
            use_bc, bc_top, text_cy = label_layout(opts, letter)
            cy = (top + text_cy) * s

            c.create_text(area_l, cy, text=letter, font=big, anchor="w")
            c.create_text((area_l + area_r) / 2, cy, text=n1, font=big, anchor="center")
            c.create_text(area_r, cy, text=n2, font=big, anchor="e")

            if use_bc:
                bars, total = code128_bars(letter + n1 + n2)
                bc_w = total * BC_MODULE * s
                bc_x = (MARGIN_LEFT + LABEL_W / 2) * s - bc_w / 2
                y0 = (top + bc_top) * s
                for start, width in bars:
                    c.create_rectangle(bc_x + start * BC_MODULE * s, y0,
                                       bc_x + (start + width) * BC_MODULE * s,
                                       y0 + BC_HEIGHT * s, fill="black", width=0)

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
