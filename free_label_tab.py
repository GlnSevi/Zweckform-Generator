"""The independent 'Frei / Excel' tab for the Windows application."""

import math
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from free_labels import (FORMATS, FORMATS_BY_CODE, LabelFormat, excel_items,
                         paginate, print_excel_sheets, read_excel, render_sheet)


class FreeLabelTab(ttk.Frame):
    def __init__(self, parent, app, saved=None):
        super().__init__(parent)
        self.app = app
        self.workbook = {}
        self.items = []
        self.sheets = []
        self.page_index = 0
        self.layout = None
        self.pending = None
        self.loading = False
        self.import_queue = queue.Queue()
        defaults = {"format": "L7160", "width": "63,5", "height": "38,1",
                    "columns": "3", "rows": "7", "left": "7,21", "top": "15,15",
                    "gap_x": "2,54", "gap_y": "0", "font_mm": "7",
                    "off_x": "0", "off_y": "0", "start": "1",
                    "skip_header": False, "keep_blanks": True, "border": False}
        saved = saved if isinstance(saved, dict) else {}
        self.vars = {k: (tk.BooleanVar(self, value=saved.get(k, v)) if isinstance(v, bool)
                         else tk.StringVar(self, value=saved.get(k, v))) for k, v in defaults.items()}
        if self.vars["format"].get() not in (*FORMATS_BY_CODE, "Eigenes Format"):
            self.vars["format"].set("L7160")
        self.sheet_name = tk.StringVar(self)
        self.format_name = tk.StringVar(self)
        self.build_ui()
        self.select_format(initial=True)
        for key, var in self.vars.items():
            if key != "format":
                var.trace_add("write", self.schedule_refresh)
        self.refresh()

    def settings(self):
        return {key: var.get() for key, var in self.vars.items()}

    def build_ui(self):
        controls = ttk.Frame(self, padding=12)
        controls.grid(row=0, column=0, sticky="n")
        ttk.Label(controls, text="Freie Etiketten aus Excel", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(controls, text="Spalte A: eine Zelle = ein Etikett.\nReihenfolge: links → rechts, dann nächste Reihe.",
                  foreground="#555").pack(anchor="w", pady=(4, 10))
        self.import_button = ttk.Button(controls, text="Excel-Datei öffnen…", command=self.open_excel)
        self.import_button.pack(anchor="w")
        self.file_label = ttk.Label(controls, text="Noch keine Datei geladen", wraplength=370, foreground="#555")
        self.file_label.pack(anchor="w", pady=(4, 6))
        line = ttk.Frame(controls)
        line.pack(fill="x")
        ttk.Label(line, text="Arbeitsblatt:").pack(side="left", padx=(0, 6))
        self.sheet_box = ttk.Combobox(line, textvariable=self.sheet_name, state="readonly", width=27)
        self.sheet_box.pack(side="left")
        self.sheet_box.bind("<<ComboboxSelected>>", self.source_changed)
        ttk.Checkbutton(controls, text="Erste Zeile ist eine Überschrift (überspringen)",
                        variable=self.vars["skip_header"]).pack(anchor="w", pady=(5, 0))
        ttk.Checkbutton(controls, text="Leere Zeilen als freie Etiketten beibehalten",
                        variable=self.vars["keep_blanks"]).pack(anchor="w")

        ttk.Label(controls, text="Avery Zweckform / Bogenformat:").pack(anchor="w", pady=(10, 2))
        self.format_box = ttk.Combobox(controls, state="readonly", textvariable=self.format_name,
            values=[f.description for f in FORMATS] + ["Eigenes Format"], width=49)
        self.format_box.pack(anchor="w")
        self.format_box.bind("<<ComboboxSelected>>", self.select_format)
        self.geometry_frame = ttk.Frame(controls)
        self.geometry_frame.pack(anchor="w", pady=(5, 2))
        self.geometry_entries = []
        for row, fields in enumerate([[("Breite mm", "width"), ("Höhe mm", "height")],
                                     [("Spalten", "columns"), ("Reihen", "rows")],
                                     [("Rand links mm", "left"), ("Rand oben mm", "top")],
                                     [("Abstand X mm", "gap_x"), ("Abstand Y mm", "gap_y")]]):
            for col, (label, key) in enumerate(fields):
                ttk.Label(self.geometry_frame, text=label).grid(row=row, column=col * 2, sticky="w", padx=(0, 5))
                entry = ttk.Entry(self.geometry_frame, textvariable=self.vars[key], width=7)
                entry.grid(row=row, column=col * 2 + 1, padx=(0, 10), pady=1)
                self.geometry_entries.append(entry)

        line = ttk.Frame(controls)
        line.pack(anchor="w", pady=(8, 4))
        ttk.Label(line, text="Start bei Etikett:").pack(side="left", padx=(0, 5))
        self.start_box = ttk.Spinbox(line, from_=1, to=21, width=5, textvariable=self.vars["start"])
        self.start_box.pack(side="left")
        ttk.Label(line, text="Schrift (mm):").pack(side="left", padx=(15, 5))
        ttk.Entry(line, textvariable=self.vars["font_mm"], width=5).pack(side="left")
        line = ttk.Frame(controls)
        line.pack(anchor="w", pady=(0, 4))
        ttk.Label(line, text="Feinjustierung (mm):").pack(side="left")
        for title, key in [("X", "off_x"), ("Y", "off_y")]:
            ttk.Label(line, text=title).pack(side="left", padx=(8, 4))
            ttk.Entry(line, textvariable=self.vars[key], width=6).pack(side="left")
        ttk.Checkbutton(controls, text="Umrandung drucken", variable=self.vars["border"]).pack(anchor="w")
        ttk.Label(controls, text="Drucker:").pack(anchor="w", pady=(8, 2))
        ttk.Combobox(controls, state="readonly", textvariable=self.app.vars["printer"],
                     values=self.app.printer_list(), width=49).pack(anchor="w")
        ttk.Button(controls, text="Druckereinstellungen… (Papierfach usw.)",
                   command=self.app.open_printer_settings).pack(anchor="w", pady=(3, 4))
        self.info = ttk.Label(controls, text="", wraplength=370, foreground="#444")
        self.info.pack(anchor="w", pady=(4, 5))
        buttons = ttk.Frame(controls)
        buttons.pack(anchor="w")
        self.print_button = ttk.Button(buttons, text="Drucken", command=self.do_print)
        self.print_button.pack(side="left")
        self.test_button = ttk.Button(buttons, text="Testbogen (1. Seite mit Rahmen)",
                                      command=lambda: self.do_print(test=True))
        self.test_button.pack(side="left", padx=(8, 0))
        ttk.Label(controls, text="A4 · Hochformat · 100 % Größe\nZuerst auf Normalpapier testen und mit dem Bogen vergleichen.",
                  foreground="#666", wraplength=370).pack(anchor="w", pady=(7, 0))

        right = ttk.Frame(self, padding=(0, 12, 12, 12))
        right.grid(row=0, column=1, sticky="n")
        nav = ttk.Frame(right)
        nav.pack()
        self.prev_button = ttk.Button(nav, text="◀", width=3, command=lambda: self.turn_page(-1))
        self.prev_button.pack(side="left")
        self.page_label = ttk.Label(nav, text="Seite 1/1", width=17, anchor="center")
        self.page_label.pack(side="left")
        self.next_button = ttk.Button(nav, text="▶", width=3, command=lambda: self.turn_page(1))
        self.next_button.pack(side="left")
        self.canvas = tk.Canvas(right, width=round(210 * self.app.SCALE), height=round(297 * self.app.SCALE),
                                bg="white", highlightthickness=1, highlightbackground="#999")
        self.canvas.pack(pady=(6, 0))

    def select_format(self, _event=None, initial=False):
        if not initial:
            index = self.format_box.current()
            self.vars["format"].set(FORMATS[index].code if index < len(FORMATS) else "Eigenes Format")
        layout = FORMATS_BY_CODE.get(self.vars["format"].get())
        self.format_name.set(layout.description if layout else "Eigenes Format")
        if layout:
            for key in ("width", "height", "columns", "rows", "left", "top", "gap_x", "gap_y"):
                self.vars[key].set(f"{getattr(layout, key):g}".replace(".", ","))
        for entry in self.geometry_entries:
            entry.configure(state="readonly" if layout else "normal")
        if not initial:
            self.vars["start"].set("1")
            self.page_index = 0
            self.schedule_refresh()

    def open_excel(self):
        path = filedialog.askopenfilename(parent=self, title="Excel-Datei mit Etiketten in Spalte A",
            filetypes=[("Excel-Dateien", "*.xlsx *.xlsm *.xls"), ("Alle Dateien", "*.*")])
        if not path:
            return
        self.loading = True
        self.workbook = {}
        self.sheet_name.set("")
        self.sheet_box.configure(values=[])
        self.file_label.configure(text=f"Wird eingelesen: {Path(path).name}")
        self.import_button.configure(state="disabled")
        self.refresh()

        def load():
            try:
                self.import_queue.put((path, read_excel(path), None))
            except Exception as exc:
                self.import_queue.put((path, None, exc))
        threading.Thread(target=load, daemon=True).start()
        self.after(100, self.finish_import)

    def finish_import(self):
        try:
            path, workbook, error = self.import_queue.get_nowait()
        except queue.Empty:
            self.after(100, self.finish_import)
            return
        self.loading = False
        self.import_button.configure(state="normal")
        if error:
            self.file_label.configure(text="Keine Datei geladen – Import fehlgeschlagen")
            messagebox.showerror("Excel-Import", f"Die Datei konnte nicht eingelesen werden:\n{error}", parent=self)
        else:
            self.workbook = workbook
            names = list(workbook)
            self.sheet_box.configure(values=names)
            self.sheet_name.set(names[0] if names else "")
            self.file_label.configure(text=Path(path).name)
        self.source_changed()

    def source_changed(self, _event=None):
        self.page_index = 0
        self.refresh()

    def schedule_refresh(self, *_args):
        if self.pending is not None:
            self.after_cancel(self.pending)
        self.pending = self.after(180, self.refresh)

    def number(self, key, title, minimum, maximum, integer=False):
        try:
            value = float(self.vars[key].get().replace(",", "."))
            if not math.isfinite(value) or not minimum <= value <= maximum or (integer and value != int(value)):
                raise ValueError()
        except ValueError:
            raise ValueError(f"{title}: bitte {'eine ganze Zahl' if integer else 'eine Zahl'} von {minimum:g} bis {maximum:g} eingeben.") from None
        return int(value) if integer else value

    def current_layout(self):
        preset = FORMATS_BY_CODE.get(self.vars["format"].get())
        if preset:
            return preset
        values = {key: self.number(key, title, minimum, maximum, integer) for key, title, minimum, maximum, integer in [
            ("width", "Breite", 5, 210, False), ("height", "Höhe", 5, 297, False),
            ("columns", "Spalten", 1, 40, True), ("rows", "Reihen", 1, 59, True),
            ("left", "Rand links", 0, 210, False), ("top", "Rand oben", 0, 297, False),
            ("gap_x", "Abstand X", 0, 210, False), ("gap_y", "Abstand Y", 0, 297, False)]}
        return LabelFormat("Eigenes Format", **values).validate()

    def render_options(self):
        return {"font_mm": self.number("font_mm", "Schriftgröße", 1, 100),
                "off_x": self.number("off_x", "Versatz X", -30, 30),
                "off_y": self.number("off_y", "Versatz Y", -30, 30),
                "border": self.vars["border"].get()}

    def refresh(self):
        if self.pending is not None:
            self.after_cancel(self.pending)
            self.pending = None
        self.print_button.configure(state="disabled")
        self.test_button.configure(state="disabled")
        self.items, self.sheets = [], []
        try:
            self.layout = self.current_layout().validate()
            self.start_box.configure(to=self.layout.per_sheet)
            start = self.number("start", "Startposition", 1, self.layout.per_sheet, True)
            options = self.render_options()
            self.items = excel_items(self.workbook.get(self.sheet_name.get(), []),
                                    self.vars["skip_header"].get(), self.vars["keep_blanks"].get())
            self.sheets = paginate(self.items, start, self.layout.per_sheet)
            self.page_index = min(self.page_index, len(self.sheets) - 1)
            self.draw_preview(options)
            count = sum(item is not None for item in self.items)
            if count:
                blanks = len(self.items) - count
                self.info.configure(text=f"{count} Etiketten auf {len(self.sheets)} Bogen/Bögen."
                    + (f" {blanks} leere Position(en) aus Excel." if blanks else ""), foreground="#444")
                self.print_button.configure(state="normal")
            else:
                self.info.configure(text="Excel-Datei wird eingelesen…" if self.loading else
                    "Keine Etiketten geladen. Bitte eine Excel-Datei mit Texten in Spalte A wählen.", foreground="#555")
            if not self.loading:
                self.test_button.configure(state="normal")
        except (ValueError, OSError) as exc:
            self.canvas.delete("all")
            self.page_label.configure(text="Keine Vorschau")
            self.info.configure(text=str(exc), foreground="#a32121")
            self.prev_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
            self.sheets = []
        if hasattr(self.app, "free_tab"):
            self.app.save_settings()

    def draw_preview(self, options=None):
        if not self.sheets:
            return
        image = render_sheet(self.sheets[self.page_index], self.layout, **(options or self.render_options()))
        scale = self.app.SCALE
        image = image.resize((round(210 * scale), round(297 * scale)), Image.Resampling.LANCZOS)
        self.preview_image = ImageTk.PhotoImage(image, master=self)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.preview_image, anchor="nw")
        options = options or self.render_options()
        for pos in range(self.layout.per_sheet):
            x, y, w, h = self.layout.rect(pos)
            x, y = x + options["off_x"], y + options["off_y"]
            self.canvas.create_rectangle(x * scale, y * scale, (x + w) * scale, (y + h) * scale,
                                         outline="#bbb", dash=(2, 3))
            if self.sheets[self.page_index][pos] is None:
                self.canvas.create_text((x + 1) * scale, (y + 1) * scale,
                    text=str(pos + 1), fill="#aaa", anchor="nw", font=("Arial", 7))
        self.page_label.configure(text=f"Seite {self.page_index + 1}/{len(self.sheets)}")
        self.prev_button.configure(state="normal" if self.page_index else "disabled")
        self.next_button.configure(state="normal" if self.page_index + 1 < len(self.sheets) else "disabled")

    def turn_page(self, step):
        if self.sheets:
            self.page_index = max(0, min(len(self.sheets) - 1, self.page_index + step))
            self.draw_preview()

    def do_print(self, test=False):
        self.refresh()  # apply edits even if the debounce timer has not fired yet
        if self.loading or not self.sheets or (not test and not any(self.items)):
            return
        printer = self.app.vars["printer"].get()
        if not printer:
            messagebox.showerror("Drucken", "Bitte einen Drucker auswählen.", parent=self)
            return
        try:
            options = self.render_options()
            options["frames"] = test
            sheets = self.sheets[:1] if test else self.sheets
            print_excel_sheets(printer, sheets, self.layout, options,
                               devmode=self.app.devmode_for_print(printer))
            messagebox.showinfo("Drucken", f"{len(sheets)} Bogen/Bögen an „{printer}“ gesendet.", parent=self)
        except Exception as exc:
            messagebox.showerror("Druckfehler", str(exc), parent=self)
