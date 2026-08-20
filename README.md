# Lager-Label – Regal-Etiketten & Regalschilder

Erzeugt und druckt Lagerort-Etiketten im Format **Avery Zweckform L4761-100**
(A4, 4 Etiketten je 192 × 61 mm pro Bogen), im Stil der bestehenden Regal-Beschriftung
(z. B. `K 14 09`) – und zusätzlich große **Regalschilder als STL-Datei für den
3D-Drucker** (Reiter „Regalschilder (3D-Druck)“).

## Programm (Windows-App)

**`dist\Lager-Etiketten.exe`** starten – keine Installation nötig.

### Reiter „Etiketten (Papier)“

- **Buchstabe(n)** – z. B. `K` oder `K,A` (mehrere mit Komma)
- **1. Zahl von–bis** – z. B. Feld 1 bis 20
- **2. Zahl von–bis** – z. B. Ebene 9 bis 9 (fest) oder ein Bereich
- **Drucker** direkt auswählen und mit **🖨 Drucken** ohne Umweg drucken
- **Druckereinstellungen…**: öffnet den Windows-Treiberdialog des Druckers
  (z. B. **Papierfach** wählen) – die Auswahl wird gemerkt und beim Druck verwendet
- **Umrandung** wird als Rahmen mit auf das Etikett gedruckt (abschaltbar)
- **Barcode (Code 128)** der Nummer ohne Leerzeichen (z. B. `K1409`) unter den Zahlen –
  standardmäßig nur bei Buchstabe **K** (Feld „nur bei“, leer = alle)
- Außerdem: Reihenfolge, Start-Position auf angebrochenem Bogen,
  führende Nullen, Schriftgröße, Feinjustierung X/Y in mm
- **Testbogen**-Knopf: druckt die erste Seite mit Schnittrahmen zum Prüfen auf Normalpapier

### Reiter „Regalschilder (3D-Druck)“

Große Schilder (Standard **300 × 300 mm**) zum Markieren der Regale von der
Seite – als **STL-Datei** für den 3D-Drucker:

- **Buchstabe davor (optional)** + **Zahl von–bis** – je Zahl entsteht ein Schild
  (z. B. `K01` … `K20`); bei mehreren Schildern wird ein Ordner gewählt und
  jede Datei einzeln abgelegt (`Regalschild_K01.stl` …)
- Die Zahl wird **vertieft in die Platte eingelassen** (Standard **0,6 mm**
  = 3 Druckschichten à 0,2 mm), damit sie sich sauber **einfärben** lässt:
  entweder ausmalen (Lackstift) oder beim Druck einfach das **Filament
  wechseln**, sobald der Boden der Vertiefung erreicht ist – die Höhe dafür
  zeigt das Programm nach dem Erstellen an
- **Rahmen mit einfräsen**: umlaufender Rahmen im Stil der Papier-Etiketten –
  er folgt der Eckenrundung der Platte
- **Abgerundete Ecken**: Ecken-Radius frei einstellbar (Standard 20 mm, 0 = eckig)
- **Lochung oben** (abschaltbar): 2 Durchgangslöcher (Ø 6 mm) in den oberen
  Ecken zum Anschrauben oder Aufhängen – sie weichen Rundung und Rahmen
  automatisch aus
- Plattengröße, Dicke, Schrifthöhe und Tiefe der Vertiefung sind frei
  einstellbar; die Vorschau rechts zeigt jedes Schild vorab (gelb mit
  schwarzer Schrift, wie das fertige Schild)
- Die STL-Netze entstehen per Boolescher Volumen-Verrechnung (**manifold3d**)
  und sind dadurch wasserdicht – alle Konturen geschlossen, die Vertiefung
  lässt sich sauber einfärben
- Wichtig: Vertiefungen **unter ca. 0,2 mm** (eine Druckschicht) sind im
  fertigen Druck nicht sichtbar – das Programm warnt in dem Fall
- Die STL-Datei einfach im Slicer öffnen (z. B. **OrcaSlicer**) und drucken –
  es werden keine Zusatzprogramme benötigt

Alle Einstellungen werden automatisch gespeichert
(`%APPDATA%\LagerLabel\settings.json`).

### An die Taskleiste anheften

1. `dist\Lager-Etiketten.exe` starten
2. Rechtsklick auf das Symbol in der Taskleiste → **„An Taskleiste anheften“**

(Oder die EXE z. B. nach `C:\Programme\Lager-Etiketten\` kopieren und von dort anheften –
dann funktioniert das Anheften auch nach dem Verschieben des Projektordners noch.)

## Dateien

| Datei | Zweck |
| --- | --- |
| `dist\Lager-Etiketten.exe` | Fertiges Programm (an Taskleiste anheftbar) |
| `lager_label.py` | Python-Quellcode (tkinter-GUI, Druck über pywin32/GDI, STL-Erzeugung) |
| `etiketten-generator.html` | Browser-Variante (gleiches Layout, Druck über Browser) |
| `icon.ico` | Programm-Icon |

## Neu bauen (nach Code-Änderungen)

```powershell
pip install pywin32 pillow pyinstaller manifold3d numpy
python -m PyInstaller --noconfirm --onefile --windowed --name "Lager-Etiketten" --icon icon.ico --add-data "icon.ico;." lager_label.py
```

## Tipp zum Drucken

Der Druck geht direkt über Windows-GDI, millimetergenau positioniert – ohne
Browser-Skalierung. Bei leichtem Versatz des Druckers die **Feinjustierung X/Y**
(in mm) anpassen. Erst einen **Testbogen** auf Normalpapier drucken und gegen
einen Etikettenbogen halten.
