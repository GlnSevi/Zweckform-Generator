# Lager-Label – Regal-Etiketten Generator

Erzeugt und druckt Lagerort-Etiketten im Format **Avery Zweckform L4761-100**
(A4, 4 Etiketten je 192 × 61 mm pro Bogen), im Stil der bestehenden Regal-Beschriftung
(z. B. `K 14 09`).

## Programm (Windows-App)

**`dist\Lager-Etiketten.exe`** starten – keine Installation nötig.

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
| `lager_label.py` | Python-Quellcode (tkinter-GUI, Druck über pywin32/GDI) |
| `etiketten-generator.html` | Browser-Variante (gleiches Layout, Druck über Browser) |
| `icon.ico` | Programm-Icon |

## Neu bauen (nach Code-Änderungen)

```powershell
pip install pywin32 pillow pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name "Lager-Etiketten" --icon icon.ico --add-data "icon.ico;." lager_label.py
```

## Tipp zum Drucken

Der Druck geht direkt über Windows-GDI, millimetergenau positioniert – ohne
Browser-Skalierung. Bei leichtem Versatz des Druckers die **Feinjustierung X/Y**
(in mm) anpassen. Erst einen **Testbogen** auf Normalpapier drucken und gegen
einen Etikettenbogen halten.
