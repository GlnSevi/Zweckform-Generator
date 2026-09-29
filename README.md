# Lager-Label – Regal-Etiketten & Regalschilder

Druckt Lagerort-Etiketten auf **Avery Zweckform L4761-100** (A4, 4 Etiketten je
192 × 61 mm pro Bogen) im Stil der bestehenden Regal-Beschriftung (z. B. `K 14 09`) –
und erstellt zusätzlich große **Regalschilder für den 3D-Drucker**.

## Programm starten

**`Lager-Etiketten.exe`** starten – keine Installation nötig.
(Download beim jeweils neuesten [Release](https://github.com/GlnSevi/Lager-Label/releases)
oder lokal unter `dist\`.)

Zum Anheften an die Taskleiste: Programm starten, dann Rechtsklick auf das
Taskleisten-Symbol → **„An Taskleiste anheften“**.

## Reiter „Etiketten (Papier)“

Oben umschaltbar: **Serie (Buchstabe + Zahlen)** oder **Freitext**.

### Serie

- **Buchstabe(n)** – z. B. `K` oder `K,A` (mehrere mit Komma)
- **1. Zahl von–bis** – z. B. Feld 1 bis 20
- **2. Zahl von–bis** – z. B. Ebene 9 bis 9 (fest) oder ein Bereich
- **Reihenfolge** wählbar: Buchstaben zusammen (`K 01, A 01, K 02 …`) oder
  nacheinander (erst alle K, dann alle A)
- **Barcode (Code 128)** der Nummer ohne Leerzeichen (z. B. `K1409`) unter den
  Zahlen – wahlweise nur bei bestimmten Buchstaben (Feld „nur bei“, leer = alle)
- **Führende Nullen** einstellbar (z. B. `01` statt `1`)

### Freitext

- Beliebige Texte drucken – **je Zeile ein Etikett**, mittig auf dem Etikett;
  zu lange Texte werden automatisch passend verkleinert
- **Anzahl je Text** legt fest, wie oft jede Zeile gedruckt wird (1–999);
  eine abweichende Anzahl je Zeile geht mit senkrechtem Strich,
  z. B. `Reserviert Werkstatt | 8`
- Optional **Barcode (Code 128)** mit dem Text unter dem Etikett
  (bei Umlauten/Sonderzeichen nicht möglich – das Programm weist darauf hin)

### Für beide Modi

- **Vorschau** rechts zeigt jeden Bogen genau so, wie er gedruckt wird
- **Drucker** direkt auswählen und mit **🖨 Drucken** drucken
- **Druckereinstellungen…** öffnet den Dialog des Druckers (z. B. **Papierfach**
  wählen) – die Auswahl wird gemerkt und beim Druck verwendet
- **Umrandung** wird als Rahmen mit auf das Etikett gedruckt (abschaltbar)
- **Start-Position** auf einem angebrochenen Bogen wählbar (Etikett 1–4)
- **Schriftgröße** und **Feinjustierung X/Y** in mm einstellbar
- **Testbogen**-Knopf: druckt die erste Seite mit Schnittrahmen auf Normalpapier –
  gegen einen Etikettenbogen halten und bei Versatz die Feinjustierung anpassen

## Reiter „Regalschilder (3D-Druck)“

Große Schilder (Standard **300 × 300 mm**) zum Markieren der Regale von der
Seite – als **STL-Datei** für den 3D-Drucker:

- **Buchstabe davor (optional)** + **Zahl von–bis** – je Zahl entsteht ein
  Schild (z. B. `K01` … `K20`); bei mehreren Schildern wird ein Ordner gewählt
  und jede Datei einzeln abgelegt (`Regalschild_K01.stl` …)
- Die Zahl wird **vertieft in die Platte eingelassen** (Standard **0,6 mm**),
  damit sie sich sauber **einfärben** lässt: entweder ausmalen (Lackstift)
  oder beim Druck das **Filament wechseln**, sobald der Boden der Vertiefung
  erreicht ist – die passende Höhe zeigt das Programm nach dem Erstellen an
- **Rahmen mit einfräsen**: umlaufender Rahmen im Stil der Papier-Etiketten
- **Abgerundete Ecken**: Radius frei einstellbar (Standard 20 mm, 0 = eckig)
- **Lochung oben** (abschaltbar): 2 Löcher (Ø 6 mm) in den oberen Ecken zum
  Anschrauben oder Aufhängen
- Plattengröße, Dicke, Schrifthöhe und Tiefe der Vertiefung frei einstellbar
- **Vorschau** rechts zeigt jedes Schild vorab (gelb mit schwarzer Schrift,
  wie das fertige Schild)
- Hinweis: Vertiefungen **unter ca. 0,2 mm** (eine Druckschicht) sind im
  fertigen Druck nicht sichtbar – das Programm warnt in dem Fall
- Die STL-Datei einfach im Slicer öffnen (z. B. **OrcaSlicer**) und drucken

## Reiter „Frei / Excel“

1. **Excel-Datei öffnen…** und das gewünschte **Arbeitsblatt** wählen
   (`.xlsx`, `.xlsm` oder `.xls`; Excel muss nicht installiert sein).
2. **Zweckform-Format** auswählen: L4761, L7159–L7169 oder 3666.
   Maße und Anzahl pro Bogen stehen direkt in der Auswahl.
3. Die **Vorschau** kontrollieren, bei Bedarf zunächst einen **Testbogen**
   auf Normalpapier drucken und anschließend **Drucken** wählen.

Jede Zelle in **Spalte A** ergibt genau ein Etikett. Die Reihenfolge ist
**links nach rechts, dann die nächste Reihe von oben nach unten**.
Ist der Bogen voll, geht es auf dem nächsten weiter. Weitere Excel-Spalten
werden ignoriert; die Datei wird nicht verändert.

- **Erste Zeile ist eine Überschrift**: auf Wunsch A1 überspringen;
  standardmäßig wird schon A1 als Etikett übernommen.
- **Leere Zeilen** bleiben standardmäßig als freie Etikettenpositionen erhalten.
  Abschalten, um nur gefüllte Zellen lückenlos zu verteilen. Leere Zellen nach
  dem letzten Inhalt erzeugen keine zusätzlichen Bögen.
- **Zeilenumbrüche innerhalb einer Zelle** bleiben innerhalb desselben Etiketts.
  Lange Texte werden umgebrochen und bei Bedarf passend verkleinert.
- Text wie `00123` bleibt erhalten. Einfache Zahlenformate wie `00000`
  werden ebenfalls berücksichtigt. Für Artikelnummern oder besondere
  Darstellungen (z. B. Währung, Prozent oder komplexe Excel-Zahlenformate)
  die Werte in Spalte A als **Text** ablegen.
- Excel-Formeln verwenden das zuletzt in Excel gespeicherte Ergebnis.
  Fehlt dieses oder enthält eine Zelle einen Excel-Fehler, nennt die App
  die betroffene Zelle. Dann in Excel neu berechnen und speichern oder
  die Formeln durch Werte ersetzen. Makros werden nicht ausgeführt.
- **Start bei Etikett** nutzt angebrochene Bögen weiter; **Schriftgröße**,
  **Umrandung** und **Feinjustierung X/Y** lassen sich unabhängig einstellen.
- **Eigenes Format** erlaubt weitere rechteckige A4-Bögen: Breite/Höhe,
  Spalten/Reihen, Ränder und Abstände in mm. Die zuvor ausgewählte Vorlage
  dient als Ausgangspunkt. Ungültige Raster werden vor dem Druck gemeldet.
- Der Druck erfolgt in **A4-Hochformat, in Originalgröße**. Vorschau und
  Druck verwenden dieselbe Darstellung. Die grauen Hilfslinien und Nummern
  freier Plätze in der Vorschau werden nur beim Testbogen als Rahmen gedruckt;
  Platznummern werden nie gedruckt.

Die Vorlagenmaße und Quellen stehen in [FORMAT_SOURCES.md](FORMAT_SOURCES.md).
Pro Datei sind bis zu 50.000 eingelesene Zeilen und pro Zelle bis zu
4.000 Zeichen vorgesehen. XLSM wird ohne Ausführung von Makros gelesen.

Eingaben und Einstellungen werden automatisch gespeichert und beim
nächsten Start wiederhergestellt. Im Reiter „Frei / Excel“ werden nur
die Druck-/Importoptionen gespeichert; die Excel-Datei wird für einen neuen
Auftrag erneut geöffnet.

## Neu bauen (für Entwickler)

```powershell
pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
python -m PyInstaller --noconfirm Lager-Etiketten.spec
```

Die fertige EXE liegt danach unter `dist\Lager-Etiketten.exe`.
