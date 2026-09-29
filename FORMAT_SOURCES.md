# Zweckform-Vorlagen im Reiter „Frei / Excel“

Alle Maße sind in mm. Nominale Etikettenmaße und Stückzahlen wurden am
28.09.2026 mit den Avery-Zweckform-Vorlagenseiten abgeglichen:

| Vorlage | Etikett | Spalten × Reihen | Hersteller |
| --- | --- | --- | --- |
| L4761 | 192 × 61 | 1 × 4 | [Avery](https://www.avery-zweckform.com/vorlage-l4761) |
| L7159 | 63,5 × 33,9 | 3 × 8 | [Avery](https://www.avery-zweckform.com/vorlage-l7159) |
| L7160 | 63,5 × 38,1 | 3 × 7 | [Avery](https://www.avery-zweckform.com/vorlage-l7160) |
| L7161 | 63,5 × 46,6 | 3 × 6 | [Avery](https://www.avery-zweckform.com/vorlage-l7161) |
| L7162 | 99,1 × 33,9 | 2 × 8 | [Avery](https://www.avery-zweckform.com/vorlage-l7162) |
| L7163 | 99,1 × 38,1 | 2 × 7 | [Avery](https://www.avery-zweckform.com/vorlage-l7163) |
| L7164 | 63,5 × 72 | 3 × 4 | [Avery](https://www.avery-zweckform.com/vorlage-l7164) |
| L7165 | 99,1 × 67,7 | 2 × 4 | [Avery](https://www.avery-zweckform.com/vorlage-l7165) |
| L7166 | 99,1 × 93,1 | 2 × 3 | [Avery](https://www.avery-zweckform.com/vorlage-l7166) |
| L7167 | 199,6 × 289,1 | 1 × 1 | [Avery](https://www.avery-zweckform.com/vorlage-l7167) |
| L7168 | 199,6 × 143,5 | 1 × 2 | [Avery](https://www.avery-zweckform.com/vorlage-l7168) |
| L7169 | 99,1 × 139 | 2 × 2 | [Avery](https://www.avery-zweckform.com/vorlage-l7169) |
| 3666 | 38 × 21,2 | 5 × 13 | [Avery](https://www.avery-zweckform.com/vorlage-3666) |
| FX1574 (-W / -GL) | 90 × 20 | 2 × 14 | [Flexicom Format74](https://shop.flexicom.de/media/Formate/Format74.pdf) |

Als zusätzlicher Abgleich der Raster dienen die Angaben im
[LibreOffice-Etikettenkatalog](https://github.com/LibreOffice/core/blob/master/extras/source/labels/labels.xml).
Die App verwendet die aktuellen nominalen Herstellermaße und zentrierte Raster
auf A4, bei mehrspaltigen L71xx-Bögen mit 2,54 mm horizontalem Zwischenraum.
Alte Katalogangaben mit abweichenden/gerundeten Maßen werden nicht ungeprüft
übernommen (z. B. 64 mm bei L7159 oder Letter-Seitenmaße bei L7169).
L4761 behält das Raster der bestehenden App.

## Vollständiger Katalog (`zweckform_formats.json`)

Zusätzlich zu den oben geprüften Vorlagen enthält die App alle weiteren
A4-Bögen aus
- dem LibreOffice-Katalog (Hersteller „Avery Zweckform“, „Avery A4“,
  „Avery A4/Asia“) und
- den [gLabels-Vorlagen](https://github.com/j-evins/glabels-qt/blob/master/templates/avery-iso-templates.xml)
  (Avery A4, inkl. runder/CD-Etiketten).

Die Datei wird mit `python tools/build_formats.py` aus den Kopien in
`tools/sources/` erzeugt. Übernommen werden A4-Hoch- und Querformat;
Endlos-, Letter- und A5-Formate entfallen, ebenso Bögen mit gemischten
Etikettengrößen und die wenigen Katalogeinträge, deren Raster nicht auf A4
passt (das Skript listet sie auf). Bei doppelter Artikelnummer gewinnt die
deutsche Zweckform-Angabe; die geprüften Vorlagen haben immer Vorrang.
Varianten mit gleicher Nummer, aber anderem Raster erhalten eine Endung
(z. B. `C2050-2`).

Die Katalogmaße stammen nicht direkt vom Hersteller. Vor größeren
Druckaufträgen daher immer zuerst den **Testbogen** verwenden.

Die voreingestellten Maße sind im Reiter sichtbar. Bei abweichendem Bogen:
Vorlage auswählen, danach „Eigenes Format“ wählen und Maße korrigieren.
Ein Testbogen auf Normalpapier dient zur Kontrolle am tatsächlich verwendeten
Etikettenbogen. Druckereinzug/Versatz lassen sich über X/Y ausgleichen.
