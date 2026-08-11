# Roadmap

Die Vision bleibt bewusst schlank: **sehr praktisches Dateihandling in einem
Tool** — konvertieren, umbenennen, packen, ordnen, prüfen. Keine Auswertungen,
keine Datenbank, kein Dashboard. Reine Python-Standardbibliothek, keine
Installationen.

## Umgesetzt (Stand August 2026)

### Bedienung
- ✅ **Sitzung merken**: Quellen, Ziele und Einstellungen aller Tabs werden beim
  Schließen gespeichert (`~/.einkauf_data_converter.json`) und beim Start
  wiederhergestellt.
- ✅ **Kontextmenü im Vorschau-Tree** (Rechtsklick): Datei öffnen, im Explorer
  zeigen, Zeile aus dem Plan nehmen — „Ausführen" arbeitet dann den
  editierten Plan ab.
- ✅ **Preset-Schnellzugriff**: „Presets ▾" in der Kopfzeile mit den zuletzt
  verwendeten Presets.

### Automatisierung
- ✅ **Headless-Lauf** für den Windows-Taskplaner:
  `python data_converter_gui.py --preset lauf.json --run`
- ✅ **Watch-Ordner**: `--preset lauf.json --watch 30` beobachtet die Quelle des
  ersten Ketten-Werkzeugs und führt die Kette aus, sobald neue Dateien stabil
  angekommen sind.
- ✅ **Werkzeug-Ketten**: *Datei → Kette ausführen…* führt mehrere Tabs
  nacheinander aus; die Kette wird in Presets mitgespeichert.

### Formate
- ✅ **IGES → Mesh** im Tessellierungs-Kern (NURBS-Flächen 128, Rotation 120,
  Extrusion 122, getrimmte Flächen 144/142, Transformationen 124).
- ✅ **PDF**: Seiten drehen (90/180/270) und Reihenfolge ändern (`3,1-2`).
- ✅ **Bilder**: Zuschneiden per Zone (% von links oben).
- ✅ **Packen**: Dateien auf mehrere unabhängige ZIPs mit Maximalgröße verteilen
  (E-Mail-Anhang-Limits).

## Bewusste Grenzen (bleiben)

- Kein SLDPRT/IPT/eDrawings (proprietäre Kernel-Daten), kein DWG
- Kein Office→PDF/A, kein JPEG/RAW, kein RAR/7z
- Keine Preisanalysen, keine Datenbank, kein Dashboard — dafür gibt es Excel & ERP
