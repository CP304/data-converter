# Einkauf Data Converter

Werkzeugkasten für die tägliche Einkaufsarbeit: Lieferantendaten konvertieren,
aufbereiten, umbenennen, packen und dokumentieren — **komplett mit
Python-Bordmitteln**. Keine Installationen, keine APIs, keine Abhängigkeiten
(auch kein Pillow/LibreOffice/FreeCAD).

## Start

```powershell
python data_converter_gui.py
```

Oder unter Windows per Doppelklick: `start_data_converter.bat`
(prüft `.venv`, `python` und `py`).

## Bedienlogik

Jedes Werkzeug ist ein Tab und folgt demselben Fluss:

1. **Quelle** wählen (Ordner oder mehrere Dateien) und filtern
   (Dateitypen, Namensbestandteil, max. Größe, Unterordner ja/nein).
2. **Einstellungen** setzen, Zielordner wählen (Unterordner neben der Quelle
   oder eigener Ordner).
3. **Vorschau** klicken → der komplette Plan (Quelle → Ergebnis) erscheint als
   Tabelle. Es wird noch nichts verändert.
4. **Ausführen** → Fortschrittsbalken, Live-Log unten im Fenster, Status pro
   Zeile. Zusätzlich wird eine Logdatei `lauf_log_…txt` im Zielordner abgelegt.
   **Stop** bricht sauber ab.

Bestehende Dateien werden nie überschrieben (automatisch `_2`, `_3`, …).
Wiederkehrende Aufgaben lassen sich über *Datei → Preset speichern/laden* als
JSON ablegen.

## Die Werkzeuge

### Tabellen (das Herzstück)
Preislisten, Stücklisten und Lieferantenlisten wandeln und aufbereiten.

- **Lesen:** CSV, TSV, TXT (Trennzeichen und Encoding werden automatisch
  erkannt, auch cp1252-Altdaten), XLSX, JSON, XML
- **Schreiben:** CSV, TSV, **XLSX** (eigener Writer auf zipfile-Basis, öffnet
  sauber in Excel, Kopfzeile fett, Spaltenbreiten angepasst), JSON, XML,
  HTML-Bericht, Markdown
- **Aufbereitung:** Werte trimmen, Leerzeilen/Duplikate entfernen,
  Dezimalzeichen normalisieren (`1.234,50` ↔ `1234.50`), Spalten auswählen und
  umbenennen (`Alt>Neu; SKU`), mehrere Dateien zu einer Tabelle
  zusammenführen (mit Spalte „Quelle")

### CAD
Neutrale 3D-Formate konvertieren — mit **eigenem B-Rep-Tessellierungs-Kern**
für STEP (Part-21-Parser, Flächen-Evaluatoren für Ebene/Zylinder/Kegel/Kugel/
Torus/NURBS/Extrusion/Rotation, Trimmung im UV-Raum, Ear-Clipping mit
Verfeinerung — alles Standardbibliothek).

- **Lesen:** STL (ASCII/binär), OBJ, PLY (ASCII/binär), 3MF, **STEP** (AP203/214/242)
- **Schreiben:** STL, OBJ, PLY, 3MF, GLB und eine **eigenständige HTML-3D-Ansicht**
  (eigener WebGL-Viewer, läuft offline in jedem Browser — ideal zum Weiterleiten
  an Kollegen ohne CAD-Lizenz)
- **STEP/IGES-Prüfbericht:** Schema, Produkte, Einheiten, Autor, Entitäten-Statistik,
  Geometrieart — als CSV/HTML für die Wareneingangsprüfung von CAD-Daten
- Qualitätsstufen grob/mittel/fein für die STEP-Tessellierung

Grenzen (ehrlich): Native Formate wie SLDPRT/IPT/eDrawings enthalten
proprietäre, undokumentierte Kernel-Daten (Parasolid/ShapeManager) — die kann
ohne Hersteller-Kernel niemand lesen. Praxis-Tipp: beim Lieferanten STEP oder
STL anfordern; eDrawings selbst kann Teile als STL speichern. STEP-Tessellierung
liefert ein Sichtmodell/Fertigungs-Mesh, kein exaktes B-Rep;
Baugruppen-Transformationen werden nicht angewendet.

### Umbenennen
Suchen/Ersetzen (optional Regex), Präfix/Suffix, Nummerierung,
Groß-/Kleinschreibung, Änderungsdatum voranstellen. Vorher/Nachher in der
Vorschau, Journal-Datei je Lauf und **Rückgängig**-Button.

### Packen
Jede Datei einzeln zippen, alles als ZIP-/TAR.GZ-Lieferantenpaket bündeln oder
Archive entpacken (ZIP, TAR, TGZ, GZ, BZ2, XZ) — mit Schutz gegen unsichere
Pfade und optionalem Glätten der Ordnerstruktur.

### Ordnen
Downloads und Lieferantenordner aufräumen: nach Dateityp oder Datum (JJJJ-MM)
in Unterordner sortieren, verstreute Dateien flach in einen Zielordner
zusammenführen, leere Unterordner entfernen — wahlweise kopieren statt
verschieben, mit Journal und **Rückgängig**.

### Inventar
Lieferantenpaket dokumentieren: Manifest mit Pfad, Typ, Größe, Änderungsdatum,
SHA-256 und **Duplikat-Erkennung** — als CSV, HTML-Bericht und/oder XLSX.
Optional eine `SHA256SUMS.txt` im Standardformat für die Übergabe.

### E-Mail
Anhänge aus `.eml`-Dateien (Export aus Outlook/Thunderbird) gesammelt
extrahieren (optional ein Unterordner pro Mail) — oder Mails komplett
archivieren: **als HTML** (eingebettete Bilder werden als Data-URIs
mitgenommen, die Datei ist allein lebensfähig) oder als lesbare Textdatei.

### Text/Encoding
Encoding (UTF-8, UTF-8-BOM, cp1252, latin-1) und Zeilenenden (CRLF/LF) für
ERP-Importe vereinheitlichen. Quell-Encoding wird automatisch erkannt.

## Bewusst nicht enthalten

Diese Version verspricht nur, was sie ohne Installationen hält. Nicht enthalten
sind deshalb: native CAD-Formate (SLDPRT/IPT/eDrawings — proprietäre
Kernel-Daten), DWG/DXF-Geometrie, Office→PDF, Bildbearbeitung,
RAR/7z-Entpacken, Medien-Konvertierung.

## Tests

```powershell
python test_converter_backend.py
```

```powershell
python data_converter_gui.py --self-test
```

Beispieldaten für manuelle Tests liegen unter `mock_files/`.

## Projektstruktur

- `data_converter_gui.py` — Einstiegspunkt
- `converter_app/gui.py` — Oberfläche (tkinter/ttk)
- `converter_app/tabular.py` — Tabellen lesen/aufbereiten/schreiben
- `converter_app/xlsx_io.py` — XLSX-Reader/-Writer (zipfile + ElementTree)
- `converter_app/cad_io.py` — STL/OBJ/PLY/3MF/GLB lesen & schreiben, HTML-3D-Viewer, STEP/IGES-Bericht
- `converter_app/step_mesh.py` — STEP-B-Rep-Tessellierungs-Kern
- `converter_app/filetools.py` — Umbenennen, Archive, Ordnen, Inventar, EML, Encoding
- `converter_app/selftest.py` — Kurz-Selbsttest ohne GUI
