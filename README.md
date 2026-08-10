# Einkauf Data Converter

Kleines Python-GUI-Tool fuer Einkaeufer, um Dateien aus Ordnern oder Einzeldateien gesammelt zu konvertieren.

## Start

```powershell
python data_converter_gui.py
```

Oder unter Windows:

```powershell
.\start_data_converter.bat
```

Falls Windows `python` nicht kennt, Python installieren oder im Projekt eine `.venv` anlegen. Der Starter prueft `.venv`, `python` und `py`.

## Bedienlogik

- Jede Konvertierungsregel ist eine kompakte Tabellenzeile.
- Pfad als Ordner oder Einzeldatei waehlen.
- Inputformate ueber `Input` gruppiert nach CAD, Tabellen, Office, Dokumenten, Bildern, Print, PDF, Archiven und Medien ausklappen und markieren.
- Optional nach Namensbestandteil oder maximaler MB-Groesse filtern.
- Sinnvolle Outputformate ueber `Output` anhaken. Die Liste richtet sich nach den Inputs.
- Ueber `Zahnrad` Spezialoperationen und Namenskonventionen aktivieren.
- Mit `+ Zeile` weitere Konvertierungsregeln hinzufuegen.
- Unten Output im selben Ordner und rekursive Suche steuern.
- `Go` startet, `Stop` fordert Abbruch an.
- Logs werden nicht in der GUI angezeigt, sondern als `conversion_log_YYYYMMDD_HHMMSS.txt` im Outputordner gespeichert.

## Direkt nutzbar

Ohne Zusatzsoftware sind im Prototyp vor allem diese Funktionen robust:

- CSV/TSV/JSON untereinander konvertieren.
- Dateien in ZIP packen.
- Gleiche Dateiendung kopieren.

Mit `Pillow`:

- Bilder zwischen JPG, PNG, TIFF, WebP und PDF konvertieren.
- Bilder komprimieren.
- Demo-Zone weiss ueberdecken.
- Outputnamen per Suchen/Ersetzen, Praefix und Suffix vereinheitlichen.
- Optional auch Inputdateien nach derselben Namensregel umbenennen.

Installation:

```powershell
pip install pillow
```

Mit LibreOffice im PATH:

- DOC/DOCX/XLS/XLSX/PPT/PPTX/ODT/ODS/ODP nach PDF oder moderne Office-Formate konvertieren.

Mit FreeCADCmd im PATH:

- Einige 3D-CAD-Formate nach STEP/STP/STL/OBJ/IFC exportieren, soweit FreeCAD den Import unterstuetzt.

## Formatgruppen

- CAD 3D Native: SolidWorks, CATIA, Inventor, NX/Creo-nahe Dateitypen und weitere native CAD-Endungen.
- CAD 3D Neutral: STEP/STP, IGES/IGS, JT, Parasolid, SAT/SAB, STL, OBJ, 3MF, IFC, GLB/GLTF.
- CAD 2D: DWG, DXF, DWF/DWFX, DGN, HPGL/PLT.
- CAM/NC: NC, CNC, G-Code, TAP, DRL.
- Elektronik: Gerber, KiCad PCB, BRD, SCH, DSN.
- Tabellen/ERP: CSV, TSV, JSON, XML, YAML, XLS/XLSX/XLSM/XLSB, ODS, MDB/ACCDB, SQLite/DB, Parquet.
- MS Office: Word, Excel, PowerPoint, Visio und OpenDocument-Formate.
- Dokumente/Text: TXT, MD, HTML, XHTML, XML, EML, MSG, TEX.
- Bilder: JPG, PNG, TIFF, WebP, HEIC/HEIF, AVIF, BMP, GIF, PSD, RAW-nahe Formate.
- Vektor/Print: SVG, EPS, AI, CDR, PS, INDD.
- PDF: PDF in PDF, Bilder, Text oder ZIP.
- Archive: ZIP, 7Z, RAR, TAR, TGZ, GZ, BZ2, XZ.
- Medien: MP4, MOV, AVI, WMV, WebM, MP3, WAV.

Nicht jedes Format ist ohne Zusatzsoftware direkt konvertierbar. Die GUI zeigt bewusst den Einkaufsbedarf breit an; die eigentliche Umsetzung haengt je Format von Backends wie LibreOffice, FreeCAD, ODA, ImageMagick, Ghostscript, FFmpeg oder Hersteller-CAD ab.

## Tests

```powershell
python data_converter_gui.py --self-test
python test_converter_backend.py
```

Fuer manuelle GUI-Tests liegen einfache Mockdateien unter `mock_files/`. Die STEP/DXF-Dateien dort sind Platzhalter fuer UI-, Filter-, ZIP- und Kopierlaeufe, nicht fuer echte CAD-Geometriepruefung.

## Empfehlungen fuer den Einkauf

- `STEP`/`STP`: bester neutraler Standard fuer 3D-CAD-Austausch.
- `PDF`: Zeichnungen, Spezifikationen, Angebote, Freigaben.
- `XLSX`: editierbare Preislisten, Lieferantenlisten, Stuecklisten.
- `CSV`: robuste Datenuebergabe an ERP, PIM, BI und Skripte.
- `ZIP`: Lieferantenpakete gesammelt und nachvollziehbar ablegen.
- `PNG`/`JPG`/`WebP`: Bilder, Etiketten, Screenshots, Shopdaten.

## Was geht sonst noch?

Technisch sinnvoll erweiterbar sind:

- PDF schwaerzen/weissen nach frei definierter Zone.
- PDF splitten/mergen/komprimieren.
- Bilder skalieren, Wasserzeichen entfernen/setzen, DPI normalisieren.
- Archive entpacken, neu zippen, Struktur vereinheitlichen.
- CAD-Daten vereinheitlichen: native Formate nach STEP, Zeichnungen nach PDF/DXF.
- Tabellen normalisieren: Spaltenmapping, Duplikate, Zeichensatz, Dezimaltrennzeichen.
- Office-Pakete automatisch als PDF/A archivieren.

Native CAD-Formate wie SolidWorks, Inventor, CATIA, NX oder Creo sind lizenz- und kernelabhaengig. Fuer eine produktive Einkaufsversion empfiehlt sich ein klar gewaehltes Backend: Hersteller-CAD, FreeCAD fuer offene Formate, ODA fuer DWG/DXF oder ein kommerzieller Batch-Konverter.
