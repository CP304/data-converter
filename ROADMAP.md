# Roadmap

Die Vision bleibt bewusst schlank: **sehr praktisches Dateihandling in einem
Tool** — konvertieren, umbenennen, packen, ordnen, prüfen. Keine Auswertungen,
keine Datenbank, kein Dashboard. Reine Python-Standardbibliothek, keine
Installationen.

Sinnvolle nächste Schritte, alle innerhalb dieser Vision:

## Bedienung

- **Zuletzt verwendete Quellen/Ziele** je Tab merken (schneller Wiedereinstieg)
- **Kontextmenü im Vorschau-Tree**: Datei öffnen, Ordner öffnen, Zeile aus dem
  Plan nehmen
- **Presets als Schnellzugriff** in der Kopfzeile (die drei Wochen-Standardläufe
  mit einem Klick)

## Automatisierung (Dateihandling ohne Klicks)

- **Headless-Lauf**: `python data_converter_gui.py --preset lauf.json --run`
  ohne Fenster → planbar über den Windows-Taskplaner
- **Watch-Ordner**: Eingangsordner beobachten, Preset automatisch anwenden
  (Lieferantenpaket kommt an → entpacken → ordnen → Inventar)
- **Werkzeug-Ketten**: mehrere Tabs als eine Kette speichern und in einem
  Rutsch ausführen

## Formate abrunden

- **IGES → Mesh** im Tessellierungs-Kern (Flächen-Evaluatoren existieren schon)
- **PDF**: Seiten drehen und Reihenfolge ändern
- **Bilder**: Zuschneiden (feste Ränder/Zone)
- **Packen**: ZIP in Teile fester Größe splitten (E-Mail-Anhang-Limits)

## Bewusste Grenzen (bleiben)

- Kein SLDPRT/IPT/eDrawings (proprietäre Kernel-Daten), kein DWG
- Kein Office→PDF/A, kein JPEG/RAW, kein RAR/7z
- Keine Preisanalysen, keine Datenbank, kein Dashboard — dafür gibt es Excel & ERP
