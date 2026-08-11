# Roadmap: Vom Konverter zum Einkaufs-Cockpit

Leitidee: Das Tool soll nicht nur Dateien **umwandeln**, sondern die tägliche
Einkaufsarbeit rund um Lieferantendaten **verstehen, automatisieren und
dokumentieren**. Alles bleibt bei der harten Vorgabe: reine
Python-Standardbibliothek, keine Installationen, keine APIs.

## Stufe 1 — Daten verstehen (nächste Ausbaustufe)

Der größte Hebel: nicht Formate wandeln, sondern **Inhalte auswerten**.

- **Preislisten-Diff**: alte vs. neue Preisliste über SKU abgleichen →
  Preisänderungen absolut/%, neue und entfallene Artikel, sortiert nach
  größtem Kostenhebel. Bericht als XLSX/HTML. (Tabellen-Engine ist da,
  fehlt nur die Diff-Logik + Tab.)
- **Angebotsvergleich**: 2–5 Lieferanten-Tabellen über SKU-Spalte mappen →
  Bestpreis-Matrix mit Ampel, Summenzeile pro Lieferant, Einsparpotenzial.
- **BOM-Abgleich**: Stückliste gegen einen Ordner prüfen — zu welcher
  Position fehlt die Zeichnung/das STEP? Ergebnis als Fehlliste für die
  Lieferantenanfrage.
- **PDF-Textextraktion**: eigener Content-Stream-Parser (Tj/TJ-Operatoren,
  Standard-Fonts) → Angebots-PDFs durchsuchbar machen, Preise/Positionen
  herausziehen. Grundlage: der vorhandene PDF-Parser in `pdf_io.py`.

## Stufe 2 — Automatisieren

- **CLI/Headless-Modus**: `python data_converter_gui.py --preset wochenlauf.json --run`
  → läuft ohne Fenster, Exit-Code + Log. Damit über den Windows-Taskplaner
  planbar (z. B. jeden Montag 6:00 Eingangsordner verarbeiten).
- **Watch-Ordner**: Eingangsordner überwachen (Polling, stdlib); neue Dateien
  lösen automatisch ein Preset aus (z. B. „Lieferantenpaket eingegangen →
  entpacken → Inventar → Ablage").
- **Workflow-Ketten**: mehrere Werkzeuge als eine Pipeline speichern
  (Entpacken → Ordnen → Tabellen normalisieren → Inventar) — ein Klick statt vier.

## Stufe 3 — Wissen behalten

- **Lauf-Historie in SQLite** (stdlib `sqlite3`): jeder Lauf mit Zeitpunkt,
  Quelle, Aktionen, Ergebnis — durchsuchbar („Wann habe ich das
  Müller-Paket verarbeitet?").
- **Teile-/Lieferanten-Verknüpfung light**: SKU ↔ Zeichnungsdatei ↔ Lieferant
  als kleine lokale Datenbank, gefüttert aus Inventar-Läufen und Preislisten.
- **Berichts-Dashboard**: eine HTML-Startseite, die alle erzeugten Berichte
  (Inventare, Preisdiffs, CAD-Prüfungen) einsammelt und verlinkt.

## Stufe 4 — CAD tiefer

- **IGES → Mesh** im Tessellierungs-Kern (Flächen-Evaluatoren existieren,
  es fehlt der IGES-Entity-Parser).
- **Kennzahlen aus dem Mesh**: Volumen, Oberfläche, Bounding-Box →
  Rohteilgewicht und grobe Kostenschätzung direkt beim Konvertieren.
- **Mesh-Vergleich**: Revision alt vs. neu → „hat sich die Geometrie
  geändert?" (Hausdorff-Abstand, Bericht) — Gold wert bei Zeichnungsänderungen.
- **2D-Ableitungen**: Projektionen des Meshes (vorn/oben/seitlich) als SVG —
  schnelle Ansichts-Blätter ohne CAD.

## Bewusste Grenzen (bleiben)

- Kein Lesen von SLDPRT/IPT/eDrawings (proprietäre Kernel-Daten ohne Spezifikation).
- Kein Office→PDF/A (bräuchte LibreOffice), kein JPEG/RAW (bräuchte Codecs),
  kein RAR/7z. Lieber wenige Dinge zu 100 % als viele zu 80 %.

## Reihenfolge-Empfehlung

1. Preislisten-Diff (größter täglicher Nutzen, kleinster Aufwand)
2. CLI/Headless + Taskplaner (Automatisierung ohne neue UI)
3. Angebotsvergleich, BOM-Abgleich
4. Lauf-Historie (SQLite), dann Watch-Ordner
5. CAD-Kennzahlen und Mesh-Vergleich
