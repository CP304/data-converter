# SOLIDWORKS Costing produktiv machen — Recherche & Startdatensatz

Stand: August 2026 · Zielgruppe: Konstruktion/Arbeitsvorbereitung im deutschen Maschinen-/Blechbau

Ausgangslage: Costing ist installiert, aber **alle Vorlagen sind leer** → keine Ergebnisse.
Dieses Dokument beschreibt (1) wie Costing intern tickt, (2) den manuellen Weg,
(3) den halbautomatischen Excel-Weg (offiziell unterstützt, für Massendaten),
(4) den API-/Makro-Weg inkl. dessen Grenzen und (5) marktübliche Startwerte.

---

## 1. Wie Costing funktioniert (und warum aktuell nichts rauskommt)

Costing rechnet **nicht** aus dem Feature-Baum, sondern erkennt Geometrie
(Volumen, Flächen, Bohrungen, Biegungen, Schnittkonturen) und bewertet diese mit
Sätzen, die **ausschließlich in einer Costing-Vorlage (Template)** stehen. Das Teil
selbst enthält keine Preise. Leere Vorlage = 0 € bzw. „keine Kosten".

Die Vorlage liefert im Wesentlichen vier Datenblöcke:

| Block | Inhalt | Beispiel |
|---|---|---|
| Material | Materialklassen, Werkstoffnamen, Preis pro kg (bzw. pro Volumen), Dichte, Rohteil-/Tafelformate, Schrottwert | „Stahl / S235JR / 1,60 €/kg" |
| Maschinen & Sätze | Stundensatz je Maschine/Verfahren, Zeitspanvolumen (MRR), Vorschübe, Schnittgeschwindigkeiten | „3-Achs-BAZ 85 €/h, MRR Stahl 50 cm³/min" |
| Rüsten/Operationen | Rüstkosten je Aufspannung/Werkzeug/Programm, Handling, Zuschlag je Feature | „Fräsen Rüsten 45 min je Aufspannung" |
| Benutzerdefiniert & Aufschläge | Nicht-geometrische Kosten (Lackieren, Prüfen, Verpacken), Zu-/Abschläge, Losgröße | „Pulverbeschichten 25 €/m²" |

Vorlagentypen und Endungen:

| Endung | Vorlage für |
|---|---|
| `.sldctm` | Bearbeitete Teile (Zerspanung) — Basis auch für Guss, Kunststoff-Spritzguss, 3D-Druck |
| `.sldcts` | Blechteile |
| `.sldctc` | Mehrkörper-Teile / Schweißkonstruktionen / Baugruppen |

**Speicherort (Standard):**
`C:\ProgramData\SOLIDWORKS\SOLIDWORKS <Jahr>\lang\german\Costing templates`
(bei englischer Installation `…\lang\english\…`).

**Verzeichnis anmelden:** *Extras > Optionen > Systemoptionen > Dateipositionen >
Costing-Vorlagen* → Ordner hinzufügen. Für ein Team: **Vorlagen auf ein Netzlaufwerk
legen** und dort schreibgeschützt pflegen — sonst hat jeder Anwender eigene Preise.

**Lizenz:** Costing ist Bestandteil von SOLIDWORKS **Professional und Premium**
(Standard hat es nicht). Blech + Zerspanung sind der Kern; Mehrkörper/Schweißkonstruktionen/
Baugruppen bzw. Guss, Kunststoff und 3D-Druck sind je nach Paket/Version unterschiedlich
freigeschaltet — bitte im konkreten Haus-Setup gegenprüfen.

### Die drei häufigsten Gründe für „keine Kosten" trotz gefüllter Vorlage

1. **Werkstoff des Teils ist in der Vorlage nicht hinterlegt.** Costing matcht über
   Materialklasse + Werkstoffname gegen die SOLIDWORKS-Werkstoffdatenbank. Heißt der
   Werkstoff im Teil „1.4301" und in der Vorlage „AISI 304", gibt es keinen Treffer.
   → In der Vorlage **exakt die Namen** aus eurer (ggf. eigenen) Werkstoffbibliothek verwenden.
2. **Blechdicke fehlt.** In der Blechvorlage ist jede Kombination aus Werkstoff **und
   Dicke** eine eigene Zeile (inkl. Schnittgeschwindigkeit und Biegekosten). Fehlt 2,5 mm,
   wird das Teil nicht bewertet.
3. **Kein Rohteil/Blechformat definiert** → kein Materialpreis, damit auch kein Verschnitt.

---

## 2. Manueller Weg (funktioniert immer, keine Freigabe nötig)

Der **Costing-Vorlagen-Editor** ist ein eigenständiges Programm — kein SOLIDWORKS-Fenster.

**Start:** Windows-Startmenü → *SOLIDWORKS Tools <Jahr>* → *Costing Template Editor*,
oder aus SOLIDWORKS im Costing-Task-Fenster über *Vorlage bearbeiten*. Mehrere Editoren
lassen sich parallel öffnen (praktisch, um Zerspanung/Blech/Mehrkörper nebeneinander zu pflegen).

### Vorgehen

1. Mitgelieferte Beispielvorlage öffnen (z. B. `machining_template_default.sldctm`) und
   **sofort unter neuem Namen speichern**, z. B. `<Firma>_Zerspanung_2026.sldctm`
   (in den Netzwerkordner). Nie die Auslieferungsvorlage überschreiben — sonst ist sie
   nach dem nächsten Service Pack weg oder überschrieben.
2. Register **Material**: Materialklassen und Werkstoffe anlegen, Preis pro kg,
   Dichte, Rohteilformen (Rund/Flach/Platte) bzw. Tafelformate, Schrottwert.
3. Register **Fräsen / Drehen / Bohren** (bei Blech: **Schneiden / Biegen /
   Bibliotheks-Features**): pro Materialklasse Maschinenstundensatz und
   Leistungsdaten (Zeitspanvolumen, Vorschub, Schnittgeschwindigkeit, Kosten je Hub
   bei Stanzen, Kosten je Biegung).
4. Register **Operationen/Rüsten**: Rüstkosten je Aufspannung/Werkzeugwechsel/Programm,
   und wie Zusatzkosten je Feature aufgeschlagen werden.
5. Register **Benutzerdefiniert**: alles, was keine Geometrie ist — Lackieren, Eloxieren,
   Prüfen, Verpacken. Diese Operationen lassen sich an eine Materialklasse oder an
   „alle Materialien" hängen und automatisch mitrechnen.
6. Aufschläge/Rabatte werden im Costing-Task-Fenster je Teil gesetzt (Aufschlag positiv,
   Rabatt als negativer Prozentwert, wahlweise bezogen auf Gesamt- oder Materialkosten).
7. Vorlage speichern → in SOLIDWORKS unter *Dateipositionen* eintragen → im Costing-Fenster
   als Standardvorlage wählen.

### Empfohlene Ordnerstruktur

```
\\server\CAD\Costing\
├─ 01_Produktiv\      <Firma>_Zerspanung_2026.sldctm  (schreibgeschützt für User)
│                     <Firma>_Blech_2026.sldcts
│                     <Firma>_Mehrkoerper_2026.sldctc
├─ 02_Entwurf\        Arbeitsstände der AV
└─ 03_Excel\          Export/Import-Dateien + Änderungshistorie
```

---

## 3. Halbautomatisch: Excel-Import/Export (der praktikable Massen-Weg)

Seit **SOLIDWORKS 2016** hat der Vorlagen-Editor oben rechts **Export nach Excel** und
**Import**. Damit lassen sich die **Materialdaten** von Zerspanungs- und Blechvorlagen als
`.xlsx` herausschreiben, in Excel massenhaft bearbeiten (Formeln, Preislisten des
Händlers, SVERWEIS) und wieder einlesen — auch in eine leere oder teilbefüllte Vorlage.

**Das ist der einzige offiziell unterstützte Bulk-Weg** und der, den ich für euren Start
empfehle:

1. Neue leere Vorlage anlegen → **Export nach Excel** → ihr seht die exakte Spaltenstruktur
   eurer Version (die Spaltenbezeichnungen unterscheiden sich zwischen Versionen und
   Sprachständen).
2. Werte aus `Costing-Startwerte-DE-2026.xlsx` (liegt neben diesem Dokument) in die
   exportierte Datei übertragen — Spalten zuordnen, **nicht** die Exportdatei durch meine
   Datei ersetzen. (Die berechneten Spalten der Mappe füllt Excel beim ersten Öffnen selbst;
   erzeugt wird die Mappe von `build_costing_xlsx.py`, falls ihr sie neu generieren wollt.)
3. **Import** in die Vorlage, speichern, mit 2–3 Referenzteilen gegenrechnen.

Grenzen: Import/Export deckt **materialbezogene Daten** ab. Maschinenstundensätze,
Rüstzeiten und benutzerdefinierte Operationen werden in der Regel weiterhin im Editor
eingetragen — das sind aber nur wenige Dutzend Werte, einmalig.

---

## 4. API-/Makro-Weg — was geht und was nicht

**Wichtigste Erkenntnis vorweg: Die Costing-API befüllt keine Vorlagen.**
Es gibt keine dokumentierte Schnittstelle, um Stundensätze, Materialpreise oder
Rüstkosten in eine `.sldctm/.sldcts/.sldctc` zu schreiben. Die Dateien sind ein
geschlossenes Binärformat — sie lassen sich nicht mit Excel, Notepad oder als XML öffnen,
und Fremdschreiben ist weder unterstützt noch versionsstabil.

Was die API **kann** (Namespace `SolidWorks.Interop.sldcostingapi`, Einstieg über
`IModelDocExtension::GetCostingManager` → `ICostManager`):

- Costing für Teile/Körper **automatisiert ausführen** (`ICostBody::CreateCostAnalysis`,
  `ICostAnalysis`), auch im Stapel über ganze Ordner oder eine Baugruppe
- Vorlage je Dokument setzen, Losgröße/Stückzahl, Aufschläge und Standardwerte lesen/setzen
  (API-Hilfe: *Get and Set Costing Default Values*)
- Ergebnisse auslesen (`ICostFeature::CombinedCost` u. a.) und in **benutzerdefinierte
  Eigenschaften**, CSV/Excel oder PDM-Variablen schreiben
- Costing-Berichte automatisiert erzeugen

Daraus folgt die realistische Arbeitsteilung:

| Aufgabe | Weg |
|---|---|
| Vorlage **einmalig befüllen** | Excel-Import (Material) + Vorlagen-Editor (Sätze, Rüsten, Custom) |
| Preise **regelmäßig aktualisieren** | Excel-Datei pflegen → Import in die Vorlage |
| Kosten **für viele Teile berechnen** | VBA-Makro / Add-in über die Costing-API |
| Kosten **in Metadaten/ERP** bringen | Makro schreibt Eigenschaften, Export nach CSV |

Ein einsatzfähiges Makro-Gerüst für den Stapellauf liegt unter
[`makros/CostingBatch.bas`](makros/CostingBatch.bas). Die exakten Member-Namen der
Costing-API variieren zwischen den Versionen — vor dem Produktivlauf gegen die **lokale**
API-Hilfe (`Hilfe > API-Hilfe > Costing API`) eures Releases prüfen; im Makro sind die
betroffenen Stellen markiert.

### Was ihr bei der IT/Systembetreuung freigeben lassen müsst

| Weg | Freigabe nötig? | Anmerkung |
|---|---|---|
| Vorlagen-Editor + manuelle Eingabe | nein | Teil der Installation |
| Excel-Export/Import in der Vorlage | nein | reiner Dateizugriff |
| Netzwerkordner für Vorlagen | Schreibrecht für AV, Leserecht für alle | organisatorisch |
| VBA-Makro (`.swp`) in SOLIDWORKS | meist ja (Makro-Sicherheitsstufe/Vertrauenswürdige Pfade) | einfachster Automatisierungsweg |
| Eigenes Add-in (C#/VB.NET, DLL-Registrierung) | ja, Admin-Rechte | nur wenn Makro nicht reicht |
| Task-Scheduler-Batchlauf (SOLIDWORKS Professional) | nein | für nächtliche Massenläufe kombinierbar |

**Frageliste für die IT** (kurz, damit die Antwort schnell kommt):
1. Sind VBA-Makros in SOLIDWORKS erlaubt, und gibt es einen vertrauenswürdigen Makropfad?
2. Dürfen wir ein Netzlaufwerk als zentralen Costing-Vorlagenpfad eintragen?
3. Haben wir Professional oder Premium (und wie viele Lizenzen mit Costing)?
4. Dürfen Add-ins (DLL) installiert werden, falls die Automatisierung ausgebaut wird?

---

## 5. Marktübliche Startwerte (Deutschland, Stand 2026)

Alle Werte sind **Richtwerte für die deutsche Lohnfertigung/Eigenfertigung** und als
Startpunkt gedacht — sie ersetzen keine eigene Kostenstellenrechnung. Die vollständigen,
einspielbaren Tabellen liegen in `Costing-Startwerte-DE-2026.xlsx`.

### 5.1 Maschinenstundensätze (inkl. Bediener, ohne Gewinn)

| Verfahren | Spanne €/h | Startwert €/h |
|---|---|---|
| Bearbeitungszentrum 3-Achs | 70–100 | **85** |
| Bearbeitungszentrum 5-Achs | 100–150 | **115** |
| CNC-Drehen (2-Achs) | 60–90 | **75** |
| CNC-Drehen mit angetriebenen Werkzeugen | 85–115 | **95** |
| Säge / Trennen | 40–60 | **50** |
| Faserlaser 4–6 kW | 130–180 | **150** |
| Stanz-/Nibbelmaschine | 90–130 | **110** |
| Abkantpresse | 70–95 | **80** |
| Entgraten / Schleifen manuell | 45–65 | **55** |
| Schweißen MAG/WIG | 55–80 | **65** |
| Montage / Handarbeitsplatz | 45–65 | **55** |
| Messraum / Prüfung | 70–95 | **80** |

Erhebungen für das Laserschneiden zeigen eine sehr große Streuung (40–300 €/h,
Mittelwert ~137 €/h); üblich sind 130–180 €/h. Für Fräsen werden 3-Achs typisch mit
70–100 €/h und 5-Achs mit 100–150 €/h angesetzt.

### 5.2 Rüstzeiten (je Los)

| Vorgang | Startwert |
|---|---|
| Fräsen, je Aufspannung | 45 min |
| Drehen, je Aufspannung | 30 min |
| Zusätzlicher Werkzeugwechsel | 3 min |
| Laser, je Programm/Tafel | 10 min |
| Abkanten, Grundrüsten | 15 min |
| Abkanten, je zusätzliches Biegewerkzeug | 6 min |
| Schweißvorrichtung einrichten | 20 min |

### 5.3 Materialpreise (Handelsware, Klein-/Mittelmengen, inkl. Zuschnitt-Aufschlag)

| Werkstoff | €/kg | Schrotterlös €/kg |
|---|---|---|
| S235JR (Blech/Flach) | 1,40–1,90 → **1,65** | 0,15 |
| S355J2 | 1,60–2,10 → **1,80** | 0,15 |
| DC01 (kaltgewalzt) | 1,50–2,00 → **1,75** | 0,15 |
| DX51D verzinkt | 1,60–2,20 → **1,90** | 0,12 |
| C45 / 1.0503 | 1,70–2,30 → **1,95** | 0,15 |
| 42CrMo4 | 2,20–3,00 → **2,50** | 0,15 |
| 1.4301 (V2A) | 8–13 → **10,50** | 1,20 |
| 1.4404 (V4A) | 10–15 → **12,50** | 1,40 |
| EN AW-5754 / AlMg3 | 5,50–8,00 → **6,50** | 1,10 |
| EN AW-6082 / AlMgSi1 | 6,00–9,00 → **7,00** | 1,10 |
| CuZn39Pb3 (Messing) | 9–14 → **11,00** | 5,00 |
| Cu-ETP (Kupfer) | 12–17 → **14,00** | 7,00 |
| POM-C | 6–10 → **8,00** | 0 |
| PA6 | 7–11 → **9,00** | 0 |
| PE-HD | 4–6 → **5,00** | 0 |
| Ti Grade 5 | 45–85 → **60,00** | 6,00 |

Einordnung: Europäisches Warmband lag Mitte 2026 bei rund **1.000 €/t**; auf Blechtafel,
Zuschnitt und Handelsmarge gerechnet landet man bei den oben genannten 1,4–1,9 €/kg.
Für Edelstahlblech werden im Handel 8–15 €/kg genannt, V2A üblicherweise 10–13 €/kg —
Achtung: **Legierungszuschlag** schwankt monatlich, das ist bei V2A/V4A der größte Hebel.

### 5.4 Zeitspanvolumen (MRR) für die Zerspanungsvorlage

Costing rechnet Fräsen/Bohren über abgetragenes Volumen. Wichtig: **nicht** das
Spitzen-MRR des Werkzeugherstellers eintragen, sondern ein **effektives** MRR inkl.
Luftschnitten, Werkzeugwechseln und Schlichten. Faustregel: Herstellerwert × 0,4–0,5.

| Materialklasse | Schruppen cm³/min | Schlichten cm³/min | Startwert effektiv |
|---|---|---|---|
| Aluminium | 150–400 | 20–40 | **120** |
| Baustahl / Automatenstahl | 40–80 | 8–15 | **35** |
| Vergütungsstahl 42CrMo4 | 25–50 | 5–10 | **22** |
| Edelstahl 1.4301/1.4404 | 20–40 | 4–8 | **16** |
| Titan | 8–15 | 2–4 | **7** |
| Kunststoff | 300–600 | 50–80 | **200** |
| Guss GG/GGG | 50–90 | 10–18 | **40** |

Bohren (Vorschubgeschwindigkeit, VHM/HSS gemischt): Alu **500 mm/min**,
Stahl **180 mm/min**, Edelstahl **110 mm/min**, Kunststoff **600 mm/min**.

### 5.5 Blech: Schnittgeschwindigkeiten Faserlaser (4–6 kW), mm/min

| Dicke | S235 (O₂) | 1.4301 (N₂) | AlMg3 (N₂) |
|---|---|---|---|
| 1,0 mm | 8.000 | 25.000 | 20.000 |
| 1,5 mm | 6.500 | 18.000 | 14.000 |
| 2,0 mm | 5.000 | 12.000 | 9.500 |
| 3,0 mm | 3.200 | 7.000 | 5.500 |
| 4,0 mm | 2.600 | 4.800 | 3.800 |
| 5,0 mm | 2.200 | 3.600 | 2.800 |
| 6,0 mm | 1.900 | 2.900 | 2.200 |
| 8,0 mm | 1.500 | 2.100 | 1.500 |
| 10,0 mm | 1.250 | 1.600 | 1.000 |
| 12,0 mm | 1.000 | 1.150 | 700 |
| 15,0 mm | 800 | 800 | — |
| 20,0 mm | 600 | — | — |

(Referenzpunkte aus Herstellerangaben: 6 kW, 10 mm Edelstahl ≈ 2.000–2.500 mm/min;
20 mm Baustahl ≈ 950 mm/min. Die Tabelle ist bewusst konservativ, weil Costing damit
inkl. Beschleunigungs- und Konturverlusten rechnet.)

**Einstechzeit (Pierce)** je Loch/Kontur: bis 3 mm **0,3 s**, 4–6 mm **0,8 s**,
8–10 mm **1,5 s**, 12–15 mm **2,5 s**, 20 mm **4,0 s**.

**Stanzen (falls vorhanden):** 0,04–0,08 € je Hub, Standardwerkzeug; Umformwerkzeuge
(Durchzug, Kiemen) 0,15–0,30 € je Hub.

### 5.6 Blech: Biegen

Kalkulationsbasis 80 €/h Abkantpresse:

| Teilegröße / Kantenlänge | Zeit je Biegung | Kosten je Biegung |
|---|---|---|
| klein, < 500 mm, 1–3 mm | 0,10–0,20 min | **0,20 €** |
| mittel, 500–1.500 mm | 0,20–0,35 min | **0,35 €** |
| groß, > 2.000 mm oder 2-Mann | 0,50–0,90 min | **0,90 €** |
| enger Radius / schwer zugänglich / Aufschlag | +30–60 % | — |

Zusätzlich Handling je Teil 0,15–0,40 min. Hinweis aus der Praxis: ab Losgröße 50–100
sinken die Stückkosten je Biegung um 20–40 % gegenüber Einzelteilen — das bildet ihr in
Costing über die **Losgröße** und die Rüstkostenverteilung ab, nicht über kleinere
Biegepreise.

### 5.7 Benutzerdefinierte Operationen (Register „Benutzerdefiniert")

| Operation | Richtwert |
|---|---|
| Pulverbeschichten | 22–35 €/m², Mindestlos 40–70 € |
| Nasslackieren | 30–55 €/m² |
| Eloxieren (natur) | 25–45 €/m² |
| Verzinken galvanisch | 0,80–1,50 €/kg, Mindestlos 30 € |
| Feuerverzinken | 0,55–0,95 €/kg |
| Brünieren | 0,60–1,10 €/kg |
| Entgraten manuell | 0,5–3 min/Teil (≈ 0,45–2,75 €) |
| Gewinde M4–M12 schneiden | 0,10–0,30 €/Gewinde |
| Senken | 0,05–0,12 €/Senkung |
| Einpressmutter/-bolzen setzen | 0,25–0,55 €/Stück inkl. Teil |
| Lasergravur/Beschriftung | 0,50–2,00 €/Teil |
| Erstmusterprüfbericht | 60–150 €/Los |
| Verpackung | 1,50–6,00 €/Teil |

### 5.8 Aufschläge

| Position | Startwert |
|---|---|
| Ausschuss/Nacharbeit | 2 % |
| Verwaltungs-/Vertriebsgemeinkosten | 12 % |
| Gewinn | 8 % (Einzelteil 12–15 %, Serie 5–8 %) |
| Werkzeugverschleiß | im Stundensatz enthalten |

---

## 6. Empfohlene Reihenfolge für die Einführung

1. **Woche 1 — Grundgerüst:** Netzwerkordner + drei Vorlagen anlegen, Dateiposition in
   den Systemoptionen eintragen. Nur die Werkstoffe aufnehmen, die ihr wirklich verbaut
   (typisch 8–15 statt 200).
2. **Woche 1 — Material per Excel:** Vorlage → Export nach Excel → Werte aus der
   Startwert-Mappe übernehmen → Import.
3. **Woche 2 — Sätze und Rüsten:** Stundensätze aus 5.1/5.2 im Editor eintragen, dabei
   eigene Maschinen benennen (nicht „Mill 1", sondern „DMU 50", „TruLaser 3030").
4. **Woche 2 — Kalibrieren:** 3 typische Teile wählen, zu denen ihr **echte** Angebote
   oder Nachkalkulationen habt (1× Frästeil, 1× Drehteil, 1× Blechteil). Costing dagegen
   rechnen und die Sätze so lange anpassen, bis die Abweichung < 10 % ist. **Dieser
   Schritt entscheidet über die Akzeptanz** — die Absolutwerte in Kapitel 5 sind nur der
   Startpunkt.
5. **Woche 3 — Ausrollen:** Vorlagen schreibgeschützt setzen, Kurz-Anleitung für die
   Konstruktion, Costing-Ergebnis als benutzerdefinierte Eigenschaft in die Teile.
6. **Danach — Automatisieren:** Makro aus `makros/` für Stapelläufe, Excel-Datei
   quartalsweise mit den Händlerpreisen aktualisieren und neu importieren.

**Pflegeintervall:** Materialpreise quartalsweise (bei V2A/V4A wegen Legierungszuschlag
eher monatlich), Stundensätze jährlich mit der Kostenrechnung.

---

## 7. Quellen

- [Costing-Vorlagen – SOLIDWORKS Hilfe 2025](https://help.solidworks.com/2025/english/solidworks/sldworks/c_Costing_Templates.htm)
- [Machining Costing Vorlagen-Editor – SOLIDWORKS Hilfe (deutsch)](http://help.solidworks.com/2018/german/SolidWorks/sldworks/c_costing_template_editor_machining.htm)
- [Costing Vorlagen-Editor für Blechteile – SOLIDWORKS Hilfe (deutsch)](https://help.solidworks.com/2019/german/SolidWorks/sldworks/c_costing_template_editor_sheet_metal.htm)
- [Costing Template Editor for Sheet Metal Parts – 2026](https://help.solidworks.com/2026/english/SolidWorks/Sldworks/c_costing_template_editor_sheet_metal.htm)
- [Importing and Exporting in Costing Templates](https://help.solidworks.com/2024/english/SolidWorks/sldworks/c_importing_and_exporting_in_costing_templates.htm)
- [Updating Template Material Cost Data (Excel)](https://help.solidworks.com/2022/english/SolidWorks/sldworks/t_costing_import_export_template_excel.htm)
- [Creating a New Costing Template](https://help.solidworks.com/2025/english/solidworks/sldworks/t_create_new_costing_template.htm)
- [Kostenberechnung für benutzerdefinierte Operationen für ausgewählte Materialien](https://help.solidworks.com/2019/german/SolidWorks/sldworks/t_Cost_select_material_mat_class.htm)
- [SOLIDWORKS Costing API Help (sldcostingapi)](https://help.solidworks.com/2024/english/api/SWHelp_List.html?id=d8090a89d8b040459d280a3216ca4e08)
- [Get and Set Costing Default Values Example (VBA)](https://help.solidworks.com/2022/English/api/swcostingapi/Get_and_Set_Costing_Default_Values_Example_VB.htm)
- [Create Machining Costing Analysis Example (VBA)](https://help.solidworks.com/2024/english/api/sldworksapi/Create_Machining_Costing_Analyses_Example_VB.htm)
- [CombinedCost Property (ICostFeature)](https://help.solidworks.com/2024/english/api/swcostingapi/SolidWorks.Interop.sldcostingapi~SolidWorks.Interop.sldcostingapi.ICostFeature~CombinedCost.html)
- [SOLIDWORKS Costing Template Editor – CATI](https://www.cati.com/blog/solidworks-costing-template-editor/)
- [Costing Templates 101: Editing Templates – TriMech](https://trimech.com/blog/costing-templates-101-editing-templates)
- [SOLIDWORKS Costing; automating the cost estimation – PLM Group](https://support.plmgroup.eu/hc/en-us/articles/360015578337-SOLIDWORKS-Costing-automating-the-cost-estimation)
- [SOLIDWORKS Costing – Übersicht (deutsch)](https://help.solidworks.com/2024/German/SWConnected/swdotworks/c_costing_overview.htm)
- [Maschinenstundensatz korrekt kalkulieren – pos.de](https://pos.de/blog/fertigungs-know-how/maschinenstundensatz-korrekt-kalkulieren-schluss-mit-bauchgefuehl/)
- [CNC Fräsen Preisliste 2025 – CNC Magazin](https://cnc-and-more.blog/cnc-fraesen-preisliste-2025-kostenfaktoren-und-tipps/)
- [Stundensatz CNC Fräsen Rechner – CNCRechner.de](https://www.cncrechner.de/rechner/cnc-kosten/)
- [Stundensätze für das Laserschneiden – orderspot](https://orderspot.de/stundensaetze-fuer-das-laserschneiden-bei-orderspot/)
- [Was kostet Laserschneiden – MicroStep](https://microstep.com/de/Expertenwissen/Laserschneiden/was-kostet-laserschneiden)
- [Schnitt- und Betriebskosten der Schneidverfahren – Schneidforum](https://www.schneidforum.de/schneidwissen/schneidkosten/)
- [Abkanten von Blech: Toleranzen und Kosten – Futronika](https://www.futronika.de/news/abkanten-blech/)
- [Stahlpreise Prognose 2025/2026 – Maschine & Werkzeug](https://www.maschinewerkzeug.de/stahlpreise-prognose-entwicklung-chart/)
- [Edelstahlblech Preis: Kosten pro kg](https://evek.top/blog/post/228-edelstahlblech-preis-kosten-pro-kg-qm-und-nach-tafelformat)
- [Legierungszuschlag 1.4301](https://legierungszuschlag.info/en/wkst/4301)
- [Faserlaser Schnittgeschwindigkeiten – ADHMT](https://www.adhmt.com/fiber-laser-cutting-speed-raycus-laser-source/)
- [Zeitspanvolumen Fräsen/Bohren – Meusburger](https://schnittdaten.meusburger.com/zeitspanvolumen-fraesen-bohren/)
