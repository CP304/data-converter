# SOLIDWORKS Costing zum Laufen bringen

**Warum kommt aktuell nichts raus?**
Costing holt sich jede Zahl aus einer **Vorlage**. Im Teil selbst stehen keine Preise.
Leere Vorlage = 0 €. Ihr müsst also einmal drei Vorlagen befüllen, dann rechnet Costing
für jedes Teil automatisch.

**Was ihr braucht:** ca. einen halben Tag, die Excel-Mappe aus diesem Ordner
(`Costing-Startwerte-DE-2026.xlsx`) und jemanden, der eure Stundensätze kennt.

---

## In 6 Schritten

**1. Vorlagen-Editor öffnen**
Windows-Startmenü → *SOLIDWORKS Tools <Jahr>* → *Costing Template Editor*.
Das ist ein eigenes Programm, nicht SOLIDWORKS.

**2. Vorlage anlegen**
Mitgelieferte Beispielvorlage öffnen → sofort **Speichern unter** mit eigenem Namen,
am besten auf einem Netzlaufwerk. Drei Stück braucht ihr:

| Datei | wofür |
|---|---|
| `Firma_Zerspanung.sldctm` | Frästeile, Drehteile |
| `Firma_Blech.sldcts` | Blechteile |
| `Firma_Mehrkoerper.sldctc` | Schweißkonstruktionen, Baugruppen |

**3. Werkstoffe eintragen**
Nur die, die ihr wirklich verwendet — typisch 8–15 Stück, nicht 200.
Der Name muss **exakt** so heißen wie in eurer SOLIDWORKS-Werkstoffdatenbank,
sonst findet Costing den Werkstoff nicht.
→ Werte: Blatt `10_Material_Zerspanung` und `11_Material_Blech` der Excel-Mappe.

**4. Stundensätze und Rüstzeiten eintragen**
Benennt die Maschinen wie im Betrieb („DMU 50" statt „Mill 1").
→ Werte: Blatt `20_Maschinen_Stundensaetze` und `21_Ruestzeiten`.

**5. Vorlagen in SOLIDWORKS anmelden**
*Extras > Optionen > Systemoptionen > Dateipositionen > Costing-Vorlagen* → Ordner hinzufügen.
Danach im Costing-Fenster die Vorlage auswählen.

**6. Kalibrieren — der wichtigste Schritt**
Drei Teile nehmen, zu denen ihr ein **echtes** Angebot oder eine Nachkalkulation habt
(1 Frästeil, 1 Drehteil, 1 Blechteil). Costing dagegen rechnen und die Sätze anpassen,
bis die Abweichung unter 10 % liegt. Erst dann ausrollen.

---

## Die wichtigsten Startwerte

Alle Werte sind Richtwerte für Deutschland, Stand 2026. Die vollständigen Tabellen
(inkl. Laser-Schnittgeschwindigkeiten, Zeitspanvolumen, Biegekosten, Oberflächen)
stehen in `Costing-Startwerte-DE-2026.xlsx`.

**Stundensätze**

| Maschine | €/h |
|---|---|
| Bearbeitungszentrum 3-Achs | 85 |
| Bearbeitungszentrum 5-Achs | 115 |
| CNC-Drehen | 75 |
| Faserlaser 4–6 kW | 150 |
| Abkantpresse | 80 |
| Schweißen | 65 |
| Montage / Handarbeit | 55 |

**Material (Handelsware, Kleinmengen)**

| Werkstoff | €/kg |
|---|---|
| S235JR | 1,65 |
| C45 | 1,95 |
| 1.4301 (V2A) | 10,50 |
| 1.4404 (V4A) | 12,50 |
| AlMg3 (5754) | 6,50 |
| AlMgSi1 (6082) | 7,00 |
| POM | 8,00 |

**Rüstzeiten:** Fräsen 45 min je Aufspannung · Drehen 30 min · Laser 10 min je Tafel ·
Abkanten 15 min je Los

**Aufschläge:** Ausschuss 2 % · Verwaltung/Vertrieb 12 % · Gewinn 8 % (Einzelteil 12–15 %)

---

## Geht das auch automatisch?

**Materialpreise ja — per Excel.** Im Vorlagen-Editor gibt es oben rechts
**Export nach Excel** und **Import**. Damit könnt ihr die Werkstoffe massenhaft pflegen
und später Preise aktualisieren, ohne jede Zeile anzuklicken. Empfohlener Ablauf:
Vorlage exportieren → Werte aus der Mappe in die exportierte Datei übertragen → importieren.

**Stundensätze und Rüstzeiten: nein.** Die tippt ihr einmal im Editor ein — es sind nur
ein paar Dutzend Zahlen.

**Vorlagen per Programmierung befüllen: geht nicht.** Die Costing-API kann Vorlagen weder
lesen noch schreiben, und die Vorlagendateien sind ein geschlossenes Binärformat.

**Was die API kann:** Costing für viele Teile automatisch durchrechnen und die Kosten in
Dateieigenschaften oder eine CSV schreiben — z. B. nachts über einen ganzen Ordner.
Ein VBA-Gerüst dafür liegt in [`makros/CostingBatch.bas`](makros/CostingBatch.bas).
Dafür müssen bei euch Makros erlaubt sein.

---

## Wenn trotzdem 0,00 € rauskommt

1. **Werkstoff fehlt in der Vorlage** oder heißt anders als im Teil (z. B. „1.4301" gegen „AISI 304").
2. **Blechdicke fehlt.** In der Blechvorlage ist jede Kombination aus Werkstoff *und* Dicke
   eine eigene Zeile. Ohne 2,5 mm wird ein 2,5-mm-Teil nicht bewertet.
3. **Kein Rohteil bzw. Blechformat hinterlegt** → kein Materialpreis.

---

## Was ihr noch klären müsst

- Habt ihr **Professional oder Premium**? Costing ist in beiden enthalten, der Umfang
  (Baugruppen, Schweißkonstruktionen) unterscheidet sich je nach Paket und Version.
- Dürfen die Vorlagen auf ein **Netzlaufwerk**, damit alle mit denselben Preisen rechnen?
- Sind **VBA-Makros** freigegeben? Nur nötig, wenn ihr die Berechnung automatisieren wollt.

---

## Mehr Details

Ausführliche Fassung mit Begründungen, Spannbreiten und allen Quellen:
[`Hintergrund-und-Quellen.md`](Hintergrund-und-Quellen.md)

Dateien in diesem Ordner:

| Datei | Inhalt |
|---|---|
| `Costing-Startwerte-DE-2026.xlsx` | alle Startwerte, 10 Blätter |
| `makros/CostingBatch.bas` | VBA-Makro für Stapelberechnung |
| `build_costing_xlsx.py` | erzeugt die Excel-Mappe neu |
| `Hintergrund-und-Quellen.md` | Langfassung inkl. Quellenliste |
