' =====================================================================
' CostingBatch.bas  -  Stapel-Kostenberechnung mit SOLIDWORKS Costing
' ---------------------------------------------------------------------
' Zweck:
'   Alle Teile (*.sldprt) eines Ordners oeffnen, Costing rechnen lassen,
'   das Ergebnis in benutzerdefinierte Eigenschaften schreiben und
'   zusaetzlich als CSV ausgeben.
'
' WICHTIG - Was dieses Makro NICHT kann:
'   Die Costing-API kann KEINE Vorlagen befuellen. Stundensaetze,
'   Materialpreise und Ruestkosten muessen vorher im Costing-Vorlagen-
'   Editor bzw. per Excel-Import in die Vorlage eingetragen werden
'   (siehe ..\README.md bzw. ..\Hintergrund-und-Quellen.md, Kapitel 3 und 4).
'
' Einsatz:
'   SOLIDWORKS > Extras > Makro > Neu ... > Code einfuegen > als .swp
'   speichern. Costing-Add-In muss geladen sein (Extras > Zusatzanwendungen).
'   Voraussetzung: SOLIDWORKS Professional oder Premium.
'
' Versionshinweis:
'   Die mit [VERSION] markierten Aufrufe der Costing-API unterscheiden
'   sich je nach SOLIDWORKS-Release. Vor dem Produktivlauf gegen die
'   lokale API-Hilfe pruefen: Hilfe > API-Hilfe > Costing API >
'   "Get and Set Costing Default Values" und "Create Machining Costing
'   Analysis Example".
' =====================================================================

Option Explicit

' ------------------------- Konfiguration -----------------------------
Const QUELL_ORDNER      As String = "C:\Daten\Teile"                       ' zu bewertende Teile
Const CSV_AUSGABE       As String = "C:\Daten\Costing_Ergebnis.csv"        ' Ergebnisdatei
Const VORLAGE_ZERSPANUNG As String = "\\server\CAD\Costing\01_Produktiv\Firma_Zerspanung_2026.sldctm"
Const VORLAGE_BLECH      As String = "\\server\CAD\Costing\01_Produktiv\Firma_Blech_2026.sldcts"
Const LOSGROESSE        As Long = 1                                        ' Stueckzahl fuer die Kalkulation
Const EIGENSCHAFT_KOSTEN As String = "Costing_Stueckkosten"
Const EIGENSCHAFT_STAND  As String = "Costing_Stand"
Const UNTERORDNER_MIT   As Boolean = True
' ---------------------------------------------------------------------

Dim swApp       As SldWorks.SldWorks
Dim csvHandle   As Integer
Dim anzahlOk    As Long
Dim anzahlFehler As Long

Sub main()

    Set swApp = Application.SldWorks
    anzahlOk = 0
    anzahlFehler = 0

    If Not CostingAddInGeladen() Then
        MsgBox "Das Costing-Add-In ist nicht geladen." & vbCrLf & _
               "Extras > Zusatzanwendungen > SOLIDWORKS Costing aktivieren.", vbExclamation
        Exit Sub
    End If

    csvHandle = FreeFile
    Open CSV_AUSGABE For Output As #csvHandle
    Print #csvHandle, "Datei;Werkstoff;Losgroesse;Stueckkosten_EUR;Status"

    VerarbeiteOrdner QUELL_ORDNER

    Close #csvHandle

    MsgBox "Fertig." & vbCrLf & _
           "Berechnet: " & anzahlOk & vbCrLf & _
           "Fehler/uebersprungen: " & anzahlFehler & vbCrLf & vbCrLf & _
           "Ergebnis: " & CSV_AUSGABE, vbInformation

End Sub


' --------------------------------------------------------------------
' Ordner rekursiv durchlaufen
' --------------------------------------------------------------------
Private Sub VerarbeiteOrdner(ByVal ordner As String)

    Dim datei As String
    Dim unterordner As Collection
    Dim eintrag As Variant

    If Right$(ordner, 1) <> "\" Then ordner = ordner & "\"

    datei = Dir$(ordner & "*.sldprt", vbNormal)
    Do While Len(datei) > 0
        VerarbeiteTeil ordner & datei
        datei = Dir$()
    Loop

    If Not UNTERORDNER_MIT Then Exit Sub

    Set unterordner = New Collection
    datei = Dir$(ordner, vbDirectory)
    Do While Len(datei) > 0
        If datei <> "." And datei <> ".." Then
            If (GetAttr(ordner & datei) And vbDirectory) = vbDirectory Then
                unterordner.Add ordner & datei
            End If
        End If
        datei = Dir$()
    Loop

    For Each eintrag In unterordner
        VerarbeiteOrdner CStr(eintrag)
    Next

End Sub


' --------------------------------------------------------------------
' Ein Teil oeffnen, rechnen, Ergebnis sichern
' --------------------------------------------------------------------
Private Sub VerarbeiteTeil(ByVal pfad As String)

    Dim swModel     As SldWorks.ModelDoc2
    Dim swPart      As SldWorks.PartDoc
    Dim swExt       As SldWorks.ModelDocExtension
    Dim swCostMgr   As Object            ' SldCostingApi.CostManager
    Dim fehler      As Long, warnung As Long
    Dim werkstoff   As String
    Dim kosten      As Double
    Dim istBlech    As Boolean
    Dim status      As String

    Set swModel = swApp.OpenDoc6(pfad, swDocPART, swOpenDocOptions_Silent, "", fehler, warnung)

    If swModel Is Nothing Then
        Print #csvHandle, Dateiname(pfad) & ";;;;Datei konnte nicht geoeffnet werden"
        anzahlFehler = anzahlFehler + 1
        Exit Sub
    End If

    Set swPart = swModel
    Set swExt = swModel.Extension

    werkstoff = swPart.GetMaterialPropertyName2("", "")
    istBlech = HatBlechkoerper(swPart)

    If Len(werkstoff) = 0 Then
        status = "Kein Werkstoff zugewiesen - Costing kann nicht rechnen"
        Print #csvHandle, Dateiname(pfad) & ";;;;" & status
        anzahlFehler = anzahlFehler + 1
        swApp.CloseDoc swModel.GetTitle
        Exit Sub
    End If

    ' --- Costing-Manager holen (dokumentierter Einstiegspunkt) ---------
    Set swCostMgr = swExt.GetCostingManager

    If swCostMgr Is Nothing Then
        status = "Costing nicht verfuegbar (Lizenz/Add-In pruefen)"
        Print #csvHandle, Dateiname(pfad) & ";" & werkstoff & ";;;" & status
        anzahlFehler = anzahlFehler + 1
        swApp.CloseDoc swModel.GetTitle
        Exit Sub
    End If

    kosten = -1
    status = "OK"

    On Error Resume Next
    Err.Clear

    ' ================== [VERSION] ab hier pruefen ======================
    ' Die folgenden Aufrufe existieren sinngemaess in allen aktuellen
    ' Releases, heissen aber je nach Version leicht anders. Referenz:
    '   ICostManager      (Extension.GetCostingManager)
    '   ICostBody         .CreateCostAnalysis / .GetName
    '   ICostAnalysis     Ergebnis der Analyse
    '   ICostFeature      .CombinedCost
    '
    ' 1) Vorlage setzen (Zerspanung oder Blech)
    If istBlech Then
        swCostMgr.TemplateName = VORLAGE_BLECH          ' [VERSION]
    Else
        swCostMgr.TemplateName = VORLAGE_ZERSPANUNG     ' [VERSION]
    End If

    ' 2) Losgroesse setzen
    swCostMgr.TotalPartsCount = LOSGROESSE               ' [VERSION]

    ' 3) Analyse ausfuehren und Gesamtkosten je Stueck lesen
    kosten = HoleStueckkosten(swCostMgr)
    ' ==================================================================

    If Err.Number <> 0 Then
        status = "Costing-API Fehler " & Err.Number & ": " & Err.Description
        Err.Clear
        kosten = -1
    End If
    On Error GoTo 0

    If kosten >= 0 Then
        swExt.CustomPropertyManager("").Add3 EIGENSCHAFT_KOSTEN, swCustomInfoText, _
                                             Format$(kosten, "0.00"), swCustomPropertyReplaceValue
        swExt.CustomPropertyManager("").Add3 EIGENSCHAFT_STAND, swCustomInfoText, _
                                             Format$(Now, "yyyy-mm-dd"), swCustomPropertyReplaceValue
        swModel.Save3 swSaveAsOptions_Silent, fehler, warnung
        anzahlOk = anzahlOk + 1
    Else
        anzahlFehler = anzahlFehler + 1
    End If

    Print #csvHandle, Dateiname(pfad) & ";" & werkstoff & ";" & LOSGROESSE & ";" & _
                      IIf(kosten >= 0, Format$(kosten, "0.00"), "") & ";" & status

    swApp.CloseDoc swModel.GetTitle

End Sub


' --------------------------------------------------------------------
' [VERSION] Analyse ausfuehren und Gesamtkosten je Stueck ermitteln.
' In eine eigene Funktion gekapselt, damit nur hier angepasst werden muss.
' --------------------------------------------------------------------
Private Function HoleStueckkosten(ByVal swCostMgr As Object) As Double

    Dim swCostAnalysis As Object
    Dim ergebnis As Double

    ergebnis = -1

    On Error Resume Next

    ' Variante A: Manager rechnet direkt fuer das aktive Dokument
    Set swCostAnalysis = swCostMgr.CreateCostAnalysis      ' [VERSION]
    If Not swCostAnalysis Is Nothing Then
        ergebnis = swCostAnalysis.TotalCost                ' [VERSION]
    End If

    ' Variante B (falls A leer bleibt): ueber die Koerper gehen
    ' Dim bodies As Variant, i As Long, swCostBody As Object
    ' bodies = swCostMgr.GetCostBodies                     ' [VERSION]
    ' For i = 0 To UBound(bodies)
    '     Set swCostBody = bodies(i)
    '     Set swCostAnalysis = swCostBody.CreateCostAnalysis
    '     ergebnis = ergebnis + swCostAnalysis.TotalCost
    ' Next

    On Error GoTo 0

    HoleStueckkosten = ergebnis

End Function


' --------------------------------------------------------------------
' Hilfsfunktionen
' --------------------------------------------------------------------
Private Function HatBlechkoerper(ByVal swPart As SldWorks.PartDoc) As Boolean

    Dim swFeat As SldWorks.Feature

    HatBlechkoerper = False
    Set swFeat = swPart.FirstFeature

    Do While Not swFeat Is Nothing
        Select Case swFeat.GetTypeName2
            Case "SheetMetal", "SMBaseFlange", "SolidToSheetMetal"
                HatBlechkoerper = True
                Exit Function
        End Select
        Set swFeat = swFeat.GetNextFeature
    Loop

End Function


Private Function CostingAddInGeladen() As Boolean

    Dim swExt As Object

    On Error Resume Next
    ' Costing meldet sich erst zurueck, wenn das Add-In aktiv ist.
    CostingAddInGeladen = swApp.IsAddInLoaded("SldCosting.SldCostingAddin")
    If Err.Number <> 0 Then
        Err.Clear
        CostingAddInGeladen = True   ' Aeltere Releases kennen die Abfrage nicht -> weiterlaufen
    End If
    On Error GoTo 0

End Function


Private Function Dateiname(ByVal pfad As String) As String
    Dim p As Long
    p = InStrRev(pfad, "\")
    If p > 0 Then Dateiname = Mid$(pfad, p + 1) Else Dateiname = pfad
End Function
