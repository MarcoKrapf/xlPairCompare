VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGUI 
   Caption         =   "[Titel]"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   OleObjectBlob   =   "frmGUI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Hauptcode des Tools, steuert die GUI
'------------------------------------

'Variablen definieren
Dim wksAusgabe As Worksheet
Dim shp As Shape
Dim objMail As Object 'Shell-Objekt für E-Mail
Dim blnArea1 As Boolean 'Kennzeichen ob Bereich 1 eingelesen ist
Dim blnArea2 As Boolean 'Kennzeichen ob Bereich 2 eingelesen ist
Dim rngCheck As Range 'Range für den Check ob die Bereiche sauber sind
Dim varCheckZelle As Variant 'Variable für eine einzelne Zelle der Range für den Check
Dim lngCheck As Long 'Variable für den Check jedes einzelnen Zeichens in den Bereichen (Position)
Dim strCheckZeichen As String 'Variable für den Check jedes einzelnen Zeichens in den Bereichen (Wert)
Dim blnCheckFund As Boolean 'Kennzeichen dass ein unerwünschtes Zeichen gefunden wurde (benötigt für SchnellCheck)
Dim blnCheckLeer As Boolean 'Kennzeichen ob schon ein unnötiges Leerzeichen gefunden wurde
Dim blnCheckSteuer As Boolean 'Kennzeichen ob schon ein Steuerzeichen gefunden wurde
Dim blnCheckGesch As Boolean 'Kennzeichen ob schon ein geschütztes Leerzeichen gefunden wurde
Dim byteCheckSumme As Byte 'Variable zum Speichern der Check-Ergebnisse
Dim lngMaxAnzahlZeilenBereich As Long 'Zähler für die Längste Spalte eines Bereichs
Dim lngAnzahlLinienAktuell As Long 'Anzahl der Linien des aktuellen Vergleichs
Dim lngFunde As Long 'Zählvariable für die Funde an unnötigen Leerzeichen bzw. nicht druckbaren Zeichen
Dim lngAnzLeer As Long 'Zählvariable für die Anzahl Zellen mit unnötigen Leerzeichen im Gesamt-Check
Dim lngAnzSteuer As Long 'Zählvariable für die Anzahl Zellen mit Steuerzeichen im Gesamt-Check
Dim lngAnzGesch As Long 'Zählvariable für die Anzahl Zellen mit geschützten Leerzeichen im Gesamt-Check
Dim strTimestamp As String 'Timestamp wann der Vergleich gestartet wurde, wird benötigt für die Namen der Linien
Dim varAuswahl1 As Variant 'Array mit dem eingelesenen Bereich 1
Dim varAuswahl2 As Variant 'Array mit dem eingelesenen Bereich 2
Dim str1 As String 'Variable um Bereich 1 einzulesen
Dim str2 As String 'Variable um Bereich 2 einzulesen
Dim lngNr As Long 'Nummer des Treffers
Dim lngE1 As Long, lngE2 As Long, lngP As Long 'Anzahl Einzelwerte und Wertpaare
Dim i As Long, j As Long, k As Long 'Zählvariablen für Schleifen
Dim M As Long, s As Long, p As Long, z As Long 'Variablen für die Ausgabe


Sub Start(area1 As Variant, area2 As Variant)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Variablen neu dimensionieren
        'Spalte 1: Vergleichswert bzw. "#areac#"
        'Spalte 2: Letzte Zelle des ersten bzw. erste Zelle des zweiten Bereichs
        'Spalte 3: Nummer des Treffers
    ReDim varAuswahl1(1 To g_varSelection1.Rows.Count, 1 To 3)
    ReDim varAuswahl2(1 To g_varSelection2.Rows.Count, 1 To 3)
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
        Call BereicheEinlesen
        Call WerteSuchen
        
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub BereicheEinlesen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    str1 = ""
    str2 = ""
    
    'Ersten Bereich einlesen
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "AA0") '"Paare finden"
        g_strSanduhrNummer = "[1/5]"
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "AA1") '"Ersten Bereich einlesen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / (g_varSelection1.Rows.Count)
        
        For i = g_varSelection1.Row To g_varSelection1.Row + g_varSelection1.Rows.Count - 1
            For j = g_varSelection1.Column To g_varSelection1.Column + g_varSelection1.Columns.Count - 1
                str1 = str1 & ZelleEinlesen()
            Next j
            varAuswahl1(i - g_varSelection1.Row + 1, 1) = str1 'Wert
            varAuswahl1(i - g_varSelection1.Row + 1, 2) = Cells(i, j - 1).AddressLocal 'Zellposition
            str1 = "" 'String zurücksetzen
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i
    
    'Zweiten Bereich einlesen
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "AA0") '"Paare finden"
        g_strSanduhrNummer = "[2/5]"
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "AA2") '"Zweiten Bereich einlesen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / (g_varSelection2.Rows.Count)
        
        For i = g_varSelection2.Row To g_varSelection2.Row + g_varSelection2.Rows.Count - 1
            For j = g_varSelection2.Column To g_varSelection2.Column + g_varSelection2.Columns.Count - 1
                str2 = str2 & ZelleEinlesen()
            Next j
            varAuswahl2(i - g_varSelection2.Row + 1, 1) = str2  'Wert
            varAuswahl2(i - g_varSelection2.Row + 1, 2) = Cells(i, j - g_varSelection2.Columns.Count).AddressLocal  'Zellposition
            str2 = "" 'String zurücksetzen
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Function ZelleEinlesen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Suchoptionen berücksichtigen
    If checkboxGrossKleinBuchstaben.Value = True And checkboxLeerzeichen.Value = True Then 'wenn beide Checkboxen aktiv
        ZelleEinlesen = Replace(LCase(Cells(i, j).Value), " ", "") 'alles zu Kleinbuchstaben konvertieren und Leerzeichen entfernen
    ElseIf checkboxGrossKleinBuchstaben.Value = True And checkboxLeerzeichen.Value = False Then 'nur Checkbox "Groß-/Kleinbuchstaben ignorieren" aktiv
        ZelleEinlesen = LCase(Cells(i, j).Value) 'alles zu Kleinbuchstaben konvertieren, Leerzeichen beibehalten
    ElseIf checkboxGrossKleinBuchstaben.Value = False And checkboxLeerzeichen.Value = True Then 'nur Checkbox "Leerzeichen ignorieren" aktiv
        ZelleEinlesen = Replace(Cells(i, j).Value, " ", "") 'alle Leerzeichen löschen, Groß-/Kleinschreibung beibehalten
    Else 'wenn keine Checkboxen aktiv
        ZelleEinlesen = Cells(i, j).Value 'Groß-/Kleinschreibung und Leerzeichen beibehalten
    End If

Exit Function
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Function

Private Sub WerteSuchen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    'Zähler zurücksetzen
    lngNr = 1
    lngE1 = 0
    lngE2 = 0
    lngP = 0
    
    'Gleiche Werte suchen und als gefunden markieren
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "AA0") '"Paare finden"
        g_strSanduhrNummer = "[3/5]"
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "AA3") '"Wertpaare suchen und als gefunden markieren"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / UBound(varAuswahl1)
    
        For i = 1 To UBound(varAuswahl1)
            For j = 1 To UBound(varAuswahl2)
                If varAuswahl1(i, 1) = varAuswahl2(j, 1) _
                    And varAuswahl1(i, 1) <> "#areac#" _
                    And varAuswahl1(i, 1) <> "" Then
                        varAuswahl1(i, 1) = "#areac#"
                        varAuswahl2(j, 1) = "#areac#"
                        varAuswahl1(i, 3) = lngNr
                        varAuswahl2(j, 3) = lngNr
                        lngNr = lngNr + 1
                End If
            Next j
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i
    
    'Werte zählen
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "AA0") '"Paare finden"
        g_strSanduhrNummer = "[4/5]"
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "AA4") '"Werte des ersten Bereichs zählen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / UBound(varAuswahl1)
        
        For i = 1 To UBound(varAuswahl1)
            If varAuswahl1(i, 1) <> "" And varAuswahl1(i, 1) <> "#areac#" Then
                lngE1 = lngE1 + 1 'Einzelwert in Bereich 1
            ElseIf varAuswahl1(i, 1) <> "" Then
                lngP = lngP + 1 'Wertpaar
            End If
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i
        
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "AA0") '"Paare finden"
        g_strSanduhrNummer = "[5/5]"
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "AA5") '"Werte des zweiten Bereichs zählen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / UBound(varAuswahl2)
        
        For i = 1 To UBound(varAuswahl2)
            If varAuswahl2(i, 1) <> "" And varAuswahl2(i, 1) <> "#areac#" Then
                lngE2 = lngE2 + 1 'Einzelwert in Bereich 2
            End If
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i
        
    'Sanduhr ausblenden
    Call SanduhrAusblenden
                
    'Einzelne Werte farbig hervorheben wenn angehakt
    If checkboxHervorhebenEinzeln.Value = True Then
        Call EinzelwerteHervorheben
    End If
    
    'Wertpaare farbig hervorheben wenn angehakt
    If checkboxHervorhebenPaare.Value = True Then
        Call PaareHervorheben
    End If
    
    'Linien zeichnen wenn angehakt
    If checkboxLinien.Value = True Then
        Call LinieZeichnen2
    End If
    
    'Buttons aktivieren
    If lngE1 > 0 Or lngE2 > 0 Then
        btnHervorhebenEinzeln.Enabled = True
        btnHervorhebenEinzeln.BackColor = &H80FF80    'grün
        btnAusgabe1.Enabled = True
        btnAusgabe1.BackColor = &H80FF80              'grün
    End If
    If lngP > 0 Then
        btnHervorhebenPaare.Enabled = True
        btnHervorhebenPaare.BackColor = &H80FFFF      'gelb
        btnLinienZeichnen.Enabled = True
        btnLinienZeichnen.BackColor = &HC000C0        'lila
        btnAusgabe2.Enabled = True
        btnAusgabe2.BackColor = &H80FFFF              'gelb
    End If

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub WerteAusgeben(byteWas As Byte, blnAusgNur1 As Boolean, blnAusgNur2 As Boolean)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If optAusgNeu.Value = True Then
        Call WerteAusgebenNeuesBlatt(byteWas, blnAusgNur1, blnAusgNur2)
    Else
        Call WerteAusgebenCursor(byteWas, blnAusgNur1, blnAusgNur2)
    End If

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub WerteAusgebenNeuesBlatt(byteWas As Byte, blnAusgNur1 As Boolean, blnAusgNur2 As Boolean)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Neues Tabellenblatt erstellen
    Set wksAusgabe = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    wksAusgabe.Name = modINFOtexte.SonstigerText(g_strSprache, "TXT1") & Worksheets.Count
    wksAusgabe.Activate
    
    'Zeile und Spalte für die Ausgabe festlegen
    s = 1 'Spaltenkorrektur
    z = 1 'Zeilenkorrektur
        If blnAusgNur1 = True Then
            p = 0 'Spaltenkorrektur
            Call AusgebenBereich1(byteWas)
        ElseIf blnAusgNur2 = True Then
            p = 0 'Spaltenkorrektur
            Call AusgebenBereich2(byteWas)
        Else
            p = g_varSelection1.Columns.Count + 1 'Spaltenkorrektur
            Call AusgebenBereich1(byteWas)
            Call AusgebenBereich2(byteWas)
        End If
        
    'Sanduhr ausblenden
    Call SanduhrAusblenden

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub WerteAusgebenCursor(byteWas As Byte, blnAusgNur1 As Boolean, blnAusgNur2 As Boolean)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Zeile und Spalte für die Ausgabe festlegen
    s = Selection.Column 'Spaltenkorrektur
    z = Selection.Row 'Zeilenkorrektur
        If blnAusgNur1 = True Then
            p = 0 'Spaltenkorrektur
            Call AusgebenBereich1(byteWas)
        ElseIf blnAusgNur2 = True Then
            p = 0 'Spaltenkorrektur
            Call AusgebenBereich2(byteWas)
        Else
            p = g_varSelection1.Columns.Count + 1 'Spaltenkorrektur
            Call AusgebenBereich1(byteWas)
            Call AusgebenBereich2(byteWas)
        End If
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub AusgebenBereich1(byteWas As Byte)
    'Ausgeben Bereich 1
    
    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
    g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "BB0") '"Ausgeben"
    g_strSanduhrNummer = "[1/2]"
    g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "BB1") '"Ausgeben des ersten Bereichs"
    'Fortschrittsbalken zurücksetzen
    Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    'Stückelung des Balkens berechnen
    g_dblBalkenAnteil = 100 / UBound(varAuswahl1)
    
    k = 0 'Zähler für die aktuelle Ausgabezeile zurücksetzen
    For i = 1 To UBound(varAuswahl1, 1)
        'Einzelwerte ausgeben
        If byteWas = 1 Then
            If varAuswahl1(i, 1) <> "#areac#" Then
                M = s 'Zähler für die aktuelle Ausgabespalte zurücksetzen
                For j = 1 To g_varSelection1.Columns.Count
                    Cells(k + z, M).Value = g_varSelection1(i, j)
                    M = M + 1 'Zähler für die aktuelle Ausgabespalte hochzählen
                Next j
                k = k + 1 'Zähler für die aktuelle Ausgabezeile hochzählen
            End If
        'Wertpaare ausgeben
        Else
            If varAuswahl1(i, 1) = "#areac#" Then
                M = s 'Zähler für die aktuelle Ausgabespalte zurücksetzen
                For j = 1 To g_varSelection1.Columns.Count
                    Cells(k + z, M).Value = g_varSelection1(i, j)
                    M = M + 1 'Zähler für die aktuelle Ausgabespalte hochzählen
                Next j
                k = k + 1 'Zähler für die aktuelle Ausgabezeile hochzählen
            End If
        End If
        
        'Sanduhr aktualisieren
        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next i
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub AusgebenBereich2(byteWas As Byte)
    'Ausgeben Bereich 2
    
    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
    g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "BB0") '"Ausgeben"
    g_strSanduhrNummer = "[2/2]"
    g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "BB2") '"Ausgeben des zweiten Bereichs"
    'Fortschrittsbalken zurücksetzen
    Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    'Stückelung des Balkens berechnen
    g_dblBalkenAnteil = 100 / UBound(varAuswahl2)
    
    k = 0 'Zähler für die aktuelle Ausgabezeile zurücksetzen
    For i = 1 To UBound(varAuswahl2, 1)
        'Einzelwerte ausgeben
        If byteWas = 1 Then
            If varAuswahl2(i, 1) <> "#areac#" Then
                M = s + p 'Zähler für die aktuelle Ausgabespalte zurücksetzen
                For j = 1 To g_varSelection2.Columns.Count
                    Cells(k + z, M).Value = g_varSelection2(i, j)
                    M = M + 1 'Zähler für die aktuelle Ausgabespalte hochzählen
                Next j
                k = k + 1 'Zähler für die aktuelle Ausgabezeile hochzählen
            End If
        'Wertpaare ausgeben
        Else
            If varAuswahl2(i, 1) = "#areac#" Then
                M = s + p 'Zähler für die aktuelle Ausgabespalte zurücksetzen
                For j = 1 To g_varSelection2.Columns.Count
                    Cells(k + z, M).Value = g_varSelection2(i, j)
                    M = M + 1 'Zähler für die aktuelle Ausgabespalte hochzählen
                Next j
                k = k + 1 'Zähler für die aktuelle Ausgabezeile hochzählen
            End If
        End If
        
        'Sanduhr aktualisieren
        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next i
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub AreaAuswaehlen(ByVal byteArea As Byte)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    Select Case byteArea
        Case 1: Call Area1Auswaehlen
        Case 2: Call Area2Auswaehlen
        Case 3: Call AreasAuswaehlenBeide
    End Select

    Call ButtonsCheck
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub Area1Auswaehlen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If Selection.Areas.Count > 1 Then
        MsgBox modINFOtexte.MsgBoxText(g_strSprache, "MA1"), vbExclamation, g_strTool & " " & g_strVersion
        Exit Sub
    Else
        'Bereich einlesen
        Set g_varSelection1 = Selection
        'Check ob ganze Spalte selektiert ist
        If g_varSelection1.Rows.Count = 1048576 Then
            Call BereichAnpassen(g_varSelection1, 1)
        End If

        'GUI beschriften
        lblArea1.Caption = Replace(g_varSelection1.AddressLocal, "$", "")
            'Beschriftung in Tab "Bereinigung"
            lblCheckArea1.Caption = lblArea1.Caption
            lblAnzahlLeer.Caption = "(?)"
            lblAnzahlSteuer.Caption = "(?)"
            lblAnzahlGesch.Caption = "(?)"
            'Beschriftung im Tab 'Visualisierung"
            Call modGUItexte.TextButtonHervorhebungenLoeschen(g_strSprache)
        'GUI aktualisieren
        Call ButtonsAktivieren
        'Start-Button aktualisieren
        blnArea1 = True
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub Area2Auswaehlen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If Selection.Areas.Count > 1 Then
        MsgBox modINFOtexte.MsgBoxText(g_strSprache, "MA1"), vbExclamation, g_strTool & " " & g_strVersion
        Exit Sub
    Else
        'Bereich einlesen
        Set g_varSelection2 = Selection
        'Check ob ganze Spalte selektiert ist
        If g_varSelection2.Rows.Count = 1048576 Then
            Call BereichAnpassen(g_varSelection2, 2)
        End If

        'GUI beschriften
        lblArea2.Caption = Replace(g_varSelection2.AddressLocal, "$", "")
            'Beschriftung in Tab "Bereinigung"
            lblCheckArea2.Caption = lblArea2.Caption
            lblAnzahlLeer.Caption = "(?)"
            lblAnzahlSteuer.Caption = "(?)"
            lblAnzahlGesch.Caption = "(?)"
            'Beschriftung im Tab 'Visualisierung"
            Call modGUItexte.TextButtonHervorhebungenLoeschen(g_strSprache)
        'GUI aktualisieren
        Call ButtonsAktivieren
        'Start-Button aktualisieren
        blnArea2 = True
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub AreasAuswaehlenBeide()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If Selection.Areas.Count <> 2 Then
        MsgBox modINFOtexte.MsgBoxText(g_strSprache, "MA2"), vbExclamation, g_strTool & " " & g_strVersion
        Exit Sub
    Else
        'Bereiche einlesen
        Set g_varSelection1 = Selection.Areas(1)
        Set g_varSelection2 = Selection.Areas(2)
        'Check ob ganze Spalten selektiert sind
        If g_varSelection1.Rows.Count = 1048576 Then
            Call BereichAnpassen(g_varSelection1, 1)
        End If
        If g_varSelection2.Rows.Count = 1048576 Then
            Call BereichAnpassen(g_varSelection2, 2)
        End If

        'GUI beschriften
        lblArea1.Caption = Replace(g_varSelection1.AddressLocal, "$", "")
        lblArea2.Caption = Replace(g_varSelection2.AddressLocal, "$", "")
            'Beschriftung in Tab "Bereinigung"
            lblCheckArea1.Caption = lblArea1.Caption
            lblCheckArea2.Caption = lblArea2.Caption
            lblAnzahlLeer.Caption = "(?)"
            lblAnzahlSteuer.Caption = "(?)"
            lblAnzahlGesch.Caption = "(?)"
            'Beschriftung im Tab 'Visualisierung"
            Call modGUItexte.TextButtonHervorhebungenLoeschen(g_strSprache)
        'GUI aktualisieren
        Call ButtonsAktivieren
        'Start-Button aktualisieren
        blnArea1 = True
        blnArea2 = True
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub ButtonsAktivieren()
    btnBereicheAnzeigen.Enabled = True
    btnSchnellCheck.Enabled = True
    btnSchnellCheck.BackColor = &HC0C000       'grün
    btnCheckAll.Enabled = True
    btnCheckAll.BackColor = &HFFC0C0   'blau
    btnCheckLeer.Enabled = True
    btnCheckLeer.BackColor = &HFFC0C0   'blau
    btnCheckSteuer.Enabled = True
    btnCheckSteuer.BackColor = &HFFC0C0   'blau
    btnCheckGesch.Enabled = True
    btnCheckGesch.BackColor = &HFFC0C0   'blau
    btnFixAll.Enabled = True
    btnFixAll.BackColor = &HFFC0C0   'blau
    btnFixLeer.Enabled = True
    btnFixLeer.BackColor = &HFFC0C0   'blau
    btnFixSteuer.Enabled = True
    btnFixSteuer.BackColor = &HFFC0C0   'blau
    btnFixGesch.Enabled = True
    btnFixGesch.BackColor = &HFFC0C0   'blau
End Sub

Private Sub BereichAnpassen(ByRef rngBereich As Range, ByVal byteNr As Byte)
    'Wenn ganze Zeilen selektiert sind kürzen bis zur letzten gefüllten Zeile
    
    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Zähler zurücksetzen
    lngMaxAnzahlZeilenBereich = 0
    
    'Längste Spalte des Bereichs ermitteln
    For j = 1 To rngBereich.Columns.Count
        If Cells(rngBereich.Rows.Count, rngBereich.Column + j - 1).End(xlUp).Row > lngMaxAnzahlZeilenBereich Then
            If IsEmpty(Cells(1048576, rngBereich.Column + j - 1)) Then
                lngMaxAnzahlZeilenBereich = Cells(rngBereich.Rows.Count, rngBereich.Column + j - 1).End(xlUp).Row
            Else
                lngMaxAnzahlZeilenBereich = 1048576
            End If
        End If
    Next j
            
    'Bereich anpassen, kürzen bis zur letzten Zeile mit Daten
    Set rngBereich = Range( _
        Cells(1, rngBereich.Column), _
        Cells(lngMaxAnzahlZeilenBereich, rngBereich.Column + rngBereich.Columns.Count - 1))
    
    'Selektion auf dem Tabellenblatt anpassen
    Call AreasSelektieren
        
    'Hinweis anzeigen
    MsgBox (modINFOtexte.MsgBoxText(g_strSprache, "MB1") & byteNr & _
        modINFOtexte.MsgBoxText(g_strSprache, "MB2") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MB3") & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MB4") & lngMaxAnzahlZeilenBereich), _
        vbInformation, g_strTool & " " & g_strVersion
        
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub AngrenzenderBereich()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If Selection.Areas.Count > 2 Then 'Wenn zu viele Bereiche selektiert sind
        MsgBox "Für diese Funktion dürfen maximal zwei Bereiche selektiert sein", vbExclamation, g_strTool & " " & g_strVersion
        Exit Sub
    ElseIf Selection.Areas.Count = 2 Then 'Wenn zwei Bereiche selektiert sind
        Set g_varSelection = Union(Selection.Areas(1).CurrentRegion, Selection.Areas(2).CurrentRegion)
        g_varSelection.Select
    Else
        'Angrenzenden Bereich selektieren
        Set g_varSelection = Selection.CurrentRegion
        g_varSelection.Select
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub LinieZeichnen(ByVal rngA As Range, ByVal rngE As Range, ByRef strName As String)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Linie zeichnen und formatieren
    Set shp = ActiveSheet.Shapes.AddLine(rngA.Left + rngA.Width, rngA.Top + rngA.Height / 2, _
                        rngE.Left, rngE.Top + rngE.Height / 2)
    shp.line.ForeColor.RGB = &HC000C0 'lila
    shp.line.Weight = 2
    shp.Name = strName
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub LinieZeichnen2()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Sanduhr neu starten
    g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "CC0") '"Linien zeichnen"
    g_strSanduhrNummer = ""
    g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "CC1") '"Linien zwischen den Wertpaaren zeichnen"
    'Fortschrittsbalken zurücksetzen
    Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    'Stückelung des Balkens berechnen
    g_dblBalkenAnteil = 100 / UBound(varAuswahl1)
    
    'Zähler der Anzahl Linien des aktuellen Vergleichs zurücksetzen
    lngAnzahlLinienAktuell = 0
    
    For i = 1 To UBound(varAuswahl1)
        For j = 1 To UBound(varAuswahl2)
            If varAuswahl1(i, 3) = varAuswahl2(j, 3) And varAuswahl1(i, 3) <> "" Then
                Call LinieZeichnen(Range(varAuswahl1(i, 2)), Range(varAuswahl2(j, 2)), "xlPairC" & "-" & strTimestamp & "-" & CStr(i))
                'Button aktivieren
                btnLinienLoeschenAktuelle.Enabled = True
                btnLinienLoeschenAktuelle.BackColor = &HFFC0C0          'blau
                'Zähler der Anzahl Linien des aktuellen Vergleichs hochzählen
                lngAnzahlLinienAktuell = lngAnzahlLinienAktuell + 1
            End If
        Next j
        
        'Sanduhr aktualisieren
        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
    Next i

    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub Linienloeschen(strLoeschtyp As String)
    
    On Error Resume Next
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Festlegen welche Linien gelöscht werden sollen
    Select Case strLoeschtyp
        Case "aktuelle"
            'Sanduhr neu starten
            g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "DD0") '"Linien löschen"
            g_strSanduhrNummer = ""
            g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "DD1") '"Alle Linien des aktuellen Vergleichs löschen"
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / ActiveSheet.Shapes.Count 'Ungefähre Anzahl, zählt alle Objekte auf dem Tabellenblatt
            
            'Alle Linien dieses Tools auf dem Tabellenblatt löschen
            For Each shp In ActiveSheet.Shapes
                If shp.Type = 9 And shp.Name Like "xlPairC*" & strTimestamp & "*" Then
                    shp.Delete
                End If
      
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
                
            Next
            
        Case "alle"
            'Sanduhr neu starten
            g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "DD0") '"Linien löschen"
            g_strSanduhrNummer = ""
            g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "DD2") '"Alle Linien dieses Tools löschen"
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / ActiveSheet.Shapes.Count 'Ungefähre Anzahl, zählt alle Objekte auf dem Tabellenblatt
            
            'Alle Linien dieses Tools auf dem Tabellenblatt löschen
            For Each shp In ActiveSheet.Shapes
                If shp.Type = 9 And shp.Name Like "xlPairC*" Then
                    shp.Delete
                End If
      
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
                
            Next
    End Select
    
    lngAnzahlLinienAktuell = 0 'Zähler zurücksetzen
    
    'Button deaktivieren
    btnLinienLoeschenAktuelle.Enabled = False
    btnLinienLoeschenAktuelle.BackColor = &H8000000F    'grau
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    On Error GoTo 0
End Sub

Private Sub HervorhebungLoeschen()

    On Error Resume Next
    
    g_varSelection1.ClearFormats
    g_varSelection2.ClearFormats
    
    On Error GoTo 0
End Sub

Private Sub EinzelwerteHervorheben()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Bereich 1
        If lngE1 > 0 Then
        
            'Sanduhr neu starten
            g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "EE0") '"Visualisieren"
            g_strSanduhrNummer = "[1/2]"
            g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "EE1") '"Einzelwerte des ersten Bereichs hervorheben"
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / UBound(varAuswahl1)
        
            For i = 1 To UBound(varAuswahl1, 1)
                If varAuswahl1(i, 1) <> "#areac#" And varAuswahl1(i, 1) <> "" Then
                    For j = Range(varAuswahl1(i, 2)).Column To _
                            Range(varAuswahl1(i, 2)).Column - g_varSelection1.Columns.Count + 1 Step -1
                        Cells(Range(varAuswahl1(i, 2)).Row, j).Interior.Color = &H80FF80    'grün
                    Next j
                End If
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
            Next i
        End If
    
    'Bereich 2
        If lngE2 > 0 Then
        
            'Sanduhr neu starten
            g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "EE0") '"Visualisieren"
            g_strSanduhrNummer = "[2/2]"
            g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "EE2") '"Einzelwerte des zweiten Bereichs hervorheben"
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / UBound(varAuswahl2)
        
            For i = 1 To UBound(varAuswahl2, 1)
                If varAuswahl2(i, 1) <> "#areac#" And varAuswahl2(i, 1) <> "" Then
                    For j = Range(varAuswahl2(i, 2)).Column To _
                            Range(varAuswahl2(i, 2)).Column + g_varSelection2.Columns.Count - 1 Step 1
                        Cells(Range(varAuswahl2(i, 2)).Row, j).Interior.Color = &H80FF80    'grün
                    Next j
                End If
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
            Next i
        End If
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub PaareHervorheben()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If lngP > 0 Then
    
        'Sanduhr einblenden
        Call SanduhrEinblenden
    
        'Bereich 1
            'Sanduhr neu starten
            g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "EE0") '"Visualisieren"
            g_strSanduhrNummer = "[1/2]"
            g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "EE3") '"Paare des ersten Bereichs hervorheben"
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / UBound(varAuswahl1)
                
            For i = 1 To UBound(varAuswahl1, 1)
                If varAuswahl1(i, 1) = "#areac#" Then
                    For j = Range(varAuswahl1(i, 2)).Column To _
                            Range(varAuswahl1(i, 2)).Column - g_varSelection1.Columns.Count + 1 Step -1
                        Cells(Range(varAuswahl1(i, 2)).Row, j).Interior.Color = &H80FFFF 'gelb
                    Next j
                End If
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
            Next i
            
        'Bereich 2
            'Sanduhr neu starten
            g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "EE0") '"Visualisieren"
            g_strSanduhrNummer = "[2/2]"
            g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "EE4") '"Paare des zweiten Bereichs hervorheben"
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / UBound(varAuswahl2)
            
            For i = 1 To UBound(varAuswahl2, 1)
                If varAuswahl2(i, 1) = "#areac#" Then
                    For j = Range(varAuswahl2(i, 2)).Column To _
                            Range(varAuswahl2(i, 2)).Column + g_varSelection2.Columns.Count - 1 Step 1
                        Cells(Range(varAuswahl2(i, 2)).Row, j).Interior.Color = &H80FFFF 'gelb
                    Next j
                End If
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
            Next i
        
        'Sanduhr ausblenden
        Call SanduhrAusblenden
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub AreasSelektieren()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If g_varSelection1 Is Nothing And g_varSelection2 Is Nothing Then
        Exit Sub
    ElseIf g_varSelection1 Is Nothing Then
        g_varSelection2.Select
    ElseIf g_varSelection2 Is Nothing Then
        g_varSelection1.Select
    Else
        'Beide ausgewählten Bereiche auf dem Tabellenblatt selektieren
        Union(g_varSelection1, g_varSelection2).Select
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub ButtonsCheck()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Buttons aktivieren wenn beide Bereiche ausgewählt
    If blnArea1 And blnArea2 Then
        btnStart.Enabled = True
        btnStart.BackColor = &HC0C000       'grün
    Else
        btnStart.Enabled = False
        btnStart.BackColor = &HC0&          'rot
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub CheckAdresseEinlesen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If g_varSelection1 Is Nothing Then
        Set rngCheck = Range(g_varSelection2.AddressLocal)
    ElseIf g_varSelection2 Is Nothing Then
        Set rngCheck = Range(g_varSelection1.AddressLocal)
    Else
        Set rngCheck = Union(Range(g_varSelection1.AddressLocal), Range(g_varSelection2.AddressLocal))
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub CheckAll(strType As String)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Adresse des selektierten Bereichs in Variable einlesen
    Call CheckAdresseEinlesen
    
    'Variablen zurücksetzen
    blnCheckFund = False
    blnCheckLeer = False
    blnCheckSteuer = False
    blnCheckGesch = False
    byteCheckSumme = 0
    lngAnzLeer = 0
    lngAnzSteuer = 0
    lngAnzGesch = 0
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Sanduhr neu starten
    g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "FF0") '"Daten prüfen"
    g_strSanduhrNummer = ""
    g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "FF1") '"Auf alle kritischen Zeichen prüfen"
    'Fortschrittsbalken zurücksetzen
    Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    'Stückelung des Balkens berechnen
    g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
    
    'Prüfen ob der ausgewählte Bereich unnötige Leerzeichen oder nicht druckbare Zeichen enthält
    For Each varCheckZelle In rngCheck
        If varCheckZelle <> Application.WorksheetFunction.Trim(varCheckZelle.Value) Then 'Check nach unnötigen Leerzeichen
            If blnCheckLeer = False Then
                byteCheckSumme = byteCheckSumme + 1
                blnCheckLeer = True
                blnCheckFund = True 'für Schnell-Check
            End If
            lngAnzLeer = lngAnzLeer + 1 'Zähler hochsetzen
        End If
        For lngCheck = 1 To Len(varCheckZelle)
             strCheckZeichen = Mid(varCheckZelle, lngCheck, 1)
              Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                  Case 1 To 31, 127, 129, 141, 143, 144, 157 'Steuerzeichen gefunden
                    If blnCheckSteuer = False Then
                        byteCheckSumme = byteCheckSumme + 2
                        blnCheckSteuer = True
                        blnCheckFund = True 'für Schnell-Check
                    End If
                    lngAnzSteuer = lngAnzSteuer + 1 'Zähler hochsetzen
                    Exit For
                End Select
        Next lngCheck
        For lngCheck = 1 To Len(varCheckZelle)
             strCheckZeichen = Mid(varCheckZelle, lngCheck, 1)
              Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                  Case 160 'geschütztes Leerzeichen gefunden
                    If blnCheckGesch = False Then
                        byteCheckSumme = byteCheckSumme + 4
                        blnCheckGesch = True
                        blnCheckFund = True 'für Schnell-Check
                    End If
                    lngAnzGesch = lngAnzGesch + 1 'Zähler hochsetzen
                    Exit For
                End Select
        Next lngCheck
        
        'Sanduhr aktualisieren
        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
        'Vorzeitiger Abbruch beim Schnell-Check wenn ein unerwünschtes Zeichen gefunden wurden
        If strType = "schnell" And blnCheckFund = True Then
            Exit For 'Schleife verlassen
        End If
    Next varCheckZelle
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    'Check-Summe auswerten und Info ausgeben
    If strType = "alles" Then
        'Beschriftung aktualisieren
        lblAnzahlLeer.Caption = lngAnzLeer
        lblAnzahlSteuer.Caption = lngAnzSteuer
        lblAnzahlGesch.Caption = lngAnzGesch
        
        Select Case byteCheckSumme
            Case 1 'unnötige Leerzeichen
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB1"), byteCheckSumme)
            Case 2 'Steuerzeichen
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB2"), byteCheckSumme)
            Case 3 'unnötige Leerzeichen und Steuerzeichen
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB3"), byteCheckSumme)
            Case 4 'geschützte Leerzeichen
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB4"), byteCheckSumme)
            Case 5 'unnötige Leerzeichen und geschützte Leerzeichen
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB5"), byteCheckSumme)
            Case 6 'Steuerzeichen und geschützte Leerzeichen
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB6"), byteCheckSumme)
            Case 7 'unnötige Leerzeichen, Steuerzeichen und geschützte Leerzeichen
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB7"), byteCheckSumme)
            Case Else 'alle Zellen sauber
                Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB8"), byteCheckSumme)
        End Select
    Else 'wenn Schnell-Scheck
        If blnCheckFund = True Then
            Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB9"), byteCheckSumme)
        Else
            Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB8"), byteCheckSumme)
        End If
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub CheckInfo(strText As String, byteCheckSumme As Byte)

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If byteCheckSumme = 0 Then
        MsgBox strText, vbOKOnly, modINFOtexte.MsgBoxText(g_strSprache, "MC1")
    Else
        MsgBox (modINFOtexte.MsgBoxText(g_strSprache, "MC2") & strText & _
            modINFOtexte.MsgBoxText(g_strSprache, "MC3") & vbNewLine & vbNewLine & _
            modINFOtexte.MsgBoxText(g_strSprache, "MC4")), _
            vbExclamation + vbOKOnly, modINFOtexte.MsgBoxText(g_strSprache, "MC1")
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub CheckUnnoetigeLeerzeichen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Adresse des selektierten Bereichs in Variable einlesen
    Call CheckAdresseEinlesen
    
    'Variablen zurücksetzen
    byteCheckSumme = 0
    lngFunde = 0
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Sanduhr neu starten
    g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "FF0") '"Daten prüfen"
    g_strSanduhrNummer = ""
    g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "FF2") '"Auf unnötige Leerzeichen prüfen"
    'Fortschrittsbalken zurücksetzen
    Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    'Stückelung des Balkens berechnen
    g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
    
    'Prüfen ob der ausgewählte Bereich unnötige Leerzeichen enthält
    For Each varCheckZelle In rngCheck
        If varCheckZelle <> Application.WorksheetFunction.Trim(varCheckZelle.Value) Then 'Check nach unnötigen Leerzeichen
            lngFunde = lngFunde + 1
            byteCheckSumme = 1
        End If
        
        'Sanduhr aktualisieren
        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next varCheckZelle
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    'Anzahl Fundstellen anzeigen
    lblAnzahlLeer.Caption = lngFunde
    
    'Info ausgeben
    If byteCheckSumme > 0 Then
        Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB1"), byteCheckSumme)
    Else
        Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB10"), byteCheckSumme)
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub CheckSteuerzeichen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Adresse des selektierten Bereichs in Variable einlesen
    Call CheckAdresseEinlesen
    
    'Variablen zurücksetzen
    byteCheckSumme = 0
    lngFunde = 0
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Sanduhr neu starten
    g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "FF0") '"Daten prüfen"
    g_strSanduhrNummer = ""
    g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "FF3") '"Auf Steuerzeichen prüfen"
    'Fortschrittsbalken zurücksetzen
    Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    'Stückelung des Balkens berechnen
    g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
    
    'Prüfen ob der ausgewählte Bereich Steuerzeichen enthält
    For Each varCheckZelle In rngCheck
        For lngCheck = 1 To Len(varCheckZelle)
            strCheckZeichen = Mid(varCheckZelle, lngCheck, 1)
            Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                Case 1 To 31, 127, 129, 141, 143, 144, 157 'Steuerzeichen gefunden
                    lngFunde = lngFunde + 1 'Zähler hochzählen
                    byteCheckSumme = 2
                    Exit For
            End Select
        Next lngCheck
        
        'Sanduhr aktualisieren
        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next varCheckZelle
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    'Anzahl Fundstellen anzeigen
    lblAnzahlSteuer.Caption = lngFunde
    
    'Info ausgeben
    If byteCheckSumme > 0 Then
        Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB2"), byteCheckSumme)
    Else
        Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB11"), byteCheckSumme)
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub CheckGeschuetzteLeerzeichen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Adresse des selektierten Bereichs in Variable einlesen
    Call CheckAdresseEinlesen
    
    'Variablen zurücksetzen
    byteCheckSumme = 0
    lngFunde = 0
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Sanduhr neu starten
    g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "FF0") '"Daten prüfen"
    g_strSanduhrNummer = ""
    g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "FF4") '"Auf geschützte Leerzeichen prüfen"
    'Fortschrittsbalken zurücksetzen
    Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    'Stückelung des Balkens berechnen
    g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
    
    'Prüfen ob der ausgewählte Bereich geschützte Leerzeichen enthält
    For Each varCheckZelle In rngCheck
        For lngCheck = 1 To Len(varCheckZelle)
            strCheckZeichen = Mid(varCheckZelle, lngCheck, 1)
            Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                Case 160 'geschütztes Leerzeichen gefunden
                    lngFunde = lngFunde + 1 'Zähler hochzählen
                    byteCheckSumme = 4
                    Exit For
            End Select
        Next lngCheck
        
        'Sanduhr aktualisieren
        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next varCheckZelle
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    'Anzahl Fundstellen anzeigen
    lblAnzahlGesch.Caption = lngFunde
    
    'Info ausgeben
    If byteCheckSumme > 0 Then
        Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB4"), byteCheckSumme)
    Else
        Call CheckInfo(modINFOtexte.MsgBoxText(g_strSprache, "SB12"), byteCheckSumme)
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub FixAll()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sicherheitsabfrage
    If MsgBox(modINFOtexte.MsgBoxText(g_strSprache, "MD1") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD2") & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD3") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD4") & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD5") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD6") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD7"), vbExclamation + vbOKCancel, _
        modINFOtexte.MsgBoxText(g_strSprache, "MD8")) = vbOK Then
        
        'Adresse des selektierten Bereichs in Variable einlesen
        Call CheckAdresseEinlesen
        
        'Sanduhr einblenden
        Call SanduhrEinblenden
        
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "GG0") '"Daten bereinigen"
        g_strSanduhrNummer = ""
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "GG1") '"Alle kritischen Zeichen entfernen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
        
        'Alle Zellen durchlaufen
        For Each varCheckZelle In rngCheck
            For lngCheck = 1 To Len(varCheckZelle)
                strCheckZeichen = Mid(varCheckZelle, lngCheck, 1)
                Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                    Case 1 To 31, 127, 129, 141, 143, 144, 157 'Steuerzeichen gefunden
                        varCheckZelle.Value = Application.WorksheetFunction _
                            .Replace(varCheckZelle.Value, lngCheck, 1, Chr(9)) 'Steuerzeichen durch horizontalen Tab ersetzen
                    Case 160 'Geschütztes Leerzeichen gefunden
                        varCheckZelle.Value = Application.WorksheetFunction _
                            .Replace(varCheckZelle.Value, lngCheck, 1, Chr(32)) 'Geschütztes Leerzeichen durch normales ersetzen
                End Select
            Next lngCheck
            
            'Alle unerwünschten Zeichen entfernen
            varCheckZelle.Value = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(varCheckZelle.Value))
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next varCheckZelle
        
        'Sanduhr ausblenden
        Call SanduhrAusblenden
        
        'Zähler zurücksetzen und Anzahl Fundstellen auf null setzen
        lngAnzLeer = 0
        lngAnzSteuer = 0
        lngAnzGesch = 0
        lblAnzahlLeer.Caption = lngAnzLeer
        lblAnzahlSteuer.Caption = lngAnzSteuer
        lblAnzahlGesch.Caption = lngAnzGesch
    
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub FixUnnoetigeLeerzeichen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sicherheitsabfrage
    If MsgBox(modINFOtexte.MsgBoxText(g_strSprache, "MD9") & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD10") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD7"), _
        vbExclamation + vbOKCancel, modINFOtexte.MsgBoxText(g_strSprache, "MD11")) = vbOK Then

        'Adresse des selektierten Bereichs in Variable einlesen
        Call CheckAdresseEinlesen
        
        'Sanduhr einblenden
        Call SanduhrEinblenden
        
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "GG0") '"Daten bereinigen"
        g_strSanduhrNummer = ""
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "GG2") '"Unnötige Leerzeichen entfernen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
        
        'Alle Zellen durchlaufen
        For Each varCheckZelle In rngCheck
            varCheckZelle.Value = Application.WorksheetFunction.Trim(varCheckZelle.Value) 'Unnötige Leerzeichen entfernen
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next varCheckZelle
        
        'Sanduhr ausblenden
        Call SanduhrAusblenden
        
        lngFunde = 0 'Zähler zurücksetzen
        
        'Anzahl Fundstellen auf null setzen
        lblAnzahlLeer.Caption = lngFunde
    
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub FixSteuerzeichen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sicherheitsabfrage
    If MsgBox(modINFOtexte.MsgBoxText(g_strSprache, "MD12") & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD13") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD7"), _
        vbExclamation + vbOKCancel, modINFOtexte.MsgBoxText(g_strSprache, "MD14")) = vbOK Then

        'Adresse des selektierten Bereichs in Variable einlesen
        Call CheckAdresseEinlesen
        
        'Sanduhr einblenden
        Call SanduhrEinblenden
        
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "GG0") '"Daten bereinigen"
        g_strSanduhrNummer = ""
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "GG3") '"Steuerzeichen entfernen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
        
        'Alle Zellen durchlaufen
        For Each varCheckZelle In rngCheck
            lngFunde = 0 'Markierung zurücksetzen
            
            For lngCheck = 1 To Len(varCheckZelle)
                strCheckZeichen = Mid(varCheckZelle, lngCheck, 1)
                Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                    Case 1 To 31, 127, 129, 141, 143, 144, 157 'Steuerzeichen gefunden
                        varCheckZelle.Value = Application.WorksheetFunction _
                            .Replace(varCheckZelle.Value, lngCheck, 1, Chr(9)) 'Steuerzeichen durch horizontalen Tab ersetzen
                
                    lngFunde = 1 'Markieren, dass die Zelle bereinigt werden muss
                End Select
            Next lngCheck
            
            If lngFunde = 1 Then
                varCheckZelle.Value = Application.WorksheetFunction.Clean(varCheckZelle.Value) 'Horizontalen Tab entfernen
            End If
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next varCheckZelle
        
        'Sanduhr ausblenden
        Call SanduhrAusblenden
        
        lngFunde = 0 'Zähler zurücksetzen
        
        'Anzahl Fundstellen auf null setzen
        lblAnzahlSteuer.Caption = lngFunde
    
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub FixGeschuetzteLeerzeichen()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sicherheitsabfrage
    If MsgBox(modINFOtexte.MsgBoxText(g_strSprache, "MD15") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "MD7"), _
        vbExclamation + vbOKCancel, modINFOtexte.MsgBoxText(g_strSprache, "MD16")) = vbOK Then
        
        'Adresse des selektierten Bereichs in Variable einlesen
        Call CheckAdresseEinlesen
        
        'Sanduhr einblenden
        Call SanduhrEinblenden
        
        'Sanduhr neu starten
        g_strSanduhrAktion = modSANDUHRtexte.SanduhrText(g_strSprache, "GG0") '"Daten bereinigen"
        g_strSanduhrNummer = ""
        g_strSanduhrSchritt = modSANDUHRtexte.SanduhrText(g_strSprache, "GG4") '"Geschützte Leerzeichen entfernen"
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / rngCheck.Count 'Anzahl der Zellen
        
        'Alle Zellen durchlaufen
        For Each varCheckZelle In rngCheck
            lngFunde = 0 'Markierung zurücksetzen
            
            For lngCheck = 1 To Len(varCheckZelle)
                strCheckZeichen = Mid(varCheckZelle, lngCheck, 1)
                Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                    Case 160 'Steuerzeichen gefunden
                        varCheckZelle.Value = Application.WorksheetFunction _
                            .Replace(varCheckZelle.Value, lngCheck, 1, Chr(32)) 'Geschütztes Leerzeichen durch normales ersetzen
                
                    lngFunde = 1 'Markieren, dass die Zelle bereinigt werden muss
                End Select
            Next lngCheck
            
            If lngFunde = 1 Then
                varCheckZelle.Value = Application.WorksheetFunction.Trim(varCheckZelle.Value) 'Unnötige Leerzeichen entfernen
            End If
            
            'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next varCheckZelle
        
        'Sanduhr ausblenden
        Call SanduhrAusblenden
        
        lngFunde = 0 'Zähler zurücksetzen
        
        'Anzahl Fundstellen auf null setzen
        lblAnzahlGesch.Caption = lngFunde
    
    End If
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub VergleichReset()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Buttons zurücksetzen
    btnAusgabe1.Enabled = False
    btnAusgabe1.BackColor = &H8000000F              'grau
    btnAusgabe2.Enabled = False
    btnAusgabe2.BackColor = &H8000000F              'grau
    btnHervorhebenEinzeln.Enabled = False
    btnHervorhebenEinzeln.BackColor = &H8000000F    'grau
    btnHervorhebenPaare.Enabled = False
    btnHervorhebenPaare.BackColor = &H8000000F      'grau
    btnHervorhebungLoeschen.Enabled = True
    btnHervorhebungLoeschen.BackColor = &HFFC0C0    'blau
    btnLinienZeichnen.Enabled = False
    btnLinienZeichnen.BackColor = &H8000000F        'grau
    btnLinienLoeschenAktuelle.Enabled = False
    btnLinienLoeschenAktuelle.BackColor = &H8000000F    'grau
    btnLinienLoeschenAlle.Enabled = True
    btnLinienLoeschenAlle.BackColor = &HFFC0C0          'blau

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub Reset(strLinienLoeschen As String)
    
    'Erste Seite aktivieren
    MultiPageGUI.Value = 0
    
    'GUI-Beschriftungen
    Call modGUItexte.Sprache(g_strSprache)
    lblAusgeben1Beide.Caption = ""
    lblAusgeben1Area2.Caption = ""
    lblAusgeben1Area1.Caption = ""
    lblAusgeben2Beide.Caption = ""
    lblAusgeben2Area2.Caption = ""
    lblAusgeben2Area1.Caption = ""
    lblArea1.Caption = ""
    lblArea2.Caption = ""
    lblCheckArea1.Caption = "-"
    lblCheckArea2.Caption = "-"
    lblAnzahlLeer.Caption = "(?)"
    lblAnzahlSteuer.Caption = "(?)"
    lblAnzahlGesch.Caption = "(?)"
    'Beschriftung im Tab 'Visualisierung"
    Call modGUItexte.TextButtonHervorhebungenLoeschen(g_strSprache)

    'Areas zurücksetzen
    blnArea1 = False
    blnArea2 = False
    
    'Buttons setzen
    Call VergleichReset
    btnSchnellCheck.Enabled = False
    btnSchnellCheck.BackColor = &HC0&        'rot
    btnCheckAll.Enabled = False
    btnCheckAll.BackColor = &H8000000F    'grau
    btnCheckLeer.Enabled = False
    btnCheckLeer.BackColor = &H8000000F    'grau
    btnCheckSteuer.Enabled = False
    btnCheckSteuer.BackColor = &H8000000F    'grau
    btnCheckGesch.Enabled = False
    btnCheckGesch.BackColor = &H8000000F    'grau
    btnFixAll.Enabled = False
    btnFixAll.BackColor = &H8000000F    'grau
    btnFixLeer.Enabled = False
    btnFixLeer.BackColor = &H8000000F    'grau
    btnFixSteuer.Enabled = False
    btnFixSteuer.BackColor = &H8000000F    'grau
    btnFixGesch.Enabled = False
    btnFixGesch.BackColor = &H8000000F    'grau
    btnStart.Enabled = False
    btnStart.BackColor = &HC0&              'rot
    btnArea1.BackColor = &HC0C000           'grün
    btnArea2.BackColor = &HC0C000           'grün
    btnAreaBeide.BackColor = &HC0C000       'grün
    btnCurrentRegion.BackColor = &HFFC0C0   'blau
    btnBereicheAnzeigen.Enabled = False
    
    'Evtl. vorhandene Formatierungen und Linien löschen
    If strLinienLoeschen = "aktuelle" Then
        Call HervorhebungLoeschen
        Call Linienloeschen(strLinienLoeschen)
    End If
    
    On Error Resume Next
    
    'Variablen zurücksetzen
    Set g_varSelection1 = Nothing
    Set g_varSelection2 = Nothing
    Erase varAuswahl1
    Erase varAuswahl2
    
    On Error GoTo 0
    
End Sub

Private Sub AnleitungAnzeigen() 'Öffnen bzw. schließen des Popups4

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If frmAnleitung.Visible = False Then
        Load frmAnleitung
        frmAnleitung.StartUpPosition = 2 'Zentriert auf dem gesamten Bildschirm
        frmAnleitung.Show
    Else
        Unload frmAnleitung
    End If

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub FeaturesAnzeigen() 'Öffnen bzw. schließen des Popups

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If frmVersionshinweise.Visible = False Then
        Load frmVersionshinweise
        frmVersionshinweise.StartUpPosition = 2 'Zentriert auf dem gesamten Bildschirm
        frmVersionshinweise.Show
    Else
        Unload frmVersionshinweise
    End If

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub DisclaimerAnzeigen() 'Öffnen bzw. schließen des Popups

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If frmDisclaimer.Visible = False Then
        Load frmDisclaimer
        frmDisclaimer.StartUpPosition = 2 'Zentriert auf dem gesamten Bildschirm
        frmDisclaimer.Show
    Else
        Unload frmDisclaimer
    End If

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub Tooltips()

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    If checkboxTooltip.Value = True Then
        Call modTOOLTIPtexte.TooltipsON
    Else
        Call modTOOLTIPtexte.TooltipsOFF
    End If

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub


'Sanduhr
'-------

Private Sub SanduhrEinblenden() 'Sanduhr einblenden

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    Load frmSanduhr
    With frmSanduhr
        .StartUpPosition = 1 'Zentriert im Element, zu dem das UserForm-Objekt gehört
        .Show
    End With
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Private Sub SanduhrAusblenden() 'Sanduhr ausblenden
    Unload frmSanduhr
End Sub

Public Sub FortschrittsbalkenReset(strAktion As String, strNr As String, strSchritt As String) 'Fortschrittsbalken der Sanduhr zurücksetzen

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    With frmSanduhr
        .Caption = strAktion
        .lblFortschrittBalken.Width = 0 'Breite des Balkens
        .lblFortschrittProzent.Caption = "" 'Anzeige des Prozentanteils
        .lblFortschrittNr.Caption = strNr 'Nummer des Einzelschritts
        .lblFortschrittSchritt.Caption = strSchritt 'Einzelschritt
    End With
    g_dblBalkenAktuell = 0 'Länge des Balkens zurücksetzen
    DoEvents 'neu zeichnen
    
Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub

Public Sub FortschrittsbalkenAktualisieren(dblProzent As Double) 'Fortschrittsbalken der Sanduhr aktualisieren

    'Falls ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    With frmSanduhr
        .lblFortschrittBalken.Width = CInt(dblProzent) 'Breite des Balkens
        .lblFortschrittProzent.Caption = CInt(dblProzent) & "%" 'Anzeige des Prozentanteils
    End With
    DoEvents 'neu zeichnen

Exit Sub
    
Fehlerbehandlung:
    Call FehlerSammelmeldung
End Sub


'Fehlerbehandlung
'----------------

Private Sub FehlerSammelmeldung()
    MsgBox (modINFOtexte.MsgBoxText(g_strSprache, "SC1") & vbNewLine & vbNewLine & _
        modINFOtexte.MsgBoxText(g_strSprache, "SC2") & Err.Number & vbNewLine & vbNewLine & _
        Err.Description), vbCritical, g_strTool & " " & g_strVersion & modINFOtexte.MsgBoxText(g_strSprache, "SC3")
End Sub


'Hyperlinks
'----------

Private Sub SourceCodeURL()
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:="https://github.com/MarcoKrapf/xlPairCompare"
    On Error GoTo 0
End Sub

Private Sub SpendenLinkURLaufrufen()
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:="http://www.ghfkh.de/"
    On Error GoTo 0
End Sub


'E-Mails
'-------

Private Sub eMail() 'Feedback E-Mail
    On Error Resume Next
        Set objMail = CreateObject("Shell.Application")
        objMail.ShellExecute "mailto:" & "excel@marco-krapf.de" _
            & "&subject=" & "Feedback: " & g_strTool & " " & g_strVersion & " / " _
            & Application.OperatingSystem & " / Excel-Version " & Application.Version
    On Error GoTo 0
End Sub


'Klicks
'------

Private Sub btnStart_Click()
    Call AreasSelektieren
    Call VergleichReset
    strTimestamp = CStr(Now) 'Zeitpunkt des Starts als String
    Call Start(g_varSelection1, g_varSelection2)
End Sub

Private Sub btnCurrentRegion_Click()
    Call AngrenzenderBereich
    Call modGUItexte.TexteDynamisch
End Sub

Private Sub btnArea1_Click()
    Call AreaAuswaehlen(1)
    Call modGUItexte.TexteDynamisch
End Sub

Private Sub btnArea2_Click()
    Call AreaAuswaehlen(2)
    Call modGUItexte.TexteDynamisch
End Sub

Private Sub btnAreaBeide_Click()
    Call AreaAuswaehlen(3)
    Call modGUItexte.TexteDynamisch
End Sub

Private Sub btnBereicheAnzeigen_Click()
    Call AreasSelektieren
End Sub

Private Sub btnHervorhebenEinzeln_Click()
    Call EinzelwerteHervorheben
End Sub

Private Sub btnHervorhebenPaare_Click()
    Call PaareHervorheben
End Sub

Private Sub btnHervorhebungLoeschen_Click()
    Call HervorhebungLoeschen
End Sub

Private Sub btnLinienZeichnen_Click()
    Call LinieZeichnen2
End Sub

Private Sub btnLinienLoeschenAktuelle_Click()
    Call Linienloeschen("aktuelle")
End Sub

Private Sub btnLinienLoeschenAlle_Click()
    Call Linienloeschen("alle")
End Sub

Private Sub btnAusgabe1_Click()
    Call WerteAusgeben(1, optAusg1Nur1, optAusg1Nur2)
End Sub

Private Sub btnAusgabe2_Click()
    Call WerteAusgeben(2, optAusg2Nur1, optAusg2Nur2)
End Sub

Private Sub btnSchnellCheck_Click()
    Call CheckAll("schnell")
End Sub

Private Sub btnCheckAll_Click()
    Call CheckAll("alles")
End Sub

Private Sub btnCheckLeer_Click()
    Call CheckUnnoetigeLeerzeichen
End Sub

Private Sub btnCheckSteuer_Click()
    Call CheckSteuerzeichen
End Sub

Private Sub btnCheckGesch_Click()
    Call CheckGeschuetzteLeerzeichen
End Sub

Private Sub btnFixAll_Click()
    Call FixAll
End Sub

Private Sub btnFixLeer_Click()
    Call FixUnnoetigeLeerzeichen
End Sub

Private Sub btnFixSteuer_Click()
    Call FixSteuerzeichen
End Sub

Private Sub btnFixGesch_Click()
    Call FixGeschuetzteLeerzeichen
End Sub

Private Sub btnAnleitung_Click()
    Call AnleitungAnzeigen
End Sub

Private Sub btnReset_Click()
    Call Reset("aktuelle")
End Sub

Private Sub btnFeatures_Click()
    Call FeaturesAnzeigen
End Sub

Private Sub btnDisclaimer_Click()
    Call DisclaimerAnzeigen
End Sub

Private Sub btnSourceCode_Click()
    Call SourceCodeURL
End Sub

Private Sub btnFeedback_Click()
    Call eMail
End Sub

Private Sub checkboxTooltip_Click()
    Call Tooltips
End Sub

Private Sub imgFlaggeDE_Click()
    g_strSprache = "DE"
    Call modGUItexte.Sprache(g_strSprache)
    If checkboxTooltip.Value = True Then Call modTOOLTIPtexte.TooltipsON
End Sub

Private Sub imgFlaggeEN_Click()
    g_strSprache = "EN"
    Call modGUItexte.Sprache(g_strSprache)
    If checkboxTooltip.Value = True Then Call modTOOLTIPtexte.TooltipsON
End Sub

Private Sub imgPrinz_Click()
    Call SpendenLinkURLaufrufen
End Sub

Private Sub lblSpendenLink_Click()
    Call SpendenLinkURLaufrufen
End Sub


'GUI initialisieren
'------------------

Private Sub UserForm_Initialize()
    'Startwerte setzen
    Call Reset("keine")
End Sub


'GUI schließen
'-------------

Private Sub UserForm_Terminate()
    Unload frmAnleitung
    Unload frmDisclaimer
    Unload frmVersionshinweise
    Unload frmSanduhr
    Unload frmGUI
End Sub
