Attribute VB_Name = "modGUItexte"
Option Explicit

'Modulbeschreibung:
'Anpassung der statischen und dynamischen GUI-Beschriftungen je nach gewählter Sprache
'-------------------------------------------------------------------------------------

Public Sub Sprache(strSprachwahl As String)

    Select Case strSprachwahl
        Case "DE"
            With frmGUI
                'Titel
                .Caption = g_strTool & " " & g_strVersion
                .lblTitel = g_strTool
                .checkboxTooltip.Caption = "Tooltips"
                'Buttons unten
                .btnReset.Caption = "Reset"
                .btnBereicheAnzeigen.Caption = "Selektierte Bereiche anzeigen"
                
                'MultiPage
                .MultiPageGUI.Pages(0).Caption = "Vergleich"
                .MultiPageGUI.Pages(1).Caption = "Visualisierung"
                .MultiPageGUI.Pages(2).Caption = "Ausgabe"
                .MultiPageGUI.Pages(3).Caption = "Bereinigung"
                .MultiPageGUI.Pages(4).Caption = "Info"
                .MultiPageGUI.Pages(5).Caption = "Spende"
                
                'Page "Vergleichen"
                .frameAreasVergleich.Caption = "Bereiche zu vergleichen"
                .lblAreaVergleich1.Caption = "Bereich 1:"
                .lblAreaVergleich2.Caption = "Bereich 2:"
                .btnArea1.Caption = "Bereich 1 einlesen"
                .btnArea2.Caption = "Bereich 2 einlesen"
                .btnAreaBeide.Caption = "Beide Bereiche einlesen"
                .btnSchnellCheck.Caption = "Auf kritische Zeichen prüfen"
                .btnCurrentRegion.Caption = "Angrenzende Zellen selektieren"
                .btnStart.Caption = "Vergleich starten"
                .frameVergleichsoptionen.Caption = "Optionen"
                .checkboxGrossKleinBuchstaben.Caption = "Groß-/Kleinbuchstaben ignorieren"
                .checkboxLeerzeichen.Caption = "Leerzeichen ignorieren"
                .frameAnzeige.Caption = "Visualisierung auf dem Tabellenblatt"
                .checkboxLinien.Caption = "Linien zwischen den Wertpaaren zeichnen"
                .checkboxHervorhebenEinzeln.Caption = "Einzelne Werte jedes Bereichs hervorheben"
                .checkboxHervorhebenPaare.Caption = "Wertpaare hervorheben"
                
                'Page "Hervorbehen"
                .btnHervorhebenEinzeln.Caption = "Einzelwerte hervorheben"
                .btnHervorhebenPaare.Caption = "Wertpaare hervorheben"
                Call TextButtonHervorhebungenLoeschen("DE")
                .btnLinienZeichnen.Caption = "Linien zwischen den Wertpaaren zeichnen"
                .btnLinienLoeschenAktuelle.Caption = "Linien des aktuellen Vergleichs löschen"
                .btnLinienLoeschenAlle.Caption = "Alle Linien auf diesem Tabellenblatt löschen"
                
                'Page "Ausgeben"
                .frameAusgeben1.Caption = "Ausgabeort"
                .optAusgNeu.Caption = "in neuem Tabellenblatt"
                .optAusgCursor.Caption = "ab aktueller Cursorposition"
                .frameAusgeben2.Caption = "Einzelwerte ausgeben"
                .optAusg1Beide.Caption = "beide Bereiche"
                .optAusg1Nur1.Caption = "nur den ersten Bereich"
                .optAusg1Nur2.Caption = "nur den zweiten Bereich"
                .btnAusgabe1.Caption = "Ausgeben"
                .frameAusgeben3.Caption = "Wertpaare ausgeben"
                .optAusg2Beide.Caption = "beide Bereiche"
                .optAusg2Nur1.Caption = "nur den ersten Bereich"
                .optAusg2Nur2.Caption = "nur den zweiten Bereich"
                .btnAusgabe2.Caption = "Ausgeben"
                
                'Page "Bereinigen"
                .frameBereinigen.Caption = "Daten prüfen und bereinigen"
                .lblCheckBereichText1.Caption = "Bereich 1:"
                .lblCheckBereichText2.Caption = "Bereich 2:"
                .btnCheckAll.Caption = "Auf alle kritschen Zeichen prüfen"
                .btnCheckLeer.Caption = "Auf unnötige Leerzeichen prüfen"
                .btnCheckSteuer.Caption = "Auf Steuerzeichen prüfen"
                .btnCheckGesch.Caption = "Auf geschützte Leerzeichen prüfen"
                .btnFixAll.Caption = "Alle kritschen Zeichen entfernen"
                .btnFixLeer.Caption = "Unnötige Leerzeichen entfernen"
                .btnFixSteuer.Caption = "Steuerzeichen entfernen"
                .btnFixGesch.Caption = "Geschützte Leerzeichen entfernen"
                .lblAnzahlText.Caption = "Anzahl Zellen"
                
                'Page "Info"
                .btnAnleitung.Caption = "Anleitung"
                .btnFeatures.Caption = "Features"
                .btnSourceCode.Caption = "Quellcode auf GitHub"
                .btnDisclaimer.Caption = "Nutzungsbedingungen"
                .btnFeedback.Caption = "Feedback"
                .lblInfo1.Caption = "xlPairCompare " & g_strVersion & " (Nov 2016)"
                .lblInfo2.Caption = "Autor: Marco Krapf - E-Mail: excel@marco-krapf.de"
                
                'Page "Spenden"
                .lblSpendenLink.Caption = "Info und Spende an die Stiftung 'Große Hilfe für kleine Helden'"
                .lblSpendenText.Caption = "Das Excel-Add-in 'xlPairCompare' wird privat entwickelt und unter " & _
                    "http://marco-krapf.de/excel/ kostenlos zum Download angeboten." & vbNewLine & vbNewLine & _
                    "Über eine kleine Spende an die Stiftung 'Große Hilfe für kleine Helden' für kranke Kinder " & _
                    "in der Region Heilbronn würde ich mich sehr freuen."
            End With
            
        Case "EN"
            With frmGUI
                'Titel
                .Caption = g_strTool & " " & g_strVersion
                .lblTitel = g_strTool
                .checkboxTooltip.Caption = "Tooltips"
                'Buttons unten
                .btnReset.Caption = "Reset"
                .btnBereicheAnzeigen.Caption = "Show selected areas"
                
                'MultiPage
                .MultiPageGUI.Pages(0).Caption = "Comparison"
                .MultiPageGUI.Pages(1).Caption = "Visualisation"
                .MultiPageGUI.Pages(2).Caption = "Output"
                .MultiPageGUI.Pages(3).Caption = "Cleanup"
                .MultiPageGUI.Pages(4).Caption = "Info"
                .MultiPageGUI.Pages(5).Caption = "Donation"
                
                'Page "Vergleichen"
                .frameAreasVergleich.Caption = "Areas to compare"
                .lblAreaVergleich1.Caption = "Area 1:"
                .lblAreaVergleich2.Caption = "Area 2:"
                .btnArea1.Caption = "Read in area 1"
                .btnArea2.Caption = "Read in area 2"
                .btnAreaBeide.Caption = "Read in both areas"
                .btnSchnellCheck.Caption = "Check for critical characters"
                .btnCurrentRegion.Caption = "Select adjacent cells"
                .btnStart.Caption = "Start comparison"
                .frameVergleichsoptionen.Caption = "Options"
                .checkboxGrossKleinBuchstaben.Caption = "Ignore lower/upper-case letters"
                .checkboxLeerzeichen.Caption = "Ignore spaces"
                .frameAnzeige.Caption = "Visualisation on the worksheet"
                .checkboxLinien.Caption = "Draw connection lines between value pairs"
                .checkboxHervorhebenEinzeln.Caption = "Highlight single values"
                .checkboxHervorhebenPaare.Caption = "Highlight value pairs"
                
                'Page "Hervorbehen"
                .btnHervorhebenEinzeln.Caption = "Highlight single values"
                .btnHervorhebenPaare.Caption = "Highlight value pairs"
                Call TextButtonHervorhebungenLoeschen("EN")
                .btnLinienZeichnen.Caption = "Draw connection lines between value pairs"
                .btnLinienLoeschenAktuelle.Caption = "Remove connection lines of the current run"
                .btnLinienLoeschenAlle.Caption = "Remove all connection lines in this worksheet"
                
                'Page "Ausgeben"
                .frameAusgeben1.Caption = "Output location"
                .optAusgNeu.Caption = "new worksheet"
                .optAusgCursor.Caption = "at current cursor position"
                .frameAusgeben2.Caption = "Print single values"
                .optAusg1Beide.Caption = "both areas"
                .optAusg1Nur1.Caption = "only the first area"
                .optAusg1Nur2.Caption = "only the second area"
                .btnAusgabe1.Caption = "Print"
                .frameAusgeben3.Caption = "Print value pairs"
                .optAusg2Beide.Caption = "both areas"
                .optAusg2Nur1.Caption = "only the first area"
                .optAusg2Nur2.Caption = "only the second area"
                .btnAusgabe2.Caption = "Print"
                
                'Page "Bereinigen"
                .frameBereinigen.Caption = "Data check and cleanup"
                .lblCheckBereichText1.Caption = "Area 1:"
                .lblCheckBereichText2.Caption = "Area 2:"
                .btnCheckAll.Caption = "Check for all critical characters"
                .btnCheckLeer.Caption = "Check for unnecessary spaces"
                .btnCheckSteuer.Caption = "Check for control characters"
                .btnCheckGesch.Caption = "Check for non-breaking spaces"
                .btnFixAll.Caption = "Remove all critical characters"
                .btnFixLeer.Caption = "Remove unnecessary spaces"
                .btnFixSteuer.Caption = "Remove control characters"
                .btnFixGesch.Caption = "Remove non-breaking spaces"
                .lblAnzahlText.Caption = "Number of cells"
                
                'Page "Info"
                .btnAnleitung.Caption = "User's guide"
                .btnFeatures.Caption = "Features"
                .btnSourceCode.Caption = "Source code on GitHub"
                .btnDisclaimer.Caption = "Terms of use"
                .btnFeedback.Caption = "Feedback"
                .lblInfo1.Caption = "xlPairCompare " & g_strVersion & " (Nov 2016)"
                .lblInfo2.Caption = "Author: Marco Krapf - email: excel@marco-krapf.de"
                
                'Page "Spenden"
'                .lblSpendeQRcode.Caption = "Donate online via SmartPhone now"

                .lblSpendenLink.Caption = "Info and donation to the foundation"
                .lblSpendenText.Caption = "This add-in is being developed and maintained with private effort " & _
                    "and provided for free download on http://marco-krapf.de/excel/" & vbNewLine & vbNewLine & _
                    "I would be very happy about a small donation to this foundation for sick children in the " & _
                    "region of Heilbronn/Germany."
            End With
        Case Else
    End Select
End Sub

Public Sub TexteDynamisch()

    On Error Resume Next
    
    With frmGUI
        .lblAusgeben1Beide.Caption = Replace(g_varSelection1.AddressLocal, "$", "") & _
                                    " / " & Replace(g_varSelection2.AddressLocal, "$", "")
        .lblAusgeben1Area1.Caption = Replace(g_varSelection1.AddressLocal, "$", "")
        .lblAusgeben1Area2.Caption = Replace(g_varSelection2.AddressLocal, "$", "")
        .lblAusgeben2Beide.Caption = Replace(g_varSelection1.AddressLocal, "$", "") & _
                                    " / " & Replace(g_varSelection2.AddressLocal, "$", "")
        .lblAusgeben2Area1.Caption = Replace(g_varSelection1.AddressLocal, "$", "")
        .lblAusgeben2Area2.Caption = Replace(g_varSelection2.AddressLocal, "$", "")
    End With
    
    On Error GoTo 0
End Sub

Public Sub TextButtonHervorhebungenLoeschen(strSprachwahl As String)
    Select Case strSprachwahl
        Case "DE"
            frmGUI.btnHervorhebungLoeschen.Caption = "Alle farblichen Hervorhebungen" & vbNewLine & _
                    "in den selektierten" & vbNewLine & "Bereichen löschen" & vbNewLine & vbNewLine & _
                    frmGUI.lblArea1.Caption & vbNewLine & frmGUI.lblArea2.Caption
        Case "EN"
            frmGUI.btnHervorhebungLoeschen.Caption = "Remove all highlightings" & vbNewLine & _
                    "in the selected areas" & vbNewLine & vbNewLine & _
                    frmGUI.lblArea1.Caption & vbNewLine & frmGUI.lblArea2.Caption
        Case Else
    End Select
End Sub
