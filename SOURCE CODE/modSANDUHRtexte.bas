Attribute VB_Name = "modSANDUHRtexte"
Option Explicit

'Modulbeschreibung:
'Rückgabe der Texte für die Sanduhr
'----------------------------------

Public Function SanduhrText(strSprachwahl As String, strText As String) As String
    Select Case strSprachwahl
        Case "DE"
            Select Case strText
                Case "AA0"
                    SanduhrText = "Paare finden"
                Case "AA1"
                    SanduhrText = "Ersten Bereich einlesen"
                Case "AA2"
                    SanduhrText = "Zweiten Bereich einlesen"
                Case "AA3"
                    SanduhrText = "Wertpaare suchen und als gefunden markieren"
                Case "AA4"
                    SanduhrText = "Werte des ersten Bereichs zählen"
                Case "AA5"
                    SanduhrText = "Werte des zweiten Bereichs zählen"
                
                Case "BB0"
                    SanduhrText = "Ausgeben"
                Case "BB1"
                    SanduhrText = "Ausgeben des ersten Bereichs"
                Case "BB2"
                    SanduhrText = "Ausgeben des zweiten Bereichs"
                    
                Case "CC0"
                    SanduhrText = "Linien zeichnen"
                Case "CC1"
                    SanduhrText = "Linien zwischen den Wertpaaren zeichnen"
                
                Case "DD0"
                    SanduhrText = "Linien löschen"
                Case "DD1"
                    SanduhrText = "Alle Linien des aktuellen Vergleichs löschen"
                Case "DD2"
                    SanduhrText = "Alle Linien dieses Tools löschen"
                
                Case "EE0"
                    SanduhrText = "Visualisieren"
                Case "EE1"
                    SanduhrText = "Einzelwerte des ersten Bereichs hervorheben"
                Case "EE2"
                    SanduhrText = "Einzelwerte des zweiten Bereichs hervorheben"
                Case "EE3"
                    SanduhrText = "Paare des ersten Bereichs hervorheben"
                Case "EE4"
                    SanduhrText = "Paare des zweiten Bereichs hervorheben"
                
                Case "FF0"
                    SanduhrText = "Daten prüfen"
                Case "FF1"
                    SanduhrText = "Auf alle kritischen Zeichen prüfen"
                Case "FF2"
                    SanduhrText = "Auf unnötige Leerzeichen prüfen"
                Case "FF3"
                    SanduhrText = "Auf Steuerzeichen prüfen"
                Case "FF4"
                    SanduhrText = "Auf geschützte Leerzeichen prüfen"
                    
                Case "GG0"
                    SanduhrText = "Daten bereinigen"
                Case "GG1"
                    SanduhrText = "Alle kritischen Zeichen entfernen"
                Case "GG2"
                    SanduhrText = "Unnötige Leerzeichen entfernen"
                Case "GG3"
                    SanduhrText = "Steuerzeichen entfernen"
                Case "GG4"
                    SanduhrText = "Geschützte Leerzeichen entfernen"
                    
                Case Else
                    SanduhrText = "[FEHLER]"
            End Select
        
        Case "EN"
            Select Case strText
                Case "AA0"
                    SanduhrText = "Find pairs"
                Case "AA1"
                    SanduhrText = "Read in area #1"
                Case "AA2"
                    SanduhrText = "Read in area #2"
                Case "AA3"
                    SanduhrText = "Look for record pairs and tag as found"
                Case "AA4"
                    SanduhrText = "Count values in area #1"
                Case "AA5"
                    SanduhrText = "Count values in area #2"
                
                Case "BB0"
                    SanduhrText = "Printing"
                Case "BB1"
                    SanduhrText = "Printing area #1"
                Case "BB2"
                    SanduhrText = "Printing area #2"
                
                Case "CC0"
                    SanduhrText = "Drawing lines"
                Case "CC1"
                    SanduhrText = "Drawing connection lines between record pairs"
                
                Case "DD0"
                    SanduhrText = "Removing lines"
                Case "DD1"
                    SanduhrText = "Removing connection lines of the current run"
                Case "DD2"
                    SanduhrText = "Removing all connection lines drawn by this tool"
                
                Case "EE0"
                    SanduhrText = "Visualising"
                Case "EE1"
                    SanduhrText = "Highlighting single values in area #1"
                Case "EE2"
                    SanduhrText = "Highlighting single values in area #2"
                Case "EE3"
                    SanduhrText = "Highlighting values of pairs in area #1"
                Case "EE4"
                    SanduhrText = "Highlighting values of pairs in area #2"
                
                Case "FF0"
                    SanduhrText = "Checking data"
                Case "FF1"
                    SanduhrText = "Checking for all critical characters"
                Case "FF2"
                    SanduhrText = "Checking for unnecessary spaces"
                Case "FF3"
                    SanduhrText = "Checking for control characters"
                Case "FF4"
                    SanduhrText = "Checking for non-breaking spaces"
                
                Case "GG0"
                    SanduhrText = "Cleaning up data"
                Case "GG1"
                    SanduhrText = "Removing all critical characters"
                Case "GG2"
                    SanduhrText = "Removing unnecessary spaces"
                Case "GG3"
                    SanduhrText = "Removing control characters"
                Case "GG4"
                    SanduhrText = "Removing non-breaking spaces"
                    
                Case Else
                    SanduhrText = "[ERROR]"
            End Select
        
        Case Else
            SanduhrText = "[FEHLER]"
    End Select
End Function
