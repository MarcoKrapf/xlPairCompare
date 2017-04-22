Attribute VB_Name = "modINFOtexte"
Option Explicit

'Modulbeschreibung:
'R�ckgabe der Texte f�r Messageboxen und andere Infos
'----------------------------------------------------

Public Function MsgBoxText(strSprachwahl As String, strText As String) As String
    Select Case strSprachwahl
        Case "DE"
            Select Case strText
                Case "MA1"
                    MsgBoxText = "F�r diese Funktion muss genau ein Bereich selektiert sein"
                Case "MA2"
                    MsgBoxText = "F�r diese Funktion m�ssen genau zwei Bereiche selektiert sein"
                
                Case "MB1"
                    MsgBoxText = "Bereich "
                Case "MB2"
                    MsgBoxText = " wird angepasst."
                Case "MB3"
                    MsgBoxText = "Die letzte Zelle mit Inhalt befindet sich"
                Case "MB4"
                    MsgBoxText = "in Zeile "
                
                Case "MC1"
                    MsgBoxText = "Check durchgef�hrt"
                Case "MC2"
                    MsgBoxText = "Die Zellen enthalten "
                Case "MC3"
                    MsgBoxText = ", wodurch der Vergleich der Bereiche beeinflusst werden kann."
                Case "MC4"
                    MsgBoxText = "Eine Datenbereinigung in der Registerkarte 'Bereinigung' wird empfohlen."
                
                Case "MD1"
                    MsgBoxText = "Entfernen aller kritischen Zeichen"
                Case "MD2"
                    MsgBoxText = "1) Leerzeichen am Anfang und Ende einer Zelle"
                Case "MD3"
                    MsgBoxText = "sowie mehrfach aufeinanderfolgende Leerzeichen"
                Case "MD4"
                    MsgBoxText = "2) Steuerzeichen (7-Bit-ASCII-Zeichen 0-31"
                Case "MD5"
                    MsgBoxText = "und Unicode-Zeichen 127, 129, 141, 143, 144 und 157)"
                Case "MD6"
                    MsgBoxText = "3) Gesch�tzte Leerzeichen (Unicode-Zeichen 160)"
                Case "MD7"
                    MsgBoxText = "DIESE AKTION KANN NICHT R�CKG�NGIG GEMACHT WERDEN"
                Case "MD8"
                    MsgBoxText = "Daten bereinigen - Alle kritischen Zeichen entfernen"
                Case "MD9"
                    MsgBoxText = "Entfernen von Leerzeichen am Anfang und Ende einer Zelle"
                Case "MD10"
                    MsgBoxText = "sowie von mehrfach aufeinanderfolgenden Leerzeichen"
                Case "MD11"
                    MsgBoxText = "Daten bereinigen - Unn�tige Leerzeichen entfernen"
                Case "MD12"
                    MsgBoxText = "Entfernen von 7-Bit-ASCII-Zeichen (Zeichencodes 0-31)"
                Case "MD13"
                    MsgBoxText = "und Unicode-Zeichen (Zeichencodes 127, 129, 141, 143, 144 und 157)"
                Case "MD14"
                    MsgBoxText = "Daten bereinigen - Steuerzeichen entfernen"
                Case "MD15"
                    MsgBoxText = "Entfernen von gesch�tzten Leerzeichen (Unicode-Zeichen 160)"
                Case "MD16"
                    MsgBoxText = "Daten bereinigen - Gesch�tzte Leerzeichen entfernen"
                
                Case "SB1"
                    MsgBoxText = "unn�tige Leerzeichen"
                Case "SB2"
                    MsgBoxText = "Steuerzeichen"
                Case "SB3"
                    MsgBoxText = "unn�tige Leerzeichen und Steuerzeichen"
                Case "SB4"
                    MsgBoxText = "gesch�tzte Leerzeichen"
                Case "SB5"
                    MsgBoxText = "unn�tige Leerzeichen und gesch�tzte Leerzeichen"
                Case "SB6"
                    MsgBoxText = "Steuerzeichen und gesch�tzte Leerzeichen"
                Case "SB7"
                    MsgBoxText = "unn�tige Leerzeichen, Steuerzeichen und gesch�tzte Leerzeichen"
                Case "SB8"
                    MsgBoxText = "Keine kritischen Zeichen gefunden."
                Case "SB9"
                    MsgBoxText = "kritische Zeichen"
                Case "SB10"
                    MsgBoxText = "Keine unn�tigen Leerzeichen gefunden."
                Case "SB11"
                    MsgBoxText = "Keine Steuerzeichen gefunden."
                Case "SB12"
                    MsgBoxText = "Keine gesch�tzten Leerzeichen gefunden."
                    
                Case "SC1"
                    MsgBoxText = "Sorry, hier ist was schiefgelaufen..."
                Case "SC2"
                    MsgBoxText = "Fehler Nr. "
                Case "SC3"
                    MsgBoxText = " - Fehler im Code"
                    
                Case Else
                    MsgBoxText = "[FEHLER]"
            End Select
        
        Case "EN"
            Select Case strText
                Case "MA1"
                    MsgBoxText = "To perform this function, exactly one area must be selected"
                Case "MA2"
                    MsgBoxText = "To perform this function, exactly two areas must be selected"
                
                Case "MB1"
                    MsgBoxText = "Area "
                Case "MB2"
                    MsgBoxText = " has been adjusted."
                Case "MB3"
                    MsgBoxText = "The last cell with content is located"
                Case "MB4"
                    MsgBoxText = "in row "
                
                Case "MC1"
                    MsgBoxText = "Check performed"
                Case "MC2"
                    MsgBoxText = "Some cells contain "
                Case "MC3"
                    MsgBoxText = " which can impact the comparison of the areas."
                Case "MC4"
                    MsgBoxText = "We recommend cleaning up the data in the 'Cleanup' tab."
                
                Case "MD1"
                    MsgBoxText = "Removing all critical characters"
                Case "MD2"
                    MsgBoxText = "1) Spaces at the beginning and the end of a cell"
                Case "MD3"
                    MsgBoxText = "as well as multiple spaces within a cell"
                Case "MD4"
                    MsgBoxText = "2) Control characters (7-bit ASCII code characters 0-31"
                Case "MD5"
                    MsgBoxText = "and unicode characters 127, 129, 141, 143, 144 and 157)"
                Case "MD6"
                    MsgBoxText = "3) Non-breaking spaces (unicode character 160)"
                Case "MD7"
                    MsgBoxText = "THIS ACTION CANNOT BE UNDONE"
                Case "MD8"
                    MsgBoxText = "Cleaning up data - Removing all critical characters"
                Case "MD9"
                    MsgBoxText = "Removal of spaces at the beginning and the end of a cell"
                Case "MD10"
                    MsgBoxText = "as well as multiple spaces within a cell"
                Case "MD11"
                    MsgBoxText = "Cleaning up data - Removing unnecessary spaces"
                Case "MD12"
                    MsgBoxText = "Removal of 7-bit ASCII code characters 0-31"
                Case "MD13"
                    MsgBoxText = "and unicode characters 127, 129, 141, 143, 144 and 157"
                Case "MD14"
                    MsgBoxText = "Cleaning up data - Removing control characters"
                Case "MD15"
                    MsgBoxText = "Removal of non-breaking spaces (unicode character 160)"
                Case "MD16"
                    MsgBoxText = "Cleaning up data - Removing non-breaking spaces"
                
                Case "SB1"
                    MsgBoxText = "unnecessary spaces"
                Case "SB2"
                    MsgBoxText = "control characters"
                Case "SB3"
                    MsgBoxText = "unnecessary spaces and control characters"
                Case "SB4"
                    MsgBoxText = "non-breaking spaces"
                Case "SB5"
                    MsgBoxText = "unnecessary spaces and gesch�tzte Leerzeichen"
                Case "SB6"
                    MsgBoxText = "control characters und non-breaking spaces"
                Case "SB7"
                    MsgBoxText = "unnecessary spaces, control characters and non-breaking spaces"
                Case "SB8"
                    MsgBoxText = "No critical characters found."
                Case "SB9"
                    MsgBoxText = "critical characters"
                Case "SB10"
                    MsgBoxText = "No unnecessary spaces found."
                Case "SB11"
                    MsgBoxText = "No control characters found."
                Case "SB12"
                    MsgBoxText = "No non-breaking spaces found."
                    
                Case "SC1"
                    MsgBoxText = "Sorry, something went wrong..."
                Case "SC2"
                    MsgBoxText = "Error number "
                Case "SC3"
                    MsgBoxText = " - Programming error"
                    
                Case Else
                    MsgBoxText = "[FEHLER]"
        
            End Select
            
        Case Else
            MsgBoxText = "[FEHLER]"
    End Select
End Function

Public Function SonstigerText(strSprachwahl As String, strText As String) As String
    Select Case strSprachwahl
        Case "DE"
            Select Case strText
                Case "TXT1"
                    SonstigerText = "Ausgabe_"
                    
                Case Else
                    SonstigerText = "[FEHLER]"
            End Select
        
        Case "EN"
            Select Case strText
                Case "TXT1"
                    SonstigerText = "Output_"
                
                Case Else
                    SonstigerText = "[ERROR]"
            End Select
        
        Case Else
            SonstigerText = "[FEHLER]"
    End Select
End Function
