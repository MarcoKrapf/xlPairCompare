Attribute VB_Name = "modTool"
Option Explicit

'Modulbeschreibung:
'Globale Variablen festlegen, Tool starten und GUI aufrufen
'----------------------------------------------------------

Public Const g_strTool As String = "xlPairCompare" 'Tool-Name
Public Const g_strVersion As String = "Version 1.1" 'Tool-Version
Public g_strSprache As String 'Kennzeichen für die Sprache der GUI
Public g_varSelection As Range 'Adresse des gesamten selektierten Bereichs auf dem Tabellenblatt
Public g_varSelection1 As Range 'Adresse auf dem Tabellenblatt von Bereich 1
Public g_varSelection2 As Range 'Adresse auf dem Tabellenblatt von Bereich 2
Public g_strSanduhrAktion As String 'Aktion, bei der die Sanduhr aufgerufen wird
Public g_strSanduhrNummer As String 'Nummer des Einzelschritts, bei der die Sanduhr aufgerufen wird
Public g_strSanduhrSchritt As String 'Einzelschritt, bei der die Sanduhr aufgerufen wird
Public g_dblBalkenAnteil As Double 'Breite des Fortschrittsbalkens der Sanduhr pro Schleifendurchlauf
Public g_dblBalkenAktuell As Double 'Aktuelle Breite des Fortschrittsbalkens der Sanduhr


Sub ToolStartenIcon(control As IRibbonControl) 'Aufruf durch den Button im Ribbon
    Call ToolStarten
End Sub

Sub ToolStarten() 'Diese Prozedur manuell starten zum Testen der Entwicklung
    'Sprache einstellen
        g_strSprache = "DE"
        frmGUI.checkboxTooltip = False
    'Sanduhr initialisieren
        With frmSanduhr
            .Caption = "xlPairCompare"
            .lblFortschrittSchritt = ""
            .lblFortschrittProzent = ""
            .lblFortschrittBalken.Width = 0
        End With
    'GUI laden und starten
        Load frmGUI
        frmGUI.Show
End Sub
