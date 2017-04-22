VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVersionshinweise 
   Caption         =   "[xl PairCompare - Features und Versionshistorie]"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7755
   OleObjectBlob   =   "frmVersionshinweise.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmVersionshinweise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Texte für die Versionshinweise, die beim Aufrufen gezogen werden
'----------------------------------------------------------------

Private Sub UserForm_Initialize()

    Select Case g_strSprache
        Case "DE"
            With frmVersionshinweise
                .Caption = g_strTool & " - Features und Versionshistorie"
                .lblVersionsInfo10a.Caption = "Version 1.0 (19.11.2016)"
                .lblVersionsInfo10b.Caption = "- Finden von Wertpaaren, also Werten die in beiden selektierten Bereichen vorkommen" & vbNewLine & _
                                                "- Farbiges Hervorheben von Einzelwerten und/oder Wertpaaren auf dem Tabellenblatt" & vbNewLine & _
                                                "- Visualisieren von Wertpaaren durch Zeichnen von Verbindungslinien auf dem Tabellenblatt" & vbNewLine & _
                                                "- Ausgeben von gefundenen Einzelwerten und/oder Wertpaaren auf dem Tabellenblatt" & vbNewLine & _
                                                "- Entfernen von kritischen Zeichen wie Steuerzeichen und geschützten Leerzeichen in Zellen möglich" & vbNewLine & _
                                                "- Ignorieren von Groß-/Kleinbuchstaben und/oder Leerzeichen beim Vergleichen möglich" & vbNewLine & _
                                                "- Schnelle Selektion der Bereiche durch automatisches Markieren angrenzender Zellen möglich"
                .lblVersionsInfo11a.Caption = "Version 1.1 (26.11.2016)"
                .lblVersionsInfo11b.Caption = "- GUI kann zwischen englisch und deutsch umgeschaltet werden" & vbNewLine & _
                                                "- Bugfixes"
            End With
            
        Case "EN"
            With frmVersionshinweise
                .Caption = g_strTool & " - Features and version history"
                .lblVersionsInfo10a.Caption = "Version 1.0 (19.11.2016)"
                .lblVersionsInfo10b.Caption = "- Finding value pairs, i.e. data records that occur in both selected areas" & vbNewLine & _
                                                "- Highlighting of single values and/or value pairs on the worksheet" & vbNewLine & _
                                                "- Visualise data record pairs on the worksheet" & vbNewLine & _
                                                "- Output of detected single values and/or value pairs on the worksheet" & vbNewLine & _
                                                "- Removing critical characters such as control characters possible" & vbNewLine & _
                                                "- Ignore lower/upper-case letters and/or spaces possible when comparing" & vbNewLine & _
                                                "- Quick selection of the areas possible by automatically marking adjacent cells"
                .lblVersionsInfo11a.Caption = "Version 1.1 (26.11.2016)"
                .lblVersionsInfo11b.Caption = "- GUI can be switched between English and German" & vbNewLine & _
                                                "- Bugfixes"
            End With
            
        Case Else
        
    End Select
End Sub
