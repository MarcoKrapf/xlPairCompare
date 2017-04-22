VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAnleitung 
   Caption         =   "[xl PairCompare - Anleitung]"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8565
   OleObjectBlob   =   "frmAnleitung.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAnleitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Texte für die Anleitung, die beim Aufrufen gezogen werden
'---------------------------------------------------------

Private Sub UserForm_Initialize()

    'Erste Seite aktivieren
    MultiPageAnleitung.Value = 0
    
    Select Case g_strSprache
    
        Case "DE"
            With frmAnleitung
                'Allgemein
                .Caption = g_strTool & " " & g_strVersion & " - Anleitung"
                .lblInfo1.Caption = "Anleitung"
                .lblInfo2.Caption = "Das Tool sucht in zwei auf dem Tabellenblatt selektierten Bereichen " & _
                    "nach Werten, die in beiden Bereichen vorkommen. Jede Übereinstimmung bildet " & _
                    "ein Wertpaar. Die Bereiche werden von oben nach unten verglichen, somit wird der jeweils " & _
                    "erste Wert für ein Wertpaar herangezogen. Im Fall von Duplikaten werden so lange Wertpaare " & _
                    "gebildet wie dieser Wert auf beiden Seiten vorhanden ist. " & _
                    "Details zu den Funktionen sind in den jeweiligen Registerkarten beschrieben."
                'MultiPage
                .MultiPageAnleitung.Pages(0).Caption = "Vergleich"
                .MultiPageAnleitung.Pages(1).Caption = "Visualisierung"
                .MultiPageAnleitung.Pages(2).Caption = "Ausgabe"
                .MultiPageAnleitung.Pages(3).Caption = "Bereinigung"
                'Registerkarte "Vergleich"
                .lblAnl1.Caption = "In der Sektion 'Bereiche zu vergleichend' werden zwei auf dem Tabellenblatt " & _
                    "markierte Bereiche in das Tool eingelesen. Mit den Buttons 'Bereich 1 einlesen' bzw. 'Bereich 2 einlesen' " & _
                    "wird jeweils ein einzelner Bereich eingelesen, folglich darf dafür auch nur jeweils ein Bereich markiert " & _
                    "sein. Der Button 'Beide Bereiche einlesen' wird verwendet wenn schon beide Bereiche markiert sind. " & vbNewLine & vbNewLine & _
                    "Um große Bereiche schneller markieren zu können gibt es den Button 'Angrenzende Zellen selektieren'. " & _
                    "Damit werden alle Zellen, die an den bzw. die markierten Bereiche angrenzen, selektiert. " & _
                    "Werden ganze Spalten markiert, dann grenzt das Tool die Selektion automatisch bis zur letzten " & _
                    "benutzten Zeile ein. " & vbNewLine & vbNewLine & _
                    "Mit dem Button 'Auf kritische Zeichen prüfen' wird ein Schnell-Check " & _
                    "auf Leerzeichen am Anfang und am Ende einer Zelle, mehrfach aufeinanderfolgende Leerzeichen, " & _
                    "geschützte Leerzeichen und (teils nicht sichtbare) Steuerzeichen durchgeführt, welche den Vergleich " & _
                    "beeinflussen können. Solche Zeichen werden etwa durch importierte Daten aus anderen Systemen " & _
                    "in Excel eingeschleppt und können in der Registerkarte 'Bereinigung' genauer geprüft und wenn " & _
                    "nötig entfernt werden. Der Button ist nur aktiv wenn mindestens ein Bereich eingelesen ist." & vbNewLine & vbNewLine & _
                    "Im Bereich 'Visualisierung auf dem Tabellenblatt' kann angehakt werden, wie die optische " & _
                    "Darstellung auf dem Tabellenblatt beim Start des Vergleichs aussehen soll. " & _
                    "Alle Optionen können auch im Anschluss an den Vergleich jederzeit einzeln in der Registerkarte " & _
                    "'Visualisierung' an- bzw. ausgeschaltet werden. " & vbNewLine & vbNewLine & _
                    "Der Button 'Vergleich starten' ist nur aktiv wenn genau zwei Bereiche in das Tool eingelesen " & _
                    "sind. Für den Vergleich kann mit den nebenstehenden Optionen festgelegt werden, ob " & _
                    "Groß-/Kleinschreibung eine Rolle spielt und ob sämtliche(!) Leerzeichen in den Zellen " & _
                    "ignoriert werden sollen."
                'Registerkarte "Visualisierung"
                .lblAnl2.Caption = "Hier kann die optische Darstellung auf dem Tabellenblatt angepasst werden. " & vbNewLine & vbNewLine & _
                    "Je nachdem ob der Vergleich ergeben hat, dass in mindestens einem der Bereiche Werte " & _
                    "vorhanden sind, die es im anderen Bereich nicht gibt (Einzelwerte) oder ob identische Werte " & _
                    "in beiden Bereichen gefunden wurden (Wertpaare) sind die jeweiligen Buttons aktiviert. " & _
                    "Einzelwerte und Wertpaare können farblich hervorgehoben werden und es können Linien " & _
                    "zwischen den Wertpaaren gezeichnet werden. " & vbNewLine & vbNewLine & _
                    "Die Hervorhebungen können wieder gelöscht " & _
                    "werden, wobei hier sämtliche Formatierungen in den selektierten Bereichen entfernt werden " & _
                    "(die aktuell eingelesenen Bereiche werden im Button angezeigt). " & vbNewLine & vbNewLine & _
                    "Der Button 'Linien des aktuellen Vergleichs löschen' entfernt nur die zuletzt eingezeichneten " & _
                    "Verbindungslinien, während 'Alle Linien dieses Tools löschen' auch Linien früherer " & _
                    "Ausführungen löscht. Damit besteht die Möglichkeit, auf einem Tabellenblatt mehrere " & _
                    "Vergleiche hintereinander durchzuführen und beim Löschen der Linien des letzten Vergleichs " & _
                    "die anderen beizubehalten."
                        
                'Registerkarte "Ausgabe"
                .lblAnl3.Caption = "Für die Ausgabe der jeweiligen Werte wird unter 'Ausgabeort' festgelegt, " & _
                    "ob diese in einem neuen Tabellenblatt, das automatisch erzeugt wird, oder ab dem " & _
                    "Ort, an dem der Cursor steht, ausgegeben werden (die Zelle mit dem Cursor ist dann die " & _
                    "linke obere Ecke der Ausgabe)." & vbNewLine & vbNewLine & _
                    "Die Ausgabe kann entweder für die gefundenen Einzelwerte oder für die Wertpaare gestartet " & _
                    "werden. Dabei muss angegeben werden ob nur der erste, nur der zweite oder beide Bereiche " & _
                    "nebeneinander ausgegeben werden sollen. Die Buttons sind nur aktiviert wenn es die jeweiligen " & _
                    "Werte auch gibt."
                
                'Registerkarte "Bereinigung"
                .lblAnl4.Caption = "Mit dieser Registerkarte kann geprüft werden, ob es in den selektierten " & _
                    "Bereichen kritische Zeichen gibt, die den Vergleich beeinflussen können, etwa durch den " & _
                    "Import von Daten aus anderern Systemen." & vbNewLine & vbNewLine & _
                    "'Unnötige Leerzeichen' (ASCII-Zeichen-Code 32) sind solche am Anfang einer Zelle, " & _
                    "am Ende einer Zelle sowie mehrfach aufeinanderfolgende Leerzeichen. " & _
                    "'Steuerzeichen' sind alle Zeichen des 7-Bit-ASCII-Zeichensatzes mit den Zeichencodes 0 bis 31 " & _
                    "sowie die Unicode-Zeichen 127, 129, 141, 143, 144 und 157. " & _
                    "Das 'geschützte Leerzeichen' hat den Code 160 des Unicode-Zeichensatzes." & vbNewLine & vbNewLine & _
                    "Mit den Prüf-Buttons kann ein Einzel-Check auf eine Kategorie oder ein Gesamt-Check " & _
                    "durchgeführt werden. Die Anzahl gefundener Zellen der jeweiligen Kategorie wird nach dem Check " & _
                    "neben den jeweiligen Buttons angezeigt." & vbNewLine & vbNewLine & _
                    "Die Buttons zum Entfernen kritischer Zeichen führen nach einer Sicherheitsabfrage die Aktion unwiderruflich aus, " & _
                    "das Entfernen kann also nicht wieder rückgängig gemacht werden (was in der Regel auch nicht nötig ist). " & vbNewLine & vbNewLine & _
                    "Die Bereinigung unnötiger Leerzeichen entspricht der Funktion GLÄTTEN() und entfernt Leerzeichen " & _
                    "am Anfang und am Ende einer Zelle sowie alle mehrfach aufeinanderfolgenden Leerzeichen mit dem Zeichencode 32 " & _
                    "im Innern einer Zelle, so dass nur noch eines stehenbleibt." & vbNewLine & vbNewLine & _
                    "Das Entfernen der nicht druckbaren Steuerzeichen erfolgt in zwei Schritten. Zuerst wird jedes dieser " & _
                    "Zeichen durch einen horizontalen Tabulator (Zeichencode 9) ersetzt, der dann analog der Funktion " & _
                    "SÄUBERN() gelöscht wird." & vbNewLine & vbNewLine & _
                    "Geschützte Leerzeichen (Zeichencode 160) werden beim Entfernen zuerst durch normale Leerzeichen " & _
                    "mit dem Zeichencode 32 ausgetauscht. Anschließend wird die Funktion GLÄTTEN() ausgeführt, so dass eventuell " & _
                    "vorhandene mehrfache Leerzeichen entfernt werden. " & vbNewLine & vbNewLine & _
                    "Mit 'Alle kritischen Zeichen entfernen' werden alle drei zuvor beschriebenen Schritte gleichzeitig " & _
                    "ausgeführt."
            End With
            
        Case "EN"
            With frmAnleitung
                'Allgemein
                .Caption = g_strTool & " " & g_strVersion & " - User's guide"
                .lblInfo1.Caption = "User's guide"
                .lblInfo2.Caption = "The tool searches for data records which occur in both of exactly " & _
                    "two selected areas on a worksheet. Each match forms a pair of data records. " & _
                    "The areas are compared from top to bottom, therefore the first record " & _
                    "in each area is used for the value pair. In the case of duplicates, " & _
                    "value are formed as long as this data record is present in both areas. " & _
                    "Details on the functions are described on the respective tab pages."
                'MultiPage
                .MultiPageAnleitung.Pages(0).Caption = "Comparison"
                .MultiPageAnleitung.Pages(1).Caption = "Visualisation"
                .MultiPageAnleitung.Pages(2).Caption = "Output"
                .MultiPageAnleitung.Pages(3).Caption = "Cleanup"
                'Registerkarte "Vergleich"
                .lblAnl1.Caption = "In the section 'Areas to compare' two areas marked on the worksheet " & _
                    "are read into the tool. When clicking the buttons 'Read in area 1' and 'Read in area 2' " & _
                    "a single area is read, so only one area can be marked for this. " & _
                    "sein. The button 'Read in both areas' is used when two areas are selected on the worksheet. " & vbNewLine & vbNewLine & _
                    "For marking large areas more quickly, the button 'Select adjacent cells' can be applied. " & _
                    "This selects all cells which surround the cell where the cursor is placed on the worksheet. " & _
                    "If entire columns are selected, the tool automatically adjusts the selection to the lowest line used. " & vbNewLine & vbNewLine & _
                    "The 'Check for critical characters' button performs a quick check on spaces at the beginning " & _
                    "and the end of a cell, consecutive spaces inside a cell, non-breaking spaces and control characters " & _
                    "which can influence the comparison. Such characters are dragged into Excel e.g. by imported data " & _
                    "from other systems. They can be checked more detailed and removed in the 'Cleanup' tab if necessary. " & _
                    "This button is only active if at least one area is read into the tool." & vbNewLine & vbNewLine & _
                    "In the section 'Visualisation on the worksheet' the presentation of the results can be set. " & _
                    "All options can also be switched on and off individually in the 'Visualisation' tab " & _
                    "after the comparison has been executed. " & vbNewLine & vbNewLine & _
                    "The button 'Start comparison' is only active if exactly two areas are loaded into the tool. " & _
                    "Before starting, the checkboxes in the 'Options' section can be used to determine whether " & _
                    "the comparison is case-sensitive and whether all(!) spaces in the cells should be ignored."
        
                'Registerkarte "Visualisierung"
                .lblAnl2.Caption = "The optical representation on the table sheet can be adapted using this tab. " & vbNewLine & vbNewLine & _
                    "Depending on whether the comparison has shown that there are values in at least one of the " & _
                    "ranges that do not exist in the other area (single values) or if identical values " & _
                    "have been found in both areas (value pairs), the respective buttons are activated. " & _
                    "Single values and value pairs can be highlighted, and lines between the pairs of values " & _
                    "can be drawn. " & vbNewLine & vbNewLine & _
                    "The highlightings can be removed, which deletes all formatting in the selected areas " & _
                    "(the areas to apply this function are shown in the button)." & vbNewLine & vbNewLine & _
                    "The button 'Remove connenction lines of the current run' deletes the lines of the last " & _
                    "comparison, while 'Remove all connection lines in this worksheet' also deletes lines from " & _
                    "previous executions. This makes it possible to perform several comparisons on a worksheet " & _
                    "and to delete the lines of the last comparison while keeping the others."
                        
                'Registerkarte "Ausgabe"
                .lblAnl3.Caption = "For the output of the respective values, the 'Output location' is used to determine " & _
                    "whether the output is generated in a new worksheet that is generated automatically " & _
                    "or at the location where the cursor is located (the cell with the current cursor position " & _
                    "is the upper left corner)." & vbNewLine & vbNewLine & _
                    "The output can be performed either for the single values found or for the data record pairs. " & _
                    "It must be specified whether only the first, only the second or both areas should be printed. " & _
                    "The buttons are only activated if the respective values exist."
                
                'Registerkarte "Bereinigung"
                .lblAnl4.Caption = "Use this tab to check whether there are critical characters in the selected areas " & _
                    "that can influence the comparison, e.g. after importing data from other systems." & vbNewLine & vbNewLine & _
                    "'Unnecessary spaces' (ASCII character code 32) are those at the beginning of a cell, " & _
                    "at the end of a cell, and consecutive spaces inside a cell. 'Control characters' are all characters " & _
                    "of the 7-bit ASCII character set with the character codes 0 to 31 as well as the unicode " & _
                    "characters 127, 129, 141, 143, 144 and 157. The 'non-breaking space' has code 160 of the unicode " & _
                    "character set." & vbNewLine & vbNewLine & _
                    "The check buttons can be used to perform a single check on a category or an overall check. " & _
                    "The number of cells found in the respective category is displayed next to the " & _
                    "respective buttons after the check." & vbNewLine & vbNewLine & _
                    "The buttons for removing critical characters will irrevocably perform the respective operation " & _
                    "after showing a confirmation prompt, so the removal can not be undone." & vbNewLine & vbNewLine & _
                    "The cleansing of unnecessary spaces corresponds to the worksheet function TRIM() " & _
                    "and removes spaces at the beginning and the end of a cell, as well as all consecutive " & _
                    "spaces with the character code 32 inside a cell, leaving only one remaining." & vbNewLine & vbNewLine & _
                    "The non-printable control characters are removed in two steps. First, each of these " & _
                    "characters is replaced by a horizontal tab (character code 9), which is then deleted " & _
                    "in the same way as the worksheet function CLEAN()." & vbNewLine & vbNewLine & _
                    "Non-breaking spaces (character code 160) are first replaced by normal spaces with the " & _
                    "character code 32, followed by executing the TRIM() function so that any existing multiple " & _
                    "spaces are removed." & vbNewLine & vbNewLine & _
                    "'Remove all critical characters' executes all three steps described above simultaneously."
            End With
        
        Case Else
        
    End Select
End Sub
