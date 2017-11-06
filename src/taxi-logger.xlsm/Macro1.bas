Attribute VB_Name = "Macro1"

Sub beautify(writePos)
    Columns("B:B").ColumnWidth = 2.75
    Range("D:D,F:F,H:H").ColumnWidth = 35.25
    Range("D:D,F:F,H:H").WrapText = True
    Range("B4:H" & writePos).Select
    Selection.Borders.Weight = xlMedium
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintArea = "$B:$H"
        .PrintTitleRows = "$1:$4"
        .PrintTitleColumns = ""
        .FitToPagesWide = 1
    End With
    Application.PrintCommunication = True
End Sub
