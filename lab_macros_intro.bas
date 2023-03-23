Attribute VB_Name = "Módulo1"
Sub lab_macros_intro()
Attribute lab_macros_intro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' lab_macros_intro Macro
'

'

' Set a loop to format all pages
Dim sheet_count As Integer
sheet_count = Sheets.Count

For i = 1 To sheet_count
    
    ' Start from sheet 1
    Sheets(i).Activate

    ' Delete comumn A
    Columns("A:A").Delete Shift:=xlToLeft
    
    ' Set horizonal alignment for columns A-E
    Columns("A:E").HorizontalAlignment = xlCenter
    
    ' Set columns widths
    Columns("A:A").ColumnWidth = 20.3
    Columns("B:B").ColumnWidth = 27.7
    Columns("C:C").ColumnWidth = 38.9
    Columns("D:E").ColumnWidth = 23.1
    
    ' Set formatting for the header row
    Range("A1:E1").Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent2
    End With
    Selection.Font.Bold = True
    
    ' Set borders
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
    End With

Next i

End Sub
