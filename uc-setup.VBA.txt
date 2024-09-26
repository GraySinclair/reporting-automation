Sub UCSetup()
'
' UC SETUP
'
 
'
    'Prefix
    Application.DisplayAlerts = False
   
    '
    Sheets("Receivables Radius Export").Select
    ActiveWorkbook.Worksheets("Receivables Radius Export").AutoFilter.Sort. _
        SortFields.Clear
    Sheets("V3BD").Select
    ActiveWorkbook.Worksheets("V3BD").AutoFilter.Sort. _
        SortFields.Clear
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Receivables Radius Export").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Application.Goto Reference:="R1C1"
    Rows("1:1").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    ActiveSheet.Range("$A$1:$AH$9999").AutoFilter Field:=9, Criteria1:="STL19"
    Sheets("Receivables Radius Export").Select
    Sheets("Receivables Radius Export").Name = "UC"
    Sheets("V3BD").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("UC").Select
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("J:J").EntireColumn.AutoFit
    Selection.NumberFormat = "0"
    Range("J1").Select
    Selection.FormulaArray = "PU Control"
    Range("A1").Select
    Application.DisplayAlerts = True
End Sub