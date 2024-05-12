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