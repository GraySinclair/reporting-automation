Sub AR_Setup()
    Dim items As Worksheet
    Dim last As Worksheet
 
' PreWork --------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
 
    Set items = ActiveWorkbook.Sheets("ITEMS")
    'Set last = ActiveWorkbook.Sheets("ITEMS (2)")
   
    RemoveBlankColumns items, 35
    DeleteExtraRows items
    CreateTable items, "today"
    RemoveColumnsByHeaders items, Array("1.0 Ticket #/Ody Document #", "CURRENCY" & vbLf, "RES #", "Car Class", "1 To 30 Days", "31 to 60 Days", "61 to 90 Days", "91+ Days", "COLLECTOR", "DEPTID")
    RenameTableHeader items, "BUSINESS UNIT", "Unit"
    RenameTableHeader items, "Last Billing Ref #", "LastBillRef"
    AddColumnsByHeaders
   
    'TODO TextToCol
 
   
'--------------------------------------------------------------------------------
'CLEANUP                                                                        |
'--------------------------------------------------------------------------------
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
' Create a table from a given range and name it
Sub CreateTable(ws As Worksheet, tblName As String)
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
 
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
 
    ' Add the table
    With ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tblRange, XlListObjectHasHeaders:=xlYes)
        .Name = tblName
        .TableStyle = "" ' Apply your preferred style
    End With
End Sub
' Remove blank columns from a worksheet
Sub RemoveBlankColumns(ws As Worksheet, startCol As Long)
    Dim col As Long
    For col = startCol To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
            ws.Columns(col).Delete
        End If
    Next col
End Sub
' Delete rows after the last used row
Sub DeleteExtraRows(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < ws.Rows.Count Then
        ws.Rows(lastRow + 1 & ":" & ws.Rows.Count).Delete
    End If
End Sub
' Remove columns based on header values
Sub RemoveColumnsByHeaders(ws As Worksheet, headers As Variant)
    Dim col As Long
    Dim header As Range
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set header = ws.Cells(1, col)
        If Not IsError(Application.Match(header.Value, headers, 0)) Then
            ws.Columns(col).Delete
            col = col - 1
        End If
    Next col
End Sub
Sub AddColumnsToTable(ws As Worksheet, numColumns As Integer, columnNames As Variant)
    Dim tbl As ListObject
    Dim i As Integer
    Dim newColumn As ListColumn
  
    Set tbl = ws.ListObjects(1)
  
    For i = 1 To numColumns
        Set newColumn = tbl.ListColumns.Add
        newColumn.Name = columnNames(i - 1)
    Next i
End Sub
' Takes worksheet, column name, formula. Then applies to table column
Sub SetFormula(ws As Worksheet, colName As String, formula As String)
    Dim tbl As ListObject
    Dim col As ListColumn
  
    Set tbl = ws.ListObjects(1)
    Set col = tbl.ListColumns(colName)
  
    ' Set formula in the body of the column
    col.DataBodyRange.formula = formula
End Sub
Sub MoveColumnBeforeX(ws As Worksheet, movingcol As String, beforecol As String)
    Dim tbl As ListObject
    Dim rng As Range
   
    Set tbl = ws.ListObjects(1)
    Set rng = tbl.ListColumns(movingcol).Range
   
    rng.Cut
    tbl.ListColumns(beforecol).Range.Insert Shift:=xlToRight
 
    ' Clean up
    Application.CutCopyMode = False
End Sub
Sub FormatTable(ws As Worksheet)
    Dim tbl As ListObject
    Dim headerCell As Range
    Dim dataRange As Range
   
 
    Set tbl = ws.ListObjects(1)
       
    tbl.Range.Style = "Normal"
       
    ' Format headers (bold and centered)
    Set headerCell = tbl.HeaderRowRange
    headerCell.Font.Bold = True
    headerCell.HorizontalAlignment = xlCenter
    headerCell.VerticalAlignment = xlCenter
       
    ' Format body of data
    Set dataRange = tbl.DataBodyRange
'    dataRange.Font.Size = 11
'    dataRange.Columns.AutoFit
    dataRange.HorizontalAlignment = xlCenter
    dataRange.VerticalAlignment = xlCenter
End Sub
Sub FormatTableColumn(ws As Worksheet, format As String, tblcol As Variant)
    Dim tbl As ListObject
    Dim col As Variant
    Dim columnRange As Range
 
    Set tbl = ws.ListObjects(1)
 
    For Each col In tblcol
        Set columnRange = tbl.ListColumns(col).DataBodyRange
       
        columnRange.NumberFormat = format
    Next col
End Sub
Sub CopyPasteTableAsValues(ws As Worksheet)
    Dim tbl As ListObject
    Dim tblRange As Range
   
    Set tbl = ws.ListObjects(1)
    Set tblRange = tbl.Range
   
    tblRange.Copy
    tblRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub
Sub RenameTableHeader(ws As Worksheet, oldHeader As String, newHeader As String)
    Dim tbl As ListObject
    Dim headerCell As Range
   
    If ws.ListObjects.Count > 0 Then
        Set tbl = ws.ListObjects(1)
       
        For Each headerCell In tbl.HeaderRowRange
            If headerCell.Value = oldHeader Then
                headerCell.Value = newHeader
                Exit For
            End If
        Next headerCell
    End If
End Sub
'    Dim it1 As Worksheet
'    Set it1 = ActiveWorkbook.Sheets("ITEMS")
'    With it1
'        .Activate
'
'        ' COLUMN NAMES
'        .Cells(1, 1).Value = "Unit"
'        .Cells(1, 9).Value = "Last Bill Ref"
'        .Cells(1, 11).Value = "OdyNum"
'
'        'Text to col which helps make data consistent and removes white spaces.
'        Dim columnsToFormat As Variant
'        Dim col As Variant
'
'        columnsToFormat = Array("D", "F", "G", "H", "I", "J", "K", "W", "X", "AF")
'
'        For Each col In columnsToFormat
'            With .Columns(col)
'                .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
'                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
'                    Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False
'            End With
'        Next col
'
'
'        ' Notes xLookup --------------------------------------------------------------------------------
'        Dim lastRow As Long
'        Dim noteRange As Range
'
'        lastRow = it1.Cells(it1.Rows.Count, "F").End(xlUp).Row
'        Set noteRange = .Range("AG2:AG" & lastRow)
'        .Range("AG2").formula = "=XLOOKUP(RC[-27],'ITEMS (2)'!C[-27],'ITEMS (2)'!C,0)"
'        noteRange.FillDown
'        noteRange.Value = noteRange.Value
'
'        ' delete cells in notes column that were not pulled from yesterdays file aka 0's
'        For Each cell In noteRange
'            If cell.Value = 0 Then
'                cell.ClearContents
'            End If
'        Next cell
'
'        ' Formatting--------------------------------------------------------------------------
'        With Cells 'Applies to all cells
'                .WrapText = False
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                With .Borders
'                    .LineStyle = xlContinuous
'                    .Weight = xlThin ' Optional: set the weight of the borders
'                    .ColorIndex = xlAutomatic ' Optional: set the color of the borders
'                End With
'        End With
'
'        .Columns("H:I").NumberFormat = "0"
'        .Columns("P:R").NumberFormat = "mm/dd/yy"
'        .Columns("S:U").NumberFormat = "$#,##0.00"
'        .Columns("AA:AD").NumberFormat = "$#,##0.00"
'
'        With Columns
'            .AutoFit
'            .AutoFilter ' Needs to be deactivated on Mac Dev Environment
'        End With
'
'        With .Range("A1:AG1")
'            .Font.Bold = True
'            .Interior.Color = RGB(128, 128, 128) ' Gray color
'        End With
'        .Columns("AG").HorizontalAlignment = xlLeft
'        'Move Data
'        .Columns("K").Cut
'        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("R:S").Cut
'        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("S:U").Cut
'        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("S:T").Cut
'        .Columns("N").Offset(0, 1).Insert Shift:=xlToRight
'
'        'Hide columns
'        .Columns("C").EntireColumn.Hidden = True
'        .Columns("M:P").EntireColumn.Hidden = True
'        .Columns("R:S").EntireColumn.Hidden = True
'        .Columns("U:AF").EntireColumn.Hidden = True
'
'        .Columns("AG").ColumnWidth = 30
'    End With 'END IT1 WITH
''----------------------------------------------------------------
''CUSTOMER TOTALS
''-----------------------------------------------------------------
'    Dim ct As Worksheet
'    Set ct = ActiveWorkbook.Sheets("Customer Totals")
'
'    With ct
'        .Activate
'        ' COLUMN NAMES
'        .Cells(1, 5).Value = "RentCount"
'        .Cells(1, 15).Value = "LastPayment"
'        .Cells(1, 16).Value = "LastPayment"
'        .Cells(1, 20).Value = "BA Created"
'        .Cells(1, 22).Value = "BA Updated"
'
'        With .Columns
'            .Hidden = False
'            .ClearFormats
'            .AutoFit
'            .AutoFilter ' Needs to be deactivated on Mac Dev Environment
'        End With
'
'        With .Columns("C")
'            .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
'                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
'                Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False
'        End With
'
'        With .Cells 'Applies to all cells
'                .WrapText = False
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                With .Borders
'                    .LineStyle = xlContinuous
'                    .Weight = xlThin ' Optional: set the weight of the borders
'                    .ColorIndex = xlAutomatic ' Optional: set the color of the borders
'                End With
'        End With
'        With .Range("A1:V1")
'            .Font.Bold = True
'            .Interior.Color = RGB(128, 128, 128) ' Gray color
'        End With
'        .Columns("C").NumberFormat = "0" 'BA COLUMN
'        .Columns("F:L").NumberFormat = "$#,##0.00"
'        .Columns("O").NumberFormat = "mm/dd/yy"
'        .Columns("P").NumberFormat = "$#,##0.00"
'        .Columns("R").NumberFormat = "$#,##0.00"
'        .Columns("T").NumberFormat = "mm/dd/yy"
'        .Columns("V").NumberFormat = "mm/dd/yy"
'
'        'Move Data
'        .Columns("G").Cut
'        .Columns("V").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("K").Cut
'        .Columns("V").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("M:N").Cut
'        .Columns("J").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("N:Q").Cut
'        .Columns("V").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("M").ColumnWidth = 30
'        .Columns("X").ColumnWidth = 10
'        .Columns("T").ColumnWidth = 30
'        .Columns("B").EntireColumn.Hidden = True
'        .Columns("Q:W").EntireColumn.Hidden = True
'        .Columns("M").HorizontalAlignment = xlLeft
'
'
'        Dim lastPayment As String
'        Dim balance As Double
'        Dim updateText As String
'
'        ' Format balance and payment date
'        balance = Application.WorksheetFunction.Sum(.Range("I2:J2"))
'        lastPayment = ". Last payment: " & format(.Range("K2").Value, "mm/dd") & " for " & format(.Range("L2").Value, "$#,##0.00")
'        updateText = ". Update: "
'
'        .Range("M2").formula = _
'        "= ""As of "" & TEXT(TODAY(), ""mm/dd"") & "", account has "" & E2 & "" vehicles. 61+ Balance: "" & TEXT(SUM(I2:J2), ""#,##0.00"") & "". Last Payment: "" & TEXT(K2, ""mm/dd"") & "" for "" & TEXT(L2, ""#,##0.00"") & "" Update: """
'        Dim lastAcc As Long
'        Dim conversationRange As Range
'        lastAcc = ct.Cells(ct.Rows.Count, "B").End(xlUp).Row - 1
'        Set conversationRange = .Range("M2:M" & lastAcc)
'        conversationRange.FillDown
'        conversationRange.Value = conversationRange.Value
'    End With 'End ct With
''----------------------------------------------------------------------
''                              Credits                                '
''----------------------------------------------------------------------
'    Dim cr As Worksheet
'    Set cr = ActiveWorkbook.Sheets("Credits")
'
'    With cr
'        .Activate
'        'Format Removal
''        With Columns("A:AG")
''            .ClearFormats
''        End With
'
'        ' COLUMN NAMES
'        .Cells(1, 1).Value = "Unit"
'        .Cells(1, 9).Value = "Last Bill Ref"
'        .Cells(1, 11).Value = "OdyNum"
'
'        'Text to col which helps make data consistent and removes white spaces.
'
'        columnsToFormat = Array("D", "F", "G", "H", "I", "J", "K", "W", "X", "AF")
'
'        For Each col In columnsToFormat
'            With .Columns(col)
'                .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
'                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
'                    Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False
'            End With
'        Next col
'
'
'        ' Notes xLookup --------------------------------------------------------------------------------
'        Dim lastcr As Long
'        Dim crRange As Range
'
'        lastcr = cr.Cells(cr.Rows.Count, "F").End(xlUp).Row
'        Set crRange = .Range("AG2:AG" & lastcr)
'        .Range("AG2").FormulaR1C1 = "=XLOOKUP(RC[-27],ITEMS!C[-27],ITEMS!C,0)"
'        crRange.FillDown
'        ' Formatting--------------------------------------------------------------------------
'        With Cells 'Applies to all cells
'                .WrapText = False
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                With .Borders
'                    .LineStyle = xlContinuous
'                    .Weight = xlThin ' Optional: set the weight of the borders
'                    .ColorIndex = xlAutomatic ' Optional: set the color of the borders
'                End With
'        End With
'
'        .Columns("H:I").NumberFormat = "0"
'        .Columns("P:R").NumberFormat = "mm/dd/yy"
'        .Columns("S:U").NumberFormat = "$#,##0.00"
'        .Columns("AA:AD").NumberFormat = "$#,##0.00"
''        .Range("F2:F" & lastcr).Interior.Color = RGB(0, 0, 128)
'        With Columns
'            .AutoFit
'            .AutoFilter ' Needs to be deactivated on Mac Dev Environment
'        End With
'
'        With .Range("A1:AG1")
'            .Font.Bold = True
'            .Interior.Color = RGB(128, 128, 128) ' Gray color
'        End With
'        .Columns("AG").HorizontalAlignment = xlLeft
'        'Move Data
'        .Columns("K").Cut
'        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("R:S").Cut
'        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("S:U").Cut
'        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
'        .Columns("S:T").Cut
'        .Columns("N").Offset(0, 1).Insert Shift:=xlToRight
'
'        'Hide columns
'        .Columns("C").EntireColumn.Hidden = True
'        .Columns("M:P").EntireColumn.Hidden = True
'        .Columns("R:S").EntireColumn.Hidden = True
'        .Columns("U:AF").EntireColumn.Hidden = True
'
'        .Columns("AG").ColumnWidth = 30
'    End With 'END CR WITH
'
'    Dim sheetNames As Variant
'    Dim sheetName As Variant
'    Dim ws As Worksheet
'    Dim wsfr As Worksheet
'
'    ' List of sheets to delete
'    sheetNames = Array("ITEMS (2)", "Top 20 60+Previous")
'
'    ' Loop through each sheet name in the array
'    For Each sheetName In sheetNames
'        On Error Resume Next ' Ignore errors if the sheet doesn't exist
'        Set ws = ActiveWorkbook.Sheets(sheetName)
'        If Not ws Is Nothing Then
'            ws.Delete ' Delete the sheet
'        End If
'        On Error GoTo 0 ' Turn error handling back on
'    Next sheetName
'
'    'applied to all worksheets
'    For Each wsfr In ActiveWorkbook.Sheets
'        With wsfr
'            .Activate ' Activate the sheet
'            ActiveWindow.FreezePanes = False ' Unfreeze any existing frozen panes
'            .Rows("2:2").Select ' Select the row to freeze
'            ActiveWindow.FreezePanes = True ' Apply the freeze panes setting
'            ActiveWindow.Zoom = 85 ' Set the zoom level to 85%
'            .Range("A1").Activate
'        End With
'    Next wsfr
'    With it1
'        .Activate
'        .Range("A1").Activate
'    End With