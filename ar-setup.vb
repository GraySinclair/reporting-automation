Sub AR_Setup()
'
' AR Setup Macro
'
' TO DO:
'
'
'

' PreWork --------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
   
'--------------------------------------------------------------------------
'ITEMS 1
'------------------------------------------------------------------------
    Dim it1 As Worksheet
    Set it1 = ActiveWorkbook.Sheets("ITEMS")
    With it1
        .Activate
 
        'test processing time with format removal for tolls macro
'        With Columns("A:AG")
'            .ClearFormats
'        End With

        ' COLUMN NAMES
        .Cells(1, 1).Value = "Unit"
        .Cells(1, 9).Value = "Last Bill Ref"
        .Cells(1, 11).Value = "OdyNum"
        'Text to col to make data consistent and removes white spaces.
        Dim columnsToFormat As Variant
        Dim col As Variant
 
        columnsToFormat = Array("D", "F", "G", "H", "I", "J", "K", "W", "X", "AF")
       
        For Each col In columnsToFormat
            With .Columns(col)
                .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                    Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False
            End With
        Next col

        ' Notes xLookup --------------------------------------------------------------------------------
        Dim lastRow As Long
        Dim noteRange As Range
       
        lastRow = it1.Cells(it1.Rows.Count, "F").End(xlUp).Row
        Set noteRange = .Range("AG2:AG" & lastRow)
        .Range("AG2").Formula = "=XLOOKUP(RC[-27],'ITEMS (2)'!C[-27],'ITEMS (2)'!C,0)"
        noteRange.FillDown
        noteRange.Value = noteRange.Value
        ' delete cells in notes column that were not pulled from yesterdays file (nulls)
        For Each cell In noteRange
            If cell.Value = 0 Then
                cell.ClearContents
            End If
        Next cell

        ' Formatting--------------------------------------------------------------------------
        With Cells
                .WrapText = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
        End With

        .Columns("H:I").NumberFormat = "0"
        .Columns("P:R").NumberFormat = "mm/dd/yy"
        .Columns("S:U").NumberFormat = "$#,##0.00"
        .Columns("AA:AD").NumberFormat = "$#,##0.00"
       
        With Columns
            .AutoFit
            .AutoFilter ' Needs to be deactivated on Mac Dev Environment
        End With
       
        With .Range("A1:AG1")
            .Font.Bold = True
            .Interior.Color = RGB(128, 128, 128) ' Gray color
        End With
        .Columns("AG").HorizontalAlignment = xlLeft
        
        'Move Data
        .Columns("K").Cut
        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("R:S").Cut
        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("S:U").Cut
        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("S:T").Cut
        .Columns("N").Offset(0, 1).Insert Shift:=xlToRight
   
        'Hide columns
        .Columns("C").EntireColumn.Hidden = True
        .Columns("M:P").EntireColumn.Hidden = True
        .Columns("R:S").EntireColumn.Hidden = True
        .Columns("U:AF").EntireColumn.Hidden = True
        
        .Columns("AG").ColumnWidth = 30
    End With 'End With for IT1


'CUSTOMER TOTALS
'-----------------------------------------------------------------
    Dim ct As Worksheet
    Set ct = ActiveWorkbook.Sheets("Customer Totals")
   
    With ct
        .Activate
        ' column headers
        .Cells(1, 5).Value = "RentCount"
        .Cells(1, 15).Value = "LastPayment"
        .Cells(1, 16).Value = "LastPayment"
        .Cells(1, 20).Value = "BA Created"
        .Cells(1, 22).Value = "BA Updated"
           
        With .Columns
            .Hidden = False
            .ClearFormats
            .AutoFit
            .AutoFilter ' Needs to be deactivated on Mac Dev Environment
        End With
       
        With .Columns("C")
            .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False
        End With
       
        With .Cells
                .WrapText = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
        End With
        With .Range("A1:V1")
            .Font.Bold = True
            .Interior.Color = RGB(128, 128, 128) ' Gray color
        End With
        .Columns("C").NumberFormat = "0" 'BA COLUMN
        .Columns("F:L").NumberFormat = "$#,##0.00"
        .Columns("O").NumberFormat = "mm/dd/yy"
        .Columns("P").NumberFormat = "$#,##0.00"
        .Columns("R").NumberFormat = "$#,##0.00"
        .Columns("T").NumberFormat = "mm/dd/yy"
        .Columns("V").NumberFormat = "mm/dd/yy"
       
        'Move Data
        .Columns("G").Cut
        .Columns("V").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("K").Cut
        .Columns("V").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("M:N").Cut
        .Columns("J").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("N:Q").Cut
        .Columns("V").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("M").ColumnWidth = 30
        .Columns("X").ColumnWidth = 10
        .Columns("T").ColumnWidth = 30
        .Columns("B").EntireColumn.Hidden = True
        .Columns("Q:W").EntireColumn.Hidden = True
        .Columns("M").HorizontalAlignment = xlLeft
        
        Dim lastPayment As String
        Dim balance As Double
        Dim updateText As String

        ' Format balance and payment date
        balance = Application.WorksheetFunction.Sum(.Range("I2:J2"))
        lastPayment = ". Last payment: " & Format(.Range("K2").Value, "mm/dd") & " for " & Format(.Range("L2").Value, "$#,##0.00")
        updateText = ". Update: "

        .Range("M2").Formula = _
        "= ""As of "" & TEXT(TODAY(), ""mm/dd"") & "", account has "" & E2 & "" vehicles. 61+ Balance: "" & TEXT(SUM(I2:J2), ""#,##0.00"") & "". Last Payment: "" & TEXT(K2, ""mm/dd"") & "" for "" & TEXT(L2, ""#,##0.00"") & "" Update: """
        Dim lastAcc As Long
        Dim conversationRange As Range
        lastAcc = ct.Cells(ct.Rows.Count, "B").End(xlUp).Row - 1
        Set conversationRange = .Range("M2:M" & lastAcc)
        conversationRange.FillDown
        conversationRange.Value = conversationRange.Value
    End With 'End With for CTs

'                              Credits                                
'----------------------------------------------------------------------
    Dim cr As Worksheet
    Set cr = ActiveWorkbook.Sheets("Credits")
   
    With cr
        .Activate
        'Format Removal
'        With Columns("A:AG")
'            .ClearFormats
'        End With
        ' COLUMN NAMES
        .Cells(1, 1).Value = "Unit"
        .Cells(1, 9).Value = "Last Bill Ref"
        .Cells(1, 11).Value = "OdyNum"
 
        columnsToFormat = Array("D", "F", "G", "H", "I", "J", "K", "W", "X", "AF")
       
       For Each col In columnsToFormat
            With .Columns(col)
                .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                    Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False
            End With
        Next col
       
        ' Notes xLookup --------------------------------------------------------------------------------
        Dim lastcr As Long
        Dim crRange As Range
       
        lastcr = cr.Cells(cr.Rows.Count, "F").End(xlUp).Row
        Set crRange = .Range("AG2:AG" & lastcr)
        .Range("AG2").FormulaR1C1 = "=XLOOKUP(RC[-27],ITEMS!C[-27],ITEMS!C,0)"
        crRange.FillDown
        ' Formatting--------------------------------------------------------------------------
        With Cells
                .WrapText = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
        End With

        .Columns("H:I").NumberFormat = "0"
        .Columns("P:R").NumberFormat = "mm/dd/yy"
        .Columns("S:U").NumberFormat = "$#,##0.00"
        .Columns("AA:AD").NumberFormat = "$#,##0.00"
'        .Range("F2:F" & lastcr).Interior.Color = RGB(0, 0, 128)
        With Columns
            .AutoFit
            .AutoFilter ' Needs to be deactivated on Mac Dev Environment
        End With

        With .Range("A1:AG1")
            .Font.Bold = True
            .Interior.Color = RGB(128, 128, 128) ' Gray color
        End With
        .Columns("AG").HorizontalAlignment = xlLeft
        
        'Move Data
        .Columns("K").Cut
        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("R:S").Cut
        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("S:U").Cut
        .Columns("AF").Offset(0, 1).Insert Shift:=xlToRight
        .Columns("S:T").Cut
        .Columns("N").Offset(0, 1).Insert Shift:=xlToRight

        'Hide columns
        .Columns("C").EntireColumn.Hidden = True
        .Columns("M:P").EntireColumn.Hidden = True
        .Columns("R:S").EntireColumn.Hidden = True
        .Columns("U:AF").EntireColumn.Hidden = True

        .Columns("AG").ColumnWidth = 30
    End With 'End With for Cr

    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim wsfr As Worksheet

    ' Array of sheets to delete(make sure alerts are off!)
    sheetNames = Array("ITEMS (2)", "Top 20 60+Previous")

    ' Loop through each sheet name in the array
    For Each sheetName In sheetNames
        On Error Resume Next ' Ignore errors if the sheet doesn't exist
        Set ws = ActiveWorkbook.Sheets(sheetName)
        If Not ws Is Nothing Then
            ws.Delete
        End If
        On Error GoTo 0 ' Turn error handling back on
    Next sheetName
   
    'applied to all worksheets
    For Each wsfr In ActiveWorkbook.Sheets
        With wsfr
            .Activate ' Activate the sheet
            ActiveWindow.FreezePanes = False ' Unfreeze any existing frozen panes in case of manual setting
            .Rows("2:2").Select ' Select the row to freeze
            ActiveWindow.FreezePanes = True ' Reapply the freeze panes setting
            ActiveWindow.Zoom = 85
            .Range("A1").Select
        End With
    Next wsfr
    it1.Activate

'CLEANUP
'--------------------------------------------------------------------------------
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub