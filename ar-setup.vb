Sub AR_Setup()
'
' AR Setup Macro
'
' TO DO:
 
' PreWork --------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
   
'--------------------------------------------------------------------------
'ITEMS 1
'------------------------------------------------------------------------
    Dim it1 As Worksheet
    Set it1 = ActiveWorkbook.Sheets("ITEMS")
    With it1
        .Activate
 
        'Format Removal
'        With Columns("A:AG")
'            .ClearFormats
'        End With

        ' COLUMN NAMES
        .Cells(1, 1).Value = "Unit"
        .Cells(1, 9).Value = "Last Bill Ref"
        .Cells(1, 11).Value = "OdyNum"
        'Text to col which helps make data consistent and removes white spaces.
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
        ' delete cells in notes column that were not pulled from yesterdays file aka 0's
        For Each cell In noteRange
            If cell.Value = 0 Then
                cell.ClearContents
            End If
        Next cell

        ' Formatting--------------------------------------------------------------------------
        With Cells 'Applies to all cells
                .WrapText = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin ' Optional: set the weight of the borders
                    .ColorIndex = xlAutomatic ' Optional: set the color of the borders
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