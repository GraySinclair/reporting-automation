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
    End With 'END IT1 WITH  
'----------------------------------------------------------------
'CUSTOMER TOTALS
'-----------------------------------------------------------------
    Dim ct As Worksheet
    Set ct = ActiveWorkbook.Sheets("Customer Totals")
   
    With ct
        .Activate
        ' COLUMN NAMES
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
       
        With .Cells 'Applies to all cells
                .WrapText = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin ' Optional: set the weight of the borders
                    .ColorIndex = xlAutomatic ' Optional: set the color of the borders
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
       
 