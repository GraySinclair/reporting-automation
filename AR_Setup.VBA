Sub AR_Setup()
'
' AR Setup Macro
'
' TO DO:
' Make numbers to plain text after Text to columns Function
' Fix notes lookup
'
    Dim it1 As Worksheet
    Set it1 = ActiveWorkbook.Sheets("ITEMS")
 
' Notifications OFF --------------------------------------------------------------------------
    Application.DisplayAlerts = False
   
' Text to Columns Number Data ----------------------------------------------------------------
    With it1
        With Columns("D") 'BA COLUMN
            .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False
        End With
        With Columns("F") 'TICKET COLUMN
            .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False
        End With
        With Columns("G") 'INVOICE NUMBER COLUMN
            .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False
        End With
        With Columns("H") 'BILL REF NUMBER COLUMN
            .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
               TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False
        End With
        With Columns("I") 'LAST BILL REF NUMBER COLUMN
            .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False
        End With
 
        .Cells(1, 1).Value = "Unit"
        .Cells(1, 9).Value = "Last Bill Ref"
        .Cells(1, 11).Value = "OdyNum"
        .Cells(1, 16).Value = "AS OF"
       
' Notes Lookup --------------------------------------------------------------------------------
        Dim lastRow As Long
        lastRow = it1.Cells(it1.Rows.Count, "F").End(xlUp).Row
        .Range("AG2").Formula = "=XLOOKUP(RC[-27],'ITEMS (2)'!C[-27],'ITEMS (2)'!C,0)"
        .Range("AG2:AG" & lastRow).FillDown
        .Range("AG2:AG" & lastRow).Copy
        .Range("AG2:AG" & lastRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        .Range("$A$1:$AG" & lastRow).AutoFilter Field:=33, Criteria1:="0"
' Columns Formatting --------------------------------------------------------------------------
        With Columns
            .AutoFit
        End With
        Columns("C").EntireColumn.Hidden = True
'        Columns("K").Delete 'Cant delete columns until other macros are adjusted to suit new AR
       
        .Activate ' Activate the sheet
        With ActiveWindow
            .FreezePanes = False ' Unfreeze any existing frozen panes
            ActiveSheet.Rows("2:2").Select '.SplitRow = 1       ' Set the row to freeze
            .FreezePanes = True ' Apply the freeze panes setting
            .Zoom = 85 ' Set the zoom level to 85%
            With Cells 'Applies to all cells
                .WrapText = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .AutoFilter
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin ' Optional: set the weight of the borders
                    .ColorIndex = xlAutomatic ' Optional: set the color of the borders
                End With
            End With
        End With
    End With
'------------------------------
 
 
    Columns("K:K").Select
    Selection.Cut
    Columns("V:V").Select
    Selection.Insert Shift:=xlToRight
    Columns("P:S").Select
    Range("S1").Activate
    Selection.EntireColumn.Hidden = True
    Columns("Y:Z").Select
    Range("Z1").Activate
    Selection.Cut
    Columns("O:P").Select
    Selection.Insert Shift:=xlToRight
    Columns("M:P").Select
    Range("P1").Activate
    Selection.EntireColumn.Hidden = True
    Columns("W:AF").Select
    Range("AF1").Activate
    Selection.EntireColumn.Hidden = True
   
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$AG$9999").AutoFilter Field:=33
    Sheets(Array("Top 20 60+Previous", "ITEMS (2)")).Select
    Sheets("Top 20 60+Previous").Activate
    ActiveWindow.SelectedSheets.Delete
        With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
   
'--------------------------------------------
 
'----------------------------------------------
   
    'Credits
    Sheets("Credits").Select
    With ActiveWindow
        .FreezePanes = False ' Unfreeze any existing frozen panes
        .SplitRow = 1       ' Set the row to freeze
        .FreezePanes = True ' Apply the freeze panes setting
    End With
    Selection.FormulaArray = "Unit"
    Range("I1").Select
    Selection.FormulaArray = "Last Bill Ref"
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Selection.AutoFilter
    With Selection
       .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:F").EntireColumn.AutoFit
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    Columns("Q:T").Select
    Selection.EntireColumn.Hidden = True
    Columns("V:X").Select
    Selection.EntireColumn.Hidden = True
    Columns("Y:Z").Select
    Selection.Cut
    Columns("P:Q").Select
    Selection.Insert Shift:=xlToRight
    Columns("AA:AF").Select
    Selection.EntireColumn.Hidden = True
    Columns("AG:AG").ColumnWidth = 17.14
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
       .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("AG2").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=XLOOKUP(RC[-27],ITEMS!C[-27],ITEMS!C,0)"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
   
'---------------------------------CUSTOMER TOTALS--------------------------------------------------------------------------------------------------------
    Dim ct As Worksheet
    Set ct = Sheets("Customer Totals")
 
    ' Freeze panes
    With ct
        .Activate
        ActiveWindow.SplitRow = 1 ' Set the row to freeze
        ActiveWindow.FreezePanes = True ' Apply the freeze panes setting
    End With
 
    ' Formatting cells
    With ct.Cells
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
 
        ' Removing diagonal borders
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
 
        ' Setting border properties
        Dim border As Variant
        For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
            With .Borders(border)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        Next border
 
        ' Autofit columns
        .EntireColumn.AutoFit
    End With
   
    ' Check if AutoFilter is already applied
    If Not ct.AutoFilterMode Then
        ' Apply AutoFilter to the first row (header row)
        ct.Rows(1).AutoFilter
    End If
   
    'Reset to first sheet
    Sheets("ITEMS").Select
    Range("A2").Select
    Application.DisplayAlerts = True
End Sub
