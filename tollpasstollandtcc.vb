Sub CorpBillTollAndTCCFees()
    Dim tollfee As Worksheet
    Dim tccfee As Worksheet
    Dim access As Worksheet
    Dim historic As Worksheet
    'TODO: Add error handling
    'TODO: Add check to ensure balances match prior to processing data. Throw alert and exit sub if they do not match.
    'TODO: Trim Unit lookup without checking individual cells

    ' PreWork ----------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
 
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
   
    Set tollfee = ActiveWorkbook.Sheets("Corporate Billing Toll Fee")
    Set tccfee = ActiveWorkbook.Sheets("CorpBilling TCC Fee @new rate")
    Set access = ActiveWorkbook.Sheets("Master Access File")
    Set historic = ActiveWorkbook.Sheets("Historic File")
   
    RemoveBlankColumns tollfee, 20
    RemoveBlankColumns tccfee, 26
   
    ' Delete rows after the last used row
    DeleteExtraRows tollfee
    DeleteExtraRows tccfee
    DeleteExtraRows access
    DeleteExtraRows historic
   
    CreateTable historic, "historictable"
    CreateTable access, "accesstable"
    CreateTable tollfee, "tollfeetable"
    CreateTable tccfee, "tccfeetable"
   
    AddColumnsToTable tollfee, 8, Array("BA", "Frequency", "City", "State", "PO", "PO1", "PO2", "Unit")
    AddColumnsToTable tccfee, 2, Array("BA", "Frequency")
   
    'Luca, does the historic file need to be filtered for the correct info to top of file or as is?
    ' Set formula in (sheet, column, formula)
    SetFormula tollfee, "City", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Rental City],0)"
    SetFormula tollfee, "State", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Rental State],0)"
    SetFormula tollfee, "PO", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Claim '# Field],0)"
    SetFormula tollfee, "PO1", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[PO 1],0)"
    SetFormula tollfee, "PO2", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[PO 2],0)"
    SetFormula tollfee, "BA", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[BA'#],0)"
    SetFormula tollfee, "Frequency", "=XLOOKUP([@[BA]],accesstable[BA],accesstable[Frequency],0)"
    SetFormula tollfee, "Unit", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Veh Unit Nbr],0)"
    SetFormula tollfee, "License Plate", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[License Plate '#],0)"
   
    SetFormula tccfee, "BA", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[BA'#],0)"
    SetFormula tccfee, "Frequency", "=XLOOKUP([@[BA]],accesstable[BA],accesstable[Frequency],0)"
   
    RemoveColumnsByHeaders tollfee, Array("BillingRefNum", "Brand", "CheckOutLocation", "Lic State", "Invoice Ending")
    RemoveColumnsByHeaders tccfee, Array("BillingRefNum", "Brand", "CheckOutLocation", "Transponder HTA ID", "Transponder_Install_Date", "IsTag_Installed", "Is_EzPass_Region_Toll", "Invoice Ending")
   
    MoveColumnBeforeX tollfee, "RA#", "CorpID"
    MoveColumnBeforeX tollfee, "BA", "AccountName"
    MoveColumnBeforeX tollfee, "Frequency", "AccountName"
    MoveColumnBeforeX tollfee, "Unit", "Toll Ref ID"
    MoveColumnBeforeX tollfee, "RA#", "CorpID"
    MoveColumnBeforeX tollfee, "BA", "AccountName"
    MoveColumnBeforeX tollfee, "Frequency", "AccountName"
   
    MoveColumnBeforeX tccfee, "RA#", "CorpID"
    MoveColumnBeforeX tccfee, "BA", "AccountName"
    MoveColumnBeforeX tccfee, "Frequency", "AccountName"
   
    FormatTable tollfee
    FormatTable tccfee
   
    ' Formats columns ws, format(VBA), Array(column names)
    FormatTableColumn tollfee, "@", Array("BA", "Frequency", "City", "State", "PO", "PO1", "PO2", "Unit")
    FormatTableColumn tollfee, "$0.00", Array("Toll Amount")
    FormatTableColumn tollfee, "mm/dd/yyyy", Array("VehicleRental_CheckOutDate", "VehicleRental_CheckInDate")
   
    FormatTableColumn tollfee, "mm/dd/yyyy hh:mm:ss", Array("Toll Date")
   
    
    ' Cleanup ----------------------------------------
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ' force a recalculation prior to pasting as values 
    Application.Calculate
    
    CopyPasteTableAsValues tollfee
End Sub
' Create a table from a given range and name it
Sub CreateTable(ws As Worksheet, tblName As String)
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
 
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
   
    ' Remove filters if they exist
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
        ws.AutoFilterMode = False
    End If
   
    ' Add the table
    With ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tblRange, XlListObjectHasHeaders:=xlYes)
        .Name = tblName
        .TableStyle = ""
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
' Add columns to a table
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
' Move a column before another column
Sub MoveColumnBeforeX(ws As Worksheet, movingcol As String, beforecol As String)
    Dim tbl As ListObject
    Dim rng As Range
   
    Set tbl = ws.ListObjects(1)
    Set rng = tbl.ListColumns(movingcol).Range
   
    rng.Cut
    tbl.ListColumns(beforecol).Range.Insert Shift:=xlToRight
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
    dataRange.Font.Size = 11
    dataRange.Columns.AutoFit
    dataRange.HorizontalAlignment = xlCenter
    dataRange.VerticalAlignment = xlCenter
End Sub
' Format columns in a table
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
' Copy and paste a table as values
Sub CopyPasteTableAsValues(ws As Worksheet)
    Dim tbl As ListObject
    Dim tblRange As Range
   
    Set tbl = ws.ListObjects(1)
    Set tblRange = tbl.Range
   
    tblRange.Copy
    tblRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub