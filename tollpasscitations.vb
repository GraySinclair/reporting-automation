Sub CorpBillCitationAdminFees()
    'worksheets
    Dim tssfee As Worksheet
    Dim tsstotal As Worksheet
    Dim access As Worksheet
    Dim historic As Worksheet
   
    ' PreWork ----------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Turn off automatic calculation
    Application.EnableEvents = False ' Disable events temporarily
 
    Set tssfee = ActiveWorkbook.Sheets("TSS Fee")
    Set tsstotal = ActiveWorkbook.Sheets("TSS Fee_Total")
    Set access = ActiveWorkbook.Sheets("Master Access File")
    Set historic = ActiveWorkbook.Sheets("Historic File")
   
    ' Remove blank columns from tssfee and tsstotal
    RemoveBlankColumns tssfee, 23
    RemoveBlankColumns tsstotal, 19
   
    ' Delete extra rows after last used row
    DeleteExtraRows tssfee
    DeleteExtraRows tsstotal
    DeleteExtraRows access
    DeleteExtraRows historic
   
    ' Create tables dynamically for historic, access, tssfee, tsstotal
    CreateTable historic, "historictable"
    CreateTable access, "accesstable"
    CreateTable tssfee, "tssfeetable"
    CreateTable tsstotal, "tsstotaltable"
   
    ' Add Columns
    AddColumnsToTable tssfee, 4, Array("BA", "Frequency", "Unit", "Datetime")
    AddColumnsToTable tsstotal, 2, Array("BA", "Frequency")
   
    ' Remove specific columns from tssfee and tsstotal based on headers
    RemoveColumnsByHeaders tssfee, Array("BillingRefNum", "Brand", "CheckOutLocation", "Lic State", "Invoice Ending")
    RemoveColumnsByHeaders tsstotal, Array("BillingRefNum", "Brand", "CheckOutLocation", "Lic State", "Usage Days", "Invoice Ending")
   
    ' Set formula in (sheet, column, formula)
    SetFormula tssfee, "City", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Rental City],0)"
    SetFormula tssfee, "State", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Rental State],0)"
    SetFormula tssfee, "PO", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Claim '# Field],0)"
    SetFormula tssfee, "PO1", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[PO 1],0)"
    SetFormula tssfee, "PO2", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[PO 2],0)"
    SetFormula tssfee, "BA", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[BA'#],0)"
    SetFormula tssfee, "Frequency", "=XLOOKUP([@[BA]],accesstable[BA],accesstable[Frequency],0)"
    SetFormula tssfee, "Unit", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Veh Unit Nbr],0)"
    SetFormula tssfee, "Datetime", "=TEXT([@[Toll Date]],""mm/dd/yyyy"")&TEXT([@[ISSUE TIME]],"" hh:mm:ss"")"
   
    'START -----------------------------------
 
 
            ' Move the "RA#" column to be the first column
            ranumcol.Range.Cut
            firstcol.Range.Insert Shift:=xlToRight
 
 
'tsstotal table-----------------------------------------------
 
    If Not corpIDColumn Is Nothing Then
        ' Get the index of the CorpID column
        corpIDIndex = corpIDColumn.Index
 
        tbl.ListColumns(corpIDIndex + 11).Delete
       
        ' Find the column "RA#" in the table
        On Error Resume Next
        Set ranumcol = tbl.ListColumns("RA#")
        On Error GoTo 0
       
        ' Check if the "RA#" column exists
        If Not ranumcol Is Nothing Then
            ' Get the first column in the table
            Set firstcol = tbl.ListColumns(1)
       
            ' Move the "RA#" column to be the first column
            ranumcol.Range.Cut
            firstcol.Range.Insert Shift:=xlToRight
        End If
    End If
   
    'TSSFEE XLOOKUPS------------------------------------------------
    'question: does the historic file need to be filtered for the correct info?
 
   
    'XLOOKUP FOR TSSFEE - Unit# from Historic
    'NEEDS CONVERT TO TEXT & Trim BEFORE COPY/PASTE values
    formula = "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[Veh Unit Nbr],0)"
    tbl.ListColumns("Unit #").DataBodyRange.FormulaLocal = formula
    tbl.ListColumns("Unit #").DataBodyRange.NumberFormat = "@"
   
 
    'functional end----------------------------------------------
   
    Set tbl = tsstotal.ListObjects("tsstotaltable")
   
    'XLOOKUP FOR TSSTOTAL - BA col from Historic
    formula = "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[BA'#],0)"
    tbl.ListColumns("BA").DataBodyRange.formula = formula
    tbl.ListColumns("BA").DataBodyRange.NumberFormat = "@"
 
    'XLOOKUP FOR TSSTOTAL - Frequency col from access
    formula = "=XLOOKUP([@[BA]],accesstable[BA],accesstable[Frequency],0)"
    tbl.ListColumns("Frequency").DataBodyRange.formula = formula
    tbl.ListColumns("Frequency").DataBodyRange.NumberFormat = "@"
   
    ' Re-enable automatic calculation after macro is done
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ' Force a full calculation to update XLOOKUP results
    Application.Calculate
   
    ' Copy the entire table range (including headers)
    tbl.Range.Copy
    ' Paste values only, effectively replacing formulas with their values
    tbl.Range.PasteSpecial Paste:=xlPasteValues
    ' Clear the clipboard (optional)
    Application.CutCopyMode = False
    ' AutoFit all columns in the table
    tbl.Range.Columns.AutoFit
   
    ' AutoFit all columns in the table
    tbl.Range.Columns.AutoFit
    tbl.DataBodyRange.NumberFormat = "@"
    ' Cleanup ----------------------------------------
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ' remove pre-indexing when complete
    '  ----------------------------------------
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
            col = col - 1 ' Adjust column index after deletion
        End If
    Next col
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
Sub AddColumnsToTable(ws As Worksheet, numColumns As Integer, columnNames As Variant)
    Dim tbl As ListObject
    Dim i As Integer
    Dim newColumn As ListColumn
   
    ' Check if table exists on the worksheet
    On Error Resume Next
    Set tbl = ws.ListObjects(1)  ' Assuming there is only one table on the sheet
    On Error GoTo 0
   
    ' Add the specified number of columns with the given names
    For i = 1 To numColumns
        ' Add the column
        Set newColumn = tbl.ListColumns.Add
       
        ' Set the name for the new column
        newColumn.Name = columnNames(i - 1)
    Next i
End Sub
Sub SetFormula(ws As Worksheet, colName As String, formula As String)
    Dim tbl As ListObject
    Dim col As ListColumn
   
    Set tbl = ws.ListObjects(1)
    Set col = tbl.ListColumns(colName)
   
    ' Set formula in the body of the column
    col.DataBodyRange.formula = formula
End Sub