Sub CorpBillCitationAdminFees()
    'worksheets
    Dim tssfee As Worksheet
    Dim tsstotal As Worksheet
    Dim access As Worksheet
    Dim historic As Worksheet
  
    ' PreWork ----------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual 
    Application.EnableEvents = False 
 
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True

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
  
    'Luca, does the historic file need to be filtered for the correct info to top of file or as is?
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
    SetFormula tsstotal, "BA", "=XLOOKUP([@[RA'#]],historictable[Ticket '#],historictable[BA'#],0)"
    SetFormula tsstotal, "Frequency", "=XLOOKUP([@[BA]],accesstable[BA],accesstable[Frequency],0)"
 
    ' Remove specific columns from tssfee and tsstotal based on headers
    RemoveColumnsByHeaders tssfee, Array("BillingRefNum", "Brand", "CheckOutLocation", "Lic State", "Invoice Ending")
    RemoveColumnsByHeaders tsstotal, Array("BillingRefNum", "Brand", "CheckOutLocation", "Lic State", "Usage Days", "Invoice Ending")
   
    ' Can you use an Array to move multiple cols example: BA & freq together before AccountName?
    MoveColumnBeforeX tssfee, "RA#", "CorpID"
    MoveColumnBeforeX tssfee, "BA", "AccountName"
    MoveColumnBeforeX tssfee, "Frequency", "AccountName"
    MoveColumnBeforeX tssfee, "Unit", "Toll Ref ID"
    MoveColumnBeforeX tssfee, "Toll Date", "Toll Road"
    MoveColumnBeforeX tsstotal, "RA#", "CorpID"
    MoveColumnBeforeX tsstotal, "BA", "AccountName"
    MoveColumnBeforeX tsstotal, "Frequency", "AccountName"
   
    ' removes weird formats after transformations to restore blank state. could  potentially just use tbl.clearformats as long as tests dont change certain datum.
    FormatTable tssfee
    FormatTable tsstotal
   
    ' Formats columns ws, format(VBA), Array(column names)
    FormatTableColumn tssfee, "@", Array("BA", "Frequency", "City", "State", "PO", "PO1", "PO2", "Unit")
    FormatTableColumn tssfee, "0.00", Array("Toll Amount", "TSS Fee Amt")
    FormatTableColumn tssfee, "mm/dd/yyyy", Array("VehicleRental_CheckOutDate", "VehicleRental_CheckInDate")
   
    

    ' Cleanup ----------------------------------------
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.Calculate
   
  
    CopyPasteTableAsValues tssfee
   
    RemoveColumnsByHeaders tssfee, Array("ISSUE TIME", "Toll Date")
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
            ws.ShowAllData ' Remove active filter
        End If
        ws.AutoFilterMode = False ' Turn off AutoFilter
    End If
   
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