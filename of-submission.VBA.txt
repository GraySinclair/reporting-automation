Sub OF_Submission()
    Dim wb As Workbook
    Dim wsOFSubmission As Worksheet
    Dim wsItems As Worksheet
    Dim headerRow As Range
    Dim copyTickets As Range
    Dim copyBas As Range
    Dim copyAmt As Range
    Dim lastRow As Long
    Dim colARange As Range
    Dim colBRange As Range
    Dim colNRange As Range
    Dim colRRange As Range
    Dim colVRange As Range
    Dim colERange As Range
    Dim colFRange As Range
    Dim colGRange As Range
    Dim colHRange As Range
'------------------------
    Dim tomorrowDate As String
    Dim noterange As Range
   
    ' Set reference to the workbook
    Set wb = ActiveWorkbook
   
    ' Set references
    Set wsOFSubmission = Nothing
    On Error Resume Next
    Set wsOFSubmission = wb.Sheets("OF Submission")
    On Error GoTo 0
    Set wsItems = Nothing
    On Error Resume Next
    Set wsItems = ActiveWorkbook.Worksheets("ITEMS")
    On Error GoTo 0
   
    Set headerRow = wsItems.Range("A1:AG1")
    
    ' Create worksheet "OF Submission" if it doesn't exist
    If wsOFSubmission Is Nothing Then
        Set wsOFSubmission = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsOFSubmission.Name = "OF Submission"
    End If
   
    ' Set column headers in "OF Submission" worksheet
    With wsOFSubmission
        .Cells.Clear  ' Clear existing content
       
        ' Set column headers
        .Cells(1, 1).Value = "Program Name"
        .Cells(1, 2).Value = "Expense Name"
        .Cells(1, 3).Value = "Contractor ID"
        .Cells(1, 4).Value = "Client Code"
        .Cells(1, 5).Value = "CMS Client ID"
        .Cells(1, 6).Value = "Business Name"
        .Cells(1, 7).Value = "First Name"
        .Cells(1, 8).Value = "Last Name"
        .Cells(1, 9).Value = "Approved"
        .Cells(1, 10).Value = "BILLACCOUNT"
        .Cells(1, 11).Value = "Vendor Ref. 1"
        .Cells(1, 12).Value = "Vendor Ref. 2"
        .Cells(1, 13).Value = "Is Active"
        .Cells(1, 14).Value = "Expense Type"
        .Cells(1, 15).Value = "Requested Amount"
        .Cells(1, 16).Value = "Target Amount"
        .Cells(1, 17).Value = "Balance Amount"
        .Cells(1, 18).Value = "Number of Payments"
        .Cells(1, 19).Value = "Creation Date"
        .Cells(1, 20).Value = "Distribution Date"
        .Cells(1, 21).Value = "Contractor Reference"
        .Cells(1, 22).Value = "Lookup Method"
        .Cells(2, 1).Value = "TopHAT Logistical Solutions, LLC."
        .Cells(2, 2).Value = "Enterprise Direct Truck Lease"
        .Cells(2, 14).Value = "Installment"
        .Cells(2, 18).Value = "1"
        .Cells(2, 22).Value = "cmsclientid"
       
        With .Range("A:V")
            ' Change font to Calibri for the entire column range
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Columns.AutoFilter
            .Columns.AutoFit
        End With
       
        ' Format columns
        .Range("O:Q").NumberFormat = "$#,##0.00"
        .Range("O:Q").HorizontalAlignment = xlCenter
        .Range("R2:V2").HorizontalAlignment = xlLeft
        .Range("R2:V2").Font.Italic = True
       
        ' Format the header row (assuming it's row 1)
        With .Range("A1:V1")
            .Interior.Color = RGB(192, 192, 192) ' Gray color
            .Font.Bold = False ' Make header text bold
        End With
    End With
    'END OF SHEET SETUP------------------------------------------------------------------------------
   
    With wsItems.Range("$A$1:$AG$9999")
   
        On Error Resume Next
        ActiveSheet.ShowAllData
 
'-------------------------------------------COPY FILTERS----------------------------------------
        .AutoFilter Field:=2, Criteria1:="XZ4312Y"
        'Check to see if there are visible cells in the filtered range
        On Error Resume Next
        .AutoFilter Field:=33, Criteria1:="="
        On Error Resume Next
       
'--------------------------ERRRRR@@@@@@@@@@
        ' Get tomorrow's date
        'tomorrowDate = Format(Date + 1, "mm.dd.yy")  ' Adjust date format as needed
   
        'Set noterange = .Offset(1).Columns("AG").SpecialCells(xlCellTypeVisible)
        'On Error GoTo 0
        ' Write the string to cell AH1
        'noterange.Value = "OF Submission #1xx " & tomorrowDate
'-------------------------------------------COPY TICKETS----------------------------------------
        Set copyTickets = .Resize(.Rows.Count - 1).Offset(1).Columns("F").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        ' Copy visible cells to destination sheet starting at K2
        copyTickets.Copy
       wsOFSubmission.Range("K2").PasteSpecial Paste:=xlPasteValues
        ' Clear clipboard to avoid memory issues
        Application.CutCopyMode = False
       
'-------------------------------------------COPY BA#S----------------------------------------
        Set copyBas = .Resize(.Rows.Count - 1).Offset(1).Columns("D").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        ' Copy visible cells to destination sheet starting at J2
        copyBas.Copy
        wsOFSubmission.Range("J2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
 
'-------------------------------------------COPY AMT----------------------------------------
        Set copyAmt = .Resize(.Rows.Count - 1).Offset(1).Columns("V").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        ' Copy visible cells to destination sheet starting at O2
        copyAmt.Copy
        wsOFSubmission.Range("O2").PasteSpecial Paste:=xlPasteValues
        wsOFSubmission.Range("P2").PasteSpecial Paste:=xlPasteValues
        wsOFSubmission.Range("Q2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End With
 
    ' Find the last used row in column A
    lastRow = wsOFSubmission.Cells(wsOFSubmission.Rows.Count, "K").End(xlUp).Row
   
    'TOPHAT INFO XLOOKUPS
    Range("E2").Formula = _
    "=XLOOKUP(J2, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$BH:$BH, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$DM:$DM, 0)"
    Range("F2").Formula = _
    "=XLOOKUP(J2, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$BH:$BH, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$DN:$DN, 0)"
    Range("G2").Formula = _
    "=XLOOKUP(J2, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$BH:$BH, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$DO:$DO, 0)"
    Range("H2").Formula = _
    "=XLOOKUP(J2, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$BH:$BH, '[Tophat Acc List.xlsx]XZ4312Y(IC)'!$DP:$DP, 0)"
 
 
    ' Check if there are enough rows to fill down
    If lastRow >= 2 Then
        ' Define the fill range from Row 2 to the last used row in each column
        Set colARange = wsOFSubmission.Range("A2:A" & lastRow)
        Set colBRange = wsOFSubmission.Range("B2:B" & lastRow)
        Set colNRange = wsOFSubmission.Range("N2:N" & lastRow)
        Set colRRange = wsOFSubmission.Range("R2:R" & lastRow)
        Set colVRange = wsOFSubmission.Range("V2:V" & lastRow)
        Set colERange = wsOFSubmission.Range("E2:E" & lastRow)
        Set colFRange = wsOFSubmission.Range("F2:F" & lastRow)
        Set colGRange = wsOFSubmission.Range("G2:G" & lastRow)
        Set colHRange = wsOFSubmission.Range("H2:H" & lastRow)
       
        ' Fill down the value from Row 2
        wsOFSubmission.Range("A2").AutoFill Destination:=colARange, Type:=xlFillDefault
        wsOFSubmission.Range("B2").AutoFill Destination:=colBRange, Type:=xlFillDefault
        wsOFSubmission.Range("N2").AutoFill Destination:=colNRange, Type:=xlFillDefault
        wsOFSubmission.Range("R2").AutoFill Destination:=colRRange, Type:=xlFillDefault
        wsOFSubmission.Range("V2").AutoFill Destination:=colVRange, Type:=xlFillDefault
        wsOFSubmission.Range("E2").AutoFill Destination:=colERange, Type:=xlFillDefault
        wsOFSubmission.Range("F2").AutoFill Destination:=colFRange, Type:=xlFillDefault
        wsOFSubmission.Range("G2").AutoFill Destination:=colGRange, Type:=xlFillDefault
        wsOFSubmission.Range("H2").AutoFill Destination:=colHRange, Type:=xlFillDefault
    End If
   
    
    With wsOFSubmission
        With Cells
            .Copy
        .PasteSpecial Paste:=xlPasteValues
        End With
        Application.CutCopyMode = False
        Columns("J").Delete
        With .Range("E:H")
            .Columns.AutoFit
        End With
    End With
   
    Set OFS = Workbooks.Add
    wsOFSubmission.Move After:=OFS.Sheets(1)
    Application.DisplayAlerts = False
    OFS.Sheets(1).Delete
   Application.DisplayAlerts = True
 
    ' Open the Save As dialog
    Application.Dialogs(xlDialogSaveAs).Show
End Sub