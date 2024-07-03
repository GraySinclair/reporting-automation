Sub OF_Submission()
    Dim it1 As Worksheet
    Dim OFS As Worksheet
  
' PreWork --------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationManual
  
    Set it1 = ActiveWorkbook.Worksheets("ITEMS")
'--------------------------------------------------------------------------
'OF Submission                                                            |
'--------------------------------------------------------------------------
    'Create worksheet named "OF Submission" if it doesn't exist
    If OFS Is Nothing Then
        On Error Resume Next
        ActiveWorkbook.Sheets("OF Submission").Delete
        Set OFS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        OFS.Name = "OF Submission"
    End If
' OF report setup
    ' Set column headers in "OF Submission" worksheet
    With OFS
        .Cells.Clear  ' Clear existing content in case of macro rerun
      
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
            .Columns.AutoFilter 'Disable on mac dev environment
            .Columns.AutoFit
        End With