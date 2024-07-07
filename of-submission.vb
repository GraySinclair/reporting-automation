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
'END OF NEW SHEET SETUP------------------------------------------------------------------------------
'--------------------------------------------------------------------------
'ITEMS 1                                                                  |
'--------------------------------------------------------------------------
    'New Code
    Dim lastrowinb As Long
    Dim row2 As Long
    Dim i As Long
  
    ' Find the last row in column B
    lastrowinb = it1.Cells(it1.Rows.Count, "B").End(xlUp).Row
    row2 = 2
   
    'Leave ITEMS sheet filtered post-run macro
   
        
        
        it1.ShowAllData 'Testif this works otherwise add ActiveSheet.
        it1.Range("A1").AutoFilter Field:=2, Criteria1:="XZ4312Y"
        'Check to see if there are visible cells in the filtered range
        it1.Range("A1").AutoFilter Field:=33, Criteria1:="="
    ' Loop through each row
    For i = 1 To lastrowinb
        ' Check if column B contains "XZ4312Y" and column AG is blank
        If it1.Cells(i, "B").Value = "XZ4312Y" And it1.Cells(i, "AG").Value = "" Then
            ' Copy BA#s from column D to column J in OFS
            OFS.Cells(row2, "J").Value = it1.Cells(i, "D").Value
            ' Copy tickets
            OFS.Cells(row2, "K").Value = it1.Cells(i, "F").Value
            ' Copy amt
            OFS.Cells(row2, "O").Value = it1.Cells(i, "T").Value
            ' Copy amt
            OFS.Cells(row2, "P").Value = it1.Cells(i, "T").Value
            ' Copy amt
            OFS.Cells(row2, "Q").Value = it1.Cells(i, "T").Value
            'Write notes in notes column 'NEED TO ADD OF SUB NUMBER ----------------------------------------------------------------@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'it1.Cells(i, "AG").Value = "OF Submission " & Format(Date, "mm.dd.yy")
            row2 = row2 + 1 ' Move to the next row in OFS
        End If
    Next i