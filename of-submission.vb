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