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
   
 
    RemoveBlankColumns tssfee, 23
    RemoveBlankColumns tsstotal, 19
   
    ' deleting blank columns if they exist in tssfees and tsstotal---------------------------------
    'Dim col As Long
    ' Loop through columns from X to A (backward to avoid shifting)
    ' deleting blank columns in tssfee
'    For col = 23 To 1 Step -1
'        If WorksheetFunction.CountA(tssfee.Columns(col)) = 0 Then
'            tssfee.Columns(col).Delete
'        End If
'    Next col
   
    ' deleting blank columns in tsstotal
    For col = 19 To 1 Step -1
        If WorksheetFunction.CountA(tsstotal.Columns(col)) = 0 Then
            tsstotal.Columns(col).Delete
        End If
    Next col
End Sub