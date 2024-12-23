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
   
End Sub