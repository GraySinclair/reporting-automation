Sub AR_Setup()
'
' AR Setup Macro
'
' TO DO:
 
' PreWork --------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
   
'--------------------------------------------------------------------------
'ITEMS 1
'------------------------------------------------------------------------
    Dim it1 As Worksheet
    Set it1 = ActiveWorkbook.Sheets("ITEMS")
    With it1
        .Activate
 
        'Format Removal
'        With Columns("A:AG")
'            .ClearFormats
'        End With