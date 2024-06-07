' PreWork --------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
   
'--------------------------------------------------------------------------
'FedEx Remit                                                              '
'--------------------------------------------------------------------------
   
    Dim aws As Worksheet
    Set aws = ActiveWorkbook.ActiveSheet
   
    With aws
        .Range("B1:H1").UnMerge
        ' Insert 3 blank columns at Column A
        .Columns("C:E").Insert Shift:=xlToRight
       
        .Columns("A").NumberFormat = "@"
        .Columns("B").NumberFormat = "mmm dd, yyyy"
        .Columns("C:F").NumberFormat = "@"
        .Columns("G").NumberFormat = "#,##0.00"
        .Columns("H").NumberFormat = "@"
        .Columns("I:J").NumberFormat = "#,##0.00"