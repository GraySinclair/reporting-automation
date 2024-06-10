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

         With .Columns("A") 'FieldInfo:=Array(Array(0, 2)... The "0" represents where the data begins aka where the TTC will split the data the "2" represents the data output. (Options- 1=General, 2=Text(string), 3=Date(if applicable))
            .TextToColumns Destination:=.Cells(1, 3), DataType:=xlFixedWidth, _
                FieldInfo:=Array(Array(0, 2), Array(4, 2), Array(10, 2)), _
                TrailingMinusNumbers:=True
        End With
       
        .Cells(2, 3).Value = "GPBR"
        .Cells(2, 4).Value = "Ticket #"
        .Cells(2, 5).Value = "Submission Type"
       
        With Cells 'Applies to all cells
            .WrapText = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        
        
        .Cells(2, 12).Value = "Notes"
' Copy formatting from column K to column L
        .Columns("K").Copy
        .Columns("L").PasteSpecial Paste:=xlPasteFormats
        .Columns("K").NumberFormat = "#,##0.00"
        .Columns("L").NumberFormat = "@"
       
        ' Clear the clipboard to avoid the "marching ants" border
        Application.CutCopyMode = False
   
        Dim sumRange As Double
        Dim targetCell As Range
        Dim formattedSum As String
       
        ' Set the target cell where the output will be written
        Set targetCell = aws.Range("L3")
   