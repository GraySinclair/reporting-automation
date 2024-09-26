Sub Fedex_Remit()
    Dim aws As Worksheet
    Set aws = ActiveWorkbook.ActiveSheet
   
    Application.DisplayAlerts = False
    With aws
        ' Insert 3 blank columns at Column A
        .Columns("B:D").Insert Shift:=xlToRight
       
        With .Columns("A")
            .TextToColumns Destination:=.Cells(1, 2), DataType:=xlFixedWidth, _
                FieldInfo:=Array(Array(0, 1), Array(4, 1), Array(10, 1)), _
                TrailingMinusNumbers:=True
        End With
       
        .Cells(2, 2).Value = "GPBR"
        .Cells(2, 3).Value = "Ticket #"
        .Cells(2, 4).Value = "Submission Type"
       
        With .Columns
            .AutoFit
        End With
        With Cells 'Applies to all cells
            .WrapText = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With .Columns
            .AutoFit
        End With
    End With
    Application.DisplayAlerts = True
End Sub