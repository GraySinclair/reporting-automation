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
   
        ' Find the last row in column K with data minus total row
        lastRowk = aws.Cells(Rows.Count, "K").End(xlUp).Row - 1
   
        ' Calculate the sum of the range K3:K(lastRowk)
        sumRange = Application.WorksheetFunction.Sum(aws.Range("K3:K" & lastRowk))
        formattedSum = Format(sumRange, "$#,##0.00")
       
        ' Create the output string
        targetCell.Value = "Fwd2ASKSSbyGray: " & Format(Date, "mm/dd/yy") & " - Amt:" & formattedSum
        .Range("L3:L" & lastRowk).FillDown
       
        With .Columns
            .AutoFit
        End With
       
    End With

        Dim targetWorkbook As Workbook
        Dim targetSheet As Worksheet
        Dim lastRow As Long
        Dim lastRowToCopy As Long
        Dim targetLastRow As Long
   
        ' Find the last row in the source sheet
        lastRow = aws.Cells(aws.Rows.Count, 1).End(xlUp).Row
   
        ' Define the range to copy (from row 3 to the last row)
        lastRowToCopy = lastRow - 2 ' Subtracting 2 to start from row 3
   
        ' Open the target workbook
        Set targetWorkbook = Workbooks.Open("W:\Corporate DB Program\Account Specialists\Gray\FedEx\Payment Archive\FedexPmtHist.xlsx")
   
        ' Set the target sheet (assuming you want to copy to the first sheet)
        Set targetSheet = targetWorkbook.Sheets(1)
   
        ' Find the last row in the target sheet
        targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
   
        ' Copy the range from the source sheet to the target sheet
        aws.Rows("3:" & lastRow).Copy Destination:=targetSheet.Rows(targetLastRow)
   
        ' Optional: Save and close the target workbook
        targetWorkbook.Save
        targetWorkbook.Close
   
'Cleanup --------------------------------------------------------------------------
        Set targetSheet = Nothing
        Set targetWorkbook = Nothing
        Set aws = Nothing
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        'Application.Calculation = xlCalculationManual
        'MsgBox "Rows copied successfully!"
End Sub