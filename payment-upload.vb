Sub PU()
    Dim awb As Workbook
    Dim itemsheet As Worksheet
    Dim puSheet As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim cell As Range

    Set awb = ActiveWorkbook
    Set itemsheet = awb.Sheets("ITEMS")

    ' Check if "PU" sheet exists
    On Error Resume Next
    Set puSheet = awb.Sheets("PU")
    On Error GoTo 0

    ' If "PU" doesn't exist, create it
    If puSheet Is Nothing Then
        Set puSheet = awb.Sheets.Add
        puSheet.Name = "PU"
    End If
    'TODO: Make array once functionality is proven
    With puSheet
        .Cells(1, 1).Value = "Payment Reference"
        .Cells(2, 1).Value = "Check Amount/EFT AMOUNT"
        .Cells(3, 1).Value = "Payment Date"
        .Cells(5, 1).Value = "Currency"
        .Cells(5, 2).Value = "Payer First Name"
        .Cells(5, 3).Value = "Payer Last/Company Name"
        .Cells(5, 4).Value = "Business Unit"
        .Cells(5, 5).Value = "Customer ID"
        .Cells(5, 6).Value = "Item Amount"
        .Cells(5, 7).Value = "Item"
        .Cells(5, 8).Value = "Reference Qualifier"
        .Cells(6, 1).Value = "USD"
        .Cells(6, 4).Value = "300NB"
        .Cells(6, 8).Value = "I"

        With .Range("A1:B3")
            With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
            End With
        End With

        .Columns("A").ColumnWidth = 28
        .Columns("B").ColumnWidth = 17
        .Columns("C").ColumnWidth = 26
        .Columns("D:F").ColumnWidth = 13
        .Columns("G").ColumnWidth = 19
        .Columns("H").ColumnWidth = 13

        With .Cells
            .WrapText = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("A5:H5")
            .WrapText = True
            .Interior.Color = RGB(128, 128, 128) ' Gray color
            .Font.Bold = True
        End With

        .Columns("A:D").NumberFormat = "@"
        .Columns("E").NumberFormat = "0"
        .Columns("F").NumberFormat = "$#,##0.00"
        .Columns("G:H").NumberFormat = "@"
    End With 'End PU Sheet

    ' Find the last row in column F
    lastRow = itemsheet.Cells(itemsheet.Rows.Count, "F").End(xlUp).Row
    ' Set the destination starting row in "PU"
    destRow = 6

    With itemsheet
        For Each cell In itemsheet.Range("AG1:AG" & lastRow)
            If cell.Value = "test" Then
                puSheet.Cells(destRow, "G").Value = itemsheet.Cells(cell.Row, "F").Value
                destRow = destRow + 1
            End If
        Next cell
    End With
'    TURN OFF UNTIL FIXED@@@@@
'    Clear previous data in column G from row 6
'    puSheet.Range("G6:G" & puSheet.Cells(Rows.Count, "G").End(xlUp).Row).ClearContents

'    ' Loop through each cell in column AG
'    For Each cell In ThisWorkbook.Sheets(1).Range("AG1:AG" & lastRow)
'        If cell.Value = "test" Then
'            ' Copy corresponding value from column F to column G in "PU"
'            puSheet.Cells(destRow, "G").Value = ThisWorkbook.Sheets(1).Cells(cell.Row, "F").Value
'            destRow = destRow + 1
'        End If
'    Next cell
End Sub