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