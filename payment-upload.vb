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
