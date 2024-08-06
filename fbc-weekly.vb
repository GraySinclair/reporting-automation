Sub FBC_Report()
    Dim wb As Workbook
    Dim wsFBCData As Worksheet
    Dim wsItems As Worksheet
    Dim wsFBCReport As Worksheet

    Dim filterCriteria As String
    Dim copyRange As Range

    ' References
    Set wb = ActiveWorkbook
    Set wsFBCData = Nothing
    On Error Resume Next
    Set wsFBCData = wb.Sheets("FBC Data")
    On Error GoTo 0
    Set wsFBCReport = Nothing
    On Error Resume Next
    Set wsFBCReport = wb.Sheets("FBC Report")
    On Error GoTo 0
    Set wsItems = Nothing
    On Error Resume Next
    Set wsItems = wb.Worksheets("ITEMS")
    On Error GoTo 0

    ' Create worksheet "FBC Data" if it doesn't exist
    If wsFBCData Is Nothing Then
        Set wsFBCData = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsFBCData.Name = "FBC Data"
    End If