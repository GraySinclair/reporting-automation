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

    ' Create worksheet "FBC Report" if it doesn't exist
    If wsFBCReport Is Nothing Then
        Set wsFBCReport = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsFBCReport.Name = "FBC Report"
    End If

    ' Filter the range in ITEMS sheet to pull filtered FBC data
    On Error Resume Next
    ActiveSheet.ShowAllData
    With wsItems.Range("$A$1:$AG$9999") 'TODO: change this to a from start to "lastrow" to improve code functionality
        .AutoFilter Field:=2, Criteria1:="XZ55FFT"
        ' Check if there are visible cells in the filtered range
        On Error Resume Next
        Set copyRange = .SpecialCells(xlCellTypeVisible) 'TODO: change from visible cells to requirement based on filters
        On Error GoTo 0
    End With