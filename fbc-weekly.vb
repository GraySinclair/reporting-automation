Sub FBC_Report()
    Dim wb As Workbook
    Dim wsFBCData As Worksheet
    Dim wsItems As Worksheet
    Dim wsFBCReport As Worksheet

    Dim filterCriteria As String
    Dim copyRange As Range

'TODO: Set Refs