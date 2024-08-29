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

    With wsFBCReport
    ' If there are visible cells, copy them to FBC Data sheet
    If Not copyRange Is Nothing Then
        copyRange.Copy
        wsFBCData.Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If

    'TODO: place holder hardcoded values, test array functionality here 
    'TODO: test using FBC Account data.xlsx as make-shift database with xlookup functionality
    .Cells.Clear

        ' Set Headers
        .Cells(1, 1).Value = "CID #"
        .Cells(1, 2).Value = "BA #"
        .Cells(1, 3).Value = "Company Name"
        .Cells(1, 4).Value = "On Rent Count"
        .Cells(1, 5).Value = "Bal Amt"
        .Cells(1, 6).Value = "1-30 Days"
        .Cells(1, 7).Value = "31-60 Days"
        .Cells(1, 8).Value = "61-90 Days"
        .Cells(1, 9).Value = "91+ Days"
        .Cells(1, 10).Value = "Notes"

        .Cells(2, 1).Value = "XZ55FFT"
        .Cells(3, 1).Value = "XZ55FFT"
        .Cells(4, 1).Value = "XZ55FFT"
        .Cells(5, 1).Value = "XZ55FFT"
        .Cells(6, 1).Value = "XZ55FFT"
        .Cells(7, 1).Value = "XZ55FFT"
        .Cells(8, 1).Value = "XZ55FFT"
        .Cells(9, 1).Value = "XZ55FFT"
        .Cells(10, 1).Value = "XZ55FFT"
        .Cells(11, 1).Value = "XZ55FFT"
        .Cells(12, 1).Value = "XZ55FFT"
        .Cells(13, 1).Value = "XZ55FFT"
        .Cells(14, 1).Value = "XZ55FFT"
        .Cells(15, 1).Value = "XZ55FFT"
        .Cells(16, 1).Value = "XZ55FFT"
        .Cells(17, 1).Value = "XZ55FFT"
        .Cells(18, 1).Value = "XZ55FFT"
        .Cells(19, 1).Value = "XZ55FFT"
        .Cells(20, 1).Value = "XZ55FFT"
        .Cells(21, 1).Value = "XZ55FFT"
        .Cells(22, 1).Value = "XZ55FFT"
        .Cells(23, 1).Value = "XZ55FFT"
        .Cells(24, 1).Value = "XZ55FFT"
        .Cells(25, 1).Value = "XZ55FFT"
        .Cells(26, 1).Value = "XZ55FFT"
        .Cells(27, 1).Value = "XZ55FFT"
        .Cells(28, 1).Value = "XZ55FFT"
        .Cells(29, 1).Value = "XZ55FFT"
        .Cells(30, 1).Value = "XZ55FFT"
        .Cells(31, 1).Value = "XZ55FFT"
        .Cells(32, 1).Value = "XZ55FFT"
        .Cells(2, 2).Value = "18051500"
        .Cells(3, 2).Value = "18051503"
        .Cells(4, 2).Value = "18051505"
        .Cells(5, 2).Value = "18051506"
        .Cells(6, 2).Value = "18051507"
        .Cells(7, 2).Value = "18051508"
        .Cells(8, 2).Value = "18051509"
        .Cells(9, 2).Value = "18051540"
        .Cells(10, 2).Value = "18051542"
        .Cells(11, 2).Value = "18051544"
        .Cells(12, 2).Value = "18051545"
        .Cells(13, 2).Value = "18051546"
        .Cells(14, 2).Value = "18051548"
        .Cells(15, 2).Value = "18051549"
        .Cells(16, 2).Value = "18051550"
        .Cells(17, 2).Value = "18051552"
        .Cells(18, 2).Value = "18051553"
        .Cells(19, 2).Value = "18051555"
        .Cells(20, 2).Value = "18051557"
        .Cells(21, 2).Value = "18051558"
        .Cells(22, 2).Value = "18051559"
        .Cells(23, 2).Value = "18051562"
        .Cells(24, 2).Value = "18051563"
        .Cells(25, 2).Value = "18051564"
        .Cells(26, 2).Value = "18051566"
        .Cells(27, 2).Value = "18051567"
        .Cells(28, 2).Value = "18051568"
        .Cells(29, 2).Value = "18051569"
        .Cells(30, 2).Value = "18051570"
        .Cells(31, 2).Value = "18051572"
        .Cells(32, 2).Value = "18051573"
        .Cells(2, 3).Value = "Derst Baking Co. (Savannah)"
        .Cells(3, 3).Value = "FBC of Bardstown"
        .Cells(4, 3).Value = "FBC of Batesville"
        .Cells(5, 3).Value = "FBC of Baton Rouge"
        .Cells(6, 3).Value = "FBC of Birmingham"
        .Cells(7, 3).Value = "FBC of Bradenton"
        .Cells(8, 3).Value = "FBC of Denton"
        .Cells(9, 3).Value = "FBC of Denver"
        .Cells(10, 3).Value = "FBC of El Paso"
        .Cells(11, 3).Value = "FBC of Henderson"
        .Cells(12, 3).Value = "FBC of Houston"
        .Cells(13, 3).Value = "FBC of Jacksonville"
        .Cells(14, 3).Value = "FBC of Jamestown"
        .Cells(15, 3).Value = "FBC of Knoxville"
        .Cells(16, 3).Value = "FBC of Lenexa"
        .Cells(17, 3).Value = "FBC of Newton"
        .Cells(18, 3).Value = "FBC of Miami"
        .Cells(19, 3).Value = "Flowers Baking Sales of NorCal"
        .Cells(20, 3).Value = "FBC of New Orleans"
        .Cells(21, 3).Value = "FBC of Norfolk"
        .Cells(22, 3).Value = "FBC of Ohio"
        .Cells(23, 3).Value = "FBC of Oxford"
        .Cells(24, 3).Value = "FBC of Portland"
        .Cells(25, 3).Value = "FBC of San Antonio"
        .Cells(26, 3).Value = "FBC of Thomasville"
        .Cells(27, 3).Value = "FBC of Tyler"
        .Cells(28, 3).Value = "FBC of Villa Rica"
        .Cells(29, 3).Value = "Franklin Baking Co. (Goldsboro)"
        .Cells(30, 3).Value = "Holsum Bakery (Phoenix)"
        .Cells(31, 3).Value = "Lepage Bakery"
        .Cells(32, 3).Value = "Tasty"

        With .Range("A1:J1")
            .Interior.Color = RGB(192, 192, 192) ' Gray color
            .Font.Bold = True ' Make header text bold
        End With

        With .Range("E2:I32")
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End With

        'FBC Report Data Calculations
        '--------------------On Rent Count
        Range("D2").FormulaR1C1 = "=XLOOKUP(RC[-2], 'Customer Totals'!C[-1], 'Customer Totals'!C[1], 0)"
        Range("D2:D32").FillDown
        '--------------------Bal Amt Sum
        Range("E2").FormulaR1C1 = "=SUM(RC[1]:RC[4])"
        Range("E2:E32").FillDown
        '--------------------1-30 Days
        Range("F2").Formula = "=SUMIFS('FBC Data'!$M:$M, 'FBC Data'!$C:$C, ""="" & 'FBC Report'!$B2, 'FBC Data'!$L:$L, "">"" & TODAY() - 31)"
        Range("F2:F32").FillDown
        '--------------------31-60 Days
        Range("G2").Formula = "=SUMIFS('FBC DATA'!$M:$M, 'FBC DATA'!$C:$C, ""="" & 'FBC Report'!$B2, 'FBC DATA'!$L:$L, "">="" & (TODAY() - 60), 'FBC DATA'!$L:$L, ""<="" & (TODAY() - 31))"
        Range("G2:G32").FillDown
        '--------------------61-90 Days
        Range("H2").Formula = "=SUMIFS('FBC DATA'!$M:$M, 'FBC DATA'!$C:$C, ""="" & 'FBC Report'!$B2, 'FBC DATA'!$L:$L, "">="" & (TODAY() - 90), 'FBC DATA'!$L:$L, ""<="" & (TODAY() - 60))"
        Range("H2:H32").FillDown
        '--------------------91+ Days
        Range("I2").Formula = "=SUMIFS('FBC DATA'!$M:$M, 'FBC DATA'!$C:$C, ""="" & 'FBC Report'!$B2, 'FBC DATA'!$L:$L, ""<"" & (TODAY() - 90))"
        Range("I2:I32").FillDown

        With .Range("A:J")
            ' Change font
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Columns.AutoFilter
            .Columns.AutoFit
        End With
    End With
