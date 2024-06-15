Sub TollConsolidation()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowDest As Long
    Dim i As Integer
   
    ' TODO: 
    ' prevent enterprise logo from being copied over during loop
    ' use filepath to consolidate all files in folder instead of dropping all files into 1 workbook requirement
    '
    '

    ' Built by Gray Sinclair - E66CVG
    ' Loop through each sheet starting from sheet 2
    For i = 2 To Sheets.Count
        Set wsSource = Sheets(i)
        Set wsDest = Sheets(1)
       
        ' Delete the first 9 rows in the source sheet
        wsSource.Rows("1:9").Delete
       
        ' Find the last row in the destination sheet
        lastRowDest = wsDest.Cells(wsDest.Rows.Count, "E").End(xlUp).Row
       
        ' Copy data from columns A to T in the source sheet to the bottom of the destination sheet
        wsSource.Range("A1:T" & wsSource.Cells(wsSource.Rows.Count, "E").End(xlUp).Row).Copy _
            Destination:=wsDest.Range("A" & lastRowDest + 1)
    Next i
    'Format Section
    Sheets("Sheet1").Select
    Sheets("Sheet1").Move
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Sheet1").Range("A1").Select
End Sub