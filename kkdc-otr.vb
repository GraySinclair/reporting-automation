Sub KKDC_OTR()
'
' KKDC_OTR Macro
' Requires Access to "Krispy Kreme Acc Info.xlsx" (Access can be granted by Gray Sinclair - E66CVG. To update files, email Taylor Furgalack from Krispy Kreme)


    Range("L2:L99").Select
    Selection.ClearContents
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-2],'[Krispy Kreme Acc Info.xlsx]Krispy Kreme Info'!C1,'[Krispy Kreme Acc Info.xlsx]Krispy Kreme Info'!C2,0)"
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=IF(K2<>L2,1,0)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Application.CutCopyMode = False
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$D$1:$D$99").AutoFilter Field:=4, Criteria1:="17192263"
End Sub