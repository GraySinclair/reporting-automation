Sub AR_Setup()
'
' AR Setup Macro
'
' TO DO:
 
' PreWork --------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
   
'--------------------------------------------------------------------------
'ITEMS 1
'------------------------------------------------------------------------
    Dim it1 As Worksheet
    Set it1 = ActiveWorkbook.Sheets("ITEMS")
    With it1
        .Activate
 
        'Format Removal
'        With Columns("A:AG")
'            .ClearFormats
'        End With

        ' COLUMN NAMES
        .Cells(1, 1).Value = "Unit"
        .Cells(1, 9).Value = "Last Bill Ref"
        .Cells(1, 11).Value = "OdyNum"
        'Text to col which helps make data consistent and removes white spaces.
        Dim columnsToFormat As Variant
        Dim col As Variant
 
        columnsToFormat = Array("D", "F", "G", "H", "I", "J", "K", "W", "X", "AF")
       
        For Each col In columnsToFormat
            With .Columns(col)
                .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                    Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False
            End With
        Next col