Sub TollConsolidation()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowDest As Long
    Dim i As Integer
   
 
    ' Built by Gray Sinclair - E66CVG
    ' Loop through each sheet starting from sheet 2
    For i = 2 To Sheets.Count
        Set wsSource = Sheets(i)
        Set wsDest = Sheets(1)
       
        ' Delete the first 9 rows in the source sheet
        wsSource.Rows("1:9").Delete