Sub CorpBillCitationAdminFees()
    'worksheets
    Dim tssfee As Worksheet
    Dim tsstotal As Worksheet
    Dim access As Worksheet
    Dim historic As Worksheet
   
    
    ' PreWork ----------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Turn off automatic calculation
    Application.EnableEvents = False ' Disable events temporarily
 
    Set tssfee = ActiveWorkbook.Sheets("TSS Fee")
    Set tsstotal = ActiveWorkbook.Sheets("TSS Fee_Total")
    Set access = ActiveWorkbook.Sheets("Master Access File")
    Set historic = ActiveWorkbook.Sheets("Historic File")
   
 
    RemoveBlankColumns tssfee, 23
    RemoveBlankColumns tsstotal, 19
   
    ' deleting blank columns if they exist in tssfees and tsstotal---------------------------------
    'Dim col As Long
    ' Loop through columns from X to A (backward to avoid shifting)
    ' deleting blank columns in tssfee
'    For col = 23 To 1 Step -1
'        If WorksheetFunction.CountA(tssfee.Columns(col)) = 0 Then
'            tssfee.Columns(col).Delete
'        End If
'    Next col
   
    ' deleting blank columns in tsstotal
    For col = 19 To 1 Step -1
        If WorksheetFunction.CountA(tsstotal.Columns(col)) = 0 Then
            tsstotal.Columns(col).Delete
        End If
    Next col

   
'    ' deleting blank columns in access
'    For col = 20 To 1 Step -1
'        If WorksheetFunction.CountA(access.Columns(col)) = 0 Then
'            access.Columns(col).Delete
'        End If
'    Next col
'    ' deleting blank columns in historic
'    For col = 26 To 1 Step -1
'        If WorksheetFunction.CountA(historic.Columns(col)) = 0 Then
'            historic.Columns(col).Delete
'        End If
'    Next col
'
 
    'for deleting extra space------------------------------------------------------
    Dim lastrowintssfee As Long
    Dim lastrowintsstotal As Long
    Dim lastrowinhistoric As Long
    Dim lastrowinaccess As Long
   
    lastrowintssfee = tssfee.Cells(tssfee.Rows.Count, "A").End(xlUp).Row
    lastrowintsstotal = tsstotal.Cells(tsstotal.Rows.Count, "A").End(xlUp).Row
    lastrowinhistoric = historic.Cells(historic.Rows.Count, "A").End(xlUp).Row
    lastrowinaccess = access.Cells(access.Rows.Count, "A").End(xlUp).Row
   
    ' Delete rows after the last row with data
    If lastrowintssfee < tssfee.Rows.Count Then
        tssfee.Rows(lastrowintssfee + 1 & ":" & tssfee.Rows.Count).Delete
    End If
   
    If lastrowintsstotal < tsstotal.Rows.Count Then
        tsstotal.Rows(lastrowintsstotal + 1 & ":" & tsstotal.Rows.Count).Delete
    End If
   
    If lastrowinhistoric < historic.Rows.Count Then
        historic.Rows(lastrowinhistoric + 1 & ":" & historic.Rows.Count).Delete
    End If
   
    If lastrowinaccess < access.Rows.Count Then
        access.Rows(lastrowinaccess + 1 & ":" & access.Rows.Count).Delete
    End If
   

    Dim header As Range
    ' Loop through columns A to W (1 to 23)
    With tssfee
        For col = 1 To 23
            ' Check if the header in row 1 matches any of the specified values
            Set header = tssfee.Cells(1, col)
       
            If header.Value = "BillingRefNum" Or header.Value = "Brand" Or header.Value = "CheckOutLocation" Or header.Value = "Lic State" Or header.Value = "Invoice Ending" Then
                tssfee.Columns(col).Delete
                col = col - 1 ' Adjust column index after deletion
            End If
        Next col
    End With
   

    With tsstotal
        For col = 1 To 23
            ' Check if the header in row 1 matches any of the specified values
            Set header = tsstotal.Cells(1, col)
       
            If header.Value = "BillingRefNum" Or header.Value = "Brand" Or header.Value = "CheckOutLocation" Or header.Value = "Lic State" Or header.Value = "Usage Days" Or header.Value = "Invoice Ending" Then
                tsstotal.Columns(col).Delete
                col = col - 1 ' Adjust column index after deletion
            End If
        Next col
    End With
 
    Dim tblRange As Range
    Dim tbl As ListObject
    Dim lastcolintssfee As Long
    Dim lastcolintsstotal As Long
    Dim lastcolinaccess As Long
    Dim lastcolinhistoric As Long
   
    Dim tolldatetime As Integer
    Dim formula As String
    Dim datarange As Range
   
    Dim ranumcol As ListColumn
    Dim firstcol As ListColumn
    'for adding columns
    Dim corpIDColumn As ListColumn
    Dim corpIDIndex As Integer
   
    ' Find the last used columns
    lastcolintssfee = tssfee.Cells(1, tssfee.Columns.Count).End(xlToLeft).Column ' Last column in row 1
    lastcolintsstotal = tsstotal.Cells(1, tsstotal.Columns.Count).End(xlToLeft).Column
    lastcolinaccess = access.Cells(1, access.Columns.Count).End(xlToLeft).Column
    lastcolinhistoric = historic.Cells(1, historic.Columns.Count).End(xlToLeft).Column
   
    'historic table-----------------------------------------------
    ' Define the dynamic range (adjust based on lastRow and lastCol)
    Set tblRange = historic.Range(historic.Cells(1, 1), historic.Cells(lastrowinhistoric, lastcolinhistoric))
   
    ' Add a table to the range
    Set tbl = historic.ListObjects.Add(SourceType:=xlSrcRange, Source:=tblRange, _
                                  XlListObjectHasHeaders:=xlYes)
   
    ' Optional: Name the table
    tbl.Name = "historictable"
 
    ' Optional: Format the table (apply a style)
    tbl.TableStyle = "" ' You can choose a different table style here
 
    'access table-----------------------------------------------
    ' Define the dynamic range (adjust based on lastRow and lastCol)
    Set tblRange = access.Range(access.Cells(1, 1), access.Cells(lastrowinaccess, lastcolinaccess))
   
    ' Add a table to the range
    Set tbl = access.ListObjects.Add(SourceType:=xlSrcRange, Source:=tblRange, _
                                  XlListObjectHasHeaders:=xlYes)
   
    ' Optional: Name the table
    tbl.Name = "accesstable"
     
    ' Optional: Format the table (apply a style)
    tbl.TableStyle = "" ' You can choose a different table style here
   
    'tssfee table-----------------------------------------------
    ' Define the dynamic range (adjust based on lastRow and lastCol)
    Set tblRange = tssfee.Range(tssfee.Cells(1, 1), tssfee.Cells(lastrowintssfee, lastcolintssfee))
   
    ' Add a table to the range
    Set tbl = tssfee.ListObjects.Add(SourceType:=xlSrcRange, Source:=tblRange, _
                                  XlListObjectHasHeaders:=xlYes)
   
    ' Optional: Name the table
    tbl.Name = "tssfeetable"

    ' Optional: Format the table (apply a style)
    tbl.TableStyle = "" ' You can choose a different table style here
    'tbl.Range.ClearFormats
   
    ' Find the CorpID column
    On Error Resume Next
    Set corpIDColumn = tbl.ListColumns("CorpID")
    On Error GoTo 0
   
    If Not corpIDColumn Is Nothing Then
        ' Get the index of the CorpID column
        corpIDIndex = corpIDColumn.Index
       
        ' Insert new columns after the CorpID column
        tbl.ListColumns.Add (corpIDIndex + 1) ' Insert first new column after CorpID
        tbl.ListColumns.Add (corpIDIndex + 2) ' Insert second new column after CorpID
        tbl.ListColumns.Add (corpIDIndex + 13) ' Insert third new column after CorpID
        ' Insert a new column after the 18th column (or where you want it)
        tolldatetime = corpIDIndex + 18 ' 18 columns after CorpID
        ' Insert a new column
        tbl.ListColumns.Add (tolldatetime)
       
        ' Set the headers for the new columns
        tbl.HeaderRowRange.Cells(1, corpIDIndex + 1).Value = "BA"
        tbl.HeaderRowRange.Cells(1, corpIDIndex + 2).Value = "Frequency"
        tbl.HeaderRowRange.Cells(1, corpIDIndex + 13).Value = "Unit #"
        ' Set the header for the new column
        tbl.HeaderRowRange.Cells(1, tolldatetime).Value = "Toll Date/Time"
       
 
        formula = "=TEXT([@[Toll Date]],""mm/dd/yyyy"")&TEXT([@[ISSUE TIME]],"" hh:mm:ss"")"
        Set datarange = tbl.ListColumns(tolldatetime).DataBodyRange
        datarange.formula = formula
       

        ' Copy the calculated values
        datarange.Copy
 
        ' Paste as values (overwriting the formulas with their results)
        datarange.PasteSpecial Paste:=xlPasteValues
        ' Clear the clipboard to avoid the "marching ants" effect
        Application.CutCopyMode = False
       
        If tolldatetime > 2 Then
            tbl.ListColumns(tolldatetime - 1).Delete
            tbl.ListColumns(tolldatetime - 2).Delete
        End If
       
        ' Find the column "RA#" in the table
        On Error Resume Next
        Set ranumcol = tbl.ListColumns("RA#")
        On Error GoTo 0
       
        ' Check if the "RA#" column exists
        If Not ranumcol Is Nothing Then
            ' Get the first column in the table
            Set firstcol = tbl.ListColumns(1)
       
            ' Move the "RA#" column to be the first column
            ranumcol.Range.Cut
            firstcol.Range.Insert Shift:=xlToRight
        End If
    End If
 
End Sub