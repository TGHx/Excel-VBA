Sub ImportCSVAndAppendAsSheet()
    Dim csvPath As String
    Dim wb As Workbook, ws As Worksheet
    Dim newSheetName As String
    Dim lastRow As Long
    
    ' Prompt user to select CSV file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        If .Show = -1 Then ' if OK is pressed
            csvPath = .SelectedItems(1)
        Else
            MsgBox "No file selected.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Set reference to active workbook
    Set wb = ThisWorkbook
    
    ' Generate sheet name in YYYYMMDD format
    newSheetName = Format(Now, "YYYYMMDD")
    
    ' Check if sheet already exists
    On Error Resume Next
    Set ws = wb.Sheets(newSheetName)
    On Error GoTo 0
    
    ' If sheet exists, prompt to overwrite
    If Not ws Is Nothing Then
        If MsgBox("Sheet " & newSheetName & " already exists. Overwrite?", vbYesNo) = vbYes Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        Else
            MsgBox "Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Add new sheet with specified name
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = newSheetName
    
    ' Import CSV into the new sheet
    With ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With
    
    MsgBox "CSV imported successfully as " & newSheetName, vbInformation
End Sub




