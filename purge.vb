Sub PurgeUnreconciledRows()
    Dim wsCSV As Worksheet, wsReconciled As Worksheet
    Dim lastRowCSV As Long, lastRowReconciled As Long, r As Long
    Dim serialNumber As String
    Dim found As Range
    
    ' Set references to the worksheets
    Set wsCSV = ThisWorkbook.Sheets("NameOfYourCSVSheet") ' Change to your actual CSV sheet name
    Set wsReconciled = ThisWorkbook.Sheets("Reconciled List")
    
    ' Find the last row of data in both sheets
    lastRowCSV = wsCSV.Cells(wsCSV.Rows.Count, "A").End(xlUp).Row ' Assuming serial numbers are in column A
    lastRowReconciled = wsReconciled.Cells(wsReconciled.Rows.Count, "A").End(xlUp).Row
    
    ' Loop from the last row to the second row in the CSV data sheet
    For r = lastRowCSV To 2 Step -1
        serialNumber = wsCSV.Cells(r, 1).Value ' Assuming serial numbers are in column A
        
        ' Attempt to find the serial number in the "Reconciled List"
        Set found = wsReconciled.Range("A1:A" & lastRowReconciled).Find(serialNumber, LookIn:=xlValues)
        
        ' If not found, delete the row from the CSV data sheet
        If found Is Nothing Then
            wsCSV.Rows(r).Delete
        End If
    Next r
    
    MsgBox "Unreconciled rows purged.", vbInformation
End Sub