Sub PurgeUnmatchedSerialNumbers()
    Dim purgeSheet As Worksheet
    Dim reconcileSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim serialNumber As String
    Dim found As Range

    ' Set references to the sheets
    Set purgeSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set reconcileSheet = ThisWorkbook.Sheets("Reconciled List")

    ' Find the last row with data in the purgeSheet
    lastRow = purgeSheet.Cells(purgeSheet.Rows.Count, "A").End(xlUp).Row

    ' Loop from the last row to the second row
    For i = lastRow To 2 Step -1
        serialNumber = purgeSheet.Cells(i, 1).Value ' Assuming serial numbers are in column A

        ' Check if the serial number exists in the Reconciled List
        Set found = reconcileSheet.Columns(1).Find(What:=serialNumber, LookIn:=xlValues, LookAt:=xlWhole)

        ' If not found, delete the row
        If found Is Nothing Then
            purgeSheet.Rows(i).Delete
        End If
    Next i

    MsgBox "Purge complete.", vbInformation
End Sub