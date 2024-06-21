Sub InsertAndPopulateCMDBID()
    Dim lastSheet As Worksheet
    Dim reconcileSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim serialNumber As String
    Dim found As Range

    ' Set references to the sheets
    Set lastSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set reconcileSheet = ThisWorkbook.Sheets("Reconciled List")

    ' Insert new column G for "CMDB ID"
    lastSheet.Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    lastSheet.Cells(1, "G").Value = "CMDB ID"

    ' Find the last row with data in the lastSheet
    lastRow = lastSheet.Cells(lastSheet.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in the lastSheet starting from row 2
    For i = 2 To lastRow
        serialNumber = lastSheet.Cells(i, "A").Value ' Assuming serial numbers are in column A

        ' Search for the serial number in column B of the Reconciled List
        Set found = reconcileSheet.Columns("B:B").Find(What:=serialNumber, LookIn:=xlValues, LookAt:=xlWhole)

        ' If found, copy the corresponding CMDB ID from column A of Reconciled List to column G of the lastSheet
        If Not found Is Nothing Then
            lastSheet.Cells(i, "G").Value = reconcileSheet.Cells(found.Row, "A").Value
        End If
    Next i

    MsgBox "CMDB ID column populated.", vbInformation
End Sub