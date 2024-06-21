Sub RemoveUnwantedColumns()
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim col As Range
    Dim i As Integer
    Dim keepColumns As Object
    Set keepColumns = CreateObject("Scripting.Dictionary")
    
    ' Initialize the dictionary with the headers to keep
    keepColumns.Add "Serial", 1
    keepColumns.Add "Last Check In", 1
    keepColumns.Add "CMDB ID", 1
    
    ' Set reference to the last sheet in the workbook
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' Assuming the headers are in the first row
    Set headerRow = ws.Rows(1)
    
    ' Loop through each column in reverse order
    For i = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
        If Not keepColumns.exists(ws.Cells(1, i).Value) Then
            ws.Columns(i).Delete
        End If
    Next i
    
    MsgBox "Columns cleaned.", vbInformation
End Sub