Sub PurgeDatesWithinTwoDays()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim currentDate As Date
    Dim cellDate As Date

    ' Set reference to the last sheet in the workbook
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    ' Get today's date
    currentDate = Date

    ' Find the last row with data in column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Loop from the last row to the second row
    For i = lastRow To 2 Step -1
        ' Convert the date string from column E to a date
        On Error Resume Next ' In case of conversion error, skip to the next iteration
        cellDate = CDate(ws.Cells(i, "E").Value)
        On Error GoTo 0 ' Turn back on regular error handling

        ' Check if the date is within 2 days of today (either before or after)
        If Abs(DateDiff("d", currentDate, cellDate)) <= 2 Then
            ws.Rows(i).Delete
        End If
    Next i

    MsgBox "Purge complete.", vbInformation
End Sub