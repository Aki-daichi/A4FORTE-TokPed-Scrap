Attribute VB_Name = "Module3"

Sub CheckAndClearDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Ganti "Sheet1" dengan nama lembar yang sesuai
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Loop through each row
    For i = 2 To lastRow
        ' Check if value in column C equals value in column D
        If ws.Cells(i, "C").Value = ws.Cells(i, "D").Value Then
            ' Clear value in column D
            ws.Cells(i, "D").Value = ""
        End If
    Next i

    MsgBox "Duplicates checked and cleared!", vbInformation
End Sub

