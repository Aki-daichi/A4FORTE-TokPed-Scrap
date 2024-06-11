Attribute VB_Name = "Module1"
Sub AddReviewToURLs()
    Dim lastRow As Long
    Dim i As Long
    
    ' Menentukan baris terakhir yang berisi data di kolom A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop melalui setiap baris yang berisi URL di kolom A
    For i = 1 To lastRow
        ' Menambahkan "/review" pada URL di kolom A dan menyimpannya di kolom B
        Cells(i, 2).Value = Cells(i, 1).Value & "/review"
    Next i
End Sub
