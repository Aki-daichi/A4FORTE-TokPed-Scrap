Attribute VB_Name = "Module2"
Sub AmbilDataTokopedia()
    Dim IE As Object
    Dim html As Object
    Dim ele As Object
    Dim rating As String
    Dim ulasan As String
    Dim totalRating As String
    Dim totalUlasan As String
    Dim pesananProses As String
    Dim jamOperasi As String
    Dim shopName As String
    Dim shopLocation As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Buat instance baru Internet Explorer
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False ' Atur ke True jika Anda ingin melihat browser beroperasi

    ' Dapatkan worksheet aktif
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Ganti "Sheet1" dengan nama sheet Anda

    ' Dapatkan baris terakhir di kolom B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Loop melalui setiap tautan di kolom B
    For i = 2 To lastRow ' Mulai dari baris kedua jika ada header
        ' Buka tautan
        IE.navigate ws.Cells(i, 2).Value
        Do While IE.Busy Or IE.readyState <> 4
            DoEvents
        Loop

        ' Dapatkan HTML dokumen
        Set html = IE.document

        ' Cari elemen yang berisi jumlah ulasan dan rating
        On Error Resume Next
        Set ele = html.querySelector("p.css-1bhobcm-unf-heading.e1qvo2ff8")
        If Not ele Is Nothing Then
            ulasan = ele.innerText ' Mengambil teks jumlah ulasan
        Else
            ulasan = ""
        End If
        On Error GoTo 0
        
        ' Extract the review count and rating count from the text
        If InStr(ulasan, "•") > 0 Then
            totalUlasan = Trim(Split(Split(ulasan, "•")(1), " ")(1))
            totalRating = Trim(Split(ulasan, " ")(0))
        Else
            totalUlasan = "N/A"
            totalRating = "N/A"
        End If
        
        ' Remove any thousand separators (comma or dot)
        totalUlasan = Replace(totalUlasan, ".", "")
        totalUlasan = Replace(totalUlasan, ",", "")
        
        totalRating = Replace(totalRating, ".", "")
        totalRating = Replace(totalRating, ",", "")

        ' Cari elemen yang berisi rata-rata rating
        On Error Resume Next
        Set ele = html.querySelector(".css-124sv7s .score")
        If Not ele Is Nothing Then
            rating = ele.innerText ' Mengambil teks rata-rata rating
        Else
            rating = ""
        End If
        On Error GoTo 0

        ' Cari elemen yang berisi nama toko dan lokasi toko
        On Error Resume Next
        Set ele = html.querySelector("h1[data-testid='shopNameHeader']")
        If Not ele Is Nothing Then
            shopName = ele.innerText ' Mengambil teks nama toko
        Else
            shopName = ""
        End If
        Set ele = html.querySelector("span[data-testid='shopLocationHeader']")
        If Not ele Is Nothing Then
            shopLocation = ele.innerText ' Mengambil teks lokasi toko
        Else
            shopLocation = ""
        End If
        On Error GoTo 0

        ' Cari elemen yang berisi pesanan diproses
        On Error Resume Next
        Set ele = html.querySelector(".css-p8je3v.e1wfhb0y0 + div")
        If Not ele Is Nothing Then
            pesananProses = ele.innerText ' Mengambil teks pesanan diproses
            pesananProses = Trim(Replace(pesananProses, "Pesanan diproses", ""))
            pesananProses = Replace(pesananProses, vbCrLf, "")
        Else
            pesananProses = "N/A"
        End If
        On Error GoTo 0

        ' Cari elemen yang berisi jam operasi
        On Error Resume Next
        Set ele = html.querySelector("[data-testid='shopOperationalHourHeader'] > .css-6x4cyu > p")
        If Not ele Is Nothing Then
            jamOperasi = ele.innerText ' Mengambil teks jam operasi
            jamOperasi = Replace(jamOperasi, vbCrLf, "")
        Else
            jamOperasi = "N/A"
        End If
        On Error GoTo 0

        ' Masukkan data ke kolom yang sesuai
        ws.Cells(i, 3).Value = shopName
        ws.Cells(i, 4).Value = shopLocation
        ws.Cells(i, 5).Value = rating
        ws.Cells(i, 6).Value = totalRating
        ws.Cells(i, 7).Value = totalUlasan
        ws.Cells(i, 8).Value = pesananProses
        ws.Cells(i, 9).Value = jamOperasi

        ' Pause for a short time to avoid overloading the server
        Application.Wait (Now + TimeValue("0:00:02"))
    Next i

    ' Tutup Internet Explorer
    IE.Quit
    Set IE = Nothing

    MsgBox "Shop info, review counts, ratings, order processing, and operational hours have been updated!", vbInformation
End Sub

