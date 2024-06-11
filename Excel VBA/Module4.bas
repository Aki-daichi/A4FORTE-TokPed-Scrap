Attribute VB_Name = "Module4"

Sub AmbilProduk()
    Dim IE As Object
    Dim HTMLDoc As Object
    Dim URL As String
    Dim productName As String
    Dim productPrice As String
    Dim productSales As String
    Dim lastRow As Long
    Dim i As Long
    Dim productIndex As Long
    Dim hasSauceProduct As Boolean
    
    ' Menentukan baris terakhir di kolom A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Membuat instance Internet Explorer
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False ' Tidak menampilkan jendela browser
    
    ' Loop melalui setiap baris dari 2 hingga lastRow
    For i = 2 To lastRow
        ' Mengambil URL dari kolom A
        URL = Cells(i, 1).Value
        
        ' Membuka URL
        IE.navigate URL
        
        ' Menunggu hingga halaman selesai dimuat
        Do While IE.Busy Or IE.readyState <> 4
            DoEvents
        Loop
        
        ' Mengambil dokumen HTML
        Set HTMLDoc = IE.document
        
        ' Reset flag pencarian produk saus
        hasSauceProduct = False
        
        ' Mengambil data produk
        On Error Resume Next
        ' Cari produk dengan kata kunci "sambal", "sambel", "saos", atau "saus"
        For productIndex = 0 To 4
            productName = HTMLDoc.querySelectorAll("[data-testid='linkProductName']").Item(productIndex).innerText
            If InStr(1, productName, "sambal", vbTextCompare) > 0 Or _
               InStr(1, productName, "sambel", vbTextCompare) > 0 Or _
               InStr(1, productName, "saos", vbTextCompare) > 0 Or _
               InStr(1, productName, "saus", vbTextCompare) > 0 Then
                productPrice = HTMLDoc.querySelectorAll("[data-testid='linkProductPrice']").Item(productIndex).innerText
                If HTMLDoc.querySelectorAll(".prd_label-integrity.css-1sgek4h").Length > productIndex Then
                    productSales = HTMLDoc.querySelectorAll(".prd_label-integrity.css-1sgek4h").Item(productIndex).innerText
                Else
                    productSales = "0" ' Atur penjualan menjadi 0 jika tidak ditemukan
                End If
                hasSauceProduct = True
                Exit For
            End If
        Next productIndex
        
        ' Jika tidak ada produk saus, ambil produk pertama
        If Not hasSauceProduct Then
            productName = HTMLDoc.querySelectorAll("[data-testid='linkProductName']").Item(0).innerText
            productPrice = HTMLDoc.querySelectorAll("[data-testid='linkProductPrice']").Item(0).innerText
            If HTMLDoc.querySelectorAll(".prd_label-integrity.css-1sgek4h").Length > 0 Then
                productSales = HTMLDoc.querySelectorAll(".prd_label-integrity.css-1sgek4h").Item(0).innerText
            Else
                productSales = "0" ' Atur penjualan menjadi 0 jika tidak ditemukan
            End If
        End If
        On Error GoTo 0
        
        ' Menyimpan data ke kolom J, K, L (Kolom 10, 11, 12)
        Cells(i, 10).Value = productName
        Cells(i, 11).Value = productPrice
        Cells(i, 12).Value = productSales
        
        ' Produk tambahan
        ' Menyimpan data produk 2 ke kolom M, N, O (Kolom 13, 14, 15)
        For productIndex = 1 To 4
            If HTMLDoc.querySelectorAll("[data-testid='linkProductName']").Length > productIndex Then
                productName = HTMLDoc.querySelectorAll("[data-testid='linkProductName']").Item(productIndex).innerText
                productPrice = HTMLDoc.querySelectorAll("[data-testid='linkProductPrice']").Item(productIndex).innerText
                If HTMLDoc.querySelectorAll(".prd_label-integrity.css-1sgek4h").Length > productIndex Then
                    productSales = HTMLDoc.querySelectorAll(".prd_label-integrity.css-1sgek4h").Item(productIndex).innerText
                Else
                    productSales = "0" ' Atur penjualan menjadi 0 jika tidak ditemukan
                End If
                
                Cells(i, (productIndex * 3) + 10).Value = productName
                Cells(i, (productIndex * 3) + 11).Value = productPrice
                Cells(i, (productIndex * 3) + 12).Value = productSales
            Else
                Exit For ' Keluar dari loop jika tidak ada cukup data
            End If
        Next productIndex
        
    Next i
    
    ' Menutup instance Internet Explorer
    IE.Quit
    Set IE = Nothing
    Set HTMLDoc = Nothing
End Sub



