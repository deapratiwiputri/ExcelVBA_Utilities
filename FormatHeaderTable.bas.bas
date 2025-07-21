Attribute VB_Name = "Module1"
Sub FormatHeaderOnlyCellsWithData()
    Dim ws As Worksheet ' Mendeklarasikan variabel ws sebagai objek Worksheet
    Dim firstDataRow As Long ' Mendeklarasikan variabel firstDataRow untuk menyimpan nomor baris pertama yang berisi data
    Dim rng As Range ' Mendeklarasikan variabel rng sebagai objek Range (akan menyimpan sel-sel header)
    Dim cell As Range ' Mendeklarasikan variabel cell untuk iterasi sel

    For Each ws In ThisWorkbook.Worksheets ' Melakukan iterasi melalui setiap sheet di workbook yang aktif
        With ws ' Menggunakan blok With untuk merujuk ke sheet yang sedang diproses tanpa harus menulis ws. berulang kali
            firstDataRow = 0 ' Menginisialisasi firstDataRow ke 0 untuk setiap sheet

            ' Cari baris pertama yang memiliki isi (dari atas)
            ' Memeriksa 1000 baris pertama di kolom A. Ini bisa dioptimalkan.
            For Each cell In .Range("A1:A1000")
                ' Memeriksa apakah baris tersebut memiliki data dengan menghitung sel tidak kosong di seluruh baris
                If Application.WorksheetFunction.CountA(.Rows(cell.Row)) > 0 Then
                    firstDataRow = cell.Row ' Jika ada data, simpan nomor barisnya
                    Exit For ' Keluar dari loop karena baris pertama sudah ditemukan
                End If
            Next cell

            ' Kalau ditemukan baris yang berisi data
            If firstDataRow > 0 Then
                Set rng = Nothing ' Mengatur rng ke Nothing untuk memastikan bersih sebelum digunakan
                ' Ambil hanya cell yang berisi data pada baris tersebut
                For Each cell In .Rows(firstDataRow).Cells ' Iterasi melalui setiap sel di baris pertama yang ditemukan
                    If cell.Value <> "" Then ' Jika sel tidak kosong
                        If rng Is Nothing Then ' Jika rng belum diatur (ini sel data pertama yang ditemukan di baris itu)
                            Set rng = cell ' Atur rng ke sel ini
                        Else
                            Set rng = Union(rng, cell) ' Jika rng sudah diatur, tambahkan sel ini ke rentang rng
                        End If
                    End If
                Next cell

                ' Format hanya cell yang berisi data
                If Not rng Is Nothing Then ' Jika rng (yaitu, sel-sel header yang berisi data) tidak kosong
                    With rng
                        .Interior.Color = RGB(197, 217, 241) ' Mengatur warna latar belakang (biru muda)
                        .Font.Color = RGB(0, 0, 0)         ' Mengatur warna font (hitam)
                        .Font.Bold = True                  ' Mengatur font menjadi tebal
                    End With
                End If
            End If
        End With
    Next ws
End Sub
