Attribute VB_Name = "Module1"
Sub AutoNumberAllSheets()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim colNo As Long
    Dim lastRow As Long
    Dim i As Long
    Dim headerFound As Boolean

    ' Menonaktifkan pembaruan layar untuk mempercepat eksekusi makro
    Application.ScreenUpdating = False

    ' Loop melalui setiap sheet dalam workbook yang aktif
    For Each ws In ThisWorkbook.Worksheets
        headerFound = False ' Reset flag untuk setiap sheet
        colNo = 0           ' Reset kolom nomor untuk setiap sheet

        ' Mencari baris header pertama yang berisi data (maksimal 100 baris pertama)
        ' Ini asumsi bahwa header Anda ada di 100 baris teratas
        For headerRow = 1 To 100
            ' Mencari "No", "no", atau "No Urut" di baris tersebut
            ' xlWhole: mencari kecocokan persis
            ' xlValues: mencari nilai sel, bukan rumus
            ' xlByRows: mencari berdasarkan baris
            ' xlNext: mencari selanjutnya (dari kiri ke kanan)
            On Error Resume Next ' Menghindari error jika Find tidak menemukan apa-apa
            Dim foundCell As Range
            Set foundCell = ws.Rows(headerRow).Find(What:="No", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False)
            If foundCell Is Nothing Then
                Set foundCell = ws.Rows(headerRow).Find(What:="no", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False)
            End If
            If foundCell Is Nothing Then
                Set foundCell = ws.Rows(headerRow).Find(What:="No Urut", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False)
            End If
            On Error GoTo 0 ' Mengembalikan penanganan error normal

            If Not foundCell Is Nothing Then
                colNo = foundCell.Column ' Dapatkan nomor kolom dari header yang ditemukan
                headerFound = True       ' Set flag bahwa header ditemukan
                Exit For                 ' Keluar dari loop pencarian header
            End If
        Next headerRow

        ' Jika header "No" atau "No Urut" ditemukan di sheet ini
        If headerFound Then
            With ws
                ' Mencari baris terakhir yang berisi data di kolom yang ditemukan (colNo)
                ' xlUp akan mencari dari bawah ke atas sampai menemukan sel yang berisi data
                lastRow = .Cells(.Rows.Count, colNo).End(xlUp).Row

                ' Memastikan ada data di bawah header (lebih dari satu baris)
                ' Dan pastikan baris terakhir tidak sama dengan baris header
                If lastRow > headerRow Then
                    ' Memulai penomoran dari baris setelah header
                    For i = headerRow + 1 To lastRow
                        .Cells(i, colNo).Value = i - headerRow ' Isi sel dengan nomor urut
                    Next i
                ElseIf lastRow = headerRow And .Cells(headerRow, colNo).Value <> "" Then
                    ' Kondisi khusus: jika hanya ada header dan tidak ada data lain di bawahnya,
                    ' atau hanya ada satu baris data (header + 1 data) tapi lastRow masih sama dengan headerRow.
                    ' Kita bisa menambahkan logika di sini jika perlu menomori hanya 1 data di bawah header.
                    ' Untuk kasus ini, kita tidak akan menomori jika hanya ada header.
                    ' Jika Anda ingin menomori 1 data, Anda bisa menambahkan:
                    ' If .Cells(headerRow + 1, colNo).Value <> "" Then .Cells(headerRow + 1, colNo).Value = 1
                End If
            End With
            Debug.Print "Sheet: " & ws.Name & " - Penomoran selesai di kolom " & colNo
        Else
            Debug.Print "Sheet: " & ws.Name & " - Header 'No' atau 'No Urut' tidak ditemukan. Dilewati."
        End If
    Next ws

    Application.ScreenUpdating = True ' Mengaktifkan kembali pembaruan layar
    MsgBox "Penomoran di semua sheet telah selesai.", vbInformation
End Sub
