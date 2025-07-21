Attribute VB_Name = "Module1"
Sub DeleteEmptySheets()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If WorksheetFunction.CountA(ws.Cells) = 0 Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True
    MsgBox "Semua sheet kosong telah dihapus.", vbInformation
End Sub
