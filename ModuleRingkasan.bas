Attribute VB_Name = "ModuleRingkasan"
'@IgnoreModule AssignmentNotUsed
'@Folder "resume"

Sub TampilkanLogRingkas()
    Dim wsLog As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim startRow As Long
    Dim currentActionID As String
    Dim processedActions As Collection
    Dim actionCount As Long
    Dim sheetCount As Long
    Dim materialReplaced As String
    Dim materialNew As String
    
    Set wsLog = ThisWorkbook.Sheets("HISTORY_UNDO")
    Set wsTarget = ThisWorkbook.Sheets("CALCULATE")
    Set processedActions = New Collection
    
    startRow = 26                                ' Mulai dari baris 26 di CALCULATE
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    
    ' Header yang dipadatkan
    With wsTarget
        .Range("B" & startRow).Value = "No"
        .Range("C" & startRow).Value = "Tanggal"
        .Range("D" & startRow).Value = "Action ID"
        .Range("E" & startRow).Value = "Material Lama"
        .Range("F" & startRow).Value = "Material Baru"
        .Range("G" & startRow).Value = "Sheet"
        .Range("H" & startRow).Value = "Perubahan"
        .Range("I" & startRow).Value = "Jenis"
        .Range("J" & startRow).Value = "Nilai Lama"
        .Range("K" & startRow).Value = "Nilai Baru"
        
        ' Bersihkan area log data sebelumnya (hanya sampai baris 31)
        .Range("B" & startRow + 1 & ":K31").ClearContents
       ' .Range("B" & startRow & ":K" & startRow).Font.Bold = True
    End With
    
    ' Cek jika belum ada history (artinya hanya header di baris 1)
    If lastRow < 2 Then
        With wsTarget
            
            .Range("B" & startRow + 1).Value = "Tidak ada history perubahan material"
            
        End With
        Exit Sub
    End If
    
    ' Tampilkan maksimal 5 entry terakhir (sampai baris 31)
    j = 1
    For i = lastRow To 2 Step -1
        If j > 5 Or (startRow + j) > 31 Then Exit For
        
        With wsTarget
            .Cells(startRow + j, "B").Value = j  ' No
            .Cells(startRow + j, "C").Value = Format(wsLog.Cells(i, "A").Value, "dd/mm hh:mm") ' Tanggal singkat
            .Cells(startRow + j, "D").Value = wsLog.Cells(i, "H").Value ' Action ID
            
            ' Tentukan material berdasarkan jenis aksi
            If wsLog.Cells(i, "J").Value = "REPLACE" Then
                .Cells(startRow + j, "E").Value = wsLog.Cells(i, "E").Value ' Material yang diganti
                .Cells(startRow + j, "F").Value = wsLog.Cells(i, "I").Value ' Material baru
            Else
                .Cells(startRow + j, "E").Value = "-"
                .Cells(startRow + j, "F").Value = wsLog.Cells(i, "E").Value ' Material yang ditambah
            End If
            
            .Cells(startRow + j, "G").Value = wsLog.Cells(i, "B").Value ' Sheet
            .Cells(startRow + j, "H").Value = "Row " & wsLog.Cells(i, "C").Value & " Col " & wsLog.Cells(i, "D").Value ' Lokasi perubahan
            
            ' Singkatan jenis aksi
            Select Case wsLog.Cells(i, "J").Value
            Case "REPLACE"
                .Cells(startRow + j, "I").Value = "GANTI"
            Case "ADD_EXISTING"
                .Cells(startRow + j, "I").Value = "TAMBAH"
            Case "INSERT_ROW"
                .Cells(startRow + j, "I").Value = "BARIS"
            Case "ADD_NEW"
                .Cells(startRow + j, "I").Value = "BARU"
            Case Else
                .Cells(startRow + j, "I").Value = "LAIN"
            End Select
            
            ' Nilai lama dan baru
            If IsNumeric(wsLog.Cells(i, "F").Value) Then
                .Cells(startRow + j, "J").Value = Format(wsLog.Cells(i, "F").Value, "0.00")
            Else
                .Cells(startRow + j, "J").Value = wsLog.Cells(i, "F").Value
            End If
            
            If IsNumeric(wsLog.Cells(i, "G").Value) Then
                .Cells(startRow + j, "K").Value = Format(wsLog.Cells(i, "G").Value, "0.00")
            Else
                .Cells(startRow + j, "K").Value = wsLog.Cells(i, "G").Value
            End If
        End With
        j = j + 1
    Next i
    
    ' Format kolom untuk tampilan yang lebih rapi
    With wsTarget.Range("C" & startRow + 1 & ":C31")
        .NumberFormat = "dd/mm hh:mm"
    End With
    
    With wsTarget.Range("J" & startRow + 1 & ":K31")
        .NumberFormat = "0.00"
    End With
End Sub

