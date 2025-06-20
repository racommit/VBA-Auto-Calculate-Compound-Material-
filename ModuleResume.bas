Attribute VB_Name = "ModuleResume"
'@IgnoreModule AssignmentNotUsed
'@Folder "resume"
Sub UpdateResumeSheet()
    Dim wsResume As Worksheet
    Dim ws As Worksheet
    Dim resumeIndex As Long
    Dim i As Long
    Dim colOffset As Long
    Dim insertCol As Long
    Dim sheetList As Collection
    Dim wsName As Variant
    Dim targetCol As Long
    Const afterCols As Long = 3 ' "sisa nwt" + 2 kolom setelahnya

    Set sheetList = New Collection
    Set wsResume = ThisWorkbook.Sheets("RESUME")
    
    ' Temukan index sheet Resume
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i

    ' Ambil semua sheet sebelum Resume
    For i = 1 To resumeIndex - 1
        sheetList.Add ThisWorkbook.Sheets(i).Name
    Next i

    ' Cari kolom "sisa nwt" di baris ke-3
    insertCol = wsResume.Cells(3, wsResume.Columns.Count).End(xlToLeft).Column + 1 ' Default if not found
    For i = 3 To wsResume.Cells(3, wsResume.Columns.Count).End(xlToLeft).Column
        If LCase(Trim(wsResume.Cells(3, i).Value)) = "sisa nwt" Then
            insertCol = i
            Exit For
        End If
    Next i

    ' Pindahkan kolom "sisa nwt" beserta 2 kolom setelahnya ke paling kanan
    targetCol = 3 + sheetList.Count
    If insertCol <> targetCol Then
        wsResume.Columns(insertCol & ":" & insertCol + afterCols - 1).Cut
        wsResume.Columns(targetCol).Insert Shift:=xlToRight
    End If

    ' Hapus konten lama pada kolom SPEC
    If targetCol > 3 Then
        wsResume.Range(wsResume.Cells(3, 3), wsResume.Cells(16, targetCol - 1)).ClearContents
    End If

    ' Isi ulang kolom-kolom SPEC
    colOffset = 3
    For Each wsName In sheetList
        wsResume.Cells(3, colOffset).Value = wsName
        wsResume.Cells(3, colOffset).Font.Bold = True

        For i = 4 To 16
            If wsResume.Cells(i, 2).Value <> "" Then
                wsResume.Cells(i, colOffset).Formula = "=VLOOKUP($B" & i & ",'" & wsName & "'!$L$2:$N$15,3,FALSE)"
                wsResume.Cells(i, colOffset).NumberFormat = "0.00%"
            End If
        Next i

        colOffset = colOffset + 1
    Next wsName

    ' Tambahkan border sampai sebelum "sisa nwt"
    If colOffset > 3 Then
        wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset - 1)).BorderAround ColorIndex:=1, Weight:=xlMedium
        wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset - 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset - 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
    End If

    ShowInfo "Update selesai: " & sheetList.Count & " SPEC sheet berhasil dimasukkan!"
End Sub


Sub CekSheetSebelumResume()
    Dim i As Long, resumeIndex As Long
    Dim logMsg As String
    Dim ws As Worksheet

    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i

    logMsg = "Sheet sebelum Resume:" & vbNewLine
    For i = 1 To resumeIndex - 1
        Set ws = ThisWorkbook.Sheets(i)
        logMsg = logMsg & ws.Name & vbNewLine
    Next i

    MsgBox logMsg, vbInformation
End Sub
Sub UpdateResumeSheet_Dinamis2()
    Dim wsResume As Worksheet
    Dim resumeIndex As Long
    Dim i As Long
    Dim colOffset As Long
    Dim sheetList As Collection
    Dim wsName As Variant
    Dim lastCol As Long
    Dim sisaNWTCol As Long
    Dim endCol As Long
    Dim sisaNWTWidth As Long

    Set sheetList = New Collection
    Set wsResume = ThisWorkbook.Sheets("RESUME")
    
    ' Temukan index sheet Resume
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i

    ' Ambil semua sheet sebelum Resume
    For i = 1 To resumeIndex - 1
        sheetList.Add ThisWorkbook.Sheets(i).Name
    Next i

    ' Cari kolom "sisa nwt" di baris ke-3
    lastCol = wsResume.Cells(3, wsResume.Columns.Count).End(xlToLeft).Column
    sisaNWTCol = 0
    For i = 3 To lastCol
        If LCase(Trim(wsResume.Cells(3, i).Value)) = "sisa nwt" Then
            sisaNWTCol = i
            Exit For
        End If
    Next i
    If sisaNWTCol = 0 Then
        MsgBox "Kolom 'sisa nwt' tidak ditemukan!"
        Exit Sub
    End If
    endCol = lastCol
    sisaNWTWidth = endCol - sisaNWTCol + 1

    ' 1. Copy semua kolom mulai "sisa nwt" sampai terakhir ke array sementara
    Dim tempArr As Variant
    tempArr = wsResume.Range(wsResume.Cells(3, sisaNWTCol), wsResume.Cells(16, endCol)).Value

    ' 2. Clear semua kolom SPEC & kolom SISA NWT dst
    wsResume.Range(wsResume.Cells(3, 3), wsResume.Cells(16, endCol)).Clear

    ' 3. Masukkan sheet spec sebanyak mungkin
    colOffset = 3
    For Each wsName In sheetList
        wsResume.Cells(3, colOffset).Value = wsName
        wsResume.Cells(3, colOffset).Font.Bold = True

        For i = 4 To 16
            If wsResume.Cells(i, 2).Value <> "" Then
                wsResume.Cells(i, colOffset).Formula = "=VLOOKUP($B" & i & ",'" & wsName & "'!$L$2:$N$15,3,FALSE)"
                wsResume.Cells(i, colOffset).NumberFormat = "0.00%"
            End If
        Next i
        colOffset = colOffset + 1
    Next wsName

    ' 4. Paste kembali semua kolom SISA NWT, 2023, dst ke kanan setelah kolom SPEC terakhir
    wsResume.Range(wsResume.Cells(3, colOffset), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).Value = tempArr

    ' 5. Update border sampai kolom terakhir
    wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).BorderAround ColorIndex:=1, Weight:=xlMedium
    wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).Borders(xlInsideVertical).LineStyle = xlContinuous

    MsgBox "Update selesai: " & sheetList.Count & " SPEC sheet berhasil dimasukkan!", vbInformation
End Sub

Sub UpdateResumeSheet_Dinamis3()
    Dim wsResume As Worksheet
    Dim resumeIndex As Long
    Dim i As Long
    Dim colOffset As Long
    Dim sheetList As Collection
    Dim wsName As Variant
    Dim lastCol As Long
    Dim sisaNWTCol As Long
    Dim endCol As Long
    Dim sisaNWTWidth As Long

    Set sheetList = New Collection
    Set wsResume = ThisWorkbook.Sheets("RESUME")
    
    ' Temukan index sheet Resume
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i

    ' Ambil semua sheet sebelum Resume
    For i = 1 To resumeIndex - 1
        sheetList.Add ThisWorkbook.Sheets(i).Name
    Next i

    ' Cari kolom "sisa nwt" di baris ke-3
    lastCol = wsResume.Cells(3, wsResume.Columns.Count).End(xlToLeft).Column
    sisaNWTCol = 0
    For i = 3 To lastCol
        If LCase(Trim(wsResume.Cells(3, i).Value)) = "sisa nwt" Then
            sisaNWTCol = i
            Exit For
        End If
    Next i
    If sisaNWTCol = 0 Then
        MsgBox "Kolom 'sisa nwt' tidak ditemukan!"
        Exit Sub
    End If
    endCol = lastCol
    sisaNWTWidth = endCol - sisaNWTCol + 1

    ' 1. Copy semua kolom mulai "sisa nwt" sampai terakhir ke array sementara
    Dim tempArr As Variant
    tempArr = wsResume.Range(wsResume.Cells(3, sisaNWTCol), wsResume.Cells(16, endCol)).Value

    ' 2. Clear semua kolom SPEC & kolom SISA NWT dst
    wsResume.Range(wsResume.Cells(3, 3), wsResume.Cells(16, endCol)).Clear

    ' 3. Masukkan sheet spec sebanyak mungkin
    colOffset = 3
    For Each wsName In sheetList
        wsResume.Cells(3, colOffset).Value = wsName
        wsResume.Cells(3, colOffset).Font.Bold = True

        For i = 4 To 16
            If wsResume.Cells(i, 2).Value <> "" Then
                wsResume.Cells(i, colOffset).Formula = "=VLOOKUP($B" & i & ",'" & wsName & "'!$L$2:$N$15,3,FALSE)"
                wsResume.Cells(i, colOffset).NumberFormat = "0.00%"
            End If
        Next i
        colOffset = colOffset + 1
    Next wsName

    ' 4. Paste kembali semua kolom SISA NWT, 2023, dst ke kanan setelah kolom SPEC terakhir
    wsResume.Range(wsResume.Cells(3, colOffset), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).Value = tempArr

    ' 5. Update border sampai kolom terakhir
    wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).BorderAround ColorIndex:=1, Weight:=xlMedium
    wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    wsResume.Range(wsResume.Cells(3, 2), wsResume.Cells(16, colOffset + sisaNWTWidth - 1)).Borders(xlInsideVertical).LineStyle = xlContinuous

    MsgBox "Update selesai: " & sheetList.Count & " SPEC sheet berhasil dimasukkan!", vbInformation
End Sub


