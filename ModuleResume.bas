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
    Dim lastCol As Long

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

    ' Hapus kolom lama hanya dari kolom ke-3 sampai sebelum kolom "sisa nwt"
    If insertCol > 3 Then
        wsResume.Range(wsResume.Cells(3, 3), wsResume.Cells(16, insertCol - 1)).ClearContents
    End If

    ' Masukkan ulang kolom-kolom baru sebelum kolom "sisa nwt"
    colOffset = 3
    For Each wsName In sheetList
        If colOffset >= insertCol Then Exit For ' Hindari menimpa "sisa nwt" dan setelahnya

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

    ShowInfo logMsg
End Sub

