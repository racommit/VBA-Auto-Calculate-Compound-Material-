Attribute VB_Name = "ModuleNWTResume"
'@Folder "resume"
Option Explicit

'==== Konfigurasi Awal ====
Const COL_LABEL As Long = 2        ' Kolom untuk label baris (col B)
Const ROW_HEADER As Long = 3       ' Baris header yang memuat nama spec
Const COL_FIRST_DATA As Long = 3   ' Kolom data pertama setelah label

' Label baris perhitungan
Const LABEL_TOTAL As String = "Total (NWT) Production Tires"
Const LABEL_PORTION As String = "Portion per size (%)"
Const LABEL_PORTION_SUS As String = "Portion Material sustainability"
Const LABEL_SUSTAIN_SPEC As String = "Material sustainability"

Sub TarikNWT_danKalkulasi()
    Dim fd As FileDialog
    Dim srcPath As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Pilih file sumber NWT"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        srcPath = .SelectedItems(1)
    End With

    TarikNWT_dinamis srcPath
    KalkulasiResumeAuto

    MsgBox "Proses penarikan dan kalkulasi selesai!", vbInformation
End Sub

' Tarik data NWT per tahun/spec secara dinamis
Sub TarikNWT_dinamis(ByVal srcPath As String)
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim specCols As Object, yearRows As Object
    Dim srcSpecRows As Object, srcYearCols As Object
    Dim spec As Variant, yr As Variant
    Dim oldCalc As XlCalculation

    BeginFastMode oldCalc

    Set wsDest = ThisWorkbook.Sheets("RESUME")

    ' Petakan kolom spec pada sheet RESUME
    Set specCols = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = COL_FIRST_DATA To wsDest.Cells(ROW_HEADER, wsDest.Columns.Count).End(xlToLeft).Column
        spec = Trim(wsDest.Cells(ROW_HEADER, c).Value)
        If spec <> "" Then specCols(spec) = c
    Next c

    ' Petakan baris per tahun (mengacu pada angka 4 digit di kolom B)
    Set yearRows = CreateObject("Scripting.Dictionary")
    Dim r As Long, lastRow As Long
    lastRow = wsDest.Cells(wsDest.Rows.Count, COL_LABEL).End(xlUp).Row
    For r = 1 To lastRow
        If IsNumeric(wsDest.Cells(r, COL_LABEL).Value) And Len(wsDest.Cells(r, COL_LABEL).Value) = 4 Then
            yearRows(CLng(wsDest.Cells(r, COL_LABEL).Value)) = r
        End If
    Next r

    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
    Set wsSrc = wbSrc.Sheets(1)

    ' Petakan baris spec dan kolom tahun pada file sumber
    Set srcSpecRows = CreateObject("Scripting.Dictionary")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        spec = Trim(wsSrc.Cells(r, 1).Value)
        If spec <> "" Then srcSpecRows(spec) = r
    Next r

    Set srcYearCols = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    For c = 2 To lastCol
        If IsNumeric(wsSrc.Cells(1, c).Value) Then
            srcYearCols(CLng(wsSrc.Cells(1, c).Value)) = c
        End If
    Next c

    ' Salin data sesuai spec dan tahun yang cocok
    For Each spec In specCols.Keys
        If srcSpecRows.Exists(spec) Then
            For Each yr In yearRows.Keys
                If srcYearCols.Exists(yr) Then
                    wsDest.Cells(yearRows(yr) + 1, specCols(spec)).Value = wsSrc.Cells(srcSpecRows(spec), srcYearCols(yr)).Value
                End If
            Next yr
        End If
    Next spec

    wbSrc.Close False
    EndFastMode oldCalc
End Sub

' Hitung Portion per size dan Portion Material sustainability
Sub KalkulasiResumeAuto()
    Dim ws As Worksheet
    Dim specCols As Object, yearRows As Object
    Dim rowSustain As Long
    Dim spec As Variant, yr As Variant
    Dim total As Double
    Dim col As Long

    Set ws = ThisWorkbook.Sheets("RESUME")

    ' Peta kolom spec
    Set specCols = CreateObject("Scripting.Dictionary")
    For col = COL_FIRST_DATA To ws.Cells(ROW_HEADER, ws.Columns.Count).End(xlToLeft).Column
        spec = Trim(ws.Cells(ROW_HEADER, col).Value)
        If spec <> "" Then specCols(spec) = col
    Next col

    ' Peta baris tahun
    Set yearRows = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_LABEL).End(xlUp).Row
    For r = 1 To lastRow
        If IsNumeric(ws.Cells(r, COL_LABEL).Value) And Len(ws.Cells(r, COL_LABEL).Value) = 4 Then
            yearRows(CLng(ws.Cells(r, COL_LABEL).Value)) = r
        End If
    Next r

    ' Cari baris material sustainability per spec
    rowSustain = 0
    For r = 1 To lastRow
        If LCase(Trim(ws.Cells(r, COL_LABEL).Value)) = LCase(LABEL_SUSTAIN_SPEC) Then
            rowSustain = r
            Exit For
        End If
    Next r

    If rowSustain = 0 Then
        MsgBox "Baris '" & LABEL_SUSTAIN_SPEC & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If

    ' Hitung untuk setiap tahun
    For Each yr In yearRows.Keys
        Dim rowNWT As Long, rowPortion As Long, rowPortionMat As Long
        rowNWT = yearRows(yr) + 1
        rowPortion = yearRows(yr) + 2
        rowPortionMat = yearRows(yr) + 3

        ' Hitung total NWT tahun tersebut
        total = 0
        For Each spec In specCols.Keys
            total = total + val(ws.Cells(rowNWT, specCols(spec)).Value)
        Next spec

        ' Hitung portion dan portion sustainability per spec
        For Each spec In specCols.Keys
            Dim nwtVal As Double
            nwtVal = val(ws.Cells(rowNWT, specCols(spec)).Value)
            If total <> 0 Then
                ws.Cells(rowPortion, specCols(spec)).Value = nwtVal / total
                ws.Cells(rowPortionMat, specCols(spec)).Value = _
                    ws.Cells(rowSustain, specCols(spec)).Value * (nwtVal / total)
            Else
                ws.Cells(rowPortion, specCols(spec)).Value = 0
                ws.Cells(rowPortionMat, specCols(spec)).Value = 0
            End If
        Next spec
    Next yr
End Sub


