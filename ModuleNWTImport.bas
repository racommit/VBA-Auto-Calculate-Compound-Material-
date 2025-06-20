Attribute VB_Name = "ModuleNWTImport"
'@Folder "resume"
Option Explicit
Option Compare Text

'==== Konfigurasi Awal ====
Const COL_LABEL As Long = 2        ' Kolom untuk label baris (col B)
Const ROW_HEADER As Long = 3       ' Baris header yang memuat nama spec
Const COL_FIRST_DATA As Long = 3   ' Kolom data pertama setelah label

' Label baris perhitungan
Const LABEL_TOTAL As String = "Total (NWT) Production Tires"
Const LABEL_PORTION As String = "Portion per size (%)"
Const LABEL_PORTION_SUS As String = "Portion Material sustainability"
Const LABEL_SUSTAIN_SPEC As String = "Material sustainability"
Const LABEL_TOTAL_NWT As String = "Total NWT"

Sub ImportAndTransposeNWT()
    Dim fd As FileDialog
    Dim srcPath As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim arrB As Variant, arrC As Variant

    ' Pilih file sumber
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Pilih file sumber NWT"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        srcPath = .SelectedItems(1)
    End With

    ' Buka file sumber, sheet pertama
    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
    Set wsSrc = wbSrc.Sheets(1)
    Set wsDest = ThisWorkbook.Sheets("RESUME")

    ' Tentukan data terakhir di kolom B (diasumsikan kolom C sama panjang)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 2).End(xlUp).Row

    ' Ambil array data kolom B2:B(last) dan C2:C(last) dari sumber
    arrB = wsSrc.Range("B2:B" & lastRow).Value
    arrC = wsSrc.Range("C2:C" & lastRow).Value

    ' Paste secara transpose ke sheet RESUME, mulai C17 dan C20
    wsDest.Range("C17").Resize(1, UBound(arrB, 1)).Value = _
        Application.WorksheetFunction.Transpose(arrB)
    wsDest.Range("C20").Resize(1, UBound(arrC, 1)).Value = _
        Application.WorksheetFunction.Transpose(arrC)

    wbSrc.Close False
    MsgBox "Data B2:Bx sumber masuk ke C17, C2:Cx ke C20, secara transpose!", _
        vbInformation
End Sub

Sub ImportNWT_DenganMappingCerdas()
    Dim fd As FileDialog
    Dim srcPath As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim arrSpecs() As Variant, arrTahun() As Variant
    Dim dictSpecCol As Object
    Dim dictTahunCol As Object
    Dim r As Long, c As Long
    Dim destHeader As Variant, destSpecPos As Object
    Dim destYearCol As Object
    Dim destRows() As Long
    Dim spec As Variant
    Dim lastSpecCol As Long
    Dim yearIdx As Long
    Dim yearKey As String
    Dim destRowCount As Long
    Dim rowSustain As Long
    Dim dictTotalNWT As Object
    Dim rowTotalNWT As Long

    '--- Konfigurasi
    Const DEST_SHEET As String = "RESUME"
    Const DEST_ROW_HEADER As Long = 3

    '--- Pilih file sumber
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Pilih file sumber NWT"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        srcPath = .SelectedItems(1)
    End With

    '--- Buka file sumber
    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
    Set wsSrc = wbSrc.Sheets(1)
    Set wsDest = ThisWorkbook.Sheets(DEST_SHEET)

    '--- Identifikasi header tahun pada sumber (baris 1, kolom 2 ke kanan)
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    arrTahun = wsSrc.Range(wsSrc.Cells(1, 2), wsSrc.Cells(1, lastCol)).Value
    '--- Buat dictionary kolom tahun pada sumber
    Set dictTahunCol = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(arrTahun, 2)
        dictTahunCol(Trim(arrTahun(1, c))) = c + 1 ' kolom 1=spec, kolom 2=tahun1, dst
    Next

    '--- Identifikasi seluruh spec pada sumber (kolom 1, baris 2 ke bawah)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    arrSpecs = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, 1)).Value

    '--- Buat dictionary baris spec pada sumber
    Set dictSpecCol = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(arrSpecs, 1)
        dictSpecCol(Trim(arrSpecs(r, 1))) = r + 1 ' baris 1 = header
    Next

    '--- Baca nilai Total NWT bila ada
    Set dictTotalNWT = CreateObject("Scripting.Dictionary")
    rowTotalNWT = 0
    If dictSpecCol.Exists(LABEL_TOTAL_NWT) Then
        rowTotalNWT = dictSpecCol(LABEL_TOTAL_NWT)
    Else
        'periksa kolom B jika label bukan di kolom A
        For r = 2 To lastRow
            If Trim(wsSrc.Cells(r, 2).Value) = LABEL_TOTAL_NWT Then
                rowTotalNWT = r
                Exit For
            End If
        Next r
    End If
    If rowTotalNWT > 0 Then
        For c = 1 To UBound(arrTahun, 2)
            yearKey = Trim(arrTahun(1, c))
            If dictTahunCol.Exists(yearKey) Then
                dictTotalNWT(yearKey) = wsSrc.Cells(rowTotalNWT, dictTahunCol(yearKey)).Value
            End If
        Next c
    End If

    '--- Ambil header spec pada sheet RESUME (baris 3, kolom 3 ke kanan)
    lastCol = wsDest.Cells(DEST_ROW_HEADER, wsDest.Columns.Count).End(xlToLeft).Column
    destHeader = wsDest.Range(wsDest.Cells(DEST_ROW_HEADER, 3), _
                              wsDest.Cells(DEST_ROW_HEADER, lastCol)).Value

    '--- Buat dictionary posisi kolom spec di sheet tujuan dan kolom total per tahun
    Set destSpecPos = CreateObject("Scripting.Dictionary")
    Set destYearCol = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(destHeader, 2)
        Dim hdrVal As String
        hdrVal = Trim(destHeader(1, c))
        If hdrVal <> "" Then
            If IsNumeric(hdrVal) Then
                destYearCol(CLng(hdrVal)) = c + 2
            Else
                destSpecPos(hdrVal) = c + 2
            End If
        End If
    Next c

    ' Tentukan kolom terakhir untuk spec yang ada di sumber
    lastSpecCol = COL_FIRST_DATA
    For Each spec In destSpecPos.Keys
        If dictSpecCol.Exists(Trim(spec)) Then
            If destSpecPos(spec) > lastSpecCol Then lastSpecCol = destSpecPos(spec)
        End If
    Next spec

    '--- Tambahkan kolom "sisa nwt" jika ada di sumber namun belum ada di tujuan
    If dictSpecCol.Exists("sisa nwt") Then
        If Not destSpecPos.Exists("sisa nwt") Then
            lastCol = lastCol + 1
            wsDest.Cells(DEST_ROW_HEADER, lastCol).Value = "sisa nwt"
            destSpecPos("sisa nwt") = lastCol
        End If
    End If

    '--- Temukan baris tujuan untuk setiap tahun berdasarkan label
    '    Gunakan baris dengan label "Total (NWT) Production Tires" karena
    '    baris tersebut merupakan awal blok data untuk suatu tahun.
    lastRow = wsDest.Cells(wsDest.Rows.Count, COL_LABEL).End(xlUp).Row
    rowSustain = 0
    For r = 1 To lastRow
        If Trim(wsDest.Cells(r, COL_LABEL).Value) = LABEL_TOTAL Then
            destRowCount = destRowCount + 1
            ReDim Preserve destRows(1 To destRowCount)
            destRows(destRowCount) = r
        End If
        If rowSustain = 0 And Trim(wsDest.Cells(r, COL_LABEL).Value) = LABEL_SUSTAIN_SPEC Then
            rowSustain = r
        End If
    Next

    '--- Pindahkan data sesuai tahun dan spec
    For yearIdx = 1 To Application.WorksheetFunction.Min(UBound(arrTahun, 2), destRowCount)
        yearKey = Trim(arrTahun(1, yearIdx))
        If dictTahunCol.Exists(yearKey) Then
            c = dictTahunCol(yearKey)
            For Each spec In destSpecPos.Keys
                If dictSpecCol.Exists(Trim(spec)) Then
                    r = dictSpecCol(Trim(spec))
                    wsDest.Cells(destRows(yearIdx), destSpecPos(spec)).Value = wsSrc.Cells(r, c).Value
                End If
            Next spec

            'Total NWT untuk tahun ini diambil dari sumber
            If destYearCol.Exists(CLng(yearKey)) And dictTotalNWT.Exists(yearKey) Then
                wsDest.Cells(destRows(yearIdx), destYearCol(CLng(yearKey))).Value = dictTotalNWT(yearKey)
            End If

            ' -- Isi baris "Portion per size (%)" dengan rumus persentase
            If Trim(wsDest.Cells(destRows(yearIdx) + 1, COL_LABEL).Value) = LABEL_PORTION Then
                For Each spec In destSpecPos.Keys
                    If dictSpecCol.Exists(Trim(spec)) And destYearCol.Exists(CLng(yearKey)) Then
                        Dim numAddr As String, totalAddr As String
                        numAddr = wsDest.Cells(destRows(yearIdx), destSpecPos(spec)).Address(False, False)
                        totalAddr = wsDest.Cells(destRows(yearIdx), destYearCol(CLng(yearKey))).Address(True, True)
                        wsDest.Cells(destRows(yearIdx) + 1, destSpecPos(spec)).Formula = "=" & numAddr & "/" & totalAddr
                        wsDest.Cells(destRows(yearIdx) + 1, destSpecPos(spec)).NumberFormat = "0.00%"
                    End If
                Next spec
            End If

            ' -- Hitung Portion Material sustainability
            If rowSustain > 0 Then
                If Trim(wsDest.Cells(destRows(yearIdx) + 2, COL_LABEL).Value) = LABEL_PORTION_SUS Then
                    For Each spec In destSpecPos.Keys
                        If dictSpecCol.Exists(Trim(spec)) Then
                            Dim susAddr As String, portAddr As String
                            susAddr = wsDest.Cells(rowSustain, destSpecPos(spec)).Address(False, False)
                            portAddr = wsDest.Cells(destRows(yearIdx) + 1, destSpecPos(spec)).Address(False, False)
                            wsDest.Cells(destRows(yearIdx) + 2, destSpecPos(spec)).Formula = "=" & susAddr & "*" & portAddr
                            wsDest.Cells(destRows(yearIdx) + 2, destSpecPos(spec)).NumberFormat = "0.00%"
                        End If
                    Next spec
                End If
            End If
        End If
    Next yearIdx

    '--- Hitung total baris Portion per size (%) dan Portion Material sustainability
    For yearIdx = 1 To destRowCount
        yearKey = Trim(arrTahun(1, yearIdx))
        If destYearCol.Exists(CLng(yearKey)) Then
            Dim addrStart As String, addrEnd As String

            'Total untuk Portion per size (%)
            If Trim(wsDest.Cells(destRows(yearIdx) + 1, COL_LABEL).Value) = LABEL_PORTION Then
                addrStart = wsDest.Cells(destRows(yearIdx) + 1, COL_FIRST_DATA).Address(False, False)
                addrEnd = wsDest.Cells(destRows(yearIdx) + 1, lastSpecCol).Address(False, False)
                wsDest.Cells(destRows(yearIdx) + 1, destYearCol(CLng(yearKey))).Formula = "=SUM(" & addrStart & ":" & addrEnd & ")"
            End If

            'Total untuk Portion Material sustainability
            If Trim(wsDest.Cells(destRows(yearIdx) + 2, COL_LABEL).Value) = LABEL_PORTION_SUS Then
                addrStart = wsDest.Cells(destRows(yearIdx) + 2, COL_FIRST_DATA).Address(False, False)
                addrEnd = wsDest.Cells(destRows(yearIdx) + 2, lastSpecCol).Address(False, False)
                wsDest.Cells(destRows(yearIdx) + 2, destYearCol(CLng(yearKey))).Formula = "=SUM(" & addrStart & ":" & addrEnd & ")"
        End If
        End If
    Next yearIdx

    wbSrc.Close False
    MsgBox "Impor dan mapping selesai!", vbInformation
End Sub
