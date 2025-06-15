Attribute VB_Name = "RunKombinasi"

'@NoIndent
'@IgnoreModule AssignmentNotUsed
'@Folder "Core_calculate"
Option Explicit
'@Ignore ModuleScopeDimKeyword
Dim GlobalActionID As String

' Variabel global untuk menyimpan perhitungan sustainability
Dim GlobalSustainabilityBefore As Double
Dim GlobalSustainabilityAfter As Double
Dim GlobalTotalBefore As Double
Dim GlobalTotalAfter As Double
Dim GlobalPercentageBefore As Double
Dim GlobalPercentageAfter As Double

Dim totalallportion As Double
Dim totalallportionBefore As Double

' Dictionary untuk menyimpan backup data material
Dim MaterialBackupData As Object

' material count
Dim MaterialOldBefore(1 To 3) As Double
Dim MaterialOldAfter(1 To 3) As Double
Dim MaterialNewBefore(1 To 3) As Double
Dim MaterialNewAfter(1 To 3) As Double


' Structure untuk menyimpan data replacement
Type replacementData
    material_replaced As String
    percentage_material As Double
    new_material As String
    percentage_new_material As Double
    new_material_class As String
    isValid As Boolean
End Type

' Structure untuk menyimpan backup data
Type MaterialBackup
    sheetName As String
    rowNumber As Long
    columnLetter As String
    originalValue As Double
    materialName As String
    backupType As String                         ' "MODIFIED" atau "ADDED"
End Type

Type MaterialEdit
    rowIndex As Long
    sheetName As String
    oldValue As Double
    newValue As Double
    replacementIdx As Long
End Type
Sub submit_multiple_replacement3()
    Dim replacements(1 To 3) As replacementData
    Dim i As Long
    Dim validCount As Long
    
    ' Initialize backup dictionary
    Set MaterialBackupData = CreateObject("Scripting.Dictionary")
    
    ' Input ranges: C5:C9, G5:G9, K5:K9
    Dim inputColumns As Variant
    inputColumns = Array(3, 7, 11)               ' C=3, G=7, K=11
    Dim WSsim As Worksheet
    Set WSsim = ThisWorkbook.Sheets("SIMULATION_PROCESS")
    
    
    WSsim.Range("E71:E73").ClearContents
    WSsim.Range("F71:F73").ClearContents
  
    
    Dim kelasLama As String, kelasBaru As String


For i = 1 To 3
    If replacements(i).isValid Then
        kelasLama = CariKategoriKelas(replacements(i).material_replaced)
        kelasBaru = UCase(Trim(replacements(i).new_material_class))
        
        Debug.Print kelasLama
        Debug.Print kelasBaru
        
       If kelasLama <> "" And kelasBaru <> "" Then
           If InStr(1, kelasLama, kelasBaru, vbTextCompare) = 0 And InStr(1, kelasBaru, kelasLama, vbTextCompare) = 0 Then
                MsgBox "Kategori kelas material lama (" & kelasLama & ") berbeda dengan kelas material baru (" & kelasBaru & ") pada set-" & i & "." & vbCrLf & _
                       "Proses dibatalkan!", vbCritical, "Validasi Kelas Tidak Cocok"
                Exit Sub
            End If
        Else
            MsgBox "Kategori kelas material lama atau baru tidak ditemukan di sheet CATEGORY.", vbCritical, "Error Validasi Kelas"
            Exit Sub
        End If
    End If
Next i

    
    ' Baca dan validasi semua input
    For i = 1 To 3
        With replacements(i)
            .material_replaced = UCase(Trim(Cells(5, inputColumns(i - 1)).Value))
            .new_material = UCase(Trim(Cells(7, inputColumns(i - 1)).Value))
            .new_material_class = UCase(Trim(Cells(9, inputColumns(i - 1)).Value))
            
            
            
            
            ' Validasi numerik
            If IsNumeric(Cells(6, inputColumns(i - 1)).Value) And IsNumeric(Cells(8, inputColumns(i - 1)).Value) Then
                .percentage_material = CDbl(Cells(6, inputColumns(i - 1)).Value)
                .percentage_new_material = CDbl(Cells(8, inputColumns(i - 1)).Value)
            Else
                .percentage_material = 0
                .percentage_new_material = 0
            End If
            
            If .percentage_material <> .percentage_new_material Then
                MsgBox "Nilai pergantian tidak sesuai", vbExclamation
            
                Exit Sub
            End If
            
            ' Cek validitas input
            .isValid = (.material_replaced <> "" And .new_material <> "" And _
                        .new_material_class <> "" And .percentage_material > 0 And _
                        .percentage_new_material > 0)
            
            If .isValid Then validCount = validCount + 1
        End With
    Next i
    
    WSsim.Range("D8").Value = replacements(1).new_material
    WSsim.Range("D9").Value = replacements(2).new_material
    WSsim.Range("D10").Value = replacements(3).new_material
            
    WSsim.Range("D13").Value = replacements(1).new_material_class
    WSsim.Range("D14").Value = replacements(2).new_material_class
    WSsim.Range("D15").Value = replacements(3).new_material_class
    
    WSsim.Range("C65").Value = replacements(1).material_replaced
    WSsim.Range("C66").Value = replacements(2).material_replaced
    WSsim.Range("C67").Value = replacements(3).material_replaced
    
    If validCount = 0 Then
        MsgBox "Tidak ada input yang valid untuk diproses.", vbExclamation
        Exit Sub
    End If
    
    ' Validasi material di CATEGORY SPECIFICATION
    Dim wsCategory As Worksheet
    Set wsCategory = ThisWorkbook.Sheets("CATEGORY SPESIFICATION")
    
    For i = 1 To 3
        If replacements(i).isValid Then
            Dim foundCell As Range
            Set foundCell = wsCategory.Range("D3:D144").Find(What:=replacements(i).material_replaced, LookIn:=xlValues, LookAt:=xlWhole)
            
            If foundCell Is Nothing Then
                MsgBox "Kode material " & replacements(i).material_replaced & " (Set " & i & ") tidak ditemukan.", vbExclamation
                Exit Sub
            End If
        End If
    Next i
    
    ' Hitung sustainability BEFORE replacement
    Call CalculateSustainabilityBefore
    
    ' Generate unique Action ID
    GlobalActionID = CLng(Timer * 1000)
    
    ' Process semua replacement (dengan backup)
    Call update_multiple_material_data_with_backup(replacements)
    
    ' Hitung sustainability AFTER replacement
    Call CalculateSustainabilityAfter
    
    ' Tampilkan hasil sustainability
    Call DisplaySustainabilityResults
    
    ' Simpan hasil dulu sebelum restore
    Dim tempAfter As Double
    tempAfter = GlobalPercentageAfter

    ' **RESTORE MATERIAL VALUES TO ORIGINAL STATE**
    Call RestoreMaterialValues

    ' Set kembali nilai After yang sudah dihitung
    GlobalPercentageAfter = tempAfter
    
    
    WSsim.Range("D27").Value = GlobalPercentageBefore ' Total Sustainability Before
    WSsim.Range("E8").Value = GlobalPercentageBefore
    WSsim.Range("E9").Value = GlobalPercentageBefore
    WSsim.Range("E10").Value = GlobalPercentageBefore
    WSsim.Range("D28").Value = GlobalPercentageAfter ' Total Sustainability After
    
    If replacements(1).isValid Then
        WSsim.Range("E65").Value = MaterialOldBefore(1)
        WSsim.Range("F65").Value = MaterialOldAfter(1)
    Else
        WSsim.Range("E65").ClearContents
        WSsim.Range("F65").ClearContents
    End If

    If replacements(2).isValid Then
        WSsim.Range("E66").Value = MaterialOldBefore(2)
        WSsim.Range("F66").Value = MaterialOldAfter(2)
    Else
        WSsim.Range("E66").ClearContents
        WSsim.Range("F66").ClearContents
    End If

    If replacements(3).isValid Then
        WSsim.Range("E67").Value = MaterialOldBefore(3)
        WSsim.Range("F67").Value = MaterialOldAfter(3)
    Else
        WSsim.Range("E67").ClearContents
        WSsim.Range("F67").ClearContents
    End If

    If replacements(1).isValid Then
        WSsim.Range("E71").Value = MaterialNewBefore(1)
        WSsim.Range("F71").Value = MaterialNewAfter(1)
    Else
        WSsim.Range("E71").ClearContents
        WSsim.Range("F71").ClearContents
    End If

    If replacements(2).isValid Then
        WSsim.Range("E72").Value = MaterialNewBefore(2)
        WSsim.Range("F72").Value = MaterialNewAfter(2)
    Else
        WSsim.Range("E72").ClearContents
        WSsim.Range("F72").ClearContents
    End If

    If replacements(3).isValid Then
        WSsim.Range("E73").Value = MaterialNewBefore(3)
        WSsim.Range("F73").Value = MaterialNewAfter(3)
    Else
        WSsim.Range("E73").ClearContents
        WSsim.Range("F73").ClearContents
    End If






    
    ' MsgBox "Proses " & validCount & " material replacement berhasil (Action ID: " & GlobalActionID & ")" & vbCrLf & _
    "Sustainability Before: " & Format(GlobalPercentageBefore, "0.00%") & vbCrLf & _
    "Sustainability After: " & Format(GlobalPercentageAfter, "0.00%") & vbCrLf & _
    "Data material telah dikembalikan ke kondisi semula.", vbInformation
End Sub
Function CariKategoriKelas(materialName As String) As String
    Dim wsCat As Worksheet
    Set wsCat = ThisWorkbook.Sheets("CATEGORY SPESIFICATION")
    Dim rng As Range, found As Range
    Set rng = wsCat.Range("D3:D144") ' D = nama material, E = kategori kelas
    Set found = rng.Find(What:=materialName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not found Is Nothing Then
        CariKategoriKelas = UCase(Trim(wsCat.Cells(found.Row, "E").Value)) ' Ambil kategori kelas di kolom E
    Else
        CariKategoriKelas = ""
    End If
End Function

Sub CalculateSustainabilityBefore()
    ' Reset variabel global
    GlobalSustainabilityBefore = 0
    GlobalTotalBefore = 0
    GlobalPercentageBefore = 0
    
    ' Dictionary untuk menyimpan perhitungan per spesifikasi
    Dim specSustainabilityBefore As Object
    Dim specTotalBefore As Object
    Set specSustainabilityBefore = CreateObject("Scripting.Dictionary")
    Set specTotalBefore = CreateObject("Scripting.Dictionary")
    
    ' Nama kategori sustainability (6 kategori)
    Dim sustainableCategoryNames(1 To 6) As String
    sustainableCategoryNames(1) = "NATURAL RUBBER"
    sustainableCategoryNames(2) = "RECLAIM RUBBER"
    sustainableCategoryNames(3) = "SUSTAINABLE FILLER"
    sustainableCategoryNames(4) = "SUSTAINABLE OIL"
    sustainableCategoryNames(5) = "SUSTAINABLE CHEMICAL"
    sustainableCategoryNames(6) = "SUSTAINABLE REINFORCEMENT"
    
    Dim sh As Worksheet
    Dim resumeIndex As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' Cari index sheet RESUME
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i
    
    ' Loop semua sheet sebelum RESUME untuk menghitung per spesifikasi
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" Then
            Dim specName As String
            specName = sh.Name
            
            ' Inisialisasi dictionary untuk spesifikasi ini
            specSustainabilityBefore.Add specName, 0
            specTotalBefore.Add specName, 0
            
            lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
            
            Dim currentCategory As String
            currentCategory = ""
            
            ' Loop setiap baris di sheet
            Dim Row As Long
            For Row = 3 To lastRow
                ' Cek jika baris kategori (ada nilai di kolom J)
                If Not IsEmpty(sh.Cells(Row, "J").Value) And IsNumeric(sh.Cells(Row, "J").Value) Then
                    currentCategory = Trim(UCase(sh.Cells(Row, "H").Value))
                    
                    ' Ambil nilai total kategori dari kolom J
                    Dim nilaiKategoriJ As Double
                    nilaiKategoriJ = sh.Cells(Row, "J").Value
                    
                    ' Tambahkan ke total spesifikasi ini
                    specTotalBefore(specName) = specTotalBefore(specName) + nilaiKategoriJ
                    
                    ' Cek jika kategori ini adalah sustainability category
                    Dim catIdx As Long
                    For catIdx = 1 To 6
                        If sustainableCategoryNames(catIdx) = currentCategory Then
                            specSustainabilityBefore(specName) = specSustainabilityBefore(specName) + nilaiKategoriJ
                        End If
                    Next catIdx
                End If
            Next Row
        End If
    Next i
    
    Dim validSheetCount As Long
    validSheetCount = 0
    For i = 1 To resumeIndex - 1
        If ThisWorkbook.Sheets(i).Name <> "CATEGORY SPESIFICATION" Then
            validSheetCount = validSheetCount + 1
        End If
    Next i
    
    ' ---------------------------------------------
    ' HITUNG SUSTAINABILITY BEFORE (TERTIMBANG)
    ' ---------------------------------------------
    Dim daftarspecBefore() As Variant
    ReDim daftarspecBefore(1 To validSheetCount + 1)

    Dim totalPersentaseSustainabilityBefore As Double
    totalPersentaseSustainabilityBefore = 0
    Dim specKey As Variant
    Dim idx As Long
    idx = 1

    For Each specKey In specSustainabilityBefore.Keys
        If specTotalBefore(specKey) > 0 Then
            daftarspecBefore(idx) = specSustainabilityBefore(specKey) / specTotalBefore(specKey)
        Else
            daftarspecBefore(idx) = 0
        End If
        totalPersentaseSustainabilityBefore = totalPersentaseSustainabilityBefore + daftarspecBefore(idx)
        GlobalSustainabilityBefore = GlobalSustainabilityBefore + specSustainabilityBefore(specKey)
        GlobalTotalBefore = GlobalTotalBefore + specTotalBefore(specKey)
        idx = idx + 1
    Next specKey
    
    ' Gunakan daftarspec() untuk dikalikan portion()
    Dim hitungportion() As Variant
    Dim portion() As Variant, n As Long
    portion = AmbilMassHorizontal()
    n = UBound(portion)
    
    Debug.Print UBound(daftarspecBefore)
    Debug.Print UBound(portion)
    
    Dim wssum As Worksheet
    Set wssum = ThisWorkbook.Sheets("RESUME")    ' Ganti sesuai nama sheet jika berbeda

    Dim colSisaNWT As Long
    Dim rowMaterialSustainability As Long
    Dim nilaiGabungan As Variant

   

    ' --- Cari kolom yang berisi "sisa nwt" di baris ke-3 ---
    For i = 1 To wssum.Cells(3, wssum.Columns.Count).End(xlToLeft).Column
        If LCase(Trim(wssum.Cells(3, i).Value)) = "sisa nwt" Then
            colSisaNWT = i
            Exit For
        End If
    Next i

    If colSisaNWT = 0 Then
        MsgBox "Kolom 'sisa nwt' tidak ditemukan di baris 3.", vbExclamation
        Exit Sub
    End If

    ' --- Cari baris yang berisi "material sustainability" di kolom B ---
    For i = 1 To wssum.Cells(wssum.Rows.Count, 2).End(xlUp).Row
        If LCase(Trim(wssum.Cells(i, 2).Value)) = "material sustainability" Then
            rowMaterialSustainability = i
            Exit For
        End If
    Next i

    If rowMaterialSustainability = 0 Then
        MsgBox "Baris 'material sustainability' tidak ditemukan di kolom B.", vbExclamation
        Exit Sub
    End If

    ' --- Ambil nilai dari sel perpotongan ---
    nilaiGabungan = wssum.Cells(rowMaterialSustainability, colSisaNWT).Value
    If nilaiGabungan <> 0 Or nilaiGabungan <> "" Then
        daftarspecBefore(validSheetCount + 1) = nilaiGabungan
    End If
   
    
    ReDim hitungportionBefore(1 To n)

    For i = 1 To n
       
        hitungportionBefore(i) = daftarspecBefore(i) * portion(i)
       
    Next i

    
    totalallportionBefore = 0
    For i = 1 To UBound(hitungportionBefore)
        totalallportionBefore = totalallportionBefore + hitungportionBefore(i)
    Next i

    Debug.Print "TOTAL Sustainability Before (tertimbang): " & totalallportionBefore
    ThisWorkbook.Sheets("SIMULATION_PROCESS").Range("E27").Value = totalallportionBefore
    Debug.Print "test"
    If GlobalTotalBefore > 0 Then
        GlobalPercentageBefore = totalPersentaseSustainabilityBefore
    Else
        GlobalPercentageBefore = 0
        
    End If
    
    Debug.Print "Sustainability Before Calculation Completed"
    Debug.Print "Total Sustainability Before: " & GlobalSustainabilityBefore
    Debug.Print "Total Before: " & GlobalTotalBefore
    Debug.Print "Percentage Before: " & GlobalPercentageBefore
    
    ' === RATA-RATA PER KATEGORI (ACROSS ALL SPEC SHEETS) ===
    Dim categorySum As Object
    Dim categoryCount As Object
    Set categorySum = CreateObject("Scripting.Dictionary")
    Set categoryCount = CreateObject("Scripting.Dictionary")

    ' Loop semua sheet spesifikasi
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" Then
            lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
        
            For Row = 3 To lastRow
                If Not IsEmpty(sh.Cells(Row, "J").Value) And IsNumeric(sh.Cells(Row, "J").Value) Then
                    Dim catName As String
                    Dim catValue As Double
                    catName = Trim(UCase(sh.Cells(Row, "H").Value))
                    catValue = sh.Cells(Row, "J").Value
                   
                    If categorySum.exists(catName) Then
                        Debug.Print i
                        Debug.Print catName
                        Debug.Print catValue
                        categorySum(catName) = categorySum(catName) + catValue
                        Debug.Print categorySum(catName)
                        categoryCount(catName) = categoryCount(catName) + 1
                    Else
                        categorySum.Add catName, catValue
                        categoryCount.Add catName, 1
                    End If
                End If
            Next Row
        End If
    Next i

  ' Array urutan kategori yang diinginkan (urutannya seperti laporan)
Dim orderedCategories As Variant
orderedCategories = Array( _
    "NATURAL RUBBER", _
    "SYNTHETIC RUBBER", _
    "RECLAIM RUBBER", _
    "FILLER", _
    "SUSTAINABLE FILLER", _
    "OIL", _
    "SUSTAINABLE OIL", _
    "CHEMICAL", _
    "SUSTAINABLE CHEMICAL", _
    "REINFORCEMENT", _
    "SUSTAINABLE REINFORCEMENT")

Dim persentaseKategoriPerSpec As Object
Set persentaseKategoriPerSpec = CreateObject("Scripting.Dictionary")

For i = 1 To resumeIndex - 1
    Set sh = ThisWorkbook.Sheets(i)
    If sh.Name <> "CATEGORY SPESIFICATION" Then
        Dim totalAll As Double
        totalAll = 0
        
        lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
        
        ' Simpan total per kategori
        Dim kategoriTotalSpec As Object
        Set kategoriTotalSpec = CreateObject("Scripting.Dictionary")
        Dim catval As Double
        Dim catKey As Variant
        
        For Row = 3 To lastRow
            If Not IsEmpty(sh.Cells(Row, "J").Value) And IsNumeric(sh.Cells(Row, "J").Value) Then
                
                
                catName = Trim(UCase(sh.Cells(Row, "H").Value))
                catval = sh.Cells(Row, "J").Value
                
                totalAll = totalAll + catval
                
                If kategoriTotalSpec.exists(catName) Then
                    kategoriTotalSpec(catName) = kategoriTotalSpec(catName) + catval
                Else
                    kategoriTotalSpec.Add catName, catval
                End If
            End If
        Next Row

        If totalAll > 0 Then
            Dim urut As Long
            Debug.Print "=== Persentase per Kategori untuk SPEC: " & sh.Name & " ==="
            For urut = LBound(orderedCategories) To UBound(orderedCategories)
                catName = orderedCategories(urut)
                Dim nilaiCat As Double: nilaiCat = 0
                
                If kategoriTotalSpec.exists(catName) Then
                    nilaiCat = kategoriTotalSpec(catName)
                End If
                
                Dim persentase As Double
                persentase = nilaiCat / totalAll
                
                Dim kunciGabung As String
                kunciGabung = catName & "|" & sh.Name
                persentaseKategoriPerSpec(kunciGabung) = persentase
                
                Debug.Print catName & "|" & sh.Name & " = " & Format(nilaiCat, "0.0000") & " / " & Format(totalAll, "0.0000") & _
                            " = " & Format(persentase, "0.00%")
            Next urut
        End If
    End If
Next i

' === RATA-RATA PERSENTASE PER KATEGORI DI SELURUH SPEC ===
Dim totalPersenPerKategori As Object
Dim countPerKategori As Object
Set totalPersenPerKategori = CreateObject("Scripting.Dictionary")
Set countPerKategori = CreateObject("Scripting.Dictionary")

Dim gabungKey As Variant
For Each gabungKey In persentaseKategoriPerSpec.Keys
    Dim parts() As String
    parts = Split(gabungKey, "|")
    Dim namaKategori As Variant
    namaKategori = parts(0)
    
    If totalPersenPerKategori.exists(namaKategori) Then
        totalPersenPerKategori(namaKategori) = totalPersenPerKategori(namaKategori) + persentaseKategoriPerSpec(gabungKey)
        countPerKategori(namaKategori) = countPerKategori(namaKategori) + 1
    Else
        totalPersenPerKategori.Add namaKategori, persentaseKategoriPerSpec(gabungKey)
        countPerKategori.Add namaKategori, 1
    End If
Next gabungKey

Debug.Print "=== AVERAGE PERSENTASE PER KATEGORI DI SELURUH SPEC ==="
Dim avgKategori As Double
For Each namaKategori In totalPersenPerKategori.Keys
    avgKategori = totalPersenPerKategori(namaKategori) / countPerKategori(namaKategori)
    Debug.Print namaKategori & ": " & Format(avgKategori, "0.00%")
Next namaKategori


Dim avgTotalSustainability As Double
avgTotalSustainability = 0


For Each namaKategori In totalPersenPerKategori.Keys
    For i = 1 To 6
        If sustainableCategoryNames(i) = namaKategori Then
            avgKategori = totalPersenPerKategori(namaKategori) / countPerKategori(namaKategori)
            avgTotalSustainability = avgTotalSustainability + avgKategori
            Exit For
        End If
    Next i
Next namaKategori

Debug.Print "TOTAL AVERAGE SUSTAINABILITY CATEGORY: " & Format(avgTotalSustainability, "0.00%")



End Sub

Sub CalculateSustainabilityAfter()
    ' Reset variabel global untuk after
    GlobalSustainabilityAfter = 0
    GlobalTotalAfter = 0
    GlobalPercentageAfter = 0
    
    ' Dictionary untuk menyimpan perhitungan per spesifikasi
    Dim specSustainabilityAfter As Object
    Dim specTotalAfter As Object
    Set specSustainabilityAfter = CreateObject("Scripting.Dictionary")
    Set specTotalAfter = CreateObject("Scripting.Dictionary")
    
    ' Nama kategori sustainability (6 kategori)
    Dim sustainableCategoryNames(1 To 6) As String
    sustainableCategoryNames(1) = "NATURAL RUBBER"
    sustainableCategoryNames(2) = "RECLAIM RUBBER"
    sustainableCategoryNames(3) = "SUSTAINABLE FILLER"
    sustainableCategoryNames(4) = "SUSTAINABLE OIL"
    sustainableCategoryNames(5) = "SUSTAINABLE CHEMICAL"
    sustainableCategoryNames(6) = "SUSTAINABLE REINFORCEMENT"
    
    Dim sh As Worksheet
    Dim resumeIndex As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' Cari index sheet RESUME
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i
    
    ' Loop semua sheet sebelum RESUME untuk menghitung per spesifikasi
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" Then
            Dim specName As String
            specName = sh.Name
            
            ' Inisialisasi dictionary untuk spesifikasi ini
            specSustainabilityAfter.Add specName, 0
            specTotalAfter.Add specName, 0
            
            lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
            
            Dim currentCategory As String
            currentCategory = ""
            
            ' Loop setiap baris di sheet
            Dim Row As Long
            For Row = 3 To lastRow
                ' Cek jika baris kategori (ada nilai di kolom J)
                If Not IsEmpty(sh.Cells(Row, "J").Value) And IsNumeric(sh.Cells(Row, "J").Value) Then
                    currentCategory = Trim(UCase(sh.Cells(Row, "H").Value))
                    
                    ' Ambil nilai total kategori dari kolom J
                    Dim nilaiKategoriJ As Double
                    nilaiKategoriJ = sh.Cells(Row, "J").Value
                    
                    ' Tambahkan ke total spesifikasi ini
                    specTotalAfter(specName) = specTotalAfter(specName) + nilaiKategoriJ
                    
                    ' Cek jika kategori ini adalah sustainability category
                    Dim catIdx As Long
                    For catIdx = 1 To 6
                        If sustainableCategoryNames(catIdx) = currentCategory Then
                            specSustainabilityAfter(specName) = specSustainabilityAfter(specName) + nilaiKategoriJ
                        End If
                    Next catIdx
                End If
            Next Row
        End If
    Next i
    
    ' Hitung total persentase sustainability per spesifikasi
    Dim totalPersentaseSustainabilityAfter As Double
    totalPersentaseSustainabilityAfter = 0

    Dim validSheetCount As Long
    validSheetCount = 0
    For i = 1 To resumeIndex - 1
        If ThisWorkbook.Sheets(i).Name <> "CATEGORY SPESIFICATION" Then
            validSheetCount = validSheetCount + 1
        End If
    Next i

    Dim daftarspec() As Variant
    ReDim daftarspec(1 To validSheetCount + 1)

    Dim specNames() As String
    ReDim specNames(1 To validSheetCount)

    Dim specKey As Variant
    Dim idx As Long
    idx = 1

    For Each specKey In specSustainabilityAfter.Keys
        If specTotalAfter(specKey) > 0 Then
            daftarspec(idx) = (specSustainabilityAfter(specKey) / specTotalAfter(specKey))
            Debug.Print daftarspec(idx)
        Else
            daftarspec(idx) = 0
        End If
        specNames(idx) = specKey                 ' Simpan nama spesifikasinya juga jika ingin tracking
        totalPersentaseSustainabilityAfter = totalPersentaseSustainabilityAfter + daftarspec(idx)
        GlobalSustainabilityAfter = GlobalSustainabilityAfter + specSustainabilityAfter(specKey)
        GlobalTotalAfter = GlobalTotalAfter + specTotalAfter(specKey)
        idx = idx + 1
    Next specKey

   
    
    ' Simpan hasil ke variabel global
    If GlobalTotalAfter > 0 Then
        GlobalPercentageAfter = totalPersentaseSustainabilityAfter
    Else
        GlobalPercentageAfter = 0
    End If
    
    Debug.Print "Sustainability After Calculation Completed"
    Debug.Print "Total Sustainability After: " & GlobalSustainabilityAfter
    Debug.Print "Total After: " & GlobalTotalAfter
    Debug.Print "Percentage After: " & GlobalPercentageAfter
    
    ' Debug detail perhitungan per spesifikasi
    Debug.Print "=== DETAIL PERHITUNGAN PER SPESIFIKASI ==="
    Debug.Print "Program: KOMBINASI"
    Debug.Print "Global Sustainability After: " & GlobalSustainabilityAfter
    Debug.Print "Global Total After: " & GlobalTotalAfter
    Debug.Print "Global Percentage After: " & GlobalPercentageAfter

    ' Loop semua sheet untuk debug detail
    'Dim sh As Worksheet
    'Dim resumeIndex As Long
    ' Dim i As Long

    ' Cari index sheet RESUME
    ' Cari index sheet RESUME
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i


   
    Dim idxSpec As Long
    idxSpec = 1

    ' Debug per spesifikasi
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" Then
            Debug.Print "--- Sheet: " & sh.Name & " ---"
        
            'Dim lastRow As Long
            lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
        
            Dim sustainabilityTotal As Double
            Dim grandTotal As Double
       
        
            sustainabilityTotal = 0
            grandTotal = 0
        
            ' Nama kategori sustainability
            ' Dim sustainableCategoryNames(1 To 6) As String
            sustainableCategoryNames(1) = "NATURAL RUBBER"
            sustainableCategoryNames(2) = "RECLAIM RUBBER"
            sustainableCategoryNames(3) = "SUSTAINABLE FILLER"
            sustainableCategoryNames(4) = "SUSTAINABLE OIL"
            sustainableCategoryNames(5) = "SUSTAINABLE CHEMICAL"
            sustainableCategoryNames(6) = "SUSTAINABLE REINFORCEMENT"
        
            For Row = 3 To lastRow
                If Not IsEmpty(sh.Cells(Row, "J").Value) And IsNumeric(sh.Cells(Row, "J").Value) Then
                    Dim categoryName As String
                    Dim categoryValue As Double
                    categoryName = Trim(UCase(sh.Cells(Row, "H").Value))
                    categoryValue = sh.Cells(Row, "J").Value
                
                    grandTotal = grandTotal + categoryValue
                
                    ' Cek sustainability category
                    ' Dim catIdx As Long
                    For catIdx = 1 To 6
                        If sustainableCategoryNames(catIdx) = categoryName Then
                            sustainabilityTotal = sustainabilityTotal + categoryValue
                            Debug.Print "    Sustainable: " & categoryName & " = " & categoryValue
                            Exit For
                        End If
                    Next catIdx
                
                    Debug.Print "    Category: " & categoryName & " = " & categoryValue
                End If
            Next Row
        
            Debug.Print "    Sustainability Total: " & sustainabilityTotal
            Debug.Print "    Grand Total: " & grandTotal
       
       
            If grandTotal > 0 Then
                
                Debug.Print "    Percentage: " & daftarspec(idxSpec)
            Else
                daftarspec(idxSpec) = 0
            End If
            
            idxSpec = idxSpec + 1
        End If
    Next i

    Debug.Print "=== END DEBUG ==="
    Debug.Print "Sample: daftarspec(4) = " & daftarspec(4)
    
    ' === RATA-RATA PER KATEGORI (ACROSS ALL SPEC SHEETS) ===
    Dim categorySum As Object
    Dim categoryCount As Object
    Set categorySum = CreateObject("Scripting.Dictionary")
    Set categoryCount = CreateObject("Scripting.Dictionary")

    ' Loop semua sheet spesifikasi
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" Then
            lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
        
            For Row = 3 To lastRow
                If Not IsEmpty(sh.Cells(Row, "J").Value) And IsNumeric(sh.Cells(Row, "J").Value) Then
                    Dim catName As String
                    Dim catValue As Double
                    catName = Trim(UCase(sh.Cells(Row, "H").Value))
                    catValue = sh.Cells(Row, "J").Value
                   
                    If categorySum.exists(catName) Then
                        Debug.Print i
                        Debug.Print catName
                        Debug.Print catValue
                        categorySum(catName) = categorySum(catName) + catValue
                        Debug.Print categorySum(catName)
                        categoryCount(catName) = categoryCount(catName) + 1
                    Else
                        categorySum.Add catName, catValue
                        categoryCount.Add catName, 1
                    End If
                End If
            Next Row
        End If
    Next i

  ' Array urutan kategori yang diinginkan (urutannya seperti laporan)
Dim orderedCategories As Variant
orderedCategories = Array( _
    "NATURAL RUBBER", _
    "SYNTHETIC RUBBER", _
    "RECLAIM RUBBER", _
    "FILLER", _
    "SUSTAINABLE FILLER", _
    "OIL", _
    "SUSTAINABLE OIL", _
    "CHEMICAL", _
    "SUSTAINABLE CHEMICAL", _
    "REINFORCEMENT", _
    "SUSTAINABLE REINFORCEMENT")

Dim persentaseKategoriPerSpec As Object
Set persentaseKategoriPerSpec = CreateObject("Scripting.Dictionary")

For i = 1 To resumeIndex - 1
    Set sh = ThisWorkbook.Sheets(i)
    If sh.Name <> "CATEGORY SPESIFICATION" Then
        Dim totalAll As Double
        totalAll = 0
        
        lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
        
        ' Simpan total per kategori
        Dim kategoriTotalSpec As Object
        Set kategoriTotalSpec = CreateObject("Scripting.Dictionary")
        Dim catval As Double
        Dim catKey As Variant
        
        For Row = 3 To lastRow
            If Not IsEmpty(sh.Cells(Row, "J").Value) And IsNumeric(sh.Cells(Row, "J").Value) Then
                
                
                catName = Trim(UCase(sh.Cells(Row, "H").Value))
                catval = sh.Cells(Row, "J").Value
                
                totalAll = totalAll + catval
                
                If kategoriTotalSpec.exists(catName) Then
                    kategoriTotalSpec(catName) = kategoriTotalSpec(catName) + catval
                Else
                    kategoriTotalSpec.Add catName, catval
                End If
            End If
        Next Row

        If totalAll > 0 Then
            Dim urut As Long
            Debug.Print "=== Persentase per Kategori untuk SPEC: " & sh.Name & " ==="
            For urut = LBound(orderedCategories) To UBound(orderedCategories)
                catName = orderedCategories(urut)
                Dim nilaiCat As Double: nilaiCat = 0
                
                If kategoriTotalSpec.exists(catName) Then
                    nilaiCat = kategoriTotalSpec(catName)
                End If
                
                Dim persentase As Double
                persentase = nilaiCat / totalAll
                
                Dim kunciGabung As String
                kunciGabung = catName & "|" & sh.Name
                persentaseKategoriPerSpec(kunciGabung) = persentase
                
                Debug.Print catName & "|" & sh.Name & " = " & Format(nilaiCat, "0.0000") & " / " & Format(totalAll, "0.0000") & _
                            " = " & Format(persentase, "0.00%")
            Next urut
        End If
    End If
Next i

' === RATA-RATA PERSENTASE PER KATEGORI DI SELURUH SPEC ===
Dim totalPersenPerKategori As Object
Dim countPerKategori As Object
Set totalPersenPerKategori = CreateObject("Scripting.Dictionary")
Set countPerKategori = CreateObject("Scripting.Dictionary")

Dim gabungKey As Variant
For Each gabungKey In persentaseKategoriPerSpec.Keys
    Dim parts() As String
    parts = Split(gabungKey, "|")
    Dim namaKategori As Variant
    namaKategori = parts(0)
    
    If totalPersenPerKategori.exists(namaKategori) Then
        totalPersenPerKategori(namaKategori) = totalPersenPerKategori(namaKategori) + persentaseKategoriPerSpec(gabungKey)
        countPerKategori(namaKategori) = countPerKategori(namaKategori) + 1
    Else
        totalPersenPerKategori.Add namaKategori, persentaseKategoriPerSpec(gabungKey)
        countPerKategori.Add namaKategori, 1
    End If
Next gabungKey

Debug.Print "=== AVERAGE PERSENTASE PER KATEGORI DI SELURUH SPEC ==="
Dim avgKategori As Double
For Each namaKategori In totalPersenPerKategori.Keys
    avgKategori = totalPersenPerKategori(namaKategori) / countPerKategori(namaKategori)
    Debug.Print namaKategori & ": " & Format(avgKategori, "0.00%")
Next namaKategori


Dim avgTotalSustainability As Double
avgTotalSustainability = 0


For Each namaKategori In totalPersenPerKategori.Keys
    For i = 1 To 6
        If sustainableCategoryNames(i) = namaKategori Then
            avgKategori = totalPersenPerKategori(namaKategori) / countPerKategori(namaKategori)
            avgTotalSustainability = avgTotalSustainability + avgKategori
            Exit For
        End If
    Next i
Next namaKategori

Debug.Print "TOTAL AVERAGE SUSTAINABILITY CATEGORY: " & Format(avgTotalSustainability, "0.00%")

 ' Gunakan daftarspec() untuk dikalikan portion()
    Dim hitungportion() As Variant
    Dim portion() As Variant, n As Long, u As Long
    portion = AmbilMassHorizontal()
    n = UBound(portion)
    u = UBound(daftarspec)
    ReDim hitungportion(1 To n)
    
 
    
    daftarspec(validSheetCount + 1) = avgTotalSustainability
   
    Debug.Print
    For i = 1 To n
        
        Debug.Print "Persentase: " & daftarspec(i)
        Debug.Print "Portion: " & portion(i)
        hitungportion(i) = daftarspec(i) * portion(i)
    Next i
    
    
    totalallportion = 0
   
    For i = 1 To UBound(hitungportion)
        totalallportion = totalallportion + hitungportion(i)
    Next i
    Debug.Print "hasil " & totalallportion
    
    Dim WSsim2 As Worksheet
    Set WSsim2 = ThisWorkbook.Sheets("SIMULATION_PROCESS")
    
    WSsim2.Cells.Range("E28").Value = totalallportion


End Sub

Sub DisplaySustainabilityResults()
    ' Menampilkan hasil sustainability ke worksheet atau konsol
    Debug.Print "=== SUSTAINABILITY CALCULATION RESULTS ==="
    Debug.Print "Before Replacement:"
    Debug.Print "  Total Sustainability: " & GlobalSustainabilityBefore
    Debug.Print "  Total Weight: " & GlobalTotalBefore
    Debug.Print "  Percentage: " & Format(GlobalPercentageBefore, "0.00%")
    Debug.Print ""
    Debug.Print "After Replacement:"
    Debug.Print "  Total Sustainability: " & GlobalSustainabilityAfter
    Debug.Print "  Total Weight: " & GlobalTotalAfter
    Debug.Print "  Percentage: " & Format(GlobalPercentageAfter, "0.00%")
    Debug.Print ""
    Debug.Print "Difference:"
    Debug.Print "  Sustainability Change: " & (GlobalSustainabilityAfter - GlobalSustainabilityBefore)
    Debug.Print "  Percentage Change: " & Format(GlobalPercentageAfter - GlobalPercentageBefore, "0.00%")
    Debug.Print "================================================"
    
    ' Opsional: Simpan ke worksheet tertentu (misalnya CALCULATE sheet)
    Dim wsCalc As Worksheet
    Set wsCalc = ThisWorkbook.Sheets("CALCULATE")
    
    ' Simpan hasil ke cell tertentu (sesuaikan dengan kebutuhan)
    wsCalc.Range("B35").Value = "Sustainability Before:"
    wsCalc.Range("C35").Value = GlobalPercentageBefore
    wsCalc.Range("B36").Value = "Sustainability After:"
    wsCalc.Range("C36").Value = GlobalPercentageAfter
    wsCalc.Range("B37").Value = "Change:"
    wsCalc.Range("C37").Value = GlobalPercentageAfter - GlobalPercentageBefore
End Sub
Sub RestoreMaterialValues()
    On Error GoTo ErrorHandler

    If MaterialBackupData Is Nothing Then Exit Sub

    Dim key As Variant, backupInfo As Variant, ws As Worksheet
    Dim arrDel As Collection
    Set arrDel = New Collection

    ' 1. Kumpulkan semua info row material baru (ADDED)
    For Each key In MaterialBackupData.Keys
        backupInfo = MaterialBackupData(key)
        Dim parts As Variant: parts = Split(backupInfo, "|")
        If UBound(parts) >= 5 Then
            If parts(5) = "ADDED" Then
                arrDel.Add Array(parts(0), parts(4)) ' sheet, nama material
            End If
        End If
    Next key

    ' 2. Hapus semua row material baru berdasarkan nama, bukan row index!
    Dim i As Long
    For i = 1 To arrDel.Count
        Set ws = ThisWorkbook.Sheets(arrDel(i)(0))
        Dim nm As String: nm = arrDel(i)(1)
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        Dim rr As Long
        For rr = lastRow To 3 Step -1            ' dari bawah ke atas
            If Trim(UCase(ws.Cells(rr, "H").Value)) = Trim(UCase(nm)) Then
                ws.Rows(rr).Delete Shift:=xlUp
                Exit For
            End If
        Next rr
    Next i

    ' 3. Restore nilai MODIFIED berdasar nama material, bukan row index!
    For Each key In MaterialBackupData.Keys
        backupInfo = MaterialBackupData(key)
        Dim parts2 As Variant: parts2 = Split(backupInfo, "|")
        If UBound(parts2) >= 5 Then
            ' Di bagian restore:
If parts2(5) = "MODIFIED" Then
    Set ws = ThisWorkbook.Sheets(parts2(0))
    Dim nmMod As String: nmMod = parts2(4)
    Dim colLetter As String: colLetter = parts2(2)
    Dim valRestore As Variant: valRestore = parts2(3)
    Dim lastRow2 As Long: lastRow2 = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    Dim r2 As Long
    For r2 = 3 To lastRow2
        If Trim(UCase(ws.Cells(r2, "H").Value)) = Trim(UCase(nmMod)) Then
            Dim fixVal As Variant
            fixVal = valRestore
            If InStr(CStr(fixVal), ",") > 0 And InStr(CStr(fixVal), ".") = 0 Then
                fixVal = Replace(CStr(fixVal), ",", ".")
            End If
            If IsNumeric(fixVal) Then
                ws.Cells(r2, colLetter).Value = val(fixVal)
            Else
                ws.Cells(r2, colLetter).Value = fixVal
            End If
            ws.Cells(r2, colLetter).NumberFormat = "General"
            Exit For
        End If
    Next r2
End If


        End If
    Next key

    Set MaterialBackupData = Nothing
    Exit Sub
    
    Call KonversiKolomIkeNumberSemuaSheet
ErrorHandler:
    MsgBox "Restore ERROR: " & Err.Description, vbCritical
End Sub
Sub KonversiKolomIkeNumberSemuaSheet()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim i As Long, lastSheet As Long
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, val As Variant, newVal As Variant

    ' Cari posisi sheet "RESUME"
    lastSheet = 0
    For i = 1 To wb.Sheets.Count
        If wb.Sheets(i).Name = "RESUME" Then
            lastSheet = i
            Exit For
        End If
    Next i

    If lastSheet = 0 Then
        MsgBox """RESUME"" sheet tidak ditemukan!", vbCritical
        Exit Sub
    End If

    For i = 1 To lastSheet - 1
        Set ws = wb.Sheets(i)
        If ws.Name <> "CATEGORY SPESIFICATION" Then
            lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
            For r = 3 To lastRow
                val = ws.Cells(r, "I").Value
                If VarType(val) = vbString Then
                    val = Trim(val)
                    If val <> "" Then
                        ' Jika format Indo/Eropa: koma desimal tanpa titik
                        If InStr(val, ",") > 0 And InStr(val, ".") = 0 Then
                            val = Replace(val, ",", ".")
                        End If
                        ' Hanya parse bila hasilnya valid dan tidak integer besar
                        If IsNumeric(val) Then
                            newVal = val(val) ' <-- ini penting, Val lebih aman dari CDbl!
                            ws.Cells(r, "I").Value = newVal
                        End If
                    End If
                End If
            Next r
            ws.Range(ws.Cells(3, "I"), ws.Cells(lastRow, "I")).NumberFormat = "General"
        End If
    Next i
End Sub






' Fungsi untuk mendapatkan nilai sustainability (dapat dipanggil dari modul lain)
Function GetSustainabilityBefore() As Double
    GetSustainabilityBefore = GlobalPercentageBefore
End Function

Function GetSustainabilityAfter() As Double
    GetSustainabilityAfter = GlobalPercentageAfter
End Function

Function GetSustainabilityChange() As Double
    GetSustainabilityChange = GlobalPercentageAfter - GlobalPercentageBefore
End Function
Sub update_multiple_material_data_with_backup(replacements() As replacementData)
    Dim simpan_data As Double
    Dim nilai_pengganti As Double
    Dim found As Boolean
    Dim resumeIndex As Long
    Dim specSheets As Collection
    Dim sh As Worksheet
    Dim i As Long, j As Long
    
    Set specSheets = New Collection
    found = False
    
    ' Initialize backup dictionary if not already done
    If MaterialBackupData Is Nothing Then
        Set MaterialBackupData = CreateObject("Scripting.Dictionary")
    End If
    
    ' Cari posisi RESUME sheet
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i
    
    ' Kumpulkan spec sheets
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" Then
            specSheets.Add sh.Name
        End If
    Next i
    
    Dim lastHistRow As Long
    Dim ws As Worksheet
    Dim wsName As Variant
    
    ' Process setiap sheet
    For Each wsName In specSheets
        Set ws = ThisWorkbook.Sheets(wsName)
        
        ' Dictionary untuk akumulasi material baru per sheet
        Dim newMaterialAccumulation As Object
        Set newMaterialAccumulation = CreateObject("Scripting.Dictionary")
        
        ' Collection untuk menyimpan info material yang akan diproses
        Dim materialToProcess As Collection
        Set materialToProcess = New Collection
        
        ' STEP 1: SCAN DAN KUMPULKAN DATA TANPA MENGUBAH APAPUN DULU
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        ' Scan untuk material yang akan diganti
        For i = 3 To lastRow
            Dim cellValue As String
            cellValue = Trim(UCase(ws.Cells(i, "H").Value))
            
            For j = 1 To 3
                If replacements(j).isValid Then
                    If cellValue = UCase(replacements(j).material_replaced) Then
                        If IsNumeric(ws.Cells(i, "I").Value) Then
                            ' Simpan info untuk diproses nanti
                            Dim materialInfo As Object
                            Set materialInfo = CreateObject("Scripting.Dictionary")
                            materialInfo("row") = i
                            materialInfo("replacementIndex") = j
                            materialInfo("originalValue") = ws.Cells(i, "I").Value
                            materialInfo("materialName") = replacements(j).material_replaced
                            materialToProcess.Add materialInfo
                        End If
                    End If
                End If
            Next j
        Next i
        
        ' STEP 2: PROSES MATERIAL REPLACEMENT (DARI BAWAH KE ATAS UNTUK AVOID INDEX SHIFT)
        Dim k As Long
        For k = materialToProcess.Count To 1 Step -1
            Dim matInfo As Object
            Set matInfo = materialToProcess(k)
            
            Dim rowNum As Long: rowNum = matInfo("row")
            Dim repIdx As Long: repIdx = matInfo("replacementIndex")
            Dim origValue As Double: origValue = matInfo("originalValue")
            
            ' Backup original value
            Dim backupKey As String
            backupKey = ws.Name & "_" & rowNum & "_I_" & matInfo("materialName") & "_" & k
            Dim backupValue As String
            backupValue = ws.Name & "|" & rowNum & "|I|" & origValue & "|" & matInfo("materialName") & "|MODIFIED"
            
            If Not MaterialBackupData.exists(backupKey) Then
                MaterialBackupData.Add backupKey, backupValue
            End If
            
            ' Calculate new value
            nilai_pengganti = 1 - replacements(repIdx).percentage_material
            Dim newValue As Double: newValue = origValue * nilai_pengganti
            
            ' Update cell
            ws.Cells(rowNum, "I").Value = newValue
            found = True
            
            ' Store before/after values
            MaterialOldBefore(repIdx) = origValue
            MaterialOldAfter(repIdx) = newValue
            
            ' Accumulate new material
            Dim addedAmount As Double
            addedAmount = replacements(repIdx).percentage_new_material * origValue
            
            Dim newMatName As String: newMatName = replacements(repIdx).new_material
            If newMaterialAccumulation.exists(newMatName) Then
                newMaterialAccumulation(newMatName) = newMaterialAccumulation(newMatName) + addedAmount
            Else
                newMaterialAccumulation.Add newMatName, addedAmount
                ' Store class info
                newMaterialAccumulation.Add newMatName & "_CLASS", replacements(repIdx).new_material_class
            End If
        Next k
        
        ' STEP 3: PROSES NEW MATERIALS (SETELAH SEMUA REPLACEMENT SELESAI)
        ' Refresh lastRow setelah semua modifikasi
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        Dim materialKeys As Collection
        Set materialKeys = New Collection
        
        ' Collect material keys (exclude _CLASS keys)
        Dim matKey As Variant
        For Each matKey In newMaterialAccumulation.Keys
            If Right(matKey, 6) <> "_CLASS" Then
                materialKeys.Add matKey
            End If
        Next matKey
        
        ' Process new materials
        Dim keyIndex As Long
        For keyIndex = 1 To materialKeys.Count
            Dim materialName As String: materialName = materialKeys(keyIndex)
            Dim totalAmount As Double: totalAmount = newMaterialAccumulation(materialName)
            Dim materialClass As String: materialClass = newMaterialAccumulation(materialName & "_CLASS")
            
            ' Cari apakah material sudah ada
            Dim materialFoundRow As Long: materialFoundRow = 0
            Dim materialExists As Boolean: materialExists = False
            
            ' Re-scan current sheet state
            lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
            
            For i = 3 To lastRow
                If Trim(UCase(ws.Cells(i, "H").Value)) = UCase(materialName) Then
                    If IsNumeric(ws.Cells(i, "I").Value) Then
                        materialFoundRow = i
                        materialExists = True
                        Exit For
                    End If
                End If
            Next i
            
            If materialExists Then
                ' Material exists - update value
                Dim oldExistingValue As Double: oldExistingValue = ws.Cells(materialFoundRow, "I").Value
                
                ' Backup existing value
                Dim existingBackupKey As String
                existingBackupKey = ws.Name & "_" & materialFoundRow & "_I_" & materialName & "_EXISTING_" & keyIndex
                Dim existingBackupValue As String
                existingBackupValue = ws.Name & "|" & materialFoundRow & "|I|" & oldExistingValue & "|" & materialName & "|MODIFIED"
                
                If Not MaterialBackupData.exists(existingBackupKey) Then
                    MaterialBackupData.Add existingBackupKey, existingBackupValue
                End If
                
                ws.Cells(materialFoundRow, "I").Value = oldExistingValue + totalAmount
                
                ' Store new material before/after values
                Dim idxNew1 As Long
                For idxNew1 = 1 To 3
                    If replacements(idxNew1).isValid And replacements(idxNew1).new_material = materialName Then
                        MaterialNewBefore(idxNew1) = oldExistingValue
                        MaterialNewAfter(idxNew1) = oldExistingValue + totalAmount
                        Exit For
                    End If
                Next idxNew1
                
            Else
                ' Material doesn't exist - find class and add new material
                Dim classFoundRow As Long: classFoundRow = 0
                
                ' Re-scan for class
                lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
                
                For i = 3 To lastRow
                    If Trim(UCase(ws.Cells(i, "H").Value)) = UCase(materialClass) Then
                        If IsNumeric(ws.Cells(i, "J").Value) Then ' Class indicator
                            classFoundRow = i
                            Exit For
                        End If
                    End If
                Next i
                
                If classFoundRow > 0 Then
                    ' Find correct insert position
                    Dim insertRow As Long: insertRow = classFoundRow + 1
                    
                    ' Scan untuk posisi insert yang tepat
                    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
                    
                    For i = classFoundRow + 1 To lastRow + 1
                        If i > lastRow Then
                            insertRow = i
                            Exit For
                        End If
                        
                        Dim nextVal As String: nextVal = Trim(ws.Cells(i, "H").Value)
                        If nextVal = "" Then
                            insertRow = i
                            Exit For
                        ElseIf IsNumeric(ws.Cells(i, "J").Value) Then
                            ' Found another class, insert here
                            insertRow = i
                            Exit For
                        End If
                    Next i
                    
                    ' Insert new row
                    If insertRow <= lastRow Then
                        ws.Rows(insertRow).Insert Shift:=xlDown
                    End If
                    
                    ' Clear formatting and add data
                    ws.Rows(insertRow).Font.Bold = False
                    ws.Cells(insertRow, "H").Value = materialName
                    ws.Cells(insertRow, "I").Value = totalAmount
                    
                    ' Backup for added row
                    Dim addedBackupKey As String
                    addedBackupKey = ws.Name & "_" & insertRow & "_ADDED_" & materialName & "_" & keyIndex
                    Dim addedBackupValue As String
                    addedBackupValue = ws.Name & "|" & insertRow & "|I|0|" & materialName & "|ADDED"
                    
                    If Not MaterialBackupData.exists(addedBackupKey) Then
                        MaterialBackupData.Add addedBackupKey, addedBackupValue
                    End If
                    
                    ' Store new material values
                    Dim idxNew2 As Long
                    For idxNew2 = 1 To 3
                        If replacements(idxNew2).isValid And replacements(idxNew2).new_material = materialName Then
                            MaterialNewBefore(idxNew2) = 0
                            MaterialNewAfter(idxNew2) = totalAmount
                            Exit For
                        End If
                    Next idxNew2
                    
                Else
                    MsgBox "Class '" & materialClass & "' tidak ditemukan di sheet " & wsName, vbExclamation
                End If
            End If
        Next keyIndex
        
        ' Clean up objects
        Set newMaterialAccumulation = Nothing
        Set materialToProcess = Nothing
        Set materialKeys = Nothing
        
    Next wsName
    
    If found Then
        Debug.Print "Material update completed with backup. Total backup items: " & MaterialBackupData.Count
        MsgBox "Update material berhasil! Total backup: " & MaterialBackupData.Count, vbInformation
    Else
        MsgBox "Tidak ada material yang ditemukan untuk diganti.", vbExclamation
    End If
End Sub



Sub RunFullSimulationAndSubmit()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Backup input (C5:9, G5:9, K5:9)
    Dim backupInputs(0 To 2, 0 To 4) As Variant
    Dim inputCols As Variant
    inputCols = Array(3, 7, 11)

    Dim colIdx As Integer, rowOffset As Integer
    For colIdx = 0 To 2
        For rowOffset = 0 To 4
            backupInputs(colIdx, rowOffset) = ws.Cells(5 + rowOffset, inputCols(colIdx)).Value
        Next rowOffset
    Next colIdx

    ' ============================
    ' STEP 1: JALANKAN SIMULASI
    ' ============================
    Call SimulateStepwiseReplacement_WithBackup(backupInputs)

    ' ============================
    ' STEP 2: KEMBALIKAN INPUT
    ' ============================
    For colIdx = 0 To 2
        For rowOffset = 0 To 4
            ws.Cells(5 + rowOffset, inputCols(colIdx)).Value = backupInputs(colIdx, rowOffset)
        Next rowOffset
    Next colIdx

    ' ============================
    ' STEP 3: JALANKAN PROSES UTAMA
    ' ============================
    Call submit_multiple_replacement3

  '  MsgBox "Simulasi + submit selesai.", vbInformation
End Sub

Sub prepare_partial_replacement(i As Integer, ByRef backupInputs As Variant)
    ' Restore semua input dari backup
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    Dim inputCols As Variant
    inputCols = Array(3, 7, 11)                  ' C=3, G=7, K=11

    Dim colIdx As Integer, rowOffset As Integer
    For colIdx = 0 To 2
        For rowOffset = 0 To 4
            ws.Cells(5 + rowOffset, inputCols(colIdx)).Value = backupInputs(colIdx, rowOffset)
        Next rowOffset
    Next colIdx

    ' Hapus kolom input yang lebih dari i material
    For colIdx = i To 2
        For rowOffset = 0 To 4
            ws.Cells(5 + rowOffset, inputCols(colIdx)).Value = ""
        Next rowOffset
    Next colIdx
End Sub

Sub SimulateStepwiseReplacement_WithBackup(ByRef backupInputs As Variant)

    Dim simResults(1 To 3) As Double
    Dim portionkombinasi(1 To 3) As Double
    Dim portionkombinasi2(1 To 3) As Double
    
    Dim i As Integer, x As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Tentukan jumlah material yang valid (x)
    If Trim(ws.Range("C5").Value) <> "" And Trim(ws.Range("G5").Value) = "" And Trim(ws.Range("K5").Value) = "" Then
        x = 1
    ElseIf Trim(ws.Range("C5").Value) <> "" And Trim(ws.Range("G5").Value) <> "" And Trim(ws.Range("K5").Value) = "" Then
        x = 2
    ElseIf Trim(ws.Range("C5").Value) <> "" And Trim(ws.Range("G5").Value) <> "" And Trim(ws.Range("K5").Value) <> "" Then
        x = 3
    Else
        MsgBox "Minimal satu set material (di kolom C) harus diisi untuk simulasi.", vbExclamation
        Exit Sub
    End If
Debug.Print totalallportion
    ' Jalankan simulasi stepwise
    For i = 1 To x
        Call reset_globaldata
        Call prepare_partial_replacement(i, backupInputs)
        Call submit_multiple_replacement3
        simResults(i) = GlobalPercentageAfter
        portionkombinasi(i) = totalallportion
        portionkombinasi2(i) = totalallportionBefore
        Debug.Print portionkombinasi2(i)
    Next i

    ' Tampilkan hasil ke SIMULATION_PROCESS
    Dim WSsim As Worksheet
    Set WSsim = ThisWorkbook.Sheets("SIMULATION_PROCESS")
    WSsim.Range("F8").Value = IIf(x >= 1, simResults(1), "")
    WSsim.Range("F9").Value = IIf(x >= 2, simResults(2), "")
    WSsim.Range("F10").Value = IIf(x >= 3, simResults(3), "")
    
    WSsim.Range("G8").Value = IIf(x >= 1, portionkombinasi2(1), "")
    WSsim.Range("G9").Value = IIf(x >= 2, portionkombinasi2(2), "")
    WSsim.Range("G10").Value = IIf(x >= 3, portionkombinasi2(3), "")
    
    WSsim.Range("H8").Value = IIf(x >= 1, portionkombinasi(1), "")
    WSsim.Range("H9").Value = IIf(x >= 2, portionkombinasi(2), "")
    WSsim.Range("H10").Value = IIf(x >= 3, portionkombinasi(3), "")

    ' Pesan hasil
    Dim resultMsg As String
    resultMsg = "Simulasi stepwise selesai!" & vbCrLf
    If x >= 1 Then resultMsg = resultMsg & "Material 1: " & Format(simResults(1), "0.00%") & vbCrLf
    If x >= 2 Then resultMsg = resultMsg & "Material 1+2: " & Format(simResults(2), "0.00%") & vbCrLf
    If x >= 3 Then resultMsg = resultMsg & "Material 1+2+3: " & Format(simResults(3), "0.00%")

    MsgBox resultMsg, vbInformation
End Sub

Sub reset_globaldata()
    GlobalSustainabilityBefore = 0
    GlobalSustainabilityAfter = 0
    GlobalTotalBefore = 0
    GlobalTotalAfter = 0
    GlobalPercentageBefore = 0
    GlobalPercentageAfter = 0
End Sub


