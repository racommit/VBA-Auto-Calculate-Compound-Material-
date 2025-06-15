Attribute VB_Name = "SimulasiBeforeAfter"
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

' material count
Dim MaterialOldBefore(1 To 3) As Double
Dim MaterialOldAfter(1 To 3) As Double
Dim MaterialNewBefore(1 To 3) As Double
Dim MaterialNewAfter(1 To 3) As Double


' Dictionary untuk menyimpan backup data material
Dim MaterialBackupData As Object

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

Sub submit_multiple_replacement2()
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
    
    
    Dim kelasLama As String, kelasBaru As String


For i = 1 To 3
    If replacements(i).isValid Then
        kelasLama = CariKategoriKelas(replacements(i).material_replaced)
        kelasBaru = UCase(Trim(replacements(i).new_material_class))
        
        DebugLog kelasLama
        DebugLog kelasBaru
        
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
    
    For i = 1 To 3
    WSsim.Cells(64 + i, "E").Value = MaterialOldBefore(i) ' E65–E67
    WSsim.Cells(64 + i, "F").Value = MaterialOldAfter(i)

    WSsim.Cells(70 + i, "E").Value = MaterialNewBefore(i) ' E71–E73
    WSsim.Cells(70 + i, "F").Value = MaterialNewAfter(i)
    Next i


    
     
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
    WSsim.Range("D28").Value = GlobalPercentageAfter ' Total Sustainability After
    
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
    
    DebugLog UBound(daftarspecBefore)
    DebugLog UBound(portion)
    
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

    Dim totalallportionBefore As Double
    totalallportionBefore = 0
    For i = 1 To UBound(hitungportionBefore)
        totalallportionBefore = totalallportionBefore + hitungportionBefore(i)
    Next i

    DebugLog "TOTAL Sustainability Before (tertimbang): " & totalallportionBefore
    ThisWorkbook.Sheets("SIMULATION_PROCESS").Range("E27").Value = totalallportionBefore
    If GlobalTotalBefore > 0 Then
        GlobalPercentageBefore = totalPersentaseSustainabilityBefore
    Else
        GlobalPercentageBefore = 0
        
    End If
    
    DebugLog "Sustainability Before Calculation Completed"
    DebugLog "Total Sustainability Before: " & GlobalSustainabilityBefore
    DebugLog "Total Before: " & GlobalTotalBefore
    DebugLog "Percentage Before: " & GlobalPercentageBefore
    
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
                        DebugLog i
                        DebugLog catName
                        DebugLog catValue
                        categorySum(catName) = categorySum(catName) + catValue
                        DebugLog categorySum(catName)
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
            DebugLog "=== Persentase per Kategori untuk SPEC: " & sh.Name & " ==="
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
                
                DebugLog catName & "|" & sh.Name & " = " & Format(nilaiCat, "0.0000") & " / " & Format(totalAll, "0.0000") & _
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

DebugLog "=== AVERAGE PERSENTASE PER KATEGORI DI SELURUH SPEC ==="
Dim avgKategori As Double
For Each namaKategori In totalPersenPerKategori.Keys
    avgKategori = totalPersenPerKategori(namaKategori) / countPerKategori(namaKategori)
    DebugLog namaKategori & ": " & Format(avgKategori, "0.00%")
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

DebugLog "TOTAL AVERAGE SUSTAINABILITY CATEGORY: " & Format(avgTotalSustainability, "0.00%")



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
            DebugLog daftarspec(idx)
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
    
    DebugLog "Sustainability After Calculation Completed"
    DebugLog "Total Sustainability After: " & GlobalSustainabilityAfter
    DebugLog "Total After: " & GlobalTotalAfter
    DebugLog "Percentage After: " & GlobalPercentageAfter
    
    ' Debug detail perhitungan per spesifikasi
    DebugLog "=== DETAIL PERHITUNGAN PER SPESIFIKASI ==="
    DebugLog "Program: KOMBINASI"
    DebugLog "Global Sustainability After: " & GlobalSustainabilityAfter
    DebugLog "Global Total After: " & GlobalTotalAfter
    DebugLog "Global Percentage After: " & GlobalPercentageAfter

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
            DebugLog "--- Sheet: " & sh.Name & " ---"
        
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
                            DebugLog "    Sustainable: " & categoryName & " = " & categoryValue
                            Exit For
                        End If
                    Next catIdx
                
                    DebugLog "    Category: " & categoryName & " = " & categoryValue
                End If
            Next Row
        
            DebugLog "    Sustainability Total: " & sustainabilityTotal
            DebugLog "    Grand Total: " & grandTotal
       
       
            If grandTotal > 0 Then
                
                DebugLog "    Percentage: " & daftarspec(idxSpec)
            Else
                daftarspec(idxSpec) = 0
            End If
            
            idxSpec = idxSpec + 1
        End If
    Next i

    DebugLog "=== END DEBUG ==="
    DebugLog "Sample: daftarspec(4) = " & daftarspec(4)
    
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
                        DebugLog i
                        DebugLog catName
                        DebugLog catValue
                        categorySum(catName) = categorySum(catName) + catValue
                        DebugLog categorySum(catName)
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
            DebugLog "=== Persentase per Kategori untuk SPEC: " & sh.Name & " ==="
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
                
                DebugLog catName & "|" & sh.Name & " = " & Format(nilaiCat, "0.0000") & " / " & Format(totalAll, "0.0000") & _
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

DebugLog "=== AVERAGE PERSENTASE PER KATEGORI DI SELURUH SPEC ==="
Dim avgKategori As Double
For Each namaKategori In totalPersenPerKategori.Keys
    avgKategori = totalPersenPerKategori(namaKategori) / countPerKategori(namaKategori)
    DebugLog namaKategori & ": " & Format(avgKategori, "0.00%")
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

DebugLog "TOTAL AVERAGE SUSTAINABILITY CATEGORY: " & Format(avgTotalSustainability, "0.00%")

 ' Gunakan daftarspec() untuk dikalikan portion()
    Dim hitungportion() As Variant
    Dim portion() As Variant, n As Long, u As Long
    portion = AmbilMassHorizontal()
    n = UBound(portion)
    u = UBound(daftarspec)
    ReDim hitungportion(1 To n)
    
 
    
    daftarspec(validSheetCount + 1) = avgTotalSustainability
   
    DebugLog
    For i = 1 To n
        
        DebugLog "Persentase: " & daftarspec(i)
        DebugLog "Portion: " & portion(i)
        hitungportion(i) = daftarspec(i) * portion(i)
    Next i
    
    Dim totalallportion As Double
    totalallportion = 0
   
    For i = 1 To UBound(hitungportion)
        totalallportion = totalallportion + hitungportion(i)
    Next i
    DebugLog "hasil " & totalallportion
    
    Dim WSsim2 As Worksheet
    Set WSsim2 = ThisWorkbook.Sheets("SIMULATION_PROCESS")
    
    WSsim2.Cells.Range("E28").Value = totalallportion


End Sub

Sub DisplaySustainabilityResults()
    ' Menampilkan hasil sustainability ke worksheet atau konsol
    DebugLog "=== SUSTAINABILITY CALCULATION RESULTS ==="
    DebugLog "Before Replacement:"
    DebugLog "  Total Sustainability: " & GlobalSustainabilityBefore
    DebugLog "  Total Weight: " & GlobalTotalBefore
    DebugLog "  Percentage: " & Format(GlobalPercentageBefore, "0.00%")
    DebugLog ""
    DebugLog "After Replacement:"
    DebugLog "  Total Sustainability: " & GlobalSustainabilityAfter
    DebugLog "  Total Weight: " & GlobalTotalAfter
    DebugLog "  Percentage: " & Format(GlobalPercentageAfter, "0.00%")
    DebugLog ""
    DebugLog "Difference:"
    DebugLog "  Sustainability Change: " & (GlobalSustainabilityAfter - GlobalSustainabilityBefore)
    DebugLog "  Percentage Change: " & Format(GlobalPercentageAfter - GlobalPercentageBefore, "0.00%")
    DebugLog "================================================"
    
    ' Opsional: Simpan ke worksheet tertentu (misalnya CALCULATE sheet)
    Dim wsCalc As Worksheet
    Set wsCalc = ThisWorkbook.Sheets("CALCULATE")
    
    ' Simpan hasil ke cell tertentu (sesuaikan dengan kebutuhan)
    wsCalc.Range("B35").Value = "Sustainability Now:"
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
        For rr = lastRow To 3 Step -1 ' dari bawah ke atas
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
        DebugLog "Material update completed with backup. Total backup items: " & MaterialBackupData.Count
        ShowInfo "Update material berhasil! Total backup: " & MaterialBackupData.Count
    Else
        MsgBox "Tidak ada material yang ditemukan untuk diganti.", vbExclamation
    End If
End Sub

Sub HitungWeightedSustainabilityAfter()
    Dim portion() As Variant
    Dim susAfter() As Variant
    Dim hasilGabungan() As Double
    Dim i As Long, n As Long

    portion = AmbilMassHorizontal()
    n = UBound(portion)
    For i = 1 To n - 1
        DebugLog portion(i)
    Next i
     
End Sub

Function AmbilMassHorizontal() As Variant

    Dim ws As Worksheet
    Dim sel As Range
    Dim barisMaterial As Long, barisMass As Long
    Dim i As Long
    Dim kolStart As Long, kol As Long

    Dim daftarspec() As String
    Dim daftarMass() As Variant
    Dim portion() As Variant
    Set ws = ThisWorkbook.Sheets("RESUME")

    ' Cari baris "Material"
    For Each sel In ws.UsedRange
        If LCase(Trim(sel.Value)) = "material" Then
            barisMaterial = sel.Row
            Exit For
        End If
    Next sel

    If barisMaterial = 0 Then
        MsgBox "Material' tidak ditemukan", vbCritical
        Exit Function
    End If

    ' Cari baris terakhir dari "Total (NWT) Production Tires"
    For Each sel In ws.UsedRange
        If Trim(sel.Value) = "Total (NWT) Production Tires" Then
            barisMass = sel.Row
        End If
    Next sel

    If barisMass = 0 Then
        MsgBox "Total (NWT) Production Tires Tidak Ditemukan", vbCritical
        Exit Function
    End If

    ' Kolom awal (diasumsikan mulai dari kolom 3)
    kolStart = 3
    kol = kolStart

    ' Looping untuk hitung jumlah kolom valid (sampai kolom kosong atau "sisa nwt")
    Do While True
        Dim val As String
        val = LCase(Trim(ws.Cells(barisMass, kol).Value))
        
        If val = "" Or val = "sisa nwt" Then Exit Do
        
        kol = kol + 1
    Loop

    ' Inisialisasi array
    ReDim daftarspec(1 To kol - kolStart)
    ReDim daftarMass(1 To kol - kolStart)
    ReDim portion(1 To kol - kolStart)

    ' Ambil data spec dan sustainability dari baris yang sesuai
    For i = 1 To kol - kolStart
        daftarspec(i) = ws.Cells(barisMaterial, kolStart + i - 1).Value
        If IsNumeric(ws.Cells(barisMass, kolStart + i - 1).Value) Then
            daftarMass(i) = CDbl(ws.Cells(barisMass, kolStart + i - 1).Value)
        Else
            daftarMass(i) = 0
        End If
    Next i

    'hitung total Nwt
    Dim totalMass As Double
    totalMass = 0
    For i = 1 To UBound(daftarMass)
        totalMass = totalMass + daftarMass(i)
    Next i

    For i = 1 To UBound(daftarMass)
        If totalMass > 0 Then
            portion(i) = daftarMass(i) / totalMass
            DebugLog portion(i)
        Else
            portion(i) = 0
        End If
    Next i

    AmbilMassHorizontal = portion
    
   
End Function


