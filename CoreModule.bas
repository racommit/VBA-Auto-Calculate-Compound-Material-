Attribute VB_Name = "CoreModule"
'@NoIndent
'@IgnoreModule AssignmentNotUsed
'@Folder "Core_calculate"
Option Explicit
'@Ignore ModuleScopeDimKeyword
Dim GlobalActionID As String

' Variabel global untuk menyimpan perhitungan sustainability
Dim GlobalSustainabilityBefore3 As Double
Dim GlobalSustainabilityAfter3 As Double
Dim GlobalTotalBefore3 As Double
Dim GlobalTotalAfter3 As Double
Dim GlobalPercentageBefore3 As Double
Dim GlobalPercentageAfter3 As Double

' Structure untuk menyimpan data replacement
Type replacementData
    material_replaced As String
    percentage_material As Double
    new_material As String
    percentage_new_material As Double
    new_material_class As String
    isValid As Boolean
End Type

Sub submit_multiple_replacement()
    Dim replacements(1 To 3) As replacementData
    Dim i As Long
    Dim validCount As Long
    
    ' Input ranges: C5:C9, G5:G9, K5:K9
    Dim inputColumns As Variant
    inputColumns = Array(3, 7, 11)               ' C=3, G=7, K=11
    
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
    
    ' Copy sheet RESUME
    Call CopySheetByValue("Multiple_" & GlobalActionID)
    
    ' Process semua replacement
    Call update_multiple_material_data(replacements)
    
    ' Hitung sustainability AFTER replacement
    Call CalculateSustainabilityAfter
    
    ' Tampilkan hasil sustainability
    Call DisplaySustainabilityResults
    
    ShowInfo "Proses " & validCount & " material replacement berhasil (Action ID: " & GlobalActionID & ")" & vbCrLf & _
           "Sustainability Before: " & Format(GlobalPercentageBefore3, "0.00%") & vbCrLf & _
           "Sustainability After: " & Format(GlobalPercentageAfter3, "0.00%")
           
    ' Catat ke HISTORY_CHANGE
    Dim lastActionID As String
    Dim lastRow As Long
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets("HISTORY_UNDO")
    
    
    
    lastActionID = Now()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HISTORY_CHANGE")
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If ws.Cells(2, 1).Value = "" Then
        ws.Cells(2, 1).Value = 1
        nextRow = 2
    Else
        ws.Cells(nextRow, 1).Value = ws.Cells(nextRow - 1, 1).Value + 1
    End If
        
    ws.Cells(nextRow, 2).Value = Now
    ws.Cells(nextRow, 3).Value = Sheets("CALCULATE").Range("B33").Value
    ws.Cells(nextRow, 4).Value = "Material Replacement (Action ID: " & lastActionID & ")"
End Sub

Sub CalculateSustainabilityBefore()
    ' Reset variabel global
    GlobalSustainabilityBefore3 = 0
    GlobalTotalBefore3 = 0
    GlobalPercentageBefore3 = 0
    
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
    
    ' Hitung total persentase sustainability per spesifikasi
    Dim totalPersentaseSustainabilityBefore As Double
    totalPersentaseSustainabilityBefore = 0
    
    Dim specKey As Variant
    For Each specKey In specSustainabilityBefore.Keys
        Dim persentaseSpecBefore As Double
        persentaseSpecBefore = 0
        
        If specTotalBefore(specKey) > 0 Then
            persentaseSpecBefore = (specSustainabilityBefore(specKey) / specTotalBefore(specKey))
        End If
        
        totalPersentaseSustainabilityBefore = totalPersentaseSustainabilityBefore + persentaseSpecBefore
        GlobalSustainabilityBefore3 = GlobalSustainabilityBefore3 + specSustainabilityBefore(specKey)
        GlobalTotalBefore3 = GlobalTotalBefore3 + specTotalBefore(specKey)
    Next specKey
    
    ' Simpan hasil ke variabel global
    If GlobalTotalBefore3 > 0 Then
        GlobalPercentageBefore3 = totalPersentaseSustainabilityBefore
    Else
        GlobalPercentageBefore3 = 0
    End If
    
    DebugLog "Sustainability Before Calculation Completed"
    DebugLog "Total Sustainability Before: " & GlobalSustainabilityBefore3
    DebugLog "Total Before: " & GlobalTotalBefore3
    DebugLog "Percentage Before: " & GlobalPercentageBefore3
End Sub

Sub CalculateSustainabilityAfter()
    ' Reset variabel global untuk after
    GlobalSustainabilityAfter3 = 0
    GlobalTotalAfter3 = 0
    GlobalPercentageAfter3 = 0
    
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
    
    Dim specKey As Variant
    For Each specKey In specSustainabilityAfter.Keys
        Dim persentaseSpecAfter As Double
        persentaseSpecAfter = 0
        
        If specTotalAfter(specKey) > 0 Then
            persentaseSpecAfter = (specSustainabilityAfter(specKey) / specTotalAfter(specKey))
        End If
        
        totalPersentaseSustainabilityAfter = totalPersentaseSustainabilityAfter + persentaseSpecAfter
        GlobalSustainabilityAfter3 = GlobalSustainabilityAfter3 + specSustainabilityAfter(specKey)
        GlobalTotalAfter3 = GlobalTotalAfter3 + specTotalAfter(specKey)
    Next specKey
    
    ' Simpan hasil ke variabel global
    If GlobalTotalAfter3 > 0 Then
        GlobalPercentageAfter3 = totalPersentaseSustainabilityAfter
    Else
        GlobalPercentageAfter3 = 0
    End If
    
    DebugLog "Sustainability After Calculation Completed"
    DebugLog "Total Sustainability After: " & GlobalSustainabilityAfter3
    DebugLog "Total After: " & GlobalTotalAfter3
    DebugLog "Percentage After: " & GlobalPercentageAfter3
    
    ' Debug detail perhitungan per spesifikasi
    DebugLog "=== DETAIL PERHITUNGAN PER SPESIFIKASI ==="
    DebugLog "Program: Calculate"
    DebugLog "Global Sustainability After: " & GlobalSustainabilityAfter3
    DebugLog "Global Total After: " & GlobalTotalAfter3
    DebugLog "Global Percentage After: " & GlobalPercentageAfter3

    ' Loop semua sheet untuk debug detail

    ' Cari index sheet RESUME
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i

    ' Debug per spesifikasi
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" Then
            DebugLog "--- Sheet: " & sh.Name & " ---"
        
       
            lastRow = sh.Cells(sh.Rows.Count, "H").End(xlUp).Row
        
            Dim sustainabilityTotal As Double
            Dim grandTotal As Double
            sustainabilityTotal = 0
            grandTotal = 0
        
            ' Nama kategori sustainability
       
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
                DebugLog "    Percentage: " & (sustainabilityTotal / grandTotal)
            End If
        End If
    Next i

    DebugLog "=== END DEBUG ==="
End Sub

Sub DisplaySustainabilityResults()
    ' Menampilkan hasil sustainability ke worksheet atau konsol
    DebugLog "=== SUSTAINABILITY CALCULATION RESULTS ==="
    DebugLog "Before Replacement:"
    DebugLog "  Total Sustainability: " & GlobalSustainabilityBefore3
    DebugLog "  Total Weight: " & GlobalTotalBefore3
    DebugLog "  Percentage: " & Format(GlobalPercentageBefore3, "0.00%")
    DebugLog ""
    DebugLog "After Replacement:"
    DebugLog "  Total Sustainability: " & GlobalSustainabilityAfter3
    DebugLog "  Total Weight: " & GlobalTotalAfter3
    DebugLog "  Percentage: " & Format(GlobalPercentageAfter3, "0.00%")
    DebugLog ""
    DebugLog "Difference:"
    DebugLog "  Sustainability Change: " & (GlobalSustainabilityAfter3 - GlobalSustainabilityBefore3)
    DebugLog "  Percentage Change: " & Format(GlobalPercentageAfter3 - GlobalPercentageBefore3, "0.00%")
    DebugLog "================================================"
    
    ' Opsional: Simpan ke worksheet tertentu (misalnya CALCULATE sheet)
    Dim wsCalc As Worksheet
    Set wsCalc = ThisWorkbook.Sheets("CALCULATE")
    
    ' Simpan hasil ke cell tertentu (sesuaikan dengan kebutuhan)
    wsCalc.Range("B35").Value = "Sustainability Before:"
    wsCalc.Range("C35").Value = GlobalPercentageBefore3
    wsCalc.Range("B36").Value = "Sustainability After:"
    wsCalc.Range("C36").Value = GlobalPercentageAfter3
    wsCalc.Range("B37").Value = "Change:"
    wsCalc.Range("C37").Value = GlobalPercentageAfter3 - GlobalPercentageBefore3
End Sub

' Fungsi untuk mendapatkan nilai sustainability (dapat dipanggil dari modul lain)
Function GetSustainabilityBefore() As Double
    GetSustainabilityBefore = GlobalPercentageBefore3
End Function

Function GetSustainabilityAfter() As Double
    GetSustainabilityAfter = GlobalPercentageAfter3
End Function

Function GetSustainabilityChange() As Double
    GetSustainabilityChange = GlobalPercentageAfter3 - GlobalPercentageBefore3
End Function

Sub update_multiple_material_data(replacements() As replacementData)
    Dim simpan_data As Double
    Dim nilai_pengganti As Double
    Dim found As Boolean
    Dim resumeIndex As Long
    Dim specSheets As Collection
    Dim sh As Worksheet
    Dim i As Long, j As Long
    
    Set specSheets = New Collection
    found = False
    
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
    
    Dim histWS As Worksheet
    Set histWS = ThisWorkbook.Sheets("HISTORY_UNDO")
    Dim lastHistRow As Long
    Dim ws As Worksheet
    Dim wsName As Variant
    
    ' Process setiap sheet
    For Each wsName In specSheets
        Set ws = ThisWorkbook.Sheets(wsName)
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        ' Dictionary untuk akumulasi material baru per sheet
        Dim newMaterialAccumulation As Object
        Set newMaterialAccumulation = CreateObject("Scripting.Dictionary")
        
        ' STEP 1: Proses material yang diganti dan kumpulkan akumulasi material baru
        For j = 1 To 3
            If replacements(j).isValid Then
                For i = 3 To lastRow
                    If Trim(UCase(ws.Cells(i, "H").Value)) = UCase(replacements(j).material_replaced) Then
                        If IsNumeric(ws.Cells(i, "I").Value) Then
                            simpan_data = ws.Cells(i, "I").Value
                            Dim oldValue As Double: oldValue = simpan_data
                            nilai_pengganti = 1 - replacements(j).percentage_material
                            Dim newValue As Double: newValue = simpan_data * nilai_pengganti
                            
                            ' Simpan histori untuk material yang diganti
                            lastHistRow = histWS.Cells(histWS.Rows.Count, 1).End(xlUp).Row + 1
                            With histWS
                                .Cells(lastHistRow, 1).Value = Now
                                .Cells(lastHistRow, 2).Value = ws.Name
                                .Cells(lastHistRow, 3).Value = i
                                .Cells(lastHistRow, 4).Value = "I"
                                .Cells(lastHistRow, 5).Value = replacements(j).material_replaced
                                .Cells(lastHistRow, 6).Value = oldValue
                                .Cells(lastHistRow, 7).Value = newValue
                                .Cells(lastHistRow, 8).Value = GlobalActionID
                                .Cells(lastHistRow, 9).Value = replacements(j).new_material
                                .Cells(lastHistRow, 10).Value = "REPLACE"
                            End With
                            
                            ws.Cells(i, "I").Value = newValue
                            found = True
                            
                            ' Akumulasi material baru
                            Dim addedAmount As Double
                            addedAmount = replacements(j).percentage_new_material * simpan_data
                            
                            If newMaterialAccumulation.exists(replacements(j).new_material) Then
                                newMaterialAccumulation(replacements(j).new_material) = _
                                                                                      newMaterialAccumulation(replacements(j).new_material) + addedAmount
                            Else
                                newMaterialAccumulation.Add replacements(j).new_material, addedAmount
                            End If
                            
                            ' Simpan class untuk referensi
                            If Not newMaterialAccumulation.exists(replacements(j).new_material & "_CLASS") Then
                                newMaterialAccumulation.Add replacements(j).new_material & "_CLASS", replacements(j).new_material_class
                            End If
                        End If
                    End If
                Next i
            End If
        Next j
        
        ' STEP 2: Process material baru yang terakumulasi
        Dim materialKey As Variant
        For Each materialKey In newMaterialAccumulation.Keys
            If Right(materialKey, 6) <> "_CLASS" Then ' Skip class entries
                Dim materialName As String: materialName = materialKey
                Dim totalAmount As Double: totalAmount = newMaterialAccumulation(materialKey)
                Dim materialClass As String: materialClass = newMaterialAccumulation(materialKey & "_CLASS")
                
                ' PERBAIKAN UTAMA: Cari material yang sudah ada di seluruh sheet terlebih dahulu
                Dim materialFoundRow As Long: materialFoundRow = 0
                Dim materialExists As Boolean: materialExists = False
                
                ' Scan seluruh sheet untuk mencari material yang sudah ada
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
                    ' Material sudah ada - update nilai
                    Dim oldExistingValue As Double: oldExistingValue = ws.Cells(materialFoundRow, "I").Value
                    
                    ' Simpan histori untuk material yang sudah ada
                    lastHistRow = histWS.Cells(histWS.Rows.Count, 1).End(xlUp).Row + 1
                    With histWS
                        .Cells(lastHistRow, 1).Value = Now
                        .Cells(lastHistRow, 2).Value = ws.Name
                        .Cells(lastHistRow, 3).Value = materialFoundRow
                        .Cells(lastHistRow, 4).Value = "I"
                        .Cells(lastHistRow, 5).Value = materialName
                        .Cells(lastHistRow, 6).Value = oldExistingValue
                        .Cells(lastHistRow, 7).Value = oldExistingValue + totalAmount
                        .Cells(lastHistRow, 8).Value = GlobalActionID
                        .Cells(lastHistRow, 9).Value = materialName
                        .Cells(lastHistRow, 10).Value = "ADD_TO_EXISTING"
                    End With
                    
                    ws.Cells(materialFoundRow, "I").Value = oldExistingValue + totalAmount
                    
                Else
                    ' Material belum ada - cari class dan tambahkan material baru
                    Dim classFoundRow As Long: classFoundRow = 0
                    For i = 3 To lastRow
                        If Trim(UCase(ws.Cells(i, "H").Value)) = UCase(materialClass) Then
                            If IsNumeric(ws.Cells(i, "J").Value) Then
                                classFoundRow = i
                                Exit For
                            End If
                        End If
                    Next i
                    
                    If classFoundRow > 0 Then
                        ' Cari posisi untuk insert material baru di bawah class
                        Dim insertRow As Long: insertRow = classFoundRow + 1
                        
                        ' Cari posisi yang tepat untuk insert (setelah material terakhir dalam class ini)
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
                                ' Ketemu class baru
                                If Trim(UCase(nextVal)) <> UCase(materialClass) Then
                                    insertRow = i
                                    Exit For
                                End If
                            End If
                        Next i
                        
                        ' Insert material baru
                        If insertRow <= lastRow Then
                            ws.Rows(insertRow).Insert Shift:=xlDown
                        End If
                        
                        ' Simpan histori untuk material baru
                        lastHistRow = histWS.Cells(histWS.Rows.Count, 1).End(xlUp).Row + 1
                        With histWS
                            .Cells(lastHistRow, 1).Value = Now
                            .Cells(lastHistRow, 2).Value = ws.Name
                            .Cells(lastHistRow, 3).Value = insertRow
                            .Cells(lastHistRow, 4).Value = "H"
                            .Cells(lastHistRow, 5).Value = materialName
                            .Cells(lastHistRow, 6).Value = ""
                            .Cells(lastHistRow, 7).Value = materialName
                            .Cells(lastHistRow, 8).Value = GlobalActionID
                            .Cells(lastHistRow, 9).Value = materialName
                            .Cells(lastHistRow, 10).Value = "INSERT_NEW_MATERIAL"
                        End With
                        
                        lastHistRow = histWS.Cells(histWS.Rows.Count, 1).End(xlUp).Row + 1
                        With histWS
                            .Cells(lastHistRow, 1).Value = Now
                            .Cells(lastHistRow, 2).Value = ws.Name
                            .Cells(lastHistRow, 3).Value = insertRow
                            .Cells(lastHistRow, 4).Value = "I"
                            .Cells(lastHistRow, 5).Value = materialName
                            .Cells(lastHistRow, 6).Value = 0
                            .Cells(lastHistRow, 7).Value = totalAmount
                            .Cells(lastHistRow, 8).Value = GlobalActionID
                            .Cells(lastHistRow, 9).Value = materialName
                            .Cells(lastHistRow, 10).Value = "SET_NEW_MATERIAL_VALUE"
                        End With
                        
                        ws.Rows(insertRow).Font.Bold = False
                        ws.Cells(insertRow, "H").Value = materialName
                        ws.Cells(insertRow, "I").Value = totalAmount
                        
                    Else
                        MsgBox "Class '" & materialClass & "' tidak ditemukan di sheet " & wsName, vbExclamation
                    End If
                End If
            End If
        Next materialKey
    Next wsName
    
    If found Then
       Application.ScreenUpdating = False
On Error GoTo Cleanup

'Call CreateMultipleSummarySheet(replacements)

Cleanup:
    Application.ScreenUpdating = True

        Call ModuleRingkasan.TampilkanLogRingkas
        
        ' Catat ke HISTORY_CHANGE
        Set ws = ThisWorkbook.Sheets("HISTORY_CHANGE")
        Dim nextRow As Long
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        
        If ws.Cells(2, 1).Value = "" Then
            ws.Cells(2, 1).Value = 1
            nextRow = 2
        Else
            ws.Cells(nextRow, 1).Value = ws.Cells(nextRow - 1, 1).Value + 1
        End If
        
        ws.Cells(nextRow, 2).Value = Now
        ws.Cells(nextRow, 3).Value = Sheets("CALCULATE").Range("B33").Value
        ws.Cells(nextRow, 4).Value = "Multiple Material Replacement (Action ID: " & GlobalActionID & ")"
        
        ShowInfo "Multiple update berhasil dan histori tersimpan."
    Else
        MsgBox "Tidak ada material yang ditemukan untuk diganti.", vbExclamation
    End If
End Sub

Sub CreateMultipleSummarySheet(replacements() As replacementData)
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim sh As Worksheet
    Dim specSheets As Collection
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim sheetName As String
    
    sheetName = "Multiple_Summary_" & GlobalActionID
    Set specSheets = New Collection
    
    ' Hapus sheet jika sudah ada
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Buat sheet baru
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    Dim resumeIndex As Long: resumeIndex = 0
    
    ' Cari posisi sheet RESUME
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "RESUME" Then
            resumeIndex = i
            Exit For
        End If
    Next i
    
    ' Ambil semua sheet sebelum RESUME yang valid
    For i = 1 To resumeIndex - 1
        Set sh = ThisWorkbook.Sheets(i)
        If sh.Name <> "CATEGORY SPESIFICATION" And Left(sh.Name, 6) <> "Before" And Left(sh.Name, 8) <> "Multiple" Then
            specSheets.Add sh.Name
        End If
    Next i
    
    With ws
        ' Header informasi
        .Range("B2").Value = "MULTIPLE MATERIAL REPLACEMENT SUMMARY"
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Size = 14
        
        ' Detail replacements
        Dim startRow As Long: startRow = 4
        For j = 1 To 3
            If replacements(j).isValid Then
                .Range("B" & startRow).Value = "SET " & j & ":"
                .Range("B" & startRow).Font.Bold = True
                
                .Range("C" & startRow).Value = "Replaced Material:"
                .Range("D" & startRow).Value = replacements(j).material_replaced
                
                .Range("C" & (startRow + 1)).Value = "Replacement %:"
                .Range("D" & (startRow + 1)).Value = replacements(j).percentage_material
                .Range("D" & (startRow + 1)).NumberFormat = "0.00%"
                
                .Range("C" & (startRow + 2)).Value = "New Material:"
                .Range("D" & (startRow + 2)).Value = replacements(j).new_material
                
                .Range("C" & (startRow + 3)).Value = "New Material %:"
                .Range("D" & (startRow + 3)).Value = replacements(j).percentage_new_material
                .Range("D" & (startRow + 3)).NumberFormat = "0.00%"
                
                .Range("C" & (startRow + 4)).Value = "Material Class:"
                .Range("D" & (startRow + 4)).Value = replacements(j).new_material_class
                
                startRow = startRow + 6
            End If
        Next j
        
        ' Tabel ringkasan
        startRow = startRow + 2                  'Used for range positioning
        .Range("B" & startRow).Value = "SUSTAINABILITY SUMMARY"
        .Range("B" & startRow).Font.Bold = True
        .Range("B" & startRow).Interior.Color = RGB(173, 216, 230)
        
        startRow = startRow + 2                  'Used for range positioning
        .Range("B" & startRow).Value = "Material"
        .Range("B" & startRow).Font.Bold = True
        .Range("B" & startRow).Interior.Color = RGB(173, 216, 230)
        .Range("B" & startRow).HorizontalAlignment = xlCenter
        
        ' Header untuk setiap sheet SPEC
        For i = 1 To specSheets.Count
            .Cells(startRow, i + 2).Value = specSheets(i)
            .Cells(startRow, i + 2).Font.Bold = True
            .Cells(startRow, i + 2).HorizontalAlignment = xlCenter
            .Cells(startRow, i + 2).Interior.Color = RGB(173, 216, 230)
        Next i
        
        ' Daftar kategori material
        Dim materials As Variant
        materials = Array( _
                    Array("Natural Rubber", RGB(255, 0, 0)), _
                    Array("Synthetic Rubber", RGB(0, 0, 255)), _
                    Array("Reclaim Rubber", RGB(255, 0, 0)), _
                    Array("Filler", RGB(0, 0, 255)), _
                    Array("Sustainable Filler", RGB(255, 0, 0)), _
                    Array("Oil", RGB(0, 0, 255)), _
                    Array("Sustainable Oil", RGB(255, 0, 0)), _
                    Array("Chemical", RGB(0, 0, 255)), _
                    Array("Sustainable Chemical", RGB(255, 0, 0)), _
                    Array("Reinforcement", RGB(0, 0, 255)), _
                    Array("Sustainable Reinforcement", RGB(255, 0, 0)), _
                    Array("Total Material", RGB(0, 0, 0)), _
                    Array("Material Sustainability", RGB(0, 0, 0)) _
                    )
        
        ' Tulis material dan formula
        For i = 0 To UBound(materials)
            rowIndex = startRow + 1 + i
            .Range("B" & rowIndex).Value = materials(i)(0)
            If materials(i)(0) = "Total Material" Then
                .Range("B" & rowIndex & ":" & Chr(64 + 2 + specSheets.Count) & rowIndex).Interior.Color = RGB(255, 200, 100)
                .Range("B" & rowIndex).Font.Bold = True
            ElseIf materials(i)(0) = "Material Sustainability" Then
                .Range("B" & rowIndex & ":" & Chr(64 + 2 + specSheets.Count) & rowIndex).Interior.Color = RGB(255, 255, 0)
                .Range("B" & rowIndex).Font.Bold = True
            Else
                .Range("B" & rowIndex).Font.Color = materials(i)(1)
            End If
        Next i
        
        ' VLOOKUP untuk tiap SPEC sheet
        For i = 1 To specSheets.Count
            For rowIndex = startRow + 1 To startRow + 13
                Dim material As String: material = .Range("B" & rowIndex).Value
                .Cells(rowIndex, i + 2).Formula = "=VLOOKUP(""" & material & """,'" & specSheets(i) & "'!L2:N15,3,FALSE)"
                .Cells(rowIndex, i + 2).NumberFormat = "0.00%"
            Next rowIndex
        Next i
        
        ' Border
        lastCol = 2 + specSheets.Count
        .Range(.Cells(startRow, 2), .Cells(startRow + 13, lastCol)).BorderAround ColorIndex:=1, Weight:=xlMedium
        .Range(.Cells(startRow, 2), .Cells(startRow + 13, lastCol)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(startRow, 2), .Cells(startRow + 13, lastCol)).Borders(xlInsideVertical).LineStyle = xlContinuous
        
        .Columns("A:" & Chr(64 + lastCol)).AutoFit
    End With
    
    ws.Activate
End Sub

Sub CopySheetByValue(new_material As String)
    On Error GoTo ErrorHandler
    Dim srcSheet As Worksheet
    Dim dstSheet As Worksheet
    Dim rngSrc As Range
    Dim rngDst As Range

    Set srcSheet = ThisWorkbook.Sheets("RESUME")

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Before " & new_material).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set dstSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    dstSheet.Name = "Before " & new_material

    With srcSheet.UsedRange
        Set rngSrc = srcSheet.Range("A1:XFD100")
        Set rngDst = dstSheet.Range("A1:XFD100")

        rngDst.Value = rngSrc.Value

        rngSrc.Copy
        rngDst.PasteSpecial Paste:=xlPasteFormats
        rngDst.PasteSpecial Paste:=xlPasteColumnWidths
        Application.CutCopyMode = False
    End With

    ShowInfo "Sheet RESUME berhasil disalin ke 'Before " & new_material
    
       
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    MsgBox "Error dalam copy sheet: " & Err.Description, vbCritical
    
End Sub

Sub redo_last_action()
    Application.ScreenUpdating = False
   
    
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets("HISTORY_UNDO")
    Dim lastActionID As String
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        ShowInfo "Tidak ada histori undo."
        Call ModuleRingkasan.TampilkanLogRingkas
        Exit Sub
    End If
    
    lastActionID = wsLog.Cells(lastRow, "H").Value
    
    ' Kumpulkan semua action dalam ActionID yang sama untuk diproses bersamaan
    Dim actionsToProcess As Collection
    Set actionsToProcess = New Collection
    
    For i = lastRow To 2 Step -1
        If wsLog.Cells(i, "H").Value = lastActionID Then
            Dim actionData As Object
            Set actionData = CreateObject("Scripting.Dictionary")
            actionData("Row") = i
            actionData("Sheet") = wsLog.Cells(i, "B").Value
            actionData("TargetRow") = wsLog.Cells(i, "C").Value
            actionData("TargetCol") = wsLog.Cells(i, "D").Value
            actionData("MaterialName") = wsLog.Cells(i, "E").Value
            actionData("OldValue") = wsLog.Cells(i, "F").Value
            actionData("NewValue") = wsLog.Cells(i, "G").Value
            actionData("ActionType") = wsLog.Cells(i, "J").Value
            actionsToProcess.Add actionData
        End If
    Next i
    
    ' Process actions dalam urutan yang benar (terbalik dari yang dikumpulkan)
    Dim actionItem As Object
    For i = actionsToProcess.Count To 1 Step -1
        Set actionItem = actionsToProcess(i)
        
        Dim wsTarget As Worksheet
        Dim rowT As Long, colT As String, actionType As String
        Dim oldVal As Variant, materialName As String
        
        Set wsTarget = ThisWorkbook.Sheets(actionItem("Sheet"))
        rowT = actionItem("TargetRow")
        colT = actionItem("TargetCol")
        oldVal = actionItem("OldValue")
        actionType = actionItem("ActionType")
        materialName = actionItem("MaterialName")
        
        Select Case actionType
            ' Untuk material yang diganti - kembalikan ke nilai original
        Case "REPLACE"
            wsTarget.Cells(rowT, colT).Value = oldVal
                
            ' Untuk material existing yang ditambahkan - kembalikan ke nilai sebelumnya
        Case "ADD_EXISTING", "ADD_EXISTING_ACCUMULATED", "ADD_TO_EXISTING"
            wsTarget.Cells(rowT, colT).Value = oldVal
                
            ' Untuk material baru yang diinsert - hapus row
        Case "INSERT_ROW", "ADD_NEW", "INSERT_ROW_ACCUMULATED", "ADD_NEW_ACCUMULATED", "INSERT_NEW_MATERIAL", "INSERT_NEW_MATERIAL_WITH_VALUE"
            If colT = "H" Then
                wsTarget.Rows(rowT).Delete Shift:=xlUp
            End If
                
            ' Untuk action lainnya (backward compatibility)
        Case Else
            ' Coba restore nilai jika ada
            If colT <> "" And oldVal <> "" Then
                wsTarget.Cells(rowT, colT).Value = oldVal
            End If
        End Select
    Next i
    
    ' Hapus semua log entries untuk ActionID ini
    For i = lastRow To 2 Step -1
        If wsLog.Cells(i, "H").Value = lastActionID Then
            wsLog.Rows(i).Delete
        End If
    Next i
    
    ' Catat ke HISTORY_CHANGE
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HISTORY_CHANGE")
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If ws.Cells(2, 1).Value = "" Then
        ws.Cells(2, 1).Value = 1
        nextRow = 2
    Else
        ws.Cells(nextRow, 1).Value = ws.Cells(nextRow - 1, 1).Value + 1
    End If
        
    ws.Cells(nextRow, 2).Value = Now
    ws.Cells(nextRow, 3).Value = Sheets("CALCULATE").Range("B33").Value
    ws.Cells(nextRow, 4).Value = "Undo Material Replacement (Action ID: " & lastActionID & ")"
    
    Call ModuleRingkasan.TampilkanLogRingkas
    ShowInfo "Undo untuk action terakhir berhasil."
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub


End Sub

Sub reset_globaldata()
    GlobalSustainabilityBefore3 = 0
    GlobalSustainabilityAfter3 = 0
    GlobalTotalBefore3 = 0
    GlobalTotalAfter3 = 0
    GlobalPercentageBefore3 = 0
    GlobalPercentageAfter3 = 0
End Sub

' Permanenkan perubahan terakhir dengan menghapus histori undo untuk Action ID
' paling baru. Jalankan sub ini setelah proses submit berhasil apabila data
' sudah dianggap final.
Sub permanenkan_terakhir()
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim lastActionID As String
    Dim i As Long

    Set wsLog = ThisWorkbook.Sheets("HISTORY_UNDO")
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        ShowInfo "Tidak ada histori yang dapat dihapus."
        Exit Sub
    End If

    lastActionID = wsLog.Cells(lastRow, "H").Value

    ' Hapus semua baris dengan Action ID terakhir
    For i = lastRow To 2 Step -1
        If wsLog.Cells(i, "H").Value = lastActionID Then
            wsLog.Rows(i).Delete
        Else
            Exit For
        End If
    Next i

    ' Catat aksi permanenkan pada HISTORY_CHANGE
    Dim ws As Worksheet
    Dim nextRow As Long
    Set ws = ThisWorkbook.Sheets("HISTORY_CHANGE")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If ws.Cells(2, 1).Value = "" Then
        ws.Cells(2, 1).Value = 1
        nextRow = 2
    Else
        ws.Cells(nextRow, 1).Value = ws.Cells(nextRow - 1, 1).Value + 1
    End If

    ws.Cells(nextRow, 2).Value = Now
    ws.Cells(nextRow, 3).Value = Sheets("CALCULATE").Range("B33").Value
    ws.Cells(nextRow, 4).Value = "Permanenkan Perubahan (Action ID: " & lastActionID & ")"

    Call ModuleRingkasan.TampilkanLogRingkas
    ShowInfo "Perubahan telah dipermanenkan dan histori undo dihapus."
End Sub


