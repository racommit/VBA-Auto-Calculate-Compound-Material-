Attribute VB_Name = "ModuleReset"
'@Folder "button_features"
Option Explicit

Sub Button3_Click()
    historymodal.Show
End Sub

Sub reset_input()
    Dim wsCategory As Worksheet
    Set wsCategory = ThisWorkbook.Sheets("CALCULATE")

    wsCategory.Range("C5:C9").Value = ""
    wsCategory.Range("G5:G9").Value = ""
    wsCategory.Range("K5:K9").Value = ""
    Call btnResetAll_Click
    
End Sub
Sub ResetAllData()
    Dim wsCalc As Worksheet
    Dim WSsim As Worksheet

    Set wsCalc = ThisWorkbook.Sheets("CALCULATE")
    Set WSsim = ThisWorkbook.Sheets("SIMULATION_PROCESS")

    Dim response As VbMsgBoxResult
    response = MsgBox("Apakah Anda yakin ingin mereset semua data input dan hasil kalkulasi?" & vbCrLf & vbCrLf & _
                      "Tindakan ini tidak dapat dibatalkan!", vbYesNo + vbQuestion + vbDefaultButton2, "Konfirmasi Reset")

    If response = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Mereset data..."

    ' ===== RESET SHEET CALCULATE =====
    wsCalc.Range("C5:C9").ClearContents
    wsCalc.Range("G5:G9").ClearContents
    wsCalc.Range("K5:K9").ClearContents

    ' ===== RESET SHEET SIMULATION_PROCESS =====
    Call ClearSimulationResults(WSsim)

    wsCalc.Activate
    wsCalc.Range("C5").Select

    Application.ScreenUpdating = True
    Application.StatusBar = False

    ShowInfo "Reset berhasil!" & vbCrLf & vbCrLf & _
           "Semua data input dan hasil kalkulasi telah dihapus." & vbCrLf & _
           "Anda dapat memasukkan data baru sekarang.", "Reset Selesai"
End Sub


Sub ResetInputOnly()
    ' Sub untuk mereset hanya input data (tanpa hasil kalkulasi)
    
    Dim wsCalc As Worksheet
    Set wsCalc = ThisWorkbook.Sheets("CALCULATE")
    
    ' Konfirmasi reset input
    Dim response As VbMsgBoxResult
    response = MsgBox("Reset hanya data input (hasil kalkulasi akan tetap ada)?", _
                      vbYesNo + vbQuestion, "Reset Input Saja")
    
    If response = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Reset hanya input di sheet CALCULATE
    wsCalc.Range("C5:C9").ClearContents          ' Material A
    wsCalc.Range("G5:G9").ClearContents          ' Material B
    wsCalc.Range("K5:K9").ClearContents          ' Material C
    
    wsCalc.Activate
    wsCalc.Range("C5").Select
    
    Application.ScreenUpdating = True
    
    ShowInfo "Input data berhasil direset!", "Reset Input Selesai"
    
End Sub

Sub ResetResultsOnly()
    ' Sub untuk mereset hanya hasil kalkulasi (input tetap ada)
    
    Dim WSsim As Worksheet
    Set WSsim = ThisWorkbook.Sheets("SIMULATION_PROCESS")
    
    ' Konfirmasi reset results
    Dim response As VbMsgBoxResult
    response = MsgBox("Reset hanya hasil kalkulasi (input data akan tetap ada)?", _
                      vbYesNo + vbQuestion, "Reset Hasil Saja")
    
    If response = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Reset semua hasil di SIMULATION_PROCESS
    Call ClearSimulationResults(WSsim)
    
    Application.ScreenUpdating = True
    
    ShowInfo "Hasil kalkulasi berhasil direset!" & vbCrLf & _
           "Silakan jalankan simulasi kembali.", "Reset Hasil Selesai"
    
End Sub
Private Sub ClearSimulationResults(ByRef ws As Worksheet)
    ' Clear all simulation output in SIMULATION_PROCESS

    ' Clear stepwise/combination output (before/after/tertimbang)
    ws.Range("E8:H10").ClearContents      ' F8:H10 (output stepwise)
    ws.Range("D8:D10").ClearContents      ' Output material baru stepwise
    ws.Range("D13:D15").ClearContents     ' Output kelas baru stepwise

    ' Clear old/new material before/after (per-material A/B/C)
    ws.Range("E65:F67").ClearContents     ' Material lama before/after A/B/C
    ws.Range("E71:F73").ClearContents     ' Material baru before/after A/B/C
    ws.Range("G65:G67").ClearContents     ' Status lama A/B/C (jika ada)
    ws.Range("G71:G73").ClearContents     ' Status baru A/B/C (jika ada)
    ws.Range("C65:C67").ClearContents     ' Output kode material lama

    ' Clear category results
    ws.Range("D46:F48").ClearContents

    ' Clear sustainability total & tertimbang
    ws.Range("D27:D28").ClearContents     ' Total Sustainability Before/After
    ws.Range("E27:E28").ClearContents     ' Sustainability tertimbang Before/After

    ' (Jika ada tambahan area baru di masa depan, tambahkan di sini)
End Sub



' ===== BUTTON EVENT HANDLERS =====
' Tambahkan sub ini jika menggunakan button controls

Sub btnResetAll_Click()
    ' Event handler untuk tombol Reset All
    Call ResetAllData
End Sub

Sub btnResetInput_Click()
    ' Event handler untuk tombol Reset Input Only
    Call ResetInputOnly
End Sub

Sub btnResetResults_Click()
    ' Event handler untuk tombol Reset Results Only
    Call ResetResultsOnly
End Sub

Sub reset_globaldata()
    GlobalSustainabilityBefore = 0
    GlobalSustainabilityAfter = 0
    GlobalTotalBefore = 0
    GlobalTotalAfter = 0
    GlobalPercentageBefore = 0
    GlobalPercentageAfter = 0
    GlobalAvgPercentageBefore = 0
    GlobalAvgPercentageAfter = 0
End Sub



