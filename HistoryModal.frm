VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HistoryModal 
   ClientHeight    =   3936
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "HistoryModal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "historymodal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Modal"
Option Explicit

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range

    Set ws = ThisWorkbook.Sheets("HISTORY_CHANGE")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Ambil data dari kolom B:D mulai dari baris 2
    Set dataRange = ws.Range("B2:D" & lastRow)

    ListBox1.ColumnCount = 3
    ListBox1.ColumnWidths = "80;60;80"
    ListBox1.list = dataRange.Value
End Sub

