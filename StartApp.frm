VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartApp 
   Caption         =   "Start App"
   ClientHeight    =   2280
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "StartApp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule EncapsulatePublicField
Option Explicit
Public UserInputValue As String

Private Sub CommandButton1_Click()
    If Trim(TextBox1.Value) = "" Then
        MsgBox "Input tidak boleh kosong!", vbExclamation, "Peringatan"
        TextBox1.SetFocus
    Else
        UserInputValue = TextBox1.Value
        Sheets("CALCULATE").Range("B17").Value = "Author : " + UserInputValue
        Sheets("CALCULATE").Range("B33").Value = "Author : " + UserInputValue
        Me.Hide
    End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then                         ' Enter ditekan
        If Trim(TextBox1.Value) <> "" Then
            UserInputValue = TextBox1.Value
            Sheets("CALCULATE").Range("B17").Value = "Author : " + UserInputValue
            Sheets("CALCULATE").Range("B33").Value = "Author : " + UserInputValue
            Me.Hide
        Else
            MsgBox "Isi input terlebih dahulu sebelum menekan Enter!", vbExclamation
        End If
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then        ' Artinya user klik tombol [X]
        If Trim(TextBox1.Value) = "" Then
            MsgBox "Silakan isi input terlebih dahulu sebelum menutup form!", vbExclamation
            Cancel = True                        ' Gagalkan penutupan form
        End If
    End If
End Sub

