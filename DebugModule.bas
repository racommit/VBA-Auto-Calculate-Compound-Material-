Attribute VB_Name = "DebugModule"
Option Explicit
Public Const DEBUG_MODE As Boolean = False
Public Const SHOW_INFO As Boolean = False

Public Sub DebugLog(ByVal msg As String)
    If DEBUG_MODE Then Debug.Print msg
End Sub

Public Sub ShowInfo(ByVal msg As String, Optional ByVal title As String = "")
    If SHOW_INFO Then
        MsgBox msg, vbInformation, title
    Else
        DebugLog msg
    End If
End Sub
