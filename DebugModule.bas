Attribute VB_Name = "DebugModule"
Option Explicit

Public Const DEBUG_MODE As Boolean = False
Public Const SHOW_INFO As Boolean = True

Public Sub DebugLog(ByVal msg As String)
    If DEBUG_MODE Then Debug.Print msg
End Sub

Public Sub ShowInfo(ByVal msg As String, Optional ByVal title As String = "Info")
    If SHOW_INFO Then MsgBox msg, vbInformation, title
End Sub

Public Function IsValidPercentage(ByVal pct As Double) As Boolean
    IsValidPercentage = (pct >= 0 And pct <= 1)
End Function
