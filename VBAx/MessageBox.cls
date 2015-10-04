VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Class: MessageBox
' Wrapper for the MsgBox function.

Option Explicit

' Displays a message in a MsgBox
'
' @param String Msg
' @param String Title
Public Sub Display(ByVal Msg As String, Optional Style As VbMsgBoxStyle = vbOKOnly, Optional Title As String)
Attribute Display.VB_UserMemId = 0
    VBA.MsgBox Msg, Style, Title
End Sub

Public Function YesNo(ByVal Message As String, Optional ByVal Title As String = "Confirm") As Boolean
    Dim result As VbMsgBoxResult
    result = VBA.MsgBox(Message, vbYesNo + vbQuestion, Title)
    YesNo = result = vbYes
End Function

Public Sub Info(ByVal Message As String, Optional ByVal Title As String = "Information")
    VBA.MsgBox Message, vbInformation + vbOKOnly, Title
End Sub

Public Sub Warning(ByVal Message As String, Optional ByVal Title As String = "Warning")
    VBA.MsgBox Message, vbExclamation + vbOKOnly, Title
End Sub

Public Sub Error(ByVal Message As String, Optional ByVal Title As String = "Error")
    VBA.MsgBox Message, vbOKOnly + vbCritical, Title
End Sub
