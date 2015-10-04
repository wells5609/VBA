VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISessionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ISessionType
'
' Interface for a session type.

Option Explicit

' Returns the session type name.
'
' @return String
Public Property Get Name() As String
End Property

' Called on Session.Start
'
' @param Session obj
Public Sub Start(obj As Session)
End Sub

' Called on Session.Destroy
'
' @param Session obj
Public Sub Destroy(obj As Session)
End Sub
