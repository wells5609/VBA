VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SessionTypeDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISessionType

Public Property Get Name() As String
    Name = "default"
End Property

Public Sub Start(obj As Session)
    ' Does nothing
End Sub

Public Sub Destroy(obj As Session)
    ' Does nothing
End Sub

Private Property Get ISessionType_Name() As String
    ISessionType_Name = Name
End Property

Private Sub ISessionType_Start(obj As Session)
    Start obj
End Sub

Private Sub ISessionType_Destroy(obj As Session)
    Destroy obj
End Sub
