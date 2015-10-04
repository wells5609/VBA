VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Interface: IAction
' Contract for an action object.

Option Explicit

' Returns the action name.
'
' @return String
Public Property Get Name() As String
End Property

' Executes the action.
'
' @param IActionable obj [Optional]
Public Sub DoAction(Optional obj As IActionable)
End Sub
