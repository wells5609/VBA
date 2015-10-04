VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IActionable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Interface: IActionable
' Contract for an object with actions.

Option Explicit

' Called before an action is executed via DoAction().
'
' @param IAction obj
Public Sub BeforeAction(obj As IAction)
End Sub

' Called after an action is executed via DoAction().
'
' @param IAction obj
Public Sub AfterAction(obj As IAction)
End Sub
