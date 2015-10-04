VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Interface: IChild
' Contract for an object with a parent object.

Option Explicit

' Returns the object's parent.
'
' @param Variant Index [Optional] Identifier for the desired parent.
' @return Object
Public Property Get Parent(Optional Index As Variant) As Object
End Property
