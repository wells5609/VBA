VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDIAware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Interface: IDIAware
' Contract for a DI-aware object.

Option Explicit

' Returns the DIContainer.
'
' @return DIContainer
Public Property Get di() As DIContainer
End Property

' Sets the DIContainer.
'
' @param DIContainer obj
Public Property Set di(obj As DIContainer)
End Property
