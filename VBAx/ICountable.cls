VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICountable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Interface: ICountable
' Contract for an object with a count of something.

Option Explicit

' Returns the number of items in the object.
'
' @return Long
Public Property Get Count() As Long
End Property
