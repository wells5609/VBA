VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDependencyProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Interface: IDependencyProvider
' Provides objects to a DIContainer by resolving one or more identifiers.

Option Explicit

' Returns the ID (or list of IDs) provided by this object.
'
' @return Variant(Of Array)|String
Public Property Get ProvidedID() As Variant
End Property

' Whether the object(s) provided are to be shared.
'
' Returns boolean if object provides only 1 object (i.e. returns a string from ProvidedID)
' Returns an array of IDs if multiple objects are provided
'
' @return Variant(Of Boolean|Array)
Public Property Get Share() As Variant
End Property

' Resolves the object and returns it to the container.
'
' @param Variant ID
' @return Object
Public Function Resolve(ID As Variant) As Object
End Function
