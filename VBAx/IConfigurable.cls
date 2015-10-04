VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfigurable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Interface: IConfigurable
' Contract for an object that can be configured.

Option Explicit

' Returns the Config instance.
Public Property Get Config() As Config
End Property

' Sets the Config instance and configures the object.
'
' @param Config obj
Public Sub Configure(obj As Config)
End Sub
