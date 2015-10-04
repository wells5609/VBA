VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Class: Config
' Basic configuration object.

Option Explicit
Implements ICountable

Private Data As Scripting.Dictionary

Private Sub Class_Initialize()
    Set Data = New Scripting.Dictionary
End Sub
Private Sub Class_Terminate()
    Set Data = Nothing
End Sub

' Creates and returns a new Config object.
'
' @return Config
Public Function Create() As Config
    Set Create = New Config
End Function

' Returns a configuration value.
'
' @param String Key
' @return Variant
Public Property Get Item(ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(Data(Key)) Then
        Set Item = Data(Key)
    Else
        Let Item = Data(Key)
    End If
End Property

' Sets an object configuration value.
'
' @param String Key
' @param Object Value
Public Property Set Item(ByVal Key As String, Value As Object)
    Set Data(Key) = Value
End Property

' Sets a non-object configuration value.
'
' @param String Key
' @param Variant Value
Public Property Let Item(ByVal Key As String, Value As Variant)
    Let Data(Key) = Value
End Property

' Returns whether a value exists with the given key.
'
' @param String Key
' @return Boolean
Public Property Get Exists(ByVal Key As String) As Boolean
    Let Exists = Data.Exists(Key)
End Property

' Returns an array of config values.
'
' @return Variant(Of Array)
Public Property Get Items() As Variant()
    Let Items = Data.Items
End Property

' Returns an array of config keys.
'
' @return Variant(Of Array)
Public Property Get Keys() As Variant()
    Let Keys = Data.Keys
End Property

' Returns the number of items in the config.
'
' @return Long
Public Property Get Count() As Long
    Count = Data.Count
End Property

' Adds a value to the object.
'
' @param String Key
' @param Variant Value
Public Sub Add(ByVal Key As String, Value As Variant)
    Data.Add Key, Value
End Sub

' Assigns a value.
'
' @param String Key
' @param Variant Value
Public Sub Assign(ByVal Key As String, Value As Variant)
    If VBA.IsObject(Value) Then
        Set Data(Key) = Value
    Else
        Let Data(Key) = Value
    End If
End Sub

' Removes an item from the object.
'
' @param String Key
Public Sub Remove(ByVal Key As String)
    Data.Remove Key
End Sub

' Clears the object of all values.
Public Sub Clear()
    Data.RemoveAll
    Set Data = New Scripting.Dictionary
End Sub

' Adds items from a Dictionary.
'
' @param Dictionary obj
Public Sub AddFromDictionary(obj As Dictionary)
    Dim Key As Variant
    For Each Key In obj.Keys: Assign Key, obj(Key): Next
End Sub

' Adds a sub-configuration object.
'
' @param String Key
' @param Object(Of Config|Dictionary) obj
Public Sub AddConfig(ByVal Key As String, Optional obj As Dictionary)
    Dim conf As Config: Set conf = New Config
    If Not obj Is Nothing Then conf.AddFromDictionary obj
    Assign Key, conf
End Sub

' Returns a sub-configuration object.
'
' @param String Key
' @return Config
Public Function GetConfig(ByVal Key As String) As Config
    Set GetConfig = Data(Key)
End Function

' Returns a copy of the Config with all Config objects converted to Dictionary objects.
'
' @return Scripting.Dictionary
Public Function toDictionary() As Dictionary
    Dim obj As New Dictionary
    Dim Key As Variant
    For Each Key In Data.Keys
        If TypeOf Data(Key) Is Config Then
            obj.Add Key, Data(Key).toDictionary()
        Else
            obj.Add Key, Data(Key)
        End If
    Next
    Set toDictionary = obj
End Function


Private Property Get ICountable_Count() As Long
    ICountable_Count = Count
End Property

