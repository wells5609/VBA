VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Implements ICountable

Private Type CacheEntry
    Name As String
    Value As Variant
    Expires As Date
End Type

Public DefaultTTL As Long

Private Storage As Dictionary
Private Expirations As Dictionary

Private Sub Class_Initialize()
    Set Storage = New Dictionary
    Set Expirations = New Dictionary
    DefaultTTL = 3600
End Sub

Private Sub Class_Terminate()
    Set Storage = Nothing
    Set Expirations = Nothing
End Sub

Public Function Create() As Cache
    Set Create = New Cache
End Function

' Returns the expiration of an item.
'
' @param String Key
' @return Variant|Date
Public Property Get Expiration(ByVal Key As String) As Variant
    Expiration = Expirations(Key)
End Property

' Returns the number of items in the cache.
'
' @return Long
Public Property Get Count() As Long
    Count = Storage.Count
End Property

' Checks whether an item exists in the cache and is not expired.
'
' @param String Key
' @return Boolean
Public Function Exists(ByVal Key As String) As Boolean
    If Storage.Exists(Key) Then
        Exists = Expirations(Key) >= Now
    Else
        Exists = False
    End If
End Function

' Fetches a value from the cache.
'
' If the item does not exist or is expired, Empty is returned.
'
' @param String Key
' @return Variant|Empty
Public Function Fetch(ByVal Key As String) As Variant
    
    Fetch = Empty
    
    If Storage.Exists(Key) Then
        If Expirations(Key) < Now Then
            Delete Key
        ElseIf VBA.IsObject(Storage(Key)) Then
            Set Fetch = Storage(Key)
        Else
            Let Fetch = Storage(Key)
        End If
    End If
    
End Function

' Stores a value in the cache for the given length of time.
'
' If no TTL is given, the default is used.
'
' @param String Key
' @param Variant Value
' @param Long TTL [Optional]
Public Sub Store(ByVal Key As String, Value As Variant, Optional TTL As Long)
    
    If VBA.IsObject(Value) Then
        Set Storage(Key) = Value
    Else
        Let Storage(Key) = Value
    End If
    
    If TTL = 0 Then Let TTL = DefaultTTL
    
    Let Expirations(Key) = VBA.Now + VBA.TimeSerial(0, 0, TTL)
    
End Sub

' Updates an existing value without changing the expiration.
'
' If the item does not exist, it is stored and its expiration is set using the default TTL.
'
' @param String Key
' @param Variant Value
Public Sub Update(ByVal Key As String, Value As Variant)
    
    If Not Storage.Exists(Key) Then
        Store Key, Value
    ElseIf VBA.IsObject(Value) Then
        Set Storage(Key) = Value
    Else
        Let Storage(Key) = Value
    End If
    
End Sub

' Deletes an item from the cache and clears its expiration time.
'
' @param String Key
Public Sub Delete(ByVal Key As String)
    Expirations.Remove Key
    Storage.Remove Key
End Sub

' Sets an item's expiration time.
'
' @param String Key
' @param Date Expiration
Public Sub SetExpiration(ByVal Key As String, Expiration As Date)
    If Expiration < VBA.Now Then
        VBA.Err.Raise 1726, "Cache.SetExpiration", "Date cannot be in the past"
    End If
    Let Expirations(Key) = Expiration
End Sub

' Removes expired items from the cache.
Public Sub CleanUp()
    Dim Key As Variant
    For Each Key In Expirations.Keys
        If Expirations(Key) < VBA.Now Then Delete Key
    Next
End Sub

' Removes expired items from the cache.
Public Sub PurgeExpired()
    Call CleanUp
End Sub

' Checks whether the given item is expired.
'
' @param String Key
' @return Boolean
Public Function IsExpired(ByVal Key As String) As Boolean
    IsExpired = Expirations(Key) < VBA.Now
End Function


Private Property Get ICountable_Count() As Long
    ICountable_Count = Count
End Property
