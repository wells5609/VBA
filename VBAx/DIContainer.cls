VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DIContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Class: DIContainer
' VBA dependency injection/service locator container.

Option Explicit
Implements ICountable

Private myObjects As Scripting.Dictionary
Private mySharedIDs As Scripting.Dictionary
Private myProviders As Scripting.Dictionary
Private myCallbacks As Scripting.Dictionary

' Create the internal storage objects.
Private Sub Class_Initialize()
    Set myObjects = New Scripting.Dictionary
    Set mySharedIDs = New Scripting.Dictionary
    Set myProviders = New Scripting.Dictionary
    Set myCallbacks = New Scripting.Dictionary
End Sub
Private Sub Class_Terminate()
    Set myObjects = Nothing
    Set mySharedIDs = Nothing
    Set myProviders = Nothing
    Set myCallbacks = Nothing
End Sub

' Returns an instance of the container.
'
' The default instance is returned by default. Pass True to create a new instance.
'
' @param Boolean CreateNew [Optional] Default = False
' @return DIContainer
Public Function Instance(Optional CreateNew As Boolean = False) As DIContainer
    Static firstInstance As DIContainer
    If firstInstance Is Nothing Then
        Set firstInstance = New DIContainer
        Set Instance = firstInstance
        CreateNew = True ' Allow check by reference
    ElseIf CreateNew Then
        Set Instance = New DIContainer
    Else
        Set Instance = firstInstance
    End If
End Function

' Creates and returns a new DIContainer.
'
' @return DIContainer
Public Function Create() As DIContainer
    Set Create = Instance(True)
End Function

' Sets a shared object in the container.
'
' @param Variant ID
' @param Object obj
Public Property Set Object(ID As Variant, obj As Object)
    If TypeOf obj Is IDIAware Then Set obj.di = Me
    mySharedIDs(ID) = 1
    Set myObjects(ID) = obj
End Property

' Resolves an object by ID.
'
' @param Variant ID
' @return Object
Public Property Get Object(ID As Variant) As Object
Attribute Object.VB_UserMemId = 0
    If myObjects.Exists(ID) Then
        Set Object = myObjects(ID)
    Else
        Set Object = ResolveFromProvider(ID)
    End If
End Property

' Checks whether a shared object exists.
'
' @param Variant ID
' @return Boolean
Public Property Get HasObject(ID As Variant) As Boolean
    Let HasObject = myObjects.Exists(ID)
End Property

' Checks whether an object has a provider.
'
' @param Variant ID
' @return Boolean
Public Property Get HasProvider(ID As Variant) As Boolean
    If myProviders.Exists(ID) Or myCallbacks.Exists(ID) Then
        Let HasProvider = True
    Else
        Let HasProvider = False
    End If
End Property

' Checks whether an object is resolvable.
'
' @param Variant ID
' @return Boolean
Public Property Get Has(ID As Variant) As Boolean
    If HasObject(ID) Or HasProvider(ID) Then
        Let Has = True
    Else
        Let Has = False
    End If
End Property

' Returns the number of resolved shared objects.
'
' @return Long
Public Property Get Count() As Long
    Let Count = myObjects.Count
End Property

' Resolves an object by ID.
'
' @alias of Object()
'
' @param Variant ID
' @return Object
Public Function Resolve(ID As Variant) As Object
    Set Resolve = Object(ID)
End Function

' Checks whether the container has the given object.
'
' @param Object obj
' @return Boolean
Public Function Contains(obj As Object) As Boolean
    Dim Key As Variant
    For Each Key In myObjects.Keys
        If myObjects(Key) Is obj Then
            Let Contains = True
            Exit Function
        End If
    Next
    Let Contains = False
End Function

' Finds the container's index/key for the given object.
'
' @param Object obj
' @return Variant|Null
Public Function IndexOf(obj As Object) As Variant
    Dim Key As Variant
    For Each Key In myObjects.Keys
        If myObjects(Key) Is obj Then
            Let IndexOf = Key
            Exit Function
        End If
    Next
    Let IndexOf = Null
End Function

' Designates an object as shared.
'
' @param Variant ObjectOrID The object or ID.
Public Sub Share(ObjectOrID As Variant)
    If VBA.IsObject(ObjectOrID) Then
        Dim Key As Variant
        For Each Key In myObjects.Keys
            If myObjects(Key) Is ObjectOrID Then
                mySharedIDs(Key) = 1
                Exit For
            End If
        Next
    Else
        mySharedIDs(ObjectOrID) = 1
    End If
End Sub

' Registers an IDependencyProvider object.
'
' @param IDependencyProvider Provider
Public Sub Register(Provider As IDependencyProvider)
    
    Dim ID As Variant:      ID = Provider.ProvidedID
    Dim Share As Variant:   Share = Provider.Share
    
    If IsArray(ID) Then
        Dim i As Long
        For i = LBound(ID) To UBound(ID)
            Set myProviders(ID(i)) = Provider
            If Share = True Then mySharedIDs(ID(i)) = 1
        Next
    Else
        Set myProviders(ID) = Provider
        If Share = True Then mySharedIDs(ID) = 1
    End If
    
    If IsArray(Share) Then
        Dim j As Long
        For j = LBound(Share) To UBound(Share)
            mySharedIDs(Share(j)) = 1
        Next
    End If
    
End Sub

' Registers a callback to provide an object.
'
' @param Variant(Of String|Array) ID
' @param Variant(Of Object|Array|String) Func
' @param VbCallType CallType [Optional] Defaults to VbMethod.
Public Sub RegisterCallback( _
    ID As Variant, _
    Func As Variant, _
    Optional CallType As VbCallType = VbMethod)
    
    Dim cbObj As Callback
    If TypeOf Func Is Callback Then
        Set cbObj = Func
    Else
        Set cbObj = Callback.Create(Func, CallType)
    End If
    
    If IsArray(ID) Then
        Dim i As Long
        For i = LBound(ID) To UBound(ID)
            Set myCallbacks(ID(i)) = cbObj
        Next
    Else
        Set myCallbacks(ID) = cbObj
    End If
        
End Sub


' Resolves an object from a provider.
'
' @param Variant ID
' @return Object
Private Function ResolveFromProvider(ID As Variant) As Object
    
    Dim obj As Object
    
    If myProviders.Exists(ID) Then
        Set obj = myProviders(ID).Resolve(ID)
    ElseIf myCallbacks.Exists(ID) Then
        Set obj = myCallbacks(ID).Exec(ID)
    Else
        GoTo Escape
    End If
    
    If TypeOf obj Is IDIAware Then Set obj.di = Me
    If mySharedIDs.Exists(ID) Then Set myObjects(ID) = obj
    
Escape:
    Set ResolveFromProvider = obj

End Function

' Implements ICountable
Private Property Get ICountable_Count() As Long
    Let ICountable_Count = Count
End Property
