VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICountable

Private myStartTime As Date
Private myType As ISessionType
Private myVars As Dictionary

Public Event BeforeStart(SessionType As ISessionType, Cancel As Boolean)
Public Event Start(SessionType As ISessionType)
Public Event BeforeDestroy(SessionType As ISessionType, Cancel As Boolean)
Public Event Destroy()

Public Sub Start(Optional obj As ISessionType)
    
    If obj Is Nothing Then Set obj = New SessionTypeDefault
    
    Dim Cancel As Boolean: Cancel = False
    RaiseEvent BeforeStart(obj, Cancel)
    If Cancel Then Exit Sub
    
    If Not myType Is Nothing Then Destroy
    
    Let myStartTime = VBA.Now
    Set myVars = New Dictionary
    Set myType = obj
    myType.Start Me
    
    RaiseEvent Start(myType)
    
End Sub

Public Sub Destroy()
    
    Dim Cancel As Boolean: Cancel = False
    RaiseEvent BeforeDestroy(myType, Cancel)
    If Cancel Then Exit Sub
    
    myType.Destroy Me
    Let myStartTime = Empty
    Set myType = Nothing
    Set myVars = Nothing
    
    RaiseEvent Destroy
    
End Sub

Public Property Get SessionType() As ISessionType
    Set SessionType = myType
End Property

Public Property Get StartTime() As Date
    StartTime = myStartTime
End Property

Public Property Get Var(ByVal Key As String) As Variant
    If VBA.IsObject(myVars(Key)) Then
        Set Var = myVars(Key)
    Else
        Let Var = myVars(Key)
    End If
End Property

Public Property Let Var(ByVal Key As String, Value As Variant)
    Let myVars(Key) = Value
End Property

Public Property Set Var(ByVal Key As String, Value As Object)
    Set myVars(Key) = Value
End Property

Public Property Get Count() As Long
    Count = myVars.Count
End Property

Public Property Get Uptime() As Double
    Uptime = CDbl(VBA.Now) - CDbl(myStartTime)
End Property

Public Function HasVar(ByVal Key As String) As Boolean
    Let HasVar = myVars.Exists(Key)
End Function

Public Sub DeleteVar(ByVal Key As String)
    myVars.Remove Key
End Sub


Private Sub Class_Terminate()
    If Not myType Is Nothing Then Destroy
End Sub

Private Property Get ICountable_Count() As Long
    ICountable_Count = Count
End Property
