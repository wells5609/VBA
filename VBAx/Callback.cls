VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Callback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private myObject As Object
Private myProcedure As String
Private myCallType As VbCallType

' Creates a new Callback
'
' @param Variant(Of Array|Object|String) Func
' @param VbCallType CallType [Optional] Default = VbMethod
' @return Callback
Public Function Create(Func As Variant, Optional CallType As VbCallType = VbMethod) As Callback
    Set Create = New Callback
    Create.Init Func, CallType
End Function

' Invokes the callback.
'
' @param Variant ...
' @return Variant
Public Function Exec(ParamArray args() As Variant) As Variant
Attribute Exec.VB_UserMemId = 0
    Dim argsArr() As Variant: Let argsArr = args
    If myObject Is Nothing Then
        CALL_FUNC_ARGS Exec, argsArr, "Callback.Exec"
    Else
        CALL_OBJECT_ARGS Exec, argsArr, "Callback.Exec"
    End If
End Function

' Invokes the callback with an array of arguments.
'
' @param Variant(Of Array) args
' @return Variant
Public Function ExecArray(args As Variant) As Variant
    Dim argsArr() As Variant: Let argsArr = args
    If myObject Is Nothing Then
        CALL_FUNC_ARGS ExecArray, argsArr, "Callback.ExecArray"
    Else
        CALL_OBJECT_ARGS ExecArray, argsArr, "Callback.ExecArray"
    End If
End Function


Friend Sub Init(Func As Variant, Optional CallType As VbCallType = VbMethod)
    If IsArray(Func) Then
        Set myObject = Func(LBound(Func))
        Let myProcedure = Func(UBound(Func))
    ElseIf VBA.IsObject(Func) Then
        Set myObject = Func
        Let myProcedure = "Invoke"
    Else
        Let myProcedure = Func
    End If
    myCallType = CallType
End Sub

Private Sub CALL_FUNC_ARGS(ByRef AssignTo As Variant, ByRef args() As Variant, ByVal Caller As String)
    
    Select Case UBound(args) + 1
        Case 0
            Assign AssignTo, Application.Run(myProcedure)
        Case 1
            Assign AssignTo, Application.Run(myProcedure, args(0))
        Case 2
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1))
        Case 3
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2))
        Case 4
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2), args(3))
        Case 5
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2), args(3), args(4))
        Case 6
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 7
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 8
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 9
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 10
            Assign AssignTo, Application.Run(myProcedure, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case Else
            Throw InvalidProcedureCall, Caller, myProcedure
    End Select

End Sub

Private Sub CALL_OBJECT_ARGS(ByRef AssignTo As Variant, ByRef args() As Variant, ByVal Caller As String)

    Select Case UBound(args) + 1
        Case 0
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType)
        Case 1
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0))
        Case 2
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1))
        Case 3
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2))
        Case 4
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2), args(3))
        Case 5
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2), args(3), args(4))
        Case 6
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 7
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 8
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 9
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 10
            Assign AssignTo, VBA.CallByName(myObject, myProcedure, myCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case Else
            Throw InvalidProcedureCall, Caller, myObject, myProcedure
    End Select

End Sub

Private Sub Class_Terminate()
    Set myObject = Nothing
End Sub
