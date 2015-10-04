Attribute VB_Name = "stdLib"
Option Explicit

' Checks whether x is an array.
'
' @param Variant x
' @return Boolean
Public Function IsArray(ByVal x As Variant) As Boolean
    Let IsArray = (VBA.TypeName(x) Like "*()")
End Function

' Checks whether a variable is a string.
'
' @param Variant x
' @return Boolean
Public Function IsString(ByVal x As Variant) As Boolean
    Let IsString = VBA.VarType(x) = vbString
End Function

' Checks whether a variable is an integer.
'
' @param Variant(Of Numeric Or Date) x
' @return Boolean
Public Function IsInt(ByVal x As Variant) As Boolean
    If VBA.IsDate(x) Then x = CDbl(x)
    If VBA.IsNumeric(x) Then
        Let IsInt = x = VBA.Fix(x)
    Else
        Let IsInt = False
    End If
End Function

' Checks whether a given directory exists.
'
' @param String Path
' @return Boolean
Public Function IsDir(ByVal Path As String) As Boolean
    Let IsDir = VBA.Len(VBA.Dir(Path, vbDirectory)) > 0
End Function

' Checks whether a given file exists.
'
' @param String Path
' @return Boolean
Public Function IsFile(ByVal Path As String) As Boolean
    Let IsFile = VBA.Len(VBA.Dir(Path)) > 0
End Function

' Checks whether a procedure is callable.
'
' If an object is given as the first argument and no method is given in the second argument,
' the function checks for the existance of a "Invoke" method on the object.
'
' @param Variant(Of String|Object) Callback: A user-defined function as a string, or an object.
' @param String Method [Optional]: The method to call on the Callback object.
' @return Boolean
Public Function IsCallable(ByVal Callback As Variant, Optional Method As String) As Boolean
    
    On Error GoTo ErrHandler
    
    If IsArray(Callback) Then
        Call VBA.CallByName(Callback(LBound(Callback)), Callback(UBound(Callback)), VbMethod)
    ElseIf VBA.IsObject(Callback) Then
        If Method = vbNullString Then Method = "Invoke"
        Call VBA.CallByName(Callback, Method, VbMethod)
    ElseIf VBA.VarType(Callback) = vbString Then
        Call Application.Run(Callback)
    Else
        Let IsCallable = False
        GoTo Escape
    End If
    
    Let IsCallable = True
    GoTo Escape
    
ErrHandler:
    If Err.Number = 449 Then
        Let IsCallable = True
    Else
        Let IsCallable = False
    End If

Escape:
    On Error GoTo 0
End Function

' Checks whether the object has the specified property.
'
' The propety must be a publicly accessible "Get" property.
'
' @param Object Object
' @param String Property
' @return Boolean
Public Function PropertyExists(Object As Object, ByVal Property As String) As Boolean
    
    On Error GoTo ErrHandler
    Call VBA.CallByName(Object, Property, VbGet)
    Let PropertyExists = True
    GoTo Escape
    
ErrHandler:
    If VBA.Err.Number = 449 Then
        Let PropertyExists = True
    Else
        Let PropertyExists = False
        If VBA.Err.Number <> 438 Then Dev.Report Err
    End If

Escape:
    On Error GoTo 0
End Function

' Checks whether the object has the specified method.
'
' @param Object Object
' @param String Method
' @return Boolean
Public Function MethodExists(Object As Object, ByVal Method As String) As Boolean
    
    On Error GoTo ErrHandler
    Call VBA.CallByName(Object, Method, VbMethod)
    Let MethodExists = True
    GoTo Escape

ErrHandler:
    If VBA.Err.Number = 449 Then
        Let MethodExists = True
    Else
        Let MethodExists = False
    End If

Escape:
    On Error GoTo 0
End Function

' Checks whether a variable is Empty, Null, Nothing, an empty array, or empty string.
'
' @param Variant x
' @return Boolean True if variable is empty (as defined above), otherwise False.
Public Function VarIsEmpty(ByVal x As Variant) As Boolean
    Select Case VBA.VarType(x)
        Case vbEmpty, vbNull
            Let VarIsEmpty = True
        Case vbObject
            Let VarIsEmpty = x Is Nothing
        Case vbArray
            Let VarIsEmpty = Arrays.Count(x) < 1
        Case Else
            Let VarIsEmpty = x = vbNullString
    End Select
End Function

' Returns the count of an array, object, or string.
'
' @param Variant x
' @return Long
Public Function Count(ByVal x As Variant) As Long
    On Error GoTo ErrHandler
    If IsArray(x) Then
        Count = UBound(x) - LBound(x) + 1
    ElseIf VBA.IsObject(x) Then
        Count = x.Count
    Else
        Count = VBA.Len(x)
    End If
    GoTo Escape
ErrHandler:
    Let Count = -1
Escape:
    On Error GoTo 0
End Function

' Applies a callback function to the given array or object.
'
' @param Variant(Of Array|Object) x
' @param String func
' @return Variant
Public Function Map(x As Variant, ByVal Procedure As String) As Variant
    If IsArray(x) Then
        Let Map = Arrays.Map(x, Procedure)
    ElseIf VBA.IsObject(x) Then
        Set Map = ObjectMap(x, Procedure)
    Else
        ERR_INVALID_ARG "VBAx.stdLib.Map", "Array|Object", VBA.TypeName(x)
    End If
End Function

' Applies a callback function to the given object.
'
' @param Variant(Of Object) obj
' @param String Procedure
' @return Object(Of Collection|Scripting.Dictionary)
Public Function ObjectMap(obj As Variant, ByVal Procedure As String) As Object
    
    If Not VBA.IsObject(obj) Then GoTo InvalidArgErr
    
    If Not IsCallable(Procedure) Then
        ERR_NOT_CALLABLE "VBAx.stdLib.ObjectMap", Procedure
        Exit Function
    End If
    
    Dim rtn As Object
    Dim Item As Variant
    
    If TypeOf obj Is Dictionary Then
        Set rtn = New Dictionary
        For Each Item In obj.Keys
            rtn.Add Item, Application.Run(Procedure, obj(Item))
        Next
    Else
        Set rtn = New VBA.Collection
        For Each Item In obj
            rtn.Add Application.Run(Procedure, Item)
        Next
    End If
    
    Set ObjectMap = rtn
    Exit Function

InvalidArgErr:
    ERR_INVALID_ARG "VBAx.stdLib.ObjectMap", "Object(Of Collection|Dictionary)", VBA.TypeName(obj)
    
End Function

' Casts a variable to a string.
'
' @param Variant x
' @return String
Public Function StrVal(ByVal x As Variant) As String
    
    If VBA.IsObject(x) Then
        On Error GoTo Err438
        StrVal = x.toString()
        On Error GoTo 0
    Else
        StrVal = x
    End If
    GoTo Escape
    
Err438:
    Dim e As VBA.ErrObject: Set e = Err
    If e.Number = 438 Then
        ' "Unsupported property or method"
        StrVal = VBA.TypeName(x): Resume Next
    Else
        ReRaise e
    End If
Escape:
End Function

' Assigns a value to a variable.
'
' @param Variant var
' @param Variant Value
Public Sub Assign(ByRef Var As Variant, Value As Variant)
    If VBA.IsObject(Value) Then
        Set Var = Value
    Else
        Let Var = Value
    End If
End Sub

' Calls a user callback.
'
' @param Variant Func
' @param Variant ...
' @return Variant
Public Function CallFunc(ByVal Func As Variant, ParamArray args() As Variant) As Variant
    Dim argsArr() As Variant: Let argsArr = args
    Dim cb As Callback
    Set cb = Callback.Create(Func)
    Assign CallFunc, cb.ExecArray(argsArr)
    Set cb = Nothing
End Function

' Calls a user callback with an array of arguments.
'
' @param Variant Func
' @param Variant(Of Array) args
' @return Variant
Public Function CallFuncArray(ByVal Func As Variant, args As Variant) As Variant
    Dim argsArr() As Variant: Let argsArr = args
    Dim cb As Callback
    Set cb = Callback.Create(Func)
    Assign CallFuncArray, cb.ExecArray(argsArr)
    Set cb = Nothing
End Function

' @param x As Variant(Of T)
' @param y As Variant(Of T)
' @return As Variant(Of Boolean Or Null Or Empty)
Public Function Equals(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim xIsObj As Boolean: xIsObj = VBA.IsObject(x)
    Dim yIsObj As Boolean: yIsObj = VBA.IsObject(y)
    If xIsObj Xor yIsObj Then
        Let Equals = Empty
    ElseIf xIsObj And yIsObj Then
        Let Equals = x.Equals(y)
    Else
        If VBA.TypeName(x) = VBA.TypeName(y) Then
            Let Equals = x = y
        ElseIf VBA.IsNull(x) Or VBA.IsNull(y) Then
            Let Equals = Null
        Else
            Let Equals = Empty
        End If
    End If
End Function

' Evaluates a string command and returns the result.
'
' @param String cmd
' @return Variant
Public Function Eval(ByVal cmd As String) As Variant

    On Error GoTo ObjectError
    Let Eval = Application.Evaluate(cmd)
    GoTo Escape

ObjectError:
    Dim e As VBA.ErrObject: Set e = VBA.Err
    If e.Number = 91 Then
        ' "Object variable or With block variable not set."
        Set Eval = Application.Evaluate(cmd): Resume Next
    Else
        ReRaise e
    End If

Escape:
    On Error GoTo 0
End Function

' Evaluates a comparison between the two given values and returns boolean.
'
' @param Variant Value1
' @param Variant Value2
' @param String Comparison [Optional] Default = "="
' @return Boolean
Public Function Compare(ByVal Value1 As Variant, ByVal Value2 As Variant, Optional ByVal Comparison As String = "=") As Boolean
    Let Compare = Application.Evaluate("IF(" _
        & evalPrepValue(Value1) & " " & VBA.Trim$(Comparison) & " " _
        & evalPrepValue(Value2) & ", TRUE, FALSE)")
End Function

' Throws an error.
'
' @param ErrorEnum Error
' @param String Source
' @param Variant Arg1 [Optional]
' @param Variant Arg2 [Optional]
Public Sub Throw(Error As ErrorEnum, ByVal Source As String, Optional ByVal Arg1 As Variant, Optional ByVal Arg2 As Variant)
    Select Case Error
        Case InvalidArgument
            ERR_INVALID_ARG Source, Arg1, Arg2
        Case BadFunctionCall, InvalidProcedureCall
            ERR_NOT_CALLABLE Source, Arg1, Arg2
        Case Else
            VBA.Err.Raise Error, Source, ErrorEnumMsg(Error) & ": " & Arg1
    End Select
End Sub

' Re-raises an error object.
'
' @param VBA.ErrObject e
Public Sub ReRaise(e As VBA.ErrObject)
    VBA.Err.Raise e.Number, e.Source, e.Description, e.HelpFile, e.HelpContext
End Sub

' Returns a message corresponding to a given ErrorEnum.
'
' @param ErrorEnum Error
' @return String
Public Function ErrorEnumMsg(Error As ErrorEnum) As String
    Select Case Error
        Case ErrorEnum.ArgumentNotOptional
            ErrorEnumMsg = "Argument not optional"
        Case ErrorEnum.InvalidArgument
            ErrorEnumMsg = "Invalid argument"
        Case ErrorEnum.InvalidProcedureCall
            ErrorEnumMsg = "Invalid procedure call"
        Case ErrorEnum.BadFunctionCall
            ErrorEnumMsg = "Bad function call"
        Case ErrorEnum.SubFuncNotDefined
            ErrorEnumMsg = "Sub or function not defined"
        Case ErrorEnum.FileNotFound
            ErrorEnumMsg = "File not found"
        Case ErrorEnum.PathNotFound
            ErrorEnumMsg = "Path not found"
        Case ErrorEnum.TypeMismatch
            ErrorEnumMsg = "Type mismatch"
        Case ErrorEnum.PropertyNotFound
            ErrorEnumMsg = "Property not found"
        Case ErrorEnum.UnsupportedPropertyOrMethod
            ErrorEnumMsg = "Unsupported property or method"
        Case ErrorEnum.InvalidPropertyValue
            ErrorEnumMsg = "Invalid property value"
        Case ErrorEnum.ReadOnlyProperty
            ErrorEnumMsg = "Read-only property"
        Case ErrorEnum.WriteOnlyProperty
            ErrorEnumMsg = "Write-only property"
        Case ErrorEnum.ObjectRequired
            ErrorEnumMsg = "Object required"
        Case ErrorEnum.SubscriptOutOfRange
            ErrorEnumMsg = "Subscript out of range"
        Case ErrorEnum.Overflow
            ErrorEnumMsg = "Overflow error"
        Case ErrorEnum.ApplicationOrObjectDefined
            ErrorEnumMsg = "Application or object-defined error"
        Case ErrorEnum.RuntimeError
            ErrorEnumMsg = "Runtime error"
        Case ErrorEnum.LogicError
            ErrorEnumMsg = "Logic error"
        Case Else
            ErrorEnumMsg = "Error"
    End Select
End Function


Private Function evalPrepValue(ByVal Value As Variant) As String
    If VBA.IsNumeric(Value) Then
        evalPrepValue = VBA.Val(Value)
    Else
        evalPrepValue = """" & Value & """"
    End If
End Function

Private Sub ERR_INVALID_ARG( _
    Optional ByVal Source As String, _
    Optional ByVal ExpectedType As String, _
    Optional ByVal GivenType As String)
    
    Dim Message As String: Message = "Invalid argument error"
    If Source <> vbNullString Then _
        Message = Message & " in '" & Source & "'"
    If ExpectedType <> vbNullString Then _
        Message = Message & vbCrLf & "  Expected type: '" & ExpectedType & "'"
    If GivenType <> vbNullString Then _
        Message = Message & vbCrLf & "  Given type: '" & GivenType & "'"
    
    VBA.Err.Raise ErrorEnum.InvalidArgument, Source, Message
    
End Sub

Private Sub ERR_NOT_CALLABLE( _
    Optional ByVal Source As String, _
    Optional ByVal Procedure As Variant, _
    Optional ByVal Method As String)
    
    Dim Msg As String: Msg = "Uncallable procedure"
    If Not VBA.IsMissing(Procedure) Then
        If VBA.IsObject(Procedure) Then
            If Method = vbNullString Then Method = "Invoke"
            Procedure = VBA.TypeName(Procedure) & "." & Method
        End If
        Msg = Msg & ": '" & Procedure & "'"
    End If
    If Source <> vbNullString Then Msg = Msg & " in '" & Source & "'"
    
    VBA.Err.Raise ErrorEnum.InvalidProcedureCall, Source, Msg
    
End Sub
