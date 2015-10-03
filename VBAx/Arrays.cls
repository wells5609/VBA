VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Arrays
' Utilities for Arrays.

Option Explicit

' Returns the number of items in an array.
'
' @param Variant arr
' @return Long
' @throws 574
Public Function Count(arr As Variant) As Long
    
    If Not IsArray(arr) Then
        Throw InvalidArgument, "Arrays.Count", "Array", VBA.TypeName(arr)
        Exit Function
    End If
    
    On Error Resume Next
    Count = UBound(arr) - LBound(arr) + 1
    On Error GoTo 0
    
End Function

' Returns the first item in an array.
'
' @param Variant arr
' @return Variant
' @throws 574
Public Function First(arr As Variant) As Variant

    If Not IsArray(arr) Then
        Throw InvalidArgument, "Arrays.First", "Array", VBA.TypeName(arr)
        Exit Function
    End If
    
    Dim i As Long: i = LBound(arr)
    If VBA.IsObject(arr(i)) Then
        Set First = arr(i)
    Else
        First = arr(i)
    End If

End Function

' Returns the last item in an array.

' @param Variant arr
' @return Variant
' @throws 574
Public Function Last(arr As Variant) As Variant

    If Not IsArray(arr) Then
        Throw InvalidArgument, "Arrays.Last", "Array", VBA.TypeName(arr)
        Exit Function
    End If
    
    Dim i As Long: i = UBound(arr)
    If VBA.IsObject(arr(i)) Then
        Set Last = arr(i)
    Else
        Last = arr(i)
    End If
    
End Function

' Whether a given item exists in an array.

' @param Variant needle
' @param Variant haystack
' @return Boolean
' @throws 574
Public Function IsIn(ByVal Find As Variant, arr As Variant) As Boolean

    If Not IsArray(arr) Then
        Throw InvalidArgument, "Arrays.IsIn", "Array", VBA.TypeName(arr)
        Exit Function
    End If
    
    Let IsIn = False
    
    Dim v As Variant
    For Each v In arr
        If Find = v Then
           Let IsIn = True
           Exit Function
        End If
    Next

End Function

' Applies a callback function to each item in an array.

' @param Variant arr
' @param Variant(Of String|Object) Callback
' @param String method [Optional] Method of Callback object to call.
' @return Variant
' @throws 574
Public Function Map(arr As Variant, ByVal Callback As Variant, Optional ByVal Method As String) As Variant
    
    If Not IsArray(arr) Then
        Throw InvalidArgument, "Arrays.Map", "Array", VBA.TypeName(arr)
        Exit Function
    End If
    
    'If Not IsCallable(Callback, Method) Then
    '    Call ERR_NOT_CALLABLE("VBAUtilities.ArrayMap", Callback, Method)
    '    Exit Function
    'End If
        
    Dim i As Long: i = LBound(arr)
    Dim Values As Variant: Values = Array()
    ReDim Values(i To UBound(arr))
    Dim v As Variant
    
    If VBA.IsObject(Callback) Then
        If Method = vbNullString Then Method = "Invoke"
        For Each v In arr
            Values(i) = VBA.CallByName(Callback, Method, VbMethod, v)
            i = i + 1
        Next
    Else
        For Each v In arr
            Values(i) = Application.Run(Callback, v)
            i = i + 1
        Next
    End If
    
    Let Map = Values
    
End Function

' Returns the index of the given value in the given array, or one less than
' the lower bound of the array if the value is not found in the array.
'
' @param Variant arr The array to search through.
' @param Variant Find The value to search for.
' @return Long
' @throws 574
Public Function IndexOf(arr As Variant, Find As Variant) As Long

    If Not IsArray(arr) Then
        Throw InvalidArgument, "Arrays.IndexOf", "Array", VBA.TypeName(arr)
        Exit Function
    End If
    
    Let IndexOf = LBound(arr) - 1
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = Find Then
            Let IndexOf = i
            Exit Function
        End If
    Next

End Function

' Returns an array of unique values from an array.
'
' @param Array arr
' @return Array
' @throws 574
Public Function Unique(arr As Variant) As Variant()
    
    If Not IsArray(arr) Then
        Throw InvalidArgument, "Arrays.Unique", "Array", VBA.TypeName(arr)
        Exit Function
    End If
    
    Dim tmp As New Scripting.Dictionary
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Not tmp.Exists(arr(i)) Then
            tmp.Add arr(i), i
        End If
    Next
    
    Let Unique = tmp.Keys
    Set tmp = Nothing
    
End Function

' @param Variant(Of Array(Of T)) arr
' @return Integer
Public Function Rank(ByVal arr As Variant) As Integer
    If Not IsArray(arr) Then Err.Raise 13
    Dim x As Long
    Dim i As Integer: i = 0
    On Error Resume Next
    While VBA.Err.Number = 0
        i = i + 1
        x = UBound(arr, i)
    Wend
    Let Rank = i - 1
    On Error GoTo 0
End Function

' Converts an array to a Collection.
'
' @param Array arr
' @return Collection
Public Function ToCollection(arr() As Variant) As Collection
    Set ToCollection = New Collection
    Dim i As Long
    For i = LBound(arr) To UBound(arr): ToCollection.Add arr(i), i: Next
End Function

' Converts a Collection to an array of strings.
'
' @param Collection coll
' @return Array(Of String)
Public Function FromStringCollection(coll As Collection) As String()
    Dim i As Integer: i = 1
    Dim obj As Variant
    Dim arr() As String
    ReDim arr(1 To coll.Count)
    For Each obj In coll
        arr(i) = obj
        i = i + 1
    Next
    Let FromStringCollection = arr
End Function
