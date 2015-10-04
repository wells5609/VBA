VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Strings
' Utilities for strings.

Option Explicit

Private Const QS = """"

' Wraps a string in double quotes.
'
' @param String str
' @return String
Public Function Quote(ByVal Str As String) As String
    Let Quote = QS & Str & QS
End Function

' Joins an array of strings by delimiter.
'
' @param Variant Strings()
' @param String Delimiter [Optional] Default = ","
' @return String
Public Function Join(Strings() As Variant, Optional Delimiter As String = ",") As String
    Let Join = VBA.Join(Strings, Delimiter)
End Function

' Changes a string to uppercase.
'
' @param String str
' @return String
Public Function toUpper(ByVal Str As String) As String
    Let toUpper = VBA.UCase$(Str)
End Function

' Changes a string to lowercase.
'
' @param String str
' @return String
Public Function toLower(ByVal Str As String) As String
    Let toLower = VBA.LCase$(Str)
End Function

' Returns the portion of a starting at position Start and optionally trimmed to Length number of characters.
'
' @param String str
' @param Long Start
' @param Length Long [Optional]
' @return String
Public Function Substr( _
    ByVal Str As String, _
    ByVal Start As Long, _
    Optional ByVal Length As Long) As String
    
    If Start < 0 Then
        Substr = VBA.Right$(Str, VBA.Abs(Start))
    Else
        Substr = VBA.Right$(Str, VBA.Len(Str) - Start)
    End If
    
    If Length <> 0 And VBA.Len(Substr) > Length Then
        Substr = VBA.Left$(Substr, Length)
    End If
    
End Function

' Replaces substrings within a string.
'
' @param String str
' @param String Find
' @param String Repl
' @return String
Public Function Replace( _
    ByVal Str As String, _
    ByVal Find As String, _
    ByVal Repl As String) As String
    Let Replace = VBA.Replace(Str, Find, Repl)
End Function

' Replaces several substrings within a string.
'
' @param String Str
' @param Variant(Of Array) Find()
' @param Variant Repl
' @return String
Public Function ReplaceArray( _
    ByVal Str As String, _
    Find() As Variant, _
    Repl As Variant) As String
    
    Dim f As Variant
    If IsArray(Repl) Then
        Dim r As Long: r = UBound(Repl)
        Dim i As Long
        For i = LBound(Find) To UBound(Find)
            If i > r Then
                Str = VBA.Replace(Str, CStr(Find(i)), "")
            Else
                Str = VBA.Replace(Str, CStr(Find(i)), CStr(Repl(i)))
            End If
        Next
    Else
        For Each f In Find: Str = VBA.Replace(Str, CStr(f), Repl): Next
    End If
    
    ReplaceArray = Str
     
End Function

' Concatenates an array of strings by reference.
'
' @param String ByRef str
' @ParamArray Strings()
' @return String
Public Function Concat(ByRef Str As String, ParamArray Strings() As Variant) As String
    Dim v As Variant
    For Each v In Strings: Str = Str & CStr(v): Next
    Let Concat = Str
End Function

' Determines whether a string starts with a given prefix.
'
' @param String str
' @param String Find
' @param Boolean CaseSensitive [Optional] Default = True
' @return Boolean
Public Function StartsWith( _
    ByVal Str As String, _
    ByVal Find As String, _
    Optional CaseSensitive As Boolean = True) As Boolean
    
    If CaseSensitive Then
        Let StartsWith = (VBA.Left$(Str, VBA.Len(Find)) = Find)
    Else
        Let StartsWith = (VBA.Left$(VBA.LCase$(Str), VBA.Len(Find)) = VBA.LCase$(Find))
    End If
End Function

' Determines whether a string ends with a given suffix.
'
' @param String str
' @param String Find
' @param Boolean CaseSensitive [Optional] Default = True
' @return Boolean
Public Function EndsWith( _
    ByVal Str As String, _
    ByVal Find As String, _
    Optional CaseSensitive As Boolean = True) As Boolean
    
    If CaseSensitive Then
        Let EndsWith = (VBA.Right$(Str, VBA.Len(Find)) = Find)
    Else
        Let EndsWith = (VBA.Right$(VBA.LCase$(Str), VBA.Len(Find)) = VBA.LCase$(Find))
    End If
End Function

' Trims a specified set of characters from the beginning of the given string.
'
' @param String str
' @param String Chars
' @return String
Public Function LTrimChars(ByVal Str As String, ByVal Chars As String) As String
    If Str = vbNullString Then
        Let LTrimChars = vbNullString
        Exit Function
    End If
    Dim i As Integer: i = 1
    While VBA.InStr(Chars, VBA.Mid$(Str, i, 1)) > 0 And i <= VBA.Len(Str): i = i + 1: Wend
    Let LTrimChars = VBA.Mid$(Str, i)
End Function

' Trims a specified set of characters from the end of the given string.
'
' @param String str
' @param String Chars
' @return String
Public Function RTrimChars(ByVal Str As String, ByVal Chars As String) As String
    If Str = vbNullString Then
        Let RTrimChars = vbNullString
        Exit Function
    End If
    Dim i As Integer: i = VBA.Len(Str)
    While VBA.InStr(Chars, VBA.Mid$(Str, i, 1)) > 0 And i >= 1: i = i - 1: Wend
    Let RTrimChars = VBA.Left$(Str, i)
End Function

' Trims a specified set of characters from the beginning and end of a string.
'
' @param String str
' @param String Chars
' @return String
Public Function TrimChars(ByVal Str As String, ByVal Chars As String) As String
    If Str = vbNullString Then
        Let TrimChars = vbNullString
        Exit Function
    End If
    Dim i As Integer: i = 1
    While VBA.InStr(Chars, VBA.Mid$(Str, i, 1)) > 0 And i <= VBA.Len(Str): i = i + 1: Wend
    Str = VBA.Mid$(Str, i): i = VBA.Len(Str)
    While VBA.InStr(Chars, VBA.Mid$(Str, i, 1)) > 0 And i >= 1: i = i - 1: Wend
    Let TrimChars = VBA.Left$(Str, i)
End Function

' Checks whether a given string is found within a concatenated list of strings.
'
' @param String Find
' @param String List
' @param String Separator [Optional] Default = "|"
' @return Boolean
Public Function InStrList( _
    ByVal Find As String, _
    ByVal List As String, _
    Optional ByVal Separator As String = "|") As Boolean
    
    Dim Value As Variant
    For Each Value In VBA.Split(List, Separator)
        If Find = VBA.Trim$(Value) Then
           Let InStrList = True
           Exit Function
        End If
    Next
    Let InStrList = False
End Function

' Parses an argument string to a Scripting.Dictionary object.
'
' @param String str
' @param String Separator [Optional] Default = ";"
' @return Scripting.Dictionary
Public Function Parse(ByVal Str As String, Optional ByVal Separator As String = ";") As Scripting.Dictionary
    Set Parse = New Scripting.Dictionary
    Dim part As Variant
    Dim parts() As String
    For Each part In VBA.Split(Str, Separator)
        Let parts = VBA.Split(part, "=")
        Parse.Add VBA.Trim(parts(0)), Trim(parts(1))
    Next
End Function

' Splits a string into an array of characters.
'
' @param String str
' @return Array(Of String)
Public Function SplitChars(ByVal Str As String) As String()
    Dim arr() As String
    ReDim arr(VBA.Len(Str) - 1)
    Dim i As Long
    For i = 1 To VBA.Len(Str): arr(i - 1) = VBA.Mid$(Str, i, 1): Next
    Let SplitChars = arr
End Function
