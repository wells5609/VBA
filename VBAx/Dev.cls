VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

#Const DEVELOPING = True

Public Sub Log(ByVal x As Variant)
#If DEVELOPING Then
    If VBA.VarType(x) = vbString Then
        Debug.Print x
    ElseIf TypeOf x Is VBA.ErrObject Then
        Me.Report x
    Else
        Debug.Print VarDump(x)
    End If
#End If
End Sub

Public Sub Dump(ByVal x As Variant)
#If DEVELOPING Then
    Debug.Print VarDump(x)
#End If
End Sub

Public Sub Report(ByVal e As VBA.ErrObject)
#If DEVELOPING Then
    Debug.Print "[" & Format(Now, "hh:mm:ss") & "] Error #" & e.Number & ": " & e.Description & " in " & e.Source
#Else
    Err.Raise e.Number, e.Source, e.Description, e.HelpFile, e.HelpContext
#End If
End Sub

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
        VBA.Err.Raise e.Number, e.Source, e.Description, e.HelpFile, e.HelpContext
    End If
Escape:
End Function

' Dumps a variable to a string.
'
' @param x As Variant
' @return As String
Public Function VarDump(ByVal x As Variant) As String

    If VBA.IsObject(x) Then
        If x Is Nothing Then
            VarDump = "[Nothing]"
        Else
            VarDump = "[Object] " & StrVal(x)
        End If
        GoTo Escape
    End If
    
    Dim ty As String: ty = VBA.TypeName(x)
    
    Select Case ty
        Case "Boolean", "Integer", "Long", "Single", "Double", "Currency", "Byte", "Decimal"
            VarDump = "[" & ty & "] " & x
        #If VBA7 And Win64 Then
        Case "LongLong"
            VarDump = "[LongLong] " & x
        #End If
        Case "Date"
            Dim dtStr As String, tmStr As String
            If Abs(x) >= 1 Then dtStr = VBA.Month(x) & "/" & VBA.Day(x) & "/" & VBA.Year(x)
            If x <> VBA.Fix(x) Then tmStr = VBA.Format$(x, "h:nn:ss AM/PM")
            VarDump = "[Date] " & VBA.Trim$(dtStr & " " & tmStr)
        Case "String"
            If x = vbNullString Then
                VarDump = "[vbNullString]"
            Else
                VarDump = "[String] """ & x & """"
            End If
        Case "Empty", "Null", "Nothing"
            VarDump = "[" & ty & "]"
        Case "Error"
            If VBA.IsMissing(x) Then
                VarDump = "[Missing]"
            Else
                VarDump = "[CVErr] " & CStr(x)
            End If
        Case "ErrObject"
            VarDump = "[ErrObject] #" & x.Number
        Case "Unknown"
            VarDump = ty
        Case Else
            If Not IsArray(x) Then
                VarDump = ""
                GoTo Escape
            End If
            Dim rnk As Integer: rnk = VBAx.Arrays.Rank(x)
            If rnk = 1 Then
                Dim lBnd As Long: lBnd = LBound(x)
                Dim uBnd As Long: uBnd = UBound(x)
                Dim arr As Variant
                If uBnd - lBnd < 0 Then
                    arr = Array()
                Else
                    Dim mx As Long: mx = 8 - 1
                    Dim xb As Long: xb = VBA.IIf(uBnd - lBnd < mx, uBnd, lBnd + mx)
                    ReDim arr(lBnd To xb)
                    Dim i As Long
                    For i = lBnd To xb: arr(i) = VarDump(x(i)): Next
                End If
                VarDump = "[Array] (" & VBA.Join(arr, ", ") & VBA.IIf(xb < uBnd, ", ...", "") & ")"
            Else
                VarDump = "[Array] " & VBA.Replace(ty, "()", "(" & VBA.String(rnk - 1, ",") & ")")
            End If
    End Select
    
Escape:
End Function
