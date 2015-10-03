VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Type VersionStruct
    maj As Integer
    min As Integer
    pat As Integer
    rel As String
    Str As String
End Type

Private vMajor As Integer
Private vMinor As Integer
Private vPatch As Integer
Private vRelease As String
Private vString As String

' Returns a new VersionObject with the given settings.
'
' @param Integer Major
' @param Integer Minor
' @param Integer Patch
' @param String ReleaseString [Optional]
' @return VersionObject
Public Function Create( _
    ByVal Major As Integer, _
    ByVal Minor As Integer, _
    ByVal Patch As Integer, _
    Optional ByVal ReleaseString As String) As VersionObject
    
    Set Create = New VersionObject
    Create.Init Major, Minor, Patch, ReleaseString
End Function

Public Sub Init( _
    ByVal Major As Integer, _
    ByVal Minor As Integer, _
    ByVal Patch As Integer, _
    Optional ByVal ReleaseString As String)
    
    vMajor = Major
    vMinor = Minor
    vPatch = Patch
    vString = vMajor & "." & vMinor & "." & vPatch
    
    If ReleaseString <> vbNullString Then
        vRelease = VBA.LCase$(ReleaseString)
        vString = vString & "-" & vRelease
    End If
    
End Sub

Public Property Get toString() As String
Attribute toString.VB_UserMemId = 0
    Let toString = vString
End Property

Public Property Get Major() As Integer
    Let Major = vMajor
End Property

Public Property Get Minor() As Integer
    Let Minor = vMinor
End Property

Public Property Get Patch() As Integer
    Let Patch = vPatch
End Property

Public Property Get ReleaseString() As String
    Let ReleaseString = vRelease
End Property

Public Property Get ReleaseCandidate() As Integer
    If IsReleaseCandidate Then
        Let ReleaseCandidate = VBA.Right(vRelease, VBA.Len(vRelease) - 2)
    End If
End Property

Public Property Get IsRelease() As Boolean
    Let IsRelease = VBA.Len(vRelease) > 0
End Property

Public Property Get IsReleaseCandidate() As Boolean
    IsReleaseCandidate = VBA.Left$(vRelease, 2) = "rc"
End Property

Public Property Get IsAlpha() As Boolean
    Let IsAlpha = vRelease = "alpha"
End Property

Public Property Get IsBeta() As Boolean
    Let IsBeta = vRelease = "beta"
End Property

Public Property Get IsStable() As Boolean
    Let IsStable = vRelease = vbNullString
End Property

' Compares a version string
'
' Returns 1 if Version is lower than the current version.
' Returns 0 if the Version is equal to the current version.
' Returns -1 if Version is higher than the current version.
'
' @param String Version
' @return Integer
Public Function Compare(ByVal Version As String) As Integer
    
    If VBA.LCase$(Version) = vString Then
        Compare = 0
        Exit Function
    End If
    
    Dim obj As VersionStruct: obj = parseString(Version)
    
    If obj.maj < vMajor Then
        Compare = 1
    ElseIf obj.maj > vMajor Then
        Compare = -1
    Else
        If obj.min > vMinor Then
            Compare = -1
        ElseIf obj.pat > vPatch Then
            Compare = -1
        ElseIf vRelease = vbNullString Or obj.rel = vRelease Then
            Compare = 1
        Else
            Compare = -1
        End If
    End If
End Function

Public Function IsAtLeast(ByVal Version As String) As Boolean
    IsAtLeast = Compare(Version) >= 0
End Function


Private Function parseString(ByVal Version As String) As VersionStruct
    
    Dim obj As VersionStruct
    obj.Str = Version
    
    Dim relVer As String
    If VBA.InStr(1, Version, "-") > 0 Then
        relVer = VBA.LCase$(VBA.Right$(Version, VBA.Len(Version) - VBA.InStrRev(Version, "-")))
        Version = VBA.Left$(Version, VBA.Len(Version) - VBA.Len(relVer) - 1)
    End If
    
    Dim parts() As String: parts = VBA.Split(Version, ".", 3)
    Dim lBnd As Long: lBnd = LBound(parts)
    
    With obj
        .maj = parts(lBnd)
        .min = parts(lBnd + 1)
        .pat = parts(lBnd + 2)
        .rel = relVer
    End With
    
    parseString = obj
    
End Function
