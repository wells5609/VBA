Option Explicit

' REQUIRED REFERENCES:
'  1) Microsoft Scripting Runtime

' Checks if a file exists.
Function FileExists(ByVal Path As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    FileExists = fso.FileExists(Path)
    Set fso = Nothing
End Function

' Checks if a directory (folder) exists.
Function DirExists(ByVal Path As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    DirExists = fso.FolderExists(Path)
    Set fso = Nothing
End Function

' Creates a directory (folder), recursively creating parent dirs as necessary.
Sub CreateDir(ByVal Path As String, Optional fso As Scripting.FileSystemObject)
    If IsMissing(fso) Then Set fso = New Scripting.FileSystemObject
    If Not fso.FolderExists(Path) Then
        CreateDir fso.GetParentFolderName(Path), fso
        fso.CreateFolder (Path)
    End If
End Sub

' "Normalizes" a directory path to use backslashes.
Function NormalizeDirpath(ByVal Path As String) As String
    Path = Replace(Path, "/", "\")
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    NormalizeDirpath = Replace(Path, "\\", "\")
End Function
