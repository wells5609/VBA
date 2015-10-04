VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' FileSystem
' Utilities for the filesystem.

Option Explicit

' Returns the directory/path separator.
'
' On Windows, the separator is "\".
' On Unix/Mac, it is "/".
'
' @return String
Public Property Get Separator() As String
    Separator = Application.PathSeparator
End Property

' The wrong path separator for the current environment.
'
' @return String
Private Property Get WrongSeparator() As String
    If Application.PathSeparator = "\" Then
        WrongSeparator = "\"
    Else
        WrongSeparator = "/"
    End If
End Property

' Checks whether a given directory exists.
'
' Uses native VBA functions.
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

' Checks if a directory (folder) exists.
'
' @param String Path
' @return Boolean
Public Function DirExists(ByVal Path As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    Let DirExists = fso.FolderExists(Path)
    Set fso = Nothing
End Function

' Checks if a file exists.
'
' @param String Path
' @return Boolean
Public Function FileExists(ByVal Path As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    Let FileExists = fso.FileExists(Path)
    Set fso = Nothing
End Function

' Returns the path to the directory containing the given file.
'
' @param String FilePath
' @return String
Public Function DirName(ByVal FilePath As String) As String
    Dim fso As New Scripting.FileSystemObject
    Let DirName = fso.GetParentFolderName(FilePath)
    Set fso = Nothing
End Function

' Returns the file name, with or without extension. Works with directories.
'
' @param String Path
' @param Boolean WithExtension [Optional] Default = True
' @return String
Public Function FileName(ByVal Path As String, Optional WithExtension As Boolean = True) As String
    Dim fso As New Scripting.FileSystemObject
    If Not WithExtension Then
        Let FileName = fso.GetBaseName(Path)
    Else
        Let FileName = fso.GetFileName(Path)
    End If
    Set fso = Nothing
End Function

' Returns the file basename (filename without extension).
'
' For directories, this function is the same as Filename().
'
' @param String Path
' @return String
Public Function BaseName(ByVal Path As String) As String
    Dim fso As New Scripting.FileSystemObject
    Let BaseName = fso.GetBaseName(Path)
    Set fso = Nothing
End Function

' Returns the file extension.
'
' @param String FilePath
' @return String
Public Function FileExtension(ByVal FilePath As String) As String
    Dim fso As New Scripting.FileSystemObject
    Let FileExtension = fso.GetExtensionName(FilePath)
    Set fso = Nothing
End Function

' Returns the path to the directory 1 level up.
'
' @param String Path
' @return String
Public Function DirParent(ByVal Path As String) As String
    Path = StripTrailingSlash(Path)
    Let DirParent = VBA.Left$(Path, VBA.InStrRev(Path, Separator) - 1)
End Function

' Creates a directory (folder), recursively creating parent dirs as necessary.
'
' @param String Path
' @param Scripting.FileSystemObject [Optional]
Public Sub CreateDir(ByVal Path As String, Optional ByVal fso As Scripting.FileSystemObject = Nothing)
    If fso Is Nothing Then Set fso = New Scripting.FileSystemObject
    If Not fso.FolderExists(Path) Then
        CreateDir fso.GetParentFolderName(Path), fso
        fso.CreateFolder Path
    End If
    Set fso = Nothing
End Sub

' Creates a directory if it does not exist.
'
' @param String Path
Public Sub EnsureDirExists(ByVal Path As String)
    If Not IsDir(Path) Then CreateDir Path
End Sub

' Strips trailing slash from a file path.
'
' @param String Path
' @return String
Public Function StripTrailingSlash(ByVal Path As String) As String
    If VBA.Right$(Path, 1) = "\" Or VBA.Right$(Path, 1) = "/" Then
        StripTrailingSlash = VBA.Left$(Path, VBA.Len(Path) - 1)
    Else
        StripTrailingSlash = Path
    End If
End Function

' Adds a trailing slash to a file path if it does not exist.
'
' @param String Path
' @return String
Public Function AddTrailingSlash(ByVal Path As String) As String
    AddTrailingSlash = StripTrailingSlash(Path) & Application.PathSeparator
End Function

' "Normalizes" a file path to use the system directory separator.
'
' Any existing trailing slash will be retained (but converted if necessary).
'
' @param String Path
' @return String
Public Function NormalizePath(ByVal Path As String) As String
    Dim ds As String: ds = Application.PathSeparator
    Dim ws As String: ws = WrongSeparator
    Let NormalizePath = VBA.Replace(VBA.Replace(Path, ws, ds), ds & ds, ds)
End Function

' "Normalizes" a path to use the system directory separator and adds a trailing slash.
'
' @param String Path
' @return String
Public Function NormalizeDirpath(ByVal Path As String) As String
    Let NormalizeDirpath = NormalizePath(AddTrailingSlash(Path))
End Function

' Removes illegal characters from filenames.
'
' @param String Path
' @return String
Public Function SanitizeFileName(ByVal Path As String) As String
    Let SanitizeFileName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Path, "|", vbNullString), ">", vbNullString), "<", vbNullString), Chr(34), vbNullString), "?", vbNullString), "*", vbNullString), ":", vbNullString)
End Function

' Prompts the user to select a workbook and returns an open Excel.Workbook object.
'
' @param String InitialPath [Optional]
' @param String WindowTitle [Optional] Default = "Select a workbook"
' @return Object If user cancels operation, returns Nothing.
Public Function PromptForWorkbook( _
    Optional ByVal InitialPath As String = vbNullString, _
    Optional ByVal WindowTitle As String = "Select a workbook") As Excel.Workbook
    
    Dim FilePath As String
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = WindowTitle
        .InitialFileName = InitialPath
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
        End If
    End With
    
    If Len(FilePath) > 0 Then
        Set PromptForWorkbook = Excel.Workbooks.Open(FileName:=FilePath)
    End If
    
End Function

Public Function PromptForDirectory( _
    Optional ByVal InitialPath As String, _
    Optional ByVal WindowTitle As String = "Select a workbook") As String
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = WindowTitle
        .InitialFileName = InitialPath
        If .Show = -1 Then
            PromptForDirectory = .SelectedItems(1)
        End If
    End With
    Set fd = Nothing
End Function

Public Function PromptForFileType( _
    ByVal FileDescription As String, _
    ByVal FileExtensions As String, _
    Optional ByVal InitialPath As String = vbNullString, _
    Optional ByVal WindowTitle As String = vbNullString) As Variant
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Filters.Add FileDescription, FileExtensions, 1
        If WindowTitle <> vbNullString Then .Title = WindowTitle
        If InitialPath <> vbNullString Then .InitialFileName = InitialPath
        If .Show = -1 Then
            PromptForFileType = .SelectedItems(1)
        End If
    End With
    Set fd = Nothing
End Function
