Attribute VB_Name = "fs_"
Option Explicit

'Returns an array of file names within a directory (not within subdirectories)
'Path qualifier is a path to the directory that may include a filename qualifier such as *.txt
Function getFiles(pathQualifier As String)
    Dim result As String

    Dim file As Variant
    file = Dir(pathQualifier)
    Do While Len(file) > 0
        result = result & "/" & file '/ is delimiter, see https://superuser.com/questions/358855/what-characters-are-safe-in-cross-platform-file-names-for-linux-windows-and-os
        file = Dir
    Loop
    
    'Remove the first "/"
    If Len(result) > 0 Then
        result = Right(result, Len(result) - 1)
    End If
    
    getFiles = Split(result, "/")
End Function
