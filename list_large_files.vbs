Option Explicit

Const ONE_MB = 1048576 ' 1024 * 1024

Dim targetPath
If WScript.Arguments.Count > 0 Then
    targetPath = WScript.Arguments(0)
Else
    targetPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
End If

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FolderExists(targetPath) Then
    WScript.Echo "Ordner nicht gefunden: " & targetPath
    WScript.Quit 1
End If

Dim rootFolder
Set rootFolder = fso.GetFolder(targetPath)

WScript.Echo "Dateien größer als 1 MB in: " & rootFolder.Path
ProcessFolder rootFolder

Sub ProcessFolder(folder)
    Dim file
    For Each file In folder.Files
        If file.Size > ONE_MB Then
            WScript.Echo file.Path & " (" & FormatSize(file.Size) & ")"
        End If
    Next

    Dim subFolder
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder
    Next
End Sub

Function FormatSize(sizeInBytes)
    Dim sizeInMB
    sizeInMB = sizeInBytes / ONE_MB
    FormatSize = FormatNumber(sizeInMB, 2) & " MB"
End Function
