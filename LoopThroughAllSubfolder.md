Loop Through All Subfolders Using VBA 

```VB
Dim FileSystem As Object
Dim HostFolder As String

HostFolder = "C:\"

Set FileSystem = CreateObject("Scripting.FileSystemObject")
DoFolder FileSystem.GetFolder(HostFolder)

Sub DoFolder(Folder)
    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
    Next
    Dim File
    For Each File In Folder.Files
        ' Operate on each file
    Next
End Sub
```
