Set objArgs = WScript.Arguments
If objArgs.Count < 2 Then WScript.Quit
mode = objArgs(0)  ' 模式：folder 或 background
path = objArgs(1)  ' 路径

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

If mode = "folder" Then
    ' 逻辑A：提取选中的那个文件夹
    ExtractFolder path
ElseIf mode = "background" Then
    ' 逻辑B：遍历当前目录下所有的子文件夹并提取
    Set mainFolder = objFSO.GetFolder(path)
    For Each subFolder In mainFolder.SubFolders
        ExtractFolder subFolder.Path
    Next
End If

Sub ExtractFolder(targetPath)
    If Not objFSO.FolderExists(targetPath) Then Exit Sub
    destPath = objFSO.GetParentFolderName(targetPath)
    Set objSource = objShell.NameSpace(targetPath)
    Set objDest = objShell.NameSpace(destPath)
    
    If Not objSource Is Nothing Then
        objDest.MoveHere objSource.Items(), 16 + 1024
    End If
    
    WScript.Sleep 100
    On Error Resume Next
    objFSO.DeleteFolder targetPath, True
End Sub
