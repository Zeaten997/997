Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' 在空白处右键时，第二个参数传入当前目录路径
If WScript.Arguments.Count > 0 Then
    currentPath = WScript.Arguments(0)
    
    folder1 = currentPath & "\01+审核"
    folder2 = currentPath & "\02+验证"
    
    If Not fso.FolderExists(folder1) Then fso.CreateFolder(folder1)
    If Not fso.FolderExists(folder2) Then fso.CreateFolder(folder2)
End If
