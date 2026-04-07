Set shell = CreateObject("WScript.Shell")
Set args = WScript.Arguments

If args.Count < 1 Then WScript.Quit
targetPath = args(0)

' 1. 安全确认
pwd = InputBox("警告：将删除多余文件，仅保留空目录及特定目录表。" & vbCrLf & vbCrLf & "确认密码：", "997工具")
If pwd <> "997" Then WScript.Quit

' 2. 逻辑：
' A. 找出所有文件。
' B. 找出需要保留的文件（文件名含关键字，或者是目录下唯一的非Excel文件）。
' C. 差集计算：[所有文件] 减去 [保留文件] = [待删除文件]。
psCmd = "powershell.exe -ExecutionPolicy Bypass -Command "" " & _
    "$all = Get-ChildItem -Path '" & targetPath & "' -Recurse -File -Force; " & _
    "$keep = $all | Where-Object { " & _
    "    $f = $_; " & _
    "    $p = $f.DirectoryName; " & _
    "    $isKey = ($f.Name -like '*图纸目录*' -or $f.Name -like '*图纸总目录*'); " & _
    "    $hasOtherTypes = (Get-ChildItem -Path $p -File -Force | Where-Object { $_.Extension -notlike '.xls*' }); " & _
    "    $isLoneExcel = ($f.Extension -like '.xls*' -and -not $hasOtherTypes); " & _
    "    $isKey -or $isLoneExcel " & _
    "}; " & _
    "$toDelete = $all | Where-Object { $_.FullName -notin $keep.FullName }; " & _
    "$toDelete | Remove-Item -Force -ErrorAction SilentlyContinue """

' 运行（0隐藏窗口，True同步执行）
shell.Run psCmd, 0, True

MsgBox "清理完成！" & vbCrLf & "已清空所有文件，仅保留目录结构。", 64, "执行成功"
