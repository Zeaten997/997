Set shell = CreateObject("WScript.Shell")
Set args = WScript.Arguments

' 检查参数是否足够（压缩包路径 + 模式）
If args.Count < 2 Then
    MsgBox "请将压缩包拖拽到此脚本上，并指定模式。", 48, "参数缺失"
    WScript.Quit
End If

zipPath = args(0)
mode = args(1)

' 修正逻辑：1.换行符使用 `n  2.如果目标目录存在则清空内容
psCode = "$zipPath = '" & zipPath & "'; $mode = " & mode & "; " & _
         "$ProgressPreference = 'SilentlyContinue'; " & _
         "try { " & _
         "  $parentDir = Split-Path -Parent $zipPath; " & _
         "  $tempDir = Join-Path $parentDir ('Temp_Work_' + (Get-Date -Format 'HHmmss')); " & _
         "  Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force; " & _
         "  $newFolderName = ''; $files = Get-ChildItem -Path $tempDir -Recurse -File; " & _
         "  foreach ($f in $files) { if ($f.BaseName -match '^(12\.[^-]+)') { $newFolderName = $Matches[1]; break } }; " & _
         "  if (-not $newFolderName) { $newFolderName = (Get-Item $zipPath).BaseName }; " & _
         "  $targetPath = Join-Path $parentDir $newFolderName; " & _
         "  if (Test-Path $targetPath) { Remove-Item (Join-Path $targetPath '*') -Recurse -Force } " & _
         "  else { New-Item -Path $targetPath -ItemType Directory -Force }; " & _
         "  switch ($mode) { " & _
         "    1 { $dwgDir = New-Item -Path (Join-Path $targetPath 'DWG') -ItemType Directory -Force; $pdfDir = New-Item -Path (Join-Path $targetPath 'PDF') -ItemType Directory -Force; foreach ($f in $files) { $ext = $f.Extension.ToLower(); if ($ext -eq '.dwg') { Move-Item $f.FullName $dwgDir.FullName -Force } elseif ($ext -eq '.pdf') { Move-Item $f.FullName $pdfDir.FullName -Force } } } " & _
         "    2 { $files | Where-Object { $_.Extension -eq '.dwg' } | ForEach-Object { Move-Item $_.FullName $targetPath -Force } } " & _
         "    3 { $files | Where-Object { $_.Extension -eq '.pdf' } | ForEach-Object { Move-Item $_.FullName $targetPath -Force } } " & _
         "  }; " & _
         "  if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }; " & _
         "  Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('处理完成！' + [char]10 + '目录：' + $newFolderName, '成果整理') " & _
         "} catch { " & _
         "  Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('出错：' + $_.Exception.Message) " & _
         "}"

' 运行 PowerShell，参数 0 表示隐藏窗口
shell.Run "powershell.exe -ExecutionPolicy Bypass -Command " & Chr(34) & psCode & Chr(34), 0, True
