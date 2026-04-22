Attribute VB_Name = "modTrackerLogic"
'====================================================================
' modTrackerLogic.bas
' Excel VBA 模块：修改追踪启动停止逻辑 + 业务逻辑实现
' 说明
'   1. 追踪启动逻辑 (StartTrackerLogic)：初始化环境、备份原始数据至隐藏表、激活监听
'   2. 追踪停止逻辑 (StopTrackerLogic)：关闭监听开关并更新界面状态
'   3. 标记删除功能 (ApplyDeleteMark)：对选定区域执行删除线及红色标记
'   4. 手动标记功能 (ApplyManualMark)：对选定区域应用红字黄底并写入日期
'   5. 清除功能：清除格式、标记等
'   6. 辅助工具函数：自动识别或创建"修改日期"列，执行日期写入逻辑
'====================================================================

Option Explicit

Public myRibbon As IRibbonUI           ' 引用 Ribbon UI 对象，用于刷新按钮状态
Public bIsTracking As Boolean          ' 全局布尔值：当前是否处于修改追踪模式
Public EventWatcher As clsAppEvents   ' 类模块实例：真正Excel 事件监听

Private Const HiddenLogPrefix As String = "HiddenLog_"
Private Const MaxTargetCells As Long = 10000

' Sub: StartTrackerLogic
' 功能：系统的启动仪式
Sub StartTrackerLogic()
    Dim ws As Worksheet
    Dim hiddenSh As Worksheet
    Dim hiddenShName As String
    Dim headers As Variant, h As Variant
    Dim cell As Range

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ' 1. 安全校验：不能在日志表上启动
    If ws.Name Like HiddenLogPrefix & "*" Then
        MsgBox "无法在系统日志表上启动功能！", vbExclamation
        Exit Sub
    End If

    hiddenShName = HiddenLogPrefix & ws.Name

    ' 2. 强制删除旧的备份
    Application.DisplayAlerts = False
    On Error Resume Next
    Set hiddenSh = ws.Parent.Worksheets(hiddenShName)
    Do While Not hiddenSh Is Nothing
        hiddenSh.Visible = xlSheetVisible
        hiddenSh.Delete
        Set hiddenSh = Nothing
        Set hiddenSh = ws.Parent.Worksheets(hiddenShName)
    Loop
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' [注：原有的旧标记检测和提示逻辑已脱离至独立的按钮中]

    ' 3. 创建原始数据快照"隐藏
    ws.Copy After:=ws.Parent.Sheets(ws.Parent.Sheets.Count)
    Set hiddenSh = ActiveSheet

    On Error Resume Next
    hiddenSh.Name = hiddenShName
    If Err.Number <> 0 Then
        hiddenSh.Name = hiddenShName & "_" & Format(Now, "hhmmss")
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' 4. 深度隐藏备份表，防止用户手动取消隐藏
    hiddenSh.Visible = xlSheetVeryHidden
    ws.Activate

    ' 5. 启动类模块监听
    If EventWatcher Is Nothing Then Set EventWatcher = New clsAppEvents
    Set EventWatcher.App = Application

    ' 6. 更新 UI 和全局变量
    bIsTracking = True
    RefreshRibbon ' 刷新 UI 状态

CleanUp:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
        
    MsgBox "修改追踪已启动！", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "启动失败 & Err.Description", vbCritical
    Resume CleanUp
End Sub

' Sub: StopTrackerLogic
' 功能：关闭追踪系统
' 逻辑：修改全局状态变量并刷新 UI。出于数据安全考虑，该过程不执行删除备份表的动作
Sub StopTrackerLogic()
    bIsTracking = False
    RefreshRibbon
End Sub

' Sub: RefreshRibbon - 强制刷新 Ribbon 按钮状态，使其重新调用 GetEnabledState
Sub RefreshRibbon()
    If Not myRibbon Is Nothing Then
        myRibbon.Invalidate
    End If
End Sub

' Sub: ProcessSheetActivate
' 功能：处理工作表激活事件
Sub ProcessSheetActivate(ByVal Sh As Object)
    StopTrackerLogic
    MsgBox "已切换工作表，修改追踪已自动停止。如需监控当前表，请重新点击启动", vbInformation
End Sub

' Sub: ProcessSheetChange
' 功能：处理工作表变更事件的核心逻辑
Sub ProcessSheetChange(ByVal Sh As Object, ByVal Target As Range)
    If IsHiddenLogSheet(Sh.Name) Then Exit Sub
    If Target.Cells.CountLarge > MaxTargetCells Then Exit Sub

    Dim hiddenSh As Worksheet
    Set hiddenSh = GetHiddenLogSheet(Sh)
    If hiddenSh Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim cell As Range
    Dim originalValue As Variant
    Dim modDateCol As Long
    Dim headerRow As Long

    FindOrCreateModDateHeader Sh, modDateCol, headerRow

    For Each cell In Target
        If cell.Column <> modDateCol Then
            originalValue = GetWorksheetCellValue(hiddenSh, cell.Row, cell.Column)
            If IsValueChanged(cell.Value, originalValue) Then
                MarkCellModified cell
                WriteModifyDate Sh, cell.Row, modDateCol, headerRow
            Else
                ClearCellModification cell
                If Not HasRowModifiedCells(Sh, cell.Row, modDateCol) Then
                    ClearDateCell Sh, cell.Row, modDateCol, headerRow
                End If
            End If
        End If
    Next cell

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Sub: ApplyDeleteMark
' 功能：执行删除标记的核心逻辑
' 逻辑
'   1. 提取选中区域涉及的所有行号（通过 Collection 去重）
'   2. 更新这些行的日期列
'   3. 给选中单元格添加红色、删除线和黄色背景
Sub ApplyDeleteMark(targetRange As Range)
    Dim ws As Worksheet
    Dim cell As Range
    Dim rowIds As Collection
    Dim rowNum As Variant
    Dim modDateCol As Long
    Dim headerRow As Long

    Set ws = targetRange.Worksheet
    ' 禁止在日志表内操作
    If ws.Name Like HiddenLogPrefix & "*" Then
        MsgBox "无法对系统日志表执行标记删除操作", vbExclamation
        Exit Sub
    End If

    Set rowIds = New Collection
    On Error Resume Next
    ' 收集唯一行号，避免对同一行多次重复写日期操作
    For Each cell In targetRange.Cells
        rowIds.Add CStr(cell.Row), CStr(cell.Row)
    Next cell
    On Error GoTo 0

    FindOrCreateModDateHeader ws, modDateCol, headerRow

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 写入修改日期
    For Each rowNum In rowIds
        WriteModifyDate ws, CLng(rowNum), modDateCol, headerRow
    Next rowNum

    ' 设置单元格样式：删除线+ 变红 + 黄底
    For Each cell In targetRange.Cells
        With cell.Font
            .Strikethrough = True
            .Color = vbRed
        End With
        cell.Interior.Color = vbYellow
    Next cell

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Sub: ApplyManualMark
' 功能：对子范围应用手动标记（红字黄底）并写入日期
Sub ApplyManualMark(targetRange As Range)
    Dim ws As Worksheet
    Dim cell As Range
    Dim rowIds As Collection
    Dim rowNum As Variant
    Dim modDateCol As Long
    Dim headerRow As Long

    Set ws = targetRange.Worksheet
    If IsHiddenLogSheet(ws.Name) Then
        MsgBox "无法对系统日志表执行手动标记", vbInformation
        Exit Sub
    End If

    Set rowIds = New Collection
    On Error Resume Next
    For Each cell In targetRange.Cells
        rowIds.Add CStr(cell.Row), CStr(cell.Row)
    Next cell
    On Error GoTo 0

    FindOrCreateModDateHeader ws, modDateCol, headerRow

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each rowNum In rowIds
        WriteModifyDate ws, CLng(rowNum), modDateCol, headerRow
    Next rowNum

    For Each cell In targetRange.Cells
        MarkCellModified cell
    Next cell

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

' 按钮1：清除所有格式
' 功能：对旧标记的红底涂黄格式进行清除，不清除日期
Sub ClearFormatsOnly()
    Dim ws As Worksheet
    Dim dataRow As Range
    Dim cell As Range
    Dim modDateCol As Long, headerRow As Long

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    If Not FindModDateHeader(ws, modDateCol, headerRow) Then
        MsgBox "未找到修改日期列，无法进行扫描！", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 保持原有的寻找逻辑
    For Each dataRow In ws.UsedRange.Rows
        If IsDate(ws.Cells(dataRow.Row, modDateCol).Value) Or dataRow.Row = headerRow Then
            For Each cell In dataRow.Cells
                If IsModifiedCell(cell) Then
                    ' 仅清除格式，不碰日期
                    ClearCellModification cell
                End If
            Next cell
        End If
    Next dataRow

    Application.ScreenUpdating = True
End Sub

' 按钮2：清除所有标记
' 功能：只要日期列有日期，就清除整行的红字黄底格式，并清除该日期
Sub ClearMarksAndDates()
    Dim ws As Worksheet
    Dim lastRow As Long, modDateCol As Long, headerRow As Long
    Dim i As Long
    Dim targetCell As Range

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ' 1. 寻找日期
    If Not FindModDateHeader(ws, modDateCol, headerRow) Then
        MsgBox "未找到修改日期列，无法进行扫描！", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 2. 获取最后一行
    lastRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1

    ' 3. 循环遍历所有行
    For i = headerRow + 1 To lastRow
        ' 判断日期列是否有内容 (如果是日期类型)
        If IsDate(ws.Cells(i, modDateCol).Value) Then
            
            ' 清除该行所有单元格的红字黄底格式
            ' 注意：仅清除符合条件的格式，不影响其他格式
            For Each targetCell In ws.Rows(i).Cells
               targetCell.Font.Strikethrough = False
                ' 如果是红字且是黄底，则恢复
                If IsModifiedCell(targetCell) Then
                    ClearCellModification targetCell
                End If
                
                If targetCell.Column > ws.UsedRange.Columns.Count Then Exit For
            Next targetCell

            ' 清除日期及其格式
            With ws.Cells(i, modDateCol)
                .ClearContents
                .Font.Strikethrough = False
                .Interior.ColorIndex = xlNone
            End With
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

' 按钮3：手动清除标记
' 功能：对选中单元格的红底涂黄格式进行清除，并清除其对应行的日期
Sub ClearManualSelection()
    Dim ws As Worksheet
    Dim cell As Range
    Dim modDateCol As Long, headerRow As Long
    Dim dictRows As Object

    Set ws = ActiveSheet
    If ws Is Nothing Or TypeName(Selection) <> "Range" Then Exit Sub

    ' 寻找日期
    FindModDateHeader ws, modDateCol, headerRow

    Application.ScreenUpdating = False
    
    Set dictRows = CreateObject("Scripting.Dictionary")

    ' 1. 遍历选中的单元格
    For Each cell In Selection
        ' 修正点：判断条件和清除逻辑都要指向 .Font.Strikethrough
        If IsModifiedCell(cell) Then
            ClearCellModification cell
            cell.Font.Strikethrough = False ' 修正：添加了 .Font
            
            If modDateCol > 0 And cell.Row <> headerRow Then
                If Not dictRows.exists(cell.Row) Then
                    dictRows.Add cell.Row, True
                End If
            End If
        End If
    Next cell
    
    ' 2. 批量清除对应行的日期列格式
    If modDateCol > 0 Then
        Dim varRow As Variant
        For Each varRow In dictRows.keys
            With ws.Cells(varRow, modDateCol)
                .ClearContents
                .Interior.ColorIndex = xlNone
                .Font.Strikethrough = False ' 修正：添加了 .Font
            End With
        Next varRow
    End If

    Application.ScreenUpdating = True
End Sub

' Sub: WriteModifyDate
' 功能：在指定的行和日期列位置写入当前日期
' 逻辑：包含格式化日期和单元格底色设置
Public Sub WriteModifyDate(ws As Worksheet, rowIndex As Long, modDateCol As Long, headerRow As Long)
    If rowIndex = headerRow Then Exit Sub ' 防止覆盖标题
    With ws.Cells(rowIndex, modDateCol)
        .Value = Date
        .NumberFormat = "yyyy/mm/dd"
        .Font.Color = vbRed
        .Interior.Color = vbYellow
    End With
End Sub
