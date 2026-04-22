Attribute VB_Name = "modTrackerHelpers"
Option Explicit

' 定义隐藏日志表的前缀常量
Private Const HiddenLogPrefix As String = "HiddenLog_"

' ============================================================
' 函数: GetModDateHeaderKeywords
' 功能: 返回日期列可能的标题关键字数组，用于在表中定位日期列
' ============================================================
Public Function GetModDateHeaderKeywords() As Variant
    GetModDateHeaderKeywords = Array("修改日期", "日期", "修改")
End Function

' ============================================================
' 函数: IsHiddenLogSheet
' 功能: 根据表名判断是否为系统生成的隐藏日志备份表
' ============================================================
Public Function IsHiddenLogSheet(sheetName As String) As Boolean
    IsHiddenLogSheet = (sheetName Like HiddenLogPrefix & "*")
End Function

' ============================================================
' 函数: GetHiddenLogSheet
' 功能: 获取指定工作表对应的隐藏备份工作表对象
' ============================================================
Public Function GetHiddenLogSheet(ws As Worksheet) As Worksheet
    On Error Resume Next
    ' 通过表名拼接寻找对应的备份表
    Set GetHiddenLogSheet = ws.Parent.Worksheets(HiddenLogPrefix & ws.Name)
    On Error GoTo 0
End Function

' ============================================================
' 函数: FindModDateHeader
' 功能: 在工作表的前10行内搜索预设的关键字，定位“日期列”的位置
' ============================================================
Public Function FindModDateHeader(ws As Worksheet, ByRef modDateCol As Long, ByRef headerRow As Long) As Boolean
    Dim headers As Variant, h As Variant
    Dim foundCell As Range

    headers = GetModDateHeaderKeywords()
    ' 遍历关键字数组进行查找
    For Each h In headers
        On Error Resume Next
        ' 在1到10行内进行全单元格匹配查找
        Set foundCell = ws.Range("1:10").Find(What:=h, LookAt:=xlWhole)
        On Error GoTo 0

        ' 如果找到，则记录列号和行号并返回True
        If Not foundCell Is Nothing Then
            modDateCol = foundCell.Column
            headerRow = foundCell.Row
            FindModDateHeader = True
            Exit Function
        End If
    Next h

    FindModDateHeader = False
End Function

' ============================================================
' 函数: FindOrCreateModDateHeader
' 功能: 查找日期列，如果不存在则在表格最后一列自动创建
' ============================================================
Public Function FindOrCreateModDateHeader(ws As Worksheet, ByRef modDateCol As Long, ByRef headerRow As Long) As Boolean
    ' 首先尝试查找现有列
    If FindModDateHeader(ws, modDateCol, headerRow) Then
        FindOrCreateModDateHeader = True
        Exit Function
    End If

    ' 如果没找到，则在最后一列新建“修改日期”标题
    modDateCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column
    headerRow = 1
    With ws.Cells(headerRow, modDateCol)
        .Value = "修改日期"
        .Font.Bold = True ' 标题加粗
    End With

    FindOrCreateModDateHeader = True
End Function

' ============================================================
' 函数: IsValueChanged
' 功能: 比较单元格修改前后的值是否真的不同（强制转为字符串比较）
' ============================================================
Public Function IsValueChanged(newValue As Variant, oldValue As Variant) As Boolean
    IsValueChanged = (CStr(newValue) <> CStr(oldValue))
End Function

' ============================================================
' 函数: IsModifiedCell
' 功能: 判断一个单元格是否已经被标记为“修改状态”（根据红字黄底判断）
' ============================================================
Public Function IsModifiedCell(cell As Range) As Boolean
    IsModifiedCell = (cell.Font.Color = vbRed And cell.Interior.Color = vbYellow)
End Function

' ============================================================
' 子过程: MarkCellModified
' 功能: 将单元格设置为修改标记样式（红色字体，黄色背景）
' ============================================================
Public Sub MarkCellModified(cell As Range)
    cell.Font.Color = vbRed
    cell.Interior.Color = vbYellow
End Sub

' ============================================================
' 子过程: ClearCellModification
' 功能: 清除单元格的修改标记样式，恢复自动颜色和无填充背景
' ============================================================
Public Sub ClearCellModification(cell As Range)
    cell.Font.ColorIndex = xlAutomatic
    cell.Interior.ColorIndex = xlNone
End Sub

' ============================================================
' 函数: HasRowModifiedCells
' 功能: 检查某一行内（除了日期列外）是否还存在任何带有修改标记的单元格
' ============================================================
Public Function HasRowModifiedCells(ws As Worksheet, rowIndex As Long, dateCol As Long) As Boolean
    Dim colIdx As Long

    ' 遍历当前行所有有数据的列
    For colIdx = 1 To ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        If colIdx <> dateCol Then ' 排除日期列本身
            ' 如果发现该行还有其他单元格是红色黄底，则返回True
            If IsModifiedCell(ws.Cells(rowIndex, colIdx)) Then
                HasRowModifiedCells = True
                Exit Function
            End If
        End If
    Next colIdx

    HasRowModifiedCells = False
End Function

' ============================================================
' 子过程: ClearDateCell
' 功能: 清除日期列中特定行的内容及其格式（避开标题行）
' ============================================================
Public Sub ClearDateCell(ws As Worksheet, rowIndex As Long, dateCol As Long, headerRow As Long)
    If rowIndex = headerRow Then Exit Sub ' 防止清除标题

    With ws.Cells(rowIndex, dateCol)
        ' 仅在单元格确实包含日期时执行清除
        If IsDate(.Value) Then
            .ClearContents
            .Font.ColorIndex = xlAutomatic
            .Interior.ColorIndex = xlNone
        End If
    End With
End Sub

' ============================================================
' 函数: GetWorksheetCellValue
' 功能: 安全地从单元格读取数值，带有错误处理机制
' ============================================================
Public Function GetWorksheetCellValue(ws As Worksheet, rowIndex As Long, colIndex As Long) As Variant
    On Error Resume Next
    GetWorksheetCellValue = ws.Cells(rowIndex, colIndex).Value
    On Error GoTo 0
End Function

