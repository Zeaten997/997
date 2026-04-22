Sub 外边框加粗()
    Dim ws As Worksheet
    Dim i As Integer, rowStart As Long, rowEnd As Long
    Dim colStart As Long, colEnd As Long
    Dim printRange As Range, pageArea As Range
    Dim headerCount As Integer
    
    Set ws = ActiveSheet
    
    ' 强制进入分页预览模式
    ActiveWindow.View = xlPageBreakPreview
    
    ' 1. 定位打印区域（蓝框范围）
    On Error Resume Next
    Set printRange = ws.Range(ws.PageSetup.PrintArea)
    On Error GoTo 0
    
    ' 兜底：如果没设打印区域，则取当前蓝框可见的极限范围
    If printRange Is Nothing Then
        Dim maxRow As Long, maxCol As Long
        If ws.HPageBreaks.Count > 0 Then maxRow = ws.HPageBreaks(ws.HPageBreaks.Count).Location.Row - 1 Else maxRow = ws.UsedRange.Rows.Count
        If ws.VPageBreaks.Count > 0 Then maxCol = ws.VPageBreaks(ws.VPageBreaks.Count).Location.Column - 1 Else maxCol = ws.UsedRange.Columns.Count
        Set printRange = ws.Range(ws.Cells(1, 1), ws.Cells(maxRow, maxCol))
    End If
    
    colStart = printRange.Column
    colEnd = printRange.Columns(printRange.Columns.Count).Column
    
    ' 2. 识别标题行数
    If ws.PageSetup.PrintTitleRows <> "" Then
        headerCount = ws.Range(ws.PageSetup.PrintTitleRows).Rows.Count
    Else
        headerCount = 3
    End If
    
    ' 3. 全局清理并画细线
    With printRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = xlAutomatic
    End With
    
    ' 4. 核心：仅加粗每一页的“大外框”
    For i = 0 To ws.HPageBreaks.Count
        If i = 0 Then
            rowStart = printRange.Row
        Else
            rowStart = ws.HPageBreaks(i).Location.Row
        End If
        
        If i = ws.HPageBreaks.Count Then
            rowEnd = printRange.Rows(printRange.Rows.Count).Row
        Else
            rowEnd = ws.HPageBreaks(i + 1).Location.Row - 1
        End If
        
        If rowEnd >= rowStart Then
            ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd)).BorderAround _
                LineStyle:=xlContinuous, Weight:=xlMedium
        End If
    Next i
    
    ' 5. 标题加粗：仅针对全表最顶端加粗一次
    With ws.Range(ws.Cells(printRange.Row, colStart), ws.Cells(printRange.Row + headerCount - 1, colEnd))
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    ' 最终要求的简洁提示
    MsgBox "●边框自动调整完成！", vbInformation, "提示"
End Sub


