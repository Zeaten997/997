Attribute VB_Name = "modAutoCleanDeleted"
Sub AutoCleanDeleted_Click(control As IRibbonControl)
    Dim wsData As Worksheet
    Dim lastRow As Long, paramLastRow As Long
    Dim i As Long, j As Long, k As Long
    Dim keywords As Variant
    Dim deleteThisRow As Boolean
    Dim cellData As String
    Dim usedColCount As Long
    Dim hasText As Boolean
    Dim isFullyStruck As Boolean
    Dim lastCell As Range
    
    ' ================= 参数设置区 =================
    ' 你可以在这里直接修改、增加或删除关键字，用英文逗号和双引号隔开即可
    keywords = Array("删除", "取消", "作废", "清理")
    ' ==============================================
    
    ' 1. 设置当前活动工作表
    Set wsData = ActiveSheet
    
    ' 2. 获取数据表的使用范围
    Set lastCell = wsData.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        MsgBox "当前工作表是空的！", vbInformation, "提示"
        Exit Sub
    End If
    lastRow = lastCell.Row
    usedColCount = wsData.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    ' 关闭屏幕更新以提高运行速度
    Application.ScreenUpdating = False
    
    ' 3. 核心判断逻辑 (从下往上遍历，防止删除行后索引错乱)
    For i = lastRow To 1 Step -1
        deleteThisRow = False
        hasText = False
        isFullyStruck = True ' 假设整行都有删除线，一旦发现反例则设为 False
        
        ' 遍历该行的所有已使用列
        For j = 1 To usedColCount
            If Not IsEmpty(wsData.Cells(i, j).Value) Then
                hasText = True
                cellData = wsData.Cells(i, j).Value
                
                ' ---------- 条件 1：判断是否包含内置关键字 ----------
                For k = LBound(keywords) To UBound(keywords)
                    If InStr(1, cellData, keywords(k), vbTextCompare) > 0 Then
                        deleteThisRow = True
                        Exit For
                    End If
                Next k
                
                ' 如果已经因为关键字判定删除，无需继续检查该行其他单元格
                If deleteThisRow Then Exit For
                
                ' ---------- 条件 2：判断删除线 ----------
                ' 如果字体没有删除线(False)，或者只有部分文字有删除线(Null)
                If IsNull(wsData.Cells(i, j).Font.Strikethrough) Or wsData.Cells(i, j).Font.Strikethrough = False Then
                    isFullyStruck = False
                End If
            End If
        Next j
        
        ' 如果该行有内容，没有因为关键字被删，但该行所有非空文字都具有完整的删除线标记
        If hasText And (Not deleteThisRow) And isFullyStruck Then
            deleteThisRow = True
        End If
        
        ' 4. 执行删除并下方单元格上移
        If deleteThisRow Then
            wsData.rows(i).Delete Shift:=xlUp
        End If
    Next i
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    MsgBox "数据清理完成！", vbInformation, "完成"
End Sub
