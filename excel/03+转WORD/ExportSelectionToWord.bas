Attribute VB_Name = "ExportSelectionToWord"
Option Explicit

' ==========================================================
' 第一部分：全局参数设置区 (随时可修改这里的配置)
' ==========================================================
' 1. 专业与命名关键词
Private Const CFG_MAJOR_1 As String = "自动化"
Private Const CFG_MAJOR_2 As String = "电信"
Private Const CFG_FILE_PREFIX As String = "2+"             ' 生成文件名前缀
Private Const CFG_FILE_MID As String = "+设备清单"         ' 生成文件名固定中间段
Private Const CFG_FILE_TARGET As String = "自动化电信"     ' 需要被智能替换的旧词汇


' 2. Word 页面排版参数 (单位：厘米 cm) - 底层会自动换算为磅
Private Const CFG_MARGIN_TOP_CM As Single = 2.54           ' 上边距 (原72磅)
Private Const CFG_MARGIN_BOTTOM_CM As Single = 2.54        ' 下边距 (原72磅)
Private Const CFG_MARGIN_LEFT_CM As Single = 1.91          ' 左边距 (原54.15磅)
Private Const CFG_MARGIN_RIGHT_CM As Single = 1.91         ' 右边距 (原54.15磅)

' 3. 字体与字号参数
Private Const CFG_FONT_CN As String = "仿宋"
Private Const CFG_FONT_EN As String = "Times New Roman"
Private Const CFG_SIZE_TITLE As Single = 12                ' 一级标题字号
Private Const CFG_SIZE_TABLE As Single = 10.5              ' 表格正文字号
Private Const CFG_ROW_HEIGHT_CM As Single = 0.8            ' 表格行高 (厘米)
Private excelHeaderRowsCount As Long                        ' 标题行数初始值
Private Const CFG_TITLE_ROW_HEIGHT_CM As Single = 1.6       ' 标题行总高度(厘米)

' 4. 列宽排版参数 (单位：厘米 cm)
' 配置格式严格按照："序号,名称,性能,单位,数量1,数量2(总计),单重,合重,备注" (0表示该列不存在)
Private Const CFG_WIDTH_PORTRAIT_A As String = "1.2, 3.0, 7.0, 1.2, 1.2, 1.2, 0.0, 0.0, 2.4" ' 纵向A：无重量，有总计
Private Const CFG_WIDTH_PORTRAIT_B As String = "1.2, 4.0, 7.2, 1.2, 1.2, 0.0, 0.0, 0.0, 2.4" ' 纵向B：无重量，无总计
Private Const CFG_WIDTH_LANDSCAPE_A As String = "1.2, 4.8,10.7, 1.2, 1.2, 1.2, 1.2, 1.2, 3.2" ' 横向A：有重量，有总计
Private Const CFG_WIDTH_LANDSCAPE_B As String = "1.2, 5.0,11.1, 1.2, 1.2, 0.0, 1.2, 1.2, 3.8" ' 横向B：有重量，无总计
Private Const CFG_WIDTH_LANDSCAPE_C As String = "1.2, 5.0,12.3, 1.2, 1.2, 1.2, 0.0, 0.0, 3.8" ' 横向C：无重量，有总计
Private Const CFG_WIDTH_LANDSCAPE_D As String = "1.2, 5.0,12.3, 1.2, 1.2, 0.0, 0.0, 0.0, 5.0" ' 横向D：无重量，无总计


' 5. 全局状态变量
Public g_RibbonUI As IRibbonUI
Public g_PageOrientation As Long ' 0=纵向, 1=横向

Public Const EXPORT_MODE_MANUAL As Long = 0
Public Const EXPORT_MODE_AUTO As Long = 1
Public Const EXPORT_MODE_TELECOM As Long = 2
Public Const EXPORT_MODE_ALL As Long = 3
' 6. 区块过滤配置 (遇到Start关键字后，中间内容不输出，直到遇到End关键字恢复输出)
Private Const CFG_BLOCK_SKIP_START As String = "计算机控制系统软硬件"
Private Const CFG_BLOCK_SKIP_END As String = "计算机控制系统应用软件"
Private Const CFG_EXCLUDE_NAMES As String = "设备及安装,材料及安装" ' 需要过滤不输出的项目名称关键字 (多个词用英文逗号隔开)


' 在这里填写你允许保留的规格关键词（支持英文、数字、汉字及特殊符号）
Private Property Get specKeywords() As Variant
    specKeywords = Array("防爆", "成套")
End Property




' 7. 获取需要过滤的尺寸正则规则 (由于VBA不支持数组常量，用函数返回)
Private Function GetSizePatterns() As Variant
    GetSizePatterns = Array("DN\d+(?:\.\d+)?[，。、；：,.;:]?", _
                            "D\d+(?:\.\d+)?[Xx×]\d+(?:\.\d+)?[，。、；：,.;:]?", _
                            "φ\d+(?:\.\d+)?[Xx×]\d+(?:\.\d+)?[，。、；：,.;:]?", _
                            "Φ\d+(?:\.\d+)?[Xx×]\d+(?:\.\d+)?[，。、；：,.;:]?", _
                            "统计IO用")
End Function



' ==========================================================
' 第三部分：主控流程 (逻辑编排)
' ==========================================================
Sub StartExportProcess(Optional ByVal exportMode As Variant)
    Dim mode As Long
    mode = ResolveExportMode(exportMode)
    Dim exportRanges As Collection
    Set exportRanges = New Collection
    
    If mode = EXPORT_MODE_MANUAL Then
        Dim selRange As Range
        On Error Resume Next
        Set selRange = Application.Selection
        On Error GoTo 0
        
        If Not IsValidExportRange(selRange) Then
            MsgBox "请先在Excel中框选有效的工程量清单区域！" & vbCrLf & _
                   "有效区域应包含从【序号】列到【备注】列的完整范围。", vbExclamation
            Exit Sub
        End If
        exportRanges.Add selRange
    Else
        Dim autoRange As Range
        If mode = EXPORT_MODE_AUTO Or mode = EXPORT_MODE_ALL Then
            Set autoRange = FindMajorTableRange(CFG_MAJOR_1)
            If autoRange Is Nothing Then
                MsgBox "未自动找到【" & CFG_MAJOR_1 & "专业工程量清单】范围，请检查工作簿前5行标题、序号列、备注列和材料及安装结束行。", vbExclamation
                If mode <> EXPORT_MODE_ALL Then Exit Sub
            Else
                exportRanges.Add autoRange
            End If
        End If
        
        If mode = EXPORT_MODE_TELECOM Or mode = EXPORT_MODE_ALL Then
            Set autoRange = FindMajorTableRange(CFG_MAJOR_2)
            If autoRange Is Nothing Then
                MsgBox "未自动找到【" & CFG_MAJOR_2 & "专业工程量清单】范围，请检查工作簿前5行标题、序号列、备注列和材料及安装结束行。", vbExclamation
                If mode <> EXPORT_MODE_ALL Then Exit Sub
            Else
                exportRanges.Add autoRange
            End If
        End If
        
        If exportRanges.Count = 0 Then Exit Sub
    End If
    
    ' 1. 初始化 Word 并设置页面
    Dim wdApp As Object, wdDoc As Object
    Call InitWordDocument(wdApp, wdDoc)
    
    Application.ScreenUpdating = False
    wdApp.ScreenUpdating = False
    
    ' 2. 核心状态追踪
    Dim firstTitle As String, currentTitle As String
    Dim hasMajor1 As Boolean, hasMajor2 As Boolean
    
    ' 3. 按模式处理表格
    Dim i As Long
    For i = 1 To exportRanges.Count
        If i > 1 Then wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1).InsertBreak Type:=7
        
        currentTitle = ProcessSingleTable(exportRanges(i), wdApp, wdDoc)
        If firstTitle = "" Then firstTitle = currentTitle
        
        If InStr(currentTitle, CFG_MAJOR_1) > 0 Then hasMajor1 = True
        If InStr(currentTitle, CFG_MAJOR_2) > 0 Then hasMajor2 = True
    Next i
    
    If mode = EXPORT_MODE_ALL And hasMajor1 And hasMajor2 Then
        MsgBox "已成功转换【" & CFG_MAJOR_1 & "】和【" & CFG_MAJOR_2 & "】两个专业的清单，全部转换完成！", vbInformation, "转换完成"
    End If
    
    ' 4. 执行保存
    Call SaveExportedDocument(wdApp, wdDoc, firstTitle, hasMajor1, hasMajor2)
    
    wdApp.ScreenUpdating = True
    Application.ScreenUpdating = True
    wdApp.Visible = True
    wdApp.Activate
    wdDoc.Activate
End Sub

Private Function ResolveExportMode(ByVal exportMode As Variant) As Long
    If IsMissing(exportMode) Then
        ResolveExportMode = EXPORT_MODE_MANUAL
    ElseIf VarType(exportMode) = vbBoolean Then
        If CBool(exportMode) Then
            ResolveExportMode = EXPORT_MODE_ALL
        Else
            ResolveExportMode = EXPORT_MODE_MANUAL
        End If
    Else
        ResolveExportMode = CLng(exportMode)
    End If
End Function

Private Function FindMajorTableRange(ByVal majorName As String) As Range
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Dim titleRow As Long
        titleRow = FindMajorTitleRow(ws, majorName)
        If titleRow > 0 Then
            Dim endRow As Long, startCol As Long, endCol As Long
            endRow = FindRowContainingText(ws, titleRow, ws.UsedRange.Row + ws.UsedRange.rows.Count - 1, "材料及安装")
            If endRow = 0 Then endRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
            
            startCol = FindColumnContainingText(ws, titleRow, endRow, "序号")
            endCol = FindColumnContainingText(ws, titleRow, endRow, "备注")
            
            If startCol > 0 And endCol > 0 Then
                If endCol < startCol Then
                    Dim tmpCol As Long
                    tmpCol = startCol
                    startCol = endCol
                    endCol = tmpCol
                End If
                Set FindMajorTableRange = ws.Range(ws.Cells(titleRow, startCol), ws.Cells(endRow, endCol))
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function FindMajorTitleRow(ByVal ws As Worksheet, ByVal majorName As String) As Long
    Dim maxRow As Long, lastCol As Long, r As Long, c As Long
    maxRow = WorksheetFunction.Min(5, ws.UsedRange.Row + ws.UsedRange.rows.Count - 1)
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    
    For r = 1 To maxRow
        For c = 1 To lastCol
            Dim txt As String
            txt = GetCellText(ws.Cells(r, c))
            If InStr(txt, majorName & "专业工程量清单") > 0 Then
                FindMajorTitleRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindRowContainingText(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal keyword As String) As Long
    Dim lastCol As Long, r As Long, c As Long
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    
    For r = firstRow To lastRow
        For c = 1 To lastCol
            If InStr(GetCellText(ws.Cells(r, c)), keyword) > 0 Then
                FindRowContainingText = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindColumnContainingText(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal keyword As String) As Long
    Dim lastCol As Long, r As Long, c As Long
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    
    For r = firstRow To lastRow
        For c = 1 To lastCol
            If InStr(GetCellText(ws.Cells(r, c)), keyword) > 0 Then
                FindColumnContainingText = c
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function GetCellText(ByVal cell As Range) As String
    On Error Resume Next
    If cell.MergeCells Then
        GetCellText = CStr(cell.mergeArea.Cells(1, 1).Value)
    Else
        GetCellText = CStr(cell.Value)
    End If
    On Error GoTo 0
    GetCellText = Replace(Replace(Trim(GetCellText), " ", ""), Chr(160), "")
End Function

Private Function IsValidExportRange(ByVal selRange As Range) As Boolean
    On Error GoTo CleanFail
    If selRange Is Nothing Then Exit Function
    If selRange.Areas.Count > 1 Then Exit Function
    
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    firstRow = selRange.Row
    lastRow = selRange.Row + selRange.rows.Count - 1
    firstCol = selRange.Column
    lastCol = selRange.Column + selRange.Columns.Count - 1
    
    Dim ws As Worksheet
    Set ws = selRange.Worksheet
    
    Dim colNo As Long, colMemo As Long
    colNo = FindColumnContainingText(ws, firstRow, lastRow, "序号")
    colMemo = FindColumnContainingText(ws, firstRow, lastRow, "备注")
    
    IsValidExportRange = (colNo >= firstCol And colNo <= lastCol And _
                          colMemo >= firstCol And colMemo <= lastCol And _
                          colNo < colMemo)
    Exit Function
    
CleanFail:
    IsValidExportRange = False
End Function

' ==========================================================
' 第四部分：核心业务 (表格数据提取与 Word 排版)
' ==========================================================
Function ProcessSingleTable(ByVal selRange As Range, ByVal wdApp As Object, ByVal wdDoc As Object) As String
    Dim titleText As String
    Dim firstCellText As String
    
    ' 1. 获取并插入一级标题
    firstCellText = Replace(selRange.Cells(1, 1).mergeArea.Cells(1, 1).Value, " ", "")
    If InStr(firstCellText, "工程量清单") > 0 Then
        titleText = Replace(firstCellText, "工程量清单", "")
        Set selRange = selRange.Offset(1, 0).Resize(selRange.rows.Count - 1, selRange.Columns.Count)
    Else
        titleText = InputBox("选区首行未检测到包含“工程量清单”的标题行。" & vbCrLf & _
                             "请输入要作为 Word 一级标题的专业名称：" & vbCrLf & _
                             "（如果不想要标题，请直接点确定留空）", "手动输入专业标题")
    End If
    ProcessSingleTable = titleText
    
    Dim insertRange As Object
    Set insertRange = wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1)
    
    If titleText <> "" Then
        insertRange.Text = titleText & vbCrLf
        With insertRange.Paragraphs(1).Range
            .Style = wdDoc.Styles(-2)
            .Font.NameFarEast = CFG_FONT_CN
            .Font.NameAscii = CFG_FONT_EN
            .Font.Size = CFG_SIZE_TITLE
            .Font.Bold = True
            .Font.Color = 0 ' 黑色
            .ParagraphFormat.Alignment = 0
        End With
        Set insertRange = wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1)
    End If
    
    ' 2. 初始化数据与行列状态
    Dim srcRows As Long, srcCols As Long, r As Long, c As Long
    srcRows = selRange.rows.Count
    srcCols = selRange.Columns.Count
    
    Dim serialRowCnt As Long: serialRowCnt = 0
    Dim qtyRowCnt As Long: qtyRowCnt = 0
    excelHeaderRowsCount = 1
    For c = 1 To srcCols
        If Not selRange.Columns(c).Hidden Then
            If selRange.Cells(1, c).MergeCells Then
                Dim areaRows As Long: areaRows = selRange.Cells(1, c).mergeArea.rows.Count
                
                ' 记录序号列行数
                If InStr(selRange.Cells(1, c).Value, "序号") > 0 Then serialRowCnt = areaRows
                ' 记录工程量列行数
                If InStr(selRange.Cells(1, c).Value, "工程量") > 0 Then qtyRowCnt = areaRows
                
                ' 原有的最大值逻辑保持不变，确保不会漏掉其他列的合并情况
                If areaRows > excelHeaderRowsCount Then excelHeaderRowsCount = areaRows
            End If
        End If
    Next c
    
    ' --- 新增判断：如果两者行数相同，则强制修正标题行数 ---
    If serialRowCnt > 0 And qtyRowCnt > 0 And serialRowCnt = qtyRowCnt Then
        excelHeaderRowsCount = 1
    End If
    
    ' --- 新增代码：输出计算结果 ---
    Debug.Print "当前识别的标题行数为: " & excelHeaderRowsCount
    'Debug.Print "当前识别的wHeaderRows为: " & wHeaderRows
    ' 如果想在屏幕上直接看到，可以使用下面这行 (调试完记得注释掉)：
    ' MsgBox "程序识别到的标题行数为: " & excelHeaderRowsCount
    ' ----------------------------
    
    Dim dataArr() As Variant: ReDim dataArr(1 To srcRows, 1 To srcCols)
    Dim colIdxName As Long, colIdxSpec As Long, colIdxMemo As Long
    Dim isQtyCol() As Boolean: ReDim isQtyCol(1 To srcCols)
    Dim tmpTab As String: tmpTab = Chr(1)
    Dim tmpLf As String: tmpLf = Chr(2)
    Dim patterns As Variant: patterns = GetSizePatterns()
    
    ' 3. 数据清洗装载到数组
    For r = 1 To srcRows
        For c = 1 To srcCols
            If selRange.Cells(r, c).EntireRow.Hidden Or selRange.Cells(r, c).EntireColumn.Hidden Then
                dataArr(r, c) = ""
            Else
                Dim txt As String
                ' 【核心优化】改用 .Text 属性，彻底避免单元格含公式错误值（如 #N/A）时报错
                txt = selRange.Cells(r, c).Text
                txt = Replace(txt, vbTab, tmpTab)
                txt = Replace(txt, vbCrLf, tmpLf)
                txt = Replace(txt, vbLf, tmpLf)
                ' 清除全角空格，保证输出排版整洁及正则匹配的准确性
                txt = Replace(txt, "　", "")
                
                If r <= excelHeaderRowsCount Then
                    Select Case txt
                        Case "项目名称": txt = "设备名称"
                        Case "项目技术特征": txt = "型号、规格及技术性能"
                        Case "计量单位": txt = "单位"
                        Case "工程量": txt = "数量"
                    End Select
                    If r = 1 Then
                        Select Case txt
                            Case "设备名称": colIdxName = c
                            Case "型号、规格及技术性能": colIdxSpec = c
                            Case "备注": colIdxMemo = c
                        End Select
                    End If
                Else
                    txt = RemoveSizeInfo(txt, patterns)
                End If
                dataArr(r, c) = txt
            End If
        Next c
    Next r
    
     ' 4. 识别数值列与合并同类项
    For c = 1 To srcCols
        If Not selRange.Columns(c).Hidden Then
            Dim headVal As String: headVal = selRange.Cells(1, c).mergeArea.Cells(1, 1).Value
            If headVal = "工程量" Or headVal = "数量" Or dataArr(1, c) = "数量" Then isQtyCol(c) = True
        End If
    Next c
    
    If colIdxName > 0 And colIdxSpec > 0 Then
        ' 前置处理：符合成套条件的强制修改规格
        For r = (excelHeaderRowsCount + 1) To srcRows
            If Left(dataArr(r, colIdxName), 1) = "*" Or InStr(dataArr(r, colIdxName), "成套") > 0 Then
                dataArr(r, colIdxSpec) = "设备成套提供"
            End If
        Next r
        
' 自下而上合并相邻行
        For r = srcRows To (excelHeaderRowsCount + 2) Step -1
            ' 【条件一】相邻的两行都必须是显示状态
            If Not (selRange.rows(r).Hidden Or selRange.rows(r - 1).Hidden) Then
                
                Dim nameCurr As String: nameCurr = CleanString(dataArr(r, colIdxName))
                Dim namePrev As String: namePrev = CleanString(dataArr(r - 1, colIdxName))
                
                ' 使用新规则清洗规格：只保留白名单内的关键词
                Dim specCurr As String: specCurr = CleanSpecWithKeywords(dataArr(r, colIdxSpec), specKeywords)
                Dim specPrev As String: specPrev = CleanSpecWithKeywords(dataArr(r - 1, colIdxSpec), specKeywords)
                
                ' 【条件二】清洗后的“名称”必须完全相同
                ' 【条件三·新】基于关键词清洗后的“规格”必须完全相同
                ' 【条件四】名称不能为空
                If nameCurr = namePrev And specCurr = specPrev And nameCurr <> "" Then
                    
                    ' 1. 数量累加
                    Dim subCol As Long
                    For subCol = 1 To srcCols
                        If isQtyCol(subCol) Then dataArr(r - 1, subCol) = Val(dataArr(r, subCol)) + Val(dataArr(r - 1, subCol))
                    Next subCol
                    
                    ' ==========================================================
                    ' 2. 【新增修复】处理合并后的原始规格显示
                    ' 如果两行的原始规格字符串不完全一样，说明是不同细分规格被合并了。
                    ' 为了避免保留单一规格造成误导，将其清空（或者你可以改为 "-"）
                    If dataArr(r, colIdxSpec) <> dataArr(r - 1, colIdxSpec) Then
                        dataArr(r - 1, colIdxSpec) = ""
                        ' 如果你想把它变成拼接形式，可以把上面那句换成：
                        ' dataArr(r - 1, colIdxSpec) = dataArr(r - 1, colIdxSpec) & " / " & dataArr(r, colIdxSpec)
                    End If
                    ' ==========================================================
                    
                    ' 3. 打上合并标记并清空备注
                    dataArr(r, colIdxName) = "[MERGED_ROW]"
                    If colIdxMemo > 0 Then dataArr(r, colIdxMemo) = "": dataArr(r - 1, colIdxMemo) = ""
                    
                End If
                
            End If ' <--- 就是这里！补上【条件一】的 End If
            
            Next r
        End If
    
    ' 5. 输出数据到 Word 表格
    Dim excelToWordRow() As Long: ReDim excelToWordRow(1 To srcRows)
    Dim excelToWordCol() As Long: ReDim excelToWordCol(1 To srcCols)
    Dim tRows As Long, wHeaderRows As Long, tCols As Long
    Dim serialCol As Long: serialCol = FindSerialColumn(dataArr, srcCols, excelHeaderRowsCount)
    Dim skipBlockActive As Boolean
    skipBlockActive = False
    
    For r = 1 To srcRows
        ' 1. 检查是否遇到“结束关键字”，如果遇到，则提前关闭跳过状态 (保留本行)
        If CFG_BLOCK_SKIP_END <> "" Then
            If InStr(CStr(dataArr(r, IIf(colIdxName > 0, colIdxName, 1))), CFG_BLOCK_SKIP_END) > 0 Or _
               InStr(CStr(dataArr(r, 1)), CFG_BLOCK_SKIP_END) > 0 Then
                skipBlockActive = False
            End If
        End If
        
        ' 2. 判断当前行是否需要过滤 (标题行不参与过滤)
        Dim skipExclusion As Boolean: skipExclusion = False
        If r > excelHeaderRowsCount Then
            If colIdxName > 0 Then
                skipExclusion = IsExcludedName(CStr(dataArr(r, colIdxName)))
            End If
            ' 如果当前处于“区块跳过”状态，强制标记为需要过滤
            If skipBlockActive Then skipExclusion = True
        End If
        
        ' 3. 检查是否遇到“开始关键字”，如果遇到，则从下一行开始开启跳过状态 (保留本行)
        If CFG_BLOCK_SKIP_START <> "" Then
            If InStr(CStr(dataArr(r, IIf(colIdxName > 0, colIdxName, 1))), CFG_BLOCK_SKIP_START) > 0 Or _
               InStr(CStr(dataArr(r, 1)), CFG_BLOCK_SKIP_START) > 0 Then
                skipBlockActive = True
            End If
        End If
        
        ' 4. 最终决定是否输出该行
        If Not selRange.rows(r).Hidden And dataArr(r, IIf(colIdxName > 0, colIdxName, 1)) <> "[MERGED_ROW]" _
           And Not skipExclusion _
           And ShouldExportRow(r, excelHeaderRowsCount, dataArr, srcCols, isQtyCol, serialCol) Then
            tRows = tRows + 1: excelToWordRow(r) = tRows
            If r <= excelHeaderRowsCount Then wHeaderRows = wHeaderRows + 1
        End If
    Next r
    For c = 1 To srcCols
        If Not selRange.Columns(c).Hidden Then tCols = tCols + 1: excelToWordCol(c) = tCols
    Next c
    
    If tRows = 0 Or tCols = 0 Then
        MsgBox "选区中没有可导出的有效数据。", vbExclamation
        Exit Function
    End If
    
    Dim rowArray() As String: ReDim rowArray(1 To tRows)
    Dim colArray() As String: ReDim colArray(1 To tCols)
    Dim wR As Long: wR = 1
    
    For r = 1 To srcRows
        If excelToWordRow(r) > 0 Then
            Dim wC As Long: wC = 1
            For c = 1 To srcCols
                If excelToWordCol(c) > 0 Then
                    colArray(wC) = dataArr(r, c)
                    wC = wC + 1
                End If
            Next c
            rowArray(wR) = Join(colArray, vbTab)
            wR = wR + 1
        End If
    Next r
    
    insertRange.Text = Join(rowArray, vbCrLf) & vbCrLf
    Dim wdTable As Object
    Set wdTable = insertRange.ConvertToTable(Separator:=1, numRows:=tRows, NumColumns:=tCols)
    
' 6. 统一格式化与表格对齐
    Call ApplyWordTableStyles(wdTable, wdApp, tCols, wHeaderRows, dataArr, srcCols, excelToWordCol, isQtyCol, tmpTab, tmpLf)
    Call ApplyCustomColumnWidths(wdTable, g_PageOrientation, dataArr, excelToWordCol, srcCols, excelHeaderRowsCount)
    
 ' =========================================================================
    ' 6.5 大类序号行整体加粗 (已增加特定行排除逻辑)
    ' =========================================================================
    If serialCol > 0 Then
        Dim currRow As Long
        For currRow = (excelHeaderRowsCount + 1) To srcRows
            If excelToWordRow(currRow) > 0 Then
                ' 1. 判断原数组中对应行的序号列是否符合汉字序号正则
                If IsChineseHeading(CStr(dataArr(currRow, serialCol))) Then
                    
                    ' 2. 检查当前行是否包含需要排除加粗的关键字
                    Dim rowTextName As String, rowTextCol1 As String
                    Dim skipBold As Boolean: skipBold = False
                    
                    ' 提取当前行的名称列和第一列文本进行双重比对
                    rowTextName = CStr(dataArr(currRow, IIf(colIdxName > 0, colIdxName, 1)))
                    rowTextCol1 = CStr(dataArr(currRow, 1))
                    
                    If CFG_BLOCK_SKIP_START <> "" And (InStr(rowTextName, CFG_BLOCK_SKIP_START) > 0 Or InStr(rowTextCol1, CFG_BLOCK_SKIP_START) > 0) Then
                        skipBold = True
                    End If
                    
                    If CFG_BLOCK_SKIP_END <> "" And (InStr(rowTextName, CFG_BLOCK_SKIP_END) > 0 Or InStr(rowTextCol1, CFG_BLOCK_SKIP_END) > 0) Then
                        skipBold = True
                    End If
                    
                    ' 3. 如果不是被排除的特殊行，则执行加粗
                    If Not skipBold Then
                        wdTable.rows(excelToWordRow(currRow)).Range.Font.Bold = True
                    End If
                    
                End If
            End If
        Next currRow
    End If
    
    ' 7. 重建合并单元格
    For r = srcRows To 1 Step -1
        For c = srcCols To 1 Step -1
            Dim cellExcel As Range: Set cellExcel = selRange.Cells(r, c)
            If cellExcel.MergeCells And cellExcel.Address = cellExcel.mergeArea.Cells(1, 1).Address Then
                Dim startWRow As Long: startWRow = excelToWordRow(r)
                Dim startWCol As Long: startWCol = excelToWordCol(c)
                If startWRow > 0 And startWCol > 0 Then
                    Dim endWRow As Long, endWCol As Long, i As Long
                    For i = r + cellExcel.mergeArea.rows.Count - 1 To r Step -1
                        If excelToWordRow(i) > 0 Then endWRow = excelToWordRow(i): Exit For
                    Next i
                    For i = c + cellExcel.mergeArea.Columns.Count - 1 To c Step -1
                        If excelToWordCol(i) > 0 Then endWCol = excelToWordCol(i): Exit For
                    Next i
                    If endWRow > 0 And endWCol > 0 Then
                        If (endWRow > startWRow) Or (endWCol > startWCol) Then
                            On Error Resume Next
                            wdTable.cell(startWRow, startWCol).Merge wdTable.cell(endWRow, endWCol)
                            On Error GoTo 0
                        End If
                    End If
                End If
            End If
        Next c
    Next r
    
' 8. 应用行高 (原有代码)
    Call ApplyMinimumRowHeight(wdTable)
    
    
    wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1).Select
End Function

Private Function ShouldExportRow(ByVal rowIndex As Long, ByVal headerRows As Long, ByRef dataArr() As Variant, ByVal srcCols As Long, ByRef isQtyCol() As Boolean, ByVal serialCol As Long) As Boolean
    If rowIndex <= headerRows Then
        ShouldExportRow = True
        Exit Function
    End If
    
    Dim hasQtyCol As Boolean, hasQtyValue As Boolean, c As Long
    For c = 1 To srcCols
        If isQtyCol(c) Then
            hasQtyCol = True
            If CleanString(CStr(dataArr(rowIndex, c))) <> "" Then
                hasQtyValue = True
                Exit For
            End If
        End If
    Next c
    
    If hasQtyCol Then
        ShouldExportRow = hasQtyValue Or (serialCol > 0 And CleanString(CStr(dataArr(rowIndex, serialCol))) <> "")
    Else
        ShouldExportRow = True
    End If
End Function

Private Function FindSerialColumn(ByRef dataArr() As Variant, ByVal srcCols As Long, ByVal headerRows As Long) As Long
    Dim r As Long, c As Long
    For c = 1 To srcCols
        For r = 1 To headerRows
            If InStr(CStr(dataArr(r, c)), "序号") > 0 Then
                FindSerialColumn = c
                Exit Function
            End If
        Next r
    Next c
End Function

' ==========================================================
' 第五部分：独立的工具与服务层接口 (解耦具体操作)
' ==========================================================
Private Sub InitWordDocument(ByRef wdApp As Object, ByRef wdDoc As Object)
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Add
    
    ' 自动将顶部的 CM 参数换算为 Word 需要的磅 (1 cm = 28.35 磅)
    With wdDoc.PageSetup
        .Orientation = g_PageOrientation
        .TopMargin = CFG_MARGIN_TOP_CM * 28.35
        .BottomMargin = CFG_MARGIN_BOTTOM_CM * 28.35
        .LeftMargin = CFG_MARGIN_LEFT_CM * 28.35
        .RightMargin = CFG_MARGIN_RIGHT_CM * 28.35
    End With
End Sub

    ' 自动命名 保存文件
Private Sub SaveExportedDocument(ByVal wdApp As Object, ByVal wdDoc As Object, ByVal firstTitle As String, ByVal hasMajor1 As Boolean, ByVal hasMajor2 As Boolean)
    Dim fPath As String, fName As String, baseSuffix As String
    Dim regEx As Object
     
    fName = ActiveWorkbook.Name
    If InStrRev(fName, ".") > 0 Then fName = Left(fName, InStrRev(fName, ".") - 1)
    
    If InStr(fName, "工程量清单") > 0 Then
        baseSuffix = Mid(fName, InStr(fName, "工程量清单") + Len("工程量清单"))
    Else
        baseSuffix = ""
    End If
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = "\d{4}[\.-]\d{1,2}[\.-]\d{1,2}|\d{8}"
    baseSuffix = regEx.Replace(baseSuffix, "")
    
    fName = CFG_FILE_PREFIX & Format(Date, "yyyymmdd") & CFG_FILE_MID & baseSuffix
    
    Dim isBothCompleted As Boolean
    isBothCompleted = (hasMajor1 And hasMajor2)
    
    If Not isBothCompleted Then
        If InStr(firstTitle, CFG_MAJOR_1) > 0 Then
            fName = Replace(fName, CFG_FILE_TARGET, CFG_MAJOR_1)
        ElseIf InStr(firstTitle, CFG_MAJOR_2) > 0 Then
            fName = Replace(fName, CFG_FILE_TARGET, CFG_MAJOR_2)
        Else
            fName = Replace(fName, CFG_FILE_TARGET, firstTitle)
        End If
    End If
    
    fName = fName & ".docx"
    fPath = ActiveWorkbook.Path & "\" & fName
    
    Application.DisplayAlerts = False
    On Error Resume Next
    wdDoc.SaveAs2 fPath, 16
    wdApp.Visible = False
    wdApp.ScreenUpdating = True
    wdDoc.Activate
    DoEvents
    If Err.Number <> 0 Then
        MsgBox "保存失败，请检查文件是否已打开或路径权限！" & vbCrLf & "尝试保存的路径为: " & fPath, vbCritical
    Else
        ' 如果只完成了一个专业（或中途取消），弹出单专业完成提示
        If Not isBothCompleted Then
            If InStr(firstTitle, CFG_MAJOR_1) > 0 Then
                MsgBox "已成功转换【" & CFG_MAJOR_1 & "】专业的清单！", vbInformation, "转换完成"
            ElseIf InStr(firstTitle, CFG_MAJOR_2) > 0 Then
                MsgBox "已成功转换【" & CFG_MAJOR_2 & "】专业的清单！", vbInformation, "转换完成"
            Else
                MsgBox "已成功转换【" & firstTitle & "】专业的清单！", vbInformation, "转换完成"
            End If
        End If
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub ApplyWordTableStyles(ByVal wdTable As Object, ByVal wdApp As Object, ByVal tCols As Long, ByVal wHeaderRows As Long, ByRef dataArr() As Variant, ByVal srcCols As Long, ByRef excelToWordCol() As Long, ByRef isQtyCol() As Boolean, ByVal tmpTab As String, ByVal tmpLf As String)
    Dim c As Long, r As Long, wC As Long
    With wdTable.Range
        .Font.NameFarEast = CFG_FONT_CN
        .Font.NameAscii = CFG_FONT_EN
        .Font.Size = CFG_SIZE_TABLE
        .Font.Bold = False
        .ParagraphFormat.Alignment = 0
    End With
    
    Dim titleCache() As String: ReDim titleCache(1 To tCols)
    Dim isWordColQty() As Boolean: ReDim isWordColQty(1 To tCols)
    For c = 1 To srcCols
        If excelToWordCol(c) > 0 Then
            wC = excelToWordCol(c)
            titleCache(wC) = dataArr(1, c)
            If isQtyCol(c) Then isWordColQty(wC) = True
        End If
    Next c
    
    For c = 1 To tCols
        If isWordColQty(c) Or InStr(titleCache(c), "序号") > 0 Or InStr(titleCache(c), "单位") > 0 Or InStr(titleCache(c), "数量") > 0 Then
            wdTable.Columns(c).Select
            wdApp.Selection.ParagraphFormat.Alignment = 1
        End If
    Next c
    
    For r = 1 To wHeaderRows
        wdTable.rows(r).Range.ParagraphFormat.Alignment = 1
        wdTable.rows(r).Range.Font.Bold = True
    Next r
    
    With wdTable.Range.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = tmpTab: .Replacement.Text = "^t": .Execute Replace:=2
        .Text = tmpLf: .Replacement.Text = "^p": .Execute Replace:=2
    End With
    
    With wdTable
        .Borders.InsideLineStyle = 1: .Borders.InsideLineWidth = 4: .Borders.InsideColor = 0
        .Borders.OutsideLineStyle = 1: .Borders.OutsideLineWidth = 8: .Borders.OutsideColor = 0
        .Range.Cells.VerticalAlignment = 1
        .rows.AllowBreakAcrossPages = False
        ApplyMinimumRowHeight wdTable
        On Error Resume Next
        .cell(1, 1).Select
        wdApp.Selection.rows.HeadingFormat = True
        On Error GoTo 0
    End With
End Sub

Private Sub ApplyMinimumRowHeight(ByVal wdTable As Object)
    On Error Resume Next
    With wdTable.Range.Cells
        .HeightRule = 1
        .Height = CFG_ROW_HEIGHT_CM * 28.35
    End With
    With wdTable.rows
        .HeightRule = 1
        .Height = CFG_ROW_HEIGHT_CM * 28.35
    End With
            
        ' 强制控制标题区域的总高度，使用配置区的常量
Dim r1 As Long
    Dim totalTargetHeightCm As Single: totalTargetHeightCm = CFG_TITLE_ROW_HEIGHT_CM
        Dim perRowHeight As Single
        perRowHeight = totalTargetHeightCm / excelHeaderRowsCount
        
        For r1 = 1 To excelHeaderRowsCount
            ' 【关键修改】：这里直接使用 wdTable.Rows，不要用点开头的 .Rows
            With wdTable.rows(r1)
                .HeightRule = 1 ' wdRowHeightExactly
                .Height = perRowHeight * 28.35
            End With
        Next r1
    On Error GoTo 0
End Sub

Private Sub ApplyCustomColumnWidths(ByVal wdTable As Object, ByVal pageOrientation As Long, ByRef dataArr() As Variant, ByRef excelToWordCol() As Long, ByVal srcCols As Long, ByVal headerRows As Long)
    Dim c As Long, r As Long, wCol As Long
    Dim colNo As Long, colName As Long, colSpec As Long, colUnit As Long
    Dim colQty1 As Long, colQty2 As Long, colWt1 As Long, colWt2 As Long, colMemo As Long
    
    wdTable.AllowAutoFit = False
    
    ' 1. 识别对应列在 Word 表格中的实际列号
    For c = 1 To srcCols
        wCol = excelToWordCol(c)
        If wCol > 0 Then
            Dim headerText As String: headerText = ""
            For r = 1 To headerRows
                headerText = headerText & "|" & dataArr(r, c)
            Next r
            If InStr(headerText, "序号") > 0 Then colNo = wCol
            If InStr(headerText, "设备名称") > 0 Or InStr(headerText, "项目名称") > 0 Then colName = wCol
            If InStr(headerText, "技术性能") > 0 Or InStr(headerText, "技术特征") > 0 Then colSpec = wCol
            If InStr(headerText, "单位") > 0 Then colUnit = wCol
            If InStr(headerText, "备注") > 0 Then colMemo = wCol
            
            If InStr(headerText, "一个机组") > 0 Then
                colQty1 = wCol
            ElseIf InStr(headerText, "总计") > 0 And InStr(headerText, "重量") = 0 Then
                colQty2 = wCol
            ElseIf InStr(headerText, "数量") > 0 Or InStr(headerText, "工程量") > 0 Then
                If colQty1 = 0 Then colQty1 = wCol
            End If
            
            If InStr(headerText, "单重") > 0 Then colWt1 = wCol
            If InStr(headerText, "合重") > 0 Then colWt2 = wCol
        End If
    Next c
    
    ' 2. 根据判断规则，选择全局配置中的哪种策略矩阵
    Dim cfgStr As String
    If pageOrientation = 0 Then ' 纵向
        If colQty2 > 0 Then
            cfgStr = CFG_WIDTH_PORTRAIT_A
        Else
            cfgStr = CFG_WIDTH_PORTRAIT_B
        End If
    Else ' 横向
        If colWt1 > 0 Or colWt2 > 0 Then
            If colQty2 > 0 Then
                cfgStr = CFG_WIDTH_LANDSCAPE_A
            Else
                cfgStr = CFG_WIDTH_LANDSCAPE_B
            End If
        Else
            If colQty2 > 0 Then
                cfgStr = CFG_WIDTH_LANDSCAPE_C
            Else
                cfgStr = CFG_WIDTH_LANDSCAPE_D
            End If
        End If
    End If
    
    ' 3. 解析配置并应用宽度 (将CM换算为磅)
    Dim wArr() As String
    wArr = Split(cfgStr, ",")
    
    Dim w_No As Double: w_No = Val(wArr(0))
    Dim w_Name As Double: w_Name = Val(wArr(1))
    Dim w_Spec As Double: w_Spec = Val(wArr(2))
    Dim w_Unit As Double: w_Unit = Val(wArr(3))
    Dim w_Qty1 As Double: w_Qty1 = Val(wArr(4))
    Dim w_Qty2 As Double: w_Qty2 = Val(wArr(5))
    Dim w_Wt1 As Double: w_Wt1 = Val(wArr(6))
    Dim w_Wt2 As Double: w_Wt2 = Val(wArr(7))
    Dim w_Memo As Double: w_Memo = Val(wArr(8))
    
    On Error Resume Next
    If colNo > 0 And w_No > 0 Then wdTable.Columns(colNo).Width = w_No * 28.35
    If colName > 0 And w_Name > 0 Then wdTable.Columns(colName).Width = w_Name * 28.35
    If colSpec > 0 And w_Spec > 0 Then wdTable.Columns(colSpec).Width = w_Spec * 28.35
    If colUnit > 0 And w_Unit > 0 Then wdTable.Columns(colUnit).Width = w_Unit * 28.35
    If colQty1 > 0 And w_Qty1 > 0 Then wdTable.Columns(colQty1).Width = w_Qty1 * 28.35
    If colQty2 > 0 And w_Qty2 > 0 Then wdTable.Columns(colQty2).Width = w_Qty2 * 28.35
    If colWt1 > 0 And w_Wt1 > 0 Then wdTable.Columns(colWt1).Width = w_Wt1 * 28.35
    If colWt2 > 0 And w_Wt2 > 0 Then wdTable.Columns(colWt2).Width = w_Wt2 * 28.35
    If colMemo > 0 And w_Memo > 0 Then wdTable.Columns(colMemo).Width = w_Memo * 28.35
    On Error GoTo 0
End Sub

Function RemoveSizeInfo(ByVal inputStr As String, ByVal patterns As Variant) As String
    Static regEx As Object
    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True: regEx.IgnoreCase = True
    End If
    Dim p As Variant
    For Each p In patterns
        regEx.Pattern = CStr(p)
        inputStr = regEx.Replace(inputStr, "")
    Next p
    RemoveSizeInfo = inputStr
End Function

Function CleanString(ByVal inputStr As String) As String
    Static regEx As Object
    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.Pattern = "[^\a-zA-Z0-9\u4e00-\u9fa5]"
    End If
    CleanString = regEx.Replace(inputStr, "")
End Function

'规格专用清洗函数
Private Function CleanSpecWithKeywords(ByVal inputStr As String, ByVal keywords As Variant) As String
    Dim kw As Variant
    Dim cleanRes As String: cleanRes = ""
    
    ' 如果输入为空，直接返回
    If Trim(inputStr) = "" Then
        CleanSpecWithKeywords = ""
        Exit Function
    End If
    
    ' 遍历关键词白名单
    For Each kw In keywords
        If Trim(kw) <> "" Then
            ' 不区分大小写比对：如果规格中包含该关键词
            If InStr(1, inputStr, kw, vbTextCompare) > 0 Then
                cleanRes = kw
                Exit For ' 匹配到第一个关键词后立即退出，防止冲突
            End If
        End If
    Next kw
    
    CleanSpecWithKeywords = cleanRes
End Function

'过滤关键字行
Private Function IsExcludedName(ByVal cellText As String) As Boolean
    Dim excludeArr() As String
    Dim i As Long
    
    ' 如果配置为空，则默认不拦截
    If CFG_EXCLUDE_NAMES = "" Then
        IsExcludedName = False
        Exit Function
    End If
    
    ' 拆分逗号隔开的关键字配置
    excludeArr = Split(CFG_EXCLUDE_NAMES, ",")
    For i = 0 To UBound(excludeArr)
        ' 只要项目名称中包含该关键字，就触发拦截 (忽略空格影响)
        If Trim(excludeArr(i)) <> "" And InStr(cellText, Trim(excludeArr(i))) > 0 Then
            IsExcludedName = True
            Exit Function
        End If
    Next i
    
    IsExcludedName = False
End Function
' ==========================================
' 辅助函数：判断文本是否为汉字序号 (如 一、二、(一)、（二）)
' ==========================================
Function IsChineseHeading(ByVal txt As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' 先清理文本中的半角和全角空格（防止因排版空格导致匹配失败）
    txt = Replace(txt, " ", "")
    txt = Replace(txt, "　", "")
    
    ' 正则表达式匹配：纯汉字数字，或者带括号的汉字数字，或者带顿号的汉字数字
    regEx.Pattern = "^([一二三四五六七八九十]+|[（\(][一二三四五六七八九十]+[）\)]|[一二三四五六七八九十]+、)$"
    IsChineseHeading = regEx.Test(txt)
End Function


