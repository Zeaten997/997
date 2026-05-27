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

' 6. 获取需要过滤的尺寸正则规则 (由于VBA不支持数组常量，用函数返回)
Private Function GetSizePatterns() As Variant
    GetSizePatterns = Array("DN\d+[，。、；：,.;:]?", "D\d+[Xx×]\d+[，。、；：,.;:]?", _
                            "φ\d+[Xx×]\d+[，。、；：,.;:]?", "Φ\d+[Xx×]\d+[，。、；：,.;:]?", "统计IO用")
End Function


' ==========================================================
' 第三部分：主控流程 (逻辑编排)
' ==========================================================
Sub StartExportProcess(isAll As Boolean)
    Dim selRange As Range
    Set selRange = Selection
    
    If selRange.Cells.Count = 1 Then
        MsgBox "请先在Excel中框选需要导出的工程量清单区域！", vbExclamation
        Exit Sub
    End If
    
    ' 1. 初始化 Word 并设置页面
    Dim wdApp As Object, wdDoc As Object
    Call InitWordDocument(wdApp, wdDoc)
    
    Application.ScreenUpdating = False
    wdApp.ScreenUpdating = False
    
    ' 2. 核心状态追踪
    Dim firstTitle As String, currentTitle As String
    Dim hasMajor1 As Boolean, hasMajor2 As Boolean
    
    ' 3. 处理首个表格
    currentTitle = ProcessSingleTable(selRange, wdApp, wdDoc)
    firstTitle = currentTitle
    
    If InStr(currentTitle, CFG_MAJOR_1) > 0 Then hasMajor1 = True
    If InStr(currentTitle, CFG_MAJOR_2) > 0 Then hasMajor2 = True
    
    ' 4. 全部转换模式下的循环抓取
    If isAll Then
        If Not (hasMajor1 And hasMajor2) Then
            Dim nextRange As Range
            Do
                Application.ScreenUpdating = True
                On Error Resume Next
                Set nextRange = Application.InputBox("已完成当前专业转换。" & vbCrLf & _
                    "请用鼠标框选【下一个专业】的工程量清单区域（可点击下方标签切换工作表）。" & vbCrLf & _
                    "(点击取消结束并保存文档)", "全部转换 - 选择下一个区域", Type:=8)
                On Error GoTo 0
                
                If nextRange Is Nothing Then Exit Do
                
                Application.ScreenUpdating = False
                wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1).InsertBreak Type:=7
                
                currentTitle = ProcessSingleTable(nextRange, wdApp, wdDoc)
                
                If InStr(currentTitle, CFG_MAJOR_1) > 0 Then hasMajor1 = True
                If InStr(currentTitle, CFG_MAJOR_2) > 0 Then hasMajor2 = True
                
                If hasMajor1 And hasMajor2 Then
                    MsgBox "已成功转换【" & CFG_MAJOR_1 & "】和【" & CFG_MAJOR_2 & "】两个专业的清单，全部转换完成！", vbInformation, "转换完成"
                    Exit Do
                End If
                Set nextRange = Nothing
            Loop
        Else
            MsgBox "已成功转换【" & CFG_MAJOR_1 & "】和【" & CFG_MAJOR_2 & "】两个专业的清单，全部转换完成！", vbInformation, "转换完成"
        End If
    End If
    
    ' 5. 执行保存
    Call SaveExportedDocument(wdApp, wdDoc, firstTitle, hasMajor1, hasMajor2)
    
    wdApp.ScreenUpdating = True
    Application.ScreenUpdating = True
End Sub

' ==========================================================
' 第四部分：核心业务 (表格数据提取与 Word 排版)
' ==========================================================
Function ProcessSingleTable(ByVal selRange As Range, ByVal wdApp As Object, ByVal wdDoc As Object) As String
    Dim titleText As String
    Dim firstCellText As String
    
    ' 1. 获取并插入一级标题
    firstCellText = Replace(selRange.Cells(1, 1).MergeArea.Cells(1, 1).Value, " ", "")
    If InStr(firstCellText, "工程量清单") > 0 Then
        titleText = Replace(firstCellText, "工程量清单", "")
        Set selRange = selRange.Offset(1, 0).Resize(selRange.Rows.Count - 1, selRange.Columns.Count)
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
    srcRows = selRange.Rows.Count
    srcCols = selRange.Columns.Count
    
    Dim excelHeaderRowsCount As Long: excelHeaderRowsCount = 1
    For c = 1 To srcCols
        If Not selRange.Columns(c).Hidden Then
            If selRange.Cells(1, c).MergeCells Then
                Dim areaRows As Long: areaRows = selRange.Cells(1, c).MergeArea.Rows.Count
                If areaRows > excelHeaderRowsCount Then excelHeaderRowsCount = areaRows
            End If
        End If
    Next c
    
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
                txt = CStr(selRange.Cells(r, c).Value)
                txt = Replace(txt, vbTab, tmpTab)
                txt = Replace(txt, vbCrLf, tmpLf)
                txt = Replace(txt, vbLf, tmpLf)
                
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
            Dim headVal As String: headVal = selRange.Cells(1, c).MergeArea.Cells(1, 1).Value
            If headVal = "工程量" Or headVal = "数量" Or dataArr(1, c) = "数量" Then isQtyCol(c) = True
        End If
    Next c
    
    If colIdxName > 0 And colIdxSpec > 0 Then
        For r = (excelHeaderRowsCount + 1) To srcRows
            If Left(dataArr(r, colIdxName), 1) = "*" Or InStr(dataArr(r, colIdxName), "成套") > 0 Then
                dataArr(r, colIdxSpec) = "设备成套提供"
            End If
        Next r
        For r = srcRows To (excelHeaderRowsCount + 2) Step -1
            If Not (selRange.Rows(r).Hidden Or selRange.Rows(r - 1).Hidden) Then
                If CleanString(dataArr(r, colIdxName)) = CleanString(dataArr(r - 1, colIdxName)) And _
                   CleanString(dataArr(r, colIdxSpec)) = CleanString(dataArr(r - 1, colIdxSpec)) And _
                   CleanString(dataArr(r, colIdxName)) <> "" Then
                    Dim subCol As Long
                    For subCol = 1 To srcCols
                        If isQtyCol(subCol) Then dataArr(r - 1, subCol) = Val(dataArr(r, subCol)) + Val(dataArr(r - 1, subCol))
                    Next subCol
                    dataArr(r, colIdxName) = "[MERGED_ROW]"
                    If colIdxMemo > 0 Then dataArr(r, colIdxMemo) = "": dataArr(r - 1, colIdxMemo) = ""
                End If
            End If
        Next r
    End If
    
    ' 5. 输出数据到 Word 表格
    Dim excelToWordRow() As Long: ReDim excelToWordRow(1 To srcRows)
    Dim excelToWordCol() As Long: ReDim excelToWordCol(1 To srcCols)
    Dim tRows As Long, wHeaderRows As Long, tCols As Long
    
    For r = 1 To srcRows
        If Not selRange.Rows(r).Hidden And dataArr(r, IIf(colIdxName > 0, colIdxName, 1)) <> "[MERGED_ROW]" Then
            tRows = tRows + 1: excelToWordRow(r) = tRows
            If r <= excelHeaderRowsCount Then wHeaderRows = wHeaderRows + 1
        End If
    Next r
    For c = 1 To srcCols
        If Not selRange.Columns(c).Hidden Then tCols = tCols + 1: excelToWordCol(c) = tCols
    Next c
    
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
    
    ' 7. 重建合并单元格
    For r = srcRows To 1 Step -1
        For c = srcCols To 1 Step -1
            Dim cellExcel As Range: Set cellExcel = selRange.Cells(r, c)
            If cellExcel.MergeCells And cellExcel.Address = cellExcel.MergeArea.Cells(1, 1).Address Then
                Dim startWRow As Long: startWRow = excelToWordRow(r)
                Dim startWCol As Long: startWCol = excelToWordCol(c)
                If startWRow > 0 And startWCol > 0 Then
                    Dim endWRow As Long, endWCol As Long, i As Long
                    For i = r + cellExcel.MergeArea.Rows.Count - 1 To r Step -1
                        If excelToWordRow(i) > 0 Then endWRow = excelToWordRow(i): Exit For
                    Next i
                    For i = c + cellExcel.MergeArea.Columns.Count - 1 To c Step -1
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
    
    wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1).Select
End Function

' ==========================================================
' 第五部分：独立的工具与服务层接口 (解耦具体操作)
' ==========================================================
Private Sub InitWordDocument(ByRef wdApp As Object, ByRef wdDoc As Object)
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    wdApp.Visible = True
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
    wdApp.Visible = True
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
        wdTable.Rows(r).Range.ParagraphFormat.Alignment = 1
        wdTable.Rows(r).Range.Font.Bold = True
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
        .Rows.AllowBreakAcrossPages = False
        .Rows.HeightRule = 1
        .Rows.Height = CFG_ROW_HEIGHT_CM * 28.35
        On Error Resume Next
        .cell(1, 1).Select
        wdApp.Selection.Rows.HeadingFormat = True
        On Error GoTo 0
    End With
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

