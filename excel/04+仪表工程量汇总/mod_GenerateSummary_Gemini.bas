Attribute VB_Name = "mod_GenerateSummary_Gemini"
Option Explicit '20260529 V0.1

' ==========================================
' 7.3 参数设置区 (全局参数)
' ==========================================
Dim TechSpecKeywords As Variant
Dim SortKeywords As Variant
Dim RemarkClearKeywords As Variant
Dim ExcludeKeywords As Variant
Dim TripleGroupMap As Object

Private Sub InitParams()
    ' 4.1 需要区分技术特征的关键字库 (若特征中包含这些词，即使设备名相同也单独汇总)
    TechSpecKeywords = Array("防爆", "隔爆", "本安")  ' 示例：可根据实际需求增删
    
    ' 5.2.1 排序关键字列表
    SortKeywords = Array("温度计", "热电阻", "热电偶", "压力表", "变送器", "流量计", "物位", "料位", "液位", "开关阀", "调节阀", "快切阀", "泄露", "探测器", "分析仪")
    
    ' 5.3 需要清除备注的关键字
    RemarkClearKeywords = Array("流量计", "阀")
    
    ' 3.3 排除汇总的关键字 (如果项目名称包含这些词，则排除)
    ExcludeKeywords = Array("*", "成套")
    ' 初始化三字符聚类映射（用于将包含相同连续三字的设备分为一组）
    Set TripleGroupMap = CreateObject("Scripting.Dictionary")
End Sub

' ==========================================
' 主程序
' ==========================================
Sub GenerateSummary_gemini()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 1. 初始化所有字典与参数
    InitParams
    
    Dim wsSource As Worksheet
    Set wsSource = ActiveSheet
    
    ' 2. 寻找表头及其对应的列号
    Dim headerRow As Long, colNo As Integer, colName As Integer, colSpec As Integer
    Dim colUnit As Integer, colQty As Integer, colRemark As Integer
    Dim r As Long, c As Integer, cellVal As String
    
    For r = 1 To 10
        For c = 1 To 20
            If Trim(wsSource.Cells(r, c).Value) = "项目名称" Then
                headerRow = r
                colName = c
                Exit For
            End If
        Next c
        If headerRow > 0 Then Exit For
    Next r
    
    If headerRow = 0 Then
        MsgBox "未找到包含'项目名称'的表头行，请检查表格格式！", vbExclamation
        GoTo ExitSub
    End If
    
    ' 获取其它列号
    For c = 1 To 20
        cellVal = Trim(wsSource.Cells(headerRow, c).Value)
        If cellVal = "序号" Then colNo = c
        If cellVal = "项目技术特征" Then colSpec = c
        If cellVal = "计量单位" Then colUnit = c
        If cellVal = "备注" Then colRemark = c
        If cellVal = "工程量" Then
            colQty = c
            ' 2. 如果工程量是个合并单元格，寻找"总计"子列
            If wsSource.Cells(headerRow, c).MergeCells Then
                Dim mergeArea As Range, cc As Integer
                Set mergeArea = wsSource.Cells(headerRow, c).mergeArea
                For cc = mergeArea.Column To mergeArea.Column + mergeArea.Columns.Count - 1
                        Dim hdr As String
                        hdr = Trim(wsSource.Cells(headerRow + 1, cc).Value)
                        If hdr = "总计" Or hdr = "合计" Then
                            colQty = cc
                            Exit For
                        End If
                Next cc
            End If
        End If
    Next c

    ' 必要列校验（防止后续因找不到列而崩溃）
    If colNo = 0 Or colSpec = 0 Or colUnit = 0 Or colQty = 0 Or colRemark = 0 Then
        MsgBox "未能识别必要的表头列（序号/项目技术特征/计量单位/工程量/备注），请检查表头！", vbExclamation
        GoTo ExitSub
    End If
    
    ' 3. 初始化字典对象
    Dim dictPart1 As Object, dictPart3 As Object
    Set dictPart1 = CreateObject("Scripting.Dictionary")
    Set dictPart3 = CreateObject("Scripting.Dictionary")
    
    Dim part2Row1 As Variant, part2Row2 As Variant
    Dim currentPart As Integer: currentPart = 1
    Dim currentPart3Cat As String: currentPart3Cat = ""
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.rows.Count, colName).End(xlUp).Row
    
    ' 4. 遍历数据
    For r = headerRow + 1 To lastRow
        Dim projectName As String
        projectName = Trim(CStr(wsSource.Cells(r, colName).Value))
        
        If projectName = "" Then GoTo NextRow
        ' 3.2 遇到材料及安装，直接终止
        If InStr(projectName, "材料及安装") > 0 Then Exit For
        
        ' ---------- 区域判定与过渡逻辑 ----------
        If currentPart = 1 Then
            ' 5.2.1 统计至计算机控制系统前的数据
            If projectName = "计算机控制系统" Or InStr(projectName, "计算机控制系统软硬件") > 0 Then
                currentPart = 2
            End If
        End If
        
        If currentPart = 2 Then
            ' 5.2.2 仅保留特定的两行
            If InStr(projectName, "计算机控制系统软硬件") > 0 Then
                part2Row1 = ExtractRowData(wsSource, r, colName, colSpec, colUnit, colQty, colRemark)
            ElseIf InStr(projectName, "计算机控制系统应用软件") > 0 Then
                part2Row2 = ExtractRowData(wsSource, r, colName, colSpec, colUnit, colQty, colRemark)
                currentPart = 3 ' 从下一行开始进入第三部分
            End If
            GoTo NextRow ' 第二部分的其他杂项行直接跳过
        End If
        
        ' 第三大类主类名提取
        If currentPart = 3 Then
            Dim rawQty As String, rawUnit As String
            rawQty = Trim(CStr(wsSource.Cells(r, colQty).Value))
            rawUnit = Trim(CStr(wsSource.Cells(r, colUnit).Value))
            ' 提取第三部分汉字大类：没有单位、工程量，且不是小计/合计等
            If rawUnit = "" And rawQty = "" And Not (projectName Like "*计*" Or projectName Like "*费*") Then
                currentPart3Cat = projectName
                If Not dictPart3.Exists(currentPart3Cat) Then
                    Dim subD As Object
                    Set subD = CreateObject("Scripting.Dictionary")
                    dictPart3.Add currentPart3Cat, subD
                End If
                GoTo NextRow
            End If
        End If
        
        ' ---------- 第1和第3部分 常规数据收集 ----------
        Dim qtyStr As String
        qtyStr = Trim(CStr(wsSource.Cells(r, colQty).Value))
        
        ' 3.1 没有数值则不汇总
        If qtyStr = "" Or qtyStr = "-" Or Val(qtyStr) = 0 Then GoTo NextRow

        ' 3.3 排除关键字判定
        Dim isExclude As Boolean: isExclude = False
        Dim k As Variant
        For Each k In ExcludeKeywords
            If InStr(projectName, k) > 0 Then isExclude = True: Exit For
        Next k
        If isExclude Then GoTo NextRow
        
        Dim spec As String, unitStr As String, remarkStr As String
        spec = Trim(CStr(wsSource.Cells(r, colSpec).Value))
        unitStr = Trim(CStr(wsSource.Cells(r, colUnit).Value))
        remarkStr = Trim(CStr(wsSource.Cells(r, colRemark).Value))
        
        ' 4.3 管径清洗与标点处理
        Dim cleanSp As String
        cleanSp = CleanTechSpec(spec)
        
        ' 4.1 技术特征关键字区分判断
        Dim hasSpKw As Boolean: hasSpKw = False
        For Each k In TechSpecKeywords
            If InStr(cleanSp, k) > 0 Then hasSpKw = True: Exit For
        Next k
        
        Dim groupKey As String
        If hasSpKw Then
            groupKey = projectName & "|SPECMATCH|" & cleanSp
        Else
            groupKey = projectName & "|NOSPEC"
        End If
        
        ' 确定存放的字典
        Dim targetDict As Object
        If currentPart = 1 Then
            Set targetDict = dictPart1
        Else
            If currentPart3Cat = "" Then currentPart3Cat = "未分类"
            If Not dictPart3.Exists(currentPart3Cat) Then
                Dim subD2 As Object
                Set subD2 = CreateObject("Scripting.Dictionary")
                dictPart3.Add currentPart3Cat, subD2
            End If
            Set targetDict = dictPart3(currentPart3Cat)
        End If
        
        ' 数据写入或累加
        If targetDict.Exists(groupKey) Then
            Dim itm As Variant
            itm = targetDict(groupKey)
            itm(3) = itm(3) + Val(qtyStr) ' 累加数量
            ' 4.2 若合并项目中特征不同，进行标记(后续清空)
            If itm(1) <> cleanSp Then itm(6) = True
            targetDict(groupKey) = itm
        Else
            Dim sortOrd As Integer
            sortOrd = GetSortOrder(projectName)
            
            ' 5.3 遇到流量计、阀等，清除备注
            For Each k In RemarkClearKeywords
                If InStr(projectName, k) > 0 Then remarkStr = "": Exit For
            Next k
            
            ' Array 结构: 0:名称, 1:特征, 2:单位, 3:数量, 4:备注, 5:排序权重, 6:特征是否冲突清空
            targetDict.Add groupKey, Array(projectName, cleanSp, unitStr, Val(qtyStr), remarkStr, sortOrd, False)
        End If

NextRow:
    Next r
    
    ' 5. 输出并格式化生成的表格
    Call WriteSummarySheet(wsSource, colNo, colName, colSpec, colUnit, colQty, colRemark, dictPart1, part2Row1, part2Row2, dictPart3)
ExitSub:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' ==========================================
' 小功能1：技术特征清洗 (利用正则剔除管径口径及残留标点)
' ==========================================
Function CleanTechSpec(ByVal spec As String) As String
    If Len(Trim(spec)) = 0 Then
        CleanTechSpec = ""
        Exit Function
    End If
    
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.IgnoreCase = True
    ' 匹配 DN带数字 或 Φ带数字并包含可选厚度(x或*)及小数点的管道
    reg.Pattern = "(DN\d+|[Φφ]\d+(\.\d+)?([xX*]\d+(\.\d+)?)?)"
    spec = reg.Replace(spec, "")
    
    ' 清除替换后残留的多余逗号或顿号
    reg.Pattern = "\s*[,，、]+\s*"
    spec = reg.Replace(spec, ",")
    
    spec = Trim(spec)
    ' 剥除首尾残留的标点
    Do While Left(spec, 1) = "," Or Left(spec, 1) = "，" Or Left(spec, 1) = "、"
        spec = Mid(spec, 2)
    Loop
    Do While Right(spec, 1) = "," Or Right(spec, 1) = "，" Or Right(spec, 1) = "、"
        spec = Left(spec, Len(spec) - 1)
    Loop
    
    CleanTechSpec = Trim(spec)
End Function

' 新增：提取字符串中首次出现的连续三个中文字符子串（用于聚类）
Function ExtractChineseTriplet(ByVal s As String) As String
    Dim i As Long
    Dim sLen As Long
    Dim ch1 As String, ch2 As String, ch3 As String
    sLen = Len(s)
    If sLen < 3 Then
        ExtractChineseTriplet = ""
        Exit Function
    End If
    For i = 1 To sLen - 2
        ch1 = Mid(s, i, 1)
        ch2 = Mid(s, i + 1, 1)
        ch3 = Mid(s, i + 2, 1)
        If AscW(ch1) > 255 And AscW(ch2) > 255 And AscW(ch3) > 255 Then
            ExtractChineseTriplet = ch1 & ch2 & ch3
            Exit Function
        End If
    Next i
    ExtractChineseTriplet = ""
End Function

' 提取所有连续三个中文字符子串（去重），返回数组（可能为空数组）
Function ExtractAllChineseTriplets(ByVal s As String) As Variant
    Dim i As Long, sLen As Long
    Dim ch1 As String, ch2 As String, ch3 As String, triple As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    sLen = Len(Trim(s))
    If sLen < 3 Then
        ExtractAllChineseTriplets = Array()
        Exit Function
    End If
    For i = 1 To sLen - 2
        ch1 = Mid(s, i, 1)
        ch2 = Mid(s, i + 1, 1)
        ch3 = Mid(s, i + 2, 1)
        If AscW(ch1) > 255 And AscW(ch2) > 255 And AscW(ch3) > 255 Then
            triple = ch1 & ch2 & ch3
            If Not d.Exists(triple) Then d.Add triple, True
        End If
    Next i
    If d.Count = 0 Then
        ExtractAllChineseTriplets = Empty
    Else
        Dim arr() As String
        ReDim arr(0 To d.Count - 1)
        Dim idx As Integer: idx = 0
        Dim k As Variant
        For Each k In d.keys
            arr(idx) = k
            idx = idx + 1
        Next k
        ExtractAllChineseTriplets = arr
    End If
End Function

' 可选：用户提供的两字符串三连中文子串匹配函数（保留以备需要）
Private Function HasThreeConsecutiveChars(ByVal s1 As String, ByVal s2 As String) As Boolean
    Dim i As Integer, subStr As String
    If Len(Trim(s1)) < 3 Or Len(Trim(s2)) < 3 Then
        HasThreeConsecutiveChars = False
        Exit Function
    End If
    For i = 1 To Len(s1) - 2
        subStr = Mid(s1, i, 3)
        If AscW(Left(subStr, 1)) > 255 Then
            If InStr(1, s2, subStr) > 0 Then
                HasThreeConsecutiveChars = True
                Exit Function
            End If
        End If
    Next i
    HasThreeConsecutiveChars = False
End Function

' ==========================================
' 小功能2：获取排序优先级权重
' ==========================================
Function GetSortOrder(ByVal projectName As String) As Integer
    Dim i As Integer
    Dim triples As Variant
    Dim t As Variant
    Dim nextWeight As Integer

    ' 优先按预设关键词排序（保持原有行为）
    For i = LBound(SortKeywords) To UBound(SortKeywords)
        If InStr(projectName, SortKeywords(i)) > 0 Then
            GetSortOrder = i + 1
            Exit Function
        End If
    Next i

    ' 获取所有不重复的连续三中文字符子串
    triples = ExtractAllChineseTriplets(projectName)
    If Not IsEmpty(triples) Then
        ' 若已有任一子串在映射中，优先使用该映射权重（保持已有分组一致性）
        For Each t In triples
            If t <> "" Then
                If TripleGroupMap.Exists(t) Then
                    GetSortOrder = TripleGroupMap(t)
                    Exit Function
                End If
            End If
        Next t
        ' 否则为该名称分配一个新的组权重，并将该名称的所有子串映射到该权重
        nextWeight = 100 + TripleGroupMap.Count + 1
        For Each t In triples
            If Not TripleGroupMap.Exists(t) Then TripleGroupMap.Add t, nextWeight
        Next t
        GetSortOrder = nextWeight
        Exit Function
    End If

    GetSortOrder = 999 ' 未列入关键字，优先级靠后
End Function

' ==========================================
' 小功能3：提取第二部分的特殊整行数据（应对可能出现的合并单元格空值）
' ==========================================
Function ExtractRowData(ws As Worksheet, r As Long, cName As Integer, cSpec As Integer, cUnit As Integer, cQty As Integer, cRemark As Integer) As Variant
    Dim arr(0 To 4) As Variant
    arr(0) = Trim(ws.Cells(r, cName).Value)
    arr(1) = Trim(ws.Cells(r, cSpec).Value)
    arr(2) = Trim(ws.Cells(r, cUnit).Value)
    
    Dim qty As String
    qty = Trim(ws.Cells(r, cQty).Value)
    ' 应对套件工程量被填写在主合并列而非“总计”子列的情况
    If qty = "" And cQty > 1 Then qty = Trim(ws.Cells(r, cQty - 1).Value)
    arr(3) = qty
    arr(4) = Trim(ws.Cells(r, cRemark).Value)
    
    ExtractRowData = arr
End Function

' ==========================================
' 小功能4：获取汉字大类数字
' ==========================================
Function GetChineseNumeral(ByVal num As Integer) As String
    Dim arr As Variant
    arr = Array("", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五")
    If num > 0 And num <= UBound(arr) Then
        GetChineseNumeral = arr(num)
    Else
        GetChineseNumeral = CStr(num)
    End If
End Function

' ==========================================
' 小功能5：核心写入与格式排版
' ==========================================
Sub WriteSummarySheet(wsSource As Worksheet, cNo As Integer, cName As Integer, cSpec As Integer, cUnit As Integer, cQty As Integer, cRem As Integer, dict1 As Object, p2r1 As Variant, p2r2 As Variant, dict3 As Object)
    Dim wsDest As Worksheet
    Dim sheetName As String
    sheetName = "工程量汇总"
    
    ' 5.5 如果表格重名则自动后缀重命名，不覆盖
    Dim appendIdx As Long: appendIdx = 1
    Dim sht As Worksheet
    Dim existFlag As Boolean
    Do
        existFlag = False
        For Each sht In wsSource.Parent.Sheets
            If sht.Name = sheetName Then existFlag = True: Exit For
        Next sht
        If existFlag Then
            sheetName = "工程量汇总_" & appendIdx
            appendIdx = appendIdx + 1
        End If
    Loop While existFlag

    Set wsDest = wsSource.Parent.Sheets.Add(After:=wsSource)
    wsDest.Name = sheetName
    
    ' 5.1 表头
    wsDest.Cells(1, 1).Value = "序号"
    wsDest.Cells(1, 2).Value = "项目名称"
    wsDest.Cells(1, 3).Value = "项目技术特征"
    wsDest.Cells(1, 4).Value = "计量单位"
    wsDest.Cells(1, 5).Value = "工程量"
    wsDest.Cells(1, 6).Value = "备注"
    
    ' 6.2 & 6.3 表头排版
    With wsDest.Range("A1:F1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.Name = "宋体"
        .Font.Size = 10.5
    End With
    
    Dim destRow As Long: destRow = 2
    Dim mainCatIdx As Long: mainCatIdx = 1
    Dim subCatIdx As Long, startRow As Long, r As Long
    Dim k As Variant, itm As Variant
    
    ' ------------------ 写入第一部分 ------------------
    wsDest.Cells(destRow, 1).Value = GetChineseNumeral(mainCatIdx)
    wsDest.Cells(destRow, 2).Value = "现场仪表"
    wsDest.Range(wsDest.Cells(destRow, 1), wsDest.Cells(destRow, 6)).Font.Bold = True
    destRow = destRow + 1
    
    startRow = destRow
    subCatIdx = 1
    For Each k In dict1.keys
        itm = dict1(k)
        wsDest.Cells(destRow, 1).Value = subCatIdx
        wsDest.Cells(destRow, 2).Value = itm(0)
        ' 4.2 特征冲突则抹除
        wsDest.Cells(destRow, 3).Value = IIf(itm(6), "", itm(1))
        wsDest.Cells(destRow, 4).Value = itm(2)
        wsDest.Cells(destRow, 5).Value = itm(3)
        wsDest.Cells(destRow, 6).Value = itm(4)
        wsDest.Cells(destRow, 26).Value = itm(5) ' Z列写入隐形排序权重
        destRow = destRow + 1
        subCatIdx = subCatIdx + 1
    Next k
    
    ' 第一部分拼音聚类排序 (5.2.1 & 5.4)
    If destRow > startRow Then
        wsDest.Range(wsDest.Cells(startRow, 1), wsDest.Cells(destRow - 1, 26)).Sort _
            Key1:=wsDest.Cells(startRow, 26), Order1:=xlAscending, _
            Key2:=wsDest.Cells(startRow, 2), Order2:=xlAscending, _
            Header:=xlNo, SortMethod:=xlPinYin
        ' 刷新序号
        For r = startRow To destRow - 1
            wsDest.Cells(r, 1).Value = r - startRow + 1
        Next r
    End If
    
    ' ------------------ 写入第二部分 ------------------
    mainCatIdx = mainCatIdx + 1
    wsDest.Cells(destRow, 1).Value = GetChineseNumeral(mainCatIdx)
    wsDest.Cells(destRow, 2).Value = "计算机控制系统"
    wsDest.Range(wsDest.Cells(destRow, 1), wsDest.Cells(destRow, 6)).Font.Bold = True
    destRow = destRow + 1
    
    If Not IsEmpty(p2r1) Then
        wsDest.Cells(destRow, 1).Value = 1
        wsDest.Cells(destRow, 2).Value = p2r1(0)
        wsDest.Cells(destRow, 3).Value = p2r1(1)
        wsDest.Cells(destRow, 4).Value = p2r1(2)
        wsDest.Cells(destRow, 5).Value = p2r1(3)
        wsDest.Cells(destRow, 6).Value = p2r1(4)
        destRow = destRow + 1
    End If
    If Not IsEmpty(p2r2) Then
        wsDest.Cells(destRow, 1).Value = IIf(IsEmpty(p2r1), 1, 2)
        wsDest.Cells(destRow, 2).Value = p2r2(0)
        wsDest.Cells(destRow, 3).Value = p2r2(1)
        wsDest.Cells(destRow, 4).Value = p2r2(2)
        wsDest.Cells(destRow, 5).Value = p2r2(3)
        wsDest.Cells(destRow, 6).Value = p2r2(4)
        destRow = destRow + 1
    End If
    
    ' ------------------ 写入第三部分 ------------------
    For Each k In dict3.keys
        mainCatIdx = mainCatIdx + 1
        wsDest.Cells(destRow, 1).Value = GetChineseNumeral(mainCatIdx)
        wsDest.Cells(destRow, 2).Value = k
        wsDest.Range(wsDest.Cells(destRow, 1), wsDest.Cells(destRow, 6)).Font.Bold = True
        destRow = destRow + 1
        
        Dim subDict As Object: Set subDict = dict3(k)
        startRow = destRow
        subCatIdx = 1
        
        Dim sk As Variant
        For Each sk In subDict.keys
            itm = subDict(sk)
            wsDest.Cells(destRow, 1).Value = subCatIdx
            wsDest.Cells(destRow, 2).Value = itm(0)
            wsDest.Cells(destRow, 3).Value = IIf(itm(6), "", itm(1))
            wsDest.Cells(destRow, 4).Value = itm(2)
            wsDest.Cells(destRow, 5).Value = itm(3)
            wsDest.Cells(destRow, 6).Value = itm(4)
            wsDest.Cells(destRow, 26).Value = itm(5)
            destRow = destRow + 1
            subCatIdx = subCatIdx + 1
        Next sk
        
        ' 第三部分子类同样采用拼音与连续词组自适应排序
        If destRow > startRow Then
            wsDest.Range(wsDest.Cells(startRow, 1), wsDest.Cells(destRow - 1, 26)).Sort _
                Key1:=wsDest.Cells(startRow, 26), Order1:=xlAscending, _
                Key2:=wsDest.Cells(startRow, 2), Order2:=xlAscending, _
                Header:=xlNo, SortMethod:=xlPinYin
                
            For r = startRow To destRow - 1
                wsDest.Cells(r, 1).Value = r - startRow + 1
            Next r
        End If
    Next k
    
    wsDest.Columns(26).ClearContents ' 清理权重缓存列
    
    ' ------------------ 整体格式与排版 ------------------
    
    ' 6.1 字体设定
    With wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(destRow - 1, 6))
        .Font.Name = "宋体"
        .Font.Size = 10.5
    End With
    With wsDest.Range(wsDest.Cells(2, 5), wsDest.Cells(destRow - 1, 5))
        .Font.Name = "Arial"
        .Font.Size = 11
    End With
    
    ' 6.2 居中设定
    wsDest.Columns(1).HorizontalAlignment = xlCenter
    wsDest.Columns(4).HorizontalAlignment = xlCenter
    wsDest.Columns(5).HorizontalAlignment = xlCenter
        '绘制全表实线边框
        With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(destRow - 1, 6)).Borders
        .LineStyle = xlContinuous  ' 设置为实线
        '.weight = xlThin          ' 设置为细线条
        '.ColorIndex = xlAutomatic  ' 设置为自动颜色（通常为黑色/深灰）
    End With
    
    ' 6.5 还原原表列宽
    wsDest.Columns(1).ColumnWidth = wsSource.Columns(cNo).ColumnWidth
    wsDest.Columns(2).ColumnWidth = wsSource.Columns(cName).ColumnWidth
    wsDest.Columns(3).ColumnWidth = wsSource.Columns(cSpec).ColumnWidth
    wsDest.Columns(4).ColumnWidth = wsSource.Columns(cUnit).ColumnWidth
    wsDest.Columns(5).ColumnWidth = wsSource.Columns(cQty).ColumnWidth
    wsDest.Columns(6).ColumnWidth = wsSource.Columns(cRem).ColumnWidth
    
    ' 6.4 渲染引擎探针探测行高折行
    Dim targetCells As Range
    For r = 1 To destRow - 1
        Set targetCells = wsDest.Range(wsDest.Cells(r, 1), wsDest.Cells(r, 6))
        
        ' 开启换行探针并自动调整
        targetCells.WrapText = True
        wsDest.rows(r).AutoFit
        
        ' 探测高度阈值并死锁
        If wsDest.rows(r).RowHeight > 20 Then
            wsDest.rows(r).RowHeight = 30
            targetCells.WrapText = True
        Else
            wsDest.rows(r).RowHeight = 20
            targetCells.WrapText = False
        End If
    Next r
End Sub





