Attribute VB_Name = "mod_GenerateSummary_Codex"
Option Explicit

'============================================================
' 工程量汇总模块
' 入口过程：GenerateSummary_codex
'============================================================
' ' 说明：此模块负责将工程量清单按设备名称/技术特征/单位进行汇总输出。
' ' 核心要点：
' '  - 支持自动识别表头（含合并单元格与“总计”子列）
' '  - 将表分为三部分：现场仪表、计算机控制系统、第三部分（分类与条目）
' '  - 使用字典合并同名条目，支持技术特征区分与备注清理
' '  - 排序策略：优先预设关键字，再回退到三字连续中文子串聚类（V0.1 逻辑）
' ' 使用方式：在源表激活时运行 GenerateSummary_codex。

Private Const COL_SEQ As Long = 1
Private Const COL_NAME As Long = 2
Private Const COL_TECH As Long = 3
Private Const COL_UNIT As Long = 4
Private Const COL_QTY As Long = 5
Private Const COL_REMARK As Long = 6
' TripleGroupMap：用于将连续三字中文子串映射到组权重的字典。
' ' 该映射来自 V0.1 中的三字聚类逻辑，用于在未命中预设 SortKeys 时进行分组排序。
Private TripleGroupMap As Object

Public Sub GenerateSummary_codex()
    Dim cfg As Object
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim cols As Object
    Dim blocks As Object
    Dim firstDict As Object
    Dim secondRows As Collection
    Dim thirdRows As Collection
    Dim oldCalc As XlCalculation

    On Error GoTo EH

    ' 主流程概览：
    ' 1) 构建配置（关键字、表头、格式等）
    ' 2) 识别源表与表头列，探测分区行
    ' 3) 暂停界面刷新与自动计算提高性能
    ' 4) 收集并合并数据（firstDict/secondRows/thirdRows）
    ' 5) 生成唯一的输出表，写入并格式化
    ' 6) 恢复应用程序设置并提示完成

    Set cfg = BuildConfig()
    Set wsSrc = ResolveSourceWorksheet(ActiveSheet, cfg)
    Set cols = DetectSourceColumns(wsSrc, cfg)

    If cols("Name") = 0 Or cols("Tech") = 0 Or cols("Unit") = 0 Or cols("Qty") = 0 Then
        Err.Raise vbObjectError + 100, , "未能识别项目名称、项目技术特征、计量单位或工程量列。"
    End If
    Set blocks = DetectBlockRows(wsSrc, cols, cfg)

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set firstDict = CreateObject("Scripting.Dictionary")
    Set secondRows = New Collection
    Set thirdRows = New Collection

    ' 初始化三字连续子串映射（用于可选聚类排序）
    Set TripleGroupMap = CreateObject("Scripting.Dictionary")

    CollectSourceData wsSrc, cols, blocks, cfg, firstDict, secondRows, thirdRows

    Set wsOut = CreateUniqueSheet(wsSrc, CStr(cfg("OutputSheetBaseName")))
    SaveSourceSheetName wsOut, wsSrc
    WriteSummary wsOut, firstDict, secondRows, thirdRows, cfg
    FormatSummary wsOut, wsSrc, cols, cfg

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = oldCalc

    MsgBox "工程量汇总完成，已生成工作表：" & wsOut.Name, vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    If oldCalc <> 0 Then Application.Calculation = oldCalc
    MsgBox "工程量汇总失败：" & Err.Description, vbCritical
End Sub

'========================
' 参数设置区
'========================

Private Function BuildConfig() As Object
    ' BuildConfig: 返回一个包含可配置参数的字典（集中维护脚本行为与表头关键字）
    ' 关键配置说明：
    '  - Headers: 输出表头文本数组
    '  - NameHeaders/TechHeaders/UnitHeaders/QtyHeaders: 源表中可能出现的表头名（用于自动识别列）
    '  - QtyTotalHeaders: 若工程量为合并表头，则在下一行查找“总计”或“合计”子列
    '  - StartKeyword/StopKeyword: 用于识别采集起止的关键行（过滤范围）
    '  - Computer...Keyword: 识别计算机控制系统区块并保留软硬件与应用软件两行
    '  - TechDistinctKeys: 若技术特征包含其中词，则视为需单独保留/区分的特征
    '  - SortKeys/ClearRemarkKeys: 排序优先关键词与需要清空备注的关键词
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")

    cfg("OutputSheetBaseName") = "工程量汇总"
    cfg("HeaderSearchRows") = 8
    cfg("HeaderSearchCols") = 40
    cfg("DefaultRemarkColumn") = 15

    cfg("Headers") = Array("序号", "项目名称", "项目技术特征", "计量单位", "工程量", "备注")

    cfg("NameHeaders") = Array("项目名称", "设备")
    cfg("TechHeaders") = Array("项目技术特征", "技术特征")
    cfg("UnitHeaders") = Array("计量单位", "单位")
    cfg("QtyHeaders") = Array("工程量", "数量")
    cfg("QtyTotalHeaders") = Array("总计", "合计")
    cfg("RemarkHeaders") = Array("备注", "备 注")

    cfg("StartKeyword") = "设备及安装"
    cfg("StopKeyword") = "材料及安装"
    cfg("ComputerSystemKeyword") = "计算机控制系统"
    cfg("ComputerHardwareKeyword") = "计算机控制系统软硬件"
    cfg("ComputerSoftwareKeyword") = "计算机控制系统应用软件"

    cfg("MajorFirstName") = "现场仪表"
    cfg("MajorSecondName") = "计算机控制系统"

    cfg("IgnoreNameKeys") = Array("*", "成套")
    cfg("IgnoreCategoryKeys") = Array("小计", "合计", "运杂费")

    cfg("TechDistinctKeys") = Array("防爆", "隔爆", "本安")
    cfg("SortKeys") = Array("温度计", "热电阻", "热电偶", "压力表", "变送器", "流量计", "物位", "料位", "液位", "开关阀", "调节阀", "快切阀", "泄露", "探测器", "分析仪")
    cfg("ClearRemarkKeys") = Array("流量计", "阀", "调节阀", "切断阀", "球阀", "蝶阀", "闸阀", "截止阀")

    cfg("DefaultFontName") = "宋体"
    cfg("DefaultFontSize") = 10.5
    cfg("QtyFontName") = "Arial"
    cfg("QtyFontSize") = 11
    cfg("NormalRowHeight") = 20
    cfg("WrapRowHeight") = 30

    Set BuildConfig = cfg
End Function

'========================
' 数据采集
'========================

Private Sub CollectSourceData(ByVal ws As Worksheet, ByVal cols As Object, ByVal blocks As Object, ByVal cfg As Object, _
                              ByVal firstDict As Object, ByVal secondRows As Collection, ByVal thirdRows As Collection)
    ' CollectSourceData: 遍历源表数据行并按三部分规则收集条目。
    ' 输入：ws（源表）、cols（表头列映射）、blocks（分区行号）、cfg（配置字典）
    ' 输出：firstDict（第一部分合并字典）、secondRows（计算机系统保留行）、thirdRows（第三部分条目）
    Dim r As Long
    Dim lastRow As Long
    Dim nameText As String
    Dim techText As String
    Dim unitText As String
    Dim remarkText As String
    Dim qty As Double
    Dim rowKind As String

    ' 获取数据区域的最后一行（避免逐行扫描整个表）
    lastRow = GetLastUsedRow(ws)

    For r = blocks("StartRow") + 1 To lastRow
        nameText = CleanText(ws.Cells(r, cols("Name")).Text)
        If ContainsAny(nameText, Array(CStr(cfg("StopKeyword")))) Then Exit For

        ' --------- 第1部分：现场仪表（在计算机控制系统头之前） ---------
        ' 常规条目：读取工程量/单位/技术特征并按规则合并
        If r < blocks("ComputerHeaderRow") Then
            qty = ReadRowQuantity(ws, r, cols)
            techText = CleanPipeSize(CleanText(ws.Cells(r, cols("Tech")).Text))
            unitText = CleanText(ws.Cells(r, cols("Unit")).Text)
            If IsValidItem(nameText, qty, unitText, cfg) Then
                remarkText = CleanText(ws.Cells(r, cols("Remark")).Text)
                If ContainsAny(nameText, cfg("ClearRemarkKeys")) Then remarkText = vbNullString
                AddOrMergeItem firstDict, nameText, techText, unitText, qty, remarkText, cfg
            End If

        ' --------- 第2部分：计算机控制系统（只保留软硬件与应用软件两行） ---------
        ElseIf r >= blocks("ComputerHardwareRow") And r <= blocks("ComputerSoftwareRow") Then
            ' 这一段内的明细一般不输出，只保留软硬件与应用软件两条作为整体项
            If r = blocks("ComputerHardwareRow") Or r = blocks("ComputerSoftwareRow") Then
                AddComputerSystemRow ws, r, cols, secondRows
            End If

        ' --------- 第3部分：第三大类 ---------
        ' 第三部分对行进行分类：若为大类标题（CATEGORY）则记录为分类；若为条目（ITEM）则收集为条目待合并
        ElseIf r > blocks("ComputerSoftwareRow") Then
            rowKind = DetectThirdBlockRowKind(ws, r, cols, cfg)
            If rowKind = "CATEGORY" Then
                If Len(nameText) > 0 Then thirdRows.Add Array("CATEGORY", nameText)
            ElseIf rowKind = "ITEM" Then
                qty = ReadRowQuantity(ws, r, cols)
                techText = CleanPipeSize(CleanText(ws.Cells(r, cols("Tech")).Text))
                unitText = CleanText(ws.Cells(r, cols("Unit")).Text)
                If IsValidItem(nameText, qty, unitText, cfg) Then
                    remarkText = CleanText(ws.Cells(r, cols("Remark")).Text)
                    If ContainsAny(nameText, cfg("ClearRemarkKeys")) Then remarkText = vbNullString
                    thirdRows.Add Array("ITEM", nameText, techText, unitText, qty, remarkText)
                End If
            End If
        End If
    Next r
End Sub

Private Sub AddComputerSystemRow(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Object, ByVal rows As Collection)
    ' AddComputerSystemRow：提取计算机控制系统的固定两行（软硬件、应用软件）
    ' 说明：部分表格中“工程量”列为合并单元格，实际数量可能写在主列（QtyMain），
    '      因此读取工程量时需要回退到主列作为兜底。
    Dim nameText As String
    Dim techText As String
    Dim unitText As String
    Dim remarkText As String
    Dim qty As Double

    nameText = CleanText(ws.Cells(r, cols("Name")).Text)
    If Len(nameText) = 0 Then Exit Sub

    techText = CleanPipeSize(CleanText(ws.Cells(r, cols("Tech")).Text))
    unitText = CleanText(ws.Cells(r, cols("Unit")).Text)
    remarkText = CleanText(ws.Cells(r, cols("Remark")).Text)

    '这两行必须保留；总计列为空时，回退读取工程量主列。
    qty = ReadRowQuantity(ws, r, cols)
    If qty = 0 And cols.Exists("QtyMain") Then qty = ReadQuantity(ws.Cells(r, cols("QtyMain")))

    rows.Add Array(nameText, techText, unitText, qty, remarkText)
End Sub
Private Function DetectThirdBlockRowKind(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Object, ByVal cfg As Object) As String
    ' DetectThirdBlockRowKind：判断当前行在第三部分中是大类标题（CATEGORY）还是具体条目（ITEM）
    ' 判定规则（启发式）：
    '  - 若序号列为中文大类（如“一、二...”），视为 CATEGORY
    '  - 若工程量非零且计量单位存在，且名称不为忽略类（小计/合计/运杂费等），视为 ITEM
    '  - 若工程量与单位均为空（且非忽略类），也可能是 CATEGORY（目录标题）
    Dim seqText As String
    Dim nameText As String
    Dim unitText As String
    Dim qty As Double

    seqText = CleanText(ws.Cells(r, cols("Seq")).Text)
    nameText = CleanText(ws.Cells(r, cols("Name")).Text)
    unitText = CleanText(ws.Cells(r, cols("Unit")).Text)
    qty = ReadRowQuantity(ws, r, cols)

    If Len(nameText) = 0 Then Exit Function
    If IsChineseMajorSeq(seqText) Then
        DetectThirdBlockRowKind = "CATEGORY"
    ElseIf qty <> 0 And Len(unitText) > 0 And Not ContainsAny(nameText, cfg("IgnoreCategoryKeys")) Then
        DetectThirdBlockRowKind = "ITEM"
    ElseIf qty = 0 And Len(unitText) = 0 And Not ContainsAny(nameText, cfg("IgnoreCategoryKeys")) Then
        DetectThirdBlockRowKind = "CATEGORY"
    End If
End Function

Private Sub AddOrMergeItem(ByVal dict As Object, ByVal itemName As String, ByVal tech As String, ByVal unitName As String, _
                           ByVal qty As Double, ByVal remark As String, ByVal cfg As Object)
    ' AddOrMergeItem：将单行数据加入字典或合并到已有项。
    ' 合并策略：
    '  - 若技术特征包含 cfg("TechDistinctKeys") 中的词，则将 itemName 与 tech 一并作为 key，保持特征独立
    '  - 否则以 itemName 为 key，将数量累加
    '  - 遇到技术特征或备注不一致时，使用标记（it(5)=True）或清空对应字段以避免错误信息传播
    Dim key As String
    Dim it As Variant
    Dim keepTechSeparate As Boolean

    keepTechSeparate = ContainsAny(tech, cfg("TechDistinctKeys"))
    If keepTechSeparate Then
        key = itemName & ChrW(30) & tech
    Else
        key = itemName
    End If

    If dict.Exists(key) Then
        it = dict(key)
        it(3) = CDbl(it(3)) + qty
        If Not keepTechSeparate Then
            If NormalizeCompareText(CStr(it(1))) <> NormalizeCompareText(tech) Then
                it(5) = True
                it(1) = vbNullString
            End If
        End If
        If Len(CStr(it(2))) = 0 Then it(2) = unitName
        If NormalizeCompareText(CStr(it(4))) <> NormalizeCompareText(remark) Then it(4) = vbNullString
        dict(key) = it
    Else
        dict.Add key, Array(itemName, tech, unitName, qty, remark, False)
    End If
End Sub

Private Function IsValidItem(ByVal itemName As String, ByVal qty As Double, ByVal unitText As String, ByVal cfg As Object) As Boolean
    ' IsValidItem：校验是否为可被汇总的有效条目
    ' 条件：名称非空、工程量非零、单位存在，且名称不包含应忽略的关键词（如成套、合计、小计等）
    If Len(itemName) = 0 Then Exit Function
    If qty = 0 Then Exit Function
    If Len(unitText) = 0 Then Exit Function
    If ContainsAny(itemName, cfg("IgnoreNameKeys")) Then Exit Function
    If ContainsAny(itemName, cfg("IgnoreCategoryKeys")) Then Exit Function
    IsValidItem = True
End Function

'========================
' 输出
'========================

Private Sub WriteSummary(ByVal wsOut As Worksheet, ByVal firstDict As Object, ByVal secondRows As Collection, _
                         ByVal thirdRows As Collection, ByVal cfg As Object)
    ' WriteSummary：负责把已收集的三部分数据写入输出表。
    ' 流程：
    '  1) 清空输出表并写入表头
    '  2) 写入第一部分（现场仪表）——调用 WriteDictItems
    '  3) 写入第二部分（计算机控制系统）——调用 WriteCollectionItems
    '  4) 写入第三部分（按分类分组）——调用 WriteThirdBlocks
    Dim outRow As Long

    wsOut.Cells.Clear
    wsOut.Range("A1:F1").Value = cfg("Headers")
    outRow = 2

    outRow = WriteMajorRow(wsOut, outRow, "一", CStr(cfg("MajorFirstName")))
    outRow = WriteDictItems(wsOut, firstDict, outRow, cfg)

    outRow = WriteMajorRow(wsOut, outRow, "二", CStr(cfg("MajorSecondName")))
    outRow = WriteCollectionItems(wsOut, secondRows, outRow)

    outRow = WriteThirdBlocks(wsOut, thirdRows, outRow, cfg)
End Sub

Private Function WriteThirdBlocks(ByVal wsOut As Worksheet, ByVal thirdRows As Collection, ByVal outRow As Long, ByVal cfg As Object) As Long
    Dim i As Long
    Dim item As Variant
    Dim dict As Object
    Dim majorNo As Long
    Dim hasCategory As Boolean

    Set dict = CreateObject("Scripting.Dictionary")
    majorNo = 3

    For i = 1 To thirdRows.Count
        item = thirdRows(i)
        If item(0) = "CATEGORY" Then
            If dict.Count > 0 Or hasCategory Then
                outRow = WriteDictItems(wsOut, dict, outRow, cfg)
                Set dict = CreateObject("Scripting.Dictionary")
            End If
            outRow = WriteMajorRow(wsOut, outRow, ChineseNumber(majorNo), CStr(item(1)))
            majorNo = majorNo + 1
            hasCategory = True
        ElseIf item(0) = "ITEM" Then
            If Not hasCategory Then
                outRow = WriteMajorRow(wsOut, outRow, ChineseNumber(majorNo), "其他")
                majorNo = majorNo + 1
                hasCategory = True
            End If
            AddOrMergeItem dict, CStr(item(1)), CStr(item(2)), CStr(item(3)), CDbl(item(4)), CStr(item(5)), cfg
        End If
    Next i

    If dict.Count > 0 Then outRow = WriteDictItems(wsOut, dict, outRow, cfg)
    WriteThirdBlocks = outRow
End Function

Private Function WriteMajorRow(ByVal ws As Worksheet, ByVal outRow As Long, ByVal seqText As String, ByVal titleText As String) As Long
    ws.Cells(outRow, COL_SEQ).Value = seqText
    ws.Cells(outRow, COL_NAME).Value = titleText
    WriteMajorRow = outRow + 1
End Function

Private Function WriteCollectionItems(ByVal ws As Worksheet, ByVal rows As Collection, ByVal outRow As Long) As Long
    Dim i As Long
    Dim item As Variant

    For i = 1 To rows.Count
        item = rows(i)
        ws.Cells(outRow, COL_SEQ).Value = i
        ws.Cells(outRow, COL_NAME).Value = item(0)
        ws.Cells(outRow, COL_TECH).Value = item(1)
        ws.Cells(outRow, COL_UNIT).Value = item(2)
        ws.Cells(outRow, COL_QTY).Value = item(3)
        ws.Cells(outRow, COL_REMARK).Value = item(4)
        outRow = outRow + 1
    Next i

    WriteCollectionItems = outRow
End Function

Private Function WriteDictItems(ByVal ws As Worksheet, ByVal dict As Object, ByVal outRow As Long, ByVal cfg As Object) As Long
    ' WriteDictItems：将字典中的合并项写入输出表并按排序规则排序。
    ' 实现要点：
    '  - 在临时列 G/H/I 写入排序辅助值（组权重、常用名分组、原名），用于三键排序
    '  - 调用 SortOutputBlock 进行排序后刷新序号，并清除临时列内容
    Dim startRow As Long
    Dim i As Long
    Dim k As Variant
    Dim it As Variant

    If dict.Count = 0 Then
        WriteDictItems = outRow
        Exit Function
    End If

    startRow = outRow
    i = 1

    For Each k In dict.keys
        it = dict(k)
        ws.Cells(outRow, COL_SEQ).Value = i
        ws.Cells(outRow, COL_NAME).Value = it(0)
        ws.Cells(outRow, COL_TECH).Value = it(1)
        ws.Cells(outRow, COL_UNIT).Value = it(2)
        ws.Cells(outRow, COL_QTY).Value = it(3)
        ws.Cells(outRow, COL_REMARK).Value = it(4)
        ws.Cells(outRow, 7).Value = SortGroupOf(CStr(it(0)), cfg("SortKeys"))
        ws.Cells(outRow, 8).Value = CommonNameGroup(CStr(it(0)))
        ws.Cells(outRow, 9).Value = it(0)
        outRow = outRow + 1
        i = i + 1
    Next k

    SortOutputBlock ws, startRow, outRow - 1

    For i = startRow To outRow - 1
        ws.Cells(i, COL_SEQ).Value = i - startRow + 1
    Next i
    ws.Range(ws.Cells(startRow, 7), ws.Cells(outRow - 1, 9)).ClearContents

    WriteDictItems = outRow
End Function

Private Sub SortOutputBlock(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    ' SortOutputBlock：对输出块按三键排序
    ' 排序键（优先级由高到低）：
    '  1) 列 G - 组权重（SortGroupOf）
    '  2) 列 H - 常用名分组（CommonNameGroup）
    '  3) 列 I - 原名（保证排序稳定性）
    If lastRow < firstRow Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(firstRow, 7), ws.Cells(lastRow, 7)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(firstRow, 8), ws.Cells(lastRow, 8)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(firstRow, 9), ws.Cells(lastRow, 9)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, 9))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin ' 使用拼音排序以兼容中文
        .Apply
    End With
End Sub

'========================
' 表结构识别
'========================

Private Function DetectSourceColumns(ByVal ws As Worksheet, ByVal cfg As Object) As Object
    ' DetectSourceColumns：扫描表头附近区域，自动识别关键列位置（序号/名称/技术特征/单位/工程量/备注）
    '  - 支持在一定行/列范围内查找多个可能的表头名称（由 cfg 提供）
    '  - 若工程量为合并表头，记录主工程量列（QtyMain）并解析“总计”子列（Qty）
    Dim cols As Object
    Dim headerCell As Range
    Dim qtyMain As Range

    Set cols = CreateObject("Scripting.Dictionary")
    cols("Seq") = FindHeaderColumn(ws, Array("序号"), cfg)
    cols("Name") = FindHeaderColumn(ws, cfg("NameHeaders"), cfg)
    cols("Tech") = FindHeaderColumn(ws, cfg("TechHeaders"), cfg)
    cols("Unit") = FindHeaderColumn(ws, cfg("UnitHeaders"), cfg)
    cols("Remark") = FindHeaderColumn(ws, cfg("RemarkHeaders"), cfg)
    If cols("Remark") = 0 Then cols("Remark") = CLng(cfg("DefaultRemarkColumn"))

    Set qtyMain = FindExactHeaderCell(ws, cfg("QtyHeaders"), cfg)
    If qtyMain Is Nothing Then
        cols("Qty") = 0
    Else
        '工程量有子列时，只取“总计”子列；列宽也随这个列。
        cols("QtyMain") = qtyMain.Column
        cols("Qty") = ResolveQuantityColumn(ws, qtyMain, cfg)
    End If

    Set DetectSourceColumns = cols
End Function

Private Function DetectBlockRows(ByVal ws As Worksheet, ByVal cols As Object, ByVal cfg As Object) As Object
    Dim blocks As Object
    Dim lastRow As Long
    Dim r As Long
    Dim nameText As String

    Set blocks = CreateObject("Scripting.Dictionary")
    lastRow = GetLastUsedRow(ws)

    blocks("StartRow") = 1
    blocks("ComputerHeaderRow") = lastRow + 1
    blocks("ComputerHardwareRow") = lastRow + 1
    blocks("ComputerSoftwareRow") = lastRow + 1
    blocks("StopRow") = lastRow + 1

    For r = 1 To lastRow
        nameText = CleanText(ws.Cells(r, cols("Name")).Text)
        If blocks("StartRow") = 1 And ContainsAny(nameText, Array(CStr(cfg("StartKeyword")))) Then blocks("StartRow") = r
        If blocks("ComputerHeaderRow") = lastRow + 1 And nameText = CStr(cfg("ComputerSystemKeyword")) Then blocks("ComputerHeaderRow") = r
        If blocks("ComputerHardwareRow") = lastRow + 1 And ContainsAny(nameText, Array(CStr(cfg("ComputerHardwareKeyword")))) Then blocks("ComputerHardwareRow") = r
        If blocks("ComputerSoftwareRow") = lastRow + 1 And ContainsAny(nameText, Array(CStr(cfg("ComputerSoftwareKeyword")))) Then blocks("ComputerSoftwareRow") = r
        If blocks("StopRow") = lastRow + 1 And ContainsAny(nameText, Array(CStr(cfg("StopKeyword")))) Then blocks("StopRow") = r
    Next r

    If blocks("ComputerHeaderRow") = lastRow + 1 Then
        blocks("ComputerHeaderRow") = blocks("ComputerHardwareRow")
    End If
    If blocks("ComputerHardwareRow") = lastRow + 1 Then
        blocks("ComputerHardwareRow") = blocks("ComputerHeaderRow")
    End If

    Set DetectBlockRows = blocks
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal keys As Variant, ByVal cfg As Object) As Long
    Dim c As Range
    Set c = FindHeaderCell(ws, keys, cfg)
    If Not c Is Nothing Then FindHeaderColumn = c.Column
End Function

Private Function FindHeaderCell(ByVal ws As Worksheet, ByVal keys As Variant, ByVal cfg As Object) As Range
    Dim r As Long
    Dim c As Long
    Dim textValue As String

    For r = 1 To CLng(cfg("HeaderSearchRows"))
        For c = 1 To CLng(cfg("HeaderSearchCols"))
            textValue = CleanText(ws.Cells(r, c).Text)
            If ContainsAny(textValue, keys) Then
                Set FindHeaderCell = ws.Cells(r, c)
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindExactHeaderCell(ByVal ws As Worksheet, ByVal keys As Variant, ByVal cfg As Object) As Range
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim textValue As String

    For r = 1 To CLng(cfg("HeaderSearchRows"))
        For c = 1 To CLng(cfg("HeaderSearchCols"))
            textValue = CleanText(ws.Cells(r, c).Text)
            For i = LBound(keys) To UBound(keys)
                If textValue = CStr(keys(i)) Then
                    Set FindExactHeaderCell = ws.Cells(r, c)
                    Exit Function
                End If
            Next i
        Next c
    Next r
End Function
Private Function ResolveQuantityColumn(ByVal ws As Worksheet, ByVal qtyHeader As Range, ByVal cfg As Object) As Long
    ' ResolveQuantityColumn：当工程量表头为合并单元格时，优先查找下一行的“总计/合计”子列
    ' 返回实际用于读取工程量的列号（若无子列则退回主表头列）
    Dim headerArea As Range
    Dim firstCol As Long
    Dim lastCol As Long
    Dim childRow As Long
    Dim c As Long

    Set headerArea = qtyHeader.mergeArea
    firstCol = headerArea.Column
    lastCol = GetQuantityHeaderLastColumn(ws, qtyHeader)
    childRow = headerArea.Row + headerArea.rows.Count

    ' 优先查找工程量下一行的“总计”子列。
    For c = firstCol To lastCol
        If ContainsAny(CleanText(ws.Cells(childRow, c).Text), cfg("QtyTotalHeaders")) Then
            ResolveQuantityColumn = c
            Exit Function
        End If
    Next c

    ' 没有子列或没有“总计”时，退回工程量表头所在列。
    ResolveQuantityColumn = qtyHeader.Column
End Function

Private Function GetQuantityHeaderLastColumn(ByVal ws As Worksheet, ByVal qtyHeader As Range) As Long
    ' GetQuantityHeaderLastColumn：返回工程量主表头所在区域的最后一列，用于确定子列扫描范围
    Dim headerArea As Range
    Dim c As Long

    Set headerArea = qtyHeader.mergeArea
    If headerArea.Columns.Count > 1 Then
        GetQuantityHeaderLastColumn = headerArea.Column + headerArea.Columns.Count - 1
        Exit Function
    End If

    ' 若未合并，则向右扫描至下一个主表头前一列作为边界（最多扫描 8 列）
    For c = qtyHeader.Column + 1 To qtyHeader.Column + 8
        If Len(CleanText(ws.Cells(qtyHeader.Row, c).Text)) > 0 Then
            GetQuantityHeaderLastColumn = c - 1
            Exit Function
        End If
    Next c
    GetQuantityHeaderLastColumn = qtyHeader.Column
End Function

'========================
' 格式
'========================

Private Sub FormatSummary(ByVal wsOut As Worksheet, ByVal wsSrc As Worksheet, ByVal cols As Object, ByVal cfg As Object)
    Dim lastRow As Long
    Dim rng As Range

    lastRow = GetLastUsedRow(wsOut)
    If lastRow < 1 Then Exit Sub

    Set rng = wsOut.Range("A1:F" & lastRow)
    With rng
        .Font.Name = CStr(cfg("DefaultFontName"))
        .Font.Size = CDbl(cfg("DefaultFontSize"))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    wsOut.Range("E1:E" & lastRow).Font.Name = CStr(cfg("QtyFontName"))
    wsOut.Range("E1:E" & lastRow).Font.Size = CDbl(cfg("QtyFontSize"))

    wsOut.rows(1).Font.Bold = True
    BoldMajorRows wsOut, lastRow

    wsOut.rows(1).HorizontalAlignment = xlCenter
    wsOut.Columns(COL_SEQ).HorizontalAlignment = xlCenter
    wsOut.Columns(COL_UNIT).HorizontalAlignment = xlCenter
    wsOut.Columns(COL_QTY).HorizontalAlignment = xlCenter

    MapColumnWidths wsOut, wsSrc, cols
    ' 调整行高：使用 ApplySmartRowHeight 自动根据单元格内容决定是否换行并采用正常/折行高度
    ' NormalRowHeight: 默认行高（不换行时）
    ' WrapRowHeight: 超出默认高度时的固定折行高度（以保持表格整齐）
    ApplySmartRowHeight wsOut, lastRow, CSng(cfg("NormalRowHeight")), CSng(cfg("WrapRowHeight"))

    ' 保持临时列 G:I 可见（这些列在 WriteDictItems 完成排序后会被清空），便于调试时查看排序权重
    wsOut.Columns("G:I").Hidden = False
End Sub

Private Sub BoldMajorRows(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    For r = 2 To lastRow
        If IsChineseMajorSeq(CleanText(ws.Cells(r, COL_SEQ).Text)) Then
            ws.rows(r).Font.Bold = True
        End If
    Next r
End Sub

Private Sub MapColumnWidths(ByVal wsOut As Worksheet, ByVal wsSrc As Worksheet, ByVal cols As Object)
    wsOut.Columns(COL_SEQ).ColumnWidth = wsSrc.Columns(cols("Seq")).ColumnWidth
    wsOut.Columns(COL_NAME).ColumnWidth = wsSrc.Columns(cols("Name")).ColumnWidth
    wsOut.Columns(COL_TECH).ColumnWidth = wsSrc.Columns(cols("Tech")).ColumnWidth
    wsOut.Columns(COL_UNIT).ColumnWidth = wsSrc.Columns(cols("Unit")).ColumnWidth
    wsOut.Columns(COL_QTY).ColumnWidth = wsSrc.Columns(cols("Qty")).ColumnWidth
    wsOut.Columns(COL_REMARK).ColumnWidth = wsSrc.Columns(cols("Remark")).ColumnWidth
End Sub

Private Sub ApplySmartRowHeight(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal normalHeight As Single, ByVal wrapHeight As Single)
    Dim r As Long
    ' ApplySmartRowHeight：智能设置每行的换行与高度
    ' 算法：先临时启用 WrapText 并 AutoFit 行高
    ' 若自动测得的行高大于 normalHeight，则认为需要折行，固定为 wrapHeight 并保持换行
    ' 否则设置为 normalHeight 并关闭换行以节省空间
    For r = 1 To lastRow
        With ws.rows(r)
            .WrapText = True
            .EntireRow.AutoFit
            If .RowHeight > normalHeight Then
                .RowHeight = wrapHeight
                .WrapText = True
            Else
                .RowHeight = normalHeight
                .WrapText = False
            End If
        End With
    Next r
End Sub

'========================
' 通用工具
'========================

Private Function ReadQuantity(ByVal cell As Range) As Double
    ' ReadQuantity：安全读取单元格数值，忽略错误值并确保返回 Double
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then ReadQuantity = CDbl(v)
End Function

Private Function ReadRowQuantity(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal cols As Object) As Double
    '数据行只读已识别的工程量列；若存在“总计”子列，这里就是总计列。
    ReadRowQuantity = ReadQuantity(ws.Cells(rowIndex, cols("Qty")))
End Function

Private Function CleanText(ByVal textValue As String) As String
    ' CleanText：统一清理文本，替换全角空格与换行符，并做 Trim
    textValue = Replace(textValue, ChrW(12288), " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    CleanText = Trim(textValue)
End Function

Private Function NormalizeCompareText(ByVal textValue As String) As String
    ' NormalizeCompareText：用于比较文本时的标准化（去掉空格并清理）
    NormalizeCompareText = Replace(CleanText(textValue), " ", vbNullString)
End Function

Private Function CleanPipeSize(ByVal textValue As String) As String
    ' CleanPipeSize：使用正则去除管道口径/公称直径等说明（如 DN100、Φ50×2 等），并清理分隔符
    ' 目的是留下设备的技术特征主体，便于聚类和比较
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "((DN|De|Dg)\s*\d+(\.\d+)?|[Φφ]\s*\d+(\.\d+)?(\s*[xX×*]\s*\d+(\.\d+)?)?)(\s*[,，;；、/]\s*)?"
    CleanPipeSize = Trim(re.Replace(textValue, vbNullString))
End Function

Private Function ContainsAny(ByVal textValue As String, ByVal keys As Variant) As Boolean
    ' ContainsAny：判断 textValue 是否包含 keys 中任一子串（不区分大小写）
    Dim i As Long
    If IsArray(keys) Then
        For i = LBound(keys) To UBound(keys)
            If Len(CStr(keys(i))) > 0 Then
                If InStr(1, textValue, CStr(keys(i)), vbTextCompare) > 0 Then
                    ContainsAny = True
                    Exit Function
                End If
            End If
        Next i
    ElseIf Len(CStr(keys)) > 0 Then
        ContainsAny = (InStr(1, textValue, CStr(keys), vbTextCompare) > 0)
    End If
End Function

Private Function ExtractAllChineseTriplets(ByVal s As String) As Variant
    ' ExtractAllChineseTriplets：提取字符串中所有连续的三字中文子串并去重
    ' 返回：若无满足的子串则返回 Empty，否则返回字符串数组
    Dim i As Long, sLen As Long
    Dim ch1 As String, ch2 As String, ch3 As String, triple As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    s = Trim(s)
    sLen = Len(s)
    If sLen < 3 Then
        ExtractAllChineseTriplets = Empty
        Exit Function
    End If
    For i = 1 To sLen - 2
        ch1 = Mid(s, i, 1)
        ch2 = Mid(s, i + 1, 1)
        ch3 = Mid(s, i + 2, 1)
        ' 通过 AscW > 255 判断为中文字符
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
        Dim idx As Long: idx = 0
        Dim k As Variant
        For Each k In d.keys
            arr(idx) = k
            idx = idx + 1
        Next k
        ExtractAllChineseTriplets = arr
    End If
End Function

Private Function SortGroupOf(ByVal itemName As String, ByVal sortKeys As Variant) As Long
    Dim i As Long
    Dim triples As Variant
    Dim t As Variant
    Dim nextWeight As Long

    If IsArray(sortKeys) Then
        For i = LBound(sortKeys) To UBound(sortKeys)
            If InStr(1, itemName, CStr(sortKeys(i)), vbTextCompare) > 0 Then
                SortGroupOf = i + 1
                Exit Function
            End If
        Next i
    End If

    ' 尝试使用连续三字子串进行分组（与 V0.1 的聚类逻辑兼容）
    triples = ExtractAllChineseTriplets(itemName)
    If Not IsEmpty(triples) Then
        For Each t In triples
            If Len(CStr(t)) > 0 Then
                If Not TripleGroupMap Is Nothing Then
                    If TripleGroupMap.Exists(CStr(t)) Then
                        SortGroupOf = TripleGroupMap(CStr(t))
                        Exit Function
                    End If
                End If
            End If
        Next t
        ' 若没有已存在映射，则为该名称分配新的组权重，并将其子串映射到该权重
        If TripleGroupMap Is Nothing Then Set TripleGroupMap = CreateObject("Scripting.Dictionary")
        nextWeight = 100 + TripleGroupMap.Count + 1
        For Each t In triples
            If Not TripleGroupMap.Exists(CStr(t)) Then TripleGroupMap.Add CStr(t), nextWeight
        Next t
        SortGroupOf = nextWeight
        Exit Function
    End If

    SortGroupOf = 999
End Function

Private Function CommonNameGroup(ByVal itemName As String) As String
    ' CommonNameGroup：用于从名称中提取一个用于分组/二次排序的短标识（优先取左侧 3~5 字）
    ' 目的是把语义上相近的名称聚在一起，辅助拼音排序提高可读性
    Dim i As Long
    Dim maxLen As Long
    Dim token As String

    itemName = NormalizeCompareText(itemName)
    If Len(itemName) < 3 Then
        CommonNameGroup = itemName
        Exit Function
    End If

    maxLen = IIf(Len(itemName) >= 5, 5, Len(itemName))
    For i = maxLen To 3 Step -1
        token = Left$(itemName, i)
        If Len(token) >= 3 Then
            CommonNameGroup = token
            Exit Function
        End If
    Next i

    CommonNameGroup = itemName
End Function

Private Function IsChineseMajorSeq(ByVal seqText As String) As Boolean
    ' IsChineseMajorSeq：检查序号文本是否为中文大类序号（用于识别第一列为“大类”行，例如“一、二、三”）
    Dim chineseNums As String
    chineseNums = "一二三四五六七八九十"
    seqText = CleanText(seqText)
    If Len(seqText) = 0 Or Len(seqText) > 3 Then Exit Function
    IsChineseMajorSeq = (InStr(1, chineseNums, Left$(seqText, 1), vbBinaryCompare) > 0)
End Function

Private Function ChineseNumber(ByVal n As Long) As String
    ' ChineseNumber：将整数转换为常用中文数字表示（用于输出大类序号）
    Dim arr As Variant
    arr = Array("", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    If n >= 1 And n <= UBound(arr) Then
        ChineseNumber = arr(n)
    Else
        ChineseNumber = CStr(n)
    End If
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    ' GetLastUsedRow：返回工作表中最后有内容的行号（用于限制遍历范围，避免扫描大量空行）
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        GetLastUsedRow = 1
    Else
        GetLastUsedRow = lastCell.Row
    End If
End Function

Private Function ResolveSourceWorksheet(ByVal activeWs As Worksheet, ByVal cfg As Object) As Worksheet
    Dim sourceName As String

    '若当前在已生成的汇总表上再次运行，自动回到记录的源表取数。
    If Left$(activeWs.Name, Len(CStr(cfg("OutputSheetBaseName")))) = CStr(cfg("OutputSheetBaseName")) Then
        sourceName = GetSourceSheetName(activeWs)
        If Len(sourceName) > 0 Then
            Set ResolveSourceWorksheet = activeWs.Parent.Worksheets(sourceName)
            Exit Function
        End If
        If activeWs.Index > 1 Then
            Set ResolveSourceWorksheet = activeWs.Parent.Worksheets(activeWs.Index - 1)
            Exit Function
        End If
    End If
    ' 默认直接使用激活工作表作为源表（若不是已生成的汇总表）
    Set ResolveSourceWorksheet = activeWs
End Function

Private Function GetSourceSheetName(ByVal ws As Worksheet) As String
    ' GetSourceSheetName：从已生成的汇总表的自定义属性中读取记录的源表名称（用于再次运行时自动回溯）
    On Error Resume Next
    GetSourceSheetName = CStr(ws.CustomProperties("SourceSheetName").Value)
    On Error GoTo 0
End Function

Private Sub SaveSourceSheetName(ByVal wsOut As Worksheet, ByVal wsSrc As Worksheet)
    ' SaveSourceSheetName：在输出工作表上写入自定义属性记录源表名，便于之后从汇总表回溯到源数据
    On Error Resume Next
    wsOut.CustomProperties("SourceSheetName").Delete
    On Error GoTo 0
    wsOut.CustomProperties.Add "SourceSheetName", wsSrc.Name
End Sub
Private Function CreateUniqueSheet(ByVal afterSheet As Worksheet, ByVal baseName As String) As Worksheet
    Dim wb As Workbook
    Dim nameTry As String
    Dim idx As Long

    Set wb = afterSheet.Parent
    nameTry = baseName
    idx = 1

    Do While SheetExists(wb, nameTry)
        idx = idx + 1
        nameTry = baseName & idx
    Loop

    Set CreateUniqueSheet = wb.Worksheets.Add(After:=afterSheet)
    CreateUniqueSheet.Name = nameTry
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function


