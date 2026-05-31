Attribute VB_Name = "modTrackerCallbacks"
Option Explicit

'------------------------------------------------------------
' Ribbon 回调与按钮控制
'------------------------------------------------------------
   Public g_SummaryMode As Integer ' 记录汇总模式：0 代表 Codex (默认), 1 代表 Gemini
' Sub: OnRibbonLoad - Excel 加载 Ribbon 时触发，获取 UI 句柄并初始化状态
Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
    bIsTracking = False
    g_PageOrientation = 1 'Word导出默认横向
    g_SummaryMode = 0 ' 汇总默认勾选 Codex
End Sub

' Sub: GetEnabledState - 决定 Ribbon 按钮的可用性（灰色或彩色）
' 逻辑：开启时禁用"启动"按钮，启动停止/删除标记"按钮；反之亦然
Sub GetEnabledState(control As IRibbonControl, ByRef returnedVal)
    Select Case control.id
        Case "BtnStart"
            returnedVal = Not bIsTracking
        Case "BtnStop", "BtnMarkDelete"
            returnedVal = bIsTracking
        Case Else
            returnedVal = True
    End Select
End Sub

' Sub: StartTracker_Click - Ribbon 启动按钮的点击回调
Sub StartTracker_Click(control As IRibbonControl)
    StartTrackerLogic
End Sub

' Sub: StopTracker_Click - Ribbon 停止按钮的点击回调
Sub StopTracker_Click(control As IRibbonControl)
    StopTrackerLogic
End Sub

' Sub: MarkDelete_Click - Ribbon "标记删除"按钮的回调逻辑
' 判断：若未启动追踪或未选择有效区域，则弹出警告
Sub MarkDelete_Click(control As IRibbonControl)
    Dim sel As Range
    On Error Resume Next
    Set sel = Application.Selection
    On Error GoTo 0

    If sel Is Nothing Or TypeName(sel) <> "Range" Then
        MsgBox "请选择单个或多个单元格范围", vbExclamation
        Exit Sub
    End If

    ApplyDeleteMark sel
End Sub

' Sub: ManualMark_Click - 手动标记选中内容为红字黄底并写入修改日期
Sub ManualMark_Click(control As IRibbonControl)
    Dim sel As Range
    On Error Resume Next
    Set sel = Application.Selection
    On Error GoTo 0

    If sel Is Nothing Or TypeName(sel) <> "Range" Then
        MsgBox "请选择单个或多个单元格范围", vbExclamation
        Exit Sub
    End If

    ApplyManualMark sel
End Sub

' Sub: RefreshRibbon - 强制刷新 Ribbon 全局状态
Sub RefreshRibbon()
    If Not myRibbon Is Nothing Then
        myRibbon.Invalidate ' 刷新整个 Ribbon
    Else
        MsgBox "Ribbon 连接已丢失，请重新打开工作簿恢复功能。", vbExclamation, "状态提醒"
    End If
End Sub

' 回调 1：清除所有格式
Sub ClearFormatsOnly_Click(control As IRibbonControl)
    ClearFormatsOnly
End Sub

' 回调 2：清除所有标记
Sub ClearMarksAndDates_Click(control As IRibbonControl)
    ClearMarksAndDates
End Sub

' 回调 3：手动清除标记
Sub ClearManualSelection_Click(control As IRibbonControl)
    ClearManualSelection
End Sub

Public Sub HandleSheetActivate(ByVal Sh As Object)
    ProcessSheetActivate Sh
End Sub

Public Sub HandleSheetChange(ByVal Sh As Object, ByVal Target As Range)
    ProcessSheetChange Sh, Target
End Sub


' ==========================================
' WORD导出
' ==========================================
Sub GetPressed(control As IRibbonControl, ByRef returnedVal)
    If control.id = "chkPortrait" Then returnedVal = (g_PageOrientation = 0)
    If control.id = "chkLandscape" Then returnedVal = (g_PageOrientation = 1)
End Sub

Sub OnAction_CheckBox(control As IRibbonControl, pressed As Boolean)
    ' 【新增拦截】如果 Ribbon 对象丢失，拦截点击并提示，防止出现双选假象
    If myRibbon Is Nothing Then
        MsgBox "后台 Ribbon 状态已重置。" & vbCrLf & "请保存并重新打开工作簿。", vbCritical, "无法执行"
        Exit Sub
    End If

    If Not pressed Then
        myRibbon.InvalidateControl "chkPortrait"
        myRibbon.InvalidateControl "chkLandscape"
        Exit Sub
    End If
    
    If control.id = "chkPortrait" Then g_PageOrientation = 0
    If control.id = "chkLandscape" Then g_PageOrientation = 1
    
    myRibbon.InvalidateControl "chkPortrait"
    myRibbon.InvalidateControl "chkLandscape"
End Sub

Sub OnBtnManual(control As IRibbonControl)
    Call StartExportProcess(EXPORT_MODE_MANUAL)
End Sub

Sub OnBtnAuto(control As IRibbonControl)
    Call StartExportProcess(EXPORT_MODE_AUTO)
End Sub

Sub OnBtnTelecom(control As IRibbonControl)
    Call StartExportProcess(EXPORT_MODE_TELECOM)
End Sub

Sub OnBtnAll(control As IRibbonControl)
    Call StartExportProcess(EXPORT_MODE_ALL)
End Sub
' ==========================================
' 仪表设备汇总
' ==========================================
Sub GetPressed_Summary(control As IRibbonControl, ByRef returnedVal)
    If control.id = "chkCodex" Then returnedVal = (g_SummaryMode = 0)
    If control.id = "chkGemini" Then returnedVal = (g_SummaryMode = 1)
End Sub

Sub OnAction_CheckBox_Summary(control As IRibbonControl, pressed As Boolean)
       If myRibbon Is Nothing Then
        MsgBox "后台 Ribbon 状态已重置。" & vbCrLf & "请保存并重新打开工作簿。", vbCritical, "无法执行"
        Exit Sub
    End If

    If Not pressed Then
        myRibbon.InvalidateControl "chkCodex"
        myRibbon.InvalidateControl "chkGemini"
        Exit Sub
    End If
    
    If control.id = "chkCodex" Then g_SummaryMode = 0
    If control.id = "chkGemini" Then g_SummaryMode = 1
    
    myRibbon.InvalidateControl "chkCodex"
    myRibbon.InvalidateControl "chkGemini"
End Sub

Sub OnBtnExecuteSummary(control As IRibbonControl)
    If g_SummaryMode = 0 Then
        GenerateSummary_codex
    Else
        GenerateSummary_gemini
    End If
End Sub


