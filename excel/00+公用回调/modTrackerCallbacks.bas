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

' Sub: RefreshRibbon - 强制刷新 Ribbon 按钮状态，使其重新调用 GetEnabledState
Sub RefreshRibbon()
    RefreshRibbon
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
    If Not pressed Then
        If Not myRibbon Is Nothing Then
            myRibbon.InvalidateControl "chkPortrait"
            myRibbon.InvalidateControl "chkLandscape"
        End If
        Exit Sub
    End If
    
    If control.id = "chkPortrait" Then g_PageOrientation = 0
    If control.id = "chkLandscape" Then g_PageOrientation = 1
    
    If Not myRibbon Is Nothing Then
        myRibbon.InvalidateControl "chkPortrait"
        myRibbon.InvalidateControl "chkLandscape"
    End If
End Sub

Sub OnBtnSingle(control As IRibbonControl)
    Call StartExportProcess(False) ' 调用主模块的导出单页方法
End Sub

Sub OnBtnAll(control As IRibbonControl)
    Call StartExportProcess(True)  ' 调用主模块的导出所有方法
End Sub
' ==========================================
' 仪表设备汇总
' ==========================================
Sub GetPressed_Summary(control As IRibbonControl, ByRef returnedVal)
    If control.id = "chkCodex" Then returnedVal = (g_SummaryMode = 0)
    If control.id = "chkGemini" Then returnedVal = (g_SummaryMode = 1)
End Sub

Sub OnAction_CheckBox_Summary(control As IRibbonControl, pressed As Boolean)
    If Not pressed Then
        If Not myRibbon Is Nothing Then
            myRibbon.InvalidateControl "chkCodex"
            myRibbon.InvalidateControl "chkGemini"
        End If
        Exit Sub
    End If
    
    If control.id = "chkCodex" Then g_SummaryMode = 0
    If control.id = "chkGemini" Then g_SummaryMode = 1
    
    If Not myRibbon Is Nothing Then
        myRibbon.InvalidateControl "chkCodex"
        myRibbon.InvalidateControl "chkGemini"
    End If
End Sub

Sub OnBtnExecuteSummary(control As IRibbonControl)
    If g_SummaryMode = 0 Then
        GenerateSummary_codex
    Else
        GenerateSummary_gemini
    End If
End Sub


