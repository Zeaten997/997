VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

'====================================================================
' clsAppEvents.cls
' Excel VBA 类模块：应用事件监听
' ------------------------------------------------------------------
' 说明：
'   - App_SheetActivate 用于检测工作表切换并停止追踪
'   - App_SheetChange 用于检测单元格变更并执行修改标记
' ------------------------------------------------------------------
' 功能汇总：
'   1. 跨表监控：利用 WithEvents 监听整个 Excel 实例的单元格变动。
'   2. 自动停止：当用户切换工作表时，出于逻辑安全考虑自动停止追踪。
'   3. 修改追踪：检测单元格内容是否与隐藏备份表（HiddenLog_前缀）不一致。
'   4. 视觉标记：对修改过的单元格进行标红、加黄底色，并自动记录修改日期。
'   5. 状态还原：若单元格改回原始值且该行无其他修改，自动清除视觉标记和日期。
'====================================================================

Option Explicit

' 定义隐藏日志表的前缀，用于比对原始数据
Private Const HiddenLogPrefix As String = "HiddenLog_"
' 设置最大处理单元格数量，防止一次性粘贴大量数据时导致程序卡死
Private Const MaxTargetCells As Long = 10000

'--------------------------------------------------------------------
' Sub: App_SheetActivate
' 功能：监听工作表激活事件。
' 逻辑：当用户点击切换到其他工作表时，如果当前正处于追踪模式（bIsTracking），
'       则调用 StopTrackerLogic 停止追踪，以避免在错误的工作表上执行逻辑。
'--------------------------------------------------------------------
Private Sub App_SheetActivate(ByVal Sh As Object)
    If bIsTracking Then HandleSheetActivate Sh
End Sub

'--------------------------------------------------------------------
' Sub: App_SheetChange
' 功能：核心逻辑——监听单元格内容变更。
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If bIsTracking Then HandleSheetChange Sh, Target
End Sub
