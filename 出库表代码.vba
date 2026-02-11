Option Explicit

' ============================================
' 出库表工作表代码模块
' 功能：监控出库表的变化，自动联动更新入库表的实时库存
' 创建日期：2026-02-09
' ============================================

' 上次的行数（用于检测删除/新增操作）
Private previousRowCount As Long

' 防止重复触发刷新的标志位
Private isRefreshing As Boolean

' 🆕 用于SelectionChange检测
Private lastCheckTime As Date
Private lastKnownRowCount As Long

' 工作表激活时初始化
Private Sub Worksheet_Activate()
    previousRowCount = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row
    lastKnownRowCount = previousRowCount
    isRefreshing = False
    lastCheckTime = Now
End Sub

' 监控出库表的变化
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler

    ' 如果正在刷新中，直接退出（防止死循环）
    If isRefreshing Then Exit Sub

    ' 如果是表头行被修改，直接退出
    If Target.Row <= 1 Then Exit Sub

    Dim currentRowCount As Long

    ' 获取当前行数
    currentRowCount = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row

    ' 如果是首次加载，初始化并退出
    If previousRowCount = 0 Then
        previousRowCount = currentRowCount
        Exit Sub
    End If

    ' 检测是否有变化（行数变化 = 删除或新增）
    Dim hasChange As Boolean
    hasChange = (currentRowCount <> previousRowCount)

    ' 更新行数
    previousRowCount = currentRowCount

    ' 如果没有变化，退出
    If Not hasChange Then Exit Sub

    ' 延迟500ms后刷新入库表库存（避免频繁触发）
    Application.OnTime Now + TimeValue("00:00:00.5"), "RefreshInventoryDelayed"

    ' 🆕 延迟1秒后刷新车间结存
    Application.OnTime Now + TimeValue("00:00:01"), "RefreshAllWorkshopStockQuietly"

    Exit Sub

ErrorHandler:
    Debug.Print "出库表监控错误: " & Err.Description
End Sub

' 🆕 监听选择变化（用于检测删除行）
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    
    ' 每5秒检查一次（避免频繁触发）
    If DateDiff("s", lastCheckTime, Now) < 5 Then Exit Sub
    
    Dim currentRowCount As Long
    currentRowCount = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row
    
    ' 初始化
    If lastKnownRowCount = 0 Then
        lastKnownRowCount = currentRowCount
        lastCheckTime = Now
        Exit Sub
    End If
    
    ' 检测行数减少（删除操作）
    If currentRowCount < lastKnownRowCount Then
        lastKnownRowCount = currentRowCount
        lastCheckTime = Now
        
        ' 延迟刷新入库表库存
        Application.OnTime Now + TimeValue("00:00:01"), "RefreshInventoryDelayed"
        
        ' 🆕 延迟刷新车间结存
        Application.OnTime Now + TimeValue("00:00:02"), "RefreshAllWorkshopStockQuietly"
        
        ' 显示提示
        Application.StatusBar = "检测到出库记录删除，正在更新库存和车间结存..."
        Application.OnTime Now + TimeValue("00:00:05"), "ClearStatusBarQuiet"
    ElseIf currentRowCount > lastKnownRowCount Then
        ' 行数增加（新增记录）
        lastKnownRowCount = currentRowCount
        lastCheckTime = Now
    End If
End Sub

' 延迟刷新函数（在标准模块中）
' 这个函数会被OnTime调用，放在标准模块中
