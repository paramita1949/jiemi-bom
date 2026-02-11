' ============================================
' 实时库存表工作表代码
' 将此代码复制到"实时库存"工作表模块中
' ============================================

Private Sub CommandButton1_Click()
    On Error GoTo ErrorHandler
    
    ' 显示提示
    Application.StatusBar = "正在刷新实时库存表..."
    
    ' 调用刷新函数
    Call RefreshRealTimeInventoryQuietly
    
    ' 显示完成提示
    Application.StatusBar = "实时库存表已刷新！" & Format(Now(), "hh:mm:ss")
    
    ' 2秒后清除状态栏
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBarQuiet"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "刷新实时库存表时发生错误: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub
