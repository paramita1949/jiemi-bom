' ============================================
' ThisWorkbook 模块代码（可选）
' 如果需要启用定时器自动刷新，将此代码复制到 ThisWorkbook 模块
' ============================================

' 工作簿打开时启动定时器
Private Sub Workbook_Open()
    ' 启动自动刷新定时器（每2分钟刷新一次）
    Call StartAutoRefreshTimer
End Sub

' 工作簿关闭前停止定时器
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 停止自动刷新定时器
    Call StopAutoRefreshTimer
End Sub
