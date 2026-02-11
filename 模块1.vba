Option Explicit

' ============================================
' 库存管理模块
' 功能：实时库存的计算和维护
' 创建日期：2026-02-09
' ============================================

Private Const DEBUG_LOG As Boolean = False

' 定时器相关变量
Private nextScheduleTime As Date

' 启动自动刷新定时器（每2分钟检查一次）
Public Sub StartAutoRefreshTimer()
    On Error Resume Next

    ' 先取消之前的定时器
    Call StopAutoRefreshTimer

    ' 设置新的定时器（2分钟后）
    nextScheduleTime = Now + TimeValue("00:02:00")
    Application.OnTime nextScheduleTime, "AutoRefreshInventory"

    Debug.Print "自动刷新定时器已启动，下次刷新时间: " & nextScheduleTime
End Sub

' 停止自动刷新定时器
Public Sub StopAutoRefreshTimer()
    On Error Resume Next

    If nextScheduleTime > 0 Then
        Application.OnTime nextScheduleTime, "AutoRefreshInventory", , False
        nextScheduleTime = 0
        Debug.Print "自动刷新定时器已停止"
    End If
End Sub

' 定时器回调函数（每2分钟自动调用）
Sub AutoRefreshInventory()
    On Error Resume Next

    ' 静默刷新库存
    Call RefreshAllInventoryQuietly

    ' 在状态栏显示提示
    Application.StatusBar = "自动刷新：库存已更新 " & Format(Now(), "hh:mm:ss")

    ' 2秒后清除状态栏
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBarQuiet"

    ' 重新启动定时器（下次2分钟后）
    Call StartAutoRefreshTimer
End Sub

' 获取列索引的辅助函数（根据表头名称）
Function GetColumnIndex(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim col As Long
    Dim lastCol As Long

    ' 查找最后一列
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    ' 遍历表头行，查找匹配的列名
    For col = 1 To lastCol
        If Trim(ws.Cells(headerRow, col).Value) = Trim(headerName) Then
            GetColumnIndex = col
            Exit Function
        End If
    Next col

    ' 如果没找到，返回 0
    GetColumnIndex = 0
End Function

' 调试日志输出函数
Sub DebugLog(category As String, message As String)
    Dim logPath As String
    Dim fso As Object
    Dim logFile As Object

    logPath = "c:\Users\Administrator\Desktop\vba2\调试.md"

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 如果文件不存在，创建新文件
    If Not fso.FileExists(logPath) Then
        Set logFile = fso.CreateTextFile(logPath, True)
    Else
        ' 追加模式打开文件
        Set logFile = fso.OpenTextFile(logPath, 8, True)  ' 8 = ForAppending
    End If

    ' 写入日志
    logFile.WriteLine Now & vbTab & category & vbTab & message
    logFile.Close

    Set logFile = Nothing
    Set fso = Nothing
End Sub

' ============================================
' 主函数：刷新所有物料的实时库存
' ============================================
Sub RefreshAllInventory()
    On Error GoTo ErrorHandler

    Dim wsInbound As Worksheet
    Dim wsOutbound As Worksheet
    Dim inboundLastRow As Long
    Dim outboundLastRow As Long
    Dim i As Long, j As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsInbound = ThisWorkbook.Worksheets("入库")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")

    ' 获取入库表列索引
    Dim colInDate As Long
    Dim colInMaterialCode As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    Dim colInAlreadyOut As Long
    Dim colInStock As Long

    colInDate = GetColumnIndex(wsInbound, 1, "日期")
    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")
    colInAlreadyOut = GetColumnIndex(wsInbound, 1, "已出库")
    colInStock = GetColumnIndex(wsInbound, 1, "实时库存")

    ' 验证必需列
    If colInMaterialCode = 0 Or colInBatch = 0 Or colInQty = 0 Or colInStock = 0 Then
        MsgBox "入库表缺少必需的列（物料编号、批次、入库数量或实时库存）", vbCritical, "错误"
        GoTo CleanUp
    End If

    ' 获取出库表列索引
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long

    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")

    ' 验证出库表必需列
    If colOutMaterialCode = 0 Or colOutBatch = 0 Or colOutQty = 0 Then
        MsgBox "出库表缺少必需的列（物料编号、批次或出库数量）", vbCritical, "错误"
        GoTo CleanUp
    End If

    ' 获取入库表最后一行
    inboundLastRow = wsInbound.Cells(wsInbound.Rows.Count, colInMaterialCode).End(xlUp).Row
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row

    ' 遍历入库表的每一行，计算实时库存
    For i = 2 To inboundLastRow  ' 第1行是表头
        Dim inMaterialCode As String
        Dim inBatch As String
        Dim inQty As Double
        Dim totalOutbound As Double

        inMaterialCode = Trim(wsInbound.Cells(i, colInMaterialCode).Value)
        inBatch = Trim(wsInbound.Cells(i, colInBatch).Value)
        inQty = Val(wsInbound.Cells(i, colInQty).Value)

        ' 跳过空行
        If inMaterialCode = "" Or inBatch = "" Then GoTo NextInboundRow

        ' 计算该批次的累计出库量
        totalOutbound = 0
        For j = 2 To outboundLastRow
            If Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = inMaterialCode And _
               Trim(wsOutbound.Cells(j, colOutBatch).Value) = inBatch Then
                totalOutbound = totalOutbound + Val(wsOutbound.Cells(j, colOutQty).Value)
            End If
        Next j

        ' 计算实时库存 = 入库数量 - 累计出库
        Dim realStock As Double
        realStock = inQty - totalOutbound

        ' 确保库存不为负数
        If realStock < 0 Then realStock = 0

        ' 更新入库表的实时库存
        wsInbound.Cells(i, colInStock).Value = realStock

        ' 更新"已出库"列（如果没有出库记录，自动填0）
        If colInAlreadyOut > 0 Then
            wsInbound.Cells(i, colInAlreadyOut).Value = totalOutbound
        End If

        If DEBUG_LOG Then
            DebugLog "RefreshStock", _
                     "Material=" & inMaterialCode & _
                     ", Batch=" & inBatch & _
                     ", InQty=" & inQty & _
                     ", OutQty=" & totalOutbound & _
                     ", RealStock=" & realStock
        End If

NextInboundRow:
    Next i

    MsgBox "实时库存已刷新完成！" & vbCrLf & "共处理 " & (inboundLastRow - 1) & " 条入库记录", vbInformation, "完成"

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "刷新实时库存时发生错误: " & Err.Description, vbCritical, "错误"
    Resume CleanUp
End Sub

' ============================================
' 单个物料的实时库存计算
' 返回指定批次物料的实时库存
' ============================================
Function GetRealTimeStock(materialCode As String, batchNumber As String) As Double
    On Error GoTo ErrorHandler

    Dim wsInbound As Worksheet
    Dim wsOutbound As Worksheet
    Dim inboundLastRow As Long
    Dim outboundLastRow As Long
    Dim i As Long

    Set wsInbound = ThisWorkbook.Worksheets("入库")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")

    ' 获取列索引
    Dim colInMaterialCode As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long

    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")

    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")

    Dim inboundQty As Double
    Dim totalOutbound As Double

    ' 查找入库记录
    inboundLastRow = wsInbound.Cells(wsInbound.Rows.Count, colInMaterialCode).End(xlUp).Row
    For i = 2 To inboundLastRow
        If Trim(wsInbound.Cells(i, colInMaterialCode).Value) = materialCode And _
           Trim(wsInbound.Cells(i, colInBatch).Value) = batchNumber Then
            inboundQty = Val(wsInbound.Cells(i, colInQty).Value)
            Exit For
        End If
    Next i

    ' 计算累计出库
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    For i = 2 To outboundLastRow
        If Trim(wsOutbound.Cells(i, colOutMaterialCode).Value) = materialCode And _
           Trim(wsOutbound.Cells(i, colOutBatch).Value) = batchNumber Then
            totalOutbound = totalOutbound + Val(wsOutbound.Cells(i, colOutQty).Value)
        End If
    Next i

    ' 返回实时库存
    GetRealTimeStock = inboundQty - totalOutbound
    If GetRealTimeStock < 0 Then GetRealTimeStock = 0

    Exit Function

ErrorHandler:
    GetRealTimeStock = 0
End Function

' ============================================
' 刷新指定物料的实时库存
' ============================================
Sub RefreshMaterialInventory(materialCode As String)
    On Error GoTo ErrorHandler

    Dim wsInbound As Worksheet
    Dim wsOutbound As Worksheet
    Dim inboundLastRow As Long
    Dim outboundLastRow As Long
    Dim i As Long, j As Long

    Set wsInbound = ThisWorkbook.Worksheets("入库")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")

    ' 获取列索引
    Dim colInMaterialCode As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    Dim colInAlreadyOut As Long
    Dim colInStock As Long
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long

    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")
    colInAlreadyOut = GetColumnIndex(wsInbound, 1, "已出库")
    colInStock = GetColumnIndex(wsInbound, 1, "实时库存")

    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")

    inboundLastRow = wsInbound.Cells(wsInbound.Rows.Count, colInMaterialCode).End(xlUp).Row
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row

    ' 遍历入库表中该物料的所有批次
    For i = 2 To inboundLastRow
        If Trim(wsInbound.Cells(i, colInMaterialCode).Value) = materialCode Then
            Dim inBatch As String
            Dim inQty As Double
            Dim totalOutbound As Double

            inBatch = Trim(wsInbound.Cells(i, colInBatch).Value)
            inQty = Val(wsInbound.Cells(i, colInQty).Value)

            ' 计算该批次的累计出库量
            totalOutbound = 0
            For j = 2 To outboundLastRow
                If Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = materialCode And _
                   Trim(wsOutbound.Cells(j, colOutBatch).Value) = inBatch Then
                    totalOutbound = totalOutbound + Val(wsOutbound.Cells(j, colOutQty).Value)
                End If
            Next j

            ' 更新实时库存（如果没有出库记录，实时库存 = 入库数量）
            Dim realStock As Double
            realStock = inQty - totalOutbound
            If realStock < 0 Then realStock = 0

            wsInbound.Cells(i, colInStock).Value = realStock

            ' 更新"已出库"列（如果没有出库记录，自动填0）
            If colInAlreadyOut > 0 Then
                wsInbound.Cells(i, colInAlreadyOut).Value = totalOutbound
            End If
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "刷新物料库存时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' ============================================
' 延迟刷新函数（由出库表的OnTime调用）
' ============================================
Sub RefreshInventoryDelayed()
    On Error Resume Next

    ' 静默刷新库存，不显示任何提示
    Call RefreshAllInventoryQuietly

    ' 在状态栏显示简短提示
    Application.StatusBar = "库存已刷新 " & Format(Now(), "hh:mm:ss")

    ' 2秒后清除状态栏
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBarQuiet"
End Sub

' 静默刷新库存（无弹窗）
Sub RefreshAllInventoryQuietly()
    On Error GoTo ErrorHandler

    Dim wsInbound As Worksheet
    Dim wsOutbound As Worksheet
    Dim inboundLastRow As Long
    Dim outboundLastRow As Long
    Dim i As Long, j As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsInbound = ThisWorkbook.Worksheets("入库")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")

    ' 获取入库表列索引
    Dim colInDate As Long
    Dim colInMaterialCode As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    Dim colInAlreadyOut As Long
    Dim colInStock As Long

    colInDate = GetColumnIndex(wsInbound, 1, "日期")
    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")
    colInAlreadyOut = GetColumnIndex(wsInbound, 1, "已出库")
    colInStock = GetColumnIndex(wsInbound, 1, "实时库存")

    ' 验证必需列
    If colInMaterialCode = 0 Or colInBatch = 0 Or colInQty = 0 Or colInStock = 0 Then
        GoTo CleanUp
    End If

    ' 获取出库表列索引
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long

    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")

    ' 验证出库表必需列
    If colOutMaterialCode = 0 Or colOutBatch = 0 Or colOutQty = 0 Then
        GoTo CleanUp
    End If

    ' 获取入库表最后一行
    inboundLastRow = wsInbound.Cells(wsInbound.Rows.Count, colInMaterialCode).End(xlUp).Row
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row

    ' 遍历入库表的每一行，计算实时库存
    For i = 2 To inboundLastRow  ' 第1行是表头
        Dim inMaterialCode As String
        Dim inBatch As String
        Dim inQty As Double
        Dim totalOutbound As Double

        inMaterialCode = Trim(wsInbound.Cells(i, colInMaterialCode).Value)
        inBatch = Trim(wsInbound.Cells(i, colInBatch).Value)
        inQty = Val(wsInbound.Cells(i, colInQty).Value)

        ' 跳过空行
        If inMaterialCode = "" Or inBatch = "" Then GoTo NextInboundRow

        ' 计算该批次的累计出库量
        totalOutbound = 0
        For j = 2 To outboundLastRow
            If Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = inMaterialCode And _
               Trim(wsOutbound.Cells(j, colOutBatch).Value) = inBatch Then
                totalOutbound = totalOutbound + Val(wsOutbound.Cells(j, colOutQty).Value)
            End If
        Next j

        ' 计算实时库存 = 入库数量 - 累计出库
        Dim realStock As Double
        realStock = inQty - totalOutbound

        ' 确保库存不为负数
        If realStock < 0 Then realStock = 0

        ' 更新入库表的实时库存
        wsInbound.Cells(i, colInStock).Value = realStock

        ' 更新"已出库"列（如果没有出库记录，自动填0）
        If colInAlreadyOut > 0 Then
            wsInbound.Cells(i, colInAlreadyOut).Value = totalOutbound
        End If

NextInboundRow:
    Next i

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Resume CleanUp
End Sub

' 清除状态栏（静默版）
Sub ClearStatusBarQuiet()
    Application.StatusBar = False
End Sub


' ============================================
' 辅助函数：从规格字符串中提取数字
' 例如："30000只/袋" -> 30000
' ============================================
Public Function ExtractSpecQuantity(spec As String) As Double
    On Error Resume Next

    ' 特殊情况处理：如果规格是 "/" 或空，视为 1
    If Trim(spec) = "/" Or Trim(spec) = "" Or Trim(spec) = "无" Then
        ExtractSpecQuantity = 1
        Exit Function
    End If

    Dim regex As Object
    Dim matches As Object

    ' 创建正则表达式对象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(\d+)"  ' 匹配数字
    regex.Global = False

    ' 执行匹配
    If regex.Test(spec) Then
        Set matches = regex.Execute(spec)
        ExtractSpecQuantity = CDbl(matches(0).Value)
    Else
        ' 如果没有匹配到数字，返回1作为默认值
        ExtractSpecQuantity = 1
    End If

    On Error GoTo 0
End Function

' ============================================
' 生成实时库存表（分组折叠显示）
' 表头位置: D4-L4
' 数据起始: D5
' ============================================
Sub GenerateRealTimeInventory()
    On Error GoTo ErrorHandler
    
    Dim wsInbound As Worksheet
    Dim wsOutbound As Worksheet
    Dim wsInventory As Worksheet
    Dim inboundLastRow As Long
    Dim outboundLastRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 获取实时库存表（必须已存在）
    On Error Resume Next
    Set wsInventory = ThisWorkbook.Worksheets("实时库存")
    On Error GoTo ErrorHandler
    
    If wsInventory Is Nothing Then
        MsgBox "请先手动创建"实时库存"工作表", vbCritical
        GoTo CleanUp
    End If
    
    ' 清除D5以下的数据和分组（保留表头D4-L4）
    Dim lastDataRow As Long
    lastDataRow = wsInventory.Cells(wsInventory.Rows.Count, 4).End(xlUp).Row  ' D列
    If lastDataRow >= 5 Then
        wsInventory.Rows("5:" & lastDataRow).ClearOutline
        wsInventory.Range("D5:L" & lastDataRow).ClearContents
        wsInventory.Range("D5:L" & lastDataRow).Interior.ColorIndex = xlNone
        wsInventory.Range("D5:L" & lastDataRow).Font.Bold = False
    End If
    
    Set wsInbound = ThisWorkbook.Worksheets("入库")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 获取入库表列索引
    Dim colInMaterialCode As Long
    Dim colInMaterialName As Long
    Dim colInManufacturer As Long
    Dim colInUnit As Long
    Dim colInSpec As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    
    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInMaterialName = GetColumnIndex(wsInbound, 1, "物料名称")
    colInManufacturer = GetColumnIndex(wsInbound, 1, "生产厂家")
    colInUnit = GetColumnIndex(wsInbound, 1, "单位")
    colInSpec = GetColumnIndex(wsInbound, 1, "规格")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")
    
    ' 验证必需列
    If colInMaterialCode = 0 Or colInBatch = 0 Or colInQty = 0 Then
        MsgBox "入库表缺少必需的列", vbCritical
        GoTo CleanUp
    End If
    
    ' 获取出库表列索引
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long
    
    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
    
    If colOutMaterialCode = 0 Or colOutBatch = 0 Or colOutQty = 0 Then
        MsgBox "出库表缺少必需的列", vbCritical
        GoTo CleanUp
    End If
    
    ' 收集数据
    inboundLastRow = wsInbound.Cells(wsInbound.Rows.Count, colInMaterialCode).End(xlUp).Row
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    
    ' 使用字典来组织数据结构
    Dim dictMaterials As Object
    Set dictMaterials = CreateObject("Scripting.Dictionary")
    
    ' 遍历入库表收集数据
    Dim i As Long, j As Long
    For i = 2 To inboundLastRow
        Dim matCode As String
        Dim matName As String
        Dim manufacturer As String
        Dim unit As String
        Dim spec As String
        Dim batch As String
        Dim inQty As Double
        
        matCode = Trim(wsInbound.Cells(i, colInMaterialCode).Value)
        If matCode = "" Then GoTo NextInRow
        
        matName = ""
        If colInMaterialName > 0 Then matName = Trim(wsInbound.Cells(i, colInMaterialName).Value)
        
        manufacturer = ""
        If colInManufacturer > 0 Then manufacturer = Trim(wsInbound.Cells(i, colInManufacturer).Value)
        
        unit = ""
        If colInUnit > 0 Then unit = Trim(wsInbound.Cells(i, colInUnit).Value)
        
        spec = ""
        If colInSpec > 0 Then spec = Trim(wsInbound.Cells(i, colInSpec).Value)
        
        batch = Trim(wsInbound.Cells(i, colInBatch).Value)
        inQty = Val(wsInbound.Cells(i, colInQty).Value)
        
        ' 计算该批次的出库量
        Dim outQty As Double
        outQty = 0
        For j = 2 To outboundLastRow
            If Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = matCode And _
               Trim(wsOutbound.Cells(j, colOutBatch).Value) = batch Then
                outQty = outQty + Val(wsOutbound.Cells(j, colOutQty).Value)
            End If
        Next j
        
        Dim stockQty As Double
        stockQty = inQty - outQty
        If stockQty < 0 Then stockQty = 0
        
        ' 创建唯一键：物料编号|生产厂家|规格|批次
        Dim uniqueKey As String
        uniqueKey = matCode & "|" & manufacturer & "|" & spec & "|" & batch
        
        ' 存储到字典
        If Not dictMaterials.Exists(matCode) Then
            Set dictMaterials(matCode) = CreateObject("Scripting.Dictionary")
            dictMaterials(matCode)("Name") = matName
            dictMaterials(matCode)("Unit") = unit
            Set dictMaterials(matCode)("Details") = CreateObject("Scripting.Dictionary")
        End If
        
        ' 存储明细
        Dim detailInfo(6) As Variant
        detailInfo(0) = manufacturer
        detailInfo(1) = spec
        detailInfo(2) = batch
        detailInfo(3) = inQty
        detailInfo(4) = outQty
        detailInfo(5) = stockQty
        
        dictMaterials(matCode)("Details")(uniqueKey) = detailInfo
        
NextInRow:
    Next i
    
    ' 写入数据到实时库存表
    Dim currentRow As Long
    currentRow = 5  ' 从第5行开始（D5）
    
    Dim matKeys As Variant
    matKeys = dictMaterials.Keys
    
    Dim k As Long
    For k = 0 To dictMaterials.Count - 1
        Dim matKey As String
        matKey = matKeys(k)
        
        Dim matData As Object
        Set matData = dictMaterials(matKey)
        
        ' 计算汇总数据
        Dim totalIn As Double, totalOut As Double, totalStock As Double
        totalIn = 0
        totalOut = 0
        totalStock = 0
        
        Dim detailKeys As Variant
        detailKeys = matData("Details").Keys
        
        Dim d As Long
        For d = 0 To matData("Details").Count - 1
            Dim detailArr As Variant
            detailArr = matData("Details")(detailKeys(d))
            totalIn = totalIn + detailArr(3)
            totalOut = totalOut + detailArr(4)
            totalStock = totalStock + detailArr(5)
        Next d
        
        ' 写入汇总行（从D列开始）
        Dim summaryRow As Long
        summaryRow = currentRow
        
        wsInventory.Cells(summaryRow, 4).Value = matKey               ' D列：物料编号
        wsInventory.Cells(summaryRow, 5).Value = matData("Name")      ' E列：物料名称
        wsInventory.Cells(summaryRow, 6).Value = "(汇总)"             ' F列：生产厂家
        wsInventory.Cells(summaryRow, 7).Value = matData("Unit")      ' G列：单位
        wsInventory.Cells(summaryRow, 8).Value = ""                   ' H列：规格（留空）
        wsInventory.Cells(summaryRow, 9).Value = ""                   ' I列：批次（留空）
        wsInventory.Cells(summaryRow, 10).Value = totalIn             ' J列：入库
        wsInventory.Cells(summaryRow, 11).Value = totalOut            ' K列：出库
        wsInventory.Cells(summaryRow, 12).Value = totalStock          ' L列：实时库存
        
        ' 格式化汇总行
        With wsInventory.Range(wsInventory.Cells(summaryRow, 4), wsInventory.Cells(summaryRow, 12))
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)  ' 浅蓝色背景
        End With
        
        currentRow = currentRow + 1
        
        ' 写入明细行
        Dim detailStartRow As Long
        detailStartRow = currentRow
        
        For d = 0 To matData("Details").Count - 1
            detailArr = matData("Details")(detailKeys(d))
            
            wsInventory.Cells(currentRow, 4).Value = matKey               ' D列：物料编号
            wsInventory.Cells(currentRow, 5).Value = matData("Name")      ' E列：物料名称
            wsInventory.Cells(currentRow, 6).Value = detailArr(0)         ' F列：生产厂家
            wsInventory.Cells(currentRow, 7).Value = matData("Unit")      ' G列：单位
            wsInventory.Cells(currentRow, 8).Value = detailArr(1)         ' H列：规格
            wsInventory.Cells(currentRow, 9).Value = detailArr(2)         ' I列：批次
            wsInventory.Cells(currentRow, 10).Value = detailArr(3)        ' J列：入库
            wsInventory.Cells(currentRow, 11).Value = detailArr(4)        ' K列：出库
            wsInventory.Cells(currentRow, 12).Value = detailArr(5)        ' L列：实时库存
            
            currentRow = currentRow + 1
        Next d
        
        ' 创建分组（只有当有明细行时）
        If matData("Details").Count > 0 Then
            wsInventory.Rows(detailStartRow & ":" & (currentRow - 1)).Group
        End If
        
        ' 每个物料编号之间空一行
        currentRow = currentRow + 1
    Next k
    
    ' 自动调整列宽
    wsInventory.Columns("D:L").AutoFit
    
    MsgBox "实时库存表生成完成！" & vbCrLf & "共 " & dictMaterials.Count & " 个物料", vbInformation
    
CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "生成实时库存表时发生错误: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' ============================================
' 静默刷新实时库存表（无弹窗）
' 表头位置: D4-L4
' 数据起始: D5
' ============================================
Sub RefreshRealTimeInventoryQuietly()
    On Error GoTo ErrorHandler
    
    Dim wsInbound As Worksheet
    Dim wsOutbound As Worksheet
    Dim wsInventory As Worksheet
    Dim inboundLastRow As Long
    Dim outboundLastRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 获取实时库存表
    On Error Resume Next
    Set wsInventory = ThisWorkbook.Worksheets("实时库存")
    On Error GoTo ErrorHandler
    
    If wsInventory Is Nothing Then
        ' 如果表不存在，直接退出
        GoTo CleanUp
    End If
    
    ' 清除D5以下的数据和分组（保留表头D4-L4）
    Dim lastDataRow As Long
    lastDataRow = wsInventory.Cells(wsInventory.Rows.Count, 4).End(xlUp).Row  ' D列
    If lastDataRow >= 5 Then
        wsInventory.Rows("5:" & lastDataRow).ClearOutline
        wsInventory.Range("D5:L" & lastDataRow).ClearContents
        wsInventory.Range("D5:L" & lastDataRow).Interior.ColorIndex = xlNone
        wsInventory.Range("D5:L" & lastDataRow).Font.Bold = False
    End If
    
    Set wsInbound = ThisWorkbook.Worksheets("入库")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 获取入库表列索引
    Dim colInMaterialCode As Long
    Dim colInMaterialName As Long
    Dim colInManufacturer As Long
    Dim colInUnit As Long
    Dim colInSpec As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    
    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInMaterialName = GetColumnIndex(wsInbound, 1, "物料名称")
    colInManufacturer = GetColumnIndex(wsInbound, 1, "生产厂家")
    colInUnit = GetColumnIndex(wsInbound, 1, "单位")
    colInSpec = GetColumnIndex(wsInbound, 1, "规格")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")
    
    ' 验证必需列
    If colInMaterialCode = 0 Or colInBatch = 0 Or colInQty = 0 Then
        GoTo CleanUp
    End If
    
    ' 获取出库表列索引
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long
    
    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
    
    If colOutMaterialCode = 0 Or colOutBatch = 0 Or colOutQty = 0 Then
        GoTo CleanUp
    End If
    
    ' 收集数据
    inboundLastRow = wsInbound.Cells(wsInbound.Rows.Count, colInMaterialCode).End(xlUp).Row
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    
    ' 使用字典来组织数据结构
    Dim dictMaterials As Object
    Set dictMaterials = CreateObject("Scripting.Dictionary")
    
    ' 遍历入库表收集数据
    Dim i As Long, j As Long
    For i = 2 To inboundLastRow
        Dim matCode As String
        Dim matName As String
        Dim manufacturer As String
        Dim unit As String
        Dim spec As String
        Dim batch As String
        Dim inQty As Double
        
        matCode = Trim(wsInbound.Cells(i, colInMaterialCode).Value)
        If matCode = "" Then GoTo NextInRow2
        
        matName = ""
        If colInMaterialName > 0 Then matName = Trim(wsInbound.Cells(i, colInMaterialName).Value)
        
        manufacturer = ""
        If colInManufacturer > 0 Then manufacturer = Trim(wsInbound.Cells(i, colInManufacturer).Value)
        
        unit = ""
        If colInUnit > 0 Then unit = Trim(wsInbound.Cells(i, colInUnit).Value)
        
        spec = ""
        If colInSpec > 0 Then spec = Trim(wsInbound.Cells(i, colInSpec).Value)
        
        batch = Trim(wsInbound.Cells(i, colInBatch).Value)
        inQty = Val(wsInbound.Cells(i, colInQty).Value)
        
        ' 计算该批次的出库量
        Dim outQty As Double
        outQty = 0
        For j = 2 To outboundLastRow
            If Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = matCode And _
               Trim(wsOutbound.Cells(j, colOutBatch).Value) = batch Then
                outQty = outQty + Val(wsOutbound.Cells(j, colOutQty).Value)
            End If
        Next j
        
        Dim stockQty As Double
        stockQty = inQty - outQty
        If stockQty < 0 Then stockQty = 0
        
        ' 创建唯一键：物料编号|生产厂家|规格|批次
        Dim uniqueKey As String
        uniqueKey = matCode & "|" & manufacturer & "|" & spec & "|" & batch
        
        ' 存储到字典
        If Not dictMaterials.Exists(matCode) Then
            Set dictMaterials(matCode) = CreateObject("Scripting.Dictionary")
            dictMaterials(matCode)("Name") = matName
            dictMaterials(matCode)("Unit") = unit
            Set dictMaterials(matCode)("Details") = CreateObject("Scripting.Dictionary")
        End If
        
        ' 存储明细
        Dim detailInfo(6) As Variant
        detailInfo(0) = manufacturer
        detailInfo(1) = spec
        detailInfo(2) = batch
        detailInfo(3) = inQty
        detailInfo(4) = outQty
        detailInfo(5) = stockQty
        
        dictMaterials(matCode)("Details")(uniqueKey) = detailInfo
        
NextInRow2:
    Next i
    
    ' 写入数据到实时库存表
    Dim currentRow As Long
    currentRow = 5  ' 从第5行开始（D5）
    
    Dim matKeys As Variant
    matKeys = dictMaterials.Keys
    
    Dim k As Long
    For k = 0 To dictMaterials.Count - 1
        Dim matKey As String
        matKey = matKeys(k)
        
        Dim matData As Object
        Set matData = dictMaterials(matKey)
        
        ' 计算汇总数据
        Dim totalIn As Double, totalOut As Double, totalStock As Double
        totalIn = 0
        totalOut = 0
        totalStock = 0
        
        Dim detailKeys As Variant
        detailKeys = matData("Details").Keys
        
        Dim d As Long
        For d = 0 To matData("Details").Count - 1
            Dim detailArr As Variant
            detailArr = matData("Details")(detailKeys(d))
            totalIn = totalIn + detailArr(3)
            totalOut = totalOut + detailArr(4)
            totalStock = totalStock + detailArr(5)
        Next d
        
        ' 写入汇总行（从D列开始）
        Dim summaryRow As Long
        summaryRow = currentRow
        
        wsInventory.Cells(summaryRow, 4).Value = matKey               ' D列：物料编号
        wsInventory.Cells(summaryRow, 5).Value = matData("Name")      ' E列：物料名称
        wsInventory.Cells(summaryRow, 6).Value = "(汇总)"             ' F列：生产厂家
        wsInventory.Cells(summaryRow, 7).Value = matData("Unit")      ' G列：单位
        wsInventory.Cells(summaryRow, 8).Value = ""                   ' H列：规格（留空）
        wsInventory.Cells(summaryRow, 9).Value = ""                   ' I列：批次（留空）
        wsInventory.Cells(summaryRow, 10).Value = totalIn             ' J列：入库
        wsInventory.Cells(summaryRow, 11).Value = totalOut            ' K列：出库
        wsInventory.Cells(summaryRow, 12).Value = totalStock          ' L列：实时库存
        
        ' 格式化汇总行
        With wsInventory.Range(wsInventory.Cells(summaryRow, 4), wsInventory.Cells(summaryRow, 12))
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)  ' 浅蓝色背景
        End With
        
        currentRow = currentRow + 1
        
        ' 写入明细行
        Dim detailStartRow As Long
        detailStartRow = currentRow
        
        For d = 0 To matData("Details").Count - 1
            detailArr = matData("Details")(detailKeys(d))
            
            wsInventory.Cells(currentRow, 4).Value = matKey               ' D列：物料编号
            wsInventory.Cells(currentRow, 5).Value = matData("Name")      ' E列：物料名称
            wsInventory.Cells(currentRow, 6).Value = detailArr(0)         ' F列：生产厂家
            wsInventory.Cells(currentRow, 7).Value = matData("Unit")      ' G列：单位
            wsInventory.Cells(currentRow, 8).Value = detailArr(1)         ' H列：规格
            wsInventory.Cells(currentRow, 9).Value = detailArr(2)         ' I列：批次
            wsInventory.Cells(currentRow, 10).Value = detailArr(3)        ' J列：入库
            wsInventory.Cells(currentRow, 11).Value = detailArr(4)        ' K列：出库
            wsInventory.Cells(currentRow, 12).Value = detailArr(5)        ' L列：实时库存
            
            currentRow = currentRow + 1
        Next d
        
        ' 创建分组（只有当有明细行时）
        If matData("Details").Count > 0 Then
            wsInventory.Rows(detailStartRow & ":" & (currentRow - 1)).Group
        End If
        
        ' 每个物料编号之间空一行
        currentRow = currentRow + 1
    Next k
    
    ' 自动调整列宽
    wsInventory.Columns("D:L").AutoFit
    
CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Resume CleanUp
End Sub

' ============================================
' 车间结存管理函数
' 创建日期：2026-02-11
' ============================================

' 获取车间结存量
' 优先返回动态计算的实时结存
Function GetWorkshopStock(materialCode As String) As Double
    On Error GoTo ErrorHandler
    
    ' 直接调用动态计算函数
    GetWorkshopStock = CalculateWorkshopRealStock(materialCode)
    Exit Function
    
ErrorHandler:
    GetWorkshopStock = 0
End Function

' 更新车间结存表的实时结存
Sub UpdateWorkshopStock(materialCode As String, newStock As Double)
    On Error GoTo ErrorHandler
    
    Dim wsWorkshop As Worksheet
    Dim wsMaterial As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim colMaterialCode As Long
    Dim colMaterialName As Long
    Dim colRealStock As Long
    Dim found As Boolean
    
    Set wsWorkshop = ThisWorkbook.Worksheets("车间结存")
    
    colMaterialCode = GetColumnIndex(wsWorkshop, 1, "物料编号")
    colMaterialName = GetColumnIndex(wsWorkshop, 1, "物料名称")
    colRealStock = GetColumnIndex(wsWorkshop, 1, "实时结存")
    
    If colMaterialCode = 0 Or colRealStock = 0 Then Exit Sub
    
    lastRow = wsWorkshop.Cells(wsWorkshop.Rows.Count, colMaterialCode).End(xlUp).Row
    found = False
    
    ' 查找物料编号
    For i = 2 To lastRow
        If Trim(wsWorkshop.Cells(i, colMaterialCode).Value) = materialCode Then
            ' 更新实时结存
            wsWorkshop.Cells(i, colRealStock).Value = newStock
            found = True
            Exit For
        End If
    Next i
    
    ' 如果没找到，新增一行
    If Not found Then
        Dim newRow As Long
        newRow = lastRow + 1
        wsWorkshop.Cells(newRow, colMaterialCode).Value = materialCode
        wsWorkshop.Cells(newRow, colRealStock).Value = newStock
        
        ' 从物料表获取物料名称
        If colMaterialName > 0 Then
            Set wsMaterial = ThisWorkbook.Worksheets("物料")
            Dim matColCode As Long
            Dim matColName As Long
            Dim matLastRow As Long
            Dim j As Long
            
            matColCode = GetColumnIndex(wsMaterial, 1, "物料编号")
            matColName = GetColumnIndex(wsMaterial, 1, "物料名称")
            
            If matColCode > 0 And matColName > 0 Then
                matLastRow = wsMaterial.Cells(wsMaterial.Rows.Count, matColCode).End(xlUp).Row
                
                For j = 2 To matLastRow
                    If Trim(wsMaterial.Cells(j, matColCode).Value) = materialCode Then
                        wsWorkshop.Cells(newRow, colMaterialName).Value = wsMaterial.Cells(j, matColName).Value
                        Exit For
                    End If
                Next j
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "更新车间结存时发生错误: " & Err.Description
End Sub

' ============================================
' 基于出库记录动态计算车间实时结存
' 创建日期：2026-02-11
' ============================================
Function CalculateWorkshopRealStock(materialCode As String) As Double
    On Error GoTo ErrorHandler
    
    Dim wsWorkshop As Worksheet
    Dim wsOutbound As Worksheet
    Dim initStock As Double
    Dim totalPickup As Double
    Dim totalUsage As Double
    Dim i As Long
    Dim lastRow As Long
    
    Set wsWorkshop = ThisWorkbook.Worksheets("车间结存")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 获取初期结存量
    Dim colMaterialCode As Long
    Dim colInitStock As Long
    
    colMaterialCode = GetColumnIndex(wsWorkshop, 1, "物料编号")
    colInitStock = GetColumnIndex(wsWorkshop, 1, "初期结存量")
    
    If colMaterialCode = 0 Or colInitStock = 0 Then
        CalculateWorkshopRealStock = 0
        Exit Function
    End If
    
    lastRow = wsWorkshop.Cells(wsWorkshop.Rows.Count, colMaterialCode).End(xlUp).Row
    
    ' 查找该物料的初期结存
    initStock = 0
    For i = 2 To lastRow
        If Trim(wsWorkshop.Cells(i, colMaterialCode).Value) = materialCode Then
            initStock = Val(wsWorkshop.Cells(i, colInitStock).Value)
            Exit For
        End If
    Next i
    
    ' 从出库表统计该物料的累计领用量和累计使用量
    Dim colOutMaterialCode As Long
    Dim colOutQty As Long
    Dim colOutWorkshopUsage As Long
    
    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
    colOutWorkshopUsage = GetColumnIndex(wsOutbound, 1, "车间使用量")
    
    If colOutMaterialCode = 0 Or colOutQty = 0 Then
        CalculateWorkshopRealStock = initStock
        Exit Function
    End If
    
    totalPickup = 0
    totalUsage = 0
    lastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(wsOutbound.Cells(i, colOutMaterialCode).Value) = materialCode Then
            totalPickup = totalPickup + Val(wsOutbound.Cells(i, colOutQty).Value)
            
            If colOutWorkshopUsage > 0 Then
                totalUsage = totalUsage + Val(wsOutbound.Cells(i, colOutWorkshopUsage).Value)
            End If
        End If
    Next i
    
    ' 计算实时结存 = 初期结存 + 累计领用 - 累计使用
    CalculateWorkshopRealStock = initStock + totalPickup - totalUsage
    
    ' 确保不为负数
    If CalculateWorkshopRealStock < 0 Then CalculateWorkshopRealStock = 0
    
    Exit Function
    
ErrorHandler:
    CalculateWorkshopRealStock = 0
End Function

' ============================================
' 刷新所有物料的车间实时结存
' ============================================
Sub RefreshAllWorkshopStock()
    On Error GoTo ErrorHandler
    
    Dim wsWorkshop As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim materialCode As String
    Dim realStock As Double
    Dim updateCount As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set wsWorkshop = ThisWorkbook.Worksheets("车间结存")
    
    Dim colMaterialCode As Long
    Dim colRealStock As Long
    
    colMaterialCode = GetColumnIndex(wsWorkshop, 1, "物料编号")
    colRealStock = GetColumnIndex(wsWorkshop, 1, "实时结存")
    
    If colMaterialCode = 0 Or colRealStock = 0 Then
        MsgBox "车间结存表缺少必需的列", vbCritical
        GoTo CleanUp
    End If
    
    lastRow = wsWorkshop.Cells(wsWorkshop.Rows.Count, colMaterialCode).End(xlUp).Row
    updateCount = 0
    
    For i = 2 To lastRow
        materialCode = Trim(wsWorkshop.Cells(i, colMaterialCode).Value)
        
        If materialCode <> "" Then
            ' 动态计算实时结存
            realStock = CalculateWorkshopRealStock(materialCode)
            
            ' 更新到车间结存表
            wsWorkshop.Cells(i, colRealStock).Value = realStock
            updateCount = updateCount + 1
        End If
    Next i
    
    MsgBox "车间实时结存已刷新！共更新 " & updateCount & " 个物料", vbInformation
    
CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "刷新车间结存时发生错误: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' ============================================
' 静默刷新车间结存（无弹窗）
' ============================================
Sub RefreshAllWorkshopStockQuietly()
    On Error Resume Next
    
    Dim wsWorkshop As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim materialCode As String
    Dim realStock As Double
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set wsWorkshop = ThisWorkbook.Worksheets("车间结存")
    
    Dim colMaterialCode As Long
    Dim colRealStock As Long
    
    colMaterialCode = GetColumnIndex(wsWorkshop, 1, "物料编号")
    colRealStock = GetColumnIndex(wsWorkshop, 1, "实时结存")
    
    If colMaterialCode = 0 Or colRealStock = 0 Then GoTo CleanUp
    
    lastRow = wsWorkshop.Cells(wsWorkshop.Rows.Count, colMaterialCode).End(xlUp).Row
    
    For i = 2 To lastRow
        materialCode = Trim(wsWorkshop.Cells(i, colMaterialCode).Value)
        
        If materialCode <> "" Then
            realStock = CalculateWorkshopRealStock(materialCode)
            wsWorkshop.Cells(i, colRealStock).Value = realStock
        End If
    Next i
    
    Application.StatusBar = "车间结存已刷新 " & Format(Now, "hh:mm:ss")
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBarQuiet"
    
CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

