Option Explicit

' ============================================
' 车间结存表工作表代码模块
' 功能：输入物料编号时，自动从物料表获取名称
'       激活表时自动刷新实时结存
'       ?? 历史实时结存倒查（通过DTPicker和K3产品编号）
' 创建日期：2026-02-11
' 修改日期：2026-02-11 - 添加历史库存查询功能
'           2026-02-11 - 修复空白行显示0的问题
'           2026-02-11 - 修复DTPicker选择日期不触发查询的问题
'           2026-02-11 - 优化：初期结存量为空时默认按0计算
'           2026-02-11 - 再次优化：解决大量空白行显示0的问题
'           2026-02-11 - 修正：历史数据写入起始行从8行改为7行
' ============================================

' ?? 上次出库表的行数（用于检测出库表变化）
Private lastOutboundRowCount As Long

' ?? 查询进行中标志（防止重复触发）
Private isQuerying As Boolean

' ?? 工作表激活时检测出库表变化并刷新实时结存
Private Sub Worksheet_Activate()
    On Error Resume Next
    
    Dim wsOutbound As Worksheet
    Dim currentOutboundRowCount As Long
    
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    currentOutboundRowCount = wsOutbound.Cells(wsOutbound.Rows.Count, 1).End(xlUp).Row
    
    ' 如果是首次加载，记录当前行数
    If lastOutboundRowCount = 0 Then
        lastOutboundRowCount = currentOutboundRowCount
        ' 首次加载也刷新一次
        Call RefreshAllWorkshopStockQuietly
        Exit Sub
    End If
    
    ' 检查出库表行数是否有变化
    If currentOutboundRowCount <> lastOutboundRowCount Then
        ' 出库表有变化，刷新实时结存
        lastOutboundRowCount = currentOutboundRowCount
        Call RefreshAllWorkshopStockQuietly
        
        ' 显示提示
        Application.StatusBar = "车间结存已自动刷新 " & Format(Now, "hh:mm:ss")
        Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBarQuiet"
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' ?? 检查是否修改了K3单元格（产品编号）
    If Not Intersect(Target, Me.Range("K3")) Is Nothing Then
        Application.EnableEvents = False
        Call QueryHistoricalStockAuto
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Dim colMaterialCode As Long
    Dim colMaterialName As Long
    
    ' 获取列索引
    colMaterialCode = GetColumnIndex(Me, 1, "物料编号")
    colMaterialName = GetColumnIndex(Me, 1, "物料名称")
    
    If colMaterialCode = 0 Then Exit Sub
    
    ' 检查是否修改了物料编号列
    If Intersect(Target, Me.Columns(colMaterialCode)) Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    
    Dim cell As Range
    For Each cell In Target
        If cell.Column = colMaterialCode Then
            Dim materialCode As String
            materialCode = Trim(cell.Value)
            
            If materialCode <> "" Then
                ' 从物料表获取物料名称
                Dim wsMaterial As Worksheet
                Set wsMaterial = ThisWorkbook.Worksheets("物料")
                
                Dim matColCode As Long
                Dim matColName As Long
                Dim matLastRow As Long
                Dim i As Long
                
                matColCode = GetColumnIndex(wsMaterial, 1, "物料编号")
                matColName = GetColumnIndex(wsMaterial, 1, "物料名称")
                
                If matColCode > 0 And matColName > 0 Then
                    matLastRow = wsMaterial.Cells(wsMaterial.Rows.Count, matColCode).End(xlUp).Row
                    
                    For i = 2 To matLastRow
                        If Trim(wsMaterial.Cells(i, matColCode).Value) = materialCode Then
                            Me.Cells(cell.Row, colMaterialName).Value = wsMaterial.Cells(i, matColName).Value
                            Exit For
                        End If
                    Next i
                End If
            Else
                ' 清空物料名称
                If colMaterialName > 0 Then
                    Me.Cells(cell.Row, colMaterialName).ClearContents
               End If
            End If
        End If
    Next cell
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "填充物料名称时发生错误: " & Err.Description, vbCritical
End Sub

' ?? DTPicker日期选择器 - 日期值改变时触发（修复核心：改用Change事件）
' 点击/选择/修改日期都会触发此事件，替代原有的CallbackKeyDown
Private Sub DTPicker1_Change()
    On Error Resume Next
    
    ' 检查K3是否有产品编号
    If Trim(Me.Range("K3").Value) <> "" Then
        ' 有产品编号，执行查询
        Call QueryHistoricalStockAuto
    Else
        ' 没有产品编号，清空结果区域
        Call ClearHistoricalResultArea
    End If
End Sub

' ============================================
' ?? 自动查询历史实时结存
' 触发条件：DTPicker日期改变 或 K3产品编号改变
' ============================================
Private Sub QueryHistoricalStockAuto()
    On Error GoTo ErrorHandler
    
    ' 如果正在查询中，跳过
    If isQuerying Then Exit Sub
    
    Dim productCode As String
    Dim queryDate As Date
    
    ' 读取产品编号
    productCode = Trim(Me.Range("K3").Value)
    
    ' 如果产品编号为空，清空结果区域并退出
    If productCode = "" Then
        Call ClearHistoricalResultArea
        Exit Sub
    End If
    
    ' 读取查询日期
    On Error Resume Next
    queryDate = DTPicker1.Value
    If Err.Number <> 0 Then
        ' DTPicker不可用或未选择日期，使用当前日期
        queryDate = Date
    End If
    On Error GoTo ErrorHandler
    
    ' 设置查询标志
    isQuerying = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 执行查询
    Call QueryHistoricalStock(productCode, queryDate)
    
    ' 重置标志
    isQuerying = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    isQuerying = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' 静默错误，不弹窗
End Sub

' ============================================
' ?? 查询历史实时结存（核心函数）
' 参数：productCode - 产品编号
'       queryDate - 查询日期
' 修改日期：2026-02-11 - 新增K列（生产批号）、L列（上批结存）、M列（上批生产批号）、N列（差异）
' ============================================
Private Sub QueryHistoricalStock(productCode As String, queryDate As Date)
    On Error GoTo ErrorHandler
    
    Dim wsBOM As Worksheet
    Dim wsWorkshop As Worksheet
    Dim wsOutbound As Worksheet
    Dim bomList() As Variant
    Dim bomCount As Long
    Dim i As Long
    Dim currentRow As Long
    
    Set wsBOM = ThisWorkbook.Worksheets("BOM")
    Set wsWorkshop = ThisWorkbook.Worksheets("车间结存")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 步骤1：清空旧数据
   Call ClearHistoricalResultArea
    
    ' 步骤2：从BOM表获取该产品的物料清单
    bomList = GetProductBOM(productCode)
    bomCount = UBound(bomList, 1)
    
    If bomCount = 0 Then
        ' 未找到BOM数据
        Exit Sub
    End If
    
    ' 步骤3：从第7行开始填写结果（核心修正：从8行改为7行）
    currentRow = 7
    
    ' 步骤3.5：获取当前查询日期的生产批号和上批生产批号
    Dim currentProductionBatch As String
    Dim previousProductionBatch As String
    
    currentProductionBatch = GetProductionBatchByDate(productCode, queryDate)
    previousProductionBatch = GetPreviousProductionBatch(productCode, queryDate)
    
    For i = 1 To bomCount
        Dim materialCode As String
        Dim materialName As String
        Dim historicalStock As Variant ' 修改为Variant类型，支持空值
        Dim previousStock As Double    ' 上批结存
        Dim difference As Double       ' 差异
        
        materialCode = bomList(i, 1)  ' 物料编号
        materialName = bomList(i, 2)  ' 物料名称
        
        ' 计算历史实时结存
        historicalStock = CalculateHistoricalStock(materialCode, queryDate)
        
        ' 填写结果（仅当有有效数值时才填充，避免大量空白行）
        If historicalStock <> "" Then
            ' 计算上批结存（基于上一个出库日期）
            previousStock = CalculatePreviousBatchStock(materialCode, queryDate)
            
            ' 计算差异
            difference = IIf(historicalStock = "", 0, historicalStock) - previousStock
            
            Me.Cells(currentRow, "F").Value = materialCode                                  ' F列：物料编号
            Me.Cells(currentRow, "H").Value = materialName                                  ' H列：物料名称
            Me.Cells(currentRow, "J").Value = IIf(historicalStock = "", 0, historicalStock) ' J列：历史实时结存
            Me.Cells(currentRow, "K").Value = currentProductionBatch                        ' K列：生产批号
            Me.Cells(currentRow, "L").Value = previousStock                                 ' L列：上批结存
            Me.Cells(currentRow, "M").Value = previousProductionBatch                       ' M列：上批生产批号
            Me.Cells(currentRow, "N").Value = difference                                    ' N列：差异
            currentRow = currentRow + 1
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    ' 静默错误
End Sub

' ============================================
' ?? 从BOM表获取产品的物料清单
' 返回：二维数组(物料编号, 物料名称)
' ============================================
Private Function GetProductBOM(productCode As String) As Variant
    On Error Resume Next
    
    Dim wsBOM As Worksheet
    Dim colProductCode As Long
    Dim colMaterialCode As Long
    Dim colMaterialName As Long
    Dim lastRow As Long
    Dim i As Long
    Dim resultArray() As Variant
    Dim resultCount As Long
    
    Set wsBOM = ThisWorkbook.Worksheets("BOM")
    
    ' 获取列索引
    colProductCode = GetColumnIndex(wsBOM, 1, "产品编号")
    colMaterialCode = GetColumnIndex(wsBOM, 1, "物料编号")
    colMaterialName = GetColumnIndex(wsBOM, 1, "物料名称")
    
    If colProductCode = 0 Or colMaterialCode = 0 Then
        ' 返回空数组
        ReDim resultArray(0, 2)
        GetProductBOM = resultArray
        Exit Function
    End If
    
    lastRow = wsBOM.Cells(wsBOM.Rows.Count, colProductCode).End(xlUp).Row
    
    ' 初步分配数组空间
    ReDim resultArray(1 To 100, 1 To 2)
    resultCount = 0
    
    ' 遍历BOM表查找匹配的产品编号
    For i = 2 To lastRow
        If Trim(wsBOM.Cells(i, colProductCode).Value) = productCode Then
            resultCount = resultCount + 1
            resultArray(resultCount, 1) = Trim(wsBOM.Cells(i, colMaterialCode).Value)
            
            If colMaterialName > 0 Then
                resultArray(resultCount, 2) = Trim(wsBOM.Cells(i, colMaterialName).Value)
            Else
                resultArray(resultCount, 2) = ""
            End If
        End If
    Next i
    
    ' 调整数组大小
    If resultCount > 0 Then
        ReDim Preserve resultArray(1 To resultCount, 1 To 2)
    Else
        ReDim resultArray(0, 2)
    End If
    
    GetProductBOM = resultArray
End Function

' ============================================
' ?? 计算物料的历史实时结存
' 公式：历史结存 = 初期结存 + 截止日期累计领用 - 截止日期累计使用
' 优化点：
' 1. 初期结存量为空时默认按0计算。
' 2. 找不到物料时返回空字符串，以便上层函数判断是否填充行。
' ============================================
Private Function CalculateHistoricalStock(materialCode As String, queryDate As Date) As Variant
    On Error Resume Next
    
    Dim wsWorkshop As Worksheet
    Dim wsOutbound As Worksheet
    Dim initStock As Double
    Dim cumulativePickup As Double
    Dim cumulativeUsage As Double
    Dim lastRow As Long
    Dim i As Long
    Dim isMaterialFound As Boolean ' 标记是否找到对应物料行
    
    Set wsWorkshop = ThisWorkbook.Worksheets("车间结存")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 步骤1：从车间结存表获取初期结存
    Dim colMaterialCode As Long
    Dim colInitStock As Long
    
    colMaterialCode = GetColumnIndex(wsWorkshop, 1, "物料编号")
    colInitStock = GetColumnIndex(wsWorkshop, 1, "初期结存量")
    
    If colMaterialCode = 0 Or colInitStock = 0 Then
        ' 找不到物料列/初期结存列，返回空字符串
        CalculateHistoricalStock = ""
        Exit Function
    End If
    
    lastRow = wsWorkshop.Cells(wsWorkshop.Rows.Count, colMaterialCode).End(xlUp).Row
    initStock = 0
    isMaterialFound = False
    
    For i = 2 To lastRow
        If Trim(wsWorkshop.Cells(i, colMaterialCode).Value) = materialCode Then
            isMaterialFound = True
            ' 单元格为空时，Val()返回0，实现默认等于0的逻辑
            initStock = Val(wsWorkshop.Cells(i, colInitStock).Value)
            Exit For
        End If
    Next i
    
    ' 如果找不到物料行，返回空字符串，以便上层函数跳过该行
    If Not isMaterialFound Then
        CalculateHistoricalStock = ""
        Exit Function
    End If
    
    ' 步骤2：从出库表统计截止日期的累计数据
    Dim colOutDate As Long
    Dim colOutMaterialCode As Long
    Dim colOutQty As Long
    Dim colOutWorkshopUsage As Long
    Dim outDate As Date
    
    colOutDate = GetColumnIndex(wsOutbound, 1, "日期")
    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
    colOutWorkshopUsage = GetColumnIndex(wsOutbound, 1, "车间使用量")
    
    If colOutDate = 0 Or colOutMaterialCode = 0 Or colOutQty = 0 Then
        CalculateHistoricalStock = initStock
        Exit Function
    End If
    
    lastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    cumulativePickup = 0
    cumulativeUsage = 0
    
    ' 遍历出库表，只统计日期 <= queryDate 的记录
    For i = 2 To lastRow
        If Trim(wsOutbound.Cells(i, colOutMaterialCode).Value) = materialCode Then
            ' 检查日期
            If Not IsEmpty(wsOutbound.Cells(i, colOutDate)) Then
                On Error Resume Next
                outDate = CDate(wsOutbound.Cells(i, colOutDate).Value)
                If Err.Number = 0 Then
                    ' 只统计 <= queryDate 的记录
                    If outDate <= queryDate Then
                        cumulativePickup = cumulativePickup + Val(wsOutbound.Cells(i, colOutQty).Value)
                        
                        If colOutWorkshopUsage > 0 Then
                            cumulativeUsage = cumulativeUsage + Val(wsOutbound.Cells(i, colOutWorkshopUsage).Value)
                        End If
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' 步骤3：计算历史实时结存
    CalculateHistoricalStock = initStock + cumulativePickup - cumulativeUsage
    
    ' 确保不为负数
    If CalculateHistoricalStock < 0 Then CalculateHistoricalStock = 0
    
    ' 四舍五入到2位小数
    CalculateHistoricalStock = Round(CalculateHistoricalStock, 2)
End Function

' ============================================
' ?? 计算物料的上批结存（基于上一个出库日期）
' 参数：materialCode - 物料编号
'       queryDate - 查询日期
' 返回：上一个出库日期的实时结存
' 创建日期：2026-02-11
' ============================================
Private Function CalculatePreviousBatchStock(materialCode As String, queryDate As Date) As Double
    On Error Resume Next
    
    Dim wsWorkshop As Worksheet
    Dim wsOutbound As Worksheet
    Dim initStock As Double
    Dim cumulativePickup As Double
    Dim cumulativeUsage As Double
    Dim lastRow As Long
    Dim i As Long
    Dim previousDate As Date
    Dim foundPreviousDate As Boolean
    
    Set wsWorkshop = ThisWorkbook.Worksheets("车间结存")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 步骤1：查找该物料在queryDate之前的最后一个出库日期
    Dim colOutDate As Long
    Dim colOutMaterialCode As Long
    Dim outDate As Date
    
    colOutDate = GetColumnIndex(wsOutbound, 1, "日期")
    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    
    If colOutDate = 0 Or colOutMaterialCode = 0 Then
        CalculatePreviousBatchStock = 0
        Exit Function
    End If
    
    lastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    previousDate = #1/1/1900#  ' 初始化为最小日期
    foundPreviousDate = False
    
    ' 遍历出库表，查找该物料在queryDate之前的最后一个出库日期
    For i = 2 To lastRow
        If Trim(wsOutbound.Cells(i, colOutMaterialCode).Value) = materialCode Then
            If Not IsEmpty(wsOutbound.Cells(i, colOutDate)) Then
                On Error Resume Next
                outDate = CDate(wsOutbound.Cells(i, colOutDate).Value)
                If Err.Number = 0 Then
                    ' 找到小于queryDate的日期，且大于当前记录的previousDate
                    If outDate < queryDate And outDate > previousDate Then
                        previousDate = outDate
                        foundPreviousDate = True
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' 步骤2：如果没找到上一个出库日期，返回初期结存
    If Not foundPreviousDate Then
        ' 从车间结存表获取初期结存
        Dim colMaterialCode As Long
        Dim colInitStock As Long
        
        colMaterialCode = GetColumnIndex(wsWorkshop, 1, "物料编号")
        colInitStock = GetColumnIndex(wsWorkshop, 1, "初期结存量")
        
        If colMaterialCode = 0 Or colInitStock = 0 Then
            CalculatePreviousBatchStock = 0
            Exit Function
        End If
        
        lastRow = wsWorkshop.Cells(wsWorkshop.Rows.Count, colMaterialCode).End(xlUp).Row
        initStock = 0
        
        For i = 2 To lastRow
            If Trim(wsWorkshop.Cells(i, colMaterialCode).Value) = materialCode Then
                initStock = Val(wsWorkshop.Cells(i, colInitStock).Value)
                Exit For
            End If
        Next i
        
        CalculatePreviousBatchStock = initStock
        Exit Function
    End If
    
    ' 步骤3：计算previousDate的实时结存
    ' 逻辑：初期结存 + 截止previousDate的累计领用 - 截止previousDate的累计使用
    
    ' 获取初期结存
    colMaterialCode = GetColumnIndex(wsWorkshop, 1, "物料编号")
    colInitStock = GetColumnIndex(wsWorkshop, 1, "初期结存量")
    
    If colMaterialCode = 0 Or colInitStock = 0 Then
        CalculatePreviousBatchStock = 0
        Exit Function
    End If
    
    lastRow = wsWorkshop.Cells(wsWorkshop.Rows.Count, colMaterialCode).End(xlUp).Row
    initStock = 0
    
    For i = 2 To lastRow
        If Trim(wsWorkshop.Cells(i, colMaterialCode).Value) = materialCode Then
            initStock = Val(wsWorkshop.Cells(i, colInitStock).Value)
            Exit For
        End If
    Next i
    
    ' 从出库表统计截止previousDate的累计数据
    Dim colOutQty As Long
    Dim colOutWorkshopUsage As Long
    
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
    colOutWorkshopUsage = GetColumnIndex(wsOutbound, 1, "车间使用量")
    
    If colOutQty = 0 Then
        CalculatePreviousBatchStock = initStock
        Exit Function
    End If
    
    lastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    cumulativePickup = 0
    cumulativeUsage = 0
    
    ' 遍历出库表，只统计日期 <= previousDate 的记录
    For i = 2 To lastRow
        If Trim(wsOutbound.Cells(i, colOutMaterialCode).Value) = materialCode Then
            If Not IsEmpty(wsOutbound.Cells(i, colOutDate)) Then
                On Error Resume Next
                outDate = CDate(wsOutbound.Cells(i, colOutDate).Value)
                If Err.Number = 0 Then
                    ' 只统计 <= previousDate 的记录
                    If outDate <= previousDate Then
                        cumulativePickup = cumulativePickup + Val(wsOutbound.Cells(i, colOutQty).Value)
                        
                        If colOutWorkshopUsage > 0 Then
                            cumulativeUsage = cumulativeUsage + Val(wsOutbound.Cells(i, colOutWorkshopUsage).Value)
                        End If
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' 计算上批结存
    CalculatePreviousBatchStock = initStock + cumulativePickup - cumulativeUsage
    
    ' 确保不为负数
    If CalculatePreviousBatchStock < 0 Then CalculatePreviousBatchStock = 0
    
    ' 四舍五入到2位小数
    CalculatePreviousBatchStock = Round(CalculatePreviousBatchStock, 2)
End Function

' ============================================
' ?? 获取指定日期对应的生产批号
' 参数：productCode - 产品编号
'       queryDate - 查询日期
' 返回：该日期的生产批号（如果有）
' 创建日期：2026-02-11
' ============================================
Private Function GetProductionBatchByDate(productCode As String, queryDate As Date) As String
    On Error Resume Next
    
    Dim wsProduction As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim colDate As Long
    Dim colProductCode As Long
    Dim colProductionBatch As Long
    Dim recordDate As Date
    
    Set wsProduction = ThisWorkbook.Worksheets("生产记录")
    
    ' 获取列索引
    colDate = GetColumnIndex(wsProduction, 1, "日期")
    colProductCode = GetColumnIndex(wsProduction, 1, "产品编号")
    colProductionBatch = GetColumnIndex(wsProduction, 1, "生产批号")
    
    If colDate = 0 Or colProductCode = 0 Or colProductionBatch = 0 Then
        GetProductionBatchByDate = ""
        Exit Function
    End If
    
    lastRow = wsProduction.Cells(wsProduction.Rows.Count, colProductCode).End(xlUp).Row
    
    ' 遍历生产记录表，查找匹配的产品编号和日期
    For i = 2 To lastRow
        If Trim(wsProduction.Cells(i, colProductCode).Value) = productCode Then
            If Not IsEmpty(wsProduction.Cells(i, colDate)) Then
                On Error Resume Next
                recordDate = CDate(wsProduction.Cells(i, colDate).Value)
                If Err.Number = 0 Then
                    ' 检查日期是否匹配
                    If Int(recordDate) = Int(queryDate) Then
                        GetProductionBatchByDate = Trim(wsProduction.Cells(i, colProductionBatch).Value)
                        Exit Function
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' 如果没找到，返回空字符串
    GetProductionBatchByDate = ""
End Function

' ============================================
' ?? 获取上一批的生产批号
' 参数：productCode - 产品编号
'       queryDate - 查询日期
' 返回：上一批的生产批号
' 创建日期：2026-02-11
' ============================================
Private Function GetPreviousProductionBatch(productCode As String, queryDate As Date) As String
    On Error Resume Next
    
    Dim wsProduction As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim colDate As Long
    Dim colProductCode As Long
    Dim colProductionBatch As Long
    Dim recordDate As Date
    Dim previousDate As Date
    Dim previousBatch As String
    Dim foundPrevious As Boolean
    
    Set wsProduction = ThisWorkbook.Worksheets("生产记录")
    
    ' 获取列索引
    colDate = GetColumnIndex(wsProduction, 1, "日期")
    colProductCode = GetColumnIndex(wsProduction, 1, "产品编号")
    colProductionBatch = GetColumnIndex(wsProduction, 1, "生产批号")
    
    If colDate = 0 Or colProductCode = 0 Or colProductionBatch = 0 Then
        GetPreviousProductionBatch = ""
        Exit Function
    End If
    
    lastRow = wsProduction.Cells(wsProduction.Rows.Count, colProductCode).End(xlUp).Row
    previousDate = #1/1/1900#
    foundPrevious = False
    
    ' 遍历生产记录表，查找该产品在queryDate之前的最后一个生产日期
    For i = 2 To lastRow
        If Trim(wsProduction.Cells(i, colProductCode).Value) = productCode Then
            If Not IsEmpty(wsProduction.Cells(i, colDate)) Then
                On Error Resume Next
                recordDate = CDate(wsProduction.Cells(i, colDate).Value)
                If Err.Number = 0 Then
                    ' 找到小于queryDate的日期，且大于当前记录的previousDate
                    If recordDate < queryDate And recordDate > previousDate Then
                        previousDate = recordDate
                        previousBatch = Trim(wsProduction.Cells(i, colProductionBatch).Value)
                        foundPrevious = True
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    If foundPrevious Then
        GetPreviousProductionBatch = previousBatch
    Else
        GetPreviousProductionBatch = ""
    End If
End Function

' ============================================
' ?? 清空历史查询结果区域
' 修改点：
' 1. 清空起始行从8行改为7行，与数据写入行保持一致
' 2. 清空固定大区域（F7:N1000），确保所有旧数据都被清空
' 修改日期：2026-02-11 - 扩展清空范围到K、L、M、N列
' ============================================
Private Sub ClearHistoricalResultArea()
    On Error Resume Next
    
    ' 清空从第7行到第1000行的F-N列，确保所有旧数据都被清除
    Me.Range("F7:N1000").ClearContents
End Sub

' 辅助函数：获取列索引
Private Function GetColumnIndex(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim col As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        If Trim(ws.Cells(headerRow, col).Value) = Trim(headerName) Then
            GetColumnIndex = col
            Exit Function
        End If
    Next col
    GetColumnIndex = 0
End Function