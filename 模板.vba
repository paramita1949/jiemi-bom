Option Explicit

' 调试开关
Private Const DEBUG_LOG As Boolean = False

' 模块级变量：保存配方数据，用于计算需求量
Private formulaData() As Variant
Private currentProductCode As String
Private mustDivideData() As Variant  ' 保存整除标志（Y/N）


' 当 E3 或 C4 或 R3 单元格内容改变时触发
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    
    ' ?? 检查是否是 R3 单元格被修改（生产批号）
    If Not Intersect(Target, Me.Range("R3")) Is Nothing Then
        Call QueryProductionHistory
    End If
    
    ' 检查是否是 E3 单元格被修改
    If Not Intersect(Target, Me.Range("E3")) Is Nothing Then
        Call FillBOMData
    End If
    
    ' 检查是否是 C4 单元格被修改（成品需求量）
    If Not Intersect(Target, Me.Range("C4")) Is Nothing Then
        Call CalculateRequirements
    End If
    
    ' ?? 检查是否是 E列（需求量）被修改，同步到M列（入库）
    Dim colRequirement As Long
    colRequirement = 5  ' E列
    
    If Not Intersect(Target, Me.Columns(colRequirement)) Is Nothing Then
        If Target.Row >= 6 Then  ' 只处理数据行
            Dim cell As Range
            For Each cell In Intersect(Target, Me.Columns(colRequirement))
                If cell.Row >= 6 And Not IsEmpty(Me.Cells(cell.Row, "B")) Then
                    ' E列变化时，同步到M列（入库）
                    Me.Cells(cell.Row, 13).Value = cell.Value  ' M列 = E列
                End If
            Next cell
        End If
    End If
    
    Application.EnableEvents = True
End Sub

' 工作表激活或打开时自动填充日期
Private Sub Worksheet_Activate()
    Call FillDates
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

' 主要的数据填充逻辑（优化版）
Sub FillBOMData()
    Dim wsTemplate As Worksheet
    Dim wsData As Worksheet
    Dim productCode As String
    Dim productName As String
    Dim dataRow As Long
    Dim templateRow As Long
    Dim bomCount As Integer
    Dim lastRow As Long
    
    ' BOM 表的列索引（动态获取）
    Dim colProductCode As Long
    Dim colProductName As Long
    Dim colMaterialCode As Long
    Dim colMaterialName As Long
    Dim colSpec As Long
    Dim colUnit As Long
    Dim colManufacturer As Long
    Dim colFormula As Long  ' 配方列
    Dim colMustDivide As Long  ' 整除列
    
    ' 用于批量处理的数组
    Dim bomData() As Variant
    Dim i As Long
    
    ' 关闭屏幕更新和事件，提高性能
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual  ' 关闭自动计算
    
    On Error GoTo ErrorHandler
    
    ' 设置工作表
    Set wsTemplate = ThisWorkbook.Worksheets("模板")
    Set wsData = ThisWorkbook.Worksheets("BOM") ' BOM 数据表
    
    ' 获取 BOM 表的列索引（从第1行表头读取）
    colProductCode = GetColumnIndex(wsData, 1, "产品编号")
    colProductName = GetColumnIndex(wsData, 1, "产品名称")
    colMaterialCode = GetColumnIndex(wsData, 1, "物料编号")
    colMaterialName = GetColumnIndex(wsData, 1, "物料名称")
    colSpec = GetColumnIndex(wsData, 1, "规格")
    colUnit = GetColumnIndex(wsData, 1, "单位")
    colManufacturer = GetColumnIndex(wsData, 1, "生产厂家")
    colFormula = GetColumnIndex(wsData, 1, "配方")  ' 获取配方列索引
    colMustDivide = GetColumnIndex(wsData, 1, "整除")  ' 获取整除列索引
    
    ' 验证必需的列是否存在
    If colProductCode = 0 Or colMaterialCode = 0 Then
        MsgBox "BOM 表中缺少必需的列（产品编号或物料编号），请检查表头！", vbCritical, "错误"
        GoTo CleanUp
    End If
    
    ' 获取产品编码
    productCode = Trim(wsTemplate.Range("E3").Value)
    
    ' 如果产品编码为空，清空数据区域并退出
    If productCode = "" Then
        Call ClearBOMArea
        GoTo CleanUp
    End If
    
    ' 清空之前的 BOM 数据
    Call ClearBOMArea
    
    ' 第一步：先遍历数据，计算匹配的记录数
    bomCount = 0
    productName = ""
    lastRow = wsData.Cells(wsData.Rows.Count, colProductCode).End(xlUp).Row
    
    For dataRow = 2 To lastRow
        If Trim(wsData.Cells(dataRow, colProductCode).Value) = productCode Then
            If productName = "" And colProductName > 0 Then
                productName = wsData.Cells(dataRow, colProductName).Value
            End If
            bomCount = bomCount + 1
        End If
    Next dataRow
    
    ' 如果没有找到数据，退出
    If bomCount = 0 Then
        MsgBox "未找到产品编号: " & productCode, vbExclamation, "提示"
        GoTo CleanUp
    End If
    
    ' 第二步：一次性插入所需的行数（从第7行开始插入bomCount-1行）
    If bomCount > 1 Then
        wsTemplate.Rows("7:" & (6 + bomCount - 1)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ' 一次性复制格式
        wsTemplate.Rows(6).Copy
        wsTemplate.Rows("7:" & (6 + bomCount - 1)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' 第三步：准备数据数组（避免循环中多次访问单元格）
    ' 确保bomCount大于0，避免创建无效数组
    If bomCount > 0 Then
        ReDim bomData(1 To bomCount, 1 To 9)
        ReDim formulaData(1 To bomCount)  ' 保存每行的配方数字
        ReDim mustDivideData(1 To bomCount)  ' 保存每行的整除标志（Y/N）
    Else
        ' 如果bomCount为0，创建1个元素的数组作为占位符
        ReDim bomData(1 To 1, 1 To 9)
        ReDim formulaData(1 To 1)
        ReDim mustDivideData(1 To 1)
    End If
    
    ' 重新遍历，填充数组
    i = 0
    For dataRow = 2 To lastRow
        If Trim(wsData.Cells(dataRow, colProductCode).Value) = productCode Then
            i = i + 1
            
            ' A列：序号
            bomData(i, 1) = i
            
            ' B列：物料编号
            If colMaterialCode > 0 Then
                bomData(i, 2) = wsData.Cells(dataRow, colMaterialCode).Value
            End If
            
            ' C列：物料名称
            If colMaterialName > 0 Then
                bomData(i, 3) = wsData.Cells(dataRow, colMaterialName).Value
            End If
            
            ' D列：规格
            If colSpec > 0 Then
                bomData(i, 4) = wsData.Cells(dataRow, colSpec).Value
            End If
            
            ' E列：需求量 - 保持空白（稍后通过C4计算）
            bomData(i, 5) = ""
            
            ' F列：单位
            If colUnit > 0 Then
                bomData(i, 6) = wsData.Cells(dataRow, colUnit).Value
            End If
            
            ' G列：车间结存量 - ?? 从车间结存表获取
            Dim workshopStock As Double
            Dim materialCodeForStock As String
            If colMaterialCode > 0 Then
                materialCodeForStock = Trim(wsData.Cells(dataRow, colMaterialCode).Value)
                workshopStock = GetWorkshopStock(materialCodeForStock)
                bomData(i, 7) = workshopStock
            Else
                bomData(i, 7) = ""
            End If
            
            ' H列：本次领用量 - 保持空白
            bomData(i, 8) = ""
            
            ' I列：生产厂家
            If colManufacturer > 0 Then
                bomData(i, 9) = wsData.Cells(dataRow, colManufacturer).Value
            End If
            
            ' 保存配方数字到隐藏数组（用于后续计算）
            If colFormula > 0 Then
                formulaData(i) = wsData.Cells(dataRow, colFormula).Value
            Else
                formulaData(i) = 1  ' 默认值
            End If

            ' 保存整除标志到隐藏数组（用于后续计算）
            If colMustDivide > 0 Then
                mustDivideData(i) = Trim(wsData.Cells(dataRow, colMustDivide).Value)
            Else
                mustDivideData(i) = "Y"  ' 默认为Y，表示必须整除
            End If
        End If
    Next dataRow
    
    ' 第四步：一次性写入所有数据
    wsTemplate.Range("A6").Resize(bomCount, 9).Value = bomData



    ' 保存当前产品编号（供CalculateRequirements使用）
    currentProductCode = productCode

    ' 填充产品名称到 C3
    If productName <> "" Then
        wsTemplate.Range("C3").Value = productName
    End If
    
    ' 填充日期
    Call FillDates
    
    ' 计算需求量（如果C4已填写）
    Call CalculateRequirements
    
    GoTo CleanUp
    
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "发生错误: " & Err.Description & vbCrLf & "错误编号: " & Err.Number, vbCritical, "错误"
    End If
    
CleanUp:
    ' 恢复设置
    Application.Calculation = xlCalculationAutomatic  ' 恢复自动计算
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 清空 BOM 数据区域
Sub ClearBOMArea()
    Dim wsTemplate As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wsTemplate = ThisWorkbook.Worksheets("模板")
    
    ' 清空 C3 产品名称
    wsTemplate.Range("C3").ClearContents
    
    ' 查找数据区域的实际最后一行（从A列查找）
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "A").End(xlUp).Row
    
    ' 如果有多于一行的数据（第6行之后还有数据）
    If lastRow > 6 Then
        ' 删除第7行到最后一行之间包含数据的行
        ' 但要保留备注和领料人行（通过检查内容判断）
        Dim deleteStartRow As Long
        Dim deleteEndRow As Long
        Dim foundRemark As Boolean
        
        deleteStartRow = 7
        deleteEndRow = 6
        foundRemark = False
        
        ' 从第7行开始查找，找到"备注"行为止
        For i = 7 To lastRow
            If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
                deleteEndRow = i - 1
                foundRemark = True
                Exit For
            End If
        Next i
        
        ' 如果找到了备注行，且在第7行之后，删除中间的数据行
        If foundRemark And deleteEndRow >= deleteStartRow Then
            wsTemplate.Rows(deleteStartRow & ":" & deleteEndRow).Delete Shift:=xlUp
        ElseIf Not foundRemark And lastRow > 6 Then
            ' 如果没找到备注行，但有数据，删除第7行到最后一行
            wsTemplate.Rows("7:" & lastRow).Delete Shift:=xlUp
        End If
    End If
    
    ' 清空第6行的数据内容（保留格式作为模板）
    ' A-I列：基本BOM数据，J列：批号，K列：备用，L列：报废，M列：入库，N列：抽检，O列：下次结存
    wsTemplate.Range("A6:O6").ClearContents
End Sub

' 自动填充领料日期和生产日期
Sub FillDates()
    Dim wsTemplate As Worksheet
    Dim pickupDate As Date
    Dim productionDate As Date
    
    Set wsTemplate = ThisWorkbook.Worksheets("模板")
    
    ' 获取当前日期作为领料日期
    pickupDate = Date
    
    ' 生产日期 = 领料日期 + 1天
    productionDate = pickupDate + 1
    
    ' 格式化并填充 I4（领料日期）- 格式：2026.02.03
    wsTemplate.Range("I4").Value = Format(pickupDate, "yyyy.mm.dd")
    
    ' 格式化并填充 E4（生产日期）- 格式：2026.02.03
    wsTemplate.Range("E4").Value = Format(productionDate, "yyyy.mm.dd")
End Sub

' 自动计算需求量（C4 / 配方数字）
Sub CalculateRequirements()
    Dim wsTemplate As Worksheet
    Dim wsData As Worksheet
    Dim productCode As String
    Dim finishedProductQty As Double
    Dim lastRow As Long
    Dim i As Long
    
    ' 关闭屏幕更新和事件，提高性能
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Set wsTemplate = ThisWorkbook.Worksheets("模板")
    
    ' 获取成品需求量（C4）
    finishedProductQty = Val(wsTemplate.Range("C4").Value)
    
    ' 如果成品需求量为空或0，清空需求量列并退出
    If finishedProductQty = 0 Then
        lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 6 Then
            ' 查找最后一行数据（备注行之前）
            For i = lastRow To 6 Step -1
                If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
                    lastRow = i - 1
                    Exit For
                End If
            Next i
            
            ' 清空需求量列（E列）
            If lastRow >= 6 Then
                wsTemplate.Range("E6:E" & lastRow).ClearContents
            End If
        End If
        GoTo CleanUp
    End If
    
    ' 获取当前产品编号
    productCode = Trim(wsTemplate.Range("E3").Value)
    
    ' 如果没有产品编号，退出
    If productCode = "" Then
        GoTo CleanUp
    End If
    
    ' 如果配方数据不存在或产品编号变了，重新从BOM表读取
    If Not IsArray(formulaData) Or currentProductCode <> productCode Then
        Set wsData = ThisWorkbook.Worksheets("BOM")
        
        ' 获取配方列索引
        Dim colProductCode As Long
        Dim colFormula As Long
        Dim dataRow As Long
        Dim bomCount As Integer
        
        colProductCode = GetColumnIndex(wsData, 1, "产品编号")
        colFormula = GetColumnIndex(wsData, 1, "配方")
        
        If colProductCode = 0 Then
            GoTo CleanUp
        End If
        
        ' 查找匹配的产品并收集配方数据
        lastRow = wsData.Cells(wsData.Rows.Count, colProductCode).End(xlUp).Row
        bomCount = 0
        
        ' 先计算数量
        For dataRow = 2 To lastRow
            If Trim(wsData.Cells(dataRow, colProductCode).Value) = productCode Then
                bomCount = bomCount + 1
            End If
        Next dataRow

        ' 重新分配数组
        If bomCount > 0 Then
            ReDim formulaData(1 To bomCount)
        Else
            ReDim formulaData(1 To 1)  ' 创建占位符数组，避免空数组
            GoTo SkipFormulaFill
        End If
        
        ' 填充配方数据
        i = 0
        For dataRow = 2 To lastRow
            If Trim(wsData.Cells(dataRow, colProductCode).Value) = productCode Then
                i = i + 1
                If colFormula > 0 Then
                    formulaData(i) = Val(wsData.Cells(dataRow, colFormula).Value)
                Else
                    formulaData(i) = 1  ' 默认值
                End If
            End If
        Next dataRow

SkipFormulaFill:
        currentProductCode = productCode
    End If
    
    ' 计算并填入需求量
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "A").End(xlUp).Row
    
    ' 查找最后一行数据（备注行之前）
    For i = lastRow To 6 Step -1
        If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    ' 如果有数据行，计算需求量
    If lastRow >= 6 Then
        Dim rowCount As Long
        rowCount = lastRow - 5  ' 从第6行开始

        ' 确保数组已初始化且有效
        If IsArray(formulaData) Then
            Dim dataUpperBound As Long
            dataUpperBound = UBound(formulaData)

            ' 确保不会超出配方数据的范围
            If rowCount > dataUpperBound Then
                rowCount = dataUpperBound
            End If

            ' 只有当rowCount大于0时才执行循环
            If rowCount > 0 Then
                ' 批量计算并填入需求量
                For i = 1 To rowCount
                    Dim formulaValue As Double
                    Dim requirementQty As Double
                    formulaValue = Val(formulaData(i))

                    ' 避免除零错误
                    If formulaValue > 0 Then
                        requirementQty = finishedProductQty / formulaValue
                        wsTemplate.Cells(5 + i, 5).Value = requirementQty  ' E列：需求量
                    Else
                        requirementQty = finishedProductQty
                        wsTemplate.Cells(5 + i, 5).Value = requirementQty  ' E列：需求量
                    End If
                    
                    ' ?? M列（入库）默认等于E列（需求量）
                    wsTemplate.Cells(5 + i, 13).Value = requirementQty  ' M列：入库
                Next i
            End If
        End If
    End If
    
    GoTo CleanUp
    
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "计算需求量时发生错误: " & Err.Description & vbCrLf & "错误编号: " & Err.Number, vbCritical, "错误"
    End If
    
CleanUp:
    ' 恢复设置
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ========== 批号计算功能 ==========

' CommandButton1 点击事件：计算批号
Private Sub CommandButton1_Click()
    On Error GoTo ErrorHandler

    Dim wsTemplate As Worksheet
    Dim wsInbound As Worksheet
    Dim wsOutbound As Worksheet

    ' 关闭屏幕更新和事件，提高性能
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsTemplate = ThisWorkbook.Worksheets("模板")
    Set wsInbound = ThisWorkbook.Worksheets("入库")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")

    ' 验证成品需求量是否填写
    If IsEmpty(wsTemplate.Range("C4")) Or Val(wsTemplate.Range("C4").Value) <= 0 Then
        MsgBox "请先填写成品需求量（C4单元格）", vbExclamation, "提示"
        GoTo CleanUp
    End If

    ' 验证BOM数据是否已加载（通过检查B列）
    Dim lastRow As Long
    Dim i As Long
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row

    Dim hasData As Boolean
    hasData = False
    For i = 6 To lastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then
            hasData = True
            
            ' 优化：如果车间结存量（G列）为空，默认为0
            If IsEmpty(wsTemplate.Cells(i, "G")) Or Trim(wsTemplate.Cells(i, "G").Value) = "" Then
                wsTemplate.Cells(i, "G").Value = 0
            End If
        End If
    Next i

    If Not hasData Then
        MsgBox "请先填写产品编号（E3单元格）以加载BOM数据", vbExclamation, "提示"
        GoTo CleanUp
    End If

    ' 检查 mustDivideData 数组是否已正确初始化
    Dim arrayOK As Boolean
    Dim arraySize As Long
    Dim dataRowCount As Long
    arrayOK = False
    arraySize = 0

    ' 计算实际数据行数
    dataRowCount = 0
    For i = 6 To lastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then
            dataRowCount = dataRowCount + 1
        End If
    Next i

    ' 检查数组是否已初始化
    On Error Resume Next
    If IsArray(mustDivideData) Then
        arraySize = UBound(mustDivideData)
        If Err.Number = 0 And arraySize >= dataRowCount And arraySize > 0 Then
            arrayOK = True
        End If
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    ' 如果数组未初始化或大小不匹配，重新加载数组（不覆盖模板数据）
    If Not arrayOK Then
        Call ReloadBOMArrays
    End If

    ' 步骤0：刷新实时库存（基于出库记录重新计算）
    Call RefreshAllInventory

    ' 步骤1：计算本次领用量（H列）
    Call CalculatePickupQuantity(wsTemplate)

    ' 步骤2：分配批号并生成出库记录
    Call AllocateBatchNumbers(wsTemplate, wsInbound, wsOutbound)

    ' 步骤3：填写J列批号显示
    Call FillBatchNumberDisplay(wsTemplate, wsOutbound)

    ' 步骤4：计算下次结存（O列）
    Call CalculateNextBatchStock

    ' 步骤5：保存生产记录
    Call SaveProductionRecords

    ' 批号计算完成
    GoTo CleanUp

ErrorHandler:
    MsgBox "发生错误：" & Err.Description & vbCrLf & "错误编号：" & Err.Number, vbCritical, "错误"

CleanUp:
    ' 恢复设置
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 重新加载BOM数组（不覆盖模板表数据）
' 用于在点击"计算批号"时，如果数组未初始化，只重新加载数组而不清空用户填写的数据
Sub ReloadBOMArrays()
    On Error GoTo ErrorHandler

    Dim wsTemplate As Worksheet
    Dim wsData As Worksheet
    Dim productCode As String
    Dim lastRow As Long
    Dim dataRow As Long
    Dim i As Long
    Dim bomCount As Long

    Dim colProductCode As Long
    Dim colMaterialCode As Long
    Dim colFormula As Long
    Dim colMustDivide As Long

    Set wsTemplate = ThisWorkbook.Worksheets("模板")
    Set wsData = ThisWorkbook.Worksheets("BOM")

    ' 获取产品编号
    productCode = Trim(wsTemplate.Range("E3").Value)

    If productCode = "" Then
        Exit Sub
    End If

    ' 获取BOM表的列索引
    colProductCode = GetColumnIndex(wsData, 1, "产品编号")
    colMaterialCode = GetColumnIndex(wsData, 1, "物料编号")
    colFormula = GetColumnIndex(wsData, 1, "配方")
    colMustDivide = GetColumnIndex(wsData, 1, "整除")

    If colProductCode = 0 Or colMaterialCode = 0 Then
        Exit Sub
    End If

    ' 计算BOM数据行数
    bomCount = 0
    lastRow = wsData.Cells(wsData.Rows.Count, colProductCode).End(xlUp).Row

    For dataRow = 2 To lastRow
        If Trim(wsData.Cells(dataRow, colProductCode).Value) = productCode Then
            bomCount = bomCount + 1
        End If
    Next dataRow

    If bomCount = 0 Then
        Exit Sub
    End If

    ' 重新初始化数组
    ReDim formulaData(1 To bomCount)
    ReDim mustDivideData(1 To bomCount)

    ' 填充数组
    i = 0
    For dataRow = 2 To lastRow
        If Trim(wsData.Cells(dataRow, colProductCode).Value) = productCode Then
            i = i + 1

            ' 保存配方数字
            If colFormula > 0 Then
                formulaData(i) = wsData.Cells(dataRow, colFormula).Value
            Else
                formulaData(i) = 1
            End If

            ' 保存整除标志
            If colMustDivide > 0 Then
                mustDivideData(i) = Trim(wsData.Cells(dataRow, colMustDivide).Value)
            Else
                mustDivideData(i) = "Y"
            End If
        End If
    Next dataRow

    ' 调试日志
    If DEBUG_LOG Then
        DebugLog "ReloadArrays", "ProductCode=" & productCode & ", BOMCount=" & bomCount & ", ArraySize=" & UBound(mustDivideData)
        For i = 1 To bomCount
            Dim materialCode As String
            materialCode = Trim(wsTemplate.Cells(5 + i, "B").Value)
            DebugLog "ReloadArrays_Data", "Index=" & i & ", Material=" & materialCode & ", MustDivide=" & mustDivideData(i)
        Next i
    End If

    Exit Sub

ErrorHandler:
    If DEBUG_LOG Then
        DebugLog "ReloadArrays_Error", "Error: " & Err.Description
    End If
End Sub



' 计算本次领用量（H列）
' 公式：X × 规格数量 + 车间结存量 >= 需求量，X为最小正整数
Sub CalculatePickupQuantity(wsTemplate As Worksheet)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim i As Long
    Dim spec As String
    Dim specQty As Double
    Dim requirement As Double
    Dim workshopStock As Double
    Dim needOutbound As Double
    Dim X As Long
    Dim pickupQty As Double

    ' 查找最后一行
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row

    ' 查找备注行，排除备注行之后的内容
    For i = lastRow To 6 Step -1
        If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i

    ' 遍历每一行计算本次领用量
    For i = 6 To lastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then  ' B列：物料编号
            spec = Trim(wsTemplate.Cells(i, "D").Value)  ' D列：规格
            requirement = Val(wsTemplate.Cells(i, "E").Value)  ' E列：需求量
            workshopStock = Val(wsTemplate.Cells(i, "G").Value)  ' G列：车间结存量

            ' 从规格中提取数字
            specQty = ExtractSpecQuantity(spec)

            ' 获取该物料的整除标志（数组索引 = 行号 - 5，因为数据从第6行开始，数组从1开始）
            Dim mustDivide As String
            mustDivide = "Y"  ' 默认为Y

            ' 检查数组是否已初始化
            On Error Resume Next
            If IsArray(mustDivideData) Then
                Dim arrayIndex As Long
                arrayIndex = i - 5

                ' 更安全的边界检查
                If Err.Number = 0 Then
                    If arrayIndex >= LBound(mustDivideData) And arrayIndex <= UBound(mustDivideData) Then
                        mustDivide = Trim(mustDivideData(arrayIndex))
                        If DEBUG_LOG Then
                            DebugLog "PickupQty_Array", "Row=" & i & ", Index=" & arrayIndex & ", ArraySize=" & UBound(mustDivideData) & ", MustDivide=" & mustDivide
                        End If
                    Else
                        ' 数组越界
                        If DEBUG_LOG Then
                            DebugLog "PickupQty_Error", "Row=" & i & ", Index=" & arrayIndex & " out of bounds [" & LBound(mustDivideData) & " to " & UBound(mustDivideData) & "], Using default mustDivide=Y"
                        End If
                        mustDivide = "Y"
                    End If
                End If
            End If
            If Err.Number <> 0 Then
                ' 如果出错，记录日志并使用默认值
                If DEBUG_LOG Then
                    DebugLog "PickupQty_Error", "Row=" & i & ", Err=" & Err.Description & ", Using default mustDivide=Y"
                End If
                mustDivide = "Y"
            End If
            On Error GoTo ErrorHandler

            ' 新的计算逻辑：根据整除标志采用不同的计算方法
            If specQty > 0 Then
                If mustDivide = "Y" Then
                    ' ==================== 整除=Y 的逻辑 ====================
                    ' 步骤1：计算净需求量 = 需求量 - 车间结存
                    Dim netReq As Double
                    netReq = requirement - workshopStock

                    ' 确保净需求量不为负数
                    If netReq < 0 Then netReq = 0

                    ' 步骤2：将净需求量向上取整到规格的整数倍
                    Dim needUnits As Long
                    needUnits = -Int(-netReq / specQty)
                    pickupQty = needUnits * specQty

                    ' 步骤3：如果规格单位数的个位数是9，+1凑成整十
                    ' 例如：319捆 → 320捆，19板 → 20板
                    If needUnits Mod 10 = 9 Then
                        needUnits = needUnits + 1
                        pickupQty = needUnits * specQty
                    End If

                    needOutbound = pickupQty
                Else
                    ' ==================== 整除=N 的逻辑 ====================
                    ' 按(需求量-车间结存量)/规格 向上取整，再乘规格
                    ' 领用量=各批次出库量之和
                    Dim netReqN As Double
                    netReqN = requirement - workshopStock
                    If netReqN < 0 Then netReqN = 0

                    If specQty > 0 Then
                        Dim needUnitsN As Long
                        needUnitsN = -Int(-netReqN / specQty)
                        pickupQty = needUnitsN * specQty
                    Else
                        pickupQty = netReqN
                    End If

                    needOutbound = pickupQty
                End If
            Else
                pickupQty = 0
                needOutbound = 0
            End If

            ' 填写到H列
            wsTemplate.Cells(i, "H").Value = pickupQty

            If DEBUG_LOG Then
                DebugLog "PickupQty", _
                         "Row=" & i & _
                         ", Code=" & Trim(wsTemplate.Cells(i, "B").Value) & _
                         ", Spec=" & spec & _
                         ", SpecQty=" & specQty & _
                         ", Req=" & requirement & _
                         ", Stock=" & workshopStock & _
                         ", MustDivide=" & mustDivide & _
                         ", Need=" & needOutbound & _
                         ", X=" & X & _
                         ", Pickup=" & pickupQty
            End If
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "计算本次领用量时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' 分配批号并生成出库记录（FIFO先进先出）
Sub AllocateBatchNumbers(wsTemplate As Worksheet, wsInbound As Worksheet, wsOutbound As Worksheet)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim i As Long
    Dim materialCode As String
    Dim pickupQty As Double
    Dim workshopStock As Double
    Dim needOutbound As Double
    Dim spec As String
    Dim specQty As Double

    ' 查找模板表最后一行
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row

    ' 查找备注行
    For i = lastRow To 6 Step -1
        If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i

    ' 遍历每一行物料
    For i = 6 To lastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then
            materialCode = Trim(wsTemplate.Cells(i, "B").Value)  ' B列：物料编号
            pickupQty = Val(wsTemplate.Cells(i, "H").Value)  ' H列：本次领用量
            workshopStock = Val(wsTemplate.Cells(i, "G").Value)  ' G列：车间结存量
            spec = Trim(wsTemplate.Cells(i, "D").Value)  ' D列：规格

            ' 计算需要出库的数量
            ' pickupQty 已经是按(需求量-结存量)向上取整后的结果
            ' 这里不要再扣一次结存量，否则会少出库
            needOutbound = pickupQty

            ' 获取该物料的整除标志
            Dim mustDivide As String
            mustDivide = "Y"  ' 默认为Y

            ' 检查数组是否已初始化
            On Error Resume Next
            If IsArray(mustDivideData) Then
                Dim arrayIndex As Long
                arrayIndex = i - 5

                ' 更安全的边界检查
                If Err.Number = 0 Then
                    If arrayIndex >= LBound(mustDivideData) And arrayIndex <= UBound(mustDivideData) Then
                        mustDivide = Trim(mustDivideData(arrayIndex))
                        If DEBUG_LOG Then
                            DebugLog "Allocate_Array", "Row=" & i & ", Index=" & arrayIndex & ", ArraySize=" & UBound(mustDivideData) & ", MustDivide=" & mustDivide
                        End If
                    Else
                        ' 数组越界
                        If DEBUG_LOG Then
                            DebugLog "Allocate_Error", "Row=" & i & ", Index=" & arrayIndex & " out of bounds [" & LBound(mustDivideData) & " to " & UBound(mustDivideData) & "], Using default mustDivide=Y"
                        End If
                        mustDivide = "Y"
                    End If
                End If
            End If
            If Err.Number <> 0 Then
                ' 如果出错，记录日志并使用默认值
                If DEBUG_LOG Then
                    DebugLog "Allocate_Error", "Row=" & i & ", Err=" & Err.Description & ", Using default mustDivide=Y"
                End If
                mustDivide = "Y"
            End If
            On Error GoTo ErrorHandler

            If needOutbound > 0 Then
                ' 提取规格数量
                specQty = ExtractSpecQuantity(spec)

                ' 执行FIFO分配，传入整除标志
                Call AllocateBatchesFIFO(materialCode, needOutbound, specQty, _
                                        wsTemplate, wsInbound, wsOutbound, i, mustDivide)
            End If
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "分配批号时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' FIFO批次分配核心函数
Sub AllocateBatchesFIFO(materialCode As String, needOutbound As Double, specQty As Double, _
                        wsTemplate As Worksheet, wsInbound As Worksheet, wsOutbound As Worksheet, _
                        templateRow As Long, mustDivide As String)
    On Error GoTo ErrorHandler

    Dim inboundLastRow As Long
    Dim outboundLastRow As Long
    Dim i As Long
    Dim remainingNeed As Double
    Dim currentBatchStock As Double
    Dim thisOutbound As Double
    Dim batchNumber As String
    Dim materialName As String
    Dim manufacturer As String
    Dim unit As String
    Dim auxUnit As String
    Dim inboundQty As Double
    Dim alreadyOutbound As Double
    Dim currentDate As Date

    ' 获取入库表列索引
    Dim colInDate As Long
    Dim colInMaterialCode As Long
    Dim colInMaterialName As Long
    Dim colInManufacturer As Long
    Dim colInUnit As Long
    Dim colInAuxUnit As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    Dim colInAuxQty As Long
    Dim colInAlreadyOut As Long
    Dim colInStock As Long

    colInDate = GetColumnIndex(wsInbound, 1, "日期")
    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInMaterialName = GetColumnIndex(wsInbound, 1, "物料名称")
    colInManufacturer = GetColumnIndex(wsInbound, 1, "生产厂家")
    colInUnit = GetColumnIndex(wsInbound, 1, "单位")
    colInAuxUnit = GetColumnIndex(wsInbound, 1, "辅单位")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")
    colInAuxQty = GetColumnIndex(wsInbound, 1, "辅数量")
    colInAlreadyOut = GetColumnIndex(wsInbound, 1, "已出库")
    colInStock = GetColumnIndex(wsInbound, 1, "实时库存")

    ' 验证必需列是否存在
    If colInMaterialCode = 0 Or colInBatch = 0 Or colInStock = 0 Then
        MsgBox "入库表缺少必需的列（物料编号、批次或实时库存），请检查表头！", vbCritical, "错误"
        Exit Sub
    End If

    remainingNeed = needOutbound
    currentDate = Date
    inboundLastRow = wsInbound.Cells(wsInbound.Rows.Count, colInMaterialCode).End(xlUp).Row

    ' 遍历入库表，按FIFO原则分配（假设入库表已按日期排序）
    For i = 2 To inboundLastRow  ' 第1行是表头
        ' 获取当前行的物料编号和批次
        Dim currentMaterialCode As String
        Dim currentBatch As String
        currentMaterialCode = Trim(wsInbound.Cells(i, colInMaterialCode).Value)
        currentBatch = Trim(wsInbound.Cells(i, colInBatch).Value)

        ' 调试：记录遍历的所有行
        If DEBUG_LOG Then
            If currentMaterialCode <> "" Then
                DebugLog "FIFO_Scan", _
                         "Row=" & i & _
                         ", Material=" & currentMaterialCode & _
                         ", Batch=" & currentBatch & _
                         ", Stock=" & Val(wsInbound.Cells(i, colInStock).Value) & _
                         ", Match=" & IIf(currentMaterialCode = materialCode, "YES", "NO")
            End If
        End If

        ' 检查是否是目标物料且有库存
        If currentMaterialCode = materialCode Then
            currentBatchStock = Val(wsInbound.Cells(i, colInStock).Value)

            If currentBatchStock > 0 And remainingNeed > 0 Then
                ' 确定本次从该批次出库的数量
                If specQty > 0 Then
                    If mustDivide = "Y" Then
                        ' ==================== 整除=Y 的逻辑 ====================
                        ' 第二重逻辑：如果计算出的板数是19、29、39等，+1凑成整十
                        Dim needUnits As Double
                        needUnits = remainingNeed / specQty

                        ' 向上取整
                        Dim needUnitsInt As Long
                        needUnitsInt = -Int(-needUnits)

                        ' 如果个位数是9（19、29、39...），则+1
                        If needUnitsInt Mod 10 = 9 Then
                            needUnitsInt = needUnitsInt + 1
                        End If

                        ' 计算实际需要量
                        Dim actualNeed As Double
                        actualNeed = needUnitsInt * specQty

                        ' 检查库存是否足够
                        If currentBatchStock >= actualNeed Then
                            thisOutbound = actualNeed
                        ElseIf currentBatchStock >= remainingNeed Then
                            ' 库存不够完整的规格单位，但够满足剩余需求，使用剩余需求
                            thisOutbound = remainingNeed
                        Else
                            ' 库存连剩余需求都不够，使用全部库存
                            thisOutbound = currentBatchStock
                        End If
                    Else
                        ' ==================== 整除=N 的逻辑 ====================
                        ' 批次之和整除即可，单个批次可以不整除
                        ' 例如：需要84000（7箱），批次1出47445，批次2出36555

                        ' 直接按库存和剩余需求分配，不做向上取整
                        If currentBatchStock >= remainingNeed Then
                            ' 库存足够，取出剩余需求
                            thisOutbound = remainingNeed
                        Else
                            ' 库存不够，取出全部库存
                            thisOutbound = currentBatchStock
                        End If
                    End If
                Else
                    ' 没有规格限制，按原逻辑
                    If currentBatchStock >= remainingNeed Then
                        thisOutbound = remainingNeed
                    Else
                        thisOutbound = currentBatchStock
                    End If
                End If

                ' 如果计算后的出库量为0，跳过
                If thisOutbound <= 0 Then
                    GoTo NextBatch
                End If

                ' 获取批次信息
                batchNumber = Trim(wsInbound.Cells(i, colInBatch).Value)
                If colInMaterialName > 0 Then materialName = wsInbound.Cells(i, colInMaterialName).Value
                If colInManufacturer > 0 Then manufacturer = wsInbound.Cells(i, colInManufacturer).Value
                If colInUnit > 0 Then unit = wsInbound.Cells(i, colInUnit).Value
                If colInAuxUnit > 0 Then auxUnit = wsInbound.Cells(i, colInAuxUnit).Value

                ' 生成出库记录
                outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, 1).End(xlUp).Row
                If outboundLastRow = 1 And IsEmpty(wsOutbound.Cells(1, 1)) Then
                    outboundLastRow = 1  ' 如果是空表，从第1行开始（假设有表头）
                Else
                    outboundLastRow = outboundLastRow + 1
                End If

                ' 获取出库表列索引
                Dim colOutDate As Long
                Dim colOutMaterialCode As Long
                Dim colOutMaterialName As Long
                Dim colOutManufacturer As Long
                Dim colOutUnit As Long
                Dim colOutAuxUnit As Long
                Dim colOutBatch As Long
                Dim colOutQty As Long
                Dim colOutAuxQty As Long
                Dim colOutSpec As Long
                Dim colOutStock As Long
                Dim colOutProductionBatch As Long

                colOutDate = GetColumnIndex(wsOutbound, 1, "日期")
                colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
                colOutMaterialName = GetColumnIndex(wsOutbound, 1, "物料名称")
                colOutSpec = GetColumnIndex(wsOutbound, 1, "规格")
                colOutManufacturer = GetColumnIndex(wsOutbound, 1, "生产厂家")
                colOutUnit = GetColumnIndex(wsOutbound, 1, "单位")
                colOutAuxUnit = GetColumnIndex(wsOutbound, 1, "辅单位")
                colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
                colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
                colOutAuxQty = GetColumnIndex(wsOutbound, 1, "辅数量")
                colOutStock = GetColumnIndex(wsOutbound, 1, "实时库存")
                colOutProductionBatch = GetColumnIndex(wsOutbound, 1, "生产批号")

                ' ?? 获取模板表I3的生产批号
                Dim productionBatchNumber As String
                productionBatchNumber = Trim(wsTemplate.Range("I3").Value)

                ' 写入出库记录
                If colOutDate > 0 Then wsOutbound.Cells(outboundLastRow, colOutDate).Value = currentDate
                If colOutMaterialCode > 0 Then wsOutbound.Cells(outboundLastRow, colOutMaterialCode).Value = materialCode
                If colOutMaterialName > 0 Then wsOutbound.Cells(outboundLastRow, colOutMaterialName).Value = materialName
                If colOutSpec > 0 Then wsOutbound.Cells(outboundLastRow, colOutSpec).Value = wsTemplate.Cells(templateRow, "D").Value
                If colOutManufacturer > 0 Then wsOutbound.Cells(outboundLastRow, colOutManufacturer).Value = manufacturer
                If colOutUnit > 0 Then wsOutbound.Cells(outboundLastRow, colOutUnit).Value = unit
                If colOutAuxUnit > 0 Then wsOutbound.Cells(outboundLastRow, colOutAuxUnit).Value = auxUnit
                If colOutBatch > 0 Then wsOutbound.Cells(outboundLastRow, colOutBatch).Value = batchNumber
                If colOutQty > 0 Then wsOutbound.Cells(outboundLastRow, colOutQty).Value = thisOutbound
                If colOutAuxQty > 0 And specQty > 0 Then wsOutbound.Cells(outboundLastRow, colOutAuxQty).Value = thisOutbound / specQty
                ' ?? 写入生产批号
                If colOutProductionBatch > 0 Then wsOutbound.Cells(outboundLastRow, colOutProductionBatch).Value = productionBatchNumber

                ' 更新入库表
                If colInAlreadyOut > 0 Then
                    alreadyOutbound = Val(wsInbound.Cells(i, colInAlreadyOut).Value)
                    wsInbound.Cells(i, colInAlreadyOut).Value = alreadyOutbound + thisOutbound
                End If

                wsInbound.Cells(i, colInStock).Value = currentBatchStock - thisOutbound

                ' 写入出库后的实时库存
                If colOutStock > 0 Then wsOutbound.Cells(outboundLastRow, colOutStock).Value = currentBatchStock - thisOutbound

                ' ?? 计算并写入车间使用量和车间实时结存
                Dim colOutWorkshopUsage As Long
                Dim colOutWorkshopStock As Long
                Dim requirement As Double
                Dim scrap As Double
                Dim inspection As Double
                Dim workshopUsage As Double
                Dim totalPickup As Double
                Dim workshopStockAfter As Double
                
                colOutWorkshopUsage = GetColumnIndex(wsOutbound, 1, "车间使用量")
                colOutWorkshopStock = GetColumnIndex(wsOutbound, 1, "实时结存")
                
                ' 获取模板表中该物料的入库、报废、抽检
                Dim inbound As Double
                inbound = Val(wsTemplate.Cells(templateRow, "M").Value)      ' M列：入库
                scrap = Val(wsTemplate.Cells(templateRow, "L").Value)        ' L列：报废
                inspection = Val(wsTemplate.Cells(templateRow, "N").Value)   ' N列：抽检
                
                ' 获取本次领用量总计
                totalPickup = Val(wsTemplate.Cells(templateRow, "H").Value)  ' H列：本次领用量
                
                ' 按出库比例分配车间使用量（使用入库量代替需求量）
                If totalPickup > 0 Then
                    workshopUsage = (thisOutbound / totalPickup) * (inbound + scrap + inspection)
                    workshopUsage = Round(workshopUsage, 2)  ' ?? 四舍五入到2位小数
                Else
                    workshopUsage = 0
                End If
                
                ' 写入车间使用量
                If colOutWorkshopUsage > 0 Then
                    wsOutbound.Cells(outboundLastRow, colOutWorkshopUsage).Value = workshopUsage
                End If
                
                ' 计算并写入出库后的车间实时结存
                ' 实时结存 = 当前车间结存 + 本次出库 - 本次使用
                Dim currentWorkshopStock As Double
                currentWorkshopStock = Val(wsTemplate.Cells(templateRow, "G").Value)  ' G列：车间结存量
                
                workshopStockAfter = currentWorkshopStock + thisOutbound - workshopUsage
                workshopStockAfter = Round(workshopStockAfter, 2)  ' ?? 四舍五入到2位小数
                
                If colOutWorkshopStock > 0 Then
                    wsOutbound.Cells(outboundLastRow, colOutWorkshopStock).Value = workshopStockAfter
                End If

                ' 更新剩余需求
                remainingNeed = remainingNeed - thisOutbound

                If DEBUG_LOG Then
                    DebugLog "FIFO", _
                             "Code=" & materialCode & _
                             ", InRow=" & i & _
                             ", Batch=" & batchNumber & _
                             ", Stock=" & currentBatchStock & _
                             ", NeedBefore=" & (remainingNeed + thisOutbound) & _
                             ", Out=" & thisOutbound & _
                             ", NeedAfter=" & remainingNeed & _
                             ", MustDivide=" & mustDivide
                End If

                If remainingNeed <= 0 Then Exit For
            End If
        End If
NextBatch:
    Next i

    ' 检查是否满足需求
    If remainingNeed > 0 Then
        MsgBox "物料编号 " & materialCode & " 库存不足，还缺少 " & Format(remainingNeed, "#,##0.00") & " 单位", _
               vbExclamation, "库存不足"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "FIFO分配批次时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' 调试日志写入到"调试"工作表
Private Sub DebugLog(tag As String, message As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("调试")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "调试"
        ws.Range("A1:C1").Value = Array("Time", "Tag", "Message")
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = tag
    ws.Cells(nextRow, 3).Value = message
End Sub

' ============================================
' 刷新所有物料的实时库存（本地版本）
' 基于出库记录重新计算入库表的实时库存
' ============================================
Private Sub RefreshAllInventory()
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
    Dim colInMaterialCode As Long
    Dim colInBatch As Long
    Dim colInQty As Long
    Dim colInAlreadyOut As Long
    Dim colInStock As Long

    colInMaterialCode = GetColumnIndex(wsInbound, 1, "物料编号")
    colInBatch = GetColumnIndex(wsInbound, 1, "批次")
    colInQty = GetColumnIndex(wsInbound, 1, "入库数量")
    colInAlreadyOut = GetColumnIndex(wsInbound, 1, "已出库")
    colInStock = GetColumnIndex(wsInbound, 1, "实时库存")

    ' 验证必需列
    If colInMaterialCode = 0 Or colInBatch = 0 Or colInQty = 0 Or colInStock = 0 Then
        DebugLog "RefreshStock", "入库表缺少必需的列（物料编号、批次、入库数量或实时库存）"
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
        DebugLog "RefreshStock", "出库表缺少必需的列（物料编号、批次或出库数量）"
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

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    DebugLog "RefreshStock", "Error: " & Err.Description
    Resume CleanUp
End Sub

' 填写H列（本次领用量）和J列（批号）显示（多批号用强制换行分隔）
Sub FillBatchNumberDisplay(wsTemplate As Worksheet, wsOutbound As Worksheet)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim materialCode As String
    Dim batchList As String
    Dim qtyList As String
    Dim outboundLastRow As Long
    Dim currentDate As Date

    ' 获取出库表列索引
    Dim colOutDate As Long
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long
    Dim colOutAuxQty As Long
    Dim colOutProductionBatch As Long  ' ?? 生产批号列
    Dim productionBatch As String      ' ?? 当前生产批号

    colOutDate = GetColumnIndex(wsOutbound, 1, "日期")
    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
    colOutAuxQty = GetColumnIndex(wsOutbound, 1, "辅数量")
    colOutProductionBatch = GetColumnIndex(wsOutbound, 1, "生产批号")  ' ?? 获取生产批号列索引

    If colOutMaterialCode = 0 Or colOutBatch = 0 Or colOutQty = 0 Then
        MsgBox "出库表缺少必需的列（物料编号、批次或出库数量），请检查表头！", vbCritical, "错误"
        Exit Sub
    End If

    currentDate = Date
    productionBatch = Trim(wsTemplate.Range("I3").Value)  ' ?? 读取I3单元格的生产批号
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row

    ' 查找备注行
    For i = lastRow To 6 Step -1
        If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i

    ' 遍历模板表的每一行
    For i = 6 To lastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then
            materialCode = Trim(wsTemplate.Cells(i, "B").Value)
            batchList = ""
            qtyList = ""

            ' 从出库表中查找当天该物料的所有批号和数量
            For j = 2 To outboundLastRow  ' 假设第1行是表头
                ' ?? 检查日期、物料编号和生产批号是否匹配
                If colOutDate > 0 Then
                    ' 增加生产批号筛选条件：只读取当前生产批号的出库记录
                    If wsOutbound.Cells(j, colOutDate).Value = currentDate And _
                       Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = materialCode And _
                       (colOutProductionBatch = 0 Or productionBatch = "" Or _
                        Trim(wsOutbound.Cells(j, colOutProductionBatch).Value) = productionBatch) Then

                        ' 获取批号
                        Dim batchInfo As String
                        batchInfo = Trim(wsOutbound.Cells(j, colOutBatch).Value)

                        ' 获取出库数量
                        Dim outQty As Double
                        outQty = Val(wsOutbound.Cells(j, colOutQty).Value)

                        ' 组合批号列表（J列）
                        If batchList = "" Then
                            batchList = batchInfo
                        Else
                            ' 使用分隔符 "--------"
                            batchList = batchList & vbLf & "--------" & vbLf & batchInfo
                        End If

                        ' 组合数量列表（H列）
                        If qtyList = "" Then
                            qtyList = CStr(CLng(outQty))
                        Else
                            ' 使用分隔符 "--------"
                            qtyList = qtyList & vbLf & "--------" & vbLf & CStr(CLng(outQty))
                        End If
                    End If
                Else
                    ' ?? 如果没有日期列，匹配物料编号和生产批号
                    If Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = materialCode And _
                       (colOutProductionBatch = 0 Or productionBatch = "" Or _
                        Trim(wsOutbound.Cells(j, colOutProductionBatch).Value) = productionBatch) Then
                        Dim batchInfo2 As String
                        batchInfo2 = Trim(wsOutbound.Cells(j, colOutBatch).Value)

                        Dim outQty2 As Double
                        outQty2 = Val(wsOutbound.Cells(j, colOutQty).Value)

                        If batchList = "" Then
                            batchList = batchInfo2
                        Else
                            ' 使用分隔符 "--------"
                            batchList = batchList & vbLf & "--------" & vbLf & batchInfo2
                        End If

                        If qtyList = "" Then
                            qtyList = Format(outQty2, "#,##0")
                        Else
                            ' 使用分隔符 "--------"
                            qtyList = qtyList & vbLf & "--------" & vbLf & Format(outQty2, "#,##0")
                        End If
                    End If
                End If
            Next j

            ' 填写到H列（本次领用量 - 分批号显示）
            If qtyList <> "" Then
                wsTemplate.Cells(i, "H").Value = qtyList
                wsTemplate.Cells(i, "H").WrapText = True  ' 启用自动换行
            End If

            ' 填写到J列（批号）
            If batchList <> "" Then
                wsTemplate.Cells(i, "J").Value = batchList
                wsTemplate.Cells(i, "J").WrapText = True  ' 启用自动换行
            End If
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "填写批号显示时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' ============================================
' ?? 计算下批结存（O列）
' 从车间结存表读取实时结存
' 创建日期：2026-02-11
' 修改日期：2026-02-11 - 简化逻辑，直接读取车间结存表
'          2026-02-11 - 修改填写位置为O列
' ============================================
Sub CalculateNextBatchStock()
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim materialCode As String
    Dim realStock As Double
    
    Application.EnableEvents = False
    
    lastRow = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
    
    ' 查找备注行
    For i = lastRow To 6 Step -1
        If InStr(1, Me.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    ' 遍历每一行，从车间结存表获取实时结存
    For i = 6 To lastRow
        If Not IsEmpty(Me.Cells(i, "B")) Then
            materialCode = Trim(Me.Cells(i, "B").Value)
            
            ' ?? 直接调用模块1的GetWorkshopStock函数（该函数会动态计算实时结存）
            realStock = GetWorkshopStock(materialCode)
            
            ' 填写到O列（下次结存）
            Me.Cells(i, "O").Value = realStock
        End If
    Next i
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "计算下批结存时发生错误: " & Err.Description, vbCritical
End Sub

' ============================================
' ?? 从模板表批量更新车间结存表
' 创建日期：2026-02-11
' ============================================
Sub UpdateWorkshopStockFromTemplate()
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim materialCode As String
    Dim nextStock As Double
    Dim updateCount As Long
    
    Application.ScreenUpdating = False
    
    lastRow = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
    updateCount = 0
    
    ' 查找备注行
    For i = lastRow To 6 Step -1
        If InStr(1, Me.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    ' 遍历模板表每一行
    For i = 6 To lastRow
        If Not IsEmpty(Me.Cells(i, "B")) Then
            materialCode = Trim(Me.Cells(i, "B").Value)
            nextStock = Val(Me.Cells(i, "N").Value)
            
            ' 更新车间结存表
            Call UpdateWorkshopStock(materialCode, nextStock)
            updateCount = updateCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    ' 车间结存已更新
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "更新车间结存时发生错误: " & Err.Description, vbCritical
End Sub

' ============================================
' ?? 反查历史生产批号
' 当R3单元格修改时触发
' 从生产记录表和生产记录明细表读取数据并填充模板表
' 创建日期：2026-02-11
' ============================================
Sub QueryProductionHistory()
    On Error GoTo ErrorHandler
    
    Dim wsTemplate As Worksheet
    Dim wsProduction As Worksheet
    Dim wsProductionDetail As Worksheet
    Dim wsOutbound As Worksheet
    Dim productionBatch As String
    Dim productCode As String
    Dim productName As String
    Dim requirementQty As Double
    Dim productionDate As Date
    Dim pickupDate As Date
    Dim i As Long
    Dim currentRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set wsTemplate = Me
    Set wsProduction = ThisWorkbook.Worksheets("生产记录")
    Set wsProductionDetail = ThisWorkbook.Worksheets("生产记录明细")
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 读取生产批号
    productionBatch = Trim(wsTemplate.Range("R3").Value)
    
    ' 如果生产批号为空，不执行反查
    If productionBatch = "" Then
        Exit Sub
    End If
    
    ' 步骤1：从生产记录表查找主记录
    Dim colProDate As Long
    Dim colProProductCode As Long
    Dim colProProductName As Long
    Dim colProProductionBatch As Long
    Dim colProRequirementQty As Long
    Dim colProPickupDate As Long
    Dim colProProductionDate As Long
    Dim lastRow As Long
    Dim found As Boolean
    
    colProDate = GetColumnIndex(wsProduction, 1, "日期")
    colProProductCode = GetColumnIndex(wsProduction, 1, "产品编号")
    colProProductName = GetColumnIndex(wsProduction, 1, "产品名称")
    colProProductionBatch = GetColumnIndex(wsProduction, 1, "生产批号")
    colProRequirementQty = GetColumnIndex(wsProduction, 1, "需求数量")
    
    ' 尝试获取领料日期和生产日期列（如果存在）
    On Error Resume Next
    colProPickupDate = GetColumnIndex(wsProduction, 1, "领料日期")
    colProProductionDate = GetColumnIndex(wsProduction, 1, "生产日期")
    On Error GoTo ErrorHandler
    
    If colProProductCode = 0 Or colProProductionBatch = 0 Then
        MsgBox "生产记录表缺少必需的列", vbCritical
        GoTo CleanUp
    End If
    
    lastRow = wsProduction.Cells(wsProduction.Rows.Count, colProProductCode).End(xlUp).Row
    found = False
    
    ' 查找匹配的生产批号
    For i = 2 To lastRow
        If Trim(wsProduction.Cells(i, colProProductionBatch).Value) = productionBatch Then
            ' 找到了，读取数据
            productCode = Trim(wsProduction.Cells(i, colProProductCode).Value)
            
            If colProProductName > 0 Then
                productName = Trim(wsProduction.Cells(i, colProProductName).Value)
            End If
            
            If colProRequirementQty > 0 Then
                requirementQty = Val(wsProduction.Cells(i, colProRequirementQty).Value)
            End If
            
            ' 读取日期
            If colProPickupDate > 0 And Not IsEmpty(wsProduction.Cells(i, colProPickupDate)) Then
                pickupDate = CDate(wsProduction.Cells(i, colProPickupDate).Value)
            Else
                pickupDate = Date
            End If
            
            If colProProductionDate > 0 And Not IsEmpty(wsProduction.Cells(i, colProProductionDate)) Then
                productionDate = CDate(wsProduction.Cells(i, colProProductionDate).Value)
            Else
                productionDate = Date + 1
            End If
            
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "未找到生产批号：" & productionBatch, vbExclamation
        GoTo CleanUp
    End If
    
    ' 步骤2：填写模板表主信息
    wsTemplate.Range("E3").Value = productCode           ' 产品编号
    wsTemplate.Range("C3").Value = productName           ' 产品名称
    wsTemplate.Range("C4").Value = requirementQty        ' 成品需求量
    wsTemplate.Range("I4").Value = Format(pickupDate, "yyyy.mm.dd")      ' 领料日期
    wsTemplate.Range("E4").Value = Format(productionDate, "yyyy.mm.dd")  ' 生产日期
    
    ' 步骤3：清空BOM数据区域
    Call ClearBOMArea
    
    ' 步骤4：从生产记录明细表读取物料明细
    Dim colDetailProductionBatch As Long
    Dim colDetailMaterialCode As Long
    Dim colDetailMaterialName As Long
    Dim colDetailSpec As Long
    Dim colDetailRequirement As Long
    Dim colDetailPickup As Long
    Dim colDetailScrap As Long
    Dim colDetailInspection As Long
    Dim colDetailInbound As Long
    Dim colDetailWorkshopStock As Long
    Dim detailLastRow As Long
    
    colDetailProductionBatch = GetColumnIndex(wsProductionDetail, 1, "生产批号")
    colDetailMaterialCode = GetColumnIndex(wsProductionDetail, 1, "物料编号")
    colDetailMaterialName = GetColumnIndex(wsProductionDetail, 1, "物料名称")
    colDetailSpec = GetColumnIndex(wsProductionDetail, 1, "规格")
    colDetailRequirement = GetColumnIndex(wsProductionDetail, 1, "需求量")
    colDetailPickup = GetColumnIndex(wsProductionDetail, 1, "本次领用量")
    colDetailScrap = GetColumnIndex(wsProductionDetail, 1, "报废")
    colDetailInspection = GetColumnIndex(wsProductionDetail, 1, "抽检")
    colDetailInbound = GetColumnIndex(wsProductionDetail, 1, "入库")
    
    ' 尝试获取车间结存量列（如果存在）
    On Error Resume Next
    colDetailWorkshopStock = GetColumnIndex(wsProductionDetail, 1, "车间结存量")
    On Error GoTo ErrorHandler
    
    If colDetailProductionBatch = 0 Or colDetailMaterialCode = 0 Then
        MsgBox "生产记录明细表缺少必需的列", vbCritical
        GoTo CleanUp
    End If
    
    detailLastRow = wsProductionDetail.Cells(wsProductionDetail.Rows.Count, colDetailProductionBatch).End(xlUp).Row
    currentRow = 6  ' 从第6行开始填写
    
    ' 遍历生产记录明细表，查找匹配的生产批号
    ' ?? 增加逻辑：确保每个物料编号在模板中只占一行
    Dim processedMaterials As Object
    Set processedMaterials = CreateObject("Scripting.Dictionary")
    
    For i = 2 To detailLastRow
        If Trim(wsProductionDetail.Cells(i, colDetailProductionBatch).Value) = productionBatch Then
            Dim materialCode As String
            materialCode = Trim(wsProductionDetail.Cells(i, colDetailMaterialCode).Value)
            
            ' 如果该物料还没处理过，则添加到模板
            If Not processedMaterials.Exists(materialCode) Then
                processedMaterials.Add materialCode, True
                
                Dim materialName As String
                Dim spec As String
                Dim requirement As Double
                Dim scrap As Double
                Dim inspection As Double
                Dim inbound As Double
                Dim workshopStock As Double
                Dim unit As String
                Dim manufacturer As String
                
                If colDetailMaterialName > 0 Then materialName = Trim(wsProductionDetail.Cells(i, colDetailMaterialName).Value)
                If colDetailSpec > 0 Then spec = Trim(wsProductionDetail.Cells(i, colDetailSpec).Value)
                If colDetailRequirement > 0 Then requirement = Val(wsProductionDetail.Cells(i, colDetailRequirement).Value)
                If colDetailScrap > 0 Then scrap = Val(wsProductionDetail.Cells(i, colDetailScrap).Value)
                If colDetailInspection > 0 Then inspection = Val(wsProductionDetail.Cells(i, colDetailInspection).Value)
                If colDetailInbound > 0 Then inbound = Val(wsProductionDetail.Cells(i, colDetailInbound).Value)
                
                ' 获取当前最新的车间结存
                workshopStock = GetWorkshopStock(materialCode)
                
                ' 从BOM表获取单位和生产厂家（为了保持资料最新）
                Dim wsBOM As Worksheet
                Set wsBOM = ThisWorkbook.Worksheets("BOM")
                Dim colBOMProductCode As Long, colBOMMaterialCode As Long, colBOMUnit As Long, colBOMManufacturer As Long
                Dim bomLastRow As Long, j As Long
                
                colBOMProductCode = GetColumnIndex(wsBOM, 1, "产品编号")
                colBOMMaterialCode = GetColumnIndex(wsBOM, 1, "物料编号")
                colBOMUnit = GetColumnIndex(wsBOM, 1, "单位")
                colBOMManufacturer = GetColumnIndex(wsBOM, 1, "生产厂家")
                
                If colBOMProductCode > 0 And colBOMMaterialCode > 0 Then
                    bomLastRow = wsBOM.Cells(wsBOM.Rows.Count, colBOMProductCode).End(xlUp).Row
                    For j = 2 To bomLastRow
                        If Trim(wsBOM.Cells(j, colBOMProductCode).Value) = productCode And _
                           Trim(wsBOM.Cells(j, colBOMMaterialCode).Value) = materialCode Then
                            If colBOMUnit > 0 Then unit = Trim(wsBOM.Cells(j, colBOMUnit).Value)
                            If colBOMManufacturer > 0 Then manufacturer = Trim(wsBOM.Cells(j, colBOMManufacturer).Value)
                            Exit For
                        End If
                    Next j
                End If
                
                ' 写入到模板表
                wsTemplate.Cells(currentRow, "A").Value = currentRow - 5    ' 序号
                wsTemplate.Cells(currentRow, "B").Value = materialCode      ' 物料编号
                wsTemplate.Cells(currentRow, "C").Value = materialName      ' 物料名称
                wsTemplate.Cells(currentRow, "D").Value = spec              ' 规格
                wsTemplate.Cells(currentRow, "E").Value = requirement       ' 需求量
                wsTemplate.Cells(currentRow, "F").Value = unit              ' 单位
                wsTemplate.Cells(currentRow, "G").Value = workshopStock     ' 车间结存量
                wsTemplate.Cells(currentRow, "I").Value = manufacturer      ' 生产厂家
                wsTemplate.Cells(currentRow, "L").Value = scrap             ' 报废
                wsTemplate.Cells(currentRow, "M").Value = inbound           ' 入库
                wsTemplate.Cells(currentRow, "N").Value = inspection        ' 抽检
                wsTemplate.Cells(currentRow, "O").Value = workshopStock     ' 初始结存
                
                currentRow = currentRow + 1
            End If
        End If
    Next i
    
    ' 步骤5：从出库表填充J列（批号）
    Call FillBatchNumbersFromOutbound(productionBatch)
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "反查生产批号时发生错误: " & Err.Description, vbCritical
    
CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ============================================
' ?? 从出库表填充J列批号
' 参数：productionBatch - 生产批号
' 创建日期：2026-02-11
' ============================================
Sub FillBatchNumbersFromOutbound(productionBatch As String)
    On Error GoTo ErrorHandler
    
    Dim wsTemplate As Worksheet
    Dim wsOutbound As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim materialCode As String
    Dim batchList As String
    Dim qtyList As String
    Dim outboundLastRow As Long
    Dim j As Long
    
    Set wsTemplate = Me
    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    
    ' 获取出库表列索引
    Dim colOutProductionBatch As Long
    Dim colOutMaterialCode As Long
    Dim colOutBatch As Long
    Dim colOutQty As Long
    
    colOutProductionBatch = GetColumnIndex(wsOutbound, 1, "生产批号")
    colOutMaterialCode = GetColumnIndex(wsOutbound, 1, "物料编号")
    colOutBatch = GetColumnIndex(wsOutbound, 1, "批次")
    colOutQty = GetColumnIndex(wsOutbound, 1, "出库数量")
    
    If colOutProductionBatch = 0 Or colOutMaterialCode = 0 Or colOutBatch = 0 Then
        Exit Sub
    End If
    
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row
    outboundLastRow = wsOutbound.Cells(wsOutbound.Rows.Count, colOutMaterialCode).End(xlUp).Row
    
    ' 查找备注行
    For i = lastRow To 6 Step -1
        If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    ' 遍历模板表的每一行
    For i = 6 To lastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then
            materialCode = Trim(wsTemplate.Cells(i, "B").Value)
            batchList = ""
            qtyList = ""
            
            ' 从出库表中查找该生产批号和该物料的所有批号和数量
            For j = 2 To outboundLastRow
                If Trim(wsOutbound.Cells(j, colOutProductionBatch).Value) = productionBatch And _
                   Trim(wsOutbound.Cells(j, colOutMaterialCode).Value) = materialCode Then
                    
                    ' 获取批号
                    Dim batchInfo As String
                    batchInfo = Trim(wsOutbound.Cells(j, colOutBatch).Value)
                    
                    ' 获取出库数量
                    Dim outQty As Double
                    outQty = Val(wsOutbound.Cells(j, colOutQty).Value)
                    
                    ' 组合批号列表（J列）
                    If batchList = "" Then
                        batchList = batchInfo
                    Else
                        batchList = batchList & vbLf & "--------" & vbLf & batchInfo
                    End If
                    
                    ' 组合数量列表（H列）
                    If qtyList = "" Then
                        qtyList = CStr(CLng(outQty))
                    Else
                        qtyList = qtyList & vbLf & "--------" & vbLf & CStr(CLng(outQty))
                    End If
                End If
            Next j
            
            ' 填写到J列（批号）
            If batchList <> "" Then
                wsTemplate.Cells(i, "J").Value = batchList
                wsTemplate.Cells(i, "J").WrapText = True
            End If
            
            ' 更新H列（本次领用量 - 分批号显示）
            If qtyList <> "" Then
                wsTemplate.Cells(i, "H").Value = qtyList
                wsTemplate.Cells(i, "H").WrapText = True
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    ' 静默错误
End Sub

' ============================================
' ?? 保存生产记录到生产记录表和生产记录明细表
' 在计算批号按钮点击后调用
' 创建日期：2026-02-11
' ============================================
Sub SaveProductionRecord()
    On Error GoTo ErrorHandler
    
    Dim wsTemplate As Worksheet
    Dim wsProduction As Worksheet
    Dim wsProductionDetail As Worksheet
    Dim productCode As String
    Dim productName As String
    Dim productionBatch As String
    Dim requirementQty As Double
    Dim productionDate As Date
    Dim pickupDate As Date
    Dim i As Long
    Dim lastRow As Long
    Dim newRow As Long
    
    Application.ScreenUpdating = False
    
    Set wsTemplate = Me
    Set wsProduction = ThisWorkbook.Worksheets("生产记录")
    Set wsProductionDetail = ThisWorkbook.Worksheets("生产记录明细")
    
    ' 步骤1：读取模板表的主信息
    productCode = Trim(wsTemplate.Range("E3").Value)
    productName = Trim(wsTemplate.Range("C3").Value)
    productionBatch = Trim(wsTemplate.Range("R3").Value)
    requirementQty = Val(wsTemplate.Range("C4").Value)
    
    ' 读取日期
    On Error Resume Next
    pickupDate = CDate(wsTemplate.Range("I4").Value)
    If Err.Number <> 0 Then pickupDate = Date
    Err.Clear
    
    productionDate = CDate(wsTemplate.Range("E4").Value)
    If Err.Number <> 0 Then productionDate = Date + 1
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' 验证必需数据
    If productCode = "" Or productionBatch = "" Then
        ' 缺少必需数据，不保存
        Exit Sub
    End If
    
    ' 步骤2：检查生产批号是否已存在
    Dim colProProductionBatch As Long
    Dim colProProductCode As Long
    Dim colProProductName As Long
    Dim colProRequirementQty As Long
    Dim colProPickupDate As Long
    Dim colProProductionDate As Long
    Dim colProDate As Long
    Dim productionLastRow As Long
    Dim found As Boolean
    Dim existingRow As Long
    
    colProDate = GetColumnIndex(wsProduction, 1, "日期")
    colProProductCode = GetColumnIndex(wsProduction, 1, "产品编号")
    colProProductName = GetColumnIndex(wsProduction, 1, "产品名称")
    colProProductionBatch = GetColumnIndex(wsProduction, 1, "生产批号")
    colProRequirementQty = GetColumnIndex(wsProduction, 1, "需求数量")
    
    ' 尝试获取领料日期和生产日期列
    On Error Resume Next
    colProPickupDate = GetColumnIndex(wsProduction, 1, "领料日期")
    colProProductionDate = GetColumnIndex(wsProduction, 1, "生产日期")
    On Error GoTo ErrorHandler
    
    If colProProductionBatch = 0 Then
        ' 生产记录表缺少必需列
        Exit Sub
    End If
    
    productionLastRow = wsProduction.Cells(wsProduction.Rows.Count, 1).End(xlUp).Row
    found = False
    
    ' 查找是否已存在该生产批号
    For i = 2 To productionLastRow
        If Trim(wsProduction.Cells(i, colProProductionBatch).Value) = productionBatch Then
            found = True
            existingRow = i
            Exit For
        End If
    Next i
    
    ' 步骤3：写入或更新生产记录主表
    If found Then
        ' 更新现有记录
        newRow = existingRow
    Else
        ' 新增记录
        newRow = productionLastRow + 1
    End If
    
    If colProDate > 0 Then wsProduction.Cells(newRow, colProDate).Value = pickupDate
    If colProProductCode > 0 Then wsProduction.Cells(newRow, colProProductCode).Value = productCode
    If colProProductName > 0 Then wsProduction.Cells(newRow, colProProductName).Value = productName
    If colProProductionBatch > 0 Then wsProduction.Cells(newRow, colProProductionBatch).Value = productionBatch
    If colProRequirementQty > 0 Then wsProduction.Cells(newRow, colProRequirementQty).Value = requirementQty
    If colProPickupDate > 0 Then wsProduction.Cells(newRow, colProPickupDate).Value = pickupDate
    If colProProductionDate > 0 Then wsProduction.Cells(newRow, colProProductionDate).Value = productionDate
    
    ' 步骤4：删除生产记录明细表中该生产批号的旧记录（如果是更新）
    If found Then
        Dim colDetailProductionBatch As Long
        Dim detailLastRow As Long
        Dim j As Long
        
        colDetailProductionBatch = GetColumnIndex(wsProductionDetail, 1, "生产批号")
        
        If colDetailProductionBatch > 0 Then
            detailLastRow = wsProductionDetail.Cells(wsProductionDetail.Rows.Count, colDetailProductionBatch).End(xlUp).Row
            
            ' 从后往前遍历删除（避免删除行后索引错位）
            For j = detailLastRow To 2 Step -1
                If Trim(wsProductionDetail.Cells(j, colDetailProductionBatch).Value) = productionBatch Then
                    wsProductionDetail.Rows(j).Delete Shift:=xlUp
                End If
            Next j
        End If
    End If
    
    ' 步骤5：写入生产记录明细表
    Dim colDetailMaterialCode As Long
    Dim colDetailMaterialName As Long
    Dim colDetailSpec As Long
    Dim colDetailRequirement As Long
    Dim colDetailPickup As Long
    Dim colDetailBatch As Long
    Dim colDetailScrap As Long
    Dim colDetailInspection As Long
    Dim colDetailInbound As Long
    Dim colDetailWorkshopStock As Long
    Dim detailNewRow As Long
    
    colDetailProductionBatch = GetColumnIndex(wsProductionDetail, 1, "生产批号")
    colDetailMaterialCode = GetColumnIndex(wsProductionDetail, 1, "物料编号")
    colDetailMaterialName = GetColumnIndex(wsProductionDetail, 1, "物料名称")
    colDetailSpec = GetColumnIndex(wsProductionDetail, 1, "规格")
    colDetailRequirement = GetColumnIndex(wsProductionDetail, 1, "需求量")
    colDetailPickup = GetColumnIndex(wsProductionDetail, 1, "本次领用量")
    colDetailBatch = GetColumnIndex(wsProductionDetail, 1, "批号")
    colDetailScrap = GetColumnIndex(wsProductionDetail, 1, "报废")
    colDetailInspection = GetColumnIndex(wsProductionDetail, 1, "抽检")
    colDetailInbound = GetColumnIndex(wsProductionDetail, 1, "入库")
    
    ' 尝试获取车间结存量列
    On Error Resume Next
    colDetailWorkshopStock = GetColumnIndex(wsProductionDetail, 1, "车间结存量")
    On Error GoTo ErrorHandler
    
    If colDetailProductionBatch = 0 Or colDetailMaterialCode = 0 Then
        ' 生产记录明细表缺少必需列
        Exit Sub
    End If
    
    detailNewRow = wsProductionDetail.Cells(wsProductionDetail.Rows.Count, 1).End(xlUp).Row + 1
    
    ' 遍历模板表，写入明细
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row
    
    ' 查找备注行
    For i = lastRow To 6 Step -1
        If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    For i = 6 To lastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then
            Dim materialCode As String
            Dim materialName As String
            Dim spec As String
            Dim requirement As Double
            Dim pickup As Double
            Dim batchNumber As String
            Dim scrap As Double
            Dim inspection As Double
            Dim inbound As Double
            Dim workshopStock As Double
            
            materialCode = Trim(wsTemplate.Cells(i, "B").Value)
            materialName = Trim(wsTemplate.Cells(i, "C").Value)
            spec = Trim(wsTemplate.Cells(i, "D").Value)
            requirement = Val(wsTemplate.Cells(i, "E").Value)
            scrap = Val(wsTemplate.Cells(i, "L").Value)
            inbound = Val(wsTemplate.Cells(i, "M").Value)
            inspection = Val(wsTemplate.Cells(i, "N").Value)
            workshopStock = Val(wsTemplate.Cells(i, "G").Value)
            
            ' ?? 处理多批号情况：拆分H列和J列
            Dim pickupStr As String
            Dim batchStr As String
            Dim pickupArray() As String
            Dim batchArray() As String
            Dim k As Long
            Dim batchCount As Long
            
            pickupStr = Trim(wsTemplate.Cells(i, "H").Value)
            batchStr = Trim(wsTemplate.Cells(i, "J").Value)
            
            ' 拆分批号（用"--------"分隔）
            If InStr(batchStr, "--------") > 0 Then
                ' 多批号情况
                batchArray = Split(batchStr, vbLf & "--------" & vbLf)
                pickupArray = Split(pickupStr, vbLf & "--------" & vbLf)
                batchCount = UBound(batchArray) + 1
            Else
                ' 单批号情况
                ReDim batchArray(0)
                ReDim pickupArray(0)
                batchArray(0) = batchStr
                pickupArray(0) = pickupStr
                batchCount = 1
            End If
            
            ' ?? 遍历每个批号，每个批号写入一条明细记录
            For k = 0 To batchCount - 1
                batchNumber = Trim(batchArray(k))
                pickup = Val(Trim(pickupArray(k)))
                
                ' 写入明细表
                If colDetailProductionBatch > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailProductionBatch).Value = productionBatch
                If colDetailMaterialCode > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailMaterialCode).Value = materialCode
                If colDetailMaterialName > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailMaterialName).Value = materialName
                If colDetailSpec > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailSpec).Value = spec
                If colDetailRequirement > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailRequirement).Value = requirement
                If colDetailPickup > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailPickup).Value = pickup
                If colDetailBatch > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailBatch).Value = batchNumber
                If colDetailScrap > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailScrap).Value = scrap
                If colDetailInspection > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailInspection).Value = inspection
                If colDetailInbound > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailInbound).Value = inbound
                If colDetailWorkshopStock > 0 Then wsProductionDetail.Cells(detailNewRow, colDetailWorkshopStock).Value = workshopStock
                
                detailNewRow = detailNewRow + 1
            Next k
        End If
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ' 静默错误（不影响主流程）
End Sub


' ============================================
' ?? 保存生产记录到 "生产记录" 和 "生产记录明细" 表
' ============================================
Sub SaveProductionRecords()
    On Error GoTo ErrorHandler

    Dim wsTemplate As Worksheet
    Dim wsRecord As Worksheet
    Dim wsDetail As Worksheet
    Dim recordLastRow As Long
    Dim detailLastRow As Long
    Dim templateLastRow As Long
    Dim i As Long
    Dim productionBatch As String
    Dim productCode As String
    
    ' 设置工作表
    Set wsTemplate = ThisWorkbook.Worksheets("模板")
    
    ' 尝试获取记录表，如果不存在则提示
    On Error Resume Next
    Set wsRecord = ThisWorkbook.Worksheets("生产记录")
    Set wsDetail = ThisWorkbook.Worksheets("生产记录明细")
    On Error GoTo ErrorHandler
    
    If wsRecord Is Nothing Or wsDetail Is Nothing Then
        MsgBox "未找到 '生产记录' 或 '生产记录明细' 工作表，无法保存记录。" & vbCrLf & _
               "请确保这两个工作表已经创建，并且名称完全一致。", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 获取关键信息
    productionBatch = Trim(wsTemplate.Range("I3").Value)
    productCode = Trim(wsTemplate.Range("E3").Value)
    
    If productionBatch = "" Or productCode = "" Then
        ' 如果没有生产批号或产品编号，提示用户
        MsgBox "生产批号(I3)或产品编号(E3)为空，无法保存生产记录。", vbExclamation, "提示"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' ================= 1. 清理旧记录（防止重复） =================
    
    ' A. 清理主记录表 (生产批号在第4列)
    recordLastRow = wsRecord.Cells(wsRecord.Rows.Count, 1).End(xlUp).Row
    If recordLastRow > 1 Then
        For i = recordLastRow To 2 Step -1
            If Trim(wsRecord.Cells(i, 4).Value) = productionBatch Then
                wsRecord.Rows(i).Delete
            End If
        Next i
    End If
    
    ' B. 清理明细记录表 (生产批号在第1列)
    detailLastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row
    If detailLastRow > 1 Then
        For i = detailLastRow To 2 Step -1
            If Trim(wsDetail.Cells(i, 1).Value) = productionBatch Then
                wsDetail.Rows(i).Delete
            End If
        Next i
    End If
    
    ' ================= 2. 保存新主记录 =================
    ' 重新查找最后一行
    recordLastRow = wsRecord.Cells(wsRecord.Rows.Count, 1).End(xlUp).Row
    Dim newRecordRow As Long
    newRecordRow = recordLastRow + 1
    
    ' 写入主记录
    ' 日期 | 产品编号 | 产品名称 | 生产批号 | 需求数量 | 领料日期 | 生产日期
    wsRecord.Cells(newRecordRow, 1).Value = Date  ' 日期
    wsRecord.Cells(newRecordRow, 2).Value = wsTemplate.Range("E3").Value ' 产品编号
    wsRecord.Cells(newRecordRow, 3).Value = wsTemplate.Range("C3").Value ' 产品名称
    wsRecord.Cells(newRecordRow, 4).Value = productionBatch              ' 生产批号
    wsRecord.Cells(newRecordRow, 5).Value = wsTemplate.Range("C4").Value ' 需求数量
    wsRecord.Cells(newRecordRow, 6).Value = wsTemplate.Range("I4").Value ' 领料日期
    wsRecord.Cells(newRecordRow, 7).Value = wsTemplate.Range("E4").Value ' 生产日期
    
    ' ================= 3. 保存新明细记录 =================
    ' 查找模板表数据范围
    templateLastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row
    
    ' 查找备注行
    For i = templateLastRow To 6 Step -1
        If InStr(1, wsTemplate.Cells(i, 1).Value, "备注", vbTextCompare) > 0 Then
            templateLastRow = i - 1
            Exit For
        End If
    Next i
    
    ' 查找明细表最后一行
    detailLastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row
    Dim currentDetailRow As Long
    currentDetailRow = detailLastRow
    
    ' 遍历模板表行
    For i = 6 To templateLastRow
        If Not IsEmpty(wsTemplate.Cells(i, "B")) Then
            Dim rawBatch As String
            Dim rawQty As String
            Dim batches() As String
            Dim qties() As String
            Dim k As Long
            
            rawBatch = wsTemplate.Cells(i, "J").Value
            rawQty = wsTemplate.Cells(i, "H").Value
            
            ' 使用 vbLf & "--------" & vbLf 分割。注意可能只有单个记录没有分隔符。
            ' 如果包含分隔符则分割，否则作为数组单个元素处理
            If InStr(rawBatch, "--------") > 0 Then
                batches = Split(rawBatch, vbLf & "--------" & vbLf)
                qties = Split(rawQty, vbLf & "--------" & vbLf)
            Else
                ReDim batches(0 To 0)
                ReDim qties(0 To 0)
                batches(0) = rawBatch
                qties(0) = rawQty
            End If
            
            ' 为每个批次记录写入一行
            For k = LBound(batches) To UBound(batches)
                currentDetailRow = currentDetailRow + 1
                
                ' 生产记录明细结构：
                ' 1:生产批号 | 2:物料编号 | 3:物料名称 | 4:规格 | 5:需求量 | 6:本次领用量 | 7:批号 | 8:报废 | 9:抽检 | 10:入库 | 11:车间结存量
                With wsDetail
                    .Cells(currentDetailRow, 1).Value = productionBatch                    ' 生产批号
                    .Cells(currentDetailRow, 2).Value = wsTemplate.Cells(i, "B").Value     ' 物料编号
                    .Cells(currentDetailRow, 3).Value = wsTemplate.Cells(i, "C").Value     ' 物料名称
                    .Cells(currentDetailRow, 4).Value = wsTemplate.Cells(i, "D").Value     ' 规格
                    .Cells(currentDetailRow, 5).Value = wsTemplate.Cells(i, "E").Value     ' 需求量
                    
                    ' 写入拆分后的领用量和批号
                    .Cells(currentDetailRow, 6).Value = qties(k)                          ' 本次领用量
                    .Cells(currentDetailRow, 7).Value = batches(k)                        ' 批号
                    
                    ' 其他字段（报废、入库、抽检、结存）通常是按物料汇总的，目前按行重复写入
                    .Cells(currentDetailRow, 8).Value = wsTemplate.Cells(i, "L").Value     ' 报废 (L列)
                    .Cells(currentDetailRow, 9).Value = wsTemplate.Cells(i, "N").Value     ' 抽检 (N列)
                    .Cells(currentDetailRow, 10).Value = wsTemplate.Cells(i, "M").Value    ' 入库 (M列)
                    .Cells(currentDetailRow, 11).Value = wsTemplate.Cells(i, "O").Value    ' 车间结存量 (O列：下次结存)
                End With
            Next k
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "生产记录已保存！" & vbCrLf & "已更新生产批号: " & productionBatch, vbInformation, "成功"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "保存生产记录时发生错误: " & Err.Description, vbCritical, "错误"
End Sub


