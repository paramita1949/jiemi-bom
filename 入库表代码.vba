Option Explicit

' ============================================
' 入库表工作表代码模块
' 功能：每次激活入库表时，自动刷新实时库存
' 创建日期：2026-02-09
' ============================================

' 上次刷新的时间（防止频繁刷新）
Private lastRefreshTime As Date

' 上次的出库表行数（用于检测出库表变化）
Private lastOutboundRowCount As Long

' 工作表激活时刷新库存
Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler

    ' 检查出库表是否有变化
    Dim wsOutbound As Worksheet
    Dim currentOutboundRowCount As Long

    Set wsOutbound = ThisWorkbook.Worksheets("出库")
    currentOutboundRowCount = wsOutbound.Cells(wsOutbound.Rows.Count, 1).End(xlUp).Row

    ' 如果是首次加载，记录当前行数
    If lastOutboundRowCount = 0 Then
        lastOutboundRowCount = currentOutboundRowCount
    End If

    ' 检查出库表行数是否有变化
    Dim hasOutboundChanged As Boolean
    hasOutboundChanged = (currentOutboundRowCount <> lastOutboundRowCount)

    ' 更新记录的行数
    lastOutboundRowCount = currentOutboundRowCount

    ' 如果出库表有变化，或者距离上次刷新超过5分钟，则刷新
    Dim needRefresh As Boolean
    needRefresh = hasOutboundChanged Or _
                  (DateDiff("s", lastRefreshTime, Now) > 300)  ' 300秒 = 5分钟

    If needRefresh Then
        ' 静默刷新库存
        Call RefreshAllInventoryQuietly

        ' 记录刷新时间
        lastRefreshTime = Now

        ' 在状态栏显示提示
        Application.StatusBar = "入库表：库存已自动刷新 " & Format(Now(), "hh:mm:ss")

        ' 2秒后清除状态栏
        Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBarQuiet"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "入库表激活监控错误: " & Err.Description
End Sub

' 工作表打开时初始化
Private Sub Worksheet_Open()
    lastRefreshTime = 0
    lastOutboundRowCount = 0
End Sub

' 工作表计算前（可选：在数据重新计算时刷新）
Private Sub Worksheet_Calculate()
    ' 这个事件在某些情况下会触发，可以根据需要启用
    ' Call RefreshAllInventoryQuietly
End Sub

' ============================================
' 监听单元格变更事件
' ============================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler

    Dim wsBOM As Worksheet
    Dim changedRow As Long
    Dim changedCol As Long
    Dim materialCode As String
    Dim inboundQty As Double
    Dim specValue As String
    Dim specQty As Double
    
    ' 获取关键列的索引
    Dim colMaterialCode As Long
    Dim colMaterialName As Long
    Dim colManufacturer As Long
    Dim colUnit As Long
    Dim colSpec As Long
    Dim colInQty As Long
    Dim colAuxQty As Long
    
    ' 获取当前工作表的列索引
    colMaterialCode = GetColumnIndex(Me, 1, "物料编号")
    colMaterialName = GetColumnIndex(Me, 1, "物料名称")
    colManufacturer = GetColumnIndex(Me, 1, "生产厂家")
    colUnit = GetColumnIndex(Me, 1, "单位")
    colSpec = GetColumnIndex(Me, 1, "规格")
    colInQty = GetColumnIndex(Me, 1, "入库数量")
    colAuxQty = GetColumnIndex(Me, 1, "辅数量")
    
    ' 如果关键列未找到，直接退出
    If colMaterialCode = 0 Then Exit Sub
    
    ' 检查变更单元格是否在关键列范围内
    If Intersect(Target, Me.Range(Me.Cells(2, colMaterialCode), Me.Cells(Me.Rows.Count, colMaterialCode))) Is Nothing And _
       Intersect(Target, Me.Range(Me.Cells(2, colInQty), Me.Cells(Me.Rows.Count, colInQty))) Is Nothing Then
        Exit Sub
    End If
    
    Application.EnableEvents = False
    
    ' 遍历每一个变更的单元格（支持批量操作）
    Dim cell As Range
    For Each cell In Target
        changedRow = cell.Row
        changedCol = cell.Column
        
        ' ----------------------------------------------------
        ' 场景1：输入物料编号 -> 从BOM表获取信息
        ' ----------------------------------------------------
        ' ----------------------------------------------------
        ' 场景1：输入物料编号 -> 从BOM表获取信息
        ' ----------------------------------------------------
        If changedCol = colMaterialCode Then
            materialCode = Trim(Me.Cells(changedRow, colMaterialCode).Value)
            
            If materialCode <> "" Then
                ' 查找 物料 表中的物料信息
                Dim wsMaterial As Worksheet
                Set wsMaterial = ThisWorkbook.Worksheets("物料")
                Dim matLastRow As Long
                Dim matRow As Long
                
                ' 获取 物料 表的关键列索引
                Dim matColMaterialCode As Long
                Dim matColMaterialName As Long
                Dim matColManufacturer As Long
                Dim matColUnit As Long
                Dim matColSpec As Long
                
                matColMaterialCode = GetColumnIndex(wsMaterial, 1, "物料编号")
                matColMaterialName = GetColumnIndex(wsMaterial, 1, "物料名称")
                matColManufacturer = GetColumnIndex(wsMaterial, 1, "生产厂家")
                matColUnit = GetColumnIndex(wsMaterial, 1, "单位")
                matColSpec = GetColumnIndex(wsMaterial, 1, "规格")
                
                If matColMaterialCode > 0 Then
                    matLastRow = wsMaterial.Cells(wsMaterial.Rows.Count, matColMaterialCode).End(xlUp).Row
                    
                    ' 使用字典来去重收集信息
                    Dim dictMfr As Object
                    Dim dictSpec As Object
                    
                    Set dictMfr = CreateObject("Scripting.Dictionary")
                    Set dictSpec = CreateObject("Scripting.Dictionary")
                    
                    Dim matchFound As Boolean
                    matchFound = False
                    
                    Dim tempName As String
                    Dim tempUnit As String
                    
                    ' 遍历 物料 表查找所有匹配项
                    For matRow = 2 To matLastRow
                        If Trim(wsMaterial.Cells(matRow, matColMaterialCode).Value) = materialCode Then
                            matchFound = True
                            
                            ' 收集物料名称和单位
                            If matColMaterialName > 0 Then tempName = wsMaterial.Cells(matRow, matColMaterialName).Value
                            If matColUnit > 0 Then tempUnit = wsMaterial.Cells(matRow, matColUnit).Value
                            
                            ' 收集厂家（支持逗号分隔）
                            If matColManufacturer > 0 Then
                                Dim rawMfr As String
                                rawMfr = Trim(wsMaterial.Cells(matRow, matColManufacturer).Value)
                                If rawMfr <> "" Then
                                    rawMfr = Replace(rawMfr, "，", ",")
                                    Dim mfrArr() As String
                                    mfrArr = Split(rawMfr, ",")
                                    Dim m As Variant
                                    For Each m In mfrArr
                                        If Trim(m) <> "" Then dictMfr(Trim(m)) = 1
                                    Next m
                                End If
                            End If
                            
                            ' 收集规格（支持逗号分隔）
                            If matColSpec > 0 Then
                                Dim rawSpec As String
                                rawSpec = Trim(wsMaterial.Cells(matRow, matColSpec).Value)
                                If rawSpec <> "" Then
                                    rawSpec = Replace(rawSpec, "，", ",")
                                    Dim specArr() As String
                                    specArr = Split(rawSpec, ",")
                                    Dim s As Variant
                                    For Each s In specArr
                                        If Trim(s) <> "" Then dictSpec(Trim(s)) = 1
                                    Next s
                                End If
                            End If
                        End If
                    Next matRow
                    
                    If matchFound Then
                        ' 1. 填充基本信息
                        If colMaterialName > 0 Then Me.Cells(changedRow, colMaterialName).Value = tempName
                        If colUnit > 0 Then Me.Cells(changedRow, colUnit).Value = tempUnit
                        
                        ' 2. 处理厂家选择
                        If colManufacturer > 0 And dictMfr.Count > 0 Then
                            Dim finalMfr As String
                            If dictMfr.Count = 1 Then
                                finalMfr = dictMfr.Keys()(0)
                            Else
                                ' 弹出选择框
                                Dim promptMfr As String
                                promptMfr = "发现多个生产厂家，请输入序号选择：" & vbCrLf
                                Dim k As Integer
                                Dim keysMfr As Variant
                                keysMfr = dictMfr.Keys
                                For k = 0 To UBound(keysMfr)
                                    promptMfr = promptMfr & (k + 1) & ". " & keysMfr(k) & vbCrLf
                                Next k
                                
                                Dim selMfr As String
                                Dim validMfr As Boolean
                                validMfr = False
                                Do
                                    selMfr = InputBox(promptMfr, "选择生产厂家", "1")
                                    If selMfr = "" Then
                                        finalMfr = keysMfr(0) ' 默认第一个
                                        validMfr = True
                                    ElseIf IsNumeric(selMfr) Then
                                        Dim idxMfr As Integer
                                        idxMfr = CInt(selMfr)
                                        If idxMfr >= 1 And idxMfr <= dictMfr.Count Then
                                            finalMfr = keysMfr(idxMfr - 1)
                                            validMfr = True
                                        End If
                                    End If
                                Loop Until validMfr
                            End If
                            Me.Cells(changedRow, colManufacturer).Value = finalMfr
                        End If
                        
                        ' 3. 处理规格选择
                        If colSpec > 0 And dictSpec.Count > 0 Then
                            Dim finalSpec As String
                            If dictSpec.Count = 1 Then
                                finalSpec = dictSpec.Keys()(0)
                            Else
                                ' 弹出选择框
                                Dim promptSpec As String
                                promptSpec = "发现多种规格，请输入序号选择：" & vbCrLf
                                Dim j As Integer
                                Dim keysSpec As Variant
                                keysSpec = dictSpec.Keys
                                For j = 0 To UBound(keysSpec)
                                    promptSpec = promptSpec & (j + 1) & ". " & keysSpec(j) & vbCrLf
                                Next j
                                
                                Dim selSpec As String
                                Dim validSpec As Boolean
                                validSpec = False
                                Do
                                    selSpec = InputBox(promptSpec, "选择规格", "1")
                                    If selSpec = "" Then
                                        finalSpec = keysSpec(0) ' 默认第一个
                                        validSpec = True
                                    ElseIf IsNumeric(selSpec) Then
                                        Dim idxSpec As Integer
                                        idxSpec = CInt(selSpec)
                                        If idxSpec >= 1 And idxSpec <= dictSpec.Count Then
                                            finalSpec = keysSpec(idxSpec - 1)
                                            validSpec = True
                                        End If
                                    End If
                                Loop Until validSpec
                            End If
                            Me.Cells(changedRow, colSpec).Value = finalSpec
                        End If
                    End If
                End If
            Else
                ' 如果清空了物料编号，也清空其他列
                If colMaterialName > 0 Then Me.Cells(changedRow, colMaterialName).ClearContents
                If colManufacturer > 0 Then Me.Cells(changedRow, colManufacturer).ClearContents
                If colUnit > 0 Then Me.Cells(changedRow, colUnit).ClearContents
                If colSpec > 0 Then Me.Cells(changedRow, colSpec).ClearContents
                If colAuxQty > 0 Then Me.Cells(changedRow, colAuxQty).ClearContents
            End If
        End If
        
        ' ----------------------------------------------------
        ' 场景2：输入入库数量 或 物料编号改变（导致规格改变） -> 计算辅数量
        ' ----------------------------------------------------
        If changedCol = colMaterialCode Or changedCol = colInQty Then
             ' 确保有规格和入库数量列
             If colSpec > 0 And colInQty > 0 And colAuxQty > 0 Then
                inboundQty = Val(Me.Cells(changedRow, colInQty).Value)
                specValue = Trim(Me.Cells(changedRow, colSpec).Value)
                
                ' 如果有入库数量和规格，计算辅数量
                If inboundQty > 0 And specValue <> "" Then
                    ' 使用 Module1 中的公共函数提取每个包装的数量
                    specQty = ExtractSpecQuantity(specValue)
                    
                    If specQty > 0 Then
                        Me.Cells(changedRow, colAuxQty).Value = inboundQty / specQty
                    Else
                        Me.Cells(changedRow, colAuxQty).Value = inboundQty
                    End If
                Else
                    Me.Cells(changedRow, colAuxQty).ClearContents
                End If
             End If
        End If
    Next cell
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "自动填充时发生错误: " & Err.Description, vbCritical
    
CleanUp:
    Application.EnableEvents = True
End Sub

' 辅助函数：获取列索引（复用 Module1 的逻辑，或者直接调用 Module1.GetColumnIndex）
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
