Option Explicit

' ============================================
' BOM表工作表代码模块
' 功能：输入物料编号时，自动从物料表获取信息
' 创建日期：2026-02-10
' ============================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler

    Dim wsMaterial As Worksheet
    Dim changedRow As Long
    Dim changedCol As Long
    Dim materialCode As String
    
    ' 获取关键列的索引
    Dim colProductCode As Long
    Dim colProductName As Long
    Dim colMaterialCode As Long
    Dim colMaterialName As Long
    Dim colUnit As Long
    Dim colSpec As Long
    Dim colManufacturer As Long
    
    ' 获取当前工作表的列索引
    colProductCode = GetColumnIndex(Me, 1, "产品编号")
    colProductName = GetColumnIndex(Me, 1, "产品名称")
    colMaterialCode = GetColumnIndex(Me, 1, "物料编号")
    colMaterialName = GetColumnIndex(Me, 1, "物料名称")
    colUnit = GetColumnIndex(Me, 1, "单位")
    colSpec = GetColumnIndex(Me, 1, "规格")
    colManufacturer = GetColumnIndex(Me, 1, "生产厂家")
    
    ' 如果关键列未找到，直接退出
    If colMaterialCode = 0 Then Exit Sub
    
    ' 检查变更单元格是否在关键列范围内
    If Intersect(Target, Me.Range(Me.Cells(2, colMaterialCode), Me.Cells(Me.Rows.Count, colMaterialCode))) Is Nothing Then
        Exit Sub
    End If
    
    Application.EnableEvents = False
    
    ' 遍历每一个变更的单元格
    Dim cell As Range
    For Each cell In Target
        changedRow = cell.Row
        changedCol = cell.Column
        
        If changedCol = colMaterialCode Then
            materialCode = Trim(Me.Cells(changedRow, colMaterialCode).Value)
            
            If materialCode <> "" Then
                ' 查找 物料 表中的信息
                Set wsMaterial = ThisWorkbook.Worksheets("物料")
                Dim matLastRow As Long
                Dim matRow As Long
                Dim foundCount As Integer
                Dim foundRows() As Long ' 保存找到的行号
                Dim manufacturers As String
                Dim prompt As String
                Dim userSelection As String
                Dim selectedRow As Long
                Dim selIdx As Integer
                
                ' 获取 物料 表的关键列索引
                Dim matColCode As Long
                Dim matColName As Long
                Dim matColManufacturer As Long
                Dim matColUnit As Long
                Dim matColSpec As Long
                
                matColCode = GetColumnIndex(wsMaterial, 1, "物料编号")
                matColName = GetColumnIndex(wsMaterial, 1, "物料名称")
                matColManufacturer = GetColumnIndex(wsMaterial, 1, "生产厂家")
                matColUnit = GetColumnIndex(wsMaterial, 1, "单位")
                matColSpec = GetColumnIndex(wsMaterial, 1, "规格")
                
                If matColCode > 0 Then
                    matLastRow = wsMaterial.Cells(wsMaterial.Rows.Count, matColCode).End(xlUp).Row
                    foundCount = 0
                    ReDim foundRows(1 To 10) ' 初始大小
                    
                    ' 遍历查找匹配的物料编号
                    For matRow = 2 To matLastRow
                        If Trim(wsMaterial.Cells(matRow, matColCode).Value) = materialCode Then
                            foundCount = foundCount + 1
                            If foundCount > UBound(foundRows) Then ReDim Preserve foundRows(1 To foundCount + 10)
                            foundRows(foundCount) = matRow
                        End If
                    Next matRow
                    
                    If foundCount > 0 Then
                        Dim finalManufacturer As String
                        Dim mfrRaw As String
                        selectedRow = foundRows(1) ' 默认行

                        ' 获取第一条记录的厂家信息
                        If matColManufacturer > 0 Then
                            mfrRaw = wsMaterial.Cells(selectedRow, matColManufacturer).Value
                        End If

                        ' ----------------------------
                        ' 逻辑处理：单条还是多条
                        ' ----------------------------
                        Dim mfrList() As String
                        Dim isMultiRow As Boolean
                        Dim isMultiMfrInOneCell As Boolean
                        
                        isMultiRow = (foundCount > 1)
                        isMultiMfrInOneCell = (InStr(mfrRaw, ",") > 0 Or InStr(mfrRaw, "，") > 0)
                        
                        ' 初始化最终厂家为获取到的原始值
                        finalManufacturer = mfrRaw

                        ' 场景1：多行记录 -> 弹出选择
                        If isMultiRow Then
                            prompt = "发现多个生产厂家（多行记录），请输入对应序号进行选择：" & vbCrLf & vbCrLf
                            Dim i As Integer
                            Dim mfrName As String
                            
                            For i = 1 To foundCount
                                mfrName = ""
                                If matColManufacturer > 0 Then
                                    mfrName = wsMaterial.Cells(foundRows(i), matColManufacturer).Value
                                End If
                                prompt = prompt & i & ". " & mfrName & vbCrLf
                            Next i
                            
                            ' 调用选择逻辑
                            Call SelectManufacturer(prompt, foundCount, userSelection, selectedRow, foundRows)
                            
                            ' 更新选中的厂家名称
                            If matColManufacturer > 0 Then
                                finalManufacturer = wsMaterial.Cells(selectedRow, matColManufacturer).Value
                            End If
                            
                        ' 场景2：单行记录但包含逗号 -> 拆分并弹出选择
                        ElseIf isMultiMfrInOneCell Then
                            ' 统一分隔符
                            mfrRaw = Replace(mfrRaw, "，", ",")
                            mfrList = Split(mfrRaw, ",")
                            
                            prompt = "发现多个生产厂家（单行多值），请输入对应序号进行选择：" & vbCrLf & vbCrLf
                            Dim k As Integer
                            For k = 0 To UBound(mfrList)
                                prompt = prompt & (k + 1) & ". " & Trim(mfrList(k)) & vbCrLf
                            Next k
                            
                            ' 简单的输入验证
                            Dim validSelection As Boolean
                            validSelection = False
                            Do
                                userSelection = InputBox(prompt, "选择生产厂家", "1")
                                If userSelection = "" Then
                                    finalManufacturer = Trim(mfrList(0)) ' 默认第一个
                                    validSelection = True
                                Else
                                    If IsNumeric(userSelection) Then
                                        selIdx = CInt(userSelection)
                                        If selIdx >= 1 And selIdx <= UBound(mfrList) + 1 Then
                                            finalManufacturer = Trim(mfrList(selIdx - 1))
                                            validSelection = True
                                        Else
                                            MsgBox "请输入有效的序号 (1-" & (UBound(mfrList) + 1) & ")", vbExclamation
                                        End If
                                    Else
                                        MsgBox "请输入数字序号", vbExclamation
                                    End If
                                End If
                            Loop Until validSelection
                        End If
                        
                        ' ----------------------------
                        ' 填充数据
                        ' ----------------------------
                        ' 填充 生产厂家 (使用 finalManufacturer)
                        If colManufacturer > 0 Then
                            Me.Cells(changedRow, colManufacturer).Value = finalManufacturer
                        End If
                        
                        ' 填充 物料名称
                        If colMaterialName > 0 And matColName > 0 Then
                            Me.Cells(changedRow, colMaterialName).Value = wsMaterial.Cells(selectedRow, matColName).Value
                        End If
                        
                        ' 填充 单位
                        If colUnit > 0 And matColUnit > 0 Then
                            Me.Cells(changedRow, colUnit).Value = wsMaterial.Cells(selectedRow, matColUnit).Value
                        End If
                        
                        ' 填充 规格 (处理多规格选择)
                        If colSpec > 0 And matColSpec > 0 Then
                            Dim specRaw As String
                            Dim finalSpec As String
                            Dim specList() As String
                            Dim isMultiSpec As Boolean
                            
                            specRaw = wsMaterial.Cells(selectedRow, matColSpec).Value
                            finalSpec = specRaw
                            
                            ' 检查是否包含逗号分隔的多规格
                            isMultiSpec = (InStr(specRaw, ",") > 0 Or InStr(specRaw, "，") > 0)
                            
                            If isMultiSpec Then
                                ' 统一分隔符
                                specRaw = Replace(specRaw, "，", ",")
                                specList = Split(specRaw, ",")
                                
                                prompt = "发现多种规格，请输入对应序号进行选择：" & vbCrLf & vbCrLf
                                Dim s As Integer
                               For s = 0 To UBound(specList)
                                    prompt = prompt & (s + 1) & ". " & Trim(specList(s)) & vbCrLf
                                Next s
                                
                                ' 调用选择逻辑（复用之前的逻辑结构，但因为变量类型不同需要内联或泛型，这里内联简单点）
                                Dim validSpecSelection As Boolean
                                validSpecSelection = False
                                Do
                                    userSelection = InputBox(prompt, "选择规格", "1")
                                    If userSelection = "" Then
                                        finalSpec = Trim(specList(0)) ' 默认第一个
                                        validSpecSelection = True
                                    Else
                                        If IsNumeric(userSelection) Then
                                            selIdx = CInt(userSelection)
                                            If selIdx >= 1 And selIdx <= UBound(specList) + 1 Then
                                                finalSpec = Trim(specList(selIdx - 1))
                                                validSpecSelection = True
                                            Else
                                                MsgBox "请输入有效的序号 (1-" & (UBound(specList) + 1) & ")", vbExclamation
                                            End If
                                        Else
                                            MsgBox "请输入数字序号", vbExclamation
                                        End If
                                    End If
                                Loop Until validSpecSelection
                            End If
                            
                            Me.Cells(changedRow, colSpec).Value = finalSpec
                        End If
                        
                    Else
                        ' 未找到
                        ' MsgBox "未在物料表中找到编号: " & materialCode, vbExclamation
                    End If
                End If
            Else
                ' 清空内容
                If colMaterialName > 0 Then Me.Cells(changedRow, colMaterialName).ClearContents
                If colUnit > 0 Then Me.Cells(changedRow, colUnit).ClearContents
                If colSpec > 0 Then Me.Cells(changedRow, colSpec).ClearContents
                If colManufacturer > 0 Then Me.Cells(changedRow, colManufacturer).ClearContents
            End If
        End If
    Next cell
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "BOM自动填充错误: " & Err.Description, vbCritical
    
CleanUp:
    Application.EnableEvents = True
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

' 辅助过程：处理生产厂家选择（仅用于多行记录场景）
Private Sub SelectManufacturer(prompt As String, foundCount As Integer, ByRef userSelection As String, ByRef selectedRow As Long, ByRef foundRows() As Long)
    Dim validSelection As Boolean
    validSelection = False
    
    Do
        userSelection = InputBox(prompt, "选择生产厂家", "1")
        ' 如果用户点了取消或空，默认选择第一个
        If userSelection = "" Then
            selectedRow = foundRows(1)
            validSelection = True
        Else
            If IsNumeric(userSelection) Then
                Dim selIdx As Integer
                selIdx = CInt(userSelection)
                If selIdx >= 1 And selIdx <= foundCount Then
                    selectedRow = foundRows(selIdx)
                    validSelection = True
                Else
                    MsgBox "请输入有效的序号 (1-" & foundCount & ")", vbExclamation
                End If
            Else
                MsgBox "请输入数字序号", vbExclamation
            End If
        End If
    Loop Until validSelection
End Sub
