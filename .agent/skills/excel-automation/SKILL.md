---
name: Excel Automation Expert
description: Excel 自动化专家 - 快速生成 VBA 宏代码，处理常见 Excel 自动化任务
---

# Excel Automation Expert

这个技能帮助您快速生成高质量的 Excel VBA 自动化代码，处理各种常见和复杂的 Excel 操作任务。

## 核心功能

### 1. 数据处理自动化
- **数据导入导出**: CSV, TXT, 其他工作簿
- **数据清洗**: 去重、格式化、验证
- **数据转换**: 透视、合并、拆分
- **批量操作**: 多文件处理、批量更新

### 2. 报表生成
- **动态报表**: 根据模板生成报表
- **图表自动化**: 创建和更新图表
- **格式化**: 条件格式、样式应用
- **PDF导出**: 批量导出为 PDF

### 3. 工作表操作
- **工作表管理**: 创建、删除、复制、重命名
- **数据筛选**: 高级筛选、自动筛选
- **排序**: 多列排序、自定义排序
- **查找替换**: 批量查找替换

### 4. 用户界面
- **UserForm设计**: 数据录入表单
- **自定义菜单**: 功能区按钮
- **进度提示**: 长时间操作的进度条
- **数据验证**: 输入验证和提示

## 常用代码模板

### 模板 1: 数据导入与清洗
```vba
Option Explicit

Sub ImportAndCleanData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dataArray As Variant
    
    ' 禁用屏幕更新提高性能
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set ws = ThisWorkbook.Worksheets("数据")
    
    ' 导入数据 (示例: 从 CSV)
    With ws.QueryTables.Add( _
        Connection:="TEXT;" & ThisWorkbook.Path & "\data.csv", _
        Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    
    ' 获取数据范围
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    dataArray = ws.Range("A2:Z" & lastRow).Value
    
    ' 数据清洗
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        ' 去除空格
        dataArray(i, 1) = Trim(dataArray(i, 1))
        ' 格式化日期
        If IsDate(dataArray(i, 2)) Then
            dataArray(i, 2) = Format(dataArray(i, 2), "yyyy-mm-dd")
        End If
        ' 自定义清洗逻辑...
    Next i
    
    ' 写回数据
    ws.Range("A2:Z" & lastRow).Value = dataArray
    
    ' 去重
    ws.Range("A1:Z" & lastRow).RemoveDuplicates _
        Columns:=1, Header:=xlYes
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set ws = Nothing
    MsgBox "数据导入并清洗完成！", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
```

### 模板 2: 批量文件处理
```vba
Sub ProcessMultipleFiles()
    On Error GoTo ErrorHandler
    
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim masterWs As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 设置文件夹路径
    folderPath = ThisWorkbook.Path & "\数据文件\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' 准备汇总工作表
    Set masterWs = ThisWorkbook.Worksheets("汇总")
    masterWs.Cells.Clear
    
    ' 写入表头
    masterWs.Range("A1:E1").Value = Array("文件名", "日期", "数量", "金额", "备注")
    lastRow = 2
    
    ' 遍历文件夹
    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        If fileName <> ThisWorkbook.Name Then
            Set wb = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
            Set ws = wb.Worksheets(1)
            
            ' 提取数据
            With masterWs
                .Cells(lastRow, 1).Value = fileName
                .Cells(lastRow, 2).Value = ws.Range("B2").Value
                .Cells(lastRow, 3).Value = ws.Range("C2").Value
                .Cells(lastRow, 4).Value = ws.Range("D2").Value
                .Cells(lastRow, 5).Value = ws.Range("E2").Value
            End With
            
            wb.Close SaveChanges:=False
            lastRow = lastRow + 1
        End If
        
        fileName = Dir()
    Loop
    
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set wb = Nothing
    Set ws = Nothing
    Set masterWs = Nothing
    MsgBox "处理完成，共处理 " & lastRow - 2 & " 个文件", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "错误: " & Err.Description & vbCrLf & "文件: " & fileName, vbCritical
    Resume CleanUp
End Sub
```

### 模板 3: 动态报表生成
```vba
Sub GenerateReport()
    On Error GoTo ErrorHandler
    
    Dim wsData As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, i As Long
    Dim reportDate As Date
    
    Application.ScreenUpdating = False
    
    Set wsData = ThisWorkbook.Worksheets("原始数据")
    
    ' 创建或清空报表工作表
    On Error Resume Next
    Set wsReport = ThisWorkbook.Worksheets("报表_" & Format(Date, "yyyymmdd"))
    On Error GoTo ErrorHandler
    
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Worksheets.Add
        wsReport.Name = "报表_" & Format(Date, "yyyymmdd")
    Else
        wsReport.Cells.Clear
    End If
    
    ' 设置报表标题
    With wsReport
        .Range("A1:F1").Merge
        .Range("A1").Value = "月度销售报表"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' 设置表头
        .Range("A3:F3").Value = Array("序号", "产品名称", "销售数量", "单价", "金额", "日期")
        .Range("A3:F3").Font.Bold = True
        .Range("A3:F3").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' 复制数据并计算
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        With wsReport
            .Cells(i + 2, 1).Value = i - 1  ' 序号
            .Cells(i + 2, 2).Value = wsData.Cells(i, 1).Value  ' 产品名称
            .Cells(i + 2, 3).Value = wsData.Cells(i, 2).Value  ' 数量
            .Cells(i + 2, 4).Value = wsData.Cells(i, 3).Value  ' 单价
            .Cells(i + 2, 5).Formula = "=C" & i + 2 & "*D" & i + 2  ' 金额
            .Cells(i + 2, 6).Value = wsData.Cells(i, 4).Value  ' 日期
        End With
    Next i
    
    ' 添加汇总行
    With wsReport
        .Cells(lastRow + 3, 2).Value = "合计:"
        .Cells(lastRow + 3, 2).Font.Bold = True
        .Cells(lastRow + 3, 3).Formula = "=SUM(C4:C" & lastRow + 2 & ")"
        .Cells(lastRow + 3, 5).Formula = "=SUM(E4:E" & lastRow + 2 & ")"
        
        ' 格式化
        .Range("D4:E" & lastRow + 3).NumberFormat = "#,##0.00"
        .Range("F4:F" & lastRow + 2).NumberFormat = "yyyy-mm-dd"
        .Columns("A:F").AutoFit
        
        ' 添加边框
        .Range("A3:F" & lastRow + 3).Borders.LineStyle = xlContinuous
    End With
    
CleanUp:
    Application.ScreenUpdating = True
    Set wsData = Nothing
    Set wsReport = Nothing
    MsgBox "报表生成完成！", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
```

### 模板 4: 进度条提示
```vba
Sub LongProcessWithProgress()
    Dim i As Long, total As Long
    Dim progressForm As Object
    
    total = 10000
    
    ' 创建简单的进度提示
    Application.StatusBar = "处理中... 0%"
    
    For i = 1 To total
        ' 执行操作...
        DoSomething i
        
        ' 更新进度 (每100次更新一次)
        If i Mod 100 = 0 Then
            Application.StatusBar = "处理中... " & _
                Format(i / total, "0%") & " (" & i & "/" & total & ")"
        End If
    Next i
    
    Application.StatusBar = False
    MsgBox "处理完成！", vbInformation
End Sub

Private Sub DoSomething(index As Long)
    ' 实际操作...
End Sub
```

### 模板 5: Worksheet_Change 事件（智能响应）
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim changedCell As Range
    
    ' 禁用事件避免循环触发
    Application.EnableEvents = False
    
    ' 处理多个单元格变化
    For Each changedCell In Target
        ' 示例1: C列变化时自动计算D列
        If changedCell.Column = 3 And changedCell.Row > 1 Then
            If IsNumeric(changedCell.Value) Then
                changedCell.Offset(0, 1).Value = changedCell.Value * 1.1
            End If
        End If
        
        ' 示例2: 日期格式转换
        If changedCell.Address = "$C$7" Then
            If changedCell.Value <> "" Then
                ' 从 "YYYY.M.D" 转换为 "YYYYMMDD"
                Me.Range("F4").Value = Format(DateValue(Replace(changedCell.Value, ".", "-")), "yyyymmdd")
            End If
        End If
        
        ' 示例3: 数据验证
        If changedCell.Column = 5 Then
            If Not IsNumeric(changedCell.Value) Or changedCell.Value < 0 Then
                MsgBox "请输入有效的数值", vbExclamation
                changedCell.Value = ""
            End If
        End If
    Next changedCell
    
CleanUp:
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "错误: " & Err.Description, vbCritical
End Sub
```

## 使用指南

当您需要自动化 Excel 任务时，告诉我：

1. **任务描述**: 您想要实现什么功能
2. **数据结构**: 数据在哪些列、格式是什么
3. **期望结果**: 最终想要得到什么
4. **特殊要求**: 性能、错误处理等

我将为您生成：
- ✅ 完整的 VBA 代码
- ✅ 详细的注释说明
- ✅ 错误处理机制
- ✅ 性能优化建议
- ✅ 使用说明

## 最佳实践

### 性能优化清单
```vba
' 在批量操作前
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayAlerts = False

' 执行操作...

' 操作完成后恢复
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayAlerts = True
```

### 安全的文件操作
```vba
Dim wb As Workbook
On Error Resume Next
Set wb = Workbooks.Open(filePath)
On Error GoTo ErrorHandler

If wb Is Nothing Then
    MsgBox "无法打开文件: " & filePath, vbCritical
    Exit Sub
End If

' 操作文件...

wb.Close SaveChanges:=True
Set wb = Nothing
```

## 快速参考

### 常用对象层次结构
```
Application
  └─ Workbooks
      └─ Worksheets
          └─ Range
              └─ Cell
```

### 常用属性和方法
- `Range().Value` - 获取/设置值
- `Range().Formula` - 获取/设置公式
- `Range().Copy` - 复制
- `Range().PasteSpecial` - 选择性粘贴
- `Cells(row, col)` - 单元格引用
- `Range().End(xlUp/xlDown)` - 查找边界

### 调试技巧
```vba
Debug.Print variable  ' 输出到立即窗口
Stop  ' 设置断点
MsgBox variable  ' 显示变量值
```
