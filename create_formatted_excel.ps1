# 读取CSV数据并创建带格式的Excel文件
$csvPath = "C:\Users\Administrator\Desktop\zhaopin\秋招公司累计汇总.csv"
$excelPath = "C:\Users\Administrator\Desktop\zhaopin\秋招公司累计汇总.xlsx"

# 创建Excel对象
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 创建工作簿
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Name = "秋招公司汇总"

# 读取CSV数据
$data = Import-Csv -Path $csvPath -Encoding UTF8

# 写入表头
$headers = $data[0].PSObject.Properties.Name
for ($i = 0; $i -lt $headers.Count; $i++) {
    $worksheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
}

# 写入数据
for ($row = 0; $row -lt $data.Count; $row++) {
    for ($col = 0; $col -lt $headers.Count; $col++) {
        $worksheet.Cells.Item($row + 2, $col + 1).Value2 = $data[$row].($headers[$col])
    }
}

# 设置表头样式
$headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $headers.Count))
$headerRange.Font.Bold = $true
$headerRange.Font.Color = 0xFFFFFF  # 白色字体
$headerRange.Interior.Color = 0x4472C4  # 蓝色背景
$headerRange.HorizontalAlignment = -4108  # 居中对齐
$headerRange.VerticalAlignment = -4108  # 垂直居中

# 设置列宽
$columnWidths = @{
    "公司名称" = 20
    "公司类型" = 18
    "招聘项目/批次" = 35
    "计算机相关岗位" = 50
    "招聘链接/地址" = 45
    "截止时间" = 25
    "工作地点" = 30
    "信息来源" = 22
    "信息可信度" = 12
    "首次收录日期" = 14
    "最近更新日期" = 14
    "备注" = 35
}

for ($i = 0; $i -lt $headers.Count; $i++) {
    $colName = $headers[$i]
    if ($columnWidths.ContainsKey($colName)) {
        $worksheet.Columns.Item($i + 1).ColumnWidth = $columnWidths[$colName]
    } else {
        $worksheet.Columns.Item($i + 1).ColumnWidth = 15
    }
}

# 设置自动换行和垂直居中
$dataRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item($data.Count + 1, $headers.Count))
$dataRange.WrapText = $true
$dataRange.VerticalAlignment = -4108  # 垂直居中

# 给信息可信度列设置颜色
$credibilityCol = 0
for ($i = 0; $i -lt $headers.Count; $i++) {
    if ($headers[$i] -eq "信息可信度") {
        $credibilityCol = $i + 1
        break
    }
}

if ($credibilityCol -gt 0) {
    for ($row = 2; $row -le $data.Count + 1; $row++) {
        $cell = $worksheet.Cells.Item($row, $credibilityCol)
        $value = $cell.Value2
        switch ($value) {
            "高" { $cell.Interior.Color = 0x92D050 }  # 绿色
            "中" { $cell.Interior.Color = 0xFFC000 }  # 黄色
            "低" { $cell.Interior.Color = 0xFF0000 }  # 红色
            "待核实" { $cell.Interior.Color = 0xFF0000 }  # 红色
        }
        $cell.HorizontalAlignment = -4108  # 居中
    }
}

# 添加边框
$borderRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item($data.Count + 1, $headers.Count))
$borderRange.Borders.LineStyle = 1  # 细线
$borderRange.Borders.Weight = 2  # 细

# 冻结首行
$worksheet.Application.ActiveWindow.SplitRow = 1
$worksheet.Application.ActiveWindow.FreezePanes = $true

# 自动调整行高
$worksheet.Rows.AutoFit() | Out-Null

# 保存文件
$workbook.SaveAs($excelPath, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
$workbook.Close()
$excel.Quit()

# 释放COM对象
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Excel文件已创建: $excelPath"
Write-Host "共 $($data.Count) 条记录"
