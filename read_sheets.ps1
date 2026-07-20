
# Read Excel file and dump sheet names + data structure
$excelPath = "C:\Users\Acer\OneDrive\^Gاحصاء ال اطفيحه.xlsx"
$altPath = "C:\Users\Acer\OneDrive\ احصاء ال اطفيحه.xlsx"

# Try primary path
if (-not (Test-Path $excelPath)) {
    $excelPath = $altPath
}

if (-not (Test-Path $excelPath)) {
    Write-Host "File not found at either path"
    # List all xlsx files in OneDrive
    Get-ChildItem "C:\Users\Acer\OneDrive\" -Filter "*.xlsx" | ForEach-Object { Write-Host "Found: $($_.FullName)" }
    exit 1
}

Write-Host "Reading file: $excelPath"
Write-Host "File size: $((Get-Item $excelPath).Length) bytes"
Write-Host ""

# Use ExcelPackage / EPPlus if available, otherwise use COM object
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($excelPath)
    
    Write-Host "=== SHEET NAMES ==="
    foreach ($sheet in $workbook.Sheets) {
        Write-Host "- '$($sheet.Name)'"
    }
    Write-Host ""
    
    # Look for receipts / payment sheet
    $targetSheets = @("وصولات الدفع", "وصولات العائلات", "وصولات", "التحصيل", "دفع", "sheet1", "Sheet1")
    
    foreach ($sheetName in $targetSheets) {
        $sheet = $workbook.Sheets | Where-Object { $_.Name -like "*$($sheetName)*" }
        if ($sheet) {
            Write-Host "=== SHEET: $($sheet.Name) ==="
            $usedRange = $sheet.UsedRange
            $rowCount = [Math]::Min($usedRange.Rows.Count, 30)
            $colCount = [Math]::Min($usedRange.Columns.Count, 20)
            
            for ($row = 1; $row -le $rowCount; $row++) {
                $rowData = @()
                for ($col = 1; $col -le $colCount; $col++) {
                    $cell = $usedRange.Cells.Item($row, $col)
                    $val = if ($cell.Value2) { $cell.Value2.ToString() } else { "" }
                    $rowData += $val
                }
                Write-Host "Row $($row): $($rowData -join ' | ')"
            }
            Write-Host ""
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Host "ERROR: $_"
}
