Write-Host "=== Excel VSCode Template Configuration ==="

# Ask for workbook name
$name = Read-Host "Enter the name of your workbook (without extension)"
$path = Join-Path (Get-Location) "$name.xlsm"

# Create workbook if missing
if (-not (Test-Path $path)) {
    Write-Host "Creating workbook: $path"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($path, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close()
    $excel.Quit()
}

# Build config.json
$config = @{
    workbookPath = $path
}

$config | ConvertTo-Json | Set-Content "config.json"

Write-Host "Configuration complete."
Write-Host "Workbook path saved to config.json"
