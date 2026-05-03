$config = Get-Content "config.json" | ConvertFrom-Json
$wbPath = $config.workbookPath

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$wb = $excel.Workbooks.Open($wbPath)

# Rubberduck COM automation
$rd = New-Object -ComObject Rubberduck.API

Write-Host "Running Rubberduck tests..."
$results = $rd.RunTests($wb)

foreach ($r in $results) {
    Write-Host "$($r.TestName): $($r.Outcome)"
}

$wb.Close()
$excel.Quit()
