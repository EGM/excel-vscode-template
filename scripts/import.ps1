$config = Get-Content "config.json" | ConvertFrom-Json
$wbPath = $config.workbookPath

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$wb = $excel.Workbooks.Open($wbPath)

foreach ($file in Get-ChildItem "./src" -Filter *.bas) {
    $moduleName = $file.BaseName
    Write-Host "Importing $moduleName"
    $wb.VBProject.VBComponents.Import($file.FullName)
}

foreach ($file in Get-ChildItem "./tests" -Filter *.bas) {
    $moduleName = $file.BaseName
    Write-Host "Importing test module $moduleName"
    $wb.VBProject.VBComponents.Import($file.FullName)
}

$wb.Save()
$wb.Close()
$excel.Quit()

Write-Host "Import complete."
