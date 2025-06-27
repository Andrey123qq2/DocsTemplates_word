$ConfigFile = "Config_all.json"
$CurrentFolder = (Split-Path $MyInvocation.MyCommand.Path -Parent)
Import-Module "$CurrentFolder\Utils.psm1" -Force
$Config = Get-Config -Folder $CurrentFolder -ConfigFile $ConfigFile
$ExcelFilePath = $Config.ExcelFilePath
$SheetName = $Config.ExcelSheetName
$OutputCsvPath = "$CurrentFolder\$($Config.CSVFile_users)"

# Start Excel application
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Open the workbook
$workbook = $excel.Workbooks.Open($ExcelFilePath)

# Get the worksheet by name
$worksheet = $workbook.Sheets.Item($SheetName)
Write-Host "worksheet.name: $($worksheet.name)"

if (-not $worksheet) {
    Write-Error "Worksheet '$SheetName' not found."
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

# Copy the worksheet to a new workbook
$tempWorkbook = $excel.Workbooks.Add()
$worksheet.Copy($tempWorkbook.WorkSheets.Item(1))

# Save as CSV (Excel uses Windows-1252 encoding by default)
$tempCsv = "$env:TEMP\temp_output.csv"
$tempWorkbook.SaveAs($tempCsv, 6)  # 6 = xlCSV

# Close workbooks
$tempWorkbook.Close($false)
# $workbook.Close($false)
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# Re-encode CSV to UTF-8
Get-Content -Path $tempCsv | Set-Content -Encoding UTF8 -Path $OutputCsvPath
# Remove-Item $tempCsv

Write-Host "Exported '$SheetName' from '$ExcelFilePath' to '$OutputCsvPath' as UTF-8 CSV."
Read-Host "Press Enter to exit"