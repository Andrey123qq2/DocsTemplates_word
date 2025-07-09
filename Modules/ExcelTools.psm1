function Export-ExcelSheetToCsv {
    param (
        [string]$ExcelFilePath,
        [string]$SheetName,
        [string]$OutputCsvPath
    )

    $tempCsvPath = "$env:TEMP\temp_output.csv"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Open($ExcelFilePath, [ref]0, [ref]$true)
        $worksheet = $workbook.Sheets.Item($SheetName)

        if (-not $worksheet) {
            throw "Worksheet '$SheetName' not found."
        }

        Write-Host "Worksheet: $($worksheet.Name)"

        # Copy worksheet to a new workbook
        $tempWorkbook = $excel.Workbooks.Add()
        $worksheet.Copy($tempWorkbook.WorkSheets.Item(1))

        # Save copied sheet as CSV
        $tempWorkbook.SaveAs($tempCsvPath, 6)  # 6 = xlCSV

        # Re-encode to UTF-8
        Get-Content -Path $tempCsvPath | Set-Content -Encoding UTF8 -Path $OutputCsvPath

        Write-Host "Exported '$SheetName' from '$ExcelFilePath' to '$OutputCsvPath' as UTF-8 CSV."
    }
    catch {
        Write-Error $_
    }
    finally {
        if ($tempWorkbook) { $tempWorkbook.Close($false) }
        if ($workbook)     { $workbook.Close($false) }
        if ($excel)        { $excel.Quit() }

        # Clean COM objects
        foreach ($comObj in @($worksheet, $workbook, $tempWorkbook, $excel)) {
            if ($comObj) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($comObj) | Out-Null
            }
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if (Test-Path $tempCsvPath) {
            Remove-Item $tempCsvPath -Force
        }
    }
}
