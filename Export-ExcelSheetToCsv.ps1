$ConfigFile = "Config_all.json"
$CurrentFolder = (Split-Path $MyInvocation.MyCommand.Path -Parent)

Import-Module "$CurrentFolder\Modules\Config.psm1" -Force
Import-Module "$CurrentFolder\Modules\ExcelTools.psm1" -Force
Import-Module "$CurrentFolder\Modules\Helpers.psm1" -Force

$Config = Get-Configuration -Folder $CurrentFolder -File $ConfigFile

Export-ExcelSheetToCsv `
    -ExcelFilePath $Config.ExcelFilePath `
    -SheetName $Config.ExcelSheetName `
    -OutputCsvPath "$CurrentFolder\$($Config.CSVFile_users)"

Read-Host "Press Enter to exit"