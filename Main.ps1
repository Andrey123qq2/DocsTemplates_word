$CurrentFolder = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load Modules
Import-Module "$CurrentFolder\Modules\Config.psm1" -Force
Import-Module "$CurrentFolder\Modules\CsvProcessor.psm1" -Force
Import-Module "$CurrentFolder\Modules\TemplateSelector.psm1" -Force
Import-Module "$CurrentFolder\Modules\WordProcessor.psm1" -Force
Import-Module "$CurrentFolder\Modules\Helpers.psm1" -Force

# Load Configuration
$ConfigFile = "Config_all.json"
$Config = Get-Configuration -Folder $CurrentFolder -File $ConfigFile

# Prepare environment
$DstPath = Initialize-Destination -BaseFolder $CurrentFolder -SubFolder $Config.DstFolder
Test-FileExistence @("$CurrentFolder\$($Config.CSVFile_users)")

# Load Data
$Users = Import-UsersCsv -Path "$CurrentFolder\$($Config.CSVFile_users)"
$TemplateSource = Select-TemplateSource -BasePath "$CurrentFolder\$($Config.TemplatesFolder)"

# Input
$surnames = Read-ListFromInput $Config.Prompt_csv_keyfield
$surname_2 = Read-Host $Config.Prompt_csv_keyfield_2

foreach ($surname in $surnames) {
    Write-Host "`nProcessing $surname"

    $user_row = Find-UserRow -Users $Users -Surname $surname
    if (-not $user_row) { continue }

    $VariableMap = New-VariableMap -User $user_row -Defaults @{ Surname = $surname }

    if ($surname_2) {
        $user2_row = Find-UserRow -Users $Users -Surname $surname_2
        $VariableMap_2 = New-VariableMap -User $user2_row -Defaults @{ Surname = $surname_2 }
    }

    Invoke-TemplateProcessing -TemplateSource $TemplateSource `
                      -DstFolder $DstPath `
                      -UserVars $VariableMap `
                      -UserVars2 $VariableMap_2 `
                      -Surname2 $surname_2 `
                      -Config $Config
}

Read-Host "Press Enter to exit"
