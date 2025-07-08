$ConfigFile = "Config_all.json"

function Get-AdditionalVariables {
    param (
        $objSelection,
        [hashtable]$VariableMap,
        $Descriptions
    )

    $doc_vars = Get-VariablesFromDocx -objDoc $objDoc
    $doc_vars_unique = $doc_vars | Where-Object { $_ -notin $VariableMap.Keys }
    $vars_description_names = $(Get-Member -InputObject $Descriptions -MemberType NoteProperty).Name

    foreach ($var in $doc_vars_unique) {
        if (-not $VariableMap.ContainsKey($var)) {
            # Prompt and store
            if ($var -in $vars_description_names) {
                $var_descr = $Config.vars_description | Select-Object -ExpandProperty $var
            } else {
                $var_descr = $var
            }
            $value = Read-Host "$($var_descr)"
            if (-not $value) {
                if ($var.Contains("date_string") -or $var.Contains("date_month_string")) {
                    $value = (Get-Date).ToString("dd MMMM yyyy")
                    Write-Host "-- Using current date value: $value"
                } elseif ($var.Contains("date")) {
                    $value = (Get-Date).ToString("dd.MM.yyyy")
                    Write-Host "-- Using current date value: $value"
                } else {
                    Write-Warning "Variable '$var' is not set. Skipping."
                }
            }
            $VariableMap[$var] = $value
        }
    }
    return $VariableMap
}

function Update-WordFile{
    param(
        [string]$FilePath,
        $VariableMap,
        $VariableMap_2,
        $surname_2
    )
    $objects = Get-WordObject -FilePath $FilePath
    $objWord = $objects[0]
    $objDoc = $objects[1]
    $objSelection = $objWord.Selection

    $VariableMap = Get-AdditionalVariables -objDoc $objDoc `
        -VariableMap $VariableMap `
        -Descriptions $Config.vars_description

    Replace-VariablesInDocx -VariableMap $VariableMap -objSelection $objSelection
    if ($surname_2) {
        Replace-VariablesInDocx -VariableMap $VariableMap_2 -VarMark '$2' -objSelection $objSelection
    }
    $objDoc.save()
    $objDoc.close()
    $objWord.Quit()
}

# --- Main ---
$CurrentFolder = (Split-Path $MyInvocation.MyCommand.Path -Parent)
Import-Module "$CurrentFolder\Utils.psm1" -Force
Import-Module "$CurrentFolder\DocxHelpers.psm1" -Force
$Config = Get-Config -Folder $CurrentFolder -ConfigFile $ConfigFile

$DstPath = "$CurrentFolder\$($Config.DstFolder)"
if (-not (Test-Path $DstPath)) {
    Write-Information "Dst Folder not found."
    New-Item -ItemType Directory -Path $DstPath
}

$CSVFile_users = "$CurrentFolder\$($Config.CSVFile_users)"
Validate-Files @($CSVFile_users)
$CSVFile_users_Content = Import-Csv -Delimiter ';' -Path $CSVFile_users -Encoding 'UTF8'

$TemplateSource = Select-TemplateSource -BasePath "$CurrentFolder\$($Config.TemplatesFolder)"
$IsFolder = $TemplateSource.PSIsContainer
$TemplateSourcePath = $TemplateSource.FullName

Validate-Files @($TemplateSourcePath)

$surnames_input = Read-Host $Config.Prompt_csv_keyfield
$surnames = $surnames_input -split '\s*,\s*'
$surname_2 = Read-Host $Config.Prompt_csv_keyfield_2

$VariableMap = @{}
$VariableMap_2 = @{}
foreach ($surname in $surnames) {
    $VariableMap["Surname"] = $surname
    Write-Host "`nProcessing $($VariableMap.Surname)"

    $user_row = $CSVFile_users_Content | Where-Object { $_.Surname -eq $surname }
    if (-not $user_row) {
        Write-Warning "Row '$surname' is not found in CSV file. Continue."
        continue
    }
    if ($user_row.GetType().BaseType.Name -eq 'Array') {
        Write-Warning "Multiple rows found for '$surname'."
        $user_row = Get-ItemFromArray -items $user_row -ItemType "user"
    }
    $VariableMap = Convert-CSVToHashtable -csv_obj $user_row -VariableMap $VariableMap

    if ($surname_2) {
        $VariableMap_2["Surname"] = $surname_2
        $user_2_row = $CSVFile_users_Content | Where-Object { $_.Surname -eq $surname_2 }
        $VariableMap_2 = Convert-CSVToHashtable -csv_obj $user_2_row -VariableMap $VariableMap_2
    }
    
    if ($IsFolder) {
        $UserDstFolder = "$DstPath\$($TemplateSource.Name)_$surname"
        if (Test-Path $UserDstFolder) {
            Remove-Item -Path $UserDstFolder -Force -Recurse
        }
        Copy-Item -Recurse -Path $TemplateSourcePath -Destination $UserDstFolder -Force

        $WordFiles = Get-ChildItem -Recurse -Path $UserDstFolder -Filter *.docx
        foreach ($file in $WordFiles) {
            $NewFileName = $file.Name.Replace("`${$($Config.FileNameReplaceVar)}", $VariableMap.Surname)
            if ($file.Name -ne $NewFileName) {
                Rename-Item -Path $file.FullName -NewName $NewFileName
                $file = Get-Item "$($file.Directory.FullName)\$NewFileName"
            }
            Write-Host "Processing file $($file.FullName)"
            Update-WordFile -FilePath $file.FullName -VariableMap $VariableMap `
                -VariableMap_2 $VariableMap_2 -surname_2 $surname_2
        }

    } else {
        $FileNameNew = "$CurrentFolder\$($Config.DstFolder)\$TemplateSource".`
            Replace("`${$($Config.FileNameReplaceVar)}", $VariableMap.Surname)
        Copy-Item $TemplateSourcePath -Destination $FileNameNew -Verbose
        Write-Host "Processing file $FileNameNew"
        Update-WordFile -FilePath $FileNameNew -VariableMap $VariableMap `
            -VariableMap_2 $VariableMap_2 -surname_2 $surname_2
    }
}

Read-Host "Press Enter to exit"