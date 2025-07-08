function Select-TemplateSource {
    param (
        [string]$BasePath
    )
    $items = Get-ChildItem -Path $BasePath
    if ($items.Count -eq 0) {
        throw "No templates found in $BasePath"
    }

    $items = $items | Sort-Object Name

    Write-Host "Choose a template (file or folder):"
    for ($i = 0; $i -lt $items.Count; $i++) {
        $type = if ($items[$i].PSIsContainer) { "Folder" } else { "File" }
        Write-Host "$($i + 1). $($items[$i].Name) [$type]"
    }

    $choice = Read-Host "Enter number of template to use"
    return $items[$choice - 1]
}


function Get-ItemFromArray {
    param (
        [array]$items,
        [string]$ItemType = "user",
        [Array]$Properties = @("Surname", "Name", "LastName")
    )
    Write-Host "Choose a $($ItemType):"
    for ($i = 0; $i -lt $items.Count; $i++) {
        Write-Host "$($i + 1). $($items[$i] | Select-Object -Property $Properties | foreach-object {
            $_.psobject.Properties.Value -join ' ' 
        })"
    }

    $choice = Read-Host "Enter number of $($ItemType) to use"
    return $items[$choice - 1] 
}

function Validate-Files {
    param (
        [Array]$FilePathes
    )
    foreach ($path in $FilePathes) {
        if (-Not (Test-Path $path)) {
            Write-Error "$path file not found."
            exit 1
        }
    }
}

function Get-Config {
    param (
        [string]$Folder,
        [string]$ConfigFile
    )
    $ConfigFile = "$Folder\$ConfigFile"
    Validate-Files @($ConfigFile)
    return Get-Content $ConfigFile | ConvertFrom-Json
}

function Convert-CSVToHashtable {
    param(
        $csv_obj,
        $VariableMap        
    )
    foreach ($prop in $csv_obj.psobject.properties) {
        $VariableMap[$prop.Name] = $prop.Value
    }
    return $VariableMap
}