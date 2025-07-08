function Read-ListFromInput {
    param ([string]$Prompt)
    $inputData = Read-Host $Prompt
    return $inputData -split '\s*,\s*'
}

function Test-FileExistence {
    param ([array]$FilePaths)
    foreach ($path in $FilePaths) {
        if (-not (Test-Path $path)) {
            Write-Error "$path not found."
            exit 1
        }
    }
}

function Select-FromList {
    param (
        [array]$Items,
        [string]$ItemType = "item",
        [array]$Properties = @("Surname", "Name", "LastName")
    )
    Write-Host "Choose a $($ItemType):"
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $line = ($Items[$i] | Select-Object -Property $Properties | ForEach-Object {
            $_.PSObject.Properties.Value -join ' '
        })
        Write-Host "$($i + 1). $line"
    }
    $choice = Read-Host "Enter number of $ItemType to use"
    return $Items[$choice - 1]
}

function Resolve-StringPlaceholders {
    param (
        [string]$InputString,
        [hashtable]$VarMap,
        [string]$VarMark = "$"  # e.g. '$' or '$2'
    )

    $escapedMark = [regex]::Escape($VarMark)
    $pattern = "$escapedMark\{(.*?)\}"

    # Get all matches of the pattern
    $matches_all = [regex]::Matches($InputString, $pattern)

    foreach ($match in $matches_all) {
        $placeholder = $match.Value
        $key = $match.Groups[1].Value

        if ($VarMap.ContainsKey($key)) {
            $value = $VarMap[$key]
            $InputString = $InputString -replace [regex]::Escape("$placeholder"), $value
        }
    }

    return $InputString
}
