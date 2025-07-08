function Import-UsersCsv {
    param (
        [string]$Path
    )
    return Import-Csv -Delimiter ';' -Path $Path -Encoding 'UTF8'
}

function Find-UserRow {
    param (
        [array]$Users,
        [string]$Surname
    )
    $row = $Users | Where-Object { $_.Surname -eq $Surname }
    if (-not $row) {
        Write-Warning "User '$Surname' not found in CSV. Skipping."
        return $null
    }
    if ($row.Count -gt 1) {
        Write-Warning "Multiple users found for '$Surname'."
        return Select-FromList -Items $row -ItemType "user"
    }
    return $row
}

function New-VariableMap {
    param (
        $User,
        [hashtable]$Defaults = @{}
    )
    foreach ($prop in $User.PSObject.Properties) {
        $Defaults[$prop.Name] = $prop.Value
    }
    return $Defaults
}
