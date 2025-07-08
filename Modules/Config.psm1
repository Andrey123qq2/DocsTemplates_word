function Get-Configuration {
    param (
        [string]$Folder,
        [string]$File
    )
    $configPath = Join-Path $Folder $File
    Test-FileExistence @($configPath)
    return Get-Content $configPath | ConvertFrom-Json
}

function Initialize-Destination {
    param (
        [string]$BaseFolder,
        [string]$SubFolder
    )
    $fullPath = Join-Path $BaseFolder $SubFolder
    if (-not (Test-Path $fullPath)) {
        New-Item -ItemType Directory -Path $fullPath | Out-Null
    }
    return $fullPath
}
