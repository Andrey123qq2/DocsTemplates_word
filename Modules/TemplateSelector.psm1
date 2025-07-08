function Select-TemplateSource {
    param ([string]$BasePath)

    $items = Get-ChildItem -Path $BasePath | Sort-Object Name
    if ($items.Count -eq 0) {
        throw "No templates found in $BasePath"
    }

    Write-Host "Choose a template (file or folder):"
    for ($i = 0; $i -lt $items.Count; $i++) {
        $type = if ($items[$i].PSIsContainer) { "Folder" } else { "File" }
        Write-Host "$($i + 1). $($items[$i].Name) [$type]"
    }

    $choice = Read-Host "Enter number of template to use"
    return $items[$choice - 1]
}