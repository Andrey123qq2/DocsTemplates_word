function Get-WordObject {
    param ([string]$FilePath)
    $word = New-Object -ComObject word.application
    $word.Visible = $false
    $doc = $word.Documents.Open($FilePath)
    return $word, $doc
}

function Get-VariablesFromDocx {
    param ($objDoc, [string]$VarMark = '$')
    $text = $objDoc.Content.Text
    $var_matches = [regex]::Matches($text, "\$VarMark\{(.*?)\}")
    return ($var_matches | ForEach-Object { $_.Groups[1].Value }) | Sort-Object -Unique
}

function Update-VarsPlaceholdersInDocx {
    param (
        $objSelection,
        [hashtable]$VariableMap,
        [string]$VarMark = '$'
    )
    $wdFindContinue = 1
    $ReplaceAll = 2
    foreach ($Variable in $VariableMap.Keys) {
        $FindText = "$VarMark{$Variable}"
        $ReplaceWith = $VariableMap[$Variable]
        $ReplaceParts = if ($ReplaceWith.Length -lt 255) {
            @($ReplaceWith)
        } else {
            $chunks = [regex]::Matches($ReplaceWith, ".{1,200}") | ForEach-Object { $_.Value }
            for ($i = 0; $i -lt $chunks.Count; $i++) {
                if ($i -lt $chunks.Count - 1) {
                    $chunks[$i] += $FindText
                }
            }
            $chunks
        }

        foreach ($part in $ReplaceParts) {
            $objSelection.Find.Execute($FindText, $false, $true, $false, $false, $false, $true, $wdFindContinue, $false, $part, $ReplaceAll) | Out-Null
        }
    }
    $objSelection.Find.Execute("     ", $false, $true, $false, $false, $false, $true, $wdFindContinue, $false, " ", $ReplaceAll) | Out-Null
    $objSelection.Find.Execute("    ", $false, $true, $false, $false, $false, $true, $wdFindContinue, $false, " ", $ReplaceAll) | Out-Null
    $objSelection.Find.Execute("   ", $false, $true, $false, $false, $false, $true, $wdFindContinue, $false, " ", $ReplaceAll) | Out-Null
    $objSelection.Find.Execute("  ", $false, $true, $false, $false, $false, $true, $wdFindContinue, $false, " ", $ReplaceAll) | Out-Null
}

function Request-MissingVariables {
    param (
        $objDoc,
        [hashtable]$VariableMap,
        $Descriptions
    )
    $docVars = Get-VariablesFromDocx -objDoc $objDoc
    $missingVars = $docVars | Where-Object { -not $VariableMap.ContainsKey($_) }
    foreach ($var in $missingVars) {
        $desc = if ($Descriptions.PSObject.Properties.Name -contains $var) {
            $Descriptions.$var
        } else {
            $var
        }
        $value = Read-Host "$desc"
        if (-not $value) {
            if ($var -like "*date_string*" -or $var -like "*date_month_string*") {
                $value = (Get-Date).ToString("dd MMMM yyyy")
                Write-Host "-- Using current date value: $value"
            } elseif ($var -like "*date*") {
                $value = (Get-Date).ToString("dd.MM.yyyy")
                Write-Host "-- Using current date value: $value"
            } else {
                Write-Warning "Variable '$var' not set. Skipping."
                continue
            }
        }
        $VariableMap[$var] = $value
    }
    return $VariableMap
}

function Invoke-TemplateProcessing {
    param (
        $TemplateSource,
        $DstFolder,
        $UserVars,
        $UserVars2,
        $Surname2,
        $Config
    )
    if ($TemplateSource.PSIsContainer) {
        $NewFolderName = Resolve-StringPlaceholders -InputString $TemplateSource.Name -VarMap $UserVars
        $NewFolderName = Resolve-StringPlaceholders -InputString $NewFolderName -VarMap $UserVars2 -VarMark '$2'
        $dst = Join-Path $DstFolder $NewFolderName
        if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
        Copy-Item $TemplateSource.FullName -Destination $dst -Recurse -Force

        Get-ChildItem -Recurse -Path $dst -Filter *.docx | ForEach-Object {
            $file = $_
            $newFileName = Resolve-StringPlaceholders -InputString $file.Name -VarMap $UserVars
            $newFileName = Resolve-StringPlaceholders -InputString $newFileName -VarMap $UserVars2 -VarMark '$2'
            if ($file.Name -ne $newName) {
                Rename-Item -Path $file.FullName -NewName $newFileName
                $file = Get-Item (Join-Path $file.Directory.FullName $newFileName)
            }
            Write-Host "Processing file $($file.FullName)"
            Invoke-WordFileProcessing -FilePath $file.FullName -UserVars $UserVars -UserVars2 $UserVars2 -Surname2 $Surname2 -Descriptions $Config.vars_description
        }
    } else {
        $dstFileNewName = Resolve-StringPlaceholders -InputString $TemplateSource.Name -VarMap $UserVars
        $dstFileNewName = Resolve-StringPlaceholders -InputString $dstFileNewName -VarMap $UserVars2 -VarMark '$2'
        $dst = Join-Path $DstFolder $dstFileNewName
        Copy-Item $TemplateSource.FullName -Destination $dst -Force
        Write-Host "Processing file $dst"
        Invoke-WordFileProcessing -FilePath $dst -UserVars $UserVars -UserVars2 $UserVars2 -Surname2 $Surname2 -Descriptions $Config.vars_description
    }
}

function Invoke-WordFileProcessing {
    param (
        [string]$FilePath,
        [hashtable]$UserVars,
        [hashtable]$UserVars2,
        [string]$Surname2,
        $Descriptions
    )
    $word, $doc = Get-WordObject -FilePath $FilePath
    $selection = $word.Selection
    $UserVars = Request-MissingVariables -objDoc $doc -VariableMap $UserVars -Descriptions $Descriptions
    Update-VarsPlaceholdersInDocx -VariableMap $UserVars -objSelection $selection
    if ($Surname2) {
        Update-VarsPlaceholdersInDocx -VariableMap $UserVars2 -VarMark '$2' -objSelection $selection
    }
    $doc.Save()
    $doc.Close()
    $word.Quit()
}
