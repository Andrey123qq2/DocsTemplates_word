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
        [string]$VarMark = '$',
        [string]$emptyValue = "0"
    )
    $wdFindContinue = 1
    $ReplaceAll = 2
    $highlightMarker = "-–EMPTY_VALUE--"
    $wdYellow = 7
    foreach ($Variable in $VariableMap.Keys) {
        $FindText = "$VarMark{$Variable}"
        $ReplaceWith = $VariableMap[$Variable]
        $ReplaceParts = if ($ReplaceWith.Length -lt 255) {
            if ($ReplaceWith -eq $emptyValue) {
                $ReplaceWith = $highlightMarker
            }
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
            $objSelection.Find.Execute(
                $FindText, $false, $true, $false, $false, $false, $true, 
                $wdFindContinue, $false, $part, $ReplaceAll
            ) | Out-Null
            if ($part -eq $highlightMarker) {
                $objSelection.HomeKey(6) | Out-Null # wdStory = 6, move to start
                $objSelection.Find.ClearFormatting()
                $objSelection.Find.Text = $highlightMarker
                $objSelection.Find.Replacement.ClearFormatting()
                while ($objSelection.Find.Execute()) {
                    if ($objSelection.Text -eq $highlightMarker) {
                        $objSelection.Range.HighlightColorIndex = $wdYellow
                    }
                    $objSelection.Collapse(0) # wdCollapseEnd = 0
                }
            }
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

function Convert-DocxToPdf {
    param (
        [string]$docxPath,
        [string]$pdfPath
    )
    # Create Word COM object
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    try {
        $doc = $word.Documents.Open($docxPath)
        $doc.SaveAs([ref]$pdfPath, [ref]17)  # 17 = wdFormatPDF
        $doc.Close()
        Write-Host " Converted to PDF: $pdfPath"
    } catch {
        Write-Error " Failed to convert document: $_"
    } finally {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-TemplateProcessing {
    param (
        $TemplateSource,
        $DstFolder,
        $UserVars,
        $UserVars2,
        $Config
    )
    if ($TemplateSource.PSIsContainer) {
        $NewFolderName = Resolve-StringPlaceholders -InputString $TemplateSource.Name -VarMap $UserVars
        if ($Config.files_prefix) {
            $NewFolderName = "$($Config.files_prefix)$NewFolderName"
        }
        $NewFolderName = Resolve-StringPlaceholders -InputString $NewFolderName -VarMap $UserVars2[0] -VarMark '$2'
        $dst = Join-Path $DstFolder $NewFolderName
        if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
        Copy-Item $TemplateSource.FullName -Destination $dst -Recurse -Force

        Get-ChildItem -Recurse -Path $dst -Filter *.docx | ForEach-Object {
            $file = $_
            $newFileName = Resolve-StringPlaceholders -InputString $file.Name -VarMap $UserVars
            $i = 1
            foreach ($var in $UserVars2) {
                $i++
                $newFileName = Resolve-StringPlaceholders -InputString $newFileName -VarMap $var -VarMark "`$$i"
            }
            if ($file.Name -ne $newName) {
                Rename-Item -Path $file.FullName -NewName $newFileName
                $file = Get-Item (Join-Path $file.Directory.FullName $newFileName)
            }
            Write-Host "Processing file $($file.FullName)"
            Invoke-WordFileProcessing `
                -FilePath $file.FullName `
                -UserVars $UserVars `
                -UserVars2 $UserVars2 `
                -Descriptions $Config.vars_description

            if ($file.Name -like "*_pdf.docx") {
                Convert-DocxToPdf `
                    -docxPath $file.FullName `
                    -pdfPath (Join-Path $file.Directory.FullName ($file.Name -replace "_pdf.docx", ".pdf"))
                Rename-Item -Path $file.FullName -NewName ($file.Name -replace "_pdf.docx", ".docx")
            }
        }
    } else {
        if ($Config.files_prefix) {
            $dstFileNewName = "$($Config.files_prefix)$($TemplateSource.Name)"
        }
        $dstFileNewName = Resolve-StringPlaceholders -InputString $dstFileNewName -VarMap $UserVars

        foreach ($var in $UserVars2) {
            $i++
            $dstFileNewName = Resolve-StringPlaceholders -InputString $dstFileNewName -VarMap $var -VarMark "`$$i"
        }
        $dst = Join-Path $DstFolder $dstFileNewName
        Copy-Item $TemplateSource.FullName -Destination $dst -Force
        Write-Host "Processing file $dst"
        Invoke-WordFileProcessing `
            -FilePath $dst `
            -UserVars $UserVars `
            -UserVars2 $UserVars2 `
            -Descriptions $Config.vars_description

        if ($dst.Name -like "*_pdf.docx") {
            Convert-DocxToPdf `
                -docxPath $file.FullName `
                -pdfPath (Join-Path $dst.Directory.FullName ($dst.Name -replace "_pdf.docx", ".pdf"))
            Rename-Item -Path $dst.FullName -NewName ($dst.Name -replace "_pdf.docx", ".docx")
        }
    }
}

function Invoke-WordFileProcessing {
    param (
        [string]$FilePath,
        [hashtable]$UserVars,
        [array]$UserVars2,
        $Descriptions
    )
    $word, $doc = Get-WordObject -FilePath $FilePath
    $selection = $word.Selection
    $UserVars = Request-MissingVariables -objDoc $doc -VariableMap $UserVars -Descriptions $Descriptions
    Update-VarsPlaceholdersInDocx -VariableMap $UserVars -objSelection $selection
    if ($UserVars2.Count -gt 0) {
        $i = 1
        foreach ($usrVars in $UserVars2) {
            $i++
            Update-VarsPlaceholdersInDocx -VariableMap $usrVars -VarMark "`$$i" -objSelection $selection    
        }
    }
    $doc.Save()
    $doc.Close()
    $word.Quit()
}