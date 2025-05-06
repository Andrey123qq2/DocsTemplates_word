$ConfigFile = "Config_all.json"

function Select-WordFile {
    param (
        [string]$Directory = $PSScriptRoot  # Default to the script's directory
    )
    # Get list of .word files in the specified directory
    $wordFiles = Get-ChildItem -Path $Directory -Filter "*.docx"

    # Check if any word files are found
    if ($wordFiles.Count -eq 0) {
        Write-Host "No .docx files found in the directory: $Directory"
        return $null
    }

    # Display the list of .docx files
    Write-Host "Select a Word file:"
    for ($i = 0; $i -lt $wordFiles.Count; $i++) {
        Write-Host "$($i + 1): $($wordFiles[$i].Name)"
    }

    # Ask the user to choose a file
    $choice = Read-Host "Enter the number of the file"

    # Validate the choice
    if ($choice -match '^\d+$' -and $choice -ge 1 -and $choice -le $wordFiles.Count) {
        $selectedFile = $wordFiles[$choice - 1].Name
        Write-Host "You selected: $selectedFile"
        return $selectedFile
    } else {
        Write-Host "Invalid selection. Please try again."
        return $null
    }
}

# Define function to replace variables in a .docx file
function Replace-VariablesInDocx {
    param (
        [string]$FilePath,
        [hashtable]$VariableMap
    )
    $objWord = New-Object -ComObject word.application
    $objWord.Visible = $False
    $objDoc = $objWord.Documents.Open($FilePath)
    $objSelection = $objWord.Selection
    
    $MatchCase = $false
    $MatchWholeWord = $true
    $MatchWildcards = $false
    $MatchSoundsLike = $false
    $MatchAllWordForms = $false
    $Forward = $true
    $wrap = $wdFindContinue
    $wdFindContinue = 1
    $Format = $false
    $ReplaceAll = 2

    foreach ($Variable in $VariableMap.Keys) {
        $FindText = "`${$Variable}"
        $ReplaceWith = $VariableMap[$Variable]

        # Split the variable into manageable parts if necessary
        $ReplaceWithParts = @()
        if ($ReplaceWith.Length -lt 255) {
            # If the variable is within the allowed limit, add directly
            $ReplaceWithParts = @("$ReplaceWith")
        } else {
            # Split into chunks of 255 characters or less
            $Chunks = [regex]::Matches($ReplaceWith, '.{1,200}').Value
            $i = 0
            foreach ($Chunk in $Chunks) {
                $i++
                if ($i -lt $Chunks.Length) {
                    $ReplaceWithParts += "$Chunk$FindText"
                } else {
                    $ReplaceWithParts += "$Chunk"
                }
            }
        }
        # Execute find/replace for each part
        foreach ($ReplacePart in $ReplaceWithParts) {
            $objSelection.Find.Execute(
                $FindText, 
                $MatchCase, 
                $MatchWholeWord, 
                $MatchWildcards, 
                $MatchSoundsLike, 
                $MatchAllWordForms, 
                $Forward, 
                $wrap, 
                $Format, 
                $ReplacePart, 
                $ReplaceAll
            ) |  Out-Null
        }
        $objSelection.Find.Execute("     ", $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $wrap, $Format, " ", $ReplaceAll) |  Out-Null
        $objSelection.Find.Execute("    ", $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $wrap, $Format, " ", $ReplaceAll) |  Out-Null
        $objSelection.Find.Execute("   ", $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $wrap, $Format, " ", $ReplaceAll) |  Out-Null
        $objSelection.Find.Execute("  ", $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $wrap, $Format, " ", $ReplaceAll) |  Out-Null
    }

    $objDoc.save()
    $objDoc.close()
    $objWord.Quit()
}

function Get-VariablesFromDocx {
    param (
        [string]$FilePath
    )
    $objWord = New-Object -ComObject word.application
    $objWord.Visible = $False
    $objDoc = $objWord.Documents.Open($FilePath)
    # $objSelection = $objWord.Selection
    $text = $objDoc.Content.Text
    $objDoc.close()
    $objWord.Quit()

    $var_matches = [regex]::Matches($text, '\$\{.*?\}')

    # Put all matches into an array
    $vars_array = @()
    foreach ($match in $var_matches) {
        $cleaned = $match.Value -replace '^\$\{', '' -replace '\}$', ''
        $vars_array += $cleaned
    }
    $vars_array = $vars_array | Sort-Object -Unique
    return $vars_array
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

# Get Config
$CurrentFolder = (Split-Path $MyInvocation.MyCommand.Path -Parent)
$ConfigFile = "$CurrentFolder\$ConfigFile"
Validate-Files @($ConfigFile)
$Config = Get-Content $ConfigFile | ConvertFrom-Json

# Get and validate files pathes
$CSVFile_users = "$CurrentFolder\$($Config.CSVFile_users)"
$TemplateFile = Select-WordFile -Directory "$CurrentFolder\$($Config.TemplatesFolder)"
$TemplateFilePath = "$CurrentFolder\$($Config.TemplatesFolder)\$TemplateFile"
Validate-Files @($CSVFile_users, $TemplateFilePath)
if (-Not (Test-Path "$CurrentFolder\$($Config.DstFolder)")) {
    Write-Information "Dst Folder not found."
    New-Item -ItemType Directory -Path "$CurrentFolder\$($Config.DstFolder)"
}

# Read CSV users file
$CSVFile_users_Content = Import-Csv -Delimiter ';' -Path $CSVFile_users -Encoding 'UTF8'

$VariableMap = @{}
$SharedVariables = @{}

# Ask for surnames (comma separated)
$surnames_input = Read-Host $Config.Prompt_csv_keyfield
$surnames = $surnames_input -split '\s*,\s*'

foreach ($surname in $surnames) {
    $VariableMap = @{}
    $VariableMap["Surname"] = $surname

    $user_row = $CSVFile_users_Content | Where-Object { $_.Surname -eq $surname }
    if ($user_row -eq $null) {
        # Write-Warning "Користувача з прізвищем '$surname' не знайдено в CSV. Пропуск."
        Write-Warning "Row $($VariableMap["Surname"]) is not found in CSV file. Continue"
        continue
    }

    foreach ($prop in $user_row.psobject.properties) {
        $VariableMap[$prop.Name] = $prop.Value
    }

    $DstFilePath = "$CurrentFolder\$($Config.DstFolder)\$TemplateFile"
    $FileNameNew = $DstFilePath.Replace($Config.FileNameReplaceVar, $VariableMap.Surname)

    Copy-Item $TemplateFilePath -Destination $FileNameNew -Verbose
    $doc_vars = Get-VariablesFromDocx -FilePath $FileNameNew
    $doc_vars_unique = $doc_vars | Where-Object { $_ -notin $VariableMap.Keys }

    Write-Host "`nProcessing $surname"
    $vars_description_names = $(Get-Member -InputObject $Config.vars_description -MemberType NoteProperty).Name

    foreach ($var in $doc_vars_unique) {
        # Check if already entered
        if ($SharedVariables.ContainsKey($var)) {
            $VariableMap[$var] = $SharedVariables[$var]
        } else {
            # Prompt and store
            if ($var -in $vars_description_names) {
                $var_descr = $Config.vars_description | Select-Object -ExpandProperty $var
            } else {
                $var_descr = $var
            }
            $value = Read-Host "$($var_descr)"
            $VariableMap[$var] = $value
            $SharedVariables[$var] = $value  # Save for reuse
        }
    }

    Write-Output "`nGenerating file: $FileNameNew"
    Replace-VariablesInDocx -FilePath $FileNameNew -VariableMap $VariableMap
}

Read-Host "Press Enter to exit"