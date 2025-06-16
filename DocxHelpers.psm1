function Replace-VariablesInDocx {
    param (
        $objSelection,
        [hashtable]$VariableMap,
        [string]$VarMark = '$'
    )
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
        $FindText = "$VarMark{$Variable}"
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
}

function Get-VariablesFromDocx {
    param (
        $objDoc,
        [string]$VarMark = '$'
    )
    $text = $objDoc.Content.Text
    $var_matches = [regex]::Matches($text, "\$($VarMark)\{.*?\}")
    # Put all matches into an array
    $vars_array = @()
    foreach ($match in $var_matches) {
        $cleaned = $match.Value -replace "^\$VarMark\{", '' -replace "\}$VarMark", ''
        $vars_array += $cleaned
    }
    $vars_array = $vars_array | Sort-Object -Unique
    return $vars_array
}

function Get-WordObject {
    param (
        [string]$FilePath
    )
    $objWord = New-Object -ComObject word.application
    $objWord.Visible = $False
    $objDoc = $objWord.Documents.Open($FilePath)
    return $objWord, $objDoc
}