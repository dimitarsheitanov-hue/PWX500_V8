param(
    [Parameter(Mandatory=$true)][string]$ProductsCsv,
    [Parameter(Mandatory=$true)][string]$IndexCsv,
    [Parameter(Mandatory=$true)][string]$ResultsCsv,
    [Parameter(Mandatory=$true)][string]$ProgressFile,
    [string]$CategoriesCsv = ''
)

$ErrorActionPreference = 'Stop'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Write-ProgressText([string]$Text) {
    try { [System.IO.File]::WriteAllText($ProgressFile, $Text, $utf8NoBom) } catch {}
}
function CsvField([string]$s) {
    if ($null -eq $s) { $s = '' }
    if ($s.Contains('"')) { $s = $s.Replace('"','""') }
    if ($s.Contains(';') -or $s.Contains('"') -or $s.Contains("`r") -or $s.Contains("`n")) { return '"' + $s + '"' }
    return $s
}
function Normalize([string]$Value) {
    if ($null -eq $Value) { return '' }
    $s = $Value.Trim().ToLowerInvariant()
    $s = [regex]::Replace($s, '(?i)(\d+)\s*[/x*]\s*(\d+)', '$1x$2')
    # Preserve indexed variants exactly: F != F#, B != B#, etc.
    $s = [regex]::Replace($s, '(?i)([a-z])#', '$1hash')
    $s = $s.Replace('#', ' hash ')
    $s = $s.Replace('_',' ').Replace('-',' ').Replace('\\',' ').Replace('/',' ')
    $s = [regex]::Replace($s, '[^a-z0-9а-яё ]', ' ')
    $s = [regex]::Replace($s, '\s+', ' ')
    return $s.Trim()
}
function TokensFromNorm([string]$Norm) {
    if ([string]::IsNullOrWhiteSpace($Norm)) { return @() }
    return @($Norm -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}
function Is-Dim([string]$t) { return (-not [string]::IsNullOrWhiteSpace($t) -and $t -match '^\d{2,4}x\d{2,4}$') }
function Is-Variant([string]$t) {
    if ([string]::IsNullOrWhiteSpace($t)) { return $false }
    return @('f','b','c','l','r','a','fhash','bhash','chash','lhash','rhash','ahash','hash') -contains $t.ToLowerInvariant()
}
function Is-Generic([string]$t) {
    if ([string]::IsNullOrWhiteSpace($t)) { return $true }
    return @('and','with','set','furniture','bed','bedroom','hallway','night','stand','dresser','wardrobe','cabinet','section','coffee','dining','table','desk','shelf','bottom','top','upgrade','drawer','shoes') -contains $t.ToLowerInvariant()
}
function Unique-Array($items) {
    $seen = @{}
    $out = New-Object System.Collections.ArrayList
    foreach ($x in @($items)) {
        $k = ([string]$x).ToLowerInvariant()
        if (-not $seen.ContainsKey($k)) { $seen[$k]=$true; [void]$out.Add([string]$x) }
    }
    return @($out.ToArray())
}

# V5.1.13: category mapping is a HARD safety filter, loaded from external CSV.
function Load-CategoryRules([string]$Path) {
    $out=New-Object System.Collections.ArrayList
    if([string]::IsNullOrWhiteSpace($Path) -or -not [System.IO.File]::Exists($Path)){return @()}
    foreach($r in @(Import-Csv -LiteralPath $Path -Delimiter ';')){
        $phrase=([string]$r.Phrase).Trim();$targets=([string]$r.Targets).Trim();$active=([string]$r.Active).Trim();$prio=100
        [void][int]::TryParse(([string]$r.Priority),[ref]$prio)
        if([string]::IsNullOrWhiteSpace($phrase) -or [string]::IsNullOrWhiteSpace($targets)){continue}
        if($active -in @('0','false','False','FALSE','no','NO')){continue}
        [void]$out.Add([pscustomobject]@{Phrase=(Normalize $phrase);Targets=@($targets -split '\|' | ForEach-Object {$_.Trim()} | Where-Object {-not [string]::IsNullOrWhiteSpace($_)});Priority=$prio})
    }
    return @($out | Sort-Object @{Expression='Priority';Descending=$true},@{Expression={$_.Phrase.Length};Descending=$true})
}
function Get-CategoryTargets([string]$ProductNorm,$Rules){
    foreach($r in @($Rules)){
        $p=[string]$r.Phrase
        if($ProductNorm -eq $p -or $ProductNorm.StartsWith($p+' ')){return @($r.Targets)}
    }
    return @()
}
function Path-MatchesTargets([string]$Path,$Targets){
    if($null -eq $Targets -or @($Targets).Count -eq 0){return $false}
    $parts=@(([string]$Path).ToLowerInvariant() -split '\\')
    foreach($t0 in @($Targets)){
        $t=([string]$t0).Trim().ToLowerInvariant();if([string]::IsNullOrWhiteSpace($t)){continue}
        foreach($part in $parts){if(([string]$part).Contains($t)){return $true}}
    }
    return $false
}
function Build-Meta([string]$Folder, [string]$Path) {
    $norm = Normalize $Folder
    $tok = @(TokensFromNorm $norm)
    $core = New-Object System.Collections.ArrayList
    $vars = New-Object System.Collections.ArrayList
    $dims = New-Object System.Collections.ArrayList
    foreach ($t in $tok) {
        if (Is-Dim $t) { [void]$dims.Add($t); continue }
        if (Is-Variant $t) { [void]$vars.Add($t); continue }
        if (Is-Generic $t) { continue }
        if ($t -match '^\d+$') { continue }
        if ($t.Length -lt 2) { continue }
        [void]$core.Add($t)
    }
    return [pscustomobject]@{
        Folder = $Folder
        Path = $Path
        Norm = $norm
        Core = @(Unique-Array $core)
        Variants = @(Unique-Array $vars)
        Dims = @(Unique-Array $dims)
    }
}
function Array-ContainsCI($arr, [string]$value) {
    foreach ($x in @($arr)) { if ([string]::Equals([string]$x, $value, [System.StringComparison]::OrdinalIgnoreCase)) { return $true } }
    return $false
}
function Score-Meta([string]$ProductNorm, $pCore, $pVariants, $pDims, $m) {
    if ($m.Core.Count -eq 0) { return -1 }
    foreach ($t in @($m.Core)) { if (-not (Array-ContainsCI $pCore $t)) { return -1 } }
    if ($m.Variants.Count -gt 0) {
        if ($pVariants.Count -eq 0) { return -1 }
        foreach ($v in @($m.Variants)) { if (-not (Array-ContainsCI $pVariants $v)) { return -1 } }
    }
    if ($m.Dims.Count -gt 0 -and $pDims.Count -gt 0) {
        foreach ($d in @($m.Dims)) { if (-not (Array-ContainsCI $pDims $d)) { return -1 } }
    }
    $score = 25 * $m.Core.Count
    if ($ProductNorm -eq $m.Norm -or $ProductNorm.EndsWith(' ' + $m.Norm)) { $score += 120 }
    foreach ($v in @($m.Variants)) { if (Array-ContainsCI $pVariants $v) { $score += 25 } }
    foreach ($d in @($m.Dims)) { if (Array-ContainsCI $pDims $d) { $score += 30 } }
    $score += [Math]::Min(20, [int]($m.Norm.Length / 3))
    return [int]$score
}

try {
    Write-ProgressText 'Търсене: зареждане на локалния каталог...'
    if (-not [System.IO.File]::Exists($ProductsCsv)) { throw 'Липсва списъкът с изделия от CSV.' }
    if (-not [System.IO.File]::Exists($IndexCsv)) { throw 'Липсва локалният каталог. Първо обновете каталога.' }

    $products = @(Import-Csv -LiteralPath $ProductsCsv -Delimiter ';')
    $catalogRows = @(Import-Csv -LiteralPath $IndexCsv -Delimiter ';')
    $categoryRules = @(Load-CategoryRules $CategoriesCsv)

    # Precompute folder metadata once. This is the main speed improvement in V0.2.7.
    $catalog = New-Object System.Collections.ArrayList
    $tokenIndex = @{}
    $ci = [System.StringComparer]::OrdinalIgnoreCase
    $idx = 0
    foreach ($it in $catalogRows) {
        $folder = ([string]$it.Folder).Trim()
        $path = ([string]$it.Path).Trim()
        if ([string]::IsNullOrWhiteSpace($folder) -or [string]::IsNullOrWhiteSpace($path)) { continue }
        $m = Build-Meta -Folder $folder -Path $path
        [void]$catalog.Add($m)
        $myIndex = $catalog.Count - 1
        foreach ($t in @($m.Core)) {
            $key = ([string]$t).ToLowerInvariant()
            if (-not $tokenIndex.ContainsKey($key)) { $tokenIndex[$key] = New-Object System.Collections.ArrayList }
            [void]$tokenIndex[$key].Add($myIndex)
        }
        $idx++
        if (($idx % 250) -eq 0) { Write-ProgressText ('Търсене: подготвен каталог ' + $idx + ' / ' + $catalogRows.Count + '...') }
    }

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('Product;Candidates;Status;BestScore;Paths')
    $done = 0

    foreach ($row in $products) {
        $product = ([string]$row.Product).Trim()
        if ([string]::IsNullOrWhiteSpace($product)) { continue }
        $pNorm = Normalize $product
        $pTok = @(TokensFromNorm $pNorm)
        $pCoreList = New-Object System.Collections.ArrayList
        $pVarList = New-Object System.Collections.ArrayList
        $pDimList = New-Object System.Collections.ArrayList
        foreach ($t in $pTok) {
            if (Is-Dim $t) { [void]$pDimList.Add($t); continue }
            if (Is-Variant $t) { [void]$pVarList.Add($t); continue }
            if (Is-Generic $t) { continue }
            if ($t -match '^\d+$') { continue }
            if ($t.Length -lt 2) { continue }
            [void]$pCoreList.Add($t)
        }
        $pCore = @(Unique-Array $pCoreList)
        $pVariants = @(Unique-Array $pVarList)
        $pDims = @(Unique-Array $pDimList)
        $categoryTargets = @(Get-CategoryTargets -ProductNorm $pNorm -Rules $categoryRules)
        if($categoryTargets.Count -eq 0){
            $line=(CsvField $product)+';0;'+(CsvField 'НЕЯСНА КАТЕГОРИЯ - ИЗБЕРИ ПАПКА')+';;'
            [void]$lines.Add($line)
            $done++
            if (($done % 10) -eq 0 -or $done -eq $products.Count) { Write-ProgressText ('Търсене: ' + $done + ' / ' + $products.Count + ' изделия... (неясните категории не се търсят автоматично)') }
            continue
        }

        # Candidate pool = posting list for the rarest meaningful product token.
        $pool = @()
        $bestPostingCount = [int]::MaxValue
        foreach ($t in $pCore) {
            $key = ([string]$t).ToLowerInvariant()
            if ($tokenIndex.ContainsKey($key)) {
                $arr = @($tokenIndex[$key])
                if ($arr.Count -lt $bestPostingCount) {
                    $bestPostingCount = $arr.Count
                    $pool = $arr
                }
            }
        }

        $scored = New-Object System.Collections.ArrayList
        foreach ($ciIndex in @($pool)) {
            $m = $catalog[[int]$ciIndex]
            $score = Score-Meta -ProductNorm $pNorm -pCore $pCore -pVariants $pVariants -pDims $pDims -m $m
            if ($score -ge 0 -and (Path-MatchesTargets -Path $m.Path -Targets $categoryTargets)) {
                [void]$scored.Add([pscustomobject]@{Path=$m.Path;Score=[int]$score})
            }
        }

        $sorted = @($scored | Sort-Object -Property @{Expression={$_.Score};Descending=$true}, @{Expression={$_.Path};Descending=$false})
        $keep = @()
        if ($sorted.Count -gt 0) {
            $top = [int]$sorted[0].Score
            $keep = @($sorted | Where-Object { [int]$_.Score -ge ($top - 10) } | Select-Object -First 8)
        }

        $status = 'НЯМА КАНДИДАТ - ПОСОЧИ ПАПКА'
        if ($keep.Count -eq 1) { $status = '1 КАНДИДАТ В ПРАВИЛНА КАТЕГОРИЯ - НУЖНА HOP ПРОВЕРКА' }
        elseif ($keep.Count -gt 1) { $status = 'НЯКОЛКО КАНДИДАТА В ПРАВИЛНА КАТЕГОРИЯ - НУЖНА HOP ПРОВЕРКА' }

        $best = ''
        if ($keep.Count -gt 0) { $best = [string]$keep[0].Score }
        $pathParts = @()
        foreach ($k in $keep) { $pathParts += ('[' + [string]$k.Score + ']' + [string]$k.Path) }
        $paths = $pathParts -join '~|~'
        $line = (CsvField $product) + ';' + (CsvField ([string]$keep.Count)) + ';' + (CsvField $status) + ';' + (CsvField $best) + ';' + (CsvField $paths)
        [void]$lines.Add($line)

        $done++
        if (($done % 10) -eq 0 -or $done -eq $products.Count) {
            Write-ProgressText ('Търсене: ' + $done + ' / ' + $products.Count + ' изделия... (ускорен индекс)')
        }
    }

    [System.IO.File]::WriteAllLines($ResultsCsv, [string[]]$lines.ToArray(), $utf8Bom)
    Write-ProgressText ('Търсенето е готово: ' + $done + ' изделия.')
    exit 0
}
catch {
    $lineNo = ''
    try { $lineNo = [string]$_.InvocationInfo.ScriptLineNumber } catch {}
    $typeName = ''
    try { $typeName = $_.Exception.GetType().FullName } catch {}
    Write-ProgressText ('Търсене ERROR: ' + $_.Exception.Message + ' | ред: ' + $lineNo + ' | тип: ' + $typeName)
    exit 1
}
