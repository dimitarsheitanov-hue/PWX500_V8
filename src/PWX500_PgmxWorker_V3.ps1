param(
    [Parameter(Mandatory=$true)][string]$RowsCsv,
    [Parameter(Mandatory=$true)][string]$MappingsCsv,
    [Parameter(Mandatory=$true)][string]$SearchResultsCsv,
    [Parameter(Mandatory=$true)][string]$SettingsCsv,
    [Parameter(Mandatory=$true)][string]$ResultsCsv,
    [Parameter(Mandatory=$true)][string]$ProgressFile,
    [Parameter(Mandatory=$true)][string]$DoneFile,
    [Parameter(Mandatory=$true)][ValidateSet('Check','Execute')][string]$Mode,
    [Parameter(Mandatory=$true)][string]$SettingsVersion,
    [Parameter(Mandatory=$true)][string]$ConverterExe,
    [string]$ConversionLogCsv = '',
    [string]$OnlyProduct = ''
)

# PWX500 Agent V3.2.11 - fast automatic PGMX preparation + batched HOP -> PGMX conversion per product.
# ABSOLUTE SAFETY RULES:
# - Word files are NEVER enumerated/opened/parsed/launched.
# - HOP content is opened ONLY after an exact .hop extension gate and only inside the confirmed
#   product\03_Mashini\1_Holzher subtree. If multiple nested HOP folders exist, dimensions are
#   used only to select the correct folder; after folder selection V2.9 receives HOP by №/suffix.
# - Check mode NEVER changes the server and NEVER launches the converter.
# - Execute mode may create ONLY confirmed product\03_Mashini\2_PWX_500, run the bundled
#   HOP_to_PGMX_2_9.exe directly to that folder, and copy the produced base PGMX to a missing
#   ProgramCodeF PGMX. ONLY products with the special nested structure are centrally staged:
#   A) HOP files live below 1_Holzher in a nested folder (e.g. 90x200), OR
#   B) the confirmed product is a nested child inside a family folder that contains multiple
#      product children with their own direct 03_Mashini folders (Katerina-like structure).
#   For those products only, the final ProgramCodeF is copied (never moved) to
#   V:\30_Biesse_edging\scm\SCM_PWX_500. Normal products stay only in their own 2_PWX_500.
#   Base/detail-name PGMX files are NEVER copied there. Existing files are NEVER overwritten, edited, renamed or deleted.

$ErrorActionPreference = 'Stop'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$latin1 = [System.Text.Encoding]::GetEncoding(28591)
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$CentralProgramFolder = 'V:\30_Biesse_edging\scm\SCM_PWX_500'

try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch {}

function Write-Progress([string]$Text) {
    try { [System.IO.File]::WriteAllText($ProgressFile,$Text,$utf8NoBom) } catch {}
}
function S($Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}
function CsvField([string]$Text) {
    if ($null -eq $Text) { $Text='' }
    if ($Text.Contains('"')) { $Text=$Text.Replace('"','""') }
    if ($Text.Contains(';') -or $Text.Contains('"') -or $Text.Contains("`r") -or $Text.Contains("`n")) { return '"'+$Text+'"' }
    return $Text
}
function Add-ResultLine {
    param([System.Collections.ArrayList]$Output,[string[]]$Values)
    $fields=New-Object System.Collections.ArrayList
    foreach ($value in $Values) { [void]$fields.Add((CsvField $value)) }
    [void]$Output.Add((@($fields) -join ';'))
}
function Parse-Double([string]$Value) {
    $x=(S $Value).Replace(',','.')
    $v=0.0
    if ([double]::TryParse($x,[System.Globalization.NumberStyles]::Float,$inv,[ref]$v)) { return $v }
    return [double]::NaN
}
function Same-Dim([double]$A,[double]$B) { return ([Math]::Abs($A-$B) -le 0.2) }
function Same-PanelSize([double]$A1,[double]$A2,[double]$B1,[double]$B2) {
    # L/B orientation is intentionally ignored. A rectangular part is identified by
    # its two physical face dimensions as an unordered pair, plus thickness separately.
    return (((Same-Dim $A1 $B1) -and (Same-Dim $A2 $B2)) -or ((Same-Dim $A1 $B2) -and (Same-Dim $A2 $B1)))
}
function Get-PanelDimensionKey([double]$L,[double]$W,[double]$T) {
    $a=[Math]::Min($L,$W);$b=[Math]::Max($L,$W)
    return ($a.ToString('R',$inv)+'|'+$b.ToString('R',$inv)+'|'+$T.ToString('R',$inv))
}
function Normalize-Path([string]$Path) {
    $x=S $Path
    if ([string]::IsNullOrWhiteSpace($x)) { return '' }
    $x=$x.Replace('/','\')
    while ($x.EndsWith('\')) { $x=$x.Substring(0,$x.Length-1) }
    return $x.ToLowerInvariant()
}
# Compare mapped drives Y:, V: and Z: as equivalent ONLY below their
# shared 1_Izpolzvani root. Keep the physical path unchanged for all IO.
# Authorization still requires a match to a specifically approved subtree.
function Normalize-AllowedRootAliasPath {
    param([string]$Path)
    $p = Normalize-Path $Path
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    # Do not let a lexical ../ segment escape an approved child folder.
    if ([regex]::IsMatch($p, '(^|\\)\.\.(\\|$)')) { return '' }
    $driveRoot = [regex]::Match($p, '^[yvz]:\\1_izpolzvani(?=\\|$)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($driveRoot.Success) {
        return 'y:\1_izpolzvani' + $p.Substring($driveRoot.Length)
    }
    return $p
}

function Test-IsUnderAllowedRoot([string]$Path,[object[]]$Roots) {
    $p=Normalize-AllowedRootAliasPath $Path
    if ([string]::IsNullOrWhiteSpace($p)) { return $false }
    foreach ($rootItem in $Roots) {
        $r=Normalize-AllowedRootAliasPath ([string]$rootItem)
        if ([string]::IsNullOrWhiteSpace($r)) { continue }
        if ($p -eq $r) { return $true }
        if ($p.StartsWith(($r+'\'),[System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}
function Parse-SingleSearchPath([string]$Paths) {
    $t=S $Paths
    if ([string]::IsNullOrWhiteSpace($t)) { return '' }
    $parts=@($t -split '~\|~')
    if ($parts.Count -ne 1) { return '' }
    $x=S $parts[0]
    $m=[regex]::Match($x,'^\[(\d+)\](.*)$')
    if ($m.Success) { return (S $m.Groups[2].Value) }
    return $x
}
function Normalize-Name([string]$Text) {
    $x=(S $Text).ToLowerInvariant()
    $x=$x.Replace('_',' ').Replace('-',' ')
    $x=[regex]::Replace($x,'[^a-z0-9а-яё ]',' ')
    $x=[regex]::Replace($x,'\s+',' ').Trim()
    return $x
}
function Get-SafeTargetFile([string]$ProgramCodeF) {
    $p=S $ProgramCodeF
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    if ($p.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) { return '' }
    if ([System.IO.Path]::GetFileName($p) -ne $p) { return '' }
    $ext=[System.IO.Path]::GetExtension($p)
    if ([string]::IsNullOrWhiteSpace($ext)) { return $p+'.pgmx' }
    if ($ext.ToLowerInvariant() -eq '.pgmx') { return $p }
    return ''
}

# ----- Safe HOP dimension parser: extension gate comes BEFORE content access. -----
function Split-TopArgs([string]$Text) {
    $out=New-Object System.Collections.Generic.List[string]
    $cur=New-Object System.Text.StringBuilder
    $depth=0; $quote=[char]0
    foreach($ch in $Text.ToCharArray()) {
        if($quote -ne [char]0) {
            [void]$cur.Append($ch)
            if($ch -eq $quote){$quote=[char]0}
            continue
        }
        if($ch -eq '"' -or $ch -eq "'") {$quote=$ch;[void]$cur.Append($ch);continue}
        if($ch -eq '('){$depth++;[void]$cur.Append($ch);continue}
        if($ch -eq ')'){$depth--;[void]$cur.Append($ch);continue}
        if($ch -eq ',' -and $depth -eq 0){$out.Add($cur.ToString().Trim());[void]$cur.Clear();continue}
        [void]$cur.Append($ch)
    }
    if($cur.Length -gt 0){$out.Add($cur.ToString().Trim())}
    return @($out.ToArray())
}
function Eval-SimpleExpr([string]$Expr,[hashtable]$Names) {
    $x=S $Expr
    if([string]::IsNullOrWhiteSpace($x)){throw 'empty expression'}
    foreach($key in @($Names.Keys | Sort-Object { $_.Length } -Descending)) {
        $val=([double]$Names[$key]).ToString('R',$inv)
        $x=[regex]::Replace($x,'(?i)\b'+[regex]::Escape([string]$key)+'\b',$val)
    }
    if($x -match '[A-Za-z_]'){throw "unknown name in expression: $x"}
    $x=$x.Replace(',','.')
    if($x -notmatch '^[0-9\.\+\-\*\/\(\)\s]+$'){throw "unsupported expression: $x"}
    $dt=New-Object System.Data.DataTable
    $v=$dt.Compute($x,$null)
    return [Convert]::ToDouble($v,$inv)
}
function Get-HopDimensions([string]$Path) {
    if([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.hop'){ throw 'BLOCKED: HOP reader received non-HOP file.' }
    $bytes=[System.IO.File]::ReadAllBytes($Path)
    $text=$latin1.GetString($bytes)
    $names=@{}
    foreach($m in [regex]::Matches($text,'(?im)^\s*(DX|DY|DZ)\s*:=\s*([^;\r\n]+)')) {
        $key=$m.Groups[1].Value.ToUpperInvariant()
        try { $names[$key]=Eval-SimpleExpr $m.Groups[2].Value $names } catch {}
    }
    $fm=[regex]::Match($text,'(?im)^\s*FERTIGTEIL\s*\(([^)]*)\)')
    if($fm.Success) {
        $args=@(Split-TopArgs $fm.Groups[1].Value)
        if($args.Count -ge 3) {
            try {
                $vals=@((Eval-SimpleExpr $args[0] $names),(Eval-SimpleExpr $args[1] $names),(Eval-SimpleExpr $args[2] $names))
                if(-not $names.ContainsKey('DX') -or [double]$names['DX'] -le 0){$names['DX']=$vals[0]}
                if(-not $names.ContainsKey('DY') -or [double]$names['DY'] -le 0){$names['DY']=$vals[1]}
                if(-not $names.ContainsKey('DZ') -or [double]$names['DZ'] -le 0){$names['DZ']=$vals[2]}
            } catch {}
        }
    }
    if(-not $names.ContainsKey('DX') -or -not $names.ContainsKey('DY') -or -not $names.ContainsKey('DZ')){return $null}
    if([double]$names['DX'] -le 0 -or [double]$names['DY'] -le 0 -or [double]$names['DZ'] -le 0){return $null}
    return [pscustomobject]@{DX=[double]$names['DX'];DY=[double]$names['DY'];DZ=[double]$names['DZ']}
}
function Get-HopFileMeta([string]$Path) {
    if([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.hop'){ return $null }
    $stem=[System.IO.Path]::GetFileNameWithoutExtension($Path)
    $m=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?_(?<name>.+)$')
    if(-not $m.Success){$m=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?\s*(?<name>.*)$')}
    if(-not $m.Success){return $null}
    $no=0;if(-not [int]::TryParse($m.Groups['no'].Value,[ref]$no)){return $null}
    $proc=0
    if($m.Groups['proc'].Success -and -not [string]::IsNullOrWhiteSpace($m.Groups['proc'].Value)){[void][int]::TryParse($m.Groups['proc'].Value,[ref]$proc)}
    $dims=Get-HopDimensions $Path
    return [pscustomobject]@{Path=$Path;File=[System.IO.Path]::GetFileName($Path);No=$no;Proc=$proc;NameNorm=(Normalize-Name $m.Groups['name'].Value);Dims=$dims}
}


# ----- HOP source-folder resolver for nested 1_Holzher structures. -----
# A confirmed product may keep HOP files directly in 1_Holzher OR in one/more nested
# variant folders such as 90x200 / 120x200 / 160x200. We first identify the correct
# DIRECT HOP-containing folder, then conversion uses only that folder by №/suffix.
function Get-HopContainingFolders([string]$HolzRoot) {
    $out=New-Object System.Collections.ArrayList
    if(-not [System.IO.Directory]::Exists($HolzRoot)){return @()}
    $seen=New-Object 'System.Collections.Generic.HashSet[string]'
    try {
        # Enumerate ONLY *.hop paths recursively. No other file type is opened or parsed.
        foreach($hp in [System.IO.Directory]::EnumerateFiles($HolzRoot,'*.hop',[System.IO.SearchOption]::AllDirectories)) {
            $dir=[System.IO.Path]::GetDirectoryName($hp)
            if(-not [string]::IsNullOrWhiteSpace($dir) -and $seen.Add($dir)){[void]$out.Add($dir)}
        }
    } catch { return @() }
    return @($out | Sort-Object)
}

function Get-RelativeDepthUnderAllowedRoot([string]$Path,[object[]]$Roots) {
    $p=Normalize-AllowedRootAliasPath $Path
    if([string]::IsNullOrWhiteSpace($p)){return -1}
    $best=-1
    foreach($rootItem in @($Roots)) {
        $r=Normalize-AllowedRootAliasPath ([string]$rootItem)
        if([string]::IsNullOrWhiteSpace($r)){continue}
        if($p.Equals($r,[System.StringComparison]::OrdinalIgnoreCase)){if($best -lt 0){$best=0};continue}
        $prefix=$r+'\\'
        if($p.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)) {
            $rel=$p.Substring($prefix.Length)
            $depth=@($rel.Split('\\',[System.StringSplitOptions]::RemoveEmptyEntries)).Count
            if($depth -gt $best){$best=$depth}
        }
    }
    return $best
}
function Test-FamilyNestedProductStructure([string]$ProductPath,[object[]]$Roots) {
    if(-not [System.IO.Directory]::Exists($ProductPath)){return $false}
    $depth=Get-RelativeDepthUnderAllowedRoot -Path $ProductPath -Roots $Roots
    # Normal products are typically category/product (depth 2). A family-child product is one level deeper.
    if($depth -lt 3){return $false}
    try {
        $parent=[System.IO.Directory]::GetParent($ProductPath)
        if($null -eq $parent){return $false}
        # The family folder itself must not already be a product end-point.
        if([System.IO.Directory]::Exists((Join-Path $parent.FullName '03_Mashini'))){return $false}
        $childrenWithMachines=0
        foreach($child in [System.IO.Directory]::EnumerateDirectories($parent.FullName)) {
            if([System.IO.Directory]::Exists((Join-Path $child '03_Mashini'))){$childrenWithMachines++}
            if($childrenWithMachines -ge 2){return $true}
        }
    } catch { return $false }
    return $false
}
function Test-CentralStagingRequired([string]$ProductPath,[string]$HolzRoot,[object[]]$Roots) {
    # Special structure type A: HOP files are kept in one or more nested folders below 1_Holzher.
    if([System.IO.Directory]::Exists($HolzRoot)) {
        try {
            $holzNorm=Normalize-Path $HolzRoot
            foreach($d in @(Get-HopContainingFolders $HolzRoot)) {
                if((Normalize-Path ([string]$d)) -ne $holzNorm){return $true}
            }
        } catch {}
    }
    # Special structure type B: Katerina-like family folder with multiple child products.
    if(Test-FamilyNestedProductStructure -ProductPath $ProductPath -Roots $Roots){return $true}
    return $false
}

function Get-RelativeFolderLabel([string]$Root,[string]$Folder) {
    $r=(S $Root).TrimEnd('\');$f=S $Folder
    if([string]::Equals($r,$f,[System.StringComparison]::OrdinalIgnoreCase)){return '.'}
    if($f.StartsWith(($r+'\'),[System.StringComparison]::OrdinalIgnoreCase)){return $f.Substring($r.Length+1)}
    return $f
}
function Build-ExpectedHopChecks([object[]]$ProductRows,[hashtable]$PhysicalCount) {
    $out=New-Object System.Collections.ArrayList
    $seen=New-Object 'System.Collections.Generic.HashSet[string]'
    foreach($r in @($ProductRows)) {
        $no=0;if(-not [int]::TryParse((S $r.DetailNo),[ref]$no)){continue}
        $l=Parse-Double (S $r.Length);$w=Parse-Double (S $r.Width);$t=Parse-Double (S $r.Thickness)
        if([double]::IsNaN($l) -or [double]::IsNaN($w) -or [double]::IsNaN($t)){continue}
        $pn=1;if(-not [int]::TryParse((S $r.ProcessingNo),[ref]$pn) -or $pn -lt 1){$pn=1}
        $physicalKey=(S $r.CsvFile)+'|'+(S $r.Product)+'|'+(S $r.DetailNo)+'|'+(S $r.DetailName)
        $cnt=1;if($PhysicalCount.ContainsKey($physicalKey)){$cnt=[int]$PhysicalCount[$physicalKey]}
        if($cnt -gt 2){continue}
        $exp=0;if($cnt -gt 1){$exp=$pn}
        $dimKey=Get-PanelDimensionKey $l $w $t
        $key=([string]$no)+'|'+([string]$exp)+'|'+$dimKey+'|'+(Normalize-Name (S $r.DetailName))
        if($seen.Add($key)){
            [void]$out.Add([pscustomobject]@{No=$no;ExpectedProc=$exp;L=$l;W=$w;T=$t;NameNorm=(Normalize-Name (S $r.DetailName));PhysicalKey=([string]$no+'|'+$dimKey)})
        }
    }
    return @($out.ToArray())
}
function Get-HopFolderMatchScore([string]$Folder,[object[]]$ExpectedChecks) {
    $metas=New-Object System.Collections.ArrayList
    $expectedNos=New-Object 'System.Collections.Generic.HashSet[int]'
    foreach($e in @($ExpectedChecks)){[void]$expectedNos.Add([int]$e.No)}
    $readable=0;$unreadable=0
    try {
        foreach($hp in [System.IO.Directory]::EnumerateFiles($Folder,'*.hop',[System.IO.SearchOption]::TopDirectoryOnly)) {
            $stem=[System.IO.Path]::GetFileNameWithoutExtension($hp)
            $m=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?')
            if(-not $m.Success){continue}
            $n=0;if(-not [int]::TryParse($m.Groups['no'].Value,[ref]$n)){continue}
            if(-not $expectedNos.Contains($n)){continue}
            try{$meta=Get-HopFileMeta $hp;if($null -ne $meta -and $null -ne $meta.Dims){[void]$metas.Add($meta);$readable++}else{$unreadable++}}catch{$unreadable++}
        }
    } catch {
        return [pscustomobject]@{Folder=$Folder;Matched=0;MatchedPhysical=0;IdentityMismatch=0;Readable=$readable;Unreadable=$unreadable;Error=$_.Exception.Message}
    }
    $matched=0;$mismatchPhysical=New-Object 'System.Collections.Generic.HashSet[string]';$matchedPhysical=New-Object 'System.Collections.Generic.HashSet[string]'
    foreach($e in @($ExpectedChecks)) {
        $byId=@($metas | Where-Object { [int]$_.No -eq [int]$e.No -and [int]$_.Proc -eq [int]$e.ExpectedProc })
        if($byId.Count -eq 0){continue}
        $ok=$false
        foreach($h in $byId) {
            if((Same-PanelSize ([double]$h.Dims.DX) ([double]$h.Dims.DY) ([double]$e.L) ([double]$e.W)) -and (Same-Dim ([double]$h.Dims.DZ) ([double]$e.T))){$ok=$true;break}
        }
        if($ok){$matched++;[void]$matchedPhysical.Add([string]$e.PhysicalKey)}else{[void]$mismatchPhysical.Add([string]$e.PhysicalKey)}
    }
    return [pscustomobject]@{Folder=$Folder;Matched=$matched;MatchedPhysical=$matchedPhysical.Count;IdentityMismatch=$mismatchPhysical.Count;Readable=$readable;Unreadable=$unreadable;Error=''}
}
function Resolve-ProductHopFolder([string]$HolzRoot,[object[]]$ProductRows,[hashtable]$PhysicalCount) {
    if(-not [System.IO.Directory]::Exists($HolzRoot)){return [pscustomobject]@{Status='NO_HOLZHER';Path='';Count=0;Reason='Липсва 03_Mashini\\1_Holzher.'}}
    $dirs=@(Get-HopContainingFolders $HolzRoot)
    if($dirs.Count -eq 0){return [pscustomobject]@{Status='MISSING';Path='';Count=0;Reason='В 1_Holzher и подпапките му няма .hop файлове.'}}
    if($dirs.Count -eq 1){
        return [pscustomobject]@{Status='SELECTED';Path=[string]$dirs[0];Count=1;Reason=('Единствена HOP папка: '+(Get-RelativeFolderLabel $HolzRoot ([string]$dirs[0])))}
    }

    $checks=@(Build-ExpectedHopChecks -ProductRows $ProductRows -PhysicalCount $PhysicalCount)
    if($checks.Count -eq 0){
        $labels=@($dirs | ForEach-Object {Get-RelativeFolderLabel $HolzRoot ([string]$_)})
        return [pscustomobject]@{Status='AMBIGUOUS';Path='';Count=$dirs.Count;Reason=('Няколко HOP подпапки и няма достатъчно CSV размери за сравнение: '+($labels -join ', '))}
    }
    $expectedPhysical=New-Object 'System.Collections.Generic.HashSet[string]'
    foreach($e in $checks){[void]$expectedPhysical.Add([string]$e.PhysicalKey)}
    $scores=New-Object System.Collections.ArrayList
    foreach($d in $dirs){[void]$scores.Add((Get-HopFolderMatchScore -Folder ([string]$d) -ExpectedChecks $checks))}

    $perfect=@($scores | Where-Object { [int]$_.Matched -eq $checks.Count -and [int]$_.IdentityMismatch -eq 0 })
    if($perfect.Count -eq 1){
        $one=$perfect[0]
        return [pscustomobject]@{Status='SELECTED';Path=[string]$one.Folder;Count=$dirs.Count;Reason=('HOP подпапката е избрана по пълно съвпадение на №/суфикс/размери: '+(Get-RelativeFolderLabel $HolzRoot ([string]$one.Folder))+' ('+[string]$one.Matched+'/'+[string]$checks.Count+')')}
    }
    if($perfect.Count -gt 1){
        $labels=@($perfect | ForEach-Object {Get-RelativeFolderLabel $HolzRoot ([string]$_.Folder)})
        return [pscustomobject]@{Status='AMBIGUOUS';Path='';Count=$perfect.Count;Reason=('Повече от една HOP подпапка съвпада напълно: '+($labels -join ', ')+'. Нужен е ръчен избор.')}
    }

    $ordered=@($scores | Sort-Object @{Expression='MatchedPhysical';Descending=$true},@{Expression='Matched';Descending=$true})
    $top=$ordered[0]
    $second=if($ordered.Count -gt 1){$ordered[1]}else{$null}
    $threshold=[Math]::Min(2,[Math]::Max(1,$expectedPhysical.Count))
    $secondMatched=if($null -ne $second){[int]$second.MatchedPhysical}else{0}
    # Safe partial auto-choice: at least two distinct physical details (or all if product has only one),
    # no same-identity dimension mismatch, and a clear lead of >=2 physical matches over the next folder
    # (or the other folders have zero matching physical details).
    $clearLead=($secondMatched -eq 0 -or ([int]$top.MatchedPhysical -ge ($secondMatched+2)))
    if([int]$top.MatchedPhysical -ge $threshold -and [int]$top.IdentityMismatch -eq 0 -and $clearLead){
        return [pscustomobject]@{Status='SELECTED';Path=[string]$top.Folder;Count=$dirs.Count;Reason=('HOP подпапката е избрана по размери на детайлите: '+(Get-RelativeFolderLabel $HolzRoot ([string]$top.Folder))+'; съвпаднали физически детайли '+[string]$top.MatchedPhysical+'.')}
    }

    $summary=New-Object System.Collections.ArrayList
    foreach($sc in $ordered){[void]$summary.Add((Get-RelativeFolderLabel $HolzRoot ([string]$sc.Folder))+'='+[string]$sc.MatchedPhysical+' физ. / '+[string]$sc.Matched+' операции')}
    return [pscustomobject]@{Status='AMBIGUOUS';Path='';Count=$dirs.Count;Reason=('Няколко HOP подпапки без еднозначен победител: '+(@($summary) -join ' | ')+'. Нужен е ръчен избор.')}
}
function Resolve-HopCandidateSet([System.Collections.ArrayList]$Candidates,[string]$DetailName,[string]$ModeReason) {
    if($Candidates.Count -eq 0){return $null}
    if($Candidates.Count -eq 1) {
        $one=$Candidates[0]
        return [pscustomobject]@{Status='FOUND';Path=[string]$one.Path;File=[string]$one.File;Count=1;Reason=$ModeReason}
    }

    # Name is only a tie-breaker. Number/suffix + physical dimensions are already mandatory.
    $dn=Normalize-Name $DetailName
    if(-not [string]::IsNullOrWhiteSpace($dn)) {
        $nameMatches=New-Object System.Collections.ArrayList
        foreach($meta in @($Candidates)) {
            $n=[string]$meta.NameNorm
            if(-not [string]::IsNullOrWhiteSpace($n) -and ($n -eq $dn -or $n.Contains($dn) -or $dn.Contains($n))){[void]$nameMatches.Add($meta)}
        }
        if($nameMatches.Count -eq 1) {
            $one=$nameMatches[0]
            return [pscustomobject]@{Status='FOUND';Path=[string]$one.Path;File=[string]$one.File;Count=$Candidates.Count;Reason=($ModeReason+' Няколко кандидата; името даде еднозначен избор.')}
        }
    }
    return [pscustomobject]@{Status='AMBIGUOUS';Path='';File='';Count=$Candidates.Count;Reason=('Няколко HOP с еднакви №/суфикс и физически размери: '+[string]$Candidates.Count)}
}

function Resolve-SourceHop([string]$HolzFolder,[int]$DetailNo,[int]$ExpectedProc,[string]$DetailName,[double]$L,[double]$W,[double]$T) {
    if(-not [System.IO.Directory]::Exists($HolzFolder)){return [pscustomobject]@{Status='NO_HOLZHER';Path='';File='';Count=0;Reason='Липсва 03_Mashini\\1_Holzher.'}}
    $sizeMatches=New-Object System.Collections.ArrayList
    $identityCount=0; $unreadable=0
    $dimensionEvidence=New-Object System.Collections.ArrayList
    try {
        # Enumerate ONLY *.hop. No other extension is opened or parsed.
        foreach($hp in [System.IO.Directory]::EnumerateFiles($HolzFolder,'*.hop',[System.IO.SearchOption]::TopDirectoryOnly)) {
            $stem=[System.IO.Path]::GetFileNameWithoutExtension($hp)
            $prefix=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?')
            if(-not $prefix.Success){continue}
            $no=0;$proc=0
            if(-not [int]::TryParse($prefix.Groups['no'].Value,[ref]$no)){continue}
            if($prefix.Groups['proc'].Success -and -not [string]::IsNullOrWhiteSpace($prefix.Groups['proc'].Value)){[void][int]::TryParse($prefix.Groups['proc'].Value,[ref]$proc)}
            if($no -ne $DetailNo -or $proc -ne $ExpectedProc){continue}
            $identityCount++
            try {
                $meta=Get-HopFileMeta $hp
                if($null -eq $meta -or $null -eq $meta.Dims){$unreadable++;continue}
                $dx=[double]$meta.Dims.DX;$dy=[double]$meta.Dims.DY;$dz=[double]$meta.Dims.DZ
                [void]$dimensionEvidence.Add(('{0}: HOP {1} x {2} x {3}; CSV {4} x {5} x {6}' -f $meta.File,$dx,$dy,$dz,$L,$W,$T))

                # Permanent rule: L/B orientation is NOT identity. The two face dimensions
                # must match as an unordered pair. Detail number, processing suffix and thickness
                # remain mandatory. This accepts both DX=L,DY=W and DX=W,DY=L equally.
                $sizeOk=(Same-PanelSize $dx $dy $L $W) -and (Same-Dim $dz $T)
                if($sizeOk){[void]$sizeMatches.Add($meta)}
            } catch {$unreadable++}
        }
    } catch { return [pscustomobject]@{Status='ERROR';Path='';File='';Count=0;Reason=('Грешка при HOP списъка: '+$_.Exception.Message)} }

    $resolved=Resolve-HopCandidateSet -Candidates $sizeMatches -DetailName $DetailName -ModeReason '№/суфикс + физически размери + дебелина съвпадат; ориентацията L/B спрямо DX/DY не се използва.'
    if($null -ne $resolved){return $resolved}

    $why='Няма HOP със задължителни №/суфикс и съвпадащи физически размери независимо от ориентацията L/B.'
    if($identityCount -eq 0){$why='Няма HOP със задължителните № и суфикс.'}
    elseif($unreadable -gt 0){$why+=' Неразчетени HOP: '+[string]$unreadable+'.'}
    if($dimensionEvidence.Count -gt 0){$why+=' '+(@($dimensionEvidence) -join ' | ')}
    return [pscustomobject]@{Status='MISSING';Path='';File='';Count=$identityCount;Reason=$why}
}


# ----- Conversion source resolver: NO HOP content parsing, NO dimension pre-filter. -----
# Rule V3.2.11: first resolve the correct DIRECT HOP-containing folder under 1_Holzher.
# If there is only one HOP folder it is accepted. If there are several, CSV dimensions
# choose a unique folder where possible; otherwise the product stays manual. AFTER that,
# every CSV detail that still needs a base PGMX is passed to V2.9 by filename identity
# (detail number + required _1/_2 suffix), with no size pre-filter before conversion.
function Resolve-SourceHopForConversion([string]$HolzFolder,[int]$DetailNo,[int]$ExpectedProc,[string]$DetailName) {
    if(-not [System.IO.Directory]::Exists($HolzFolder)){return [pscustomobject]@{Status='NO_HOLZHER';Path='';File='';Count=0;Reason='Липсва 03_Mashini\\1_Holzher.'}}
    $matches=New-Object System.Collections.ArrayList
    try {
        foreach($hp in [System.IO.Directory]::EnumerateFiles($HolzFolder,'*.hop',[System.IO.SearchOption]::TopDirectoryOnly)) {
            # Extension filter is applied by EnumerateFiles before any content access. No HOP is opened here.
            $stem=[System.IO.Path]::GetFileNameWithoutExtension($hp)
            $m=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?(?:[_\s-]+)?(?<name>.*)$')
            if(-not $m.Success){continue}
            $no=0;$proc=0
            if(-not [int]::TryParse($m.Groups['no'].Value,[ref]$no)){continue}
            if($m.Groups['proc'].Success -and -not [string]::IsNullOrWhiteSpace($m.Groups['proc'].Value)){[void][int]::TryParse($m.Groups['proc'].Value,[ref]$proc)}
            if($no -ne $DetailNo -or $proc -ne $ExpectedProc){continue}
            [void]$matches.Add([pscustomobject]@{Path=$hp;File=[System.IO.Path]::GetFileName($hp);NameNorm=(Normalize-Name $m.Groups['name'].Value)})
        }
    } catch {return [pscustomobject]@{Status='ERROR';Path='';File='';Count=0;Reason=('Грешка при HOP списъка: '+$_.Exception.Message)}}
    if($matches.Count -eq 0){return [pscustomobject]@{Status='MISSING';Path='';File='';Count=0;Reason='Няма HOP със задължителните № и суфикс.'}}
    if($matches.Count -eq 1){$one=$matches[0];return [pscustomobject]@{Status='FOUND';Path=[string]$one.Path;File=[string]$one.File;Count=1;Reason='HOP е намерен по №/суфикс. Размерите не се използват като филтър преди V2.9.'}}
    $dn=Normalize-Name $DetailName
    if(-not [string]::IsNullOrWhiteSpace($dn)) {
        $nameMatches=New-Object System.Collections.ArrayList
        foreach($x in @($matches)) {
            $n=[string]$x.NameNorm
            if(-not [string]::IsNullOrWhiteSpace($n) -and ($n -eq $dn -or $n.Contains($dn) -or $dn.Contains($n))){[void]$nameMatches.Add($x)}
        }
        if($nameMatches.Count -eq 1){$one=$nameMatches[0];return [pscustomobject]@{Status='FOUND';Path=[string]$one.Path;File=[string]$one.File;Count=$matches.Count;Reason='Няколко HOP по №/суфикс; името даде еднозначен избор. Размерите не се използват като филтър.'}}
    }
    return [pscustomobject]@{Status='AMBIGUOUS';Path='';File='';Count=$matches.Count;Reason=('Няколко HOP със същите №/суфикс: '+[string]$matches.Count+'. Не се избира на сляпо.')}
}

# ----- Fast PGMX name lookup. No PGMX content is opened in normal V3.2.11 processing. -----
function Get-PgmxNameMeta([string]$Path) {
    if([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.pgmx'){return $null}
    $stem=[System.IO.Path]::GetFileNameWithoutExtension($Path)
    $m=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?(?:[_\s-]+)(?<name>.+)$')
    if(-not $m.Success){return $null}
    $no=0;if(-not [int]::TryParse($m.Groups['no'].Value,[ref]$no)){return $null}
    $proc=0
    if($m.Groups['proc'].Success -and -not [string]::IsNullOrWhiteSpace($m.Groups['proc'].Value)){[void][int]::TryParse($m.Groups['proc'].Value,[ref]$proc)}
    return [pscustomobject]@{Path=$Path;File=[System.IO.Path]::GetFileName($Path);No=$no;Proc=$proc;NameNorm=(Normalize-Name $m.Groups['name'].Value)}
}
function Resolve-BasePgmxFast([string]$PwxFolder,[int]$DetailNo,[int]$ExpectedProc,[string]$DetailName) {
    if(-not [System.IO.Directory]::Exists($PwxFolder)){return [pscustomobject]@{Status='MISSING';Path='';File='';Count=0;Reason=''}}
    $identity=New-Object System.Collections.ArrayList
    try {
        foreach($pgmxPath in [System.IO.Directory]::EnumerateFiles($PwxFolder,'*.pgmx',[System.IO.SearchOption]::TopDirectoryOnly)) {
            $meta=Get-PgmxNameMeta $pgmxPath
            if($null -ne $meta -and [int]$meta.No -eq $DetailNo -and [int]$meta.Proc -eq $ExpectedProc){[void]$identity.Add($meta)}
        }
    } catch { return [pscustomobject]@{Status='ERROR';Path='';File='';Count=0;Reason=('PGMX списък: '+$_.Exception.Message)} }
    if($identity.Count -eq 0){return [pscustomobject]@{Status='MISSING';Path='';File='';Count=0;Reason=''}}
    if($identity.Count -eq 1){$one=$identity[0];return [pscustomobject]@{Status='FOUND';Path=[string]$one.Path;File=[string]$one.File;Count=1;Reason='Еднозначен основен PGMX по №/суфикс.'}}

    # При повече от един основен PGMX името е само tie-breaker. Не гадаем.
    $dn=Normalize-Name $DetailName
    if(-not [string]::IsNullOrWhiteSpace($dn)) {
        $nameMatches=New-Object System.Collections.ArrayList
        foreach($meta in @($identity)) {
            $n=[string]$meta.NameNorm
            if(-not [string]::IsNullOrWhiteSpace($n) -and ($n -eq $dn -or $n.Contains($dn) -or $dn.Contains($n))){[void]$nameMatches.Add($meta)}
        }
        if($nameMatches.Count -eq 1){$one=$nameMatches[0];return [pscustomobject]@{Status='FOUND';Path=[string]$one.Path;File=[string]$one.File;Count=$identity.Count;Reason='Няколко по №/суфикс; името даде еднозначен основен PGMX.'}}
    }
    return [pscustomobject]@{Status='AMBIGUOUS';Path='';File='';Count=$identity.Count;Reason=('Няколко основни PGMX по №/суфикс: '+[string]$identity.Count)}
}

function Invoke-ConverterCliOnce([string[]]$HopPaths,[string]$OutDir,[int]$TimeoutMs) {
    $stdout=Join-Path ([System.IO.Path]::GetTempPath()) ('pwx500_conv_out_'+[Guid]::NewGuid().ToString('N')+'.txt')
    $stderr=Join-Path ([System.IO.Path]::GetTempPath()) ('pwx500_conv_err_'+[Guid]::NewGuid().ToString('N')+'.txt')
    $args=New-Object System.Collections.ArrayList
    [void]$args.Add('--outdir')
    [void]$args.Add('"'+$OutDir.Replace('"','\"')+'"')
    foreach($hp in @($HopPaths)){[void]$args.Add('"'+([string]$hp).Replace('"','\"')+'"')}
    $argLine=@($args) -join ' '
    $exit=-999;$launchError=''
    try {
        $p=Start-Process -FilePath $ConverterExe -ArgumentList $argLine -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        $finished=$p.WaitForExit($TimeoutMs)
        if(-not $finished){
            try{$p.Kill()}catch{};try{$p.WaitForExit()}catch{}
            $launchError='Конверторът надхвърли допустимото време и беше спрян.';$exit=-2
        } else { try{$p.WaitForExit()}catch{};try{$exit=$p.ExitCode}catch{} }
        try{$p.Dispose()}catch{}
    } catch {$launchError=$_.Exception.Message}

    $outText='';try{if([System.IO.File]::Exists($stdout)){$outText=[System.IO.File]::ReadAllText($stdout).Trim()}}catch{}
    $errText='';try{if([System.IO.File]::Exists($stderr)){$errText=[System.IO.File]::ReadAllText($stderr).Trim()}}catch{}
    try{if([System.IO.File]::Exists($stdout)){[System.IO.File]::Delete($stdout)}}catch{}
    try{if([System.IO.File]::Exists($stderr)){[System.IO.File]::Delete($stderr)}}catch{}
    return [pscustomobject]@{ExitCode=$exit;Out=$outText;Err=$errText;LaunchError=$launchError}
}

function Invoke-ConverterBatchFast([string[]]$HopPaths,[string]$OutDir) {
    if(-not [System.IO.File]::Exists($ConverterExe)){return [pscustomobject]@{OK=$false;ExitCode=-1;Count=0;Message='Липсва HOP_to_PGMX_2_9.exe.'}}
    if($null -eq $HopPaths -or $HopPaths.Count -eq 0){return [pscustomobject]@{OK=$true;ExitCode=0;Count=0;Message='Няма HOP файлове за конверсия.'}}
    if(-not [System.IO.Directory]::Exists($OutDir)){[void][System.IO.Directory]::CreateDirectory($OutDir)}

    $unique=New-Object System.Collections.ArrayList
    $seen=@{}
    foreach($hp0 in @($HopPaths)) {
        $hp=S $hp0
        if([string]::IsNullOrWhiteSpace($hp)){continue}
        if([System.IO.Path]::GetExtension($hp).ToLowerInvariant() -ne '.hop'){return [pscustomobject]@{OK=$false;ExitCode=-1;Count=0;Message=('BLOCKED: converter input is not .hop: '+$hp)}}
        if(-not [System.IO.File]::Exists($hp)){return [pscustomobject]@{OK=$false;ExitCode=-1;Count=0;Message=('Липсва HOP: '+$hp)}}
        $key=$hp.ToLowerInvariant()
        if(-not $seen.ContainsKey($key)){$seen[$key]=$true;[void]$unique.Add($hp)}
    }
    if($unique.Count -eq 0){return [pscustomobject]@{OK=$true;ExitCode=0;Count=0;Message='Няма HOP файлове за конверсия.'}}

    $pending=New-Object System.Collections.ArrayList
    foreach($hp in @($unique)) {
        $outFile=Join-Path $OutDir (([System.IO.Path]::GetFileNameWithoutExtension([string]$hp))+'.pgmx')
        if(-not [System.IO.File]::Exists($outFile)){[void]$pending.Add([string]$hp)}
    }
    if($pending.Count -eq 0){return [pscustomobject]@{OK=$true;ExitCode=0;Count=0;Message='Всички основни PGMX вече съществуват.'}}

    # First pass: fast batch, exactly as in V3.2.9.
    $timeoutMs=[Math]::Max(60000,[Math]::Min(600000,([int]$pending.Count*15000)))
    $first=Invoke-ConverterCliOnce -HopPaths @($pending) -OutDir $OutDir -TimeoutMs $timeoutMs

    # Second pass V3.2.11: ONLY files still missing after the batch are retried one-by-one.
    # This preserves the fast normal path while reproducing the user's proven one-file-at-a-time behavior
    # for exceptional HOPs. Existing outputs are never overwritten by the agent.
    $missing=New-Object System.Collections.ArrayList
    foreach($hp in @($pending)) {
        $outFile=Join-Path $OutDir (([System.IO.Path]::GetFileNameWithoutExtension([string]$hp))+'.pgmx')
        if(-not [System.IO.File]::Exists($outFile)){[void]$missing.Add([string]$hp)}
    }
    $retryTotal=$missing.Count;$retryOk=0
    $retryNotes=New-Object System.Collections.ArrayList
    foreach($hp in @($missing)) {
        $single=Invoke-ConverterCliOnce -HopPaths @([string]$hp) -OutDir $OutDir -TimeoutMs 90000
        $outFile=Join-Path $OutDir (([System.IO.Path]::GetFileNameWithoutExtension([string]$hp))+'.pgmx')
        if([System.IO.File]::Exists($outFile)){
            $retryOk++
            [void]$retryNotes.Add(('RETRY OK '+[System.IO.Path]::GetFileName([string]$hp)))
        } else {
            $note=('RETRY FAIL '+[System.IO.Path]::GetFileName([string]$hp)+' ExitCode='+[string]$single.ExitCode)
            if(-not [string]::IsNullOrWhiteSpace([string]$single.LaunchError)){$note+=' | '+[string]$single.LaunchError}
            if(-not [string]::IsNullOrWhiteSpace([string]$single.Out)){$note+=' | OUT: '+[string]$single.Out}
            if(-not [string]::IsNullOrWhiteSpace([string]$single.Err)){$note+=' | ERR: '+[string]$single.Err}
            [void]$retryNotes.Add($note)
        }
    }

    $created=0
    foreach($hp in @($pending)) {
        $outFile=Join-Path $OutDir (([System.IO.Path]::GetFileNameWithoutExtension([string]$hp))+'.pgmx')
        if([System.IO.File]::Exists($outFile)){$created++}
    }
    $ok=($created -eq $pending.Count)
    $finalExit=if($ok){0}else{[int]$first.ExitCode}
    $msg=('Пакетна конверсия: '+[string]$created+'/'+[string]$pending.Count+' PGMX. FirstExitCode='+[string]$first.ExitCode)
    if($retryTotal -gt 0){$msg+=' | Втори опит поединично: '+[string]$retryOk+'/'+[string]$retryTotal}
    if(-not [string]::IsNullOrWhiteSpace([string]$first.LaunchError)){$msg+=' | '+[string]$first.LaunchError}
    if(-not [string]::IsNullOrWhiteSpace([string]$first.Out)){$msg+=' | OUT: '+[string]$first.Out}
    if(-not [string]::IsNullOrWhiteSpace([string]$first.Err)){$msg+=' | ERR: '+[string]$first.Err}
    if($retryNotes.Count -gt 0){$msg+=' | '+(@($retryNotes) -join ' | ')}
    return [pscustomobject]@{OK=$ok;ExitCode=$finalExit;Count=$created;Message=$msg}
}
function Append-ConversionLog([string]$Product,[string]$DetailNo,[string]$Program,[string]$Hop,[string]$Base,[string]$Result,[string]$Info,[int]$ExitCode) {
    if([string]::IsNullOrWhiteSpace($ConversionLogCsv)){return}
    try {
        if(-not [System.IO.File]::Exists($ConversionLogCsv)){
            [System.IO.File]::WriteAllText($ConversionLogCsv,'Timestamp;Product;DetailNo;ProgramCodeF;HOP;BasePGMX;Result;Info;ConverterExitCode'+[Environment]::NewLine,$utf8Bom)
        }
        $line=@((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),$Product,$DetailNo,$Program,$Hop,$Base,$Result,$Info,[string]$ExitCode) | ForEach-Object { CsvField ([string]$_) }
        [System.IO.File]::AppendAllText($ConversionLogCsv,(($line -join ';')+[Environment]::NewLine),$utf8NoBom)
    } catch {}
}

function Is-TrustedMappingSource([string]$Source){
    $x=(S $Source).ToUpperInvariant()
    return ($x -in @('HISTORICAL_PGMX','HOP_AUTO','HOP_VERIFIED_MANUAL','HOP_VERIFIED_USER'))
}

try {
    Write-Progress 'PGMX V3.2.11: зареждане на локалните CSV резултати...'
    if(-not [System.IO.File]::Exists($RowsCsv)){throw 'Липсват прочетените CSV редове.'}
    if(-not [System.IO.File]::Exists($SearchResultsCsv)){throw 'Липсват резултатите от търсенето на изделия.'}
    if(-not [System.IO.File]::Exists($SettingsCsv)){throw 'Липсва списъкът с разрешени папки.'}

    $rows=@(Import-Csv -LiteralPath $RowsCsv -Delimiter ';')
    if(-not [string]::IsNullOrWhiteSpace((S $OnlyProduct))){
        $rows=@($rows | Where-Object { (S $_.Product) -eq (S $OnlyProduct) })
        if($rows.Count -eq 0){throw ('OnlyProduct филтър: не са намерени CSV редове за изделието: '+(S $OnlyProduct))}
        Write-Progress ('PGMX V3.2.11: OnlyProduct филтър: '+(S $OnlyProduct))
    }
    $search=@(Import-Csv -LiteralPath $SearchResultsCsv -Delimiter ';')
    $settings=@(Import-Csv -LiteralPath $SettingsCsv -Delimiter ';')
    $maps=@();if(-not [string]::IsNullOrWhiteSpace($MappingsCsv) -and [System.IO.File]::Exists($MappingsCsv)){$maps=@(Import-Csv -LiteralPath $MappingsCsv -Delimiter ';')}

    $roots=New-Object System.Collections.ArrayList
    foreach($setting in $settings){if((S $setting.Version) -ne $SettingsVersion){continue};$rp=S $setting.Path;if(-not [string]::IsNullOrWhiteSpace($rp) -and -not $roots.Contains($rp)){[void]$roots.Add($rp)}}
    if($roots.Count -eq 0){throw 'Няма разрешени папки за текущата версия на настройките.'}

    $mapped=@{};foreach($m in $maps){
        $product=S $m.Product;$path=S $m.Path;$source=S $m.Source
        if(-not [string]::IsNullOrWhiteSpace($product) -and -not [string]::IsNullOrWhiteSpace($path) -and (Is-TrustedMappingSource $source)){$mapped[$product]=$path}
    }
    # V5.1.13 safety: a raw single search candidate is NEVER enough for PGMX processing.
    # It must first become a trusted mapping through HOP verification.
    $singleSearch=@{}

    # Multi-processing rule for conversion: based only on CSV identity, not dimensions.
    # Within one CSV/product/detail-name/detail-number, one row => no suffix; two rows => _1/_2.
    # Dimensions are deliberately not used to decide which HOPs are sent to V2.9.
    $physicalCount=@{}
    foreach($row in $rows){
        $key=(S $row.CsvFile)+'|'+(S $row.Product)+'|'+(S $row.DetailNo)+'|'+(S $row.DetailName)
        if(-not $physicalCount.ContainsKey($key)){$physicalCount[$key]=0}
        $physicalCount[$key]=[int]$physicalCount[$key]+1
    }
    $productRowsByProduct=@{}
    foreach($row in $rows){
        $pp=S $row.Product
        if([string]::IsNullOrWhiteSpace($pp)){continue}
        if(-not $productRowsByProduct.ContainsKey($pp)){$productRowsByProduct[$pp]=New-Object System.Collections.ArrayList}
        [void]$productRowsByProduct[$pp].Add($row)
    }
    # HOP subfolder decision is calculated once per confirmed product path and reused for all rows.
    $hopFolderSelectionCache=@{}
    # When one confirmed product contains several HOP variant folders under the same 03_Mashini,
    # raw base PGMX names may collide (e.g. 01_Bottom_Shelf in 90x200 and 120x200). In that case
    # conversion is done in a LOCAL temporary folder and only ProgramCodeF is copied to 2_PWX_500.
    $variantConversionOutDir=@{}
    $tempConversionRoot=Join-Path ([System.IO.Path]::GetTempPath()) ('PWX500_Agent_V3210_'+[Guid]::NewGuid().ToString('N'))

    $output=New-Object System.Collections.ArrayList
    [void]$output.Add('CSV;Product;DetailNo;DetailName;ProcessingNo;ExpectedSuffix;ProgramCodeF;ProductPath;PwxFolder;SourceHop;BasePgmx;TargetPgmx;FolderStatus;HopStatus;BaseStatus;TargetStatus;Validation;Action;Result;Error;CentralFolder;CentralStatus;CentralError')
    $todo=@($rows | Where-Object { -not [string]::IsNullOrWhiteSpace((S $_.ProgramCodeF)) })
    $done=0
    # Execute mode: converter is launched at most once per product, with ALL CSV-linked missing HOPs in one batch. No dimension pre-filter.
    $conversionBatchDone=@{}
    $centralRequirementCache=@{}

    foreach($row in $todo){
        $csv=S $row.CsvFile;$product=S $row.Product;$detailNoText=S $row.DetailNo;$detailName=S $row.DetailName;$program=S $row.ProgramCodeF
        $l=Parse-Double (S $row.Length);$w=Parse-Double (S $row.Width);$t=Parse-Double (S $row.Thickness)
        $processingNo=1;if(-not [int]::TryParse((S $row.ProcessingNo),[ref]$processingNo) -or $processingNo -lt 1){$processingNo=1}
        $detailNo=0;$validDetailNo=[int]::TryParse($detailNoText,[ref]$detailNo)
        $physicalKey=$csv+'|'+$product+'|'+$detailNoText+'|'+$detailName
        $procCount=1;if($physicalCount.ContainsKey($physicalKey)){$procCount=[int]$physicalCount[$physicalKey]}
        $expectedProc=0;$suffix='-';if($procCount -gt 1){$expectedProc=$processingNo;$suffix='_'+[string]$processingNo}

        $productPath='';$pwxFolder='';$sourceHop='';$baseFile='';$targetFile=Get-SafeTargetFile $program;$targetPath=''
        $folderStatus='';$hopStatus='';$baseStatus='';$targetStatus='';$validation='';$action='';$result='';$errorText=''
        $centralStatus='';$centralError=''

        if(-not $validDetailNo){$result='ГРЕШКА';$action='ПРОПУСНАТ';$errorText='Невалиден № детайл.'}
        elseif($procCount -gt 2){$result='ГРЕШКА';$action='РЪЧНА ПРОВЕРКА';$errorText=('За физическия детайл има '+[string]$procCount+' CSV реда. Очаквани са 1 или 2 обработки.')}
        elseif([string]::IsNullOrWhiteSpace($targetFile)){$result='ГРЕШКА';$action='ПРОПУСНАТ';$errorText='Невалиден ProgramCodeF за име на файл.'}
        elseif(-not $mapped.ContainsKey($product)){$result='НЯМА ПОТВЪРДЕНА ПАПКА';$action='ПЪРВО HOP/МИГРАЦИЯ ПОТВЪРДИ ИЗДЕЛИЕТО'}
        else {
            $productPath=S $mapped[$product]
            if(-not (Test-IsUnderAllowedRoot -Path $productPath -Roots @($roots))){$result='ГРЕШКА';$action='ПРОПУСНАТ';$errorText='Папката е извън разрешените корени.'}
            elseif(-not [System.IO.Directory]::Exists($productPath)){$result='ГРЕШКА';$action='ПРОПУСНАТ';$errorText='Папката на изделието липсва или не е достъпна.'}
            else {
                $machineFolder=Join-Path $productPath '03_Mashini'
                if(-not [System.IO.Directory]::Exists($machineFolder)){$result='ГРЕШКА';$action='ПРОПУСНАТ';$errorText='Липсва 03_Mashini.'}
                else {
                    $pwxFolder=Join-Path $machineFolder '2_PWX_500';$holzRoot=Join-Path $machineFolder '1_Holzher';$holzFolder=$holzRoot
                    $folderExists=[System.IO.Directory]::Exists($pwxFolder)
                    $folderStatus=if($folderExists){'СЪЩЕСТВУВА'}else{'ЛИПСВА'}
                    $targetPath=Join-Path $pwxFolder $targetFile

                    if($folderExists -and [System.IO.File]::Exists($targetPath)){
                        $targetStatus='СЪЩЕСТВУВА';$baseStatus='НЕ Е НУЖЕН';$hopStatus='НЕ Е НУЖЕН';$validation='ProgramCodeF вече съществува.';$action='БЕЗ ДЕЙСТВИЕ';$result='ГОТОВ'
                    }
                    else {
                        # Resolve the HOP source folder once per CSV product + confirmed product path.
                        # This is needed BEFORE trusting a base PGMX: when several HOP variant folders share
                        # one 2_PWX_500, raw base names can collide between sizes.
                        $selKey=(S $product).ToLowerInvariant()+'|'+(Normalize-Path $productPath)
                        if(-not $hopFolderSelectionCache.ContainsKey($selKey)){
                            $pr=@();if($productRowsByProduct.ContainsKey($product)){$pr=@($productRowsByProduct[$product])}
                            $hopFolderSelectionCache[$selKey]=Resolve-ProductHopFolder -HolzRoot $holzRoot -ProductRows $pr -PhysicalCount $physicalCount
                        }
                        $hopFolderSelection=$hopFolderSelectionCache[$selKey]
                        $multiHopFolders=([int]$hopFolderSelection.Count -gt 1)

                        # If several variant folders exist and we cannot uniquely select one, DO NOT trust/copy
                        # a generic base PGMX from the shared 2_PWX_500. Leave the row manual instead.
                        if($multiHopFolders -and $hopFolderSelection.Status -ne 'SELECTED'){
                            $baseStatus='НЕ СЕ ИЗПОЛЗВА';$hopStatus='НЕЕДНОЗНАЧНА HOP ПАПКА';$action='РЪЧНА ПРОВЕРКА';$result='НЕЕДНОЗНАЧНА HOP ПОДПАПКА';$errorText=[string]$hopFolderSelection.Reason
                        }
                        else {
                            if($multiHopFolders){
                                $base=[pscustomobject]@{Status='MISSING';Path='';File='';Count=0;Reason='При няколко HOP вариантни папки основният PGMX не се преизползва, за да няма смесване между размери.'}
                            } else {
                                $base=Resolve-BasePgmxFast -PwxFolder $pwxFolder -DetailNo $detailNo -ExpectedProc $expectedProc -DetailName $detailName
                            }

                            if($base.Status -eq 'FOUND'){
                                $baseFile=[string]$base.File;$baseStatus='НАМЕРЕН';$hopStatus='НЕ Е НУЖЕН';$validation=[string]$base.Reason
                                if($Mode -eq 'Check'){$targetStatus='ЛИПСВА';$action='ЗА КОПИРАНЕ';$result='ГОТОВ ЗА СЪЗДАВАНЕ'}
                                else {
                                    try {
                                        if(-not $folderExists){[void][System.IO.Directory]::CreateDirectory($pwxFolder);$folderExists=$true;$folderStatus='СЪЗДАДЕНА'}
                                        $targetPath=Join-Path $pwxFolder $targetFile
                                        if([System.IO.File]::Exists($targetPath)){$targetStatus='СЪЩЕСТВУВА';$action='БЕЗ ДЕЙСТВИЕ';$result='ГОТОВ'}
                                        else {[System.IO.File]::Copy([string]$base.Path,$targetPath,$false);$targetStatus='СЪЗДАДЕН';$action='КОПИРАН';$result='СЪЗДАДЕН ProgramCodeF'}
                                    } catch {$targetStatus='ГРЕШКА';$action='ПРОПУСНАТ';$result='ГРЕШКА';$errorText='Копиране на ProgramCodeF: '+$_.Exception.Message}
                                }
                            }
                            elseif($base.Status -eq 'AMBIGUOUS'){$baseStatus='НЯКОЛКО ОСНОВНИ PGMX ('+[string]$base.Count+')';$action='РЪЧНА ПРОВЕРКА';$result='НЕЕДНОЗНАЧЕН ОСНОВЕН PGMX';$validation=[string]$base.Reason}
                            elseif($base.Status -eq 'ERROR'){$baseStatus='ГРЕШКА';$action='ПРОПУСНАТ';$result='ГРЕШКА';$errorText=[string]$base.Reason}
                            else {
                                $baseStatus=if($multiHopFolders){'НЕ СЕ ПРЕИЗПОЛЗВА'}else{'ЛИПСВА'}
                                if($hopFolderSelection.Status -ne 'SELECTED'){
                                    $hopStatus=if($hopFolderSelection.Status -eq 'AMBIGUOUS'){'НЕЕДНОЗНАЧНА HOP ПАПКА'}else{'HOP ПАПКА ЛИПСВА'}
                                    $action='РЪЧНА ПРОВЕРКА';$result=if($hopFolderSelection.Status -eq 'AMBIGUOUS'){'НЕЕДНОЗНАЧНА HOP ПОДПАПКА'}else{'НЯМА ВАЛИДНА HOP ПАПКА'}
                                    $errorText=[string]$hopFolderSelection.Reason
                                }
                                else {
                                    $holzFolder=[string]$hopFolderSelection.Path
                                    # AFTER folder choice, no dimension/content filter before V2.9.
                                    $hop=Resolve-SourceHopForConversion -HolzFolder $holzFolder -DetailNo $detailNo -ExpectedProc $expectedProc -DetailName $detailName
                                    if($hop.Status -eq 'FOUND'){
                                        $sourceHop=[string]$hop.Path;$hopStatus='НАМЕРЕН';$validation=([string]$hopFolderSelection.Reason+' | '+[string]$hop.Reason)
                                        if($Mode -eq 'Check'){
                                            $targetStatus='ЛИПСВА';$action=if($folderExists){'ГОТОВ ЗА HOP -> PGMX'}else{'ЩЕ СЕ СЪЗДАДЕ 2_PWX_500 + КОНВЕРСИЯ'};$result='ГОТОВ ЗА КОНВЕРСИЯ'
                                        }
                                        else {
                                            $exitCode=-999
                                            try {
                                                if(-not $folderExists){[void][System.IO.Directory]::CreateDirectory($pwxFolder);$folderExists=$true;$folderStatus='СЪЗДАДЕНА'}
                                                $convOutDir=$pwxFolder
                                                if($multiHopFolders){
                                                    if(-not $variantConversionOutDir.ContainsKey($selKey)){
                                                        if(-not [System.IO.Directory]::Exists($tempConversionRoot)){[void][System.IO.Directory]::CreateDirectory($tempConversionRoot)}
                                                        $variantDir=Join-Path $tempConversionRoot ([Guid]::NewGuid().ToString('N'))
                                                        [void][System.IO.Directory]::CreateDirectory($variantDir)
                                                        $variantConversionOutDir[$selKey]=$variantDir
                                                    }
                                                    $convOutDir=[string]$variantConversionOutDir[$selKey]
                                                }
                                                if(-not $conversionBatchDone.ContainsKey($selKey)) {
                                                    Write-Progress ("PGMX V3.2.11: пакетна конверсия за {0}..." -f $product)
                                                    $batchHops=New-Object System.Collections.ArrayList
                                                    foreach($br in @($todo | Where-Object { (S $_.Product) -eq $product })) {
                                                        $bdnText=S $br.DetailNo;$bdn=0
                                                        if(-not [int]::TryParse($bdnText,[ref]$bdn)){continue}
                                                        $bproc=1;if(-not [int]::TryParse((S $br.ProcessingNo),[ref]$bproc) -or $bproc -lt 1){$bproc=1}
                                                        $bkey=(S $br.CsvFile)+'|'+$product+'|'+$bdnText+'|'+(S $br.DetailName)
                                                        $bcnt=1;if($physicalCount.ContainsKey($bkey)){$bcnt=[int]$physicalCount[$bkey]}
                                                        if($bcnt -gt 2){continue}
                                                        $bexp=0;if($bcnt -gt 1){$bexp=$bproc}
                                                        $btgt=Get-SafeTargetFile (S $br.ProgramCodeF)
                                                        if([string]::IsNullOrWhiteSpace($btgt)){continue}
                                                        $btgtPath=Join-Path $pwxFolder $btgt
                                                        if([System.IO.File]::Exists($btgtPath)){continue}
                                                        if(-not $multiHopFolders){
                                                            $bbase=Resolve-BasePgmxFast -PwxFolder $pwxFolder -DetailNo $bdn -ExpectedProc $bexp -DetailName (S $br.DetailName)
                                                            if($bbase.Status -eq 'FOUND'){continue}
                                                            if($bbase.Status -eq 'AMBIGUOUS' -or $bbase.Status -eq 'ERROR'){continue}
                                                        }
                                                        $bhop=Resolve-SourceHopForConversion -HolzFolder $holzFolder -DetailNo $bdn -ExpectedProc $bexp -DetailName (S $br.DetailName)
                                                        if($bhop.Status -eq 'FOUND' -and -not [string]::IsNullOrWhiteSpace((S $bhop.Path))){[void]$batchHops.Add([string]$bhop.Path)}
                                                    }
                                                    $conv=Invoke-ConverterBatchFast -HopPaths @($batchHops) -OutDir $convOutDir
                                                    $conversionBatchDone[$selKey]=$conv
                                                } else {$conv=$conversionBatchDone[$selKey]}
                                                $exitCode=[int]$conv.ExitCode
                                                $baseDest=Join-Path $convOutDir (([System.IO.Path]::GetFileNameWithoutExtension([string]$hop.Path))+'.pgmx')
                                                if(-not [System.IO.File]::Exists($baseDest)){throw ([string]$conv.Message)}
                                                $baseFile=if($multiHopFolders){'TEMP:'+[System.IO.Path]::GetFileName($baseDest)}else{[System.IO.Path]::GetFileName($baseDest)}
                                                $baseStatus=if($multiHopFolders){'КОНВЕРТИРАН ВРЕМЕННО'}else{'КОНВЕРТИРАН ПАКЕТНО'}
                                                $targetPath=Join-Path $pwxFolder $targetFile
                                                if(-not [System.IO.File]::Exists($targetPath)){[System.IO.File]::Copy($baseDest,$targetPath,$false)}
                                                $targetStatus=if([System.IO.File]::Exists($targetPath)){'СЪЗДАДЕН'}else{'ЛИПСВА'}
                                                if($targetStatus -ne 'СЪЗДАДЕН'){throw 'ProgramCodeF не беше създаден.'}
                                                $action=if($multiHopFolders){'ПАКЕТЕН HOP -> TEMP PGMX -> ProgramCodeF'}else{'ПАКЕТЕН HOP -> PGMX + КОПИЕ'}
                                                $result='КОНВЕРТИРАН + ProgramCodeF'
                                                Append-ConversionLog $product $detailNoText $program ([string]$hop.Path) $baseDest $result ([string]$conv.Message) $exitCode
                                            }
                                            catch {
                                                $targetStatus='НЕ Е СЪЗДАДЕН';$action='ПРОПУСНАТ';$result='ГРЕШКА';$errorText=$_.Exception.Message
                                                Append-ConversionLog $product $detailNoText $program ([string]$hop.Path) $baseFile $result $errorText $exitCode
                                            }
                                        }
                                    }
                                    elseif($hop.Status -eq 'AMBIGUOUS'){$hopStatus='НЕЕДНОЗНАЧЕН ('+[string]$hop.Count+')';$action='РЪЧНА ПРОВЕРКА';$result='НЕЕДНОЗНАЧЕН HOP';$errorText=[string]$hop.Reason}
                                    else {$hopStatus='ЛИПСВА / НЕ СЪВПАДА';$action='РЪЧНА ПРОВЕРКА';$result='НЯМА ВАЛИДЕН HOP';$errorText=[string]$hop.Reason}
                                }
                            }
                        }
                    }
                }
            }
        }

        # Central staging rule V3.2.11:
        # ONLY special nested-structure products are copied to SCM_PWX_500.
        # Normal products remain ONLY in their own 2_PWX_500. Base/detail-name PGMX are never copied centrally.
        if(-not [string]::IsNullOrWhiteSpace($targetFile)){
            $centralRequired=$false
            if(-not [string]::IsNullOrWhiteSpace($productPath)) {
                $centralKey=Normalize-Path $productPath
                if($centralRequirementCache.ContainsKey($centralKey)){$centralRequired=[bool]$centralRequirementCache[$centralKey]}
                else {
                    $holzForCentral=''
                    try{$holzForCentral=Join-Path (Join-Path $productPath '03_Mashini') '1_Holzher'}catch{}
                    $centralRequired=Test-CentralStagingRequired -ProductPath $productPath -HolzRoot $holzForCentral -Roots @($roots)
                    $centralRequirementCache[$centralKey]=$centralRequired
                }
            }
            if(-not $centralRequired){
                $centralStatus='НЕ СЕ ИЗИСКВА'
            }
            else {
                $centralTarget=Join-Path $CentralProgramFolder $targetFile
                if($Mode -eq 'Check'){
                    if([System.IO.File]::Exists($centralTarget)){$centralStatus='СЪЩЕСТВУВА'}
                    elseif($result -eq 'ГОТОВ' -or $result -eq 'ГОТОВ ЗА СЪЗДАВАНЕ' -or $result -eq 'ГОТОВ ЗА КОНВЕРСИЯ'){$centralStatus='ЩЕ СЕ КОПИРА'}
                    else{$centralStatus='НЯМА ДА СЕ КОПИРА'}
                }
                else {
                    $rowSuccessful=($result -eq 'ГОТОВ' -or $result -eq 'СЪЗДАДЕН ProgramCodeF' -or $result -eq 'КОНВЕРТИРАН + ProgramCodeF')
                    if($rowSuccessful -and -not [string]::IsNullOrWhiteSpace($targetPath) -and [System.IO.File]::Exists($targetPath)){
                        try{
                            if(-not [System.IO.Directory]::Exists($CentralProgramFolder)){[void][System.IO.Directory]::CreateDirectory($CentralProgramFolder)}
                            if([System.IO.File]::Exists($centralTarget)){$centralStatus='СЪЩЕСТВУВА'}
                            else{[System.IO.File]::Copy($targetPath,$centralTarget,$false);$centralStatus='КОПИРАН'}
                        }catch{$centralStatus='ГРЕШКА';$centralError=$_.Exception.Message}
                    } else {$centralStatus='ПРОПУСНАТ'}
                }
            }
        }

        Add-ResultLine -Output $output -Values @($csv,$product,$detailNoText,$detailName,[string]$processingNo,$suffix,$program,$productPath,$pwxFolder,$sourceHop,$baseFile,$targetFile,$folderStatus,$hopStatus,$baseStatus,$targetStatus,$validation,$action,$result,$errorText,$CentralProgramFolder,$centralStatus,$centralError)
        $done++
        if(($done % 10) -eq 0 -or $done -eq $todo.Count){Write-Progress ("PGMX V3.2.11 {0}: {1} / {2} ProgramCodeF..." -f $Mode,$done,$todo.Count)}
    }

    try{if([System.IO.Directory]::Exists($tempConversionRoot)){[System.IO.Directory]::Delete($tempConversionRoot,$true)}}catch{}
    [System.IO.File]::WriteAllLines($ResultsCsv,@($output),$utf8Bom)
    try{[System.IO.File]::WriteAllText($DoneFile,("OK|{0}" -f $done),$utf8NoBom)}catch{}
    Write-Progress ("PGMX V3.2.11 {0} готово: {1} ProgramCodeF." -f $Mode,$done)
    [Environment]::ExitCode=0
    return
}
catch {
    $msg='PGMX V3.2.11 ERROR: '+$_.Exception.Message
    try{$msg+=' | Line: '+[string]$_.InvocationInfo.ScriptLineNumber}catch{}
    try{[System.IO.File]::WriteAllText($DoneFile,('ERROR|'+$msg),$utf8NoBom)}catch{}
    Write-Progress $msg
    [Environment]::ExitCode=1
    return
}
