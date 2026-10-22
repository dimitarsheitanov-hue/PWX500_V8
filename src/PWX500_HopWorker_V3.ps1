param(
    [Parameter(Mandatory=$true)][string]$RowsCsv,
    [Parameter(Mandatory=$true)][string]$SearchResultsCsv,
    [Parameter(Mandatory=$true)][string]$ResultsCsv,
    [Parameter(Mandatory=$true)][string]$ProgressFile,
    [string]$MappingsCsv = ''
)

# SAFETY V3.0:
# - Reads ONLY local CSV control files plus *.hop files.
# - HOP content is read ONLY inside an exact candidate path: 03_Mashini\1_Holzher subtree.
# - Enumerates ONLY *.hop file names recursively there, then opens only HOPs
#   whose detail number occurs in the CSV product subset.
# - Word files are NEVER enumerated for content and NEVER opened/parsed/launched.
# - No source file is modified, copied, renamed or deleted.
# - HOP auto-confirm requires ALL expected CSV checks to match exact detail number, processing suffix rule, physical L/B dimensions regardless of orientation, and thickness.
$ErrorActionPreference = 'Stop'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$latin1 = [System.Text.Encoding]::GetEncoding(28591)
$inv = [System.Globalization.CultureInfo]::InvariantCulture

function Progress([string]$t) { try { [System.IO.File]::WriteAllText($ProgressFile,$t,$utf8NoBom) } catch {} }
function CsvField([string]$s) {
    if ($null -eq $s) { $s='' }
    if ($s.Contains('"')) { $s=$s.Replace('"','""') }
    if ($s.Contains(';') -or $s.Contains('"') -or $s.Contains("`r") -or $s.Contains("`n")) { return '"'+$s+'"' }
    return $s
}
function S($v) { if ($null -eq $v) { return '' }; return ([string]$v).Trim() }


# Mapping trust model V5.1.13.
# Historical quick migration and already HOP-verified mappings do not need another HOP pass.
function Is-TrustedSource([string]$Source){
    $x=(S $Source).ToUpperInvariant()
    return ($x -in @('HISTORICAL_PGMX','HOP_AUTO','HOP_VERIFIED_MANUAL','HOP_VERIFIED_USER'))
}
function Parse-Double([string]$s) {
    $x=(S $s).Replace(',','.')
    $v=0.0
    if ([double]::TryParse($x,[System.Globalization.NumberStyles]::Float,$inv,[ref]$v)) { return $v }
    return [double]::NaN
}
function Normalize-Name([string]$s) {
    $x=(S $s).ToLowerInvariant()
    $x=$x.Replace('_',' ').Replace('-',' ')
    $x=[regex]::Replace($x,'[^a-z0-9а-яё ]',' ')
    $x=[regex]::Replace($x,'\s+',' ').Trim()
    return $x
}
function Split-TopArgs([string]$s) {
    $out=New-Object System.Collections.Generic.List[string]
    $cur=New-Object System.Text.StringBuilder
    $depth=0; $quote=[char]0
    foreach($ch in $s.ToCharArray()) {
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

# Safe arithmetic evaluator: variable names are substituted first, then only
# digits/decimal points/parentheses and + - * / are permitted. No PowerShell code is evaluated.
function Eval-SimpleExpr([string]$expr,[hashtable]$names) {
    $x=(S $expr)
    if([string]::IsNullOrWhiteSpace($x)){throw 'empty expression'}
    foreach($key in @($names.Keys | Sort-Object { $_.Length } -Descending)) {
        $val=([double]$names[$key]).ToString('R',$inv)
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
    # Absolute extension gate BEFORE content access.
    if([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.hop'){ throw 'BLOCKED: HOP reader received non-HOP file.' }
    $bytes=[System.IO.File]::ReadAllBytes($Path)
    $text=$latin1.GetString($bytes)
    $names=@{}
    $matches=[regex]::Matches($text,'(?im)^\s*(DX|DY|DZ)\s*:=\s*([^;\r\n]+)')
    foreach($m in $matches){
        $key=$m.Groups[1].Value.ToUpperInvariant()
        try{$names[$key]=Eval-SimpleExpr $m.Groups[2].Value $names}catch{}
    }
    $fm=[regex]::Match($text,'(?im)^\s*FERTIGTEIL\s*\(([^)]*)\)')
    if($fm.Success){
        $args=@(Split-TopArgs $fm.Groups[1].Value)
        if($args.Count -ge 3){
            try{$vals=@((Eval-SimpleExpr $args[0] $names),(Eval-SimpleExpr $args[1] $names),(Eval-SimpleExpr $args[2] $names));if(-not $names.ContainsKey('DX') -or [double]$names['DX'] -le 0){$names['DX']=$vals[0]};if(-not $names.ContainsKey('DY') -or [double]$names['DY'] -le 0){$names['DY']=$vals[1]};if(-not $names.ContainsKey('DZ') -or [double]$names['DZ'] -le 0){$names['DZ']=$vals[2]}}catch{}
        }
    }
    if(-not $names.ContainsKey('DX') -or -not $names.ContainsKey('DY') -or -not $names.ContainsKey('DZ')){return $null}
    if([double]$names['DX'] -le 0 -or [double]$names['DY'] -le 0 -or [double]$names['DZ'] -le 0){return $null}
    return [pscustomobject]@{DX=[double]$names['DX'];DY=[double]$names['DY'];DZ=[double]$names['DZ']}
}
function Get-HopFileMeta([string]$Path) {
    $stem=[System.IO.Path]::GetFileNameWithoutExtension($Path)
    $m=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?_(?<name>.+)$')
    if(-not $m.Success){$m=[regex]::Match($stem,'^(?<no>\d+)(?:_(?<proc>\d+))?\s*(?<name>.*)$')}
    if(-not $m.Success){return $null}
    $no=0;if(-not [int]::TryParse($m.Groups['no'].Value,[ref]$no)){return $null}
    $proc=0
    if($m.Groups['proc'].Success -and -not [string]::IsNullOrWhiteSpace($m.Groups['proc'].Value)){
        [void][int]::TryParse($m.Groups['proc'].Value,[ref]$proc)
    }
    $nm=Normalize-Name $m.Groups['name'].Value
    $dims=Get-HopDimensions $Path
    return [pscustomobject]@{Path=$Path;DetailNo=$no;ProcessingNo=$proc;NameNorm=$nm;Dims=$dims}
}
function Same-Dim([double]$a,[double]$b){return ([Math]::Abs($a-$b) -le 0.2)}
function Same-PanelSize([double]$a1,[double]$a2,[double]$b1,[double]$b2){return (((Same-Dim $a1 $b1) -and (Same-Dim $a2 $b2)) -or ((Same-Dim $a1 $b2) -and (Same-Dim $a2 $b1)))}
function Get-PanelDimensionKey([double]$l,[double]$w,[double]$t){$a=[Math]::Min($l,$w);$b=[Math]::Max($l,$w);return ($a.ToString('R',$inv)+'|'+$b.ToString('R',$inv)+'|'+$t.ToString('R',$inv))}
function Parse-Paths([string]$s){
    $out=New-Object System.Collections.Generic.List[object]
    if([string]::IsNullOrWhiteSpace($s)){return @()}
    foreach($part in @($s -split '~\|~')){
        $p=(S $part);if([string]::IsNullOrWhiteSpace($p)){continue}
        $score='';$path=$p
        $m=[regex]::Match($p,'^\[(\d+)\](.*)$');if($m.Success){$score=$m.Groups[1].Value;$path=(S $m.Groups[2].Value)}
        $out.Add([pscustomobject]@{Score=$score;Path=$path})
    }
    return @($out.ToArray())
}
function Get-HopContainingFolders([string]$HolzRoot){
    $out=New-Object System.Collections.ArrayList
    if(-not [System.IO.Directory]::Exists($HolzRoot)){return @()}
    $seen=New-Object 'System.Collections.Generic.HashSet[string]'
    try{
        foreach($hp in [System.IO.Directory]::EnumerateFiles($HolzRoot,'*.hop',[System.IO.SearchOption]::AllDirectories)){
            $dir=[System.IO.Path]::GetDirectoryName($hp)
            if(-not [string]::IsNullOrWhiteSpace($dir) -and $seen.Add($dir)){[void]$out.Add($dir)}
        }
    }catch{return @()}
    return @($out | Sort-Object)
}
function Relative-HopFolder([string]$Root,[string]$Folder){
    $r=(S $Root).TrimEnd('\');$f=S $Folder
    if([string]::Equals($r,$f,[System.StringComparison]::OrdinalIgnoreCase)){return '.'}
    if($f.StartsWith(($r+'\'),[System.StringComparison]::OrdinalIgnoreCase)){return $f.Substring($r.Length+1)}
    return $f
}

try {
    Progress 'HOP проверка: зареждане на локалните CSV резултати...'
    if(-not [System.IO.File]::Exists($RowsCsv)){throw 'Липсват прочетените CSV редове.'}
    if(-not [System.IO.File]::Exists($SearchResultsCsv)){throw 'Липсват резултатите от търсенето на изделия.'}
    $rows=@(Import-Csv -LiteralPath $RowsCsv -Delimiter ';')
    $search=@(Import-Csv -LiteralPath $SearchResultsCsv -Delimiter ';')

    $mapped=@{}
    if(-not [string]::IsNullOrWhiteSpace($MappingsCsv) -and [System.IO.File]::Exists($MappingsCsv)){
        foreach($r in @(Import-Csv -LiteralPath $MappingsCsv -Delimiter ';')){
            $mp=S $r.Product;if([string]::IsNullOrWhiteSpace($mp)){continue}
            $mapped[$mp]=[pscustomobject]@{Path=(S $r.Path);Source=(S $r.Source)}
        }
    }

    $rowsByProduct=@{}
    foreach($r in $rows){$p=S $r.Product;if([string]::IsNullOrWhiteSpace($p)){continue};if(-not $rowsByProduct.ContainsKey($p)){$rowsByProduct[$p]=New-Object System.Collections.ArrayList};[void]$rowsByProduct[$p].Add($r)}

    $out=New-Object System.Collections.ArrayList
    [void]$out.Add('Product;CandidatePath;SearchScore;ExpectedDetails;MatchedFull;MatchedName;MatchedDimensions;RelevantHopFiles;ReadableHopFiles;UnreadableHopFiles;CandidateStatus;Decision;AutoPath')
    $todo=@($search | Where-Object {
        $p=S $_.Product
        if($mapped.ContainsKey($p)){return (-not (Is-TrustedSource ([string]$mapped[$p].Source)))}
        $n=0; return ([int]::TryParse((S $_.Candidates),[ref]$n) -and $n -ge 1)
    })
    $prodDone=0

    foreach($sr in $todo){
        $product=S $sr.Product
        if(-not $rowsByProduct.ContainsKey($product)){continue}
        # Build expected HOP checks from CSV. Repeated colour/ERP variants are deduplicated,
        # but two rows for the same physical detail are preserved as processing 1 / processing 2.
        # Universal rule: when a physical detail has two CSV processing rows, row 1 MUST match
        # <detail>_1_*.hop and row 2 MUST match <detail>_2_*.hop. For a single processing row,
        # the HOP file must have no _1/_2 processing suffix. Detail number + physical L/B size + thickness are mandatory; L/B orientation itself is ignored.
        $records=@{}
        $procSets=@{}
        foreach($r in @($rowsByProduct[$product])){
            $no=0;if(-not [int]::TryParse((S $r.DetailNo),[ref]$no)){continue}
            $l=Parse-Double (S $r.Length);$w=Parse-Double (S $r.Width);$t=Parse-Double (S $r.Thickness)
            if([double]::IsNaN($l) -or [double]::IsNaN($w) -or [double]::IsNaN($t)){continue}
            $nameNorm=Normalize-Name (S $r.DetailName)
            $pn=1;if(-not [int]::TryParse((S $r.ProcessingNo),[ref]$pn) -or $pn -lt 1){$pn=1}
            $baseKey=([string]$no)+'|'+$nameNorm+'|'+(Get-PanelDimensionKey $l $w $t)
            if(-not $procSets.ContainsKey($baseKey)){
                $procSets[$baseKey]=New-Object 'System.Collections.Generic.HashSet[int]'
            }
            [void]$procSets[$baseKey].Add($pn)
            $recordKey="$baseKey|$pn"
            if(-not $records.ContainsKey($recordKey)){
                $records[$recordKey]=[pscustomobject]@{BaseKey=$baseKey;No=$no;NameNorm=$nameNorm;L=$l;W=$w;T=$t;CsvProcessingNo=$pn}
            }
        }
        $expectedList=New-Object System.Collections.ArrayList
        foreach($rec in @($records.Values)){
            $procCount=0
            if($procSets.ContainsKey([string]$rec.BaseKey)){$procCount=$procSets[[string]$rec.BaseKey].Count}
            $expectedProc=0
            if($procCount -gt 1){$expectedProc=[int]$rec.CsvProcessingNo}
            [void]$expectedList.Add([pscustomobject]@{No=[int]$rec.No;NameNorm=[string]$rec.NameNorm;L=[double]$rec.L;W=[double]$rec.W;T=[double]$rec.T;ExpectedProcessingNo=$expectedProc})
        }
        $expected=@($expectedList | Sort-Object No,ExpectedProcessingNo)
        $expectedNos=New-Object 'System.Collections.Generic.HashSet[int]'
        foreach($e in $expected){[void]$expectedNos.Add([int]$e.No)}
        $candidates=@()
        if($mapped.ContainsKey($product) -and -not (Is-TrustedSource ([string]$mapped[$product].Source))){
            $candidates=@([pscustomobject]@{Score='MANUAL';Path=[string]$mapped[$product].Path})
        } else { $candidates=@(Parse-Paths (S $sr.Paths)) }
        $candRows=New-Object System.Collections.ArrayList

        foreach($c in $candidates){
            $path=S $c.Path;$holz=Join-Path (Join-Path $path '03_Mashini') '1_Holzher'
            $full=0;$nameM=0;$dimM=0;$relevant=0;$readable=0;$unreadable=0;$status='OK'
            if(-not [System.IO.Directory]::Exists($holz)){$status='НЯМА 03_Mashini\1_Holzher'}
            else {
                $hopDirs=@(Get-HopContainingFolders $holz)
                if($hopDirs.Count -eq 0){$status='НЯМА HOP В 1_Holzher И ПОДПАПКИТЕ'}
                else {
                    $folderScores=New-Object System.Collections.ArrayList
                    foreach($hopDir in $hopDirs){
                        $hopMetas=New-Object System.Collections.ArrayList
                        $fr=0;$fread=0;$funread=0
                        try {
                            foreach($hp in [System.IO.Directory]::EnumerateFiles([string]$hopDir,'*.hop',[System.IO.SearchOption]::TopDirectoryOnly)){
                                $stem=[System.IO.Path]::GetFileNameWithoutExtension($hp)
                                $mm=[regex]::Match($stem,'^(\d+)')
                                if(-not $mm.Success){continue}
                                $dn=0;if(-not [int]::TryParse($mm.Groups[1].Value,[ref]$dn)){continue}
                                if(-not $expectedNos.Contains($dn)){continue}
                                $fr++
                                try{$meta=Get-HopFileMeta $hp;if($null -ne $meta){[void]$hopMetas.Add($meta);$fread++}else{$funread++}}catch{$funread++}
                            }
                        } catch {$funread++}
                        $ffull=0;$fname=0;$fdim=0
                        foreach($e in $expected){
                            $byNo=@($hopMetas | Where-Object {[int]$_.DetailNo -eq [int]$e.No -and [int]$_.ProcessingNo -eq [int]$e.ExpectedProcessingNo})
                            $hasName=$false;$hasDim=$false;$hasFull=$false
                            foreach($h in $byNo){
                                $nm=([string]$h.NameNorm)
                                $nameOk=(-not [string]::IsNullOrWhiteSpace($e.NameNorm) -and ($nm -eq $e.NameNorm -or $nm.Contains($e.NameNorm) -or $e.NameNorm.Contains($nm)))
                                $dimOk=$false
                                if($null -ne $h.Dims){$dimOk=(Same-PanelSize ([double]$h.Dims.DX) ([double]$h.Dims.DY) ([double]$e.L) ([double]$e.W)) -and (Same-Dim ([double]$h.Dims.DZ) ([double]$e.T))}
                                if($nameOk){$hasName=$true};if($dimOk){$hasDim=$true;$hasFull=$true;break}
                            }
                            if($hasName){$fname++};if($hasDim){$fdim++};if($hasFull){$ffull++}
                        }
                        [void]$folderScores.Add([pscustomobject]@{Folder=[string]$hopDir;Full=$ffull;Name=$fname;Dims=$fdim;Relevant=$fr;Readable=$fread;Unreadable=$funread})
                    }
                    $best=@($folderScores | Sort-Object @{Expression='Full';Descending=$true},@{Expression='Dims';Descending=$true},@{Expression='Readable';Descending=$true})
                    if($best.Count -gt 0){
                        $b=$best[0];$full=[int]$b.Full;$nameM=[int]$b.Name;$dimM=[int]$b.Dims;$relevant=[int]$b.Relevant;$readable=[int]$b.Readable;$unreadable=[int]$b.Unreadable
                        $ties=@($best | Where-Object {[int]$_.Full -eq $full -and [int]$_.Dims -eq $dimM})
                        if($hopDirs.Count -eq 1){$status='HOP ПАПКА: '+(Relative-HopFolder $holz ([string]$b.Folder))}
                        elseif($ties.Count -eq 1){$status='HOP ПОДПАПКА ПО РАЗМЕРИ: '+(Relative-HopFolder $holz ([string]$b.Folder))+' | '+[string]$hopDirs.Count+' кандидата'}
                        else{$status='HOP ПОДПАПКИ: '+[string]$hopDirs.Count+' | най-добро съвпадение '+[string]$full+'/'+[string]$expected.Count}
                    }
                }
            }
            $candRows.Add([pscustomobject]@{Product=$product;CandidatePath=$path;SearchScore=(S $c.Score);Expected=$expected.Count;Full=$full;Name=$nameM;Dims=$dimM;Relevant=$relevant;Readable=$readable;Unreadable=$unreadable;Status=$status}) | Out-Null
        }

        $fullCands=@($candRows | Where-Object { $_.Expected -gt 0 -and $_.Full -eq $_.Expected })
        $auto=''
        $autoReason=''
        if($fullCands.Count -eq 1){
            $auto=[string]$fullCands[0].CandidatePath
            $autoReason='HOP'
        }
        elseif($fullCands.Count -gt 1){
            # More than one 100% HOP match stays manual.
            $auto=''
            $autoReason=''
        }
        foreach($cr in $candRows){
            $decision='НЯМА ПЪЛНО HOP СЪВПАДЕНИЕ'
            $rowAuto=''
            if(-not [string]::IsNullOrWhiteSpace($auto)){
                if([string]::Equals([string]$cr.CandidatePath,$auto,[System.StringComparison]::OrdinalIgnoreCase)){
                    $decision='✓ АВТОМАТИЧНО ИЗБРАН ПО HOP'
                    $rowAuto=$auto
                } else {
                    $isFull=($cr.Expected -gt 0 -and $cr.Full -eq $cr.Expected)
                    $decision='ОТХВЪРЛЕН ПО HOP'
                }
            }
            elseif($fullCands.Count -gt 1){
                $isFull=($cr.Expected -gt 0 -and $cr.Full -eq $cr.Expected)
                if($isFull){$decision='ПЪЛНО СЪВПАДЕНИЕ - РЪЧНО'}else{$decision='ОТХВЪРЛЕН ПО HOP'}
            }
            $line=@(
                CsvField $cr.Product;CsvField $cr.CandidatePath;CsvField $cr.SearchScore;CsvField ([string]$cr.Expected);CsvField ([string]$cr.Full);CsvField ([string]$cr.Name);CsvField ([string]$cr.Dims);CsvField ([string]$cr.Relevant);CsvField ([string]$cr.Readable);CsvField ([string]$cr.Unreadable);CsvField $cr.Status;CsvField $decision;CsvField $rowAuto
            ) -join ';'
            [void]$out.Add($line)
        }
        $prodDone++
        Progress "HOP проверка: $prodDone / $($todo.Count) изделия..."
    }

    [System.IO.File]::WriteAllLines($ResultsCsv,[string[]]$out.ToArray(),$utf8Bom)
    Progress "HOP проверката е готова: $prodDone изделия."
    exit 0
}
catch {
    $ln='';try{$ln=[string]$_.InvocationInfo.ScriptLineNumber}catch{}
    Progress ("HOP ERROR: "+$_.Exception.Message+" | ред: "+$ln)
    exit 1
}
