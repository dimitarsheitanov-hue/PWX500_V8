param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$IndexCsv,
    [Parameter(Mandatory=$true)][string]$SettingsCsv,
    [Parameter(Mandatory=$true)][string]$SettingsVersion,
    [Parameter(Mandatory=$true)][string]$CodesOut,
    [Parameter(Mandatory=$true)][string]$ProductsOut,
    [Parameter(Mandatory=$true)][string]$ProgressFile,
    [Parameter(Mandatory=$true)][string]$DoneFile
)

# PWX500 Agent V5.1.35 - safe one-pass global PGMX filename index; reduced SMB metadata calls.
# SAFETY:
# - Reads ONLY the selected .xlsx workbook and local settings CSV.
# - Traverses DIRECTORY metadata only under explicit whitelist roots.
# - Inside direct 03_Mashini\2_PWX_500 folders it enumerates *.pgmx FILENAMES only.
# - PGMX content is NEVER opened or parsed by this historical worker.
# - Word files are NEVER opened, parsed, previewed, indexed or launched.
# - Nothing on the server is created, edited, renamed, moved or deleted.
# MATCHING:
# - Exact ProgramCodeF filename -> physical product directory.
# - Evidence is counted by DISTINCT physical detail number/name, not by row count.
# - Two distinct physical details in one product directory are enough to confirm it.
# - All historical codes are matched against the in-memory index, so there is no
#   per-product recursive server search and no four-probe limitation.

$ErrorActionPreference='Stop'
$utf8Bom=New-Object System.Text.UTF8Encoding($true)
$utf8NoBom=New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$MinEvidenceDetails=2
$excludedPrefixes=@(
    'Y:\1_Izpolzvani\01_KM\_0_10_F\01_Aksesoari',
    'Y:\1_Izpolzvani\01_KM\_06_Pac_s12_B#\01_Aksesoari'
)

function S($v){ if($null -eq $v){return ''}; return ([string]$v).Trim() }
function CsvField([string]$s){
    if($null -eq $s){$s=''}
    if($s.Contains('"')){$s=$s.Replace('"','""')}
    if($s.Contains(';') -or $s.Contains('"') -or $s.Contains("`r") -or $s.Contains("`n")){return '"'+$s+'"'}
    return $s
}
function Write-ProgressText([string]$text){ try{[System.IO.File]::WriteAllText($ProgressFile,$text,$utf8NoBom)}catch{} }
function Normalize-Path([string]$path){
    if([string]::IsNullOrWhiteSpace($path)){return ''}
    return $path.Trim().Replace('/','\').TrimEnd('\').ToLowerInvariant()
}
function Is-ExcludedPath([string]$path){
    $p=Normalize-Path $path
    foreach($ex in $excludedPrefixes){
        $e=Normalize-Path $ex
        if($p -eq $e -or $p.StartsWith($e+'\',[System.StringComparison]::OrdinalIgnoreCase)){return $true}
    }
    return $false
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

function Is-UnderAllowedRoot([string]$path,$roots){
    $p=Normalize-AllowedRootAliasPath $path
    if([string]::IsNullOrWhiteSpace($p)){return $false}
    foreach($root in @($roots)){
        $r=Normalize-AllowedRootAliasPath ([string]$root)
        if([string]::IsNullOrWhiteSpace($r)){continue}
        if($p -eq $r -or $p.StartsWith($r+'\',[System.StringComparison]::OrdinalIgnoreCase)){return $true}
    }
    return $false
}
function Normalize-Code([string]$code){
    $c=S $code
    if([string]::IsNullOrWhiteSpace($c)){return ''}
    $c=[System.IO.Path]::GetFileName($c)
    if($c -match '^\d+$'){$c=$c+'.pgmx'}
    if(-not $c.EndsWith('.pgmx',[System.StringComparison]::OrdinalIgnoreCase)){return ''}
    return $c
}
function Get-ZipText($zip,[string]$name){
    $entry=$zip.GetEntry($name)
    if($null -eq $entry){return $null}
    $stream=$null;$reader=$null
    try{
        $stream=$entry.Open()
        $reader=New-Object System.IO.StreamReader -ArgumentList @($stream,[System.Text.Encoding]::UTF8,$true)
        return $reader.ReadToEnd()
    } finally {
        if($null -ne $reader){$reader.Dispose()} elseif($null -ne $stream){$stream.Dispose()}
    }
}
function New-ReadOnlyWorkbookSnapshot([string]$path){
    $lastError=''
    for($attempt=1;$attempt -le 4;$attempt++){
        $source=$null;$dest=$null;$tempPath=''
        try{
            $tempPath=[System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),('PWX500_History_'+[guid]::NewGuid().ToString('N')+'.xlsx'))
            # FileShare.ReadWrite + Delete lets us read the last saved file even while Excel keeps it open.
            # The original workbook is never modified; all parsing happens on this local temporary snapshot.
            $share=[System.IO.FileShare]7
            $source=[System.IO.File]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,$share)
            $dest=[System.IO.File]::Open($tempPath,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None)
            $source.CopyTo($dest)
            $dest.Flush()
            $dest.Dispose();$dest=$null
            $source.Dispose();$source=$null
            return $tempPath
        }
        catch{
            $lastError=$_.Exception.Message
            try{if($null -ne $dest){$dest.Dispose()}}catch{}
            try{if($null -ne $source){$source.Dispose()}}catch{}
            try{if(-not [string]::IsNullOrWhiteSpace($tempPath) -and [System.IO.File]::Exists($tempPath)){[System.IO.File]::Delete($tempPath)}}catch{}
            if($attempt -lt 4){Start-Sleep -Milliseconds (250 * $attempt)}
        }
    }
    throw ('Excel workbook could not be opened for shared read after 4 attempts. '+$lastError)
}
function Get-CellText($cell,$shared){
    if($null -eq $cell){return ''}
    $type=S ($cell.GetAttribute('t'))
    if($type -eq 's'){
        $v=$cell.SelectSingleNode("./*[local-name()='v']")
        if($null -eq $v){return ''}
        $i=0
        if([int]::TryParse((S $v.InnerText),[ref]$i) -and $i -ge 0 -and $i -lt $shared.Count){return [string]$shared[$i]}
        return ''
    }
    if($type -eq 'inlineStr'){
        $texts=@($cell.SelectNodes(".//*[local-name()='t']"))
        $parts=New-Object System.Collections.ArrayList
        foreach($t in $texts){[void]$parts.Add([string]$t.InnerText)}
        return (@($parts) -join '')
    }
    $v=$cell.SelectSingleNode("./*[local-name()='v']")
    if($null -ne $v){return [string]$v.InnerText}
    $texts=@($cell.SelectNodes(".//*[local-name()='t']"))
    if($texts.Count -gt 0){return (@($texts | ForEach-Object {$_.InnerText}) -join '')}
    return ''
}
function Parse-Description([string]$description){
    $product='';$detailName='';$detailNo=''
    $parts=@([regex]::Split((S $description),'\s+-\s+'))
    if($parts.Count -ge 3){
        $last=S $parts[$parts.Count-1]
        $n=0
        if([int]::TryParse($last,[ref]$n)){
            $detailNo=[string]$n
            $detailName=S $parts[$parts.Count-2]
            $pp=New-Object System.Collections.ArrayList
            for($i=0;$i -le $parts.Count-3;$i++){
                $x=S $parts[$i]
                if(-not [string]::IsNullOrWhiteSpace($x)){[void]$pp.Add($x)}
            }
            $product=(@($pp) -join ' - ').Trim()
        }
    }
    return [pscustomobject]@{Product=$product;DetailName=$detailName;DetailNo=$detailNo}
}
function Normalize-Name([string]$Value){
    if($null -eq $Value){return ''}
    $s=$Value.Trim().ToLowerInvariant()
    $s=[regex]::Replace($s,'(?i)(\d+)\s*[/x*]\s*(\d+)','$1x$2')
    $s=[regex]::Replace($s,'(?i)([a-z])#','$1hash')
    $s=$s.Replace('#',' hash ')
    $s=$s.Replace('_',' ').Replace('-',' ').Replace('\',' ').Replace('/',' ')
    $s=[regex]::Replace($s,'[^a-z0-9\u0400-\u04ff ]',' ')
    $s=[regex]::Replace($s,'\s+',' ')
    return $s.Trim()
}
function Get-ExpectedCategoryMarker([string]$Product){
    $n=Normalize-Name $Product
    if($n -match '^dresser(?:\s|$)'){return 'skrinove'}
    if($n -match '^bedroom and bed(?:\s|$)' -or $n -match '^bed(?:\s|$)'){return 'spalni_i_legla'}
    if($n -match '^coffee(?:\s|$)' -or $n -match '^dining table(?:\s|$)'){return 'masi'}
    if($n -match '^hallway set(?:\s|$)'){return '05_koridor'}
    if($n -match '^kitchen(?:\s|$)'){return 'kuhni'}
    if($n -match '^upgrade wardrobe(?:\s|$)' -or $n -match '^wardrobe(?:\s|$)'){return 'garderobi'}
    if($n -match '^section(?:\s|$)'){return 'sekcii'}
    return ''
}
function Path-MatchesExpectedCategory([string]$Path,[string]$Marker){
    if([string]::IsNullOrWhiteSpace($Marker)){return $false}
    foreach($segment in @($Path -split '\\')){
        if(([string]$segment).Trim().ToLowerInvariant().Contains($Marker.ToLowerInvariant())){return $true}
    }
    return $false
}
function Contains-TokenCI($tokens,[string]$value){
    foreach($t in @($tokens)){
        if([string]::Equals([string]$t,$value,[System.StringComparison]::OrdinalIgnoreCase)){return $true}
    }
    return $false
}
function Product-PathCompatible([string]$Product,[string]$Path){
    # Exact ProgramCodeF is the primary evidence. Names are used only to block explicit
    # contradictions such as F vs F# or 180x200 vs 160x200.
    $pn=Normalize-Name $Product
    $qn=Normalize-Name $Path
    $pt=@($pn -split '\s+' | Where-Object {-not [string]::IsNullOrWhiteSpace($_)})
    $qt=@($qn -split '\s+' | Where-Object {-not [string]::IsNullOrWhiteSpace($_)})
    foreach($base in @('f','b','c','a','l','r')){
        $ph=Contains-TokenCI $pt ($base+'hash')
        $pp=Contains-TokenCI $pt $base
        $qh=Contains-TokenCI $qt ($base+'hash')
        $qp=Contains-TokenCI $qt $base
        if($ph -and $qp -and -not $qh){return $false}
        if($pp -and $qh -and -not $qp){return $false}
    }
    $pd=@([regex]::Matches($pn,'\b\d{2,4}x\d{2,4}\b') | ForEach-Object {$_.Value})
    $qd=@([regex]::Matches($qn,'\b\d{2,4}x\d{2,4}\b') | ForEach-Object {$_.Value})
    if($pd.Count -gt 0 -and $qd.Count -gt 0){
        $same=$false
        foreach($a in $pd){foreach($b in $qd){if([string]::Equals([string]$a,[string]$b,[System.StringComparison]::OrdinalIgnoreCase)){$same=$true;break}};if($same){break}}
        if(-not $same){return $false}
    }
    return $true
}
function Detail-Key($r){
    $detailNo=S $r.DetailNo
    if(-not [string]::IsNullOrWhiteSpace($detailNo)){return 'N:'+($detailNo.ToLowerInvariant())}
    $name=S $r.DetailName
    if(-not [string]::IsNullOrWhiteSpace($name)){return 'T:'+($name.ToLowerInvariant())}
    return 'R:'+(S $r.ExcelRow)
}
function Get-OrCreatePathEvidence($map,[string]$path,[string]$pwx){
    $key=Normalize-Path $path
    if(-not $map.ContainsKey($key)){
        $map[$key]=[pscustomobject]@{
            Path=$path
            Pwx=$pwx
            Details=(New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase))
            Codes=(New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase))
        }
    }
    return $map[$key]
}

$zip=$null
$workbookSnapshot=''
try{
    Write-ProgressText 'V0.2.20: validating Excel and allowed roots...'
    if(-not [System.IO.File]::Exists($WorkbookPath)){throw 'Selected Excel file does not exist.'}
    if([System.IO.Path]::GetExtension($WorkbookPath).ToLowerInvariant() -ne '.xlsx'){throw 'Only .xlsx is allowed for historical migration.'}
    if(-not [System.IO.File]::Exists($SettingsCsv)){throw 'Allowed folders settings are missing.'}

    $roots=New-Object System.Collections.ArrayList
    foreach($st in @(Import-Csv -LiteralPath $SettingsCsv -Delimiter ';')){
        if((S ($st.Version)) -ne $SettingsVersion){continue}
        $rp=S ($st.Path)
        if(-not [string]::IsNullOrWhiteSpace($rp) -and -not $roots.Contains($rp)){[void]$roots.Add($rp)}
    }
    if($roots.Count -eq 0){throw 'No allowed folders for current settings version.'}

    # Read historical Excel from a local read-only snapshot. This works even when Excel has
    # the original workbook open, as long as the last saved file can be shared for reading.
    Write-ProgressText 'V0.2.20: creating read-only snapshot of the selected Excel workbook...'
    $workbookSnapshot=New-ReadOnlyWorkbookSnapshot $WorkbookPath
    $zip=[System.IO.Compression.ZipFile]::OpenRead($workbookSnapshot)
    [xml]$workbookXml=Get-ZipText $zip 'xl/workbook.xml'
    [xml]$relsXml=Get-ZipText $zip 'xl/_rels/workbook.xml.rels'
    if($null -eq $workbookXml -or $null -eq $relsXml){throw 'Invalid XLSX: workbook.xml missing.'}
    $sheetNode=$null
    foreach($n in @($workbookXml.SelectNodes("//*[local-name()='sheet']"))){
        if([string]::Equals((S ($n.GetAttribute('name'))),'PWX_500',[System.StringComparison]::OrdinalIgnoreCase)){$sheetNode=$n;break}
    }
    if($null -eq $sheetNode){throw 'Worksheet PWX_500 was not found.'}
    $rid=''
    foreach($a in @($sheetNode.Attributes)){
        if($a.LocalName -eq 'id' -and $a.NamespaceURI -like '*relationships*'){$rid=[string]$a.Value;break}
    }
    if([string]::IsNullOrWhiteSpace($rid)){throw 'PWX_500 relationship id not found.'}
    $target=''
    foreach($rel in @($relsXml.SelectNodes("//*[local-name()='Relationship']"))){
        if([string]::Equals((S ($rel.GetAttribute('Id'))),$rid,[System.StringComparison]::OrdinalIgnoreCase)){$target=S ($rel.GetAttribute('Target'));break}
    }
    if([string]::IsNullOrWhiteSpace($target)){throw 'PWX_500 worksheet XML target not found.'}
    $sheetPart=$target.Replace('\','/').TrimStart([char]'/')
    if(-not $sheetPart.StartsWith('xl/')){$sheetPart='xl/'+$sheetPart.TrimStart([char]'/')}
    [xml]$sheetXml=Get-ZipText $zip $sheetPart
    if($null -eq $sheetXml){throw 'PWX_500 worksheet could not be read.'}

    $shared=New-Object System.Collections.ArrayList
    $sharedText=Get-ZipText $zip 'xl/sharedStrings.xml'
    if(-not [string]::IsNullOrWhiteSpace($sharedText)){
        [xml]$sharedXml=$sharedText
        foreach($si in @($sharedXml.SelectNodes("//*[local-name()='si']"))){
            $parts=New-Object System.Collections.ArrayList
            foreach($t in @($si.SelectNodes(".//*[local-name()='t']"))){[void]$parts.Add([string]$t.InnerText)}
            [void]$shared.Add((@($parts) -join ''))
        }
    }

    Write-ProgressText 'V0.2.20: reading PWX_500 history...'
    $records=New-Object System.Collections.ArrayList
    $productRows=@{}
    foreach($row in @($sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']"))){
        $rowNo=0
        [void][int]::TryParse((S ($row.GetAttribute('r'))),[ref]$rowNo)
        if($rowNo -le 1){continue}
        $vals=@{}
        foreach($cell in @($row.SelectNodes("./*[local-name()='c']"))){
            $ref=S ($cell.GetAttribute('r'))
            $m=[regex]::Match($ref,'^[A-Za-z]+')
            if(-not $m.Success){continue}
            $col=$m.Value.ToUpperInvariant()
            if(@('B','V','W','X') -notcontains $col){continue}
            $vals[$col]=S (Get-CellText $cell $shared)
        }
        $desc='';if($vals.ContainsKey('B')){$desc=S $vals['B']}
        $code='';if($vals.ContainsKey('V')){$code=Normalize-Code $vals['V']}
        $excluded='';if($vals.ContainsKey('W')){$excluded=S $vals['W']}
        $date='';if($vals.ContainsKey('X')){$date=S $vals['X']}
        if([string]::IsNullOrWhiteSpace($code)){continue}
        $parsed=Parse-Description $desc
        $rec=[pscustomobject]@{
            ExcelRow=[string]$rowNo
            Description=$desc
            Product=[string]$parsed.Product
            DetailNo=[string]$parsed.DetailNo
            DetailName=[string]$parsed.DetailName
            ProgramCodeF=$code
            ExcludedPWX=$excluded
            Date=$date
        }
        [void]$records.Add($rec)
        $prod=S $rec.Product
        if(-not [string]::IsNullOrWhiteSpace($prod)){
            if(-not $productRows.ContainsKey($prod)){$productRows[$prod]=New-Object System.Collections.ArrayList}
            [void]$productRows[$prod].Add($rec)
        }
    }
    if($records.Count -eq 0){throw 'No ProgramCodeF records found in PWX_500.'}
    $zip.Dispose();$zip=$null
    try{if(-not [string]::IsNullOrWhiteSpace($workbookSnapshot) -and [System.IO.File]::Exists($workbookSnapshot)){[System.IO.File]::Delete($workbookSnapshot)}}catch{}
    $workbookSnapshot=''

    # Build the SAME global filename index, but use today's product-folder catalog
    # when it is current for the active settings. This changes only discovery speed:
    # matching/evidence rules below are untouched. If the catalog is missing/stale,
    # fall back to the original full recursive scan.
    Write-ProgressText ('V5.1.35: '+$records.Count+' Excel rows, '+$productRows.Count+' products. Preparing PGMX filename index...')
    $codeIndex=@{}
    $visited=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $foldersVisited=0;$productFolders=0;$pwxFolders=0;$pgmxFiles=0;$scanErrors=0

    function Add-PgmxFolderToIndex([string]$productDir) {
        if([string]::IsNullOrWhiteSpace($productDir)){return}
        if(Is-ExcludedPath $productDir){return}
        if(-not (Is-UnderAllowedRoot -path $productDir -roots $roots)){return}
        $machineDir=Join-Path $productDir '03_Mashini'
        if(-not [System.IO.Directory]::Exists($machineDir)){return}
        $script:productFolders++
        $pwx=Join-Path $machineDir '2_PWX_500'
        if(-not [System.IO.Directory]::Exists($pwx)){return}
        $script:pwxFolders++
        try{
            foreach($file in [System.IO.Directory]::EnumerateFiles($pwx,'*.pgmx',[System.IO.SearchOption]::TopDirectoryOnly)){
                $name=[System.IO.Path]::GetFileName($file)
                if([string]::IsNullOrWhiteSpace($name)){continue}
                $key=$name.ToLowerInvariant()
                if(-not $codeIndex.ContainsKey($key)){$codeIndex[$key]=New-Object System.Collections.ArrayList}
                $already=$false
                foreach($loc in @($codeIndex[$key])){
                    if((Normalize-Path ([string]$loc.Path)) -eq (Normalize-Path $productDir)){$already=$true;break}
                }
                if(-not $already){[void]$codeIndex[$key].Add([pscustomobject]@{Path=$productDir;Pwx=$pwx})}
                $script:pgmxFiles++
            }
        } catch {$script:scanErrors++}
    }

    $usedCatalog=$false
    try {
        if([System.IO.File]::Exists($IndexCsv)){
            $idx=Get-Item -LiteralPath $IndexCsv -ErrorAction Stop
            $set=Get-Item -LiteralPath $SettingsCsv -ErrorAction Stop
            $catalogCurrent=($idx.LastWriteTime.Date -eq (Get-Date).Date -and $idx.LastWriteTimeUtc -ge $set.LastWriteTimeUtc)
            if($catalogCurrent){
                $catalogRows=@(Import-Csv -LiteralPath $IndexCsv -Delimiter ';')
                if($catalogRows.Count -gt 0){
                    Write-ProgressText ('V5.1.35 fast migration: using current daily product catalog ('+$catalogRows.Count+' product folders)...')
                    foreach($cr in $catalogRows){
                        $productDir=S ($cr.Path)
                        if([string]::IsNullOrWhiteSpace($productDir)){continue}
                        if(-not $visited.Add($productDir)){continue}
                        $foldersVisited++
                        Add-PgmxFolderToIndex $productDir
                        if(($foldersVisited % 100) -eq 0){Write-ProgressText ('V5.1.35 fast migration: product folders '+$foldersVisited+' | PWX500 '+$pwxFolders+' | PGMX '+$pgmxFiles)}
                    }
                    if($pwxFolders -gt 0){$usedCatalog=$true}
                }
            }
        }
    } catch {
        $usedCatalog=$false
        $codeIndex=@{};$visited=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $foldersVisited=0;$productFolders=0;$pwxFolders=0;$pgmxFiles=0;$scanErrors=0
    }

    if(-not $usedCatalog){
        Write-ProgressText 'V5.1.35 migration: catalog unavailable/stale; using complete safe scan with fewer network metadata calls...'
        foreach($root in @($roots)){
            if(-not [System.IO.Directory]::Exists($root)){$scanErrors++;continue}
            $q=New-Object 'System.Collections.Generic.Queue[string]'
            $q.Enqueue([string]$root)
            while($q.Count -gt 0){
                $dir=$q.Dequeue()
                if(Is-ExcludedPath $dir){continue}
                if(-not $visited.Add($dir)){continue}
                $foldersVisited++
                try{
                    # Safe speed optimization: enumerate this directory only once.
                    # V5.1.34 first tested dir\03_Mashini and then enumerated dir,
                    # which creates an extra SMB metadata request for every folder.
                    # The product marker and traversal rules remain identical.
                    $dirInfo=New-Object System.IO.DirectoryInfo($dir)
                    $childInfos=@($dirInfo.EnumerateDirectories())
                    $hasMachine=$false
                    foreach($ci in $childInfos){
                        if([string]::Equals([string]$ci.Name,'03_Mashini',[System.StringComparison]::OrdinalIgnoreCase)){$hasMachine=$true;break}
                    }
                    if($hasMachine){Add-PgmxFolderToIndex $dir;continue}
                    foreach($ci in $childInfos){
                        $child=[string]$ci.FullName
                        if(Is-ExcludedPath $child){continue}
                        try{if(($ci.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0){continue}}catch{}
                        $q.Enqueue($child)
                    }
                } catch {
                    $scanErrors++
                    # Safety fallback: if listing this folder failed, retain the
                    # exact V5.1.34 marker check so a product is not missed only
                    # because directory enumeration was denied/interrupted.
                    try{
                        $machineDir=Join-Path $dir '03_Mashini'
                        if([System.IO.Directory]::Exists($machineDir)){Add-PgmxFolderToIndex $dir}
                    }catch{}
                }
                if(($foldersVisited % 250) -eq 0){Write-ProgressText ('V5.1.35 full indexing: folders '+$foldersVisited+' | product markers '+$productFolders+' | PWX500 folders '+$pwxFolders+' | PGMX filenames '+$pgmxFiles)}
            }
        }
    }

    if($pwxFolders -eq 0){throw 'No 03_Mashini\2_PWX_500 folders were found under the allowed roots.'}
    Write-ProgressText ('V5.1.35 index ready: '+$pwxFolders+' PWX500 folders, '+$pgmxFiles+' PGMX filenames. Matching Excel in memory...')

    $codeLines=New-Object System.Collections.ArrayList
    [void]$codeLines.Add('Product;DetailNo;DetailName;ProgramCodeF;CandidatePath;PwxFolder;Exists;EvidenceRank')
    $productLines=New-Object System.Collections.ArrayList
    [void]$productLines.Add('Product;ExcelRows;Codes;FoundUnique;NotFound;Multiple;Excluded;CandidatePaths;ProposedPath;Decision;Evidence')

    $productsDone=0;$productsWithProposal=0;$productConflicts=0;$productsNoEvidence=0
    $totalCodeLookups=0;$totalFoundCodeLookups=0;$totalDistinctDetailsChecked=0

    foreach($product in @($productRows.Keys | Sort-Object)){
        $rows=@($productRows[$product])
        $pathEvidence=@{}
        $detailSeen=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $detailFound=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $codeSeen=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $excludedCount=0

        foreach($r in $rows){
            $excluded=S $r.ExcludedPWX
            if(-not [string]::IsNullOrWhiteSpace($excluded) -and $excluded -ne '0'){$excludedCount++;continue}
            $code=Normalize-Code ([string]$r.ProgramCodeF)
            if([string]::IsNullOrWhiteSpace($code)){continue}
            $detailKey=Detail-Key $r
            [void]$detailSeen.Add($detailKey)
            $codeKey=$code.ToLowerInvariant()
            if(-not $codeSeen.Add($codeKey)){continue}
            $totalCodeLookups++

            $locations=@()
            if($codeIndex.ContainsKey($codeKey)){$locations=@($codeIndex[$codeKey])}
            if($locations.Count -gt 0){
                $rank=0
                $accepted=0
                foreach($loc in $locations){
                    $rank++
                    if(-not (Product-PathCompatible -Product $product -Path ([string]$loc.Path))){continue}
                    $accepted++
                    $ev=Get-OrCreatePathEvidence -map $pathEvidence -path ([string]$loc.Path) -pwx ([string]$loc.Pwx)
                    [void]$ev.Details.Add($detailKey)
                    [void]$ev.Codes.Add($code)
                    $vals=@([string]$product,[string]$r.DetailNo,[string]$r.DetailName,$code,[string]$loc.Path,[string]$loc.Pwx,'True',[string]$rank)
                    [void]$codeLines.Add((@($vals | ForEach-Object {CsvField ([string]$_)}) -join ';'))
                }
                if($accepted -gt 0){$totalFoundCodeLookups++;[void]$detailFound.Add($detailKey)}
                if($accepted -eq 0){
                    $vals=@([string]$product,[string]$r.DetailNo,[string]$r.DetailName,$code,'','','False','0')
                    [void]$codeLines.Add((@($vals | ForEach-Object {CsvField ([string]$_)}) -join ';'))
                }
            } else {
                $vals=@([string]$product,[string]$r.DetailNo,[string]$r.DetailName,$code,'','','False','0')
                [void]$codeLines.Add((@($vals | ForEach-Object {CsvField ([string]$_)}) -join ';'))
            }
        }

        $totalDistinctDetailsChecked+=$detailSeen.Count
        $tested=New-Object System.Collections.ArrayList
        $categoryMarker=Get-ExpectedCategoryMarker $product
        foreach($k in @($pathEvidence.Keys)){
            $ev=$pathEvidence[$k]
            $pref=0
            if(-not [string]::IsNullOrWhiteSpace($categoryMarker) -and (Path-MatchesExpectedCategory -Path ([string]$ev.Path) -Marker $categoryMarker)){$pref=1}
            [void]$tested.Add([pscustomobject]@{
                Path=[string]$ev.Path
                Pwx=[string]$ev.Pwx
                Hits=[int]$ev.Details.Count
                HitCodes=(@($ev.Codes | Sort-Object) -join ', ')
                PreferredCategory=[int]$pref
            })
        }

        $sorted=@($tested | Sort-Object -Property @{Expression={$_.Hits};Descending=$true},@{Expression={$_.PreferredCategory};Descending=$true},@{Expression={$_.Path};Descending=$false})
        $qualified=@($sorted | Where-Object {[int]$_.Hits -ge $MinEvidenceDetails})
        $bestHits=0
        if($sorted.Count -gt 0){$bestHits=[int]$sorted[0].Hits}
        $missing=[Math]::Max(0,$detailSeen.Count-$detailFound.Count)
        $multiple=$qualified.Count
        $candidatePaths=(@($sorted | ForEach-Object {'['+[string]$_.Hits+' details] '+[string]$_.Path}) -join ' | ')
        $proposed='';$decision='NO_TWO_DETAIL_EVIDENCE';$evidence=''

        if($detailSeen.Count -lt $MinEvidenceDetails){
            $decision='NOT_ENOUGH_DISTINCT_DETAILS';$productsNoEvidence++
        } elseif($qualified.Count -eq 1){
            $proposed=[string]$qualified[0].Path
            $decision='MAP_UNIQUE_PATH'
            $evidence=[string]$qualified[0].Hits+' distinct details by exact ProgramCodeF: '+[string]$qualified[0].HitCodes
            $productsWithProposal++
        } elseif($qualified.Count -gt 1){
            $topHits=[int]$qualified[0].Hits
            $sameTop=@($qualified | Where-Object {[int]$_.Hits -eq $topHits})
            if($sameTop.Count -eq 1){
                $proposed=[string]$sameTop[0].Path
                $decision='MAP_UNIQUE_PATH'
                $evidence=[string]$sameTop[0].Hits+' distinct details; strongest exact ProgramCodeF evidence'
                $productsWithProposal++
            } else {
                $preferred=@($sameTop | Where-Object {[int]$_.PreferredCategory -eq 1})
                if($preferred.Count -eq 1){
                    $proposed=[string]$preferred[0].Path
                    $decision='MAP_UNIQUE_PATH'
                    $evidence=[string]$preferred[0].Hits+' distinct details; category used only as tie-break'
                    $productsWithProposal++
                } else {
                    $decision='MULTIPLE_PRODUCT_PATHS';$productConflicts++
                }
            }
        } else {
            $productsNoEvidence++
        }

        $vals=@([string]$product,[string]$rows.Count,[string]$detailSeen.Count,[string]$bestHits,[string]$missing,[string]$multiple,[string]$excludedCount,$candidatePaths,$proposed,$decision,$evidence)
        [void]$productLines.Add((@($vals | ForEach-Object {CsvField ([string]$_)}) -join ';'))
        $productsDone++
        if(($productsDone % 25) -eq 0 -or $productsDone -eq $productRows.Count){
            Write-ProgressText ('V0.2.20 matching: '+$productsDone+' / '+$productRows.Count+' products | confirmed '+$productsWithProposal+' | conflicts '+$productConflicts)
        }
    }

    [System.IO.File]::WriteAllLines($CodesOut,[string[]]$codeLines.ToArray(),$utf8Bom)
    [System.IO.File]::WriteAllLines($ProductsOut,[string[]]$productLines.ToArray(),$utf8Bom)
    $done='OK;Records='+$totalCodeLookups+';ExcelRows='+$records.Count+';UniqueFound='+$totalFoundCodeLookups+';NotFound='+[string]([Math]::Max(0,$totalCodeLookups-$totalFoundCodeLookups))+';Multiple='+$productConflicts+';Products='+$productRows.Count+';Proposals='+$productsWithProposal+';Conflicts='+$productConflicts+';IndexedFolders='+$pwxFolders+';IndexedFiles='+$pgmxFiles+';FoldersVisited='+$foldersVisited+';DistinctDetails='+$totalDistinctDetailsChecked+';Mode=GLOBAL_PGMX_INDEX'
    [System.IO.File]::WriteAllText($DoneFile,$done,$utf8NoBom)
    Write-ProgressText ('V0.2.20 complete: '+$productsWithProposal+' directories confirmed | '+$productConflicts+' conflicts | '+$productsNoEvidence+' without two-detail evidence | indexed '+$pgmxFiles+' PGMX filenames.')
    exit 0
}
catch{
    try{if($null -ne $zip){$zip.Dispose()}}catch{}
    try{if(-not [string]::IsNullOrWhiteSpace($workbookSnapshot) -and [System.IO.File]::Exists($workbookSnapshot)){[System.IO.File]::Delete($workbookSnapshot)}}catch{}
    $msg=$_.Exception.Message;$line=''
    try{$line=[string]$_.InvocationInfo.ScriptLineNumber}catch{}
    Write-ProgressText ('History ERROR: '+$msg+' | line: '+$line)
    try{[System.IO.File]::WriteAllText($DoneFile,'ERROR;'+$msg,$utf8NoBom)}catch{}
    exit 1
}
