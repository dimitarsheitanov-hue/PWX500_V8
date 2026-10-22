param(
    [Parameter(Mandatory=$true)][string]$ManifestPath,
    [Parameter(Mandatory=$true)][string]$ProgressPath,
    [Parameter(Mandatory=$true)][string]$StopPath,
    [Parameter(Mandatory=$true)][string]$ResultPath,
    [Parameter(Mandatory=$true)][string]$RootPath,
    [switch]$ForceImport
)

$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$DbCsv=Join-Path $RootPath 'CAM_Agent_Times_DB.csv'
$BaseCsv=Join-Path $RootPath 'CAM_Agent_Times_Base.csv'
$ActiveCsv=Join-Path $RootPath 'CAM_Agent_Times_Active.csv'
$HistoryCsv=Join-Path $RootPath 'CAM_Agent_Times_History.csv'
$RegistryCsv=Join-Path $RootPath 'CAM_Agent_Times_ImportedFiles.csv'
$ErrorCsv=Join-Path $RootPath 'CAM_Agent_Times_ImportErrors.csv'
# Shared with the single-workbook writer in the main agent.
$WriteLockPath=Join-Path $RootPath 'CAM_Agent_Times_Write.lock'
$CommitJournalPath=Join-Path $RootPath 'CAM_Agent_TimesImport_Commit.json'
$script:DbRowsCache=@{}
$script:ActiveRowsCache=@{}
$script:ChangedCodes=@{}
$script:SortedCodeKeys=@()
$script:ReadProgressContext=$null
$script:ProgressWriteWarning=$false
$script:MedianFallbackSec=5.44
$script:MedianFallbackSource='MEDIAN_FALLBACK_5.44'

function CsvField {
    param([string]$Value)
    if($null -eq $Value){return '""'}
    return '"'+$Value.Replace('"','""')+'"'
}
function Normalize-Code {
    param([string]$Value)
    if([string]::IsNullOrWhiteSpace($Value)){return ''}
    $v=$Value.Trim()
    try{$v=[System.IO.Path]::GetFileName($v)}catch{}
    if(-not $v.EndsWith('.pgmx',[System.StringComparison]::OrdinalIgnoreCase)){$v += '.pgmx'}
    return $v.ToLowerInvariant()
}
function Parse-Num {
    param($Value)
    if($null -eq $Value){return $null}
    $s=([string]$Value).Trim().Replace(',','.')
    if([string]::IsNullOrWhiteSpace($s)){return $null}
    $n=0.0
    if([double]::TryParse($s,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$n) -and -not [double]::IsNaN($n) -and -not [double]::IsInfinity($n)){return $n}
    return $null
}
function Parse-DateSafe {
    param([string]$Value)
    if([string]::IsNullOrWhiteSpace($Value)){return [datetime]::MinValue}
    $d=[datetime]::MinValue
    if([datetime]::TryParse($Value,[ref]$d)){return $d}
    return [datetime]::MinValue
}
function Test-IsMedianFallbackRow {
    param($Row)
    if($null -eq $Row){return $false}
    $status=[string]$Row.Status
    $src=[string]$Row.ActiveSource
    return ($status -eq 'MEDIAN_FALLBACK' -or $src -eq $script:MedianFallbackSource)
}
function Move-PreparedFileAtomic {
    param([string]$TemporaryPath,[string]$DestinationPath,[string]$BackupPath='')
    # Replace an existing file atomically. Move-Item -Force is not an atomic
    # replacement and can fail with "file already exists" on Windows.
    for($attempt=0;$attempt -lt 8;$attempt++){
        try{
            if([System.IO.File]::Exists($DestinationPath)){
                if($BackupPath){
                    if([System.IO.File]::Exists($BackupPath)){throw 'Вече съществува резервен файл за незавършен запис.'}
                    [System.IO.File]::Replace($TemporaryPath,$DestinationPath,$BackupPath,$true)
                }else{[System.IO.File]::Replace($TemporaryPath,$DestinationPath,[System.Management.Automation.Language.NullString]::Value,$true)}
            }else{[System.IO.File]::Move($TemporaryPath,$DestinationPath)}
            return
        }catch{
            $reason=$_.Exception.Message
            if($attempt -ge 7 -or -not [System.IO.File]::Exists($TemporaryPath) -or ($BackupPath -and [System.IO.File]::Exists($BackupPath))){
                throw ("Неуспешен запис на '$DestinationPath': $reason")
            }
            Start-Sleep -Milliseconds (75+($attempt*50))
        }
    }
}
function Write-JsonAtomic {
    param([string]$Path,$Object)
    $tmp=$Path+'.tmp.'+[guid]::NewGuid().ToString('N')
    try{
        [System.IO.File]::WriteAllText($tmp,($Object|ConvertTo-Json -Depth 6 -Compress),(New-Object System.Text.UTF8Encoding($false)))
        Move-PreparedFileAtomic -TemporaryPath $tmp -DestinationPath $Path
    }finally{if([System.IO.File]::Exists($tmp)){try{[System.IO.File]::Delete($tmp)}catch{}}}
}
function Write-Progress {
    param([string]$Status,[string]$Message,[int]$CurrentIndex,[int]$TotalFiles,[string]$CurrentFile,$Stats)
    $o=[ordered]@{Status=$Status;Message=$Message;CurrentIndex=$CurrentIndex;TotalFiles=$TotalFiles;CurrentFile=$CurrentFile;FilesDone=$Stats.FilesDone;FilesSkipped=$Stats.FilesSkipped;Rows=$Stats.Rows;Positive=$Stats.Positive;Zero=$Stats.Zero;MedianFallback=$Stats.MedianFallback;NewCodes=$Stats.NewCodes;ActiveUpdated=$Stats.ActiveUpdated;ZeroIgnored=$Stats.ZeroIgnored;Errors=$Stats.Errors;ErrorLogPath=$ErrorCsv;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')}
    try{
        Write-JsonAtomic -Path $ProgressPath -Object ([pscustomobject]$o)
        $script:ProgressWriteWarning=$false
    }catch{
        # Progress is advisory. A brief reader/file lock must not abort a
        # committed import. The final result is still a required write.
        if(-not $script:ProgressWriteWarning){[Console]::Error.WriteLine('Progress: '+$_.Exception.Message)}
        $script:ProgressWriteWarning=$true
    }
}
function Report-ReadProgress {
    param([string]$Message)
    $ctx=$script:ReadProgressContext
    if($null -ne $ctx){Write-Progress -Status 'Running' -Message $Message -CurrentIndex $ctx.Index -TotalFiles $ctx.Total -CurrentFile $ctx.File -Stats $ctx.Stats}
}
function Test-StopRequested {
    return [System.IO.File]::Exists($StopPath)
}
function Ensure-Headers {
    if(-not [System.IO.File]::Exists($HistoryCsv)){
        [System.IO.File]::WriteAllText($HistoryCsv,'ProgramCodeF;MeasuredTimeSec;MeasuredAt;Source;ImportedAt'+[Environment]::NewLine,(New-Object System.Text.UTF8Encoding($true)))
    }
    if(-not [System.IO.File]::Exists($RegistryCsv)){
        [System.IO.File]::WriteAllText($RegistryCsv,'Identity;Path;ModifiedAt;ImportedAt;Rows;Positive;Zero'+[Environment]::NewLine,(New-Object System.Text.UTF8Encoding($true)))
    }
    if(-not [System.IO.File]::Exists($ErrorCsv)){
        [System.IO.File]::WriteAllText($ErrorCsv,'When;File;Error'+[Environment]::NewLine,(New-Object System.Text.UTF8Encoding($true)))
    }
}
function Load-RegistrySet {
    $set=@{}
    if([System.IO.File]::Exists($RegistryCsv)){
        foreach($r in @(Import-Csv -LiteralPath $RegistryCsv -Delimiter ';' -Encoding UTF8)){
            if($r.Identity){$set[[string]$r.Identity]=$true}
        }
    }
    return $set
}
function Get-Identity {
    param([System.IO.FileInfo]$Fi)
    return ($Fi.FullName.ToLowerInvariant()+'|'+$Fi.Length+'|'+$Fi.LastWriteTimeUtc.Ticks)
}
function Load-DbMap {
    $map=@{}
    if([System.IO.File]::Exists($DbCsv)){
        foreach($r in @(Import-Csv -LiteralPath $DbCsv -Delimiter ';' -Encoding UTF8)){
            $k=Normalize-Code $r.ProgramCodeF
            if($k){$map[$k]=$r}
        }
        return $map
    }
    if([System.IO.File]::Exists($BaseCsv)){
        foreach($r in @(Import-Csv -LiteralPath $BaseCsv -Encoding UTF8)){
            $k=Normalize-Code $r.ProgramCodeF
            if(-not $k){continue}
            $t=Parse-Num $r.Canonical_Time_sec
            $active=''
            if($null -ne $t -and $t -gt 0){$active=[string]$t}
            $status='NO_VALID_TIME'
            if($active){$status='BASE_VALID'}
            $map[$k]=[pscustomobject]@{ProgramCodeF=[System.IO.Path]::GetFileName([string]$r.ProgramCodeF);PreviousValidTimeSec='';PreviousValidMeasuredAt='';PreviousValidSource='';LastMeasuredTimeSec=$active;LastMeasuredAt='2000-01-01 00:00:00';LastSource='Initial 71-file database';ActiveTimeSec=$active;ActiveMeasuredAt='2000-01-01 00:00:00';ActiveSource='Initial 71-file database';Status=$status;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')}
        }
    }
    return $map
}
function Get-ExistingFileLength {
    param([string]$Path)
    if([System.IO.File]::Exists($Path)){return [long](Get-Item -LiteralPath $Path).Length}
    return [long]-1
}
function Restore-AppendLength {
    param([string]$Path,[long]$Length)
    if($Length -lt 0){if([System.IO.File]::Exists($Path)){[System.IO.File]::Delete($Path)};return}
    if(-not [System.IO.File]::Exists($Path)){throw ("Липсва файл за възстановяване: $Path")}
    $stream=[System.IO.File]::Open($Path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Write,[System.IO.FileShare]::Read)
    try{
        if($stream.Length -lt $Length){throw ("Файлът е по-къс от последната сигурна точка: $Path")}
        $stream.SetLength($Length)
        $stream.Flush($true)
    }finally{$stream.Dispose()}
}
function Remove-CommitScratchFiles {
    param([string]$Token)
    foreach($path in @(($DbCsv+'.tmp.'+$Token),($DbCsv+'.bak.'+$Token),($ActiveCsv+'.tmp.'+$Token),($ActiveCsv+'.bak.'+$Token))){
        if([System.IO.File]::Exists($path)){try{[System.IO.File]::Delete($path)}catch{}}
    }
}
function Restore-PendingCommit {
    if(-not [System.IO.File]::Exists($CommitJournalPath)){return}
    $j=([System.IO.File]::ReadAllText($CommitJournalPath,[System.Text.Encoding]::UTF8)|ConvertFrom-Json)
    if($j.Version -ne 1 -or [string]$j.Token -notmatch '^[0-9a-f]{32}$'){throw 'Невалиден журнал за незавършен импорт. Записването е спряно.'}
    $token=[string]$j.Token
    $items=@(
        [pscustomobject]@{Path=$DbCsv;Existed=[bool]$j.DbExisted},
        [pscustomobject]@{Path=$ActiveCsv;Existed=[bool]$j.ActiveExisted}
    )
    foreach($item in $items){
        $tmp=$item.Path+'.tmp.'+$token
        $bak=$item.Path+'.bak.'+$token
        if([System.IO.File]::Exists($bak)){
            Move-PreparedFileAtomic -TemporaryPath $bak -DestinationPath $item.Path
        }elseif(-not $item.Existed -and -not [System.IO.File]::Exists($tmp)){
            # A new snapshot was moved into place, but the file was never
            # committed to the registry. Remove only that new snapshot.
            if([System.IO.File]::Exists($item.Path)){[System.IO.File]::Delete($item.Path)}
        }
    }
    Restore-AppendLength -Path $HistoryCsv -Length ([long]$j.HistoryLength)
    Restore-AppendLength -Path $RegistryCsv -Length ([long]$j.RegistryLength)
    [System.IO.File]::Delete($CommitJournalPath)
    Remove-CommitScratchFiles -Token $token
}
function Save-DbAndActive {
    param($Map,[System.Collections.Generic.List[string]]$HistoryLines,[string]$RegistryLine)
    # Reuse serialized rows that the current file did not change. Previously
    # every file executed a PowerShell pipeline for every field of 17k+ rows.
    if($script:SortedCodeKeys.Count -ne $Map.Count){
        $script:SortedCodeKeys=[string[]]@($Map.Keys)
        [Array]::Sort($script:SortedCodeKeys,[System.StringComparer]::OrdinalIgnoreCase)
    }
    $dbLines=New-Object System.Collections.Generic.List[string]
    $activeLines=New-Object System.Collections.Generic.List[string]
    $dbLines.Add('ProgramCodeF;PreviousValidTimeSec;PreviousValidMeasuredAt;PreviousValidSource;LastMeasuredTimeSec;LastMeasuredAt;LastSource;ActiveTimeSec;ActiveMeasuredAt;ActiveSource;Status;UpdatedAt')
    $activeLines.Add('ProgramCodeF;ActiveTimeSec')
    foreach($k in $script:SortedCodeKeys){
        if(-not $script:DbRowsCache.ContainsKey($k) -or $script:ChangedCodes.ContainsKey($k)){
            $r=$Map[$k]
            $vals=@($r.ProgramCodeF,$r.PreviousValidTimeSec,$r.PreviousValidMeasuredAt,$r.PreviousValidSource,$r.LastMeasuredTimeSec,$r.LastMeasuredAt,$r.LastSource,$r.ActiveTimeSec,$r.ActiveMeasuredAt,$r.ActiveSource,$r.Status,$r.UpdatedAt)
            for($v=0;$v -lt $vals.Count;$v++){$vals[$v]='"'+([string]$vals[$v]).Replace('"','""')+'"'}
            $script:DbRowsCache[$k]=[string]::Join(';',[string[]]$vals)
            $t=Parse-Num $r.ActiveTimeSec
            if($null -ne $t -and $t -gt 0){
                $code=[System.IO.Path]::GetFileName([string]$r.ProgramCodeF)
                $script:ActiveRowsCache[$k]='"'+$code.Replace('"','""')+'";"'+$t.ToString('R',[System.Globalization.CultureInfo]::InvariantCulture)+'"'
            }else{[void]$script:ActiveRowsCache.Remove($k)}
        }
        $dbLines.Add([string]$script:DbRowsCache[$k])
        if($script:ActiveRowsCache.ContainsKey($k)){$activeLines.Add([string]$script:ActiveRowsCache[$k])}
    }
    $token=[guid]::NewGuid().ToString('N')
    $tmp=$DbCsv+'.tmp.'+$token
    $tmp2=$ActiveCsv+'.tmp.'+$token
    $enc=New-Object System.Text.UTF8Encoding($true)
    $journalWritten=$false
    try{
        [System.IO.File]::WriteAllLines($tmp,$dbLines.ToArray(),$enc)
        [System.IO.File]::WriteAllLines($tmp2,$activeLines.ToArray(),$enc)
        $journal=[pscustomobject]@{Version=1;Token=$token;DbExisted=[System.IO.File]::Exists($DbCsv);ActiveExisted=[System.IO.File]::Exists($ActiveCsv);HistoryLength=(Get-ExistingFileLength $HistoryCsv);RegistryLength=(Get-ExistingFileLength $RegistryCsv)}
        Write-JsonAtomic -Path $CommitJournalPath -Object $journal
        $journalWritten=$true
        Move-PreparedFileAtomic -TemporaryPath $tmp -DestinationPath $DbCsv -BackupPath ($DbCsv+'.bak.'+$token)
        Move-PreparedFileAtomic -TemporaryPath $tmp2 -DestinationPath $ActiveCsv -BackupPath ($ActiveCsv+'.bak.'+$token)
        if($HistoryLines.Count -gt 0){[System.IO.File]::AppendAllLines($HistoryCsv,$HistoryLines.ToArray(),(New-Object System.Text.UTF8Encoding($false)))}
        # Registry is last: only a fully saved file can be skipped next time.
        [System.IO.File]::AppendAllText($RegistryCsv,$RegistryLine+[Environment]::NewLine,(New-Object System.Text.UTF8Encoding($false)))
        [System.IO.File]::Delete($CommitJournalPath)
        $journalWritten=$false
        $script:ChangedCodes.Clear()
    }catch{
        $writeError=$_.Exception.Message
        if($journalWritten){
            try{Restore-PendingCommit}catch{throw ($writeError+' | Възстановяването не завърши: '+$_.Exception.Message+' | Продължи след освобождаване на заключените файлове.')}
        }
        throw $writeError
    }finally{
        # Never delete backups while a recovery journal still needs them.
        if(-not [System.IO.File]::Exists($CommitJournalPath)){Remove-CommitScratchFiles -Token $token}
    }
}
function Merge-Observation {
    param($Map,[string]$ProgramCodeF,[double]$TimeSec,[datetime]$MeasuredAt,[string]$Source,[System.Collections.Generic.List[string]]$HistoryLines)
    $k=Normalize-Code $ProgramCodeF
    if(-not $k){return [pscustomobject]@{NewCode=$false;ActiveUpdated=$false;ZeroIgnored=$false;MedianFallback=$false}}
    $now=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $isNew=$false
    if($Map.ContainsKey($k)){$r=$Map[$k]}else{
        $isNew=$true
        $r=[pscustomobject]@{ProgramCodeF=[System.IO.Path]::GetFileName($ProgramCodeF);PreviousValidTimeSec='';PreviousValidMeasuredAt='';PreviousValidSource='';LastMeasuredTimeSec='';LastMeasuredAt='';LastSource='';ActiveTimeSec='';ActiveMeasuredAt='';ActiveSource='';Status='NO_VALID_TIME';UpdatedAt=$now}
        $Map[$k]=$r
    }
    $lastDate=Parse-DateSafe ([string]$r.LastMeasuredAt)
    if($MeasuredAt -ge $lastDate){$r.LastMeasuredTimeSec=$TimeSec.ToString('R',[System.Globalization.CultureInfo]::InvariantCulture);$r.LastMeasuredAt=$MeasuredAt.ToString('yyyy-MM-dd HH:mm:ss');$r.LastSource=$Source}
    $active=Parse-Num $r.ActiveTimeSec
    $activeDate=Parse-DateSafe ([string]$r.ActiveMeasuredAt)
    $activeWasFallback=Test-IsMedianFallbackRow $r
    $activeUpdated=$false
    $zeroIgnored=$false
    $medianFallback=$false

    if($TimeSec -gt 0){
        # A real measured positive time ALWAYS replaces a median fallback,
        # even if the real value is lower than 5.44. For real-vs-real values,
        # keep the existing newest-measurement rule.
        $acceptReal=($activeWasFallback -or $MeasuredAt -ge $activeDate)
        if($acceptReal -and -not ($MeasuredAt -eq $activeDate -and $Source -eq [string]$r.ActiveSource -and $null -ne $active -and $TimeSec -eq $active -and -not $activeWasFallback)){
            if($null -ne $active -and $active -gt 0 -and -not $activeWasFallback){
                $r.PreviousValidTimeSec=$active.ToString('R',[System.Globalization.CultureInfo]::InvariantCulture)
                $r.PreviousValidMeasuredAt=$r.ActiveMeasuredAt
                $r.PreviousValidSource=$r.ActiveSource
            }
            $r.ActiveTimeSec=$TimeSec.ToString('R',[System.Globalization.CultureInfo]::InvariantCulture)
            $r.ActiveMeasuredAt=$MeasuredAt.ToString('yyyy-MM-dd HH:mm:ss')
            $r.ActiveSource=$Source
            $r.Status='VALID'
            $activeUpdated=$true
        }
    }else{
        if($null -ne $active -and $active -gt 0 -and -not $activeWasFallback){
            # Existing real time wins over a new zero.
            if($MeasuredAt -ge $lastDate){$r.Status='ZERO_IGNORED'}
            $zeroIgnored=$true
        }else{
            # No prior real time: zero is never kept as the active production
            # time. Use the agreed temporary median 5.44 sec until a real
            # positive measurement arrives.
            $r.ActiveTimeSec=$script:MedianFallbackSec.ToString('R',[System.Globalization.CultureInfo]::InvariantCulture)
            $r.ActiveMeasuredAt=$MeasuredAt.ToString('yyyy-MM-dd HH:mm:ss')
            $r.ActiveSource=$script:MedianFallbackSource
            $r.Status='MEDIAN_FALLBACK'
            $medianFallback=$true
            if(-not $activeWasFallback){$activeUpdated=$true}
        }
    }

    # A newer zero may be the latest observation while an older REAL positive
    # remains active. Do not relabel a median fallback as ZERO_IGNORED.
    $latest=Parse-Num $r.LastMeasuredTimeSec
    $activeAfter=Parse-Num $r.ActiveTimeSec
    if(-not (Test-IsMedianFallbackRow $r) -and $null -ne $activeAfter -and $activeAfter -gt 0 -and $null -ne $latest -and $latest -eq 0){$r.Status='ZERO_IGNORED'}
    $r.UpdatedAt=$now
    $script:ChangedCodes[$k]=$true
    $HistoryLines.Add((@($r.ProgramCodeF,$TimeSec.ToString('R',[System.Globalization.CultureInfo]::InvariantCulture),$MeasuredAt.ToString('yyyy-MM-dd HH:mm:ss'),$Source,$now)|ForEach-Object{CsvField ([string]$_)}) -join ';')
    return [pscustomobject]@{NewCode=$isNew;ActiveUpdated=$activeUpdated;ZeroIgnored=$zeroIgnored;MedianFallback=$medianFallback}
}
function Read-ObservationsFromCsv {
    param([string]$Path)
    $out=New-Object System.Collections.Generic.List[object]
    $first=[System.IO.File]::ReadLines($Path,[System.Text.Encoding]::UTF8)|Select-Object -First 1
    $delimiter=','
    if($first -match ';'){$delimiter=';'}elseif($first -match "`t"){$delimiter="`t"}
    $foundColumns=$false
    $rowIndex=0
    foreach($r in @(Import-Csv -LiteralPath $Path -Delimiter $delimiter -Encoding UTF8)){
        $rowIndex++
        if(($rowIndex % 500) -eq 0 -and (Test-StopRequested)){throw '__STOP__'}
        $pc='';$tv=$null;$hasProgram=$false;$hasTime=$false
        foreach($n in @('Project','ProgramCodeF','Program','PGMX')){if($r.PSObject.Properties.Name -contains $n){$hasProgram=$true;$pc=[string]$r.$n;if($pc){break}}}
        foreach($n in @('Time (sec)','Time_sec','TimeSec','Canonical_Time_sec','ActiveTimeSec','MeasuredTimeSec')){if($r.PSObject.Properties.Name -contains $n){$hasTime=$true;$tv=Parse-Num $r.$n;if($null -ne $tv){break}}}
        if($hasProgram -and $hasTime){$foundColumns=$true}
        if($pc -and $null -ne $tv){
            if($tv -lt 0){throw ("Отрицателно време на ред $rowIndex в CSV файла.")}
            $out.Add([pscustomobject]@{ProgramCodeF=$pc;TimeSec=[double]$tv})
        }
    }
    if(-not $foundColumns){throw 'Не намирам колони за програма и време в CSV файла.'}
    return @($out.ToArray())
}
function Release-ExcelObject {
    param($Object)
    if($null -ne $Object){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)}catch{}}
}
function Read-ExcelRangeValues {
    param($Worksheet,[int]$FirstRow,[int]$FirstColumn,[int]$LastRow,[int]$LastColumn)
    $cells=$null;$a=$null;$b=$null;$range=$null
    try{
        $cells=$Worksheet.Cells
        $a=$cells.Item($FirstRow,$FirstColumn)
        $b=$cells.Item($LastRow,$LastColumn)
        $range=$Worksheet.Range($a,$b)
        $values=$range.Value2
        # Preserve the 2-D SAFEARRAY; do not let PowerShell enumerate it.
        return ,$values
    }finally{Release-ExcelObject $range;Release-ExcelObject $b;Release-ExcelObject $a;Release-ExcelObject $cells}
}
function Read-ObservationsFromExcel {
    param([string]$Path,$Excel)
    $out=New-Object System.Collections.Generic.List[object]
    $wb=$null;$workbooks=$null;$sheets=$null;$ws=$null;$used=$null;$cells=$null;$last=$null
    try{
        $workbooks=$Excel.Workbooks
        $wb=$workbooks.Open($Path,0,$true)
        $sheets=$wb.Worksheets
        $ws=$sheets.Item(1)
        $used=$ws.UsedRange
        $firstR=[int]$used.Row
        $firstC=[int]$used.Column
        $cells=$ws.Cells
        # Ignore formatting-only rows/columns in a bloated UsedRange.
        # All Find options are explicit so a user's saved Find settings
        # cannot change the import's bounds.
        $last=$cells.Find('*',[Type]::Missing,-4123,2,1,2,$false,$false,$false)
        if($null -eq $last){throw 'Excel листът е празен.'}
        $lastR=[int]$last.Row
        Release-ExcelObject $last;$last=$null
        $last=$cells.Find('*',[Type]::Missing,-4123,2,2,2,$false,$false,$false)
        $lastC=[int]$last.Column
        Release-ExcelObject $last;$last=$null
        $headerLastR=[Math]::Min($firstR+24,$lastR)
        $headerLastC=[Math]::Min($firstC+255,$lastC)
        $values=Read-ExcelRangeValues -Worksheet $ws -FirstRow $firstR -FirstColumn $firstC -LastRow $headerLastR -LastColumn $headerLastC
        $isMatrix=($values -is [Array] -and $values.Rank -eq 2)
        $baseR=0;$baseC=0
        if($isMatrix){$baseR=$values.GetLowerBound(0);$baseC=$values.GetLowerBound(1)}
        $hdr=0;$pcCol=0;$timeCol=0
        for($r=$firstR;$r -le $headerLastR;$r++){
            $pcCol=0;$timeCol=0
            for($c=$firstC;$c -le $headerLastC;$c++){
                if($isMatrix){$v=([string]$values.GetValue($baseR+$r-$firstR,$baseC+$c-$firstC)).Trim()}else{$v=([string]$values).Trim()}
                if($v -match '^(Project|ProgramCodeF)$'){$pcCol=$c}
                if($v -eq 'Time (sec)'){$timeCol=$c}
            }
            if($pcCol -gt 0 -and $timeCol -gt 0){$hdr=$r;break}
        }
        if($hdr -eq 0){throw 'Не намирам заглавия Project/ProgramCodeF и Time (sec).'}
        # Two column reads per block, no per-cell COM calls and no full
        # UsedRange array containing unused columns.
        for($start=$hdr+1;$start -le $lastR;$start+=5000){
            if(Test-StopRequested){throw '__STOP__'}
            $end=[Math]::Min($start+4999,$lastR)
            Report-ReadProgress ("Чета редове $start - $end от $lastR...")
            $codes=Read-ExcelRangeValues -Worksheet $ws -FirstRow $start -FirstColumn $pcCol -LastRow $end -LastColumn $pcCol
            $times=Read-ExcelRangeValues -Worksheet $ws -FirstRow $start -FirstColumn $timeCol -LastRow $end -LastColumn $timeCol
            $codeMatrix=($codes -is [Array] -and $codes.Rank -eq 2)
            $timeMatrix=($times -is [Array] -and $times.Rank -eq 2)
            $cr=0;$cc=0;$tr=0;$tc=0
            if($codeMatrix){$cr=$codes.GetLowerBound(0);$cc=$codes.GetLowerBound(1)}
            if($timeMatrix){$tr=$times.GetLowerBound(0);$tc=$times.GetLowerBound(1)}
            for($offset=0;$offset -le ($end-$start);$offset++){
                if(($offset % 500) -eq 0 -and (Test-StopRequested)){throw '__STOP__'}
                if($codeMatrix){$pc=([string]$codes.GetValue($cr+$offset,$cc)).Trim()}else{$pc=([string]$codes).Trim()}
                if(-not $pc){continue}
                if($timeMatrix){$raw=$times.GetValue($tr+$offset,$tc)}else{$raw=$times}
                $tv=Parse-Num $raw
                if($null -eq $tv){continue}
                if($tv -lt 0){throw ('Отрицателно време на ред '+($start+$offset)+'.')}
                $out.Add([pscustomobject]@{ProgramCodeF=$pc;TimeSec=[double]$tv})
            }
        }
        return @($out.ToArray())
    }finally{
        try{if($wb){$wb.Close($false)}}catch{}
        Release-ExcelObject $last;Release-ExcelObject $cells;Release-ExcelObject $used
        Release-ExcelObject $ws;Release-ExcelObject $sheets;Release-ExcelObject $wb;Release-ExcelObject $workbooks
    }
}
function New-ImportExcelSession {
    $app=$null
    try{
        $app=New-Object -ComObject Excel.Application
        $app.Visible=$false;$app.DisplayAlerts=$false;$app.ScreenUpdating=$false;$app.EnableEvents=$false
        # Never execute macros or update links in imported measurement files.
        try{$app.AutomationSecurity=3}catch{}
        try{$app.AskToUpdateLinks=$false}catch{}
        return $app
    }catch{
        try{if($app){$app.Quit()}}catch{}
        Release-ExcelObject $app
        throw
    }
}
function Write-ImportFileError {
    param([string]$Path,[string]$Phase,[string]$Message)
    $errLine=@((Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Path,($Phase+': '+$Message))|ForEach-Object{CsvField ([string]$_)}
    [System.IO.File]::AppendAllText($ErrorCsv,(($errLine -join ';')+[Environment]::NewLine),(New-Object System.Text.UTF8Encoding($false)))
}

$stats=@{FilesDone=0;FilesSkipped=0;Rows=0;Positive=0;Zero=0;MedianFallback=0;NewCodes=0;ActiveUpdated=0;ZeroIgnored=0;Errors=0}
$excel=$null;$writeLock=$null
$path='';$phase='Startup';$currentIndex=0;$total=0
$failedFiles=New-Object System.Collections.Generic.List[string]
$fallbackCodes=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
try{
    try{$writeLock=[System.IO.File]::Open($WriteLockPath,[System.IO.FileMode]::OpenOrCreate,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)}catch{throw 'Базата с времена се използва от друг импорт/запис. Изчакай той да приключи.'}
    Restore-PendingCommit
    Ensure-Headers
    $files=@([System.IO.File]::ReadAllLines($ManifestPath,[System.Text.Encoding]::UTF8)|Where-Object{-not [string]::IsNullOrWhiteSpace($_)})
    $total=$files.Count
    Write-Progress -Status 'Running' -Message 'Зареждане на локалната база...' -CurrentIndex 0 -TotalFiles $total -CurrentFile '' -Stats $stats
    $registry=Load-RegistrySet
    $map=Load-DbMap
    for($i=0;$i -lt $files.Count;$i++){
        $currentIndex=$i+1
        if(Test-StopRequested){
            Write-Progress -Status 'Stopped' -Message 'Импортът е спрян безопасно.' -CurrentIndex $i -TotalFiles $total -CurrentFile '' -Stats $stats
            Write-JsonAtomic -Path $ResultPath -Object ([pscustomobject]@{Status='Stopped';Stats=$stats;FailedFiles=$failedFiles.ToArray();FallbackCodes=@($fallbackCodes);MedianFallbackSec=$script:MedianFallbackSec;ErrorLogPath=$ErrorCsv;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')})
            exit 0
        }
        $path=[string]$files[$i]
        $displayName=[System.IO.Path]::GetFileName($path)
        $fi=$null;$phase='Reading'
        try{
            # Missing, inaccessible, or unsupported inputs are per-file
            # errors. They must not terminate the remaining manifest.
            $fi=Get-Item -LiteralPath $path -ErrorAction Stop
            if($fi.PSIsContainer){throw 'Избраният път е папка, а не файл.'}
            $identity=Get-Identity $fi
            if(-not $ForceImport -and $registry.ContainsKey($identity)){
                $stats.FilesSkipped++
                Write-Progress -Status 'Running' -Message 'Файлът вече е импортнат - пропускам.' -CurrentIndex $currentIndex -TotalFiles $total -CurrentFile $fi.Name -Stats $stats
                continue
            }
            $script:ReadProgressContext=@{Index=$currentIndex;Total=$total;File=$fi.Name;Stats=$stats}
            Write-Progress -Status 'Running' -Message 'Чета файл с времена...' -CurrentIndex $currentIndex -TotalFiles $total -CurrentFile $fi.Name -Stats $stats
            $ext=$fi.Extension.ToLowerInvariant()
            if($ext -eq '.csv'){$obs=@(Read-ObservationsFromCsv -Path $fi.FullName)}elseif($ext -in @('.xls','.xlsx','.xlsm','.xlsb')){
                if($null -eq $excel){$excel=New-ImportExcelSession}
                $obs=@(Read-ObservationsFromExcel -Path $fi.FullName -Excel $excel)
            }else{throw ('Неподдържан формат: '+$ext)}
            if(Test-StopRequested){throw '__STOP__'}
            $phase='Merging'
            $historyLines=New-Object System.Collections.Generic.List[string]
            $fileRows=0;$filePos=0;$fileZero=0;$fileFallback=0;$fileNew=0;$fileUpdated=0;$fileIgnored=0
            Report-ReadProgress ('Обединявам '+$obs.Count+' записа...')
            foreach($o in $obs){
                if(($fileRows % 500) -eq 0 -and (Test-StopRequested)){throw '__STOP__'}
                $fileRows++
                if([double]$o.TimeSec -gt 0){$filePos++}else{$fileZero++}
                $m=Merge-Observation -Map $map -ProgramCodeF ([string]$o.ProgramCodeF) -TimeSec ([double]$o.TimeSec) -MeasuredAt $fi.LastWriteTime -Source $fi.Name -HistoryLines $historyLines
                if($m.NewCode){$fileNew++}
                if($m.ActiveUpdated){$fileUpdated++}
                if($m.ZeroIgnored){$fileIgnored++}
                if($m.MedianFallback){$fileFallback++;[void]$fallbackCodes.Add((Normalize-Code ([string]$o.ProgramCodeF)))}elseif([double]$o.TimeSec -gt 0){[void]$fallbackCodes.Remove((Normalize-Code ([string]$o.ProgramCodeF)))}
            }
            if(Test-StopRequested){throw '__STOP__'}
            $phase='Saving'
            Report-ReadProgress ('Записвам '+$fileRows+' реда и обновената база...')
            $line=@($identity,$fi.FullName,$fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'),(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$fileRows,$filePos,$fileZero)|ForEach-Object{CsvField ([string]$_)}
            Save-DbAndActive -Map $map -HistoryLines $historyLines -RegistryLine ($line -join ';')
            $registry[$identity]=$true
            $stats.FilesDone++;$stats.Rows+=$fileRows;$stats.Positive+=$filePos;$stats.Zero+=$fileZero
            $stats.NewCodes+=$fileNew;$stats.ActiveUpdated+=$fileUpdated;$stats.ZeroIgnored+=$fileIgnored;$stats.MedianFallback+=$fileFallback
            $phase='Committed'
            Write-Progress -Status 'Running' -Message 'Файлът е обработен и записан.' -CurrentIndex $currentIndex -TotalFiles $total -CurrentFile $fi.Name -Stats $stats
        }catch{
            if($_.Exception.Message -eq '__STOP__'){
                Write-Progress -Status 'Stopped' -Message 'Импортът е спрян безопасно. Текущият файл ще се повтори при продължаване.' -CurrentIndex $currentIndex -TotalFiles $total -CurrentFile $displayName -Stats $stats
                Write-JsonAtomic -Path $ResultPath -Object ([pscustomobject]@{Status='Stopped';Stats=$stats;FailedFiles=$failedFiles.ToArray();FallbackCodes=@($fallbackCodes);MedianFallbackSec=$script:MedianFallbackSec;ErrorLogPath=$ErrorCsv;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')})
                exit 0
            }
            $message=$_.Exception.Message
            $stats.Errors++
            $failedFiles.Add($path)
            try{Write-ImportFileError -Path $path -Phase $phase -Message $message}catch{[Console]::Error.WriteLine('Error log: '+$_.Exception.Message)}
            if($phase -ne 'Reading'){
                # A database/commit failure is not a bad input. Stop before
                # another file can inherit an uncommitted in-memory map.
                throw ('Импортът е спрян при запис/обединяване на '+$displayName+': '+$message)
            }
            Write-Progress -Status 'Running' -Message ('Пропускам файл с грешка: '+$message) -CurrentIndex $currentIndex -TotalFiles $total -CurrentFile $displayName -Stats $stats
        }
    }
    $completionMessage='Всички избрани файлове са проверени.'
    if($stats.Errors -gt 0){$completionMessage=('Проверката приключи с '+$stats.Errors+' файла с грешка. Те не са маркирани като импортирани.')}
    Write-Progress -Status 'Completed' -Message $completionMessage -CurrentIndex $total -TotalFiles $total -CurrentFile '' -Stats $stats
    Write-JsonAtomic -Path $ResultPath -Object ([pscustomobject]@{Status='Completed';Message=$completionMessage;Stats=$stats;FailedFiles=$failedFiles.ToArray();FallbackCodes=@($fallbackCodes);MedianFallbackSec=$script:MedianFallbackSec;ErrorLogPath=$ErrorCsv;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')})
}catch{
    $fatalMessage=$_.Exception.Message
    try{Write-Progress -Status 'Error' -Message $fatalMessage -CurrentIndex $currentIndex -TotalFiles $total -CurrentFile ([System.IO.Path]::GetFileName($path)) -Stats $stats}catch{}
    try{Write-JsonAtomic -Path $ResultPath -Object ([pscustomobject]@{Status='Error';Error=$fatalMessage;Phase=$phase;CurrentFile=$path;Stats=$stats;FailedFiles=$failedFiles.ToArray();FallbackCodes=@($fallbackCodes);MedianFallbackSec=$script:MedianFallbackSec;ErrorLogPath=$ErrorCsv;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')})}catch{}
    exit 1
}finally{
    try{if($excel){$excel.Quit()}}catch{}
    Release-ExcelObject $excel
    try{if($writeLock){$writeLock.Dispose()}}catch{}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
