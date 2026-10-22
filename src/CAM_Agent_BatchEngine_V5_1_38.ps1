# CAM Agent V5.1.38 - independent queued PGMX -> common MIXX -> common PowerSet.
# Deliberately no modification of original product/server PGMX or old converter.
# Daily-manifest + persistent exception ledger, compatible with Windows PowerShell 5.1.
$script:AutoExcludedPgmxFile = Join-Path $script:AppRoot 'PWX500_Agent_AutoExcludedPgmx.json'
$script:AutoExcludedPgmxMap = $null
$script:BatchOutputFolderOverride = ''
$script:BatchActiveManifestPath = ''

function Write-CamJsonAtomic {
    param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)]$Data)
    $parent=[System.IO.Path]::GetDirectoryName($Path)
    if(-not [System.IO.Directory]::Exists($parent)){[void][System.IO.Directory]::CreateDirectory($parent)}
    $temp=Join-Path $parent ([System.IO.Path]::GetFileName($Path)+'.'+[guid]::NewGuid().ToString('N')+'.tmp')
    $backup=$Path+'.bak'
    try{
        $json=ConvertTo-Json -InputObject $Data -Depth 14
        [System.IO.File]::WriteAllText($temp,$json,(New-Object System.Text.UTF8Encoding($true)))
        if([System.IO.File]::Exists($Path)){
            [System.IO.File]::Replace($temp,$Path,$backup,$true)
        }else{
            [System.IO.File]::Move($temp,$Path)
        }
    }finally{try{if([System.IO.File]::Exists($temp)){[System.IO.File]::Delete($temp)}}catch{}}
}

function Read-CamJson {
    param([string]$Path)
    if(-not [System.IO.File]::Exists($Path)){return $null}
    $txt=[System.IO.File]::ReadAllText($Path,[System.Text.Encoding]::UTF8)
    if([string]::IsNullOrWhiteSpace($txt)){throw ('Празен регистър: '+$Path)}
    try{return (ConvertFrom-Json -InputObject $txt -ErrorAction Stop)}catch{throw ('Повреден регистър: '+$Path+' | '+$_.Exception.Message)}
}

function Get-PgmxAutoExclusions {
    if($null -ne $script:AutoExcludedPgmxMap){return $script:AutoExcludedPgmxMap}
    $data=Read-CamJson -Path $script:AutoExcludedPgmxFile
    $map=@{}
    if($null -ne $data){
        foreach($entry in @($data.Details)){
            if($null -eq $entry){continue}
            $code=Normalize-TimeProgramCode ([string]$entry.ProgramCodeF)
            if(-not $code){continue}
            if(([string]$entry.Reason) -notin @('PROBLEM_MIXX','NO_REAL_TIME')){continue}
            $map[$code]=$entry
        }
    }
    $script:AutoExcludedPgmxMap=$map
    return $map
}

function Test-PgmxAutoExcluded {
    param([string]$ProgramCodeF)
    $key=Normalize-TimeProgramCode $ProgramCodeF
    if(-not $key){return $false}
    $map=Get-PgmxAutoExclusions
    return [bool]$map.ContainsKey($key)
}

function Register-PgmxAutoExclusion {
    param([string]$ProgramCodeF,[ValidateSet('PROBLEM_MIXX','NO_REAL_TIME')][string]$Reason,[string]$Origin='')
    $key=Normalize-TimeProgramCode $ProgramCodeF
    if(-not $key){throw 'Не мога да запиша изключване без валиден ProgramCodeF.'}
    $map=Get-PgmxAutoExclusions
    if($map.ContainsKey($key)){return}
    $map[$key]=[pscustomobject]@{ProgramCodeF=$key;Reason=$Reason;Origin=$Origin;RecordedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')}
    try{
        $obj=[ordered]@{Version='5.1.38';Details=@($map.Values|Sort-Object ProgramCodeF)}
        Write-CamJsonAtomic -Path $script:AutoExcludedPgmxFile -Data $obj
    }catch{$script:AutoExcludedPgmxMap=$null;throw}
}

function Get-DayManifestPath {
    param([string]$DailyFolder)
    return (Join-Path $DailyFolder 'PWX500_DayPGMX_Manifest.json')
}

function Save-DayPgmxManifest {
    param($Manifest)
    if($null -eq $Manifest -or -not $Manifest.DailyFolder){throw 'Невалиден дневен PGMX регистър.'}
    $Manifest.UpdatedAt=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-CamJsonAtomic -Path (Get-DayManifestPath -DailyFolder $Manifest.DailyFolder) -Data $Manifest
}

function Prepare-Stage1PgmxDailyFiles {
    # Executed ONLY after the confirmed PGMX worker reports a completed Execute.
    # Physical PGMX copies in each daily folder are the immutable source for delayed Stage 2.
    if($script:PgmxLastMode -ne 'Execute'){throw 'PGMX още не са генерирани в Execute режим.'}
    if([System.IO.File]::Exists((Join-Path $script:AppRoot 'CAM_Agent_TimesImport_Commit.json'))){throw 'Има незавършен импорт на времена; продължи го преди подготовка на PGMX.'}
    $daily=Get-DailyMixFolder
    if(-not [System.IO.Directory]::Exists($daily)){[void][System.IO.Directory]::CreateDirectory($daily)}
    $date=Get-Date -Format 'yyyy-MM-dd'
    $path=Get-DayManifestPath -DailyFolder $daily
    $old=Read-CamJson -Path $path
    if($null -ne $old -and [string]$old.Day -ne $date){
        throw ('Дневната папка вече принадлежи на друга дата: '+$daily+'. Не смесвам години без изрично решение.')
    }
    $byCode=@{}
    if($null -ne $old){foreach($item in @($old.Details)){if($null -ne $item){$byCode[(Normalize-TimeProgramCode $item.ProgramCodeF)]=$item}}}
    $candidates=@(Get-TimeMeasurementCandidates)
    $copied=0
    foreach($item in $candidates){
        Test-OperatorStopRequested
        $code=Normalize-TimeProgramCode $item.ProgramCodeF
        if(-not $code -or (Test-PgmxAutoExcluded $code)){continue}
        $name=[System.IO.Path]::GetFileName([string]$item.Name)
        if($name -ne [string]$item.Name -or [System.IO.Path]::GetExtension($name) -ine '.pgmx'){throw ('Невалидно име на PGMX: '+$name)}
        $source=[System.IO.Path]::GetFullPath([string]$item.Source)
        if(-not [System.IO.File]::Exists($source)){throw ('PGMX източникът липсва: '+$source)}
        $target=Join-Path $daily $name
        if($byCode.ContainsKey($code) -and [string]$byCode[$code].Status -in @('COMPLETED','PROBLEM_MIXX','NO_REAL_TIME')){continue}
        # Stage 1 may revisit its same-day source, but never replaces another code.
        [System.IO.File]::Copy($source,$target,$true)
        if(-not [System.IO.File]::Exists($target)){throw ('Неуспешно дневно копие: '+$target)}
        $byCode[$code]=[pscustomobject]@{ProgramCodeF=$code;DailyPath=$target;SourcePath=$source;Status='PENDING';UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')}
        $copied++
    }
    $manifest=[pscustomobject]@{Version='5.1.38';Day=$date;DailyFolder=$daily;UpdatedAt='';Details=@($byCode.Values | Sort-Object ProgramCodeF)}
    Save-DayPgmxManifest -Manifest $manifest
    return [pscustomobject]@{DailyFolder=$daily;Copied=$copied;Pending=@($manifest.Details|Where-Object {$_.Status -eq 'PENDING'}).Count}
}

function Get-PendingDayPgmxEntries {
    $root=Get-DailyRootFolder
    $entries=New-Object System.Collections.Generic.List[object]
    foreach($folder in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name)){
        if($folder.Name -notmatch '^\d{2}\.\d{2}$'){continue}
        $manifestPath=Get-DayManifestPath -DailyFolder $folder.FullName
        if(-not [System.IO.File]::Exists($manifestPath)){continue}
        $m=Read-CamJson -Path $manifestPath
        if([string]$m.DailyFolder -ne $folder.FullName){throw ('Несъответстваща дневна папка в регистъра: '+$manifestPath)}
        if([string]$m.Day -notmatch '^\d{4}-\d{2}-\d{2}$'){throw ('Липсва дата в: '+$manifestPath)}
        foreach($e in @($m.Details)){
            if($null -eq $e -or [string]$e.Status -ne 'PENDING'){continue}
            $code=Normalize-TimeProgramCode ([string]$e.ProgramCodeF)
            if(-not $code -or (Test-PgmxAutoExcluded $code) -or (Test-TimeCodeHasRealMeasurement $code)){continue}
            $original=[System.IO.Path]::GetFullPath([string]$e.DailyPath)
            if((Normalize-PathForCompare ([System.IO.Path]::GetDirectoryName($original))) -ne (Normalize-PathForCompare $folder.FullName)){
                throw ('Некоректен изходен път в дневния регистър: '+$original)
            }
            if(-not [System.IO.File]::Exists($original)){throw ('PGMX в чакащ дневен регистър липсва: '+$original)}
            $entries.Add([pscustomobject]@{ProgramCodeF=$code;DailyPath=$original;Day=$m.Day;DailyFolder=$m.DailyFolder})
        }
    }
    return @($entries.ToArray())
}

function Get-BatchManifestPath {param([string]$Folder);return (Join-Path $Folder 'PWX500_Batch_Manifest.json')}
function Save-PgmxBatchManifest {param($Batch);$Batch.UpdatedAt=Get-Date -Format 'yyyy-MM-dd HH:mm:ss';Write-CamJsonAtomic -Path (Get-BatchManifestPath $Batch.Folder) -Data $Batch}

function Get-UnfinishedPgmxBatch {
    $base=Join-Path (Get-DailyRootFolder) 'Общи_пакети'
    if(-not [System.IO.Directory]::Exists($base)){return $null}
    $matches=New-Object System.Collections.Generic.List[object]
    foreach($dir in @(Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue | Sort-Object Name)){
        $path=Get-BatchManifestPath $dir.FullName
        if(-not [System.IO.File]::Exists($path)){continue}
        $batch=Read-CamJson -Path $path
        if([string]$batch.Status -in @('PREPARED','MIXX_READY','POWERSET_RUNNING','POWERSET_POSTPROCESS','ERROR','STOPPED')){$matches.Add($batch)}
    }
    if($matches.Count -gt 1){throw 'Има повече от един незавършен общ пакет. Нужна е ръчна проверка, за да не се смесят партиди.'}
    if($matches.Count -eq 1){return $matches[0]}
    return $null
}

function Assert-PgmxBatchNoContentCollision {
    param([object[]]$Entries)
    $hashByName=@{}
    foreach($entry in @($Entries)){
        $file=[System.IO.Path]::GetFileName([string]$entry.DailyPath)
        $key=$file.ToLowerInvariant()
        $hash=(Get-FileHash -LiteralPath $entry.DailyPath -Algorithm SHA256 -ErrorAction Stop).Hash
        if($hashByName.ContainsKey($key) -and $hashByName[$key] -ne $hash){
            throw ('Еднакво име PGMX с различно съдържание в различни дни: '+$file+'. Общият MIXX е спрян; файловете остават на място.')
        }
        $hashByName[$key]=$hash
    }
    return $true
}

function New-PgmxBatch {
    $pending=@(Get-PendingDayPgmxEntries)
    if($pending.Count -eq 0){return $null}
    [void](Assert-PgmxBatchNoContentCollision -Entries $pending)
    $dates=@($pending|ForEach-Object{$_.Day}|Sort-Object -Unique)
    $first=[datetime]::ParseExact($dates[0],'yyyy-MM-dd',[System.Globalization.CultureInfo]::InvariantCulture)
    $last=[datetime]::ParseExact($dates[$dates.Count-1],'yyyy-MM-dd',[System.Globalization.CultureInfo]::InvariantCulture)
    $range=$first.ToString('dd.MM')+'-'+$last.ToString('dd.MM')
    $base=Join-Path (Get-DailyRootFolder) 'Общи_пакети'
    if(-not [System.IO.Directory]::Exists($base)){[void][System.IO.Directory]::CreateDirectory($base)}
    $name=$range+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss')
    $dir=Join-Path $base $name
    if([System.IO.Directory]::Exists($dir)){throw ('Обща папка вече съществува, не я презаписвам: '+$dir)}
    [void][System.IO.Directory]::CreateDirectory($dir)
    $rows=New-Object System.Collections.Generic.List[object]
    $seen=@{}
    foreach($e in $pending){
        Test-OperatorStopRequested
        $name=[System.IO.Path]::GetFileName($e.DailyPath)
        $staged=Join-Path $dir $name
        if(-not $seen.ContainsKey($name.ToLowerInvariant())){
            [System.IO.File]::Copy($e.DailyPath,$staged,$false)
            $seen[$name.ToLowerInvariant()]=$true
        }
        $rows.Add([pscustomobject]@{ProgramCodeF=$e.ProgramCodeF;Day=$e.Day;DailyFolder=$e.DailyFolder;DailyPath=$e.DailyPath;BatchPath=$staged;Status='PENDING'})
    }
    $batch=[pscustomobject]@{Version='5.1.38';Folder=$dir;DateRange=$range;Status='PREPARED';MixxPath='';XlsPath='';ProblemCodes=@();NoTimeCodes=@();UpdatedAt='';Details=@($rows.ToArray())}
    Save-PgmxBatchManifest -Batch $batch
    return $batch
}

function Move-BatchOriginPgmx {
    param([object]$Entry,[ValidateSet('Проблемни_PGMX','Без_времена')][string]$Subfolder)
    $full=[System.IO.Path]::GetFullPath([string]$Entry.DailyPath)
    $daily=[System.IO.Path]::GetFullPath([string]$Entry.DailyFolder)
    if((Normalize-PathForCompare ([System.IO.Path]::GetDirectoryName($full))) -ne (Normalize-PathForCompare $daily)){
        throw ('Блокирано местене извън изходната дневна папка: '+$full)
    }
    if([System.IO.Path]::GetExtension($full) -ine '.pgmx'){throw ('Блокирано местене на не-PGMX: '+$full)}
    $destDir=Join-Path $daily $Subfolder
    if(-not [System.IO.Directory]::Exists($destDir)){[void][System.IO.Directory]::CreateDirectory($destDir)}
    $file=[System.IO.Path]::GetFileName($full)
    $target=Join-Path $destDir $file
    # If interrupted immediately after a successful physical move, do not
    # lose the file or move another one. Replay is safe for this exact path.
    if(-not [System.IO.File]::Exists($full)){
        if([System.IO.File]::Exists($target)){return $target}
        throw ('Изходният PGMX липсва както в дневната, така и в очакваната изолирана папка: '+$full)
    }
    if([System.IO.File]::Exists($target)){
        $stem=[System.IO.Path]::GetFileNameWithoutExtension($file);$ext=[System.IO.Path]::GetExtension($file);$i=2
        do{$target=Join-Path $destDir ($stem+'_'+$i+$ext);$i++}while([System.IO.File]::Exists($target))
    }
    [System.IO.File]::Move($full,$target)
    return $target
}

function Set-BatchCodeStatuses {
    param($Batch,[string[]]$Codes,[ValidateSet('PROBLEM_MIXX','NO_REAL_TIME')][string]$Reason)
    $seen=@{}
    foreach($c in @($Codes)){$k=Normalize-TimeProgramCode $c;if($k){$seen[$k]=$true}}
    foreach($row in @($Batch.Details)){
        if($row.Status -ne 'PENDING' -and $row.Status -ne $Reason){continue}
        $code=Normalize-TimeProgramCode $row.ProgramCodeF
        if(-not $seen.ContainsKey($code)){continue}
        $sub=if($Reason -eq 'PROBLEM_MIXX'){'Проблемни_PGMX'}else{'Без_времена'}
        $destination=Move-BatchOriginPgmx -Entry $row -Subfolder $sub
        $row.Status=$Reason
        Save-PgmxBatchManifest -Batch $Batch
        Register-PgmxAutoExclusion -ProgramCodeF $code -Reason $Reason -Origin $destination
        Update-DayManifestEntry -Entry $row -Status $Reason
    }
}

function Update-DayManifestEntry {
    param($Entry,[string]$Status)
    $path=Get-DayManifestPath -DailyFolder $Entry.DailyFolder
    $day=Read-CamJson -Path $path
    if($null -eq $day){throw ('Липсва дневен регистър: '+$path)}
    foreach($row in @($day.Details)){
        if((Normalize-TimeProgramCode $row.ProgramCodeF) -eq (Normalize-TimeProgramCode $Entry.ProgramCodeF)){
            $row.Status=$Status;$row.UpdatedAt=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
    }
    Save-DayPgmxManifest -Manifest $day
}

function Sync-PgmxAutoExclusionsFromDailyFolders {
    # Upgrade from V5.1.35/37: isolate folders existed before this JSON ledger.
    # Names only; no PGMX content or product/server folder is opened.
    $root=Get-DailyRootFolder
    if(-not [System.IO.Directory]::Exists($root)){return}
    foreach($day in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)){
        if($day.Name -notmatch '^\d{2}\.\d{2}$'){continue}
        $rules=@([pscustomobject]@{Folder='Проблемни_PGMX';Reason='PROBLEM_MIXX'},[pscustomobject]@{Folder='Без_времена';Reason='NO_REAL_TIME'})
        foreach($rule in $rules){
            $dir=Join-Path $day.FullName $rule.Folder
            if(-not [System.IO.Directory]::Exists($dir)){continue}
            foreach($file in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.pgmx' -ErrorAction SilentlyContinue)){
                $code=Normalize-TimeProgramCode $file.Name
                if(-not $code){continue}
                if(-not (Test-PgmxAutoExcluded $code)){
                    Register-PgmxAutoExclusion -ProgramCodeF $code -Reason $rule.Reason -Origin $file.FullName
                }
            }
        }
    }
}

function Test-PgmxPendingWorkInCurrentRoot {
    $root=Get-DailyRootFolder
    if(-not [System.IO.Directory]::Exists($root)){return $false}
    foreach($folder in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)){
        if($folder.Name -notmatch '^\d{2}\.\d{2}$'){continue}
        $path=Get-DayManifestPath -DailyFolder $folder.FullName
        if(-not [System.IO.File]::Exists($path)){continue}
        $day=Read-CamJson -Path $path
        foreach($row in @($day.Details)){if($null -ne $row -and [string]$row.Status -eq 'PENDING'){return $true}}
    }
    return ($null -ne (Get-UnfinishedPgmxBatch))
}

# V5.1.38 BUILD FIX3: a Process object may remain in memory after it exits.
# Block Stage 2 only while a worker is ACTUALLY running, or a time import
# has not finished its guarded commit/cleanup. On uncertain state, fail closed.
function Assert-CumulativeWorkersIdle {
    foreach($workerName in @('IndexWorkerProcess','CsvWorkerProcess','SearchWorkerProcess','HopWorkerProcess','PgmxWorkerProcess','HistoryWorkerProcess','TimeImportProcess')){
        $worker=Get-Variable -Name $workerName -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        if($null -eq $worker){continue}
        $hasExited=$false
        try{
            $hasExited=[bool]$worker.HasExited
        }catch{
            throw ('Cannot safely verify worker state: '+$workerName+'. '+$_.Exception.Message)
        }
        if(-not $hasExited){throw ('Друг worker още е активен: '+$workerName+'. Изчакай приключването му.')}
    }
    if($script:TimeImportRunning -or $script:TimeImportSynchronousRunning){throw 'В момента се импортират времена. Изчакай завършването им.'}
}

function Invoke-CumulativeMixPowerSet {
    if($script:AutoMixChainRunning){throw 'MIXX/PowerSet вече работи.'}
    if($script:TrainingActive){throw 'Спри обучението преди общия MIXX/PowerSet.'}
    if($script:OperatorAutoFlow){throw 'Етап 1 още работи. Общият MIXX/PowerSet не може да започне едновременно.'}
    Assert-CumulativeWorkersIdle
    $batch=Get-UnfinishedPgmxBatch
    if($null -eq $batch){$batch=New-PgmxBatch}
    if($null -eq $batch){
        [System.Windows.Forms.MessageBox]::Show('Няма чакащи PGMX в дневните регистри. Изпълни първо „Подготовка + PGMX“.','CAM Agent - Общ MIXX',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
        return
    }
    $script:OperatorStopRequested=$false
    $script:AutoMixChainRunning=$true
    $script:BatchActiveManifestPath=Get-BatchManifestPath -Folder $batch.Folder
    $script:BatchOutputFolderOverride=[string]$batch.Folder
    $script:LastProblemPgmxMoved=@();$script:LastNoTimePgmxMoved=@();$script:LastMedianFallbackCodes=@()
    try{
        $status.Text=('Общ MIXX + PowerSet: '+$batch.DateRange+' | детайли: '+@($batch.Details).Count)
        [System.Windows.Forms.Application]::DoEvents()
        if(-not [System.IO.Directory]::Exists($batch.Folder)){throw ('Липсва обща папка: '+$batch.Folder)}
        $mixTarget=Get-DailyMixFilePath
        $mixCheckpointOk=([string]$batch.MixxPath -eq [string]$mixTarget) -and [System.IO.File]::Exists($mixTarget) -and (Get-Item -LiteralPath $mixTarget).Length -gt 0
        if(-not $mixCheckpointOk){
            $files=@($batch.Details | Where-Object {$_.Status -eq 'PENDING'} | ForEach-Object {$_.BatchPath} | Sort-Object -Unique)
            if($files.Count -eq 0){throw 'Няма PGMX за общия MIXX.'}
            if([System.IO.File]::Exists($mixTarget)){
                throw ('Предишният MIXX съществува без потвърден checkpoint. Няма да го изтрия или презапиша: '+$mixTarget)
            }
            $mixPhase=[System.Diagnostics.Stopwatch]::StartNew()
            Write-MixLog ('PERF BATCH MIXX begin; files='+$files.Count+'; folder='+$batch.Folder)
            $mix=Invoke-DailyMixAutomation -PgmxFiles $files
            Write-MixLog ('PERF BATCH MIXX complete; seconds='+[Math]::Round($mixPhase.Elapsed.TotalSeconds,2))
            # Keep the warning-code list in the same durable checkpoint as the
            # finished Mixx, BEFORE moving any original dated PGMX files.
            $batch.ProblemCodes=@($script:LastProblemPgmxMoved | ForEach-Object {[System.IO.Path]::GetFileName($_)})
            $batch.MixxPath=$mix;$batch.Status='MIXX_READY';Save-PgmxBatchManifest -Batch $batch
        }
        # A restarted agent must finish every previously detected warning.
        if(@($batch.ProblemCodes).Count -gt 0){Set-BatchCodeStatuses -Batch $batch -Codes @($batch.ProblemCodes) -Reason 'PROBLEM_MIXX'}
        Test-OperatorStopRequested
        $alreadyExported=$false
        if([string]$batch.Status -eq 'POWERSET_POSTPROCESS'){
            if(-not $batch.XlsPath -or -not [System.IO.File]::Exists([string]$batch.XlsPath)){
                throw 'PowerSet има записан checkpoint за XLS, но файлът липсва. Спряно е безопасно.'
            }
            $xls=[string]$batch.XlsPath
            $alreadyExported=$true
        }
        if(-not $alreadyExported){
            $daily=$batch.Folder
            $xlsTarget=Join-Path $daily ('Mix(' + [System.IO.Path]::GetFileName($daily) + ').xls')
            if([System.IO.File]::Exists($xlsTarget)){
                throw ('Открит е XLS без завършен PowerSet checkpoint. Няма да го презапиша автоматично: '+$xlsTarget)
            }
            $batch.Status='POWERSET_RUNNING';Save-PgmxBatchManifest -Batch $batch
            $powerPhase=[System.Diagnostics.Stopwatch]::StartNew()
            Write-PowerSetLog ('PERF BATCH POWERSET begin; MIXX='+$batch.MixxPath)
            $xls=Invoke-PowerSetFullAutomation
            Write-PowerSetLog ('PERF BATCH POWERSET complete; seconds='+[Math]::Round($powerPhase.Elapsed.TotalSeconds,2))
            $batch.XlsPath=$xls
            $batch.NoTimeCodes=@($script:LastMedianFallbackCodes)
            $batch.Status='POWERSET_POSTPROCESS';Save-PgmxBatchManifest -Batch $batch
        }
        if(@($batch.NoTimeCodes).Count -gt 0){Set-BatchCodeStatuses -Batch $batch -Codes @($batch.NoTimeCodes) -Reason 'NO_REAL_TIME'}
        foreach($row in @($batch.Details)){
            if($row.Status -eq 'PENDING'){$row.Status='COMPLETED';Update-DayManifestEntry -Entry $row -Status 'COMPLETED'}
        }
        $batch.Status='COMPLETED';Save-PgmxBatchManifest -Batch $batch
        Add-OperatorProcessHistory -Stage 'BATCH_COMPLETED' -Status 'COMPLETED' -Message ('Общ Mixx/XLS: '+$batch.Folder)
        $status.Text='Общ MIXX + PowerSet + XLS е готов: '+$batch.Folder
        [System.Windows.Forms.MessageBox]::Show('Общият пакет е готов.'+[Environment]::NewLine+$batch.Folder,'CAM Agent - MIXX/PowerSet',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
    }catch{
        $msg=$_.Exception.Message
        try{
            # A verified exported XLS is a stronger resume checkpoint than an
            # error while isolating daily PGMX. Never replace it on a retry.
            if([string]$batch.Status -ne 'POWERSET_POSTPROCESS'){
                if($script:OperatorStopRequested){$batch.Status='STOPPED'}else{$batch.Status='ERROR'}
            }
            Save-PgmxBatchManifest -Batch $batch
        }catch{}
        Write-MixLog ('BATCH ERROR: '+$msg)
        $status.Text='Общ MIXX/PowerSet: '+$msg
        throw
    }finally{
        $script:BatchOutputFolderOverride=''
        $script:BatchActiveManifestPath=''
        $script:AutoMixChainRunning=$false
        try{Refresh-MainOperatorUi}catch{}
    }
}

function Finish-PgmxOnlyStage {
    $result=Prepare-Stage1PgmxDailyFiles
    try{Finalize-CsvRegistryForCurrentBatch -PipelineSucceeded $true -PipelineMessage 'PGMX са подготвени. MIXX/PowerSet очаква отделен ръчен старт.'}catch{}
    Set-OperatorCheckpoint -CurrentStage 'PGMX_READY' -LastSafeStage 'PGMX_DONE' -Status 'COMPLETED' -Message ('PGMX подготвени: '+$result.Copied+'; чакащи за MIXX: '+$result.Pending+'.')
    $script:OperatorAutoFlow=$false
    $status.Text=('Етап 1 ГОТОВ: '+$result.Copied+' PGMX; чакащи '+$result.Pending+'. MIXX и PowerSet НЕ стартират.')
    try{Refresh-OperatorDashboard;Refresh-OperatorReports -WriteFiles}catch{}
    [System.Windows.Forms.MessageBox]::Show($status.Text,'CAM Agent - PGMX',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
}
