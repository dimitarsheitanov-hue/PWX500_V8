# Run in GitHub Windows PowerShell 5.1 before publishing the EXE.
# Filesystem-only regression: no Maestro, no PowerSet, no network, no converter.
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$module=Join-Path $root 'src\CAM_Agent_BatchEngine_V5_1_38.ps1'
$main=Join-Path $root 'src\PWX500_Agent_V3.ps1'
$worker=Join-Path $root 'src\PWX500_PgmxWorker_V3.ps1'
foreach($file in @($module,$main,$worker)){
    $tok=$null;$err=$null
    [void][System.Management.Automation.Language.Parser]::ParseFile($file,[ref]$tok,[ref]$err)
    if($err.Count -gt 0){throw ('PowerShell parser errors in '+$file+': '+(($err|ForEach-Object{$_.Message}) -join ' | '))}
}
$raw=[System.IO.File]::ReadAllText($main,[System.Text.Encoding]::UTF8)
if(-not $raw.Contains('Test-PgmxAutoExcluded $pc')){throw 'Missing automatic exclusion from time candidates.'}
if(-not $raw.Contains('Finish-PgmxOnlyStage')){throw 'PGMX-only exit hook missing.'}
if(-not $raw.Contains('Invoke-CumulativeMixPowerSet')){throw 'Separate batch button missing.'}
if(-not $raw.Contains('$script:OperatorScheduleMode')){throw 'Schedule Mode field missing.'}
if(-not $raw.Contains('$effective=@($script:DefaultAllowedFolders)+@($script:AllowedFolders.ToArray())')){throw 'Worker whitelist snapshot not unified.'}

$temp=Join-Path ([System.IO.Path]::GetTempPath()) ('CAMAgentBatchTest_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp -Force | Out-Null
try{
    $script:AppRoot=$temp;$script:TestDailyRoot=Join-Path $temp 'FOUND'
    New-Item -ItemType Directory -Path $script:TestDailyRoot -Force|Out-Null
    $script:PgmxLastMode='Execute'
    $script:AutoMixChainRunning=$false
    $script:OperatorStopRequested=$false
    function Normalize-TimeProgramCode {param([string]$c);return ([System.IO.Path]::GetFileName(([string]$c)).Trim().ToLowerInvariant())}
    function Normalize-PathForCompare {param([string]$p);return (([string]$p).TrimEnd('\').ToLowerInvariant())}
    function Get-DailyRootFolder{return $script:TestDailyRoot}
    function Test-OperatorStopRequested {if($script:OperatorStopRequested){throw 'STOP'}}
    function Test-TimeCodeHasRealMeasurement {param([string]$c);return $false}
    . $module

    # BUILD FIX3 regression: an exited TimeImportProcess handle must NOT
    # block Stage 2; a genuinely running worker or active import MUST block.
    $script:TimeImportProcess=[pscustomobject]@{HasExited=$true}
    $script:TimeImportRunning=$false
    $script:TimeImportSynchronousRunning=$false
    Assert-CumulativeWorkersIdle
    $script:TimeImportProcess=[pscustomobject]@{HasExited=$false}
    $blocked=$false
    try{Assert-CumulativeWorkersIdle}catch{$blocked=$true}
    if(-not $blocked){throw 'Live TimeImportProcess was incorrectly allowed.'}
    $script:TimeImportProcess=[pscustomobject]@{HasExited=$true}
    $script:TimeImportRunning=$true
    $blocked=$false
    try{Assert-CumulativeWorkersIdle}catch{$blocked=$true}
    if(-not $blocked){throw 'Active time-import cleanup flag was incorrectly allowed.'}
    $script:TimeImportRunning=$false
    $script:TimeImportSynchronousRunning=$true
    $blocked=$false
    try{Assert-CumulativeWorkersIdle}catch{$blocked=$true}
    if(-not $blocked){throw 'Synchronous time import was incorrectly allowed.'}
    $script:TimeImportSynchronousRunning=$false
    $script:TimeImportProcess=$null
    $script:PgmxWorkerProcess=[pscustomobject]@{HasExited=$false}
    $blocked=$false
    try{Assert-CumulativeWorkersIdle}catch{$blocked=$true}
    if(-not $blocked){throw 'Live PGMX worker was incorrectly allowed.'}
    $script:PgmxWorkerProcess=[pscustomobject]@{HasExited=$true}
    Assert-CumulativeWorkersIdle
    $script:PgmxWorkerProcess=$null
    Assert-CumulativeWorkersIdle
    Write-Host 'PASS V5.1.38 FIX3 worker gate: exited handle allowed; live worker/import blocked.' -ForegroundColor Green

    $d1=Join-Path $script:TestDailyRoot '20.10';$d2=Join-Path $script:TestDailyRoot '21.10'
    [void][System.IO.Directory]::CreateDirectory($d1);[void][System.IO.Directory]::CreateDirectory($d2)
    $a=Join-Path $d1 '01111111.pgmx';$b=Join-Path $d2 '01111111.pgmx'
    [System.IO.File]::WriteAllText($a,'same PGMX')
    [System.IO.File]::WriteAllText($b,'same PGMX')
    $testPairs=@([pscustomobject]@{Folder=$d1;File=$a;Day='2026-10-20'},[pscustomobject]@{Folder=$d2;File=$b;Day='2026-10-21'})
    foreach($pair in $testPairs){
        $item=[pscustomobject]@{ProgramCodeF='01111111.pgmx';DailyPath=$pair.File;SourcePath=$pair.File;Status='PENDING';UpdatedAt=''}
        $manifest=[pscustomobject]@{Version='5.1.38';Day=$pair.Day;DailyFolder=$pair.Folder;UpdatedAt='';Details=@($item)}
        Save-DayPgmxManifest -Manifest $manifest
    }
    $pending=@(Get-PendingDayPgmxEntries)
    if($pending.Count -ne 2){throw ('Expected 2 pending daily files, found '+$pending.Count)}
    [void](Assert-PgmxBatchNoContentCollision -Entries $pending)
    $batch=New-PgmxBatch
    if(@($batch.Details).Count -ne 2){throw 'Common batch lost origins.'}
    if(@(Get-ChildItem -LiteralPath $batch.Folder -File -Filter '*.pgmx').Count -ne 1){throw 'Identical codes must be staged only once.'}
    if(-not (Test-Path -LiteralPath (Get-BatchManifestPath $batch.Folder))){throw 'Missing persistent batch manifest.'}
    $state=Get-UnfinishedPgmxBatch
    if($null -eq $state -or $state.Status -ne 'PREPARED'){throw 'Interrupted batch was not found for resume.'}

    $fake=Join-Path $d2 '02222222.pgmx';[System.IO.File]::WriteAllText($fake,'different PGMX')
    $collision=@([pscustomobject]@{DailyPath=$a},[pscustomobject]@{DailyPath=$b})
    [System.IO.File]::WriteAllText($b,'changed')
    $rejected=$false
    try{[void](Assert-PgmxBatchNoContentCollision -Entries $collision)}catch{$rejected=$true}
    if(-not $rejected){throw 'Different same-name PGMX must fail BEFORE Maestro.'}
    [System.IO.File]::WriteAllText($b,'same PGMX')

    $first=$batch.Details[0]
    $before=[string]$first.DailyPath
    $isolated=Move-BatchOriginPgmx -Entry $first -Subfolder 'Проблемни_PGMX'
    if([System.IO.File]::Exists($before) -or -not [System.IO.File]::Exists($isolated)){throw 'Original was not safely moved to its day.'}
    $replayed=Move-BatchOriginPgmx -Entry $first -Subfolder 'Проблемни_PGMX'
    if($replayed -ne $isolated){throw 'Interrupted isolation replay did not find the exact existing file.'}
    Register-PgmxAutoExclusion -ProgramCodeF $first.ProgramCodeF -Reason 'PROBLEM_MIXX' -Origin $isolated
    if(-not (Test-PgmxAutoExcluded $first.ProgramCodeF)){throw 'Isolation not remembered.'}
    $script:AutoExcludedPgmxMap=$null
    if(-not (Test-PgmxAutoExcluded $first.ProgramCodeF)){throw 'Isolation not persistent across restart.'}
    if(@(Get-PendingDayPgmxEntries).Count -ne 0){throw 'Previously isolated code was selected for automatic retest.'}

    $unmeasured=Join-Path $d2 '02222222.pgmx'
    $row=[pscustomobject]@{ProgramCodeF='02222222.pgmx';DailyPath=$unmeasured;SourcePath=$unmeasured;Status='PENDING';UpdatedAt=''}
    $manifest=Read-CamJson -Path (Get-DayManifestPath $d2)
    $manifest.Details=@($manifest.Details)+@($row)
    Save-DayPgmxManifest -Manifest $manifest
    $isolatedTime=Move-BatchOriginPgmx -Entry ([pscustomobject]@{DailyPath=$unmeasured;DailyFolder=$d2}) -Subfolder 'Без_времена'
    Register-PgmxAutoExclusion -ProgramCodeF '02222222.pgmx' -Reason 'NO_REAL_TIME' -Origin $isolatedTime
    $script:AutoExcludedPgmxMap=$null
    if(-not (Test-PgmxAutoExcluded '02222222.pgmx')){throw 'No-time exclusion not persisted.'}
    if([System.IO.File]::Exists($unmeasured)){throw 'No-time PGMX not isolated in original date folder.'}

    $legacy=Join-Path $d1 'Без_времена';[void][System.IO.Directory]::CreateDirectory($legacy)
    [System.IO.File]::WriteAllText((Join-Path $legacy '03333333.pgmx'),'legacy')
    Sync-PgmxAutoExclusionsFromDailyFolders
    if(-not (Test-PgmxAutoExcluded '03333333.pgmx')){throw 'Old isolated files were not migrated into exclusions.'}
    Write-Host 'PASS V5.1.38 batch tests: persistent ledger, multi-date mapping, deduplication, collision stop, isolation and restart.' -ForegroundColor Green
}finally{
    try{Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction Stop}catch{}
}
