$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$RepoRoot = Split-Path -Parent $PSScriptRoot
$Source = Join-Path $RepoRoot 'src\PWX500_Agent_V3.ps1'
if (-not (Test-Path -LiteralPath $Source)) { throw "Missing source: $Source" }

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Source,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    throw ('PowerShell syntax error(s): ' + (($parseErrors | ForEach-Object {$_.Message}) -join ' | '))
}

$wanted = @('Save-OperatorScheduleSettings','Load-OperatorScheduleSettings','Test-OperatorScheduleTick')
$funcs = @{}
$ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true) | ForEach-Object {
    if ($wanted -contains $_.Name) { $funcs[$_.Name] = $_.Extent.Text }
}
foreach ($name in $wanted) {
    if (-not $funcs.ContainsKey($name)) { throw "Missing function in source: $name" }
    Invoke-Expression $funcs[$name]
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('CAMAgentScheduleTest_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    $script:OperatorScheduleFile = Join-Path $tempRoot 'PWX500_Agent_Schedule.ini'
    $script:OperatorScheduleLastRunDate = ''
    $script:OperatorScheduleMode = 'PGMX_ONLY'
    $script:MainScheduleModeCombo = $null
    $operatorScheduleCheck = [pscustomobject]@{ Checked = $true }
    $operatorScheduleTime = [pscustomobject]@{ Value = [datetime]::Today.AddHours(22) }
    $status = [pscustomobject]@{ Text = '' }

    # Worker/busy state used by Test-OperatorScheduleTick.
    $script:CsvWorkerProcess=$null
    $script:SearchWorkerProcess=$null
    $script:HopWorkerProcess=$null
    $script:PgmxWorkerProcess=$null
    $script:AutoMixChainRunning=$false
    $script:TrainingActive=$false
    $script:TimeImportRunning=$false
    $script:TimeImportSynchronousRunning=$false

    # 1) Backward compatibility: old INI without LastRunDate must load and reset in-memory date.
    [System.IO.File]::WriteAllLines($script:OperatorScheduleFile,@('Enabled=1','Time=21:30'),(New-Object System.Text.UTF8Encoding($true)))
    $script:OperatorScheduleLastRunDate='2099-12-31'
    Load-OperatorScheduleSettings
    if (-not $operatorScheduleCheck.Checked) { throw 'Old INI compatibility failed: Enabled was not loaded.' }
    if ($operatorScheduleTime.Value.ToString('HH:mm') -ne '21:30') { throw 'Old INI compatibility failed: Time was not loaded.' }
    if ($script:OperatorScheduleLastRunDate -ne '') { throw 'Old INI compatibility failed: missing LastRunDate did not reset memory.' }

    # 2) Persistence round-trip.
    $today = Get-Date -Format 'yyyy-MM-dd'
    $script:OperatorScheduleLastRunDate=$today
    Save-OperatorScheduleSettings
    $iniLines=[System.IO.File]::ReadAllLines($script:OperatorScheduleFile,[System.Text.Encoding]::UTF8)
    if ($iniLines -notcontains ('LastRunDate=' + $today)) { throw 'LastRunDate was not persisted.' }
    $script:OperatorScheduleLastRunDate=''
    Load-OperatorScheduleSettings
    if ($script:OperatorScheduleLastRunDate -ne $today) { throw 'LastRunDate was not restored after load.' }

    # 3) If already run today, scheduled tick must not call Start-OperatorFullFlow.
    $script:TestStartCount=0
    function Start-OperatorFullFlow {param([string]$Mode='FULL');$script:TestStartCount++;$script:TestLastMode=$Mode}
    $operatorScheduleCheck.Checked=$true
    $operatorScheduleTime.Value=(Get-Date).AddMinutes(-5)
    Test-OperatorScheduleTick
    if ($script:TestStartCount -ne 0) { throw 'Duplicate-run guard failed while LastRunDate is today.' }

    # 4) First eligible tick starts exactly once and writes today BEFORE the flow.
    $script:OperatorScheduleLastRunDate=''
    Save-OperatorScheduleSettings
    Test-OperatorScheduleTick
    if ($script:TestStartCount -ne 1) { throw 'Eligible scheduled tick did not start exactly once.' }
    $iniLines=[System.IO.File]::ReadAllLines($script:OperatorScheduleFile,[System.Text.Encoding]::UTF8)
    if ($iniLines -notcontains ('LastRunDate=' + $today)) { throw 'Scheduled tick did not persist LastRunDate.' }
    Test-OperatorScheduleTick
    if ($script:TestStartCount -ne 1) { throw 'Second tick in the same process triggered a duplicate run.' }

    # 5) Simulate EXE restart: clear memory, reload INI, then tick. It must still not run again.
    $script:OperatorScheduleLastRunDate=''
    Load-OperatorScheduleSettings
    Test-OperatorScheduleTick
    if ($script:TestStartCount -ne 1) { throw 'Restart simulation triggered a duplicate scheduled run.' }

    if($script:TestLastMode -ne 'PGMX_ONLY'){throw 'Schedule did not start PGMX-only safe mode by default.'}
    $script:MainScheduleModeCombo=[pscustomobject]@{SelectedIndex=1}
    Save-OperatorScheduleSettings
    $script:OperatorScheduleMode='PGMX_ONLY'
    Load-OperatorScheduleSettings
    if($script:OperatorScheduleMode -ne 'FULL' -or $script:MainScheduleModeCombo.SelectedIndex -ne 1){throw 'FULL automatic schedule mode did not survive restart.'}
    $script:MainScheduleModeCombo.SelectedIndex=0
    Save-OperatorScheduleSettings
    Load-OperatorScheduleSettings
    if($script:OperatorScheduleMode -ne 'PGMX_ONLY'){throw 'PGMX-only schedule mode did not survive restart.'}

    Write-Host 'SCHEDULE PERSISTENCE TEST: PASS' -ForegroundColor Green
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
