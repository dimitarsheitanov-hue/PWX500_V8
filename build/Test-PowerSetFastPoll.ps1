# V5.1.38 FIX6 - PowerSet completion detection performance + false-success regression.
# Runs on GitHub windows-latest / Windows PowerShell 5.1 without Maestro.
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
$source=Join-Path (Split-Path -Parent $PSScriptRoot) 'src\PWX500_Agent_V3.ps1'
$tok=$null;$err=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($source,[ref]$tok,[ref]$err)
if($err.Count -gt 0){throw 'Main source parser failure in PowerSet fast polling regression.'}
$definitions=@($ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true))
$names=@('Get-PowerSetOptimizationUiSnapshot','Wait-PowerSetOptimizationComplete')
foreach($name in $names){if(@($definitions|Where-Object{$_.Name -eq $name}).Count -ne 1){throw ('PowerSet poll function missing: '+$name)}}
$snapshot=@($definitions|Where-Object{$_.Name -eq 'Get-PowerSetOptimizationUiSnapshot'})[0].Extent.Text
$wait=@($definitions|Where-Object{$_.Name -eq 'Wait-PowerSetOptimizationComplete'})[0].Extent.Text
if(@([regex]::Matches($snapshot, '[.]FindAll\(')).Count -ne 1){throw 'PowerSet state snapshot must enumerate UIA only once per poll.'}
if($wait -match 'Get-MixUiElementsByName'){throw 'PowerSet wait still repeats full UIA scans per result signal.'}
if($wait -notmatch 'ResultSignals -ge 2'){throw 'PowerSet completion threshold changed unexpectedly.'}
if($wait -notmatch 'TimeoutMinutes'){throw 'PowerSet timeout/stop behavior missing.'}
if($wait -notmatch 'Test-OperatorStopRequested'){throw 'PowerSet stop behavior missing.'}

# Run the real wait function with a mocked read-only UIA snapshot.
# Progress is NOT accepted until two independent completion labels are present.
Invoke-Expression $wait
$script:TestSnapshotCalls=0
$script:TestLog=@()
function Test-OperatorStopRequested {}
function Get-MaestroProcess {return [pscustomobject]@{Id=4711}}
function Get-PowerSetOptimizationUiSnapshot {
    param([int]$ProcessId)
    $script:TestSnapshotCalls++
    if($ProcessId -ne 4711){throw 'Unexpected process ID in PowerSet mock'}
    $signals=if($script:TestSnapshotCalls -eq 1){0}elseif($script:TestSnapshotCalls -eq 2){1}else{2}
    return [pscustomobject]@{Available=$true;Busy=($script:TestSnapshotCalls -eq 2);ResultSignals=$signals;ScanMs=5}
}
function Write-PowerSetLog {param([string]$Message);$script:TestLog+=([string]$Message)}
function Set-PowerSetTestStatus {param([string]$Text)}
$answer=Wait-PowerSetOptimizationComplete -InitialProcessId 4711 -TimeoutMinutes 1
if($answer -ne 4711 -or $script:TestSnapshotCalls -ne 3){throw ('PowerSet must continue only after two signals; calls='+$script:TestSnapshotCalls)}
if(@($script:TestLog|Where-Object{$_ -like 'Optimization completed.*'}).Count -ne 1){throw 'PowerSet completion log missing.'}
Write-Host 'PASS FIX6: single UIA snapshot per poll, no false completion, stop/timeout unchanged.' -ForegroundColor Green
