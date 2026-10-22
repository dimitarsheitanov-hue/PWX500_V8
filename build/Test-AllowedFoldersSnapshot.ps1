$ErrorActionPreference='Stop'
$src=Join-Path (Split-Path -Parent $PSScriptRoot) 'src\PWX500_Agent_V3.ps1'
$tok=$null;$errs=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($src,[ref]$tok,[ref]$errs)
if($errs.Count -gt 0){throw 'Parse error in main source.'}
$names=@('Save-AllowedFolders','Normalize-PathForCompare')
$funcs=@{}
$ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)|ForEach-Object{
    if($names -contains $_.Name){$funcs[$_.Name]=$_.Extent.Text}
}
foreach($n in $names){if(-not $funcs.ContainsKey($n)){throw ('Missing: '+$n)};Invoke-Expression $funcs[$n]}
function Convert-ToCsvField {param($x);return [string]$x}
$test=Join-Path ([System.IO.Path]::GetTempPath()) ('CAMAgentRoots_'+[guid]::NewGuid().ToString('N')+'.csv')
try{
    $script:SettingsCsv=$test
    $script:SettingsVersion='4'
    $script:DefaultAllowedFolders=@('Y:\1_Izpolzvani\01_KM\_0_10_F','Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\08_Tapicirana_mebel')
    $script:AllowedFolders=New-Object System.Collections.Generic.List[string]
    $script:AllowedFolders.Add('Y:\1_Izpolzvani\01_KM\_0_10_F')
    $script:AllowedFolders.Add('Y:\1_Izpolzvani\02_TM\Additional_Approved')
    Save-AllowedFolders
    $rows=@(Import-Csv -LiteralPath $test -Delimiter ';')
    if($rows.Count -ne 3){throw ('Expected default+custom union of 3, got '+$rows.Count)}
    foreach($required in @($script:DefaultAllowedFolders)+@($script:AllowedFolders.ToArray())){
        if($required -notin @($rows|ForEach-Object{$_.Path})){throw ('PGMX worker cannot see allowed root: '+$required)}
    }
    $stamp=(Get-Item -LiteralPath $test).LastWriteTimeUtc
    Start-Sleep -Milliseconds 100
    Save-AllowedFolders
    if((Get-Item -LiteralPath $test).LastWriteTimeUtc -ne $stamp){throw 'Unchanged whitelist caused a catalog invalidation.'}
    Write-Host 'PASS V5.1.38 allowed-root tests: default and custom same version-4 worker snapshot.' -ForegroundColor Green
}finally{Remove-Item -LiteralPath $test -Force -ErrorAction SilentlyContinue}
