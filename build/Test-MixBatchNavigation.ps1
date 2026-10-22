# V5.1.38 BUILD FIX4 - filesystem/path and functional isolation tests.
# Windows PowerShell 5.1. No Maestro or PowerSet required.
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$source=Join-Path $root 'src\PWX500_Agent_V3.ps1'
$tokens=$null;$errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors.Count -gt 0){throw ('Main script parse errors: '+(($errors|ForEach-Object{$_.Message}) -join ' | '))}
$definitions=@($ast.FindAll({param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true))
$names=@('Get-MixBatchBrowserSteps','Get-MixFileBrowserSnapshot','Get-MixBatchBrowserSnapshot','Get-MixBrowserObservedPath','Find-MixBatchCenterFolderItem','Open-MixBatchFolderInDialog','ConvertTo-MixWindowsFolderPath','Select-MixDailyFolderInDialog')
foreach($n in $names){
    $found=@($definitions|Where-Object{$_.Name -eq $n})
    if($found.Count -ne 1){throw ('Expected one batch navigation function '+$n+'; got '+$found.Count)}
}
foreach($n in @('Get-MixBatchBrowserSteps','Get-MixBrowserObservedPath','ConvertTo-MixWindowsFolderPath')){
    $f=@($definitions|Where-Object{$_.Name -eq $n})[0]
    Invoke-Expression $f.Extent.Text
}
$whole=[System.IO.File]::ReadAllText($source,[System.Text.Encoding]::UTF8)
$selector=@($definitions|Where-Object{$_.Name -eq 'Select-MixDailyFolderInDialog'})[0].Extent.Text
$addSnapshot=@($definitions|Where-Object{$_.Name -eq 'Get-MixFileBrowserSnapshot'})[0].Extent.Text
if($addSnapshot.Contains('$_.Current.IsEnabled')){throw 'A disabled Add programs Confirm must remain visible as a structural anchor.'}
# FIX6: a disabled Confirm is still a valid anchor; it is not guaranteed to
# receive UIA hit-testing. A disabled button is NOT clicked by navigation.
if($addSnapshot -notmatch 'confirmAnchorValid' -or $addSnapshot -notmatch 'Current.IsEnabled' -or
   $addSnapshot -notmatch 'scopeBox' -or $addSnapshot -notmatch 'Test-MixElementWithin \$confirmHit'){
    throw 'Disabled Add programs Confirm is not handled safely in FIX6.'
}
$center=@($definitions|Where-Object{$_.Name -eq 'Find-MixBatchCenterFolderItem'})[0].Extent.Text
if($center -notmatch 'ItemContainerPattern' -or $center -notmatch "SendWait\('\{PGDN\}'\)" -or
   $center -notmatch 'Test-OperatorStopRequested'){
    throw 'Exact center-list folder lookup is missing virtualization/scroll fallback.'
}

if($selector -notmatch '\$script:BatchOutputFolderOverride' -or $selector -notmatch 'Open-MixBatchFolderInDialog' -or $selector -notmatch 'Select-MixDailyFolderInDialogDesktopLegacy'){
    throw 'The batch-only selector or the unchanged daily route is missing.'
}
$fn=@($definitions|Where-Object{$_.Name -eq 'Open-MixBatchFolderInDialog'})[0].Extent.Text
foreach($symbol in @('Add','Choose','Save','Get-MixBrowserObservedPath','Find-MixBatchCenterFolderItem','Invoke-MixNativeClick','Get-MixBatchBrowserSnapshot','Test-MixExpectedFileVisible')){
    if(-not $fn.Contains($symbol)){throw ('Batch browser missing '+$symbol)}
}
$stable=@(
    [pscustomobject]@{Name='Select-MixDailyFolderInDialogDesktopLegacy';Hash='953330af0330d601d43cdf8b23d16964b268c0d456b26b8f1db8b7d48c306713'}
    [pscustomobject]@{Name='Select-MixExactDailyFolder';Hash='9b9735369c79fefc3cde78ca7371180268e1b409db0c75d41abffadc5ef2243b'}
    [pscustomobject]@{Name='Open-MixSaveDailyFolder';Hash='9ec352b112f0c589775cf916e739863c25fe998e062db5c45393b8d56f42e9b5'}
    [pscustomobject]@{Name='Invoke-DailyMixAutomation';Hash='1a15de834199bbec778a9122baf79c85414d425e76bfc59a1dbc6e746cdba41f'}
    [pscustomobject]@{Name='Select-AllMixPgmxFilesInDialog';Hash='603229769fe297c2a7058917e47e793d9baf6a528ab555bf5169d3389cec49c5'}
    [pscustomobject]@{Name='Invoke-PowerSetStartRunAutomation';Hash='0014d696bcaca647dbb8d0152fb91037641ded84d5ed3ff41830a7e7de939e09'}
    [pscustomobject]@{Name='Invoke-PowerSetFullAutomation';Hash='22d08e3dfce0d59bf0027a281a538a7987e89a91a4096dcb74b46fbccc346b94'}
    [pscustomobject]@{Name='Open-MixTreeAbsoluteDailyPath';Hash='d7fda241ee2daa45ead4be089939a0c97dc0ec85c7d05f2118c5ebd7254361b2'}
    [pscustomobject]@{Name='Get-MixSaveFolderSnapshot';Hash='b039cb8a3f1d218773196babcc75ac37aac2b1ec791b1ef98b4e3a3b7bea1d83'}
)
foreach($entry in $stable){
    $n=[string]$entry.Name;$expected=[string]$entry.Hash
    # Hash full function blocks (including comments and spacing to the next
    # function), just like the baseline Python extractor. Parser validation
    # above is separate and still checks the AST for every function.
    $rx='(?ms)^function[ \t]+'+[regex]::Escape($n)+'[ \t]*\{.*?(?=^function[ \t]|\z)'
    $m=[regex]::Match($whole,$rx)
    if(-not $m.Success){throw ('Stable navigation function missing: '+$n)}
    $text=$m.Value.Replace("`r`n","`n")
    $bytes=[System.Text.Encoding]::UTF8.GetBytes($text)
    $sha=[System.Security.Cryptography.SHA256]::Create()
    try{$actual=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
    if($actual -ne $expected){throw ('Stable navigation or PowerSet function changed unexpectedly: '+$n)}
}
$desktop='C:\Users\Operator\Desktop'
$shot=[pscustomobject]@{
    IsActive=$true;ContentList=[pscustomobject]@{Mock=$true};
    BreadcrumbTokens=@('Desktop','FOUND');DesktopAnchorIsRoot=$true;CurrentAbsolutePath=''
}
$found=Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop
if($found -ne 'c:\users\operator\desktop\found'){throw ('Breadcrumb parent mismatch: '+$found)}
$shot.BreadcrumbTokens=@('Desktop','FOUND','Общи_пакети')
$sub=Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop
if($sub -ne 'c:\users\operator\desktop\found\общи_пакети'){throw ('Breadcrumb batch base mismatch: '+$sub)}
$shot.BreadcrumbTokens=@('Desktop','FOUND','Общи_пакети','22.10-22.10_20261022_121647')
$final=Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop
if($final -ne 'c:\users\operator\desktop\found\общи_пакети\22.10-22.10_20261022_121647'){throw ('Breadcrumb batch leaf mismatch: '+$final)}
# Selecting a folder row while the breadcrumb is still FOUND never proves it opened.
$shot.BreadcrumbTokens=@('Desktop','FOUND')
if((Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop) -eq $final){throw 'Selected folder treated as opened.'}
# A contradictory absolute address or missing Desktop identity rejects navigation.
$shot.BreadcrumbTokens=@('Desktop','FOUND','Общи_пакети','22.10-22.10_20261022_121647')
$shot.CurrentAbsolutePath='C:\Users\Operator\Desktop\FOUND'
if(Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop){throw 'Contradictory absolute address was accepted.'}
$shot.CurrentAbsolutePath='';$shot.DesktopAnchorIsRoot=$false
if(Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop){throw 'Unverified Desktop anchor was accepted.'}
$shot.DesktopAnchorIsRoot=$true;$shot.BreadcrumbTokens=@('Desktop','FOUND','..','Общи_пакети')
if(Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop){throw 'Parent traversal token accepted.'}
$shot.IsActive=$false;$shot.BreadcrumbTokens=@('Desktop','FOUND')
if(Get-MixBrowserObservedPath -Snapshot $shot -DesktopPath $desktop){throw 'Inactive dialog accepted.'}
$temp=Join-Path ([System.IO.Path]::GetTempPath()) ('CAMAgentMixBrowser_'+[guid]::NewGuid().ToString('N'))
try{
    $root=Join-Path $temp 'FOUND'
    $base=Join-Path $root 'Общи_пакети'
    $batch=Join-Path $base '22.10-22.10_20261022_121647'
    [void][System.IO.Directory]::CreateDirectory($batch)
    $steps=@(Get-MixBatchBrowserSteps -RootPath $root -BatchPath $batch)
    if($steps.Count -ne 3 -or $steps[0] -ne $root -or $steps[1] -ne $base -or $steps[2] -ne $batch){throw 'Batch steps not strictly rooted and ordered.'}
    $stray=Join-Path $root '22.10'
    [void][System.IO.Directory]::CreateDirectory($stray)
    $failed=$false
    try{$null=Get-MixBatchBrowserSteps -RootPath $root -BatchPath $stray}catch{$failed=$true}
    if(-not $failed){throw 'A daily folder was incorrectly accepted as a cumulative batch.'}
    $failed=$false
    try{$null=Get-MixBatchBrowserSteps -RootPath $root -BatchPath (Join-Path $base 'nonexistent')}catch{$failed=$true}
    if(-not $failed){throw 'A missing folder was incorrectly accepted as a batch.'}
}finally{try{Remove-Item -LiteralPath $temp -Recurse -Force}catch{}}
Write-Host 'PASS BUILD FIX6: existing Maestro/PowerSet functions unchanged; batch navigation paths, selected-vs-open, anchor, absolute address, and dialog safety.' -ForegroundColor Green
