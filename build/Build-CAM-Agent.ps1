$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$RepoRoot = Split-Path -Parent $PSScriptRoot
$SrcDir   = Join-Path $RepoRoot 'src'
$DistDir  = Join-Path $RepoRoot 'dist'
$WorkDir  = Join-Path $RepoRoot 'build\_generated'
$PackageDir = Join-Path $DistDir 'CAM_Agent_V5_1_38_COMPILED'

$MainSource = Join-Path $SrcDir 'PWX500_Agent_V3.ps1'
$CompileSource = Join-Path $WorkDir 'PWX500_Agent_V3_COMPILE.ps1'
$IconPath = Join-Path $SrcDir 'CAM_Agent_Icon.ico'
$BatchSource = Join-Path $SrcDir 'CAM_Agent_BatchEngine_V5_1_38.ps1'
$ExePath  = Join-Path $DistDir 'CAM_Agent_V5_1_38.exe'
$ZipPath  = Join-Path $DistDir 'CAM_Agent_V5_1_38_COMPILED_PACKAGE.zip'

if (-not (Test-Path -LiteralPath $MainSource)) { throw "Missing source: $MainSource" }
if (-not (Test-Path -LiteralPath $IconPath))   { throw "Missing icon: $IconPath" }
if (-not (Test-Path -LiteralPath $BatchSource)) { throw "Missing batch engine: $BatchSource" }

Remove-Item -LiteralPath $DistDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $DistDir -Force | Out-Null
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
New-Item -ItemType Directory -Path $PackageDir -Force | Out-Null

# Keep the validated V5.1.38 source intact. Only the generated compile copy is
# adapted for PS2EXE. In a compiled EXE, $PSScriptRoot is not reliable inside
# functions, so every runtime reference is redirected to one script-scoped root.
$sourceText = Get-Content -LiteralPath $MainSource -Raw -Encoding UTF8
$marker = 'Add-Type -AssemblyName System.Windows.Forms'
if ($sourceText.IndexOf($marker, [System.StringComparison]::Ordinal) -lt 0) {
    throw 'Could not find the expected WinForms marker in PWX500_Agent_V3.ps1.'
}

# Compile-copy only: replace the automatic script variable token with a stable
# script-scoped variable. This changes paths only, not any production logic.
$compileBody = $sourceText.Replace('$PSScriptRoot', '$script:AppRoot')

# PS2EXE FIX2: inline the V5.1.38 batch/status engine ONLY in the generated
# compile copy. Runtime dot-sourcing can leave functions invisible in a
# compiled script's invocation scope, even when the external file exists.
# The original src/PWX500_Agent_V3.ps1 and external batch module are untouched.
$batchText = Get-Content -LiteralPath $BatchSource -Raw -Encoding UTF8
if ([string]::IsNullOrWhiteSpace($batchText)) { throw 'Batch engine source is empty.' }
$batchText = $batchText.TrimStart([char]0xFEFF).Replace('$PSScriptRoot', '$script:AppRoot')
$batchAnchor = '# V5.1.38 isolated batch/status engine.'
$batchStart = $compileBody.IndexOf($batchAnchor, [System.StringComparison]::Ordinal)
if ($batchStart -lt 0) { throw 'Batch engine load anchor is missing in main source.' }
if ($compileBody.IndexOf($batchAnchor, $batchStart + 1, [System.StringComparison]::Ordinal) -ge 0) {
    throw 'Batch engine load anchor is not unique.'
}
$batchEnd = $compileBody.IndexOf('$form = New-Object System.Windows.Forms.Form', $batchStart, [System.StringComparison]::Ordinal)
if ($batchEnd -lt 0) { throw 'Could not locate expected form marker after batch engine load.' }
$oldLoad = $compileBody.Substring($batchStart, $batchEnd - $batchStart)
if ($oldLoad -notmatch 'CAM_Agent_BatchEngine_V5_1_38[.]ps1' -or
    $oldLoad -notmatch '(?m)^\.[ \t]+\$batchEngine[ \t]*\r?$') {
    throw 'Unexpected batch engine load block. Aborting instead of modifying unknown source.'
}
$embeddedBlock = "# --- V5.1.38 batch engine EMBEDDED in compiled EXE ---`r`n" +
                 $batchText + "`r`n# --- End embedded V5.1.38 batch engine ---`r`n`r`n"
$compileBody = $compileBody.Substring(0, $batchStart) + $embeddedBlock + $compileBody.Substring($batchEnd)
if ($compileBody.Contains('. $batchEngine')) { throw 'Batch engine still dot-sourced in compiled copy.' }


$compat = @'
# --- GitHub EXE build compatibility only ---
# Stable directory of the compiled EXE. PS2EXE exposes $ScriptRoot, but use a
# process-image fallback as well so the value is never dependent on current dir.
$script:AppRoot = ''
$ps2exeRootVar = Get-Variable -Name ScriptRoot -ErrorAction SilentlyContinue
if ($ps2exeRootVar -and -not [string]::IsNullOrWhiteSpace([string]$ps2exeRootVar.Value)) {
    $script:AppRoot = [string]$ps2exeRootVar.Value
}
if ([string]::IsNullOrWhiteSpace([string]$script:AppRoot)) {
    try {
        $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if (-not [string]::IsNullOrWhiteSpace([string]$exePath)) {
            $script:AppRoot = Split-Path -Parent $exePath
        }
    } catch {}
}
if ([string]::IsNullOrWhiteSpace([string]$script:AppRoot)) {
    throw 'CAM Agent EXE cannot determine its application folder.'
}
# --- End GitHub EXE build compatibility ---

'@

$compileText = $compileBody.Replace($marker, ($compat + $marker))

# Build-time guard: no source runtime reference may still depend on PSScriptRoot.
if ($compileText.Contains('$PSScriptRoot')) {
    throw 'Compile compatibility failed: a $PSScriptRoot reference remains in generated EXE source.'
}
Set-Content -LiteralPath $CompileSource -Value $compileText -Encoding UTF8

# Syntax validation before compiling.
$tokens = $null
$parseErrors = $null
$compiledAst = [System.Management.Automation.Language.Parser]::ParseFile($CompileSource, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    $msg = ($parseErrors | ForEach-Object { $_.Message }) -join [Environment]::NewLine
    throw "PowerShell syntax validation failed:`r`n$msg"
}

# Build-time visibility guard: all batch module functions must exist as normal
# definitions in the main compiled script. This catches the user-reported
# "Sync-PgmxAutoExclusionsFromDailyFolders is not recognized" failure class.
$moduleFunctions = @([regex]::Matches($batchText, '(?m)^function[ \t]+([A-Za-z0-9_-]+)\b') | ForEach-Object { $_.Groups[1].Value })
if ($moduleFunctions.Count -lt 20) { throw 'Embedded batch engine function list is unexpectedly short.' }
$compiledFunctions = @($compiledAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
foreach ($fn in $moduleFunctions) {
    if (@($compiledFunctions | Where-Object { $_ -eq $fn }).Count -ne 1) {
        throw ('Embedded batch function missing or duplicated in compiled source: ' + $fn)
    }
}
if (-not $compiledFunctions.Contains('Sync-PgmxAutoExclusionsFromDailyFolders')) {
    throw 'The isolated PGMX synchronizer is not available in the compiled source.'
}
Write-Host ('PASS: ' + $moduleFunctions.Count + ' batch engine functions embedded directly in main EXE source.') -ForegroundColor Green

# Regression test: scheduled automatic run must persist LastRunDate and survive an EXE restart.
& (Join-Path $PSScriptRoot 'Test-SchedulePersistence.ps1')

# These fail the build before EXE publishing on duplicate/retry/collision regressions.
& (Join-Path $PSScriptRoot 'Test-BatchEngine.ps1')
& (Join-Path $PSScriptRoot 'Test-AllowedFoldersSnapshot.ps1')
& (Join-Path $PSScriptRoot 'Test-RootDriveAliases.ps1')
& (Join-Path $PSScriptRoot 'Test-MixBatchNavigation.ps1')
& (Join-Path $PSScriptRoot 'Test-PowerSetFastPoll.ps1')
& (Join-Path $PSScriptRoot 'Test-CoreIntegrity.ps1')

# Reproducible compiler version.
$requiredPs2Exe = [version]'1.0.18'
$installed = Get-Module -ListAvailable -Name ps2exe | Where-Object { $_.Version -eq $requiredPs2Exe } | Select-Object -First 1
if (-not $installed) {
    try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop } catch {}
    Install-Module -Name ps2exe -RequiredVersion $requiredPs2Exe -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
}
Import-Module ps2exe -RequiredVersion $requiredPs2Exe -Force

Write-Host "Compiling CAM Agent V5.1.38 with PS2EXE $requiredPs2Exe ..."
Invoke-ps2exe `
    -inputFile $CompileSource `
    -outputFile $ExePath `
    -x64 `
    -STA `
    -noConsole `
    -noOutput `
    -noError `
    -iconFile $IconPath `
    -title 'CAM Agent V5.1.38' `
    -description 'PWX500 / P800 / FTT CAM Agent' `
    -company 'KM' `
    -product 'CAM Agent' `
    -version '5.1.38.0'

if (-not (Test-Path -LiteralPath $ExePath)) { throw 'Compilation did not create CAM_Agent_V5_1_38.exe.' }
if ((Get-Item -LiteralPath $ExePath).Length -lt 100KB) { throw 'Compiled EXE is unexpectedly small; build stopped.' }

# Runtime dependencies that must stay beside the compiled EXE.
$runtimeFiles = @(
    'CAM_Agent_Branding.html',
    'CAM_Agent_BatchEngine_V5_1_38.ps1',
    'CAM_Agent_Icon.ico',
    'CAM_Agent_Icon.png',
    'CAM_Agent_TimeReport.cs',
    'CAM_Agent_TimesImportWorker_V5_1_31.ps1',
    'CAM_Agent_TimesReport_Template.xlsx',
    'CAM_Agent_Times_Active.csv',
    'CAM_Agent_Times_Base.csv',
    'CAM_Agent_Times_DB.csv',
    'HOP_to_PGMX_2_9.exe',
    'PWX500_CsvWorker_V3.ps1',
    'PWX500_HistoryWorker_V3.ps1',
    'PWX500_HopWorker_V3.ps1',
    'PWX500_IndexWorker_V3.ps1',
    'PWX500_PgmxWorker_V3.ps1',
    'PWX500_ProductCategories.csv',
    'PWX500_SearchWorker_V3.ps1'
)

Copy-Item -LiteralPath $ExePath -Destination (Join-Path $PackageDir 'CAM_Agent_V5_1_38.exe') -Force
foreach ($name in $runtimeFiles) {
    $from = Join-Path $SrcDir $name
    if (-not (Test-Path -LiteralPath $from)) { throw "Missing runtime dependency: $name" }
    Copy-Item -LiteralPath $from -Destination (Join-Path $PackageDir $name) -Force
}

$runtimeReadme = @'
CAM Agent V5.1.38 - COMPILED EXE

1. Keep the complete folder together.
2. Start CAM_Agent_V5_1_38.exe directly.
3. Do NOT move only the EXE away from the worker scripts, databases and HOP_to_PGMX_2_9.exe.
4. The folder must be writable because the agent creates local settings, indexes, logs and status files beside the EXE.
5. Use "Основна папка..." in the operator interface to choose where daily dd.MM folders are created.
6. If no daily root is selected, the existing Desktop\FOUND\dd.MM behavior remains the default.

IMPORTANT:
The main GUI is compiled into the EXE. Several dedicated worker scripts intentionally remain as external .ps1 runtime components because V5.1.35/V5.1.38 was validated with that architecture. They are launched by the compiled agent when needed. This preserves the already-tested processing logic instead of rewriting it during the EXE conversion.
'@
Set-Content -LiteralPath (Join-Path $PackageDir 'README_COMPILED_EXE_BG.txt') -Value $runtimeReadme -Encoding UTF8

# Hash manifest for the finished runtime folder.
$hashLines = Get-ChildItem -LiteralPath $PackageDir -File | Sort-Object Name | ForEach-Object {
    $h = Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256
    '{0}  {1}' -f $h.Hash.ToLowerInvariant(), $_.Name
}
Set-Content -LiteralPath (Join-Path $PackageDir 'SHA256SUMS.txt') -Value $hashLines -Encoding ASCII

Compress-Archive -Path (Join-Path $PackageDir '*') -DestinationPath $ZipPath -CompressionLevel Optimal -Force

$sourceHash = (Get-FileHash -LiteralPath $MainSource -Algorithm SHA256).Hash.ToLowerInvariant()
$exeHash = (Get-FileHash -LiteralPath $ExePath -Algorithm SHA256).Hash.ToLowerInvariant()
$zipHash = (Get-FileHash -LiteralPath $ZipPath -Algorithm SHA256).Hash.ToLowerInvariant()
$info = @(
    'CAM Agent V5.1.38 GitHub build',
    "Built: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')",
    "Windows PowerShell: $($PSVersionTable.PSVersion)",
    "PS2EXE: $requiredPs2Exe",
    "Source SHA256: $sourceHash",
    "EXE SHA256: $exeHash",
    "Package SHA256: $zipHash"
)
Set-Content -LiteralPath (Join-Path $DistDir 'BUILD_INFO.txt') -Value $info -Encoding UTF8

Write-Host ''
Write-Host 'BUILD OK' -ForegroundColor Green
Write-Host "EXE: $ExePath"
Write-Host "ZIP: $ZipPath"
