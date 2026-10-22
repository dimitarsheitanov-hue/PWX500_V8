# Windows PowerShell 5.1 regression: three mapped drives share ONLY the approved
# 1_Izpolzvani subtrees, without adding duplicate scan roots or weakening other guards.
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$repo = Split-Path -Parent $PSScriptRoot
$src = Join-Path $repo 'src'

function Get-FunctionText([string]$file,[string]$name) {
    $tokens=$null;$errs=$null
    $ast=[System.Management.Automation.Language.Parser]::ParseFile((Join-Path $src $file),[ref]$tokens,[ref]$errs)
    if($errs -and $errs.Count){
        $errorText = (@($errs | ForEach-Object { $_.Message }) -join '; ')
        throw ('Parse failure: '+$file+' : '+$errorText)
    }
    $f=@($ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)|Where-Object{$_.Name -eq $name})
    if($f.Count -ne 1){throw ('Unexpected function count: '+$file+' '+$name+' = '+$f.Count)}
    return $f[0].Extent.Text
}

$roots=@('Y:\1_Izpolzvani\01_KM\_0_10_F', 'Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\08_Tapicirana_mebel')
$approved=@(
    'Y:\1_Izpolzvani\01_KM\_0_10_F\05_Koridor\Monaco F_Mirror',
    'V:\1_Izpolzvani\01_KM\_0_10_F\05_Koridor\Monaco F_Mirror',
    'Z:\1_Izpolzvani\01_KM\_0_10_F\05_Koridor\Monaco F_Mirror',
    'v:/1_Izpolzvani/02_TM/03_15mm_kant_1mm/08_Tapicirana_mebel/Product',
    'z:\1_IZPOLZVANI\01_KM\_0_10_F'
)
$rejected=@(
    'V:\1_Izpolzvani',
    'Z:\1_Izpolzvani\01_KM\Other',
    'V:\1_Izpolzvani\01_KM\_0_10_Fake',
    'V:\1_Izpolzvani\01_KM\_0_10_F\..\Other',
    'W:\1_Izpolzvani\01_KM\_0_10_F\Product',
    'V:\1_IzpolzvaniExtra\01_KM\_0_10_F\Product',
    'V:\Other\01_KM\_0_10_F\Product',
    'C:\1_Izpolzvani\01_KM\_0_10_F\Product'
)

# Main GUI whitelist (default union with user's explicit custom roots).
& {
    foreach($fn in @('Normalize-PathForCompare','Normalize-AllowedRootAliasPath','Test-IsUnderAllowedRoot')) {
        Invoke-Expression (Get-FunctionText 'PWX500_Agent_V3.ps1' $fn)
    }
    $script:DefaultAllowedFolders=@($roots)
    $script:AllowedFolders=New-Object System.Collections.Generic.List[string]
    $script:AllowedFolders.Add('Y:\1_Izpolzvani\02_TM\Additional_Approved')
    foreach($p in $approved){ if(-not (Test-IsUnderAllowedRoot -Path $p)){throw ('MAIN rejects allowed alias: '+$p)} }
    foreach($p in $rejected){ if(Test-IsUnderAllowedRoot -Path $p){throw ('MAIN permits outside path: '+$p)} }
    foreach($d in @('Y','V','Z')) {
        $p=('{0}:\1_Izpolzvani\02_TM\Additional_Approved\Product' -f $d)
        if(-not (Test-IsUnderAllowedRoot -Path $p)){throw ('MAIN rejects custom root alias: '+$p)}
    }
}

# PGMX worker must trust exactly the same approved paths from the Version 4
# settings snapshot, while using the original V:/Z: physical path for file IO.
& {
    foreach($fn in @('S','Normalize-Path','Normalize-AllowedRootAliasPath','Test-IsUnderAllowedRoot','Get-RelativeDepthUnderAllowedRoot')) {
        Invoke-Expression (Get-FunctionText 'PWX500_PgmxWorker_V3.ps1' $fn)
    }
    foreach($p in $approved){if(-not (Test-IsUnderAllowedRoot -Path $p -Roots $roots)){throw ('PGMX rejects alias: '+$p)}}
    foreach($p in $rejected){if(Test-IsUnderAllowedRoot -Path $p -Roots $roots){throw ('PGMX permits outside path: '+$p)}}
    # Whatever the original Y: family-folder classification is, V: and Z:
    # must produce exactly the same result. Do not change legacy behavior here.
    $familyPath='Y:\1_Izpolzvani\01_KM\_0_10_F\Category\Family\Child'
    $expectedDepth=Get-RelativeDepthUnderAllowedRoot -Path $familyPath -Roots $roots
    foreach($drive in @('V','Z')){
        $otherPath=('{0}:\1_Izpolzvani\01_KM\_0_10_F\Category\Family\Child' -f $drive)
        $depth=Get-RelativeDepthUnderAllowedRoot -Path $otherPath -Roots $roots
        if($depth -ne $expectedDepth){throw ('PGMX family depth differs for '+$drive+': '+$depth+' vs '+$expectedDepth)}
    }
}

# Historical migration can safely read the same catalog-selected folder path.
& {
    foreach($fn in @('Normalize-Path','Normalize-AllowedRootAliasPath','Is-UnderAllowedRoot')) {
        Invoke-Expression (Get-FunctionText 'PWX500_HistoryWorker_V3.ps1' $fn)
    }
    foreach($p in $approved){if(-not (Is-UnderAllowedRoot -path $p -roots $roots)){throw ('HISTORY rejects alias: '+$p)}}
    foreach($p in $rejected){if(Is-UnderAllowedRoot -path $p -roots $roots){throw ('HISTORY permits outside path: '+$p)}}
}

# Don't expand the eight Y: roots to 24 scan roots: no redundant network scans.
$main = Get-Content -LiteralPath (Join-Path $src 'PWX500_Agent_V3.ps1') -Raw -Encoding UTF8
if($main.Contains("'V:\1_Izpolzvani\01_KM\_0_10_F',") -or $main.Contains("'Z:\1_Izpolzvani\01_KM\_0_10_F',")){
    throw 'Unexpected physical V:/Z: root added to the scan whitelist.'
}
Write-Host 'PASS V5.1.38 FIX5: Y/V/Z aliases work in MAIN, PGMX and HISTORY; outside paths remain blocked; no extra network scans.' -ForegroundColor Green
