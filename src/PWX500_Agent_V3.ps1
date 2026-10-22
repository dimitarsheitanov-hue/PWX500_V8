# CAM Agent V5.1.38 - PWX 500 | P800 | FTT - safe sorting + 5.44 median fallback + speed; core PWX workflow V0.2.21
# SAFETY:
# - Reads ONLY .csv files from the fixed PWX500 input folder.
# - Folder indexing scans DIRECTORY NAMES/PATHS ONLY under explicit whitelist roots.
# - Does NOT open, parse, preview, index or launch Word files.
# - May read ONLY explicitly allowed .hop files during HOP verification, and only as text.
# - Historical migration checks only exact *.pgmx paths in candidate 2_PWX_500 folders; it does not open PGMX content.
# - Check mode never modifies the server. Execute mode may create 2_PWX_500, send all CSV-linked missing HOPs for the confirmed product to the supplied V2.9 converter in one batch, and copy successful outputs to ProgramCodeF. Existing files are never overwritten.
# - Existing server source/target files are never overwritten, edited, renamed or deleted.
# - Local settings/index are stored as .csv beside the agent.
# - Optional control export writes a NEW .csv report only when the user explicitly clicks Export.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
# Training dependencies are loaded only when a training session is started (Step 4 / Step 5).
# This keeps the normal PWX500 Agent startup independent from UI Automation/user32 initialization.
try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch {}

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:AllowedReadExtensions = @('.csv')
$script:BlockedWordExtensions = @('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.rtf')
$script:LastFileSummary = $null
$script:LastAllRows = @()
$script:LastWarnings = @()
$script:LastFolder = ''
$script:DefaultCsvFolder = 'Y:\30_Biesse_edging\scm'
$script:SettingsCsv = Join-Path $PSScriptRoot 'PWX500_Agent_AllowedFolders.csv'
$script:IndexCsv = Join-Path $PSScriptRoot 'PWX500_Agent_ProductFolderCatalog_V3.csv'
$script:PreviousIndexCsv = Join-Path $PSScriptRoot 'PWX500_Agent_ProductFolderCatalog_V0219.csv'
$script:DefaultAllowedFolders = @(
    'Y:\1_Izpolzvani\01_KM\_0_10_F',
    'Y:\1_Izpolzvani\01_KM\_06_Pac_s12_B#',
    'Y:\1_Izpolzvani\01_KM\0_11_F#',
    'Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\08_Tapicirana_mebel',
    'Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\09_Tapicirana_mebel_#',
    'Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\10_Tapicirana_mebel_B#',
    'Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\11_Tapicirana_mebel_C#',
    'Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\12_Tapicirana_mebel_F'
)
$script:AllowedFolders = New-Object System.Collections.Generic.List[string]
$script:FolderIndex = @()
$script:FolderNameMap = @{}
$script:FolderTokenMap = @{}
$script:SettingsVersion = '4'
$script:IndexWorkerProcess = $null
$script:IndexWorkerTimer = $null
$script:IndexTempCsv = Join-Path $PSScriptRoot 'PWX500_Agent_ProductFolderCatalog_V3.tmp.csv'
$script:IndexProgressFile = Join-Path $PSScriptRoot 'PWX500_Agent_IndexProgress.txt'
$script:IndexWorkerScript = Join-Path $PSScriptRoot 'PWX500_IndexWorker_V3.ps1'
$script:SearchWorkerScript = Join-Path $PSScriptRoot 'PWX500_SearchWorker_V3.ps1'
$script:SearchWorkerProcess = $null
$script:SearchWorkerTimer = $null
$script:SearchProductsCsv = Join-Path $PSScriptRoot 'PWX500_Agent_SearchProducts.tmp.csv'
$script:SearchResultsCsv = Join-Path $PSScriptRoot 'PWX500_Agent_SearchResults.tmp.csv'
$script:SearchProgressFile = Join-Path $PSScriptRoot 'PWX500_Agent_SearchProgress.txt'
$script:HopWorkerScript = Join-Path $PSScriptRoot 'PWX500_HopWorker_V3.ps1'
$script:HopWorkerProcess = $null
$script:HopWorkerTimer = $null
$script:HopResultsCsv = Join-Path $PSScriptRoot 'PWX500_Agent_HopResults.tmp.csv'
$script:HopProgressFile = Join-Path $PSScriptRoot 'PWX500_Agent_HopProgress.txt'
$script:HopResultObjects = @()
$script:ConfirmDisplayToPath = @{}
$script:FolderIndexCount = 0
$script:ProductFolderMapCsv = Join-Path $PSScriptRoot 'PWX500_Agent_ProductFolders.csv'
$script:ProductCategoriesCsv = Join-Path $PSScriptRoot 'PWX500_ProductCategories.csv'
$script:ProductCategoryRules = @()
$script:ProductFolderMap = @{}
$script:SearchResultObjects = @()
$script:ConfirmationProducts = @()
$script:ConfirmationIndex = -1
$script:CsvWorkerScript = Join-Path $PSScriptRoot 'PWX500_CsvWorker_V3.ps1'
$script:CsvWorkerProcess = $null
$script:CsvWorkerTimer = $null
$script:CsvRowsOut = Join-Path $PSScriptRoot 'PWX500_Agent_CsvRows.tmp.csv'
$script:CsvSummaryOut = Join-Path $PSScriptRoot 'PWX500_Agent_CsvSummary.tmp.csv'
$script:CsvProgressFile = Join-Path $PSScriptRoot 'PWX500_Agent_CsvProgress.txt'
$script:PgmxWorkerScript = Join-Path $PSScriptRoot 'PWX500_PgmxWorker_V3.ps1'
$script:PgmxWorkerProcess = $null
$script:PgmxWorkerTimer = $null
$script:PgmxResultsCsv = Join-Path $PSScriptRoot 'PWX500_Agent_PgmxResults.tmp.csv'
$script:PgmxProgressFile = Join-Path $PSScriptRoot 'PWX500_Agent_PgmxProgress.txt'
$script:PgmxErrorFile = Join-Path $PSScriptRoot 'PWX500_Agent_PgmxError.txt'
$script:PgmxDoneFile = Join-Path $PSScriptRoot 'PWX500_Agent_PgmxDone.txt'
$script:ConverterExe = Join-Path $PSScriptRoot 'HOP_to_PGMX_2_9.exe'
$script:ConversionLogCsv = Join-Path $PSScriptRoot 'PWX500_Agent_ConversionLog.csv'
$script:PgmxResultObjects = @()
$script:PgmxLastMode = ''
# V5.1.2: exact PGMX files physically created/copy-produced by the current Execute run.
# V5.1.2 daily-folder rule: an existing Desktop\FOUND\dd.MM folder is valid. New PGMX are merged into it and the MIXX uses ALL PGMX in that folder. Existing daily MIXX is automatically rebuilt.
$script:CreatedPgmxThisRun = @()
$script:AutoMixChainRunning = $false
$script:HistoryWorkerScript = Join-Path $PSScriptRoot 'PWX500_HistoryWorker_V3.ps1'
$script:HistoryWorkerProcess = $null
$script:HistoryWorkerTimer = $null
$script:HistoryCodesCsv = Join-Path $PSScriptRoot 'PWX500_Agent_HistoricalProgramCodes.csv'
$script:HistoryProductsCsv = Join-Path $PSScriptRoot 'PWX500_Agent_HistoricalProducts.csv'
$script:HistoryProgressFile = Join-Path $PSScriptRoot 'PWX500_Agent_HistoryProgress.txt'
$script:HistoryDoneFile = Join-Path $PSScriptRoot 'PWX500_Agent_HistoryDone.txt'
$script:HistoryErrorFile = Join-Path $PSScriptRoot 'PWX500_Agent_HistoryError.txt'
$script:HistoryMarkerFile = Join-Path $PSScriptRoot 'PWX500_Agent_HistoryMigration.done.txt'
$script:HistoryResultObjects = @()
$script:HistoryLastWorkbook = ''

# V5.1.31 operator workflow / reports / daily catalog / CSV registry / local schedule + simplified main operator window.
$script:OperatorAutoFlow = $false
$script:OperatorRunHopAfterSearch = $false
$script:OperatorRunExecuteAfterHop = $false
$script:OperatorScheduleFile = Join-Path $PSScriptRoot 'PWX500_Agent_Schedule.ini'
$script:OperatorScheduleLastRunDate = ''
$script:OperatorScheduleMode = 'PGMX_ONLY'
$script:OperatorExecutionMode = 'FULL'
$script:AppRoot = $PSScriptRoot
$script:OperatorLastSkippedReport = ''
$script:OperatorLastDetailErrorReport = ''
# V5.1.31: daily folder catalog + persistent CSV processing registry.
$script:CsvRegistryCsv = Join-Path $PSScriptRoot 'PWX500_Agent_CsvRegistry.csv'
$script:CsvSkipListCsv = Join-Path $PSScriptRoot 'PWX500_Agent_CsvSkip.tmp.csv'
$script:CsvCurrentBatch = @()
$script:OperatorAfterIndexAction = ''
# V5.1.31: local stop/resume checkpoints + persistent process/migration status.
$script:OperatorRunStateFile = Join-Path $PSScriptRoot 'PWX500_Agent_RunState.json'
$script:OperatorProcessHistoryCsv = Join-Path $PSScriptRoot 'PWX500_Agent_ProcessHistory.csv'
$script:MigrationStatusFile = Join-Path $PSScriptRoot 'PWX500_Agent_MigrationStatus.json'
$script:AppIconPng = Join-Path $PSScriptRoot 'CAM_Agent_Icon.png'
$script:AppIconIco = Join-Path $PSScriptRoot 'CAM_Agent_Icon.ico'
$script:DailyRootSettingsFile = Join-Path $PSScriptRoot 'PWX500_Agent_DailyRoot.txt'
$script:OperatorStopRequested = $false
$script:OperatorCurrentStage = 'IDLE'
$script:OperatorLastSafeStage = 'IDLE'
$script:OperatorResumePromptShown = $false
# V5.1.31: persistent detail-time database and reports.
$script:TimesBaseCsv = Join-Path $PSScriptRoot 'CAM_Agent_Times_Base.csv'
$script:TimesDbCsv = Join-Path $PSScriptRoot 'CAM_Agent_Times_DB.csv'
$script:TimesActiveCsv = Join-Path $PSScriptRoot 'CAM_Agent_Times_Active.csv'
$script:TimesMeasurementsCsv = Join-Path $PSScriptRoot 'CAM_Agent_Times_History.csv'
$script:TimesImportRegistryCsv = Join-Path $PSScriptRoot 'CAM_Agent_Times_ImportedFiles.csv'
$script:TimeDbMap = @{}
$script:ActiveTimeMap = @{}
$script:TimeDbLoaded = $false
$script:LastTimeMeasurementCandidates = @()
$script:LastTimeImportStats = $null
$script:TimeImportWorkerScript = Join-Path $PSScriptRoot 'CAM_Agent_TimesImportWorker_V5_1_31.ps1'
$script:TimeImportProcess = $null
$script:TimeImportTimer = $null
$script:TimeImportProgressPath = ''
$script:TimeImportStopPath = ''
$script:TimeImportResultPath = ''
$script:TimeImportManifestPath = ''
$script:TimeImportJobDir = ''
$script:TimeImportRunning = $false
$script:TimeImportSynchronousRunning = $false
$script:TimeImportSynchronousStopPath = ''
$script:TimeImportLastJobFile = Join-Path $PSScriptRoot 'CAM_Agent_TimesImport_LastJob.json'
$script:MainStopTimesButton = $null
$script:MedianFallbackSec = 5.44
$script:MedianFallbackSource = 'MEDIAN_FALLBACK_5.44'
$script:LastMedianFallbackCodes = @()
$script:LastProblemPgmxMoved = @()
$script:LastNoTimePgmxMoved = @()

# Training Mode state - local only, beside the agent.
$script:TrainingRoot = Join-Path $PSScriptRoot 'TrainingSessions'
$script:TrainingExportRoot = Join-Path $PSScriptRoot 'TrainingExports'
$script:TrainingActive = $false
$script:TrainingSessionPath = ''
$script:TrainingSessionName = ''
$script:TrainingSessionStart = $null
$script:TrainingSeq = 0
$script:TrainingEvents = New-Object System.Collections.Generic.List[object]
$script:TrainingLastWindowHandle = [IntPtr]::Zero
$script:TrainingLastMouseLeft = $false
$script:TrainingLastMouseRight = $false
$script:TrainingLastMouseMiddle = $false
$script:TrainingKeyStates = @{}
$script:TrainingLatestZip = ''
$script:TrainingScreenshotEnabled = $true
$script:TrainingOwnProcessId = $PID
$script:TrainingNativeReady = $false
$script:TrainingNativeError = ''

function Safe-String {
    param($Value)
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Safe-Trim {
    param($Value)
    $s = Safe-String $Value
    return $s.Trim()
}

function Get-TextEncoding {
    param([Parameter(Mandatory=$true)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return (New-Object System.Text.UTF8Encoding($true))
    }

    try {
        $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$strictUtf8.GetString($bytes)
        return $strictUtf8
    }
    catch {
        return [System.Text.Encoding]::GetEncoding(1251)
    }
}

function Get-CsvDelimiter {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][System.Text.Encoding]$Encoding
    )

    $reader = New-Object System.IO.StreamReader -ArgumentList @($Path, $Encoding, $true)
    try {
        $line = $reader.ReadLine()
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
    }

    if ($null -eq $line) { return ';' }

    $candidates = @(';', ',', "`t")
    $best = ';'
    $bestCount = -1
    foreach ($d in $candidates) {
        $count = 0
        foreach ($ch in $line.ToCharArray()) {
            if ($ch -eq $d) { $count++ }
        }
        if ($count -gt $bestCount) {
            $best = $d
            $bestCount = $count
        }
    }
    return $best
}

function Read-SafeCsv {
    param([Parameter(Mandatory=$true)][string]$Path)

    $ext = [System.IO.Path]::GetExtension($Path)
    if ($null -eq $ext) { $ext = '' }
    $ext = $ext.ToLowerInvariant()

    # ABSOLUTE BLOCK before content access.
    if ($script:BlockedWordExtensions -contains $ext) {
        throw "BLOCKED FILE TYPE: $ext"
    }
    if (-not ($script:AllowedReadExtensions -contains $ext)) {
        throw "File type is not allowed for reading in V3.0: $ext"
    }

    $encoding = Get-TextEncoding -Path $Path
    $delimiter = Get-CsvDelimiter -Path $Path -Encoding $encoding

    $parser = $null
    try {
        $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path, $encoding, $true)
        if ($null -eq $parser) { throw 'TextFieldParser could not be created.' }

        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters([string[]]@([string]$delimiter))
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false

        if ($parser.EndOfData) { return @() }
        $headers = @($parser.ReadFields())
        if ($null -eq $headers -or $headers.Count -eq 0) { return @() }

        for ($h = 0; $h -lt $headers.Count; $h++) {
            $headers[$h] = (Safe-Trim $headers[$h]).Trim([char]0xFEFF)
        }

        $rows = New-Object System.Collections.Generic.List[object]
        while (-not $parser.EndOfData) {
            try {
                $fields = @($parser.ReadFields())
            }
            catch {
                throw "CSV parse error in '$([System.IO.Path]::GetFileName($Path))': $($_.Exception.Message)"
            }

            if ($null -eq $fields) { continue }
            $obj = [ordered]@{}
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $header = Safe-Trim $headers[$i]
                if ([string]::IsNullOrWhiteSpace($header)) { $header = "Column_$i" }
                $value = ''
                if ($i -lt $fields.Count -and $null -ne $fields[$i]) { $value = [string]$fields[$i] }
                $obj[$header] = $value
            }
            $rows.Add([pscustomobject]$obj)
        }
        return @($rows.ToArray())
    }
    finally {
        if ($null -ne $parser) {
            try { $parser.Close() } catch {}
            try { $parser.Dispose() } catch {}
        }
    }
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)][string]$Name
    )
    if ($null -eq $Object) { return '' }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return '' }
    return (Safe-String $p.Value)
}

function Parse-ItemDescription {
    param($Description)

    $result = [ordered]@{
        Product = ''
        DetailName = ''
        DetailNo = ''
        Parsed = $false
    }

    $desc = Safe-Trim $Description
    if ([string]::IsNullOrWhiteSpace($desc)) { return [pscustomobject]$result }

    $parts = @([regex]::Split($desc, '\s+-\s+'))
    if ($parts.Count -ge 3) {
        $last = Safe-Trim $parts[$parts.Count - 1]
        $n = 0
        if ([int]::TryParse($last, [ref]$n)) {
            $result.DetailNo = [string]$n
            $result.DetailName = Safe-Trim $parts[$parts.Count - 2]
            $productParts = New-Object System.Collections.Generic.List[string]
            for ($i = 0; $i -le $parts.Count - 3; $i++) {
                $pp = Safe-Trim $parts[$i]
                if (-not [string]::IsNullOrWhiteSpace($pp)) { $productParts.Add($pp) }
            }
            $result.Product = (($productParts.ToArray()) -join ' - ').Trim()
            $result.Parsed = -not [string]::IsNullOrWhiteSpace($result.Product)
        }
    }

    return [pscustomobject]$result
}

function New-DataTable {
    param([string[]]$Columns)

    $table = New-Object System.Data.DataTable
    foreach ($c in $Columns) { [void]$table.Columns.Add([string]$c) }
    Write-Output -NoEnumerate $table
}


function Save-AllowedFolders {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Version;Path')
    # The PGMX worker reads this snapshot, while Main trusts default+custom roots.
    # Make both use exactly the same effective whitelist (no bypass of real outside roots).
    $effective=@($script:DefaultAllowedFolders)+@($script:AllowedFolders.ToArray())
    $seen=@{}
    foreach ($p in @($effective | Sort-Object)) {
        $key=Normalize-PathForCompare $p
        if(-not $key -or $seen.ContainsKey($key)){continue}
        $seen[$key]=$true
        $lines.Add(((Convert-ToCsvField $script:SettingsVersion), (Convert-ToCsvField $p)) -join ';')
    }
    # Keep the timestamp meaningful: an unchanged list must not invalidate the catalog.
    if ([System.IO.File]::Exists($script:SettingsCsv)) {
        $old = [System.IO.File]::ReadAllLines($script:SettingsCsv, [System.Text.Encoding]::UTF8)
        if (($old -join "`n") -ceq ($lines.ToArray() -join "`n")) { return }
    }
    [System.IO.File]::WriteAllLines($script:SettingsCsv, $lines.ToArray(), (New-Object System.Text.UTF8Encoding($true)))
}

function Set-DefaultAllowedFolders {
    $script:AllowedFolders.Clear()
    foreach ($p in $script:DefaultAllowedFolders) {
        if (-not $script:AllowedFolders.Contains($p)) { $script:AllowedFolders.Add($p) }
    }
}

function Load-AllowedFolders {
    $script:AllowedFolders.Clear()
    $loaded = $false
    if ([System.IO.File]::Exists($script:SettingsCsv)) {
        try {
            $rows = @(Read-SafeCsv -Path $script:SettingsCsv)
            $hasCurrentVersion = $false
            foreach ($r in $rows) {
                $v = Safe-Trim (Get-PropertyValue -Object $r -Name 'Version')
                if ($v -eq $script:SettingsVersion) { $hasCurrentVersion = $true; break }
            }
            # V0.2.0 had an unversioned settings file. Do not inherit that legacy list,
            # because it may contain only part of the eight agreed roots.
            if ($hasCurrentVersion) {
                foreach ($r in $rows) {
                    $v = Safe-Trim (Get-PropertyValue -Object $r -Name 'Version')
                    if ($v -ne $script:SettingsVersion) { continue }
                    $p = Safe-Trim (Get-PropertyValue -Object $r -Name 'Path')
                    if (-not [string]::IsNullOrWhiteSpace($p) -and -not $script:AllowedFolders.Contains($p)) {
                        $script:AllowedFolders.Add($p)
                        $loaded = $true
                    }
                }
            }
        }
        catch { $loaded = $false }
    }
    if (-not $loaded) {
        Set-DefaultAllowedFolders
        try { Save-AllowedFolders } catch {}
    }
    # Add the requested root to existing Version 4 settings without resetting custom roots.
    $requiredRoot = 'Y:\1_Izpolzvani\02_TM\03_15mm_kant_1mm\08_Tapicirana_mebel'
    $requiredKey = Normalize-PathForCompare $requiredRoot
    $found = $false
    foreach ($existing in $script:AllowedFolders.ToArray()) {
        if ((Normalize-PathForCompare $existing) -eq $requiredKey) { $found = $true; break }
    }
    if (-not $found) {
        $script:AllowedFolders.Add($requiredRoot)
        Save-AllowedFolders
    }
}

function Refresh-AllowedFoldersList {
    if ($null -eq $allowedList) { return }
    $allowedList.Items.Clear()
    foreach ($p in @($script:AllowedFolders.ToArray() | Sort-Object)) { [void]$allowedList.Items.Add($p) }
    if ($null -ne $allowedCountLabel) { $allowedCountLabel.Text = "Разрешени папки: $($script:AllowedFolders.Count)" }
}


function Normalize-PathForCompare {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $x = $Path.Trim().Trim('"').Replace('/', '\')
    $x = $x.TrimEnd('\')
    return $x.ToLowerInvariant()
}

# Compare mapped drives Y:, V: and Z: as equivalent ONLY below their
# shared 1_Izpolzvani root. Keep the physical path unchanged for all IO.
# Authorization still requires a match to a specifically approved subtree.
function Normalize-AllowedRootAliasPath {
    param([string]$Path)
    $p = Normalize-PathForCompare $Path
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    # Do not let a lexical ../ segment escape an approved child folder.
    if ([regex]::IsMatch($p, '(^|\\)\.\.(\\|$)')) { return '' }
    $driveRoot = [regex]::Match($p, '^[yvz]:\\1_izpolzvani(?=\\|$)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($driveRoot.Success) {
        return 'y:\1_izpolzvani' + $p.Substring($driveRoot.Length)
    }
    return $p
}

function Test-IsUnderAllowedRoot {
    param([string]$Path)
    $p = Normalize-AllowedRootAliasPath $Path
    if ([string]::IsNullOrWhiteSpace($p)) { return $false }

    # V5.1.2: the eight agreed safety roots remain valid even if a stale/partial
    # PWX500_Agent_AllowedFolders.csv exists from an older test folder.
    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($root in @($script:DefaultAllowedFolders)) {
        if (-not [string]::IsNullOrWhiteSpace($root) -and -not $roots.Contains($root)) { $roots.Add($root) }
    }
    foreach ($root in @($script:AllowedFolders.ToArray())) {
        if (-not [string]::IsNullOrWhiteSpace($root) -and -not $roots.Contains($root)) { $roots.Add($root) }
    }

    foreach ($root in @($roots.ToArray())) {
        $r = Normalize-AllowedRootAliasPath $root
        if ([string]::IsNullOrWhiteSpace($r)) { continue }
        if ($p -eq $r) { return $true }
        $prefix = $r + '\'
        if ($p.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Get-DefaultProductCategoryRows {
    return @(
        [pscustomobject]@{Phrase='bedroom and bed';Targets='12_Spalni_i_Legla|02_TM';Priority='220';Active='1'},
        [pscustomobject]@{Phrase='upgrade wardrobe';Targets='03_Garderobi';Priority='220';Active='1'},
        [pscustomobject]@{Phrase='night stand';Targets='08_Noshtni_shkafove';Priority='220';Active='1'},
        [pscustomobject]@{Phrase='hallway set';Targets='05_Koridor';Priority='220';Active='1'},
        [pscustomobject]@{Phrase='shoes drawer';Targets='05_Koridor';Priority='220';Active='1'},
        [pscustomobject]@{Phrase='coffee table';Targets='07_Masi';Priority='220';Active='1'},
        [pscustomobject]@{Phrase='dining table';Targets='07_Masi';Priority='220';Active='1'},
        [pscustomobject]@{Phrase='wardrobe';Targets='03_Garderobi';Priority='160';Active='1'},
        [pscustomobject]@{Phrase='kitchen';Targets='06_Kuhni';Priority='160';Active='1'},
        [pscustomobject]@{Phrase='dresser';Targets='11_Skrinove';Priority='160';Active='1'},
        [pscustomobject]@{Phrase='section';Targets='10_Sekcii';Priority='160';Active='1'},
        [pscustomobject]@{Phrase='desk';Targets='02_Bura';Priority='160';Active='1'},
        [pscustomobject]@{Phrase='bed';Targets='12_Spalni_i_Legla';Priority='150';Active='1'},
        [pscustomobject]@{Phrase='stand';Targets='10_Sekcii';Priority='80';Active='1'}
    )
}

function Save-ProductCategoryRules {
    param([object[]]$Rows)
    if ($null -eq $Rows) { $Rows = @() }
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Phrase;Targets;Priority;Active')
    foreach($r in @($Rows)) {
        $phrase = Safe-Trim (Get-PropertyValue -Object $r -Name 'Phrase')
        $targets = Safe-Trim (Get-PropertyValue -Object $r -Name 'Targets')
        $priority = Safe-Trim (Get-PropertyValue -Object $r -Name 'Priority')
        $active = Safe-Trim (Get-PropertyValue -Object $r -Name 'Active')
        if ([string]::IsNullOrWhiteSpace($phrase) -or [string]::IsNullOrWhiteSpace($targets)) { continue }
        $prioInt=100; if(-not [int]::TryParse($priority,[ref]$prioInt)){$prioInt=100}
        if([string]::IsNullOrWhiteSpace($active)){$active='1'}
        $lines.Add(((Convert-ToCsvField $phrase),(Convert-ToCsvField $targets),(Convert-ToCsvField ([string]$prioInt)),(Convert-ToCsvField $active)) -join ';')
    }
    [System.IO.File]::WriteAllLines($script:ProductCategoriesCsv,$lines.ToArray(),(New-Object System.Text.UTF8Encoding($true)))
}

function Ensure-ProductCategoryRulesFile {
    if (-not [System.IO.File]::Exists($script:ProductCategoriesCsv)) {
        Save-ProductCategoryRules -Rows @(Get-DefaultProductCategoryRows)
    }
}

function Load-ProductCategoryRules {
    Ensure-ProductCategoryRulesFile
    $script:ProductCategoryRules=@()
    try {
        $rows=@(Read-SafeCsv -Path $script:ProductCategoriesCsv)
        $list=New-Object System.Collections.Generic.List[object]
        foreach($r in $rows){
            $phrase=Safe-Trim (Get-PropertyValue -Object $r -Name 'Phrase')
            $targets=Safe-Trim (Get-PropertyValue -Object $r -Name 'Targets')
            $active=Safe-Trim (Get-PropertyValue -Object $r -Name 'Active')
            $prio=100; [void][int]::TryParse((Safe-Trim (Get-PropertyValue -Object $r -Name 'Priority')),[ref]$prio)
            if([string]::IsNullOrWhiteSpace($phrase) -or [string]::IsNullOrWhiteSpace($targets)){continue}
            if($active -in @('0','false','False','FALSE','no','NO')){continue}
            $list.Add([pscustomobject]@{Phrase=$phrase;Targets=$targets;Priority=$prio;Active='1'})
        }
        $script:ProductCategoryRules=@($list.ToArray() | Sort-Object @{Expression='Priority';Descending=$true}, @{Expression={$_.Phrase.Length};Descending=$true})
    } catch { $script:ProductCategoryRules=@() }
}

function Refresh-ProductCategoryGrid {
    if($null -eq $categoryGrid){return}
    Ensure-ProductCategoryRulesFile
    $table=New-DataTable -Columns @('Phrase','Targets','Priority','Active')
    foreach($r in @(Read-SafeCsv -Path $script:ProductCategoriesCsv)){
        $dr=$table.NewRow();$dr['Phrase']=Safe-Trim (Get-PropertyValue $r 'Phrase');$dr['Targets']=Safe-Trim (Get-PropertyValue $r 'Targets');$dr['Priority']=Safe-Trim (Get-PropertyValue $r 'Priority');$dr['Active']=Safe-Trim (Get-PropertyValue $r 'Active');[void]$table.Rows.Add($dr)
    }
    $categoryGrid.DataSource=$table
}

function Save-ProductCategoryGrid {
    if($null -eq $categoryGrid -or $null -eq $categoryGrid.DataSource){return}
    try{$categoryGrid.EndEdit()}catch{}
    $rows=New-Object System.Collections.Generic.List[object]
    foreach($gr in $categoryGrid.Rows){
        if($gr.IsNewRow){continue}
        $phrase=Safe-Trim $gr.Cells['Phrase'].Value;$targets=Safe-Trim $gr.Cells['Targets'].Value
        if([string]::IsNullOrWhiteSpace($phrase) -or [string]::IsNullOrWhiteSpace($targets)){continue}
        $rows.Add([pscustomobject]@{Phrase=$phrase;Targets=$targets;Priority=(Safe-Trim $gr.Cells['Priority'].Value);Active=(Safe-Trim $gr.Cells['Active'].Value)})
    }
    Save-ProductCategoryRules -Rows @($rows.ToArray())
    Load-ProductCategoryRules
    Refresh-ProductCategoryGrid
}

function Test-IsTrustedProductMappingSource {
    param([string]$Source)
    $x=(Safe-Trim $Source).ToUpperInvariant()
    return ($x -in @('HISTORICAL_PGMX','HOP_AUTO','HOP_VERIFIED_MANUAL','HOP_VERIFIED_USER'))
}

function Load-ProductFolderMappings {
    $script:ProductFolderMap = @{}
    if (-not [System.IO.File]::Exists($script:ProductFolderMapCsv)) { return }
    try {
        $rows = @(Read-SafeCsv -Path $script:ProductFolderMapCsv)
        foreach ($r in $rows) {
            $product = Safe-Trim (Get-PropertyValue -Object $r -Name 'Product')
            $path = Safe-Trim (Get-PropertyValue -Object $r -Name 'Path')
            if ([string]::IsNullOrWhiteSpace($product) -or [string]::IsNullOrWhiteSpace($path)) { continue }
            $script:ProductFolderMap[$product] = [pscustomobject]@{
                Product = $product
                Path = $path
                ConfirmedAt = Safe-Trim (Get-PropertyValue -Object $r -Name 'ConfirmedAt')
                Source = Safe-Trim (Get-PropertyValue -Object $r -Name 'Source')
            }
        }
    }
    catch { $script:ProductFolderMap = @{} }
}

function Save-ProductFolderMappings {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Product;Path;ConfirmedAt;Source')
    foreach ($key in @($script:ProductFolderMap.Keys | Sort-Object)) {
        $x = $script:ProductFolderMap[$key]
        $lines.Add(((Convert-ToCsvField $x.Product),(Convert-ToCsvField $x.Path),(Convert-ToCsvField $x.ConfirmedAt),(Convert-ToCsvField $x.Source)) -join ';')
    }
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($script:ProductFolderMapCsv, $lines.ToArray(), $utf8Bom)
}

function Set-ProductFolderMapping {
    param([string]$Product,[string]$Path,[string]$Source)
    if ([string]::IsNullOrWhiteSpace($Product) -or [string]::IsNullOrWhiteSpace($Path)) { return }
    $script:ProductFolderMap[$Product] = [pscustomobject]@{
        Product = $Product
        Path = $Path
        ConfirmedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Source = $Source
    }
    Save-ProductFolderMappings
}

function Remove-ProductFolderMapping {
    param([string]$Product)
    if ($script:ProductFolderMap.ContainsKey($Product)) {
        $script:ProductFolderMap.Remove($Product)
        Save-ProductFolderMappings
    }
}

function Parse-CandidatePaths {
    param([string]$PathsText)
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($PathsText)) { return @() }
    foreach ($part in @($PathsText -split '~\|~')) {
        $p = Safe-Trim $part
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $score = ''
        $path = $p
        if ($p -match '^\[(\d+)\](.*)$') {
            $score = $matches[1]
            $path = $matches[2]
        }
        $out.Add([pscustomobject]@{ Score=$score; Path=(Safe-Trim $path) })
    }
    return @($out.ToArray())
}

function Get-SearchResultForProduct {
    param([string]$Product)
    foreach ($x in @($script:SearchResultObjects)) {
        if ($x.Product -eq $Product) { return $x }
    }
    return $null
}

function Refresh-ConfirmationQueue {
    if ($null -eq $confirmProductList) { return }
    $current = $null
    if ($confirmProductList.SelectedItem) { $current = [string]$confirmProductList.SelectedItem }
    $confirmProductList.Items.Clear()
    $queue = New-Object System.Collections.Generic.List[string]
    foreach ($x in @($script:SearchResultObjects | Sort-Object Product)) {
        $product = [string]$x.Product
        $mapped = $false
        if ($script:ProductFolderMap.ContainsKey($product)) {
            $mp = [string]$script:ProductFolderMap[$product].Path
            if ([System.IO.Directory]::Exists($mp)) { $mapped = $true }
        }
        if ($mapped) { continue }
        $cands = @(Parse-CandidatePaths $x.Paths)
        # Queue only cases that need a human decision now: zero or multiple candidates.
        if ($cands.Count -ne 1) {
            [void]$confirmProductList.Items.Add($product)
            $queue.Add($product)
        }
    }
    $script:ConfirmationProducts = @($queue.ToArray())
    if ($null -ne $confirmQueueLabel) {
        $confirmQueueLabel.Text = "За потвърждение: $($script:ConfirmationProducts.Count) изделия"
    }
    if ($confirmProductList.Items.Count -gt 0) {
        $idx = 0
        if (-not [string]::IsNullOrWhiteSpace($current)) {
            for ($i=0; $i -lt $confirmProductList.Items.Count; $i++) {
                if ([string]$confirmProductList.Items[$i] -eq $current) { $idx = $i; break }
            }
        }
        $confirmProductList.SelectedIndex = $idx
    }
    else {
        if ($null -ne $confirmProductLabel) { $confirmProductLabel.Text = 'Няма изделия, чакащи ръчно потвърждение.' }
        if ($null -ne $confirmCandidateList) { $confirmCandidateList.Items.Clear() }
    }
}

function Get-HopResultForCandidate {
    param([string]$Product,[string]$Path)
    $pp = Normalize-PathForCompare $Path
    foreach ($h in @($script:HopResultObjects)) {
        if ([string]$h.Product -eq $Product -and (Normalize-PathForCompare ([string]$h.CandidatePath)) -eq $pp) { return $h }
    }
    return $null
}

function Show-ConfirmationProduct {
    if ($null -eq $confirmProductList -or $null -eq $confirmCandidateList) { return }
    if ($null -eq $confirmProductList.SelectedItem) { return }
    $product = [string]$confirmProductList.SelectedItem
    $confirmProductLabel.Text = $product
    $confirmCandidateList.Items.Clear()
    $script:ConfirmDisplayToPath = @{}
    $x = Get-SearchResultForProduct -Product $product
    if ($null -eq $x) { return }
    foreach ($c in @(Parse-CandidatePaths $x.Paths)) {
        $hop = Get-HopResultForCandidate -Product $product -Path $c.Path
        $prefix = ''
        if ($null -ne $hop) {
            $prefix = "[HOP $($hop.MatchedFull)/$($hop.ExpectedDetails) | име $($hop.MatchedName) | размер $($hop.MatchedDimensions)] "
        }
        $scoreText = if ([string]::IsNullOrWhiteSpace($c.Score)) { '' } else { "[$($c.Score)] " }
        $label = $prefix + $scoreText + $c.Path
        $script:ConfirmDisplayToPath[$label] = $c.Path
        [void]$confirmCandidateList.Items.Add($label)
    }
    if ($confirmCandidateList.Items.Count -gt 0) { $confirmCandidateList.SelectedIndex = 0 }
    $confirmHintLabel.Text = if ($confirmCandidateList.Items.Count -eq 0) {
        'Няма предложена папка. Натисни „Посочи точна папка“.'
    } elseif ($script:HopResultObjects.Count -gt 0) {
        'HOP резултатът е показан пред всеки кандидат. Избери правилния или натисни „Посочи точна папка“.'
    } else {
        'Избери правилния кандидат и натисни „Потвърди избрания кандидат“, или първо пусни HOP проверката.'
    }
}

function Get-SelectedConfirmationCandidatePath {
    if ($null -eq $confirmCandidateList.SelectedItem) { return '' }
    $s = [string]$confirmCandidateList.SelectedItem
    if ($script:ConfirmDisplayToPath.ContainsKey($s)) { return [string]$script:ConfirmDisplayToPath[$s] }
    if ($s -match '^\[\d+\]\s+(.*)$') { return $matches[1].Trim() }
    return $s.Trim()
}

function Refresh-CandidateGridFromSearchResults {
    if ($null -eq $productCandidatesGrid) { return }
    $table = New-DataTable -Columns @('Изделие от CSV','Кандидати','Статус','Най-добра оценка','Пътища / Потвърдена папка')
    foreach ($x in @($script:SearchResultObjects | Sort-Object Product)) {
        $product = [string]$x.Product
        $count = [string]$x.Candidates
        $state = [string]$x.Status
        $best = [string]$x.BestScore
        $paths = ([string]$x.Paths).Replace('~|~',' | ')
        if ($script:ProductFolderMap.ContainsKey($product)) {
            $mp = [string]$script:ProductFolderMap[$product].Path
            if ([System.IO.Directory]::Exists($mp)) {
                $src=[string]$script:ProductFolderMap[$product].Source
                if($src -eq 'HISTORICAL_PGMX'){$state='ПОТВЪРДЕНО ОТ БЪРЗА МИГРАЦИЯ - БЕЗ HOP'}
                elseif(Test-IsTrustedProductMappingSource -Source $src){$state='ПОТВЪРДЕНО ПО HOP'}
                else{$state='РЪЧНА ПАПКА - НУЖНА HOP ПРОВЕРКА'}
                $count = '1'
                $best = ''
                $paths = $mp
            } else {
                $state = 'ЗАПОМНЕНАТА ПАПКА ЛИПСВА - ПРОВЕРИ'
            }
        }
        $r=$table.NewRow()
        $r['Изделие от CSV']=$product
        $r['Кандидати']=$count
        $r['Статус']=$state
        $r['Най-добра оценка']=$best
        $r['Пътища / Потвърдена папка']=$paths
        [void]$table.Rows.Add($r)
    }
    $productCandidatesGrid.DataSource=$table
    Refresh-ConfirmationQueue
}

function Get-NormalizedName {
    param($Value)
    $s = (Safe-Trim $Value).ToLowerInvariant()

    # Normalize common size forms first: 160/200, 160 x 200, 160*200 -> 160x200.
    $s = [regex]::Replace($s, '(?i)(\d+)\s*[/x*]\s*(\d+)', '$1x$2')
    # Preserve # as part of the variant identity: F and F# are DIFFERENT.
    $s = [regex]::Replace($s, '(?i)([a-z])#', '$1hash')
    $s = $s.Replace('#', ' hash ')
    $s = $s.Replace('_', ' ').Replace('-', ' ').Replace('\\', ' ').Replace('/', ' ')
    $s = [regex]::Replace($s, '[^a-z0-9а-яё ]', ' ')
    $s = [regex]::Replace($s, '\s+', ' ')
    return $s.Trim()
}

function Get-SearchTokens {
    param($Value)
    $n = Get-NormalizedName $Value
    if ([string]::IsNullOrWhiteSpace($n)) { return @() }
    return @($n -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Test-IsDimensionToken {
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $false }
    return ($Token -match '^\d{2,4}x\d{2,4}$')
}

function Test-IsVariantToken {
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $false }
    return @('f','b','c','l','r','a','fhash','bhash','chash','lhash','rhash','ahash','hash') -contains $Token.ToLowerInvariant()
}

function Test-IsGenericProductToken {
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $true }
    $t = $Token.ToLowerInvariant()
    $generic = @(
        'and','with','set','furniture',
        'bed','bedroom','hallway','night','stand',
        'dresser','wardrobe','cabinet','section',
        'coffee','dining','table','desk','shelf'
    )
    return ($generic -contains $t)
}

function Get-CoreTokens {
    param($Value)
    $tokens = @(Get-SearchTokens $Value)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($t in $tokens) {
        if (Test-IsDimensionToken $t) { continue }
        if (Test-IsVariantToken $t) { continue }
        if (Test-IsGenericProductToken $t) { continue }
        if ($t -match '^\d+$') { continue }
        # Single-letter tokens are too broad. Two-letter Roman numerals are allowed only as secondary tokens.
        if ($t.Length -lt 2) { continue }
        $out.Add($t)
    }
    return @($out.ToArray() | Select-Object -Unique)
}

function Get-LeafSearchName {
    param([string]$FolderName)
    return (Get-NormalizedName $FolderName)
}

function Build-FolderNameMap {
    $map = @{}
    $tokenMap = @{}
    foreach ($it in @($script:FolderIndex)) {
        $fn = Get-LeafSearchName $it.Folder
        if (-not [string]::IsNullOrWhiteSpace($fn)) {
            if (-not $map.ContainsKey($fn)) {
                $map[$fn] = New-Object System.Collections.Generic.List[object]
            }
            $map[$fn].Add($it)
        }

        foreach ($tok in @(Get-CoreTokens $it.Folder)) {
            # Avoid using tiny/generic tokens as the primary inverted-index key.
            if ($tok.Length -lt 3) { continue }
            if (-not $tokenMap.ContainsKey($tok)) {
                $tokenMap[$tok] = New-Object System.Collections.Generic.List[object]
            }
            $tokenMap[$tok].Add($it)
        }
    }
    $script:FolderNameMap = $map
    $script:FolderTokenMap = $tokenMap
}

function Load-FolderIndex {
    $script:FolderIndex = @()
    $script:FolderNameMap = @{}
    $script:FolderTokenMap = @{}
    if (-not [System.IO.File]::Exists($script:IndexCsv)) { return }
    try {
        $rows = @(Read-SafeCsv -Path $script:IndexCsv)
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows) {
            $path = Safe-Trim (Get-PropertyValue -Object $r -Name 'Path')
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            $items.Add([pscustomobject]@{
                Root = Safe-Trim (Get-PropertyValue -Object $r -Name 'Root')
                Folder = Safe-Trim (Get-PropertyValue -Object $r -Name 'Folder')
                Path = $path
            })
        }
        $script:FolderIndex = @($items.ToArray())
        $script:FolderNameMap = @{}
        $script:FolderTokenMap = @{}
    }
    catch {
        $script:FolderIndex = @()
        $script:FolderNameMap = @{}
        $script:FolderTokenMap = @{}
    }
}

function Save-FolderIndex {
    param([object[]]$Items)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Root;Folder;Path')
    foreach ($it in $Items) {
        $lines.Add(((Convert-ToCsvField $it.Root), (Convert-ToCsvField $it.Folder), (Convert-ToCsvField $it.Path)) -join ';')
    }
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($script:IndexCsv, $lines.ToArray(), $utf8Bom)
}

function Rebuild-FolderIndex {
    $items = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]
    $count = 0
    foreach ($root in @($script:AllowedFolders.ToArray() | Sort-Object -Unique)) {
        if (-not [System.IO.Directory]::Exists($root)) {
            $errors.Add("Недостъпна/липсваща разрешена папка: $root")
            continue
        }
        try {
            $rootName = [System.IO.Path]::GetFileName($root.TrimEnd('\'))
            $items.Add([pscustomobject]@{ Root=$root; Folder=$rootName; Path=$root })
            $dirs = [System.IO.Directory]::EnumerateDirectories($root, '*', [System.IO.SearchOption]::AllDirectories)
            foreach ($d in $dirs) {
                $name = [System.IO.Path]::GetFileName($d.TrimEnd('\'))
                $items.Add([pscustomobject]@{ Root=$root; Folder=$name; Path=$d })
                $count++
                if (($count % 250) -eq 0) {
                    $status.Text = "Каталогирани папки: $count ..."
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
        }
        catch {
            $errors.Add("Грешка при обхождане на '$root': $($_.Exception.Message)")
        }
    }
    $script:FolderIndex = @($items.ToArray())
    Build-FolderNameMap
    try { Save-FolderIndex -Items $script:FolderIndex } catch { $errors.Add("Неуспешно записване на локалния каталог: $($_.Exception.Message)") }
    return [pscustomobject]@{ Count=$script:FolderIndex.Count; Errors=@($errors.ToArray()) }
}


function Stop-CsvWorkerTimer {
    if ($null -ne $script:CsvWorkerTimer) {
        try { $script:CsvWorkerTimer.Stop() } catch {}
        try { $script:CsvWorkerTimer.Dispose() } catch {}
        $script:CsvWorkerTimer = $null
    }
}

function Load-CsvWorkerResults {
    $rowsRaw = @()
    $summaryRaw = @()
    if ([System.IO.File]::Exists($script:CsvRowsOut)) { $rowsRaw = @(Read-SafeCsv -Path $script:CsvRowsOut) }
    if ([System.IO.File]::Exists($script:CsvSummaryOut)) { $summaryRaw = @(Read-SafeCsv -Path $script:CsvSummaryOut) }

    $allRows = New-Object System.Collections.Generic.List[object]
    $detailTable = New-DataTable -Columns @('CSV', 'Изделие', '№', 'Детайл', 'L', 'B', 'T', 'Количество', 'Обработка', 'ProgramCodeF', 'Статус')
    foreach ($r in $rowsRaw) {
        $item = [pscustomobject]@{
            CsvFile = Get-PropertyValue -Object $r -Name 'CsvFile'
            CsvPath = Get-PropertyValue -Object $r -Name 'CsvPath'
            Product = Get-PropertyValue -Object $r -Name 'Product'
            DetailNo = Get-PropertyValue -Object $r -Name 'DetailNo'
            DetailName = Get-PropertyValue -Object $r -Name 'DetailName'
            Length = Get-PropertyValue -Object $r -Name 'Length'
            Width = Get-PropertyValue -Object $r -Name 'Width'
            Thickness = Get-PropertyValue -Object $r -Name 'Thickness'
            Quantity = Get-PropertyValue -Object $r -Name 'Quantity'
            ProcessingNo = Get-PropertyValue -Object $r -Name 'ProcessingNo'
            ProgramCodeF = Get-PropertyValue -Object $r -Name 'ProgramCodeF'
            Status = Get-PropertyValue -Object $r -Name 'Status'
        }
        $allRows.Add($item)
        $dr = $detailTable.NewRow()
        $dr['CSV']=$item.CsvFile; $dr['Изделие']=$item.Product; $dr['№']=$item.DetailNo; $dr['Детайл']=$item.DetailName
        $dr['L']=$item.Length; $dr['B']=$item.Width; $dr['T']=$item.Thickness; $dr['Количество']=$item.Quantity
        $dr['Обработка']=$item.ProcessingNo; $dr['ProgramCodeF']=$item.ProgramCodeF; $dr['Статус']=$item.Status
        [void]$detailTable.Rows.Add($dr)
    }

    $productTable = New-DataTable -Columns @('Изделие', 'CSV файлове', 'Физически детайли', 'Редове', 'ProgramCodeF', 'Без ProgramCodeF', 'Двойни обработки')
    $productGroups = @($allRows.ToArray() | Group-Object Product | Sort-Object Name)
    foreach ($g in $productGroups) {
        $rr=@($g.Group)
        $pr=$productTable.NewRow()
        $pr['Изделие']=[string]$g.Name
        $pr['CSV файлове']=[string](@($rr|ForEach-Object{$_.CsvFile}|Sort-Object -Unique).Count)
        $pr['Физически детайли']=[string](@($rr|ForEach-Object{"$($_.DetailNo)|$($_.DetailName)|$($_.Length)|$($_.Width)|$($_.Thickness)"}|Sort-Object -Unique).Count)
        $pr['Редове']=[string]$rr.Count
        $pr['ProgramCodeF']=[string](@($rr|Where-Object{-not [string]::IsNullOrWhiteSpace($_.ProgramCodeF)}).Count)
        $pr['Без ProgramCodeF']=[string](@($rr|Where-Object{[string]::IsNullOrWhiteSpace($_.ProgramCodeF)}).Count)
        $pr['Двойни обработки']=[string](@($rr|Group-Object{"$($_.DetailNo)|$($_.DetailName)|$($_.Length)|$($_.Width)|$($_.Thickness)"}|Where-Object{$_.Count -gt 1}).Count)
        [void]$productTable.Rows.Add($pr)
    }

    $fileSummaryTable = New-DataTable -Columns @('CSV', 'Редове в CSV', 'Разчетени редове', 'Изделия', 'Физически детайли', 'ProgramCodeF', 'Без ProgramCodeF', 'Двойни обработки', 'Дублиран ProgramCodeF', 'Неразчетени редове', 'Статус', 'Грешка')
    $warnings = New-Object System.Collections.Generic.List[string]
    $fileErrors=0; $filesWithCheck=0
    foreach ($x in $summaryRaw) {
        $fr=$fileSummaryTable.NewRow()
        $fr['CSV']=Get-PropertyValue -Object $x -Name 'CSV'
        $fr['Редове в CSV']=Get-PropertyValue -Object $x -Name 'RawRows'
        $fr['Разчетени редове']=Get-PropertyValue -Object $x -Name 'ParsedRows'
        $fr['Изделия']=Get-PropertyValue -Object $x -Name 'Products'
        $fr['Физически детайли']=Get-PropertyValue -Object $x -Name 'PhysicalDetails'
        $fr['ProgramCodeF']=Get-PropertyValue -Object $x -Name 'ProgramCodeF'
        $fr['Без ProgramCodeF']=Get-PropertyValue -Object $x -Name 'NoProgram'
        $fr['Двойни обработки']=Get-PropertyValue -Object $x -Name 'DoubleProcessing'
        $fr['Дублиран ProgramCodeF']=Get-PropertyValue -Object $x -Name 'DuplicateProgram'
        $fr['Неразчетени редове']=Get-PropertyValue -Object $x -Name 'UnparsedRows'
        $fr['Статус']=Get-PropertyValue -Object $x -Name 'Status'
        $fr['Грешка']=Get-PropertyValue -Object $x -Name 'ErrorText'
        [void]$fileSummaryTable.Rows.Add($fr)
        $st=[string]($fr['Статус'])
        if($st -eq 'ГРЕШКА'){$fileErrors++;$filesWithCheck++}
        elseif($st -eq 'ПРОВЕРИ'){$filesWithCheck++}
        $err=Safe-Trim (Get-PropertyValue -Object $x -Name 'ErrorText')
        if(-not [string]::IsNullOrWhiteSpace($err)){$warnings.Add("$($fr['CSV']): $err")}
    }

    $all=@($allRows.ToArray())
    $noPrograms=@($all|Where-Object{[string]::IsNullOrWhiteSpace($_.ProgramCodeF)})
    if($noPrograms.Count -gt 0){
        $warnings.Add('');$warnings.Add("РЕДОВЕ БЕЗ ProgramCodeF: $($noPrograms.Count)")
        foreach($r in $noPrograms){$warnings.Add("  $($r.Product) | №$($r.DetailNo) | $($r.DetailName) | $($r.Length)x$($r.Width)x$($r.Thickness) | CSV: $($r.CsvFile)")}
    }
    $multi=@($all|Group-Object{"$($_.CsvFile)|$($_.Product)|$($_.DetailNo)|$($_.DetailName)|$($_.Length)|$($_.Width)|$($_.Thickness)"}|Where-Object{$_.Count -gt 1})
    if($multi.Count -gt 0){
        $warnings.Add('');$warnings.Add("ДЕТАЙЛИ С ПОВЕЧЕ ОТ ЕДНА ОБРАБОТКА: $($multi.Count)")
        foreach($grp in $multi){$items=@($grp.Group|Sort-Object {[int]$_.ProcessingNo});if($items.Count -eq 0){continue};$first=$items[0];$warnings.Add("  $($first.Product) | №$($first.DetailNo) | $($first.DetailName) | CSV: $($first.CsvFile)");foreach($it in $items){$warnings.Add("      обработка $($it.ProcessingNo) -> $($it.ProgramCodeF)")}}
    }
    $warnings.Add('');$warnings.Add('КОНТРОЛ V3.0:')
    $warnings.Add("  - CSV файлове: $($summaryRaw.Count)")
    $warnings.Add("  - Изделия: $($productGroups.Count)")
    $warnings.Add("  - Разчетени редове: $($allRows.Count)")
    $warnings.Add("  - Файлове с грешка: $fileErrors")
    $warnings.Add("  - Файлове за проверка: $filesWithCheck")
    $warnings.Add('');$warnings.Add('ЗАЩИТА V3.0:')
    $warnings.Add('  - CSV worker чете само *.csv от фиксираната папка и не обхожда подпапки.')
    $warnings.Add('  - Word файлове не се отварят и не се прави опит за достъп до съдържанието им.')
    $warnings.Add('  - .hop се четат САМО в потвърденото 03_Mashini\1_Holzher поддърво; при няколко HOP подпапки размерите служат само за избор на правилната подпапка.')
    $warnings.Add('  - При нормална PGMX обработка съдържанието на .pgmx не се отваря; използват се имена/пътища и доказаният V2.9 конвертор.')
    $warnings.Add('  - В пакета е включен точно предоставеният от потребителя HOP_to_PGMX_2_9.exe.')
    $warnings.Add('  - Каталогът чете само имена/пътища на директории и спира при папка с 03_Mashini.')

    $productsGrid.DataSource=$productTable
    $detailsGrid.DataSource=$detailTable
    $fileSummaryGrid.DataSource=$fileSummaryTable
    $warningsBox.Text=($warnings.ToArray() -join [Environment]::NewLine)
    $script:LastFileSummary=$fileSummaryTable
    $script:LastAllRows=@($allRows.ToArray())
    $script:LastWarnings=@($warnings.ToArray())
    $script:LastFolder=$script:DefaultCsvFolder
    $script:SearchResultObjects=@()
    $exportButton.Enabled=$true
    $productCandidatesGrid.DataSource=$null
    Refresh-ConfirmationQueue
    if($fileSummaryTable.Rows.Count -gt 0){$fileSummaryGrid.ClearSelection();$fileSummaryGrid.Rows[0].Selected=$true;$fileSummaryGrid.CurrentCell=$fileSummaryGrid.Rows[0].Cells[0];Update-SelectedCsvProducts}
    $status.Text="CSV: $($summaryRaw.Count) | Изделия: $($productGroups.Count) | Редове: $($allRows.Count) | готово."
    $tabs.SelectedTab=$tabPerCsv
}

function Get-CsvRegistryIdentity {
    param([Parameter(Mandatory=$true)][System.IO.FileInfo]$File)
    return (($File.FullName.ToLowerInvariant()) + '|' + ([string]$File.Length) + '|' + ([string]$File.LastWriteTimeUtc.Ticks))
}

function Load-CsvRegistryRows {
    if(-not [System.IO.File]::Exists($script:CsvRegistryCsv)){ return @() }
    try { return @(Import-Csv -LiteralPath $script:CsvRegistryCsv -Delimiter ';' -Encoding UTF8) } catch { return @() }
}

function Save-CsvRegistryRows {
    param([object[]]$Rows)
    $headers=@('Csv','FullPath','Length','LastWriteUtc','Identity','Status','FirstSeen','LastAttempt','CompletedAt','Products','Details','Unconfirmed','DetailErrors','Message')
    $lines=New-Object System.Collections.Generic.List[string]
    $lines.Add((@($headers | ForEach-Object { Convert-ToCsvField $_ }) -join ';'))
    foreach($x in @($Rows)){
        $vals=New-Object System.Collections.Generic.List[string]
        foreach($h in $headers){ $vals.Add((Convert-ToCsvField (Get-PropertyValue $x $h))) }
        $lines.Add(($vals.ToArray() -join ';'))
    }
    [System.IO.File]::WriteAllLines($script:CsvRegistryCsv,$lines.ToArray(),(New-Object System.Text.UTF8Encoding($true)))
}

function New-CsvRegistryRowFromFile {
    param([Parameter(Mandatory=$true)][System.IO.FileInfo]$File)
    $now=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    return [pscustomobject]@{
        Csv=$File.Name;FullPath=$File.FullName;Length=[string]$File.Length;LastWriteUtc=$File.LastWriteTimeUtc.ToString('o');
        Identity=(Get-CsvRegistryIdentity -File $File);Status='НОВ';FirstSeen=$now;LastAttempt='';CompletedAt='';
        Products='';Details='';Unconfirmed='';DetailErrors='';Message='Открит нов CSV.'
    }
}

function Sync-CsvRegistryDiscovery {
    $folder=$script:DefaultCsvFolder
    $rows=New-Object System.Collections.Generic.List[object]
    foreach($x in @(Load-CsvRegistryRows)){ $rows.Add($x) }
    if(-not [System.IO.Directory]::Exists($folder)){ return @($rows.ToArray()) }
    $known=@{}
    foreach($x in @($rows.ToArray())){ $id=Safe-Trim (Get-PropertyValue $x 'Identity'); if(-not [string]::IsNullOrWhiteSpace($id)){$known[$id]=$true} }
    $changed=$false
    foreach($file in @(Get-ChildItem -LiteralPath $folder -File -Filter '*.csv' -ErrorAction SilentlyContinue | Sort-Object Name)){
        $id=Get-CsvRegistryIdentity -File $file
        if(-not $known.ContainsKey($id)){
            $rows.Add((New-CsvRegistryRowFromFile -File $file));$known[$id]=$true;$changed=$true
        }
    }
    if($changed){ Save-CsvRegistryRows -Rows @($rows.ToArray()) }
    return @($rows.ToArray())
}

function Refresh-CsvRegistryGrid {
    try{
        $rows=@(Sync-CsvRegistryDiscovery)
        if($null -eq $operatorCsvRegistryGrid){return}
        $table=New-DataTable -Columns @('CSV','Статус','Променен','Последен опит','Завършен','Изделия','Детайли','Непотвърдени','Грешки детайли','Съобщение','Пълен път')
        foreach($x in @($rows | Sort-Object @{Expression={Safe-Trim (Get-PropertyValue $_ 'LastAttempt')};Descending=$true}, @{Expression={Safe-Trim (Get-PropertyValue $_ 'Csv')}})){
            $r=$table.NewRow();$r['CSV']=Get-PropertyValue $x 'Csv';$r['Статус']=Get-PropertyValue $x 'Status';$r['Променен']=Get-PropertyValue $x 'LastWriteUtc';$r['Последен опит']=Get-PropertyValue $x 'LastAttempt';$r['Завършен']=Get-PropertyValue $x 'CompletedAt';$r['Изделия']=Get-PropertyValue $x 'Products';$r['Детайли']=Get-PropertyValue $x 'Details';$r['Непотвърдени']=Get-PropertyValue $x 'Unconfirmed';$r['Грешки детайли']=Get-PropertyValue $x 'DetailErrors';$r['Съобщение']=Get-PropertyValue $x 'Message';$r['Пълен път']=Get-PropertyValue $x 'FullPath';[void]$table.Rows.Add($r)
        }
        $operatorCsvRegistryGrid.DataSource=$table
        try{if($null -ne $script:MainCsvRegistryGrid){$script:MainCsvRegistryGrid.DataSource=$table}}catch{}
    }catch{}
}

function Prepare-CsvBatchPlan {
    $folder=$script:DefaultCsvFolder
    $allRows=New-Object System.Collections.Generic.List[object]
    foreach($x in @(Sync-CsvRegistryDiscovery)){ $allRows.Add($x) }
    $lookup=@{}
    foreach($x in @($allRows.ToArray())){ $id=Safe-Trim (Get-PropertyValue $x 'Identity'); if(-not [string]::IsNullOrWhiteSpace($id)){$lookup[$id]=$x} }
    $batch=New-Object System.Collections.Generic.List[object]
    $skip=New-Object System.Collections.Generic.List[string]
    $now=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    foreach($file in @(Get-ChildItem -LiteralPath $folder -File -Filter '*.csv' -ErrorAction SilentlyContinue | Sort-Object Name)){
        $id=Get-CsvRegistryIdentity -File $file
        $row=$null;if($lookup.ContainsKey($id)){$row=$lookup[$id]}
        if($null -eq $row){$row=New-CsvRegistryRowFromFile -File $file;$allRows.Add($row);$lookup[$id]=$row}
        if((Safe-Trim (Get-PropertyValue $row 'Status')) -eq 'ЗАВЪРШЕН'){
            $skip.Add($file.FullName);continue
        }
        $row.Status='В ПРОЦЕС';$row.LastAttempt=$now;$row.CompletedAt='';$row.Message='Стартирана обработка.'
        $batch.Add([pscustomobject]@{Csv=$file.Name;FullPath=$file.FullName;Identity=$id})
    }
    Save-CsvRegistryRows -Rows @($allRows.ToArray())
    $skipLines=New-Object System.Collections.Generic.List[string];$skipLines.Add('FullPath');foreach($p in $skip.ToArray()){$skipLines.Add((Convert-ToCsvField $p))}
    [System.IO.File]::WriteAllLines($script:CsvSkipListCsv,$skipLines.ToArray(),(New-Object System.Text.UTF8Encoding($true)))
    $script:CsvCurrentBatch=@($batch.ToArray())
    Refresh-CsvRegistryGrid
    return @($script:CsvCurrentBatch)
}

function Update-CsvRegistryAfterCsvRead {
    $registry=New-Object System.Collections.Generic.List[object];foreach($x in @(Load-CsvRegistryRows)){$registry.Add($x)}
    $lookup=@{};foreach($x in @($registry.ToArray())){$lookup[(Safe-Trim (Get-PropertyValue $x 'Identity'))]=$x}
    foreach($b in @($script:CsvCurrentBatch)){
        if(-not $lookup.ContainsKey($b.Identity)){continue};$row=$lookup[$b.Identity]
        $sum=$null
        if($null -ne $script:LastFileSummary){foreach($sr in $script:LastFileSummary.Rows){if((Safe-Trim $sr['CSV']) -eq $b.Csv){$sum=$sr;break}}}
        if($null -eq $sum){$row.Status='ГРЕШКА';$row.Message='CSV worker не върна обобщение за файла.';continue}
        $row.Products=Safe-Trim $sum['Изделия'];$row.Details=Safe-Trim $sum['Физически детайли']
        if((Safe-Trim $sum['Статус']) -eq 'ГРЕШКА'){$row.Status='ГРЕШКА';$row.Message=Safe-Trim $sum['Грешка'];if([string]::IsNullOrWhiteSpace($row.Message)){$row.Message='Грешка при четене на CSV.'}}
        else{$row.Status='В ПРОЦЕС';$row.Message='CSV е прочетен. Очаква потвърждение/PGMX.'}
    }
    Save-CsvRegistryRows -Rows @($registry.ToArray());Refresh-CsvRegistryGrid
}

function Set-CurrentCsvBatchRegistryError {
    param([string]$Message)
    $registry=New-Object System.Collections.Generic.List[object];foreach($x in @(Load-CsvRegistryRows)){$registry.Add($x)}
    $lookup=@{};foreach($x in @($registry.ToArray())){$lookup[(Safe-Trim (Get-PropertyValue $x 'Identity'))]=$x}
    foreach($b in @($script:CsvCurrentBatch)){if($lookup.ContainsKey($b.Identity)){$r=$lookup[$b.Identity];if((Safe-Trim $r.Status) -ne 'ЗАВЪРШЕН'){$r.Status='ГРЕШКА';$r.Message=$Message}}}
    Save-CsvRegistryRows -Rows @($registry.ToArray());Refresh-CsvRegistryGrid
}

function Finalize-CsvRegistryForCurrentBatch {
    param([bool]$PipelineSucceeded,[string]$PipelineMessage='')
    if(@($script:CsvCurrentBatch).Count -eq 0){return}
    $registry=New-Object System.Collections.Generic.List[object];foreach($x in @(Load-CsvRegistryRows)){$registry.Add($x)}
    $lookup=@{};foreach($x in @($registry.ToArray())){$lookup[(Safe-Trim (Get-PropertyValue $x 'Identity'))]=$x}
    $states=@(Get-OperatorProductStateRows);$stateMap=@{};foreach($s in $states){$stateMap[(Safe-Trim $s.Product).ToLowerInvariant()]=$s}
    $okResults=@('ГОТОВ','СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF')
    foreach($b in @($script:CsvCurrentBatch)){
        if(-not $lookup.ContainsKey($b.Identity)){continue};$row=$lookup[$b.Identity]
        if((Safe-Trim $row.Status) -eq 'ГРЕШКА' -and -not $PipelineSucceeded){$row.Message=$PipelineMessage;continue}
        $csvRows=@($script:LastAllRows | Where-Object { (Safe-Trim $_.CsvFile) -eq $b.Csv })
        # A CSV parse error has no detail rows. Preserve its ERROR status even if the other files finish successfully.
        if((Safe-Trim $row.Status) -eq 'ГРЕШКА' -and $csvRows.Count -eq 0){continue}
        $products=@($csvRows | ForEach-Object { Safe-Trim $_.Product } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        $unconfirmed=0;foreach($p in $products){$key=$p.ToLowerInvariant();if(-not $stateMap.ContainsKey($key) -or -not [bool]$stateMap[$key].Confirmed){$unconfirmed++}}
        $expectedRows=@($csvRows | Where-Object { -not [string]::IsNullOrWhiteSpace((Safe-Trim $_.ProgramCodeF)) })
        $noProgram=@($csvRows | Where-Object { [string]::IsNullOrWhiteSpace((Safe-Trim $_.ProgramCodeF)) }).Count
        $resultRows=@($script:PgmxResultObjects | Where-Object { (Safe-Trim $_.CSV) -eq $b.Csv })
        $badResults=@($resultRows | Where-Object { $okResults -notcontains (Safe-Trim $_.Result) }).Count
        $missingResults=[Math]::Max(0,$expectedRows.Count-$resultRows.Count)
        $detailErrors=$noProgram+$badResults+$missingResults
        $row.Products=[string]$products.Count;$row.Details=[string]$csvRows.Count;$row.Unconfirmed=[string]$unconfirmed;$row.DetailErrors=[string]$detailErrors
        if(-not $PipelineSucceeded){$row.Status='ГРЕШКА';$row.CompletedAt='';$row.Message=if([string]::IsNullOrWhiteSpace($PipelineMessage)){'Общият процес не завърши.'}else{$PipelineMessage}}
        elseif($unconfirmed -gt 0 -or $detailErrors -gt 0){$row.Status='ЧАСТИЧНО';$row.CompletedAt='';$row.Message="Непотвърдени изделия: $unconfirmed; проблемни/неотчетени детайли: $detailErrors. Ще се опита отново при следващо стартиране."}
        else{$row.Status='ЗАВЪРШЕН';$row.CompletedAt=Get-Date -Format 'yyyy-MM-dd HH:mm:ss';$row.Message=if([string]::IsNullOrWhiteSpace($PipelineMessage)){'Обработката е завършена успешно.'}else{$PipelineMessage}}
    }
    Save-CsvRegistryRows -Rows @($registry.ToArray());Refresh-CsvRegistryGrid
}

function Test-IsFolderIndexCurrentToday {
    if (-not [System.IO.File]::Exists($script:IndexCsv)) { return $false }
    try {
        $index = Get-Item -LiteralPath $script:IndexCsv -ErrorAction Stop
        if ($index.LastWriteTime.Date -ne (Get-Date).Date) { return $false }
        if ([System.IO.File]::Exists($script:SettingsCsv)) {
            $settings = Get-Item -LiteralPath $script:SettingsCsv -ErrorAction Stop
            if ($index.LastWriteTimeUtc -lt $settings.LastWriteTimeUtc) { return $false }
        }
        return $true
    } catch { return $false }
}

function Refresh-OperatorCatalogLabel {
    try{
        if($null -eq $operatorCatalogLabel){return}
        if(Test-IsFolderIndexCurrentToday){$dt=(Get-Item -LiteralPath $script:IndexCsv).LastWriteTime;$operatorCatalogLabel.Text='Каталог: обновен днес '+$dt.ToString('HH:mm')}
        elseif([System.IO.File]::Exists($script:IndexCsv)){$dt=(Get-Item -LiteralPath $script:IndexCsv).LastWriteTime;$operatorCatalogLabel.Text='Каталог: стар ('+$dt.ToString('dd.MM HH:mm')+') - ще се обнови автоматично'}
        else{$operatorCatalogLabel.Text='Каталог: липсва - ще се създаде автоматично'}
    }catch{}
}

function Continue-OperatorAfterIndex {
    $action=$script:OperatorAfterIndexAction;$script:OperatorAfterIndexAction=''
    Refresh-OperatorCatalogLabel
    if($action -eq 'Csv'){try{Start-BackgroundCsvRead}catch{$script:OperatorAutoFlow=$false;$status.Text='Грешка след дневния каталог: '+$_.Exception.Message}}
    elseif($action -eq 'Search'){try{Start-BackgroundProductSearch}catch{$script:OperatorAutoFlow=$false;$status.Text='Грешка след дневния каталог: '+$_.Exception.Message}}
}

function Ensure-OperatorDailyCatalog {
    param([string]$AfterAction='')
    if(Test-IsFolderIndexCurrentToday){Refresh-OperatorCatalogLabel;try{Set-OperatorCheckpoint -CurrentStage 'CATALOG_DONE' -LastSafeStage 'CATALOG_DONE' -Status 'ACTIVE' -Message 'Дневният каталог вече е актуален.'}catch{};return $true}
    if(-not [string]::IsNullOrWhiteSpace($AfterAction)){$script:OperatorAfterIndexAction=$AfterAction}
    $status.Text='Каталогът не е обновен днес. Стартирам автоматично дневно обновяване във фон...';try{Set-OperatorCheckpoint -CurrentStage 'CATALOG_RUNNING' -Status 'ACTIVE' -Message 'Обновяване на дневния каталог.'}catch{}
    Start-BackgroundFolderIndex -Silent
    return $false
}

function Start-BackgroundCsvRead {
    $folder=$script:DefaultCsvFolder
    if(-not [System.IO.Directory]::Exists($folder)){throw "CSV папката не е достъпна: $folder"}
    if(-not [System.IO.File]::Exists($script:CsvWorkerScript)){throw "Липсва помощният файл: $($script:CsvWorkerScript)"}
    if($null -ne $script:CsvWorkerProcess -and -not $script:CsvWorkerProcess.HasExited){return}
    Test-OperatorStopRequested

    $batch=@(Prepare-CsvBatchPlan)
    if($batch.Count -eq 0){
        $script:LastAllRows=@();$script:SearchResultObjects=@();$script:PgmxResultObjects=@()
        try{$productsGrid.DataSource=$null;$detailsGrid.DataSource=$null;$fileSummaryGrid.DataSource=$null;$productCandidatesGrid.DataSource=$null}catch{}
        $status.Text='Няма нови/незавършени CSV. Всички текущи CSV файлове вече са със статус ЗАВЪРШЕН.'
        $script:OperatorAutoFlow=$false;$script:OperatorRunHopAfterSearch=$false;$script:OperatorRunExecuteAfterHop=$false
        try{Complete-OperatorRunState -Message 'Няма нови/незавършени CSV за обработка.'}catch{}
        Refresh-OperatorDashboard;Refresh-CsvRegistryGrid
        return
    }

    foreach($f in @($script:CsvRowsOut,$script:CsvSummaryOut,$script:CsvProgressFile)){try{if([System.IO.File]::Exists($f)){[System.IO.File]::Delete($f)}}catch{}}
    $argLine="-NoProfile -ExecutionPolicy Bypass -File `"$($script:CsvWorkerScript)`" -InputFolder `"$folder`" -RowsOut `"$($script:CsvRowsOut)`" -SummaryOut `"$($script:CsvSummaryOut)`" -ProgressFile `"$($script:CsvProgressFile)`" -SkipListCsv `"$($script:CsvSkipListCsv)`""
    try{Set-OperatorCheckpoint -CurrentStage 'CSV_RUNNING' -Status 'ACTIVE' -Message 'Четене на CSV файловете.'}catch{}
    $script:CsvWorkerProcess=Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -WindowStyle Hidden -PassThru
    $scanButton.Enabled=$false;$status.Text="CSV четенето стартира във фон. За обработка: $($batch.Count) файла; завършените се пропускат."
    Stop-CsvWorkerTimer
    $script:CsvWorkerTimer=New-Object System.Windows.Forms.Timer
    $script:CsvWorkerTimer.Interval=400
    $script:CsvWorkerTimer.Add_Tick({
        try{if([System.IO.File]::Exists($script:CsvProgressFile)){$p=[System.IO.File]::ReadAllText($script:CsvProgressFile,[System.Text.Encoding]::UTF8).Trim();if(-not [string]::IsNullOrWhiteSpace($p)){$status.Text=$p}}}catch{}
        if($null -ne $script:CsvWorkerProcess -and $script:CsvWorkerProcess.HasExited){
            Stop-CsvWorkerTimer;$code=$script:CsvWorkerProcess.ExitCode;$script:CsvWorkerProcess.Dispose();$script:CsvWorkerProcess=$null;$scanButton.Enabled=$true
            if($code -eq 0 -and [System.IO.File]::Exists($script:CsvRowsOut) -and [System.IO.File]::Exists($script:CsvSummaryOut)){
                try{Load-CsvWorkerResults;Update-CsvRegistryAfterCsvRead;Set-OperatorCheckpoint -CurrentStage 'CSV_DONE' -LastSafeStage 'CSV_DONE' -Status 'ACTIVE' -Message ('CSV са прочетени. Редове: '+@($script:LastAllRows).Count)}catch{Set-CurrentCsvBatchRegistryError -Message ('Грешка при зареждане на CSV резултата: '+$_.Exception.Message);[System.Windows.Forms.MessageBox]::Show("CSV са прочетени, но зареждането даде грешка:`r`n$($_.Exception.Message)",'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null;$status.Text='Грешка при зареждане на CSV резултата.'}; try{Refresh-OperatorDashboard}catch{}; try{Continue-OperatorAfterCsv}catch{}
            }else{
                $detail='';try{if([System.IO.File]::Exists($script:CsvProgressFile)){$detail=[System.IO.File]::ReadAllText($script:CsvProgressFile,[System.Text.Encoding]::UTF8).Trim()}}catch{};Set-CurrentCsvBatchRegistryError -Message ("CSV worker грешка. $detail");$status.Text="Грешка при CSV worker. $detail";try{Set-OperatorRunError -Message $status.Text}catch{};$script:OperatorAutoFlow=$false
            }
        }
    })
    $script:CsvWorkerTimer.Start()
}

function Stop-IndexWorkerTimer {
    if ($null -ne $script:IndexWorkerTimer) {
        try { $script:IndexWorkerTimer.Stop() } catch {}
        try { $script:IndexWorkerTimer.Dispose() } catch {}
        $script:IndexWorkerTimer = $null
    }
}

function Start-BackgroundFolderIndex {
    param([switch]$Silent)
    if ($null -ne $script:IndexWorkerProcess -and -not $script:IndexWorkerProcess.HasExited) {
        if(-not $Silent){[System.Windows.Forms.MessageBox]::Show('Каталогирането вече работи във фон.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null}
        return
    }
    if (-not [System.IO.File]::Exists($script:IndexWorkerScript)) {
        throw "Липсва помощният файл: $($script:IndexWorkerScript)"
    }

    try { Save-AllowedFolders } catch {}
    try { if ([System.IO.File]::Exists($script:IndexTempCsv)) { [System.IO.File]::Delete($script:IndexTempCsv) } } catch {}
    try { if ([System.IO.File]::Exists($script:IndexProgressFile)) { [System.IO.File]::Delete($script:IndexProgressFile) } } catch {}

    $argLine = "-NoProfile -ExecutionPolicy Bypass -File `"$($script:IndexWorkerScript)`" -SettingsCsv `"$($script:SettingsCsv)`" -IndexOut `"$($script:IndexTempCsv)`" -ProgressFile `"$($script:IndexProgressFile)`" -SettingsVersion `"$($script:SettingsVersion)`""
    $script:IndexWorkerProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -WindowStyle Hidden -PassThru
    $reindexButton.Enabled = $false
    $status.Text = 'Каталогирането стартира във фон. Можеш да продължиш да работиш в програмата.'

    Stop-IndexWorkerTimer
    $script:IndexWorkerTimer = New-Object System.Windows.Forms.Timer
    $script:IndexWorkerTimer.Interval = 600
    $script:IndexWorkerTimer.Add_Tick({
        try {
            if ([System.IO.File]::Exists($script:IndexProgressFile)) {
                $progress = [System.IO.File]::ReadAllText($script:IndexProgressFile, [System.Text.Encoding]::UTF8).Trim()
                if (-not [string]::IsNullOrWhiteSpace($progress)) { $status.Text = $progress }
            }
        } catch {}

        if ($null -ne $script:IndexWorkerProcess -and $script:IndexWorkerProcess.HasExited) {
            Stop-IndexWorkerTimer
            $exitCode = $script:IndexWorkerProcess.ExitCode
            $script:IndexWorkerProcess.Dispose()
            $script:IndexWorkerProcess = $null
            $reindexButton.Enabled = $true

            if ($exitCode -eq 0 -and [System.IO.File]::Exists($script:IndexTempCsv)) {
                try {
                    [System.IO.File]::Copy($script:IndexTempCsv, $script:IndexCsv, $true)
                    [System.IO.File]::Delete($script:IndexTempCsv)
                    Update-FolderIndexGrid
                    $status.Text = "Каталогът е обновен във фон: $($script:FolderIndexCount) папки. Търсенето на изделия работи само по локалния каталог."
                    try{Refresh-OperatorCatalogLabel;Refresh-OperatorDashboard}catch{}
                    try{Set-OperatorCheckpoint -CurrentStage 'CATALOG_DONE' -LastSafeStage 'CATALOG_DONE' -Status 'ACTIVE' -Message ('Каталогът е обновен: '+$script:FolderIndexCount+' папки.')}catch{}
                    try{Continue-OperatorAfterIndex}catch{$script:OperatorAutoFlow=$false;$status.Text='Каталогът е готов, но продължението даде грешка: '+$_.Exception.Message}
                }
                catch {
                    $status.Text = "Каталогът приключи, но локалното зареждане даде грешка: $($_.Exception.Message)"
                }
            }
            else {
                $detail = ''
                try { if ([System.IO.File]::Exists($script:IndexProgressFile)) { $detail = [System.IO.File]::ReadAllText($script:IndexProgressFile, [System.Text.Encoding]::UTF8).Trim() } } catch {}
                $status.Text = "Грешка при фоновото каталогиране. $detail"
                try{Set-OperatorRunError -Message $status.Text}catch{}
                $script:OperatorAfterIndexAction='';$script:OperatorAutoFlow=$false;$script:OperatorRunHopAfterSearch=$false;$script:OperatorRunExecuteAfterHop=$false
                try{Refresh-OperatorCatalogLabel}catch{}
            }
        }
    })
    $script:IndexWorkerTimer.Start()
}

function Update-FolderIndexGrid {
    if ($null -eq $folderIndexGrid) { return }
    $table = New-DataTable -Columns @('Папка', 'Пълен път', 'Разрешен корен')
    $script:FolderIndexCount = 0
    $shown = 0
if ([System.IO.File]::Exists($script:IndexCsv)) {
        $reader = $null
        try {
            $reader = New-Object System.IO.StreamReader -ArgumentList @($script:IndexCsv, [System.Text.Encoding]::UTF8, $true)
            $header = $reader.ReadLine()
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $script:FolderIndexCount++
                if ($shown -lt 300) {
                    # Index paths used here never contain semicolons in the current server structure.
                    $parts = $line -split ';', 3
                    if ($parts.Count -ge 3) {
                        $r = $table.NewRow()
                        $r['Разрешен корен'] = $parts[0].Trim('"')
                        $r['Папка'] = $parts[1].Trim('"')
                        $r['Пълен път'] = $parts[2].Trim('"')
                        [void]$table.Rows.Add($r)
                        $shown++
                    }
                }
            }
        }
        catch {}
        finally { if ($null -ne $reader) { $reader.Dispose() } }
    }
    $folderIndexGrid.DataSource = $table
    if ($null -ne $indexInfoLabel) {
        $stamp = ''
        try {
            if ([System.IO.File]::Exists($script:IndexCsv)) {
                $dt = [System.IO.File]::GetLastWriteTime($script:IndexCsv)
                $stamp = " | локален каталог от: $($dt.ToString('dd.MM.yyyy HH:mm'))"
            }
        } catch {}
        $indexInfoLabel.Text = "Каталог: $($script:FolderIndexCount) папки$stamp. На екрана се показват първите $shown."
    }
}

# V5.1.13: product category is a HARD safety filter for automatic search.
# Rules live in PWX500_ProductCategories.csv and can be edited from the UI without changing code.
function Get-ExpectedCategoryTargets {
    param([string]$Product)
    if($script:ProductCategoryRules.Count -eq 0){Load-ProductCategoryRules}
    $n=Get-NormalizedName $Product
    foreach($r in @($script:ProductCategoryRules)){
        $p=Get-NormalizedName ([string]$r.Phrase)
        if([string]::IsNullOrWhiteSpace($p)){continue}
        if($n -eq $p -or $n.StartsWith($p+' ')){
            return @(([string]$r.Targets -split '\|') | ForEach-Object {$_.Trim()} | Where-Object {-not [string]::IsNullOrWhiteSpace($_)})
        }
    }
    return @()
}
function Test-PathMatchesExpectedTargets {
    param([string]$Path,[string[]]$Targets)
    if($null -eq $Targets -or $Targets.Count -eq 0){return $false}
    $pn=(Normalize-PathForCompare $Path).ToLowerInvariant()
    foreach($target0 in @($Targets)){
        $target=(Safe-Trim $target0).ToLowerInvariant()
        if([string]::IsNullOrWhiteSpace($target)){continue}
        foreach($segment in @($pn -split '\\')){
            if((Safe-Trim $segment).ToLowerInvariant().Contains($target)){return $true}
        }
    }
    return $false
}

function Get-ProductCandidateScore {
    param(
        [Parameter(Mandatory=$true)][string]$Product,
        [Parameter(Mandatory=$true)]$FolderItem
    )

    $productNorm = Get-NormalizedName $Product
    $folderNorm = Get-NormalizedName $FolderItem.Folder
    if ([string]::IsNullOrWhiteSpace($productNorm) -or [string]::IsNullOrWhiteSpace($folderNorm)) { return -1 }

    $pTokens = @(Get-SearchTokens $Product)
    $fTokens = @(Get-SearchTokens $FolderItem.Folder)
    $pCore = @(Get-CoreTokens $Product)
    $fCore = @(Get-CoreTokens $FolderItem.Folder)
    if ($fCore.Count -eq 0) { return -1 }

    $pSet = @{}
    foreach ($t in $pTokens) { $pSet[$t] = $true }
    $pCoreSet = @{}
    foreach ($t in $pCore) { $pCoreSet[$t] = $true }

    # Every meaningful token from the folder name must occur in the CSV product name.
    foreach ($t in $fCore) {
        if (-not $pCoreSet.ContainsKey($t)) { return -1 }
    }

    $score = 20 * $fCore.Count

    # A full folder phrase at the end of the product name is a very strong match.
    if ($productNorm -eq $folderNorm -or $productNorm.EndsWith(' ' + $folderNorm)) { $score += 100 }

    $folderDims = @($fTokens | Where-Object { Test-IsDimensionToken $_ } | Select-Object -Unique)
    $productDims = @($pTokens | Where-Object { Test-IsDimensionToken $_ } | Select-Object -Unique)
    if ($folderDims.Count -gt 0) {
        foreach ($d in $folderDims) {
            if ($pSet.ContainsKey($d)) { $score += 30 }
            elseif ($productDims.Count -gt 0) { return -1 } # explicit size mismatch
        }
    }

    $folderVariants = @($fTokens | Where-Object { Test-IsVariantToken $_ } | Select-Object -Unique)
    $productVariants = @($pTokens | Where-Object { Test-IsVariantToken $_ } | Select-Object -Unique)
    if ($folderVariants.Count -gt 0) {
        $variantMatch = $false
        foreach ($v in $folderVariants) { if ($pSet.ContainsKey($v)) { $variantMatch = $true; break } }
        if ($variantMatch) { $score += 15 }
        elseif ($productVariants.Count -gt 0) { return -1 }
    }

    # V3.0: candidate scoring is LOCAL ONLY. No network Directory.Exists calls here.
    # 03_Mashini / HOP validation will be a separate explicit step for the small candidate set.

    # Prefer more specific folder names over generic one-word folders.
    $score += [Math]::Min(20, $folderNorm.Length / 3)
    return [int]$score
}

function Get-FastProductCandidates {
    param([string]$Product)

    $pCore = @(Get-CoreTokens $Product)
    $categoryTargets = @(Get-ExpectedCategoryTargets -Product $Product)
    if($categoryTargets.Count -eq 0){ return @() }
    $longCore = @($pCore | Where-Object { $_.Length -ge 3 })
    if ($longCore.Count -eq 0) { return @() }
    if ($script:FolderTokenMap.Count -eq 0) { Build-FolderNameMap }

    # Start from the rarest meaningful product token. This avoids scanning ~30k folders per product.
    $bestToken = $null
    $bestCount = [int]::MaxValue
    foreach ($tok in $longCore) {
        if ($script:FolderTokenMap.ContainsKey($tok)) {
            $c = $script:FolderTokenMap[$tok].Count
            if ($c -lt $bestCount) { $bestCount = $c; $bestToken = $tok }
        }
    }
    if ($null -eq $bestToken) { return @() }

    $scored = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($it in @($script:FolderTokenMap[$bestToken])) {
        if ($seen.ContainsKey($it.Path)) { continue }
        $seen[$it.Path] = $true
        if (-not (Test-PathMatchesExpectedTargets -Path ([string]$it.Path) -Targets $categoryTargets)) { continue }
        $score = Get-ProductCandidateScore -Product $Product -FolderItem $it
        if ($score -lt 0) { continue }
        $scored.Add([pscustomobject]@{ Item=$it; Score=$score })
    }
    if ($scored.Count -eq 0) { return @() }

    $sorted = @($scored.ToArray() | Sort-Object @{Expression='Score';Descending=$true}, @{Expression={$_.Item.Path};Descending=$false})
    $top = [int]$sorted[0].Score

    # Keep only candidates close to the best result. This prevents broad matches such as all 160x200 folders.
    $keepers = @($sorted | Where-Object { ([int]$_.Score) -ge ($top - 10) } | Select-Object -First 10)
    return @($keepers)
}

function Update-ProductCandidates {
    if ($null -eq $productCandidatesGrid) { return }
    if ($script:LastAllRows.Count -eq 0) { return }
    if ($script:FolderIndex.Count -eq 0) { return }
    if ($script:FolderTokenMap.Count -eq 0) { Build-FolderNameMap }

    $table = New-DataTable -Columns @('Изделие от CSV', 'Кандидати', 'Статус', 'Най-добра оценка', 'Пътища')
    $products = @($script:LastAllRows | ForEach-Object { $_.Product } | Sort-Object -Unique)
    $i = 0
    foreach ($product in $products) {
        $arr = @(Get-FastProductCandidates -Product $product)
        $state = 'НЯМА КАНДИДАТ'
        if ($arr.Count -eq 1) { $state = '1 СИЛЕН КАНДИДАТ - ЗА HOP ПРОВЕРКА' }
        elseif ($arr.Count -gt 1) { $state = 'НЯКОЛКО БЛИЗКИ КАНДИДАТА - ЗА HOP ПРОВЕРКА' }
        $paths = ($arr | ForEach-Object { "[$($_.Score)] $($_.Item.Path)" }) -join ' | '
        $bestScore = if ($arr.Count -gt 0) { [string]$arr[0].Score } else { '' }
        $r = $table.NewRow()
        $r['Изделие от CSV'] = $product
        $r['Кандидати'] = [string]$arr.Count
        $r['Статус'] = $state
        $r['Най-добра оценка'] = $bestScore
        $r['Пътища'] = $paths
        [void]$table.Rows.Add($r)
        $i++
        if (($i % 20) -eq 0) {
            $status.Text = "Локално търсене в каталога: $i / $($products.Count) ..."
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    $productCandidatesGrid.DataSource = $table
    if ($null -ne $candidateInfoLabel) {
        $candidateInfoLabel.Text = "Проверени изделия: $($products.Count). Категорията от CSV ограничава търсенето (напр. Dresser→Skrinove, Wardrobe→Garderobi), след което HOP потвърждава № + размери."
    }
}

function Stop-SearchWorkerTimer {
    if ($null -ne $script:SearchWorkerTimer) {
        try { $script:SearchWorkerTimer.Stop() } catch {}
        try { $script:SearchWorkerTimer.Dispose() } catch {}
        $script:SearchWorkerTimer = $null
    }
}

function Start-BackgroundProductSearch {
    if ($script:LastAllRows.Count -eq 0) { throw 'Първо натиснете Прочети CSV.' }
    if (-not [System.IO.File]::Exists($script:IndexCsv)) { throw 'Няма запазен каталог на папките. Натиснете Обнови каталог (фон).' }
    if (-not [System.IO.File]::Exists($script:SearchWorkerScript)) { throw "Липсва помощният файл: $($script:SearchWorkerScript)" }
    if ($null -ne $script:SearchWorkerProcess -and -not $script:SearchWorkerProcess.HasExited) { return }
    Test-OperatorStopRequested

    $products = @($script:LastAllRows | ForEach-Object { $_.Product } | Sort-Object -Unique)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Product')
    foreach ($p in $products) { $lines.Add((Convert-ToCsvField $p)) }
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($script:SearchProductsCsv, $lines.ToArray(), $utf8Bom)
    try { if ([System.IO.File]::Exists($script:SearchResultsCsv)) { [System.IO.File]::Delete($script:SearchResultsCsv) } } catch {}
    try { if ([System.IO.File]::Exists($script:SearchProgressFile)) { [System.IO.File]::Delete($script:SearchProgressFile) } } catch {}

    $argLine = "-NoProfile -ExecutionPolicy Bypass -File `"$($script:SearchWorkerScript)`" -ProductsCsv `"$($script:SearchProductsCsv)`" -IndexCsv `"$($script:IndexCsv)`" -ResultsCsv `"$($script:SearchResultsCsv)`" -ProgressFile `"$($script:SearchProgressFile)`" -CategoriesCsv `"$($script:ProductCategoriesCsv)`""
    try{Set-OperatorCheckpoint -CurrentStage 'SEARCH_RUNNING' -Status 'ACTIVE' -Message 'Търсене на продуктови папки.'}catch{}
    $script:SearchWorkerProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -WindowStyle Hidden -PassThru
    $findProductsButton.Enabled = $false
    $status.Text = 'Търсенето на изделия стартира във фон. Интерфейсът остава свободен.'

    Stop-SearchWorkerTimer
    $script:SearchWorkerTimer = New-Object System.Windows.Forms.Timer
    $script:SearchWorkerTimer.Interval = 500
    $script:SearchWorkerTimer.Add_Tick({
        try {
            if ([System.IO.File]::Exists($script:SearchProgressFile)) {
                $progress = [System.IO.File]::ReadAllText($script:SearchProgressFile, [System.Text.Encoding]::UTF8).Trim()
                if (-not [string]::IsNullOrWhiteSpace($progress)) { $status.Text = $progress }
            }
        } catch {}

        if ($null -ne $script:SearchWorkerProcess -and $script:SearchWorkerProcess.HasExited) {
            Stop-SearchWorkerTimer
            $exitCode = $script:SearchWorkerProcess.ExitCode
            $script:SearchWorkerProcess.Dispose()
            $script:SearchWorkerProcess = $null
            $findProductsButton.Enabled = $true
            if ($exitCode -eq 0 -and [System.IO.File]::Exists($script:SearchResultsCsv)) {
                try {
                    $rows = @(Read-SafeCsv -Path $script:SearchResultsCsv)
                    $objects = New-Object System.Collections.Generic.List[object]
                    foreach ($x in $rows) {
                        $objects.Add([pscustomobject]@{
                            Product = Get-PropertyValue -Object $x -Name 'Product'
                            Candidates = Get-PropertyValue -Object $x -Name 'Candidates'
                            Status = Get-PropertyValue -Object $x -Name 'Status'
                            BestScore = Get-PropertyValue -Object $x -Name 'BestScore'
                            Paths = Get-PropertyValue -Object $x -Name 'Paths'
                        })
                    }
                    $script:SearchResultObjects = @($objects.ToArray())
                    Refresh-CandidateGridFromSearchResults
                    $candidateInfoLabel.Text = "Проверени изделия: $($rows.Count). Категорията е твърд филтър. HOP проверката е за всички непотвърдени кандидати."
                    $tabs.SelectedTab = $tabCandidates
                    $status.Text = "Фоновото търсене е готово. Изделия: $($rows.Count). Следваща стъпка: HOP проверка на всички непотвърдени кандидати."
                    try{Set-OperatorCheckpoint -CurrentStage 'SEARCH_DONE' -LastSafeStage 'SEARCH_DONE' -Status 'ACTIVE' -Message ('Търсенето е готово. Изделия: '+$rows.Count)}catch{}
                    try{Refresh-OperatorDashboard}catch{}
                    try{Continue-OperatorAfterSearch}catch{}
                }
                catch { $status.Text = "Търсенето приключи, но зареждането на резултата даде грешка: $($_.Exception.Message)";try{Set-OperatorRunError -Message $status.Text}catch{};Set-CurrentCsvBatchRegistryError -Message $status.Text;$script:OperatorAutoFlow=$false;$script:OperatorRunHopAfterSearch=$false;$script:OperatorRunExecuteAfterHop=$false }
            }
            else {
                $detail = ''
                try { if ([System.IO.File]::Exists($script:SearchProgressFile)) { $detail = [System.IO.File]::ReadAllText($script:SearchProgressFile, [System.Text.Encoding]::UTF8).Trim() } } catch {}
                $status.Text = "Грешка при фоновото търсене. $detail";try{Set-OperatorRunError -Message $status.Text}catch{};Set-CurrentCsvBatchRegistryError -Message $status.Text;$script:OperatorAutoFlow=$false;$script:OperatorRunHopAfterSearch=$false;$script:OperatorRunExecuteAfterHop=$false
            }
        }
    })
    $script:SearchWorkerTimer.Start()
}



function Apply-HopGridVisualStyles {
    if ($null -eq $hopGrid) { return }

    foreach ($gridRow in $hopGrid.Rows) {
        try {
            # Reset first so rebinding/sorting cannot leave a colour on the wrong row.
            $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
            $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            $gridRow.DefaultCellStyle.Font = $hopGrid.Font

            $d = ([string]$gridRow.Cells['Решение'].Value).Trim()

            # Match by key wording, not the whole exact string, so harmless text/encoding changes
            # cannot remove the visual distinction.
            if ($d -like '*АВТОМАТИЧНО ИЗБРАН*') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
                $gridRow.DefaultCellStyle.Font = New-Object System.Drawing.Font($hopGrid.Font,[System.Drawing.FontStyle]::Bold)
            } elseif ($d -like '*ОТХВЪРЛЕН*') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::Gainsboro
            } elseif ($d -like '*РЪЧНО*' -or $d -like '*НЯМА ПЪЛНО HOP СЪВПАДЕНИЕ*') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
            }
        } catch {}
    }

    try { $hopGrid.Invalidate() } catch {}
}

function Refresh-HopGrid {
    if ($null -eq $hopGrid) { return }
    $table = New-DataTable -Columns @('Изделие','Кандидат','Оценка','CSV позиции','Пълни','Имена','Размери','HOP релевантни','HOP разчетени','HOP грешки','Статус','Решение')
    foreach ($h in @($script:HopResultObjects | Sort-Object Product,CandidatePath)) {
        $r=$table.NewRow()
        $r['Изделие']=[string]$h.Product
        $r['Кандидат']=[string]$h.CandidatePath
        $r['Оценка']=[string]$h.SearchScore
        $r['CSV позиции']=[string]$h.ExpectedDetails
        $r['Пълни']=[string]$h.MatchedFull
        $r['Имена']=[string]$h.MatchedName
        $r['Размери']=[string]$h.MatchedDimensions
        $r['HOP релевантни']=[string]$h.RelevantHopFiles
        $r['HOP разчетени']=[string]$h.ReadableHopFiles
        $r['HOP грешки']=[string]$h.UnreadableHopFiles
        $r['Статус']=[string]$h.CandidateStatus
        $r['Решение']=[string]$h.Decision
        [void]$table.Rows.Add($r)
    }
    $hopGrid.DataSource=$table
    # Visual distinction: automatic = green, rejected = gray, manual/no-full-match = yellow.
    # Reapplied after DataBindingComplete/Sort below because WinForms may repaint after binding.
    Apply-HopGridVisualStyles

    if ($null -ne $hopInfoLabel) {
        $products=@($script:HopResultObjects | ForEach-Object {$_.Product} | Sort-Object -Unique)
        $auto=@($script:HopResultObjects | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_.AutoPath)} | ForEach-Object {$_.Product} | Sort-Object -Unique)
        $manualCount=[Math]::Max(0,$products.Count-$auto.Count)

        $hopInfoLabel.Text="HOP проверени спорни изделия: $($products.Count) | автоматично потвърдени: $($auto.Count) | за ръчно: $manualCount. Четат се само .hop от 03_Mashini\\1_Holzher."
    }
}

function Stop-HopWorkerTimer {
    if ($null -ne $script:HopWorkerTimer) {
        try {$script:HopWorkerTimer.Stop()}catch{}
        try {$script:HopWorkerTimer.Dispose()}catch{}
        $script:HopWorkerTimer=$null
    }
}

function Start-BackgroundHopCheck {
    if ($script:LastAllRows.Count -eq 0) { throw 'Първо натиснете Прочети CSV.' }
    if ($script:SearchResultObjects.Count -eq 0 -or -not [System.IO.File]::Exists($script:SearchResultsCsv)) { throw 'Първо натиснете Търси изделия.' }
    if (-not [System.IO.File]::Exists($script:HopWorkerScript)) { throw "Липсва помощният файл: $($script:HopWorkerScript)" }
    if ($null -ne $script:HopWorkerProcess -and -not $script:HopWorkerProcess.HasExited) { return }

    $ambiguous=@($script:SearchResultObjects | Where-Object {
        $product=[string]$_.Product
        if($script:ProductFolderMap.ContainsKey($product)){
            $src=[string]$script:ProductFolderMap[$product].Source
            return (-not (Test-IsTrustedProductMappingSource -Source $src))
        }
        $n=0; return ([int]::TryParse([string]$_.Candidates,[ref]$n) -and $n -ge 1)
    })
    if ($ambiguous.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Няма изделия, които чакат HOP проверка. Директориите от Бърза миграция и вече HOP-потвърдените се пропускат.','PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
        return
    }

    foreach($f in @($script:HopResultsCsv,$script:HopProgressFile)){try{if([System.IO.File]::Exists($f)){[System.IO.File]::Delete($f)}}catch{}}
    $argLine="-NoProfile -ExecutionPolicy Bypass -File `"$($script:HopWorkerScript)`" -RowsCsv `"$($script:CsvRowsOut)`" -SearchResultsCsv `"$($script:SearchResultsCsv)`" -ResultsCsv `"$($script:HopResultsCsv)`" -ProgressFile `"$($script:HopProgressFile)`" -MappingsCsv `"$($script:ProductFolderMapCsv)`""
    Test-OperatorStopRequested
    try{Set-OperatorCheckpoint -CurrentStage 'HOP_RUNNING' -Status 'ACTIVE' -Message 'HOP проверка на непотвърдените кандидати.'}catch{}
    $script:HopWorkerProcess=Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -WindowStyle Hidden -PassThru
    $hopCheckButton.Enabled=$false
    $status.Text="HOP проверката стартира във фон за $($ambiguous.Count) изделия. Проверяват се и 1-кандидатните и ръчно променените папки; Бърза миграция се пропуска."

    Stop-HopWorkerTimer
    $script:HopWorkerTimer=New-Object System.Windows.Forms.Timer
    $script:HopWorkerTimer.Interval=500
    $script:HopWorkerTimer.Add_Tick({
        try{if([System.IO.File]::Exists($script:HopProgressFile)){$p=[System.IO.File]::ReadAllText($script:HopProgressFile,[System.Text.Encoding]::UTF8).Trim();if(-not [string]::IsNullOrWhiteSpace($p)){$status.Text=$p}}}catch{}
        if($null -ne $script:HopWorkerProcess -and $script:HopWorkerProcess.HasExited){
            Stop-HopWorkerTimer
            $code=$script:HopWorkerProcess.ExitCode
            $script:HopWorkerProcess.Dispose();$script:HopWorkerProcess=$null
            $hopCheckButton.Enabled=$true
            if($code -eq 0 -and [System.IO.File]::Exists($script:HopResultsCsv)){
                try{
                    $rows=@(Read-SafeCsv -Path $script:HopResultsCsv)
                    $objs=New-Object System.Collections.Generic.List[object]
                    foreach($x in $rows){
                        $objs.Add([pscustomobject]@{
                            Product=Get-PropertyValue $x 'Product'
                            CandidatePath=Get-PropertyValue $x 'CandidatePath'
                            SearchScore=Get-PropertyValue $x 'SearchScore'
                            ExpectedDetails=Get-PropertyValue $x 'ExpectedDetails'
                            MatchedFull=Get-PropertyValue $x 'MatchedFull'
                            MatchedName=Get-PropertyValue $x 'MatchedName'
                            MatchedDimensions=Get-PropertyValue $x 'MatchedDimensions'
                            RelevantHopFiles=Get-PropertyValue $x 'RelevantHopFiles'
                            ReadableHopFiles=Get-PropertyValue $x 'ReadableHopFiles'
                            UnreadableHopFiles=Get-PropertyValue $x 'UnreadableHopFiles'
                            CandidateStatus=Get-PropertyValue $x 'CandidateStatus'
                            Decision=Get-PropertyValue $x 'Decision'
                            AutoPath=Get-PropertyValue $x 'AutoPath'
                        })
                    }
                    $script:HopResultObjects=@($objs.ToArray())
                    $autoProducts=@($script:HopResultObjects | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_.AutoPath)} | Group-Object Product)
                    $autoCount=0
                    foreach($g in $autoProducts){
                        $product=[string]$g.Name
                        $path=[string]$g.Group[0].AutoPath
                        if(-not ((Test-IsUnderAllowedRoot -Path $path) -and [System.IO.Directory]::Exists($path))){continue}
                        if($script:ProductFolderMap.ContainsKey($product)){
                            $existing=$script:ProductFolderMap[$product]
                            $src=[string]$existing.Source
                            $existingPath=[string]$existing.Path
                            if(Test-IsTrustedProductMappingSource -Source $src){continue}
                            if((Normalize-PathForCompare $existingPath) -eq (Normalize-PathForCompare $path)){
                                Set-ProductFolderMapping -Product $product -Path $path -Source 'HOP_VERIFIED_MANUAL'
                                $autoCount++
                            }
                        } else {
                            Set-ProductFolderMapping -Product $product -Path $path -Source 'HOP_AUTO'
                            $autoCount++
                        }
                    }
                    Refresh-HopGrid
                    Refresh-CandidateGridFromSearchResults
                    $tabs.SelectedTab=$tabHop
                    $status.Text="HOP проверката е готова. Автоматично потвърдени: $autoCount. Останалите са за ръчно потвърждение."
                    try{Set-OperatorCheckpoint -CurrentStage 'HOP_DONE' -LastSafeStage 'HOP_DONE' -Status 'ACTIVE' -Message ('HOP проверката е готова. Автоматично потвърдени: '+$autoCount)}catch{}
                    try{Continue-OperatorAfterHop}catch{}
                }catch{$status.Text="HOP проверката приключи, но зареждането на резултата даде грешка: $($_.Exception.Message)";try{Set-OperatorRunError -Message $status.Text}catch{};Set-CurrentCsvBatchRegistryError -Message $status.Text;$script:OperatorAutoFlow=$false;$script:OperatorRunExecuteAfterHop=$false}
            }else{
                $detail='';try{if([System.IO.File]::Exists($script:HopProgressFile)){$detail=[System.IO.File]::ReadAllText($script:HopProgressFile,[System.Text.Encoding]::UTF8).Trim()}}catch{}
                $status.Text="Грешка при HOP проверката. $detail";try{Set-OperatorRunError -Message $status.Text}catch{};Set-CurrentCsvBatchRegistryError -Message $status.Text;$script:OperatorAutoFlow=$false;$script:OperatorRunExecuteAfterHop=$false
            }
        }
    })
    $script:HopWorkerTimer.Start()
}


function Apply-PgmxGridVisualStyles {
    if ($null -eq $pgmxGrid) { return }
    foreach ($gridRow in $pgmxGrid.Rows) {
        try {
            $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
            $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            $gridRow.DefaultCellStyle.Font = $pgmxGrid.Font
            $result = ([string]$gridRow.Cells['Резултат'].Value).Trim()
            $action = ([string]$gridRow.Cells['Действие'].Value).Trim()
            $central = ([string]$gridRow.Cells['SCM_PWX_500'].Value).Trim()
            if ($central -eq 'ГРЕШКА') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            } elseif ($result -eq 'ГОТОВ' -or $result -eq 'СЪЗДАДЕН ProgramCodeF' -or $result -eq 'КОНВЕРТИРАН + ProgramCodeF') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
            } elseif ($result -eq 'ГОТОВ ЗА СЪЗДАВАНЕ') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCyan
            } elseif ($result -eq 'ГОТОВ ЗА КОНВЕРСИЯ' -or $result -eq 'НЯМА ПОТВЪРДЕНА ПАПКА') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
            } elseif ($result -like 'НЕЕДНОЗНАЧЕН*' -or $result -like 'НЕЕДНОЗНАЧНА*' -or $result -eq 'НЯМА ВАЛИДЕН HOP' -or $result -eq 'НЯМА ВАЛИДНА HOP ПАПКА' -or $action -eq 'РЪЧНА ПРОВЕРКА') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::Moccasin
            } elseif ($result -eq 'ГРЕШКА') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            }
        } catch {}
    }
    try {
        $pgmxGrid.AllowUserToResizeColumns = $true
        foreach($c in $pgmxGrid.Columns){
            try {
                if($c.AutoSizeMode -ne [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None){
                    $w=[Math]::Max(50,[int]$c.Width)
                    $c.AutoSizeMode=[System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
                    $c.Width=$w
                }
            } catch {}
        }
    } catch {}
    try { $pgmxGrid.Invalidate() } catch {}
}

function Refresh-PgmxGrid {
    if ($null -eq $pgmxGrid) { return }
    $table = New-DataTable -Columns @('CSV','Изделие','№','Детайл','Обработка','Суфикс','ProgramCodeF','HOP източник','Основен PGMX','Целеви PGMX','2_PWX_500','HOP статус','Основен статус','Целеви статус','SCM_PWX_500','Централна грешка','Валидация','Действие','Резултат','Грешка','Папка изделие')
    foreach ($x in @($script:PgmxResultObjects)) {
        $r=$table.NewRow()
        $r['CSV']=[string]$x.CSV
        $r['Изделие']=[string]$x.Product
        $r['№']=[string]$x.DetailNo
        $r['Детайл']=[string]$x.DetailName
        $r['Обработка']=[string]$x.ProcessingNo
        $r['Суфикс']=[string]$x.ExpectedSuffix
        $r['ProgramCodeF']=[string]$x.ProgramCodeF
        $r['HOP източник']=[string]$x.SourceHop
        $r['Основен PGMX']=[string]$x.BasePgmx
        $r['Целеви PGMX']=[string]$x.TargetPgmx
        $r['2_PWX_500']=[string]$x.FolderStatus
        $r['HOP статус']=[string]$x.HopStatus
        $r['Основен статус']=[string]$x.BaseStatus
        $r['Целеви статус']=[string]$x.TargetStatus
        $r['SCM_PWX_500']=[string]$x.CentralStatus
        $r['Централна грешка']=[string]$x.CentralError
        $r['Валидация']=[string]$x.Validation
        $r['Действие']=[string]$x.Action
        $r['Резултат']=[string]$x.Result
        $r['Грешка']=[string]$x.Error
        $r['Папка изделие']=[string]$x.ProductPath
        [void]$table.Rows.Add($r)
    }
    $pgmxGrid.DataSource=$table
    Apply-PgmxGridVisualStyles
    if ($null -ne $pgmxInfoLabel) {
        $total=@($script:PgmxResultObjects).Count
        $ready=@($script:PgmxResultObjects | Where-Object { $_.Result -eq 'ГОТОВ' }).Count
        $created=@($script:PgmxResultObjects | Where-Object { $_.Result -eq 'СЪЗДАДЕН ProgramCodeF' }).Count
        $converted=@($script:PgmxResultObjects | Where-Object { $_.Result -eq 'КОНВЕРТИРАН + ProgramCodeF' }).Count
        $copy=@($script:PgmxResultObjects | Where-Object { $_.Result -eq 'ГОТОВ ЗА СЪЗДАВАНЕ' }).Count
        $conv=@($script:PgmxResultObjects | Where-Object { $_.Result -eq 'ГОТОВ ЗА КОНВЕРСИЯ' }).Count
        $manual=@($script:PgmxResultObjects | Where-Object { $_.Action -eq 'РЪЧНА ПРОВЕРКА' -or $_.Result -eq 'НЯМА ПОТВЪРДЕНА ПАПКА' }).Count
        $errors=@($script:PgmxResultObjects | Where-Object { $_.Result -eq 'ГРЕШКА' }).Count
        $centralCopied=@($script:PgmxResultObjects | Where-Object { $_.CentralStatus -eq 'КОПИРАН' }).Count
        $centralExists=@($script:PgmxResultObjects | Where-Object { $_.CentralStatus -eq 'СЪЩЕСТВУВА' }).Count
        $centralPlanned=@($script:PgmxResultObjects | Where-Object { $_.CentralStatus -eq 'ЩЕ СЕ КОПИРА' }).Count
        $centralErrors=@($script:PgmxResultObjects | Where-Object { $_.CentralStatus -eq 'ГРЕШКА' }).Count
        $centralNotRequired=@($script:PgmxResultObjects | Where-Object { $_.CentralStatus -eq 'НЕ СЕ ИЗИСКВА' }).Count
        $pgmxInfoLabel.Text="PGMX V5.1.2 $($script:PgmxLastMode): общо $total | готови $ready | копирани $created | конвертирани $converted | за копиране $copy | за конверсия $conv | ръчни $manual | грешки $errors | SCM специални: копирани $centralCopied / налични $centralExists / за копиране $centralPlanned / нормални $centralNotRequired / грешки $centralErrors"
    }
}

function Stop-PgmxWorkerTimer {
    if ($null -ne $script:PgmxWorkerTimer) {
        try { $script:PgmxWorkerTimer.Stop() } catch {}
        try { $script:PgmxWorkerTimer.Dispose() } catch {}
        $script:PgmxWorkerTimer=$null
    }
}

function Get-PgmxAllTestStats {
    $rows=@($script:LastAllRows | Where-Object { -not [string]::IsNullOrWhiteSpace(([string]$_.ProgramCodeF).Trim()) })
    $products=@($rows | ForEach-Object { ([string]$_.Product).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    return [pscustomobject]@{ Rows=$rows.Count; Products=$products.Count }
}

function Start-BackgroundPgmx {
    param([ValidateSet('Check','Execute')][string]$Mode)
    if ($script:LastAllRows.Count -eq 0) { throw 'Първо натиснете Прочети CSV.' }
    if (-not [System.IO.File]::Exists($script:PgmxWorkerScript)) { throw "Липсва помощният файл: $($script:PgmxWorkerScript)" }
    if (-not [System.IO.File]::Exists($script:CsvRowsOut)) { throw 'Липсва локалният резултат от CSV. Натиснете Прочети CSV.' }
    if (-not [System.IO.File]::Exists($script:SearchResultsCsv)) { throw 'Първо натиснете Търси изделия.' }
    if ($null -ne $script:PgmxWorkerProcess -and -not $script:PgmxWorkerProcess.HasExited) { return }
    Test-OperatorStopRequested

    if ($Mode -eq 'Execute') {
        if (-not [System.IO.File]::Exists($script:ConverterExe)) { throw 'Липсва вграденият HOP_to_PGMX_2_9.exe.' }
        # A new daily Execute must never reuse the list from an earlier run in the same UI session.
        $script:CreatedPgmxThisRun = @()
    }

    try { Save-AllowedFolders } catch {}
    try { Save-ProductFolderMappings } catch {}
    foreach($f in @($script:PgmxResultsCsv,$script:PgmxProgressFile,$script:PgmxErrorFile,$script:PgmxDoneFile)){try{if([System.IO.File]::Exists($f)){[System.IO.File]::Delete($f)}}catch{}}
    $argLine="-NoProfile -ExecutionPolicy Bypass -File `"$($script:PgmxWorkerScript)`" -RowsCsv `"$($script:CsvRowsOut)`" -MappingsCsv `"$($script:ProductFolderMapCsv)`" -SearchResultsCsv `"$($script:SearchResultsCsv)`" -SettingsCsv `"$($script:SettingsCsv)`" -ResultsCsv `"$($script:PgmxResultsCsv)`" -ProgressFile `"$($script:PgmxProgressFile)`" -DoneFile `"$($script:PgmxDoneFile)`" -Mode $Mode -SettingsVersion `"$($script:SettingsVersion)`" -ConverterExe `"$($script:ConverterExe)`" -ConversionLogCsv `"$($script:ConversionLogCsv)`""
    if($Mode -eq 'Execute'){try{Set-OperatorCheckpoint -CurrentStage 'PGMX_RUNNING' -Status 'ACTIVE' -Message 'Създаване/конвертиране на PGMX.'}catch{}}
    $script:PgmxWorkerProcess=Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -WindowStyle Hidden -RedirectStandardError $script:PgmxErrorFile -PassThru
    $script:PgmxLastMode=$Mode
    $pgmxCheckButton.Enabled=$false
    $pgmxExecuteButton.Enabled=$false
    if($Mode -eq 'Check'){$status.Text='PGMX V5.1.2 TEST: преглед на ВСИЧКИ изделия. Няма промени.'}else{$status.Text='PGMX V5.1.2 TEST: обработват се ВСИЧКИ изделия с потвърдена папка; проблемните се пропускат и се отчитат.'}

    Stop-PgmxWorkerTimer
    $script:PgmxWorkerTimer=New-Object System.Windows.Forms.Timer
    $script:PgmxWorkerTimer.Interval=500
    $script:PgmxWorkerTimer.Add_Tick({
        try{if([System.IO.File]::Exists($script:PgmxProgressFile)){$p=[System.IO.File]::ReadAllText($script:PgmxProgressFile,[System.Text.Encoding]::UTF8).Trim();if(-not [string]::IsNullOrWhiteSpace($p)){$status.Text=$p}}}catch{}
        if($null -ne $script:PgmxWorkerProcess -and $script:PgmxWorkerProcess.HasExited){
            Stop-PgmxWorkerTimer
            try { $script:PgmxWorkerProcess.WaitForExit() } catch {}
            $code=$null
            try { $code=$script:PgmxWorkerProcess.ExitCode } catch { $code=-999 }
            $script:PgmxWorkerProcess.Dispose();$script:PgmxWorkerProcess=$null
            $pgmxCheckButton.Enabled=$true;$pgmxExecuteButton.Enabled=$true

            $doneText=''
            try { if([System.IO.File]::Exists($script:PgmxDoneFile)){ $doneText=[System.IO.File]::ReadAllText($script:PgmxDoneFile,[System.Text.Encoding]::UTF8).Trim() } } catch {}
            $hasResults=[System.IO.File]::Exists($script:PgmxResultsCsv)
            $workerOk=$hasResults -and (($doneText -like 'OK*') -or $code -eq 0)
            if($hasResults){ $workerOk=$true }

            if($workerOk){
                try{
                    $rows=@(Read-SafeCsv -Path $script:PgmxResultsCsv)
                    $objs=New-Object System.Collections.Generic.List[object]
                    foreach($x in $rows){
                        $objs.Add([pscustomobject]@{
                            CSV=Get-PropertyValue $x 'CSV';Product=Get-PropertyValue $x 'Product';DetailNo=Get-PropertyValue $x 'DetailNo';DetailName=Get-PropertyValue $x 'DetailName';ProcessingNo=Get-PropertyValue $x 'ProcessingNo';ExpectedSuffix=Get-PropertyValue $x 'ExpectedSuffix';ProgramCodeF=Get-PropertyValue $x 'ProgramCodeF';ProductPath=Get-PropertyValue $x 'ProductPath';PwxFolder=Get-PropertyValue $x 'PwxFolder';SourceHop=Get-PropertyValue $x 'SourceHop';BasePgmx=Get-PropertyValue $x 'BasePgmx';TargetPgmx=Get-PropertyValue $x 'TargetPgmx';FolderStatus=Get-PropertyValue $x 'FolderStatus';HopStatus=Get-PropertyValue $x 'HopStatus';BaseStatus=Get-PropertyValue $x 'BaseStatus';TargetStatus=Get-PropertyValue $x 'TargetStatus';Validation=Get-PropertyValue $x 'Validation';Action=Get-PropertyValue $x 'Action';Result=Get-PropertyValue $x 'Result';Error=Get-PropertyValue $x 'Error';CentralFolder=Get-PropertyValue $x 'CentralFolder';CentralStatus=Get-PropertyValue $x 'CentralStatus';CentralError=Get-PropertyValue $x 'CentralError'
                        })
                    }
                    $script:PgmxResultObjects=@($objs.ToArray())
                    # V5.1.2: resolve and store the exact full paths created in THIS Execute.
                    # The worker returns TargetPgmx as a file name, while PwxFolder is the full folder.
                    # Earlier builds incorrectly called File.Exists() on the bare file name and therefore
                    # reported 0 new PGMX even when the grid correctly showed COPIED/CONVERTED rows.
                    if ($script:PgmxLastMode -eq 'Execute') {
                        try { Capture-CreatedPgmxThisRun | Out-Null } catch { $script:CreatedPgmxThisRun=@() }
                        try{Set-OperatorCheckpoint -CurrentStage 'PGMX_DONE' -LastSafeStage 'PGMX_DONE' -Status 'ACTIVE' -Message ('PGMX обработката е готова. Резултати: '+$script:PgmxResultObjects.Count)}catch{}
                    }
                    Refresh-PgmxGrid
                    $tabs.SelectedTab=$tabPgmx
                    $status.Text="PGMX V5.1.2 $($script:PgmxLastMode) приключи. Заредени: $($script:PgmxResultObjects.Count) ProgramCodeF."
                    try { Update-AutoMixUi } catch {}
                    try { Refresh-OperatorDashboard } catch {}
                    if ($script:PgmxLastMode -eq 'Execute') {
                        try{Refresh-OperatorReports -WriteFiles}catch{}
                        if($script:OperatorExecutionMode -eq 'PGMX_ONLY'){
                            try { Finish-PgmxOnlyStage } catch {
                                $status.Text='PGMX са готови, но запазването на дневния регистър не успя: '+$_.Exception.Message
                                Set-OperatorRunError -Message $status.Text
                                [System.Windows.Forms.MessageBox]::Show($status.Text,'CAM Agent - PGMX',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
                            }
                        } else { Invoke-AutoMixAfterStep3 }
                    }
                }catch{
                    $status.Text="PGMX приключи, но зареждането на резултата даде грешка: $($_.Exception.Message)"
                    if($script:PgmxLastMode -eq 'Execute'){try{Finalize-CsvRegistryForCurrentBatch -PipelineSucceeded $false -PipelineMessage $status.Text}catch{}}
                }
            }else{
                $progress='';$stderr=''
                try{if([System.IO.File]::Exists($script:PgmxProgressFile)){$progress=[System.IO.File]::ReadAllText($script:PgmxProgressFile,[System.Text.Encoding]::UTF8).Trim()}}catch{}
                try{if([System.IO.File]::Exists($script:PgmxErrorFile)){$stderr=[System.IO.File]::ReadAllText($script:PgmxErrorFile).Trim()}}catch{}
                $parts=New-Object System.Collections.Generic.List[string]
                $parts.Add("ExitCode=$code")
                if(-not [string]::IsNullOrWhiteSpace($doneText)){$parts.Add("Done=$doneText")}
                if(-not [string]::IsNullOrWhiteSpace($progress)){$parts.Add($progress)}
                if(-not [string]::IsNullOrWhiteSpace($stderr)){$parts.Add($stderr)}
                $status.Text='Грешка при PGMX worker. '+($parts -join ' | ')
                if($script:PgmxLastMode -eq 'Execute'){try{Finalize-CsvRegistryForCurrentBatch -PipelineSucceeded $false -PipelineMessage $status.Text}catch{}}
            }
        }
    })
    $script:PgmxWorkerTimer.Start()
}

function Set-ManualProductFolderFromCandidates {
    # V5.1.12: manual correction is allowed for ANY selected product in the
    # "Изделия - папки" grid, including a product with exactly one strong
    # candidate. The manual mapping always overrides the automatic candidate.
    if ($null -eq $productCandidatesGrid -or $null -eq $productCandidatesGrid.CurrentRow) {
        [System.Windows.Forms.MessageBox]::Show('Избери изделие от таблицата „Изделия - папки“.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }

    $product = Safe-Trim $productCandidatesGrid.CurrentRow.Cells['Изделие от CSV'].Value
    if ([string]::IsNullOrWhiteSpace($product)) {
        [System.Windows.Forms.MessageBox]::Show('Не може да се определи избраното изделие.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $initialPath = ''
    if ($script:ProductFolderMap.ContainsKey($product)) {
        try {
            $mapped = Safe-Trim $script:ProductFolderMap[$product].Path
            if (-not [string]::IsNullOrWhiteSpace($mapped) -and [System.IO.Directory]::Exists($mapped)) { $initialPath = $mapped }
        } catch {}
    }

    if ([string]::IsNullOrWhiteSpace($initialPath)) {
        $sr = Get-SearchResultForProduct -Product $product
        if ($null -ne $sr) {
            $cands = @(Parse-CandidatePaths $sr.Paths)
            foreach ($c in $cands) {
                if (-not [string]::IsNullOrWhiteSpace([string]$c.Path) -and [System.IO.Directory]::Exists([string]$c.Path)) {
                    $initialPath = [string]$c.Path
                    break
                }
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($initialPath)) {
        foreach ($root in @($script:AllowedFolders.ToArray())) {
            if ([System.IO.Directory]::Exists($root)) { $initialPath = $root; break }
        }
    }

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Промени папката за изделието:`r`n$product`r`n`r`nМоже да избереш друга съществуваща папка. Ако е извън разрешените корени, агентът ще поиска изрично потвърждение."
    $dialog.ShowNewFolderButton = $false
    if (-not [string]::IsNullOrWhiteSpace($initialPath)) { $dialog.SelectedPath = $initialPath }

    try {
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $path = Safe-Trim $dialog.SelectedPath
        if ([string]::IsNullOrWhiteSpace($path)) { return }
        if (-not [System.IO.Directory]::Exists($path)) {
            [System.Windows.Forms.MessageBox]::Show("Папката не съществува или не е достъпна:`r`n$path", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            return
        }
        if (-not (Test-IsUnderAllowedRoot -Path $path)) {
            # V5.1.31: allowed roots remain a hard filter for AUTOMATIC search,
            # but an explicit manual decision by the user may override them.
            $ansOutside = [System.Windows.Forms.MessageBox]::Show("Избраната папка е извън разрешените корени:`r`n$path`r`n`r`nТова е ръчен избор. Сигурни ли сте, че искате да използвате тази папка?", 'PWX500 Agent - Ръчно потвърждение', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($ansOutside -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        $machineDir = Join-Path $path '03_Mashini'
        if (-not [System.IO.Directory]::Exists($machineDir)) {
            $ans = [System.Windows.Forms.MessageBox]::Show("В избраната папка не е намерена директна подпапка 03_Mashini:`r`n$path`r`n`r`nДа се запише ли въпреки това?", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        Set-ProductFolderMapping -Product $product -Path $path -Source 'ManualOverrideCandidates'
        $status.Text = "Ръчно променена папка: $product -> $path | Нужна е HOP проверка преди PGMX."
        try{Refresh-OperatorDashboard}catch{}
        Refresh-CandidateGridFromSearchResults
        $tabs.SelectedTab = $tabCandidates

        try {
            foreach ($row in $productCandidatesGrid.Rows) {
                if ((Safe-Trim $row.Cells['Изделие от CSV'].Value) -eq $product) {
                    $row.Selected = $true
                    $productCandidatesGrid.CurrentCell = $row.Cells['Изделие от CSV']
                    try { $productCandidatesGrid.FirstDisplayedScrollingRowIndex = $row.Index } catch {}
                    break
                }
            }
        } catch {}
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при промяна на папката:`r`n$($_.Exception.Message)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    finally {
        $dialog.Dispose()
    }
}

function Set-ManualProductFolderFromOperatorGrid {
    if($null -eq $script:MainSkippedGrid -or $null -eq $script:MainSkippedGrid.CurrentRow){
        [System.Windows.Forms.MessageBox]::Show('Избери изделие от таблицата „Непотвърдени / пропуснати“.','CAM Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
        return
    }
    $product=Safe-Trim $script:MainSkippedGrid.CurrentRow.Cells['Изделие'].Value
    if([string]::IsNullOrWhiteSpace($product)){return}

    $initialPath=''
    if($script:ProductFolderMap.ContainsKey($product)){
        try{$mp=Safe-Trim $script:ProductFolderMap[$product].Path;if(-not [string]::IsNullOrWhiteSpace($mp) -and [System.IO.Directory]::Exists($mp)){$initialPath=$mp}}catch{}
    }
    if([string]::IsNullOrWhiteSpace($initialPath)){
        $sr=Get-SearchResultForProduct -Product $product
        if($null -ne $sr){
            foreach($c in @(Parse-CandidatePaths $sr.Paths)){
                $cp=Safe-Trim $c.Path
                if(-not [string]::IsNullOrWhiteSpace($cp) -and [System.IO.Directory]::Exists($cp)){$initialPath=$cp;break}
            }
        }
    }
    if([string]::IsNullOrWhiteSpace($initialPath)){
        foreach($root in @($script:AllowedFolders.ToArray())){if([System.IO.Directory]::Exists($root)){$initialPath=$root;break}}
    }

    $dialog=New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description="Посочи точната папка за:`r`n$product`r`n`r`nСлед избора изделието остава непотвърдено, докато не мине HOP проверката."
    $dialog.ShowNewFolderButton=$false
    if(-not [string]::IsNullOrWhiteSpace($initialPath)){$dialog.SelectedPath=$initialPath}
    try{
        if($dialog.ShowDialog($script:OperatorMainForm) -ne [System.Windows.Forms.DialogResult]::OK){return}
        $path=Safe-Trim $dialog.SelectedPath
        if([string]::IsNullOrWhiteSpace($path)){return}
        if(-not [System.IO.Directory]::Exists($path)){
            [System.Windows.Forms.MessageBox]::Show("Папката не съществува или не е достъпна:`r`n$path",'CAM Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
            return
        }
        if(-not (Test-IsUnderAllowedRoot -Path $path)){
            $ansOutside=[System.Windows.Forms.MessageBox]::Show("Избраната папка е извън разрешените корени:`r`n$path`r`n`r`nТова е ръчен избор. Сигурни ли сте, че искате да я използвате?",'CAM Agent - Ръчно потвърждение',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
            if($ansOutside -ne [System.Windows.Forms.DialogResult]::Yes){return}
        }
        $machineDir=Join-Path $path '03_Mashini'
        if(-not [System.IO.Directory]::Exists($machineDir)){
            $ans=[System.Windows.Forms.MessageBox]::Show("В избраната папка не е намерена директна подпапка 03_Mashini:`r`n$path`r`n`r`nДа се запише ли въпреки това?",'CAM Agent',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
            if($ans -ne [System.Windows.Forms.DialogResult]::Yes){return}
        }
        $mappingSource='ManualOverrideOperator'
        try{
            $hr=Get-HopResultForCandidate -Product $product -Path $path
            if($null -ne $hr){
                $exp=0;$full=0
                [void][int]::TryParse((Safe-Trim $hr.ExpectedDetails),[ref]$exp)
                [void][int]::TryParse((Safe-Trim $hr.MatchedFull),[ref]$full)
                if($exp -gt 0 -and $full -eq $exp){$mappingSource='HOP_VERIFIED_USER'}
            }
        }catch{}
        Set-ProductFolderMapping -Product $product -Path $path -Source $mappingSource
        if($mappingSource -eq 'HOP_VERIFIED_USER'){$status.Text="Потвърдено: $product -> $path (100% HOP съвпадение)."}
        else{$status.Text="Ръчно зададена папка: $product -> $path. Натисни 'Провери HOP на зададените' за потвърждение."}
        try{Refresh-OperatorDashboard}catch{}
        try{Refresh-OperatorReports}catch{}
    }catch{
        [System.Windows.Forms.MessageBox]::Show("Грешка при задаване на папка:`r`n$($_.Exception.Message)",'CAM Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
    }finally{$dialog.Dispose()}
}

function Confirm-CurrentProductPath {
    param([string]$Path,[string]$Source)
    if ($null -eq $confirmProductList.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show('Няма избрано изделие за потвърждение.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $product = [string]$confirmProductList.SelectedItem
    $path = Safe-Trim $Path
    if ([string]::IsNullOrWhiteSpace($path)) {
        [System.Windows.Forms.MessageBox]::Show('Не е избрана папка.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    if (-not [System.IO.Directory]::Exists($path)) {
        [System.Windows.Forms.MessageBox]::Show("Папката не съществува или не е достъпна:`r`n$path", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    if (-not (Test-IsUnderAllowedRoot -Path $path)) {
        $isExplicitManual = (-not [string]::IsNullOrWhiteSpace($Source)) -and ($Source -match '^Manual')
        if ($isExplicitManual) {
            # V5.1.31: manual exact-folder selection may intentionally be outside
            # the automatic-search roots, but only after explicit Yes confirmation.
            $ansOutside = [System.Windows.Forms.MessageBox]::Show("Избраната папка е извън разрешените корени:`r`n$path`r`n`r`nТова е ръчен избор. Сигурни ли сте, че искате да използвате тази папка?", 'PWX500 Agent - Ръчно потвърждение', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($ansOutside -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Кандидатът е извън разрешените корени и няма да бъде приет автоматично:`r`n$path", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
    }
    $machineDir = Join-Path $path '03_Mashini'
    if (-not [System.IO.Directory]::Exists($machineDir)) {
        $ans = [System.Windows.Forms.MessageBox]::Show("В избраната папка не е намерена директна подпапка 03_Mashini:`r`n$path`r`n`r`nДа се запише ли въпреки това?", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    try {
        # If this exact candidate already has a 100% HOP match from the current HOP check,
        # a human selection is a trusted HOP-verified decision. A completely new manual
        # folder remains untrusted until HOP is run for it.
        $mappingSource=$Source
        try {
            $hr=Get-HopResultForCandidate -Product $product -Path $path
            if($null -ne $hr){
                $exp=0;$full=0
                [void][int]::TryParse((Safe-Trim $hr.ExpectedDetails),[ref]$exp)
                [void][int]::TryParse((Safe-Trim $hr.MatchedFull),[ref]$full)
                if($exp -gt 0 -and $full -eq $exp){$mappingSource='HOP_VERIFIED_USER'}
            }
        } catch {}
        Set-ProductFolderMapping -Product $product -Path $path -Source $mappingSource
        $status.Text = "Потвърдено: $product -> $path"
        Refresh-CandidateGridFromSearchResults
        $tabs.SelectedTab = $tabConfirm
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при запис на потвърждението:`r`n$($_.Exception.Message)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}


function Apply-HistoryGridVisualStyles {
    if ($null -eq $historyGrid) { return }
    foreach ($gridRow in $historyGrid.Rows) {
        try {
            $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
            $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            $gridRow.DefaultCellStyle.Font = $historyGrid.Font
            $result = ([string]$gridRow.Cells['Запис'].Value).Trim()
            $decision = ([string]$gridRow.Cells['Решение'].Value).Trim()
            if ($result -eq 'НОВА ВРЪЗКА ЗАПИСАНА' -or $result -eq 'ВЕЧЕ ЗАПИСАНА') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
            } elseif ($result -like 'КОНФЛИКТ*' -or $decision -like '*MULTIPLE*' -or $decision -like '*DUPLICATE*') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            } elseif ($decision -eq 'NO_FILE_FOUND' -or $decision -eq 'NO_TWO_DETAIL_EVIDENCE' -or $decision -eq 'NOT_ENOUGH_DISTINCT_DETAILS') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
            } elseif ($result -eq 'НЕ СЕ ЗАПИСВА') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::Gainsboro
            }
        } catch {}
    }
    try { $historyGrid.Invalidate() } catch {}
}

function Refresh-HistoryGrid {
    if ($null -eq $historyGrid) { return }
    $table = New-DataTable -Columns @('Изделие','Проби','Намерени детайли','Липсващи проби','Кандидати 2+','Изключени','Предложена папка','Доказателство','Решение','Запис')
    foreach ($x in @($script:HistoryResultObjects | Sort-Object Product)) {
        $r=$table.NewRow()
        $r['Изделие']=[string]$x.Product
        $r['Проби']=[string]$x.Codes
        $r['Намерени детайли']=[string]$x.FoundUnique
        $r['Липсващи проби']=[string]$x.NotFound
        $r['Кандидати 2+']=[string]$x.Multiple
        $r['Изключени']=[string]$x.Excluded
        $r['Предложена папка']=[string]$x.ProposedPath
        $r['Доказателство']=[string]$x.Evidence
        $r['Решение']=[string]$x.Decision
        $r['Запис']=[string]$x.MappingResult
        [void]$table.Rows.Add($r)
    }
    $historyGrid.DataSource=$table
    Apply-HistoryGridVisualStyles
}

function Stop-HistoryWorkerTimer {
    if ($null -ne $script:HistoryWorkerTimer) {
        try { $script:HistoryWorkerTimer.Stop() } catch {}
        try { $script:HistoryWorkerTimer.Dispose() } catch {}
        $script:HistoryWorkerTimer=$null
    }
}

function Start-BackgroundHistoryMigration {
    param([Parameter(Mandatory=$true)][string]$WorkbookPath)

    if (-not [System.IO.File]::Exists($script:HistoryWorkerScript)) { throw "Липсва помощният файл: $($script:HistoryWorkerScript)" }
    if (-not [System.IO.File]::Exists($WorkbookPath)) { throw 'Избраният Excel файл не съществува.' }
    if ([System.IO.Path]::GetExtension($WorkbookPath).ToLowerInvariant() -ne '.xlsx') { throw 'За началната миграция е разрешен само .xlsx файл.' }
    if ($null -ne $script:HistoryWorkerProcess -and -not $script:HistoryWorkerProcess.HasExited) { return }

    foreach($f in @($script:HistoryCodesCsv,$script:HistoryProductsCsv,$script:HistoryProgressFile,$script:HistoryDoneFile,$script:HistoryErrorFile)){
        try{if([System.IO.File]::Exists($f)){[System.IO.File]::Delete($f)}}catch{}
    }

    $argLine="-NoProfile -ExecutionPolicy Bypass -File `"$($script:HistoryWorkerScript)`" -WorkbookPath `"$WorkbookPath`" -IndexCsv `"$($script:IndexCsv)`" -SettingsCsv `"$($script:SettingsCsv)`" -SettingsVersion `"$($script:SettingsVersion)`" -CodesOut `"$($script:HistoryCodesCsv)`" -ProductsOut `"$($script:HistoryProductsCsv)`" -ProgressFile `"$($script:HistoryProgressFile)`" -DoneFile `"$($script:HistoryDoneFile)`""
    $script:HistoryWorkerProcess=Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -WindowStyle Hidden -RedirectStandardError $script:HistoryErrorFile -PassThru
    $script:HistoryLastWorkbook=$WorkbookPath
    Save-MigrationStatus -Result 'RUNNING' -WorkbookPath $WorkbookPath -Message 'Бързата миграция се изпълнява.'
    $historyMigrationButton.Enabled=$false
    $status.Text='V5.1.38: бърза миграция - използва актуалния дневен каталог, когато е надежден; при липсващ/стар каталог автоматично използва пълния безопасен scan. Правилото за два различни физически детайла е непроменено.'

    Stop-HistoryWorkerTimer
    $script:HistoryWorkerTimer=New-Object System.Windows.Forms.Timer
    $script:HistoryWorkerTimer.Interval=600
    $script:HistoryWorkerTimer.Add_Tick({
        try {
            if([System.IO.File]::Exists($script:HistoryProgressFile)){
                $p=[System.IO.File]::ReadAllText($script:HistoryProgressFile,[System.Text.Encoding]::UTF8).Trim()
                if(-not [string]::IsNullOrWhiteSpace($p)){ $status.Text=$p; if($null -ne $historyInfoLabel){$historyInfoLabel.Text=$p} }
            }
        } catch {}

        if($null -ne $script:HistoryWorkerProcess -and $script:HistoryWorkerProcess.HasExited){
            Stop-HistoryWorkerTimer
            try{$script:HistoryWorkerProcess.WaitForExit()}catch{}
            $code=-999;try{$code=$script:HistoryWorkerProcess.ExitCode}catch{}
            $script:HistoryWorkerProcess.Dispose();$script:HistoryWorkerProcess=$null
            $historyMigrationButton.Enabled=$true
            if($null -ne $script:MainMigrationButton){$script:MainMigrationButton.Enabled=$true}

            $doneText='';try{if([System.IO.File]::Exists($script:HistoryDoneFile)){$doneText=[System.IO.File]::ReadAllText($script:HistoryDoneFile,[System.Text.Encoding]::UTF8).Trim()}}catch{}
            $ok=($doneText.StartsWith('OK;') -and [System.IO.File]::Exists($script:HistoryProductsCsv) -and [System.IO.File]::Exists($script:HistoryCodesCsv))
            if($ok){
                try{
                    $rows=@(Read-SafeCsv -Path $script:HistoryProductsCsv)
                    $objs=New-Object System.Collections.Generic.List[object]
                    $newMappings=0;$already=0;$conflicts=0;$notMapped=0
                    $mappingChanged=$false
                    $now=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

                    foreach($x in $rows){
                        $product=Safe-Trim (Get-PropertyValue $x 'Product')
                        $proposed=Safe-Trim (Get-PropertyValue $x 'ProposedPath')
                        $decision=Safe-Trim (Get-PropertyValue $x 'Decision')
                        $mappingResult='НЕ СЕ ЗАПИСВА'

                        if($decision -eq 'MAP_UNIQUE_PATH' -and -not [string]::IsNullOrWhiteSpace($product) -and -not [string]::IsNullOrWhiteSpace($proposed)){
                            if(-not (Test-IsUnderAllowedRoot -Path $proposed)){
                                $mappingResult='КОНФЛИКТ: ИЗВЪН РАЗРЕШЕНИТЕ ПАПКИ';$conflicts++
                            } elseif($script:ProductFolderMap.ContainsKey($product)){
                                $oldPath=[string]$script:ProductFolderMap[$product].Path
                                if((Normalize-PathForCompare $oldPath) -eq (Normalize-PathForCompare $proposed)){
                                    $mappingResult='ВЕЧЕ ЗАПИСАНА';$already++
                                } else {
                                    $mappingResult='КОНФЛИКТ СЪС ЗАПОМНЕНА ПАПКА';$conflicts++
                                }
                            } else {
                                $script:ProductFolderMap[$product]=[pscustomobject]@{Product=$product;Path=$proposed;ConfirmedAt=$now;Source='HISTORICAL_PGMX'}
                                $mappingResult='НОВА ВРЪЗКА ЗАПИСАНА';$newMappings++;$mappingChanged=$true
                            }
                        } else {$notMapped++}

                        $objs.Add([pscustomobject]@{
                            Product=$product
                            Codes=Get-PropertyValue $x 'Codes'
                            FoundUnique=Get-PropertyValue $x 'FoundUnique'
                            NotFound=Get-PropertyValue $x 'NotFound'
                            Multiple=Get-PropertyValue $x 'Multiple'
                            Excluded=Get-PropertyValue $x 'Excluded'
                            CandidatePaths=Get-PropertyValue $x 'CandidatePaths'
                            ProposedPath=$proposed
                            Decision=$decision
                            Evidence=Get-PropertyValue $x 'Evidence'
                            MappingResult=$mappingResult
                        })
                    }
                    if($mappingChanged){Save-ProductFolderMappings}
                    $script:HistoryResultObjects=@($objs.ToArray())
                    Refresh-HistoryGrid
                    Refresh-CandidateGridFromSearchResults

                    $marker=@(
                        'PWX500 Agent V5.1.3 TEST-ALL PRODUCTS historical migration - GLOBAL_PGMX_INDEX',
                        'CompletedAt='+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
                        'Workbook='+$script:HistoryLastWorkbook,
                        'NewMappings='+$newMappings,
                        'AlreadyMapped='+$already,
                        'Conflicts='+$conflicts,
                        'NotMapped='+$notMapped,
                        $doneText
                    )
                    [System.IO.File]::WriteAllLines($script:HistoryMarkerFile,$marker,(New-Object System.Text.UTF8Encoding($true)))

                    $tabs.SelectedTab=$tabHistory
                    $histRecords='';$histExcelRows='';$histFound='';$histMissing='';$histMultiple='';$histProposals='';$histIndexedFiles='';$histIndexedFolders='';$histVisited=''
                    if($doneText -match 'Records=(\d+)'){$histRecords=$matches[1]}
                    if($doneText -match 'ExcelRows=(\d+)'){$histExcelRows=$matches[1]}
                    if($doneText -match 'UniqueFound=(\d+)'){$histFound=$matches[1]}
                    if($doneText -match 'NotFound=(\d+)'){$histMissing=$matches[1]}
                    if($doneText -match 'Multiple=(\d+)'){$histMultiple=$matches[1]}
                    if($doneText -match 'Proposals=(\d+)'){$histProposals=$matches[1]}
                    if($doneText -match 'IndexedFiles=(\d+)'){$histIndexedFiles=$matches[1]}
                    if($doneText -match 'IndexedFolders=(\d+)'){$histIndexedFolders=$matches[1]}
                    if($doneText -match 'FoldersVisited=(\d+)'){$histVisited=$matches[1]}
                    $summary="История V0.2.20: Excel редове $histExcelRows | индексирани PGMX $histIndexedFiles в $histIndexedFolders PWX папки | обходени папки $histVisited | намерени ProgramCodeF $histFound / $histRecords | потвърдени изделия $histProposals | нови връзки $newMappings | вече записани $already | конфликти $conflicts"
                    $historyInfoLabel.Text=$summary
                    $status.Text=$summary
                    Save-MigrationStatus -Result 'SUCCESS' -WorkbookPath $script:HistoryLastWorkbook -Message $summary -Success
                    try{Refresh-MainOperatorUi}catch{}
                }catch{
                    $status.Text="Историческата проверка приключи, но зареждането/записът на връзките даде грешка: $($_.Exception.Message)"
                    $historyInfoLabel.Text=$status.Text
                    Save-MigrationStatus -Result 'ERROR' -WorkbookPath $script:HistoryLastWorkbook -Message $status.Text
                    try{Refresh-MainOperatorUi}catch{}
                }
            } else {
                $progress='';$stderr=''
                try{if([System.IO.File]::Exists($script:HistoryProgressFile)){$progress=[System.IO.File]::ReadAllText($script:HistoryProgressFile,[System.Text.Encoding]::UTF8).Trim()}}catch{}
                try{if([System.IO.File]::Exists($script:HistoryErrorFile)){$stderr=[System.IO.File]::ReadAllText($script:HistoryErrorFile).Trim()}}catch{}
                $msg="Грешка при историческата PGMX проверка. ExitCode=$code. $progress $stderr"
                $status.Text=$msg;$historyInfoLabel.Text=$msg
                Save-MigrationStatus -Result 'ERROR' -WorkbookPath $script:HistoryLastWorkbook -Message $msg
                try{Refresh-MainOperatorUi}catch{}
            }
        }
    })
    $script:HistoryWorkerTimer.Start()
}


function Initialize-TrainingNative {
    if ('PWXTrainingNative' -as [type]) {
        $script:TrainingNativeReady = $true
        return $true
    }

    try {
        try { Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop } catch {}
        try { Add-Type -AssemblyName UIAutomationTypes -ErrorAction Stop } catch {}
        try { Add-Type -AssemblyName WindowsBase -ErrorAction Stop } catch {}

        Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class PWXTrainingNative
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
}
'@ -ErrorAction Stop

        $script:TrainingNativeReady = $true
        return $true
    }
    catch {
        $script:TrainingNativeReady = $false
        $script:TrainingNativeError = $_.Exception.Message
        return $false
    }
}

function Get-TrainingRawKeyState {
    param([Parameter(Mandatory=$true)][int]$VirtualKey)
    try { return [int][PWXTrainingNative]::GetAsyncKeyState($VirtualKey) }
    catch { return 0 }
}

function Test-TrainingKeyDown {
    param([Parameter(Mandatory=$true)][int]$VirtualKey)
    $raw = Get-TrainingRawKeyState $VirtualKey
    return (($raw -band 0x8000) -ne 0)
}

function Get-TrainingWindowContext {
    $result = [ordered]@{
        Handle = [IntPtr]::Zero
        ProcessId = 0
        Process = ''
        Window = ''
    }
    try {
        $h = [PWXTrainingNative]::GetForegroundWindow()
        $result.Handle = $h
        if ($h -eq [IntPtr]::Zero) { return [pscustomobject]$result }

        $sb = New-Object System.Text.StringBuilder 1024
        [void][PWXTrainingNative]::GetWindowText($h, $sb, $sb.Capacity)
        $result.Window = $sb.ToString()

        [uint32]$processIdValue = 0
        [void][PWXTrainingNative]::GetWindowThreadProcessId($h, [ref]$processIdValue)
        $result.ProcessId = [int]$processIdValue
        if ($processIdValue -gt 0) {
            try { $result.Process = (Get-Process -Id ([int]$processIdValue) -ErrorAction Stop).ProcessName }
            catch { $result.Process = '' }
        }
    } catch {}
    return [pscustomobject]$result
}

function Get-TrainingUiElementAtCursor {
    $result = [ordered]@{
        X = 0
        Y = 0
        Name = ''
        AutomationId = ''
        ControlType = ''
        ClassName = ''
        Parent = ''
        Bounds = ''
    }
    try {
        $pt = [System.Windows.Forms.Cursor]::Position
        $result.X = $pt.X
        $result.Y = $pt.Y

        $automationType = 'System.Windows.Automation.AutomationElement' -as [type]
        if ($null -eq $automationType) { return [pscustomobject]$result }

        $wp = New-Object System.Windows.Point([double]$pt.X, [double]$pt.Y)
        $el = [System.Windows.Automation.AutomationElement]::FromPoint($wp)
        if ($null -ne $el) {
            try { $result.Name = Safe-String $el.Current.Name } catch {}
            try { $result.AutomationId = Safe-String $el.Current.AutomationId } catch {}
            try {
                $ct = Safe-String $el.Current.ControlType.ProgrammaticName
                $result.ControlType = $ct.Replace('ControlType.','')
            } catch {}
            try { $result.ClassName = Safe-String $el.Current.ClassName } catch {}
            try {
                $br = $el.Current.BoundingRectangle
                $result.Bounds = ('{0},{1},{2},{3}' -f [math]::Round($br.X),[math]::Round($br.Y),[math]::Round($br.Width),[math]::Round($br.Height))
            } catch {}
            try {
                $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
                $parent = $walker.GetParent($el)
                if ($null -ne $parent) { $result.Parent = Safe-String $parent.Current.Name }
            } catch {}
        }
    } catch {}
    return [pscustomobject]$result
}

function Capture-TrainingScreenshot {
    param(
        [Parameter(Mandatory=$true)][int]$Sequence,
        [Parameter(Mandatory=$true)][string]$EventType,
        [IntPtr]$WindowHandle = [IntPtr]::Zero
    )

    if (-not $script:TrainingScreenshotEnabled) { return '' }
    if ([string]::IsNullOrWhiteSpace($script:TrainingSessionPath)) { return '' }

    try {
        $shotDir = Join-Path $script:TrainingSessionPath 'screenshots'
        if (-not [System.IO.Directory]::Exists($shotDir)) { [void][System.IO.Directory]::CreateDirectory($shotDir) }

        # V5.1.2: always capture the complete Windows virtual desktop.
        # This includes every connected monitor, regardless of which window is active.
        # The active window is still stored separately in the event metadata.
        $virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $left = $virtual.Left
        $top = $virtual.Top
        $width = [Math]::Max(1, $virtual.Width)
        $height = [Math]::Max(1, $virtual.Height)
        if ($width -lt 2 -or $height -lt 2) { return '' }

        $bmp = New-Object System.Drawing.Bitmap($width, $height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.CopyFromScreen($left, $top, 0, 0, (New-Object System.Drawing.Size($width,$height)))
        }
        finally { $g.Dispose() }

        $safeType = [regex]::Replace($EventType, '[^A-Za-z0-9_-]', '_')
        $fileName = ('{0:D5}_{1}_{2}.jpg' -f $Sequence, $safeType, (Get-Date -Format 'HHmmss_fff'))
        $full = Join-Path $shotDir $fileName

        $saveImage = $bmp
        $scaled = $null
        try {
            $maxW = 3200; $maxH = 1800
            if ($width -gt $maxW -or $height -gt $maxH) {
                $ratio = [Math]::Min($maxW / [double]$width, $maxH / [double]$height)
                $newW = [Math]::Max(1, [int][Math]::Round($width * $ratio))
                $newH = [Math]::Max(1, [int][Math]::Round($height * $ratio))
                $scaled = New-Object System.Drawing.Bitmap($bmp, (New-Object System.Drawing.Size($newW,$newH)))
                $saveImage = $scaled
            }
            $saveImage.Save($full, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        }
        finally {
            if ($null -ne $scaled) { $scaled.Dispose() }
            $bmp.Dispose()
        }
        return ('screenshots\' + $fileName)
    }
    catch { return '' }
}

function Add-TrainingEvent {
    param(
        [Parameter(Mandatory=$true)][string]$Type,
        [string]$Action = '',
        [string]$Note = '',
        $Context = $null,
        $Element = $null,
        [switch]$Screenshot
    )

    if ([string]::IsNullOrWhiteSpace($script:TrainingSessionPath)) { return }
    $script:TrainingSeq++
    $now = Get-Date
    if ($null -eq $Context) { $Context = Get-TrainingWindowContext }
    if ($null -eq $Element) {
        $Element = [pscustomobject]@{ X=0; Y=0; Name=''; AutomationId=''; ControlType=''; ClassName=''; Parent=''; Bounds='' }
    }

    $elapsed = 0.0
    if ($null -ne $script:TrainingSessionStart) { $elapsed = ($now - $script:TrainingSessionStart).TotalSeconds }
    $shot = ''
    if ($Screenshot) { $shot = Capture-TrainingScreenshot -Sequence $script:TrainingSeq -EventType $Type -WindowHandle $Context.Handle }

    $obj = [pscustomobject][ordered]@{
        Seq = $script:TrainingSeq
        Timestamp = $now.ToString('yyyy-MM-dd HH:mm:ss.fff')
        ElapsedSec = [Math]::Round($elapsed,3)
        Type = $Type
        Action = $Action
        Process = Safe-String $Context.Process
        ProcessId = Safe-String $Context.ProcessId
        Window = Safe-String $Context.Window
        X = Safe-String $Element.X
        Y = Safe-String $Element.Y
        ElementName = Safe-String $Element.Name
        AutomationId = Safe-String $Element.AutomationId
        ControlType = Safe-String $Element.ControlType
        ClassName = Safe-String $Element.ClassName
        Parent = Safe-String $Element.Parent
        Bounds = Safe-String $Element.Bounds
        Screenshot = $shot
        Note = $Note
    }
    $script:TrainingEvents.Add($obj)

    try {
        $csv = Join-Path $script:TrainingSessionPath 'events.csv'
        $vals = @(
            $obj.Seq,$obj.Timestamp,$obj.ElapsedSec,$obj.Type,$obj.Action,$obj.Process,$obj.ProcessId,$obj.Window,
            $obj.X,$obj.Y,$obj.ElementName,$obj.AutomationId,$obj.ControlType,$obj.ClassName,$obj.Parent,$obj.Bounds,$obj.Screenshot,$obj.Note
        ) | ForEach-Object { Convert-ToCsvField $_ }
        [System.IO.File]::AppendAllText($csv, (($vals -join ';') + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    } catch {}

    try {
        $jsonLine = ($obj | ConvertTo-Json -Compress -Depth 4)
        [System.IO.File]::AppendAllText((Join-Path $script:TrainingSessionPath 'events.jsonl'), ($jsonLine + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    } catch {}

    try {
        if ($null -ne $script:TrainingEventTable) {
            $r = $script:TrainingEventTable.NewRow()
            $r['#'] = [string]$obj.Seq
            $r['Час'] = $now.ToString('HH:mm:ss')
            $r['Тип'] = $obj.Type
            $r['Приложение'] = $obj.Process
            $display = $obj.Action
            if (-not [string]::IsNullOrWhiteSpace($obj.ElementName)) { $display += " | " + $obj.ElementName }
            if (-not [string]::IsNullOrWhiteSpace($obj.Note)) { $display += " | " + $obj.Note }
            $r['Действие / елемент'] = $display
            [void]$script:TrainingEventTable.Rows.Add($r)
            if ($null -ne $trainingEventsGrid -and $trainingEventsGrid.Rows.Count -gt 0) {
                $trainingEventsGrid.FirstDisplayedScrollingRowIndex = $trainingEventsGrid.Rows.Count - 1
            }
        }
    } catch {}
}

function Write-TrainingReadableSteps {
    if ([string]::IsNullOrWhiteSpace($script:TrainingSessionPath)) { return }
    try {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('PWX500 Agent V5.1.2 - Training session')
        $lines.Add('Задача: ' + $script:TrainingSessionName)
        $lines.Add('Начало: ' + $script:TrainingSessionStart.ToString('yyyy-MM-dd HH:mm:ss'))
        $lines.Add('')
        foreach ($e in @($script:TrainingEvents.ToArray())) {
            $prefix = ('{0,5}. [{1}] {2}' -f $e.Seq, $e.Timestamp.Substring(11,8), $e.Type)
            $line = $prefix
            if (-not [string]::IsNullOrWhiteSpace($e.Process)) { $line += ' | ' + $e.Process }
            if (-not [string]::IsNullOrWhiteSpace($e.Window)) { $line += ' | ' + $e.Window }
            if (-not [string]::IsNullOrWhiteSpace($e.Action)) { $line += ' | ' + $e.Action }
            if (-not [string]::IsNullOrWhiteSpace($e.ElementName)) { $line += ' | UI: ' + $e.ElementName }
            if (-not [string]::IsNullOrWhiteSpace($e.ControlType)) { $line += ' [' + $e.ControlType + ']' }
            if (-not [string]::IsNullOrWhiteSpace($e.Note)) { $line += ' | БЕЛЕЖКА: ' + ($e.Note -replace "`r?`n", ' / ') }
            $lines.Add($line)
        }
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllLines((Join-Path $script:TrainingSessionPath 'STEPS_READABLE.txt'), $lines.ToArray(), $utf8Bom)
    } catch {}
}

function Start-TrainingSession {
    if ($script:TrainingActive) { return }
    if (-not (Initialize-TrainingNative)) {
        $msg = 'Режимът за обучение не може да зареди Windows наблюдението.'
        if (-not [string]::IsNullOrWhiteSpace($script:TrainingNativeError)) { $msg += "`r`n`r`n" + $script:TrainingNativeError }
        [System.Windows.Forms.MessageBox]::Show($msg,'PWX500 Agent V5.1.2',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    $task = Safe-Trim $trainingTaskBox.Text
    if ([string]::IsNullOrWhiteSpace($task)) { $task = 'Създаване на Mix в Maestro' }

    try {
        if (-not [System.IO.Directory]::Exists($script:TrainingRoot)) { [void][System.IO.Directory]::CreateDirectory($script:TrainingRoot) }
        if (-not [System.IO.Directory]::Exists($script:TrainingExportRoot)) { [void][System.IO.Directory]::CreateDirectory($script:TrainingExportRoot) }

        $safeTask = [regex]::Replace($task, '[\\/:*?"<>|]', '_')
        $safeTask = [regex]::Replace($safeTask, '\s+', '_').Trim('_')
        if ($safeTask.Length -gt 55) { $safeTask = $safeTask.Substring(0,55) }
        $folderName = (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss') + '_' + $safeTask
        $sessionPath = Join-Path $script:TrainingRoot $folderName
        $suffix = 1
        while ([System.IO.Directory]::Exists($sessionPath)) {
            $sessionPath = Join-Path $script:TrainingRoot ($folderName + '_' + $suffix)
            $suffix++
        }
        [void][System.IO.Directory]::CreateDirectory($sessionPath)
        [void][System.IO.Directory]::CreateDirectory((Join-Path $sessionPath 'screenshots'))

        $script:TrainingSessionPath = $sessionPath
        $script:TrainingSessionName = $task
        $script:TrainingSessionStart = Get-Date
        $script:TrainingSeq = 0
        $script:TrainingEvents.Clear()
        $script:TrainingLatestZip = ''
        $script:TrainingLastWindowHandle = [IntPtr]::Zero
        $script:TrainingLastMouseLeft = Test-TrainingKeyDown 0x01
        $script:TrainingLastMouseRight = Test-TrainingKeyDown 0x02
        $script:TrainingLastMouseMiddle = Test-TrainingKeyDown 0x04
        $script:TrainingKeyStates = @{}
        $script:TrainingScreenshotEnabled = $trainingScreenshotCheck.Checked

        if ($null -ne $script:TrainingEventTable) { $script:TrainingEventTable.Rows.Clear() }
        $trainingChatBox.Clear()
        $trainingSessionPathBox.Text = $sessionPath
        $trainingZipPathBox.Text = ''

        $header = 'Seq;Timestamp;ElapsedSec;Type;Action;Process;ProcessId;Window;X;Y;ElementName;AutomationId;ControlType;ClassName;Parent;Bounds;Screenshot;Note'
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText((Join-Path $sessionPath 'events.csv'), ($header + "`r`n"), $utf8Bom)
        [System.IO.File]::WriteAllText((Join-Path $sessionPath 'events.jsonl'), '', $utf8Bom)
        $info = @(
            'PWX500 Agent V5.1.2 - Training Mode',
            ('Task=' + $task),
            ('Started=' + $script:TrainingSessionStart.ToString('yyyy-MM-dd HH:mm:ss')),
            ('Computer=' + $env:COMPUTERNAME),
            ('User=' + $env:USERNAME),
            ('Screenshots=' + $script:TrainingScreenshotEnabled),
            ('ScreenshotScope=ALL_MONITORS'),
            ('MonitorCount=' + [System.Windows.Forms.Screen]::AllScreens.Count),
            'KeyboardPrivacy=Only special keys and shortcuts are recorded. Typed letters/numbers are NOT captured.',
            'ControlMode=OBSERVE ONLY. Training Mode never clicks, types, or controls Maestro/Explorer.'
        )
        [System.IO.File]::WriteAllLines((Join-Path $sessionPath 'session_info.txt'), $info, $utf8Bom)

        $script:TrainingActive = $true
        $trainingStartButton.Enabled = $false
        $trainingStopButton.Enabled = $true
        $trainingAddNoteButton.Enabled = $true
        $trainingStatusLabel.Text = 'ОБУЧЕНИЕТО Е АКТИВНО - премини в Maestro и работи нормално.'
        $trainingStatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        $trainingChatBox.AppendText("Агент: Обучението започна. Работи нормално; при важно правило го напиши тук.`r`n")
        Add-TrainingEvent -Type 'SYSTEM' -Action 'TRAINING_START' -Note $task
        $trainingTimer.Start()
        $status.Text = 'Training Mode: активно обучение - ' + $task
    }
    catch {
        $script:TrainingActive = $false
        [System.Windows.Forms.MessageBox]::Show("Неуспешно стартиране на обучение:`r`n$($_.Exception.Message)",'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Stop-TrainingSession {
    param([switch]$Silent)
    if (-not $script:TrainingActive -and [string]::IsNullOrWhiteSpace($script:TrainingSessionPath)) { return }
    try { if ($null -ne $trainingTimer) { $trainingTimer.Stop() } } catch {}

    if ($script:TrainingActive) {
        Add-TrainingEvent -Type 'SYSTEM' -Action 'TRAINING_STOP'
    }
    $script:TrainingActive = $false

    try { Write-TrainingReadableSteps } catch {}
    try {
        $summaryLines = New-Object System.Collections.Generic.List[string]
        $summaryLines.Add('PWX500 Agent V5.1.2 - Training summary')
        $summaryLines.Add('Задача: ' + $script:TrainingSessionName)
        $summaryLines.Add('Начало: ' + $script:TrainingSessionStart.ToString('yyyy-MM-dd HH:mm:ss'))
        $summaryLines.Add('Край: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
        $summaryLines.Add('Общо събития: ' + $script:TrainingEvents.Count)
        foreach ($g in @($script:TrainingEvents.ToArray() | Group-Object Type | Sort-Object Name)) {
            $summaryLines.Add(('  {0}: {1}' -f $g.Name,$g.Count))
        }
        $summaryLines.Add('')
        $summaryLines.Add('Този пакет е суров обучителен запис. Правило НЕ се внедрява автоматично без преглед и потвърждение.')
        [System.IO.File]::WriteAllLines((Join-Path $script:TrainingSessionPath 'SESSION_SUMMARY.txt'), $summaryLines.ToArray(), (New-Object System.Text.UTF8Encoding($true)))
    } catch {}

    try {
        if (-not [System.IO.Directory]::Exists($script:TrainingExportRoot)) { [void][System.IO.Directory]::CreateDirectory($script:TrainingExportRoot) }
        $zipName = ([System.IO.Path]::GetFileName($script:TrainingSessionPath)) + '.zip'
        $zipPath = Join-Path $script:TrainingExportRoot $zipName
        if ([System.IO.File]::Exists($zipPath)) { [System.IO.File]::Delete($zipPath) }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($script:TrainingSessionPath, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
        $script:TrainingLatestZip = $zipPath
        $trainingZipPathBox.Text = $zipPath
    } catch {
        $trainingZipPathBox.Text = 'ZIP грешка: ' + $_.Exception.Message
    }

    try {
        $trainingStartButton.Enabled = $true
        $trainingStopButton.Enabled = $false
        $trainingAddNoteButton.Enabled = $false
        $trainingStatusLabel.Text = 'Обучението е спряно. Пакетът е готов за преглед/качване.'
        $trainingStatusLabel.ForeColor = [System.Drawing.Color]::DarkBlue
        $trainingChatBox.AppendText("Агент: Обучението приключи. Записът и ZIP пакетът са създадени.`r`n")
        $status.Text = 'Training Mode: приключи - ' + $script:TrainingSessionName
    } catch {}

    if (-not $Silent) {
        $msg = "Обучението е записано.`r`n`r`nПапка:`r`n$($script:TrainingSessionPath)"
        if (-not [string]::IsNullOrWhiteSpace($script:TrainingLatestZip)) { $msg += "`r`n`r`nZIP за качване:`r`n$($script:TrainingLatestZip)" }
        [System.Windows.Forms.MessageBox]::Show($msg,'PWX500 Agent - Обучение',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
}

function Training-TimerTick {
    if (-not $script:TrainingActive) { return }
    try {
        $ctx = Get-TrainingWindowContext
        $isOwn = ($ctx.ProcessId -eq $script:TrainingOwnProcessId)

        if (-not $isOwn -and $ctx.Handle -ne [IntPtr]::Zero -and $ctx.Handle -ne $script:TrainingLastWindowHandle) {
            $script:TrainingLastWindowHandle = $ctx.Handle
            Add-TrainingEvent -Type 'WINDOW' -Action 'Foreground window changed' -Context $ctx -Screenshot
        }

        # Use both current-down bit and the low 'pressed since last poll' bit.
        # This makes short mouse clicks much less likely to be missed by the WinForms timer.
        $leftRaw = Get-TrainingRawKeyState 0x01
        $rightRaw = Get-TrainingRawKeyState 0x02
        $middleRaw = Get-TrainingRawKeyState 0x04
        $left = (($leftRaw -band 0x8000) -ne 0); $leftPressed = (($leftRaw -band 0x0001) -ne 0)
        $right = (($rightRaw -band 0x8000) -ne 0); $rightPressed = (($rightRaw -band 0x0001) -ne 0)
        $middle = (($middleRaw -band 0x8000) -ne 0); $middlePressed = (($middleRaw -band 0x0001) -ne 0)

        if (-not $isOwn) {
            if (($left -and -not $script:TrainingLastMouseLeft) -or $leftPressed) {
                $el = Get-TrainingUiElementAtCursor
                Add-TrainingEvent -Type 'CLICK' -Action 'Left click' -Context $ctx -Element $el -Screenshot
            }
            if (($right -and -not $script:TrainingLastMouseRight) -or $rightPressed) {
                $el = Get-TrainingUiElementAtCursor
                Add-TrainingEvent -Type 'CLICK' -Action 'Right click' -Context $ctx -Element $el -Screenshot
            }
            if (($middle -and -not $script:TrainingLastMouseMiddle) -or $middlePressed) {
                $el = Get-TrainingUiElementAtCursor
                Add-TrainingEvent -Type 'CLICK' -Action 'Middle click' -Context $ctx -Element $el -Screenshot
            }
        }
        $script:TrainingLastMouseLeft = $left
        $script:TrainingLastMouseRight = $right
        $script:TrainingLastMouseMiddle = $middle

        $ctrl = Test-TrainingKeyDown 0x11
        $alt = Test-TrainingKeyDown 0x12
        $shift = Test-TrainingKeyDown 0x10
        $keyMap = [ordered]@{
            0x0D='Enter'; 0x1B='Esc'; 0x09='Tab'; 0x2E='Delete';
            0x71='F2'; 0x73='F4'; 0x74='F5'; 0x20='Space';
            0x25='Left'; 0x26='Up'; 0x27='Right'; 0x28='Down'
        }
        foreach ($kv in $keyMap.GetEnumerator()) {
            $vk = [int]$kv.Key
            $down = Test-TrainingKeyDown $vk
            $stateKey = 'VK_' + $vk
            $prev = $false
            if ($script:TrainingKeyStates.ContainsKey($stateKey)) { $prev = [bool]$script:TrainingKeyStates[$stateKey] }
            if (-not $isOwn -and $down -and -not $prev) {
                $name = [string]$kv.Value
                if ($ctrl) { $name = 'Ctrl+' + $name }
                elseif ($alt) { $name = 'Alt+' + $name }
                elseif ($shift) { $name = 'Shift+' + $name }
                Add-TrainingEvent -Type 'KEY' -Action $name -Context $ctx -Screenshot
            }
            $script:TrainingKeyStates[$stateKey] = $down
        }

        if ($ctrl -and -not $isOwn) {
            $shortcutMap = [ordered]@{ 0x43='Ctrl+C';0x56='Ctrl+V';0x58='Ctrl+X';0x5A='Ctrl+Z';0x41='Ctrl+A';0x53='Ctrl+S';0x4F='Ctrl+O';0x4E='Ctrl+N' }
            foreach ($kv in $shortcutMap.GetEnumerator()) {
                $vk = [int]$kv.Key
                $down = Test-TrainingKeyDown $vk
                $stateKey = 'CTRL_' + $vk
                $prev = $false
                if ($script:TrainingKeyStates.ContainsKey($stateKey)) { $prev = [bool]$script:TrainingKeyStates[$stateKey] }
                if ($down -and -not $prev) { Add-TrainingEvent -Type 'KEY' -Action ([string]$kv.Value) -Context $ctx -Screenshot }
                $script:TrainingKeyStates[$stateKey] = $down
            }
        }
        else {
            foreach ($vk in @(0x43,0x56,0x58,0x5A,0x41,0x53,0x4F,0x4E)) { $script:TrainingKeyStates['CTRL_' + $vk] = $false }
        }
    } catch {}
}


# ---------------- V5.1.31 DETAIL TIME DATABASE ----------------
function Normalize-TimeProgramCode {
    param([string]$Value)
    $v=Safe-Trim $Value
    if([string]::IsNullOrWhiteSpace($v)){return ''}
    try{$v=[System.IO.Path]::GetFileName($v)}catch{}
    if(-not $v.EndsWith('.pgmx',[System.StringComparison]::OrdinalIgnoreCase)){$v += '.pgmx'}
    return $v.ToLowerInvariant()
}
function Parse-TimeValue {
    param($Value)
    if($null -eq $Value){return $null}
    $t=([string]$Value).Trim().Replace(',','.')
    if([string]::IsNullOrWhiteSpace($t)){return $null}
    $n=0.0
    if([double]::TryParse($t,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$n)){return $n}
    return $null
}
function Save-TimeActiveIndex {
    $lines=New-Object System.Collections.Generic.List[string]
    $lines.Add('ProgramCodeF;ActiveTimeSec')
    foreach($k in @($script:ActiveTimeMap.Keys | Sort-Object)){
        $t=$script:ActiveTimeMap[$k]
        if($null -eq $t -or $t -le 0){continue}
        $lines.Add((Convert-ToCsvField ([System.IO.Path]::GetFileName($k)))+';'+(Convert-ToCsvField ([string]$t)))
    }
    [System.IO.File]::WriteAllLines($script:TimesActiveCsv,$lines.ToArray(),(New-Object System.Text.UTF8Encoding($true)))
}
function Initialize-ActiveTimeIndex {
    # Fast startup index: only ProgramCodeF -> active seconds. The full 2.5 MB history DB is loaded lazily.
    $script:ActiveTimeMap=@{}
    if ([System.IO.File]::Exists((Join-Path $PSScriptRoot 'CAM_Agent_TimesImport_Commit.json'))) {
        throw 'Има незавършен запис на времена. Продължи импорта, за да се възстанови базата преди дневния процес.'
    }
    if([System.IO.File]::Exists($script:TimesActiveCsv)){
        $first=$true
        foreach($line in [System.IO.File]::ReadLines($script:TimesActiveCsv,[System.Text.Encoding]::UTF8)){
            if($first){$first=$false;continue}
            if([string]::IsNullOrWhiteSpace($line)){continue}
            $p=$line.Split(';')
            if($p.Length -lt 2){continue}
            $k=Normalize-TimeProgramCode ($p[0].Trim('"'))
            $t=Parse-TimeValue ($p[1].Trim('"'))
            if($k -and $null -ne $t -and $t -gt 0){$script:ActiveTimeMap[$k]=$t}
        }
        return
    }
    # Compatibility fallback for an existing local installation upgraded from V5.1.27.
    # This scans only two needed columns and does not create 17k PowerShell CSV objects.
    if([System.IO.File]::Exists($script:TimesDbCsv)){
        $first=$true
        foreach($line in [System.IO.File]::ReadLines($script:TimesDbCsv,[System.Text.Encoding]::UTF8)){
            if($first){$first=$false;continue}
            if([string]::IsNullOrWhiteSpace($line)){continue}
            $p=$line.Split(';')
            if($p.Length -lt 8){continue}
            $k=Normalize-TimeProgramCode ($p[0].Trim('"'))
            $t=Parse-TimeValue ($p[7].Trim('"'))
            if($k -and $null -ne $t -and $t -gt 0){$script:ActiveTimeMap[$k]=$t}
        }
        # Read-only fallback. The import worker regenerates the persisted active index.
    }
}
function Save-TimeDb {
    $lines=New-Object System.Collections.Generic.List[string]
    $lines.Add('ProgramCodeF;PreviousValidTimeSec;PreviousValidMeasuredAt;PreviousValidSource;LastMeasuredTimeSec;LastMeasuredAt;LastSource;ActiveTimeSec;ActiveMeasuredAt;ActiveSource;Status;UpdatedAt')
    foreach($k in @($script:TimeDbMap.Keys | Sort-Object)){
        $r=$script:TimeDbMap[$k]
        $vals=@($r.ProgramCodeF,$r.PreviousValidTimeSec,$r.PreviousValidMeasuredAt,$r.PreviousValidSource,$r.LastMeasuredTimeSec,$r.LastMeasuredAt,$r.LastSource,$r.ActiveTimeSec,$r.ActiveMeasuredAt,$r.ActiveSource,$r.Status,$r.UpdatedAt)
        $lines.Add(($vals | ForEach-Object { Convert-ToCsvField ([string]$_) }) -join ';')
    }
    [System.IO.File]::WriteAllLines($script:TimesDbCsv,$lines.ToArray(),(New-Object System.Text.UTF8Encoding($true)))
    try{Save-TimeActiveIndex}catch{}
}
function Initialize-TimeDatabase {
    if($script:TimeDbLoaded){return}
    $script:TimeDbMap=@{}
    if([System.IO.File]::Exists($script:TimesDbCsv)){
        foreach($r in @(Import-Csv -LiteralPath $script:TimesDbCsv -Delimiter ';' -Encoding UTF8)){
            $k=Normalize-TimeProgramCode $r.ProgramCodeF;if(-not $k){continue};$script:TimeDbMap[$k]=$r
            $t=Parse-TimeValue $r.ActiveTimeSec;if($null -ne $t -and $t -gt 0){$script:ActiveTimeMap[$k]=$t}
        }
        $script:TimeDbLoaded=$true
        return
    }
    if([System.IO.File]::Exists($script:TimesBaseCsv)){
        foreach($r in @(Import-Csv -LiteralPath $script:TimesBaseCsv -Encoding UTF8)){
            $k=Normalize-TimeProgramCode $r.ProgramCodeF;if(-not $k){continue}
            $t=Parse-TimeValue $r.Canonical_Time_sec
            $active=if($null -ne $t -and $t -gt 0){[string]$t}else{''}
            $status=if($active){'BASE_VALID'}else{'NO_VALID_TIME'}
            $script:TimeDbMap[$k]=[pscustomobject]@{ProgramCodeF=[System.IO.Path]::GetFileName($r.ProgramCodeF);PreviousValidTimeSec='';PreviousValidMeasuredAt='';PreviousValidSource='';LastMeasuredTimeSec=$active;LastMeasuredAt='2000-01-01 00:00:00';LastSource='Initial 71-file database';ActiveTimeSec=$active;ActiveMeasuredAt='2000-01-01 00:00:00';ActiveSource='Initial 71-file database';Status=$status;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')}
            if($null -ne $t -and $t -gt 0){$script:ActiveTimeMap[$k]=$t}
        }
        $script:TimeDbLoaded=$true
        Save-TimeDb
        return
    }
    $script:TimeDbLoaded=$true
}
function Ensure-TimeDatabaseLoaded {
    if(-not $script:TimeDbLoaded){Initialize-TimeDatabase}
}
function Get-ActiveTimeForCode {
    param([string]$ProgramCodeF)
    $k=Normalize-TimeProgramCode $ProgramCodeF
    if(-not $k){return $null}
    if($script:ActiveTimeMap.ContainsKey($k)){
        $t=Parse-TimeValue $script:ActiveTimeMap[$k]
        if($null -ne $t -and $t -gt 0){return $t}
    }
    return $null
}
function Get-TimeDbRecordForCode {
    param([string]$ProgramCodeF)
    $k=Normalize-TimeProgramCode $ProgramCodeF
    if(-not $k){return $null}
    Ensure-TimeDatabaseLoaded
    if($script:TimeDbMap.ContainsKey($k)){return $script:TimeDbMap[$k]}
    return $null
}
function Test-TimeCodeHasRealMeasurement {
    param([string]$ProgramCodeF)
    $r=Get-TimeDbRecordForCode $ProgramCodeF
    if($null -eq $r){return $false}
    $t=Parse-TimeValue $r.ActiveTimeSec
    if($null -eq $t -or $t -le 0){return $false}
    $status=Safe-Trim $r.Status
    $src=Safe-Trim $r.ActiveSource
    if($status -eq 'MEDIAN_FALLBACK' -or $src -eq $script:MedianFallbackSource){return $false}
    return $true
}
function Test-TimeCodeUsesMedianFallback {
    param([string]$ProgramCodeF)
    $r=Get-TimeDbRecordForCode $ProgramCodeF
    if($null -eq $r){return $false}
    $t=Parse-TimeValue $r.ActiveTimeSec
    if($null -eq $t -or $t -le 0){return $false}
    return ((Safe-Trim $r.Status) -eq 'MEDIAN_FALLBACK' -or (Safe-Trim $r.ActiveSource) -eq $script:MedianFallbackSource)
}
function Test-TimeFileAlreadyImported {
    param([string]$Path)
    if(-not [System.IO.File]::Exists($Path)){return $false}
    $fi=Get-Item -LiteralPath $Path
    $id=($fi.FullName.ToLowerInvariant()+'|'+$fi.Length+'|'+$fi.LastWriteTimeUtc.Ticks)
    if(-not [System.IO.File]::Exists($script:TimesImportRegistryCsv)){return $false}
    foreach($r in @(Import-Csv -LiteralPath $script:TimesImportRegistryCsv -Delimiter ';' -Encoding UTF8)){if($r.Identity -eq $id){return $true}}
    return $false
}
function Add-TimeImportRegistryRow {
    param([string]$Path,[int]$Rows,[int]$Positive,[int]$Zero)
    $fi=Get-Item -LiteralPath $Path;$id=($fi.FullName.ToLowerInvariant()+'|'+$fi.Length+'|'+$fi.LastWriteTimeUtc.Ticks)
    if(-not [System.IO.File]::Exists($script:TimesImportRegistryCsv)){[System.IO.File]::WriteAllText($script:TimesImportRegistryCsv,'Identity;Path;ModifiedAt;ImportedAt;Rows;Positive;Zero'+[Environment]::NewLine,(New-Object System.Text.UTF8Encoding($true)))}
    $line=@($id,$fi.FullName,$fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'),(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Rows,$Positive,$Zero) | ForEach-Object { Convert-ToCsvField ([string]$_) }
    [System.IO.File]::AppendAllText($script:TimesImportRegistryCsv,(($line -join ';')+[Environment]::NewLine),(New-Object System.Text.UTF8Encoding($false)))
}
function Merge-TimeObservation {
    param([string]$ProgramCodeF,[double]$TimeSec,[datetime]$MeasuredAt,[string]$Source)
    Ensure-TimeDatabaseLoaded
    $k=Normalize-TimeProgramCode $ProgramCodeF;if(-not $k){return}
    $now=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    if($script:TimeDbMap.ContainsKey($k)){$r=$script:TimeDbMap[$k]}else{$r=[pscustomobject]@{ProgramCodeF=[System.IO.Path]::GetFileName($ProgramCodeF);PreviousValidTimeSec='';PreviousValidMeasuredAt='';PreviousValidSource='';LastMeasuredTimeSec='';LastMeasuredAt='';LastSource='';ActiveTimeSec='';ActiveMeasuredAt='';ActiveSource='';Status='NO_VALID_TIME';UpdatedAt=$now};$script:TimeDbMap[$k]=$r}
    $lastDate=[datetime]::MinValue;try{if($r.LastMeasuredAt){$lastDate=[datetime]::Parse($r.LastMeasuredAt)}}catch{}
    if($MeasuredAt -ge $lastDate){$r.LastMeasuredTimeSec=[string]$TimeSec;$r.LastMeasuredAt=$MeasuredAt.ToString('yyyy-MM-dd HH:mm:ss');$r.LastSource=$Source}
    $active=Parse-TimeValue $r.ActiveTimeSec;$activeDate=[datetime]::MinValue;try{if($r.ActiveMeasuredAt){$activeDate=[datetime]::Parse($r.ActiveMeasuredAt)}}catch{}
    if($TimeSec -gt 0 -and $MeasuredAt -ge $activeDate){if($null -ne $active -and $active -gt 0){$r.PreviousValidTimeSec=[string]$active;$r.PreviousValidMeasuredAt=$r.ActiveMeasuredAt;$r.PreviousValidSource=$r.ActiveSource};$r.ActiveTimeSec=[string]$TimeSec;$r.ActiveMeasuredAt=$MeasuredAt.ToString('yyyy-MM-dd HH:mm:ss');$r.ActiveSource=$Source;$r.Status='VALID'}elseif($TimeSec -le 0){if($null -ne $active -and $active -gt 0){$r.Status='ZERO_IGNORED'}else{$r.Status='NO_VALID_TIME'}}
    $activeNow=Parse-TimeValue $r.ActiveTimeSec
    if($null -ne $activeNow -and $activeNow -gt 0){$script:ActiveTimeMap[$k]=$activeNow}elseif($script:ActiveTimeMap.ContainsKey($k)){$script:ActiveTimeMap.Remove($k)}
    $r.UpdatedAt=$now
    if(-not [System.IO.File]::Exists($script:TimesMeasurementsCsv)){[System.IO.File]::WriteAllText($script:TimesMeasurementsCsv,'ProgramCodeF;MeasuredTimeSec;MeasuredAt;Source;ImportedAt'+[Environment]::NewLine,(New-Object System.Text.UTF8Encoding($true)))}
    $line=@($r.ProgramCodeF,[string]$TimeSec,$MeasuredAt.ToString('yyyy-MM-dd HH:mm:ss'),$Source,$now)|ForEach-Object{Convert-ToCsvField ([string]$_)}
    [System.IO.File]::AppendAllText($script:TimesMeasurementsCsv,(($line -join ';')+[Environment]::NewLine),(New-Object System.Text.UTF8Encoding($false)))
}
function Import-TimeWorkbookToDb {
    param([Parameter(Mandatory=$true)][string]$Path,[switch]$Force)
    if ($script:TimeImportRunning -or $script:TimeImportSynchronousRunning) {
        throw 'Вече се изпълнява импорт на времена. XLS файлът е запазен; добави го след приключването на текущия импорт.'
    }
    if (-not [System.IO.File]::Exists($Path)) { throw "Липсва файл с времена: $Path" }
    if (-not [System.IO.File]::Exists($script:TimeImportWorkerScript)) { throw 'Липсва модулът за импорт на времена V5.1.31.' }
    # Daily PowerSet and manual imports use the same lock, parser and commit/recovery code.
    # A separate job does not overwrite the last manual batch and its resume manifest.
    $base = Join-Path $env:LOCALAPPDATA 'PWX500_Agent\TimeImportJobs'
    [void][System.IO.Directory]::CreateDirectory($base)
    $jobDir = Join-Path $base ('Daily_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')+'_'+[guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($jobDir)
    $manifest = Join-Path $jobDir 'files.txt'
    $progress = Join-Path $jobDir 'progress.json'
    $result = Join-Path $jobDir 'result.json'
    $stopPath = Join-Path $jobDir 'stop.flag'
    [System.IO.File]::WriteAllLines($manifest, [string[]]@($Path), (New-Object System.Text.UTF8Encoding($true)))
    $arg = '-NoProfile -ExecutionPolicy Bypass -STA -File "'+$script:TimeImportWorkerScript+'" -ManifestPath "'+$manifest+'" -ProgressPath "'+$progress+'" -StopPath "'+$stopPath+'" -ResultPath "'+$result+'" -RootPath "'+$PSScriptRoot+'"'
    if ($Force) { $arg += ' -ForceImport' }
    $proc = $null
    $script:TimeImportSynchronousRunning = $true
    $script:TimeImportSynchronousStopPath = $stopPath
    try {
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $arg -WindowStyle Hidden -PassThru -ErrorAction Stop
        while (-not $proc.HasExited) {
            if ($script:OperatorStopRequested -and -not [System.IO.File]::Exists($stopPath)) {
                [System.IO.File]::WriteAllText($stopPath, 'STOP')
            }
            $pobj = Read-TimeImportJson -Path $progress
            if ($null -ne $pobj) { Set-PowerSetTestStatus ('Времена: '+[string]$pobj.Message) }
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 150
        }
        $robj = Read-TimeImportJson -Path $result
        if ($null -eq $robj) { throw ('Импортът е прекъснат без резултат. Файлът е запазен: '+$Path+' | Проверка: '+$jobDir) }
        if ([string]$robj.Status -ne 'Completed' -or [int]$robj.Stats.Errors -gt 0) {
            $reason = [string]$robj.Error
            if ([string]::IsNullOrWhiteSpace($reason)) { $reason = [string]$robj.Message }
            if ([string]$robj.Status -eq 'Stopped') { $reason = 'Спрян безопасно; текущият файл може да се добави отново.' }
            throw ('Времената от XLS не са импортнати. '+$reason+' | Лог: '+[string]$robj.ErrorLogPath)
        }
        $script:TimeDbLoaded=$false; $script:TimeDbMap=@{}
        Initialize-ActiveTimeIndex
        return [pscustomobject]@{Path=$Path;Rows=[int]$robj.Stats.Rows;Positive=[int]$robj.Stats.Positive;Zero=[int]$robj.Stats.Zero;MedianFallback=[int]$robj.Stats.MedianFallback;FallbackCodes=@($robj.FallbackCodes);MedianFallbackSec=$(if($null -ne $robj.MedianFallbackSec){[double]$robj.MedianFallbackSec}else{$script:MedianFallbackSec});Skipped=([int]$robj.Stats.FilesSkipped -gt 0)}
    } finally {
        $script:TimeImportSynchronousRunning = $false
        $script:TimeImportSynchronousStopPath = ''
        if ($null -ne $proc) { $proc.Dispose() }
    }
}
function Save-TimeImportLastJob {
    param([string]$Status,[string]$JobDir,[string]$ManifestPath,[string]$ProgressPath,[string]$StopPath,[string]$ResultPath,[int]$ProcessId)
    $tmp = $script:TimeImportLastJobFile+'.tmp.'+[guid]::NewGuid().ToString('N')
    try {
        $obj=[ordered]@{Status=$Status;JobDir=$JobDir;ManifestPath=$ManifestPath;ProgressPath=$ProgressPath;StopPath=$StopPath;ResultPath=$ResultPath;ProcessId=$ProcessId;UpdatedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')}
        [System.IO.File]::WriteAllText($tmp,($obj|ConvertTo-Json -Compress),(New-Object System.Text.UTF8Encoding($false)))
        for ($attempt=0; $attempt -lt 5; $attempt++) {
            try {
                if ([System.IO.File]::Exists($script:TimeImportLastJobFile)) {
                    [System.IO.File]::Replace($tmp,$script:TimeImportLastJobFile,[System.Management.Automation.Language.NullString]::Value,$true)
                } else { [System.IO.File]::Move($tmp,$script:TimeImportLastJobFile) }
                return
            } catch { if ($attempt -ge 4) { throw }; Start-Sleep -Milliseconds 70 }
        }
    } catch {
        try { [System.IO.File]::AppendAllText((Join-Path $PSScriptRoot 'CAM_Agent_TimeImport_UI.log'),((Get-Date -Format 's')+' Resume state: '+$_.Exception.Message+[Environment]::NewLine)) } catch {}
    } finally { if ([System.IO.File]::Exists($tmp)) { try { [System.IO.File]::Delete($tmp) } catch {} } }
}
function Get-TimeImportLastJob {
    try {
        if(-not [System.IO.File]::Exists($script:TimeImportLastJobFile)){return $null}
        return (Get-Content -LiteralPath $script:TimeImportLastJobFile -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch { return $null }
}
function Read-TimeImportJson {
    param([string]$Path)
    try {
        if([string]::IsNullOrWhiteSpace($Path) -or -not [System.IO.File]::Exists($Path)){return $null}
        return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch { return $null }
}
function Stop-TimeImportBackground {
    if(-not $script:TimeImportRunning){return}
    try {
        [System.IO.File]::WriteAllText($script:TimeImportStopPath,'STOP',(New-Object System.Text.UTF8Encoding($false)))
        if($null -ne $script:MainStatusText){$script:MainStatusText.Text='Импорт на времена: заявено е безопасно спиране след текущия файл.'}
        if($null -ne $script:MainStopTimesButton){$script:MainStopTimesButton.Enabled=$false}
    } catch {}
}
function Update-TimeImportBackgroundUi {
    if (-not $script:TimeImportRunning) { return }
    try {
        $pobj=Read-TimeImportJson -Path $script:TimeImportProgressPath
        if ($null -ne $pobj) {
            $msg='Импорт времена: '+[string]$pobj.Message
            if ([int]$pobj.TotalFiles -gt 0) { $msg += ' | файл '+[string]$pobj.CurrentIndex+' / '+[string]$pobj.TotalFiles }
            if ($pobj.CurrentFile) { $msg += ' | '+[string]$pobj.CurrentFile }
            $msg += ' | записани '+[int]$pobj.FilesDone+'; вече добавени '+[int]$pobj.FilesSkipped+'; грешки '+[int]$pobj.Errors
            if ($null -ne $script:MainStatusText) { $script:MainStatusText.Text=$msg }
            try { $status.Text=$msg } catch {}
        }
        $robj=Read-TimeImportJson -Path $script:TimeImportResultPath
        if ($null -ne $robj -and [string]$robj.Status -in @('Completed','Stopped','Error')) {
            # result.json is written before Excel cleanup and lock release in worker finally.
            if ($null -ne $script:TimeImportProcess -and -not $script:TimeImportProcess.HasExited) {
                $cleanup='Импорт времена: файловете са проверени. Приключвам Excel и освобождавам базата...'
                if ($null -ne $script:MainStatusText) { $script:MainStatusText.Text=$cleanup }
                if ($null -ne $script:MainStopTimesButton) { $script:MainStopTimesButton.Enabled=$false }
                try { $status.Text=$cleanup } catch {}
                return
            }
            $st=[string]$robj.Status
            $errors=[int]$robj.Stats.Errors
            $jobStatus=$st
            if ($st -eq 'Completed' -and $errors -gt 0) { $jobStatus='CompletedWithErrors' }
            $script:TimeImportRunning=$false
            if ($null -ne $script:TimeImportTimer) { $script:TimeImportTimer.Stop() }
            if ($null -ne $script:MainStopTimesButton) { $script:MainStopTimesButton.Enabled=$false }
            try { $script:TimeDbLoaded=$false; $script:TimeDbMap=@{}; Initialize-ActiveTimeIndex } catch {}
            Save-TimeImportLastJob -Status $jobStatus -JobDir $script:TimeImportJobDir -ManifestPath $script:TimeImportManifestPath -ProgressPath $script:TimeImportProgressPath -StopPath $script:TimeImportStopPath -ResultPath $script:TimeImportResultPath -ProcessId 0
            $doneMsg='Импортът на времена приключи успешно.'
            $icon=[System.Windows.Forms.MessageBoxIcon]::Information
            if ($st -eq 'Stopped') { $doneMsg='Импортът е спрян безопасно. Завършените файлове са запазени.' }
            elseif ($st -eq 'Error') { $doneMsg='Импортът е прекъснат. '+[string]$robj.Error; $icon=[System.Windows.Forms.MessageBoxIcon]::Error }
            elseif ($errors -gt 0) { $doneMsg='Всички избрани файлове са проверени. Има файлове с грешка.'; $icon=[System.Windows.Forms.MessageBoxIcon]::Warning }
            $totals='Записани файлове: '+[int]$robj.Stats.FilesDone+'; вече добавени: '+[int]$robj.Stats.FilesSkipped+'; с грешка: '+$errors
            $doneMsg += "`r`n"+$totals+"`r`nРедове: "+[int]$robj.Stats.Rows+'; положителни времена: '+[int]$robj.Stats.Positive+'; нули: '+[int]$robj.Stats.Zero+'; временни 5,44 s: '+[int]$robj.Stats.MedianFallback
            if ($errors -gt 0 -or $st -eq 'Error') {
                $doneMsg += "`r`n`r`nЗавършените файлове са запазени. Файловете с грешка могат да се добавят повторно."
                if ($robj.ErrorLogPath) { $doneMsg += "`r`nЛог: "+[string]$robj.ErrorLogPath }
            }
            try { $status.Text=$totals } catch {}
            try { Refresh-MainOperatorUi -StatusOnly } catch {}
            [System.Windows.Forms.MessageBox]::Show($doneMsg,'CAM Agent - времена',[System.Windows.Forms.MessageBoxButtons]::OK,$icon)|Out-Null
        } elseif ($null -eq $robj -and $null -ne $script:TimeImportProcess -and $script:TimeImportProcess.HasExited) {
            $script:TimeImportRunning=$false
            if ($null -ne $script:TimeImportTimer) { $script:TimeImportTimer.Stop() }
            if ($null -ne $script:MainStopTimesButton) { $script:MainStopTimesButton.Enabled=$false }
            Save-TimeImportLastJob -Status 'Interrupted' -JobDir $script:TimeImportJobDir -ManifestPath $script:TimeImportManifestPath -ProgressPath $script:TimeImportProgressPath -StopPath $script:TimeImportStopPath -ResultPath $script:TimeImportResultPath -ProcessId 0
            $msg='Фоновият импорт на времена е прекъснат. Завършените файлове са запазени; импортът може да се продължи.'
            if ($null -ne $script:MainStatusText) { $script:MainStatusText.Text=$msg }
            try { $status.Text=$msg; $script:TimeDbLoaded=$false; Initialize-ActiveTimeIndex } catch {}
        }
    } catch {
        try { [System.IO.File]::AppendAllText((Join-Path $PSScriptRoot 'CAM_Agent_TimeImport_UI.log'),((Get-Date -Format 's')+' UI: '+$_.Exception.Message+[Environment]::NewLine)) } catch {}
    }
}
function Start-TimeImportBackgroundFromManifest {
    param([Parameter(Mandatory=$true)][string]$ManifestPath,[string]$ExistingJobDir='')
    if($script:TimeImportRunning -or $script:TimeImportSynchronousRunning){[System.Windows.Forms.MessageBox]::Show('Вече се изпълнява импорт на времена.','CAM Agent - времена')|Out-Null;return}
    if(-not [System.IO.File]::Exists($script:TimeImportWorkerScript)){throw 'Липсва CAM_Agent_TimesImportWorker_V5_1_31.ps1.'}
    if(-not [System.IO.File]::Exists($ManifestPath)){throw 'Липсва списъкът с файлове за импорт.'}
    $jobDir=$ExistingJobDir
    if([string]::IsNullOrWhiteSpace($jobDir)){
        $base=Join-Path $env:LOCALAPPDATA 'PWX500_Agent\TimeImportJobs'
        if(-not(Test-Path $base)){[void](New-Item -ItemType Directory -Path $base -Force)}
        $jobDir=Join-Path $base ('Job_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
        [void](New-Item -ItemType Directory -Path $jobDir -Force)
    }
    $progress=Join-Path $jobDir 'progress.json'
    $stopPath=Join-Path $jobDir 'stop.flag'
    $result=Join-Path $jobDir 'result.json'
    Remove-Item -LiteralPath $progress,$stopPath,$result -Force -ErrorAction SilentlyContinue
    $arg='-NoProfile -ExecutionPolicy Bypass -STA -File "'+$script:TimeImportWorkerScript+'" -ManifestPath "'+$ManifestPath+'" -ProgressPath "'+$progress+'" -StopPath "'+$stopPath+'" -ResultPath "'+$result+'" -RootPath "'+$PSScriptRoot+'"'
    $proc=Start-Process -FilePath 'powershell.exe' -ArgumentList $arg -WindowStyle Hidden -PassThru
    $script:TimeImportProcess=$proc
    $script:TimeImportRunning=$true
    $script:TimeImportProgressPath=$progress
    $script:TimeImportStopPath=$stopPath
    $script:TimeImportResultPath=$result
    $script:TimeImportManifestPath=$ManifestPath
    $script:TimeImportJobDir=$jobDir
    if($null -eq $script:TimeImportTimer){
        $script:TimeImportTimer=New-Object System.Windows.Forms.Timer
        $script:TimeImportTimer.Interval=750
        $script:TimeImportTimer.Add_Tick({Update-TimeImportBackgroundUi})
    }
    if($null -ne $script:MainStopTimesButton){$script:MainStopTimesButton.Enabled=$true}
    Save-TimeImportLastJob -Status 'Running' -JobDir $jobDir -ManifestPath $ManifestPath -ProgressPath $progress -StopPath $stopPath -ResultPath $result -ProcessId $proc.Id
    $script:TimeImportTimer.Start()
    if($null -ne $script:MainStatusText){$script:MainStatusText.Text='Импорт на времена: стартира във фон. CAM Agent остава активен.'}
}
function Import-TimeFilesInteractive {
    if($script:TimeImportRunning -or $script:TimeImportSynchronousRunning){[System.Windows.Forms.MessageBox]::Show('Вече се изпълнява импорт на времена.','CAM Agent - времена')|Out-Null;return}
    $dlg=New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title='Добави PowerSet файлове с времена'
    $dlg.Filter='Excel / CSV (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv'
    $dlg.Multiselect=$true
    if($dlg.ShowDialog($script:OperatorMainForm) -ne [System.Windows.Forms.DialogResult]::OK){$dlg.Dispose();return}
    $files=@($dlg.FileNames)
    $dlg.Dispose()
    if($files.Count -eq 0){return}
    $base=Join-Path $env:LOCALAPPDATA 'PWX500_Agent\TimeImportJobs'
    if(-not(Test-Path $base)){[void](New-Item -ItemType Directory -Path $base -Force)}
    $jobDir=Join-Path $base ('Job_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
    [void](New-Item -ItemType Directory -Path $jobDir -Force)
    $manifest=Join-Path $jobDir 'files.txt'
    [System.IO.File]::WriteAllLines($manifest,$files,(New-Object System.Text.UTF8Encoding($true)))
    Start-TimeImportBackgroundFromManifest -ManifestPath $manifest -ExistingJobDir $jobDir
}
function Offer-ResumeTimeImportIfNeeded {
    try {
        $j=Get-TimeImportLastJob
        if($null -eq $j){return}
        if(-not $j.ManifestPath -or -not [System.IO.File]::Exists([string]$j.ManifestPath)){return}
        $st=[string]$j.Status
        if($st -eq 'Running' -and [int]$j.ProcessId -gt 0){
            try{$pp=Get-Process -Id ([int]$j.ProcessId) -ErrorAction Stop;if($null -ne $pp){return}}catch{$st='Interrupted'}
        }
        if($st -eq 'Stopped' -or $st -eq 'Interrupted' -or $st -eq 'Error' -or $st -eq 'CompletedWithErrors'){
            $ans=[System.Windows.Forms.MessageBox]::Show("Намерен е незавършен импорт на времена.`r`nЗавършените файлове са запазени.`r`n`r`nДа продължа ли с останалите?",'CAM Agent - времена',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
            if($ans -eq [System.Windows.Forms.DialogResult]::Yes){Start-TimeImportBackgroundFromManifest -ManifestPath ([string]$j.ManifestPath) -ExistingJobDir ([string]$j.JobDir)}
        }
    } catch {}
}
function Get-TimeMeasurementCandidates {
    if ([System.IO.File]::Exists((Join-Path $PSScriptRoot 'CAM_Agent_TimesImport_Commit.json'))) {
        throw 'Преди създаването на MIXX продължи импорта на времена, за да се възстанови незавършеният запис в базата.'
    }
    $list=New-Object System.Collections.Generic.List[object];$seen=@{};$ok=@('ГОТОВ','СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF')
    foreach($x in @($script:PgmxResultObjects)){
        $res=Safe-Trim $x.Result;if($ok -notcontains $res){continue}
        $pc=Safe-Trim $x.ProgramCodeF;if(-not $pc){continue}
        $k=Normalize-TimeProgramCode $pc;if($seen.ContainsKey($k)){continue};$seen[$k]=$true
        # A 5.44 median fallback is usable for reports/calculations, but it is
        # NOT a real measured time. Therefore the detail remains a PowerSet
        # measurement candidate until a real positive time is obtained.
        if(Test-TimeCodeHasRealMeasurement $pc){continue}
        if(Test-PgmxAutoExcluded $pc){continue}
        $target=Safe-Trim $x.TargetPgmx;if(-not $target){$target=$pc}
        $src='';try{if([System.IO.Path]::IsPathRooted($target)){$src=$target}else{$folder=Safe-Trim $x.PwxFolder;if($folder){$src=Join-Path $folder $target}}}catch{}
        if(-not $src -or -not [System.IO.File]::Exists($src)){continue}
        $isNew=($res -in @('СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF'))
        $list.Add([pscustomobject]@{ProgramCodeF=[System.IO.Path]::GetFileName($pc);Name=[System.IO.Path]::GetFileName($pc);Source=$src;Product=[string]$x.Product;Result=$res;IsNew=$isNew;UsesMedianFallback=(Test-TimeCodeUsesMedianFallback $pc)})
    }
    $script:LastTimeMeasurementCandidates=@($list.ToArray());return @($script:LastTimeMeasurementCandidates)
}
function Export-TimeReport {
    if (-not [System.IO.File]::Exists($script:TimesDbCsv)) { throw 'Все още няма записана база с времена. Добави файловете с времена и опитай отново.' }
    $source=Join-Path $PSScriptRoot 'CAM_Agent_TimeReport.cs'
    $template=Join-Path $PSScriptRoot 'CAM_Agent_TimesReport_Template.xlsx'
    if (-not ('CamAgentReports.TimeReport' -as [type])) {
        if (-not [System.IO.File]::Exists($source)) { throw 'Липсва CAM_Agent_TimeReport.cs от обновяването V5.1.31.' }
        if ($PSVersionTable.PSVersion.Major -le 5) {
            Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem,System.Xml.Linq -ErrorAction Stop
            Add-Type -Path $source -ReferencedAssemblies @('System.dll','System.Core.dll','System.Xml.dll','System.Xml.Linq.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll') -ErrorAction Stop
        } else { Add-Type -Path $source -ErrorAction Stop }
    }
    $dir=Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) 'CAM_Agent_Exports'
    [void][System.IO.Directory]::CreateDirectory($dir)
    $out=Join-Path $dir ('CAM_Agent_Detail_Times_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')+'.xlsx')
    # Read the on-disk committed snapshot; exporting never calls Save-TimeDb.
    $invalid=[CamAgentReports.TimeReport]::Export($script:TimesDbCsv,$template,$out)
    Start-Process explorer.exe -ArgumentList ('/select,"'+$out+'"')
    if ($invalid -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(('Отчетът е създаден. '+$invalid+' непознати стойности за време са запазени като текст за преглед; не са заменени с нула.'),'CAM Agent - отчет времена',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null
    }
    return $out
}
function Export-AllFilesRegistry {
    $dir=Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) 'CAM_Agent_Exports';if(-not(Test-Path $dir)){[void](New-Item -ItemType Directory -Path $dir)};$out=Join-Path $dir ('CAM_Agent_All_Files_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.csv');$map=@{}
    if(Test-Path $script:HistoryCodesCsv){foreach($r in @(Import-Csv -LiteralPath $script:HistoryCodesCsv -Delimiter ';' -Encoding UTF8)){$k=Normalize-TimeProgramCode $r.ProgramCodeF;if(-not $k){continue};$at=Get-ActiveTimeForCode $r.ProgramCodeF;$map[$k]=[pscustomobject]@{ProgramCodeF=[System.IO.Path]::GetFileName($r.ProgramCodeF);Product=$r.Product;DetailNo=$r.DetailNo;DetailName=$r.DetailName;PwxFolder=$r.PwxFolder;Origin='Migration';ActiveTimeSec=$(if($null -ne $at){[string]$at}else{''})}}}
    foreach($r in @($script:PgmxResultObjects)){$k=Normalize-TimeProgramCode $r.ProgramCodeF;if(-not $k){continue};if(-not $map.ContainsKey($k)){$map[$k]=[pscustomobject]@{ProgramCodeF=[System.IO.Path]::GetFileName($r.ProgramCodeF);Product=[string]$r.Product;DetailNo='';DetailName='';PwxFolder=[string]$r.PwxFolder;Origin='CAM Agent';ActiveTimeSec=''}};$t=Get-ActiveTimeForCode $r.ProgramCodeF;if($null -ne $t){$map[$k].ActiveTimeSec=[string]$t}}
    $lines=New-Object System.Collections.Generic.List[string];$lines.Add('ProgramCodeF;Product;DetailNo;DetailName;PwxFolder;Origin;ActiveTimeSec');foreach($r in @($map.Values|Sort-Object ProgramCodeF)){$lines.Add((@($r.ProgramCodeF,$r.Product,$r.DetailNo,$r.DetailName,$r.PwxFolder,$r.Origin,$r.ActiveTimeSec)|ForEach-Object{Convert-ToCsvField([string]$_)})-join ';')};[System.IO.File]::WriteAllLines($out,$lines.ToArray(),(New-Object System.Text.UTF8Encoding($true)));Start-Process explorer.exe -ArgumentList ('/select,"'+$out+'"');return $out
}
function Show-FinalProcessReport {
    $c=@($script:LastTimeMeasurementCandidates);$new=@($c|Where-Object{$_.IsNew});$old=@($c|Where-Object{-not $_.IsNew});$newOk=@($new|Where-Object{Test-TimeCodeHasRealMeasurement $_.ProgramCodeF});$oldOk=@($old|Where-Object{Test-TimeCodeHasRealMeasurement $_.ProgramCodeF});$unresolved=@($c|Where-Object{-not (Test-TimeCodeHasRealMeasurement $_.ProgramCodeF)});$fallback=@($c|Where-Object{Test-TimeCodeUsesMedianFallback $_.ProgramCodeF});$states=@(Get-OperatorProductStateRows);$unconf=@($states|Where-Object{-not $_.Confirmed});$errs=@($script:PgmxResultObjects|Where-Object{@('ГОТОВ','СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF') -notcontains (Safe-Trim $_.Result)})
    $f=New-Object System.Windows.Forms.Form;$f.Text='CAM Agent - Финален отчет';$f.StartPosition='CenterScreen';$f.Size=New-Object System.Drawing.Size(1000,720);$f.Font=New-Object System.Drawing.Font('Segoe UI',10);$f.BackColor=[System.Drawing.Color]::White
    $tb=New-Object System.Windows.Forms.TextBox;$tb.Multiline=$true;$tb.ReadOnly=$true;$tb.Dock='Top';$tb.Height=210;$tb.ScrollBars='Vertical';$tb.Text="ФИНАЛЕН ОТЧЕТ`r`n`r`nДетайли с вече реално валидно време преди текущия Mixx: $([Math]::Max(0,@($script:PgmxResultObjects).Count-$c.Count))`r`nНови детайли без реално време, включени за измерване: $($new.Count)`r`nНови детайли, получили реално време: $($newOk.Count)`r`nСтари детайли без реално време, включени за измерване: $($old.Count)`r`nСтари детайли, получили реално време: $($oldOk.Count)`r`nОстанали без реално време: $($unresolved.Count)`r`nОт тях с временен fallback 5,44 s: $($fallback.Count)`r`nНепотвърдени изделия: $($unconf.Count)`r`nДетайли с PGMX грешка: $($errs.Count)";$f.Controls.Add($tb)
    $tabs=New-Object System.Windows.Forms.TabControl;$tabs.Dock='Fill';$f.Controls.Add($tabs)
    foreach($spec in @(@('Непотвърдени',$unconf),@('Грешки',$errs),@('Без реално време',$unresolved))){$tp=New-Object System.Windows.Forms.TabPage;$tp.Text=$spec[0];$g=New-Object System.Windows.Forms.DataGridView;$g.Dock='Fill';$g.ReadOnly=$true;$g.AutoGenerateColumns=$true;$al=New-Object System.Collections.ArrayList;foreach($it in @($spec[1])){[void]$al.Add($it)};$g.DataSource=$al;$tp.Controls.Add($g);$tabs.TabPages.Add($tp)}
    $f.ShowDialog($script:OperatorMainForm)|Out-Null
}

# ---------------- STEP 4 AUTO MIXX ----------------
$script:MixAutomationReady = $false
$script:MixAutomationError = ''
$script:LastMixSaveAttemptPath = ''
$script:LastMixSaveConfirmed = $false
$script:MixLogFile = Join-Path $PSScriptRoot 'PWX500_AutoMix.log'

function Write-MixLog {
    param([string]$Message)
    try {
        $line = ('[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $Message + [Environment]::NewLine)
        [System.IO.File]::AppendAllText($script:MixLogFile,$line,(New-Object System.Text.UTF8Encoding($false)))
    } catch {}
}

function Get-MixWarningPgmxPaths {
    param([int]$ProcessId,[string]$DailyPath)
    $out=New-Object System.Collections.Generic.List[string]
    $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    try {
        if(-not (Initialize-MixAutomation)){return @()}
        $root=[System.Windows.Automation.AutomationElement]::RootElement
        $cond=New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
        $els=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$cond)
        foreach($el in @($els)){
            $texts=New-Object System.Collections.Generic.List[string]
            try{if($el.Current.Name){$texts.Add([string]$el.Current.Name)}}catch{}
            try{
                $vp=$null
                if($el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$vp)){
                    $v=[string]$vp.Current.Value;if($v){$texts.Add($v)}
                }
            }catch{}
            foreach($txt in @($texts)){
                foreach($m in [regex]::Matches($txt,'(?im)([A-Za-z]:\\[^\r\n\"]+?\.pgmx)\b')){
                    $candidate=$m.Groups[1].Value.Trim()
                    try{
                        $full=[System.IO.Path]::GetFullPath($candidate)
                        $parent=[System.IO.Path]::GetDirectoryName($full)
                        if((Normalize-PathForCompare $parent) -ne (Normalize-PathForCompare $DailyPath)){continue}
                        if([System.IO.Path]::GetExtension($full) -ine '.pgmx'){continue}
                        if(-not [System.IO.File]::Exists($full)){continue}
                        if($seen.Add($full)){$out.Add($full)}
                    }catch{}
                }
            }
        }
    } catch { Write-MixLog ('Warning PGMX read failed: '+$_.Exception.Message) }
    return @($out.ToArray())
}
function Move-DailyPgmxToSubfolder {
    param([string[]]$Paths,[Parameter(Mandatory=$true)][string]$Subfolder,[string]$Reason='')
    $daily=Get-DailyMixFolder
    $destDir=Join-Path $daily $Subfolder
    $moved=New-Object System.Collections.Generic.List[string]
    foreach($path in @($Paths)){
        if([string]::IsNullOrWhiteSpace($path)){continue}
        try{
            $full=[System.IO.Path]::GetFullPath($path)
            if([System.IO.Path]::GetExtension($full) -ine '.pgmx'){continue}
            if((Normalize-PathForCompare ([System.IO.Path]::GetDirectoryName($full))) -ne (Normalize-PathForCompare $daily)){continue}
            if(-not [System.IO.File]::Exists($full)){continue}
            if(-not [System.IO.Directory]::Exists($destDir)){[void][System.IO.Directory]::CreateDirectory($destDir)}
            $name=[System.IO.Path]::GetFileName($full)
            $dest=Join-Path $destDir $name
            if([System.IO.File]::Exists($dest)){
                # Never overwrite a previously isolated file. Keep both for diagnosis.
                $stem=[System.IO.Path]::GetFileNameWithoutExtension($name)
                $ext=[System.IO.Path]::GetExtension($name)
                $n=2
                do{$dest=Join-Path $destDir ($stem+'_'+$n+$ext);$n++}while([System.IO.File]::Exists($dest))
            }
            $done=$false;$lastError=''
            for($attempt=0;$attempt -lt 8;$attempt++){
                try{[System.IO.File]::Move($full,$dest);$done=$true;break}catch{$lastError=$_.Exception.Message;Start-Sleep -Milliseconds (100+($attempt*75))}
            }
            if($done){$moved.Add($dest);Write-MixLog ('MOVED PGMX ['+$Reason+']: '+$full+' -> '+$dest)}
            else{Write-MixLog ('MOVE PGMX skipped after retries ['+$Reason+']: '+$full+' | '+$lastError)}
        }catch{Write-MixLog ('MOVE PGMX error ['+$Reason+']: '+$path+' | '+$_.Exception.Message)}
    }
    return @($moved.ToArray())
}
function Move-DailyPgmxCodesToSubfolder {
    param([string[]]$Codes,[Parameter(Mandatory=$true)][string]$Subfolder,[string]$Reason='')
    $daily=Get-DailyMixFolder
    $paths=New-Object System.Collections.Generic.List[string]
    $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach($code in @($Codes)){
        $k=Normalize-TimeProgramCode $code
        if(-not $k){continue}
        $name=[System.IO.Path]::GetFileName($k)
        $path=Join-Path $daily $name
        if([System.IO.File]::Exists($path) -and $seen.Add($path)){$paths.Add($path)}
    }
    return @(Move-DailyPgmxToSubfolder -Paths $paths.ToArray() -Subfolder $Subfolder -Reason $Reason)
}

function Get-DefaultDailyRootFolder {
    $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    return (Join-Path $desktop 'FOUND')
}

function Normalize-DailyRootFolder {
    param([string]$Path)
    if([string]::IsNullOrWhiteSpace($Path)){throw 'Не е зададена основна папка за дневните файлове.'}
    $p=[System.Environment]::ExpandEnvironmentVariables($Path.Trim().Trim('"'))
    if(-not [System.IO.Path]::IsPathRooted($p)){throw 'Основната папка трябва да е абсолютен Windows път.'}
    # Maestro's custom browser is reliable with local/mapped-drive paths. For a
    # network location use its mapped drive letter (for example Y:\...), not UNC.
    if($p.StartsWith('\\')){throw 'За мрежова папка използвай mapped drive (напр. Y:\...), а не UNC път \\server\share.'}
    $full=[System.IO.Path]::GetFullPath($p)
    if([string]::IsNullOrWhiteSpace($full)){throw 'Невалидна основна папка.'}
    $root=[System.IO.Path]::GetPathRoot($full)
    if($root -and $full -ieq $root){return $root}
    return $full.TrimEnd('\')
}

function Get-DailyRootFolder {
    $default=Get-DefaultDailyRootFolder
    try{
        if([System.IO.File]::Exists($script:DailyRootSettingsFile)){
            $saved=([System.IO.File]::ReadAllText($script:DailyRootSettingsFile,[System.Text.Encoding]::UTF8)).Trim()
            if(-not [string]::IsNullOrWhiteSpace($saved)){return (Normalize-DailyRootFolder $saved)}
        }
    }catch{
        try{Write-MixLog ('Daily root setting invalid; using default: '+$_.Exception.Message)}catch{}
    }
    return $default
}

function Set-DailyRootFolder {
    param([Parameter(Mandatory=$true)][string]$Path)
    $full=Normalize-DailyRootFolder $Path
    if(-not [System.IO.Directory]::Exists($full)){[void][System.IO.Directory]::CreateDirectory($full)}
    $tmp=$script:DailyRootSettingsFile+'.tmp'
    [System.IO.File]::WriteAllText($tmp,$full,(New-Object System.Text.UTF8Encoding($true)))
    if([System.IO.File]::Exists($script:DailyRootSettingsFile)){
        try{[System.IO.File]::Replace($tmp,$script:DailyRootSettingsFile,$null)}catch{[System.IO.File]::Delete($script:DailyRootSettingsFile);[System.IO.File]::Move($tmp,$script:DailyRootSettingsFile)}
    }else{[System.IO.File]::Move($tmp,$script:DailyRootSettingsFile)}
    return $full
}

function Test-DailyRootCanChange {
    $rs=Get-OperatorRunState
    if($null -eq $rs){return $true}
    $st=[string]$rs.Status
    if($st -eq 'COMPLETED'){return $true}
    # Do not let a folder change invalidate a safe-stop/resume checkpoint.
    if($st -in @('ACTIVE','STOPPED','ERROR')){return $false}
    return $true
}

function Refresh-DailyRootUi {
    try{if($null -ne $autoMixFolderBox){$autoMixFolderBox.Text=Get-DailyMixFolder}}catch{}
    try{if($null -ne $script:MainDailyRootToolTip -and $null -ne $script:MainDailyRootButton){$script:MainDailyRootToolTip.SetToolTip($script:MainDailyRootButton,(Get-DailyRootFolder))}}catch{}
}

function Select-DailyRootFolderInteractive {
    param([System.Windows.Forms.IWin32Window]$Owner=$null)
    if(Test-PgmxPendingWorkInCurrentRoot){
        [System.Windows.Forms.MessageBox]::Show('Има дневни PGMX, чакащи общ MIXX/PowerSet. Първо завърши общия пакет преди смяна на основната папка.','CAM Agent - Основна папка',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null
        return $false
    }
    if(-not (Test-DailyRootCanChange)){
        [System.Windows.Forms.MessageBox]::Show('Има незавършен/прекъснат дневен процес. За безопасност основната папка не може да се сменя преди процесът да бъде приключен или да бъде стартиран наново.','CAM Agent - Основна папка',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null
        return $false
    }
    $dlg=New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description='Избери ОСНОВНАТА папка. CAM Agent ще създава вътре дневна папка dd.MM (например 20.10).'
    $dlg.ShowNewFolderButton=$true
    try{$current=Get-DailyRootFolder;if([System.IO.Directory]::Exists($current)){$dlg.SelectedPath=$current}}catch{}
    try{
        $result=if($null -ne $Owner){$dlg.ShowDialog($Owner)}else{$dlg.ShowDialog()}
        if($result -ne [System.Windows.Forms.DialogResult]::OK){return $false}
        $selected=Set-DailyRootFolder -Path ([string]$dlg.SelectedPath)
        Sync-PgmxAutoExclusionsFromDailyFolders
        Refresh-DailyRootUi
        $daily=Join-Path $selected (Get-Date -Format 'dd.MM')
        [System.Windows.Forms.MessageBox]::Show("Основната папка е записана:`r`n$selected`r`n`r`nДнешната папка ще бъде:`r`n$daily",'CAM Agent - Основна папка',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
        return $true
    }catch{
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - Основна папка',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
        return $false
    }finally{$dlg.Dispose()}
}

function Get-DailyMixFolder {
    if(-not [string]::IsNullOrWhiteSpace($script:BatchOutputFolderOverride)){return [string]$script:BatchOutputFolderOverride}
    return (Join-Path (Get-DailyRootFolder) (Get-Date -Format 'dd.MM'))
}

function Get-DailyMixFilePath {
    $folder = Get-DailyMixFolder
    return (Join-Path $folder ('Mix(' + [System.IO.Path]::GetFileName($folder) + ').mixx'))
}


function Test-DailyMixReadyForDisplay {
    # An existing empty/rejected target is not a successful Step 4.
    $target=Get-DailyMixFilePath
    $stamp=Get-MixSaveFileStamp -Path $target
    if ($null -eq $stamp -or $stamp.Length -le 0) { return $false }
    $key=Normalize-PathForCompare $target
    if ((Normalize-PathForCompare $script:LastMixSaveAttemptPath) -eq $key) { return [bool]$script:LastMixSaveConfirmed }
    $rs=Get-OperatorRunState
    if ($null -ne $rs -and (Normalize-PathForCompare ([string]$rs.DailyFolder)) -eq (Normalize-PathForCompare (Get-DailyMixFolder))) {
        if ([string]$rs.CurrentStage -eq 'MIXX_RUNNING' -and [string]$rs.LastSafeStage -eq 'PGMX_DONE') { return $false }
    }
    return $true
}

function Capture-CreatedPgmxThisRun {
    $byName = @{}
    $allowedResults = @('СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF')

    foreach ($x in @($script:PgmxResultObjects)) {
        $result = Safe-Trim $x.Result
        if ($allowedResults -notcontains $result) { continue }

        # IMPORTANT: PgmxWorker writes TargetPgmx as the target FILE NAME, not a full path.
        # Build the physical source from PwxFolder + TargetPgmx (or ProgramCodeF fallback).
        $targetValue = Safe-Trim $x.TargetPgmx
        if ([string]::IsNullOrWhiteSpace($targetValue)) { $targetValue = Safe-Trim $x.ProgramCodeF }
        if ([string]::IsNullOrWhiteSpace($targetValue)) { continue }

        $src = ''
        try {
            if ([System.IO.Path]::IsPathRooted($targetValue)) {
                $src = $targetValue
            } else {
                $folder = Safe-Trim $x.PwxFolder
                if (-not [string]::IsNullOrWhiteSpace($folder)) { $src = Join-Path $folder $targetValue }
            }
        } catch { $src = '' }
        if ([string]::IsNullOrWhiteSpace($src) -or -not [System.IO.File]::Exists($src)) { continue }

        $name = Safe-Trim $x.ProgramCodeF
        if ([string]::IsNullOrWhiteSpace($name)) { $name = [System.IO.Path]::GetFileName($src) }
        if (-not $name.EndsWith('.pgmx',[System.StringComparison]::OrdinalIgnoreCase)) { $name += '.pgmx' }

        # One ProgramCode enters the daily Mix only once even if duplicate CSV rows point to it.
        $byName[$name.ToLowerInvariant()] = [pscustomobject]@{
            Name    = $name
            Source  = $src
            Product = [string]$x.Product
            Result  = $result
        }
    }

    $script:CreatedPgmxThisRun = @($byName.Values | Sort-Object Name)
    try { Write-MixLog ('STEP3 CAPTURE: current Execute created ' + $script:CreatedPgmxThisRun.Count + ' unique PGMX.') } catch {}
    return @($script:CreatedPgmxThisRun)
}

function Get-NewPgmxForDailyMix {
    if ($script:PgmxLastMode -ne 'Execute') { return @() }
    if ($null -eq $script:CreatedPgmxThisRun -or @($script:CreatedPgmxThisRun).Count -eq 0) {
        try { Capture-CreatedPgmxThisRun | Out-Null } catch {}
    }
    return @($script:CreatedPgmxThisRun)
}

function Update-AutoMixUi {
    try {
        if ($null -ne $autoMixFolderBox) { $autoMixFolderBox.Text = Get-DailyMixFolder }
        $count = @(Get-NewPgmxForDailyMix).Count
        if ($null -ne $autoMixCountLabel) {
            if ($script:PgmxLastMode -eq 'Execute') { $autoMixCountLabel.Text = 'Нови PGMX от текущото дневно изпълнение: ' + $count }
            else { $autoMixCountLabel.Text = 'Нови PGMX от текущото дневно изпълнение: няма изпълнена Стъпка 3 в тази сесия' }
        }
    } catch {}
}

function Prepare-DailyMixFiles {
    param([switch]$UseExistingTestFiles)
    $daily = Get-DailyMixFolder
    if (-not [System.IO.Directory]::Exists($daily)) { [void][System.IO.Directory]::CreateDirectory($daily) }

    if ($UseExistingTestFiles) {
        $testFiles = @(Get-ChildItem -LiteralPath $daily -File -Filter '*.pgmx' -ErrorAction SilentlyContinue | Sort-Object Name)
        if ($testFiles.Count -eq 0) { throw "ТЕСТ режим: в дневната папка няма .pgmx файлове: $daily" }
        Write-MixLog ('TEST input: ' + $testFiles.Count + ' PGMX from ' + $daily)
        return @($testFiles | ForEach-Object { $_.FullName })
    }

    $newFiles = @(Get-TimeMeasurementCandidates)
    if ($newFiles.Count -eq 0) { throw 'Всички успешно намерени PGMX от текущите CSV вече имат валидно време. Няма файлове за нов PowerSet.' }

    # V5.1.31: copy all CURRENT CSV PGMX without a valid active time - both newly created and old existing programs.
    # manually added earlier in the same day. This is VALID and must never stop
    # the daily process. Copy/refresh the new PGMX from this Execute into the
    # existing folder, then build the common MIXX from ALL PGMX currently in it.
    $copiedNew = New-Object System.Collections.Generic.List[string]
    foreach ($x in $newFiles) {
        $dest = Join-Path $daily $x.Name
        [System.IO.File]::Copy($x.Source,$dest,$true)
        if (-not [System.IO.File]::Exists($dest)) { throw "Неуспешно копиране: $($x.Source) -> $dest" }
        $copiedNew.Add($dest)
        Write-MixLog ('COPIED/REFRESHED NEW ' + $x.Source + ' -> ' + $dest)
    }

    $allDaily = @(Get-ChildItem -LiteralPath $daily -File -Filter '*.pgmx' -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($allDaily.Count -eq 0) { throw "След копирането в дневната папка няма .pgmx файлове: $daily" }

    Write-MixLog ('DAILY MERGE: new=' + $copiedNew.Count + '; total PGMX in folder=' + $allDaily.Count + '; folder=' + $daily)
    return @($allDaily | ForEach-Object { $_.FullName })
}

function Initialize-MixAutomation {
    if ($script:MixAutomationReady) { return $true }
    try {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
        Add-Type -AssemblyName UIAutomationTypes -ErrorAction Stop
        if (-not ('PWXMixNative' -as [type])) {
            $src = @'
using System;
using System.Runtime.InteropServices;
public static class PWXMixNative
{
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
'@
            Add-Type -TypeDefinition $src -Language CSharp -ErrorAction Stop
        }
        $script:MixAutomationReady = $true
        return $true
    } catch {
        $script:MixAutomationError = $_.Exception.Message
        return $false
    }
}

function Invoke-MixNativeClick {
    param([int]$X,[int]$Y)
    [void][PWXMixNative]::SetCursorPos($X,$Y)
    Start-Sleep -Milliseconds 120
    [PWXMixNative]::mouse_event(0x0002,0,0,0,[UIntPtr]::Zero)
    [PWXMixNative]::mouse_event(0x0004,0,0,0,[UIntPtr]::Zero)
}

function Invoke-MixNativeCtrlA {
    # Native keyboard events are more reliable than SendKeys with Maestro/WPF.
    # VK_CONTROL=0x11, VK_A=0x41, KEYEVENTF_KEYUP=0x0002.
    [PWXMixNative]::keybd_event(0x11,0,0,[UIntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [PWXMixNative]::keybd_event(0x41,0,0,[UIntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [PWXMixNative]::keybd_event(0x41,0,0x0002,[UIntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [PWXMixNative]::keybd_event(0x11,0,0x0002,[UIntPtr]::Zero)
}

function Invoke-MixNativeShiftClick {
    param([int]$X,[int]$Y)
    # VK_SHIFT=0x10. Hold Shift while clicking the last PGMX row to select a
    # contiguous range from the already selected first row.
    [PWXMixNative]::keybd_event(0x10,0,0,[UIntPtr]::Zero)
    Start-Sleep -Milliseconds 70
    Invoke-MixNativeClick -X $X -Y $Y
    Start-Sleep -Milliseconds 70
    [PWXMixNative]::keybd_event(0x10,0,0x0002,[UIntPtr]::Zero)
}

function Invoke-MixPgmxRowTextClick {
    param($Element,[string]$FileName)
    # The three training sessions show a click on the ListItem followed by a
    # click on its visible filename Text. The second click gives the WPF list
    # the keyboard focus needed for Ctrl+A.
    try {
        $name=[System.IO.Path]::GetFileName($FileName)
        $all=$Element.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
        foreach($txt in $all){
            try {
                if($txt.Current.ControlType -ne [System.Windows.Automation.ControlType]::Text){ continue }
                if(([string]$txt.Current.Name) -ne $name){ continue }
                $b=$txt.Current.BoundingRectangle
                if($b.Width -gt 1 -and $b.Height -gt 1){
                    Invoke-MixNativeClick -X ([int]($b.Left + [Math]::Min(45,[Math]::Max(5,$b.Width/2)))) -Y ([int]($b.Top+$b.Height/2))
                    return $true
                }
            } catch {}
        }
    } catch {}
    return $false
}

function Invoke-MixRelativeClick {
    param([int]$ProcessId,[double]$RefX,[double]$RefY,[int]$TimeoutSec=10)

    # Maestro may recreate its main window while the application is finishing startup.
    # Never keep the first MainWindowHandle captured after launch; resolve the current
    # visible window immediately before each coordinate-based click.
    $until = (Get-Date).AddSeconds($TimeoutSec)
    do {
        try {
            $proc = Get-Process -Id $ProcessId -ErrorAction Stop
            $proc.Refresh()
            $handle = [IntPtr]$proc.MainWindowHandle
            if ($handle -ne [IntPtr]::Zero) {
                $r = New-Object PWXMixNative+RECT
                if ([PWXMixNative]::GetWindowRect($handle,[ref]$r)) {
                    $w = $r.Right - $r.Left
                    $h = $r.Bottom - $r.Top
                    if ($w -gt 300 -and $h -gt 200) {
                        $x = $r.Left + [int][Math]::Round($RefX * ($w / 1920.0))
                        $y = $r.Top + [int][Math]::Round($RefY * ($h / 1080.0))
                        Invoke-MixNativeClick -X $x -Y $y
                        return
                    }
                }

                # Fallback: UI Automation can still expose a valid bounding rectangle
                # even when GetWindowRect briefly fails during the Maestro window swap.
                try {
                    $ae = [System.Windows.Automation.AutomationElement]::FromHandle($handle)
                    if ($null -ne $ae) {
                        $b = $ae.Current.BoundingRectangle
                        if ($b.Width -gt 300 -and $b.Height -gt 200) {
                            $x = [int][Math]::Round($b.Left + ($RefX * ($b.Width / 1920.0)))
                            $y = [int][Math]::Round($b.Top + ($RefY * ($b.Height / 1080.0)))
                            Invoke-MixNativeClick -X $x -Y $y
                            return
                        }
                    }
                } catch {}
            }
        } catch {}
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $until)

    throw 'Не мога да определя актуалния прозорец MAESTRO LAB. Maestro е стартиран, но прозорецът му е бил сменен/не е готов за автоматизация.'
}

function Get-MaestroProcess {
    $p = @(Get-Process -Name 'AppMain' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like '*MAESTRO*' } | Sort-Object StartTime -Descending | Select-Object -First 1)
    if ($p.Count -gt 0) { return $p[0] }
    return $null
}

function Get-MaestroSavedLauncher {
    $cfg = Join-Path $PSScriptRoot 'MAESTROLAB_PATH.txt'
    try {
        if (Test-Path -LiteralPath $cfg -PathType Leaf) {
            $v = ([string](Get-Content -LiteralPath $cfg -Raw -ErrorAction Stop)).Trim()
            if ($v -and (Test-Path -LiteralPath $v -PathType Leaf)) { return $v }
        }
    } catch {}
    return $null
}

function Save-MaestroLauncher {
    param([string]$Path)
    if (-not $Path) { return }
    try {
        Set-Content -LiteralPath (Join-Path $PSScriptRoot 'MAESTROLAB_PATH.txt') -Value $Path -Encoding UTF8 -Force
    } catch {}
}

function Find-MaestroLauncher {
    $saved = Get-MaestroSavedLauncher
    if ($saved) { return $saved }

    $candidates = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    function Add-MaestroCandidate([string]$Path) {
        if (-not $Path) { return }
        try { $full = [System.IO.Path]::GetFullPath($Path) } catch { $full = $Path }
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { return }
        $ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()
        if ($ext -notin @('.lnk','.exe','.url','.appref-ms')) { return }
        if (-not $seen.ContainsKey($full)) { $seen[$full] = $true; $candidates.Add($full) }
    }

    # Desktop shortcuts may be on the user's Desktop or the Public Desktop.
    foreach ($folder in @(
        [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop),
        [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonDesktopDirectory)
    )) {
        if ($folder -and (Test-Path -LiteralPath $folder -PathType Container)) {
            foreach ($f in @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -like '*Maestro*' -or $_.Name -ieq 'AppMain.exe' })) {
                Add-MaestroCandidate $f.FullName
            }
        }
    }

    # Also search the user and common Start Menu trees.
    foreach ($folder in @(
        [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::StartMenu),
        [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonStartMenu)
    )) {
        if ($folder -and (Test-Path -LiteralPath $folder -PathType Container)) {
            foreach ($f in @(Get-ChildItem -LiteralPath $folder -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -like '*Maestro*' -or $_.Name -ieq 'AppMain.exe' })) {
                Add-MaestroCandidate $f.FullName
            }
        }
    }

    # Try SCM/Maestro install information from Windows uninstall registry entries.
    foreach ($regPath in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )) {
        foreach ($app in @(Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue | Where-Object { ([string]$_.DisplayName) -like '*Maestro*' })) {
            $icon = [string]$app.DisplayIcon
            if ($icon) {
                $icon = $icon.Trim().Trim('"')
                $icon = $icon -replace ',\s*\d+$',''
                Add-MaestroCandidate $icon
            }
            $loc = [string]$app.InstallLocation
            if ($loc -and (Test-Path -LiteralPath $loc -PathType Container)) {
                foreach ($f in @(Get-ChildItem -LiteralPath $loc -Filter 'AppMain.exe' -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 2)) {
                    Add-MaestroCandidate $f.FullName
                }
            }
        }
    }

    # Common SCM install roots, used only as a narrow fallback.
    foreach ($root in @(
        (Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles)) 'SCM Group'),
        (Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFilesX86)) 'SCM Group'),
        'C:\SCM Group'
    )) {
        if ($root -and (Test-Path -LiteralPath $root -PathType Container)) {
            foreach ($f in @(Get-ChildItem -LiteralPath $root -Filter 'AppMain.exe' -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 2)) {
                Add-MaestroCandidate $f.FullName
            }
        }
    }

    if ($candidates.Count -gt 0) {
        $preferred = @($candidates.ToArray() | Sort-Object @{Expression={ if ([System.IO.Path]::GetFileNameWithoutExtension($_) -like '*MaestroLab*') {0} elseif ([System.IO.Path]::GetFileNameWithoutExtension($_) -like '*Maestro*') {1} else {2} }}, @{Expression={$_}} | Select-Object -First 1)
        if ($preferred.Count -gt 0) { Save-MaestroLauncher $preferred[0]; return $preferred[0] }
    }
    return $null
}

function Select-MaestroLauncherInteractive {
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Избери MaestroLab / MAESTRO LAB за еднократно запомняне'
        $dlg.Filter = 'MaestroLab (*.lnk;*.exe;*.appref-ms)|*.lnk;*.exe;*.appref-ms|Всички файлове (*.*)|*.*'
        $desktop = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        if ($desktop -and (Test-Path -LiteralPath $desktop)) { $dlg.InitialDirectory = $desktop }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Save-MaestroLauncher $dlg.FileName
            return $dlg.FileName
        }
    } catch {}
    return $null
}

function Start-OrActivateMaestro {
    if (-not (Initialize-MixAutomation)) { throw ('Не може да се зареди автоматизацията: ' + $script:MixAutomationError) }
    $p = Get-MaestroProcess
    if ($null -eq $p) {
        $launcher = Find-MaestroLauncher
        if (-not $launcher) { $launcher = Select-MaestroLauncherInteractive }
        if (-not $launcher) { throw 'Не намирам MaestroLab автоматично и не е избран файл за стартиране.' }
        $script:LastMaestroLauncher = $launcher
        Write-MixLog ('Starting MaestroLab: ' + $launcher)
        try {
            Start-Process -FilePath $launcher -ErrorAction Stop | Out-Null
        } catch {
            try { Remove-Item -LiteralPath (Join-Path $PSScriptRoot 'MAESTROLAB_PATH.txt') -Force -ErrorAction SilentlyContinue } catch {}
            throw ('Не мога да стартирам MaestroLab от: ' + $launcher + ' | ' + $_.Exception.Message)
        }
        $limit = (Get-Date).AddSeconds(60)
        do {
            [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 500
            $p = Get-MaestroProcess
        } while ($null -eq $p -and (Get-Date) -lt $limit)
        if ($null -eq $p) { throw 'MaestroLab беше стартиран, но прозорец MAESTRO LAB не се появи до 60 секунди.' }
    }
    # Maestro can replace the first main window during startup. Re-resolve it before use.
    Start-Sleep -Seconds 2
    $current = Get-MaestroProcess
    if ($null -ne $current) { $p = $current }
    $p.Refresh()
    $hNow = [IntPtr]$p.MainWindowHandle
    if ($hNow -ne [IntPtr]::Zero) {
        [void][PWXMixNative]::ShowWindow($hNow,3)
        [void][PWXMixNative]::SetForegroundWindow($hNow)
    }
    Start-Sleep -Seconds 1

    # Keep the exact launcher/executable visible for diagnostics instead of hiding it.
    if (-not $script:LastMaestroLauncher) {
        try { $script:LastMaestroLauncher = [string]$p.Path } catch {}
    }
    return $p
}

function Get-MixUiElementsByName {
    param([int]$ProcessId,[string]$Name)
    try {
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $pc = New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
        $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$pc)
        $result = New-Object System.Collections.Generic.List[object]
        foreach ($el in $all) {
            try {
                if ([string]$el.Current.Name -eq $Name) {
                    $b = $el.Current.BoundingRectangle
                    if ($b.Width -gt 1 -and $b.Height -gt 1) { $result.Add($el) }
                }
            } catch {}
        }
        return @($result.ToArray())
    } catch { return @() }
}

function Invoke-MixUiElementByName {
    param([int]$ProcessId,[string]$Name,[int]$TimeoutSec=8)
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do {
        foreach($el in @(Get-MixUiElementsByName -ProcessId $ProcessId -Name $Name)) {
            try {
                $pat=$null
                if($el.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern,[ref]$pat)) { $pat.Invoke(); return $true }
            } catch {}
            try {
                $b=$el.Current.BoundingRectangle
                Invoke-MixNativeClick -X ([int]($b.Left+$b.Width/2)) -Y ([int]($b.Top+$b.Height/2))
                return $true
            } catch {}
            try {
                $pat=$null
                if($el.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$pat)) { $pat.Select(); return $true }
            } catch {}
        }
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100
    } while((Get-Date) -lt $until)
    return $false
}


function Test-MixUiElementByName {
    param([int]$ProcessId,[string]$Name,[int]$TimeoutSec=5)
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do {
        if (@(Get-MixUiElementsByName -ProcessId $ProcessId -Name $Name).Count -gt 0) { return $true }
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100
    } while((Get-Date) -lt $until)
    return $false
}

function Activate-MixMaestroWindow {
    param([int]$ProcessId)
    try {
        $proc=Get-Process -Id $ProcessId -ErrorAction Stop
        $proc.Refresh()
        $handle=[IntPtr]$proc.MainWindowHandle
        if($handle -eq [IntPtr]::Zero){return $false}
        [void][PWXMixNative]::ShowWindow($handle,3)
        [void][PWXMixNative]::SetForegroundWindow($handle)
        Start-Sleep -Milliseconds 350
        return $true
    } catch { return $false }
}

function Invoke-MixUiElementByParentName {
    param([int]$ProcessId,[string]$ParentName,[int]$TimeoutSec=5)
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do {
        try {
            $root=[System.Windows.Automation.AutomationElement]::RootElement
            $pc=New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
            $all=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$pc)
            $walker=[System.Windows.Automation.TreeWalker]::ControlViewWalker
            foreach($el in $all){
                try {
                    $parent=$walker.GetParent($el)
                    if($null -eq $parent){continue}
                    if([string]$parent.Current.Name -ne $ParentName){continue}
                    $b=$el.Current.BoundingRectangle
                    if($b.Width -le 1 -or $b.Height -le 1){continue}
                    Invoke-MixNativeClick -X ([int]($b.Left+$b.Width/2)) -Y ([int]($b.Top+$b.Height/2))
                    Write-MixLog ('UI parent-node click: ' + $ParentName)
                    return $true
                } catch {}
            }
        } catch {}
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 250
    } while((Get-Date) -lt $until)
    return $false
}

function Invoke-MixFolderNodeByName {
    param([int]$ProcessId,[string]$Name,[int]$TimeoutSec=5)
    # In Maestro Add programs the visible folder row is a Custom child whose
    # Automation parent carries the folder name. In Save as the visible Text
    # element itself carries the name. Support both structures.
    if(Invoke-MixUiElementByName -ProcessId $ProcessId -Name $Name -TimeoutSec ([Math]::Min(2,$TimeoutSec))){
        Write-MixLog ('UI named-node click: ' + $Name)
        return $true
    }
    $left=[Math]::Max(1,$TimeoutSec-2)
    if(Invoke-MixUiElementByParentName -ProcessId $ProcessId -ParentName $Name -TimeoutSec $left){ return $true }
    return $false
}

function Test-MixExpectedFileVisible {
    param([int]$ProcessId,[string]$ExpectedFileName)
    if ([string]::IsNullOrWhiteSpace($ExpectedFileName)) { return $true }
    try {
        $short=[System.IO.Path]::GetFileName($ExpectedFileName)
        if (Test-MixUiElementByName -ProcessId $ProcessId -Name $ExpectedFileName -TimeoutSec 1) { return $true }
        if ($short -and (Test-MixUiElementByName -ProcessId $ProcessId -Name $short -TimeoutSec 1)) { return $true }
    } catch {}
    return $false
}

function Find-MixTreeNodeByName {
    param([int]$ProcessId,[string]$Name,[int]$TimeoutSec=4)
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do {
        try {
            $root=[System.Windows.Automation.AutomationElement]::RootElement
            $pc=New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
            $all=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$pc)
            $walker=[System.Windows.Automation.TreeWalker]::ControlViewWalker

            # Prefer the actual TreeItem. In Maestro the visible row is often a
            # Custom child, while its parent is the TreeItem carrying the name.
            foreach($el in $all){
                try {
                    if([string]$el.Current.Name -ne $Name){continue}
                    $ct=[string]$el.Current.ControlType.ProgrammaticName
                    $cl=[string]$el.Current.ClassName
                    if($ct -match 'TreeItem' -or $cl -match 'TreeViewItem') { return $el }
                } catch {}
            }
            foreach($el in $all){
                try {
                    $parent=$walker.GetParent($el)
                    if($null -eq $parent){continue}
                    if([string]$parent.Current.Name -ne $Name){continue}
                    $ct=[string]$parent.Current.ControlType.ProgrammaticName
                    $cl=[string]$parent.Current.ClassName
                    if($ct -match 'TreeItem' -or $cl -match 'TreeViewItem') { return $parent }
                } catch {}
            }
            # Fallback for breadcrumb/Text elements if the TreeItem is not
            # exposed in the current UI Automation snapshot.
            foreach($el in $all){
                try { if([string]$el.Current.Name -eq $Name){ return $el } } catch {}
            }
        } catch {}
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100
    } while((Get-Date) -lt $until)
    return $null
}

function Focus-ExpandMixTreeNode {
    param([int]$ProcessId,[string]$Name,[int]$TimeoutSec=5)
    $el=Find-MixTreeNodeByName -ProcessId $ProcessId -Name $Name -TimeoutSec $TimeoutSec
    if($null -eq $el){ return $false }
    try {
        $sp=$null
        if($el.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$sp)) { $sp.ScrollIntoView() }
    } catch {}
    try {
        $sel=$null
        if($el.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$sel)) { $sel.Select() }
    } catch {}
    try { $el.SetFocus() } catch {}
    try {
        $ep=$null
        if($el.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern,[ref]$ep)) {
            if($ep.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Collapsed) { $ep.Expand() }
            Start-Sleep -Milliseconds 450
            Write-MixLog ('Tree expanded/focused: ' + $Name)
            return $true
        }
    } catch {}
    try {
        $b=$el.Current.BoundingRectangle
        if($b.Width -gt 1 -and $b.Height -gt 1) {
            Invoke-MixNativeClick -X ([int]($b.Left+$b.Width/2)) -Y ([int]($b.Top+$b.Height/2))
            Start-Sleep -Milliseconds 250
            [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}')
            Start-Sleep -Milliseconds 450
            Write-MixLog ('Tree clicked/expanded: ' + $Name)
            return $true
        }
    } catch {}
    return $false
}

function Select-MixExactDailyFolder {
    param([int]$ProcessId,[string]$DailyName,[string]$ExpectedFileName)
    $el=Find-MixTreeNodeByName -ProcessId $ProcessId -Name $DailyName -TimeoutSec 1
    if($null -eq $el){ return $false }
    try {
        $sp=$null
        if($el.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$sp)) { $sp.ScrollIntoView() }
    } catch {}
    try {
        $sel=$null
        if($el.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$sel)) { $sel.Select() }
    } catch {}
    try {
        $b=$el.Current.BoundingRectangle
        if($b.Width -gt 1 -and $b.Height -gt 1) {
            $x=[int]($b.Left+$b.Width/2)
            $y=[int]($b.Top+$b.Height/2)
            # V5.1.33: selecting dd.MM is not enough in Maestro. The folder
            # must be OPENED with a real double click before we continue.
            Invoke-MixNativeClick -X $x -Y $y
            Start-Sleep -Milliseconds 90
            Invoke-MixNativeClick -X $x -Y $y
            Write-MixLog ('Daily folder double-click sent: ' + $DailyName)
        }
    } catch {}
    # V5.1.34 speed-only: do not sleep a fixed 850 ms. Poll for the same
    # success condition and return immediately when Maestro exposes it.
    # The maximum wait remains conservative; only the successful fast path is shorter.
    if ([string]::IsNullOrWhiteSpace($ExpectedFileName)) {
        # Save-as path verification is handled by Open-MixSaveDailyFolder. Keep the
        # proven settle time here; speed-only changes must not alter this behavior.
        Start-Sleep -Milliseconds 850
        return $true
    }
    $verifyUntil=(Get-Date).AddMilliseconds(1400)
    do {
        if(Test-MixExpectedFileVisible -ProcessId $ProcessId -ExpectedFileName $ExpectedFileName) {
            Write-MixLog ('Exact daily folder opened and verified: ' + $DailyName)
            return $true
        }
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 90
    } while((Get-Date) -lt $verifyUntil)
    return $false
}

function Select-MixDailyFolderInDialogDesktopLegacy {
    param([int]$ProcessId,[string]$DailyName,[string]$ExpectedFileName='',[string]$DailyPath='')

    # Hard rule: the target is the exact filesystem folder
    # Desktop\\FOUND\\dd.MM. We do not trust whichever date Maestro happens
    # to have selected from a previous run.
    if([string]::IsNullOrWhiteSpace($DailyPath)) { $DailyPath=Get-DailyMixFolder }
    $foundRoot=Split-Path -Parent $DailyPath
    $foundName=Split-Path -Leaf $foundRoot
    $desktopRoot=Split-Path -Parent $foundRoot
    $desktopName=Split-Path -Leaf $desktopRoot
    if(-not [System.IO.Directory]::Exists($foundRoot)) { throw "Липсва основната Mix папка: $foundRoot" }
    if(-not [System.IO.Directory]::Exists($DailyPath)) { throw "Липсва дневната Mix папка: $DailyPath" }
    Write-MixLog ('Navigate exact path: ' + $DailyPath)

    # 1) Explicitly open the Desktop branch. If FOUND is not currently loaded
    # in the side tree, expanding Desktop forces Maestro to populate it.
    [void](Activate-MixMaestroWindow -ProcessId $ProcessId)
    if(-not (Focus-ExpandMixTreeNode -ProcessId $ProcessId -Name $desktopName -TimeoutSec 5)) {
        Write-MixLog ('Desktop tree node not exposed yet: ' + $desktopName)
    }
    Start-Sleep -Milliseconds 500

    # 2) Find and expand FOUND specifically under the now-loaded Desktop tree.
    if(-not (Focus-ExpandMixTreeNode -ProcessId $ProcessId -Name $foundName -TimeoutSec 6)) {
        # Fallback: clicking Desktop once can make virtualized children appear.
        [void](Invoke-MixFolderNodeByName -ProcessId $ProcessId -Name $desktopName -TimeoutSec 2)
        Start-Sleep -Milliseconds 500
        if(-not (Focus-ExpandMixTreeNode -ProcessId $ProcessId -Name $foundName -TimeoutSec 5)) { return $false }
    }
    Start-Sleep -Milliseconds 650

    # 3) Search only inside the expanded FOUND branch for the exact dd.MM.
    # First use UI Automation, then WPF incremental search/scroll to realize
    # virtualized children that are not initially visible.
    if(Select-MixExactDailyFolder -ProcessId $ProcessId -DailyName $DailyName -ExpectedFileName $ExpectedFileName) { return $true }

    $foundEl=Find-MixTreeNodeByName -ProcessId $ProcessId -Name $foundName -TimeoutSec 2
    if($null -ne $foundEl){
        try { $foundEl.SetFocus() } catch {}
        try { [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}') } catch {}
        Start-Sleep -Milliseconds 300
        try { [System.Windows.Forms.SendKeys]::SendWait($DailyName) } catch {}
        Start-Sleep -Milliseconds 500
        if(Select-MixExactDailyFolder -ProcessId $ProcessId -DailyName $DailyName -ExpectedFileName $ExpectedFileName) { return $true }

        try { [System.Windows.Forms.SendKeys]::SendWait('{HOME}') } catch {}
        Start-Sleep -Milliseconds 250
        for($i=0;$i -lt 40;$i++) {
            if(Select-MixExactDailyFolder -ProcessId $ProcessId -DailyName $DailyName -ExpectedFileName $ExpectedFileName) { return $true }
            try { [System.Windows.Forms.SendKeys]::SendWait('{PGDN}') } catch {}
            Start-Sleep -Milliseconds 180
        }
    }

    Write-MixLog ('Exact daily folder not found in Maestro tree although it exists on disk: ' + $DailyPath)
    return $false
}


function Test-PathUnderFolder {
    param([string]$Path,[string]$Parent)
    try{
        $p=(Normalize-PathForCompare ([System.IO.Path]::GetFullPath($Path))).TrimEnd('\')
        $r=(Normalize-PathForCompare ([System.IO.Path]::GetFullPath($Parent))).TrimEnd('\')
        return ($p -eq $r -or $p.StartsWith($r+'\',[System.StringComparison]::OrdinalIgnoreCase))
    }catch{return $false}
}

function Get-MixTreeItemsForProcess {
    param([int]$ProcessId)
    try{
        $root=[System.Windows.Automation.AutomationElement]::RootElement
        $pc=New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
        $all=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$pc)
        return @($all | Where-Object { try{$_.Current.ControlType -eq [System.Windows.Automation.ControlType]::TreeItem}catch{$false} })
    }catch{return @()}
}

function Find-MixTreeDriveNode {
    param([int]$ProcessId,[string]$Drive,[int]$TimeoutSec=5)
    $d=$Drive.TrimEnd('\').ToUpperInvariant()
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do{
        foreach($el in @(Get-MixTreeItemsForProcess -ProcessId $ProcessId)){
            try{
                $name=([string]$el.Current.Name).Trim()
                if($name -eq $d -or $name -match ('(?i)(?:^|[\s\(])'+[regex]::Escape($d)+'(?:[\s\)]|$)')){return $el}
            }catch{}
        }
        [System.Windows.Forms.Application]::DoEvents();Start-Sleep -Milliseconds 120
    }while((Get-Date)-lt $until)
    return $null
}

function Find-MixTreeChildNode {
    param($Parent,[string]$Name,[int]$TimeoutSec=4)
    if($null -eq $Parent){return $null}
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do{
        try{
            $walker=[System.Windows.Automation.TreeWalker]::ControlViewWalker
            $all=$Parent.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
            foreach($el in $all){
                try{
                    if($el.Current.ControlType -ne [System.Windows.Automation.ControlType]::TreeItem -or [string]$el.Current.Name -ne $Name){continue}
                    $par=$walker.GetParent($el)
                    while($null -ne $par -and $par.Current.ControlType -ne [System.Windows.Automation.ControlType]::TreeItem){$par=$walker.GetParent($par)}
                    if($null -ne $par -and [System.Windows.Automation.Automation]::Compare($par,$Parent)){return $el}
                }catch{}
            }
        }catch{}
        [System.Windows.Forms.Application]::DoEvents();Start-Sleep -Milliseconds 120
    }while((Get-Date)-lt $until)
    return $null
}

function Focus-ExpandMixTreeElement {
    param($Element,[string]$LogName='')
    if($null -eq $Element){return $false}
    try{$sp=$null;if($Element.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$sp)){$sp.ScrollIntoView()}}catch{}
    try{$sel=$null;if($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$sel)){$sel.Select()}}catch{}
    try{$Element.SetFocus()}catch{}
    try{
        $ep=$null
        if($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern,[ref]$ep)){
            if($ep.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Collapsed){$ep.Expand()}
            Start-Sleep -Milliseconds 350
            Write-MixLog ('Tree expanded/focused exact element: '+$LogName)
            return $true
        }
    }catch{}
    try{
        $b=$Element.Current.BoundingRectangle
        Invoke-MixNativeClick -X ([int]($b.Left+$b.Width/2)) -Y ([int]($b.Top+$b.Height/2))
        Start-Sleep -Milliseconds 120
        [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}')
        Start-Sleep -Milliseconds 350
        return $true
    }catch{return $false}
}

function Open-MixTreeAbsoluteDailyPath {
    param([int]$ProcessId,[string]$DailyPath,[string]$ExpectedFileName='')
    $full=[System.IO.Path]::GetFullPath($DailyPath).TrimEnd('\')
    if($full.StartsWith('\\')){return $false}
    $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    $node=$null;$parts=@()
    [void](Activate-MixMaestroWindow -ProcessId $ProcessId)
    if(Test-PathUnderFolder -Path $full -Parent $desktop){
        $desktopName=Split-Path -Leaf $desktop
        $node=Find-MixTreeNodeByName -ProcessId $ProcessId -Name $desktopName -TimeoutSec 5
        if($null -eq $node){$node=Find-MixTreeNodeByName -ProcessId $ProcessId -Name 'Desktop' -TimeoutSec 3}
        if($null -eq $node){Write-MixLog ('Desktop root not exposed in Maestro tree.');return $false}
        if(-not (Focus-ExpandMixTreeElement -Element $node -LogName 'Desktop')){return $false}
        $relative=$full.Substring($desktop.TrimEnd('\').Length).Trim('\')
        if($relative){$parts=@($relative -split '\\')}
    }else{
        $root=[System.IO.Path]::GetPathRoot($full)
        if([string]::IsNullOrWhiteSpace($root)){return $false}
        $drive=$root.TrimEnd('\')
        $relative=$full.Substring($root.Length).Trim('\')
        if($relative){$parts=@($relative -split '\\')}
        $node=Find-MixTreeDriveNode -ProcessId $ProcessId -Drive $drive -TimeoutSec 3
        if($null -eq $node){
            # A collapsed This PC/Computer node can virtualize its drives. Expand it
            # once, then retry the exact drive lookup.
            foreach($pcName in @('This PC','Computer','This computer','Този компютър')){
                if(Focus-ExpandMixTreeNode -ProcessId $ProcessId -Name $pcName -TimeoutSec 1){break}
            }
            $node=Find-MixTreeDriveNode -ProcessId $ProcessId -Drive $drive -TimeoutSec 4
        }
        if($null -eq $node){Write-MixLog ('Mapped/local drive not exposed in Maestro tree: '+$drive);return $false}
        if(-not (Focus-ExpandMixTreeElement -Element $node -LogName $drive)){return $false}
    }
    if($parts.Count -eq 0){return $false}
    for($i=0;$i -lt $parts.Count;$i++){
        $part=[string]$parts[$i]
        $child=Find-MixTreeChildNode -Parent $node -Name $part -TimeoutSec 5
        if($null -eq $child){Write-MixLog ('Path component not found under current tree node: '+$part+'; full='+$full);return $false}
        if($i -lt ($parts.Count-1)){
            if(-not (Focus-ExpandMixTreeElement -Element $child -LogName $part)){return $false}
            $node=$child
            continue
        }
        try{
            $sp=$null;if($child.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$sp)){$sp.ScrollIntoView()}
            $child.SetFocus();$b=$child.Current.BoundingRectangle
            $x=[int]($b.Left+$b.Width/2);$y=[int]($b.Top+$b.Height/2)
            Invoke-MixNativeClick -X $x -Y $y;Start-Sleep -Milliseconds 90;Invoke-MixNativeClick -X $x -Y $y
            Write-MixLog ('Configured daily folder double-click sent: '+$full)
        }catch{return $false}
        if([string]::IsNullOrWhiteSpace($ExpectedFileName)){Start-Sleep -Milliseconds 850;return $true}
        $until=(Get-Date).AddMilliseconds(1600)
        do{
            if(Test-MixExpectedFileVisible -ProcessId $ProcessId -ExpectedFileName $ExpectedFileName){return $true}
            [System.Windows.Forms.Application]::DoEvents();Start-Sleep -Milliseconds 90
        }while((Get-Date)-lt $until)
        return $false
    }
    return $false
}

function Select-MixDailyFolderInDialog {
    param([int]$ProcessId,[string]$DailyName,[string]$ExpectedFileName='',[string]$DailyPath='')
    if([string]::IsNullOrWhiteSpace($DailyPath)){$DailyPath=Get-DailyMixFolder}
    if(-not [System.IO.Directory]::Exists($DailyPath)){throw "Липсва дневната Mix папка: $DailyPath"}
    if(-not [string]::IsNullOrWhiteSpace($script:BatchOutputFolderOverride) -and
       (Normalize-PathForCompare $DailyPath) -eq (Normalize-PathForCompare $script:BatchOutputFolderOverride)){
        return (Open-MixBatchFolderInDialog -ProcessId $ProcessId -BatchPath $DailyPath -ExpectedFileName $ExpectedFileName)
    }
    $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    $dailyParent=Split-Path -Parent $DailyPath
    $dailyGrandParent=Split-Path -Parent $dailyParent
    if((Normalize-PathForCompare $dailyGrandParent) -eq (Normalize-PathForCompare $desktop)){
        # Preserve the exact, field-tested Desktop\ROOT\dd.MM path logic unchanged.
        return (Select-MixDailyFolderInDialogDesktopLegacy -ProcessId $ProcessId -DailyName $DailyName -ExpectedFileName $ExpectedFileName -DailyPath $DailyPath)
    }
    Write-MixLog ('Navigate configured daily path: '+$DailyPath)
    return (Open-MixTreeAbsoluteDailyPath -ProcessId $ProcessId -DailyPath $DailyPath -ExpectedFileName $ExpectedFileName)
}

# V5.1.38 BUILD FIX4: the batch folder is nested under ROOT\Общи_пакети.
# Maestro virtualizes the left Tree. Navigate nested batch folders using the
# actual center folder list, and verify the OPEN breadcrumb after each double
# click. This is batch-only; the field-tested daily dd.MM path remains intact.
function Get-MixBatchBrowserSteps {
    param([string]$RootPath,[string]$BatchPath)
    if([string]::IsNullOrWhiteSpace($RootPath) -or [string]::IsNullOrWhiteSpace($BatchPath)){throw 'Липсва корен или папка за общия MIXX.'}
    $root=[System.IO.Path]::GetFullPath($RootPath).TrimEnd('\')
    $batch=[System.IO.Path]::GetFullPath($BatchPath).TrimEnd('\')
    $base=Join-Path $root 'Общи_пакети'
    if([System.IO.Path]::GetDirectoryName($batch).TrimEnd('\') -ine $base){throw ('Невалиден общ пакет извън Общи_пакети: '+$batch)}
    if(-not [System.IO.Directory]::Exists($batch)){throw ('Липсва обща папка: '+$batch)}
    return @($root,$base,$batch)
}

function Get-MixFileBrowserSnapshot {
    param([int]$ProcessId)
    $s=[pscustomobject]@{IsActive=$false;Scope=$null;ContentList=$null;BreadcrumbTokens=@();DesktopAnchorIsRoot=$false;CurrentAbsolutePath='';DialogTitle=''}
    try{
        $hwnd=[PWXMixNative]::GetForegroundWindow()
        if($hwnd -eq [IntPtr]::Zero){return $s}
        $foreground=[System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
        if($foreground.Current.ProcessId -ne $ProcessId){return $s}
        $visible=@($foreground.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)|Where-Object{
            try{$b=$_.Current.BoundingRectangle; -not $_.Current.IsOffscreen -and $b.Width -gt 1 -and $b.Height -gt 1}catch{$false}
        })
        $titles=@($visible|Where-Object{([string]$_.Current.Name) -in @('Add programs','Choose the file','Choose file','Select file') -and $_.Current.BoundingRectangle.Top -lt ($foreground.Current.BoundingRectangle.Bottom-120)}|Sort-Object{$_.Current.BoundingRectangle.Top})
        $walker=[System.Windows.Automation.TreeWalker]::RawViewWalker
        foreach($title in $titles){
            $titleBox=$title.Current.BoundingRectangle
            $scope=$walker.GetParent($title)
            for($depth=0;$null -ne $scope -and $depth -lt 16;$depth++){
                $elements=@($scope.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)|Where-Object{
                    try{$b=$_.Current.BoundingRectangle; -not $_.Current.IsOffscreen -and $b.Width -gt 1 -and $b.Height -gt 1}catch{$false}
                })
                $trees=@($elements|Where-Object{
                    $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Tree -and $_.Current.BoundingRectangle.Left -lt ($titleBox.Left+200) -and $_.Current.BoundingRectangle.Top -gt $titleBox.Bottom
                }|Sort-Object{$_.Current.BoundingRectangle.Left}|Select-Object -First 1)
                # The Add-programs Confirm button is DISABLED until a PGMX is selected.
                # Use it as a geometry anchor even while disabled; never click it here.
                $confirm=@($elements|Where-Object{([string]$_.Current.Name) -eq 'Confirm' -and $_.Current.BoundingRectangle.Top -gt ($titleBox.Bottom+130)}|Sort-Object{$_.Current.BoundingRectangle.Top} -Descending|Select-Object -First 1)
                if($trees.Count -eq 1 -and $confirm.Count -eq 1){
                    $tree=$trees[0];$treeBox=$tree.Current.BoundingRectangle;$confirmBox=$confirm[0].Current.BoundingRectangle
                    $lists=@($elements|Where-Object{
                        $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::List -and
                        $_.Current.BoundingRectangle.Left -ge ($treeBox.Right-10) -and
                        $_.Current.BoundingRectangle.Top -ge ($treeBox.Top-25) -and
                        $_.Current.BoundingRectangle.Bottom -le ($confirmBox.Top+10) -and
                        $_.Current.BoundingRectangle.Left -lt ($confirmBox.Right-130)
                    }|Sort-Object{$_.Current.BoundingRectangle.Left}|Select-Object -First 1)
                    if($lists.Count -eq 1){
                        $list=$lists[0]
                        $titleHit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($titleBox.Left+$titleBox.Width/2),($titleBox.Top+$titleBox.Height/2))))
                        $confirmHit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($confirmBox.Left+$confirmBox.Width/2),($confirmBox.Top+$confirmBox.Height/2))))
                        $treeHit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($treeBox.Left+$treeBox.Width/2),($treeBox.Top+$treeBox.Height/2))))
                        $listBox=$list.Current.BoundingRectangle
                        $listHit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($listBox.Left+$listBox.Width/2),($listBox.Top+$listBox.Height/2))))
                        # FIX6: Add programs has a DISABLED Confirm button until PGMX are
                        # selected. UIA FromPoint may hit its parent rather than the
                        # disabled button: requiring a button hit rejects a valid
                        # Add programs dialog while it is still at Desktop > FOUND.
                        # Check button hit only when enabled; otherwise enforce its
                        # bounded geometry within this SAME verified dialog scope.
                        $confirmAnchorValid=$false
                        if($confirm[0].Current.IsEnabled){
                            $confirmAnchorValid=Test-MixElementWithin $confirmHit $confirm[0]
                        }else{
                            $scopeBox=$scope.Current.BoundingRectangle
                            $confirmAnchorValid=($confirmBox.Width -gt 10 -and $confirmBox.Height -gt 10 -and
                                $confirmBox.Left -ge ($scopeBox.Left-5) -and $confirmBox.Right -le ($scopeBox.Right+5) -and
                                $confirmBox.Top -ge ($titleBox.Bottom+130) -and $confirmBox.Bottom -le ($scopeBox.Bottom+5))
                        }
                        if(-not (Test-MixElementWithin $titleHit $title) -or
                           -not $confirmAnchorValid -or
                           -not (Test-MixElementWithin $treeHit $tree) -or
                           -not (Test-MixElementWithin $listHit $list)){break}
                        $s.Scope=$scope;$s.ContentList=$list;$s.IsActive=$true;$s.DialogTitle=[string]$title.Current.Name
                        # Require the Desktop root to be a true root of the navigation tree.
                        $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
                        $desktopLeaf=Split-Path -Leaf $desktop
                        $treeNodes=@($tree.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)|Where-Object{
                            $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::TreeItem -and ([string]$_.Current.Name) -in @('Desktop',$desktopLeaf)
                        })
                        foreach($node in $treeNodes){
                            $par=$walker.GetParent($node)
                            for($p=0;$null -ne $par -and $p -lt 16;$p++){
                                if([System.Windows.Automation.Automation]::Compare($par,$tree)){$s.DesktopAnchorIsRoot=$true;break}
                                if($par.Current.ControlType -eq [System.Windows.Automation.ControlType]::TreeItem){break}
                                $par=$walker.GetParent($par)
                            }
                        }
                        # The breadcrumb is a toolbar ABOVE the tree, never a selected
                        # central folder row or a preview pane label.
                        $crumbs=New-Object System.Collections.Generic.List[object]
                        $listBox=$list.Current.BoundingRectangle
                        foreach($el in $elements){
                            try{
                                $b=$el.Current.BoundingRectangle;$ct=$el.Current.ControlType
                                if($b.Top -le $titleBox.Bottom -or $b.Bottom -ge $treeBox.Top -or $b.Left -lt ($treeBox.Left+60) -or $b.Right -gt ($listBox.Right+5)){continue}
                                if($ct -eq [System.Windows.Automation.ControlType]::Edit){
                                    $v=$null
                                    if($el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$v)){
                                        $abs=ConvertTo-MixWindowsFolderPath ([string]$v.Current.Value)
                                        if($abs -and -not $s.CurrentAbsolutePath){$s.CurrentAbsolutePath=$abs}
                                    }
                                    continue
                                }
                                if($ct -notin @([System.Windows.Automation.ControlType]::Text,[System.Windows.Automation.ControlType]::Hyperlink,[System.Windows.Automation.ControlType]::Button)){continue}
                                $name=([string]$el.Current.Name).Trim()
                                if(-not $name -or $name -match '^(?:Search|>|›|»)$'){continue}
                                $hit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($b.Left+$b.Width/2),($b.Top+$b.Height/2))))
                                if(-not (Test-MixElementWithin $hit $el)){continue}
                                $duplicate=$false
                                foreach($c in $crumbs){if($c.Name -eq $name -and $b.Left -lt $c.Right -and $b.Right -gt $c.Left){$duplicate=$true;break}}
                                if(-not $duplicate){$crumbs.Add([pscustomobject]@{Name=$name;Left=$b.Left;Right=$b.Right})}
                            }catch{}
                        }
                        $parts=New-Object System.Collections.Generic.List[string]
                        foreach($crumb in @($crumbs.ToArray()|Sort-Object Left)){
                            foreach($part in @($crumb.Name -split '\s*[>›»]\s*')){if($part.Trim()){$parts.Add($part.Trim())}}
                        }
                        $s.BreadcrumbTokens=@($parts.ToArray())
                        return $s
                    }
                }
                $scope=$walker.GetParent($scope)
            }
        }
    }catch{Write-MixLog ('BATCH file-browser snapshot failed: '+$_.Exception.Message)}
    return $s
}

function Get-MixBatchBrowserSnapshot {
    param([int]$ProcessId,[ValidateSet('Add','Choose','Save')][string]$Mode)
    if($Mode -eq 'Save'){return (Get-MixSaveFolderSnapshot -ProcessId $ProcessId)}
    $s=Get-MixFileBrowserSnapshot -ProcessId $ProcessId
    if(-not $s.IsActive){return $s}
    if($Mode -eq 'Add' -and $s.DialogTitle -ne 'Add programs'){$s.IsActive=$false}
    if($Mode -eq 'Choose' -and $s.DialogTitle -notin @('Choose the file','Choose file','Select file')){$s.IsActive=$false}
    return $s
}

function Get-MixBrowserObservedPath {
    param($Snapshot,[string]$DesktopPath)
    if($null -eq $Snapshot -or -not $Snapshot.IsActive -or $null -eq $Snapshot.ContentList){return ''}
    $tokens=@($Snapshot.BreadcrumbTokens)
    if($tokens.Count -lt 1){return ''}
    $desktop=ConvertTo-MixWindowsFolderPath $DesktopPath
    $first=([string]$tokens[0]).Trim()
    $desktopLeaf=Split-Path -Leaf $DesktopPath
    $base='';$start=0
    if($first -ieq 'Desktop' -or $first -ieq $desktopLeaf){
        if(-not $Snapshot.DesktopAnchorIsRoot -or -not $desktop){return ''}
        $base=$desktop;$start=1
    }else{
        for($i=0;$i -lt $tokens.Count;$i++){
            if(([string]$tokens[$i]).Trim() -match '^[A-Za-z]:$'){$base=([string]$tokens[$i]).Trim()+'\';$start=$i+1;break}
        }
    }
    if(-not $base){return ''}
    $parts=New-Object System.Collections.Generic.List[string]
    for($i=$start;$i -lt $tokens.Count;$i++){
        $part=([string]$tokens[$i]).Trim()
        if(-not $part -or $part -eq '.' -or $part -eq '..' -or $part.IndexOfAny([char[]]'\/:') -ge 0){return ''}
        $parts.Add($part)
    }
    $path=$base
    foreach($part in $parts){$path=Join-Path $path $part}
    $observed=ConvertTo-MixWindowsFolderPath $path
    if(-not $observed){return ''}
    $abs=ConvertTo-MixWindowsFolderPath ([string]$Snapshot.CurrentAbsolutePath)
    if($abs -and $abs -ne $observed){return ''}
    return $observed
}

function Find-MixBatchCenterFolderItem {
    param($List,[string]$FolderName)
    if($null -eq $List -or [string]::IsNullOrWhiteSpace($FolderName)){return $null}
    # First try the provider's virtual-item lookup: avoid traversing or
    # scrolling dozens of UIA pages when this is supported by Maestro.
    try{
        $container=$null
        if($List.TryGetCurrentPattern([System.Windows.Automation.ItemContainerPattern]::Pattern,[ref]$container)){
            $item=$container.FindItemByProperty($null,[System.Windows.Automation.AutomationElement]::NameProperty,$FolderName)
            if($null -ne $item){
                try{$virt=$null;if($item.TryGetCurrentPattern([System.Windows.Automation.VirtualizedItemPattern]::Pattern,[ref]$virt)){$virt.Realize()}}catch{}
                try{$sp=$null;if($item.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$sp)){$sp.ScrollIntoView()}}catch{}
                $b=$item.Current.BoundingRectangle;$lb=$List.Current.BoundingRectangle
                if(-not $item.Current.IsOffscreen -and $b.Width -gt 5 -and $b.Height -gt 5 -and
                   $b.Left -ge ($lb.Left-5) -and $b.Right -le ($lb.Right+5) -and
                   $b.Top -ge ($lb.Top-5) -and $b.Bottom -le ($lb.Bottom+5)){
                    Write-MixLog ('BATCH folder found with ItemContainer: '+$FolderName)
                    return $item
                }
            }
        }
    }catch{}
    $scroll=$null;$supportsScroll=$false
    try{
        if($List.TryGetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern,[ref]$scroll) -and $scroll.Current.VerticallyScrollable){
            $supportsScroll=$true
            $scroll.SetScrollPercent(-1,0)
        }
    }catch{$scroll=$null;$supportsScroll=$false}
    # FIX6 fallback: some Maestro WPF Lists do NOT expose ScrollPattern.
    # In this case use its keyboard page navigation instead of stopping
    # after only the initially visible folders. Never click a guessed row.
    if(-not $supportsScroll){
        try{$List.SetFocus();[System.Windows.Forms.SendKeys]::SendWait('{HOME}') }catch{}
    }
    $previousPage='';$samePages=0
    for($page=0;$page -lt 32;$page++){
        Test-OperatorStopRequested
        $visibleNames=New-Object System.Collections.Generic.List[string]
        try{
            $lb=$List.Current.BoundingRectangle
            foreach($el in $List.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)){
                try{
                    if($el.Current.ControlType -ne [System.Windows.Automation.ControlType]::ListItem -or $el.Current.IsOffscreen -or -not $el.Current.IsEnabled){continue}
                    $b=$el.Current.BoundingRectangle
                    if($b.Width -le 5 -or $b.Height -le 5 -or $b.Left -lt ($lb.Left-5) -or
                       $b.Right -gt ($lb.Right+5) -or $b.Top -lt ($lb.Top-5) -or $b.Bottom -gt ($lb.Bottom+5)){continue}
                    $name=[string]$el.Current.Name
                    $visibleNames.Add($name)
                    if($name -ne $FolderName){
                        $texts=$el.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
                        if(@($texts|Where-Object{$_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Text -and [string]$_.Current.Name -eq $FolderName}).Count -eq 0){continue}
                    }
                    Write-MixLog ('BATCH exact center-list folder found: '+$FolderName+'; page='+$page)
                    return $el
                }catch{}
            }
        }catch{return $null}
        $pageSignature=($visibleNames.ToArray() -join '|')
        if($page -gt 0 -and $pageSignature -eq $previousPage){$samePages++}else{$samePages=0}
        if($samePages -ge 2){break}
        $previousPage=$pageSignature
        if($supportsScroll){
            try{
                if(-not $scroll.Current.VerticallyScrollable){break}
                $old=[double]$scroll.Current.VerticalScrollPercent
                $scroll.Scroll([System.Windows.Automation.ScrollAmount]::NoAmount,[System.Windows.Automation.ScrollAmount]::LargeIncrement)
                if([Math]::Abs([double]$scroll.Current.VerticalScrollPercent-$old) -lt 0.0001){break}
            }catch{$supportsScroll=$false}
        }
        if(-not $supportsScroll){
            try{$List.SetFocus();[System.Windows.Forms.SendKeys]::SendWait('{PGDN}')}catch{break}
        }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
    Write-MixLog ('BATCH folder not found in scrolling center list: '+$FolderName)
    return $null
}

function Open-MixBatchFolderInDialog {
    param([int]$ProcessId,[string]$BatchPath,[string]$ExpectedFileName='')
    $mode='Save'
    if(-not [string]::IsNullOrWhiteSpace($ExpectedFileName)){
        if([System.IO.Path]::GetExtension($ExpectedFileName) -ieq '.pgmx'){$mode='Add'}
        elseif([System.IO.Path]::GetExtension($ExpectedFileName) -ieq '.mixx'){$mode='Choose'}
        else{throw ('Непознат тип файл в прозореца Maestro: '+$ExpectedFileName)}
    }
    $steps=@(Get-MixBatchBrowserSteps -RootPath (Get-DailyRootFolder) -BatchPath $BatchPath)
    $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    $target=ConvertTo-MixWindowsFolderPath $BatchPath
    if(-not $target){throw 'Невалиден път за общия MIXX.'}
    [void](Activate-MixMaestroWindow -ProcessId $ProcessId)
    $snapshot=Get-MixBatchBrowserSnapshot -ProcessId $ProcessId -Mode $mode
    $observed=Get-MixBrowserObservedPath -Snapshot $snapshot -DesktopPath $desktop
    Write-MixLog ('BATCH browse '+$mode+' initial='+$observed+'; target='+$target)
    $root=ConvertTo-MixWindowsFolderPath $steps[0]
    $base=ConvertTo-MixWindowsFolderPath $steps[1]
    if($observed -notin @($root,$base,$target)){
        # Re-anchor only the known root. The old generic tree selector was
        # unreliable for the VIRTUALIZED nested batch leaf; it is not used
        # for Обши_пакети or the timestamped package folder any more.
        if(-not (Open-MixTreeAbsoluteDailyPath -ProcessId $ProcessId -DailyPath $steps[0] -ExpectedFileName '')){
            Write-MixLog ('BATCH cannot open configured root: '+$steps[0]);return $false
        }
        $until=(Get-Date).AddSeconds(4)
        do{
            $snapshot=Get-MixBatchBrowserSnapshot -ProcessId $ProcessId -Mode $mode
            $observed=Get-MixBrowserObservedPath -Snapshot $snapshot -DesktopPath $desktop
            if($observed -eq $root){break}
            Start-Sleep -Milliseconds 140
        }while((Get-Date) -lt $until)
        if($observed -ne $root){Write-MixLog ('BATCH root not confirmed in '+$mode+'; observed='+$observed);return $false}
    }
    for($index=1;$index -lt $steps.Count;$index++){
        $next=ConvertTo-MixWindowsFolderPath $steps[$index]
        if($observed -eq $next){continue}
        $parent=ConvertTo-MixWindowsFolderPath $steps[$index-1]
        if($observed -ne $parent){Write-MixLog ('BATCH unexpected folder; need '+$parent+' but open '+$observed);return $false}
        $leaf=[System.IO.Path]::GetFileName($steps[$index])
        $item=Find-MixBatchCenterFolderItem -List $snapshot.ContentList -FolderName $leaf
        if($null -eq $item){Write-MixLog ('BATCH folder not in center list: '+$leaf+' under '+$parent);return $false}
        try{
            if(-not (Activate-MixMaestroWindow -ProcessId $ProcessId)){return $false}
            $check=Get-MixBatchBrowserSnapshot -ProcessId $ProcessId -Mode $mode
            if((Get-MixBrowserObservedPath -Snapshot $check -DesktopPath $desktop) -ne $parent){return $false}
            if(-not (Test-MixElementWithin $item $check.ContentList)){return $false}
            $item.SetFocus();$b=$item.Current.BoundingRectangle
            $x=[int]($b.Left+[Math]::Min(55,$b.Width/2));$y=[int]($b.Top+$b.Height/2)
            $hit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point($x,$y)))
            if(-not (Test-MixElementWithin $hit $item)){return $false}
            Invoke-MixNativeClick -X $x -Y $y
            Start-Sleep -Milliseconds 85
            Invoke-MixNativeClick -X $x -Y $y
            Write-MixLog ('BATCH opened center-list folder by double-click: '+$next)
        }catch{Write-MixLog ('BATCH center-folder activation failed: '+$_.Exception.Message);return $false}
        $until=(Get-Date).AddSeconds(4)
        $observed=''
        do{
            Test-OperatorStopRequested
            Start-Sleep -Milliseconds 130
            $snapshot=Get-MixBatchBrowserSnapshot -ProcessId $ProcessId -Mode $mode
            $observed=Get-MixBrowserObservedPath -Snapshot $snapshot -DesktopPath $desktop
            if($observed -eq $next){break}
            if($observed -and $observed -ne $parent){break}
        }while((Get-Date) -lt $until)
        if($observed -ne $next){Write-MixLog ('BATCH breadcrumb not confirmed: '+$next+'; observed='+$observed);return $false}
    }
    if($observed -ne $target){return $false}
    if($ExpectedFileName -and -not (Test-MixExpectedFileVisible -ProcessId $ProcessId -ExpectedFileName $ExpectedFileName)){
        Write-MixLog ('BATCH expected file not visible in verified folder: '+$ExpectedFileName)
        return $false
    }
    Write-MixLog ('BATCH exact folder VERIFIED for '+$mode+': '+$target)
    return $true
}

function Get-MixPgmxListItems {
    param([int]$ProcessId)
    try {
        $root=[System.Windows.Automation.AutomationElement]::RootElement
        $pc=New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
        $all=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$pc)
        $result=New-Object System.Collections.Generic.List[object]
        foreach($el in $all){
            try {
                if($el.Current.ControlType -ne [System.Windows.Automation.ControlType]::ListItem){ continue }
                $name=[string]$el.Current.Name
                if([string]::IsNullOrWhiteSpace($name) -or -not $name.EndsWith('.pgmx',[System.StringComparison]::OrdinalIgnoreCase)){ continue }
                $b=$el.Current.BoundingRectangle
                if($b.Width -gt 1 -and $b.Height -gt 1){ $result.Add($el) }
            } catch {}
        }
        return @($result.ToArray())
    } catch { return @() }
}

function Get-MixAncestorList {
    param($Element)
    try {
        $walker=[System.Windows.Automation.TreeWalker]::ControlViewWalker
        $cur=$Element
        for($i=0;$i -lt 12 -and $null -ne $cur;$i++){
            try {
                if($cur.Current.ControlType -eq [System.Windows.Automation.ControlType]::List){ return $cur }
            } catch {}
            try { $cur=$walker.GetParent($cur) } catch { $cur=$null }
        }
    } catch {}
    return $null
}

function Test-MixExpectedPgmxSelected {
    param([int]$ProcessId,[string[]]$ExpectedFiles)
    $expected=@($ExpectedFiles | ForEach-Object { [System.IO.Path]::GetFileName([string]$_) } | Sort-Object -Unique)
    if($expected.Count -eq 0){ return $false }
    $items=@(Get-MixPgmxListItems -ProcessId $ProcessId)
    $map=@{}
    foreach($el in $items){
        try {
            $base=[System.IO.Path]::GetFileName([string]$el.Current.Name)
            if(-not [string]::IsNullOrWhiteSpace($base)){ $map[$base]=$el }
        } catch {}
    }
    foreach($name in $expected){ if(-not $map.ContainsKey($name)){ return $false } }

    # Prefer the List selection pattern when Maestro exposes it. This is more
    # reliable for virtualized WPF rows than asking every off-screen ListItem.
    try {
        $list=Get-MixAncestorList -Element $map[$expected[0]]
        if($null -ne $list){
            $sp=$null
            if($list.TryGetCurrentPattern([System.Windows.Automation.SelectionPattern]::Pattern,[ref]$sp)){
                $sel=@($sp.GetCurrentSelection())
                $selectedNames=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                foreach($s in $sel){
                    try {
                        $nm=[System.IO.Path]::GetFileName([string]$s.Current.Name)
                        if($nm.EndsWith('.pgmx',[System.StringComparison]::OrdinalIgnoreCase)){ [void]$selectedNames.Add($nm) }
                    } catch {}
                }
                $allThere=$true
                foreach($name in $expected){ if(-not $selectedNames.Contains($name)){ $allThere=$false; break } }
                if($allThere){ return $true }
            }
        }
    } catch {}

    # Fallback for Maestro builds where SelectionPattern is incomplete.
    foreach($name in $expected){
        try {
            $pat=$null
            if(-not $map[$name].TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$pat)){ return $false }
            if(-not $pat.Current.IsSelected){ return $false }
        } catch { return $false }
    }
    return $true
}

function Select-AllMixPgmxFilesInDialog {
    param([int]$ProcessId,[string[]]$ExpectedFiles,[int]$TimeoutSec=10)
    $expected=@($ExpectedFiles | ForEach-Object { [System.IO.Path]::GetFileName([string]$_) } | Sort-Object -Unique)
    if($expected.Count -eq 0){ return $false }

    # Wait until the expected PGMX rows are exposed by Maestro.
    $until=(Get-Date).AddSeconds($TimeoutSec)
    $items=@(); $map=@{}
    do {
        $items=@(Get-MixPgmxListItems -ProcessId $ProcessId)
        $map=@{}
        foreach($el in $items){
            try {
                $base=[System.IO.Path]::GetFileName([string]$el.Current.Name)
                if(-not [string]::IsNullOrWhiteSpace($base)){ $map[$base]=$el }
            } catch {}
        }
        $missing=@($expected | Where-Object { -not $map.ContainsKey($_) })
        if($missing.Count -eq 0){ break }
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 250
    } while((Get-Date) -lt $until)

    $missing=@($expected | Where-Object { -not $map.ContainsKey($_) })
    if($missing.Count -gt 0){
        Write-MixLog ('PGMX not visible in Add programs: ' + ($missing -join ', '))
        return $false
    }

    $firstName=$expected[0]
    $lastName=$expected[$expected.Count-1]
    $first=$map[$firstName]
    $last=$map[$lastName]

    # Re-activate Maestro immediately before keyboard input. Updating the agent
    # status must never leave keyboard focus in the PowerShell/WinForms window.
    [void](Activate-MixMaestroWindow -ProcessId $ProcessId)
    Start-Sleep -Milliseconds 180

    # Reproduce the operator's learned sequence precisely: click the first row,
    # click its filename Text, then Ctrl+A.
    try {
        $scroll=$null
        if($first.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$scroll)){ $scroll.ScrollIntoView() }
    } catch {}
    try {
        $sel=$null
        if($first.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$sel)){ $sel.Select() }
    } catch {}
    try { $first.SetFocus() } catch {}
    try {
        $b=$first.Current.BoundingRectangle
        if($b.Width -gt 1 -and $b.Height -gt 1){
            # Click near the filename, not in the center/right preview area.
            Invoke-MixNativeClick -X ([int]($b.Left + [Math]::Min(80,[Math]::Max(18,$b.Width*0.15)))) -Y ([int]($b.Top+$b.Height/2))
        }
    } catch {}
    Start-Sleep -Milliseconds 100
    [void](Invoke-MixPgmxRowTextClick -Element $first -FileName $firstName)
    Start-Sleep -Milliseconds 180

    # Primary method: native Ctrl+A, exactly as in all three successful training sessions.
    Invoke-MixNativeCtrlA
    Start-Sleep -Milliseconds 650
    if(Test-MixExpectedPgmxSelected -ProcessId $ProcessId -ExpectedFiles $ExpectedFiles){
        Write-MixLog ('Selected all PGMX with native Ctrl+A. Count=' + $expected.Count)
        return $true
    }

    # Some Maestro/WPF builds accept Ctrl+A only after a second explicit text click.
    [void](Activate-MixMaestroWindow -ProcessId $ProcessId)
    [void](Invoke-MixPgmxRowTextClick -Element $first -FileName $firstName)
    Start-Sleep -Milliseconds 150
    Invoke-MixNativeCtrlA
    Start-Sleep -Milliseconds 600
    if(Test-MixExpectedPgmxSelected -ProcessId $ProcessId -ExpectedFiles $ExpectedFiles){
        Write-MixLog ('Selected all PGMX with second native Ctrl+A. Count=' + $expected.Count)
        return $true
    }

    # Robust mouse/keyboard fallback: select the contiguous range first -> last
    # with Shift+click. The Add programs list contains the PGMX rows in filename
    # order, so this selects every daily program without depending on Ctrl+A.
    try {
        [void](Activate-MixMaestroWindow -ProcessId $ProcessId)
        $sf=$null
        if($first.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$sf)){ $sf.ScrollIntoView() }
        Start-Sleep -Milliseconds 100
        $bf=$first.Current.BoundingRectangle
        if($bf.Width -gt 1 -and $bf.Height -gt 1){
            Invoke-MixNativeClick -X ([int]($bf.Left + [Math]::Min(80,[Math]::Max(18,$bf.Width*0.15)))) -Y ([int]($bf.Top+$bf.Height/2))
        }
        Start-Sleep -Milliseconds 120
        $sl=$null
        if($last.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$sl)){ $sl.ScrollIntoView() }
        Start-Sleep -Milliseconds 180
        $bl=$last.Current.BoundingRectangle
        if($bl.Width -gt 1 -and $bl.Height -gt 1){
            Invoke-MixNativeShiftClick -X ([int]($bl.Left + [Math]::Min(80,[Math]::Max(18,$bl.Width*0.15)))) -Y ([int]($bl.Top+$bl.Height/2))
        }
    } catch {}
    Start-Sleep -Milliseconds 650
    if(Test-MixExpectedPgmxSelected -ProcessId $ProcessId -ExpectedFiles $ExpectedFiles){
        Write-MixLog ('Selected all PGMX with first/last Shift range. Count=' + $expected.Count)
        return $true
    }

    # Last fallback: UI Automation multi-selection where supported.
    $firstDone=$false
    foreach($name in $expected){
        $el=$map[$name]
        try {
            $scroll=$null
            if($el.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern,[ref]$scroll)){ $scroll.ScrollIntoView() }
        } catch {}
        try {
            $pat=$null
            if($el.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$pat)){
                if(-not $firstDone){ $pat.Select(); $firstDone=$true }
                else { $pat.AddToSelection() }
            }
        } catch {}
    }
    Start-Sleep -Milliseconds 400
    if(Test-MixExpectedPgmxSelected -ProcessId $ProcessId -ExpectedFiles $ExpectedFiles){
        Write-MixLog ('Selected all PGMX with UI Automation. Count=' + $expected.Count)
        return $true
    }

    Write-MixLog ('FAILED to select all PGMX after Ctrl+A, Shift-range and UIA. Expected count=' + $expected.Count)
    return $false
}

function ConvertTo-MixWindowsFolderPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p=$Path.Trim().Trim('"').Replace('/','\')
    if ($p -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))') { return '' }
    if (@($p.Split('\') | Where-Object { $_ -eq '.' -or $_ -eq '..' }).Count -gt 0) { return '' }
    return $p.TrimEnd('\').ToLowerInvariant()
}

function Get-MixSaveFolderProofDesktopLegacy {
    param($Snapshot,[string]$ExpectedPath,[string]$DesktopPath,[string]$ExpectedName='')
    $answer=[pscustomobject]@{Verified=$false;Reason='Липсва наблюдаем текущ път в Save as.';ObservedPath=''}
    $expected=ConvertTo-MixWindowsFolderPath $ExpectedPath
    $desktop=ConvertTo-MixWindowsFolderPath $DesktopPath
    if (-not $expected -or -not $desktop -or -not $expected.StartsWith($desktop+'\',[System.StringComparison]::OrdinalIgnoreCase)) { $answer.Reason='Невалиден очакван път под действителния Desktop.';return $answer }
    if ($null -eq $Snapshot -or -not $Snapshot.IsActive -or -not $Snapshot.ScopeId) { $answer.Reason='Save as не е текущият активен диалог.';return $answer }
    foreach ($field in @('NameScopeId','ConfirmScopeId','CurrentFolderScopeId')) {
        if ([string]$Snapshot.$field -ne [string]$Snapshot.ScopeId) { $answer.Reason='Името, Confirm и текущият път не са в един Save-as диалог.';return $answer }
    }
    if (-not $Snapshot.NameVisible -or -not $Snapshot.ConfirmVisible) { $answer.Reason='Липсва видимо поле Name или Confirm в Save as.';return $answer }
    if (-not $Snapshot.DesktopAnchorIsRoot -or (ConvertTo-MixWindowsFolderPath ([string]$Snapshot.DesktopAnchorPath)) -ne $desktop) { $answer.Reason='Desktop в текущия browser не е потвърден като действителния Desktop.';return $answer }
    $tokens=@($Snapshot.BreadcrumbTokens)
    if ($tokens.Count -lt 2) { return $answer }
    $desktopLeaf=($desktop -split '\\')[-1]
    if ([string]$tokens[0] -ine 'Desktop' -and [string]$tokens[0] -ine $desktopLeaf) { $answer.Reason='Текущият breadcrumb не започва от Desktop.';return $answer }
    $parts=New-Object System.Collections.Generic.List[string]
    for ($i=1; $i -lt $tokens.Count; $i++) {
        $part=([string]$tokens[$i]).Trim()
        if (-not $part -or $part -eq '.' -or $part -eq '..' -or $part.IndexOfAny([char[]]'\/:') -ge 0) { $answer.Reason='Невалиден компонент в текущия breadcrumb.';return $answer }
        $parts.Add($part)
    }
    $observed=ConvertTo-MixWindowsFolderPath ($desktop+'\'+($parts.ToArray() -join '\'))
    $answer.ObservedPath=$observed
    # A selected folder/preview is deliberately not evidence of the opened folder.
    if ($observed -ne $expected) { $answer.Reason='Текущата отворена папка е '+$observed+'; очаквана е '+$expected+'.';return $answer }
    if ($Snapshot.CurrentAbsolutePath) {
        $absolute=ConvertTo-MixWindowsFolderPath ([string]$Snapshot.CurrentAbsolutePath)
        if ($absolute -ne $expected) { $answer.Reason='Абсолютният адрес противоречи на очакваната папка: '+[string]$Snapshot.CurrentAbsolutePath;return $answer }
    }
    if ($ExpectedName -and [string]$Snapshot.NameValue -cne $ExpectedName) { $answer.Reason='Името в текущия Save as не е потвърдено: '+[string]$Snapshot.NameValue;return $answer }
    $answer.Verified=$true;$answer.Reason='Потвърдена текуща папка: '+$observed
    return $answer
}


function Get-MixSaveFolderProofGeneric {
    param($Snapshot,[string]$ExpectedPath,[string]$ExpectedName='')
    $answer=[pscustomobject]@{Verified=$false;Reason='Липсва наблюдаем текущ път в Save as.';ObservedPath=''}
    $expected=ConvertTo-MixWindowsFolderPath $ExpectedPath
    if(-not $expected){$answer.Reason='Невалиден очакван път.';return $answer}
    if($null -eq $Snapshot -or -not $Snapshot.IsActive -or -not $Snapshot.ScopeId){$answer.Reason='Save as не е текущият активен диалог.';return $answer}
    foreach($field in @('NameScopeId','ConfirmScopeId','CurrentFolderScopeId')){if([string]$Snapshot.$field -ne [string]$Snapshot.ScopeId){$answer.Reason='Името, Confirm и текущият път не са в един Save-as диалог.';return $answer}}
    if(-not $Snapshot.NameVisible -or -not $Snapshot.ConfirmVisible){$answer.Reason='Липсва видимо поле Name или Confirm в Save as.';return $answer}
    if($Snapshot.CurrentAbsolutePath){
        $absolute=ConvertTo-MixWindowsFolderPath ([string]$Snapshot.CurrentAbsolutePath)
        $answer.ObservedPath=$absolute
        if($absolute -ne $expected){$answer.Reason='Текущият абсолютен адрес е '+$absolute+'; очакван е '+$expected+'.';return $answer}
        if($ExpectedName -and [string]$Snapshot.NameValue -cne $ExpectedName){$answer.Reason='Името в текущия Save as не е потвърдено: '+[string]$Snapshot.NameValue;return $answer}
        $answer.Verified=$true;$answer.Reason='Потвърдена текуща папка по абсолютен адрес: '+$absolute;return $answer
    }
    $full=[System.IO.Path]::GetFullPath($ExpectedPath).TrimEnd('\')
    $root=[System.IO.Path]::GetPathRoot($full)
    if([string]::IsNullOrWhiteSpace($root)){$answer.Reason='Не може да се определи drive root на очаквания път.';return $answer}
    $drive=$root.TrimEnd('\').ToLowerInvariant()
    $rel=$full.Substring($root.Length).Trim('\')
    $expectedParts=@();if($rel){$expectedParts=@($rel -split '\\')}
    $tokens=@($Snapshot.BreadcrumbTokens)
    $driveIndex=-1
    for($i=0;$i -lt $tokens.Count;$i++){
        $t=([string]$tokens[$i]).Trim()
        if($t -eq $drive -or $t -match ('(?i)(?:^|[\s\(])'+[regex]::Escape($drive)+'(?:[\s\)]|$)')){$driveIndex=$i;break}
    }
    if($driveIndex -lt 0){$answer.Reason='Breadcrumb не съдържа очаквания mapped/local drive '+$drive+'.';return $answer}
    $after=New-Object System.Collections.Generic.List[string]
    for($i=$driveIndex+1;$i -lt $tokens.Count;$i++){
        $part=([string]$tokens[$i]).Trim()
        if(-not $part){continue}
        if($part -eq '.' -or $part -eq '..' -or $part.IndexOfAny([char[]]'\/:') -ge 0){$answer.Reason='Невалиден компонент в breadcrumb.';return $answer}
        $after.Add($part)
    }
    if($after.Count -ne $expectedParts.Count){$answer.Reason='Breadcrumb не съвпада по дълбочина с очаквания път.';return $answer}
    for($i=0;$i -lt $expectedParts.Count;$i++){if([string]$after[$i] -ine [string]$expectedParts[$i]){$answer.Reason='Breadcrumb не съвпада с очаквания път при '+[string]$expectedParts[$i]+'.';return $answer}}
    $answer.ObservedPath=$expected
    if($ExpectedName -and [string]$Snapshot.NameValue -cne $ExpectedName){$answer.Reason='Името в текущия Save as не е потвърдено: '+[string]$Snapshot.NameValue;return $answer}
    $answer.Verified=$true;$answer.Reason='Потвърдена текуща папка по drive breadcrumb: '+$expected
    return $answer
}

function Get-MixSaveFolderProof {
    param($Snapshot,[string]$ExpectedPath,[string]$DesktopPath,[string]$ExpectedName='')
    $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    if(Test-PathUnderFolder -Path $ExpectedPath -Parent $desktop){
        return (Get-MixSaveFolderProofDesktopLegacy -Snapshot $Snapshot -ExpectedPath $ExpectedPath -DesktopPath $desktop -ExpectedName $ExpectedName)
    }
    return (Get-MixSaveFolderProofGeneric -Snapshot $Snapshot -ExpectedPath $ExpectedPath -ExpectedName $ExpectedName)
}

function Test-MixElementWithin {
    param($Element,$Ancestor)
    if ($null -eq $Element -or $null -eq $Ancestor) { return $false }
    try {
        $walker=[System.Windows.Automation.TreeWalker]::RawViewWalker
        for ($level=0; $null -ne $Element -and $level -lt 32; $level++) {
            if ([System.Windows.Automation.Automation]::Compare($Element,$Ancestor)) { return $true }
            $Element=$walker.GetParent($Element)
        }
    } catch {}
    return $false
}

function Get-MixSaveFolderSnapshot {
    param([int]$ProcessId)
    $s=[pscustomobject]@{ScopeId='';IsActive=$false;NameScopeId='';NameVisible=$false;NameValue='';ConfirmScopeId='';ConfirmVisible=$false;CurrentFolderScopeId='';BreadcrumbTokens=@();DesktopAnchorPath='';DesktopAnchorIsRoot=$false;CurrentAbsolutePath='';Scope=$null;NameEdit=$null;ConfirmElement=$null;ContentList=$null}
    try {
        $handle=[PWXMixNative]::GetForegroundWindow()
        $foreground=[System.Windows.Automation.AutomationElement]::FromHandle($handle)
        if ($foreground.Current.ProcessId -ne $ProcessId) { return $s }
        $edit=Get-MixSaveNameEdit -ProcessId $ProcessId -RequireVisible
        if (-not (Test-MixElementWithin -Element $edit -Ancestor $foreground)) { return $s }
        $eb=$edit.Current.BoundingRectangle
        $walker=[System.Windows.Automation.TreeWalker]::RawViewWalker
        $scope=$walker.GetParent($edit)
        $visible=@();$title=$null;$confirm=$null;$tree=$null
        # Training shows a custom modal browser with a Save-as title, left Tree,
        # central List and wide bottom Name field. Use their common container.
        for ($level=0; $null -ne $scope -and $level -lt 18; $level++) {
            $visible=@($scope.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition) | Where-Object {
                try { $b=$_.Current.BoundingRectangle; -not $_.Current.IsOffscreen -and $_.Current.IsEnabled -and $b.Width -gt 1 -and $b.Height -gt 1 } catch { $false }
            })
            $title=@($visible | Where-Object { $_.Current.Name -match '^Save as(?:\.{3}|…)?$' -and $_.Current.BoundingRectangle.Top -lt $eb.Top } | Sort-Object { $_.Current.BoundingRectangle.Top } | Select-Object -Last 1)
            $confirm=@($visible | Where-Object { $_.Current.Name -eq 'Confirm' -and [Math]::Abs($_.Current.BoundingRectangle.Top-$eb.Top) -lt 45 } | Sort-Object { $_.Current.BoundingRectangle.Width } -Descending | Select-Object -First 1)
            $tree=@($visible | Where-Object { $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Tree -and $_.Current.BoundingRectangle.Top -lt $eb.Top -and $_.Current.BoundingRectangle.Left -lt ($eb.Left+100) } | Sort-Object { $_.Current.BoundingRectangle.Top } -Descending | Select-Object -First 1)
            if ($title.Count -eq 1 -and $confirm.Count -eq 1 -and $tree.Count -eq 1) { $title=$title[0];$confirm=$confirm[0];$tree=$tree[0];break }
            $scope=$walker.GetParent($scope)
        }
        if ($null -eq $scope -or $title -is [array] -or $confirm -is [array] -or $tree -is [array]) { return $s }
        $tb=$tree.Current.BoundingRectangle;$hb=$title.Current.BoundingRectangle;$cb=$confirm.Current.BoundingRectangle
        # Hit tests reject a Save-as surface hidden behind another modal window.
        $nameHit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($eb.Left+$eb.Width/2),($eb.Top+$eb.Height/2))))
        $confirmHit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($cb.Left+$cb.Width/2),($cb.Top+$cb.Height/2))))
        $treeHit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($tb.Left+$tb.Width/2),($tb.Top+$tb.Height/2))))
        if (-not (Test-MixElementWithin $nameHit $edit) -or -not (Test-MixElementWithin $confirmHit $confirm) -or -not (Test-MixElementWithin $treeHit $tree)) { return $s }
        $s.ScopeId=($scope.GetRuntimeId() -join ':')+'|'+($title.GetRuntimeId() -join ':')
        $s.IsActive=$true;$s.Scope=$scope;$s.NameEdit=$edit;$s.ConfirmElement=$confirm
        $s.NameScopeId=$s.ScopeId;$s.ConfirmScopeId=$s.ScopeId;$s.NameVisible=$true;$s.ConfirmVisible=$true
        $vp=$null;if ($edit.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$vp)) { $s.NameValue=[string]$vp.Current.Value }
        $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        $desktopLeaf=Split-Path -Leaf $desktop
        $roots=@($tree.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition) | Where-Object { $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::TreeItem -and ($_.Current.Name -eq 'Desktop' -or $_.Current.Name -eq $desktopLeaf) })
        $desktopRootIds=New-Object System.Collections.Generic.List[string]
        foreach ($node in $roots) {
            $parent=$walker.GetParent($node);$rooted=$false
            for ($depth=0; $null -ne $parent -and $depth -lt 20; $depth++) {
                if ([System.Windows.Automation.Automation]::Compare($parent,$tree)) { $rooted=$true;break }
                if ($parent.Current.ControlType -eq [System.Windows.Automation.ControlType]::TreeItem) { break }
                $parent=$walker.GetParent($parent)
            }
            if ($rooted) { $id=$node.GetRuntimeId() -join ':';if (-not $desktopRootIds.Contains($id)) { $desktopRootIds.Add($id) } }
        }
        if ($desktopRootIds.Count -eq 1) { $s.DesktopAnchorIsRoot=$true;$s.DesktopAnchorPath=$desktop }
        $lists=@($visible | Where-Object { $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::List -and $_.Current.BoundingRectangle.Left -ge ($tb.Right-10) -and $_.Current.BoundingRectangle.Top -ge ($tb.Top-20) -and $_.Current.BoundingRectangle.Bottom -le ($eb.Top+5) } | Sort-Object { $_.Current.BoundingRectangle.Left } | Select-Object -First 1)
        if ($lists.Count -eq 1) { $s.ContentList=$lists[0] }
        # Only the toolbar row BETWEEN Save-as title and folder tree is a current
        # breadcrumb. Tree labels, selected center rows and right preview cannot enter.
        $crumbs=New-Object System.Collections.Generic.List[object]
        $right=$cb.Right
        foreach ($toolbarEdit in @($visible | Where-Object { $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Edit })) {
            $b=$toolbarEdit.Current.BoundingRectangle
            $addressPattern=$null
            if ($toolbarEdit.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$addressPattern) -and (ConvertTo-MixWindowsFolderPath ([string]$addressPattern.Current.Value))) { continue }
            if ($b.Top -gt $hb.Bottom -and $b.Bottom -lt $tb.Top -and $b.Left -gt $tb.Right -and $b.Width -gt 80) { $right=[Math]::Min($right,$b.Left-5) }
        }
        foreach ($el in $visible) {
            $b=$el.Current.BoundingRectangle;$ct=$el.Current.ControlType
            if ($b.Top -le $hb.Bottom -or $b.Bottom -ge $tb.Top -or $b.Left -lt ($tb.Left+80) -or $b.Right -gt $right) { continue }
            $hit=[System.Windows.Automation.AutomationElement]::FromPoint((New-Object System.Windows.Point(($b.Left+$b.Width/2),($b.Top+$b.Height/2))))
            if (-not (Test-MixElementWithin $hit $el)) { continue }
            if ($ct -eq [System.Windows.Automation.ControlType]::Edit) {
                $v=$null
                if ($el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$v)) {
                    $p=ConvertTo-MixWindowsFolderPath ([string]$v.Current.Value)
                    if ($p) { if ($s.CurrentAbsolutePath -and $s.CurrentAbsolutePath -ne $p) { $s.IsActive=$false;return $s };$s.CurrentAbsolutePath=$p }
                }
                continue
            }
            if ($ct -ne [System.Windows.Automation.ControlType]::Text -and $ct -ne [System.Windows.Automation.ControlType]::Hyperlink -and $ct -ne [System.Windows.Automation.ControlType]::Button) { continue }
            $name=([string]$el.Current.Name).Trim()
            if (-not $name -or $name -match '^(?:Search|>|›|»)$') { continue }
            $duplicate=$false
            foreach ($old in $crumbs) { if ($old.Name -eq $name -and $b.Left -lt $old.Right -and $b.Right -gt $old.Left) { $duplicate=$true;break } }
            if (-not $duplicate) { $crumbs.Add([pscustomobject]@{Name=$name;Left=$b.Left;Right=$b.Right}) }
        }
        $parts=New-Object System.Collections.Generic.List[string]
        foreach ($crumb in @($crumbs.ToArray() | Sort-Object Left)) { foreach ($part in @($crumb.Name -split '\s*[>›»]\s*')) { if ($part.Trim()) { $parts.Add($part.Trim()) } } }
        $s.BreadcrumbTokens=@($parts.ToArray())
        if ($s.BreadcrumbTokens.Count -gt 0 -or $s.CurrentAbsolutePath) { $s.CurrentFolderScopeId=$s.ScopeId }
    } catch { Write-MixLog ('SAVE folder snapshot: '+$_.Exception.Message) }
    return $s
}

function Open-MixSaveDailyFolder {
    param([int]$ProcessId,[string]$DailyPath)
    $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    $parent=Split-Path -Parent $DailyPath;$leaf=Split-Path -Leaf $DailyPath
    for ($attempt=0; $attempt -lt 3; $attempt++) {
        Test-OperatorStopRequested
        $snapshot=Get-MixSaveFolderSnapshot -ProcessId $ProcessId
        $proof=Get-MixSaveFolderProof -Snapshot $snapshot -ExpectedPath $DailyPath -DesktopPath $desktop
        Write-MixLog ('SAVE current folder: '+$proof.Reason+'; breadcrumb='+(@($snapshot.BreadcrumbTokens) -join ' > '))
        if ($proof.Verified) { return $true }
        $parentProof=Get-MixSaveFolderProof -Snapshot $snapshot -ExpectedPath $parent -DesktopPath $desktop
        if (-not $parentProof.Verified -or $null -eq $snapshot.ContentList) { return $false }
        # Selection/preview of dd.MM leaves FOUND open. Activate the exact center
        # folder row with a double click, then require the breadcrumb to change.
        $list=$snapshot.ContentList;$target=$null;$scroll=$null
        try { if ($list.TryGetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern,[ref]$scroll) -and $scroll.Current.VerticallyScrollable) { $scroll.SetScrollPercent(-1,0) } } catch {}
        for ($page=0; $page -lt 24; $page++) {
            Test-OperatorStopRequested
            foreach ($item in $list.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)) {
                if ($item.Current.ControlType -ne [System.Windows.Automation.ControlType]::ListItem -or $item.Current.IsOffscreen -or -not $item.Current.IsEnabled) { continue }
                $name=[string]$item.Current.Name
                if ($name -eq $leaf -or (ConvertTo-MixWindowsFolderPath $name) -eq (ConvertTo-MixWindowsFolderPath $DailyPath)) { $target=$item;break }
            }
            if ($null -ne $target -or $null -eq $scroll -or -not $scroll.Current.VerticallyScrollable) { break }
            $before=$scroll.Current.VerticalScrollPercent
            $scroll.Scroll([System.Windows.Automation.ScrollAmount]::NoAmount,[System.Windows.Automation.ScrollAmount]::LargeIncrement)
            Start-Sleep -Milliseconds 150
            if ($scroll.Current.VerticalScrollPercent -eq $before) { break }
        }
        if ($null -eq $target) { Write-MixLog ('SAVE cannot find date folder in the opened FOUND list: '+$leaf);return $false }
        try {
            Test-OperatorStopRequested
            if (-not (Activate-MixMaestroWindow -ProcessId $ProcessId)) { return $false }
            $current=Get-MixSaveFolderSnapshot -ProcessId $ProcessId
            $currentParent=Get-MixSaveFolderProof -Snapshot $current -ExpectedPath $parent -DesktopPath $desktop
            if (-not $currentParent.Verified -or -not (Test-MixElementWithin $target $current.ContentList)) { return $false }
            $target.SetFocus();$b=$target.Current.BoundingRectangle
            $x=[int]($b.Left+[Math]::Min(55,$b.Width/2));$y=[int]($b.Top+$b.Height/2)
            Invoke-MixNativeClick -X $x -Y $y
            Start-Sleep -Milliseconds 80
            Invoke-MixNativeClick -X $x -Y $y
            Write-MixLog ('SAVE open date folder by double click: '+$DailyPath)
            # V5.1.34 speed-only: wait up to the same ~600 ms, but continue as
            # soon as the breadcrumb proves that the exact folder is open.
            $openUntil=(Get-Date).AddMilliseconds(650)
            do {
                Start-Sleep -Milliseconds 80
                $probe=Get-MixSaveFolderSnapshot -ProcessId $ProcessId
                $probeProof=Get-MixSaveFolderProof -Snapshot $probe -ExpectedPath $DailyPath -DesktopPath $desktop
                if($probeProof.Verified){ return $true }
            } while((Get-Date) -lt $openUntil)
        } catch { Write-MixLog ('SAVE date activation failed: '+$_.Exception.Message);return $false }
    }
    $snapshot=Get-MixSaveFolderSnapshot -ProcessId $ProcessId
    $proof=Get-MixSaveFolderProof -Snapshot $snapshot -ExpectedPath $DailyPath -DesktopPath $desktop
    Write-MixLog ('SAVE final current folder: '+$proof.Reason)
    return [bool]$proof.Verified
}

function Invoke-MixStrictSaveConfirm {
    param([int]$ProcessId,[string]$ExpectedName,[string]$DailyPath)
    # Capture afresh after SetName/TAB and immediately before the only Save click.
    $snapshot=Get-MixSaveFolderSnapshot -ProcessId $ProcessId
    $desktop=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    $proof=Get-MixSaveFolderProof -Snapshot $snapshot -ExpectedPath $DailyPath -DesktopPath $desktop -ExpectedName $ExpectedName
    Write-MixLog ('SAVE pre-Confirm proof: '+$proof.Reason+'; name='+[string]$snapshot.NameValue)
    if (-not $proof.Verified) { return $false }
    try {
        $b=$snapshot.ConfirmElement.Current.BoundingRectangle
        Invoke-MixNativeClick -X ([int]($b.Left+$b.Width/2)) -Y ([int]($b.Top+$b.Height/2))
        return $true
    } catch { return $false }
}


function Get-MixSaveNameEdit {
    param([int]$ProcessId,[switch]$RequireVisible)
    try {
        $root=[System.Windows.Automation.AutomationElement]::RootElement
        $pc=New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
        $all=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$pc)
        $edits=New-Object System.Collections.Generic.List[object]
        foreach($el in $all){
            try{
                if($el.Current.ControlType -eq [System.Windows.Automation.ControlType]::Edit){
                    if($RequireVisible -and ($el.Current.IsOffscreen -or -not $el.Current.IsEnabled)){continue}
                    $b=$el.Current.BoundingRectangle
                    # The Save-as Name field is the wide edit at the bottom of
                    # the dialog. The search box is much narrower and higher.
                    if($b.Width -gt 500 -and $b.Height -gt 15){
                        $edits.Add([pscustomobject]@{El=$el;Y=$b.Top;W=$b.Width;H=$b.Height})
                    }
                }
            }catch{}
        }
        $target=@($edits.ToArray() | Sort-Object @{Expression='Y';Descending=$true},@{Expression='W';Descending=$true} | Select-Object -First 1)
        if($target.Count -gt 0){ return $target[0].El }
    } catch {}
    return $null
}

function Test-MixSaveNameValue {
    param($Element,[string]$Value)
    if($null -eq $Element){ return $false }
    try {
        $pat=$null
        if($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$pat)){
            return ([string]$pat.Current.Value -eq $Value)
        }
    } catch {}
    return $false
}

function Set-MixSaveName {
    param([int]$ProcessId,[string]$Value,[switch]$Commit)
    $el=Get-MixSaveNameEdit -ProcessId $ProcessId -RequireVisible:$Commit
    if($null -eq $el){ return $false }

    # Preferred method: set the actual ValuePattern and then verify it.
    try {
        $pat=$null
        if($el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$pat)){
            $pat.SetValue($Value)
            Start-Sleep -Milliseconds 250
            if(Test-MixSaveNameValue -Element $el -Value $Value){
                if($Commit){
                    # WPF bindings can update their source only after LostFocus.
                    # InvokePattern on Confirm does not itself move keyboard focus.
                    $el.SetFocus()
                    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
                    Start-Sleep -Milliseconds 250
                    if(-not (Test-MixSaveNameValue -Element $el -Value $Value)){return $false}
                }
                Write-MixLog ('Save-as name set via ValuePattern: ' + $Value)
                return $true
            }
        }
    } catch {}

    # Keyboard/clipboard fallback for Maestro builds where the edit control
    # exposes ValuePattern but silently rejects SetValue.
    try {
        $el.SetFocus()
        $b=$el.Current.BoundingRectangle
        if($b.Width -gt 1 -and $b.Height -gt 1){ Invoke-MixNativeClick -X ([int]($b.Left+$b.Width/2)) -Y ([int]($b.Top+$b.Height/2)) }
        Start-Sleep -Milliseconds 150
        [System.Windows.Forms.SendKeys]::SendWait('^a')
        [System.Windows.Forms.Clipboard]::SetText($Value)
        [System.Windows.Forms.SendKeys]::SendWait('^v')
        Start-Sleep -Milliseconds 300
        $el2=Get-MixSaveNameEdit -ProcessId $ProcessId -RequireVisible:$Commit
        if(Test-MixSaveNameValue -Element $el2 -Value $Value){
            if($Commit){
                [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
                Start-Sleep -Milliseconds 250
                if(-not (Test-MixSaveNameValue -Element $el2 -Value $Value)){return $false}
            }
            Write-MixLog ('Save-as name set via clipboard: ' + $Value)
            return $true
        }
    } catch {}
    return $false
}

function Get-MixSaveFileStamp {
    param([string]$Path)
    try {
        $f=Get-Item -LiteralPath $Path -ErrorAction Stop
        if($f.PSIsContainer){return $null}
        return [pscustomobject]@{Path=$f.FullName;Length=[long]$f.Length;WriteTicks=[long]$f.LastWriteTimeUtc.Ticks;WriteUtc=$f.LastWriteTimeUtc}
    } catch {return $null}
}

function Write-MixSaveDiagnostic {
    param([int]$ProcessId,[string]$DailyPath,[string]$ExpectedName,[string]$Reason)
    Write-MixLog ('SAVE DIAGNOSTIC: '+$Reason+'; expected folder='+$DailyPath+'; expected name='+$ExpectedName)
    try {
        # List only this run's exact daily directory, never other products/dates.
        $names=@(Get-ChildItem -LiteralPath $DailyPath -File -ErrorAction Stop | Where-Object {$_.Name -like '*.mixx*' -or $_.Name -eq [System.IO.Path]::GetFileNameWithoutExtension($ExpectedName)} | Sort-Object Name)
        foreach($f in $names){Write-MixLog ('SAVE DAILY CANDIDATE: '+$f.Name+'; bytes='+$f.Length+'; UTC='+$f.LastWriteTimeUtc.ToString('o'))}
        if($names.Count -eq 0){Write-MixLog 'SAVE DAILY CANDIDATES: none'}
    } catch {Write-MixLog ('SAVE DIRECTORY INSPECTION ERROR: '+$_.Exception.Message)}
    try {
        $root=[System.Windows.Automation.AutomationElement]::RootElement
        $pc=New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
        $all=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$pc)
        $count=0
        foreach($el in $all){
            if($count -ge 80){break}
            try {
                if($el.Current.IsOffscreen){continue}
                $b=$el.Current.BoundingRectangle
                if($b.Width -le 1 -or $b.Height -le 1){continue}
                $ct=[string]$el.Current.ControlType.ProgrammaticName
                $name=[string]$el.Current.Name
                $value='';$selected=$false;$vp=$null;$sp=$null
                if($ct -eq 'ControlType.Edit' -and $el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$vp)){$value=[string]$vp.Current.Value}
                if($el.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$sp)){$selected=$sp.Current.IsSelected}
                if($ct -ne 'ControlType.Edit' -and -not $selected -and $name -notmatch '(?i)save|confirm|error|replace|already|folder|path|FOUND|Mix\(|[A-Z]:\\'){continue}
                if($name.Length -gt 220){$name=$name.Substring(0,220)}
                if($value.Length -gt 300){$value=$value.Substring(0,300)}
                Write-MixLog ('SAVE UI: type='+$ct+'; name='+$name+'; value='+$value+'; selected='+$selected+'; enabled='+$el.Current.IsEnabled+'; id='+$el.Current.AutomationId)
                if($selected -and $ct -eq 'ControlType.TreeItem'){
                    $parts=New-Object System.Collections.Generic.List[string]
                    $node=$el;$walker=[System.Windows.Automation.TreeWalker]::ControlViewWalker
                    for($level=0;$level -lt 16 -and $null -ne $node;$level++){
                        if($node.Current.ControlType -eq [System.Windows.Automation.ControlType]::TreeItem){$parts.Insert(0,[string]$node.Current.Name)}
                        $node=$walker.GetParent($node)
                    }
                    Write-MixLog ('SAVE SELECTED FOLDER BRANCH: '+($parts.ToArray() -join '\'))
                }
                $count++
            } catch {}
        }
    } catch {Write-MixLog ('SAVE UI INSPECTION ERROR: '+$_.Exception.Message)}
}

function Invoke-MixSaveConfirm {
    param([int]$ProcessId,[string]$ExpectedName)
    $edit=Get-MixSaveNameEdit -ProcessId $ProcessId -RequireVisible
    if(-not (Test-MixSaveNameValue -Element $edit -Value $ExpectedName)){return $false}
    $buttons=New-Object System.Collections.Generic.List[object]
    $eb=$edit.Current.BoundingRectangle
    foreach($el in @(Get-MixUiElementsByName -ProcessId $ProcessId -Name 'Confirm')){
        try {
            if($el.Current.IsOffscreen -or -not $el.Current.IsEnabled){continue}
            $b=$el.Current.BoundingRectangle
            if($b.Width -le 1 -or $b.Height -le 1){continue}
            # The Save-as Confirm is beside/below the filename field. Ignore
            # an unrelated Confirm that remains exposed behind the dialog.
            if($b.Bottom -lt ($eb.Top-40) -or $b.Top -gt ($eb.Bottom+180)){continue}
            $buttons.Add([pscustomobject]@{Element=$el;Distance=[Math]::Abs(($b.Top+$b.Height/2)-($eb.Top+$eb.Height/2))})
        } catch {}
    }
    foreach($candidate in @($buttons.ToArray() | Sort-Object Distance)){
        try {
            # A physical click commits the focused WPF edit before the command.
            $b=$candidate.Element.Current.BoundingRectangle
            Invoke-MixNativeClick -X ([int]($b.Left+$b.Width/2)) -Y ([int]($b.Top+$b.Height/2))
            Write-MixLog ('SAVE Confirm clicked; filename verified='+$ExpectedName)
            return $true
        } catch {}
    }
    # Keep the trained coordinate fallback, but only while the exact filename
    # is still observed in the live, enabled Save-as field.
    $edit=Get-MixSaveNameEdit -ProcessId $ProcessId -RequireVisible
    if(Test-MixSaveNameValue -Element $edit -Value $ExpectedName){
        Invoke-MixRelativeClick -ProcessId $ProcessId -RefX 1680 -RefY 890
        Write-MixLog ('SAVE Confirm trained-coordinate fallback; filename verified='+$ExpectedName)
        return $true
    }
    return $false
}

function Wait-MixSaveVerified {
    param([int]$ProcessId,[string]$TargetPath,[datetime]$StartedUtc,[hashtable]$Before,[int]$TimeoutSec=120)
    $deadline=[datetime]::UtcNow.AddSeconds($TimeoutSec)
    $stable=@{};$last=@{};$why='Файлът още не е създаден.'
    # Some custom save dialogs append their filter extension unconditionally.
    # Only this exact double-extension sibling can be normalized, after all
    # freshness/readability/stability checks and only if the target is absent.
    $paths=@($TargetPath,($TargetPath+'.mixx'))
    do {
        [System.Windows.Forms.Application]::DoEvents();Test-OperatorStopRequested
        foreach($path in $paths){
            $key=$path.ToLowerInvariant();$stamp=Get-MixSaveFileStamp -Path $path
            if($null -eq $stamp){$stable[$key]=0;continue}
            if($stamp.Length -le 0){$why='Създаденият MIXX още е празен.';$stable[$key]=0;continue}
            $old=$Before[$key]
            if($stamp.WriteUtc -lt $StartedUtc.AddSeconds(-2) -or ($null -ne $old -and $stamp.WriteTicks -eq $old.WriteTicks -and $stamp.Length -eq $old.Length)){
                $why='Намерено е старо/непроменено копие; очаквам новия запис.';$stable[$key]=0;continue
            }
            $signature=([string]$stamp.WriteTicks)+'|'+([string]$stamp.Length)
            if($last[$key] -eq $signature){$stable[$key]=[int]$stable[$key]+1}else{$last[$key]=$signature;$stable[$key]=1}
            if([int]$stable[$key] -lt 3){$why='MIXX е намерен; изчаквам записването да приключи.';continue}
            $stream=$null
            try {
                # FileShare.Read refuses an outstanding writer: a closed dialog
                # or a file's mere existence alone is not evidence of completion.
                $stream=[System.IO.File]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::Read)
                if($stream.Length -le 0){continue}
                $checked=Get-MixSaveFileStamp -Path $path
                if($null -eq $checked -or $stream.Length -ne $stamp.Length -or $checked.WriteTicks -ne $stamp.WriteTicks){
                    $why='MIXX се промени между проверките; изчаквам стабилен запис.';$stable[$key]=0;continue
                }
            } catch {$why='MIXX още е заключен за запис: '+$_.Exception.Message;continue}
            finally {if($null -ne $stream){$stream.Dispose()}}
            if($path -ne $TargetPath){
                if([System.IO.File]::Exists($TargetPath)){$why='Има отделен целеви MIXX; двойното разширение не се преименува.';continue}
                try {[System.IO.File]::Move($path,$TargetPath);Write-MixLog ('SAVE normalized fresh double extension: '+$path+' -> '+$TargetPath)}
                catch {$why='Неуспешно нормализиране на двойно разширение: '+$_.Exception.Message;continue}
            }
            Write-MixLog ('SAVE VERIFIED: '+$TargetPath+'; bytes='+$stamp.Length+'; UTC='+$stamp.WriteUtc.ToString('o'))
            return $TargetPath
        }
        Start-Sleep -Milliseconds 500
    } while([datetime]::UtcNow -lt $deadline)
    $daily=Split-Path -Parent $TargetPath
    Write-MixSaveDiagnostic -ProcessId $ProcessId -DailyPath $daily -ExpectedName ([System.IO.Path]::GetFileName($TargetPath)) -Reason $why
    throw ('Не е потвърден завършен запис на дневния MIXX: '+$TargetPath+'. '+$why+' Проверете Save as/съобщението в Maestro. Подробностите са в Mix log; PGMX остават готови.')
}

function Invoke-DailyMixAutomation {
    param([Parameter(Mandatory=$true)][string[]]$PgmxFiles)
    if ($PgmxFiles.Count -eq 0) { throw 'Няма PGMX файлове за Mix.' }
    $daily = Get-DailyMixFolder
    $dailyName = [System.IO.Path]::GetFileName($daily)
    $mixBaseName = 'Mix(' + $dailyName + ')'
    $targetMix = Join-Path $daily ($mixBaseName + '.mixx')
    $script:LastMixSaveAttemptPath = $targetMix
    $script:LastMixSaveConfirmed = $false

    $p = Start-OrActivateMaestro
    $maestroPid = $p.Id
    if ($script:LastMaestroLauncher) { Write-MixLog ('Maestro source: ' + $script:LastMaestroLauncher) }
    Write-MixLog ('AUTO START. PID=' + $maestroPid + '; inputs=' + $PgmxFiles.Count + '; folder=' + $daily)
    if ($script:LastMaestroLauncher) { $autoMixStatusLabel.Text = 'Maestro: ' + $script:LastMaestroLauncher }

    $status.Text='Стъпка 4 AUTO: отварям Mixx...'; [System.Windows.Forms.Application]::DoEvents(); Test-OperatorStopRequested
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)

    # First prefer the real UI text learned from the training sessions. This is
    # independent of monitor/DPI scaling. If Maestro does not expose the text
    # through UI Automation, fall back to the learned click position.
    $mixxOpened=$false
    if (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Mixx' -TimeoutSec 2) {
        Start-Sleep -Milliseconds 900
        if (Test-MixUiElementByName -ProcessId $maestroPid -Name 'Mix' -TimeoutSec 4) { $mixxOpened=$true }
    }
    if (-not $mixxOpened) {
        [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 403 -RefY 170
        Start-Sleep -Milliseconds 900
        if (Test-MixUiElementByName -ProcessId $maestroPid -Name 'Mix' -TimeoutSec 5) { $mixxOpened=$true }
    }
    if (-not $mixxOpened) { throw 'Не успях да отворя меню Mixx в Maestro.' }
    Write-MixLog 'Mixx menu opened.'

    $status.Text='Стъпка 4 AUTO: създавам нов Mix...'; [System.Windows.Forms.Application]::DoEvents(); Test-OperatorStopRequested
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    if (-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Mix' -TimeoutSec 3)) {
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 1183 -RefY 348
    }
    Start-Sleep -Seconds 3

    $status.Text='Стъпка 4 AUTO: Add programs...'; [System.Windows.Forms.Application]::DoEvents(); Test-OperatorStopRequested
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    if (-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Add programs' -TimeoutSec 2)) {
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 74 -RefY 181
    }
    Start-Sleep -Seconds 2

    $first=[System.IO.Path]::GetFileName($PgmxFiles[0])
    if (-not (Select-MixDailyFolderInDialog -ProcessId $maestroPid -DailyName $dailyName -ExpectedFileName $first -DailyPath $daily)) { throw "Не успях да избера и потвърдя дневната папка $dailyName в Add programs." }
    $status.Text=('Стъпка 4 AUTO: избирам всички PGMX (' + $PgmxFiles.Count + ')...'); [System.Windows.Forms.Application]::DoEvents(); Test-OperatorStopRequested
    if (-not (Select-AllMixPgmxFilesInDialog -ProcessId $maestroPid -ExpectedFiles $PgmxFiles -TimeoutSec 8)) {
        throw "Не успях да избера всички PGMX в Add programs. Очаквани: $($PgmxFiles.Count)."
    }
    Start-Sleep -Milliseconds 350
    if (-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Confirm' -TimeoutSec 5)) {
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 1680 -RefY 890
    }

    $status.Text='Стъпка 4 AUTO: зареждам програмите...'; [System.Windows.Forms.Application]::DoEvents(); Test-OperatorStopRequested
    Start-Sleep -Seconds 2
    # The known Maestro read-warning is ignored for the process, but the exact
    # PGMX paths named by that warning are isolated AFTER OK. Moving is local
    # only (daily FOUND copy); the original PGMX in the product PWX folder is untouched.
    if (Test-MixUiElementByName -ProcessId $maestroPid -Name 'OK' -TimeoutSec 10) {
        $warningPgmx=@(Get-MixWarningPgmxPaths -ProcessId $maestroPid -DailyPath $daily)
        if(Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'OK' -TimeoutSec 2){
            Write-MixLog ('Maestro warning -> OK; listed PGMX='+$warningPgmx.Count)
            Start-Sleep -Milliseconds 350
            if($warningPgmx.Count -gt 0){
                $script:LastProblemPgmxMoved=@(Move-DailyPgmxToSubfolder -Paths $warningPgmx -Subfolder 'Проблемни_PGMX' -Reason 'Maestro Warning')
            }
        }
    }

    $status.Text='Стъпка 4 AUTO: Save as...'; [System.Windows.Forms.Application]::DoEvents(); Test-OperatorStopRequested
    if (-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'File' -TimeoutSec 4)) {
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 205 -RefY 24
    }
    Start-Sleep -Milliseconds 500
    if (-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Save as...' -TimeoutSec 4)) { throw 'Не намирам File -> Save as... в Maestro.' }
    Start-Sleep -Seconds 2

    # Keep the learned navigation, then distinguish a selected date/preview from
    # an OPENED date folder. Training shows selection alone leaves FOUND open.
    if (-not (Select-MixDailyFolderInDialog -ProcessId $maestroPid -DailyName $dailyName -ExpectedFileName '' -DailyPath $daily)) { throw "Не успях да избера дневната папка $dailyName в Save as." }
    if (-not (Open-MixSaveDailyFolder -ProcessId $maestroPid -DailyPath $daily)) {
        Write-MixSaveDiagnostic -ProcessId $maestroPid -DailyPath $daily -ExpectedName ([System.IO.Path]::GetFileName($targetMix)) -Reason 'Save as не е отворил дневната папка; записването е спряно преди Confirm.'
        throw ('Не е потвърдена отворена папка '+$daily+' в Save as. Този опит е спрян преди Confirm; провери текущия път в Maestro.')
    }
    $status.Text=('Стъпка 4 AUTO: записвам ' + $mixBaseName + '.mixx...'); [System.Windows.Forms.Application]::DoEvents(); Test-OperatorStopRequested
    # Restore the trained name: this dialog displays its fixed .mixx suffix separately.
    $saveName=$mixBaseName
    if (-not (Set-MixSaveName -ProcessId $maestroPid -Value $saveName -Commit)) {
        Write-MixSaveDiagnostic -ProcessId $maestroPid -DailyPath $daily -ExpectedName ([System.IO.Path]::GetFileName($targetMix)) -Reason 'Неуспешно задаване/потвърждение на името.'
        throw ('Не успях да задам и потвърдя името в Save as: '+$saveName)
    }
    Start-Sleep -Milliseconds 350
    $saveSnapshot=Get-MixSaveFolderSnapshot -ProcessId $maestroPid
    $saveProof=Get-MixSaveFolderProof -Snapshot $saveSnapshot -ExpectedPath $daily -DesktopPath ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) -ExpectedName $saveName
    Write-MixLog ('SAVE after Name commit: '+$saveProof.Reason+'; name='+[string]$saveSnapshot.NameValue)
    if (-not $saveProof.Verified) {
        Write-MixSaveDiagnostic -ProcessId $maestroPid -DailyPath $daily -ExpectedName ([System.IO.Path]::GetFileName($targetMix)) -Reason $saveProof.Reason
        throw ('Записването е спряно преди Confirm: '+$saveProof.Reason)
    }
    $beforeSave=@{}
    foreach($path in @($targetMix,($targetMix+'.mixx'))){$beforeSave[$path.ToLowerInvariant()]=Get-MixSaveFileStamp -Path $path}
    $saveStartedUtc=[datetime]::UtcNow
    if(-not (Invoke-MixStrictSaveConfirm -ProcessId $maestroPid -ExpectedName $saveName -DailyPath $daily)){
        Write-MixSaveDiagnostic -ProcessId $maestroPid -DailyPath $daily -ExpectedName ([System.IO.Path]::GetFileName($targetMix)) -Reason 'Не е потвърден бутонът Confirm в Save as.'
        throw 'Не успях да потвърдя Save as. Дневният MIXX не е отбелязан като готов.'
    }

    $targetMix=Wait-MixSaveVerified -ProcessId $maestroPid -TargetPath $targetMix -StartedUtc $saveStartedUtc -Before $beforeSave -TimeoutSec 120
    $script:LastMixSaveConfirmed = $true
    Write-MixLog ('AUTO OK -> ' + $targetMix)
    return $targetMix
}


function Invoke-AutoMixAfterStep3 {
    if ($script:AutoMixChainRunning) { return }
    Test-OperatorStopRequested
    if ($script:PgmxLastMode -ne 'Execute') { return }

    $newFiles = @(Get-NewPgmxForDailyMix)
    $measureFiles = @(Get-TimeMeasurementCandidates)
    if ($measureFiles.Count -eq 0) {
        try {
            $autoMixStatusLabel.Text = 'Всички текущи PGMX вече имат валидно време. Mixx/PowerSet не са нужни.'
            $status.Text = 'Стъпка 3 ГОТОВО: няма детайли без валидно време за PowerSet.'
            Update-AutoMixUi
            Write-MixLog 'AUTO CHAIN: all current successful PGMX already have valid times; Steps 4 and 5 skipped.'
        } catch {}
        try{Finalize-CsvRegistryForCurrentBatch -PipelineSucceeded $true -PipelineMessage 'PGMX са готови; няма нови PGMX за дневен Mixx.'}catch{}
        $script:OperatorAutoFlow = $false
        Complete-OperatorRunState -Message 'PGMX са готови; няма нови PGMX за дневен Mixx.'
        try { Refresh-OperatorReports -WriteFiles; Refresh-OperatorDashboard } catch {}
        try { Show-FinalProcessReport } catch {}
        return
    }

    $script:AutoMixChainRunning = $true
    $script:LastProblemPgmxMoved=@()
    $script:LastMedianFallbackCodes=@()
    $script:LastNoTimePgmxMoved=@()
    try{Set-OperatorCheckpoint -CurrentStage 'MIXX_RUNNING' -LastSafeStage 'PGMX_DONE' -Status 'ACTIVE' -Message 'Създаване на дневния Mixx.'}catch{}
    $step4Done = $false
    $mixPath = $null
    $xlsPath = $null
    try {
        $daily = Get-DailyMixFolder
        if (-not [System.IO.Directory]::Exists($daily)) { [void][System.IO.Directory]::CreateDirectory($daily) }
        $target = Get-DailyMixFilePath

        $tabs.SelectedTab = $tabTraining
        $autoMixRunButton.Enabled = $false
        $autoMixStatusLabel.Text = ('Стъпка 3 приключи. Нови PGMX: ' + $newFiles.Count + '. Копирам в дневната папка...')
        $status.Text = ('Стъпка 3 -> 4 AUTO: ' + $newFiles.Count + ' нови PGMX. Подготвям дневната папка...')
        [System.Windows.Forms.Application]::DoEvents()

        if ([System.IO.File]::Exists($target)) {
            [System.IO.File]::Delete($target)
            Write-MixLog ('AUTO CHAIN deleted existing Mixx automatically for rebuild: ' + $target)
            $autoMixStatusLabel.Text = 'Съществуващият дневен Mixx е премахнат. Създавам актуален Mixx наново...'
            [System.Windows.Forms.Application]::DoEvents()
        }

        $inputs = @(Prepare-DailyMixFiles)
        if ($inputs.Count -lt $measureFiles.Count) { throw "Несъответствие при дневното копиране. PGMX без валидно време: $($measureFiles.Count); налични за Mixx: $($inputs.Count)." }

        $autoMixStatusLabel.Text = ('PGMX без валидно време: ' + $measureFiles.Count + ' (нови '+@($measureFiles|Where-Object{$_.IsNew}).Count+', стари '+@($measureFiles|Where-Object{-not $_.IsNew}).Count+'); общо PGMX в дневната папка: ' + $inputs.Count + '. Стартирам Maestro...')
        $status.Text = ('Стъпка 3 -> 4 AUTO: дневната папка съдържа ' + $inputs.Count + ' PGMX. Създавам общ дневен Mixx...')
        [System.Windows.Forms.Application]::DoEvents()

        Test-OperatorStopRequested
        $mixPath = Invoke-DailyMixAutomation -PgmxFiles $inputs
        foreach($problem in @($script:LastProblemPgmxMoved)){
            Register-PgmxAutoExclusion -ProgramCodeF ([System.IO.Path]::GetFileName($problem)) -Reason 'PROBLEM_MIXX' -Origin $problem
        }
        $step4Done = $true
        try{Set-OperatorCheckpoint -CurrentStage 'MIXX_DONE' -LastSafeStage 'MIXX_DONE' -Status 'ACTIVE' -Message ('Mixx е създаден: '+$mixPath)}catch{}
        $autoMixStatusLabel.Text = 'Стъпка 4 ГОТОВА. Стартирам автоматично Стъпка 5 - PowerSet...'
        $status.Text = 'Стъпки 3 + 4 ГОТОВИ. Стартирам PowerSet...'
        Update-AutoMixUi
        Write-MixLog ('AUTO CHAIN STEP4 OK: Step3 new=' + $newFiles.Count + '; daily total=' + $inputs.Count + '; Mixx=' + $mixPath)

        try { $tabs.SelectedTab = $tabPowerSet } catch {}
        try { $powerSetAutoRunButton.Enabled = $false } catch {}
        try{Set-OperatorCheckpoint -CurrentStage 'POWERSET_RUNNING' -LastSafeStage 'MIXX_DONE' -Status 'ACTIVE' -Message 'PowerSet обработка.'}catch{}
        Test-OperatorStopRequested
        Set-PowerSetTestStatus 'AUTO Стъпка 5: стартирам PowerSet след успешния дневен Mixx...'
        [System.Windows.Forms.Application]::DoEvents()

        $xlsPath = Invoke-PowerSetFullAutomation
        foreach($missing in @($script:LastMedianFallbackCodes)){
            Register-PgmxAutoExclusion -ProgramCodeF $missing -Reason 'NO_REAL_TIME' -Origin (Get-DailyMixFolder)
        }
        try{Set-OperatorCheckpoint -CurrentStage 'POWERSET_DONE' -LastSafeStage 'POWERSET_DONE' -Status 'ACTIVE' -Message ('PowerSet/XLS е готов: '+$xlsPath)}catch{}
        $copyPath = Join-Path 'Z:\30_Biesse_edging\scm\detail_time' ([System.IO.Path]::GetFileName($xlsPath))

        $autoMixStatusLabel.Text = 'ГОТОВО АВТОМАТИЧНО: Стъпки 3 + 4 + 5 приключиха успешно.'
        $status.Text = 'СТЪПКИ 3 + 4 + 5 ГОТОВИ: ' + $xlsPath
        Update-AutoMixUi
        Write-MixLog ('AUTO CHAIN FULL OK: Step3 new=' + $newFiles.Count + '; daily total=' + $inputs.Count + '; Mixx=' + $mixPath + '; XLS=' + $xlsPath + '; Copy=' + $copyPath)

        try{Finalize-CsvRegistryForCurrentBatch -PipelineSucceeded $true -PipelineMessage 'Целият процес PGMX + Mixx + PowerSet е завършен.'}catch{}
        Complete-OperatorRunState -Message 'Целият дневен процес е завършен успешно.'
        try { Refresh-OperatorReports -WriteFiles; Refresh-OperatorDashboard } catch {}
        $opStates=@(Get-OperatorProductStateRows)
        $opSkipped=@($opStates | Where-Object { -not $_.Confirmed }).Count
        $opSuccess=@('ГОТОВ','СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF')
        $opDetailErrors=@($script:PgmxResultObjects | Where-Object { $opSuccess -notcontains (Safe-Trim $_.Result) }).Count
        $regNow=@(Load-CsvRegistryRows);$regDone=@($regNow|Where-Object{(Safe-Trim $_.Status)-eq 'ЗАВЪРШЕН'}).Count;$regRetry=@($regNow|Where-Object{(Safe-Trim $_.Status)-in @('ЧАСТИЧНО','ГРЕШКА')}).Count
        [System.Windows.Forms.MessageBox]::Show(
            "Целият дневен процес приключи успешно.`r`n`r`nНови PGMX от Стъпка 3: $($newFiles.Count)`r`nОбщо PGMX подадени към дневния Mixx: $($inputs.Count)`r`nПреместени в Проблемни_PGMX: $(@($script:LastProblemPgmxMoved).Count)`r`nFallback 5,44 s (без реално време): $(@($script:LastMedianFallbackCodes).Count)`r`nПреместени в Без_времена: $(@($script:LastNoTimePgmxMoved).Count)`r`nПропуснати непотвърдени изделия: $opSkipped`r`nДетайли с грешка / несъздадени: $opDetailErrors`r`nCSV със статус ЗАВЪРШЕН в регистъра: $regDone`r`nCSV за повторен опит (ЧАСТИЧНО/ГРЕШКА): $regRetry`r`n`r`nMixx:`r`n$mixPath`r`n`r`nПочистен PowerSet XLS:`r`n$xlsPath`r`n`r`nКопие в detail_time:`r`n$copyPath`r`n`r`nОтчетите са в избраната дневна папка.",
            'PWX500 Agent V5.1.38 - Стъпки 3 + 4 + 5',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        try { Show-FinalProcessReport } catch {}
    } catch {
        $msg = $_.Exception.Message
        if ($script:LastMaestroLauncher) { $msg += ' | Maestro source: ' + $script:LastMaestroLauncher }
        Write-MixLog ('AUTO CHAIN ERROR: ' + $msg)
        if($script:OperatorStopRequested -or $_.Exception -is [System.OperationCanceledException]){
            try{Set-OperatorCheckpoint -CurrentStage $script:OperatorCurrentStage -LastSafeStage $script:OperatorLastSafeStage -Status 'STOPPED' -Message 'Процесът е спрян безопасно.'}catch{}
            $status.Text='ПРОЦЕСЪТ Е СПРЯН. Продължи от последната сигурна точка.'
            return
        }
        try{Finalize-CsvRegistryForCurrentBatch -PipelineSucceeded $false -PipelineMessage $msg}catch{}
        try{Set-OperatorRunError -Message $msg}catch{}

        if ($step4Done) {
            try {
                $autoMixStatusLabel.Text = 'Стъпки 3 + 4 са готови, но Стъпка 5 даде грешка: ' + $msg
                $status.Text = 'Стъпки 3 + 4 ГОТОВИ; PowerSet ГРЕШКА: ' + $msg
                $tabs.SelectedTab = $tabPowerSet
            } catch {}
            [System.Windows.Forms.MessageBox]::Show(
                "Стъпки 3 и 4 приключиха успешно, но автоматичната Стъпка 5 - PowerSet не завърши:`r`n`r`n$msg`r`n`r`nMixx остава създаден:`r`n$mixPath",
                'PWX500 Agent V5.1.38 - Стъпка 5',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } else {
            try {
                $autoMixStatusLabel.Text = 'ГРЕШКА Стъпка 3 -> 4: ' + $msg
                $status.Text = 'Стъпка 3 приключи, но автоматичната Стъпка 4 даде грешка: ' + $msg
                $tabs.SelectedTab = $tabTraining
            } catch {}
            [System.Windows.Forms.MessageBox]::Show(
                "Стъпка 3 приключи, но автоматичната Стъпка 4 не завърши:`r`n`r`n$msg",
                'PWX500 Agent V5.1.38 - Стъпка 3 -> 4',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    } finally {
        $script:AutoMixChainRunning = $false
        $script:OperatorAutoFlow = $false
        try { Refresh-OperatorDashboard } catch {}
        try { $autoMixRunButton.Enabled = $true } catch {}
        try { $powerSetAutoRunButton.Enabled = $true } catch {}
    }
}

# -------------------- STEP 5: POWERSET AUTOMATION TEST --------------------
# V5.1.4 automates the confirmed PowerSet process from SCM/Machine through
# Run, waits for the real calculation end (no fixed duration), then exports
# projects data as Mix(dd.MM).xls into the same configured daily dd.MM folder.
function Write-PowerSetLog {
    param([string]$Message)
    try {
        $log = Join-Path $PSScriptRoot 'POWERSET_AUTOMATION_LOG.txt'
        Add-Content -LiteralPath $log -Value (('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$Message)) -Encoding UTF8
    } catch {}
}

function Set-PowerSetTestStatus {
    param([string]$Text)
    try { $powerSetAutoStatusLabel.Text = $Text } catch {}
    try { $status.Text = $Text } catch {}
    try { [System.Windows.Forms.Application]::DoEvents() } catch {}
}

function Invoke-PowerSetStartRunAutomation {
    $daily = Get-DailyMixFolder
    $dailyName = [System.IO.Path]::GetFileName($daily)
    $mixPath = Get-DailyMixFilePath
    $mixName = [System.IO.Path]::GetFileName($mixPath)

    if (-not [System.IO.Directory]::Exists($daily)) {
        throw "Липсва дневната папка: $daily"
    }
    if (-not [System.IO.File]::Exists($mixPath)) {
        throw "Липсва дневният Mixx: $mixPath"
    }
    if (-not (Initialize-MixAutomation)) {
        throw ('Не мога да инициализирам автоматизацията на Maestro. ' + $script:MixAutomationError)
    }

    $p = Start-OrActivateMaestro
    $maestroPid = $p.Id
    Write-PowerSetLog ("START steps 1-7. PID=$maestroPid; Mixx=$mixPath")

    # 1. SCM - learned position from the real PowerSet training session.
    Set-PowerSetTestStatus 'PowerSet TEST 1/7: SCM...'
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 98 -RefY 67
    Start-Sleep -Milliseconds 900

    # 2. Machine.
    Set-PowerSetTestStatus 'PowerSet TEST 2/7: Machine...'
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 835 -RefY 70
    Start-Sleep -Milliseconds 1200

    # 3. Tool equipment optimizer - third icon in the left machine menu.
    Set-PowerSetTestStatus 'PowerSet TEST 3/7: Tool equipment optimizer...'
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 37 -RefY 332
    Start-Sleep -Milliseconds 1700

    # 4. Analysis criteria -> Check existing tooling.
    # Do not touch any of the other preloaded checkboxes/options.
    Set-PowerSetTestStatus 'PowerSet TEST 4/7: Check existing tooling...'
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 421 -RefY 240
    Start-Sleep -Milliseconds 350
    $criteriaSelected = Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Check existing tooling' -TimeoutSec 2
    if (-not $criteriaSelected) {
        # Maestro does not expose the selected value reliably through UI Automation.
        # Use the exact learned click on the dropdown item instead.
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 250 -RefY 302
    }
    Start-Sleep -Milliseconds 850
    # V5.1.3: do not stop on UIA verification here. Maestro can visibly show
    # Check existing tooling as selected while the dropdown item is no longer
    # present in the UI Automation tree, which caused a false error in V5.1.2.
    Write-PowerSetLog ("Analysis criteria action completed. UIA invoke result=$criteriaSelected; continuing without false-negative verification.")

    # 5. Programs mix -> ... -> exact configured daily folder\Mix(dd.MM).mixx.
    Set-PowerSetTestStatus ("PowerSet TEST 5/7: избирам $mixName...")
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    # V5.1.6: Programs mix browse button. Maestro keeps this small WPF
    # button a few pixels lower/right on the 1656x918 workstation layout than
    # a pure 1920x1080 proportional scale predicts. Use the corrected learned
    # point so the cursor lands inside the actual '...' button.
    Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 430 -RefY 400
    Start-Sleep -Milliseconds 1300

    if (-not (Select-MixDailyFolderInDialog -ProcessId $maestroPid -DailyName $dailyName -ExpectedFileName $mixName -DailyPath $daily)) {
        throw "Не успях да отворя точната дневна папка в Choose the file: $daily"
    }
    if (-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name $mixName -TimeoutSec 4)) {
        throw "В дневната папка не успях да избера файла: $mixName"
    }
    Start-Sleep -Milliseconds 450

    # 6. Confirm file selection.
    Set-PowerSetTestStatus 'PowerSet TEST 6/7: Confirm...'
    if (-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Confirm' -TimeoutSec 4)) {
        throw 'Не успях да натисна Confirm при избора на Programs mix.'
    }
    Start-Sleep -Milliseconds 1300

    # 7. Optimization -> Run.  This is the end of the V5.1.2 test.
    Set-PowerSetTestStatus 'PowerSet TEST 7/7: Optimization -> Run...'
    [void](Activate-MixMaestroWindow -ProcessId $maestroPid)
    Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 565 -RefY 140
    Start-Sleep -Milliseconds 500

    Write-PowerSetLog ("RUN clicked. Test steps 1-7 completed. Mixx=$mixPath")
    Set-PowerSetTestStatus 'PowerSet TEST: стъпки 1-7 са изпълнени. Изчисляването трябва да е стартирало.'
    return $mixPath
}



function Get-PowerSetOptimizationUiSnapshot {
    param([int]$ProcessId)
    $watch=[System.Diagnostics.Stopwatch]::StartNew()
    $signals=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    $wanted=@('General result','Projects report','Total optimisation time','Feasible projects','Unfeasible projects')
    $busy=$false
    try{
        $root=[System.Windows.Automation.AutomationElement]::RootElement
        $condition=New-Object -TypeName System.Windows.Automation.PropertyCondition -ArgumentList @([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
        # FIX6: the old loop called Get-MixUiElementsByName SIX times, and
        # each call traversed the entire UIA process tree. Traverse only ONCE.
        $all=$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,$condition)
        foreach($el in $all){
            try{
                $name=[string]$el.Current.Name
                if($name -ne 'Undo' -and $name -notin $wanted){continue}
                $b=$el.Current.BoundingRectangle
                if($b.Width -le 1 -or $b.Height -le 1){continue}
                if($name -eq 'Undo'){$busy=$true}else{[void]$signals.Add($name)}
            }catch{}
        }
        return [pscustomobject]@{Available=$true;Busy=$busy;ResultSignals=$signals.Count;ScanMs=$watch.ElapsedMilliseconds}
    }catch{
        # The safe fallback is 'not yet complete'. Never infer an optimization
        # result from a failed UIA provider call or an absent dialog.
        Write-PowerSetLog ('PowerSet UIA snapshot error: '+$_.Exception.Message)
        return [pscustomobject]@{Available=$false;Busy=$false;ResultSignals=0;ScanMs=$watch.ElapsedMilliseconds}
    }finally{$watch.Stop()}
}

function Wait-PowerSetOptimizationComplete {
    param([int]$InitialProcessId,[int]$TimeoutMinutes=60)
    $started=Get-Date
    $until=$started.AddMinutes($TimeoutMinutes)
    $lastStatus=[DateTime]::MinValue
    $sawBusy=$false
    Write-PowerSetLog ('WAIT optimization. TimeoutMinutes='+$TimeoutMinutes+'; batched UIA snapshot=FIX6')
    do {
        Test-OperatorStopRequested
        $pNow=Get-MaestroProcess
        if($null -ne $pNow){
            $pidNow=$pNow.Id
            $snap=Get-PowerSetOptimizationUiSnapshot -ProcessId $pidNow
            if($snap.Available){
                if($snap.Busy){$sawBusy=$true}
                if($snap.ResultSignals -ge 2){
                    $elapsed=[int]((Get-Date)-$started).TotalSeconds
                    Write-PowerSetLog ('Optimization completed. PID='+$pidNow+'; elapsedSec='+$elapsed+'; resultSignals='+$snap.ResultSignals+'; sawBusy='+$sawBusy+'; scanMs='+$snap.ScanMs)
                    Set-PowerSetTestStatus ('PowerSet: оптимизацията приключи за '+$elapsed+' сек. Подготвям XLS export...')
                    return $pidNow
                }
                if($snap.ScanMs -gt 2000){Write-PowerSetLog ('SLOW UIA snapshot: '+$snap.ScanMs+' ms')}
            }
        }
        if(((Get-Date)-$lastStatus).TotalSeconds -ge 5){
            $elapsed=[int]((Get-Date)-$started).TotalSeconds
            Set-PowerSetTestStatus ('PowerSet: изчаквам края на оптимизацията... '+$elapsed+' сек.')
            $lastStatus=Get-Date
        }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 250
    }while((Get-Date) -lt $until)
    throw ('PowerSet не показа край на оптимизацията до '+$TimeoutMinutes+' минути. Не е стартиран export.')
}

function Invoke-PowerSetVisibleNamedElement {
    param([int]$ProcessId,[string]$Name,[switch]$PreferRightmost,[int]$TimeoutSec=4)
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do {
        $els=@(Get-MixUiElementsByName -ProcessId $ProcessId -Name $Name)
        if($els.Count -gt 0){
            $items=New-Object System.Collections.Generic.List[object]
            foreach($el in $els){
                try {
                    $b=$el.Current.BoundingRectangle
                    if($b.Width -gt 1 -and $b.Height -gt 1){
                        $items.Add([pscustomobject]@{El=$el;Left=$b.Left;Top=$b.Top;W=$b.Width;H=$b.Height})
                    }
                } catch {}
            }
            $arr=@($items.ToArray())
            if($arr.Count -gt 0){
                if($PreferRightmost){ $arr=@($arr | Sort-Object Left -Descending) }
                else { $arr=@($arr | Sort-Object Left) }
                $pick=$arr[0]
                try {
                    $pat=$null
                    if($pick.El.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern,[ref]$pat)){
                        $pat.Invoke(); return $true
                    }
                } catch {}
                try {
                    Invoke-MixNativeClick -X ([int]($pick.Left+$pick.W/2)) -Y ([int]($pick.Top+$pick.H/2))
                    return $true
                } catch {}
            }
        }
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 200
    } while((Get-Date) -lt $until)
    return $false
}

function Test-PowerSetSaveDialogOpen {
    param([int]$ProcessId)
    try { if($null -ne (Get-MixSaveNameEdit -ProcessId $ProcessId)){ return $true } } catch {}
    try { if(Test-MixUiElementByName -ProcessId $ProcessId -Name 'Save as...' -TimeoutSec 1){ return $true } } catch {}
    return $false
}

function Wait-PowerSetSaveDialogOpen {
    param([int]$ProcessId,[int]$TimeoutSec=10)
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do {
        try {
            if($null -ne (Get-MixSaveNameEdit -ProcessId $ProcessId)){
                Write-PowerSetLog 'PowerSet Save as dialog detected by Name edit.'
                return $true
            }
        } catch {}
        try {
            if(@(Get-MixUiElementsByName -ProcessId $ProcessId -Name 'Save as...').Count -gt 0){
                Write-PowerSetLog 'PowerSet Save as dialog detected by title.'
                return $true
            }
        } catch {}
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 250
    } while((Get-Date) -lt $until)
    return $false
}

function Test-PowerSetExportMenuOpen {
    param([int]$ProcessId,[int]$TimeoutMs=2500)
    # V5.1.9: the dedicated 13.10 training proved that the menu row is exposed
    # as a real UI Automation Text element named exactly "Export projects data"
    # with a normal visible rectangle (training bounds about 815,202,141,17).
    # Require a VISIBLE/non-offscreen element so a cached/hidden item cannot be
    # mistaken for an opened Result menu.
    $until=(Get-Date).AddMilliseconds($TimeoutMs)
    do {
        try {
            $els=@(Get-MixUiElementsByName -ProcessId $ProcessId -Name 'Export projects data')
            foreach($el in $els){
                try {
                    $b=$el.Current.BoundingRectangle
                    if((-not $el.Current.IsOffscreen) -and $b.Width -gt 5 -and $b.Height -gt 5){ return $true }
                } catch {}
            }
        } catch {}
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 120
    } while((Get-Date) -lt $until)
    return $false
}

function Open-PowerSetResultMenu {
    param([int]$ProcessId)

    # V5.1.9 is based on the dedicated Result-arrow training from 13.10.
    # The user intentionally clicked the arrow repeatedly only to teach its
    # exact position; no export/save actions from that training are inferred.
    # Recorded successful arrow clicks cluster at X=823..833, Y=128..150,
    # with the repeated center around X=828..829, Y=140..141 on the learned
    # Maestro layout. Start from that center and use only tiny fallbacks.
    $points = @(
        @(829,140),
        @(825,141),
        @(832,149),
        @(828,132)
    )

    [void](Activate-MixMaestroWindow -ProcessId $ProcessId)
    Start-Sleep -Milliseconds 300

    # If the dropdown is already open, never click the arrow again because a
    # second click would close it.
    if(Test-PowerSetExportMenuOpen -ProcessId $ProcessId -TimeoutMs 300){
        Write-PowerSetLog 'Result menu was already open; no arrow click needed.'
        return $true
    }

    foreach($pt in $points){
        $rx=[double]$pt[0]; $ry=[double]$pt[1]
        Write-PowerSetLog ("Trying trained Result arrow ref=($rx,$ry)")
        try { Invoke-MixRelativeClick -ProcessId $ProcessId -RefX $rx -RefY $ry } catch {}

        # Give Maestro enough time to expose the WPF dropdown row through UIA.
        # The older 900 ms window could cause a second click before UIA had
        # finished publishing the menu, effectively toggling it closed again.
        if(Test-PowerSetExportMenuOpen -ProcessId $ProcessId -TimeoutMs 2500){
            Write-PowerSetLog ("Result menu verified open at trained ref=($rx,$ry)")
            return $true
        }

        # Restore a deterministic closed state before the next tiny fallback.
        try { [System.Windows.Forms.SendKeys]::SendWait('{ESC}') } catch {}
        Start-Sleep -Milliseconds 300
    }
    return $false
}

function Open-PowerSetExportProjectsDialog {
    param([int]$ProcessId)

    # First open the tiny Result split-arrow robustly. The earlier versions
    # used RefX=690/RefY=116, which is not the learned arrow position and can
    # leave the cursor above/left of the actual arrow on the user's screen.
    if(-not (Open-PowerSetResultMenu -ProcessId $ProcessId)){
        Write-PowerSetLog 'Result dropdown did not open after all learned-arrow probes.'
        return $false
    }

    # Once the menu is open, stop using coordinates. The training recorder saw
    # the row itself as a real UI Automation Text element named exactly
    # "Export projects data". Invoke/click that named visible row directly.
    $invoked = Invoke-PowerSetVisibleNamedElement -ProcessId $ProcessId -Name 'Export projects data' -TimeoutSec 4
    if(-not $invoked){
        Write-PowerSetLog 'Result menu opened, but named Export projects data row was not invoked.'
        return $false
    }
    Write-PowerSetLog 'Clicked visible Export projects data row after verified Result dropdown open.'

    # The dark panel on the right is only the description/tooltip. There is no
    # second submenu. Wait without touching Maestro until Save as appears.
    if(Wait-PowerSetSaveDialogOpen -ProcessId $ProcessId -TimeoutSec 12){ return $true }

    Write-PowerSetLog 'Export projects data was invoked, but Save as did not appear within 12 seconds.'
    return $false
}

function Wait-PowerSetExportCompleted {
    param([int]$ProcessId,[int]$TimeoutSec=30)
    $until=(Get-Date).AddSeconds($TimeoutSec)
    do {
        if(@(Get-MixUiElementsByName -ProcessId $ProcessId -Name 'Export operation completed').Count -gt 0){
            return $true
        }
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 300
    } while((Get-Date) -lt $until)
    return $false
}

function Finalize-PowerSetXls {
    param([Parameter(Mandatory=$true)][string]$XlsPath)

    if(-not [System.IO.File]::Exists($XlsPath)){
        throw "Не намирам PowerSet XLS за обработка: $XlsPath"
    }

    $destinationDir = 'Z:\30_Biesse_edging\scm\detail_time'
    $destinationPath = Join-Path $destinationDir ([System.IO.Path]::GetFileName($XlsPath))

    Set-PowerSetTestStatus 'PowerSet: отварям XLS, почиствам горните редове и записвам...'
    Write-PowerSetLog ('POST XLS start: ' + $XlsPath)

    # Show the generated file in its daily folder, as in the manual workflow.
    try {
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $XlsPath) | Out-Null
        Start-Sleep -Milliseconds 500
    } catch {
        Write-PowerSetLog ('Explorer open warning: ' + $_.Exception.Message)
    }

    $excel = $null
    $workbook = $null
    $sheet = $null
    $deleteRange = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($XlsPath)
        $sheet = $workbook.Worksheets.Item(1)

        # Do not assume a fixed number of PowerSet header rows. Find the real
        # table header by its confirmed first columns: No., Project, R, DY, DZ.
        $headerRow = 0
        for($r = 1; $r -le 50; $r++){
            $a = (Safe-String $sheet.Cells.Item($r,1).Text).Trim()
            $b = (Safe-String $sheet.Cells.Item($r,2).Text).Trim()
            $c = (Safe-String $sheet.Cells.Item($r,3).Text).Trim()
            $d = (Safe-String $sheet.Cells.Item($r,4).Text).Trim()
            $e = (Safe-String $sheet.Cells.Item($r,5).Text).Trim()
            if($a -eq 'No.' -and $b -eq 'Project' -and $c -eq 'R' -and $d -eq 'DY' -and $e -eq 'DZ'){
                $headerRow = $r
                break
            }
        }

        if($headerRow -le 0){
            throw 'Не намерих заглавния ред No. / Project / R / DY / DZ в първите 50 реда на PowerSet XLS.'
        }

        Write-PowerSetLog ("PowerSet XLS header found at row $headerRow")
        if($headerRow -gt 1){
            $lastDeleteRow = $headerRow - 1
            $deleteRange = $sheet.Range('A1', ('A' + $lastDeleteRow))
            [void]$deleteRange.EntireRow.Delete()
            Write-PowerSetLog ("Deleted XLS rows 1-$lastDeleteRow so the table header becomes row 1.")
        } else {
            Write-PowerSetLog 'XLS table header is already row 1; no top rows deleted.'
        }

        $workbook.Save()
        Start-Sleep -Milliseconds 300
        $workbook.Close($true)
        $excel.Quit()
    } catch {
        $excelError = $_.Exception.Message
        try { if($null -ne $workbook){ $workbook.Close($false) } } catch {}
        try { if($null -ne $excel){ $excel.Quit() } } catch {}
        throw "Не успях да почистя/запиша PowerSet XLS в Excel: $excelError"
    } finally {
        try { if($null -ne $deleteRange){ [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deleteRange) } } catch {}
        try { if($null -ne $sheet){ [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) } } catch {}
        try { if($null -ne $workbook){ [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } } catch {}
        try { if($null -ne $excel){ [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } } catch {}
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if(-not [System.IO.File]::Exists($XlsPath)){
        throw "След обработката не намирам запазения XLS: $XlsPath"
    }

    if(-not [System.IO.Directory]::Exists($destinationDir)){
        throw "Целевата директория за detail_time не е достъпна: $destinationDir"
    }

    Set-PowerSetTestStatus ('PowerSet: копирам обработения XLS в ' + $destinationDir + '...')
    try {
        [System.IO.File]::Copy($XlsPath, $destinationPath, $true)
    } catch {
        throw "XLS е почистен и запазен локално, но не успях да го копирам в ${destinationDir}: $($_.Exception.Message)"
    }

    if(-not [System.IO.File]::Exists($destinationPath)){
        throw "Копирането приключи без грешка, но не намирам файла в detail_time: $destinationPath"
    }

    $srcInfo = Get-Item -LiteralPath $XlsPath -ErrorAction Stop
    $dstInfo = Get-Item -LiteralPath $destinationPath -ErrorAction Stop
    if($srcInfo.Length -le 0 -or $dstInfo.Length -ne $srcInfo.Length){
        throw "Проверката след копиране е неуспешна. Source=$($srcInfo.Length) bytes; Copy=$($dstInfo.Length) bytes."
    }

    Write-PowerSetLog ("POST XLS OK. Cleaned=$XlsPath; copied=$destinationPath; bytes=$($srcInfo.Length)")
    return $destinationPath
}

function Apply-MedianFallbackToPowerSetXls {
    param([Parameter(Mandatory=$true)][string]$XlsPath,[string[]]$FallbackCodes,[string]$CopyPath='')
    $codes=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach($c in @($FallbackCodes)){$k=Normalize-TimeProgramCode $c;if($k){[void]$codes.Add($k)}}
    if($codes.Count -eq 0){return 0}
    if(-not [System.IO.File]::Exists($XlsPath)){throw 'Липсва PowerSet XLS за прилагане на median fallback: '+$XlsPath}
    $excel=$null;$wb=$null;$ws=$null;$used=$null
    $changed=0
    try{
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.ScreenUpdating=$false;$excel.EnableEvents=$false
        $wb=$excel.Workbooks.Open($XlsPath)
        $ws=$wb.Worksheets.Item(1)
        $used=$ws.UsedRange
        $rows=[int]$used.Rows.Count;$cols=[int]$used.Columns.Count
        $projectCol=0;$timeCol=0
        for($c=1;$c -le $cols;$c++){
            $h=(Safe-String $ws.Cells.Item(1,$c).Text).Trim()
            if($projectCol -eq 0 -and $h -in @('Project','ProgramCodeF','Program','PGMX')){$projectCol=$c}
            if($timeCol -eq 0 -and $h -in @('Time (sec)','Time_sec','TimeSec','MeasuredTimeSec')){$timeCol=$c}
        }
        if($projectCol -le 0 -or $timeCol -le 0){throw 'Не намирам колоните Project и Time (sec) в почистения PowerSet XLS.'}
        for($r=2;$r -le $rows;$r++){
            $pc=Safe-String $ws.Cells.Item($r,$projectCol).Text
            $k=Normalize-TimeProgramCode $pc
            if(-not $k -or -not $codes.Contains($k)){continue}
            $raw=$ws.Cells.Item($r,$timeCol).Value2
            $tv=Parse-TimeValue $raw
            if($null -ne $tv -and [Math]::Abs([double]$tv) -lt 0.0000001){
                $ws.Cells.Item($r,$timeCol).Value2=[double]$script:MedianFallbackSec
                $changed++
            }
        }
        $wb.Save();$wb.Close($true);$excel.Quit()
    }catch{
        $msg=$_.Exception.Message
        try{if($wb){$wb.Close($false)}}catch{}
        try{if($excel){$excel.Quit()}}catch{}
        throw ('Не успях да заменя 0 с 5,44 s в PowerSet XLS: '+$msg)
    }finally{
        try{if($used){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($used)}}catch{}
        try{if($ws){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws)}}catch{}
        try{if($wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}}catch{}
        try{if($excel){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}catch{}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
    }
    if($CopyPath){
        [System.IO.File]::Copy($XlsPath,$CopyPath,$true)
        $a=Get-Item -LiteralPath $XlsPath -ErrorAction Stop;$b=Get-Item -LiteralPath $CopyPath -ErrorAction Stop
        if($a.Length -le 0 -or $a.Length -ne $b.Length){throw 'Повторното копиране на XLS след median fallback не премина проверката.'}
    }
    Write-PowerSetLog ('MEDIAN FALLBACK XLS: changed='+$changed+'; codes='+$codes.Count+'; value='+$script:MedianFallbackSec)
    return $changed
}

function Invoke-PowerSetFullAutomation {
    Test-OperatorStopRequested
    $mixPath = Invoke-PowerSetStartRunAutomation
    $daily = Get-DailyMixFolder
    $dailyName = [System.IO.Path]::GetFileName($daily)
    $xlsBaseName = 'Mix(' + $dailyName + ')'
    $xlsPath = Join-Path $daily ($xlsBaseName + '.xls')

    $pNow = Get-MaestroProcess
    if($null -eq $pNow){ throw 'След Run не намирам активния прозорец MAESTRO LAB.' }

    Test-OperatorStopRequested
    $maestroPid = Wait-PowerSetOptimizationComplete -InitialProcessId $pNow.Id -TimeoutMinutes 60

    Set-PowerSetTestStatus 'PowerSet: Result -> Export projects data...'
    if(-not (Open-PowerSetExportProjectsDialog -ProcessId $maestroPid)){
        throw 'Оптимизацията приключи, но не успях да отворя Result -> Export projects data.'
    }

    Start-Sleep -Milliseconds 800
    Set-PowerSetTestStatus ("PowerSet: избирам дневната папка $dailyName за XLS...")
    if(-not (Select-MixDailyFolderInDialog -ProcessId $maestroPid -DailyName $dailyName -ExpectedFileName '' -DailyPath $daily)){
        throw "Не успях да избера дневната папка $daily в PowerSet Save as."
    }

    # PowerSet pre-fills Mix(dd.MM) automatically. Set it explicitly as a
    # safety check so the report is always Mix(dd.MM).xls.
    if(-not (Set-MixSaveName -ProcessId $maestroPid -Value $xlsBaseName)){
        throw "Не успях да задада името на PowerSet отчета: $xlsBaseName.xls"
    }

    # A previous report from the same date is an output file, not an input.
    # Remove only this exact daily XLS immediately before creating the new one,
    # avoiding an overwrite prompt. If removal fails (for example the XLS is
    # open in Excel), stop safely and explain the reason.
    if([System.IO.File]::Exists($xlsPath)){
        try {
            [System.IO.File]::Delete($xlsPath)
            Write-PowerSetLog ('Deleted existing daily XLS before rebuild: ' + $xlsPath)
        } catch {
            throw "Съществуващият дневен XLS не може да бъде заменен. Затвори файла, ако е отворен в Excel: $xlsPath"
        }
    }

    Set-PowerSetTestStatus ("PowerSet: записвам $xlsBaseName.xls...")
    if(-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'Confirm' -TimeoutSec 5)){
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 1650 -RefY 880
    }

    if(-not (Wait-PowerSetExportCompleted -ProcessId $maestroPid -TimeoutSec 30)){
        throw 'PowerSet не показа съобщение Export operation completed след записа.'
    }

    if(-not (Invoke-MixUiElementByName -ProcessId $maestroPid -Name 'OK' -TimeoutSec 5)){
        Invoke-MixRelativeClick -ProcessId $maestroPid -RefX 1272 -RefY 574
    }

    $fileUntil=(Get-Date).AddSeconds(15)
    do {
        Test-OperatorStopRequested
        if([System.IO.File]::Exists($xlsPath)){
            $fi = $null
            try { $fi=Get-Item -LiteralPath $xlsPath -ErrorAction Stop } catch {}
            if($null -ne $fi -and $fi.Length -gt 0){
                # From here on, post-processing errors must be visible to the user.
                $copiedPath = Finalize-PowerSetXls -XlsPath $xlsPath
                try { $script:LastTimeImportStats = Import-TimeWorkbookToDb -Path $xlsPath -Force } catch { Write-PowerSetLog ('TIME DB IMPORT ERROR: '+$_.Exception.Message); throw ('PowerSet създаде и копира XLS, но импортът на времената не е завършен: '+$_.Exception.Message) }
                $script:LastMedianFallbackCodes=@($script:LastTimeImportStats.FallbackCodes)
                if($script:LastMedianFallbackCodes.Count -gt 0){
                    $changed=Apply-MedianFallbackToPowerSetXls -XlsPath $xlsPath -FallbackCodes $script:LastMedianFallbackCodes -CopyPath $copiedPath
                    $script:LastNoTimePgmxMoved=@(Move-DailyPgmxCodesToSubfolder -Codes $script:LastMedianFallbackCodes -Subfolder 'Без_времена' -Reason 'PowerSet 0 -> median 5.44 s')
                    Write-PowerSetLog ('NO-TIME SORT: fallback codes='+$script:LastMedianFallbackCodes.Count+'; XLS changed='+$changed+'; PGMX moved='+$script:LastNoTimePgmxMoved.Count)
                }
                Write-PowerSetLog ("FULL TEST OK. Mixx=$mixPath; XLS=$xlsPath; copied=$copiedPath; bytes=$($fi.Length); medianFallback=$($script:LastMedianFallbackCodes.Count)")
                Set-PowerSetTestStatus ('PowerSet ГОТОВО: почистен XLS + копие в detail_time; fallback 5,44 s: '+$script:LastMedianFallbackCodes.Count)
                return $xlsPath
            }
        }
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 300
    } while((Get-Date) -lt $fileUntil)

    throw "PowerSet потвърди export, но не намирам създадения файл: $xlsPath"
}

function Convert-ToCsvField {
    param($Value)
    $s = Safe-String $Value
    if ($s.Contains('"')) { $s = $s.Replace('"', '""') }
    if ($s.Contains(';') -or $s.Contains('"') -or $s.Contains("`r") -or $s.Contains("`n")) {
        return '"' + $s + '"'
    }
    return $s
}


# ---------------- V5.1.31 SAFE STOP / RESUME / PERSISTENT HISTORY ----------------
function Get-OperatorRunState {
    try {
        if(-not [System.IO.File]::Exists($script:OperatorRunStateFile)){return $null}
        $txt=[System.IO.File]::ReadAllText($script:OperatorRunStateFile,[System.Text.Encoding]::UTF8)
        if([string]::IsNullOrWhiteSpace($txt)){return $null}
        return ($txt | ConvertFrom-Json)
    } catch { return $null }
}

function Add-OperatorProcessHistory {
    param([string]$Stage,[string]$Status,[string]$Message='')
    try {
        $header='Timestamp;Stage;Status;Message'
        if(-not [System.IO.File]::Exists($script:OperatorProcessHistoryCsv)){
            [System.IO.File]::WriteAllLines($script:OperatorProcessHistoryCsv,@($header),(New-Object System.Text.UTF8Encoding($true)))
        }
        $ts=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $vals=@($ts,$Stage,$Status,$Message) | ForEach-Object { Convert-ToCsvField $_ }
        [System.IO.File]::AppendAllText($script:OperatorProcessHistoryCsv,(($vals -join ';')+[Environment]::NewLine),(New-Object System.Text.UTF8Encoding($false)))
    } catch {}
}

function Set-OperatorCheckpoint {
    param(
        [Parameter(Mandatory=$true)][string]$CurrentStage,
        [string]$LastSafeStage='',
        [ValidateSet('ACTIVE','STOPPED','ERROR','COMPLETED')][string]$Status='ACTIVE',
        [string]$Message=''
    )
    try {
        $old=Get-OperatorRunState
        $now=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $started=$now;$session=[guid]::NewGuid().ToString()
        if($null -ne $old){
            if(-not [string]::IsNullOrWhiteSpace([string]$old.StartedAt)){$started=[string]$old.StartedAt}
            if(-not [string]::IsNullOrWhiteSpace([string]$old.SessionId)){$session=[string]$old.SessionId}
            if([string]::IsNullOrWhiteSpace($LastSafeStage)){$LastSafeStage=[string]$old.LastSafeStage}
        }
        if([string]::IsNullOrWhiteSpace($LastSafeStage)){$LastSafeStage='START'}
        $script:OperatorCurrentStage=$CurrentStage
        $script:OperatorLastSafeStage=$LastSafeStage
        $daily='';try{$daily=Get-DailyMixFolder}catch{}
        $obj=[ordered]@{
            Version='5.1.23';SessionId=$session;StartedAt=$started;UpdatedAt=$now;Status=$Status;
            CurrentStage=$CurrentStage;LastSafeStage=$LastSafeStage;Message=$Message;DailyFolder=$daily;FlowMode=$script:OperatorExecutionMode
        }
        $json=$obj | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($script:OperatorRunStateFile,$json,(New-Object System.Text.UTF8Encoding($true)))
        Add-OperatorProcessHistory -Stage $CurrentStage -Status $Status -Message $Message
    } catch {}
}

function Start-NewOperatorRunState {
    param([string]$Message='Нов дневен процес')
    try{if([System.IO.File]::Exists($script:OperatorRunStateFile)){[System.IO.File]::Delete($script:OperatorRunStateFile)}}catch{}
    $script:OperatorStopRequested=$false
    Set-OperatorCheckpoint -CurrentStage 'START' -LastSafeStage 'START' -Status 'ACTIVE' -Message $Message
}

function Complete-OperatorRunState {
    param([string]$Message='Дневният процес е завършен успешно.')
    Set-OperatorCheckpoint -CurrentStage 'COMPLETED' -LastSafeStage 'POWERSET_DONE' -Status 'COMPLETED' -Message $Message
}

function Set-OperatorRunError {
    param([string]$Message)
    $safe=$script:OperatorLastSafeStage
    if([string]::IsNullOrWhiteSpace($safe)){$safe='START'}
    Set-OperatorCheckpoint -CurrentStage $script:OperatorCurrentStage -LastSafeStage $safe -Status 'ERROR' -Message $Message
}

function Save-MigrationStatus {
    param([string]$Result,[string]$WorkbookPath='',[string]$Message='',[switch]$Success)
    try {
        $old=$null
        if([System.IO.File]::Exists($script:MigrationStatusFile)){
            try{$old=([System.IO.File]::ReadAllText($script:MigrationStatusFile,[System.Text.Encoding]::UTF8) | ConvertFrom-Json)}catch{}
        }
        $now=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $lastSuccess=''
        if($null -ne $old){$lastSuccess=[string]$old.LastSuccessAt}
        if($Success){$lastSuccess=$now}
        $obj=[ordered]@{LastAttemptAt=$now;LastSuccessAt=$lastSuccess;LastWorkbook=$WorkbookPath;LastResult=$Result;LastMessage=$Message}
        [System.IO.File]::WriteAllText($script:MigrationStatusFile,($obj|ConvertTo-Json -Depth 3),(New-Object System.Text.UTF8Encoding($true)))
    } catch {}
}

function Get-MigrationStatus {
    try {
        if([System.IO.File]::Exists($script:MigrationStatusFile)){
            return ([System.IO.File]::ReadAllText($script:MigrationStatusFile,[System.Text.Encoding]::UTF8) | ConvertFrom-Json)
        }
        if([System.IO.File]::Exists($script:HistoryMarkerFile)){
            $lines=[System.IO.File]::ReadAllLines($script:HistoryMarkerFile,[System.Text.Encoding]::UTF8)
            $completed='';$wb=''
            foreach($ln in $lines){if($ln -like 'CompletedAt=*'){$completed=$ln.Substring(12)}elseif($ln -like 'Workbook=*'){$wb=$ln.Substring(9)}}
            return [pscustomobject]@{LastAttemptAt=$completed;LastSuccessAt=$completed;LastWorkbook=$wb;LastResult='SUCCESS';LastMessage='Миграцията е изпълнена.'}
        }
    } catch {}
    return $null
}

function Get-FriendlyCheckpointName {
    param([string]$Stage)
    switch($Stage){
        'START' {return 'начало'}
        'CATALOG_DONE' {return 'дневен каталог'}
        'CSV_DONE' {return 'CSV прочетени'}
        'SEARCH_DONE' {return 'папки намерени'}
        'HOP_DONE' {return 'HOP проверка'}
        'PGMX_DONE' {return 'PGMX готови'}
        'MIXX_DONE' {return 'MIXX готов'}
        'POWERSET_DONE' {return 'PowerSet + XLS готови'}
        default {return $Stage}
    }
}

function Get-OperatorDoneSummaryText {
    $parts=New-Object System.Collections.Generic.List[string]
    try {
        $mig=Get-MigrationStatus
        if($null -eq $mig){$parts.Add('Бърза миграция: НЕ Е ИЗПЪЛНЯВАНА')}
        else {
            $when=if(-not [string]::IsNullOrWhiteSpace([string]$mig.LastSuccessAt)){[string]$mig.LastSuccessAt}else{[string]$mig.LastAttemptAt}
            $wb='';try{$wb=[System.IO.Path]::GetFileName([string]$mig.LastWorkbook)}catch{$wb=[string]$mig.LastWorkbook}
            $mr=[string]$mig.LastResult
            switch($mr){
                'SUCCESS' {$mr='ИЗПЪЛНЕНА УСПЕШНО'}
                'RUNNING' {$mr='В ПРОЦЕС'}
                'ERROR' {$mr='ГРЕШКА'}
                'STOPPED' {$mr='СПРЯНА'}
            }
            $parts.Add(('Бърза миграция: '+$mr+' | '+$when+' | '+$wb))
        }
    } catch {}
    try {
        if([System.IO.File]::Exists($script:IndexCsv)){$fi=Get-Item -LiteralPath $script:IndexCsv;$parts.Add('Каталог изделия: '+$fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))}
        else{$parts.Add('Каталог изделия: няма създаден каталог')}
    } catch {}
    try {
        $rs=Get-OperatorRunState
        if($null -ne $rs){$friendlySafe=Get-FriendlyCheckpointName -Stage ([string]$rs.LastSafeStage);$parts.Add(('Последен процес: '+[string]$rs.Status+' | безопасна точка: '+$friendlySafe+' | '+[string]$rs.UpdatedAt))}
        else{$parts.Add('Последен процес: няма записана сесия')}
    } catch {}
    try {
        $mix=Get-DailyMixFilePath
        if([System.IO.File]::Exists($mix)){if(Test-DailyMixReadyForDisplay){$parts.Add('MIXX: '+(Get-Item -LiteralPath $mix).LastWriteTime.ToString('HH:mm:ss'))}else{$parts.Add('MIXX: записът не е потвърден')}}
        $daily=Get-DailyMixFolder;$xls=Join-Path $daily ('Mix('+(Get-Date -Format 'dd.MM')+').xls')
        if([System.IO.File]::Exists($xls)){$parts.Add('PowerSet XLS: '+(Get-Item -LiteralPath $xls).LastWriteTime.ToString('HH:mm:ss'))}
    } catch {}
    return ($parts -join '     |     ')
}

function Stop-OperatorProcessSafely {
    param([switch]$Silent)
    $ans=[System.Windows.Forms.DialogResult]::Yes
    if(-not $Silent){$ans=[System.Windows.Forms.MessageBox]::Show('Да прекъсна ли текущия процес? Вече завършеното ще се запази и после ще може да продължиш от последната сигурна точка.','PWX500 Agent - Спиране',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)}
    if($ans -ne [System.Windows.Forms.DialogResult]::Yes){return}
    $script:OperatorStopRequested=$true
    $script:OperatorAutoFlow=$false;$script:OperatorRunHopAfterSearch=$false;$script:OperatorRunExecuteAfterHop=$false
    $hadWork=$false
    $historyWasRunning=$false;try{$historyWasRunning=($null -ne $script:HistoryWorkerProcess -and -not $script:HistoryWorkerProcess.HasExited)}catch{}
    foreach($name in @('IndexWorkerProcess','CsvWorkerProcess','SearchWorkerProcess','HopWorkerProcess','PgmxWorkerProcess','HistoryWorkerProcess')){
        try{
            $p=Get-Variable -Scope Script -Name $name -ValueOnly -ErrorAction SilentlyContinue
            if($null -ne $p -and -not $p.HasExited){$hadWork=$true;Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue}
            if($null -ne $p){try{$p.Dispose()}catch{}}
            Set-Variable -Scope Script -Name $name -Value $null
        }catch{}
    }
    try{Stop-IndexWorkerTimer}catch{};try{Stop-CsvWorkerTimer}catch{};try{Stop-SearchWorkerTimer}catch{};try{Stop-HopWorkerTimer}catch{};try{Stop-PgmxWorkerTimer}catch{};try{Stop-HistoryWorkerTimer}catch{}
    if($historyWasRunning){Save-MigrationStatus -Result 'STOPPED' -WorkbookPath $script:HistoryLastWorkbook -Message 'Миграцията е прекъсната ръчно.'}
    if($script:AutoMixChainRunning){$hadWork=$true}
    $safe=$script:OperatorLastSafeStage;if([string]::IsNullOrWhiteSpace($safe)){$safe='START'}
    Set-OperatorCheckpoint -CurrentStage $script:OperatorCurrentStage -LastSafeStage $safe -Status 'STOPPED' -Message 'Процесът е спрян ръчно от оператора.'
    $status.Text='ПРОЦЕСЪТ Е СПРЯН. Можеш да натиснеш „Продължи прекъснатия процес“.'
    try{Refresh-MainOperatorUi}catch{}
}

function Test-OperatorStopRequested {
    if($script:OperatorStopRequested){throw (New-Object System.OperationCanceledException -ArgumentList 'Процесът е спрян от оператора.')}
}

function Restore-SearchResultsFromDisk {
    if(-not [System.IO.File]::Exists($script:SearchResultsCsv)){return $false}
    try{
        $rows=@(Read-SafeCsv -Path $script:SearchResultsCsv);$objects=New-Object System.Collections.Generic.List[object]
        foreach($x in $rows){$objects.Add([pscustomobject]@{Product=Get-PropertyValue $x 'Product';Candidates=Get-PropertyValue $x 'Candidates';Status=Get-PropertyValue $x 'Status';BestScore=Get-PropertyValue $x 'BestScore';Paths=Get-PropertyValue $x 'Paths'})}
        $script:SearchResultObjects=@($objects.ToArray());Refresh-CandidateGridFromSearchResults;return $true
    }catch{return $false}
}

function Restore-PgmxResultsFromDisk {
    if(-not [System.IO.File]::Exists($script:PgmxResultsCsv)){return $false}
    try{
        $rows=@(Read-SafeCsv -Path $script:PgmxResultsCsv);$objs=New-Object System.Collections.Generic.List[object]
        foreach($x in $rows){$objs.Add([pscustomobject]@{CSV=Get-PropertyValue $x 'CSV';Product=Get-PropertyValue $x 'Product';DetailNo=Get-PropertyValue $x 'DetailNo';DetailName=Get-PropertyValue $x 'DetailName';ProcessingNo=Get-PropertyValue $x 'ProcessingNo';ExpectedSuffix=Get-PropertyValue $x 'ExpectedSuffix';ProgramCodeF=Get-PropertyValue $x 'ProgramCodeF';ProductPath=Get-PropertyValue $x 'ProductPath';PwxFolder=Get-PropertyValue $x 'PwxFolder';SourceHop=Get-PropertyValue $x 'SourceHop';BasePgmx=Get-PropertyValue $x 'BasePgmx';TargetPgmx=Get-PropertyValue $x 'TargetPgmx';FolderStatus=Get-PropertyValue $x 'FolderStatus';HopStatus=Get-PropertyValue $x 'HopStatus';BaseStatus=Get-PropertyValue $x 'BaseStatus';TargetStatus=Get-PropertyValue $x 'TargetStatus';Validation=Get-PropertyValue $x 'Validation';Action=Get-PropertyValue $x 'Action';Result=Get-PropertyValue $x 'Result';Error=Get-PropertyValue $x 'Error';CentralFolder=Get-PropertyValue $x 'CentralFolder';CentralStatus=Get-PropertyValue $x 'CentralStatus';CentralError=Get-PropertyValue $x 'CentralError'})}
        $script:PgmxResultObjects=@($objs.ToArray());$script:PgmxLastMode='Execute';Capture-CreatedPgmxThisRun|Out-Null;Refresh-PgmxGrid;return $true
    }catch{return $false}
}

function Invoke-OperatorResumeInterruptedRun {
    $rs=Get-OperatorRunState
    if($null -eq $rs){[System.Windows.Forms.MessageBox]::Show('Няма записана незавършена сесия.','PWX500 Agent')|Out-Null;return}
    if([string]$rs.Status -eq 'COMPLETED'){[System.Windows.Forms.MessageBox]::Show('Последната сесия е завършена. Стартирай нов дневен процес.','PWX500 Agent')|Out-Null;return}
    $script:OperatorStopRequested=$false
    $script:OperatorExecutionMode=if([string]$rs.FlowMode -eq 'PGMX_ONLY'){'PGMX_ONLY'}else{'FULL'}
    $safe=[string]$rs.LastSafeStage
    $status.Text='Възстановяване от последната сигурна точка: '+$safe
    try{
        switch($safe){
            'POWERSET_DONE' { Complete-OperatorRunState -Message 'Възстановено: PowerSet/XLS вече са завършени.';Refresh-MainOperatorUi;return }
            'MIXX_DONE' {
                $mix=Get-DailyMixFilePath
                if(-not [System.IO.File]::Exists($mix)){throw 'Checkpoint MIXX_DONE е записан, но Mixx файлът липсва.'}
                Set-OperatorCheckpoint -CurrentStage 'POWERSET_RUNNING' -LastSafeStage 'MIXX_DONE' -Status 'ACTIVE' -Message 'Продължаване към PowerSet.'
                $xls=Invoke-PowerSetFullAutomation
                Set-OperatorCheckpoint -CurrentStage 'POWERSET_DONE' -LastSafeStage 'POWERSET_DONE' -Status 'ACTIVE' -Message ('PowerSet XLS: '+$xls)
                Complete-OperatorRunState
                Refresh-MainOperatorUi;return
            }
            'PGMX_DONE' {
                if(Restore-PgmxResultsFromDisk){if($script:OperatorExecutionMode -eq 'PGMX_ONLY'){Finish-PgmxOnlyStage}else{Invoke-AutoMixAfterStep3};return}
            }
            'HOP_DONE' {
                if([System.IO.File]::Exists($script:CsvRowsOut)){Load-CsvWorkerResults}else{throw 'Липсва локалният CSV checkpoint.'}
                if(-not (Restore-SearchResultsFromDisk)){throw 'Липсва локалният checkpoint от търсенето.'}
                Start-BackgroundPgmx -Mode 'Execute';return
            }
            'SEARCH_DONE' {
                if([System.IO.File]::Exists($script:CsvRowsOut)){Load-CsvWorkerResults}else{throw 'Липсва локалният CSV checkpoint.'}
                if(-not (Restore-SearchResultsFromDisk)){throw 'Липсва локалният checkpoint от търсенето.'}
                $script:OperatorRunExecuteAfterHop=$true;Start-BackgroundHopCheck;return
            }
            'CSV_DONE' {
                if([System.IO.File]::Exists($script:CsvRowsOut)){Load-CsvWorkerResults}else{throw 'Липсва локалният CSV checkpoint.'}
                $script:OperatorRunHopAfterSearch=$true;$script:OperatorRunExecuteAfterHop=$true;Start-BackgroundProductSearch;return
            }
            default { Start-OperatorFullFlow -ResumeExisting -Mode $script:OperatorExecutionMode;return }
        }
        Start-OperatorFullFlow -ResumeExisting -Mode $script:OperatorExecutionMode
    }catch{
        Set-OperatorRunError -Message ('Възстановяването не успя: '+$_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show('Не успях да продължа директно от checkpoint. Ще е безопасно да стартираш дневния процес отново; завършените CSV няма да се повторят.'+[Environment]::NewLine+[Environment]::NewLine+$_.Exception.Message,'PWX500 Agent - Възстановяване',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null
    }
}

function Prompt-ResumeIfNeeded {
    if($script:OperatorResumePromptShown){return};$script:OperatorResumePromptShown=$true
    $rs=Get-OperatorRunState
    if($null -eq $rs -or [string]$rs.Status -eq 'COMPLETED'){return}
    $friendlySafe=Get-FriendlyCheckpointName -Stage ([string]$rs.LastSafeStage)
    $msg="Намерена е незавършена сесия.`r`n`r`nСтатус: $($rs.Status)`r`nПоследна сигурна точка: $friendlySafe`r`nПоследна промяна: $($rs.UpdatedAt)`r`n`r`nДа продължа ли от последната сигурна точка?`r`nИзбери Не, ако искаш да започнеш нов дневен процес."
    $ans=[System.Windows.Forms.MessageBox]::Show($msg,'PWX500 Agent - Незавършена сесия',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
    if($ans -eq [System.Windows.Forms.DialogResult]::Yes){Invoke-OperatorResumeInterruptedRun}
    else{$mode=if([string]$rs.FlowMode -eq 'PGMX_ONLY'){'PGMX_ONLY'}else{'FULL'};Start-OperatorFullFlow -Mode $mode}
}

# ---------------- V5.1.31 OPERATOR WORKFLOW / FINAL REPORTS / CSV REGISTRY ----------------
function Get-OperatorAllProducts {
    return @($script:LastAllRows | ForEach-Object { Safe-Trim $_.Product } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
}

function Get-OperatorProductStateRows {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach($product in @(Get-OperatorAllProducts)){
        $confirmed=$false; $path=''; $source=''; $reason=''; $searchStatus=''; $candidates=''
        if($script:ProductFolderMap.ContainsKey($product)){
            $m=$script:ProductFolderMap[$product]
            $path=Safe-Trim $m.Path; $source=Safe-Trim $m.Source
            if(-not [System.IO.Directory]::Exists($path)){
                $reason='Запомнената папка не е достъпна.'
            } elseif(Test-IsTrustedProductMappingSource -Source $source){
                $confirmed=$true
                if($source -eq 'HISTORICAL_PGMX'){$reason='Потвърдено от Бърза миграция.'}
                else{$reason='Потвърдено по HOP.'}
            } else {
                $reason='Има зададена папка, но още не е HOP-потвърдена.'
            }
        }
        $sr=Get-SearchResultForProduct -Product $product
        if($null -ne $sr){
            $searchStatus=Safe-Trim $sr.Status
            $candidates=Safe-Trim $sr.Candidates
            if(-not $confirmed -and [string]::IsNullOrWhiteSpace($reason)){$reason=$searchStatus}
        } elseif(-not $confirmed -and [string]::IsNullOrWhiteSpace($reason)){
            $reason='Няма резултат от търсене.'
        }
        $rows.Add([pscustomobject]@{Product=$product;Confirmed=$confirmed;Path=$path;Source=$source;Candidates=$candidates;SearchStatus=$searchStatus;Reason=$reason})
    }
    return @($rows.ToArray())
}

function Get-OperatorPendingHopCount {
    $count=0
    foreach($x in @($script:SearchResultObjects)){
        $product=[string]$x.Product
        if($script:ProductFolderMap.ContainsKey($product)){
            $src=[string]$script:ProductFolderMap[$product].Source
            if(-not (Test-IsTrustedProductMappingSource -Source $src)){$count++}
        } else {
            $n=0
            if([int]::TryParse([string]$x.Candidates,[ref]$n) -and $n -ge 1){$count++}
        }
    }
    return $count
}

function Get-OperatorDetailErrorText {
    param($X)
    $err=Safe-Trim $X.Error
    if(-not [string]::IsNullOrWhiteSpace($err)){return $err}
    $cerr=Safe-Trim $X.CentralError
    if(-not [string]::IsNullOrWhiteSpace($cerr)){return $cerr}
    $parts=New-Object System.Collections.Generic.List[string]
    foreach($v in @($X.Result,$X.Validation,$X.HopStatus,$X.FolderStatus)){
        $s=Safe-Trim $v
        if(-not [string]::IsNullOrWhiteSpace($s) -and -not $parts.Contains($s)){$parts.Add($s)}
    }
    return ($parts -join ' | ')
}

function Write-OperatorTableCsv {
    param([Parameter(Mandatory=$true)]$Table,[Parameter(Mandatory=$true)][string]$Path)
    $lines=New-Object System.Collections.Generic.List[string]
    $headers=@(); foreach($c in $Table.Columns){$headers += (Convert-ToCsvField $c.ColumnName)}
    $lines.Add(($headers -join ';'))
    foreach($row in $Table.Rows){
        $vals=@(); foreach($c in $Table.Columns){$vals += (Convert-ToCsvField $row[$c.ColumnName])}
        $lines.Add(($vals -join ';'))
    }
    [System.IO.File]::WriteAllLines($Path,$lines.ToArray(),(New-Object System.Text.UTF8Encoding($true)))
}

function Refresh-OperatorReports {
    param([switch]$WriteFiles)
    $skipTable=New-DataTable -Columns @('Изделие','Причина','Кандидати','Статус търсене','Зададена папка','Източник')
    foreach($x in @(Get-OperatorProductStateRows | Where-Object { -not $_.Confirmed })){
        $r=$skipTable.NewRow();$r['Изделие']=$x.Product;$r['Причина']=$x.Reason;$r['Кандидати']=$x.Candidates;$r['Статус търсене']=$x.SearchStatus;$r['Зададена папка']=$x.Path;$r['Източник']=$x.Source;[void]$skipTable.Rows.Add($r)
    }

    $errTable=New-DataTable -Columns @('CSV','Изделие','№','Детайл','ProgramCodeF','Резултат','Грешка','Валидация','HOP статус','Папка изделие')
    $success=@('ГОТОВ','СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF')
    foreach($x in @($script:PgmxResultObjects | Where-Object { $success -notcontains (Safe-Trim $_.Result) })){
        $r=$errTable.NewRow();$r['CSV']=Safe-Trim $x.CSV;$r['Изделие']=Safe-Trim $x.Product;$r['№']=Safe-Trim $x.DetailNo;$r['Детайл']=Safe-Trim $x.DetailName;$r['ProgramCodeF']=Safe-Trim $x.ProgramCodeF;$r['Резултат']=Safe-Trim $x.Result;$r['Грешка']=Get-OperatorDetailErrorText -X $x;$r['Валидация']=Safe-Trim $x.Validation;$r['HOP статус']=Safe-Trim $x.HopStatus;$r['Папка изделие']=Safe-Trim $x.ProductPath;[void]$errTable.Rows.Add($r)
    }

    try{if($null -ne $operatorSkippedGrid){$operatorSkippedGrid.DataSource=$skipTable}}catch{}
    try{if($null -ne $operatorErrorsGrid){$operatorErrorsGrid.DataSource=$errTable}}catch{}
    try{if($null -ne $operatorSkippedCountLabel){$operatorSkippedCountLabel.Text='Непотвърдени / пропуснати изделия: '+$skipTable.Rows.Count}}catch{}
    try{if($null -ne $operatorErrorCountLabel){$operatorErrorCountLabel.Text='Детайли, които не са направени: '+$errTable.Rows.Count}}catch{}
    try{if($null -ne $script:MainSkippedGrid){$script:MainSkippedGrid.DataSource=$skipTable}}catch{}
    try{if($null -ne $script:MainErrorsGrid){$script:MainErrorsGrid.DataSource=$errTable}}catch{}
    try{if($null -ne $script:MainUnconfirmedValue){$script:MainUnconfirmedValue.Text=[string]$skipTable.Rows.Count}}catch{}
    try{if($null -ne $script:MainErrorsValue){$script:MainErrorsValue.Text=[string]$errTable.Rows.Count}}catch{}

    if($WriteFiles){
        try{
            $folder=Get-DailyMixFolder
            if(-not [System.IO.Directory]::Exists($folder)){[void][System.IO.Directory]::CreateDirectory($folder)}
            $stamp=Get-Date -Format 'dd.MM'
            $skipPath=Join-Path $folder ('PWX500_Пропуснати_изделия_'+$stamp+'.csv')
            $errPath=Join-Path $folder ('PWX500_Грешки_детайли_'+$stamp+'.csv')
            Write-OperatorTableCsv -Table $skipTable -Path $skipPath
            Write-OperatorTableCsv -Table $errTable -Path $errPath
            $script:OperatorLastSkippedReport=$skipPath
            $script:OperatorLastDetailErrorReport=$errPath
        }catch{}
    }
}

function Refresh-OperatorDashboard {
    try{
        $products=@(Get-OperatorProductStateRows);$confirmed=@($products|Where-Object{$_.Confirmed}).Count;$unconfirmed=$products.Count-$confirmed
        $csvCount=0;try{if($null -ne $script:LastFileSummary){$csvCount=@($script:LastFileSummary.Rows).Count}}catch{}
        if($null -ne $operatorCsvSummaryLabel){$operatorCsvSummaryLabel.Text=('CSV в текущия цикъл: '+$csvCount+' | изделия: '+$products.Count+' | детайли: '+@($script:LastAllRows).Count)}
        if($null -ne $operatorProductSummaryLabel){$operatorProductSummaryLabel.Text=('Потвърдени изделия: '+$confirmed+' | непотвърдени: '+$unconfirmed)}
        if($null -ne $operatorProcessStatusLabel){$operatorProcessStatusLabel.Text=$status.Text}
        Refresh-OperatorReports;Refresh-CsvRegistryGrid;Refresh-OperatorCatalogLabel;Refresh-MainOperatorUi
    }catch{}
}


function Get-MainOperatorFriendlyStatusText {
    param([string]$Text)
    $t=Safe-Trim $Text
    if([string]::IsNullOrWhiteSpace($t)){return 'Готово за работа.'}
    if($t -match 'Building one-pass PGMX filename index'){return 'Подготвяне на индекса на PGMX файловете...'}
    if($t -match 'Excel rows' -and $t -match 'products'){return 'Бърза миграция: обработване на Excel данните и връзките към изделията...'}
    if($t -match '(?i)building.*folder.*index|indexing.*folders'){return 'Обновяване на каталога на изделията...'}
    if($t -match '(?i)reading.*csv|csv.*reading'){return 'Четене на CSV файловете...'}
    if($t -match '(?i)pgmx.*index'){return 'Подготвяне на индекса на PGMX файловете...'}
    return $t
}

function Refresh-MainOperatorUi {
    param([switch]$StatusOnly)
    try {
        if($null -eq $script:OperatorMainForm){ return }
        if ($StatusOnly) {
            if($null -ne $script:MainStatusText){$script:MainStatusText.Text=(Get-MainOperatorFriendlyStatusText -Text $status.Text)}
            if($null -ne $script:MainCatalogText){
                if(Test-IsFolderIndexCurrentToday){$dt=(Get-Item -LiteralPath $script:IndexCsv).LastWriteTime;$script:MainCatalogText.Text='Каталог: обновен днес '+$dt.ToString('HH:mm')}
                else{$script:MainCatalogText.Text='Каталог: ще се обнови автоматично'}
            }
            if($null -ne $script:MainDoneText){$script:MainDoneText.Text=Get-OperatorDoneSummaryText}
            if($null -ne $script:MainResumeButton){$rs=Get-OperatorRunState;$script:MainResumeButton.Enabled=($null -ne $rs -and [string]$rs.Status -ne 'COMPLETED')}
            return
        }
        $products=@(Get-OperatorProductStateRows)
        $confirmed=@($products | Where-Object { $_.Confirmed }).Count
        $unconfirmed=$products.Count-$confirmed
        $csvCount=0
        try{if($null -ne $script:LastFileSummary){$csvCount=@($script:LastFileSummary.Rows).Count}}catch{}
        $detailCount=@($script:LastAllRows).Count
        $pgmxTotal=@($script:PgmxResultObjects).Count
        $pgmxOk=@($script:PgmxResultObjects | Where-Object { @('ГОТОВ','СЪЗДАДЕН ProgramCodeF','КОНВЕРТИРАН + ProgramCodeF') -contains (Safe-Trim $_.Result) }).Count

        if($null -ne $script:MainCsvValue){$script:MainCsvValue.Text=[string]$csvCount}
        if($null -ne $script:MainProductsValue){$script:MainProductsValue.Text=[string]$products.Count}
        if($null -ne $script:MainConfirmedValue){$script:MainConfirmedValue.Text=[string]$confirmed}
        if($null -ne $script:MainUnconfirmedValue){$script:MainUnconfirmedValue.Text=[string]$unconfirmed}
        if($null -ne $script:MainStatusText){$script:MainStatusText.Text=(Get-MainOperatorFriendlyStatusText -Text $status.Text)}
        if($null -ne $script:MainCatalogText){
            if(Test-IsFolderIndexCurrentToday){$dt=(Get-Item -LiteralPath $script:IndexCsv).LastWriteTime;$script:MainCatalogText.Text='Каталог: обновен днес '+$dt.ToString('HH:mm')}
            elseif([System.IO.File]::Exists($script:IndexCsv)){$script:MainCatalogText.Text='Каталог: ще се обнови автоматично'}
            else{$script:MainCatalogText.Text='Каталог: липсва - ще се създаде автоматично'}
        }
        if($null -ne $script:MainDoneText){$script:MainDoneText.Text=Get-OperatorDoneSummaryText}
        if($null -ne $script:MainResumeButton){$rs=Get-OperatorRunState;$script:MainResumeButton.Enabled=($null -ne $rs -and [string]$rs.Status -ne 'COMPLETED')}

        if($null -ne $script:MainStep1Status){
            if($detailCount -gt 0){$script:MainStep1Status.Text=('ГОТОВО  •  '+$csvCount+' CSV / '+$detailCount+' детайла')}else{$script:MainStep1Status.Text='ЧАКА'}
        }
        if($null -ne $script:MainStep2Status){
            if(@($script:SearchResultObjects).Count -gt 0){$script:MainStep2Status.Text=('ГОТОВО  •  '+$products.Count+' изделия')}else{$script:MainStep2Status.Text='ЧАКА'}
        }
        if($null -ne $script:MainStep3Status){
            if($products.Count -gt 0){$script:MainStep3Status.Text=($confirmed+' потвърдени / '+$unconfirmed+' неясни')}else{$script:MainStep3Status.Text='ЧАКА'}
        }
        if($null -ne $script:MainStep4Status){
            $mixOk=$false;try{$mixOk=Test-DailyMixReadyForDisplay}catch{}
            if($mixOk){$script:MainStep4Status.Text=('MIXX ГОТОВ  •  PGMX '+$pgmxOk+'/'+$pgmxTotal)}
            elseif($pgmxTotal -gt 0){$script:MainStep4Status.Text=('PGMX '+$pgmxOk+'/'+$pgmxTotal+'  •  MIXX чака')}
            else{$script:MainStep4Status.Text='ЧАКА'}
        }
        if($null -ne $script:MainStep5Status){
            $xlsOk=$false
            try{$daily=Get-DailyMixFolder;$xls=Join-Path $daily ('Mix('+(Get-Date -Format 'dd.MM')+').xls');$xlsOk=[System.IO.File]::Exists($xls)}catch{}
            if($xlsOk){$script:MainStep5Status.Text='ГОТОВО  •  XLS създаден'}else{$script:MainStep5Status.Text='ЧАКА'}
        }
        if($null -ne $script:MainReviewButton){
            $script:MainReviewButton.Text=if($unconfirmed -gt 0){'Прегледай непотвърдените ('+$unconfirmed+')'}else{'Няма непотвърдени'}
            $script:MainReviewButton.Enabled=($unconfirmed -gt 0)
        }
        if($null -ne $script:MainContinueButton){
            $script:MainContinueButton.Text='Продължи с потвърдените ('+$confirmed+')'
            $script:MainContinueButton.Enabled=($confirmed -gt 0)
        }
    } catch {}
}

function New-PWX500OperatorMainForm {
    $f=New-Object System.Windows.Forms.Form
    $f.Text='CAM Agent V5.1.38 - PWX 500 | P800 | FTT - Дневен производствен процес'
    $f.StartPosition='CenterScreen'
    $f.Size=New-Object System.Drawing.Size(1440,860)
    $f.MinimumSize=New-Object System.Drawing.Size(1180,720)
    $f.Font=New-Object System.Drawing.Font('Segoe UI',10)
    $f.BackColor=[System.Drawing.Color]::FromArgb(248,250,252)
    $f.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39)
    $f.AutoScaleMode=[System.Windows.Forms.AutoScaleMode]::Dpi
    try{if([System.IO.File]::Exists($script:AppIconIco)){$f.Icon=New-Object System.Drawing.Icon($script:AppIconIco)}}catch{}

    # Fixed two-row root layout: header can no longer overlap the 5 process cards.
    $root=New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock='Fill';$root.ColumnCount=1;$root.RowCount=2;$root.Margin=New-Object System.Windows.Forms.Padding(0);$root.Padding=New-Object System.Windows.Forms.Padding(0)
    $root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,96)))
    $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $f.Controls.Add($root)

    # LIGHT HEADER - no dark/black areas in operator mode
    $header=New-Object System.Windows.Forms.Panel
    $header.Dock='Fill';$header.Margin=New-Object System.Windows.Forms.Padding(0);$header.Height=96
    $header.BackColor=[System.Drawing.Color]::FromArgb(239,246,255)
    $header.BorderStyle='FixedSingle'
    $root.Controls.Add($header,0,0)

    $headerIcon=New-Object System.Windows.Forms.PictureBox
    $headerIcon.Size=New-Object System.Drawing.Size(70,70)
    $headerIcon.Location=New-Object System.Drawing.Point(14,12)
    $headerIcon.SizeMode='Zoom'
    $headerIcon.BackColor=[System.Drawing.Color]::Transparent
    try{if([System.IO.File]::Exists($script:AppIconPng)){$headerIcon.Image=[System.Drawing.Image]::FromFile($script:AppIconPng)}}catch{}
    $header.Controls.Add($headerIcon)

    $h1=New-Object System.Windows.Forms.Label
    $h1.Text='CAM Agent'
    $h1.ForeColor=[System.Drawing.Color]::FromArgb(30,64,175)
    $h1.Font=New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]::Bold)
    $h1.AutoSize=$true;$h1.Location=New-Object System.Drawing.Point(96,12)
    $header.Controls.Add($h1)

    $h2=New-Object System.Windows.Forms.Label
    $h2.Text='PWX 500 | P800 | FTT  •  Дневен производствен процес'
    $h2.ForeColor=[System.Drawing.Color]::FromArgb(55,65,81)
    $h2.Font=New-Object System.Drawing.Font('Segoe UI',10)
    $h2.AutoSize=$true;$h2.Location=New-Object System.Drawing.Point(99,52)
    $header.Controls.Add($h2)

    $mainTechnicalButton=New-Object System.Windows.Forms.Button
    $mainTechnicalButton.Text='Технически режим'
    $mainTechnicalButton.Size=New-Object System.Drawing.Size(158,36)
    $mainTechnicalButton.Anchor='Top,Right'
    $mainTechnicalButton.Location=New-Object System.Drawing.Point(1248,16)
    $mainTechnicalButton.BackColor=[System.Drawing.Color]::White
    $mainTechnicalButton.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39)
    $mainTechnicalButton.UseVisualStyleBackColor=$false
    $header.Controls.Add($mainTechnicalButton)

    $schedulePanel=New-Object System.Windows.Forms.Panel
    $schedulePanel.Anchor='Top,Right'
    $schedulePanel.Location=New-Object System.Drawing.Point(650,10)
    $schedulePanel.Size=New-Object System.Drawing.Size(585,72)
    $schedulePanel.BackColor=[System.Drawing.Color]::White
    $schedulePanel.BorderStyle='FixedSingle'
    $header.Controls.Add($schedulePanel)

    $mainScheduleLabel=New-Object System.Windows.Forms.Label
    $mainScheduleLabel.Text='Автоматичен старт'
    $mainScheduleLabel.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39)
    $mainScheduleLabel.Font=New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    $mainScheduleLabel.AutoSize=$true;$mainScheduleLabel.Location=New-Object System.Drawing.Point(10,9)
    $schedulePanel.Controls.Add($mainScheduleLabel)

    $script:MainScheduleCheck=New-Object System.Windows.Forms.CheckBox
    $script:MainScheduleCheck.Text='Вкл.'
    $script:MainScheduleCheck.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39)
    $script:MainScheduleCheck.AutoSize=$true;$script:MainScheduleCheck.Location=New-Object System.Drawing.Point(12,38)
    $schedulePanel.Controls.Add($script:MainScheduleCheck)

    $script:MainScheduleTime=New-Object System.Windows.Forms.DateTimePicker
    $script:MainScheduleTime.Format=[System.Windows.Forms.DateTimePickerFormat]::Custom
    $script:MainScheduleTime.CustomFormat='HH:mm';$script:MainScheduleTime.ShowUpDown=$true
    $script:MainScheduleTime.Size=New-Object System.Drawing.Size(76,28)
    $script:MainScheduleTime.Location=New-Object System.Drawing.Point(78,33)
    $schedulePanel.Controls.Add($script:MainScheduleTime)

    $mainScheduleSave=New-Object System.Windows.Forms.Button
    $mainScheduleSave.Text='Запази';$mainScheduleSave.Size=New-Object System.Drawing.Size(82,30)
    $mainScheduleSave.Location=New-Object System.Drawing.Point(164,32)
    $mainScheduleSave.BackColor=[System.Drawing.Color]::FromArgb(219,234,254)
    $mainScheduleSave.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39)
    $mainScheduleSave.UseVisualStyleBackColor=$false
    $schedulePanel.Controls.Add($mainScheduleSave)

    $mainScheduleModeLabel=New-Object System.Windows.Forms.Label
    $mainScheduleModeLabel.Text='Режим';$mainScheduleModeLabel.AutoSize=$true
    $mainScheduleModeLabel.Location=New-Object System.Drawing.Point(258,9);$schedulePanel.Controls.Add($mainScheduleModeLabel)
    $script:MainScheduleModeCombo=New-Object System.Windows.Forms.ComboBox
    $script:MainScheduleModeCombo.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
    $script:MainScheduleModeCombo.Items.Add('Само PGMX (без MIXX)')|Out-Null
    $script:MainScheduleModeCombo.Items.Add('Целият процес 1–5')|Out-Null
    $script:MainScheduleModeCombo.SelectedIndex=0
    $script:MainScheduleModeCombo.Location=New-Object System.Drawing.Point(258,33)
    $script:MainScheduleModeCombo.Size=New-Object System.Drawing.Size(178,28)
    $schedulePanel.Controls.Add($script:MainScheduleModeCombo)

    $script:MainCatalogText=New-Object System.Windows.Forms.Label
    $script:MainCatalogText.Text='Каталог: проверка...'
    $script:MainCatalogText.ForeColor=[System.Drawing.Color]::FromArgb(30,64,175)
    $script:MainCatalogText.AutoSize=$false
    $script:MainCatalogText.Size=New-Object System.Drawing.Size(135,48)
    $script:MainCatalogText.TextAlign='MiddleLeft'
    $script:MainCatalogText.Location=New-Object System.Drawing.Point(446,17)
    $schedulePanel.Controls.Add($script:MainCatalogText)

    $body=New-Object System.Windows.Forms.Panel
    $body.Dock='Fill';$body.Padding=New-Object System.Windows.Forms.Padding(16)
    $body.BackColor=[System.Drawing.Color]::FromArgb(248,250,252)
    $root.Controls.Add($body,0,1)

    # One vertical layout fixes the clipped/hidden step cards seen in V5.1.31.
    $layout=New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock='Fill';$layout.ColumnCount=1;$layout.RowCount=6
    $layout.BackColor=$body.BackColor
    $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,116)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,166)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,92)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,78)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,92)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $body.Controls.Add($layout)

    $steps=New-Object System.Windows.Forms.TableLayoutPanel
    $steps.Dock='Fill';$steps.ColumnCount=5;$steps.RowCount=1;$steps.Padding=New-Object System.Windows.Forms.Padding(0,5,0,5)
    $steps.BackColor=$body.BackColor
    1..5 | ForEach-Object {$steps.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,20)))}
    $layout.Controls.Add($steps,0,0)

    function Add-StepCard([int]$i,[string]$title){
        $p=New-Object System.Windows.Forms.Panel
        $p.Margin=New-Object System.Windows.Forms.Padding(5)
        $p.Dock='Fill';$p.BackColor=[System.Drawing.Color]::White;$p.BorderStyle='FixedSingle'
        $n=New-Object System.Windows.Forms.Label
        $n.Text=[string]$i;$n.Font=New-Object System.Drawing.Font('Segoe UI',15,[System.Drawing.FontStyle]::Bold)
        $n.ForeColor=[System.Drawing.Color]::FromArgb(37,99,235);$n.AutoSize=$true;$n.Location=New-Object System.Drawing.Point(12,10);$p.Controls.Add($n)
        $t=New-Object System.Windows.Forms.Label
        $t.Text=$title;$t.Font=New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
        $t.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39);$t.AutoSize=$true;$t.Location=New-Object System.Drawing.Point(43,13);$p.Controls.Add($t)
        $s=New-Object System.Windows.Forms.Label
        $s.Text='ЧАКА';$s.ForeColor=[System.Drawing.Color]::FromArgb(75,85,99)
        $s.BackColor=[System.Drawing.Color]::FromArgb(249,250,251)
        $s.AutoSize=$false;$s.Size=New-Object System.Drawing.Size(235,42);$s.Location=New-Object System.Drawing.Point(12,50)
        $s.TextAlign='MiddleLeft';$p.Controls.Add($s)
        $steps.Controls.Add($p,$i-1,0)
        return $s
    }
    $script:MainStep1Status=Add-StepCard 1 'CSV'
    $script:MainStep2Status=Add-StepCard 2 'Папки + HOP'
    $script:MainStep3Status=Add-StepCard 3 'Потвърждение'
    $script:MainStep4Status=Add-StepCard 4 'PGMX + MIXX'
    $script:MainStep5Status=Add-StepCard 5 'PowerSet + XLS'

    $actions=New-Object System.Windows.Forms.Panel
    $actions.Dock='Fill';$actions.Padding=New-Object System.Windows.Forms.Padding(4)
    $actions.BackColor=[System.Drawing.Color]::White;$actions.BorderStyle='FixedSingle'
    $layout.Controls.Add($actions,0,1)

    $mainRunAllButton=New-Object System.Windows.Forms.Button
    $mainRunAllButton.Text='ЦЕЛИЯТ ПРОЦЕС 1 → 5'
    $mainRunAllButton.Font=New-Object System.Drawing.Font('Segoe UI',12,[System.Drawing.FontStyle]::Bold)
    $mainRunAllButton.BackColor=[System.Drawing.Color]::FromArgb(34,197,94);$mainRunAllButton.ForeColor=[System.Drawing.Color]::Black
    $mainRunAllButton.UseVisualStyleBackColor=$false;$mainRunAllButton.FlatStyle='Flat'
    $mainRunAllButton.Size=New-Object System.Drawing.Size(330,52);$mainRunAllButton.Location=New-Object System.Drawing.Point(8,10);$actions.Controls.Add($mainRunAllButton)

    $mainStopButton=New-Object System.Windows.Forms.Button;$mainStopButton.Text='СПРИ ПРОЦЕСА';$mainStopButton.Font=New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold);$mainStopButton.Size=New-Object System.Drawing.Size(160,42);$mainStopButton.Location=New-Object System.Drawing.Point(348,15);$mainStopButton.BackColor=[System.Drawing.Color]::FromArgb(239,68,68);$mainStopButton.ForeColor=[System.Drawing.Color]::White;$mainStopButton.UseVisualStyleBackColor=$false;$actions.Controls.Add($mainStopButton)
    $mainReadCsvButton=New-Object System.Windows.Forms.Button;$mainReadCsvButton.Text='1. Прочети CSV';$mainReadCsvButton.Size=New-Object System.Drawing.Size(140,42);$mainReadCsvButton.Location=New-Object System.Drawing.Point(518,15);$mainReadCsvButton.BackColor=[System.Drawing.Color]::FromArgb(243,244,246);$mainReadCsvButton.ForeColor=[System.Drawing.Color]::Black;$mainReadCsvButton.UseVisualStyleBackColor=$false;$actions.Controls.Add($mainReadCsvButton)
    $mainValidateButton=New-Object System.Windows.Forms.Button;$mainValidateButton.Text='2. Провери изделия';$mainValidateButton.Size=New-Object System.Drawing.Size(160,42);$mainValidateButton.Location=New-Object System.Drawing.Point(668,15);$mainValidateButton.BackColor=[System.Drawing.Color]::FromArgb(243,244,246);$mainValidateButton.ForeColor=[System.Drawing.Color]::Black;$mainValidateButton.UseVisualStyleBackColor=$false;$actions.Controls.Add($mainValidateButton)
    $script:MainReviewButton=New-Object System.Windows.Forms.Button;$script:MainReviewButton.Text='Прегледай непотвърдените';$script:MainReviewButton.BackColor=[System.Drawing.Color]::FromArgb(254,249,195);$script:MainReviewButton.ForeColor=[System.Drawing.Color]::Black;$script:MainReviewButton.UseVisualStyleBackColor=$false;$script:MainReviewButton.Size=New-Object System.Drawing.Size(210,42);$script:MainReviewButton.Location=New-Object System.Drawing.Point(838,15);$actions.Controls.Add($script:MainReviewButton)
    $script:MainContinueButton=New-Object System.Windows.Forms.Button;$script:MainContinueButton.Text='Продължи с потвърдените';$script:MainContinueButton.BackColor=[System.Drawing.Color]::FromArgb(187,247,208);$script:MainContinueButton.ForeColor=[System.Drawing.Color]::Black;$script:MainContinueButton.UseVisualStyleBackColor=$false;$script:MainContinueButton.Size=New-Object System.Drawing.Size(235,42);$script:MainContinueButton.Location=New-Object System.Drawing.Point(1058,15);$actions.Controls.Add($script:MainContinueButton)
    $mainOpenDailyButton=New-Object System.Windows.Forms.Button;$mainOpenDailyButton.Text='Отвори дневната папка';$mainOpenDailyButton.Size=New-Object System.Drawing.Size(195,34);$mainOpenDailyButton.Location=New-Object System.Drawing.Point(8,75);$mainOpenDailyButton.BackColor=[System.Drawing.Color]::FromArgb(239,246,255);$mainOpenDailyButton.ForeColor=[System.Drawing.Color]::Black;$mainOpenDailyButton.UseVisualStyleBackColor=$false;$actions.Controls.Add($mainOpenDailyButton)
    $script:MainDailyRootButton=New-Object System.Windows.Forms.Button;$script:MainDailyRootButton.Text='Основна папка...';$script:MainDailyRootButton.Size=New-Object System.Drawing.Size(125,34);$script:MainDailyRootButton.Location=New-Object System.Drawing.Point(212,75);$script:MainDailyRootButton.BackColor=[System.Drawing.Color]::FromArgb(224,231,255);$script:MainDailyRootButton.ForeColor=[System.Drawing.Color]::Black;$script:MainDailyRootButton.UseVisualStyleBackColor=$false;$actions.Controls.Add($script:MainDailyRootButton)
    $script:MainDailyRootToolTip=New-Object System.Windows.Forms.ToolTip;$script:MainDailyRootToolTip.SetToolTip($script:MainDailyRootButton,(Get-DailyRootFolder))
    $mainMigrationButton=New-Object System.Windows.Forms.Button;$mainMigrationButton.Text='Бърза миграция от Excel';$mainMigrationButton.Size=New-Object System.Drawing.Size(220,34);$mainMigrationButton.Location=New-Object System.Drawing.Point(346,75);$mainMigrationButton.BackColor=[System.Drawing.Color]::FromArgb(243,244,246);$mainMigrationButton.ForeColor=[System.Drawing.Color]::Black;$mainMigrationButton.UseVisualStyleBackColor=$false;$actions.Controls.Add($mainMigrationButton)
    $script:MainMigrationButton=$mainMigrationButton
    $mainResumeButton=New-Object System.Windows.Forms.Button;$mainResumeButton.Text='Продължи прекъснатия процес';$mainResumeButton.Size=New-Object System.Drawing.Size(245,34);$mainResumeButton.Location=New-Object System.Drawing.Point(576,75);$mainResumeButton.BackColor=[System.Drawing.Color]::FromArgb(254,240,138);$mainResumeButton.ForeColor=[System.Drawing.Color]::Black;$mainResumeButton.UseVisualStyleBackColor=$false;$actions.Controls.Add($mainResumeButton);$script:MainResumeButton=$mainResumeButton
    $mainAddTimesButton=New-Object System.Windows.Forms.Button;$mainAddTimesButton.Text='Добави файлове с времена';$mainAddTimesButton.Size=New-Object System.Drawing.Size(180,34);$mainAddTimesButton.Location=New-Object System.Drawing.Point(830,75);$actions.Controls.Add($mainAddTimesButton)
    $mainStopTimesButton=New-Object System.Windows.Forms.Button;$mainStopTimesButton.Text='Спри импорт';$mainStopTimesButton.Size=New-Object System.Drawing.Size(115,34);$mainStopTimesButton.Location=New-Object System.Drawing.Point(1018,75);$mainStopTimesButton.BackColor=[System.Drawing.Color]::FromArgb(254,226,226);$mainStopTimesButton.ForeColor=[System.Drawing.Color]::FromArgb(153,27,27);$mainStopTimesButton.UseVisualStyleBackColor=$false;$mainStopTimesButton.Enabled=$false;$actions.Controls.Add($mainStopTimesButton);$script:MainStopTimesButton=$mainStopTimesButton
    $mainExportTimesButton=New-Object System.Windows.Forms.Button;$mainExportTimesButton.Text='Отчет времена';$mainExportTimesButton.Size=New-Object System.Drawing.Size(120,34);$mainExportTimesButton.Location=New-Object System.Drawing.Point(1141,75);$actions.Controls.Add($mainExportTimesButton)
    $mainExportFilesButton=New-Object System.Windows.Forms.Button;$mainExportFilesButton.Text='Регистър файлове';$mainExportFilesButton.Size=New-Object System.Drawing.Size(135,34);$mainExportFilesButton.Location=New-Object System.Drawing.Point(1269,75);$actions.Controls.Add($mainExportFilesButton)
    $mainRunPgmxOnlyButton=New-Object System.Windows.Forms.Button
    $mainRunPgmxOnlyButton.Text='1. ПОДГОТОВКА + PGMX'
    $mainRunPgmxOnlyButton.Location=New-Object System.Drawing.Point(8,115)
    $mainRunPgmxOnlyButton.Size=New-Object System.Drawing.Size(290,39)
    $mainRunPgmxOnlyButton.BackColor=[System.Drawing.Color]::FromArgb(191,219,254)
    $mainRunPgmxOnlyButton.ForeColor=[System.Drawing.Color]::Black
    $actions.Controls.Add($mainRunPgmxOnlyButton)
    $mainRunBatchButton=New-Object System.Windows.Forms.Button
    $mainRunBatchButton.Text='2. ОБЩ MIXX + PowerSet + XLS'
    $mainRunBatchButton.Location=New-Object System.Drawing.Point(308,115)
    $mainRunBatchButton.Size=New-Object System.Drawing.Size(340,39)
    $mainRunBatchButton.BackColor=[System.Drawing.Color]::FromArgb(187,247,208)
    $mainRunBatchButton.ForeColor=[System.Drawing.Color]::Black
    $actions.Controls.Add($mainRunBatchButton)
    $mainBatchInfo=New-Object System.Windows.Forms.Label
    $mainBatchInfo.Text='Етап 2 събира чакащите PGMX от дневните папки. Не се стартира сам в режим 1.'
    $mainBatchInfo.AutoSize=$true;$mainBatchInfo.Location=New-Object System.Drawing.Point(660,125)
    $actions.Controls.Add($mainBatchInfo)

    $cards=New-Object System.Windows.Forms.TableLayoutPanel
    $cards.Dock='Fill';$cards.ColumnCount=5;$cards.RowCount=1;$cards.BackColor=$body.BackColor
    1..5 | ForEach-Object {$cards.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,20)))}
    $layout.Controls.Add($cards,0,2)
    function Add-SummaryCard([int]$idx,[string]$caption){
        $p=New-Object System.Windows.Forms.Panel;$p.Dock='Fill';$p.Margin=New-Object System.Windows.Forms.Padding(5);$p.BackColor=[System.Drawing.Color]::White;$p.BorderStyle='FixedSingle'
        $v=New-Object System.Windows.Forms.Label;$v.Text='0';$v.ForeColor=[System.Drawing.Color]::Black;$v.Font=New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold);$v.AutoSize=$true;$v.Location=New-Object System.Drawing.Point(12,8);$p.Controls.Add($v)
        $c=New-Object System.Windows.Forms.Label;$c.Text=$caption;$c.ForeColor=[System.Drawing.Color]::FromArgb(75,85,99);$c.AutoSize=$true;$c.Location=New-Object System.Drawing.Point(13,48);$p.Controls.Add($c)
        $cards.Controls.Add($p,$idx,0)
        return $v
    }
    $script:MainCsvValue=Add-SummaryCard 0 'CSV в цикъла'
    $script:MainProductsValue=Add-SummaryCard 1 'Изделия'
    $script:MainConfirmedValue=Add-SummaryCard 2 'Потвърдени'
    $script:MainUnconfirmedValue=Add-SummaryCard 3 'Непотвърдени'
    $script:MainErrorsValue=Add-SummaryCard 4 'Грешки детайли'

    $statusBox=New-Object System.Windows.Forms.GroupBox
    $statusBox.Text='Текущ статус';$statusBox.Dock='Fill';$statusBox.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39);$statusBox.BackColor=[System.Drawing.Color]::White
    $layout.Controls.Add($statusBox,0,3)
    $script:MainStatusText=New-Object System.Windows.Forms.Label
    $script:MainStatusText.Text='Готово за работа.';$script:MainStatusText.Dock='Fill';$script:MainStatusText.Padding=New-Object System.Windows.Forms.Padding(10,8,10,5)
    $script:MainStatusText.Font=New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold);$script:MainStatusText.ForeColor=[System.Drawing.Color]::Black;$script:MainStatusText.BackColor=[System.Drawing.Color]::White
    $statusBox.Controls.Add($script:MainStatusText)

    $doneBox=New-Object System.Windows.Forms.GroupBox
    $doneBox.Text='Какво е направено / последно състояние';$doneBox.Dock='Fill';$doneBox.ForeColor=[System.Drawing.Color]::FromArgb(17,24,39);$doneBox.BackColor=[System.Drawing.Color]::White
    $layout.Controls.Add($doneBox,0,4)
    $script:MainDoneText=New-Object System.Windows.Forms.Label
    $script:MainDoneText.Text='Няма записана история.';$script:MainDoneText.Dock='Fill';$script:MainDoneText.Padding=New-Object System.Windows.Forms.Padding(10,7,10,5);$script:MainDoneText.AutoEllipsis=$false;$script:MainDoneText.AutoSize=$false
    $script:MainDoneText.ForeColor=[System.Drawing.Color]::Black;$script:MainDoneText.BackColor=[System.Drawing.Color]::White
    $doneBox.Controls.Add($script:MainDoneText)

    $mainReportTabs=New-Object System.Windows.Forms.TabControl
    $mainReportTabs.Dock='Fill';$mainReportTabs.BackColor=[System.Drawing.Color]::White
    $layout.Controls.Add($mainReportTabs,0,5)
    $t1=New-Object System.Windows.Forms.TabPage;$t1.Text='Непотвърдени / пропуснати';$t1.BackColor=[System.Drawing.Color]::White;$t1.ForeColor=[System.Drawing.Color]::Black;$mainReportTabs.TabPages.Add($t1)
    $t2=New-Object System.Windows.Forms.TabPage;$t2.Text='Детайли с грешка';$t2.BackColor=[System.Drawing.Color]::White;$t2.ForeColor=[System.Drawing.Color]::Black;$mainReportTabs.TabPages.Add($t2)
    $t3=New-Object System.Windows.Forms.TabPage;$t3.Text='История CSV';$t3.BackColor=[System.Drawing.Color]::White;$t3.ForeColor=[System.Drawing.Color]::Black;$mainReportTabs.TabPages.Add($t3)

    function Style-OperatorGrid($g){
        $g.Dock='Fill';$g.ReadOnly=$true;$g.AllowUserToAddRows=$false;$g.AllowUserToDeleteRows=$false;$g.SelectionMode='FullRowSelect'
        $g.BackgroundColor=[System.Drawing.Color]::White;$g.GridColor=[System.Drawing.Color]::FromArgb(209,213,219)
        $g.DefaultCellStyle.BackColor=[System.Drawing.Color]::White;$g.DefaultCellStyle.ForeColor=[System.Drawing.Color]::Black
        $g.DefaultCellStyle.SelectionBackColor=[System.Drawing.Color]::FromArgb(219,234,254);$g.DefaultCellStyle.SelectionForeColor=[System.Drawing.Color]::Black
        $g.ColumnHeadersDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(243,244,246);$g.ColumnHeadersDefaultCellStyle.ForeColor=[System.Drawing.Color]::Black
        $g.EnableHeadersVisualStyles=$false
    }
    $mainSkippedLayout=New-Object System.Windows.Forms.TableLayoutPanel
    $mainSkippedLayout.Dock='Fill';$mainSkippedLayout.ColumnCount=1;$mainSkippedLayout.RowCount=2;$mainSkippedLayout.Margin=New-Object System.Windows.Forms.Padding(0);$mainSkippedLayout.Padding=New-Object System.Windows.Forms.Padding(0)
    $mainSkippedLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,46)))
    $mainSkippedLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $t1.Controls.Add($mainSkippedLayout)
    $mainSkippedToolbar=New-Object System.Windows.Forms.Panel;$mainSkippedToolbar.Dock='Fill';$mainSkippedToolbar.BackColor=[System.Drawing.Color]::FromArgb(248,250,252);$mainSkippedLayout.Controls.Add($mainSkippedToolbar,0,0)
    $mainSetFolderButton=New-Object System.Windows.Forms.Button;$mainSetFolderButton.Text='Посочи папка за избраното';$mainSetFolderButton.Size=New-Object System.Drawing.Size(220,32);$mainSetFolderButton.Location=New-Object System.Drawing.Point(8,7);$mainSetFolderButton.BackColor=[System.Drawing.Color]::FromArgb(219,234,254);$mainSetFolderButton.ForeColor=[System.Drawing.Color]::Black;$mainSetFolderButton.UseVisualStyleBackColor=$false;$mainSkippedToolbar.Controls.Add($mainSetFolderButton)
    $mainRunHopButton=New-Object System.Windows.Forms.Button;$mainRunHopButton.Text='Провери HOP на зададените';$mainRunHopButton.Size=New-Object System.Drawing.Size(220,32);$mainRunHopButton.Location=New-Object System.Drawing.Point(238,7);$mainRunHopButton.BackColor=[System.Drawing.Color]::FromArgb(220,252,231);$mainRunHopButton.ForeColor=[System.Drawing.Color]::Black;$mainRunHopButton.UseVisualStyleBackColor=$false;$mainSkippedToolbar.Controls.Add($mainRunHopButton)
    $mainSkippedHint=New-Object System.Windows.Forms.Label;$mainSkippedHint.Text='Двойно щракване върху изделие = Посочи папка';$mainSkippedHint.AutoSize=$true;$mainSkippedHint.ForeColor=[System.Drawing.Color]::FromArgb(75,85,99);$mainSkippedHint.Location=New-Object System.Drawing.Point(475,14);$mainSkippedToolbar.Controls.Add($mainSkippedHint)
    $script:MainSkippedGrid=New-Object System.Windows.Forms.DataGridView;Style-OperatorGrid $script:MainSkippedGrid;$script:MainSkippedGrid.AutoSizeColumnsMode='Fill';$mainSkippedLayout.Controls.Add($script:MainSkippedGrid,0,1)
    $mainSetFolderButton.Add_Click({try{Set-ManualProductFolderFromOperatorGrid}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
    $script:MainSkippedGrid.Add_CellDoubleClick({param($sender,$e);if($e.RowIndex -ge 0){try{Set-ManualProductFolderFromOperatorGrid}catch{}}})
    $mainRunHopButton.Add_Click({try{Start-BackgroundHopCheck}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
    $script:MainErrorsGrid=New-Object System.Windows.Forms.DataGridView;Style-OperatorGrid $script:MainErrorsGrid;$script:MainErrorsGrid.AutoSizeColumnsMode='DisplayedCells';$t2.Controls.Add($script:MainErrorsGrid)
    $script:MainCsvRegistryGrid=New-Object System.Windows.Forms.DataGridView;Style-OperatorGrid $script:MainCsvRegistryGrid;$script:MainCsvRegistryGrid.AutoSizeColumnsMode='DisplayedCells';$t3.Controls.Add($script:MainCsvRegistryGrid)

    $mainRunPgmxOnlyButton.Add_Click({try{Start-OperatorFullFlow -Mode 'PGMX_ONLY'}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - PGMX')|Out-Null}})
    $mainRunBatchButton.Add_Click({try{Invoke-CumulativeMixPowerSet}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - Общ MIXX/PowerSet')|Out-Null}})
    $mainRunAllButton.Add_Click({try{Start-OperatorFullFlow -Mode 'FULL'}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
    $mainStopButton.Add_Click({try{Stop-OperatorProcessSafely}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent')|Out-Null}})
    $mainResumeButton.Add_Click({try{Invoke-OperatorResumeInterruptedRun}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent')|Out-Null}})
    $mainReadCsvButton.Add_Click({try{Start-BackgroundCsvRead}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
    $mainValidateButton.Add_Click({try{Start-OperatorValidation}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
    $script:MainReviewButton.Add_Click({try{Refresh-ConfirmationQueue;$tabs.SelectedTab=$tabConfirm;$form.Show();$form.WindowState='Normal';$form.BringToFront()}catch{}})
    $script:MainContinueButton.Add_Click({try{Invoke-OperatorContinueConfirmed}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
    $mainOpenDailyButton.Add_Click({try{$d=Get-DailyMixFolder;if([System.IO.Directory]::Exists($d)){Start-Process explorer.exe -ArgumentList ('"'+$d+'"')}else{[System.Windows.Forms.MessageBox]::Show('Още няма дневна папка.','PWX500 Agent')|Out-Null}}catch{}})
    $script:MainDailyRootButton.Add_Click({try{[void](Select-DailyRootFolderInteractive -Owner $script:OperatorMainForm)}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - Основна папка')|Out-Null}})
    $mainMigrationButton.Add_Click({
        try {
            # V5.1.31: direct migration from the operator window; do not PerformClick() a hidden-form button.
            if($null -ne $script:HistoryWorkerProcess -and -not $script:HistoryWorkerProcess.HasExited){
                [System.Windows.Forms.MessageBox]::Show('Бързата миграция вече се изпълнява.','PWX500 Agent - Бърза миграция',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
                return
            }
            if([System.IO.File]::Exists($script:HistoryMarkerFile)){
                $ans=[System.Windows.Forms.MessageBox]::Show('Началната историческа миграция вече е изпълнявана. Да се повтори ли проверката? Съществуващите потвърдени връзки няма да бъдат презаписани.','PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
                if($ans -ne [System.Windows.Forms.DialogResult]::Yes){return}
            }
            $dlg=New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title='Избери „Поръчки за линия.xlsx“'
            $dlg.Filter='Excel (*.xlsx)|*.xlsx'
            $dlg.Multiselect=$false
            if($dlg.ShowDialog($script:OperatorMainForm) -eq [System.Windows.Forms.DialogResult]::OK){
                $wb=[string]$dlg.FileName
                $dlg.Dispose()
                $mainMigrationButton.Enabled=$false
                $status.Text='Бърза миграция: стартиране...'
                Refresh-MainOperatorUi
                Start-BackgroundHistoryMigration -WorkbookPath $wb
            } else {
                $dlg.Dispose()
            }
        } catch {
            $mainMigrationButton.Enabled=$true
            [System.Windows.Forms.MessageBox]::Show("Грешка при началната историческа миграция:`r`n$($_.Exception.Message)",'PWX500 Agent - Бърза миграция',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
        }
    })
    $mainAddTimesButton.Add_Click({try{Import-TimeFilesInteractive;Refresh-MainOperatorUi}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - времена')|Out-Null}})
    $mainStopTimesButton.Add_Click({try{Stop-TimeImportBackground}catch{}})
    $mainExportTimesButton.Add_Click({try{Export-TimeReport|Out-Null}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - отчет времена')|Out-Null}})
    $mainExportFilesButton.Add_Click({try{Export-AllFilesRegistry|Out-Null}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - регистър файлове')|Out-Null}})
    $mainTechnicalButton.Add_Click({try{$form.Show();$form.WindowState='Normal';$form.BringToFront()}catch{}})
    $mainScheduleSave.Add_Click({try{$operatorScheduleCheck.Checked=$script:MainScheduleCheck.Checked;$operatorScheduleTime.Value=$script:MainScheduleTime.Value;Save-OperatorScheduleSettings;$status.Text='Записан автоматичен начален час: '+$script:MainScheduleTime.Value.ToString('HH:mm');Refresh-MainOperatorUi}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent')|Out-Null}})

    $script:OperatorMainForm=$f
    return $f
}

function Start-OperatorValidation {
    $script:OperatorRunHopAfterSearch=$true
    $script:OperatorRunExecuteAfterHop=$false
    if(Ensure-OperatorDailyCatalog -AfterAction 'Search'){Start-BackgroundProductSearch}
}

function Invoke-OperatorContinueConfirmed {
    param([switch]$Silent)
    $states=@(Get-OperatorProductStateRows);$confirmed=@($states|Where-Object{$_.Confirmed});$unconfirmed=@($states|Where-Object{-not $_.Confirmed})
    if($confirmed.Count -eq 0){
        try{Finalize-CsvRegistryForCurrentBatch -PipelineSucceeded $true -PipelineMessage 'Няма потвърдени изделия; текущите CSV остават за повторен опит.'}catch{}
        $script:OperatorAutoFlow=$false;$script:OperatorRunExecuteAfterHop=$false
        if(-not $Silent){[System.Windows.Forms.MessageBox]::Show('Няма потвърдени изделия за обработка. CSV файловете остават ЧАСТИЧНО/за повторен опит.','PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null}
        Refresh-OperatorDashboard
        return
    }
    if(-not $Silent){
        $msg="Потвърдени изделия: $($confirmed.Count)`r`nНепотвърдени изделия: $($unconfirmed.Count)`r`n`r`nНепотвърдените НЯМА да спрат процеса - ще бъдат пропуснати и ще влязат в крайния отчет.`r`n`r`nДа продължа ли с потвърдените?"
        $ans=[System.Windows.Forms.MessageBox]::Show($msg,'PWX500 Agent - Продължи с потвърдените',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
        if($ans -ne [System.Windows.Forms.DialogResult]::Yes){return}
    }
    Start-BackgroundPgmx -Mode 'Execute'
}

function Start-OperatorFullFlow {
    param([switch]$ResumeExisting,[ValidateSet('FULL','PGMX_ONLY')][string]$Mode='FULL')
    $script:OperatorExecutionMode=$Mode
    if($script:TrainingActive){throw 'Спри активното обучение преди автоматичен операторски процес.'}
    if(-not $ResumeExisting){Start-NewOperatorRunState -Message 'Стартиран дневен производствен процес.'}else{$script:OperatorStopRequested=$false}
    $script:OperatorAutoFlow=$true
    $script:OperatorRunHopAfterSearch=$true
    $script:OperatorRunExecuteAfterHop=$true
    $status.Text='ОПЕРАТОР AUTO: проверявам дневния каталог и започвам Стъпка 1 - CSV.'
    if(Ensure-OperatorDailyCatalog -AfterAction 'Csv'){Start-BackgroundCsvRead}
}

function Continue-OperatorAfterCsv {
    if(-not $script:OperatorAutoFlow){return}
    if(@($script:LastAllRows).Count -eq 0){$script:OperatorAutoFlow=$false;$script:OperatorRunHopAfterSearch=$false;$script:OperatorRunExecuteAfterHop=$false;$status.Text='ОПЕРАТОР AUTO: няма разчетени редове в текущите CSV. Провери История CSV.';Refresh-OperatorDashboard;return}
    try{Start-BackgroundProductSearch}catch{$script:OperatorAutoFlow=$false;Set-CurrentCsvBatchRegistryError -Message ('Грешка при търсене: '+$_.Exception.Message);$status.Text='ОПЕРАТОР AUTO грешка при търсене: '+$_.Exception.Message;Set-OperatorRunError -Message $status.Text}
}

function Continue-OperatorAfterSearch {
    if(-not $script:OperatorRunHopAfterSearch){return}
    $script:OperatorRunHopAfterSearch=$false
    $pending=Get-OperatorPendingHopCount
    if($pending -gt 0){
        try{Start-BackgroundHopCheck}catch{$script:OperatorAutoFlow=$false;$script:OperatorRunExecuteAfterHop=$false;$status.Text='ОПЕРАТОР грешка при HOP: '+$_.Exception.Message}
    } else {
        Refresh-OperatorDashboard
        if($script:OperatorRunExecuteAfterHop){$script:OperatorRunExecuteAfterHop=$false;try{Invoke-OperatorContinueConfirmed -Silent}catch{$script:OperatorAutoFlow=$false;$status.Text='ОПЕРАТОР AUTO грешка: '+$_.Exception.Message}}
    }
}

function Continue-OperatorAfterHop {
    Refresh-OperatorDashboard
    if($script:OperatorRunExecuteAfterHop){
        $script:OperatorRunExecuteAfterHop=$false
        try{Invoke-OperatorContinueConfirmed -Silent}catch{$script:OperatorAutoFlow=$false;$status.Text='ОПЕРАТОР AUTO грешка: '+$_.Exception.Message}
    }
}

function Save-OperatorScheduleSettings {
    if($null -eq $operatorScheduleCheck -or $null -eq $operatorScheduleTime){return}
    $enabledText = if($operatorScheduleCheck.Checked){'1'}else{'0'}
    $lastRunText = [string]$script:OperatorScheduleLastRunDate
    if($lastRunText -notmatch '^\d{4}-\d{2}-\d{2}$'){$lastRunText=''}
    if($null -ne $script:MainScheduleModeCombo -and $script:MainScheduleModeCombo.SelectedIndex -ge 0){
        $script:OperatorScheduleMode=if($script:MainScheduleModeCombo.SelectedIndex -eq 1){'FULL'}else{'PGMX_ONLY'}
    }
    $lines=@(('Enabled='+$enabledText),('Time='+$operatorScheduleTime.Value.ToString('HH:mm')),('LastRunDate='+$lastRunText),('Mode='+$script:OperatorScheduleMode))
    [System.IO.File]::WriteAllLines($script:OperatorScheduleFile,$lines,(New-Object System.Text.UTF8Encoding($true)))
}

function Load-OperatorScheduleSettings {
    if($null -eq $operatorScheduleCheck -or $null -eq $operatorScheduleTime){return}
    $enabled=$false;$time='22:00';$lastRunDate='';$mode='PGMX_ONLY'
    try{
        if([System.IO.File]::Exists($script:OperatorScheduleFile)){
            foreach($line in [System.IO.File]::ReadAllLines($script:OperatorScheduleFile,[System.Text.Encoding]::UTF8)){
                if($line -match '^Enabled=(.*)$'){$enabled=($matches[1].Trim() -eq '1')}
                elseif($line -match '^Time=(\d{1,2}:\d{2})$'){$time=$matches[1]}
                elseif($line -match '^LastRunDate=(\d{4}-\d{2}-\d{2})$'){$lastRunDate=$matches[1]}
                elseif($line -match '^Mode=(PGMX_ONLY|FULL)$'){$mode=$matches[1]}
            }
        }
        $script:OperatorScheduleLastRunDate=$lastRunDate
        $script:OperatorScheduleMode=$mode
        if($null -ne $script:MainScheduleModeCombo){$script:MainScheduleModeCombo.SelectedIndex=if($mode -eq 'FULL'){1}else{0}}
        $dt=[datetime]::Today.Add([timespan]::Parse($time));$operatorScheduleTime.Value=$dt;$operatorScheduleCheck.Checked=$enabled
    }catch{}
}

function Test-OperatorScheduleTick {
    if($null -eq $operatorScheduleCheck -or -not $operatorScheduleCheck.Checked){return}
    $today=Get-Date -Format 'yyyy-MM-dd'
    if($script:OperatorScheduleLastRunDate -eq $today){return}
    $now=Get-Date;$target=$operatorScheduleTime.Value.TimeOfDay
    if($now.TimeOfDay -lt $target){return}
    $busy=($null -ne $script:CsvWorkerProcess -or $null -ne $script:SearchWorkerProcess -or $null -ne $script:HopWorkerProcess -or $null -ne $script:PgmxWorkerProcess -or $script:AutoMixChainRunning -or $script:TrainingActive -or $script:TimeImportRunning -or $script:TimeImportSynchronousRunning)
    if($busy){return}
    $script:OperatorScheduleLastRunDate=$today
    try{Save-OperatorScheduleSettings}catch{
        $script:OperatorScheduleLastRunDate=''
        $status.Text='Автоматичният старт е спрян: не може да се запише LastRunDate: '+$_.Exception.Message
        return
    }
    try{Start-OperatorFullFlow -Mode $script:OperatorScheduleMode}catch{$status.Text='Автоматичният старт не успя: '+$_.Exception.Message}
}

# V5.1.38 isolated batch/status engine. Keep current core unchanged beyond explicit hooks.
$batchEngine=Join-Path $script:AppRoot 'CAM_Agent_BatchEngine_V5_1_38.ps1'
if(-not [System.IO.File]::Exists($batchEngine)){throw ('Липсва CAM_Agent_BatchEngine_V5_1_38.ps1 до приложението: '+$batchEngine)}
. $batchEngine

$form = New-Object System.Windows.Forms.Form
$form.Text = 'PWX500 Agent V5.1.38 - Дневен каталог + CSV регистър TEST'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1360, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1180, 720)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'PWX500 Agent V5.1.38 - Дневен каталог + CSV регистър TEST'
$title.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(16, 14)
$form.Controls.Add($title)

$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = 'Папка с CSV файлове:'
$pathLabel.AutoSize = $true
$pathLabel.Location = New-Object System.Drawing.Point(18, 55)
$form.Controls.Add($pathLabel)

$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Location = New-Object System.Drawing.Point(18, 78)
$pathBox.Size = New-Object System.Drawing.Size(865, 28)
$pathBox.Anchor = 'Top,Left,Right'
$pathBox.Text = $script:DefaultCsvFolder
$pathBox.ReadOnly = $true
$form.Controls.Add($pathBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = 'Фиксирана папка'
$browseButton.Location = New-Object System.Drawing.Point(895, 75)
$browseButton.Size = New-Object System.Drawing.Size(140, 32)
$browseButton.Anchor = 'Top,Right'
$browseButton.Visible = $false
$form.Controls.Add($browseButton)

$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Text = 'Прочети CSV'
$scanButton.Location = New-Object System.Drawing.Point(895, 75)
$scanButton.Size = New-Object System.Drawing.Size(135, 32)
$scanButton.Anchor = 'Top,Right'
$form.Controls.Add($scanButton)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = 'Експорт контрол'
$exportButton.Location = New-Object System.Drawing.Point(1040, 75)
$exportButton.Size = New-Object System.Drawing.Size(145, 32)
$exportButton.Anchor = 'Top,Right'
$exportButton.Enabled = $false
$form.Controls.Add($exportButton)

$security = New-Object System.Windows.Forms.Label
$security.Text = 'БЕЗОПАСНОСТ: Word НИКОГА не се отваря. HOP се чете само след .hop филтър в потвърденото 1_Holzher. При нормална обработка PGMX съдържание не се чете. Съществуващи производствени файлове не се презаписват; изключение е дневният Mix(дд.ММ).mixx, който се създава наново.'
$security.ForeColor = [System.Drawing.Color]::DarkRed
$security.AutoSize = $false
$security.Location = New-Object System.Drawing.Point(18, 116)
$security.Size = New-Object System.Drawing.Size(1310, 40)
$security.Anchor = 'Top,Left,Right'
$form.Controls.Add($security)

$status = New-Object System.Windows.Forms.Label
$status.Text = 'Фиксирана CSV папка: ' + $script:DefaultCsvFolder
$status.AutoSize = $false
$status.Location = New-Object System.Drawing.Point(18, 156)
$status.Size = New-Object System.Drawing.Size(1310, 25)
$status.Anchor = 'Top,Left,Right'
$form.Controls.Add($status)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(18, 188)
$tabs.Size = New-Object System.Drawing.Size(1318, 610)
$tabs.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($tabs)

$tabOperator = New-Object System.Windows.Forms.TabPage
$tabOperator.Text = 'ОПЕРАТОР'
$tabs.TabPages.Add($tabOperator)

$operatorPanel=New-Object System.Windows.Forms.Panel
$operatorPanel.Dock='Fill'
$tabOperator.Controls.Add($operatorPanel)

$operatorTitle=New-Object System.Windows.Forms.Label
$operatorTitle.Text='Последователна работа - дневен каталог + без повторение на завършени CSV'
$operatorTitle.Font=New-Object System.Drawing.Font('Segoe UI',12,[System.Drawing.FontStyle]::Bold)
$operatorTitle.AutoSize=$true;$operatorTitle.Location=New-Object System.Drawing.Point(12,10);$operatorPanel.Controls.Add($operatorTitle)

$operatorReadCsvButton=New-Object System.Windows.Forms.Button
$operatorReadCsvButton.Text='1. Прочети CSV';$operatorReadCsvButton.Location=New-Object System.Drawing.Point(12,42);$operatorReadCsvButton.Size=New-Object System.Drawing.Size(175,42);$operatorPanel.Controls.Add($operatorReadCsvButton)
$operatorValidateButton=New-Object System.Windows.Forms.Button
$operatorValidateButton.Text='2. Намери + HOP проверка';$operatorValidateButton.Location=New-Object System.Drawing.Point(197,42);$operatorValidateButton.Size=New-Object System.Drawing.Size(215,42);$operatorPanel.Controls.Add($operatorValidateButton)
$operatorReviewButton=New-Object System.Windows.Forms.Button
$operatorReviewButton.Text='Прегледай неясните';$operatorReviewButton.Location=New-Object System.Drawing.Point(422,42);$operatorReviewButton.Size=New-Object System.Drawing.Size(185,42);$operatorReviewButton.BackColor=[System.Drawing.Color]::LightYellow;$operatorPanel.Controls.Add($operatorReviewButton)
$operatorContinueButton=New-Object System.Windows.Forms.Button
$operatorContinueButton.Text='3. ПРОДЪЛЖИ С ПОТВЪРДЕНИТЕ';$operatorContinueButton.Location=New-Object System.Drawing.Point(617,42);$operatorContinueButton.Size=New-Object System.Drawing.Size(265,42);$operatorContinueButton.BackColor=[System.Drawing.Color]::LightGreen;$operatorPanel.Controls.Add($operatorContinueButton)
$operatorRunAllButton=New-Object System.Windows.Forms.Button
$operatorRunAllButton.Text='AUTO: целият процес сега';$operatorRunAllButton.Location=New-Object System.Drawing.Point(892,42);$operatorRunAllButton.Size=New-Object System.Drawing.Size(215,42);$operatorPanel.Controls.Add($operatorRunAllButton)

$operatorScheduleBox=New-Object System.Windows.Forms.GroupBox
$operatorScheduleBox.Text='Автоматичен начален час (режимът се избира в главния прозорец)';$operatorScheduleBox.Location=New-Object System.Drawing.Point(1117,31);$operatorScheduleBox.Size=New-Object System.Drawing.Size(175,82);$operatorPanel.Controls.Add($operatorScheduleBox)
$operatorScheduleCheck=New-Object System.Windows.Forms.CheckBox;$operatorScheduleCheck.Text='Включено';$operatorScheduleCheck.Location=New-Object System.Drawing.Point(10,20);$operatorScheduleCheck.AutoSize=$true;$operatorScheduleBox.Controls.Add($operatorScheduleCheck)
$operatorScheduleTime=New-Object System.Windows.Forms.DateTimePicker;$operatorScheduleTime.Format=[System.Windows.Forms.DateTimePickerFormat]::Custom;$operatorScheduleTime.CustomFormat='HH:mm';$operatorScheduleTime.ShowUpDown=$true;$operatorScheduleTime.Location=New-Object System.Drawing.Point(88,18);$operatorScheduleTime.Size=New-Object System.Drawing.Size(72,25);$operatorScheduleBox.Controls.Add($operatorScheduleTime)
$operatorScheduleSaveButton=New-Object System.Windows.Forms.Button;$operatorScheduleSaveButton.Text='Запази';$operatorScheduleSaveButton.Location=New-Object System.Drawing.Point(10,47);$operatorScheduleSaveButton.Size=New-Object System.Drawing.Size(150,25);$operatorScheduleBox.Controls.Add($operatorScheduleSaveButton)

$operatorCsvSummaryLabel=New-Object System.Windows.Forms.Label;$operatorCsvSummaryLabel.Text='CSV: -';$operatorCsvSummaryLabel.Location=New-Object System.Drawing.Point(14,96);$operatorCsvSummaryLabel.Size=New-Object System.Drawing.Size(380,24);$operatorCsvSummaryLabel.Font=New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold);$operatorPanel.Controls.Add($operatorCsvSummaryLabel)
$operatorProductSummaryLabel=New-Object System.Windows.Forms.Label;$operatorProductSummaryLabel.Text='Потвърдени изделия: - | непотвърдени: -';$operatorProductSummaryLabel.Location=New-Object System.Drawing.Point(400,96);$operatorProductSummaryLabel.Size=New-Object System.Drawing.Size(420,24);$operatorProductSummaryLabel.Font=New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold);$operatorPanel.Controls.Add($operatorProductSummaryLabel)
$operatorCatalogLabel=New-Object System.Windows.Forms.Label;$operatorCatalogLabel.Text='Каталог: проверка...';$operatorCatalogLabel.Location=New-Object System.Drawing.Point(830,94);$operatorCatalogLabel.Size=New-Object System.Drawing.Size(285,22);$operatorCatalogLabel.ForeColor=[System.Drawing.Color]::DarkBlue;$operatorPanel.Controls.Add($operatorCatalogLabel)
$operatorSkippedCountLabel=New-Object System.Windows.Forms.Label;$operatorSkippedCountLabel.Text='Непотвърдени / пропуснати изделия: 0';$operatorSkippedCountLabel.Location=New-Object System.Drawing.Point(14,124);$operatorSkippedCountLabel.Size=New-Object System.Drawing.Size(390,22);$operatorPanel.Controls.Add($operatorSkippedCountLabel)
$operatorErrorCountLabel=New-Object System.Windows.Forms.Label;$operatorErrorCountLabel.Text='Детайли, които не са направени: 0';$operatorErrorCountLabel.Location=New-Object System.Drawing.Point(410,124);$operatorErrorCountLabel.Size=New-Object System.Drawing.Size(390,22);$operatorPanel.Controls.Add($operatorErrorCountLabel)
$operatorOpenReportsButton=New-Object System.Windows.Forms.Button;$operatorOpenReportsButton.Text='Отвори папката с отчетите';$operatorOpenReportsButton.Location=New-Object System.Drawing.Point(830,116);$operatorOpenReportsButton.Size=New-Object System.Drawing.Size(220,30);$operatorPanel.Controls.Add($operatorOpenReportsButton)
$operatorAdvancedLabel=New-Object System.Windows.Forms.Label;$operatorAdvancedLabel.Text='Техническите табове вдясно остават само за диагностика/настройки.';$operatorAdvancedLabel.Location=New-Object System.Drawing.Point(1060,121);$operatorAdvancedLabel.Size=New-Object System.Drawing.Size(235,34);$operatorAdvancedLabel.ForeColor=[System.Drawing.Color]::DimGray;$operatorPanel.Controls.Add($operatorAdvancedLabel)

$operatorReportTabs=New-Object System.Windows.Forms.TabControl;$operatorReportTabs.Location=New-Object System.Drawing.Point(12,154);$operatorReportTabs.Size=New-Object System.Drawing.Size(1280,360);$operatorReportTabs.Anchor='Top,Bottom,Left,Right';$operatorPanel.Controls.Add($operatorReportTabs)
$operatorSkippedTab=New-Object System.Windows.Forms.TabPage;$operatorSkippedTab.Text='Непотвърдени / пропуснати изделия';$operatorReportTabs.TabPages.Add($operatorSkippedTab)
$operatorErrorsTab=New-Object System.Windows.Forms.TabPage;$operatorErrorsTab.Text='Детайли с грешка / несъздадени';$operatorReportTabs.TabPages.Add($operatorErrorsTab)
$operatorCsvRegistryTab=New-Object System.Windows.Forms.TabPage;$operatorCsvRegistryTab.Text='История CSV';$operatorReportTabs.TabPages.Add($operatorCsvRegistryTab)
$operatorSkippedGrid=New-Object System.Windows.Forms.DataGridView;$operatorSkippedGrid.Dock='Fill';$operatorSkippedGrid.ReadOnly=$true;$operatorSkippedGrid.AllowUserToAddRows=$false;$operatorSkippedGrid.AllowUserToDeleteRows=$false;$operatorSkippedGrid.AutoSizeColumnsMode='Fill';$operatorSkippedGrid.SelectionMode='FullRowSelect';$operatorSkippedTab.Controls.Add($operatorSkippedGrid)
$operatorErrorsGrid=New-Object System.Windows.Forms.DataGridView;$operatorErrorsGrid.Dock='Fill';$operatorErrorsGrid.ReadOnly=$true;$operatorErrorsGrid.AllowUserToAddRows=$false;$operatorErrorsGrid.AllowUserToDeleteRows=$false;$operatorErrorsGrid.AutoSizeColumnsMode='DisplayedCells';$operatorErrorsGrid.SelectionMode='FullRowSelect';$operatorErrorsTab.Controls.Add($operatorErrorsGrid)
$operatorCsvRegistryGrid=New-Object System.Windows.Forms.DataGridView;$operatorCsvRegistryGrid.Dock='Fill';$operatorCsvRegistryGrid.ReadOnly=$true;$operatorCsvRegistryGrid.AllowUserToAddRows=$false;$operatorCsvRegistryGrid.AllowUserToDeleteRows=$false;$operatorCsvRegistryGrid.AutoSizeColumnsMode='DisplayedCells';$operatorCsvRegistryGrid.SelectionMode='FullRowSelect';$operatorCsvRegistryTab.Controls.Add($operatorCsvRegistryGrid)

$operatorProcessStatusLabel=New-Object System.Windows.Forms.Label;$operatorProcessStatusLabel.Text='Готово за работа.';$operatorProcessStatusLabel.Location=New-Object System.Drawing.Point(14,522);$operatorProcessStatusLabel.Size=New-Object System.Drawing.Size(1270,42);$operatorProcessStatusLabel.Anchor='Bottom,Left,Right';$operatorProcessStatusLabel.Font=New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold);$operatorPanel.Controls.Add($operatorProcessStatusLabel)

$tabProducts = New-Object System.Windows.Forms.TabPage
$tabProducts.Text = 'Изделия'
$tabs.TabPages.Add($tabProducts)

$tabDetails = New-Object System.Windows.Forms.TabPage
$tabDetails.Text = 'Детайли / ProgramCodeF'
$tabs.TabPages.Add($tabDetails)

$tabPerCsv = New-Object System.Windows.Forms.TabPage
$tabPerCsv.Text = 'Контрол по CSV'
$tabs.TabPages.Add($tabPerCsv)

$tabCandidates = New-Object System.Windows.Forms.TabPage
$tabCandidates.Text = 'Изделия - папки'
$tabs.TabPages.Add($tabCandidates)

$tabConfirm = New-Object System.Windows.Forms.TabPage
$tabConfirm.Text = 'Потвърждение'
$tabs.TabPages.Add($tabConfirm)

$tabHop = New-Object System.Windows.Forms.TabPage
$tabHop.Text = 'HOP проверка'
$tabs.TabPages.Add($tabHop)

$tabPgmx = New-Object System.Windows.Forms.TabPage
$tabPgmx.Text = 'PGMX / Конверсия'
$tabs.TabPages.Add($tabPgmx)

$tabHistory = New-Object System.Windows.Forms.TabPage
$tabHistory.Text = 'Исторически PGMX'
$tabs.TabPages.Add($tabHistory)

$tabTraining = New-Object System.Windows.Forms.TabPage
$tabTraining.Text = '4. Създаване на .mixx'
$tabs.TabPages.Add($tabTraining)

$tabPowerSet = New-Object System.Windows.Forms.TabPage
$tabPowerSet.Text = '5. PowerSet'
$tabs.TabPages.Add($tabPowerSet)

$tabAllowed = New-Object System.Windows.Forms.TabPage
$tabAllowed.Text = 'Разрешени папки'
$tabs.TabPages.Add($tabAllowed)

$tabCategories = New-Object System.Windows.Forms.TabPage
$tabCategories.Text = 'Категории търсене'
$tabs.TabPages.Add($tabCategories)

$tabIndex = New-Object System.Windows.Forms.TabPage
$tabIndex.Text = 'Каталог изделия'
$tabs.TabPages.Add($tabIndex)

$tabWarnings = New-Object System.Windows.Forms.TabPage
$tabWarnings.Text = 'Проверки'
$tabs.TabPages.Add($tabWarnings)

$productsGrid = New-Object System.Windows.Forms.DataGridView
$productsGrid.Dock = 'Fill'
$productsGrid.ReadOnly = $true
$productsGrid.AllowUserToAddRows = $false
$productsGrid.AllowUserToDeleteRows = $false
$productsGrid.AutoSizeColumnsMode = 'Fill'
$productsGrid.SelectionMode = 'FullRowSelect'
$productsGrid.MultiSelect = $false
$tabProducts.Controls.Add($productsGrid)

$detailsGrid = New-Object System.Windows.Forms.DataGridView
$detailsGrid.Dock = 'Fill'
$detailsGrid.ReadOnly = $true
$detailsGrid.AllowUserToAddRows = $false
$detailsGrid.AllowUserToDeleteRows = $false
$detailsGrid.AutoSizeColumnsMode = 'DisplayedCells'
$detailsGrid.SelectionMode = 'FullRowSelect'
$tabDetails.Controls.Add($detailsGrid)

$perCsvSplit = New-Object System.Windows.Forms.SplitContainer
$perCsvSplit.Dock = 'Fill'
$perCsvSplit.Orientation = 'Horizontal'
$perCsvSplit.SplitterDistance = 290
$tabPerCsv.Controls.Add($perCsvSplit)

$fileSummaryGrid = New-Object System.Windows.Forms.DataGridView
$fileSummaryGrid.Dock = 'Fill'
$fileSummaryGrid.ReadOnly = $true
$fileSummaryGrid.AllowUserToAddRows = $false
$fileSummaryGrid.AllowUserToDeleteRows = $false
$fileSummaryGrid.AutoSizeColumnsMode = 'Fill'
$fileSummaryGrid.SelectionMode = 'FullRowSelect'
$fileSummaryGrid.MultiSelect = $false
$perCsvSplit.Panel1.Controls.Add($fileSummaryGrid)

$fileProductsGrid = New-Object System.Windows.Forms.DataGridView
$fileProductsGrid.Dock = 'Fill'
$fileProductsGrid.ReadOnly = $true
$fileProductsGrid.AllowUserToAddRows = $false
$fileProductsGrid.AllowUserToDeleteRows = $false
$fileProductsGrid.AutoSizeColumnsMode = 'Fill'
$fileProductsGrid.SelectionMode = 'FullRowSelect'
$fileProductsGrid.MultiSelect = $false
$perCsvSplit.Panel2.Controls.Add($fileProductsGrid)

$candidateLayout = New-Object System.Windows.Forms.TableLayoutPanel
$candidateLayout.Dock = 'Fill'
$candidateLayout.RowCount = 2
$candidateLayout.ColumnCount = 1
[void]$candidateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48)))
[void]$candidateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabCandidates.Controls.Add($candidateLayout)

$candidatePanel = New-Object System.Windows.Forms.Panel
$candidatePanel.Dock = 'Fill'
$candidateLayout.Controls.Add($candidatePanel, 0, 0)

$findProductsButton = New-Object System.Windows.Forms.Button
$findProductsButton.Text = 'Търси изделия'
$findProductsButton.Location = New-Object System.Drawing.Point(8, 8)
$findProductsButton.Size = New-Object System.Drawing.Size(140, 30)
$candidatePanel.Controls.Add($findProductsButton)

$openConfirmButton = New-Object System.Windows.Forms.Button
$openConfirmButton.Text = 'Потвърди спорните'
$openConfirmButton.Location = New-Object System.Drawing.Point(156, 8)
$openConfirmButton.Size = New-Object System.Drawing.Size(150, 30)
$candidatePanel.Controls.Add($openConfirmButton)

$hopCheckButton = New-Object System.Windows.Forms.Button
$hopCheckButton.Text = 'HOP проверка'
$hopCheckButton.Location = New-Object System.Drawing.Point(314, 8)
$hopCheckButton.Size = New-Object System.Drawing.Size(135, 30)
$candidatePanel.Controls.Add($hopCheckButton)

$changeProductFolderButton = New-Object System.Windows.Forms.Button
$changeProductFolderButton.Text = 'Промени директория'
$changeProductFolderButton.Location = New-Object System.Drawing.Point(457, 8)
$changeProductFolderButton.Size = New-Object System.Drawing.Size(175, 30)
$changeProductFolderButton.BackColor = [System.Drawing.Color]::LightYellow
$candidatePanel.Controls.Add($changeProductFolderButton)

$candidateInfoLabel = New-Object System.Windows.Forms.Label
$candidateInfoLabel.Text = 'Избери ред и „Промени директория“ при нужда — работи и при 1 силен кандидат. F и F# са различни.'
$candidateInfoLabel.AutoSize = $true
$candidateInfoLabel.Location = New-Object System.Drawing.Point(642, 15)
$candidatePanel.Controls.Add($candidateInfoLabel)

$productCandidatesGrid = New-Object System.Windows.Forms.DataGridView
$productCandidatesGrid.Dock = 'Fill'
$productCandidatesGrid.ReadOnly = $true
$productCandidatesGrid.AllowUserToAddRows = $false
$productCandidatesGrid.AllowUserToDeleteRows = $false
$productCandidatesGrid.AutoSizeColumnsMode = 'Fill'
$productCandidatesGrid.SelectionMode = 'FullRowSelect'
$candidateLayout.Controls.Add($productCandidatesGrid, 0, 1)


$categoryLayout = New-Object System.Windows.Forms.TableLayoutPanel
$categoryLayout.Dock='Fill';$categoryLayout.RowCount=2;$categoryLayout.ColumnCount=1
[void]$categoryLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,70)))
[void]$categoryLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
$tabCategories.Controls.Add($categoryLayout)
$categoryPanel=New-Object System.Windows.Forms.Panel;$categoryPanel.Dock='Fill';$categoryLayout.Controls.Add($categoryPanel,0,0)
$categorySaveButton=New-Object System.Windows.Forms.Button;$categorySaveButton.Text='Запази категории';$categorySaveButton.Location=New-Object System.Drawing.Point(8,8);$categorySaveButton.Size=New-Object System.Drawing.Size(160,30);$categoryPanel.Controls.Add($categorySaveButton)
$categoryAddButton=New-Object System.Windows.Forms.Button;$categoryAddButton.Text='Добави ред';$categoryAddButton.Location=New-Object System.Drawing.Point(176,8);$categoryAddButton.Size=New-Object System.Drawing.Size(120,30);$categoryPanel.Controls.Add($categoryAddButton)
$categoryDeleteButton=New-Object System.Windows.Forms.Button;$categoryDeleteButton.Text='Изтрий ред';$categoryDeleteButton.Location=New-Object System.Drawing.Point(304,8);$categoryDeleteButton.Size=New-Object System.Drawing.Size(120,30);$categoryPanel.Controls.Add($categoryDeleteButton)
$categoryResetButton=New-Object System.Windows.Forms.Button;$categoryResetButton.Text='Начални правила';$categoryResetButton.Location=New-Object System.Drawing.Point(432,8);$categoryResetButton.Size=New-Object System.Drawing.Size(140,30);$categoryPanel.Controls.Add($categoryResetButton)
$categoryHelp=New-Object System.Windows.Forms.Label;$categoryHelp.Text='Phrase = начало на името от CSV. Targets = допустими сегменти, разделени с |. Пример Bedroom and bed -> 12_Spalni_i_Legla|02_TM. По-висок Priority печели.';$categoryHelp.Location=New-Object System.Drawing.Point(8,43);$categoryHelp.Size=New-Object System.Drawing.Size(1180,22);$categoryPanel.Controls.Add($categoryHelp)
$categoryGrid=New-Object System.Windows.Forms.DataGridView;$categoryGrid.Dock='Fill';$categoryGrid.AllowUserToAddRows=$true;$categoryGrid.AllowUserToDeleteRows=$true;$categoryGrid.AutoSizeColumnsMode='Fill';$categoryGrid.SelectionMode='FullRowSelect';$categoryLayout.Controls.Add($categoryGrid,0,1)

$hopLayout = New-Object System.Windows.Forms.TableLayoutPanel
$hopLayout.Dock = 'Fill'
$hopLayout.RowCount = 2
$hopLayout.ColumnCount = 1
[void]$hopLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
[void]$hopLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabHop.Controls.Add($hopLayout)

$hopInfoLabel = New-Object System.Windows.Forms.Label
$hopInfoLabel.Text = 'HOP проверката се пуска само за непотвърдени изделия с повече от един кандидат. Четат се само *.hop в 03_Mashini\\1_Holzher.'
$hopInfoLabel.Dock = 'Fill'
$hopInfoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hopInfoLabel.Margin = New-Object System.Windows.Forms.Padding(8,4,8,2)
$hopLayout.Controls.Add($hopInfoLabel,0,0)

$hopGrid = New-Object System.Windows.Forms.DataGridView
$hopGrid.Dock = 'Fill'
$hopGrid.ReadOnly = $true
$hopGrid.AllowUserToAddRows = $false
$hopGrid.AllowUserToDeleteRows = $false
$hopGrid.AutoSizeColumnsMode = 'DisplayedCells'
$hopGrid.SelectionMode = 'FullRowSelect'
$hopGrid.MultiSelect = $false
$hopLayout.Controls.Add($hopGrid,0,1)

# Keep HOP colours stable after bind and after user sorting.
$hopGrid.Add_DataBindingComplete({
    try { Apply-HopGridVisualStyles } catch {}
})
$hopGrid.Add_Sorted({
    try { Apply-HopGridVisualStyles } catch {}
})

$pgmxLayout = New-Object System.Windows.Forms.TableLayoutPanel
$pgmxLayout.Dock = 'Fill'
$pgmxLayout.RowCount = 2
$pgmxLayout.ColumnCount = 1
[void]$pgmxLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 92)))
[void]$pgmxLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabPgmx.Controls.Add($pgmxLayout)

$pgmxPanel = New-Object System.Windows.Forms.Panel
$pgmxPanel.Dock = 'Fill'
$pgmxLayout.Controls.Add($pgmxPanel,0,0)

$pgmxTestLabel = New-Object System.Windows.Forms.Label
$pgmxTestLabel.Text = 'ТЕСТОВ РЕЖИМ — ВСИЧКИ изделия от текущите CSV:'
$pgmxTestLabel.Location = New-Object System.Drawing.Point(8,8)
$pgmxTestLabel.Size = New-Object System.Drawing.Size(430,24)
$pgmxTestLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$pgmxTestLabel.ForeColor = [System.Drawing.Color]::DarkRed
$pgmxPanel.Controls.Add($pgmxTestLabel)

$pgmxAllModeLabel = New-Object System.Windows.Forms.Label
$pgmxAllModeLabel.Text = 'Прегледът не записва. Изпълнението пропуска проблемните случаи.'
$pgmxAllModeLabel.Location = New-Object System.Drawing.Point(445,8)
$pgmxAllModeLabel.Size = New-Object System.Drawing.Size(560,24)
$pgmxPanel.Controls.Add($pgmxAllModeLabel)

$pgmxCheckButton = New-Object System.Windows.Forms.Button
$pgmxCheckButton.Text = 'Преглед на всички'
$pgmxCheckButton.Location = New-Object System.Drawing.Point(8,42)
$pgmxCheckButton.Size = New-Object System.Drawing.Size(180,32)
$pgmxPanel.Controls.Add($pgmxCheckButton)

$pgmxExecuteButton = New-Object System.Windows.Forms.Button
$pgmxExecuteButton.Text = 'ТЕСТ: обработи всички'
$pgmxExecuteButton.Location = New-Object System.Drawing.Point(196,42)
$pgmxExecuteButton.Size = New-Object System.Drawing.Size(245,32)
$pgmxPanel.Controls.Add($pgmxExecuteButton)

$pgmxInfoLabel = New-Object System.Windows.Forms.Label
$pgmxInfoLabel.Text = 'Общ тест: Преглед на всички не записва нищо. При изпълнение САМО изделията със специална вложена структура копират крайния ProgramCodeF и в V:\30_Biesse_edging\scm\SCM_PWX_500. Нормалните изделия остават само в собствената 2_PWX_500. Основните PGMX никога не се копират централно. Проблемните случаи се пропускат. Няма презаписване.'
$pgmxInfoLabel.AutoSize = $false
$pgmxInfoLabel.Location = New-Object System.Drawing.Point(455,40)
$pgmxInfoLabel.Size = New-Object System.Drawing.Size(825,42)
$pgmxInfoLabel.Anchor = 'Top,Left,Right'
$pgmxPanel.Controls.Add($pgmxInfoLabel)

$pgmxGrid = New-Object System.Windows.Forms.DataGridView
$pgmxGrid.Dock = 'Fill'
$pgmxGrid.ReadOnly = $true
$pgmxGrid.AllowUserToAddRows = $false
$pgmxGrid.AllowUserToDeleteRows = $false
$pgmxGrid.AutoSizeColumnsMode = 'DisplayedCells'
$pgmxGrid.SelectionMode = 'FullRowSelect'
$pgmxGrid.MultiSelect = $false
$pgmxLayout.Controls.Add($pgmxGrid,0,1)
$pgmxGrid.Add_DataBindingComplete({ try { Apply-PgmxGridVisualStyles } catch {} })
$pgmxGrid.Add_Sorted({ try { Apply-PgmxGridVisualStyles } catch {} })

$historyLayout = New-Object System.Windows.Forms.TableLayoutPanel
$historyLayout.Dock = 'Fill'
$historyLayout.RowCount = 2
$historyLayout.ColumnCount = 1
[void]$historyLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 72)))
[void]$historyLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabHistory.Controls.Add($historyLayout)

$historyPanel = New-Object System.Windows.Forms.Panel
$historyPanel.Dock = 'Fill'
$historyLayout.Controls.Add($historyPanel,0,0)

$historyMigrationButton = New-Object System.Windows.Forms.Button
$historyMigrationButton.Text = 'Бърза миграция от Excel'
$historyMigrationButton.Location = New-Object System.Drawing.Point(8,12)
$historyMigrationButton.Size = New-Object System.Drawing.Size(210,34)
$historyPanel.Controls.Add($historyMigrationButton)

$historyInfoLabel = New-Object System.Windows.Forms.Label
$historyInfoLabel.Text = 'Еднократна инициализация V0.2.20: избира се „Поръчки за линия.xlsx“. Агентът прави един общ обход на разрешените подпапки, индексира само имената на *.pgmx в директните 03_Mashini\2_PWX_500 и после сравнява Excel историята в паметта. 2 различни физически детайла в една папка потвърждават директорията.'
$historyInfoLabel.AutoSize = $false
$historyInfoLabel.Location = New-Object System.Drawing.Point(232,7)
$historyInfoLabel.Size = New-Object System.Drawing.Size(1050,55)
$historyInfoLabel.Anchor = 'Top,Left,Right'
$historyPanel.Controls.Add($historyInfoLabel)

$historyGrid = New-Object System.Windows.Forms.DataGridView
$historyGrid.Dock = 'Fill'
$historyGrid.ReadOnly = $true
$historyGrid.AllowUserToAddRows = $false
$historyGrid.AllowUserToDeleteRows = $false
$historyGrid.AutoSizeColumnsMode = 'DisplayedCells'
$historyGrid.SelectionMode = 'FullRowSelect'
$historyGrid.MultiSelect = $false
$historyLayout.Controls.Add($historyGrid,0,1)
$historyGrid.Add_DataBindingComplete({ try { Apply-HistoryGridVisualStyles } catch {} })
$historyGrid.Add_Sorted({ try { Apply-HistoryGridVisualStyles } catch {} })

$confirmLayout = New-Object System.Windows.Forms.TableLayoutPanel
$confirmLayout.Dock = 'Fill'
$confirmLayout.ColumnCount = 2
$confirmLayout.RowCount = 1
[void]$confirmLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 370)))
[void]$confirmLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabConfirm.Controls.Add($confirmLayout)

# Left side: queue label + product list.
$confirmLeftLayout = New-Object System.Windows.Forms.TableLayoutPanel
$confirmLeftLayout.Dock = 'Fill'
$confirmLeftLayout.RowCount = 2
$confirmLeftLayout.ColumnCount = 1
[void]$confirmLeftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 38)))
[void]$confirmLeftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$confirmLayout.Controls.Add($confirmLeftLayout,0,0)

$confirmQueueLabel = New-Object System.Windows.Forms.Label
$confirmQueueLabel.Text = 'За потвърждение: 0 изделия'
$confirmQueueLabel.AutoSize = $true
$confirmQueueLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$confirmQueueLabel.Margin = New-Object System.Windows.Forms.Padding(8,10,4,4)
$confirmLeftLayout.Controls.Add($confirmQueueLabel,0,0)

$confirmProductList = New-Object System.Windows.Forms.ListBox
$confirmProductList.Dock = 'Fill'
$confirmProductList.Margin = New-Object System.Windows.Forms.Padding(8,2,8,8)
$confirmLeftLayout.Controls.Add($confirmProductList,0,1)

# Right side: product, instructions, candidates, buttons, safety note.
$confirmRightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$confirmRightLayout.Dock = 'Fill'
$confirmRightLayout.RowCount = 5
$confirmRightLayout.ColumnCount = 1
[void]$confirmRightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
[void]$confirmRightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
[void]$confirmRightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$confirmRightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
[void]$confirmRightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48)))
$confirmLayout.Controls.Add($confirmRightLayout,1,0)

$confirmProductLabel = New-Object System.Windows.Forms.Label
$confirmProductLabel.Text = 'Избери изделие от списъка вляво.'
$confirmProductLabel.Dock = 'Fill'
$confirmProductLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$confirmProductLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$confirmProductLabel.Margin = New-Object System.Windows.Forms.Padding(12,4,8,0)
$confirmRightLayout.Controls.Add($confirmProductLabel,0,0)

$confirmHintLabel = New-Object System.Windows.Forms.Label
$confirmHintLabel.Text = 'Избери правилния кандидат. Ако го няма в предложенията, натисни „Посочи точна папка“. '
$confirmHintLabel.Dock = 'Fill'
$confirmHintLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$confirmHintLabel.Margin = New-Object System.Windows.Forms.Padding(12,0,8,0)
$confirmRightLayout.Controls.Add($confirmHintLabel,0,1)

$confirmCandidateList = New-Object System.Windows.Forms.ListBox
$confirmCandidateList.Dock = 'Fill'
$confirmCandidateList.HorizontalScrollbar = $true
$confirmCandidateList.Margin = New-Object System.Windows.Forms.Padding(12,2,8,6)
$confirmRightLayout.Controls.Add($confirmCandidateList,0,2)

$confirmButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$confirmButtonPanel.Dock = 'Fill'
$confirmButtonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$confirmButtonPanel.WrapContents = $false
$confirmButtonPanel.Padding = New-Object System.Windows.Forms.Padding(8,6,0,4)
$confirmRightLayout.Controls.Add($confirmButtonPanel,0,3)

$confirmSelectedButton = New-Object System.Windows.Forms.Button
$confirmSelectedButton.Text = 'Потвърди избрания кандидат'
$confirmSelectedButton.Size = New-Object System.Drawing.Size(195,34)
$confirmButtonPanel.Controls.Add($confirmSelectedButton)

$chooseExactFolderButton = New-Object System.Windows.Forms.Button
$chooseExactFolderButton.Text = 'Посочи точна папка'
$chooseExactFolderButton.Size = New-Object System.Drawing.Size(150,34)
$confirmButtonPanel.Controls.Add($chooseExactFolderButton)

$skipConfirmButton = New-Object System.Windows.Forms.Button
$skipConfirmButton.Text = 'Пропусни / следващ'
$skipConfirmButton.Size = New-Object System.Drawing.Size(135,34)
$confirmButtonPanel.Controls.Add($skipConfirmButton)

$removeMappingButton = New-Object System.Windows.Forms.Button
$removeMappingButton.Text = 'Изтрий запомнената връзка'
$removeMappingButton.Size = New-Object System.Drawing.Size(180,34)
$confirmButtonPanel.Controls.Add($removeMappingButton)

$confirmContinueButton = New-Object System.Windows.Forms.Button
$confirmContinueButton.Text = 'Продължи с потвърдените'
$confirmContinueButton.Size = New-Object System.Drawing.Size(210,34)
$confirmContinueButton.BackColor = [System.Drawing.Color]::LightGreen
$confirmButtonPanel.Controls.Add($confirmContinueButton)

$confirmSafetyLabel = New-Object System.Windows.Forms.Label
$confirmSafetyLabel.Text = 'Автоматичното търсене е само в разрешените корени. Ръчен избор извън тях е разрешен след потвърждение. HOP: № + размери са задължителни; _1/_2 следват реда в CSV.'
$confirmSafetyLabel.ForeColor = [System.Drawing.Color]::DarkRed
$confirmSafetyLabel.Dock = 'Fill'
$confirmSafetyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$confirmSafetyLabel.Margin = New-Object System.Windows.Forms.Padding(12,0,8,4)
$confirmRightLayout.Controls.Add($confirmSafetyLabel,0,4)


# ---------------- TRAINING MODE UI ----------------
$trainingLayout = New-Object System.Windows.Forms.TableLayoutPanel
$trainingLayout.Dock = 'Fill'
$trainingLayout.RowCount = 3
$trainingLayout.ColumnCount = 1
[void]$trainingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 245)))
[void]$trainingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 62)))
[void]$trainingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 38)))
$tabTraining.Controls.Add($trainingLayout)

$trainingTop = New-Object System.Windows.Forms.Panel
$trainingTop.Dock = 'Fill'
$trainingLayout.Controls.Add($trainingTop,0,0)

# Automatic daily Mixx controls
$autoMixTitle = New-Object System.Windows.Forms.Label
$autoMixTitle.Text = 'СТЪПКИ 4 + 5 - пълен AUTO или отделно ръчно след готови PGMX'
$autoMixTitle.Font = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$autoMixTitle.AutoSize = $true
$autoMixTitle.Location = New-Object System.Drawing.Point(8,8)
$trainingTop.Controls.Add($autoMixTitle)

$autoMixFolderLabel = New-Object System.Windows.Forms.Label
$autoMixFolderLabel.Text = 'Дневна папка:'
$autoMixFolderLabel.AutoSize = $true
$autoMixFolderLabel.Location = New-Object System.Drawing.Point(8,38)
$trainingTop.Controls.Add($autoMixFolderLabel)

$autoMixFolderBox = New-Object System.Windows.Forms.TextBox
$autoMixFolderBox.Location = New-Object System.Drawing.Point(100,34)
$autoMixFolderBox.Size = New-Object System.Drawing.Size(395,26)
$autoMixFolderBox.ReadOnly = $true
$autoMixFolderBox.Text = Get-DailyMixFolder
$trainingTop.Controls.Add($autoMixFolderBox)

$autoMixChooseRootButton = New-Object System.Windows.Forms.Button
$autoMixChooseRootButton.Text = 'Основна папка...'
$autoMixChooseRootButton.Location = New-Object System.Drawing.Point(505,31)
$autoMixChooseRootButton.Size = New-Object System.Drawing.Size(115,32)
$trainingTop.Controls.Add($autoMixChooseRootButton)
$autoMixChooseRootButton.Add_Click({try{[void](Select-DailyRootFolderInteractive -Owner $form)}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'CAM Agent - Основна папка')|Out-Null}})

$autoMixRunButton = New-Object System.Windows.Forms.Button
$autoMixRunButton.Text = 'Ръчно: Копирай PGMX + Създай Mixx'
$autoMixRunButton.Location = New-Object System.Drawing.Point(635,31)
$autoMixRunButton.Size = New-Object System.Drawing.Size(260,32)
$trainingTop.Controls.Add($autoMixRunButton)

$autoMixOpenButton = New-Object System.Windows.Forms.Button
$autoMixOpenButton.Text = 'Отвори дневната папка'
$autoMixOpenButton.Location = New-Object System.Drawing.Point(905,31)
$autoMixOpenButton.Size = New-Object System.Drawing.Size(175,32)
$trainingTop.Controls.Add($autoMixOpenButton)

$autoMixTestCheck = New-Object System.Windows.Forms.CheckBox
$autoMixTestCheck.Text = 'ТЕСТ: използвай наличните .pgmx в дневната папка'
$autoMixTestCheck.AutoSize = $true
$autoMixTestCheck.Location = New-Object System.Drawing.Point(8,68)
$trainingTop.Controls.Add($autoMixTestCheck)

$autoMixCountLabel = New-Object System.Windows.Forms.Label
$autoMixCountLabel.Text = 'Нови PGMX от текущото дневно изпълнение: -'
$autoMixCountLabel.AutoSize = $true
$autoMixCountLabel.Location = New-Object System.Drawing.Point(390,69)
$trainingTop.Controls.Add($autoMixCountLabel)

$autoMixStatusLabel = New-Object System.Windows.Forms.Label
$autoMixStatusLabel.Text = 'Реален режим: след Стъпка 3 се създава актуален Mix(дд.ММ).mixx, след което автоматично стартира PowerSet, почиства XLS и го копира в detail_time.'
$autoMixStatusLabel.AutoSize = $false
$autoMixStatusLabel.Location = New-Object System.Drawing.Point(8,92)
$autoMixStatusLabel.Size = New-Object System.Drawing.Size(1240,22)
$autoMixStatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
$trainingTop.Controls.Add($autoMixStatusLabel)

$autoMixSeparator = New-Object System.Windows.Forms.Label
$autoMixSeparator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$autoMixSeparator.AutoSize = $false
$autoMixSeparator.Location = New-Object System.Drawing.Point(8,119)
$autoMixSeparator.Size = New-Object System.Drawing.Size(1240,2)
$trainingTop.Controls.Add($autoMixSeparator)


$trainingTaskLabel = New-Object System.Windows.Forms.Label
$trainingTaskLabel.Text = 'Стъпка 4 - какво обучаваме:'
$trainingTaskLabel.AutoSize = $true
$trainingTaskLabel.Location = New-Object System.Drawing.Point(8,132)
$trainingTop.Controls.Add($trainingTaskLabel)

$trainingTaskBox = New-Object System.Windows.Forms.TextBox
$trainingTaskBox.Text = 'Създаване на Mix в Maestro'
$trainingTaskBox.Location = New-Object System.Drawing.Point(125,129)
$trainingTaskBox.Size = New-Object System.Drawing.Size(410,26)
$trainingTop.Controls.Add($trainingTaskBox)

$trainingStartButton = New-Object System.Windows.Forms.Button
$trainingStartButton.Text = 'Старт обучение'
$trainingStartButton.Location = New-Object System.Drawing.Point(548,127)
$trainingStartButton.Size = New-Object System.Drawing.Size(145,30)
$trainingTop.Controls.Add($trainingStartButton)

$trainingStopButton = New-Object System.Windows.Forms.Button
$trainingStopButton.Text = 'Стоп + ZIP'
$trainingStopButton.Location = New-Object System.Drawing.Point(701,127)
$trainingStopButton.Size = New-Object System.Drawing.Size(120,30)
$trainingStopButton.Enabled = $false
$trainingTop.Controls.Add($trainingStopButton)

$trainingScreenshotCheck = New-Object System.Windows.Forms.CheckBox
$trainingScreenshotCheck.Text = 'Снимка при действие (всички монитори)'
$trainingScreenshotCheck.Checked = $true
$trainingScreenshotCheck.AutoSize = $true
$trainingScreenshotCheck.Location = New-Object System.Drawing.Point(835,132)
$trainingTop.Controls.Add($trainingScreenshotCheck)

$trainingOpenFolderButton = New-Object System.Windows.Forms.Button
$trainingOpenFolderButton.Text = 'Отвори обученията'
$trainingOpenFolderButton.Location = New-Object System.Drawing.Point(1010,127)
$trainingOpenFolderButton.Size = New-Object System.Drawing.Size(155,30)
$trainingTop.Controls.Add($trainingOpenFolderButton)

$trainingStatusLabel = New-Object System.Windows.Forms.Label
$trainingStatusLabel.Text = 'Обучение: режимът само наблюдава. Автоматичният бутон отгоре управлява Maestro в контролиран тест/дневен режим.'
$trainingStatusLabel.AutoSize = $false
$trainingStatusLabel.Location = New-Object System.Drawing.Point(8,165)
$trainingStatusLabel.Size = New-Object System.Drawing.Size(1240,22)
$trainingStatusLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$trainingTop.Controls.Add($trainingStatusLabel)

$trainingSessionLabel = New-Object System.Windows.Forms.Label
$trainingSessionLabel.Text = 'Сесия:'
$trainingSessionLabel.AutoSize = $true
$trainingSessionLabel.Location = New-Object System.Drawing.Point(8,194)
$trainingTop.Controls.Add($trainingSessionLabel)

$trainingSessionPathBox = New-Object System.Windows.Forms.TextBox
$trainingSessionPathBox.Location = New-Object System.Drawing.Point(58,191)
$trainingSessionPathBox.Size = New-Object System.Drawing.Size(585,26)
$trainingSessionPathBox.ReadOnly = $true
$trainingTop.Controls.Add($trainingSessionPathBox)

$trainingZipLabel = New-Object System.Windows.Forms.Label
$trainingZipLabel.Text = 'ZIP:'
$trainingZipLabel.AutoSize = $true
$trainingZipLabel.Location = New-Object System.Drawing.Point(655,194)
$trainingTop.Controls.Add($trainingZipLabel)

$trainingZipPathBox = New-Object System.Windows.Forms.TextBox
$trainingZipPathBox.Location = New-Object System.Drawing.Point(690,191)
$trainingZipPathBox.Size = New-Object System.Drawing.Size(558,26)
$trainingZipPathBox.ReadOnly = $true
$trainingTop.Controls.Add($trainingZipPathBox)

$trainingHint = New-Object System.Windows.Forms.Label
$trainingHint.Text = 'Записва: активен прозорец, кликове, UI елемент под курсора, специални клавиши/shortcut-и и снимка. НЕ записва написаните букви/цифри. Бележките се въвеждат в чата отдолу.'
$trainingHint.AutoSize = $false
$trainingHint.Location = New-Object System.Drawing.Point(8,218)
$trainingHint.Size = New-Object System.Drawing.Size(1240,20)
$trainingHint.ForeColor = [System.Drawing.Color]::DimGray
$trainingTop.Controls.Add($trainingHint)

$script:TrainingEventTable = New-DataTable -Columns @('#','Час','Тип','Приложение','Действие / елемент')
$trainingEventsGrid = New-Object System.Windows.Forms.DataGridView
$trainingEventsGrid.Dock = 'Fill'
$trainingEventsGrid.ReadOnly = $true
$trainingEventsGrid.AllowUserToAddRows = $false
$trainingEventsGrid.AllowUserToDeleteRows = $false
$trainingEventsGrid.AutoSizeColumnsMode = 'Fill'
$trainingEventsGrid.SelectionMode = 'FullRowSelect'
$trainingEventsGrid.DataSource = $script:TrainingEventTable
$trainingLayout.Controls.Add($trainingEventsGrid,0,1)
try {
    $trainingEventsGrid.Columns['#'].FillWeight = 10
    $trainingEventsGrid.Columns['Час'].FillWeight = 18
    $trainingEventsGrid.Columns['Тип'].FillWeight = 18
    $trainingEventsGrid.Columns['Приложение'].FillWeight = 24
    $trainingEventsGrid.Columns['Действие / елемент'].FillWeight = 130
} catch {}

$trainingChatPanel = New-Object System.Windows.Forms.TableLayoutPanel
$trainingChatPanel.Dock = 'Fill'
$trainingChatPanel.RowCount = 2
$trainingChatPanel.ColumnCount = 2
[void]$trainingChatPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 65)))
[void]$trainingChatPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 35)))
[void]$trainingChatPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 86)))
[void]$trainingChatPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 14)))
$trainingLayout.Controls.Add($trainingChatPanel,0,2)

$trainingChatBox = New-Object System.Windows.Forms.TextBox
$trainingChatBox.Multiline = $true
$trainingChatBox.ReadOnly = $true
$trainingChatBox.ScrollBars = 'Vertical'
$trainingChatBox.Dock = 'Fill'
$trainingChatBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
$trainingChatBox.Text = 'Агент: Натисни Старт обучение, премини в Maestro и работи както обикновено.' + [Environment]::NewLine
$trainingChatPanel.Controls.Add($trainingChatBox,0,0)
$trainingChatPanel.SetColumnSpan($trainingChatBox,2)

$trainingNoteInput = New-Object System.Windows.Forms.TextBox
$trainingNoteInput.Multiline = $true
$trainingNoteInput.ScrollBars = 'Vertical'
$trainingNoteInput.Dock = 'Fill'
$trainingChatPanel.Controls.Add($trainingNoteInput,0,1)

$trainingAddNoteButton = New-Object System.Windows.Forms.Button
$trainingAddNoteButton.Text = 'Добави бележка'
$trainingAddNoteButton.Dock = 'Fill'
$trainingAddNoteButton.Enabled = $false
$trainingChatPanel.Controls.Add($trainingAddNoteButton,1,1)

$trainingTimer = New-Object System.Windows.Forms.Timer
$trainingTimer.Interval = 60
$trainingTimer.Add_Tick({ Training-TimerTick })

# Keep stable references to the Step 4 training controls. The generic recorder functions
# are reused by Step 5 PowerSet, but only one training session can be active at a time.
$script:Step4TrainingUi = @{
    TaskBox = $trainingTaskBox
    StartButton = $trainingStartButton
    StopButton = $trainingStopButton
    ScreenshotCheck = $trainingScreenshotCheck
    StatusLabel = $trainingStatusLabel
    SessionPathBox = $trainingSessionPathBox
    ZipPathBox = $trainingZipPathBox
    ChatBox = $trainingChatBox
    NoteInput = $trainingNoteInput
    AddNoteButton = $trainingAddNoteButton
    EventTable = $script:TrainingEventTable
}

# -------------------- STEP 5: POWERSET TRAINING --------------------
$powerSetLayout = New-Object System.Windows.Forms.TableLayoutPanel
$powerSetLayout.Dock = 'Fill'
$powerSetLayout.RowCount = 3
$powerSetLayout.ColumnCount = 1
[void]$powerSetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 250)))
[void]$powerSetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 58)))
[void]$powerSetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 42)))
$tabPowerSet.Controls.Add($powerSetLayout)

$powerSetTop = New-Object System.Windows.Forms.Panel
$powerSetTop.Dock = 'Fill'
$powerSetLayout.Controls.Add($powerSetTop,0,0)

$powerSetTitle = New-Object System.Windows.Forms.Label
$powerSetTitle.Text = 'СТЪПКА 5 - ОБУЧЕНИЕ НА POWERSET'
$powerSetTitle.Font = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$powerSetTitle.AutoSize = $true
$powerSetTitle.Location = New-Object System.Drawing.Point(8,8)
$powerSetTop.Controls.Add($powerSetTitle)


$powerSetAutoRunButton = New-Object System.Windows.Forms.Button
$powerSetAutoRunButton.Text = 'ТЕСТ САМО СТЪПКА 5: PowerSet до XLS + обработка'
$powerSetAutoRunButton.Location = New-Object System.Drawing.Point(965,5)
$powerSetAutoRunButton.Size = New-Object System.Drawing.Size(280,30)
$powerSetTop.Controls.Add($powerSetAutoRunButton)

$powerSetAutoStatusLabel = New-Object System.Windows.Forms.Label
$powerSetAutoStatusLabel.Text = 'V5.1.38: след Стъпка 4 PowerSet стартира автоматично; XLS се почиства и копира в Z:\30_Biesse_edging\scm\detail_time.'
$powerSetAutoStatusLabel.AutoSize = $false
$powerSetAutoStatusLabel.Location = New-Object System.Drawing.Point(700,38)
$powerSetAutoStatusLabel.Size = New-Object System.Drawing.Size(545,40)
$powerSetAutoStatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
$powerSetTop.Controls.Add($powerSetAutoStatusLabel)

$powerSetInfo = New-Object System.Windows.Forms.Label
$powerSetInfo.Text = 'Обучението остава налично. TEST бутонът изпълнява потвърдения процес до готов дневен XLS.'
$powerSetInfo.AutoSize = $false
$powerSetInfo.Location = New-Object System.Drawing.Point(8,36)
$powerSetInfo.Size = New-Object System.Drawing.Size(675,38)
$powerSetInfo.ForeColor = [System.Drawing.Color]::DarkBlue
$powerSetTop.Controls.Add($powerSetInfo)

$powerSetTrainingTaskLabel = New-Object System.Windows.Forms.Label
$powerSetTrainingTaskLabel.Text = 'Стъпка 5 - какво обучаваме:'
$powerSetTrainingTaskLabel.AutoSize = $true
$powerSetTrainingTaskLabel.Location = New-Object System.Drawing.Point(8,70)
$powerSetTop.Controls.Add($powerSetTrainingTaskLabel)

$powerSetTrainingTaskBox = New-Object System.Windows.Forms.TextBox
$powerSetTrainingTaskBox.Text = 'PowerSet - зареждане на дневния Mixx, изчисляване и запис на отчета'
$powerSetTrainingTaskBox.Location = New-Object System.Drawing.Point(170,67)
$powerSetTrainingTaskBox.Size = New-Object System.Drawing.Size(520,26)
$powerSetTop.Controls.Add($powerSetTrainingTaskBox)

$powerSetTrainingStartButton = New-Object System.Windows.Forms.Button
$powerSetTrainingStartButton.Text = 'Старт обучение'
$powerSetTrainingStartButton.Location = New-Object System.Drawing.Point(705,65)
$powerSetTrainingStartButton.Size = New-Object System.Drawing.Size(145,30)
$powerSetTop.Controls.Add($powerSetTrainingStartButton)

$powerSetTrainingStopButton = New-Object System.Windows.Forms.Button
$powerSetTrainingStopButton.Text = 'Стоп + ZIP'
$powerSetTrainingStopButton.Location = New-Object System.Drawing.Point(858,65)
$powerSetTrainingStopButton.Size = New-Object System.Drawing.Size(120,30)
$powerSetTrainingStopButton.Enabled = $false
$powerSetTop.Controls.Add($powerSetTrainingStopButton)

$powerSetTrainingScreenshotCheck = New-Object System.Windows.Forms.CheckBox
$powerSetTrainingScreenshotCheck.Text = 'Снимка при действие (всички монитори)'
$powerSetTrainingScreenshotCheck.Checked = $true
$powerSetTrainingScreenshotCheck.AutoSize = $true
$powerSetTrainingScreenshotCheck.Location = New-Object System.Drawing.Point(990,70)
$powerSetTop.Controls.Add($powerSetTrainingScreenshotCheck)

$powerSetTrainingOpenFolderButton = New-Object System.Windows.Forms.Button
$powerSetTrainingOpenFolderButton.Text = 'Отвори обученията'
$powerSetTrainingOpenFolderButton.Location = New-Object System.Drawing.Point(8,105)
$powerSetTrainingOpenFolderButton.Size = New-Object System.Drawing.Size(155,30)
$powerSetTop.Controls.Add($powerSetTrainingOpenFolderButton)

$powerSetTrainingStatusLabel = New-Object System.Windows.Forms.Label
$powerSetTrainingStatusLabel.Text = 'Обучение: готово. Стартирай запис, премини в PowerSet и изпълни процеса точно както го правиш обикновено.'
$powerSetTrainingStatusLabel.AutoSize = $false
$powerSetTrainingStatusLabel.Location = New-Object System.Drawing.Point(175,112)
$powerSetTrainingStatusLabel.Size = New-Object System.Drawing.Size(1070,22)
$powerSetTrainingStatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
$powerSetTop.Controls.Add($powerSetTrainingStatusLabel)

$powerSetTrainingSessionLabel = New-Object System.Windows.Forms.Label
$powerSetTrainingSessionLabel.Text = 'Сесия:'
$powerSetTrainingSessionLabel.AutoSize = $true
$powerSetTrainingSessionLabel.Location = New-Object System.Drawing.Point(8,151)
$powerSetTop.Controls.Add($powerSetTrainingSessionLabel)

$powerSetTrainingSessionPathBox = New-Object System.Windows.Forms.TextBox
$powerSetTrainingSessionPathBox.Location = New-Object System.Drawing.Point(58,148)
$powerSetTrainingSessionPathBox.Size = New-Object System.Drawing.Size(585,26)
$powerSetTrainingSessionPathBox.ReadOnly = $true
$powerSetTop.Controls.Add($powerSetTrainingSessionPathBox)

$powerSetTrainingZipLabel = New-Object System.Windows.Forms.Label
$powerSetTrainingZipLabel.Text = 'ZIP:'
$powerSetTrainingZipLabel.AutoSize = $true
$powerSetTrainingZipLabel.Location = New-Object System.Drawing.Point(655,151)
$powerSetTop.Controls.Add($powerSetTrainingZipLabel)

$powerSetTrainingZipPathBox = New-Object System.Windows.Forms.TextBox
$powerSetTrainingZipPathBox.Location = New-Object System.Drawing.Point(690,148)
$powerSetTrainingZipPathBox.Size = New-Object System.Drawing.Size(558,26)
$powerSetTrainingZipPathBox.ReadOnly = $true
$powerSetTop.Controls.Add($powerSetTrainingZipPathBox)

$powerSetTrainingHint = New-Object System.Windows.Forms.Label
$powerSetTrainingHint.Text = 'Записва двата монитора, активен прозорец, кликове, UI елемент и shortcut-и. Не записва буквите/цифрите като keylogger. Важните правила ги пиши като бележка в чата.'
$powerSetTrainingHint.AutoSize = $false
$powerSetTrainingHint.Location = New-Object System.Drawing.Point(8,184)
$powerSetTrainingHint.Size = New-Object System.Drawing.Size(1240,38)
$powerSetTrainingHint.ForeColor = [System.Drawing.Color]::DimGray
$powerSetTop.Controls.Add($powerSetTrainingHint)

$script:PowerSetTrainingEventTable = New-DataTable -Columns @('#','Час','Тип','Приложение','Действие / елемент')
$powerSetTrainingEventsGrid = New-Object System.Windows.Forms.DataGridView
$powerSetTrainingEventsGrid.Dock = 'Fill'
$powerSetTrainingEventsGrid.ReadOnly = $true
$powerSetTrainingEventsGrid.AllowUserToAddRows = $false
$powerSetTrainingEventsGrid.AllowUserToDeleteRows = $false
$powerSetTrainingEventsGrid.AutoSizeColumnsMode = 'Fill'
$powerSetTrainingEventsGrid.SelectionMode = 'FullRowSelect'
$powerSetTrainingEventsGrid.DataSource = $script:PowerSetTrainingEventTable
$powerSetLayout.Controls.Add($powerSetTrainingEventsGrid,0,1)
try {
    $powerSetTrainingEventsGrid.Columns['#'].FillWeight = 10
    $powerSetTrainingEventsGrid.Columns['Час'].FillWeight = 18
    $powerSetTrainingEventsGrid.Columns['Тип'].FillWeight = 18
    $powerSetTrainingEventsGrid.Columns['Приложение'].FillWeight = 24
    $powerSetTrainingEventsGrid.Columns['Действие / елемент'].FillWeight = 130
} catch {}

$powerSetTrainingChatPanel = New-Object System.Windows.Forms.TableLayoutPanel
$powerSetTrainingChatPanel.Dock = 'Fill'
$powerSetTrainingChatPanel.RowCount = 2
$powerSetTrainingChatPanel.ColumnCount = 2
[void]$powerSetTrainingChatPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 65)))
[void]$powerSetTrainingChatPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 35)))
[void]$powerSetTrainingChatPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 86)))
[void]$powerSetTrainingChatPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 14)))
$powerSetLayout.Controls.Add($powerSetTrainingChatPanel,0,2)

$powerSetTrainingChatBox = New-Object System.Windows.Forms.TextBox
$powerSetTrainingChatBox.Multiline = $true
$powerSetTrainingChatBox.ReadOnly = $true
$powerSetTrainingChatBox.ScrollBars = 'Vertical'
$powerSetTrainingChatBox.Dock = 'Fill'
$powerSetTrainingChatBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
$powerSetTrainingChatBox.Text = 'Агент: Натисни Старт обучение, премини в PowerSet и работи както обикновено.' + [Environment]::NewLine
$powerSetTrainingChatPanel.Controls.Add($powerSetTrainingChatBox,0,0)
$powerSetTrainingChatPanel.SetColumnSpan($powerSetTrainingChatBox,2)

$powerSetTrainingNoteInput = New-Object System.Windows.Forms.TextBox
$powerSetTrainingNoteInput.Multiline = $true
$powerSetTrainingNoteInput.ScrollBars = 'Vertical'
$powerSetTrainingNoteInput.Dock = 'Fill'
$powerSetTrainingChatPanel.Controls.Add($powerSetTrainingNoteInput,0,1)

$powerSetTrainingAddNoteButton = New-Object System.Windows.Forms.Button
$powerSetTrainingAddNoteButton.Text = 'Добави бележка'
$powerSetTrainingAddNoteButton.Dock = 'Fill'
$powerSetTrainingAddNoteButton.Enabled = $false
$powerSetTrainingChatPanel.Controls.Add($powerSetTrainingAddNoteButton,1,1)

$script:PowerSetTrainingUi = @{
    TaskBox = $powerSetTrainingTaskBox
    StartButton = $powerSetTrainingStartButton
    StopButton = $powerSetTrainingStopButton
    ScreenshotCheck = $powerSetTrainingScreenshotCheck
    StatusLabel = $powerSetTrainingStatusLabel
    SessionPathBox = $powerSetTrainingSessionPathBox
    ZipPathBox = $powerSetTrainingZipPathBox
    ChatBox = $powerSetTrainingChatBox
    NoteInput = $powerSetTrainingNoteInput
    AddNoteButton = $powerSetTrainingAddNoteButton
    EventTable = $script:PowerSetTrainingEventTable
}

function Set-TrainingUiContext {
    param([ValidateSet('Step4','PowerSet')][string]$Mode)
    $ctx = if ($Mode -eq 'PowerSet') { $script:PowerSetTrainingUi } else { $script:Step4TrainingUi }
    $script:TrainingUiMode = $Mode
    $script:trainingTaskBox = $ctx.TaskBox
    $script:trainingStartButton = $ctx.StartButton
    $script:trainingStopButton = $ctx.StopButton
    $script:trainingScreenshotCheck = $ctx.ScreenshotCheck
    $script:trainingStatusLabel = $ctx.StatusLabel
    $script:trainingSessionPathBox = $ctx.SessionPathBox
    $script:trainingZipPathBox = $ctx.ZipPathBox
    $script:trainingChatBox = $ctx.ChatBox
    $script:trainingNoteInput = $ctx.NoteInput
    $script:trainingAddNoteButton = $ctx.AddNoteButton
    $script:TrainingEventTable = $ctx.EventTable
}

# Default context remains Step 4 until a PowerSet training is explicitly started.
Set-TrainingUiContext -Mode 'Step4'

$allowedLayout = New-Object System.Windows.Forms.TableLayoutPanel
$allowedLayout.Dock = 'Fill'
$allowedLayout.RowCount = 2
$allowedLayout.ColumnCount = 1
[void]$allowedLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 86)))
[void]$allowedLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabAllowed.Controls.Add($allowedLayout)

$allowedPanel = New-Object System.Windows.Forms.Panel
$allowedPanel.Dock = 'Fill'
$allowedLayout.Controls.Add($allowedPanel, 0, 0)

$addAllowedButton = New-Object System.Windows.Forms.Button
$addAllowedButton.Text = 'Добави папка'
$addAllowedButton.Location = New-Object System.Drawing.Point(8, 8)
$addAllowedButton.Size = New-Object System.Drawing.Size(130, 30)
$allowedPanel.Controls.Add($addAllowedButton)

$removeAllowedButton = New-Object System.Windows.Forms.Button
$removeAllowedButton.Text = 'Премахни'
$removeAllowedButton.Location = New-Object System.Drawing.Point(146, 8)
$removeAllowedButton.Size = New-Object System.Drawing.Size(110, 30)
$allowedPanel.Controls.Add($removeAllowedButton)

$saveAllowedButton = New-Object System.Windows.Forms.Button
$saveAllowedButton.Text = 'Запиши списъка'
$saveAllowedButton.Location = New-Object System.Drawing.Point(264, 8)
$saveAllowedButton.Size = New-Object System.Drawing.Size(135, 30)
$allowedPanel.Controls.Add($saveAllowedButton)

$reindexButton = New-Object System.Windows.Forms.Button
$reindexButton.Text = 'Обнови каталог (фон)'
$reindexButton.Location = New-Object System.Drawing.Point(407, 8)
$reindexButton.Size = New-Object System.Drawing.Size(150, 30)
$allowedPanel.Controls.Add($reindexButton)

$resetAllowedButton = New-Object System.Windows.Forms.Button
$resetAllowedButton.Text = 'Възстанови 8-те'
$resetAllowedButton.Location = New-Object System.Drawing.Point(565, 8)
$resetAllowedButton.Size = New-Object System.Drawing.Size(145, 30)
$allowedPanel.Controls.Add($resetAllowedButton)

$allowedCountLabel = New-Object System.Windows.Forms.Label
$allowedCountLabel.Text = 'Разрешени папки: 0'
$allowedCountLabel.AutoSize = $true
$allowedCountLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$allowedCountLabel.Location = New-Object System.Drawing.Point(725, 15)
$allowedPanel.Controls.Add($allowedCountLabel)

$allowedInfo = New-Object System.Windows.Forms.Label
$allowedInfo.Text = 'Обхождат се само тези папки и всички техни подпапки. Файлове НЕ се отварят при каталогиране.'
$allowedInfo.AutoSize = $true
$allowedInfo.Location = New-Object System.Drawing.Point(8, 52)
$allowedPanel.Controls.Add($allowedInfo)

$allowedList = New-Object System.Windows.Forms.ListBox
$allowedList.Dock = 'Fill'
$allowedList.HorizontalScrollbar = $true
$allowedLayout.Controls.Add($allowedList, 0, 1)

$indexLayout = New-Object System.Windows.Forms.TableLayoutPanel
$indexLayout.Dock = 'Fill'
$indexLayout.RowCount = 2
$indexLayout.ColumnCount = 1
[void]$indexLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
[void]$indexLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabIndex.Controls.Add($indexLayout)

$indexPanel = New-Object System.Windows.Forms.Panel
$indexPanel.Dock = 'Fill'
$indexLayout.Controls.Add($indexPanel, 0, 0)

$indexInfoLabel = New-Object System.Windows.Forms.Label
$indexInfoLabel.Text = 'Каталогът още не е зареден.'
$indexInfoLabel.AutoSize = $true
$indexInfoLabel.Location = New-Object System.Drawing.Point(8, 12)
$indexPanel.Controls.Add($indexInfoLabel)

$folderIndexGrid = New-Object System.Windows.Forms.DataGridView
$folderIndexGrid.Dock = 'Fill'
$folderIndexGrid.ReadOnly = $true
$folderIndexGrid.AllowUserToAddRows = $false
$folderIndexGrid.AllowUserToDeleteRows = $false
$folderIndexGrid.AutoSizeColumnsMode = 'Fill'
$folderIndexGrid.SelectionMode = 'FullRowSelect'
$indexLayout.Controls.Add($folderIndexGrid, 0, 1)

$warningsBox = New-Object System.Windows.Forms.TextBox
$warningsBox.Dock = 'Fill'
$warningsBox.Multiline = $true
$warningsBox.ScrollBars = 'Both'
$warningsBox.ReadOnly = $true
$warningsBox.Font = New-Object System.Drawing.Font('Consolas', 10)
$tabWarnings.Controls.Add($warningsBox)

function Update-SelectedCsvProducts {
    if ($null -eq $fileSummaryGrid.CurrentRow) { return }
    $csvName = Safe-Trim $fileSummaryGrid.CurrentRow.Cells['CSV'].Value
    if ([string]::IsNullOrWhiteSpace($csvName)) { return }

    $table = New-DataTable -Columns @('Изделие', 'Редове', 'Физически детайли', 'ProgramCodeF', 'Без ProgramCodeF', 'Двойни обработки')
    $rowsForCsv = @($script:LastAllRows | Where-Object { $_.CsvFile -eq $csvName })
    $groups = @($rowsForCsv | Group-Object Product | Sort-Object Name)
    foreach ($g in $groups) {
        $grpRows = @($g.Group)
        $physical = @($grpRows | ForEach-Object { "$($_.DetailNo)|$($_.DetailName)|$($_.Length)|$($_.Width)|$($_.Thickness)" } | Sort-Object -Unique).Count
        $programs = @($grpRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.ProgramCodeF) }).Count
        $noProgram = @($grpRows | Where-Object { [string]::IsNullOrWhiteSpace($_.ProgramCodeF) }).Count
        $doubleCount = @($grpRows | Group-Object { "$($_.DetailNo)|$($_.DetailName)|$($_.Length)|$($_.Width)|$($_.Thickness)" } | Where-Object { $_.Count -gt 1 }).Count
        $r = $table.NewRow()
        $r['Изделие'] = Safe-String $g.Name
        $r['Редове'] = [string]$grpRows.Count
        $r['Физически детайли'] = [string]$physical
        $r['ProgramCodeF'] = [string]$programs
        $r['Без ProgramCodeF'] = [string]$noProgram
        $r['Двойни обработки'] = [string]$doubleCount
        [void]$table.Rows.Add($r)
    }
    $fileProductsGrid.DataSource = $table
}

$fileSummaryGrid.Add_SelectionChanged({
    try { Update-SelectedCsvProducts } catch {}
})


$autoMixOpenButton.Add_Click({
    try {
        $daily=Get-DailyMixFolder
        if(-not [System.IO.Directory]::Exists($daily)){[void][System.IO.Directory]::CreateDirectory($daily)}
        Start-Process -FilePath 'explorer.exe' -ArgumentList @($daily)
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent - Mixx',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$autoMixRunButton.Add_Click({
    if($script:TrainingActive){
        [System.Windows.Forms.MessageBox]::Show('Спрете режима за обучение преди автоматичния Mixx.','PWX500 Agent - Mixx',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $autoMixRunButton.Enabled=$false
    try {
        $daily=Get-DailyMixFolder
        if(-not [System.IO.Directory]::Exists($daily)){[void][System.IO.Directory]::CreateDirectory($daily)}
        $target=Get-DailyMixFilePath
        if([System.IO.File]::Exists($target)){
            [System.IO.File]::Delete($target)
            Write-MixLog ('Deleted existing Mixx automatically for rebuild: '+$target)
            $autoMixStatusLabel.Text='Съществуващият дневен Mixx е премахнат. Създавам го наново...'
            [System.Windows.Forms.Application]::DoEvents()
        }

        $testMode=$autoMixTestCheck.Checked
        if($testMode){
            $confirm=[System.Windows.Forms.MessageBox]::Show('ТЕСТ режим: ще се използват всички .pgmx, които вече са в днешната папка. Няма да се взема списъкът от Стъпка 3. Продължаваме ли?','PWX500 Agent - Mixx TEST',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
            if($confirm -ne [System.Windows.Forms.DialogResult]::Yes){return}
        }

        $autoMixStatusLabel.Text='Подготвям дневните PGMX...'
        $status.Text='Стъпка 4: подготвям дневните PGMX...'
        [System.Windows.Forms.Application]::DoEvents()
        $inputs=@(Prepare-DailyMixFiles -UseExistingTestFiles:$testMode)
        $autoMixStatusLabel.Text=('PGMX за Mixx: '+$inputs.Count+'. Стартирам Maestro...')
        [System.Windows.Forms.Application]::DoEvents()
        $mixPath=Invoke-DailyMixAutomation -PgmxFiles $inputs
        $autoMixStatusLabel.Text='ГОТОВО: '+$mixPath
        $status.Text='Стъпка 4 ГОТОВО: '+$mixPath
        Update-AutoMixUi
        [System.Windows.Forms.MessageBox]::Show("Дневният Mixx е създаден успешно.`r`n`r`nPGMX: $($inputs.Count)`r`nMixx:`r`n$mixPath",'PWX500 Agent - Стъпка 4',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
        $msg=$_.Exception.Message
        if ($script:LastMaestroLauncher) { $msg += ' | Maestro source: ' + $script:LastMaestroLauncher }
        Write-MixLog ('AUTO ERROR: '+$msg)
        $autoMixStatusLabel.Text='ГРЕШКА: '+$msg
        $status.Text='Стъпка 4 ГРЕШКА: '+$msg
        [System.Windows.Forms.MessageBox]::Show($msg,'PWX500 Agent - Стъпка 4 Mixx',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
        $autoMixRunButton.Enabled=$true
    }
})

Update-AutoMixUi


$trainingStartButton.Add_Click({
    if ($script:TrainingActive) { return }
    Set-TrainingUiContext -Mode 'Step4'
    Start-TrainingSession
})

$trainingStopButton.Add_Click({
    Set-TrainingUiContext -Mode 'Step4'
    Stop-TrainingSession
})

$trainingAddNoteButton.Add_Click({
    Set-TrainingUiContext -Mode 'Step4'
    if (-not $script:TrainingActive) { return }
    $note = Safe-Trim $trainingNoteInput.Text
    if ([string]::IsNullOrWhiteSpace($note)) { return }
    $trainingChatBox.AppendText(('[' + (Get-Date -Format 'HH:mm:ss') + '] Ти: ' + $note + "`r`n"))
    Add-TrainingEvent -Type 'NOTE' -Action 'User training note' -Note $note
    $trainingChatBox.AppendText("Агент: записано.`r`n")
    $trainingNoteInput.Clear()
    $trainingNoteInput.Focus()
})

$trainingNoteInput.Add_KeyDown({
    param($sender,$e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $trainingAddNoteButton.PerformClick()
    }
})

$trainingOpenFolderButton.Add_Click({
    try {
        if (-not [System.IO.Directory]::Exists($script:TrainingRoot)) { [void][System.IO.Directory]::CreateDirectory($script:TrainingRoot) }
        Start-Process -FilePath 'explorer.exe' -ArgumentList @($script:TrainingRoot)
    } catch {}
})


$powerSetAutoRunButton.Add_Click({
    if ($script:TrainingActive) {
        [System.Windows.Forms.MessageBox]::Show('Спри активното обучение преди PowerSet TEST.','PWX500 Agent - PowerSet',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $powerSetAutoRunButton.Enabled = $false
    try {
        $xls = Invoke-PowerSetFullAutomation
        [System.Windows.Forms.MessageBox]::Show(
            "PowerSet процесът приключи успешно.`r`n`r`nПочистен дневен отчет:`r`n$xls`r`n`r`nКопие:`r`nZ:\30_Biesse_edging\scm\detail_time\$([System.IO.Path]::GetFileName($xls))",
            'PWX500 Agent V5.1.38 - PowerSet TEST',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        $msg = $_.Exception.Message
        Write-PowerSetLog ('ERROR: ' + $msg)
        Set-PowerSetTestStatus ('PowerSet TEST ГРЕШКА: ' + $msg)
        [System.Windows.Forms.MessageBox]::Show($msg,'PWX500 Agent V5.1.38 - PowerSet TEST',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
        $powerSetAutoRunButton.Enabled = $true
    }
})

$powerSetTrainingStartButton.Add_Click({
    if ($script:TrainingActive) {
        [System.Windows.Forms.MessageBox]::Show('Вече има активно обучение. Първо го спри със Стоп + ZIP.','PWX500 Agent - PowerSet обучение',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    Set-TrainingUiContext -Mode 'PowerSet'
    Start-TrainingSession
})

$powerSetTrainingStopButton.Add_Click({
    Set-TrainingUiContext -Mode 'PowerSet'
    Stop-TrainingSession
})

$powerSetTrainingAddNoteButton.Add_Click({
    Set-TrainingUiContext -Mode 'PowerSet'
    if (-not $script:TrainingActive) { return }
    $note = Safe-Trim $powerSetTrainingNoteInput.Text
    if ([string]::IsNullOrWhiteSpace($note)) { return }
    $powerSetTrainingChatBox.AppendText(('[' + (Get-Date -Format 'HH:mm:ss') + '] Ти: ' + $note + "`r`n"))
    Add-TrainingEvent -Type 'NOTE' -Action 'User training note' -Note $note
    $powerSetTrainingChatBox.AppendText("Агент: записано.`r`n")
    $powerSetTrainingNoteInput.Clear()
    $powerSetTrainingNoteInput.Focus()
})

$powerSetTrainingNoteInput.Add_KeyDown({
    param($sender,$e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $powerSetTrainingAddNoteButton.PerformClick()
    }
})

$powerSetTrainingOpenFolderButton.Add_Click({
    try {
        if (-not [System.IO.Directory]::Exists($script:TrainingRoot)) { [void][System.IO.Directory]::CreateDirectory($script:TrainingRoot) }
        Start-Process -FilePath 'explorer.exe' -ArgumentList @($script:TrainingRoot)
    } catch {}
})

$tabPowerSet.Add_Enter({
    if (-not $script:TrainingActive) { Set-TrainingUiContext -Mode 'PowerSet' }
})
$tabTraining.Add_Enter({
    Update-AutoMixUi
    if (-not $script:TrainingActive) { Set-TrainingUiContext -Mode 'Step4' }
})

$browseButton.Add_Click({ $pathBox.Text = $script:DefaultCsvFolder })



$changeProductFolderButton.Add_Click({
    Set-ManualProductFolderFromCandidates
})

$productCandidatesGrid.Add_CellDoubleClick({
    param($sender,$e)
    if ($e.RowIndex -ge 0) { Set-ManualProductFolderFromCandidates }
})

$openConfirmButton.Add_Click({
    Refresh-ConfirmationQueue
    $tabs.SelectedTab = $tabConfirm
})

$confirmProductList.Add_SelectedIndexChanged({
    try { Show-ConfirmationProduct } catch {}
})

$confirmSelectedButton.Add_Click({
    $p = Get-SelectedConfirmationCandidatePath
    Confirm-CurrentProductPath -Path $p -Source 'Candidate'
    try{Refresh-OperatorDashboard}catch{}
})

$confirmCandidateList.Add_DoubleClick({
    $p = Get-SelectedConfirmationCandidatePath
    if (-not [string]::IsNullOrWhiteSpace($p)) { Confirm-CurrentProductPath -Path $p -Source 'CandidateDoubleClick'; try{Refresh-OperatorDashboard}catch{} }
})

$chooseExactFolderButton.Add_Click({
    if ($null -eq $confirmProductList.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show('Няма избрано изделие.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Посочи точната папка на изделието. Ако е извън разрешените корени, ще бъде поискано изрично потвърждение.'
    $dialog.ShowNewFolderButton = $false
    $cand = Get-SelectedConfirmationCandidatePath
    if (-not [string]::IsNullOrWhiteSpace($cand) -and [System.IO.Directory]::Exists($cand)) {
        $dialog.SelectedPath = $cand
    }
    elseif ($script:AllowedFolders.Count -gt 0 -and [System.IO.Directory]::Exists($script:AllowedFolders[0])) {
        $dialog.SelectedPath = $script:AllowedFolders[0]
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Confirm-CurrentProductPath -Path $dialog.SelectedPath -Source 'Manual'
        try{Refresh-OperatorDashboard}catch{}
    }
    $dialog.Dispose()
})

$confirmContinueButton.Add_Click({
    try{Invoke-OperatorContinueConfirmed}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}
})

$skipConfirmButton.Add_Click({
    if ($confirmProductList.Items.Count -eq 0) { return }
    $i = $confirmProductList.SelectedIndex
    if ($i -lt 0) { $i = 0 } else { $i++ }
    if ($i -ge $confirmProductList.Items.Count) { $i = 0 }
    $confirmProductList.SelectedIndex = $i
})

$removeMappingButton.Add_Click({
    # This button is mainly for corrections. If a queue item is selected and has no mapping, nothing happens.
    $product = ''
    if ($confirmProductList.SelectedItem) { $product = [string]$confirmProductList.SelectedItem }
    elseif ($productCandidatesGrid.CurrentRow) { $product = Safe-Trim $productCandidatesGrid.CurrentRow.Cells['Изделие от CSV'].Value }
    if ([string]::IsNullOrWhiteSpace($product)) { return }
    if (-not $script:ProductFolderMap.ContainsKey($product)) {
        [System.Windows.Forms.MessageBox]::Show('За това изделие няма запомнена връзка.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $ans = [System.Windows.Forms.MessageBox]::Show("Да се изтрие ли запомнената връзка за:`r`n$product?", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        Remove-ProductFolderMapping -Product $product
        Refresh-CandidateGridFromSearchResults
        try{Refresh-OperatorDashboard}catch{}
    }
})

$categorySaveButton.Add_Click({
    try{Save-ProductCategoryGrid;$status.Text='Категориите са записани. Натисни „Търси изделия“ отново за новите правила.'}
    catch{[System.Windows.Forms.MessageBox]::Show("Грешка при запис на категории:`r`n$($_.Exception.Message)",'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}
})
$categoryAddButton.Add_Click({
    try{
        if($null -eq $categoryGrid.DataSource){Refresh-ProductCategoryGrid}
        $dt=[System.Data.DataTable]$categoryGrid.DataSource;$r=$dt.NewRow();$r['Phrase']='new category';$r['Targets']='';$r['Priority']='100';$r['Active']='1';[void]$dt.Rows.Add($r)
    }catch{}
})
$categoryDeleteButton.Add_Click({
    try{if($null -ne $categoryGrid.CurrentRow -and -not $categoryGrid.CurrentRow.IsNewRow){$categoryGrid.Rows.Remove($categoryGrid.CurrentRow)}}catch{}
})
$categoryResetButton.Add_Click({
    $ans=[System.Windows.Forms.MessageBox]::Show('Да се възстановят ли началните категории? Текущите промени в списъка ще бъдат заменени.','PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
    if($ans -eq [System.Windows.Forms.DialogResult]::Yes){Save-ProductCategoryRules -Rows @(Get-DefaultProductCategoryRows);Load-ProductCategoryRules;Refresh-ProductCategoryGrid}
})

# V5.1.31 operator buttons
$operatorReadCsvButton.Add_Click({try{Start-BackgroundCsvRead}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
$operatorValidateButton.Add_Click({try{Start-OperatorValidation}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
$operatorReviewButton.Add_Click({try{Refresh-ConfirmationQueue;$tabs.SelectedTab=$tabConfirm}catch{}})
$operatorContinueButton.Add_Click({try{Invoke-OperatorContinueConfirmed}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
$operatorRunAllButton.Add_Click({try{Start-OperatorFullFlow}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})
$operatorOpenReportsButton.Add_Click({try{$f=Get-DailyMixFolder;if([System.IO.Directory]::Exists($f)){Start-Process explorer.exe -ArgumentList ('"'+$f+'"')}else{[System.Windows.Forms.MessageBox]::Show('Още няма дневна папка с отчети.','PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null}}catch{}})
$operatorScheduleSaveButton.Add_Click({try{Save-OperatorScheduleSettings;$status.Text='Записан автоматичен начален час: '+$operatorScheduleTime.Value.ToString('HH:mm')}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null}})

$addAllowedButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Изберете нова разрешена папка. Агентът ще обхожда нея и всички подпапки.'
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $p = Safe-Trim $dialog.SelectedPath
        if (-not [string]::IsNullOrWhiteSpace($p) -and -not $script:AllowedFolders.Contains($p)) {
            $script:AllowedFolders.Add($p)
            Refresh-AllowedFoldersList
        }
    }
    $dialog.Dispose()
})

$removeAllowedButton.Add_Click({
    if ($null -eq $allowedList.SelectedItem) { return }
    $p = [string]$allowedList.SelectedItem
    [void]$script:AllowedFolders.Remove($p)
    Refresh-AllowedFoldersList
})

$saveAllowedButton.Add_Click({
    try {
        Save-AllowedFolders
        [System.Windows.Forms.MessageBox]::Show('Списъкът с разрешени папки е записан.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при запис: $($_.Exception.Message)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$resetAllowedButton.Add_Click({
    $answer = [System.Windows.Forms.MessageBox]::Show('Да се възстановят ли осемте начални разрешени папки?', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        Set-DefaultAllowedFolders
        Refresh-AllowedFoldersList
        try { Save-AllowedFolders } catch {}
        $status.Text = 'Възстановени са 8-те начални разрешени папки. Натиснете Обнови каталог (фон).'
    }
})

$historyMigrationButton.Add_Click({
    try {
        if([System.IO.File]::Exists($script:HistoryMarkerFile)){
            $ans=[System.Windows.Forms.MessageBox]::Show('Началната историческа миграция вече е изпълнявана. Да се повтори ли проверката? Съществуващите потвърдени връзки няма да бъдат презаписани.','PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
            if($ans -ne [System.Windows.Forms.DialogResult]::Yes){return}
        }
        $dlg=New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title='Избери „Поръчки за линия.xlsx“'
        $dlg.Filter='Excel (*.xlsx)|*.xlsx'
        $dlg.Multiselect=$false
        if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
            $wb=[string]$dlg.FileName
            $dlg.Dispose()
            Start-BackgroundHistoryMigration -WorkbookPath $wb
        } else {$dlg.Dispose()}
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при началната историческа миграция:`r`n$($_.Exception.Message)",'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
        $historyMigrationButton.Enabled=$true
    }
})

$pgmxCheckButton.Add_Click({
    try { Start-BackgroundPgmx -Mode 'Check' }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при PGMX проверка:`r`n$($_.Exception.Message)",'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
        $pgmxCheckButton.Enabled=$true;$pgmxExecuteButton.Enabled=$true
    }
})

$pgmxExecuteButton.Add_Click({
    try {
        $stats=Get-PgmxAllTestStats
        $msg="ТЕСТОВ РЕЖИМ - ВСИЧКИ ИЗДЕЛИЯ`r`n`r`nИзделия в текущите CSV: $($stats.Products)`r`nProgramCodeF редове: $($stats.Rows)`r`n`r`nЩе се обработят само случаите с потвърдена/еднозначна продуктова папка. Нееднозначните, липсващите HOP и другите проблемни случаи ще бъдат ПРОПУСНАТИ и отчетени.`r`n`r`nСъществуващи PGMX/ProgramCodeF НЕ се презаписват.`r`n`r`nСлед приключване на Стъпка 3 агентът АВТОМАТИЧНО ще копира новите <ProgramCode>.pgmx в избраната основна папка\дд.ММ, ще пресъздаде един общ Mix(дд.ММ).mixx от всички PGMX в дневната папка, след което ще стартира PowerSet, ще изчака изчислението, ще създаде и почисти Mix(дд.ММ).xls и ще го копира в Z:\30_Biesse_edging\scm\detail_time.`r`n`r`nСамо крайните ProgramCodeF на изделия със специална вложена структура ще бъдат копирани (без презаписване) и в:`r`nV:\30_Biesse_edging\scm\SCM_PWX_500`r`nНормалните изделия остават само в собствената 2_PWX_500. Основните PGMX НЕ се копират там.`r`n`r`nПродължаваме ли?"
        $ans=[System.Windows.Forms.MessageBox]::Show($msg,'PWX500 Agent - общ тест',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
        if($ans -ne [System.Windows.Forms.DialogResult]::Yes){return}
        Start-BackgroundPgmx -Mode 'Execute'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при PGMX изпълнение:`r`n$($_.Exception.Message)",'PWX500 Agent',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
        $pgmxCheckButton.Enabled=$true;$pgmxExecuteButton.Enabled=$true
    }
})

$hopCheckButton.Add_Click({
    try { Start-BackgroundHopCheck }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при стартиране на HOP проверката:`r`n$($_.Exception.Message)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $hopCheckButton.Enabled = $true
    }
})

$findProductsButton.Add_Click({
    try { Start-BackgroundProductSearch }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        $findProductsButton.Enabled = $true
    }
})

$reindexButton.Add_Click({
    try {
        Start-BackgroundFolderIndex
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при стартиране на фонов каталог:`r`n$($_.Exception.Message)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $reindexButton.Enabled = $true
    }
})

$exportButton.Add_Click({
    if ($null -eq $script:LastFileSummary -or $script:LastFileSummary.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Няма данни за експорт. Първо прочетете CSV файловете.', 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }

    $save = New-Object System.Windows.Forms.SaveFileDialog
    $save.Filter = 'CSV файл (*.csv)|*.csv'
    $save.Title = 'Запиши контролен отчет'
    $save.FileName = 'PWX500_CSV_Control_' + (Get-Date -Format 'yyyy-MM-dd_HHmm') + '.csv'
    if (-not [string]::IsNullOrWhiteSpace($script:LastFolder) -and [System.IO.Directory]::Exists($script:LastFolder)) {
        $save.InitialDirectory = $script:LastFolder
    }

    if ($save.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $lines = New-Object System.Collections.Generic.List[string]
            $headers = @()
            foreach ($c in $script:LastFileSummary.Columns) { $headers += (Convert-ToCsvField $c.ColumnName) }
            $lines.Add(($headers -join ';'))
            foreach ($row in $script:LastFileSummary.Rows) {
                $vals = @()
                foreach ($c in $script:LastFileSummary.Columns) { $vals += (Convert-ToCsvField $row[$c.ColumnName]) }
                $lines.Add(($vals -join ';'))
            }
            $utf8Bom = New-Object System.Text.UTF8Encoding($true)
            [System.IO.File]::WriteAllLines($save.FileName, $lines.ToArray(), $utf8Bom)
            [System.Windows.Forms.MessageBox]::Show("Контролният отчет е записан:`r`n$($save.FileName)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Грешка при експорт:`r`n$($_.Exception.Message)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
    $save.Dispose()
})

$scanButton.Add_Click({
    try { Start-BackgroundCsvRead }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Грешка при стартиране на фоновото CSV четене:`r`n$($_.Exception.Message)", 'PWX500 Agent', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $scanButton.Enabled = $true
    }
})

Load-AllowedFolders
Refresh-AllowedFoldersList
Load-ProductCategoryRules
Refresh-ProductCategoryGrid
Load-ProductFolderMappings
# Reuse the previous local catalog when V3.0 is extracted over V0.2.19; no server access is needed for this local catalog copy.
try {
    if (-not [System.IO.File]::Exists($script:IndexCsv) -and [System.IO.File]::Exists($script:PreviousIndexCsv)) {
        [System.IO.File]::Copy($script:PreviousIndexCsv,$script:IndexCsv,$false)
    }
} catch {}
Update-FolderIndexGrid
if ([System.IO.File]::Exists($script:HistoryMarkerFile)) {
    try {
        $historyInfoLabel.Text='Историческата миграция вече е изпълнявана. Може да я повторите при обновен файл „Поръчки за линия“; съществуващите потвърдени връзки не се презаписват.'
    } catch {}
}
$pathBox.Text = $script:DefaultCsvFolder
if ([System.IO.File]::Exists($script:IndexCsv)) {
    $idxTime = ''
    try { $idxTime = [System.IO.File]::GetLastWriteTime($script:IndexCsv).ToString('dd.MM.yyyy HH:mm') } catch {}
    $status.Text = "Фиксирана CSV папка: $($script:DefaultCsvFolder) | Локален каталог: $($script:FolderIndexCount) папки | $idxTime"
}

# V5.1.31 operator dashboard / local schedule timer
try{Load-OperatorScheduleSettings}catch{}
try{Refresh-CsvRegistryGrid;Refresh-OperatorCatalogLabel}catch{}
try{Refresh-OperatorDashboard}catch{}
try{$tabs.SelectedTab=$tabOperator}catch{}

# The old full window is now the technical/diagnostic window. The clean operator window is the default.
try { Initialize-ActiveTimeIndex } catch { $script:ActiveTimeMap=@{} }
# Carry forward excluded PGMX already isolated in earlier versions, without re-opening file content.
try { Sync-PgmxAutoExclusionsFromDailyFolders } catch { [System.Windows.Forms.MessageBox]::Show('Изолирани PGMX не могат да се регистрират: '+$_.Exception.Message,'CAM Agent - Без повторения',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null }
$form.Text='CAM Agent V5.1.38 - Технически режим / диагностика'
try{if([System.IO.File]::Exists($script:AppIconIco)){$form.Icon=New-Object System.Drawing.Icon($script:AppIconIco)}}catch{}
$script:OperatorMainClosing=$false
$form.Add_FormClosing({param($sender,$e)
    if(-not $script:OperatorMainClosing){$e.Cancel=$true;$form.Hide()}
})

[void](New-PWX500OperatorMainForm)
$operatorMainForm=$script:OperatorMainForm
if($null -eq $operatorMainForm -or -not ($operatorMainForm -is [System.Windows.Forms.Form])){ throw 'Неуспешно създаване на операторския прозорец: очакван System.Windows.Forms.Form.' }
try{$script:MainScheduleCheck.Checked=$operatorScheduleCheck.Checked;$script:MainScheduleTime.Value=$operatorScheduleTime.Value;$script:MainScheduleModeCombo.SelectedIndex=if($script:OperatorScheduleMode -eq 'FULL'){1}else{0}}catch{}
try{Refresh-OperatorDashboard;Refresh-MainOperatorUi}catch{}

$script:OperatorUiTimer=New-Object System.Windows.Forms.Timer
$script:OperatorUiTimer.Interval=3000
$script:OperatorUiTimer.Add_Tick({try{if($null -ne $operatorProcessStatusLabel){$operatorProcessStatusLabel.Text=$status.Text};Refresh-MainOperatorUi -StatusOnly;Test-OperatorScheduleTick}catch{}})
$script:OperatorUiTimer.Start()

$operatorMainForm.Add_Shown({try{Refresh-MainOperatorUi;Prompt-ResumeIfNeeded;Offer-ResumeTimeImportIfNeeded}catch{}})

$operatorMainForm.Add_FormClosing({
    $script:OperatorMainClosing=$true
    try {
        $rs=Get-OperatorRunState
        $busy=($null -ne $script:CsvWorkerProcess -or $null -ne $script:SearchWorkerProcess -or $null -ne $script:HopWorkerProcess -or $null -ne $script:PgmxWorkerProcess -or $null -ne $script:IndexWorkerProcess -or $null -ne $script:HistoryWorkerProcess -or $script:AutoMixChainRunning)
        if($busy -or ($null -ne $rs -and [string]$rs.Status -eq 'ACTIVE')){Stop-OperatorProcessSafely -Silent}
    } catch {}
    try { if($null -ne $script:OperatorUiTimer){$script:OperatorUiTimer.Stop();$script:OperatorUiTimer.Dispose()} } catch {}
    try { if($script:TrainingActive){Stop-TrainingSession -Silent} } catch {}
    try { if($script:TimeImportRunning){Stop-TimeImportBackground}; if($script:TimeImportSynchronousRunning -and $script:TimeImportSynchronousStopPath){[System.IO.File]::WriteAllText($script:TimeImportSynchronousStopPath,'STOP')} } catch {}
    try { if($null -ne $script:TimeImportTimer){$script:TimeImportTimer.Stop();$script:TimeImportTimer.Dispose()} } catch {}
    try { $form.Close() } catch {}
})

[void]$operatorMainForm.ShowDialog()
