param(
    [Parameter(Mandatory=$true)][string]$SettingsCsv,
    [Parameter(Mandatory=$true)][string]$IndexOut,
    [Parameter(Mandatory=$true)][string]$ProgressFile,
    [Parameter(Mandatory=$true)][string]$SettingsVersion
)

# SAFETY: This worker reads DIRECTORY metadata only.
# It never enumerates, opens, parses, previews or launches files.
# It stops descending when a product folder is detected by a direct 03_Mashini subfolder.
# The two explicitly excluded 01_Aksesoari trees are skipped completely, including all subfolders.
$ErrorActionPreference = 'Stop'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$excludedPrefixes = @(
    'Y:\1_Izpolzvani\01_KM\_0_10_F\01_Aksesoari',
    'Y:\1_Izpolzvani\01_KM\_06_Pac_s12_B#\01_Aksesoari'
)
function Is-ExcludedPath([string]$Path) {
    $p = $Path.TrimEnd('\')
    foreach($ex in $excludedPrefixes){
        $e = $ex.TrimEnd('\')
        if([string]::Equals($p,$e,[System.StringComparison]::OrdinalIgnoreCase)){ return $true }
        if($p.StartsWith($e+'\',[System.StringComparison]::OrdinalIgnoreCase)){ return $true }
    }
    return $false
}

function Write-ProgressText([string]$Text) {
    try { [System.IO.File]::WriteAllText($ProgressFile, $Text, $utf8NoBom) } catch {}
}
function CsvField([string]$s) {
    if ($null -eq $s) { $s = '' }
    if ($s.Contains('"')) { $s = $s.Replace('"','""') }
    if ($s.Contains(';') -or $s.Contains('"') -or $s.Contains("`r") -or $s.Contains("`n")) { return '"' + $s + '"' }
    return $s
}
function Get-Depth([string]$Root, [string]$Path) {
    $r = $Root.TrimEnd('\\')
    $p = $Path.TrimEnd('\\')
    if ($p.Length -le $r.Length) { return 0 }
    $rel = $p.Substring($r.Length).TrimStart('\\')
    if ([string]::IsNullOrWhiteSpace($rel)) { return 0 }
    return @($rel -split '\\').Count
}

try {
    Write-ProgressText 'Каталог: зареждане на разрешените папки...'
    if (-not [System.IO.File]::Exists($SettingsCsv)) { throw "Липсва файлът с разрешени папки: $SettingsCsv" }

    $roots = @()
    $rows = Import-Csv -LiteralPath $SettingsCsv -Delimiter ';'
    foreach ($r in $rows) {
        if ([string]$r.Version -ne $SettingsVersion) { continue }
        $p = ([string]$r.Path).Trim()
        if (-not [string]::IsNullOrWhiteSpace($p)) { $roots += $p }
    }
    $roots = @($roots | Sort-Object -Unique)
    if ($roots.Count -eq 0) { throw 'Няма разрешени папки за обхождане.' }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Root;Folder;Path;Depth')
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $foldersVisited = 0
    $productsFound = 0
    $errors = 0

    foreach ($root in $roots) {
        if (-not [System.IO.Directory]::Exists($root)) {
            $errors++
            Write-ProgressText "Каталог: недостъпна разрешена папка: $root"
            continue
        }

        $q = New-Object 'System.Collections.Generic.Queue[string]'
        $q.Enqueue($root)

        while ($q.Count -gt 0) {
            $dir = $q.Dequeue()
            if (Is-ExcludedPath $dir) { continue }
            if (-not $visited.Add($dir)) { continue }
            $foldersVisited++

            try {
                # V5.1.35 speed-safe scan: enumerate child directories ONCE.
                # The old code first called Directory.Exists(dir\\03_Mashini) and
                # then enumerated the same directory. On SMB this doubled many
                # metadata round-trips. The product rule itself is unchanged.
                $dirInfo = New-Object System.IO.DirectoryInfo($dir)
                $childInfos = @($dirInfo.EnumerateDirectories())
                $hasMachine = $false
                foreach ($ci in $childInfos) {
                    if ([string]::Equals([string]$ci.Name,'03_Mashini',[System.StringComparison]::OrdinalIgnoreCase)) {
                        $hasMachine = $true
                        break
                    }
                }
                if ($hasMachine) {
                    $name = [System.IO.Path]::GetFileName($dir.TrimEnd('\\'))
                    if ([string]::IsNullOrWhiteSpace($name)) { $name = $dir }
                    $depth = Get-Depth -Root $root -Path $dir
                    $lines.Add(((CsvField $root),(CsvField $name),(CsvField $dir),(CsvField ([string]$depth))) -join ';')
                    $productsFound++
                    # IMPORTANT: do not descend into product technical folders.
                    continue
                }

                foreach ($ci in $childInfos) {
                    $child = [string]$ci.FullName
                    if (Is-ExcludedPath $child) { continue }
                    try { if (($ci.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue } } catch {}
                    $q.Enqueue($child)
                }

                if (($foldersVisited % 100) -eq 0) {
                    Write-ProgressText "Каталог: проверени $foldersVisited папки, намерени $productsFound изделия... (работи във фон)"
                }
            }
            catch {
                $errors++
                # Safety fallback to the original V5.1.34 marker test if the
                # optimized one-pass child enumeration itself fails. This keeps
                # discovery correctness ahead of speed.
                try {
                    $machineDir = Join-Path $dir '03_Mashini'
                    if ([System.IO.Directory]::Exists($machineDir)) {
                        $name = [System.IO.Path]::GetFileName($dir.TrimEnd('\'))
                        if ([string]::IsNullOrWhiteSpace($name)) { $name = $dir }
                        $depth = Get-Depth -Root $root -Path $dir
                        $lines.Add(((CsvField $root),(CsvField $name),(CsvField $dir),(CsvField ([string]$depth))) -join ';')
                        $productsFound++
                    }
                } catch {}
                continue
            }
        }
    }

    Write-ProgressText "Каталог: запис на $productsFound папки на изделия..."
    [System.IO.File]::WriteAllLines($IndexOut, $lines.ToArray(), $utf8Bom)
    Write-ProgressText "Каталогът е готов: $productsFound папки на изделия. Проверени папки: $foldersVisited. Пропуснати/недостъпни: $errors."
    exit 0
}
catch {
    Write-ProgressText ("Каталог ERROR: " + $_.Exception.Message)
    exit 1
}
