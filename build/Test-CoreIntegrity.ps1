# V5.1.38: assure no tested converter/worker/time pipeline source was altered.
$ErrorActionPreference='Stop'
$src=Join-Path (Split-Path -Parent $PSScriptRoot) 'src'
$expected=@{
    'CAM_Agent_TimesImportWorker_V5_1_31.ps1'='c4c78c992395cea7cb752c97cc60a627c615d35234d396c5ea200b5a10df4003'
    'CAM_Agent_TimesReport_Template.xlsx'='9f4bd516491d360905cee2c3710e30e30208d1c4dcfb59fba835ab39cffdefa1'
    'CAM_Agent_Times_Active.csv'='a0f110226d47969db04d2de1cba46399f4014eb270f50457545dfbcd4c0a119d'
    'CAM_Agent_Times_Base.csv'='cc6381f632627c34577aff09671a5cdb926b8ac20e86dcfd6485826d9b8f24f7'
    'CAM_Agent_Times_DB.csv'='714185884f20ca3c6f006413332c666a78063b980c4144ef0eb701adc9127754'
    'HOP_to_PGMX_2_9.exe'='8848f7b82fecc6a6c903029d00d5896d790da9da11bdc16c1236e2e4b3896643'
    'PWX500_CsvWorker_V3.ps1'='f38d4c4ad280fca3feea768fa12ee431c2469d7ce7497573f5bb5e80bcdea243'
    'PWX500_HistoryWorker_V3.ps1'='738523fa2e2df5d7a0daaf01cb327dee524a46fa0f1bc83fcf36bb3263c13020'
    'PWX500_HopWorker_V3.ps1'='834cd62b655d450e3883927a5551fd7628b85a6e4642a24ebbfdcec8484ff205'
    'PWX500_IndexWorker_V3.ps1'='6bf22c241f104fbf4bdf6fd22cf67347f44480c55a4e27f1cee7200eda72b0e9'
    'PWX500_PgmxWorker_V3.ps1'='6f6a39a1d52c82bcffb8af5b2aab8df7b6e02d03879a71e723533e0d290bbbec'
    'PWX500_ProductCategories.csv'='337c5ad66f40dd53c09e0917cbb9f682eee0ed8468af1fd6d1b7c7bb31928708'
    'PWX500_SearchWorker_V3.ps1'='656cb1e1c7eb5a157c2ffd8df102916a3d233a6740af9378c42d58c70de67d32'
}

# Windows Git checkouts can change LF to CRLF (or vice versa) without changing
# a single command. The canonical hash list is pinned to the v5.1.37 source
# after CRLF->LF normalization. Other byte changes are still rejected.
$expectedCanonical=@{
    'CAM_Agent_TimesImportWorker_V5_1_31.ps1'='2303f4b18375021c9fd71da70a8da95680a41c5c2fc9c6e864e39b4a2e3951e7'
    'CAM_Agent_Times_Active.csv'='d4b37787a7afdcbc80916b04793817b3e65fe6dbe794ade9ba2421bbe85c453f'
    'CAM_Agent_Times_Base.csv'='0665e2bf1503b50b2a28c5e7b138388c32ae24b91d02f32f30a6352fd7baf57d'
    'CAM_Agent_Times_DB.csv'='51e1c3c01931620d612ac2187b0b31bad63602104a37d39fc11bddbe927a6f7e'
    'PWX500_CsvWorker_V3.ps1'='f38d4c4ad280fca3feea768fa12ee431c2469d7ce7497573f5bb5e80bcdea243'
    'PWX500_HistoryWorker_V3.ps1'='738523fa2e2df5d7a0daaf01cb327dee524a46fa0f1bc83fcf36bb3263c13020'
    'PWX500_HopWorker_V3.ps1'='834cd62b655d450e3883927a5551fd7628b85a6e4642a24ebbfdcec8484ff205'
    'PWX500_IndexWorker_V3.ps1'='6bf22c241f104fbf4bdf6fd22cf67347f44480c55a4e27f1cee7200eda72b0e9'
    'PWX500_PgmxWorker_V3.ps1'='6f6a39a1d52c82bcffb8af5b2aab8df7b6e02d03879a71e723533e0d290bbbec'
    'PWX500_ProductCategories.csv'='337c5ad66f40dd53c09e0917cbb9f682eee0ed8468af1fd6d1b7c7bb31928708'
    'PWX500_SearchWorker_V3.ps1'='656cb1e1c7eb5a157c2ffd8df102916a3d233a6740af9378c42d58c70de67d32'
}


function Get-CanonicalTextSHA256 {
    param([Parameter(Mandatory=$true)][string]$Path)
    # ISO-8859-1 maps all bytes 0..255 one-to-one, so non-UTF8 source and BOMs
    # remain byte-identical; only CRLF pairs become LF.
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $text = $latin1.GetString($bytes)
    $normalBytes = $latin1.GetBytes($text.Replace("`r`n", "`n"))
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha.ComputeHash($normalBytes))).Replace('-', '').ToLowerInvariant()
    } finally { $sha.Dispose() }
}

$eolOnlyDifferences=0
foreach($n in ($expected.Keys | Sort-Object)){
    $path=Join-Path $src $n
    if(-not [System.IO.File]::Exists($path)){ throw ('Missing stable core dependency: '+$n) }
    $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($hash -eq $expected[$n]) { continue }  # Fast exact-byte path.
    if ($expectedCanonical.ContainsKey($n)) {
        $canonicalHash = Get-CanonicalTextSHA256 -Path $path
        if ($canonicalHash -eq $expectedCanonical[$n]) {
            $eolOnlyDifferences++
            Write-Host ('INFO: LF/CRLF checkout difference only: '+$n)
            continue
        }
    }
    throw ('Changed stable core dependency beyond line endings: '+$n)
}
Write-Host ('PASS V5.1.38 stable core verification: {0} files; EOL-only differences: {1}.' -f $expected.Count, $eolOnlyDifferences) -ForegroundColor Green
