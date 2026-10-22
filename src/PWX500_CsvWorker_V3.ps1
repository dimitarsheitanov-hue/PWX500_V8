param(
    [Parameter(Mandatory=$true)][string]$InputFolder,
    [Parameter(Mandatory=$true)][string]$RowsOut,
    [Parameter(Mandatory=$true)][string]$SummaryOut,
    [Parameter(Mandatory=$true)][string]$ProgressFile,
    [string]$SkipListCsv=''
)

# SAFETY: Reads only *.csv in InputFolder TOP LEVEL. No subfolders and no other file types.
$ErrorActionPreference='Stop'
$utf8Bom=New-Object System.Text.UTF8Encoding($true)
$utf8NoBom=New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName Microsoft.VisualBasic

function Progress([string]$t){ try{[System.IO.File]::WriteAllText($ProgressFile,$t,$utf8NoBom)}catch{} }
function F([string]$s){ if($null -eq $s){$s=''}; if($s.Contains('"')){$s=$s.Replace('"','""')}; if($s.Contains(';') -or $s.Contains('"') -or $s.Contains("`r") -or $s.Contains("`n")){return '"'+$s+'"'}; return $s }
function S($v){ if($null -eq $v){return ''}; return ([string]$v).Trim() }
function Enc([string]$Path){
  $bytes=[System.IO.File]::ReadAllBytes($Path)
  if($bytes.Length -ge 3 -and $bytes[0]-eq 0xEF -and $bytes[1]-eq 0xBB -and $bytes[2]-eq 0xBF){return (New-Object System.Text.UTF8Encoding($true))}
  try{$u=New-Object System.Text.UTF8Encoding($false,$true);[void]$u.GetString($bytes);return $u}catch{return [System.Text.Encoding]::GetEncoding(1251)}
}
function Delim([string]$Path,[System.Text.Encoding]$Encoding){
  $r=New-Object System.IO.StreamReader -ArgumentList @($Path,$Encoding,$true)
  try{$line=$r.ReadLine()}finally{$r.Dispose()}
  if($null -eq $line){return ';'}
  $best=';';$bestCount=-1
  foreach($d in @(';',',',"`t")){$c=0;foreach($ch in $line.ToCharArray()){if($ch -eq $d){$c++}};if($c -gt $bestCount){$best=$d;$bestCount=$c}}
  return $best
}
function ReadCsv([string]$Path){
  if([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.csv'){throw 'Blocked: only .csv is allowed.'}
  $e=Enc $Path;$d=Delim $Path $e;$p=$null
  try{
    $p=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path,$e,$true)
    $p.TextFieldType=[Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $p.SetDelimiters([string[]]@([string]$d));$p.HasFieldsEnclosedInQuotes=$true;$p.TrimWhiteSpace=$false
    if($p.EndOfData){return @()}
    $h=@($p.ReadFields());for($i=0;$i -lt $h.Count;$i++){$h[$i]=(S $h[$i]).Trim([char]0xFEFF)}
    $out=New-Object System.Collections.Generic.List[object]
    while(-not $p.EndOfData){$fs=@($p.ReadFields());$o=[ordered]@{};for($i=0;$i -lt $h.Count;$i++){$k=S $h[$i];if([string]::IsNullOrWhiteSpace($k)){$k="Column_$i"};$v='';if($i -lt $fs.Count -and $null -ne $fs[$i]){$v=[string]$fs[$i]};$o[$k]=$v};$out.Add([pscustomobject]$o)}
    return @($out.ToArray())
  }finally{if($null -ne $p){try{$p.Close()}catch{};try{$p.Dispose()}catch{}}}
}
function V($o,[string]$n){$p=$o.PSObject.Properties[$n];if($null -eq $p){return ''};return [string]$p.Value}
function ParseDesc([string]$desc){
  $r=[ordered]@{Product='';DetailName='';DetailNo='';Parsed=$false};$desc=S $desc;if([string]::IsNullOrWhiteSpace($desc)){return [pscustomobject]$r}
  $parts=@([regex]::Split($desc,'\s+-\s+'))
  if($parts.Count -ge 3){$last=S $parts[$parts.Count-1];$n=0;if([int]::TryParse($last,[ref]$n)){$r.DetailNo=[string]$n;$r.DetailName=S $parts[$parts.Count-2];$pp=New-Object System.Collections.Generic.List[string];for($i=0;$i -le $parts.Count-3;$i++){$x=S $parts[$i];if(-not [string]::IsNullOrWhiteSpace($x)){$pp.Add($x)}};$r.Product=(($pp.ToArray()) -join ' - ').Trim();$r.Parsed=-not [string]::IsNullOrWhiteSpace($r.Product)}}
  return [pscustomobject]$r
}

try{
  if(-not [System.IO.Directory]::Exists($InputFolder)){throw "CSV folder unavailable: $InputFolder"}
  $filesAll=@([System.IO.Directory]::EnumerateFiles($InputFolder,'*.csv',[System.IO.SearchOption]::TopDirectoryOnly)|Sort-Object)
  $skip=@{}
  if(-not [string]::IsNullOrWhiteSpace($SkipListCsv) -and [System.IO.File]::Exists($SkipListCsv)){
    try{foreach($x in @(Import-Csv -LiteralPath $SkipListCsv -Encoding UTF8)){ $p=S $x.FullPath;if(-not [string]::IsNullOrWhiteSpace($p)){$skip[$p.ToLowerInvariant()]=$true} }}catch{}
  }
  $files=@($filesAll|Where-Object{-not $skip.ContainsKey(([string]$_).ToLowerInvariant())})
  Progress "CSV план: $($files.Count) за обработка; $($filesAll.Count-$files.Count) завършени се пропускат."
  $rowLines=New-Object System.Collections.Generic.List[string]
  $rowLines.Add('CsvFile;CsvPath;Product;DetailNo;DetailName;Length;Width;Thickness;Quantity;ProcessingNo;ProgramCodeF;Status')
  $sumLines=New-Object System.Collections.Generic.List[string]
  $sumLines.Add('CSV;RawRows;ParsedRows;Products;PhysicalDetails;ProgramCodeF;NoProgram;DoubleProcessing;DuplicateProgram;UnparsedRows;Status;ErrorText')
  $done=0
  foreach($file in $files){
    $name=[System.IO.Path]::GetFileName($file);$raw=0;$parsed=0;$unparsed=0;$status='OK';$err='';$items=New-Object System.Collections.Generic.List[object]
    try{
      $rows=@(ReadCsv $file);$raw=$rows.Count
      if($rows.Count -gt 0){$names=@($rows[0].PSObject.Properties|ForEach-Object{$_.Name});$req=@('ItemDescription','Length','Width','Thickness','Quantity','ProgramCodeF');$miss=@($req|Where-Object{$names -notcontains $_});if($miss.Count -gt 0){throw ('Missing columns: '+($miss -join ', '))}}
      $occ=@{}
      foreach($row in $rows){$pd=ParseDesc (V $row 'ItemDescription');if(-not $pd.Parsed){$unparsed++;continue};$key="$($pd.Product)|$($pd.DetailNo)|$($pd.DetailName)|$(V $row 'Length')|$(V $row 'Width')|$(V $row 'Thickness')";if(-not $occ.ContainsKey($key)){$occ[$key]=0};$occ[$key]=[int]$occ[$key]+1;$pn=[int]$occ[$key];$prog=S (V $row 'ProgramCodeF');$st=if([string]::IsNullOrWhiteSpace($prog)){'БЕЗ ProgramCodeF'}else{'OK'};$it=[pscustomobject]@{CsvFile=$name;CsvPath=$file;Product=$pd.Product;DetailNo=$pd.DetailNo;DetailName=$pd.DetailName;Length=V $row 'Length';Width=V $row 'Width';Thickness=V $row 'Thickness';Quantity=V $row 'Quantity';ProcessingNo=$pn;ProgramCodeF=$prog;Status=$st};$items.Add($it);$parsed++;$rowLines.Add(((F $it.CsvFile),(F $it.CsvPath),(F $it.Product),(F $it.DetailNo),(F $it.DetailName),(F $it.Length),(F $it.Width),(F $it.Thickness),(F $it.Quantity),(F ([string]$it.ProcessingNo)),(F $it.ProgramCodeF),(F $it.Status)) -join ';')}
      if($rows.Count -eq 0 -or $unparsed -gt 0){$status='ПРОВЕРИ'}
    }catch{$status='ГРЕШКА';$err=$_.Exception.Message}
    $a=@($items.ToArray());$products=@($a|ForEach-Object{$_.Product}|Sort-Object -Unique).Count;$physical=@($a|ForEach-Object{"$($_.Product)|$($_.DetailNo)|$($_.DetailName)|$($_.Length)|$($_.Width)|$($_.Thickness)"}|Sort-Object -Unique).Count;$programs=@($a|Where-Object{-not [string]::IsNullOrWhiteSpace($_.ProgramCodeF)}).Count;$no=@($a|Where-Object{[string]::IsNullOrWhiteSpace($_.ProgramCodeF)}).Count;$double=@($a|Group-Object{"$($_.Product)|$($_.DetailNo)|$($_.DetailName)|$($_.Length)|$($_.Width)|$($_.Thickness)"}|Where-Object{$_.Count -gt 1}).Count;$dup=@($a|Where-Object{-not [string]::IsNullOrWhiteSpace($_.ProgramCodeF)}|Group-Object ProgramCodeF|Where-Object{$_.Count -gt 1}).Count;if($dup -gt 0 -and $status -eq 'OK'){$status='ПРОВЕРИ'}
    $sumLines.Add(((F $name),(F ([string]$raw)),(F ([string]$parsed)),(F ([string]$products)),(F ([string]$physical)),(F ([string]$programs)),(F ([string]$no)),(F ([string]$double)),(F ([string]$dup)),(F ([string]$unparsed)),(F $status),(F $err)) -join ';')
    $done++;Progress "CSV: $done / $($files.Count) файла... (работи във фон)"
  }
  [System.IO.File]::WriteAllLines($RowsOut,$rowLines.ToArray(),$utf8Bom)
  [System.IO.File]::WriteAllLines($SummaryOut,$sumLines.ToArray(),$utf8Bom)
  Progress "CSV готови: $($files.Count) обработени; $($filesAll.Count-$files.Count) завършени пропуснати."
  exit 0
}catch{Progress ('CSV ERROR: '+$_.Exception.Message);exit 1}
