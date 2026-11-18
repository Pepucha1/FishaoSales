<#
  Fishao - denni sber cen ryb z https://www.fishao.net/wiki
  - stahne aktualni cenu kazde ryby a ulozi ji do ceny.csv (jeden radek = ryba + den)
  - vygeneruje prehled.html: dnesni cena vs. prumer, min, max
  Pouziti:
    powershell -ExecutionPolicy Bypass -File fishao-ceny.ps1            # stahnout + report
    powershell -ExecutionPolicy Bypass -File fishao-ceny.ps1 -JenReport # jen prepocitat report
#>
param([switch]$JenReport)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Dir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$CsvPath  = Join-Path $Dir 'ceny.csv'
$HtmlPath = Join-Path $Dir 'prehled.html'
$LogPath  = Join-Path $Dir 'log.txt'
$Base   = 'https://www.fishao.net'
$Utf8   = New-Object System.Text.UTF8Encoding($false)

function Log($msg) {
    $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Write-Host $line
    [IO.File]::AppendAllText($LogPath, $line + "`r`n", $Utf8)
}

function Get-Text($url) {
    $wc = New-Object Net.WebClient
    $wc.Encoding = $Utf8
    $wc.Headers['User-Agent'] = 'Mozilla/5.0 (fishao-ceny osobni statistika)'
    try { return $wc.DownloadString($url) } finally { $wc.Dispose() }
}

# ---------- 1) Katalog ryb (id, jmeno, hvezdy) z JS balicku wiki ----------
function Get-Catalog {
    $html = Get-Text "$Base/wiki/fish"
    $chunks = [regex]::Matches($html, '/_app/immutable/[^"]+\.js') | ForEach-Object { $_.Value } | Sort-Object -Unique
    foreach ($c in $chunks) {
        $js = Get-Text ($Base + $c)
        $m = [regex]::Matches($js, '\{id:(\d+),name:`([^`]*)`,rating:(\d)[^{}]*?locs:\[([\d,]*)\]')
        if ($m.Count -gt 100) {
            return $m | ForEach-Object {
                [pscustomobject]@{ id = [int]$_.Groups[1].Value; name = $_.Groups[2].Value.Trim(); stars = [int]$_.Groups[3].Value; locs = $_.Groups[4].Value -replace ',', ';' }
            }
        }
    }
    throw 'Katalog ryb nenalezen - web se asi zmenil.'
}

# ---------- 2) Cena jedne ryby z SvelteKit __data.json ----------
function Get-Price($id) {
    $j = Get-Text "$Base/wiki/fish/$id/__data.json" | ConvertFrom-Json
    $node = $j.nodes | Where-Object { $_ -and $_.data -and $_.data[0].PSObject.Properties.Name -contains 'market' } | Select-Object -First 1
    if (-not $node) { return $null }
    $d = $node.data
    $mi = $d[0].market
    if ($null -eq $mi -or $mi -lt 0) { return $null }
    $mk = $d[$mi]
    if ($null -eq $mk -or $null -eq $mk.price) { return $null }
    $price = $d[$mk.price]
    $chg = if ($null -ne $mk.change) { $d[$mk.change] } else { $null }
    if ($null -eq $price) { return $null }
    return @{ price = [double]$price; change = $chg }
}

# ryby.csv: id ryby -> id oblasti oddelene strednikem (obnovuje se pri kazdem sberu)
function Save-Areas($cat) {
    $lines = @('id,locs') + @($cat | ForEach-Object { '{0},{1}' -f $_.id, $_.locs })
    [IO.File]::WriteAllLines((Join-Path $Dir 'ryby.csv'), [string[]]$lines, $Utf8)
}

function Scrape {
    $today = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd')
    Log "Start sberu cen za $today"
    $cat = Get-Catalog
    Log "Katalog: $($cat.Count) ryb"
    Save-Areas $cat
    $rows = New-Object System.Collections.Generic.List[string]
    $ok = 0; $fail = 0
    foreach ($f in $cat) {
        $p = $null
        for ($try = 1; $try -le 3 -and -not $p; $try++) {
            try { $p = Get-Price $f.id } catch { Start-Sleep -Seconds (2 * $try) }
        }
        if ($p) {
            $name = $f.name -replace '"', '""'
            $rows.Add(('{0},{1},"{2}",{3},{4},{5}' -f $today, $f.id, $name, $f.stars, $p.price, $p.change))
            $ok++
        } else { $fail++ }
        Start-Sleep -Milliseconds 100   # setrnost k serveru
    }
    if ($ok -eq 0) { throw 'Nepodarilo se stahnout zadnou cenu.' }

    # prepis pripadnych dnesnich radku (kdyz se skript pusti 2x za den)
    $header = 'date,id,name,stars,price,change'
    $old = @()
    if (Test-Path $CsvPath) {
        $old = [IO.File]::ReadAllLines($CsvPath, $Utf8) | Select-Object -Skip 1 | Where-Object { $_ -and -not $_.StartsWith("$today,") }
    }
    [IO.File]::WriteAllLines($CsvPath, [string[]](@($header) + $old + $rows), $Utf8)
    Log "Ulozeno $ok cen, nezdarilo se $fail"
}

# ---------- 3) Report ----------
function Build-Report {
    if (-not (Test-Path $CsvPath)) { throw 'ceny.csv zatim neexistuje.' }
    $data = Import-Csv $CsvPath -Encoding UTF8
    $by = @{}
    foreach ($r in $data) {
        if (-not $by.ContainsKey($r.id)) { $by[$r.id] = New-Object System.Collections.Generic.List[object] }
        $by[$r.id].Add($r)
    }
    $allDates = @($data | Select-Object -ExpandProperty date | Sort-Object -Unique)
    $lastDate = $allDates[$allDates.Count - 1]
    $inv = [Globalization.CultureInfo]::InvariantCulture

    # ceske nazvy (nazvy-cz.csv: id,en,cz) - kdyz chybi, zustane anglicky
    $cz = @{}
    $czPath = Join-Path $Dir 'nazvy-cz.csv'
    if (Test-Path $czPath) {
        foreach ($n in (Import-Csv $czPath -Encoding UTF8)) { if ($n.cz) { $cz[$n.id] = $n.cz } }
    }

    # oblasti: oblasti.csv (id,en,cz) + ryby.csv (id ryby -> oblasti)
    $areaName = @{}
    $areaPath = Join-Path $Dir 'oblasti.csv'
    if (Test-Path $areaPath) {
        foreach ($a in (Import-Csv $areaPath -Encoding UTF8)) { $areaName[$a.id] = $(if ($a.cz) { $a.cz } else { $a.en }) }
    }
    $fishPath = Join-Path $Dir 'ryby.csv'
    if (-not (Test-Path $fishPath)) { Save-Areas (Get-Catalog) }
    $fishAreas = @{}
    foreach ($r in (Import-Csv $fishPath -Encoding UTF8)) {
        $fishAreas[$r.id] = @($r.locs -split ';' | Where-Object { $_ } | ForEach-Object { if ($areaName.ContainsKey($_)) { $areaName[$_] } else { "Oblast $_" } })
    }

    $out = foreach ($id in $by.Keys) {
        $h = $by[$id] | Sort-Object date
        $last = $h[$h.Count - 1]
        $prices = @($h | ForEach-Object { [double]::Parse($_.price, $inv) })
        $cur = $prices[$prices.Count - 1]
        $avg = ($prices | Measure-Object -Average).Average
        $p30 = @($prices | Select-Object -Last 30)
        $avg30 = ($p30 | Measure-Object -Average).Average
        $mn = ($prices | Measure-Object -Minimum).Minimum
        $mx = ($prices | Measure-Object -Maximum).Maximum
        $sd = 0
        if ($prices.Count -gt 1) { $sd = [math]::Sqrt((($prices | ForEach-Object { ($_ - $avg) * ($_ - $avg) } | Measure-Object -Sum).Sum) / ($prices.Count - 1)) }
        [ordered]@{
            id = [int]$id; name = $(if ($cz.ContainsKey($id)) { $cz[$id] } else { $last.name }); en = $last.name; stars = [int]$last.stars
            areas = $(if ($fishAreas.ContainsKey($id)) { ,$fishAreas[$id] } else { ,@() })
            date = $last.date; isCurrent = ($last.date -eq $lastDate)
            price = $cur; wikiChange = $last.change
            avg = [math]::Round($avg, 1); avg30 = [math]::Round($avg30, 1)
            min = $mn; max = $mx; days = $prices.Count
            diff = if ($avg) { [math]::Round(($cur - $avg) / $avg * 100, 1) } else { 0 }
            z = if ($sd -gt 0) { [math]::Round(($cur - $avg) / $sd, 2) } else { $null }
            hist = @($prices | Select-Object -Last 60)
        }
    }
    $json = ConvertTo-Json -InputObject @($out) -Depth 4 -Compress
    $tpl = [IO.File]::ReadAllText((Join-Path $Dir 'sablona.html'), $Utf8)
    $html = $tpl.Replace('/*DATA*/null', $json).Replace('{{LAST}}', $lastDate).Replace('{{DAYS}}', [string]@($allDates).Count).Replace('{{GEN}}', ([TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, 'Central Europe Standard Time').ToString('d. M. yyyy HH:mm')))
    [IO.File]::WriteAllText($HtmlPath, $html, $Utf8)
    Log "Report vygenerovan: $HtmlPath"
}

try {
    if (-not $JenReport) { Scrape }
    Build-Report
} catch {
    Log "CHYBA: $($_.Exception.Message)"
    exit 1
}
