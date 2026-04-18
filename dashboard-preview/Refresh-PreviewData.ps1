# Regenerates embedded-data.json + embedded-data.js from PowerBI_Assets CSVs.
# Uses hashtable grouping for O(n) aggregation (scales to large fact tables).
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$assets = Join-Path $root 'PowerBI_Assets'
$dest = $PSScriptRoot

$fact = Import-Csv (Join-Path $assets 'Fact_Safety_KPI_Monthly.csv')
$kpi = Import-Csv (Join-Path $assets 'Dim_KPI.csv')
$cat = Import-Csv (Join-Path $assets 'Dim_Category.csv')
$biz = Import-Csv (Join-Path $assets 'Dim_Business.csv')
$dates = Import-Csv (Join-Path $assets 'Dim_Date.csv') | Sort-Object { [int]$_.DateKey }

$dateMap = @{}
foreach ($d in $dates) { $dateMap[[int]$d.DateKey] = $d.YearMonth }

$bizMap = @{}
foreach ($b in $biz) {
  $st = if ($b.State) { $b.State } else { $b.Region }
  $bizMap[[int]$b.BusinessKey] = @{ Name = $b.BusinessName; State = $st }
}

$kpiMap = @{}
foreach ($k in $kpi) {
  $kpiMap[[int]$k.KPIKey] = @{
    Name = $k.KPI_Name
    Cat = [int]$k.CategoryKey
    Unit = $k.Unit_Type
  }
}

$joined = foreach ($f in $fact) {
  $k = $kpiMap[[int]$f.KPIKey]
  [PSCustomObject]@{
    DateKey = [int]$f.DateKey
    BusinessKey = [int]$f.BusinessKey
    KPIKey = [int]$f.KPIKey
    CategoryKey = $k.Cat
    Value = [double]$f.Value
    KPI_Name = $k.Name
    Unit_Type = $k.Unit
  }
}

$lastDate = [int]($dates | Select-Object -Last 1).DateKey

# Group values by category + date for fast avg
$valsByCatDate = @{}
foreach ($row in $joined) {
  $key = "$($row.CategoryKey)|$($row.DateKey)"
  if (-not $valsByCatDate.ContainsKey($key)) {
    $valsByCatDate[$key] = [System.Collections.Generic.List[double]]::new()
  }
  $valsByCatDate[$key].Add($row.Value)
}

# Group by category + date + business for biz breakdown
$valsByCatDateBiz = @{}
foreach ($row in $joined) {
  $key = "$($row.CategoryKey)|$($row.DateKey)|$($row.BusinessKey)"
  if (-not $valsByCatDateBiz.ContainsKey($key)) {
    $valsByCatDateBiz[$key] = [System.Collections.Generic.List[double]]::new()
  }
  $valsByCatDateBiz[$key].Add($row.Value)
}

function Avg-List($list) {
  if (-not $list -or $list.Count -eq 0) { return 0 }
  ($list | Measure-Object -Average).Average
}

$catSummary = foreach ($c in ($cat | Sort-Object { [int]$_.SortOrder })) {
  $ck = [int]$c.CategoryKey
  $kpis = @($kpi | Where-Object { [int]$_.CategoryKey -eq $ck })
  $kKey = "$ck|$lastDate"
  $subset = if ($valsByCatDate.ContainsKey($kKey)) { $valsByCatDate[$kKey] } else { $null }
  $avgVal = if ($subset) { Avg-List $subset } else { 0 }
  [PSCustomObject]@{
    categoryKey = $ck
    categoryName = $c.CategoryName
    sortOrder = [int]$c.SortOrder
    uxNote = $c.UX_Note
    kpiCount = $kpis.Count
    latestMonthIndex = [math]::Round($avgVal, 2)
  }
}

$monthlyArr = @()
foreach ($c in $cat) {
  $ck = [int]$c.CategoryKey
  $points = @()
  foreach ($d in $dates) {
    $dk = [int]$d.DateKey
    $kKey = "$ck|$dk"
    $sub = if ($valsByCatDate.ContainsKey($kKey)) { $valsByCatDate[$kKey] } else { $null }
    $avg = if ($sub) { Avg-List $sub } else { 0 }
    $points += [PSCustomObject]@{ yearMonth = $d.YearMonth; value = [math]::Round($avg, 2) }
  }
  $monthlyArr += [PSCustomObject]@{ categoryKey = $ck; series = $points }
}

$bizArr = @()
foreach ($c in $cat) {
  $ck = [int]$c.CategoryKey
  $bars = @()
  foreach ($b in $biz) {
    $bk = [int]$b.BusinessKey
    $kKey = "$ck|$lastDate|$bk"
    $sub = if ($valsByCatDateBiz.ContainsKey($kKey)) { $valsByCatDateBiz[$kKey] } else { $null }
    $sum = if ($sub) { ($sub | Measure-Object -Sum).Sum } else { 0 }
    $bars += [PSCustomObject]@{ business = $b.BusinessName; value = [math]::Round($sum, 2) }
  }
  $bizArr += [PSCustomObject]@{ categoryKey = $ck; bars = $bars }
}

# KPI detail: group by KPI + last date
$valsByKpiDate = @{}
foreach ($row in $joined) {
  $key = "$($row.KPIKey)|$($row.DateKey)"
  if (-not $valsByKpiDate.ContainsKey($key)) {
    $valsByKpiDate[$key] = [System.Collections.Generic.List[double]]::new()
  }
  $valsByKpiDate[$key].Add($row.Value)
}

$detailArr = @()
foreach ($c in $cat) {
  $ck = [int]$c.CategoryKey
  $rows = @()
  foreach ($k in ($kpi | Where-Object { [int]$_.CategoryKey -eq $ck })) {
    $kk = [int]$k.KPIKey
    $kKey = "$kk|$lastDate"
    $vals = if ($valsByKpiDate.ContainsKey($kKey)) { $valsByKpiDate[$kKey] } else { $null }
    $avg = if ($vals) { Avg-List $vals } else { 0 }
    $rows += [PSCustomObject]@{
      kpiKey = $kk
      kpiName = $k.KPI_Name
      unitType = $k.Unit_Type
      latestValue = [math]::Round($avg, 4)
    }
  }
  $detailArr += [PSCustomObject]@{ categoryKey = $ck; kpis = $rows }
}

# Preview UX: TRIR (KPI 21) lives under category 3 in Dim_KPI — fan out to categories 1 & 2 so
# default "Total Recordable Incident Rate (TRI)" filters have fact rows and charts render.
$triMeta = $null
$d3 = $detailArr | Where-Object { $_.categoryKey -eq 3 }
if ($d3) {
  $triMeta = @($d3.kpis) | Where-Object { $_.kpiKey -eq 21 } | Select-Object -First 1
}
if ($triMeta) {
  foreach ($ck in @(1, 2)) {
    $block = $detailArr | Where-Object { $_.categoryKey -eq $ck } | Select-Object -First 1
    if (-not $block) { continue }
    $hasTri = @($block.kpis) | Where-Object { $_.kpiKey -eq 21 }
    if (-not $hasTri) {
      $block.kpis = @($block.kpis) + [PSCustomObject]@{
        kpiKey      = $triMeta.kpiKey
        kpiName     = $triMeta.kpiName
        unitType    = $triMeta.unitType
        latestValue = $triMeta.latestValue
      }
    }
  }
}

$factRows = [System.Collections.Generic.List[object]]::new()
foreach ($f in $fact) {
  $bk = [int]$f.BusinessKey
  $kk = [int]$f.KPIKey
  $dk = [int]$f.DateKey
  $b = $bizMap[$bk]
  $k = $kpiMap[$kk]
  $tgt = $null
  if ($f.Target -ne '' -and $f.Target -ne $null) { try { $tgt = [double]$f.Target } catch {} }
  $factRows.Add([PSCustomObject]@{
      yearMonth = $dateMap[$dk]
      dateKey = $dk
      businessKey = $bk
      businessName = $b.Name
      state = $b.State
      kpiKey = $kk
      kpiName = $k.Name
      categoryKey = $k.Cat
      unitType = $k.Unit
      value = [double]$f.Value
      target = $tgt
    })
}

foreach ($row in @($factRows.ToArray())) {
  if ($row.kpiKey -ne 21 -or $row.categoryKey -ne 3) { continue }
  foreach ($altCat in @(1, 2)) {
    $factRows.Add([PSCustomObject]@{
        yearMonth     = $row.yearMonth
        dateKey       = $row.dateKey
        businessKey   = $row.businessKey
        businessName  = $row.businessName
        state         = $row.state
        kpiKey        = $row.kpiKey
        kpiName       = $row.kpiName
        categoryKey   = $altCat
        unitType      = $row.unitType
        value         = $row.value
        target        = $row.target
      })
  }
}

$monthsList = foreach ($d in $dates) {
  [PSCustomObject]@{ yearMonth = $d.YearMonth; dateKey = [int]$d.DateKey }
}
$statesList = @($biz | ForEach-Object { if ($_.State) { $_.State } else { $_.Region } } | Select-Object -Unique | Sort-Object)

$out = [PSCustomObject]@{
  meta = [PSCustomObject]@{
    dashboardTitle = 'Adani Safety Performance Dashboard'
    subtitle = 'Safety performance indicators — Interactive Preview'
    dataNote = 'Demo data from PowerBI_Assets CSVs; replace with production feeds.'
    lastDataMonth = ($dates | Select-Object -Last 1).YearMonth
    lastUpdateISO = (Get-Date).ToUniversalTime().ToString('o')
  }
  categories = @($catSummary)
  monthlyByCategory = $monthlyArr
  businessBreakdown = $bizArr
  kpiDetailByCategory = $detailArr
  factRows = $factRows.ToArray()
  months = @($monthsList)
  states = @($statesList)
}

$json = $out | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText((Join-Path $dest 'embedded-data.json'), $json)
$jsBody = "window.__DASHBOARD_DATA__ = $json;"
[System.IO.File]::WriteAllText((Join-Path $dest 'embedded-data.js'), $jsBody, [System.Text.UTF8Encoding]::new($false))
Write-Host "Updated embedded-data.json and embedded-data.js (factRows: $($factRows.Count))"
