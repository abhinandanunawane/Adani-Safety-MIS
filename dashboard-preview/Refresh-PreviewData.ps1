# Regenerates embedded-data.json + embedded-data.js from PowerBI_Assets CSVs.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$assets = Join-Path $root 'PowerBI_Assets'
$dest = $PSScriptRoot

$fact = Import-Csv (Join-Path $assets 'Fact_Safety_KPI_Monthly.csv')
$kpi = Import-Csv (Join-Path $assets 'Dim_KPI.csv')
$cat = Import-Csv (Join-Path $assets 'Dim_Category.csv')
$biz = Import-Csv (Join-Path $assets 'Dim_Business.csv')
$dates = Import-Csv (Join-Path $assets 'Dim_Date.csv') | Sort-Object DateKey

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

$catSummary = foreach ($c in ($cat | Sort-Object {[int]$_.SortOrder})) {
  $ck = [int]$c.CategoryKey
  $subset = $joined | Where-Object { $_.CategoryKey -eq $ck -and $_.DateKey -eq $lastDate }
  $kpis = @($kpi | Where-Object { [int]$_.CategoryKey -eq $ck })
  $avgVal = if ($subset.Count) { ($subset | Measure-Object -Property Value -Average).Average } else { 0 }
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
    $sub = $joined | Where-Object { $_.CategoryKey -eq $ck -and $_.DateKey -eq $dk }
    $avg = if ($sub.Count) { ($sub | Measure-Object -Property Value -Average).Average } else { 0 }
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
    $sub = $joined | Where-Object { $_.CategoryKey -eq $ck -and $_.DateKey -eq $lastDate -and $_.BusinessKey -eq $bk }
    $sum = if ($sub.Count) { ($sub | Measure-Object -Property Value -Sum).Sum } else { 0 }
    $bars += [PSCustomObject]@{ business = $b.BusinessName; value = [math]::Round($sum, 2) }
  }
  $bizArr += [PSCustomObject]@{ categoryKey = $ck; bars = $bars }
}

$detailArr = @()
foreach ($c in $cat) {
  $ck = [int]$c.CategoryKey
  $rows = @()
  foreach ($k in ($kpi | Where-Object { [int]$_.CategoryKey -eq $ck })) {
    $kk = [int]$k.KPIKey
    $vals = $joined | Where-Object { $_.KPIKey -eq $kk -and $_.DateKey -eq $lastDate }
    $avg = if ($vals.Count) { ($vals | Measure-Object -Property Value -Average).Average } else { 0 }
    $rows += [PSCustomObject]@{
      kpiKey = $kk
      kpiName = $k.KPI_Name
      unitType = $k.Unit_Type
      latestValue = [math]::Round($avg, 4)
    }
  }
  $detailArr += [PSCustomObject]@{ categoryKey = $ck; kpis = $rows }
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

$monthsList = foreach ($d in $dates) {
  [PSCustomObject]@{ yearMonth = $d.YearMonth; dateKey = [int]$d.DateKey }
}
$statesList = @($biz | ForEach-Object { if ($_.State) { $_.State } else { $_.Region } } | Select-Object -Unique | Sort-Object)

$out = [PSCustomObject]@{
  meta = [PSCustomObject]@{
    dashboardTitle = 'Adani Safety MIS'
    subtitle = 'Group Safety KPI Dashboard — Interactive Preview'
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
