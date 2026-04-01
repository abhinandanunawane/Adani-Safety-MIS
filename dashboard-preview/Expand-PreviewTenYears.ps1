# Extends Dim_Date to ~10 years (2016-01 .. 2026-02) and Fact_Safety_KPI_Monthly.csv
# by replicating the 2025-02-01 snapshot per month with mild drift (for LY / trend demos).
$ErrorActionPreference = 'Stop'
$assets = Join-Path (Split-Path $PSScriptRoot -Parent) 'PowerBI_Assets'
$factPath = Join-Path $assets 'Fact_Safety_KPI_Monthly.csv'
$datePath = Join-Path $assets 'Dim_Date.csv'

$fact = Import-Csv $factPath
$templateDateKey = 20250201
$template = @($fact | Where-Object { [int]$_.DateKey -eq $templateDateKey })
if ($template.Count -eq 0) { throw "No fact rows for template DateKey $templateDateKey" }

$dimDateOut = [System.Collections.Generic.List[object]]::new()
$dk = Get-Date -Year 2016 -Month 1 -Day 1
$end = Get-Date -Year 2026 -Month 2 -Day 1
while ($dk -le $end) {
    $key = [int]$dk.ToString('yyyyMMdd')
    $ym = $dk.ToString('yyyy-MM')
    $y = $dk.Year
    $m = $dk.Month
    $dimDateOut.Add([PSCustomObject]@{
            DateKey      = $key
            Date         = $dk.ToString('yyyy-MM-dd')
            Year         = $y
            Month        = $m
            MonthName    = $dk.ToString('MMMM')
            YearMonth    = $ym
            MonthSort    = [int]$dk.ToString('yyyyMM')
            FiscalPeriod = ('FY{0}' -f $y)
        })
    $dk = $dk.AddMonths(1)
}
$dimDateOut | Export-Csv $datePath -NoTypeInformation

$fid = 1
$factOut = [System.Collections.Generic.List[object]]::new()
foreach ($row in $dimDateOut) {
    $dk = [datetime]::ParseExact([string]$row.Date, 'yyyy-MM-dd', $null)
    $y = $dk.Year
    $m = $dk.Month
    foreach ($t in $template) {
        $bk = [int]$t.BusinessKey
        $kk = [int]$t.KPIKey
        $base = [double]$t.Value
        $targ = $null
        if ($t.Target -ne '' -and $null -ne $t.Target) { try { $targ = [double]$t.Target } catch {} }
        $wave = ([math]::Sin([double]($y * 12 + $m) * 0.08) + 1.0) * 0.015
        $yr = ($y - 2025) * 0.004
        $v = [math]::Round($base * (1.0 + $wave + $yr), 4)
        $tv = $null
        if ($null -ne $targ) { $tv = [math]::Round($targ * (1.0 + $wave * 0.45 + $yr * 0.5), 4) }
        $factOut.Add([PSCustomObject]@{
                FactKey           = $fid++
                BusinessKey       = $bk
                DateKey           = [int]$row.DateKey
                KPIKey            = $kk
                Value             = $v
                Target            = if ($null -ne $tv) { $tv } else { '' }
                DataQualityFlag   = 'OK'
                Source_Note       = 'Expanded ten-year preview span'
            })
    }
}
$factOut | Export-Csv $factPath -NoTypeInformation
Write-Host ("Dim_Date: {0} rows; Fact: {1} rows" -f $dimDateOut.Count, $factOut.Count)
