# Generates KPI table rows for kpi-reference.html from Dim_KPI.csv + Dim_Category.csv
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web
$base = Split-Path -Parent $PSScriptRoot
$csvKpi = Join-Path $base 'PowerBI_Assets\Dim_KPI.csv'
$csvCat = Join-Path $base 'PowerBI_Assets\Dim_Category.csv'
$out = Join-Path $PSScriptRoot '_kpi-tables-fragment.html'

function HtmlEsc([string]$s) {
    if ($null -eq $s) { return '' }
    [System.Web.HttpUtility]::HtmlEncode($s)
}

function UnitLbl([string]$ut) {
    switch ($ut) {
        'Count' { 'Count' }
        'PercentOrRate' { '% / rate' }
        'Days' { 'Days' }
        'Hours' { 'Hours' }
        default { HtmlEsc $ut }
    }
}

function BizDef([string]$name, [string]$ut, [string]$note) {
    $n = if ([string]::IsNullOrWhiteSpace($name)) { 'This measure' } else { $name.Trim() }
    $grain = if ($note -and $note.Trim() -ne 'Wireframe KPI') { $note.Trim() } else { 'preview wireframe model' }
    switch ($ut) {
        'Count' { "Count of $n for the reporting period and active BU / geography / vertical filters ($grain)." }
        'PercentOrRate' { "Percentage or rate for $n for the reporting period and active filters ($grain)." }
        'Days' { "Total days for $n for the reporting period and active filters ($grain)." }
        'Hours' { "Total hours for $n for the reporting period and active filters ($grain)." }
        default { "Measure for $n ($grain)." }
    }
}

function Append-Thead6([System.Text.StringBuilder]$sb) {
    [void]$sb.AppendLine('            <tr>')
    [void]$sb.AppendLine('              <th scope="col" class="doc-table__th-num">#</th>')
    [void]$sb.AppendLine('              <th scope="col">Category</th>')
    [void]$sb.AppendLine('              <th scope="col">KPI</th>')
    [void]$sb.AppendLine('              <th scope="col">Unit</th>')
    [void]$sb.AppendLine('              <th scope="col">Business definition</th>')
    [void]$sb.AppendLine('              <th scope="col">Formula (preview model)</th>')
    [void]$sb.AppendLine('            </tr>')
}

$kpis = @(Import-Csv $csvKpi -Encoding utf8)
$cats = @(Import-Csv $csvCat -Encoding utf8 | Sort-Object { [int]$_.SortOrder })
$sb = [System.Text.StringBuilder]::new()
[int]$rowNum = 0

foreach ($c in $cats) {
    $ck = $c.CategoryKey.ToString().Trim()
    $catName = HtmlEsc $c.CategoryName
    $rows = @($kpis | Where-Object { $_.CategoryKey.ToString().Trim() -eq $ck } | Sort-Object { [int]$_.KPIKey })
    [void]$sb.AppendLine(('    <section class="doc-kpi-cat" id="cat-{0}">' -f $ck))
    [void]$sb.AppendLine(('      <h2>{0}</h2>' -f $catName))
    [void]$sb.AppendLine('      <div class="doc-table-wrap">')
    [void]$sb.AppendLine('        <table class="doc-table doc-table--wide">')
    [void]$sb.AppendLine('          <thead>')
    Append-Thead6 $sb
    [void]$sb.AppendLine('          </thead>')
    [void]$sb.AppendLine('          <tbody>')
    foreach ($r in $rows) {
        $rowNum++
        $key = $r.KPIKey.ToString().Trim()
        $kn = HtmlEsc $r.KPI_Name
        $ul = UnitLbl $r.Unit_Type
        $bd = HtmlEsc (BizDef $r.KPI_Name $r.Unit_Type $r.Value_Scale_Note)
        $formula = 'CALCULATE ( SUM ( Fact[Value] ), Dim_KPI[KPIKey] = ' + $key + ' )'
        [void]$sb.AppendLine('            <tr>')
        [void]$sb.AppendLine(('              <td class="doc-table__num">{0}</td>' -f $rowNum))
        [void]$sb.AppendLine(('              <td>{0}</td>' -f $catName))
        [void]$sb.AppendLine(('              <td>{0}</td>' -f $kn))
        [void]$sb.AppendLine(('              <td>{0}</td>' -f (HtmlEsc $ul)))
        [void]$sb.AppendLine(('              <td>{0}</td>' -f $bd))
        [void]$sb.AppendLine(('              <td><code>{0}</code></td>' -f (HtmlEsc $formula)))
        [void]$sb.AppendLine('            </tr>')
    }
    [void]$sb.AppendLine('          </tbody>')
    [void]$sb.AppendLine('        </table>')
    [void]$sb.AppendLine('      </div>')
    [void]$sb.AppendLine('    </section>')
    [void]$sb.AppendLine('')
}

$assurance = @(
    @{ Key = '501'; Name = 'Incident Key Learning implementation % (preview key)'; Unit = 'PercentOrRate'; Note = 'Synthetic key for Assurance preview; align to governed measure in production' },
    @{ Key = '502'; Name = 'FRC compliance Rate (preview key)'; Unit = 'PercentOrRate'; Note = 'Synthetic key for Assurance preview; align to governed measure in production' },
    @{ Key = '503'; Name = '% Standard Implementation - top critical Safety Standards (preview key)'; Unit = 'PercentOrRate'; Note = 'Synthetic key for Assurance preview; align to governed measure in production' }
)
$assuranceCatName = HtmlEsc 'Assurance (preview model)'
[void]$sb.AppendLine('    <section class="doc-kpi-cat" id="cat-assurance-preview">')
[void]$sb.AppendLine('      <h2>Assurance (preview model)</h2>')
[void]$sb.AppendLine('      <div class="doc-table-wrap">')
[void]$sb.AppendLine('        <table class="doc-table doc-table--wide">')
[void]$sb.AppendLine('          <thead>')
Append-Thead6 $sb
[void]$sb.AppendLine('          </thead>')
[void]$sb.AppendLine('          <tbody>')
foreach ($a in $assurance) {
    $rowNum++
    $bd = HtmlEsc (BizDef $a.Name $a.Unit $a.Note)
    $formula = 'CALCULATE ( SUM ( Fact[Value] ), Dim_KPI[KPIKey] = ' + $a.Key + ' )'
    [void]$sb.AppendLine('            <tr>')
    [void]$sb.AppendLine(('              <td class="doc-table__num">{0}</td>' -f $rowNum))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f $assuranceCatName))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f (HtmlEsc $a.Name)))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f (HtmlEsc (UnitLbl $a.Unit))))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f $bd))
    [void]$sb.AppendLine(('              <td><code>{0}</code></td>' -f (HtmlEsc $formula)))
    [void]$sb.AppendLine('            </tr>')
}
[void]$sb.AppendLine('          </tbody>')
[void]$sb.AppendLine('        </table>')
[void]$sb.AppendLine('      </div>')
[void]$sb.AppendLine('    </section>')
[void]$sb.AppendLine('')

$lvKeys = @(13, 38, 39, 40, 45, 46, 53)
$lvCatName = HtmlEsc 'Location Vulnerability (preview map)'
[void]$sb.AppendLine('    <section class="doc-kpi-cat" id="cat-location-vulnerability">')
[void]$sb.AppendLine('      <h2>Location Vulnerability (preview map)</h2>')
[void]$sb.AppendLine('      <div class="doc-table-wrap">')
[void]$sb.AppendLine('        <table class="doc-table doc-table--wide">')
[void]$sb.AppendLine('          <thead>')
Append-Thead6 $sb
[void]$sb.AppendLine('          </thead>')
[void]$sb.AppendLine('          <tbody>')
foreach ($k in $lvKeys) {
    $r = $kpis | Where-Object { [int]$_.KPIKey -eq $k } | Select-Object -First 1
    if (-not $r) { continue }
    $rowNum++
    $baseBd = BizDef $r.KPI_Name $r.Unit_Type $r.Value_Scale_Note
    $bd = HtmlEsc (($baseBd.TrimEnd().TrimEnd('.').Trim() + '. Also surfaced on the Location Vulnerability map layer for the same filtered scope (preview wireframe model).'))
    $key = $r.KPIKey.ToString().Trim()
    $formula = 'CALCULATE ( SUM ( Fact[Value] ), Dim_KPI[KPIKey] = ' + $key + ' )'
    [void]$sb.AppendLine('            <tr>')
    [void]$sb.AppendLine(('              <td class="doc-table__num">{0}</td>' -f $rowNum))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f $lvCatName))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f (HtmlEsc $r.KPI_Name)))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f (HtmlEsc (UnitLbl $r.Unit_Type))))
    [void]$sb.AppendLine(('              <td>{0}</td>' -f $bd))
    [void]$sb.AppendLine(('              <td><code>{0}</code></td>' -f (HtmlEsc $formula)))
    [void]$sb.AppendLine('            </tr>')
}
[void]$sb.AppendLine('          </tbody>')
[void]$sb.AppendLine('        </table>')
[void]$sb.AppendLine('      </div>')
[void]$sb.AppendLine('    </section>')

[System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
Write-Host "Wrote $out rows total: $rowNum"
