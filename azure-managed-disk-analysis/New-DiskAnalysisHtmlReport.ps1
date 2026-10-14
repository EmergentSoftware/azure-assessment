<#
.SYNOPSIS
    Generates an HTML report from Azure managed disk analysis JSON data.

.DESCRIPTION
    Reads the JSON produced by azure-managed-disk-telemetry-analysis.ps1 and renders a
    self-contained, interactive HTML report with sortable and filterable columns.

    Savings and severity are taken straight from the analysis output rather than being
    re-derived here, so the report can never disagree with the recommendation it displays.

.PARAMETER JsonPath
    Path to the JSON file containing disk analysis results.

.PARAMETER OutputPath
    Optional path for the output HTML file. Defaults to the JSON filename with an .html extension.

.EXAMPLE
    .\New-DiskAnalysisHtmlReport.ps1 -JsonPath "disk-analysis-20260115-123456.json"

.EXAMPLE
    .\New-DiskAnalysisHtmlReport.ps1 -JsonPath "disk-analysis.json" -OutputPath "report.html"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$JsonPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $JsonPath)) { throw "JSON file not found at: $JsonPath" }

Write-Host "Loading JSON data from: $JsonPath" -ForegroundColor Cyan
$results = @(Get-Content $JsonPath -Raw | ConvertFrom-Json)
if (-not $results.Count) { throw "No data found in JSON file or invalid format" }
Write-Host "  Loaded $($results.Count) disk records" -ForegroundColor Green

if (-not $OutputPath) {
    $jsonDir = Split-Path -Parent $JsonPath
    if (-not $jsonDir) { $jsonDir = (Get-Location).Path }
    $OutputPath = Join-Path $jsonDir "$([System.IO.Path]::GetFileNameWithoutExtension($JsonPath)).html"
}

#region Helpers

# Tolerates older result files that predate a column.
function Get-Field {
    param($Record, [string]$Name, $Default = '')
    if ($Record.PSObject.Properties[$Name]) {
        $value = $Record.$Name
        if ($null -ne $value) { return $value }
    }
    return $Default
}

function ConvertTo-HtmlText {
    param($Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Format-Money {
    param([double]$Value)
    return "`$$([math]::Round($Value, 2).ToString('N2'))"
}

#endregion

#region Projection

$severityRank = @{ Action = 0; Opportunity = 1; Watch = 2; Unknown = 3; Optimal = 4 }
$severityMeta = [ordered]@{
    Action      = @{ Class = 'sev-action'; Swatch = '#c0392b'; Label = 'Act now'; Hint = 'Spend with no workload behind it' }
    Opportunity = @{ Class = 'sev-opportunity'; Swatch = '#e6a100'; Label = 'Savings available'; Hint = 'A cheaper configuration fits the observed usage' }
    Watch       = @{ Class = 'sev-watch'; Swatch = '#2a7fc4'; Label = 'Needs review'; Hint = 'At or above its limits, or bursting repeatedly' }
    Unknown     = @{ Class = 'sev-unknown'; Swatch = '#98a1ad'; Label = 'Not assessed'; Hint = 'No telemetry, or no spec for this SKU' }
    Optimal     = @{ Class = 'sev-optimal'; Swatch = '#1e9e5a'; Label = 'Optimal'; Hint = 'Already the best fit' }
}

$currencyCode = [string](Get-Field $results[0] 'CurrencyCode' 'USD')
$pricingSource = [string](Get-Field $results[0] 'PricingSource' 'n/a')

$rows = foreach ($record in $results) {
    $severity = [string](Get-Field $record 'Severity' 'Unknown')
    if (-not $severityMeta.Contains($severity)) { $severity = 'Unknown' }
    [pscustomobject]@{
        Record   = $record
        Severity = $severity
        Rank     = $severityRank[$severity]
        Savings  = [double](Get-Field $record 'PotentialMonthlySavings' 0)
        Cost     = [double](Get-Field $record 'EstMonthlyTotalCost' 0)
        RowClass = $severityMeta[$severity].Class
        Name     = [string](Get-Field $record 'Disk')
    }
}
$rows = @($rows)

$totalDisks = $rows.Count
$totalSpend = [double](($rows | Measure-Object Cost -Sum).Sum)
$totalSavings = [double](($rows | Measure-Object Savings -Sum).Sum)
$actionable = @($rows | Where-Object { $_.Savings -gt 0 }).Count
$unattached = @($rows | Where-Object { (Get-Field $_.Record 'Attached' 'Yes') -eq 'No' }).Count
$savingsPct = if ($totalSpend -gt 0) { [math]::Round(($totalSavings / $totalSpend) * 100, 1) } else { 0 }

$severityCounts = @{}
foreach ($key in $severityMeta.Keys) {
    $severityCounts[$key] = @($rows | Where-Object { $_.Severity -eq $key }).Count
}

#endregion

#region Column definitions

# A single definition drives the header, the filter row, the cells and the CSV export.
# Html = $true means the cell builder already produced safe markup.
$columns = @(
    @{ Header = 'Disk'; Get = { param($r) ConvertTo-HtmlText (Get-Field $r 'Disk') }; Html = $true; Bold = $true }
    @{ Header = 'Resource Group'; Get = { param($r) Get-Field $r 'ResourceGroup' } }
    @{ Header = 'Region'; Get = { param($r) Get-Field $r 'Location' } }
    @{ Header = 'SKU'; Get = { param($r) Get-Field $r 'SKU' } }
    @{ Header = 'State'; Get = { param($r) Get-Field $r 'DiskState' } }
    @{ Header = 'Tier'; Get = { param($r) Get-Field $r 'BillingTier' } }
    @{ Header = 'Size GiB'; Numeric = $true; Html = $true
       Sort = { param($r) [double](Get-Field $r 'ActualSizeGiB' 0) }
       Get  = { param($r)
            $actual = ConvertTo-HtmlText (Get-Field $r 'ActualSizeGiB' 0)
            if ((Get-Field $r 'CustomSize' 'No') -eq 'Yes') {
                "$actual <span class='muted'>(bills $(ConvertTo-HtmlText (Get-Field $r 'BilledSizeGiB' 0)))</span>"
            } else { $actual }
       } }
    @{ Header = 'Metrics'; Get = { param($r) Get-Field $r 'MetricsStatus' } }
    @{ Header = 'Peak IOPS'; Numeric = $true
       Sort = { param($r) [double](Get-Field $r 'PeakIOPS' 0) }
       Get  = { param($r) '{0:N0}' -f [double](Get-Field $r 'PeakIOPS' 0) } }
    @{ Header = 'P99 IOPS'; Numeric = $true
       Sort = { param($r) [double](Get-Field $r 'P99IOPS' 0) }
       Get  = { param($r) '{0:N0}' -f [double](Get-Field $r 'P99IOPS' 0) } }
    @{ Header = 'IOPS limit'; Numeric = $true
       Sort = { param($r) [double](Get-Field $r 'EffectiveMaxIOPS' 0) }
       Get  = { param($r) '{0:N0}' -f [double](Get-Field $r 'EffectiveMaxIOPS' 0) } }
    @{ Header = 'P99 MB/s'; Numeric = $true
       Sort = { param($r) [double](Get-Field $r 'P99MBps' 0) }
       Get  = { param($r) '{0:N1}' -f [double](Get-Field $r 'P99MBps' 0) } }
    @{ Header = 'Latency ms'; Numeric = $true
       Sort = { param($r) [double](Get-Field $r 'AvgLatencyMs' 0) }
       Get  = { param($r)
            $value = [double](Get-Field $r 'AvgLatencyMs' 0)
            if ($value -le 0) { '-' } else { '{0:N2}' -f $value }
       } }
    @{ Header = 'Cost / mo'; Numeric = $true
       Sort = { param($r) [double](Get-Field $r 'EstMonthlyTotalCost' 0) }
       Get  = { param($r) Format-Money ([double](Get-Field $r 'EstMonthlyTotalCost' 0)) } }
    @{ Header = 'Savings / mo'; Numeric = $true; Html = $true
       Sort = { param($r) [double](Get-Field $r 'PotentialMonthlySavings' 0) }
       Get  = { param($r)
            $value = [double](Get-Field $r 'PotentialMonthlySavings' 0)
            if ($value -le 0) { '-' } else { "<span class='savings'>$(Format-Money $value)</span>" }
       } }
    @{ Header = 'Recommendation'; Html = $true
       Get = { param($r)
            $text = ConvertTo-HtmlText (Get-Field $r 'Decision')
            $detail = Get-Field $r 'DecisionDetail'
            if ($detail) {
                $encoded = ConvertTo-HtmlText $detail
                "$text<div class='detail' title='$encoded'>$encoded</div>"
            } else { $text }
       } }
)

#endregion

#region Render

Write-Host "Generating HTML report..." -ForegroundColor Cyan
$sb = [System.Text.StringBuilder]::new(1048576)

$generatedOn = ConvertTo-HtmlText (Get-Date -Format "dd MMMM yyyy 'at' HH:mm:ss")
$sourceName = ConvertTo-HtmlText ([System.IO.Path]::GetFileName($JsonPath))

[void]$sb.AppendLine(@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Azure Managed Disk Analysis</title>
<style>
  *, *::before, *::after { box-sizing: border-box; }
  body { margin: 0; padding: 24px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
         background: #eef1f6; color: #1b1f27; }
  .container { max-width: 1900px; margin: 0 auto; background: #fff; border-radius: 10px;
               box-shadow: 0 6px 28px rgba(16,24,40,.12); overflow: hidden; }
  header { background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%); color: #fff; padding: 28px 32px; }
  header h1 { margin: 0 0 6px; font-size: 1.85em; }
  header p { margin: 2px 0; opacity: .88; font-size: .92em; }
  .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 16px;
           padding: 24px 32px; background: #f7f9fc; border-bottom: 1px solid #e3e8ef; }
  .stat { background: #fff; border: 1px solid #e3e8ef; border-radius: 8px; padding: 16px 18px; }
  .stat .n { font-size: 1.85em; font-weight: 700; line-height: 1.1; }
  .stat .l { color: #5b6471; font-size: .78em; text-transform: uppercase; letter-spacing: .6px; margin-top: 4px; }
  .stat .sub { color: #8a929e; font-size: .76em; margin-top: 3px; }
  .legend { display: flex; flex-wrap: wrap; gap: 10px; padding: 16px 32px; border-bottom: 1px solid #e3e8ef; }
  .chip { display: flex; align-items: center; gap: 8px; border: 1px solid #e3e8ef;
          border-radius: 999px; padding: 5px 12px; font-size: .82em; }
  .swatch { width: 12px; height: 12px; border-radius: 3px; flex: none; }
  .toolbar { display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
             padding: 14px 32px; background: #f7f9fc; border-bottom: 1px solid #e3e8ef; }
  .toolbar button { border: 1px solid #cfd6e0; background: #fff; color: #27303d; border-radius: 6px;
                    padding: 7px 14px; font-size: .86em; font-weight: 600; cursor: pointer; }
  .toolbar button:hover { background: #eef2f7; }
  .toolbar .count { margin-left: auto; color: #5b6471; font-size: .86em; }
  .table-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: .84em; }
  thead th { position: sticky; top: 0; z-index: 2; background: #2f3a4a; color: #fff; text-align: left;
             padding: 11px 10px; font-size: .88em; white-space: nowrap; cursor: pointer; }
  thead th:hover { background: #3b485c; }
  thead th::after { content: ' \2195'; opacity: .35; }
  thead th.asc::after { content: ' \25B2'; opacity: 1; }
  thead th.desc::after { content: ' \25BC'; opacity: 1; }
  tr.filters th { position: sticky; top: 38px; z-index: 1; background: #f2f5f9; padding: 6px; cursor: default; }
  tr.filters th::after { content: ''; }
  tr.filters input { width: 100%; min-width: 68px; border: 1px solid #cfd6e0; border-radius: 4px;
                     padding: 5px 7px; font-size: .95em; font-family: inherit; }
  tbody td { padding: 9px 10px; border-bottom: 1px solid #edf0f5; vertical-align: top; }
  tbody tr:hover { background: #f3f7ff !important; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
  .muted { color: #8a929e; font-size: .9em; }
  .detail { color: #5b6471; font-size: .9em; margin-top: 3px; max-width: 440px;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .savings { color: #0b7a42; font-weight: 700; }
  .sev-action      { background: #fdecef; border-left: 4px solid #c0392b; }
  .sev-opportunity { background: #fff8e6; border-left: 4px solid #e6a100; }
  .sev-watch       { background: #eaf4fd; border-left: 4px solid #2a7fc4; }
  .sev-unknown     { background: #f4f5f7; border-left: 4px solid #98a1ad; }
  .sev-optimal     { background: #eefaf2; border-left: 4px solid #1e9e5a; }
  footer { padding: 18px 32px; background: #f7f9fc; color: #5b6471; font-size: .82em; border-top: 1px solid #e3e8ef; }
</style>
</head>
<body>
<div class="container">
<header>
  <h1>Azure Managed Disk Analysis</h1>
  <p>Generated $generatedOn</p>
  <p>Source: $sourceName &middot; Pricing: $(ConvertTo-HtmlText $pricingSource) ($(ConvertTo-HtmlText $currencyCode))</p>
</header>
<div class="stats">
  <div class="stat"><div class="n">$totalDisks</div><div class="l">Disks analysed</div></div>
  <div class="stat"><div class="n">$(Format-Money $totalSpend)</div><div class="l">Current spend / month</div></div>
  <div class="stat"><div class="n" style="color:#0b7a42;">$(Format-Money $totalSavings)</div><div class="l">Identified savings / month</div><div class="sub">$savingsPct% of spend &middot; $(Format-Money ($totalSavings * 12)) per year</div></div>
  <div class="stat"><div class="n">$actionable</div><div class="l">Disks with savings</div></div>
  <div class="stat"><div class="n" style="color:#c0392b;">$unattached</div><div class="l">Unattached disks</div></div>
  <div class="stat"><div class="n" style="color:#1e9e5a;">$($severityCounts['Optimal'])</div><div class="l">Already optimal</div></div>
</div>
<div class="legend">
"@)

foreach ($key in $severityMeta.Keys) {
    $meta = $severityMeta[$key]
    [void]$sb.AppendLine("<span class='chip'><span class='swatch' style='background:$($meta.Swatch)'></span><strong>$(ConvertTo-HtmlText $meta.Label)</strong>&nbsp;($($severityCounts[$key])) &mdash; $(ConvertTo-HtmlText $meta.Hint)</span>")
}

[void]$sb.AppendLine(@'
</div>
<div class="toolbar">
  <button onclick="clearFilters()">Clear filters</button>
  <button onclick="showSavingsOnly()">Show savings only</button>
  <button onclick="exportCsv()">Export filtered to CSV</button>
  <span class="count">Showing <span id="vis">0</span> of <span id="tot">0</span> disks &middot; filtered savings <span id="sum">$0.00</span></span>
</div>
<div class="table-wrap"><table id="t"><thead><tr>
'@)

for ($i = 0; $i -lt $columns.Count; $i++) {
    $cls = if ($columns[$i].ContainsKey('Numeric')) { " class='num'" } else { '' }
    [void]$sb.AppendLine("<th$cls onclick='sortTable($i)'>$(ConvertTo-HtmlText $columns[$i].Header)</th>")
}
[void]$sb.AppendLine('</tr><tr class="filters">')
for ($i = 0; $i -lt $columns.Count; $i++) {
    [void]$sb.AppendLine("<th><input type='text' placeholder='filter' oninput='filterTable()'></th>")
}
[void]$sb.AppendLine('</tr></thead><tbody>')

# Highest-severity and highest-savings rows first, so the report opens on what matters.
$ordered = $rows | Sort-Object Rank, @{ Expression = 'Savings'; Descending = $true }, Name
foreach ($row in $ordered) {
    [void]$sb.Append("<tr class='$($row.RowClass)' data-savings='$($row.Savings)'>")
    foreach ($col in $columns) {
        $value = & $col.Get $row.Record
        $cell = if ($col.ContainsKey('Html')) { [string]$value } else { ConvertTo-HtmlText $value }
        if ($col.ContainsKey('Bold')) { $cell = "<strong>$cell</strong>" }
        if ($col.ContainsKey('Numeric')) {
            $sortValue = if ($col.ContainsKey('Sort')) { & $col.Sort $row.Record } else { 0 }
            [void]$sb.Append("<td class='num' data-sort='$sortValue'>$cell</td>")
        } else {
            [void]$sb.Append("<td>$cell</td>")
        }
    }
    [void]$sb.AppendLine('</tr>')
}

[void]$sb.AppendLine(@"
</tbody></table></div>
<footer>
  Savings and severity come from azure-managed-disk-telemetry-analysis.ps1 and are not recalculated here.
  Figures are retail estimates in $(ConvertTo-HtmlText $currencyCode) and exclude snapshots, egress and negotiated discounts.
</footer>
</div>
<script>
const table = document.getElementById('t');
const tbody = table.tBodies[0];
const dirs = {};

function cellSortValue(cell) {
  const attr = cell.getAttribute('data-sort');
  if (attr !== null) return parseFloat(attr) || 0;
  return cell.textContent.trim().toLowerCase();
}

function sortTable(i) {
  dirs[i] = dirs[i] === 'asc' ? 'desc' : 'asc';
  const asc = dirs[i] === 'asc';
  table.querySelectorAll('thead tr:first-child th').forEach((th, idx) => {
    th.classList.remove('asc', 'desc');
    if (idx === i) th.classList.add(asc ? 'asc' : 'desc');
  });
  const rows = Array.from(tbody.rows);
  rows.sort((a, b) => {
    const x = cellSortValue(a.cells[i]), y = cellSortValue(b.cells[i]);
    if (typeof x === 'number' && typeof y === 'number') return asc ? x - y : y - x;
    return asc ? String(x).localeCompare(String(y)) : String(y).localeCompare(String(x));
  });
  rows.forEach(r => tbody.appendChild(r));
}

function refreshCounts() {
  let visible = 0, savings = 0;
  Array.from(tbody.rows).forEach(row => {
    if (row.style.display !== 'none') { visible++; savings += parseFloat(row.dataset.savings) || 0; }
  });
  document.getElementById('vis').textContent = visible;
  document.getElementById('tot').textContent = tbody.rows.length;
  document.getElementById('sum').textContent = '$' + savings.toFixed(2);
}

function filterTable() {
  const inputs = Array.from(table.querySelectorAll('tr.filters input')).map(i => i.value.toLowerCase());
  Array.from(tbody.rows).forEach(row => {
    const show = inputs.every((f, i) => !f || row.cells[i].textContent.toLowerCase().includes(f));
    row.style.display = show ? '' : 'none';
  });
  refreshCounts();
}

function clearFilters() {
  table.querySelectorAll('tr.filters input').forEach(i => i.value = '');
  Array.from(tbody.rows).forEach(r => r.style.display = '');
  refreshCounts();
}

function showSavingsOnly() {
  table.querySelectorAll('tr.filters input').forEach(i => i.value = '');
  Array.from(tbody.rows).forEach(row => {
    row.style.display = (parseFloat(row.dataset.savings) || 0) > 0 ? '' : 'none';
  });
  refreshCounts();
}

function exportCsv() {
  const headers = Array.from(table.querySelectorAll('thead tr:first-child th'))
    .map(th => th.textContent.replace(/[\u2195\u25B2\u25BC]/g, '').trim());
  const esc = s => '"' + String(s).replace(/"/g, '""') + '"';
  const lines = [headers.map(esc).join(',')];
  Array.from(tbody.rows).filter(r => r.style.display !== 'none').forEach(r => {
    lines.push(Array.from(r.cells).map(c => esc(c.innerText.replace(/\s+/g, ' ').trim())).join(','));
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'disk-analysis-filtered-' + new Date().toISOString().slice(0, 10) + '.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

refreshCounts();
</script>
</body>
</html>
"@)

[System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))

#endregion

Write-Host "  HTML report generated: $OutputPath" -ForegroundColor Green
Write-Host ("  {0} disks | {1}/month current | {2}/month identified savings ({3}%)" -f `
        $totalDisks, (Format-Money $totalSpend), (Format-Money $totalSavings), $savingsPct) -ForegroundColor Cyan
