<#
.SYNOPSIS
    Analyzes Azure managed disk performance telemetry to identify optimization opportunities.

.DESCRIPTION
    This script retrieves Azure managed disk metrics from Azure Monitor and analyzes their performance over a specified time window.
    It identifies disks that could benefit from tier changes (upgrade/downgrade) based on:
    - Performance metrics (IOPS, throughput, latency)
    - Cost optimization opportunities
    - Utilization patterns

    Pricing is resolved per disk from the Azure Retail Prices API using the disk's own region and
    redundancy (LRS or ZRS), covering capacity, transaction, Premium on-demand burst, Premium SSD v2
    provisioned meters, and one-year Premium SSD reservations. If the API cannot be reached the
    script falls back to the rates in disk-specs.json and records that in the PricingSource column.

    The script generates JSON and CSV reports with detailed metrics and recommendations.
    Use New-DiskAnalysisHtmlReport.ps1 to generate an interactive HTML report from the JSON output.

.PARAMETER DaysToInspect
    Number of days of historical metrics to analyze. Default is 3 days.

.PARAMETER ExceedanceThreshold
    The percentage threshold (0-1) for determining if a disk consistently exceeds its limits.
    For example, 0.05 means if more than 5% of samples exceed the limit, it triggers a recommendation.
    Default is 0.05 (5%).

.PARAMETER SubscriptionId
    One or more Azure subscription IDs to analyze. If not specified, uses the current subscription context.

.PARAMETER SpecFile
    Path to the JSON file containing disk SKU specifications (IOPS and throughput limits by size)
    and the fallback billing rates used when live pricing is unavailable.
    Default is "disk-specs.json" in the script directory.

.PARAMETER PricingRegion
    Overrides the Azure region used for pricing lookups. By default each disk is priced in its
    own region, which matters because disk prices vary by region (for example a P30 is $135.17
    per month in North Central US but $122.88 in West US 2).

.PARAMETER Currency
    ISO currency code passed to the Azure Retail Prices API. Default is "USD".

.PARAMETER NoLivePricing
    Skips the Azure Retail Prices API and uses the fallback rates in disk-specs.json instead.
    Those rates are North Central US LRS only, so results for other regions will be approximate.
    Every output record carries PricingSource so the basis for each figure is visible.

.PARAMETER OutputPrefix
    Prefix for output files. Two files will be generated:
    - {OutputPrefix}-{timestamp}.json - Full analysis results in JSON format
    - {OutputPrefix}-{timestamp}.csv - Analysis results in CSV format
    Default is "disk-analysis".
    
    To generate an HTML report, run:
    .\New-DiskAnalysisHtmlReport.ps1 -JsonPath "{OutputPrefix}-{timestamp}.json"

.PARAMETER ExcludeAvd
    When specified, disks whose name contains "AVD" (case-insensitive) are excluded from the analysis.

.EXAMPLE
    .\azure-managed-disk-telemetry-analysis.ps1
    
    Analyzes all disks in the current subscription using default settings.
    Outputs JSON and CSV files with timestamp.

.EXAMPLE
    .\azure-managed-disk-telemetry-analysis.ps1 -SubscriptionId "12345678-1234-1234-1234-123456789012" -DaysToInspect 7
    
    Analyzes disks in the specified subscription over the last 7 days.

.EXAMPLE
    .\azure-managed-disk-telemetry-analysis.ps1 -ExceedanceThreshold 0.10 -OutputPrefix "weekly-report"
    
    Uses a 10% exceedance threshold and generates reports with custom filenames.
    
.EXAMPLE
    .\azure-managed-disk-telemetry-analysis.ps1 -ExcludeAvd
    
    Excludes any disks whose name contains "AVD" from the analysis.

.EXAMPLE
    .\azure-managed-disk-telemetry-analysis.ps1 -PricingRegion westus2 -Currency EUR

    Prices every disk against West US 2 retail rates in euros, regardless of where the disks live.

.EXAMPLE
    .\azure-managed-disk-telemetry-analysis.ps1 -NoLivePricing

    Runs entirely offline from the Retail Prices API using the fallback rates in disk-specs.json.

.EXAMPLE
    # Run analysis and immediately generate HTML report
    .\azure-managed-disk-telemetry-analysis.ps1
    .\New-DiskAnalysisHtmlReport.ps1 -JsonPath "disk-analysis-20250115-120000.json"

.NOTES
    Requires Azure CLI (az) to be installed and authenticated; the CLI is used only to acquire
    access tokens. Discovery runs through Azure Resource Graph and metrics through the Azure
    Monitor batch endpoint, so a run costs a handful of REST calls rather than two per disk.

    Required permissions: Reader on the target subscriptions and Monitoring Reader for metrics.
    Tokens are requested per subscription, which lets a single run span multiple tenants.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 365)]
    [int]$DaysToInspect = 7,
    [ValidateRange(0, 1)]
    [double]$ExceedanceThreshold = 0.05,
    [string[]]$SubscriptionId,
    [string]$SpecFile = "disk-specs.json",
    [string]$OutputPrefix = "disk-analysis",
    [switch]$ExcludeAvd,
    [string]$PricingRegion,
    [string]$Currency = "USD",
    [switch]$NoLivePricing
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helper Functions

# Offered managed-disk tier numbers and their billed sizes in GiB, index-aligned.
$script:DiskTierNumbers = @(1, 2, 3, 4, 6, 10, 15, 20, 30, 40, 50, 60, 70, 80)
$script:DiskTierSizes = @(4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32767)

<#
.SYNOPSIS
    Calculates the percentage of samples where a property exceeds a limit.
#>
function Get-PercentAbove {
    param($Samples, [string]$Property, [double]$Limit)
    if (-not $Samples -or -not $Samples.Count) { return 0 }
    @($Samples | Where-Object { $_.$Property -gt $Limit }).Count / $Samples.Count
}

<#
.SYNOPSIS
    Reads an aggregation value from an Azure Monitor data point.
.DESCRIPTION
    Azure Monitor omits the aggregation property entirely for intervals with no data,
    so the property cannot be dereferenced directly.
#>
function Get-MetricValue {
    param($DataPoint, [string]$Aggregation = 'average')
    if ($DataPoint -and $DataPoint.PSObject.Properties[$Aggregation]) { return $DataPoint.$Aggregation }
    return $null
}

<#
.SYNOPSIS
    Returns the requested percentile of a sample property.
#>
function Get-SamplePercentile {
    param($Samples, [string]$Property, [double]$Percentile)
    if (-not $Samples -or $Samples.Count -eq 0) { return 0 }
    $sorted = @($Samples | ForEach-Object { [double]$_.$Property } | Sort-Object)
    $index = [int][math]::Ceiling(($Percentile / 100) * $sorted.Count) - 1
    if ($index -lt 0) { $index = 0 }
    if ($index -ge $sorted.Count) { $index = $sorted.Count - 1 }
    return $sorted[$index]
}

<#
.SYNOPSIS
    Builds a performance spec for disk types that carry explicitly provisioned limits.
.DESCRIPTION
    Premium SSD v2 and Ultra have no fixed tier ladder: capacity, IOPS and throughput are set
    independently on the disk, so the provisioned values on the resource are the real limits.
    Unset values fall back to each type's free baseline.
#>
function Get-ProvisionedSpec {
    param($Disk, [string]$Tier)
    
    $isUltra = $Tier -like 'UltraSSD*'
    $baselineIops = if ($isUltra) { 100 } else { 3000 }
    $baselineMBps = if ($isUltra) { 1 } else { 125 }
    
    $iops = if ($Disk.PSObject.Properties['expandedIops'] -and $Disk.expandedIops) { [double]$Disk.expandedIops } else { $baselineIops }
    $mbps = if ($Disk.PSObject.Properties['expandedMBps'] -and $Disk.expandedMBps) { [double]$Disk.expandedMBps } else { $baselineMBps }
    
    return [pscustomobject]@{
        Sku          = if ($isUltra) { 'Ultra' } else { 'PremiumV2' }
        MaxIOPS      = $iops
        MaxMBps      = $mbps
        BaselineIOPS = $baselineIops
        BaselineMBps = $baselineMBps
    }
}

# Drives both the console colour and the HTML report's row grouping.
$script:DecisionSeverityColor = @{
    Action      = 'Magenta'
    Opportunity = 'Yellow'
    Watch       = 'Cyan'
    Optimal     = 'Green'
    Unknown     = 'DarkGray'
}

<#
.SYNOPSIS
    Packages a recommendation with the money attached to it and how urgent it is.
.DESCRIPTION
    Every branch of the recommendation logic returns one of these, so the console line is
    rendered in exactly one place and the HTML report can consume the savings and severity
    directly instead of re-deriving them from the decision text.
#>
function New-DiskDecision {
    param(
        [Parameter(Mandatory)][string]$Decision,
        [double]$Savings = 0,
        [ValidateSet('Action', 'Opportunity', 'Watch', 'Optimal', 'Unknown')][string]$Severity = 'Optimal',
        [string]$Detail = ''
    )
    return [pscustomobject]@{
        Decision = $Decision
        # 0.0 rather than 0: an int literal picks Max(int,int) and would truncate the cents away.
        Savings  = [math]::Round([math]::Max(0.0, $Savings), 2)
        Severity = $Severity
        Detail   = $Detail
    }
}

<#
.SYNOPSIS
    Formats a monthly cost delta as either a saving or an increase.
#>
function Format-CostDelta {
    param([double]$Delta)
    if ($Delta -lt 0) { return "saves `$$([math]::Abs($Delta).ToString('F2'))/month" }
    return "+`$$($Delta.ToString('F2'))/month"
}

<#
.SYNOPSIS
    Retrieves the specification (IOPS/throughput limits) for a specific disk size from the spec table.
#>
function Get-SkuSpec {
    param($Table, [int]$SizeGiB)
    if (-not $Table) { return $null }
    $prop = $Table.PSObject.Properties[$SizeGiB.ToString()]
    if ($prop) { return $prop.Value }
    return $null
}

<#
.SYNOPSIS
    Determines the billing tier for a disk based on its actual size.
.DESCRIPTION
    Azure rounds the provisioned size up to the next offered SKU size, so a 100 GiB
    Standard SSD bills as E10 (128 GiB). Standard HDD has no tier below S4 (32 GiB),
    whereas Standard SSD and Premium SSD start at 4 GiB.
#>
function Get-BillingTier {
    param([string]$DiskType, [int]$ActualSizeGB)

    $prefix = switch -Regex ($DiskType) {
        '^StandardSSD' { 'E'; break }
        '^Premium' { 'P'; break }
        default { 'S' }
    }
    $startIndex = if ($prefix -eq 'S') { 3 } else { 0 }

    for ($i = $startIndex; $i -lt $script:DiskTierSizes.Count; $i++) {
        if ($ActualSizeGB -le $script:DiskTierSizes[$i]) {
            return @{
                TierName     = "$prefix$($script:DiskTierNumbers[$i])"
                BilledSizeGB = $script:DiskTierSizes[$i]
                ActualSizeGB = $ActualSizeGB
                IsCustomSize = ($ActualSizeGB -ne $script:DiskTierSizes[$i])
            }
        }
    }

    return @{
        TierName     = "Custom"
        BilledSizeGB = $ActualSizeGB
        ActualSizeGB = $ActualSizeGB
        IsCustomSize = $true
    }
}

<#
.SYNOPSIS
    Invokes Azure CLI and returns raw stdout output.
#>
function Invoke-AzCliRaw {
    param([string[]]$Arguments)
    # Get the full path to az.exe to work with UseShellExecute = false
    $azPath = (Get-Command az -ErrorAction Stop).Source
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $azPath
    
    # Add all arguments including common flags
    foreach ($arg in ($Arguments + @("--only-show-errors","--output","json"))) {
        $psi.ArgumentList.Add($arg)
    }
    
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    
    $process = [System.Diagnostics.Process]::Start($psi)
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    
    if ($process.ExitCode -ne 0) {
        $detail = if ($stderr) { " - $($stderr.Trim())" } else { "" }
        throw "az $($Arguments -join ' ') failed with exit code $($process.ExitCode)$detail"
    }
    return $stdout
}

<#
.SYNOPSIS
    Invokes Azure CLI and returns parsed JSON output.
#>
function Invoke-AzCliJson {
    param([string[]]$Arguments)
    $json = Invoke-AzCliRaw -Arguments $Arguments
    if ([string]::IsNullOrWhiteSpace($json)) { return $null }
    return $json | ConvertFrom-Json
}

#endregion

#region Azure REST

$script:AzTokenCache = @{}
$script:ArmEndpoint = 'https://management.azure.com'
$script:MetricsAudience = 'https://metrics.monitor.azure.com'

# Azure Monitor accepts up to 50 resource IDs per batch request.
$script:MetricsBatchSize = 50

<#
.SYNOPSIS
    Acquires and caches an Azure AD access token for a resource audience.
.DESCRIPTION
    Scoping the token request to a subscription makes the CLI issue it from that
    subscription's own tenant, which is what allows a single run to span tenants.
#>
function Get-AzAccessToken {
    param(
        [Parameter(Mandatory)][string]$Resource,
        [string]$SubscriptionId
    )
    
    $key = "$Resource|$SubscriptionId"
    if ($script:AzTokenCache.ContainsKey($key)) {
        $cached = $script:AzTokenCache[$key]
        if ($cached.ExpiresOn -gt (Get-Date).AddMinutes(5)) { return $cached.Token }
    }
    
    $cliArgs = @("account", "get-access-token", "--resource", $Resource)
    if ($SubscriptionId) { $cliArgs += @("--subscription", $SubscriptionId) }
    $result = Invoke-AzCliJson $cliArgs
    
    $expires = (Get-Date).AddMinutes(45)
    if ($result.PSObject.Properties['expires_on'] -and $result.expires_on) {
        $expires = [DateTimeOffset]::FromUnixTimeSeconds([long]$result.expires_on).LocalDateTime
    }
    
    $script:AzTokenCache[$key] = [pscustomobject]@{ Token = $result.accessToken; ExpiresOn = $expires }
    return $result.accessToken
}

<#
.SYNOPSIS
    Issues an authenticated Azure REST call, retrying on throttling and transient server errors.
#>
function Invoke-AzRestCall {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        $Body,
        [string]$Resource = $script:ArmEndpoint,
        [string]$SubscriptionId,
        [int]$MaxAttempts = 5,
        [int]$TimeoutSec = 300
    )
    
    for ($attempt = 1; ; $attempt++) {
        try {
            $params = @{
                Uri        = $Uri
                Method     = $Method
                Headers    = @{ Authorization = "Bearer $(Get-AzAccessToken -Resource $Resource -SubscriptionId $SubscriptionId)" }
                TimeoutSec = $TimeoutSec
            }
            if ($null -ne $Body) {
                $params.Body = ($Body | ConvertTo-Json -Depth 6 -Compress)
                $params.ContentType = 'application/json'
            }
            return Invoke-RestMethod @params
        }
        catch {
            $response = $_.Exception.Response
            $status = if ($response) { [int]$response.StatusCode } else { 0 }
            $retryable = $status -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -ge $MaxAttempts) { throw }
            
            $delay = [math]::Pow(2, $attempt)
            if ($response -and $response.Headers.RetryAfter -and $response.Headers.RetryAfter.Delta) {
                $delay = $response.Headers.RetryAfter.Delta.TotalSeconds
            }
            Write-Verbose "HTTP $status on attempt $attempt; retrying in $delay s."
            Start-Sleep -Seconds ([math]::Min($delay, 60))
        }
    }
}

<#
.SYNOPSIS
    Runs an Azure Resource Graph query across subscriptions, following result pagination.
#>
function Invoke-ResourceGraphQuery {
    param(
        [Parameter(Mandatory)][string]$Query,
        [Parameter(Mandatory)][string[]]$Subscriptions
    )
    
    $rows = [System.Collections.Generic.List[object]]::new()
    $skipToken = $null
    
    do {
        $options = @{ '$top' = 1000 }
        if ($skipToken) { $options['$skipToken'] = $skipToken }
        $payload = @{ subscriptions = @($Subscriptions); query = $Query; options = $options }
        
        $response = Invoke-AzRestCall -Method 'POST' -SubscriptionId $Subscriptions[0] `
            -Uri "$script:ArmEndpoint/providers/Microsoft.ResourceGraph/resources?api-version=2022-10-01" -Body $payload
        
        if ($response.PSObject.Properties['data'] -and $response.data) {
            $rows.AddRange([object[]]$response.data)
        }
        $skipToken = if ($response.PSObject.Properties['`$skipToken']) { $response.'$skipToken' } else { $null }
    } while ($skipToken)
    
    return ,$rows
}

<#
.SYNOPSIS
    Fetches Azure Monitor metrics for many resources at once using the metrics batch endpoint.
.DESCRIPTION
    The batch endpoint is regional and requires every resource in a request to share both a
    subscription and a region, so callers group resources accordingly. Returns a hashtable
    keyed by lowercase resource ID; each value maps a metric name to its timeseries array.
#>
function Invoke-MetricsBatch {
    param(
        [Parameter(Mandatory)][string[]]$ResourceIds,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$Region,
        [Parameter(Mandatory)][string]$MetricNamespace,
        [Parameter(Mandatory)][string[]]$MetricNames,
        [string]$Filter
    )
    
    $results = @{}
    $utcFormat = 'yyyy-MM-ddTHH:mm:ssZ'
    $query = "api-version=2024-02-01" +
             "&starttime=$($analysisWindowStart.ToString($utcFormat))" +
             "&endtime=$($analysisWindowEnd.ToString($utcFormat))" +
             "&interval=$metricInterval" +
             "&aggregation=average" +
             "&metricnamespace=$([uri]::EscapeDataString($MetricNamespace))" +
             "&metricnames=$([uri]::EscapeDataString($MetricNames -join ','))"
    if ($Filter) { $query += "&filter=$([uri]::EscapeDataString($Filter))" }
    
    $uri = "https://$Region.metrics.monitor.azure.com/subscriptions/$SubscriptionId/metrics:getBatch?$query"
    
    for ($offset = 0; $offset -lt $ResourceIds.Count; $offset += $script:MetricsBatchSize) {
        $chunk = @($ResourceIds | Select-Object -Skip $offset -First $script:MetricsBatchSize)
        try {
            $response = Invoke-AzRestCall -Method 'POST' -Uri $uri -Body @{ resourceids = $chunk } `
                -Resource $script:MetricsAudience -SubscriptionId $SubscriptionId
        }
        catch {
            Write-Warning "Metrics batch failed for $($chunk.Count) resource(s) in $Region/$SubscriptionId - $($_.Exception.Message)"
            continue
        }
        
        if (-not $response.PSObject.Properties['values']) { continue }
        foreach ($entry in $response.values) {
            $byMetric = @{}
            foreach ($metric in $entry.value) {
                $byMetric[[string]$metric.name.value] = $metric.timeseries
            }
            $results[[string]$entry.resourceid.ToLowerInvariant()] = $byMetric
        }
    }
    
    return $results
}

#endregion


$script:DiskPriceCatalogCache = @{}

# Billing treats a month as 730 hours; hourly meters (Premium SSD v2, Ultra) are scaled by this.
$script:HoursPerMonth = 730

# Maps Retail Prices API productName values onto the short family keys used in the catalog.
$script:DiskProductFamilies = @{
    'Standard HDD Managed Disks' = 'HDD'
    'Standard SSD Managed Disks' = 'SSD'
    'Premium SSD Managed Disks'  = 'PREM'
    'Azure Premium SSD v2'       = 'PREMV2'
    'Ultra Disks'                = 'ULTRA'
}

<#
.SYNOPSIS
    Queries the Azure Retail Prices API, following pagination and retrying on throttling.
#>
function Invoke-RetailPriceQuery {
    param(
        [Parameter(Mandatory)][string]$Filter,
        [string]$CurrencyCode = 'USD'
    )
    
    $items = [System.Collections.Generic.List[object]]::new()
    $uri = "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview&currencyCode='$CurrencyCode'&`$filter=" + [uri]::EscapeDataString($Filter)
    $attempt = 0
    
    while ($uri) {
        try {
            $response = Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec 60
            $attempt = 0
        }
        catch {
            $attempt++
            if ($attempt -ge 4) { throw }
            Start-Sleep -Seconds (2 * $attempt)
            continue
        }
        
        if ($response.PSObject.Properties['Items'] -and $response.Items) {
            $items.AddRange([object[]]$response.Items)
        }
        $uri = if ($response.PSObject.Properties['NextPageLink']) { $response.NextPageLink } else { $null }
    }
    
    # Comma prevents PowerShell unrolling an empty list to $null.
    return ,$items
}

<#
.SYNOPSIS
    Records the currently effective price for a catalog key.
.DESCRIPTION
    The API publishes a meter once per price revision, so the same meterName can appear
    several times with different effectiveStartDate values. The most recently effective row
    is the one being billed today; selecting by lowest price instead silently returns a
    superseded rate (for example Standard HDD transactions at the retired $0.0005/10K).
#>
function Set-CatalogPrice {
    param($Catalog, [string]$Key, $Item, $Price = $null)
    
    $value = if ($null -ne $Price) { $Price } else { $Item.retailPrice }
    if ($null -eq $value) { return }
    
    $effective = [datetime]::MinValue
    if ($Item.PSObject.Properties['effectiveStartDate'] -and $Item.effectiveStartDate) {
        [void][datetime]::TryParse([string]$Item.effectiveStartDate, [ref]$effective)
    }
    
    if (-not $Catalog.Prices.ContainsKey($Key) -or $effective -ge $Catalog.Effective[$Key]) {
        $Catalog.Prices[$Key] = [double]$value
        $Catalog.Effective[$Key] = $effective
    }
}

<#
.SYNOPSIS
    Builds a region-scoped price catalog for every managed disk meter, cached per region.
.DESCRIPTION
    Covers capacity (LRS and ZRS), transaction, Premium on-demand burst, Premium SSD v2 and
    Ultra provisioned meters, plus one-year disk reservations. Returns $null when the region
    cannot be priced so callers can fall back to the spec file.

    Capacity meters are matched on meterName rather than skuName because each SKU also
    publishes a "<SKU> <redundancy> Disk Mount" shared-disk meter at a fraction of the
    capacity price; matching on skuName alone silently selects the wrong one.
#>
function Get-DiskPriceCatalog {
    param([Parameter(Mandatory)][string]$Region)
    
    if ($script:DiskPriceCatalogCache.ContainsKey($Region)) { return $script:DiskPriceCatalogCache[$Region] }
    
    $catalog = $null
    try {
        $productClause = "(contains(productName,'Managed Disks') or productName eq 'Azure Premium SSD v2' or productName eq 'Ultra Disks')"
        $consumption = Invoke-RetailPriceQuery -CurrencyCode $Currency -Filter "armRegionName eq '$Region' and priceType eq 'Consumption' and $productClause"
        
        if (-not $consumption.Count) { throw "No consumption meters returned for region '$Region'." }
        
        $catalog = [pscustomobject]@{
            Region     = $Region
            Source     = 'RetailAPI'
            AsOf       = (Get-Date).ToString('yyyy-MM-dd')
            Currency   = $Currency
            MeterCount = 0
            Prices     = @{}
            Effective  = @{}
        }
        
        foreach ($item in $consumption) {
            $family = $script:DiskProductFamilies[[string]$item.productName]
            if (-not $family) { continue }
            $meter = [string]$item.meterName
            
            switch -Regex ($meter) {
                '^(?<sku>[SEP]\d+) (?<red>LRS|ZRS) Disk$' {
                    Set-CatalogPrice $catalog "cap|$($Matches.sku)|$($Matches.red)" $item
                    break
                }
                '^(Disk Transactions |[SEP]\d+ )?(?<red>LRS|ZRS) Disk Operations$' {
                    Set-CatalogPrice $catalog "txn|$family|$($Matches.red)" $item
                    break
                }
                '^(?<red>LRS|ZRS) Burst Enablement$' {
                    Set-CatalogPrice $catalog "burstenable|$($Matches.red)" $item
                    break
                }
                '^(?<red>LRS|ZRS) Burst Transactions$' {
                    Set-CatalogPrice $catalog "bursttxn|$($Matches.red)" $item
                    break
                }
                '^\w+ (?<red>LRS|ZRS) Provisioned Capacity$' {
                    Set-CatalogPrice $catalog "provcap|$family|$($Matches.red)" $item
                    break
                }
                '^\w+ (?<red>LRS|ZRS) Provisioned IOPS$' {
                    Set-CatalogPrice $catalog "proviops|$family|$($Matches.red)" $item
                    break
                }
                '^\w+ (?<red>LRS|ZRS) Provisioned Throughput \(MBps\)$' {
                    Set-CatalogPrice $catalog "provmbps|$family|$($Matches.red)" $item
                    break
                }
            }
        }
        
        # Disk reservations exist for Premium SSD P30-P80 on a one-year term only.
        try {
            $reservations = Invoke-RetailPriceQuery -CurrencyCode $Currency -Filter "armRegionName eq '$Region' and priceType eq 'Reservation' and contains(productName,'Managed Disks')"
            foreach ($item in $reservations) {
                if ($item.reservationTerm -ne '1 Year') { continue }
                if ([string]$item.meterName -match '^(?<sku>P\d+) (?<red>LRS|ZRS) Disk$') {
                    # retailPrice is the full one-year commitment, so divide to a monthly equivalent.
                    Set-CatalogPrice $catalog "rsv1y|$($Matches.sku)|$($Matches.red)" $item ([double]$item.retailPrice / 12)
                }
            }
        }
        catch {
            Write-Verbose "Reservation pricing unavailable for '$Region': $($_.Exception.Message)"
        }
        
        $catalog.MeterCount = $catalog.Prices.Count
        Write-Host "  Retail pricing loaded for '$Region' ($($catalog.MeterCount) meters, $Currency)" -ForegroundColor DarkGray
    }
    catch {
        $catalog = $null
        Write-Warning "Live pricing unavailable for region '$Region' - falling back to disk-specs.json rates. ($($_.Exception.Message))"
    }
    
    $script:DiskPriceCatalogCache[$Region] = $catalog
    return $catalog
}

<#
.SYNOPSIS
    Resolves the monthly capacity cost for a disk tier, preferring live retail pricing.
.DESCRIPTION
    Falls back to the MonthlyCost baked into disk-specs.json, applying the ZRS uplift because
    the spec file only carries LRS rates for a single region.
#>
function Get-CapacityMonthlyCost {
    param(
        $Catalog,
        [string]$TierName,
        [string]$Redundancy,
        $SpecEntry,
        [int]$SizeGiB,
        [double]$FallbackPerGiB = 0
    )
    
    if ($Catalog -and $Catalog.Prices.ContainsKey("cap|$TierName|$Redundancy")) {
        return $Catalog.Prices["cap|$TierName|$Redundancy"]
    }
    
    if ($SpecEntry -and $SpecEntry.PSObject.Properties['MonthlyCost'] -and $SpecEntry.MonthlyCost) {
        $cost = [double]$SpecEntry.MonthlyCost
        if ($Redundancy -eq 'ZRS') { $cost *= $script:ZrsUpliftFactor }
        return $cost
    }
    
    return $SizeGiB * $FallbackPerGiB
}

<#
.SYNOPSIS
    Resolves a non-capacity rate (transactions, bursting, provisioned performance) with a spec fallback.
#>
function Get-CatalogRate {
    param($Catalog, [string]$Key, [double]$Fallback)
    if ($Catalog -and $Catalog.Prices.ContainsKey($Key)) { return $Catalog.Prices[$Key] }
    return $Fallback
}

#endregion

#region Initialization

# Start timing
$scriptStartTime = Get-Date

# Load disk specifications from JSON file
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$specPath = if ([System.IO.Path]::IsPathRooted($SpecFile)) { $SpecFile } else { Join-Path $scriptDir $SpecFile }
if (-not (Test-Path $specPath)) { throw "Spec file not found at $specPath" }
$spec = Get-Content $specPath -Raw | ConvertFrom-Json
$hddSpecs = $spec.StandardHDD
$ssdSpecs = $spec.StandardSSD
$premiumSpecs = $spec.PremiumSSD
$premiumV2Specs = $spec.PremiumSSDv2

# Use current subscription if none specified
if (-not $SubscriptionId -or -not $SubscriptionId.Count) {
    $SubscriptionId = @((Invoke-AzCliJson @("account","show")).id)
}

# Define the analysis time window. Azure Monitor timestamps are UTC, so the window must be too.
$analysisWindowEnd = (Get-Date).ToUniversalTime()
$analysisWindowStart = $analysisWindowEnd.AddDays(-$DaysToInspect)

# Metrics to collect from Azure Monitor
$diskMetricNames = @(
    "Composite Disk Read Operations/sec",
    "Composite Disk Write Operations/sec",
    "Composite Disk Read Bytes/sec",
    "Composite Disk Write Bytes/sec",
    "DiskPaidBurstIOPS"
)

# Billing rates come from the Rates block in disk-specs.json so they can be refreshed without touching code.
$rates = $spec.Rates
if (-not $rates) { throw "Spec file $specPath is missing the required 'Rates' section." }
$script:ZrsUpliftFactor = [double]$rates.ZrsUpliftFactor

# These are only used when the Retail Prices API cannot be reached for a region.
$fallbackHddTransactionPer10k = [double]$rates.HddTransactionPer10k
$fallbackSsdTransactionPer10k = [double]$rates.SsdTransactionPer10k
$fallbackBurstEnablementMonthly = [double]$rates.BurstEnablementMonthly
$fallbackBurstTransactionPer10k = [double]$rates.BurstTransactionPer10k

if ($NoLivePricing) {
    Write-Host "Pricing basis: disk-specs.json fallback rates ($($rates.Region), as of $($rates.AsOf)) - live pricing disabled" -ForegroundColor Yellow
} else {
    Write-Host "Pricing basis: Azure Retail Prices API ($Currency), per-disk region" -ForegroundColor Gray
}

# Interval selection honors Azure Monitor retention limits: PT1M=30d, PT5M=90d, PT1H=93d.
$metricInterval, $secondsPerSample = if ($DaysToInspect -gt 89) {
    "PT1H", 3600
} elseif ($DaysToInspect -gt 29) {
    "PT5M", 300
} else {
    "PT1M", 60
}

#endregion

#region Telemetry Collection

<#
.SYNOPSIS
    Converts prefetched Azure Monitor timeseries for one disk into IOPS/throughput samples.
.DESCRIPTION
    Metrics for every disk are retrieved up front in batches, so this is pure computation
    over the cached response rather than a per-disk API call.
#>
function Get-DiskTelemetry {
    param([string]$ResourceId)

    $emptyResult = [pscustomobject]@{ Samples = @(); Transactions = 0; BurstOperations = 0 }
    $key = $ResourceId.ToLowerInvariant()
    if (-not $script:DiskMetricsCache.ContainsKey($key)) { return $emptyResult }
    
    # Flatten the cached timeseries into a metric-name keyed map of data points.
    $map = @{}
    foreach ($entry in $script:DiskMetricsCache[$key].GetEnumerator()) {
        if ($entry.Value -and $entry.Value.Count -gt 0) { $map[$entry.Key] = $entry.Value[0].data }
    }
    
    # Return empty result if required metrics are not available
    $requiredMetrics = @(
        "Composite Disk Read Operations/sec",
        "Composite Disk Write Operations/sec",
        "Composite Disk Read Bytes/sec",
        "Composite Disk Write Bytes/sec"
    )
    
    foreach ($metricName in $requiredMetrics) {
        if (-not $map.ContainsKey($metricName) -or -not $map[$metricName]) {
            Write-Host "      WARNING: Missing metric data for '$metricName'" -ForegroundColor Yellow
            Write-Host "               Possible causes: Disk unattached, VM stopped/deallocated, or newly created disk" -ForegroundColor DarkGray
            return $emptyResult
        }
    }
    
    # Combine read and write metrics into IOPS and MBps snapshots
    # Find the minimum count across all metrics to avoid index out of bounds
    $count = @(
        $map["Composite Disk Read Operations/sec"].Count,
        $map["Composite Disk Write Operations/sec"].Count,
        $map["Composite Disk Read Bytes/sec"].Count,
        $map["Composite Disk Write Bytes/sec"].Count
    ) | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
    
    $snapshots = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $count; $i++) {
        $rOps = Get-MetricValue $map["Composite Disk Read Operations/sec"][$i]
        $wOps = Get-MetricValue $map["Composite Disk Write Operations/sec"][$i]
        $rBps = Get-MetricValue $map["Composite Disk Read Bytes/sec"][$i]
        $wBps = Get-MetricValue $map["Composite Disk Write Bytes/sec"][$i]
        $timestamp = $map["Composite Disk Read Operations/sec"][$i].timeStamp
        
        # Skip samples with missing data
        if ($null -in ($rOps,$wOps,$rBps,$wBps)) { continue }
        
        $snapshots.Add([pscustomobject]@{
            TimeStamp = $timestamp
            IOPS = $rOps + $wOps
            MBps = ($rBps + $wBps) / 1MB
        })
    }
    
    # Calculate total transactions (for billing analysis on Standard HDD/SSD)
    $transactions = 0
    foreach ($metricName in @("Composite Disk Read Operations/sec","Composite Disk Write Operations/sec")) {
        if ($map.ContainsKey($metricName)) {
            foreach ($entry in $map[$metricName]) {
                $value = Get-MetricValue $entry
                if ($null -ne $value) {
                    # Convert ops/sec to total ops (multiply by sample duration in seconds)
                    $transactions += [double]$value * $secondsPerSample
                }
            }
        }
    }
    
    # DiskPaidBurstIOPS only supports the Average aggregation, so convert the per-second rate to total operations.
    $burstOps = 0
    if ($map.ContainsKey("DiskPaidBurstIOPS")) {
        foreach ($entry in $map["DiskPaidBurstIOPS"]) {
            $value = Get-MetricValue $entry
            if ($null -ne $value) {
                $burstOps += [double]$value * $secondsPerSample
            }
        }
    }
    
    return [pscustomobject]@{
        Samples = $snapshots
        Transactions = [math]::Round($transactions,2)
        BurstOperations = [math]::Round($burstOps,2)
    }
}

<#
.SYNOPSIS
    Detects continuous burst windows where performance exceeds 80% of limits for 60+ minutes.
.DESCRIPTION
    Analyzes time-series samples to find continuous windows where IOPS or throughput
    exceeds 80% of the maximum for at least 60 consecutive minutes. Duration is derived
    from the sample count and the metric interval, so a window is measured correctly
    regardless of whether samples are 1-minute, 5-minute, or 1-hour aggregates.
#>
function Get-BurstWindows {
    param(
        $Samples,
        [double]$MaxIOPS,
        [double]$MaxMBps,
        [int]$IntervalSeconds = 60,
        [int]$MinimumDurationMinutes = 60
    )
    
    $burstWindows = [System.Collections.Generic.List[object]]::new()
    if (-not $Samples -or $Samples.Count -eq 0) { return $burstWindows }
    
    $currentWindow = $null
    $threshold = 0.80  # 80% threshold

    # Emits the window under construction when it is long enough to matter.
    $closeWindow = {
        param($Window)
        if (-not $Window) { return }
        $duration = ($Window.SampleCount * $IntervalSeconds) / 60
        if ($duration -ge $MinimumDurationMinutes) {
            $burstWindows.Add([pscustomobject]@{
                StartTime       = $Window.StartTime
                EndTime         = $Window.EndTime
                DurationMinutes = [math]::Round($duration, 1)
                MaxIOPS         = [math]::Round($Window.MaxIOPS, 0)
                MaxMBps         = [math]::Round($Window.MaxMBps, 2)
            })
        }
    }
    
    foreach ($sample in $Samples) {
        $iopsRatio = if ($MaxIOPS -gt 0) { $sample.IOPS / $MaxIOPS } else { 0 }
        $mbpsRatio = if ($MaxMBps -gt 0) { $sample.MBps / $MaxMBps } else { 0 }
        $isHighUsage = ($iopsRatio -ge $threshold) -or ($mbpsRatio -ge $threshold)
        
        if ($isHighUsage) {
            if (-not $currentWindow) {
                $currentWindow = @{
                    StartTime   = $sample.TimeStamp
                    EndTime     = $sample.TimeStamp
                    MaxIOPS     = $sample.IOPS
                    MaxMBps     = $sample.MBps
                    SampleCount = 1
                }
            } else {
                $currentWindow.EndTime = $sample.TimeStamp
                $currentWindow.SampleCount++
                if ($sample.IOPS -gt $currentWindow.MaxIOPS) { $currentWindow.MaxIOPS = $sample.IOPS }
                if ($sample.MBps -gt $currentWindow.MaxMBps) { $currentWindow.MaxMBps = $sample.MBps }
            }
        } else {
            & $closeWindow $currentWindow
            $currentWindow = $null
        }
    }
    
    & $closeWindow $currentWindow
    
    return $burstWindows
}

<#
.SYNOPSIS
    Returns the mean latency for a disk from prefetched VM-level metrics.
.DESCRIPTION
    Azure exposes disk latency on the VM rather than the disk: OS disks through
    "OS Disk Latency" and data disks through "Data Disk Latency" split by LUN dimension.
    Both are fetched in batches up front and keyed here by disk resource ID.
#>
function Get-VmDiskLatency {
    param(
        [string]$VmResourceId,
        [string]$DiskResourceId
    )
    
    if ([string]::IsNullOrEmpty($VmResourceId) -or [string]::IsNullOrEmpty($DiskResourceId)) {
        return 0  # Disk not attached to a VM
    }
    
    $key = $DiskResourceId.ToLowerInvariant()
    if ($script:DiskLatencyCache.ContainsKey($key)) { return $script:DiskLatencyCache[$key] }
    return 0
}

<#
.SYNOPSIS
    Averages the non-null data points of an Azure Monitor timeseries.
#>
function Measure-TimeseriesAverage {
    param($Timeseries)
    
    $total = 0
    $count = 0
    foreach ($point in $Timeseries.data) {
        $value = Get-MetricValue $point
        if ($null -ne $value) {
            $total += [double]$value
            $count++
        }
    }
    if ($count -eq 0) { return 0 }
    return [math]::Round($total / $count, 2)
}

#endregion

#region Disk Discovery

# One Resource Graph query replaces a per-subscription "az disk list" sweep.
Write-Host "Discovering managed disks via Resource Graph..." -ForegroundColor Cyan
$diskQuery = @'
Resources
| where type =~ 'microsoft.compute/disks'
| project id, name, rg = resourceGroup, location, subscriptionId,
          sku = tostring(sku.name),
          diskSizeGB = toint(properties.diskSizeGB),
          managedBy = tostring(managedBy),
          osType = tostring(properties.osType),
          diskState = tostring(properties.diskState),
          burstingEnabled = properties.burstingEnabled,
          perfPlus = properties.supportedCapabilities.performancePlus,
          expandedIops = toint(properties.diskIOPSReadWrite),
          expandedMBps = toint(properties.diskMBpsReadWrite),
          timeCreated = tostring(properties.timeCreated)
| order by id asc
'@
$allDisks = Invoke-ResourceGraphQuery -Query $diskQuery -Subscriptions $SubscriptionId
Write-Host "  Found $($allDisks.Count) disk(s) across $($SubscriptionId.Count) subscription(s)" -ForegroundColor Gray

if (-not $allDisks.Count) {
    Write-Warning "No managed disks found for the provided subscription scope."
    return
}

if ($ExcludeAvd) {
    $beforeCount = $allDisks.Count
    $allDisks = [System.Collections.Generic.List[object]]@($allDisks | Where-Object { ([string]$_.name) -notmatch '(?i)AVD' })
    $excluded = $beforeCount - $allDisks.Count
    if ($excluded -gt 0) {
        Write-Host "  Excluded $excluded AVD disk(s) (-ExcludeAvd)" -ForegroundColor Gray
    }
}

Write-Host "`nTotal disks discovered: $($allDisks.Count)" -ForegroundColor Green
Write-Host "Analysis window (UTC): $($analysisWindowStart.ToString('yyyy-MM-dd HH:mm')) to $($analysisWindowEnd.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Green
Write-Host "Metric granularity: $metricInterval ($secondsPerSample second samples)" -ForegroundColor Gray
Write-Host ""

#endregion

#region Telemetry Prefetch

# Metrics for every disk and VM are collected up front in batches of up to 50 resources,
# which replaces two Azure Monitor calls per disk with a handful of requests overall.
$script:DiskMetricsCache = @{}
$script:DiskLatencyCache = @{}

Write-Host "Collecting disk metrics in batches..." -ForegroundColor Cyan
$diskGroups = $allDisks | Group-Object subscriptionId, location
foreach ($group in $diskGroups) {
    $sample = $group.Group[0]
    $ids = @($group.Group.id)
    Write-Host "  $($ids.Count) disk(s) in $($sample.location)" -ForegroundColor Gray
    $batch = Invoke-MetricsBatch -ResourceIds $ids -SubscriptionId $sample.subscriptionId -Region $sample.location `
        -MetricNamespace 'microsoft.compute/disks' -MetricNames $diskMetricNames
    foreach ($item in $batch.GetEnumerator()) { $script:DiskMetricsCache[$item.Key] = $item.Value }
}

# Latency lives on the VM, so map each attached disk to its parent VM and LUN first.
$attachedDisks = @($allDisks | Where-Object { $_.managedBy })
if ($attachedDisks.Count) {
    Write-Host "Collecting VM disk latency in batches..." -ForegroundColor Cyan
    $vmIds = @($attachedDisks.managedBy | Select-Object -Unique)
    
    $vmQuery = @'
Resources
| where type =~ 'microsoft.compute/virtualmachines'
| project id, location, subscriptionId,
          osDiskId = tolower(tostring(properties.storageProfile.osDisk.managedDisk.id)),
          dataDisks = properties.storageProfile.dataDisks
'@
    $vmRows = Invoke-ResourceGraphQuery -Query $vmQuery -Subscriptions $SubscriptionId
    $vms = @($vmRows | Where-Object { $vmIds -contains $_.id })
    
    # diskId -> @{ VmId; Lun } where a null LUN means the OS disk.
    $diskToVm = @{}
    foreach ($vm in $vms) {
        if ($vm.osDiskId) { $diskToVm[$vm.osDiskId] = @{ VmId = $vm.id; Lun = $null } }
        if ($vm.dataDisks) {
            foreach ($dataDisk in $vm.dataDisks) {
                $id = [string]$dataDisk.managedDisk.id
                if ($id) { $diskToVm[$id.ToLowerInvariant()] = @{ VmId = $vm.id; Lun = [string]$dataDisk.lun } }
            }
        }
    }
    
    foreach ($group in ($vms | Group-Object subscriptionId, location)) {
        $sample = $group.Group[0]
        $ids = @($group.Group.id)
        Write-Host "  $($ids.Count) VM(s) in $($sample.location)" -ForegroundColor Gray
        
        $osLatency = Invoke-MetricsBatch -ResourceIds $ids -SubscriptionId $sample.subscriptionId -Region $sample.location `
            -MetricNamespace 'microsoft.compute/virtualmachines' -MetricNames @('OS Disk Latency')
        $dataLatency = Invoke-MetricsBatch -ResourceIds $ids -SubscriptionId $sample.subscriptionId -Region $sample.location `
            -MetricNamespace 'microsoft.compute/virtualmachines' -MetricNames @('Data Disk Latency') -Filter "LUN eq '*'"
        
        foreach ($mapping in $diskToVm.GetEnumerator()) {
            $vmKey = $mapping.Value.VmId.ToLowerInvariant()
            $lun = $mapping.Value.Lun
            
            if ($null -eq $lun) {
                if (-not $osLatency.ContainsKey($vmKey)) { continue }
                $series = $osLatency[$vmKey]['OS Disk Latency']
                if ($series -and $series.Count -gt 0) {
                    $script:DiskLatencyCache[$mapping.Key] = Measure-TimeseriesAverage $series[0]
                }
                continue
            }
            
            if (-not $dataLatency.ContainsKey($vmKey)) { continue }
            foreach ($series in $dataLatency[$vmKey]['Data Disk Latency']) {
                $seriesLun = ($series.metadatavalues | Where-Object { $_.name.value -eq 'LUN' } | Select-Object -First 1).value
                if ([string]$seriesLun -eq $lun) {
                    $script:DiskLatencyCache[$mapping.Key] = Measure-TimeseriesAverage $series
                    break
                }
            }
        }
    }
}
Write-Host ""

#endregion

#region Performance Analysis

# Analyze each disk's performance against its limits
Write-Host "Analyzing disk performance..." -ForegroundColor Cyan
if ($DaysToInspect -lt 30) {
    Write-Host "NOTE: Analyzing $DaysToInspect days of data. All cost estimates are projected to 30-day monthly values." -ForegroundColor Yellow
}
Write-Host ""
$diskNumber = 0
$results = foreach ($disk in $allDisks) {
    $diskNumber++
    Write-Host "  [$diskNumber/$($allDisks.Count)] Processing: $($disk.name) ($($disk.sku), $($disk.diskSizeGB) GiB)" -ForegroundColor Gray
    
    $tier = $disk.sku
    $actualSize = [int]$disk.diskSizeGB
    $isPerfPlus = [bool]$disk.perfPlus
    $isOsDisk = ($null -ne $disk.osType -and $disk.osType -ne "")
    $redundancy = if ($tier -match '_ZRS$') { "ZRS" } else { "LRS" }
    $diskRegion = if ($PricingRegion) { $PricingRegion } else { [string]$disk.location }
    
    # A disk with no owner still bills for its full provisioned capacity.
    $diskState = if ($disk.PSObject.Properties['diskState'] -and $disk.diskState) { [string]$disk.diskState } else { "Unknown" }
    $isAttached = [bool]$disk.managedBy
    $ageDays = 0
    if ($disk.PSObject.Properties['timeCreated'] -and $disk.timeCreated) {
        $created = [datetime]::MinValue
        if ([datetime]::TryParse([string]$disk.timeCreated, [ref]$created)) {
            $ageDays = [math]::Round(((Get-Date).ToUniversalTime() - $created.ToUniversalTime()).TotalDays, 0)
        }
    }
    
    # Resolve pricing for this disk's own region; the catalog is built once per region.
    $catalog = $null
    if (-not $NoLivePricing -and $diskRegion) { $catalog = Get-DiskPriceCatalog -Region $diskRegion }
    $pricingSource = if ($catalog) { "RetailAPI" } else { "SpecFile" }
    $pricingRegionUsed = if ($catalog) { $catalog.Region } else { [string]$rates.Region }
    
    # Determine the billing tier and whether it's a custom size
    $billingInfo = $null
    $size = $actualSize
    $isCustomSize = $false
    $billingTierName = "N/A"
    
    if ($tier -match '^(Standard_|StandardSSD_|Premium_)') {
        $billingInfo = Get-BillingTier -DiskType $tier -ActualSizeGB $actualSize
        $size = $billingInfo.BilledSizeGB
        $isCustomSize = $billingInfo.IsCustomSize
        $billingTierName = $billingInfo.TierName
        
        if ($isCustomSize) {
            Write-Host "      Custom size detected: $actualSize GiB billed as $billingTierName ($($size) GiB)" -ForegroundColor Cyan
        }
    }
    elseif ($tier -match '^(PremiumV2_|UltraSSD_)') {
        # Billed per provisioned GiB in 1 GiB steps, so there is no tier to round up to.
        $billingTierName = if ($tier -like 'UltraSSD*') { "Ultra" } else { "PremiumV2" }
    }
    
    # Get the specification for this disk tier and size
    $specForTier = switch -Regex ($tier) {
        '^Standard_(LRS|ZRS)$' { Get-SkuSpec $hddSpecs $size; break }
        '^StandardSSD_' { Get-SkuSpec $ssdSpecs $size; break }
        '^Premium_' { Get-SkuSpec $premiumSpecs $size; break }
        '^(PremiumV2_|UltraSSD_)' { Get-ProvisionedSpec -Disk $disk -Tier $tier; break }
        default { $null }
    }
    
    # Fetch performance telemetry
    $telemetry = Get-DiskTelemetry -ResourceId $disk.id
    
    # Provide context if no metrics were retrieved
    if ($telemetry.Samples.Count -eq 0) {
        if (-not $disk.managedBy) {
            Write-Host "      Disk is unattached (no VM) - metrics not available" -ForegroundColor DarkGray
        } else {
            Write-Host "      No metric samples retrieved - VM may have been stopped during analysis period" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "      Retrieved $($telemetry.Samples.Count) metric samples" -ForegroundColor DarkGray
    }
    
    # Latency comes from the VM-level metrics collected during prefetch.
    $avgLatencyMs = 0
    if ($disk.managedBy) {
        $avgLatencyMs = Get-VmDiskLatency -VmResourceId $disk.managedBy -DiskResourceId $disk.id
        if ($avgLatencyMs -gt 0) {
            Write-Host "      Average disk latency: $avgLatencyMs ms" -ForegroundColor DarkGray
        }
    }
    
    # Calculate transaction costs (for HDD and SSD)
    # Project to monthly cost for accurate comparisons
    $daysInMonth = $script:HoursPerMonth / 24
    $monthlyTransactions = ($telemetry.Transactions / $DaysToInspect) * $daysInMonth
    $hddTransactionRate = Get-CatalogRate $catalog "txn|HDD|LRS" $fallbackHddTransactionPer10k
    $ssdTransactionRate = Get-CatalogRate $catalog "txn|SSD|$redundancy" $fallbackSsdTransactionPer10k
    $hddTransactionCost = ($monthlyTransactions / 10000) * $hddTransactionRate
    $ssdTransactionCost = ($monthlyTransactions / 10000) * $ssdTransactionRate
    
    # Display transaction cost analysis if analyzing less than 30 days
    if ($DaysToInspect -lt 30 -and $telemetry.Transactions -gt 0) {
        Write-Host "      Observed $([math]::Round($telemetry.Transactions, 0).ToString("N0")) transactions over $DaysToInspect days (Projected monthly: $([math]::Round($monthlyTransactions, 0).ToString("N0")))" -ForegroundColor DarkGray
        if ($tier -eq "Standard_LRS" -and $hddTransactionCost -gt 1) {
            Write-Host "      Projected monthly HDD transaction cost: `$$($hddTransactionCost.ToString('F2'))" -ForegroundColor $(if ($hddTransactionCost -gt 5) { "Yellow" } else { "DarkGray" })
        }
        if ($tier -like "StandardSSD_*" -and $ssdTransactionCost -gt 0.1) {
            Write-Host "      Projected monthly SSD transaction cost: `$$($ssdTransactionCost.ToString('F2'))" -ForegroundColor DarkGray
        }
    }
    
    # Each disk type rounds to its own tier ladder, so resolve the billed tier per family rather
    # than reusing the current disk's billed size (Standard HDD has no tier below 32 GiB).
    $hddBilling = Get-BillingTier -DiskType "Standard_LRS" -ActualSizeGB $actualSize
    $ssdBilling = Get-BillingTier -DiskType "StandardSSD_LRS" -ActualSizeGB $actualSize
    $premiumBilling = Get-BillingTier -DiskType "Premium_LRS" -ActualSizeGB $actualSize
    
    $hddSpec = Get-SkuSpec $hddSpecs $hddBilling.BilledSizeGB
    $ssdSpec = Get-SkuSpec $ssdSpecs $ssdBilling.BilledSizeGB
    $premiumSpec = Get-SkuSpec $premiumSpecs $premiumBilling.BilledSizeGB
    
    # Premium SSD v2 is provisioned to observed demand rather than to a fixed tier.
    # P99 is used for right-sizing so a single outlier minute does not inflate the recommendation.
    $observedPeakIops = 0
    $observedPeakMBps = 0
    $observedP99Iops = 0
    $observedP99MBps = 0
    $observedP95Iops = 0
    $observedP95MBps = 0
    if ($telemetry.Samples.Count -gt 0) {
        $observedPeakIops = ($telemetry.Samples | Measure-Object -Property IOPS -Maximum).Maximum
        $observedPeakMBps = ($telemetry.Samples | Measure-Object -Property MBps -Maximum).Maximum
        $observedP99Iops = Get-SamplePercentile $telemetry.Samples "IOPS" 99
        $observedP99MBps = Get-SamplePercentile $telemetry.Samples "MBps" 99
        $observedP95Iops = Get-SamplePercentile $telemetry.Samples "IOPS" 95
        $observedP95MBps = Get-SamplePercentile $telemetry.Samples "MBps" 95
    }
    $effectiveIopsRequirement = [math]::Max(3000, $observedPeakIops)
    $effectiveMBpsRequirement = [math]::Max(125, $observedPeakMBps)
    
    # Capacity costs, preferring live retail pricing for the disk's own region and redundancy.
    # Standard HDD has no ZRS offering, so it is always priced as LRS.
    $hddStorageCost = Get-CapacityMonthlyCost -Catalog $catalog -TierName $hddBilling.TierName -Redundancy "LRS" -SpecEntry $hddSpec -SizeGiB $hddBilling.BilledSizeGB -FallbackPerGiB 0.04
    $ssdStorageCost = Get-CapacityMonthlyCost -Catalog $catalog -TierName $ssdBilling.TierName -Redundancy $redundancy -SpecEntry $ssdSpec -SizeGiB $ssdBilling.BilledSizeGB -FallbackPerGiB 0.075
    $premiumStorageCost = Get-CapacityMonthlyCost -Catalog $catalog -TierName $premiumBilling.TierName -Redundancy $redundancy -SpecEntry $premiumSpec -SizeGiB $premiumBilling.BilledSizeGB -FallbackPerGiB 0.135
    
    # Premium SSD v2 bills provisioned capacity, IOPS and throughput separately, with
    # 3,000 IOPS and 125 MB/s included at no charge.
    $v2CapacityRate = Get-CatalogRate $catalog "provcap|PREMV2|LRS" ([double]$premiumV2Specs.HourlyRatePerGiB)
    $v2IopsRate = Get-CatalogRate $catalog "proviops|PREMV2|LRS" 0
    $v2ThroughputRate = Get-CatalogRate $catalog "provmbps|PREMV2|LRS" 0
    $v2BillableIops = [math]::Max(0.0, [math]::Ceiling($effectiveIopsRequirement) - 3000)
    $v2BillableMBps = [math]::Max(0.0, [math]::Ceiling($effectiveMBpsRequirement) - 125)
    $premiumV2StorageCost = ($actualSize * $v2CapacityRate * $script:HoursPerMonth) +
                            ($v2BillableIops * $v2IopsRate * $script:HoursPerMonth) +
                            ($v2BillableMBps * $v2ThroughputRate * $script:HoursPerMonth)
    
    # For a disk that already is v2 or Ultra, cost follows what is actually provisioned on it.
    $provisionedCost = 0
    $provisionedIops = 0
    $provisionedMBps = 0
    if ($tier -match '^(PremiumV2_|UltraSSD_)') {
        $family = if ($tier -like 'UltraSSD*') { 'ULTRA' } else { 'PREMV2' }
        $provisionedIops = $specForTier.MaxIOPS
        $provisionedMBps = $specForTier.MaxMBps
        $capRate = Get-CatalogRate $catalog "provcap|$family|LRS" ([double]$premiumV2Specs.HourlyRatePerGiB)
        $iopsRate = Get-CatalogRate $catalog "proviops|$family|LRS" 0
        $mbpsRate = Get-CatalogRate $catalog "provmbps|$family|LRS" 0
        $billableIops = [math]::Max(0.0, $provisionedIops - $specForTier.BaselineIOPS)
        $billableMBps = [math]::Max(0.0, $provisionedMBps - $specForTier.BaselineMBps)
        $provisionedCost = ($actualSize * $capRate * $script:HoursPerMonth) +
                           ($billableIops * $iopsRate * $script:HoursPerMonth) +
                           ($billableMBps * $mbpsRate * $script:HoursPerMonth)
    }
    
    # Paid on-demand bursting exists only on Premium SSD P30 and above; Standard SSD bursting is free credit-based.
    # Billing is a flat monthly enablement fee plus a per-10K charge on burst transactions.
    $burstCost = 0
    $burstEligible = ($tier -like "Premium_*") -and ($premiumBilling.BilledSizeGB -ge 1024)
    if ($telemetry.BurstOperations -gt 0 -and $burstEligible) {
        $burstEnablementRate = Get-CatalogRate $catalog "burstenable|$redundancy" $fallbackBurstEnablementMonthly
        $burstTransactionRate = Get-CatalogRate $catalog "bursttxn|$redundancy" $fallbackBurstTransactionPer10k
        $monthlyBurstOps = ($telemetry.BurstOperations / $DaysToInspect) * $daysInMonth
        $burstCost = $burstEnablementRate + (($monthlyBurstOps / 10000) * $burstTransactionRate)
        Write-Host "      Paid burst operations detected: $([math]::Round($monthlyBurstOps, 0).ToString('N0'))/month (Est. monthly cost: `$$($burstCost.ToString('F2')) = `$$($burstEnablementRate.ToString('F2')) enablement + transactions)" -ForegroundColor Yellow
    }
    
    # Total cost calculation based on disk type
    $totalEstimatedMonthlyCost = 0
    $hddTotalCost = $hddStorageCost + $hddTransactionCost  # HDD = storage + transactions
    $ssdTotalCost = $ssdStorageCost + $ssdTransactionCost  # Standard SSD = storage + transactions (bursting is free)
    $premiumTotalCost = $premiumStorageCost + $burstCost   # Premium SSD v1 = storage + paid bursting when enabled
    $premiumV2TotalCost = $premiumV2StorageCost  # Premium SSD v2 = provisioned capacity + IOPS + throughput
    
    # One-year disk reservations are offered for Premium SSD P30 through P80.
    $reservationMonthly = 0
    $reservationSavings = 0
    if ($catalog -and $catalog.Prices.ContainsKey("rsv1y|$($premiumBilling.TierName)|$redundancy")) {
        $reservationMonthly = $catalog.Prices["rsv1y|$($premiumBilling.TierName)|$redundancy"]
        if ($premiumStorageCost -gt 0) { $reservationSavings = $premiumStorageCost - $reservationMonthly }
    }
    
    if ($tier -eq "Standard_LRS") {
        # Standard HDD: storage + transaction costs
        $totalEstimatedMonthlyCost = $hddTotalCost
    }
    elseif ($tier -like "StandardSSD_*") {
        # Standard SSD: storage + transaction costs
        $totalEstimatedMonthlyCost = $ssdTotalCost
    }
    elseif ($tier -like "PremiumV2_*" -or $tier -like "UltraSSD_*") {
        $totalEstimatedMonthlyCost = $provisionedCost
    }
    else {
        # Premium SSD: storage cost plus paid bursting when enabled (no transaction fees)
        $totalEstimatedMonthlyCost = $premiumTotalCost
    }
    
    # Calculate cost savings/difference comparisons
    $costSavingsHDDtoSSD = $hddTotalCost - $ssdTotalCost  # Positive = HDD more expensive, Negative = SSD more expensive
    $costDiffSSDtoPremium = $premiumTotalCost - $ssdTotalCost  # Positive = Premium v1 more expensive, Negative = Premium v1 cheaper
    $costDiffSSDtoPremiumV2 = $premiumV2TotalCost - $ssdTotalCost  # Positive = Premium v2 more expensive, Negative = Premium v2 cheaper
    
    # Determine best Premium option (v1 vs v2). Premium SSD v2 and Ultra cannot be used as
    # OS disks, so an OS disk may only ever be compared against Premium SSD v1.
    $v2Eligible = -not $isOsDisk
    $bestPremiumCost = if ($v2Eligible) { [Math]::Min($premiumTotalCost, $premiumV2TotalCost) } else { $premiumTotalCost }
    $bestPremiumType = if ($v2Eligible -and $premiumV2TotalCost -lt $premiumTotalCost) { "Premium SSD v2" } else { "Premium SSD" }
    $costDiffSSDtoBestPremium = $bestPremiumCost - $ssdTotalCost
    
    # "Tier not handled" means the disk SKU or size is not defined in the spec file (disk-specs.json).
    # This currently covers ZRS variants, Premium SSD v2, and Ultra Disk.
    $tierHandled = [bool]$specForTier
    
    # Defaults used when the tier has no spec, so every disk emits an identically shaped record.
    $effectiveIops = 0
    $effectiveMb = 0
    $percentIOPS = 0
    $percentMBps = 0
    $burstWindows = @()
    $hasBurstWindows = $false
    $burstWindowsSummary = "None"
    $txnNote = "n/a"
    $outcome = New-DiskDecision -Decision "Tier not handled" -Severity 'Unknown'
    
    # Downgrade and right-size advice must never be derived from an absence of data: with no
    # samples every percentile is 0, which otherwise makes any disk look safe to shrink.
    $hasTelemetry = $telemetry.Samples.Count -gt 0
    
    # Effective limits are useful on every handled disk, including unattached ones.
    if ($tierHandled) {
        $effectiveIops = if ($isPerfPlus -and $disk.expandedIops) { [double]$disk.expandedIops } else { $specForTier.MaxIOPS }
        $effectiveMb = if ($isPerfPlus -and $disk.expandedMBps) { [double]$disk.expandedMBps } else { $specForTier.MaxMBps }
    }
    
    if (-not $tierHandled) {
        $outcome = New-DiskDecision -Decision "Tier not handled" -Severity 'Unknown' `
            -Detail "SKU '$tier' at $size GiB has no entry in $SpecFile"
    }
    elseif (-not $isAttached) {
        # An orphaned disk bills in full for capacity it is not serving, so the size of the
        # win is its entire monthly cost and no telemetry is needed to justify it.
        $ageNote = if ($ageDays -gt 0) { " created $ageDays day(s) ago," } else { "" }
        $outcome = New-DiskDecision -Decision "Delete or snapshot - unattached" -Severity 'Action' -Savings $totalEstimatedMonthlyCost `
            -Detail "unattached disk,$ageNote reclaims `$$($totalEstimatedMonthlyCost.ToString('F2'))/month"
    }
    elseif (-not $hasTelemetry) {
        # Capacity bills whether or not the VM is running, but with no samples there is no
        # basis for a sizing call, so say so rather than implying the disk was assessed.
        $outcome = if ($diskState -eq "Reserved") {
            New-DiskDecision -Decision "Review - VM deallocated" -Severity 'Watch' `
                -Detail "disk still bills `$$($totalEstimatedMonthlyCost.ToString('F2'))/month while the VM is stopped"
        } else {
            New-DiskDecision -Decision "Insufficient telemetry" -Severity 'Unknown' `
                -Detail "no metric samples in the analysis window"
        }
    }
    else {
        # Calculate what percentage of samples exceeded the limits
        $percentIOPS = Get-PercentAbove $telemetry.Samples "IOPS" $effectiveIops
        $percentMBps = Get-PercentAbove $telemetry.Samples "MBps" $effectiveMb
    
        # Detect burst windows (60+ minute continuous periods at/above 80% capacity)
        $burstWindows = @(Get-BurstWindows -Samples $telemetry.Samples -MaxIOPS $effectiveIops -MaxMBps $effectiveMb -IntervalSeconds $secondsPerSample)
        $hasBurstWindows = $burstWindows.Count -gt 0
        $burstWindowsSummary = if ($hasBurstWindows) {
            ($burstWindows | ForEach-Object {
                "$($_.StartTime) to $($_.EndTime) ($($_.DurationMinutes)min, IOPS:$($_.MaxIOPS), MBps:$($_.MaxMBps))"
            }) -join "; "
        } else {
            "None"
        }
    
        # Determine recommendation based on tier and performance
        $outcome = New-DiskDecision -Decision "Stay" -Severity 'Optimal'
        $txnNote = "n/a"
        $exceedsLimits = ($percentIOPS -gt $ExceedanceThreshold) -or ($percentMBps -gt $ExceedanceThreshold)
        $exceedanceNote = "IOPS $([math]::Round($percentIOPS*100,1))% above, MBps $([math]::Round($percentMBps*100,1))% above"
        $burstNote = if ($hasBurstWindows) { " plus $($burstWindows.Count) sustained burst window(s)" } else { "" }
    
        switch -Regex ($tier) {
            '^Standard_(LRS|ZRS)$' {
                # Standard HDD bills per transaction, so a busy HDD can cost more than the SSD that replaces it.
                if ($exceedsLimits) {
                    $txnSaving = if ($costSavingsHDDtoSSD -gt 0) { $costSavingsHDDtoSSD } else { 0 }
                    $outcome = New-DiskDecision -Decision "Upgrade to Standard SSD" -Severity 'Opportunity' -Savings $txnSaving `
                        -Detail "$exceedanceNote$burstNote$(if ($txnSaving) { " - also saves `$$($txnSaving.ToString('F2'))/month on transactions" })"
                }
                elseif ($hasBurstWindows) {
                    $outcome = New-DiskDecision -Decision "Review burst patterns" -Severity 'Watch' `
                        -Detail "$($burstWindows.Count) sustained burst(s) at 80%+ of the tier limit"
                }
                elseif ($costSavingsHDDtoSSD -gt 5) {
                    $outcome = New-DiskDecision -Decision "Upgrade to Standard SSD (high transaction costs)" -Severity 'Opportunity' -Savings $costSavingsHDDtoSSD `
                        -Detail "transaction charges make SSD `$$($costSavingsHDDtoSSD.ToString('F2'))/month cheaper overall"
                }
                $txnNote = if ($telemetry.Transactions -gt 0) { "Yes (Standard HDD bills per transaction)" } else { "No" }
            }
            '^StandardSSD_' {
                # Evaluated in savings order: a cheaper Premium tier, then latency, then throughput,
                # then the downgrade to HDD. Latency ranks above raw throughput because a disk can sit
                # inside its IOPS limit and still be the thing slowing the workload down.
                $hddBaseline = Get-SkuSpec $hddSpecs $hddBilling.BilledSizeGB
                $hddSaving = -$costSavingsHDDtoSSD   # positive when HDD is the cheaper home
                $premiumSaving = -$costDiffSSDtoBestPremium
            
                $fitsOnHdd = $false
                if ($hddBaseline -and $hasTelemetry) {
                    $fitsOnHdd = (
                        (Get-PercentAbove $telemetry.Samples "IOPS" $hddBaseline.MaxIOPS) -le $ExceedanceThreshold -and
                        (Get-PercentAbove $telemetry.Samples "MBps" $hddBaseline.MaxMBps) -le $ExceedanceThreshold
                    )
                }
                # Above 5 ms the workload is latency-sensitive enough that HDD would be felt.
                $latencyAllowsHdd = $avgLatencyMs -le 5
                $canDowngrade = $fitsOnHdd -and -not $isOsDisk -and $latencyAllowsHdd -and $hddSaving -ge 5
            
                if ($premiumSaving -gt 0) {
                    $outcome = New-DiskDecision -Decision "Upgrade to $bestPremiumType (cost savings)" -Severity 'Opportunity' -Savings $premiumSaving `
                        -Detail "`$$($ssdTotalCost.ToString('F2'))/month on Standard SSD versus `$$($bestPremiumCost.ToString('F2')) on $bestPremiumType"
                }
                elseif ($avgLatencyMs -gt 10) {
                    $outcome = New-DiskDecision -Decision "Upgrade to $bestPremiumType (latency)" -Severity 'Opportunity' `
                        -Detail "latency $avgLatencyMs ms indicates the disk is the bottleneck; $bestPremiumType $(Format-CostDelta $costDiffSSDtoBestPremium)"
                }
                elseif ($exceedsLimits) {
                    $outcome = New-DiskDecision -Decision "Upgrade to $bestPremiumType (performance)" -Severity 'Opportunity' `
                        -Detail "$exceedanceNote$burstNote; $bestPremiumType $(Format-CostDelta $costDiffSSDtoBestPremium)"
                }
                elseif ($hasBurstWindows) {
                    $outcome = New-DiskDecision -Decision "Review burst patterns" -Severity 'Watch' `
                        -Detail "$($burstWindows.Count) sustained burst(s) at 80%+ of the tier limit"
                }
                elseif ($canDowngrade) {
                    $outcome = New-DiskDecision -Decision "Downgrade to Standard HDD" -Severity 'Opportunity' -Savings $hddSaving `
                        -Detail "underutilized data disk; HDD saves `$$($hddSaving.ToString('F2'))/month including transactions"
                }
                elseif ($fitsOnHdd -and $isOsDisk) {
                    $outcome = New-DiskDecision -Decision "Stay (OS disk)" -Severity 'Optimal' -Detail "OS disks should not run on Standard HDD"
                }
                elseif ($fitsOnHdd -and -not $latencyAllowsHdd) {
                    $outcome = New-DiskDecision -Decision "Stay (latency-sensitive)" -Severity 'Optimal' -Detail "latency $avgLatencyMs ms needs SSD response times"
                }
                elseif ($fitsOnHdd -and $hddSaving -gt 0) {
                    $outcome = New-DiskDecision -Decision "Stay (minimal savings)" -Severity 'Optimal' -Detail "HDD would save only `$$($hddSaving.ToString('F2'))/month"
                }
                elseif ($fitsOnHdd) {
                    $outcome = New-DiskDecision -Decision "Stay" -Severity 'Optimal' -Detail "underutilized, but transaction charges make HDD `$$($costSavingsHDDtoSSD.ToString('F2'))/month more expensive"
                }
            
                $txnNote = if ($ssdTransactionCost -gt 0) { "Yes (Standard SSD: `$$($ssdTransactionCost.ToString('F2'))/month)" } else { "No" }
            }
            '^Premium_' {
                # Premium SSD: look for a cheaper home on Standard SSD, a cheaper home on v2,
                # or a reservation when the disk is staying put.
                $ssdBaseline = Get-SkuSpec $ssdSpecs $ssdBilling.BilledSizeGB
                $fitsOnStandardSsd = $false
                if ($ssdBaseline -and $hasTelemetry) {
                    $fitsOnStandardSsd = (
                        (Get-PercentAbove $telemetry.Samples "IOPS" $ssdBaseline.MaxIOPS) -le $ExceedanceThreshold -and
                        (Get-PercentAbove $telemetry.Samples "MBps" $ssdBaseline.MaxMBps) -le $ExceedanceThreshold
                    )
                }
                # Premium's value is consistent low latency, so only step down when latency is comfortable.
                $latencyAllowsStandard = $avgLatencyMs -le 10
                $savingsToStandardSsd = $premiumTotalCost - $ssdTotalCost
                $savingsToV2 = $premiumTotalCost - $premiumV2TotalCost

                if ($exceedsLimits) {
                    $outcome = New-DiskDecision -Decision "Review - exceeding Premium SSD limits" -Severity 'Watch' `
                        -Detail "$exceedanceNote; consider a larger tier or Premium SSD v2"
                }
                elseif ($fitsOnStandardSsd -and $latencyAllowsStandard -and $savingsToStandardSsd -ge 5 -and -not $isOsDisk) {
                    $outcome = New-DiskDecision -Decision "Downgrade to Standard SSD" -Severity 'Opportunity' -Savings $savingsToStandardSsd `
                        -Detail "data disk fits within $($ssdBaseline.Sku) limits; saves `$$($savingsToStandardSsd.ToString('F2'))/month"
                }
                elseif ($savingsToV2 -ge 5 -and $v2Eligible) {
                    $outcome = New-DiskDecision -Decision "Migrate to Premium SSD v2 (cost savings)" -Severity 'Opportunity' -Savings $savingsToV2 `
                        -Detail "saves `$$($savingsToV2.ToString('F2'))/month at the observed performance level"
                }
                elseif ($reservationSavings -ge 5) {
                    $outcome = New-DiskDecision -Decision "Purchase 1-year disk reservation" -Severity 'Opportunity' -Savings $reservationSavings `
                        -Detail "$($premiumBilling.TierName) drops from `$$($premiumStorageCost.ToString('F2')) to `$$($reservationMonthly.ToString('F2'))/month"
                }
                elseif ($savingsToV2 -ge 5 -and $isOsDisk) {
                    $outcome = New-DiskDecision -Decision "Stay (OS disk cannot use Premium SSD v2)" -Severity 'Optimal' `
                        -Detail "v2 would be `$$($savingsToV2.ToString('F2'))/month cheaper but is not supported for OS disks"
                }
                elseif ($fitsOnStandardSsd -and $isOsDisk) {
                    $outcome = New-DiskDecision -Decision "Stay (OS disk)" -Severity 'Optimal' -Detail "underutilized, but OS disks benefit from Premium latency"
                }
                elseif ($fitsOnStandardSsd -and -not $latencyAllowsStandard) {
                    $outcome = New-DiskDecision -Decision "Stay (latency-sensitive)" -Severity 'Optimal' -Detail "latency $avgLatencyMs ms indicates the workload needs Premium"
                }
                $txnNote = "No (Premium SSD has no transaction charge)"
            }
            '^(PremiumV2_|UltraSSD_)' {
                # These bill per provisioned IOPS and MB/s above the free baseline, so the
                # saving comes from lowering the provisioned numbers rather than changing SKU.
                $family = if ($tier -like 'UltraSSD*') { 'ULTRA' } else { 'PREMV2' }
                $iopsRate = Get-CatalogRate $catalog "proviops|$family|LRS" 0
                $mbpsRate = Get-CatalogRate $catalog "provmbps|$family|LRS" 0

                # Size to P99 plus 30% headroom, never below the free baseline.
                $targetIops = [math]::Max($specForTier.BaselineIOPS, [math]::Ceiling($observedP99Iops * 1.3))
                $targetMBps = [math]::Max($specForTier.BaselineMBps, [math]::Ceiling($observedP99MBps * 1.3))
                $iopsSaving = [math]::Max(0.0, $provisionedIops - $targetIops) * $iopsRate * $script:HoursPerMonth
                $mbpsSaving = [math]::Max(0.0, $provisionedMBps - $targetMBps) * $mbpsRate * $script:HoursPerMonth
                $rightSizeSaving = $iopsSaving + $mbpsSaving

                # A fixed Premium tier is only an option if it can carry the observed load.
                $fitsOnPremiumTier = $false
                if ($premiumSpec) {
                    $fitsOnPremiumTier = ($observedP99Iops -le $premiumSpec.MaxIOPS) -and ($observedP99MBps -le $premiumSpec.MaxMBps)
                }

                if ($exceedsLimits) {
                    $outcome = New-DiskDecision -Decision "Increase provisioned performance" -Severity 'Watch' `
                        -Detail "usage exceeds the provisioned $provisionedIops IOPS / $provisionedMBps MB/s ($exceedanceNote)"
                }
                elseif ($rightSizeSaving -ge 5 -and $hasTelemetry) {
                    $outcome = New-DiskDecision -Decision "Reduce provisioned performance" -Severity 'Opportunity' -Savings $rightSizeSaving `
                        -Detail "provisioned $provisionedIops IOPS / $provisionedMBps MB/s but P99 is $([math]::Round($observedP99Iops,0)) IOPS / $([math]::Round($observedP99MBps,1)) MB/s; target $targetIops / $targetMBps saves `$$($rightSizeSaving.ToString('F2'))/month"
                }
                elseif ($premiumTotalCost -gt 0 -and ($provisionedCost - $premiumTotalCost) -ge 5 -and $fitsOnPremiumTier) {
                    $tierSaving = $provisionedCost - $premiumTotalCost
                    $outcome = New-DiskDecision -Decision "Migrate to Premium SSD $($premiumBilling.TierName)" -Severity 'Opportunity' -Savings $tierSaving `
                        -Detail "fixed tier costs `$$($premiumTotalCost.ToString('F2'))/month versus `$$($provisionedCost.ToString('F2'))"
                }
                else {
                    $outcome = New-DiskDecision -Decision "Stay" -Severity 'Optimal' -Detail "provisioned $provisionedIops IOPS / $provisionedMBps MB/s"
                }
                $txnNote = "No (billed on provisioned performance)"
            }
            default { }
        }
    }
    
    # Single render point for every branch above.
    $detailText = if ($outcome.Detail) { " - $($outcome.Detail)" } else { "" }
    Write-Host "      $($outcome.Decision)$detailText" -ForegroundColor $script:DecisionSeverityColor[$outcome.Severity]
    
    # Determine metrics availability status
    $metricsStatus = if ($telemetry.Samples.Count -eq 0) {
        if (-not $disk.managedBy) { "Unattached" } else { "No Metrics" }
    } else { "OK" }
    
    # Output analysis result for this disk
    [pscustomobject]@{
        Disk=$disk.name
        ResourceGroup=$disk.rg
        Subscription=$disk.subscriptionId
        Location=[string]$disk.location
        SKU=$tier
        Redundancy=$redundancy
        DiskState=$diskState
        AgeDays=$ageDays
        Attached=if ($isAttached) { "Yes" } else { "No" }
        ActualSizeGiB=$actualSize
        BilledSizeGiB=$size
        BillingTier=$billingTierName
        CustomSize=if ($isCustomSize) { "Yes" } else { "No" }
        Samples=$telemetry.Samples.Count
        MetricsStatus=$metricsStatus
        PercentIOPSAbove=[math]::Round($percentIOPS*100,2)
        PercentMBpsAbove=[math]::Round($percentMBps*100,2)
        PeakIOPS=[math]::Round($observedPeakIops,0)
        P99IOPS=[math]::Round($observedP99Iops,0)
        P95IOPS=[math]::Round($observedP95Iops,0)
        PeakMBps=[math]::Round($observedPeakMBps,2)
        P99MBps=[math]::Round($observedP99MBps,2)
        P95MBps=[math]::Round($observedP95MBps,2)
        ProvisionedIOPS=[math]::Round($provisionedIops,0)
        ProvisionedMBps=[math]::Round($provisionedMBps,0)
        Transactions=$telemetry.Transactions
        BurstOperations=$telemetry.BurstOperations
        BurstWindowCount=$burstWindows.Count
        BurstWindows=$burstWindowsSummary
        AvgLatencyMs=$avgLatencyMs
        EstMonthlyHDDTransactionCost=[math]::Round($hddTransactionCost,2)
        EstMonthlySSDTransactionCost=[math]::Round($ssdTransactionCost,2)
        EstMonthlyBurstCost=[math]::Round($burstCost,2)
        EstMonthlyHDDTotalCost=[math]::Round($hddTotalCost,2)
        EstMonthlySSDTotalCost=[math]::Round($ssdTotalCost,2)
        EstMonthlyPremiumTotalCost=[math]::Round($premiumTotalCost,2)
        EstMonthlyPremiumV2TotalCost=[math]::Round($premiumV2TotalCost,2)
        BestPremiumOption=$bestPremiumType
        BestPremiumCost=[math]::Round($bestPremiumCost,2)
        CostSavingsHDDtoSSD=[math]::Round($costSavingsHDDtoSSD,2)
        CostDiffSSDtoPremium=[math]::Round($costDiffSSDtoPremium,2)
        CostDiffSSDtoPremiumV2=[math]::Round($costDiffSSDtoPremiumV2,2)
        CostDiffSSDtoBestPremium=[math]::Round($costDiffSSDtoBestPremium,2)
        EstMonthlyTotalCost=[math]::Round($totalEstimatedMonthlyCost,2)
        ReservationMonthlyCost=[math]::Round($reservationMonthly,2)
        ReservationMonthlySavings=[math]::Round($reservationSavings,2)
        PricingSource=$pricingSource
        PricingRegion=$pricingRegionUsed
        CurrencyCode=$Currency
        ChargedForTransactions=$txnNote
        PerformancePlus=if ($isPerfPlus) { "Yes" } else { "No" }
        EffectiveMaxIOPS=[math]::Round($effectiveIops,0)
        EffectiveMaxMBps=[math]::Round($effectiveMb,0)
        Decision=$outcome.Decision
        DecisionDetail=$outcome.Detail
        Severity=$outcome.Severity
        PotentialMonthlySavings=$outcome.Savings
    }
}

Write-Host "`nAnalysis complete!" -ForegroundColor Green

#endregion

# Display results
Write-Host "`n=== Analysis Results ===" -ForegroundColor Cyan
$results = @($results | Sort-Object Subscription, ResourceGroup, Disk)
$results | Format-Table -AutoSize

# Export to JSON and CSV with timestamp
Write-Host "Exporting results..." -ForegroundColor Cyan
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$jsonPath = Join-Path $scriptDir "$OutputPrefix-$timestamp.json"
$csvPath = Join-Path $scriptDir "$OutputPrefix-$timestamp.csv"

$results | ConvertTo-Json -Depth 4 | Set-Content $jsonPath
$results | Export-Csv $csvPath -NoTypeInformation

Write-Host "  JSON: $jsonPath" -ForegroundColor Green
Write-Host "  CSV:  $csvPath" -ForegroundColor Green
Write-Host ""
Write-Host "To generate an HTML report, run:" -ForegroundColor Cyan
Write-Host "  .\New-DiskAnalysisHtmlReport.ps1 -JsonPath '$jsonPath'" -ForegroundColor Yellow

# Calculate and display execution time
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$timeString = if ($executionTime.TotalMinutes -ge 1) {
    "{0:N0} minutes {1:N0} seconds" -f [math]::Floor($executionTime.TotalMinutes), $executionTime.Seconds
} else {
    "{0:N1} seconds" -f $executionTime.TotalSeconds
}

Write-Host "`n⏱️  Total execution time: $timeString" -ForegroundColor Magenta
