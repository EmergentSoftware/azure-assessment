# Azure Managed Disk Analysis

Telemetry-driven cost and performance assessment for Azure managed disks.

The analyzer discovers every managed disk in scope, pulls Azure Monitor metrics for all of
them in batches, prices each disk against live Azure retail rates **for its own region and
redundancy**, and emits a recommendation with the money attached to it. A companion script
renders the results as a self-contained interactive HTML report.

---

## Contents

| File | Purpose |
| --- | --- |
| [azure-managed-disk-telemetry-analysis.ps1](azure-managed-disk-telemetry-analysis.ps1) | Analysis engine. Writes JSON + CSV. |
| [New-DiskAnalysisHtmlReport.ps1](New-DiskAnalysisHtmlReport.ps1) | Renders the HTML report from the JSON. |
| [disk-specs.json](disk-specs.json) | Per-SKU performance limits and fallback billing rates. |

---

## Quick start

```powershell
# Analyse the current subscription over the last 7 days
./azure-managed-disk-telemetry-analysis.ps1

# Then render the report from the JSON it printed
./New-DiskAnalysisHtmlReport.ps1 -JsonPath ./disk-analysis-20260115-120000.json
```

### Prerequisites

- PowerShell 7+
- Azure CLI, signed in (`az login`). The CLI is used **only** to acquire access tokens;
  all data access goes through the ARM REST APIs directly.
- **Reader** on the target subscriptions and **Monitoring Reader** for metrics.

Tokens are requested per subscription, so a single run can span multiple tenants.

---

## Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `-DaysToInspect` | `7` | Days of history to analyse (1–365). Drives the metric interval. |
| `-ExceedanceThreshold` | `0.05` | Fraction of samples above a limit before the disk is considered saturated. |
| `-SubscriptionId` | current | One or more subscription IDs to analyse. |
| `-SpecFile` | `disk-specs.json` | Path to the spec/rates file. |
| `-OutputPrefix` | `disk-analysis` | Prefix for the timestamped output files. |
| `-ExcludeAvd` | off | Skip disks whose name contains "AVD". |
| `-PricingRegion` | per-disk | Force a single region for pricing instead of each disk's own. |
| `-Currency` | `USD` | Currency code passed to the Retail Prices API. |
| `-NoLivePricing` | off | Use the fallback rates in `disk-specs.json` instead of calling the pricing API. |

### Examples

```powershell
# Two subscriptions, 30 days, skip AVD session hosts
./azure-managed-disk-telemetry-analysis.ps1 -SubscriptionId 'sub-a','sub-b' -DaysToInspect 30 -ExcludeAvd

# Price everything against West US 2 in euros
./azure-managed-disk-telemetry-analysis.ps1 -PricingRegion westus2 -Currency EUR

# Fully offline from the pricing API
./azure-managed-disk-telemetry-analysis.ps1 -NoLivePricing
```

---

## How it works

```
Resource Graph  ──►  one query returns every disk + every VM storage profile
       │
Azure Monitor   ──►  metrics:getBatch, up to 50 resources per call
       │             (disk IOPS/throughput, VM OS + data disk latency by LUN)
       │
Retail Prices   ──►  one catalog per region, cached
       │
   Analysis     ──►  JSON + CSV  ──►  HTML report
```

Everything is fetched up front in batches, so the per-disk loop is pure computation. A
33-disk estate costs roughly **six HTTP calls** and finishes in about 25 seconds.

### Metric granularity

| `-DaysToInspect` | Interval | Reason |
| --- | --- | --- |
| ≤ 29 | `PT1M` | Full resolution |
| 30–89 | `PT5M` | Azure Monitor retention limit |
| ≥ 90 | `PT1H` | Azure Monitor retention limit |

> Managed disk metrics support the **Average** aggregation only. Percentiles are therefore
> computed over interval averages, not over raw I/O.

---

## Pricing model

Prices come from the [Azure Retail Prices API](https://prices.azure.com/api/retail/prices),
resolved per region and per redundancy, and cached once per region.

| Charge | Applies to |
| --- | --- |
| Provisioned capacity | All types (LRS and ZRS) |
| Transactions (per 10K) | Standard HDD, Standard SSD |
| Burst enablement + burst transactions | Premium SSD **P30 and above** only |
| Provisioned IOPS and throughput | Premium SSD v2, Ultra |
| 1-year reservation | Premium SSD **P30–P80** only |

Two details that are easy to get wrong and are handled explicitly:

- Each SKU publishes **two** meters under the same `skuName` — `P6 LRS Disk` ($10.21) and
  `P6 LRS Disk Mount` ($0.56, the shared-disk fee). Matching on `skuName` alone under-reports
  by ~94%, so capacity is matched on `meterName`.
- Meters are republished per price revision. The **most recently effective** row is selected;
  picking the lowest price returns a retired rate.

If the API is unreachable, the script falls back to the rates in `disk-specs.json` and stamps
every record with `PricingSource = SpecFile` so a stale run is visible in the output.

> Fallback rates are North Central US LRS as of the `AsOf` date in the file. ZRS is derived by
> applying `ZrsUpliftFactor` (1.5×).

---

## Recommendations

Each disk receives a decision, a severity, and a monthly saving.

| Severity | Meaning |
| --- | --- |
| `Action` | Spend with no workload behind it (unattached disks) |
| `Opportunity` | A cheaper configuration fits the observed usage |
| `Watch` | At or above its limits, or bursting repeatedly |
| `Unknown` | No telemetry, or no spec for this SKU |
| `Optimal` | Already the best fit |

### What is evaluated per disk type

| Disk type | Checks |
| --- | --- |
| Standard HDD | Upgrade to Standard SSD on saturation or when transaction charges make SSD cheaper |
| Standard SSD | Upgrade to Premium on cost/latency/throughput; downgrade to HDD when underutilized |
| Premium SSD | Downgrade to Standard SSD, migrate to v2, or buy a 1-year reservation |
| Premium SSD v2 / Ultra | Raise or lower **provisioned** IOPS and MB/s; compare against a fixed Premium tier |
| Any, unattached | Delete or snapshot — the whole monthly cost is the saving |
| Any, VM deallocated | Flagged for review; disk bills while the VM is stopped |

### Guardrails

These prevent confidently wrong advice:

- **Premium SSD v2 and Ultra cannot be OS disks.** Every v2/Ultra recommendation is gated on
  the disk not being an OS disk, even when v2 is clearly cheaper.
- **No telemetry means no sizing call.** With zero samples every percentile is `0` and every
  lower tier appears to fit, so all downgrade and right-size advice requires observed data.
- **Right-sizing targets P99 × 1.3**, not the peak, so one outlier minute cannot drive the
  recommendation — and not the mean, so normal spikes are still covered.
- **Latency gates downgrades.** Standard SSD → HDD requires ≤ 5 ms; Premium → Standard SSD
  requires ≤ 10 ms.
- **$5/month floor** on any change, to avoid churn for trivial amounts.

---

## Output

Three files per run, timestamped:

```
disk-analysis-YYYYMMDD-HHMMSS.json   # full record, 57 fields
disk-analysis-YYYYMMDD-HHMMSS.csv    # same data, flat
disk-analysis-YYYYMMDD-HHMMSS.html   # generated separately
```

### Key fields

| Group | Fields |
| --- | --- |
| Identity | `Disk`, `ResourceGroup`, `Subscription`, `Location`, `SKU`, `Redundancy` |
| State | `DiskState`, `Attached`, `AgeDays`, `MetricsStatus`, `Samples` |
| Sizing | `ActualSizeGiB`, `BilledSizeGiB`, `BillingTier`, `CustomSize`, `PerformancePlus` |
| Observed | `PeakIOPS`, `P99IOPS`, `P95IOPS`, `PeakMBps`, `P99MBps`, `P95MBps`, `AvgLatencyMs`, `Transactions` |
| Limits | `EffectiveMaxIOPS`, `EffectiveMaxMBps`, `ProvisionedIOPS`, `ProvisionedMBps`, `PercentIOPSAbove`, `PercentMBpsAbove` |
| Bursting | `BurstOperations`, `BurstWindowCount`, `BurstWindows`, `EstMonthlyBurstCost` |
| Cost | `EstMonthlyTotalCost`, `EstMonthlyHDDTotalCost`, `EstMonthlySSDTotalCost`, `EstMonthlyPremiumTotalCost`, `EstMonthlyPremiumV2TotalCost`, `ReservationMonthlyCost`, `ReservationMonthlySavings` |
| Provenance | `PricingSource`, `PricingRegion`, `CurrencyCode` |
| Verdict | `Decision`, `DecisionDetail`, `Severity`, `PotentialMonthlySavings` |

### HTML report

- Summary cards: disk count, current monthly spend, identified savings (monthly, annual, % of spend)
- Rows ordered by severity then savings, so the top of the table is the work queue
- Per-column filters, click-to-sort, and CSV export of the filtered view
- "Show savings only" collapses to just the actionable disks and totals their savings

Savings and severity are read straight from the JSON rather than recalculated, so the report
cannot disagree with the recommendation it displays.

---

## `disk-specs.json`

Two sections:

- **Performance limits** per SKU (`StandardHDD`, `StandardSSD`, `PremiumSSD`), keyed by billed
  size in GiB, used to judge saturation. `MonthlyCost` on each entry is a *fallback* only.
- **`Rates`** — transaction, burst and ZRS uplift rates used when live pricing is unavailable.
  Update `AsOf` when refreshing them.

Premium SSD v2 and Ultra have no entries because they carry no tier ladder; their limits are
read from the disk's own `diskIOPSReadWrite` / `diskMBpsReadWrite`.

---

## Caveats

- Costs are **retail estimates**. They exclude snapshots, egress, and any negotiated or EA discount.
- Transaction counts are projected from the observed window to a 730-hour month. Azure applies
  hourly billable-transaction caps that are not modelled, so transaction cost is an upper bound.
- Standard HDD bills per 16 KiB I/O and Standard SSD per 256 KiB; the script counts operations,
  so large-I/O workloads are undercounted.
- Premium SSD v2 generally requires zonal VMs. The script checks the OS-disk constraint but not
  zonal placement — validate before migrating.
- Latency is read at the VM, so it is unavailable for unattached disks and for stopped VMs.
