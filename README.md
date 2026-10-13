# Azure Assessment

This workspace contains standalone PowerShell assessment scripts and an Azure Functions backend that wraps the VM analyzer as an API-driven job system.

## Workspace Map

- [azure-vm-analysis](azure-vm-analysis)
    PowerShell VM assessment engine, config samples, and generated report examples.
- [azure-managed-disk-analysis](azure-managed-disk-analysis)
    Managed disk telemetry assessment scripts and report assets.
- [azure-recovery-services-vault-analysis](azure-recovery-services-vault-analysis)
    Recovery Services Vault assessment scripts and sample outputs.
- [functions](functions)
    Azure Functions app for submitting VM assessments, tracking jobs, and downloading results.
- [infra](infra)
    Infrastructure definitions for the Azure-hosted deployment path.

## What Each Major Script Or App Does

### VM Assessment

- [azure-vm-analysis/azure-vm-assessment.ps1](azure-vm-analysis/azure-vm-assessment.ps1)
    Main VM analysis engine. Collects Azure Advisor findings, Azure Monitor utilization, pricing, and candidate SKU recommendations, then writes JSON, CSV, and HTML reports.
- [azure-vm-analysis/azure-vm-assessment.config.jsonc](azure-vm-analysis/azure-vm-assessment.config.jsonc)
    Sample config file for the VM analyzer.
- [azure-vm-analysis/excluded-vms.example.yaml](azure-vm-analysis/excluded-vms.example.yaml)
    Sample YAML file for excluding VM names from the assessment.

### Managed Disk Assessment

See [azure-managed-disk-analysis/README.md](azure-managed-disk-analysis/README.md) for full documentation.

- [azure-managed-disk-analysis/azure-managed-disk-telemetry-analysis.ps1](azure-managed-disk-analysis/azure-managed-disk-telemetry-analysis.ps1)
    Discovers disks via Resource Graph, batches Azure Monitor metrics, prices each disk against live Azure retail rates for its own region and redundancy, and emits a recommendation with the monthly saving attached.
- [azure-managed-disk-analysis/New-DiskAnalysisHtmlReport.ps1](azure-managed-disk-analysis/New-DiskAnalysisHtmlReport.ps1)
    Generates the interactive managed disk HTML report, including spend and savings totals.
- [azure-managed-disk-analysis/disk-specs.json](azure-managed-disk-analysis/disk-specs.json)
    Per-SKU performance limits plus the fallback billing rates used when live pricing is unavailable.

### Recovery Services Vault Assessment

- [azure-recovery-services-vault-analysis/recovery-services-vault-analysis.ps1](azure-recovery-services-vault-analysis/recovery-services-vault-analysis.ps1)
    Reviews vault backup posture, churn, and storage-related risk signals.

### Functions Backend

- [functions/README.md](functions/README.md)
    Detailed contract for the Azure Functions backend, including endpoints, app settings, and internal workers.
- [functions/src/functions/submitAssessment.ts](functions/src/functions/submitAssessment.ts)
    HTTP entrypoint for creating a VM assessment job.
- [functions/src/functions/runAssessmentJob.ts](functions/src/functions/runAssessmentJob.ts)
    Queue-triggered worker that executes the PowerShell analyzer.
- [functions/src/shared/assessmentRunner.ts](functions/src/shared/assessmentRunner.ts)
    PowerShell execution wrapper that resolves config, forwards helper paths, parses artifact paths, and uploads outputs.

## Functions API Quick Start

The Functions app is documented in detail in [functions/README.md](functions/README.md). At a high level it exposes:

- `POST /api/assessments`
- `GET /api/assessments`
- `GET /api/assessments/{jobId}`
- `GET /api/assessments/{jobId}/results`
- `GET /api/assessments/{jobId}/artifacts/{artifactName}`
- `GET /api/context`

When deployed to Azure, these HTTP functions require a function key.

## VM Assessment Quick Start

Prerequisites:

- Azure CLI installed and logged in with `az login`
- PowerShell 7 or later recommended
- Read access to the target subscriptions and resources

Run the standalone VM analyzer with a config file:

```powershell
Set-Location azure-vm-analysis
.\azure-vm-assessment.ps1 -ConfigPath .\azure-vm-assessment.config.jsonc
```

Run with explicit parameters:

```powershell
.\azure-vm-assessment.ps1 `
    -SubscriptionId "11111111-1111-1111-1111-111111111111" `
    -DaysToInspect 14 `
    -OutputPrefix vm-analysis-weekly
```

Run with exclusions:

```powershell
.\azure-vm-assessment.ps1 `
    -ConfigPath .\azure-vm-assessment.config.jsonc `
    -ExcludeVMName vmlegacy001,vmlegacy002 `
    -ExcludeVmListPath .\excluded-vms.example.yaml
```

## Local Functions Quick Start

```powershell
Set-Location functions
Copy-Item local.settings.template.json local.settings.json
npm install
npm run build
npm run start:storage
npm start
```

Important local note:

- Azurite-generated files are ignored and should not be committed.
- The VM pricing cache and VM SKU/spec cache both rely on local table and queue storage when `AzureWebJobsStorage=UseDevelopmentStorage=true`.

## Outputs

The standalone scripts and the Functions-backed VM workflow produce:

- JSON reports for machine-readable downstream processing
- CSV reports for Excel or Power BI analysis
- HTML reports for interactive review

The Functions path also produces an execution log and manifest per job.
| `UnderutilizedCpuAverageThreshold` | Double | 10 | CPU average threshold for underutilization |
| `UnderutilizedCpuP95Threshold` | Double | 25 | CPU P95 threshold for underutilization |
| `OverutilizedCpuAverageThreshold` | Double | 65 | CPU average threshold for overutilization |
| `OverutilizedCpuP95Threshold` | Double | 85 | CPU P95 threshold for overutilization |
| `LowNetworkAverageThresholdMBps` | Double | 0.5 | Network MB/s threshold for low activity |
| `LowDiskAverageThresholdMBps` | Double | 1.0 | Disk MB/s threshold for low activity |
| `BurstableLowCreditsThreshold` | Double | 20 | B-series CPU credits threshold |
| `MinimumSamplesForClassification` | Int | 12 | Minimum CPU samples required to classify a VM |

### VM Analysis Config Notes

- The sample config file is at `azure-vm-analysis/azure-vm-assessment.config.jsonc`.
- Stopped or deallocated VMs are skipped automatically before telemetry collection and pricing analysis.
- YAML exclusion files can be either a top-level list or a named list such as `ExcludeVMName:`.
- Use `.jsonc` when you want inline comments. Use plain `.json` only if you remove comments.
- Command-line parameters override config values when both are provided.
- Report behavior is configured under the `Report` object, including output directory, output format toggles, and HTML report text.
- Current-SKU pricing now records PAYG and reservation matches separately, so `RI meters not found` should only appear when the reservation price really is unavailable.
- Windows pricing includes OS license only when the VM is not using a license benefit such as Azure Hybrid Benefit. For reservations, the script derives the monthly Windows uplift from the difference between Windows PAYG and base compute PAYG, then adds that uplift to the reservation compute monthly estimate.
- The HTML report uses the recommended target SKU for the `Cost Comparison` column when a target SKU is available.
- Aggregate cards show estimated monthly cost for PAYG, 1YR RI, and 3YR RI rather than visible savings totals.
- CPU Avg % and CPU P95 % were removed from the main table and are now surfaced in the expandable rationale section under `Data`.
- The HTML report includes a best target SKU summary in the main grid and a top-3 B/D/E candidate table inside each VM detail panel.
- SKU recommendations prefer the current VM family first and only switch families when the utilization profile strongly supports it.

## 📁 Repository Structure

```
azure-assessment/
├── README.md                                    # This file
├── .gitignore                                   # Git ignore patterns
│
├── azure-managed-disk-analysis/
│   ├── README.md                                # Disk analysis documentation
│   ├── azure-managed-disk-telemetry-analysis.ps1 # Main disk analysis script
│   ├── disk-specs.json                          # Disk SKU specifications
│   └── disk-analysis-*.{html,csv,json}          # Generated reports
│
├── azure-vm-analysis/
│   ├── azure-vm-assessment.ps1                  # Main VM analysis script
│   ├── azure-vm-assessment.config.jsonc         # Sample annotated config file
│   └── vm-analysis-*.{html,csv,json}            # Generated reports
│
└── azure-recovery-services-vault-analysis/
    ├── README.md                                # Vault analysis documentation
    ├── recovery-services-vault-analysis.ps1     # Main vault analysis script
    └── vault-analysis-*.{html,csv,json}         # Generated reports
```

## 🚨 Troubleshooting

### Common Issues

#### Azure CLI Not Found
```powershell
# Install Azure CLI
winget install Microsoft.AzureCLI
# Or download from: https://aka.ms/installazurecliwindows
```

#### Authentication Errors
```powershell
# Re-authenticate with Azure
az login

# Verify current subscription
az account show

# List available subscriptions
az account list --output table

# Set default subscription
az account set --subscription "subscription-name-or-id"
```

#### No Data Returned
- Verify you have resources in the subscription (disks or vaults)
- Check that metrics are being collected (may take time for new resources)
- Ensure you have read permissions on the resources
- For disk analysis: Wait at least 24 hours after disk creation for metrics

#### Slow Performance
- Reduce `DaysToInspect` parameter for faster analysis
- Analyze one subscription at a time for large environments
- Check network connectivity to Azure

#### Permission Errors
Required Azure RBAC roles:
- **Disk Analysis**: Reader role on subscriptions/resource groups
- **Vault Analysis**: Backup Reader or Backup Operator role

## 💡 Best Practices

### Running Regular Assessments
```powershell
# Weekly disk analysis (recommended)
.\azure-managed-disk-telemetry-analysis.ps1 -DaysToInspect 7

# Monthly backup analysis (recommended)
.\recovery-services-vault-analysis.ps1 -DaysToInspect 30
```

### Multi-Subscription Organizations
```powershell
# Get all subscription IDs
$subs = (az account list --query "[].id" -o tsv)

# Run analysis across all subscriptions
.\azure-managed-disk-telemetry-analysis.ps1 -SubscriptionId $subs
```

### Automating Reports
Create a scheduled task or Azure Automation runbook to:
1. Run scripts weekly/monthly
2. Store reports in Azure Storage or SharePoint
3. Email reports to stakeholders
4. Track trends over time

### Cost Savings Tracking
1. Export initial assessment to CSV
2. Implement recommended changes
3. Run assessment again after 30 days
4. Compare results to validate savings

## 📊 Understanding the Metrics

### Disk Analysis Metrics

- **IOPS Above %**: Percentage of time disk IOPS exceeded provisioned limits
- **MBps Above %**: Percentage of time throughput exceeded limits
- **Avg Latency**: Average disk response time (lower is better)
- **Transactions**: Total read/write operations during analysis period
- **Cost Diff**: Monthly cost difference between current and recommended SKU

### Vault Analysis Metrics

- **Daily Churn GB**: Average daily change in backup size
- **Monthly Churn GB**: Estimated monthly data change (Daily × 30)
- **Total Backup Size GB**: Current total backup storage consumed
- **Recovery Points**: Number of backup snapshots available
- **Risk Level**: High/Medium/Low based on churn and health status

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## 📝 License

This project is provided as-is for Azure infrastructure assessment purposes.

## 🔗 Related Resources

- [Azure Disk Pricing](https://azure.microsoft.com/pricing/details/managed-disks/)
- [Azure Backup Pricing](https://azure.microsoft.com/pricing/details/backup/)
- [Azure CLI Documentation](https://docs.microsoft.com/cli/azure/)
- [Azure Monitor Metrics](https://docs.microsoft.com/azure/azure-monitor/essentials/metrics-supported)

## ✉️ Support
- Excel/spreadsheet analysis
- Power BI imports
- Further data processing

### 3. JSON Report (`vault-analysis-YYYYMMDD-HHMMSS.json`)

Structured JSON data for:
- Programmatic access
- API integration
- Custom reporting tools

## Contributing

Contributions are welcome! To add new assessment tools:

1. Follow the established PowerShell script structure
2. Include HTML, CSV, and JSON output formats
3. Add comprehensive documentation to this README
4. Include parameter descriptions and examples

## License

These scripts are provided as-is for Azure infrastructure assessment purposes.

## Additional Resources

- [Azure CLI Documentation](https://docs.microsoft.com/en-us/cli/azure/)
- [Azure Monitor Metrics](https://docs.microsoft.com/en-us/azure/azure-monitor/essentials/metrics-supported)
- [Azure Backup Documentation](https://docs.microsoft.com/en-us/azure/backup/)
- [Azure Cost Management](https://docs.microsoft.com/en-us/azure/cost-management-billing/)

---

For detailed documentation on each tool, see their respective sections above. For issues or questions, verify Azure CLI installation and permissions before running the scripts.
