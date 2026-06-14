# Set-AzureResourceTags

A PowerShell script to bulk-tag Azure resources from a CSV file. It applies a `WorkloadName` tag to each resource, handles all three scenarios (tag missing, tag wrong, tag already correct), and logs every action to a timestamped CSV file.

---

## Prerequisites

- PowerShell 5.1+ or PowerShell 7+
- [Az PowerShell module](https://learn.microsoft.com/en-us/powershell/azure/install-az-ps) installed and authenticated

```powershell
# Install Az module if needed
Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -Force

# Authenticate
Connect-AzAccount
```

---

## Input CSV Format

Prepare a CSV file with exactly these three columns:

| Column | Description |
|---|---|
| `ResourceName` | The name of the Azure resource |
| `ResourceGroup` | The resource group it belongs to |
| `WorkloadName` | The desired value for the `WorkloadName` tag |

**Example `resources.csv`:**

```csv
ResourceName,ResourceGroup,WorkloadName
my-vm,rg-production,Finance
my-storage,rg-staging,HR
app-service-01,rg-production,Marketing
```

---

## Usage

```powershell
# Default — looks for resources.csv in the current directory
.\Set-AzureResourceTags.ps1

# Custom input file path
.\Set-AzureResourceTags.ps1 -InputCsvPath "C:\data\resources.csv"

# Custom input and custom log output path
.\Set-AzureResourceTags.ps1 -InputCsvPath ".\resources.csv" -LogCsvPath ".\my_log.csv"

# Dry-run: preview all actions without making any changes
.\Set-AzureResourceTags.ps1 -WhatIf
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-InputCsvPath` | `.\resources.csv` | Path to the input CSV file |
| `-LogCsvPath` | `.\tagging_log_<timestamp>.csv` | Path for the output log file |
| `-WhatIf` | — | Preview mode, no changes are applied |

---

## Tagging Logic

| Scenario | Action |
|---|---|
| Tag exists with the **correct** value | ⏭️ Skipped — no API call made |
| Tag is **missing** (not set) | ✅ Tag added |
| Tag exists with the **wrong** value | 🔄 Tag updated |
| Resource not found or API error | ❌ Error logged, script continues |

> **Note:** Only the `WorkloadName` tag is touched. All other existing tags on a resource are preserved.

---

## Output Log

A CSV log file is created automatically for every run (timestamped to avoid overwrites).

**Columns:**

| Column | Description |
|---|---|
| `Timestamp` | Date and time of the action |
| `Resource` | Name of the resource |
| `Operation` | `CheckTag`, `AddTag`, `UpdateTag`, or `Validate` |
| `Result` | `Success`, `Skipped`, or `Error` |
| `ErrorReason` | Error message if the operation failed |
| `OldValue` | The tag value before the change (empty if tag was missing) |
| `NewValue` | The intended tag value |

**Example log output:**

```
Timestamp,Resource,Operation,Result,ErrorReason,OldValue,NewValue
2025-06-14 10:22:01,my-vm,AddTag,Success,,, Finance
2025-06-14 10:22:04,my-storage,CheckTag,Skipped,,,HR
2025-06-14 10:22:07,app-service-01,UpdateTag,Success,,OldTeam,Marketing
```

---

## Notes

- Rows with empty `ResourceName`, `ResourceGroup`, or `WorkloadName` fields are skipped and logged.
- Run `.\Set-AzureResourceTags.ps1 -WhatIf` first on a test subscription to validate the CSV before a production run.
- The script requires at least **Contributor** role (or a custom role with `Microsoft.Resources/tags/write`) on the target resources.

## DISCLAIMER

This sample script is not supported under any Microsoft standard support program or service. The sample script is provided "AS IS" without warranty of any kind.

Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.

The entire risk arising out of the use or performance of the sample scripts and documentation remains with you.

In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
`
