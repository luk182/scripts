#Requires -Modules Az.Resources

<#
.SYNOPSIS
    Tags Azure resources based on a CSV input file.

.DESCRIPTION
    Reads a CSV with columns ResourceName, ResourceGroup, and WorkloadName.
    Applies the tag WorkloadName=<value> to each resource.
    - Skips if the tag already has the correct value.
    - Updates if the tag is missing or has a wrong value.
    Logs all actions to a timestamped CSV file.

.PARAMETER InputCsvPath
    Path to the input CSV file. Defaults to .\resources.csv

.PARAMETER LogCsvPath
    Path for the output log CSV file. Defaults to .\tagging_log_<timestamp>.csv

.EXAMPLE
    .\Set-AzureResourceTags.ps1 -InputCsvPath "C:\data\resources.csv"
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter()]
    [string]$InputCsvPath = ".\resources.csv",

    [Parameter()]
    [string]$LogCsvPath = ".\tagging_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Helpers ────────────────────────────────────────────────────────────────────

function Write-Log {
    param (
        [string]$ResourceName,
        [string]$Operation,
        [string]$Result,
        [string]$ErrorReason = "",
        [string]$OldValue    = "",
        [string]$NewValue    = ""
    )

    $entry = [PSCustomObject]@{
        Timestamp   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Resource    = $ResourceName
        Operation   = $Operation
        Result      = $Result
        ErrorReason = $ErrorReason
        OldValue    = $OldValue
        NewValue    = $NewValue
    }

    # Append to log CSV (creates file with header on first write)
    $entry | Export-Csv -Path $LogCsvPath -Append -NoTypeInformation -Encoding UTF8

    # Mirror to console with colour
    $colour = switch ($Result) {
        "Skipped" { "Cyan"   }
        "Success" { "Green"  }
        "Error"   { "Red"    }
        default   { "White"  }
    }
    Write-Host "[$Result] $ResourceName — $Operation$(if ($ErrorReason) { " | $ErrorReason" })" -ForegroundColor $colour
}

# ── Pre-flight checks ──────────────────────────────────────────────────────────

if (-not (Test-Path $InputCsvPath)) {
    Write-Error "Input CSV not found: $InputCsvPath"
    exit 1
}

$rows = Import-Csv -Path $InputCsvPath -Encoding UTF8

$requiredColumns = @("ResourceName", "ResourceGroup", "WorkloadName")
$csvColumns      = $rows[0].PSObject.Properties.Name

foreach ($col in $requiredColumns) {
    if ($col -notin $csvColumns) {
        Write-Error "Missing required column '$col' in CSV."
        exit 1
    }
}

Write-Host "`nAzure Resource Tagger" -ForegroundColor Yellow
Write-Host "Input : $InputCsvPath"  -ForegroundColor Yellow
Write-Host "Log   : $LogCsvPath`n"  -ForegroundColor Yellow

# ── Main loop ─────────────────────────────────────────────────────────────────

$tagName = "WorkloadName"
$total = $rows.Count
$i     = 0

foreach ($row in $rows) {
    $i++
    $resourceName  = $row.ResourceName.Trim()
    $resourceGroup = $row.ResourceGroup.Trim()
    $desiredValue  = $row.WorkloadName.Trim()

    Write-Progress -Activity "Tagging resources" `
                   -Status "$i / $total : $resourceName" `
                   -PercentComplete (($i / $total) * 100)

    # Skip rows with empty required fields
    if (-not $resourceName -or -not $resourceGroup -or -not $desiredValue) {
        Write-Log -ResourceName $resourceName `
                  -Operation    "Validate" `
                  -Result       "Skipped" `
                  -ErrorReason  "One or more required fields are empty in the CSV row."
        continue
    }

    try {
        # Resolve resource
        $resource = Get-AzResource -ResourceGroupName $resourceGroup `
                                   -Name $resourceName `
                                   -ErrorAction Stop

        $currentTags = $resource.Tags
        if ($null -eq $currentTags) { $currentTags = @{} }

        $currentValue = if ($currentTags.ContainsKey($tagName)) { $currentTags[$tagName] } else { "" }

        # ── Decision logic ───────────────────────────────────────────────────
        if ($currentValue -eq $desiredValue) {
            # Tag already correct — skip
            Write-Log -ResourceName $resourceName `
                      -Operation    "CheckTag" `
                      -Result       "Skipped" `
                      -OldValue     $currentValue `
                      -NewValue     $desiredValue
            continue
        }

        $operation = if ($currentValue -eq "") { "AddTag" } else { "UpdateTag" }

        if ($PSCmdlet.ShouldProcess($resourceName, "$operation '$tagName' → '$desiredValue'")) {
            $currentTags[$tagName] = $desiredValue

            Update-AzTag -ResourceId $resource.Id `
                         -Tag        @{ $tagName = $desiredValue } `
                         -Operation  Merge `
                         -ErrorAction Stop | Out-Null

            Write-Log -ResourceName $resourceName `
                      -Operation    $operation `
                      -Result       "Success" `
                      -OldValue     $currentValue `
                      -NewValue     $desiredValue
        }

    } catch {
        Write-Log -ResourceName $resourceName `
                  -Operation    "Tag" `
                  -Result       "Error" `
                  -ErrorReason  $_.Exception.Message `
                  -OldValue     "" `
                  -NewValue     $desiredValue
    }
}

Write-Progress -Activity "Tagging resources" -Completed

Write-Host "`nDone. Log saved to: $LogCsvPath" -ForegroundColor Yellow
