<#
.SYNOPSIS
    Scans all accessible Key Vaults across all subscriptions and exports a CSV report locally.

.DESCRIPTION
    Connects to Azure using your current az login session, iterates all subscriptions,
    lists all Key Vaults, and collects keys and secrets (name, created date, enabled, expiry).
    Saves the result as a CSV file on disk.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to .\KeyVaultReport_<timestamp>.csv

.EXAMPLE
    .\Get-KeyVaultInventory.ps1
    .\Get-KeyVaultInventory.ps1 -OutputPath C:\Reports\kv-report.csv
#>

param(
    [string]$OutputPath = ".\KeyVaultReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

# Ensure required modules are available
foreach ($mod in @('Az.Accounts', 'Az.KeyVault', 'Az.Resources')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module $mod..." -ForegroundColor Yellow
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -ErrorAction Stop
}

az login

$results = [System.Collections.Generic.List[PSObject]]::new()
$subscriptions = Get-AzSubscription -ErrorAction Stop

Write-Host "Found $($subscriptions.Count) subscription(s).`n"

foreach ($sub in $subscriptions) {
    Write-Host "[$($sub.Name)] $($sub.Id)"

    try {
        Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null
        $keyVaults = Get-AzKeyVault -ErrorAction Stop
        Write-Host "  Found $($keyVaults.Count) Key Vault(s)"

        foreach ($kvRef in $keyVaults) {
            $kvName = $kvRef.VaultName
            Write-Host "  -> $kvName"

            # Keys
            try {
                $keys = Get-AzKeyVaultKey -VaultName $kvName -ErrorAction Stop
                foreach ($key in $keys) {
                    $results.Add([PSCustomObject]@{
                        SubscriptionId   = $sub.Id
                        SubscriptionName = $sub.Name
                        KeyVaultName     = $kvName
                        ItemType         = 'Key'
                        ItemName         = $key.Name
                        CreatedOn        = if ($key.Created) { $key.Created.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
                        Enabled          = $key.Enabled
                        ExpiresOn        = if ($key.Expires) { ([DateTime]$key.Expires).ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
                    })
                }
            }
            catch { Write-Warning "    Could not read keys from $kvName`: $_" }

            # Secrets
            try {
                $secrets = Get-AzKeyVaultSecret -VaultName $kvName -ErrorAction Stop
                foreach ($secret in $secrets) {
                    $results.Add([PSCustomObject]@{
                        SubscriptionId   = $sub.Id
                        SubscriptionName = $sub.Name
                        KeyVaultName     = $kvName
                        ItemType         = 'Secret'
                        ItemName         = $secret.Name
                        CreatedOn        = if ($secret.Created) { $secret.Created.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
                        Enabled          = $secret.Enabled
                        ExpiresOn        = if ($secret.Expires) { ([DateTime]$secret.Expires).ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
                    })
                }
            }
            catch { Write-Warning "    Could not read secrets from $kvName`: $_" }
        }
    }
    catch { Write-Warning "  Could not process subscription $($sub.Name): $_" }
}

# Save CSV
$results | Sort-Object SubscriptionName, KeyVaultName, ItemName |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$resolvedPath = Resolve-Path $OutputPath
Write-Host "`nDone. $($results.Count) items written to: $resolvedPath" -ForegroundColor Green
