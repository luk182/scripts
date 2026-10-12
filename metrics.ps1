#Requires -Version 7.0
<#
.SYNOPSIS
    Collects Azure resource utilization metrics (last calendar month) across one or more subscriptions
    and exports the results to a multi-tab Excel workbook. Also inventories orphaned (unattached)
    managed disks with an estimated monthly cost pulled from the public Azure Retail Prices API.

.DESCRIPTION
    Every metric used below was looked up and confirmed against the live Microsoft Learn
    "Supported metrics" reference pages (not inferred from memory). The exact metric/property
    and its documentation source is listed in the table below and is also written out to a
    "Metric-Reference" worksheet in the generated workbook for traceability.

    | Resource type            | What we report                              | Azure Monitor metric / property (REST name)                              | Source (Microsoft Learn)                                                                                                   |
    |---------------------------|----------------------------------------------|----------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------|
    | Microsoft.Compute/virtualMachines            | Avg & Max CPU %                | `Percentage CPU`                                                          | /azure/azure-monitor/reference/supported-metrics/microsoft-compute-virtualmachines-metrics                                  |
    | Microsoft.Compute/virtualMachines            | Max Memory Used % (= 100 - min Available Memory %) | `Available Memory Percentage` (guest-OS metric, requires Azure Monitor Agent) | /azure/azure-monitor/reference/supported-metrics/microsoft-compute-virtualmachines-metrics                      |
    | Microsoft.ApiManagement/service               | Scale units (capacity)         | `sku.capacity` resource property (not an Azure Monitor metric)            | /powershell/module/az.apimanagement/get-azapimanagement                                                                      |
    | Microsoft.ApiManagement/service               | Avg Gateway CPU %              | `Capacity` (classic SKUs) / `CpuPercent_Gateway` (v2 SKUs)                | /azure/azure-monitor/reference/supported-metrics/microsoft-apimanagement-service-metrics                                     |
    | Microsoft.Web/serverfarms (App Service Plan)  | Max CPU %                       | `CpuPercentage`                                                           | /azure/azure-monitor/reference/supported-metrics/microsoft-web-serverfarms-metrics                                           |
    | Microsoft Sentinel (Log Analytics workspace)  | Detect whether Sentinel is onboarded | `Microsoft.SecurityInsights/onboardingStates` List operation, queried directly per workspace | /rest/api/securityinsights/sentinel-onboarding-states/list |
    | Microsoft Sentinel (Log Analytics workspace)  | Avg daily ingestion (GB/day)   | `Usage` table KQL query (no native Azure Monitor platform metric exists)  | /azure/azure-monitor/logs/analyze-usage ; confirmed no ingestion metric exists on microsoft-operationalinsights-workspaces-metrics |
    | Microsoft Sentinel / Log Analytics workspace  | Daily commitment tier in place | `sku.capacityReservationLevel` resource property                         | /azure/templates/microsoft.operationalinsights/workspaces                                                                    |
    | Microsoft.Storage/storageAccounts/fileServices (Premium Files) | Used vs allocated per share | `FileCapacity` (used, bytes) & `FileShareCapacityQuota` (allocated, bytes), split by `FileShare` dimension | /azure/azure-monitor/reference/supported-metrics/microsoft-storage-storageaccounts-fileservices-metrics |
    | Microsoft.Sql/servers/databases (Azure SQL DB)| Avg & Max CPU %                 | `cpu_percent`                                                             | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics                                     |
    | Microsoft.Sql/servers/databases (Azure SQL DB)| Max Memory Used %              | `sql_instance_memory_percent`                                             | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics                                     |
    | Microsoft.Sql/managedInstances (SQL MI)        | Avg & Max CPU %                 | `avg_cpu_percent`                                                         | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics                                      |
    | Microsoft.Sql/managedInstances (SQL MI)        | Max Memory Used %              | Not available - no memory metric exists for this resource type in Azure Monitor (confirmed against the live docs) | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics |
    | Microsoft.CognitiveServices/accounts           | Max Provisioned-managed v2 %    | `AzureOpenAIProvisionedManagedUtilizationV2`                              | /azure/azure-monitor/reference/supported-metrics/microsoft-cognitiveservices-accounts-metrics                                |
    | Microsoft.Compute/disks                        | Orphaned (unattached) disks     | `properties.diskState == 'Unattached'` via Azure Resource Graph           | Azure Resource Graph "resources" table, community-documented, verified live against Resource Graph schema                    |
    | Managed disk pricing                           | Estimated monthly cost          | Azure Retail Prices API (`https://prices.azure.com/api/retail/prices`)    | /rest/api/cost-management/retail-prices/azure-retail-prices                                                                  |

    NOTE on aggregation technique: Azure Monitor Metrics only lets you request one time grain per
    call. To get a true Avg/Max over an entire month without exceeding the metrics API's per-call
    data point limits, this script queries at an hourly (PT1H) grain, then:
      - "Avg over last month"  = arithmetic mean of all hourly Average values
      - "Max over last month"  = highest value among all hourly Maximum values
    This is the standard approach recommended in Azure Monitor documentation for long look-back
    windows (see /azure/azure-monitor/essentials/metrics-aggregation-explained).

.PARAMETER SubscriptionId
    One or more subscription IDs to scan. Defaults to every subscription enabled for the signed-in
    account.

.PARAMETER TenantId
    Azure AD tenant (directory) ID to report on. If omitted and the signed-in account has access
    to more than one tenant, the script lists them and prompts you to pick one interactively,
    rather than silently defaulting to your home tenant.

.PARAMETER OutputPath
    Path to the .xlsx file to create. Defaults to .\AzureUtilizationReport_<yyyyMMdd>.xlsx in the
    current directory.

.PARAMETER ReportMonth
    Optional "yyyy-MM" string to report on a specific past calendar month. Defaults to last full
    calendar month relative to today.

.EXAMPLE
    .\Get-AzureUtilizationReport.ps1

.EXAMPLE
    .\Get-AzureUtilizationReport.ps1 -SubscriptionId "11111111-1111-1111-1111-111111111111" -ReportMonth "2025-09" -OutputPath "C:\Reports\Utilization.xlsx"
#>

[CmdletBinding()]
param(
    [string[]] $SubscriptionId,

    # Azure AD tenant (directory) to report on. If omitted, and the signed-in account has access
    # to more than one tenant (common for consultants with guest access to client tenants), the
    # script lists them and prompts you to pick one interactively instead of silently defaulting
    # to your home tenant.
    [string] $TenantId,

    [string] $OutputPath = ".\AzureUtilizationReport_$(Get-Date -Format 'yyyyMMdd').xlsx",

    [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')]
    [string] $ReportMonth,

    [int] $TimeGrainHours = 1
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 0. Setup: modules, auth, time window
# ---------------------------------------------------------------------------

<#
    Az.Accounts ships native/compiled assemblies that .NET cannot unload or swap once a
    different version has already been imported into the current PowerShell process. If an
    older Az.Accounts got auto-loaded earlier in this session (VS Code's PowerShell extension,
    a profile script, or a prior command), importing a module that requires a newer Az.Accounts
    will fail with "This module requires Az.Accounts version X... Please open a new session"
    (see https://aka.ms/azps-version-error). The only reliable fix is a brand-new session, so we
    detect that up front and fail fast with clear guidance instead of a confusing mid-script error.
#>
function Test-AzAccountsVersionConflict {
    $loaded = Get-Module -Name Az.Accounts
    if (-not $loaded) { return }

    $latestInstalled = Get-Module -ListAvailable -Name Az.Accounts |
    Sort-Object Version -Descending | Select-Object -First 1

    if ($latestInstalled -and $loaded.Version -lt $latestInstalled.Version) {
        throw @"
Az.Accounts version $($loaded.Version) is already loaded in this PowerShell session, but the
newer version $($latestInstalled.Version) is installed on this machine and is required by one of
the other Az modules this script uses. PowerShell cannot swap a loaded module's version mid-session.

FIX: Close this terminal/session entirely and open a brand-new PowerShell window (or a new VS Code
integrated terminal), then re-run this script. If the error persists, run:
    Update-Module -Name Az.Accounts, Az.Resources, Az.Monitor, Az.ResourceGraph, Az.OperationalInsights, Az.Storage -Force
in a fresh session first to align all Az module versions, then open another new session to run the report.
More info: https://aka.ms/azps-version-error
"@
    }
}
Test-AzAccountsVersionConflict

function Install-RequiredModule {
    param([Parameter(Mandatory)] [string] $Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing missing module: $Name" -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
    Import-Module -Name $Name -ErrorAction Stop
}

Write-Host "Checking required modules..." -ForegroundColor Cyan
# Az.Accounts must be imported first and explicitly, since every other Az.* module depends on it
# and will otherwise trigger its own (less clear) implicit import with a potentially mismatched version.
Install-RequiredModule -Name 'Az.Accounts'
foreach ($m in @('Az.Resources', 'Az.Monitor', 'Az.ResourceGraph', 'Az.OperationalInsights', 'Az.Storage', 'ImportExcel')) {
    Install-RequiredModule -Name $m
}

<#
    Wraps Connect-AzAccount and turns the common AADSTS50105 failure into a clear, actionable
    error. AADSTS50105 means the TARGET TENANT's Entra ID admin has restricted the first-party
    "Microsoft Azure PowerShell" application to require explicit user/group assignment, and the
    signed-in account hasn't been assigned. This is a tenant-side policy, not a script or client
    bug: Connect-AzAccount's interactive sign-in always uses that fixed application ID (the
    -ApplicationId parameter only applies to service-principal auth, not interactive user login),
    so there is no client-side workaround - the target tenant's admin must grant the assignment.
#>
function Connect-AzAccountWithGuidance {
    param([string] $Tenant, [switch] $Force)
    try {
        if ($Tenant) { Connect-AzAccount -Tenant $Tenant -Force:$Force | Out-Null }
        else { Connect-AzAccount | Out-Null }
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'AADSTS50105') {
            throw @"
Sign-in was blocked by the target tenant's Entra ID (Azure AD) configuration - AADSTS50105 - not by
this script. That tenant's admin has restricted the first-party "Microsoft Azure PowerShell"
application (App ID 1950a258-227b-4e31-a9cf-717495945fc2) to require explicit user/group
assignment, and your account does not currently have that assignment$(if ($Tenant) { " in tenant '$Tenant'" }).

FIX (requires the target tenant's Entra ID admin):
  1. Entra admin center > Identity > Applications > Enterprise applications
  2. Search for "Microsoft Azure PowerShell" (App ID 1950a258-227b-4e31-a9cf-717495945fc2)
  3. Open "Users and groups" > "Add user/group" > assign your account (or a group you belong to)
     (alternatively the admin can set "Assignment required?" = No under the app's Properties,
     though per-user/group assignment is the safer option)

This cannot be worked around from the script or client side - Connect-AzAccount's interactive
sign-in always uses this fixed first-party application ID. Once access is assigned, re-run this
script in a new PowerShell session.

Original error: $msg
"@
        }
        else { throw }
    }
}

if (-not (Get-AzContext)) {
    Write-Host "No active Azure session found. Launching Connect-AzAccount..." -ForegroundColor Yellow
    Connect-AzAccountWithGuidance
}

<#
    Connect-AzAccount without -Tenant authenticates to the signed-in account's HOME tenant, and a
    cached Get-AzContext from a previous session is reused as-is even if it points at the wrong
    tenant. For accounts with guest/B2B access to multiple tenants (e.g. a consultant working
    across client tenants), that means the script can silently keep targeting the wrong directory.
    To fix this, we explicitly resolve and pin a single target tenant below.
#>
if ($TenantId) {
    $targetTenantId = $TenantId
}
else {
    $availableTenants = @(Get-AzTenant)
    if ($availableTenants.Count -eq 0) {
        throw "No Azure AD tenants are visible for the signed-in account. Sign in with an account that has access to the target tenant."
    }
    elseif ($availableTenants.Count -eq 1) {
        $targetTenantId = $availableTenants[0].Id
    }
    else {
        Write-Host "`nMultiple tenants are available for this account:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $availableTenants.Count; $i++) {
            $t = $availableTenants[$i]
            Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $t.Name, $t.Id)
        }
        $selectionIndex = 0
        while ($selectionIndex -lt 1 -or $selectionIndex -gt $availableTenants.Count) {
            $rawSelection = Read-Host "`nEnter the number of the tenant to report on"
            $selectionIndex = 0
            [void][int]::TryParse($rawSelection, [ref]$selectionIndex)
        }
        $targetTenantId = $availableTenants[$selectionIndex - 1].Id
    }
}

$currentContext = Get-AzContext
if (-not $currentContext -or $currentContext.Tenant.Id -ne $targetTenantId) {
    Write-Host "Switching to tenant $targetTenantId ..." -ForegroundColor Yellow
    Connect-AzAccountWithGuidance -Tenant $targetTenantId -Force
}
Write-Host "Using tenant: $targetTenantId" -ForegroundColor Cyan

# Resolve the reporting window = last full calendar month (UTC), unless -ReportMonth given.
if ($ReportMonth) {
    $reportYear = [int]$ReportMonth.Split('-')[0]
    $reportMonthNum = [int]$ReportMonth.Split('-')[1]
    $StartDate = (Get-Date -Year $reportYear -Month $reportMonthNum -Day 1 -Hour 0 -Minute 0 -Second 0).ToUniversalTime()
}
else {
    $firstOfThisMonth = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
    $StartDate = $firstOfThisMonth.AddMonths(-1).ToUniversalTime()
}
$EndDate = $StartDate.AddMonths(1)
$TimeGrain = New-TimeSpan -Hours $TimeGrainHours

Write-Host "Reporting window: $($StartDate.ToString('u')) -> $($EndDate.ToString('u')) (UTC)" -ForegroundColor Cyan

# Resolve subscriptions to scan (explicitly pinned to the selected tenant)
if ($SubscriptionId) {
    $subs = $SubscriptionId | ForEach-Object { Get-AzSubscription -SubscriptionId $_ -TenantId $targetTenantId }
}
else {
    $subs = Get-AzSubscription -TenantId $targetTenantId | Where-Object { $_.State -eq 'Enabled' }
}
Write-Host "Scanning $($subs.Count) subscription(s)." -ForegroundColor Cyan

# Result containers
$vmRows = [System.Collections.Generic.List[object]]::new()
$apimRows = [System.Collections.Generic.List[object]]::new()
$aspRows = [System.Collections.Generic.List[object]]::new()
$sentinelRows = [System.Collections.Generic.List[object]]::new()
$fileRows = [System.Collections.Generic.List[object]]::new()
$sqlDbRows = [System.Collections.Generic.List[object]]::new()
$sqlMiRows = [System.Collections.Generic.List[object]]::new()
$cogRows = [System.Collections.Generic.List[object]]::new()
$diskRows = [System.Collections.Generic.List[object]]::new()
$errorRows = [System.Collections.Generic.List[object]]::new()

# ---------------------------------------------------------------------------
# 1. Helpers
# ---------------------------------------------------------------------------

function Add-ErrorRow {
    param($Subscription, $ResourceType, $ResourceName, $Message)
    $errorRows.Add([pscustomobject]@{
            Subscription = $Subscription
            ResourceType = $ResourceType
            ResourceName = $ResourceName
            Error        = $Message
        })
}

<#
    Queries one metric at hourly grain for the whole reporting window and returns the
    month-level Average and Maximum per the technique described in the header comment.

    NOTE: Get-AzMetric's -AggregationType parameter only accepts a single scalar aggregation value
    in this Az.Monitor module version (passing an array throws "Cannot convert 'System.String[]'
    to the type '...Nullable`1[AggregationType]'"), so each requested aggregation type is fetched
    with its own call and the results are merged here.
#>
function Get-MonthlyMetricStat {
    param(
        [Parameter(Mandatory)] [string] $ResourceId,
        [Parameter(Mandatory)] [string] $MetricName,
        [string] $MetricNamespace,
        [string[]] $AggregationType = @('Average', 'Maximum')
    )

    $result = [pscustomobject]@{ Avg = $null; Max = $null; Min = $null }

    foreach ($agg in $AggregationType) {
        $params = @{
            ResourceId      = $ResourceId
            MetricName      = $MetricName
            TimeGrain       = $TimeGrain
            StartTime       = $StartDate
            EndTime         = $EndDate
            AggregationType = $agg
            WarningAction   = 'SilentlyContinue'
            ErrorAction     = 'Stop'
        }
        if ($MetricNamespace) { $params['MetricNamespace'] = $MetricNamespace }

        $metric = Get-AzMetric @params
        $values = $metric.Data | Where-Object { $null -ne $_.$agg } | Select-Object -ExpandProperty $agg

        if ($values) {
            switch ($agg) {
                'Average' { $result.Avg = [math]::Round(($values | Measure-Object -Average).Average, 2) }
                'Maximum' { $result.Max = [math]::Round(($values | Measure-Object -Maximum).Maximum, 2) }
                'Minimum' { $result.Min = [math]::Round(($values | Measure-Object -Minimum).Minimum, 2) }
            }
        }
    }

    return $result
}

# ---------------------------------------------------------------------------
# 2. Managed disk retail-price lookup (for orphaned disks section)
# ---------------------------------------------------------------------------

# Azure managed disk size -> tier breakpoints (GiB), published on the Managed Disks pricing page.
# Same breakpoints are reused across Premium SSD (P-series), Standard SSD (E-series) and
# Standard HDD (S-series); Azure always bills the smallest tier that is >= the provisioned size.
$script:DiskTierBreakpoints = [ordered]@{
    4 = 1; 8 = 2; 16 = 3; 32 = 4; 64 = 6; 128 = 10; 256 = 15; 512 = 20
    1024 = 30; 2048 = 40; 4096 = 50; 8192 = 60; 16384 = 70; 32767 = 80
}

function Get-DiskTierNumber {
    param([Parameter(Mandatory)] [double] $SizeGB)
    foreach ($breakpoint in $script:DiskTierBreakpoints.Keys) {
        if ($SizeGB -le $breakpoint) { return $script:DiskTierBreakpoints[$breakpoint] }
    }
    return $script:DiskTierBreakpoints[32767]
}

$script:PriceCache = @{}

<#
    Returns an estimated USD monthly cost for a managed disk using the public, unauthenticated
    Azure Retail Prices API (https://prices.azure.com/api/retail/prices).
    For Premium SSD v2 / Ultra Disk (billed per-GiB/hour capacity plus separate IOPS & throughput
    meters) only the capacity component is estimated; this is called out in the Notes column.
#>
function Get-DiskMonthlyPrice {
    param(
        [Parameter(Mandatory)] [string] $SkuName,     # e.g. Premium_LRS, StandardSSD_ZRS, Standard_LRS, PremiumV2_LRS, UltraSSD_LRS
        [Parameter(Mandatory)] [double] $SizeGB,
        [Parameter(Mandatory)] [string] $Region        # ARM region name, e.g. eastus
    )

    $redundancy = if ($SkuName -match 'ZRS') { 'ZRS' } else { 'LRS' }
    $notes = $null
    $productName = $null
    $skuLabel = $null

    switch -Regex ($SkuName) {
        '^Premium_LRS$|^Premium_ZRS$' {
            $productName = 'Premium SSD Managed Disks'
            $skuLabel = "P$(Get-DiskTierNumber -SizeGB $SizeGB) $redundancy"
        }
        '^StandardSSD_LRS$|^StandardSSD_ZRS$' {
            $productName = 'Standard SSD Managed Disks'
            $skuLabel = "E$(Get-DiskTierNumber -SizeGB $SizeGB) $redundancy"
        }
        '^Standard_LRS$' {
            $productName = 'Standard HDD Managed Disks'
            $skuLabel = "S$(Get-DiskTierNumber -SizeGB $SizeGB) $redundancy"
        }
        '^PremiumV2_LRS$' {
            $productName = 'Azure Premium SSD v2'
            $skuLabel = 'Premium LRS'
            $notes = 'Premium SSD v2 is billed per-GiB/hour capacity plus separate IOPS/throughput meters; cost shown is the capacity component only.'
        }
        '^UltraSSD_LRS$' {
            $productName = 'Ultra Disks'
            $skuLabel = 'Ultra LRS'
            $notes = 'Ultra Disk is billed per-GiB/hour capacity plus separate IOPS/throughput meters; cost shown is the capacity component only (approximate, uses Ultra LRS base rate).'
        }
        default {
            return [pscustomobject]@{ MonthlyCost = $null; Currency = $null; Notes = "Unrecognized SKU '$SkuName' - no price lookup performed." }
        }
    }

    $cacheKey = "$productName|$skuLabel|$Region"
    if ($script:PriceCache.ContainsKey($cacheKey)) {
        $cached = $script:PriceCache[$cacheKey]
        return [pscustomobject]@{ MonthlyCost = $cached.MonthlyCost; Currency = $cached.Currency; Notes = $notes }
    }

    try {
        $filter = "serviceName eq 'Storage' and armRegionName eq '$Region' and productName eq '$productName' and priceType eq 'Consumption'"
        $uri = "https://prices.azure.com/api/retail/prices?`$filter=$([uri]::EscapeDataString($filter))"
        $resp = Invoke-RestMethod -Uri $uri -Method Get

        if ($productName -in @('Azure Premium SSD v2', 'Ultra Disks')) {
            # Per-GiB/hour capacity meter (exclude Mount/Operations/Snapshot/IOPS/Throughput meters)
            $item = $resp.Items | Where-Object {
                $_.skuName -eq $skuLabel -and $_.meterName -match 'Provisioned Capacity' -and $_.unitOfMeasure -match 'GiB/Hour|GB/Hour'
            } | Select-Object -First 1
            if ($item) {
                $monthlyCost = [math]::Round($item.retailPrice * $SizeGB * 730, 2) # 730 average hours/month
                $currency = $item.currencyCode
            }
        }
        else {
            # Flat monthly per-tier meter, excluding Mount (hourly multi-attach fee) and Operations (transaction) meters
            $item = $resp.Items | Where-Object {
                $_.skuName -eq $skuLabel -and $_.meterName -notmatch 'Mount|Operations|Snapshot' -and $_.unitOfMeasure -match '1/Month'
            } | Select-Object -First 1
            if ($item) {
                $monthlyCost = [math]::Round([double]$item.retailPrice, 2)
                $currency = $item.currencyCode
            }
        }

        if (-not $item) {
            $script:PriceCache[$cacheKey] = @{ MonthlyCost = $null; Currency = $null }
            return [pscustomobject]@{ MonthlyCost = $null; Currency = $null; Notes = "No matching retail price meter found for SKU '$skuLabel' in region '$Region'." }
        }

        $script:PriceCache[$cacheKey] = @{ MonthlyCost = $monthlyCost; Currency = $currency }
        return [pscustomobject]@{ MonthlyCost = $monthlyCost; Currency = $currency; Notes = $notes }
    }
    catch {
        return [pscustomobject]@{ MonthlyCost = $null; Currency = $null; Notes = "Retail Prices API lookup failed: $($_.Exception.Message)" }
    }
}

# ---------------------------------------------------------------------------
# 3. Per-subscription collection
# ---------------------------------------------------------------------------

foreach ($sub in $subs) {

    Write-Host "`n=== Subscription: $($sub.Name) ($($sub.Id)) ===" -ForegroundColor Green
    Set-AzContext -Subscription $sub.Id -Tenant $targetTenantId | Out-Null

    # --- 3.1 Virtual Machines: Percentage CPU (Avg/Max) + Max Memory Used % ---------------------
    Write-Host "  Collecting VM CPU and memory utilization..."
    $vms = Get-AzResource -ResourceType 'Microsoft.Compute/virtualMachines'
    foreach ($vm in $vms) {
        try {
            $stat = Get-MonthlyMetricStat -ResourceId $vm.ResourceId -MetricName 'Percentage CPU'

            # 'Available Memory Percentage' is a guest-OS metric: it only reports data if the Azure
            # Monitor Agent (with a Data Collection Rule collecting it) is installed on the VM, so
            # it will be blank for VMs without that agent - this is expected, not a script error.
            # Used% isn't reported directly, so Max Memory Used % = 100 - the lowest Available%
            # seen during the month (the point of greatest memory pressure).
            $memNote = $null
            $maxMemUsedPercent = $null
            try {
                $memStat = Get-MonthlyMetricStat -ResourceId $vm.ResourceId -MetricName 'Available Memory Percentage' -AggregationType @('Minimum')
                if ($null -ne $memStat.Min) {
                    $maxMemUsedPercent = [math]::Round(100 - $memStat.Min, 2)
                }
                else {
                    $memNote = 'No memory data reported - requires the Azure Monitor Agent (with a Data Collection Rule) installed on this VM.'
                }
            }
            catch { $memNote = "Memory metric query failed: $($_.Exception.Message)" }

            $vmRows.Add([pscustomobject]@{
                    Subscription         = $sub.Name
                    ResourceGroup        = $vm.ResourceGroupName
                    VMName               = $vm.Name
                    Location             = $vm.Location
                    AvgCpuPercent        = $stat.Avg
                    MaxCpuPercent        = $stat.Max
                    MaxMemoryUsedPercent = $maxMemUsedPercent
                    Notes                = $memNote
                })
        }
        catch { Add-ErrorRow $sub.Name 'VM' $vm.Name $_.Exception.Message }
    }

    # --- 3.2 API Management: units + Avg Gateway CPU % ------------------------------------------
    Write-Host "  Collecting API Management utilization..."
    $apims = Get-AzResource -ResourceType 'Microsoft.ApiManagement/service'
    foreach ($apim in $apims) {
        try {
            $full = Get-AzResource -ResourceId $apim.ResourceId
            $capacity = $full.Sku.Capacity
            $skuName = $full.Sku.Name

            # 'Capacity' metric applies to classic (v1) SKUs; 'CpuPercent_Gateway' applies to SKUv2 (Basicv2/Standardv2/Premiumv2)
            # NOTE: 'Capacity's Maximum aggregation is documented to show 0 for any SKU other than
            # Premium, so we use Average instead, which is populated correctly across all SKUs.
            $cpuStat = $null
            try { $cpuStat = Get-MonthlyMetricStat -ResourceId $apim.ResourceId -MetricName 'Capacity' -AggregationType @('Average') }
            catch { Write-Verbose "APIM '$($apim.Name)': 'Capacity' metric not available ($($_.Exception.Message)); trying 'CpuPercent_Gateway'." }
            if (-not $cpuStat -or ($null -eq $cpuStat.Avg)) {
                try { $cpuStat = Get-MonthlyMetricStat -ResourceId $apim.ResourceId -MetricName 'CpuPercent_Gateway' -AggregationType @('Average') }
                catch { Write-Verbose "APIM '$($apim.Name)': 'CpuPercent_Gateway' metric not available ($($_.Exception.Message))." }
            }

            $apimRows.Add([pscustomobject]@{
                    Subscription                   = $sub.Name
                    ResourceGroup                  = $apim.ResourceGroupName
                    APIMName                       = $apim.Name
                    Location                       = $apim.Location
                    Sku                            = $skuName
                    Units                          = $capacity
                    AvgGatewayCpuOrCapacityPercent = $cpuStat.Avg
                })
        }
        catch { Add-ErrorRow $sub.Name 'APIM' $apim.Name $_.Exception.Message }
    }

    # --- 3.3 App Service Plans: Avg & Max CPU % --------------------------------------------------
    Write-Host "  Collecting App Service Plan CPU utilization..."
    $plans = Get-AzResource -ResourceType 'Microsoft.Web/serverfarms'
    foreach ($plan in $plans) {
        try {
            $full = Get-AzResource -ResourceId $plan.ResourceId
            $stat = Get-MonthlyMetricStat -ResourceId $plan.ResourceId -MetricName 'CpuPercentage'
            $aspRows.Add([pscustomobject]@{
                    Subscription  = $sub.Name
                    ResourceGroup = $plan.ResourceGroupName
                    PlanName      = $plan.Name
                    Location      = $plan.Location
                    Sku           = $full.Sku.Name
                    WorkerCount   = $full.Sku.Capacity
                    AvgCpuPercent = $stat.Avg
                    MaxCpuPercent = $stat.Max
                })
        }
        catch { Add-ErrorRow $sub.Name 'AppServicePlan' $plan.Name $_.Exception.Message }
    }

    # --- 3.4 Sentinel / Log Analytics: Avg daily ingestion + commitment tier ---------------------
    Write-Host "  Collecting Sentinel ingestion and commitment tier..."
    $workspaces = Get-AzResource -ResourceType 'Microsoft.OperationalInsights/workspaces'
    foreach ($ws in $workspaces) {
        try {
            # Detect Sentinel onboarding via a direct per-workspace call to the Microsoft Sentinel
            # REST API ("Sentinel Onboarding States - List") rather than Azure Resource Graph.
            # Resource Graph's index for this extension resource type was found to be unreliable in
            # testing (queries for microsoft.securityinsights/onboardingstates returned no results
            # even in subscriptions with eligible workspaces), which silently caused onboarded
            # Sentinel workspaces to be skipped. The direct REST call hits the control plane itself,
            # so it cannot miss a workspace due to indexing lag or gaps.
            $onboardingResp = Invoke-AzRestMethod -Path "$($ws.ResourceId)/providers/Microsoft.SecurityInsights/onboardingStates?api-version=2025-09-01" -Method GET
            $isSentinel = $false
            if ($onboardingResp.StatusCode -eq 200) {
                $onboardingBody = $onboardingResp.Content | ConvertFrom-Json
                $isSentinel = [bool]($onboardingBody.value -and $onboardingBody.value.Count -gt 0)
            }
            if (-not $isSentinel) { continue }

            $full = Get-AzResource -ResourceId $ws.ResourceId
            $commitmentLevel = $full.Properties.sku.capacityReservationLevel
            $hasCommitment = [bool]$commitmentLevel

            # No native Azure Monitor platform metric reports ingestion volume; Microsoft's documented
            # approach is to query the workspace's own 'Usage' table via Log Analytics Query (KQL).
            $kql = @"
Usage
| where TimeGenerated >= datetime($($StartDate.ToString('o'))) and TimeGenerated < datetime($($EndDate.ToString('o')))
| where IsBillable == true
| summarize TotalGB = sum(Quantity) / 1000.0 by bin(TimeGenerated, 1d)
| summarize AvgDailyIngestionGB = avg(TotalGB)
"@
            $queryResult = Invoke-AzOperationalInsightsQuery -WorkspaceId $full.Properties.customerId -Query $kql
            $avgDailyGb = $queryResult.Results | Select-Object -ExpandProperty AvgDailyIngestionGB -First 1

            $sentinelRows.Add([pscustomobject]@{
                    Subscription           = $sub.Name
                    ResourceGroup          = $ws.ResourceGroupName
                    WorkspaceName          = $ws.Name
                    Location               = $ws.Location
                    AvgDailyIngestionGB    = if ($avgDailyGb) { [math]::Round([double]$avgDailyGb, 2) } else { $null }
                    DailyCommitmentInPlace = $hasCommitment
                    CommitmentTierGBPerDay = $commitmentLevel
                })
        }
        catch { Add-ErrorRow $sub.Name 'Sentinel/LogAnalytics' $ws.Name $_.Exception.Message }
    }

    # --- 3.5 Storage accounts (Premium Files, provisioned): used vs allocated per share ----------
    Write-Host "  Collecting Premium Files (provisioned) share capacity..."
    $storageAccounts = Get-AzStorageAccount
    # Kind 'FileStorage' is, per Microsoft's own documentation, exclusively used for Premium Files
    # provisioned storage accounts (it cannot host blob/Standard file services) - the Sku.Tier
    # check below is a redundant-but-explicit safeguard to ensure only Premium provisioned
    # accounts are included, excluding Standard file shares (incl. the newer Provisioned v2
    # billing model, which can also apply to Standard/HDD media in general-purpose v2 accounts).
    $premiumFileAccounts = $storageAccounts | Where-Object { $_.Kind -eq 'FileStorage' -and $_.Sku.Tier -eq 'Premium' }
    foreach ($sa in $premiumFileAccounts) {
        try {
            $fileServiceId = "$($sa.Id)/fileServices/default"

            $usedMetric = Get-AzMetric -ResourceId $fileServiceId -MetricName 'FileCapacity' `
                -TimeGrain $TimeGrain -StartTime $StartDate -EndTime $EndDate `
                -AggregationType Average -Dimension 'FileShare' -DetailedOutput `
                -WarningAction SilentlyContinue

            $quotaMetric = Get-AzMetric -ResourceId $fileServiceId -MetricName 'FileShareCapacityQuota' `
                -TimeGrain $TimeGrain -StartTime $StartDate -EndTime $EndDate `
                -AggregationType Average -Dimension 'FileShare' -DetailedOutput `
                -WarningAction SilentlyContinue

            $shareNames = @()
            if ($usedMetric.Data) { $shareNames += ($usedMetric.Data | ForEach-Object { $_.Timeseries.Metadatavalue.Value }) }
            $shareNames = $shareNames | Select-Object -Unique

            foreach ($series in $usedMetric.Data) {
                $shareName = $series.Timeseries.Metadatavalue.Value | Select-Object -First 1
                $usedBytes = ($series.Timeseries.Data.Average | Where-Object { $null -ne $_ } | Measure-Object -Maximum).Maximum

                $quotaSeries = $quotaMetric.Data | Where-Object { ($_.Timeseries.Metadatavalue.Value | Select-Object -First 1) -eq $shareName }
                $quotaBytes = ($quotaSeries.Timeseries.Data.Average | Where-Object { $null -ne $_ } | Measure-Object -Maximum).Maximum

                $fileRows.Add([pscustomobject]@{
                        Subscription      = $sub.Name
                        ResourceGroup     = $sa.ResourceGroupName
                        StorageAccount    = $sa.StorageAccountName
                        FileShare         = $shareName
                        UsedGiB           = if ($usedBytes) { [math]::Round($usedBytes / 1GB, 2) } else { $null }
                        AllocatedQuotaGiB = if ($quotaBytes) { [math]::Round($quotaBytes / 1GB, 2) } else { $null }
                        PercentUsed       = if ($usedBytes -and $quotaBytes) { [math]::Round(($usedBytes / $quotaBytes) * 100, 1) } else { $null }
                    })
            }
        }
        catch { Add-ErrorRow $sub.Name 'StorageAccount-PremiumFiles' $sa.StorageAccountName $_.Exception.Message }
    }

    # --- 3.6 Azure SQL Database: Avg/Max CPU % + Max Memory Used % -------------------------------
    Write-Host "  Collecting Azure SQL Database CPU and memory utilization..."
    $sqlServers = Get-AzResource -ResourceType 'Microsoft.Sql/servers'
    foreach ($srv in $sqlServers) {
        $dbs = Get-AzResource -ResourceType 'Microsoft.Sql/servers/databases' -ResourceGroupName $srv.ResourceGroupName |
        Where-Object { $_.Name -like "$($srv.Name)/*" -and $_.Name -notlike "*master" }
        foreach ($db in $dbs) {
            try {
                $stat = Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'cpu_percent'

                # 'sql_instance_memory_percent' is not applicable to data warehouse (Synapse DW)
                # databases, so a failure here is kept independent of the CPU result above.
                $memNote = $null
                $maxMemPercent = $null
                try {
                    $memStat = Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'sql_instance_memory_percent' -AggregationType @('Maximum')
                    $maxMemPercent = $memStat.Max
                    if ($null -eq $maxMemPercent) { $memNote = 'No memory data reported (metric does not apply to data warehouse/Synapse databases).' }
                }
                catch { $memNote = "Memory metric query failed: $($_.Exception.Message)" }

                $sqlDbRows.Add([pscustomobject]@{
                        Subscription         = $sub.Name
                        ResourceGroup        = $db.ResourceGroupName
                        ServerName           = $srv.Name
                        DatabaseName         = ($db.Name -split '/')[-1]
                        Location             = $db.Location
                        AvgCpuPercent        = $stat.Avg
                        MaxCpuPercent        = $stat.Max
                        MaxMemoryUsedPercent = $maxMemPercent
                        Notes                = $memNote
                    })
            }
            catch { Add-ErrorRow $sub.Name 'AzureSQLDatabase' $db.Name $_.Exception.Message }
        }
    }

    # --- 3.7 SQL Managed Instance: Avg/Max CPU % -------------------------------------------------
    # NOTE: Microsoft.Sql/managedInstances has no memory utilization metric in Azure Monitor
    # (confirmed against the live supported-metrics reference page - only CPU, worker, I/O, log
    # rate, storage and vCore count metrics are exposed), so MaxMemoryUsedPercent is always null
    # here with an explanatory note rather than a guessed/unsupported value.
    Write-Host "  Collecting SQL Managed Instance CPU utilization..."
    $sqlMis = Get-AzResource -ResourceType 'Microsoft.Sql/managedInstances'
    foreach ($mi in $sqlMis) {
        try {
            $stat = Get-MonthlyMetricStat -ResourceId $mi.ResourceId -MetricName 'avg_cpu_percent'
            $sqlMiRows.Add([pscustomobject]@{
                    Subscription         = $sub.Name
                    ResourceGroup        = $mi.ResourceGroupName
                    InstanceName         = $mi.Name
                    Location             = $mi.Location
                    AvgCpuPercent        = $stat.Avg
                    MaxCpuPercent        = $stat.Max
                    MaxMemoryUsedPercent = $null
                    Notes                = 'Microsoft.Sql/managedInstances exposes no memory utilization metric in Azure Monitor.'
                })
        }
        catch { Add-ErrorRow $sub.Name 'SQLManagedInstance' $mi.Name $_.Exception.Message }
    }

    # --- 3.8 Cognitive Services: Max Provisioned-managed v2 utilization % ------------------------
    Write-Host "  Collecting Cognitive Services provisioned-managed (PTU v2) utilization..."
    $cogAccounts = Get-AzResource -ResourceType 'Microsoft.CognitiveServices/accounts'
    foreach ($acct in $cogAccounts) {
        $stat = $null
        $note = $null
        try {
            $stat = Get-MonthlyMetricStat -ResourceId $acct.ResourceId -MetricName 'AzureOpenAIProvisionedManagedUtilizationV2' -AggregationType @('Maximum')
        }
        catch {
            # A BadRequest here means the account doesn't support this metric at all (e.g. it isn't
            # an Azure OpenAI resource, or has no provisioned-managed/PTU deployments) - that's an
            # expected, normal outcome for many Cognitive Services accounts, not a script bug, so we
            # record it as an informational note on the row instead of a Collection-Errors failure.
            if ($_.Exception.Message -match 'BadRequest') {
                $note = 'Metric not applicable to this account (likely not an Azure OpenAI resource, or has no provisioned-managed/PTU deployments).'
            }
            else {
                $note = "Metric query failed: $($_.Exception.Message)"
                Add-ErrorRow $sub.Name 'CognitiveServices' $acct.Name $_.Exception.Message
            }
        }
        $cogRows.Add([pscustomobject]@{
                Subscription                              = $sub.Name
                ResourceGroup                             = $acct.ResourceGroupName
                AccountName                               = $acct.Name
                Location                                  = $acct.Location
                MaxProvisionedManagedUtilizationV2Percent = $stat.Max
                Notes                                     = $note
            })
    }

    # --- 3.9 Orphaned (unattached) managed disks + price lookup ---------------------------------
    Write-Host "  Collecting orphaned (unattached) managed disks..."
    try {
        $unattached = Search-AzGraph -Subscription $sub.Id -Query @"
resources
| where type == 'microsoft.compute/disks'
| where tostring(properties.diskState) == 'Unattached'
| project name, resourceGroup, location, subscriptionId,
    skuName = tostring(sku.name),
    sizeGb = toint(properties.diskSizeGB),
    timeCreated = tostring(properties.timeCreated)
"@
        foreach ($disk in $unattached) {
            $price = Get-DiskMonthlyPrice -SkuName $disk.skuName -SizeGB $disk.sizeGb -Region $disk.location
            $diskRows.Add([pscustomobject]@{
                    Subscription   = $sub.Name
                    ResourceGroup  = $disk.resourceGroup
                    DiskName       = $disk.name
                    Location       = $disk.location
                    Sku            = $disk.skuName
                    SizeGB         = $disk.sizeGb
                    TimeCreated    = $disk.timeCreated
                    EstMonthlyCost = $price.MonthlyCost
                    Currency       = $price.Currency
                    Notes          = $price.Notes
                })
        }
    }
    catch { Add-ErrorRow $sub.Name 'OrphanedDisks' '(subscription-wide query)' $_.Exception.Message }
}

# ---------------------------------------------------------------------------
# 4. Metric reference sheet (for audit / traceability)
# ---------------------------------------------------------------------------

$metricReference = @(
    [pscustomobject]@{ ResourceType = 'Microsoft.Compute/virtualMachines'; Metric = 'Percentage CPU'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-compute-virtualmachines-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Compute/virtualMachines'; Metric = 'Available Memory Percentage (inverted to Max Memory Used %; requires Azure Monitor Agent)'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-compute-virtualmachines-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.ApiManagement/service'; Metric = 'sku.capacity (resource property)'; Kind = 'ARM resource property'; Source = 'learn.microsoft.com/powershell/module/az.apimanagement/get-azapimanagement' }
    [pscustomobject]@{ ResourceType = 'Microsoft.ApiManagement/service'; Metric = 'Capacity / CpuPercent_Gateway'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-apimanagement-service-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Web/serverfarms'; Metric = 'CpuPercentage'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-web-serverfarms-metrics' }
    [pscustomobject]@{ ResourceType = 'Log Analytics workspace (Sentinel)'; Metric = 'Microsoft.SecurityInsights/onboardingStates List (direct per-workspace REST call, not Resource Graph)'; Kind = 'REST API'; Source = 'learn.microsoft.com/rest/api/securityinsights/sentinel-onboarding-states/list' }
    [pscustomobject]@{ ResourceType = 'Log Analytics workspace (Sentinel)'; Metric = 'Usage table (KQL), no native platform metric exists'; Kind = 'Log Analytics query'; Source = 'learn.microsoft.com/azure/azure-monitor/logs/analyze-usage' }
    [pscustomobject]@{ ResourceType = 'Log Analytics workspace (Sentinel)'; Metric = 'sku.capacityReservationLevel (resource property)'; Kind = 'ARM resource property'; Source = 'learn.microsoft.com/azure/templates/microsoft.operationalinsights/workspaces' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Storage/storageAccounts/fileServices'; Metric = 'FileCapacity, FileShareCapacityQuota'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-storage-storageaccounts-fileservices-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'cpu_percent'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'sql_instance_memory_percent'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/managedInstances'; Metric = 'avg_cpu_percent'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/managedInstances'; Metric = 'No memory metric exists for this resource type (confirmed against live docs)'; Kind = 'N/A'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.CognitiveServices/accounts'; Metric = 'AzureOpenAIProvisionedManagedUtilizationV2'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-cognitiveservices-accounts-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Compute/disks'; Metric = "properties.diskState == 'Unattached'"; Kind = 'Azure Resource Graph query'; Source = 'Azure Resource Graph "resources" table' }
    [pscustomobject]@{ ResourceType = 'Managed disk pricing'; Metric = 'Retail Prices API'; Kind = 'REST API'; Source = 'learn.microsoft.com/rest/api/cost-management/retail-prices/azure-retail-prices' }
)

# ---------------------------------------------------------------------------
# 5. Export to Excel (one worksheet per resource type)
# ---------------------------------------------------------------------------

if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

Write-Host "`nExporting results to $OutputPath ..." -ForegroundColor Cyan

$sheets = [ordered]@{
    'VM-CPU'                = $vmRows
    'APIM'                  = $apimRows
    'AppServicePlan-CPU'    = $aspRows
    'Sentinel-Ingestion'    = $sentinelRows
    'StorageFilePremium'    = $fileRows
    'AzureSQL-CPU'          = $sqlDbRows
    'SQLMI-CPU'             = $sqlMiRows
    'CognitiveServices-PTU' = $cogRows
    'OrphanedDisks'         = $diskRows
    'Metric-Reference'      = $metricReference
    'Collection-Errors'     = $errorRows
}

foreach ($sheetName in $sheets.Keys) {
    $data = $sheets[$sheetName]
    if (-not $data -or $data.Count -eq 0) {
        $data = @([pscustomobject]@{ Info = 'No data collected for this resource type.' })
    }
    $data | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium2
}

Write-Host "Done. Report saved to: $(Resolve-Path $OutputPath)" -ForegroundColor Green
