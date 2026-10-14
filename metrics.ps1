#Requires -Version 7.0
<#
.SYNOPSIS
    Collects Azure resource utilization metrics (last calendar month) across one or more subscriptions
    and exports the results to a multi-tab Excel workbook.

.DESCRIPTION
    Every metric used below was looked up and confirmed against the live Microsoft Learn
    "Supported metrics" reference pages (not inferred from memory). The exact metric/property
    and its documentation source is listed in the table below and is also written out to a
    "Metric-Reference" worksheet in the generated workbook for traceability.

    | Resource type            | What we report                              | Azure Monitor metric / property (REST name)                              | Source (Microsoft Learn)                                                                                                   |
    |---------------------------|----------------------------------------------|----------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------|
    | Microsoft.Compute/virtualMachines            | Avg & Max CPU %                | `Percentage CPU`                                                          | /azure/azure-monitor/reference/supported-metrics/microsoft-compute-virtualmachines-metrics                                  |
    | Microsoft.Compute/virtualMachines            | Max Memory Used % (= 100 - min Available Memory %) | `Available Memory Percentage` (Host OS platform metric - no agent required; confirmed via /azure/virtual-machines/monitor-vm-reference) | /azure/azure-monitor/reference/supported-metrics/microsoft-compute-virtualmachines-metrics                      |
    | Microsoft.ApiManagement/service               | Scale units (capacity)         | `sku.capacity` resource property (not an Azure Monitor metric)            | /powershell/module/az.apimanagement/get-azapimanagement                                                                      |
    | Microsoft.ApiManagement/service               | Avg Gateway CPU %              | `Capacity` (classic SKUs) / `CpuPercent_Gateway` (v2 SKUs)                | /azure/azure-monitor/reference/supported-metrics/microsoft-apimanagement-service-metrics                                     |
    | Microsoft.Web/serverfarms (App Service Plan)  | Max CPU %                       | `CpuPercentage`                                                           | /azure/azure-monitor/reference/supported-metrics/microsoft-web-serverfarms-metrics                                           |
    | Microsoft Sentinel (Log Analytics workspace)  | Detect whether Sentinel is onboarded | `Microsoft.SecurityInsights/onboardingStates` List operation, queried directly per workspace | /rest/api/securityinsights/sentinel-onboarding-states/list |
    | Microsoft Sentinel (Log Analytics workspace)  | Avg daily ingestion (GB/day)   | `Usage` table KQL query (no native Azure Monitor platform metric exists)  | /azure/azure-monitor/logs/analyze-usage ; confirmed no ingestion metric exists on microsoft-operationalinsights-workspaces-metrics |
    | Microsoft Sentinel / Log Analytics workspace  | Daily commitment tier in place | `sku.capacityReservationLevel` resource property                         | /azure/templates/microsoft.operationalinsights/workspaces                                                                    |
    | Microsoft.Storage/storageAccounts/fileServices (Premium Files) | Used vs allocated per share | `FileCapacity` (used, bytes) & `FileShareCapacityQuota` (allocated, bytes), split by `FileShare` dimension | /azure/azure-monitor/reference/supported-metrics/microsoft-storage-storageaccounts-fileservices-metrics |
    | Microsoft.Sql/servers/databases (Azure SQL DB)| Edition, deployment type (Single/Elastic Pool), cores/DTUs, metering model | `sku.tier`, `sku.capacity`, `sku.family`, `properties.elasticPoolId` resource properties | /azure/templates/microsoft.sql/servers/databases |
    | Microsoft.Sql/servers/databases (Azure SQL DB)| Status (Online/Paused/Standby/etc.) and replica role (Geo/Named/Standby secondary) | `properties.status`, `properties.secondaryType` resource properties | /azure/templates/microsoft.sql/servers/databases |
    | Microsoft.Sql/servers/databases (Azure SQL DB)| Max CPU % (regular + full SQL instance)| `cpu_percent`, `sql_instance_cpu_percent`                                | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics                                     |
    | Microsoft.Sql/servers/databases (Azure SQL DB)| Max Memory Used %              | `sql_instance_memory_percent`                                             | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics                                     |
    | Microsoft.Sql/servers/databases (Azure SQL DB)| Max DTU % (DTU-based only), Max storage/workers/sessions % | `dtu_consumption_percent`, `storage_percent`, `workers_percent`, `sessions_percent` | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics |
    | Microsoft.Sql/servers/elasticPools              | Edition, cores/DTUs, per-DB min/max, Max CPU/storage/workers/sessions % | `sku.*` resource properties + `cpu_percent`, `storage_percent`, `workers_percent`, `sessions_percent` | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-elasticpools-metrics |
    | Microsoft.Sql/managedInstances (SQL MI)        | Edition, cores, storage size, state, Avg & Max CPU %, Max workers %, storage used | `sku.tier`, `properties.vCores`, `properties.state`, `avg_cpu_percent`, `avg_workers_percent`, `storage_space_used_mb`, `reserved_storage_mb` | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics |
    | Microsoft.Sql/managedInstances (SQL MI)        | Max Memory Used %              | Not available - no memory metric exists for this resource type in Azure Monitor (confirmed against the live docs) | /azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics |
    | Microsoft.CognitiveServices/accounts/deployments (PTU only) | Max Provisioned-managed v2 % over the last 30 days, per deployment | `AzureOpenAIProvisionedManagedUtilizationV2` filtered by the `ModelDeploymentName` dimension; deployments discovered/filtered to `*ProvisionedManaged*` SKUs via the `deployments` REST API | /azure/azure-monitor/reference/supported-metrics/microsoft-cognitiveservices-accounts-metrics                                |

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
    Path to the .xlsx file to create. Defaults to .\AzureUtilizationReport_<yyyyMMdd_HHmm>.xlsx in
    the current directory.

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

    [string] $OutputPath = ".\AzureUtilizationReport_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx",

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
    Update-Module -Name Az.Accounts, Az.Resources, Az.Monitor, Az.OperationalInsights, Az.Storage -Force
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
foreach ($m in @('Az.Resources', 'Az.Monitor', 'Az.OperationalInsights', 'Az.Storage', 'ImportExcel')) {
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

# Separate trailing 30-day window (ending "now", not the last-calendar-month window above) used
# specifically for Cognitive Services / AI Foundry PTU utilization, per client request.
$Last30dEndDate = (Get-Date).ToUniversalTime()
$Last30dStartDate = $Last30dEndDate.AddDays(-30)

Write-Host "Reporting window: $($StartDate.ToString('u')) -> $($EndDate.ToString('u')) (UTC)" -ForegroundColor Cyan
Write-Host "PTU utilization window (last 30 days): $($Last30dStartDate.ToString('u')) -> $($Last30dEndDate.ToString('u')) (UTC)" -ForegroundColor Cyan

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
$sqlPoolRows = [System.Collections.Generic.List[object]]::new()
$sqlMiRows = [System.Collections.Generic.List[object]]::new()
$cogRows = [System.Collections.Generic.List[object]]::new()
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
    Get-AzResource (used generically, e.g. "Get-AzResource -ResourceType 'Microsoft.Sql/servers/databases'")
    goes through the generic Microsoft.Resources resource-list API rather than the SQL resource
    provider's own list endpoint. That generic API is documented/known to come back with the
    type-specific `Properties` payload null or empty for many resource types - which is exactly why
    fields like Status, secondaryType, elasticPoolId, vCores, perDatabaseSettings, etc. were blank
    for every SQL row even though the field-mapping code itself was correct. Querying the SQL RP's
    own REST list endpoint directly (same direct-REST technique already used elsewhere in this
    script for Sentinel onboarding states and Cognitive Services deployments) guarantees the full,
    real payload is present. Returns plain PSCustomObjects shaped like Get-AzResource's output
    (ResourceId/Name/Location/ResourceGroupName/Sku/Properties) so the rest of the script is
    unaffected, except each object's Name is now just its own short name (not "parent/child").
#>
function Get-AzSqlChildResourceFull {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $ApiVersion
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $separator = if ($Path -match '\?') { '&' } else { '?' }
    $nextPath = "$Path$separator" + "api-version=$ApiVersion"
    while ($nextPath) {
        $resp = Invoke-AzRestMethod -Path $nextPath -Method GET
        if ($resp.StatusCode -ne 200) { break }
        $body = $resp.Content | ConvertFrom-Json
        foreach ($item in $body.value) {
            $results.Add([pscustomobject]@{
                    ResourceId        = $item.id
                    Name              = $item.name
                    Location          = $item.location
                    ResourceGroupName = ($item.id -split '/')[4]
                    Sku               = $item.sku
                    Properties        = $item.properties
                })
        }
        $nextPath = if ($body.nextLink) { ($body.nextLink -replace '^https://[^/]+', '') } else { $null }
    }
    return $results
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
        # OData filter used to split/filter a dimensioned metric down to a single dimension value,
        # e.g. "ModelDeploymentName eq 'my-deployment'" for the Cognitive Services PTU metric. Without
        # this, a dimensioned metric queried at the resource level is blended/ambiguous across all of
        # that dimension's values (e.g. across every model deployment on the account).
        [string] $MetricFilter,
        [string[]] $AggregationType = @('Average', 'Maximum'),
        # Allow callers to override the reporting window (e.g. a trailing 30-day window instead of
        # the script's default last-calendar-month window used everywhere else).
        [datetime] $StartTime = $StartDate,
        [datetime] $EndTime = $EndDate
    )

    $result = [pscustomobject]@{ Avg = $null; Max = $null; Min = $null }

    foreach ($agg in $AggregationType) {
        $params = @{
            ResourceId      = $ResourceId
            MetricName      = $MetricName
            TimeGrain       = $TimeGrain
            StartTime       = $StartTime
            EndTime         = $EndTime
            AggregationType = $agg
            WarningAction   = 'SilentlyContinue'
            ErrorAction     = 'Stop'
        }
        if ($MetricNamespace) { $params['MetricNamespace'] = $MetricNamespace }
        if ($MetricFilter) { $params['MetricFilter'] = $MetricFilter }

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
# 2. Per-subscription collection
# ---------------------------------------------------------------------------

foreach ($sub in $subs) {

    Write-Host "`n=== Subscription: $($sub.Name) ($($sub.Id)) ===" -ForegroundColor Green
    Set-AzContext -Subscription $sub.Id -Tenant $targetTenantId | Out-Null

    # --- 2.1 Virtual Machines: Percentage CPU (Avg/Max) + Max Memory Used % ---------------------
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

    # --- 2.2 API Management: units + Avg Gateway CPU % ------------------------------------------
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

    # --- 2.3 App Service Plans: Avg & Max CPU % --------------------------------------------------
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

    # --- 2.4 Sentinel / Log Analytics: Avg daily ingestion + commitment tier ---------------------
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

    # --- 2.5 Storage accounts (Premium Files, provisioned): used vs allocated per share ----------
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

            # With -DetailedOutput, Get-AzMetric returns the per-dimension breakdown on the metric
            # object's own .Timeseries property (one TimeSeriesElement per FileShare), each with its
            # own .Metadatavalues (the FileShare name) and .Data (the actual data points) - NOT
            # nested under .Data as might be assumed. (Verified directly against the installed
            # Az.Monitor assembly's PSMetric/TimeSeriesElement/MetadataValue type definitions.)
            foreach ($series in $usedMetric.Timeseries) {
                $shareName = $series.Metadatavalues | Select-Object -First 1 -ExpandProperty Value
                $usedBytes = ($series.Data.Average | Where-Object { $null -ne $_ } | Measure-Object -Maximum).Maximum

                $quotaSeries = $quotaMetric.Timeseries | Where-Object {
                    ($_.Metadatavalues | Select-Object -First 1 -ExpandProperty Value) -eq $shareName
                }
                $quotaBytes = ($quotaSeries.Data.Average | Where-Object { $null -ne $_ } | Measure-Object -Maximum).Maximum

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

    # --- 2.6 Azure SQL Database: deployment details, cores/DTUs, status, replica role, and the ----
    #      right Max usage metrics for the type (CPU, SQL-instance CPU, memory, DTU, storage,
    #      workers, sessions). Elastic pools are also reported separately (2.6b) since a pooled
    #      database shares the pool's capacity rather than having its own.
    Write-Host "  Collecting Azure SQL Database configuration and utilization..."
    $sqlServers = Get-AzResource -ResourceType 'Microsoft.Sql/servers'
    foreach ($srv in $sqlServers) {

        # --- 2.6b Elastic pools on this server: shared capacity + pool-level Max usage metrics ----
        $pools = Get-AzSqlChildResourceFull -Path "$($srv.ResourceId)/elasticPools" -ApiVersion '2023-08-01'
        $poolByName = @{}
        foreach ($pool in $pools) { $poolByName[$pool.Name] = $pool }

        foreach ($pool in $pools) {
            try {
                $poolCpuStat = Get-MonthlyMetricStat -ResourceId $pool.ResourceId -MetricName 'cpu_percent' -AggregationType @('Maximum')
                $poolStorageStat = $null
                try { $poolStorageStat = Get-MonthlyMetricStat -ResourceId $pool.ResourceId -MetricName 'storage_percent' -AggregationType @('Maximum') } catch {}
                $poolWorkersStat = $null
                try { $poolWorkersStat = Get-MonthlyMetricStat -ResourceId $pool.ResourceId -MetricName 'workers_percent' -AggregationType @('Maximum') } catch {}
                $poolSessionsStat = $null
                try { $poolSessionsStat = Get-MonthlyMetricStat -ResourceId $pool.ResourceId -MetricName 'sessions_percent' -AggregationType @('Maximum') } catch {}

                $sqlPoolRows.Add([pscustomobject]@{
                        Subscription       = $sub.Name
                        ResourceGroup      = $pool.ResourceGroupName
                        ServerName         = $srv.Name
                        ElasticPoolName    = $pool.Name
                        Location           = $pool.Location
                        Edition            = $pool.Sku.Tier
                        MeteringModel      = if ($pool.Sku.Family) { 'vCore' } else { 'DTU' }
                        Cores              = if ($pool.Sku.Family) { $pool.Sku.Capacity } else { $null }
                        DTUs               = if ($pool.Sku.Family) { $null } else { $pool.Sku.Capacity }
                        PerDbMinCapacity   = $pool.Properties.perDatabaseSettings.minCapacity
                        PerDbMaxCapacity   = $pool.Properties.perDatabaseSettings.maxCapacity
                        ZoneRedundant      = $pool.Properties.zoneRedundant
                        MaxCpuPercent      = $poolCpuStat.Max
                        MaxStoragePercent  = $poolStorageStat.Max
                        MaxWorkersPercent  = $poolWorkersStat.Max
                        MaxSessionsPercent = $poolSessionsStat.Max
                    })
            }
            catch { Add-ErrorRow $sub.Name 'SQLElasticPool' $pool.Name $_.Exception.Message }
        }

        $dbs = Get-AzSqlChildResourceFull -Path "$($srv.ResourceId)/databases" -ApiVersion '2023-08-01' |
        Where-Object { $_.Name -ne 'master' }
        foreach ($db in $dbs) {
            try {
                $isElasticPool = [bool]$db.Properties.elasticPoolId
                $elasticPoolName = if ($isElasticPool) { ($db.Properties.elasticPoolId -split '/')[-1] } else { $null }
                $poolInfo = if ($elasticPoolName) { $poolByName[$elasticPoolName] } else { $null }

                # A pooled database's own Sku is just "ElasticPool" with no real tier/capacity, so
                # fall back to the pool's Sku (GP/BC/HS/Standard/etc. + shared capacity) for it.
                $effectiveSku = if ($poolInfo) { $poolInfo.Sku } else { $db.Sku }
                $edition = $effectiveSku.Tier
                $deploymentType = if ($isElasticPool) { 'Elastic Pool' } else { 'Single Database' }
                $meteringModel = if ($edition -eq 'DataWarehouse') { 'DWU' } elseif ($effectiveSku.Family) { 'vCore' } else { 'DTU' }
                $cores = if ($meteringModel -eq 'vCore') { $effectiveSku.Capacity } else { $null }
                $dtus = if ($meteringModel -eq 'DTU') { $effectiveSku.Capacity } else { $null }

                # secondaryType is only populated on secondary/replica databases: Geo (geo-replication
                # / auto-failover group secondary), Named (Hyperscale named replica, usually readable),
                # or Standby (Hyperscale named replica kept purely for failover - not readable).
                $secondaryType = $db.Properties.secondaryType
                $isReplica = [bool]$secondaryType
                $isStandbyReplica = ($secondaryType -eq 'Standby')

                $stat = Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'cpu_percent'

                # Each extra metric is queried independently so that one not applying to this
                # edition/deployment model (e.g. DTU-only metrics on a vCore database, storage not
                # applying to Hyperscale/Data Warehouse) never blocks the others.
                $maxSqlInstanceCpuPercent = $null
                try { $maxSqlInstanceCpuPercent = (Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'sql_instance_cpu_percent' -AggregationType @('Maximum')).Max } catch {}

                $memNote = $null
                $maxMemPercent = $null
                try {
                    $maxMemPercent = (Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'sql_instance_memory_percent' -AggregationType @('Maximum')).Max
                    if ($null -eq $maxMemPercent) { $memNote = 'No memory data reported (metric does not apply to data warehouse/Synapse databases).' }
                }
                catch { $memNote = "Memory metric query failed: $($_.Exception.Message)" }

                $maxDtuPercent = $null
                if ($meteringModel -eq 'DTU') {
                    try { $maxDtuPercent = (Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'dtu_consumption_percent' -AggregationType @('Maximum')).Max } catch {}
                }

                $maxStoragePercent = $null
                if ($edition -notin @('Hyperscale', 'DataWarehouse')) {
                    try { $maxStoragePercent = (Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'storage_percent' -AggregationType @('Maximum')).Max } catch {}
                }

                $maxWorkersPercent = $null
                try { $maxWorkersPercent = (Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'workers_percent' -AggregationType @('Maximum')).Max } catch {}

                $maxSessionsPercent = $null
                try { $maxSessionsPercent = (Get-MonthlyMetricStat -ResourceId $db.ResourceId -MetricName 'sessions_percent' -AggregationType @('Maximum')).Max } catch {}

                $sqlDbRows.Add([pscustomobject]@{
                        Subscription             = $sub.Name
                        ResourceGroup            = $db.ResourceGroupName
                        ServerName               = $srv.Name
                        DatabaseName             = $db.Name
                        Location                 = $db.Location
                        Status                   = $db.Properties.status
                        Edition                  = $edition
                        DeploymentType           = $deploymentType
                        ElasticPoolName          = $elasticPoolName
                        MeteringModel            = $meteringModel
                        Cores                    = $cores
                        DTUs                     = $dtus
                        IsReplica                = $isReplica
                        ReplicaType              = $secondaryType
                        IsStandbyReplica         = $isStandbyReplica
                        AvgCpuPercent            = $stat.Avg
                        MaxCpuPercent            = $stat.Max
                        MaxSqlInstanceCpuPercent = $maxSqlInstanceCpuPercent
                        MaxMemoryUsedPercent     = $maxMemPercent
                        MaxDtuPercent            = $maxDtuPercent
                        MaxStoragePercent        = $maxStoragePercent
                        MaxWorkersPercent        = $maxWorkersPercent
                        MaxSessionsPercent       = $maxSessionsPercent
                        Notes                    = $memNote
                    })
            }
            catch { Add-ErrorRow $sub.Name 'AzureSQLDatabase' $db.Name $_.Exception.Message }
        }
    }

    # --- 2.7 SQL Managed Instance: cores, state, storage + Max CPU/worker/IO utilization ----------
    # NOTE: Microsoft.Sql/managedInstances has no memory utilization metric in Azure Monitor
    # (confirmed against the live supported-metrics reference page - only CPU, worker, I/O, log
    # rate, storage and vCore count metrics are exposed), so MaxMemoryUsedPercent is always null
    # here with an explanatory note rather than a guessed/unsupported value.
    Write-Host "  Collecting SQL Managed Instance configuration and utilization..."
    $sqlMis = Get-AzSqlChildResourceFull -Path "/subscriptions/$($sub.Id)/providers/Microsoft.Sql/managedInstances" -ApiVersion '2021-11-01'
    foreach ($mi in $sqlMis) {
        try {
            $stat = Get-MonthlyMetricStat -ResourceId $mi.ResourceId -MetricName 'avg_cpu_percent'

            $maxWorkersPercent = $null
            try { $maxWorkersPercent = (Get-MonthlyMetricStat -ResourceId $mi.ResourceId -MetricName 'avg_workers_percent' -AggregationType @('Maximum')).Max } catch {}

            # No single "storage used %" metric exists for MI, so it's derived from the Max of
            # storage_space_used_mb against the instance's reserved_storage_mb (MB-for-MB, both
            # queried at the same grain, so a same-hour Max/Max ratio is a safe worst-case estimate).
            $maxStorageUsedMb = $null
            try { $maxStorageUsedMb = (Get-MonthlyMetricStat -ResourceId $mi.ResourceId -MetricName 'storage_space_used_mb' -AggregationType @('Maximum')).Max } catch {}
            $maxReservedStorageMb = $null
            try { $maxReservedStorageMb = (Get-MonthlyMetricStat -ResourceId $mi.ResourceId -MetricName 'reserved_storage_mb' -AggregationType @('Maximum')).Max } catch {}
            $maxStoragePercent = if ($maxStorageUsedMb -and $maxReservedStorageMb) { [math]::Round(($maxStorageUsedMb / $maxReservedStorageMb) * 100, 1) } else { $null }

            $sqlMiRows.Add([pscustomobject]@{
                    Subscription         = $sub.Name
                    ResourceGroup        = $mi.ResourceGroupName
                    InstanceName         = $mi.Name
                    Location             = $mi.Location
                    Status               = $mi.Properties.state
                    Edition              = $mi.Sku.Tier
                    Cores                = $mi.Properties.vCores
                    StorageSizeGB        = $mi.Properties.storageSizeInGB
                    AvgCpuPercent        = $stat.Avg
                    MaxCpuPercent        = $stat.Max
                    MaxMemoryUsedPercent = $null
                    MaxWorkersPercent    = $maxWorkersPercent
                    MaxStorageUsedMB     = $maxStorageUsedMb
                    MaxReservedStorageMB = $maxReservedStorageMb
                    MaxStoragePercent    = $maxStoragePercent
                    Notes                = 'Microsoft.Sql/managedInstances exposes no memory utilization metric in Azure Monitor.'
                })
        }
        catch { Add-ErrorRow $sub.Name 'SQLManagedInstance' $mi.Name $_.Exception.Message }
    }

    # --- 2.8 Cognitive Services / AI Foundry: PTU (provisioned-managed) deployments only -----------
    # Only accounts that have at least one Provisioned-managed deployment (Standard/GlobalStandard/
    # etc. PAYG deployments are excluded entirely) are reported, one row per PTU deployment (not
    # per account), each showing Max utilization over the trailing 30 days (not the report month).
    # The metric is queried filtered/split by the `ModelDeploymentName` dimension so multi-deployment
    # accounts report each deployment's own number instead of one blended account-level figure -
    # this (plus simply not knowing which deployments were PTU) was the main reason the previous
    # version could fail to return a value for every PTU deployment.
    Write-Host "  Collecting Cognitive Services / AI Foundry PTU deployment utilization (last 30 days)..."
    $cogAccounts = Get-AzResource -ResourceType 'Microsoft.CognitiveServices/accounts'
    foreach ($acct in $cogAccounts) {
        try {
            # Enumerate the account's model deployments directly via the control-plane REST API
            # (there's no Resource Graph index for deployments, and this also avoids a dependency on
            # the separate Az.CognitiveServices module), following nextLink in case of pagination.
            $deployments = [System.Collections.Generic.List[object]]::new()
            $nextPath = "$($acct.ResourceId)/deployments?api-version=2024-10-01"
            while ($nextPath) {
                $resp = Invoke-AzRestMethod -Path $nextPath -Method GET
                if ($resp.StatusCode -ne 200) { break }
                $body = $resp.Content | ConvertFrom-Json
                if ($body.value) { foreach ($d in $body.value) { $deployments.Add($d) } }
                $nextPath = if ($body.nextLink) { ($body.nextLink -replace '^https://[^/]+', '') } else { $null }
            }

            # PTU SKUs are ProvisionedManaged, GlobalProvisionedManaged and DataZoneProvisionedManaged;
            # standard/PAYG deployments (Standard, GlobalStandard, DataZoneStandard, ...) are skipped.
            $ptuDeployments = $deployments | Where-Object { $_.sku.name -like '*ProvisionedManaged*' }
            if (-not $ptuDeployments -or $ptuDeployments.Count -eq 0) { continue }

            foreach ($dep in $ptuDeployments) {
                $maxUtilPercent = $null
                $note = $null
                try {
                    $stat = Get-MonthlyMetricStat -ResourceId $acct.ResourceId -MetricName 'AzureOpenAIProvisionedManagedUtilizationV2' -AggregationType @('Maximum') -MetricFilter "ModelDeploymentName eq '$($dep.name)'" -StartTime $Last30dStartDate -EndTime $Last30dEndDate
                    $maxUtilPercent = $stat.Max
                    if ($null -eq $maxUtilPercent) { $note = 'No utilization data reported in the last 30 days (deployment had no inference traffic in the window).' }
                }
                catch {
                    $note = "Metric query failed: $($_.Exception.Message)"
                    Add-ErrorRow $sub.Name 'CognitiveServices-PTU' "$($acct.Name)/$($dep.name)" $_.Exception.Message
                }

                $cogRows.Add([pscustomobject]@{
                        Subscription                                        = $sub.Name
                        ResourceGroup                                       = $acct.ResourceGroupName
                        AccountName                                         = $acct.Name
                        Location                                            = $acct.Location
                        DeploymentName                                      = $dep.name
                        ModelName                                           = $dep.properties.model.name
                        ModelVersion                                        = $dep.properties.model.version
                        PTUSkuName                                          = $dep.sku.name
                        PTUCapacity                                         = $dep.sku.capacity
                        MaxProvisionedManagedUtilizationV2PercentLast30Days = $maxUtilPercent
                        Notes                                               = $note
                    })
            }
        }
        catch { Add-ErrorRow $sub.Name 'CognitiveServices' $acct.Name $_.Exception.Message }
    }
}

# ---------------------------------------------------------------------------
# 3. Metric reference sheet (for audit / traceability)
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
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'sql_instance_cpu_percent'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'sql_instance_memory_percent'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'dtu_consumption_percent (DTU-based databases only)'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'storage_percent (not applicable to Hyperscale/Data Warehouse)'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'workers_percent, sessions_percent'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-databases-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/databases'; Metric = 'sku.tier, sku.capacity, sku.family, properties.status, properties.elasticPoolId, properties.secondaryType (Geo/Named/Standby)'; Kind = 'ARM resource property (direct REST list, not Get-AzResource - see note)'; Source = 'learn.microsoft.com/azure/templates/microsoft.sql/servers/databases' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/elasticPools'; Metric = 'cpu_percent, storage_percent, workers_percent, sessions_percent'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-servers-elasticpools-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/servers/elasticPools'; Metric = 'sku.tier, sku.capacity, sku.family, properties.perDatabaseSettings'; Kind = 'ARM resource property (direct REST list, not Get-AzResource - see note)'; Source = 'learn.microsoft.com/azure/templates/microsoft.sql/servers/elasticpools' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/managedInstances'; Metric = 'avg_cpu_percent, avg_workers_percent, storage_space_used_mb, reserved_storage_mb'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/managedInstances'; Metric = 'No memory metric exists for this resource type (confirmed against live docs)'; Kind = 'N/A'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-sql-managedinstances-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.Sql/managedInstances'; Metric = 'properties.vCores, properties.state, properties.storageSizeInGB, sku.tier'; Kind = 'ARM resource property (direct REST list, not Get-AzResource - see note)'; Source = 'learn.microsoft.com/azure/templates/microsoft.sql/managedinstances' }
    [pscustomobject]@{ ResourceType = 'Microsoft.CognitiveServices/accounts/deployments'; Metric = 'AzureOpenAIProvisionedManagedUtilizationV2, filtered/split by the ModelDeploymentName dimension, over a trailing 30-day window; only accounts/deployments on a *ProvisionedManaged* (PTU) SKU are reported'; Kind = 'Azure Monitor metric'; Source = 'learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-cognitiveservices-accounts-metrics' }
    [pscustomobject]@{ ResourceType = 'Microsoft.CognitiveServices/accounts/deployments'; Metric = 'deployments LIST (sku.name, sku.capacity, properties.model) via direct REST call'; Kind = 'REST API'; Source = 'learn.microsoft.com/rest/api/aiservices/accountmanagement/deployments/list' }
)

# ---------------------------------------------------------------------------
# 4. Export to Excel (one worksheet per resource type)
# ---------------------------------------------------------------------------

if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

Write-Host "`nExporting results to $OutputPath ..." -ForegroundColor Cyan

$sheets = [ordered]@{
    'VM-CPU'                = $vmRows
    'APIM'                  = $apimRows
    'AppServicePlan-CPU'    = $aspRows
    'Sentinel-Ingestion'    = $sentinelRows
    'StorageFilePremium'    = $fileRows
    'AzureSQL-Database'     = $sqlDbRows
    'AzureSQL-ElasticPool'  = $sqlPoolRows
    'SQLMI'                 = $sqlMiRows
    'CognitiveServices-PTU' = $cogRows
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
