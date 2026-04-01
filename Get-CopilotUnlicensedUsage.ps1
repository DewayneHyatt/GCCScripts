<#
.SYNOPSIS
    Searches the Unified Audit Log for M365 Copilot usage in Word, Excel, PowerPoint, and OneNote,
    then cross-references against Copilot license assignments to identify unlicensed users.

.DESCRIPTION
    Designed for GCC Moderate tenants where modern usage reports are unavailable.
    Queries the Unified Audit Log for CopilotInteraction events, resolves Copilot
    SKU assignments via Microsoft Graph, and exports a user-level CSV of impacted
    users (without a premium M365 Copilot license) with per-app last active dates
    and interaction counts - similar to the commercial M365 Apps usage report.

    The script establishes three separate interactive connections because each
    M365 service requires its own authentication token:

      1. Exchange Online (Connect-ExchangeOnline)
         Provides Get-AdminAuditLogConfig to verify that Unified Audit Log
         ingestion is enabled before attempting any searches.

      2. Security & Compliance (Connect-IPPSSession)
         Provides Search-UnifiedAuditLog, which lives in the Compliance
         PowerShell session - not the standard Exchange session.

      3. Microsoft Graph (Connect-MgGraph)
         Provides Get-MgSubscribedSku and Get-MgUser to discover Copilot
         license SKUs and build the set of licensed users for cross-reference.

    All three sign-in prompts occur back-to-back at the start of the script so
    the remaining work runs unattended. If the user's browser already has an
    active admin session with a satisfied MFA claim, some prompts may auto-complete.

.PARAMETER StartDate
    Beginning of the search window. Defaults to 90 days ago.

.PARAMETER EndDate
    End of the search window. Defaults to now.

.PARAMETER OutputPath
    Directory for the output CSV. Defaults to the current directory.

.PARAMETER PreflightOnly
    Connect, verify UAL is enabled, list discovered Copilot SKUs, then exit without searching.

.PARAMETER ChunkDays
    Number of days per UAL search chunk. Defaults to 7. Reduce if you hit throttling.

.EXAMPLE
    .\Get-CopilotUnlicensedUsage.ps1 -PreflightOnly
    # Validate connections and Copilot SKU discovery without running the full audit search.

.EXAMPLE
    .\Get-CopilotUnlicensedUsage.ps1 -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date)
    # Search the last 24 hours only (good for initial validation).

.EXAMPLE
    .\Get-CopilotUnlicensedUsage.ps1
    # Full 90-day search with default settings.
#>

[CmdletBinding()]
param(
    [datetime]$StartDate = (Get-Date).AddDays(-90),
    [datetime]$EndDate   = (Get-Date),
    [string]$OutputPath  = (Get-Location).Path,
    [switch]$PreflightOnly,
    [int]$ChunkDays = 7
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# ── Target apps ──────────────────────────────────────────────────────────────
$TargetApps = @('Word', 'Excel', 'PowerPoint', 'OneNote')

# ── Helper: normalise the app name from various AuditData fields ─────────────
function Resolve-AppName {
    param([object]$AuditObj)

    # Candidates in priority order
    $raw = $AuditObj.CopilotEventData.AppHost
    if (-not $raw) { $raw = $AuditObj.AppHost }
    if (-not $raw) { $raw = $AuditObj.Workload }
    if (-not $raw) { $raw = $AuditObj.Application }
    if (-not $raw) { return $null }

    # Normalise common values to friendly names
    switch -Regex ($raw) {
        'Word'       { return 'Word' }
        'Excel'      { return 'Excel' }
        'PowerPoint' { return 'PowerPoint' }
        'OneNote'    { return 'OneNote' }
        default      { return $raw }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 1 - Prerequisites & Connections
# ═══════════════════════════════════════════════════════════════════════════════

Write-Information "`n[Phase 1] Checking prerequisites..."

# ── Required modules ─────────────────────────────────────────────────────────
$requiredModules = @(
    @{ Name = 'ExchangeOnlineManagement'; MinVersion = '3.0.0' },
    @{ Name = 'Microsoft.Graph.Users';    MinVersion = '2.0.0' },
    @{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; MinVersion = '2.0.0' }
)

foreach ($mod in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $mod.Name |
                 Where-Object { $_.Version -ge [version]$mod.MinVersion } |
                 Sort-Object Version -Descending |
                 Select-Object -First 1

    if (-not $installed) {
        Write-Warning "Module '$($mod.Name)' >= $($mod.MinVersion) not found."
        $response = Read-Host "Install from PSGallery now? (Y/N)"
        if ($response -eq 'Y') {
            Install-Module -Name $mod.Name -Scope CurrentUser -Force -AllowClobber
            Write-Information "  Installed $($mod.Name)."
        } else {
            throw "Required module '$($mod.Name)' is missing. Exiting."
        }
    } else {
        Write-Information "  $($mod.Name) v$($installed.Version) - OK"
    }
}

# ── Connection 1 of 3: Exchange Online ───────────────────────────────────────
# Needed for Get-AdminAuditLogConfig (UAL ingestion check).
Write-Information "`nConnecting to Exchange Online (interactive - sign-in 1 of 3)..."
# GCC Moderate uses commercial endpoints by default.
# For GCC High, uncomment the next line and comment out the one after it:
#   Connect-ExchangeOnline -ExchangeEnvironmentName O365USGovGCCHigh -ShowBanner:$false
Connect-ExchangeOnline -ShowBanner:$false

# ── Verify UAL is enabled ───────────────────────────────────────────────────
$auditConfig = Get-AdminAuditLogConfig
if (-not $auditConfig.UnifiedAuditLogIngestionEnabled) {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    throw @"
Unified Audit Log ingestion is DISABLED in this tenant.
Enable it in the Compliance portal > Audit, or run:
  Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true
Then wait up to 24 hours for events to start flowing.
"@
}
Write-Information "  Unified Audit Log ingestion is enabled - OK"

# ── Connection 2 of 3: Security & Compliance ────────────────────────────────
# Needed for Search-UnifiedAuditLog, which is a Compliance cmdlet - not
# available through the standard Exchange Online session.
Write-Information "`nConnecting to Security & Compliance PowerShell (interactive - sign-in 2 of 3)..."
# GCC Moderate uses the default endpoint.
# For GCC High, uncomment the next line and comment out the one after it:
#   Connect-IPPSSession -ConnectionUri https://ps.compliance.protection.office365.us/powershell-liveid/
Connect-IPPSSession -ShowBanner:$false

# ── Connection 3 of 3: Microsoft Graph ───────────────────────────────────────
# Needed for Get-MgSubscribedSku (Copilot SKU discovery) and Get-MgUser
# (license assignment lookups). Uses a different token audience than Exchange.
Write-Information "`nConnecting to Microsoft Graph (interactive - sign-in 3 of 3)..."
# GCC Moderate uses the Global cloud (graph.microsoft.com).
# For GCC High, add: -Environment USGov
Connect-MgGraph -Scopes 'User.Read.All', 'Directory.Read.All' -NoWelcome

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 2 - Discover Copilot SKUs & Build Licensed-User Set
# ═══════════════════════════════════════════════════════════════════════════════

Write-Information "`n[Phase 2] Discovering Copilot SKUs..."

$allSkus = Get-MgSubscribedSku -All
$copilotSkus = @($allSkus | Where-Object {
    $_.SkuPartNumber -like '*Copilot*' -or
    $_.SkuPartNumber -like '*COPILOT*'
})

if ($copilotSkus.Count -eq 0) {
    Write-Warning @"
No subscribed SKU with 'Copilot' in the SkuPartNumber was found.
This could mean:
  - No M365 Copilot licenses are assigned in this tenant.
  - The SKU has a different name in GCC (check Get-MgSubscribedSku output).
Proceeding anyway - all users will be flagged as unlicensed.
"@
    $copilotSkuIds = @()
} else {
    $copilotSkuIds = @($copilotSkus.SkuId)
    foreach ($sku in $copilotSkus) {
        $consumed = $sku.ConsumedUnits
        $total    = ($sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning + $sku.PrepaidUnits.Suspended)
        Write-Information "  Found SKU: $($sku.SkuPartNumber)  (SkuId: $($sku.SkuId))  [$consumed / $total assigned]"
    }
}

if ($PreflightOnly) {
    Write-Information "`n[Preflight complete] Connections verified, UAL enabled, SKU discovery done."
    Write-Information "Re-run without -PreflightOnly to perform the full audit search."
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    return
}

Write-Information "`nBuilding licensed-user set from Microsoft Graph (this may take a moment)..."
$licensedUsers = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

# Page through all users with license info
$graphUsers = Get-MgUser -All -Property 'UserPrincipalName', 'AssignedLicenses' -ConsistencyLevel eventual -CountVariable userCount
foreach ($u in $graphUsers) {
    foreach ($lic in $u.AssignedLicenses) {
        if ($copilotSkuIds -contains $lic.SkuId) {
            [void]$licensedUsers.Add($u.UserPrincipalName)
            break
        }
    }
}
Write-Information "  Total users in tenant: $($graphUsers.Count)"
Write-Information "  Users with a Copilot license: $($licensedUsers.Count)"

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 3 - Search the Unified Audit Log for CopilotInteraction Events
# ═══════════════════════════════════════════════════════════════════════════════

Write-Information "`n[Phase 3] Searching Unified Audit Log for CopilotInteraction events..."
Write-Information "  Window: $($StartDate.ToString('yyyy-MM-dd HH:mm')) to $($EndDate.ToString('yyyy-MM-dd HH:mm'))"

$allRecords = [System.Collections.Generic.List[object]]::new()

# Break the date range into chunks to avoid the 50K-record cap and reduce throttling
$chunkStart = $StartDate
while ($chunkStart -lt $EndDate) {
    $chunkEnd = $chunkStart.AddDays($ChunkDays)
    if ($chunkEnd -gt $EndDate) { $chunkEnd = $EndDate }

    Write-Information "  Chunk: $($chunkStart.ToString('yyyy-MM-dd')) to $($chunkEnd.ToString('yyyy-MM-dd'))..."

    $sessionId = [guid]::NewGuid().ToString()
    $chunkRecordCount = 0
    $retryCount = 0
    $maxRetries = 3

    do {
        try {
            $results = Search-UnifiedAuditLog `
                -StartDate $chunkStart `
                -EndDate $chunkEnd `
                -Operations 'CopilotInteraction' `
                -SessionId $sessionId `
                -SessionCommand ReturnLargeSet `
                -ResultSize 5000

            if ($null -eq $results -or $results.Count -eq 0) {
                break
            }

            $allRecords.AddRange(@($results))
            $chunkRecordCount += $results.Count
            Write-Information "    Retrieved $($results.Count) records (chunk total: $chunkRecordCount)..."

            # If fewer than 5000 returned, we've exhausted this chunk
            if ($results.Count -lt 5000) {
                break
            }

            $retryCount = 0  # Reset on success
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Warning "Failed after $maxRetries retries on chunk $($chunkStart.ToString('yyyy-MM-dd')): $_"
                break
            }
            Write-Warning "  Throttled or transient error - retrying in 30s (attempt $retryCount/$maxRetries)..."
            Start-Sleep -Seconds 30
        }
    } while ($true)

    $chunkStart = $chunkEnd
    # Brief pause between chunks to reduce throttling risk in GCC
    Start-Sleep -Seconds 2
}

Write-Information "`n  Total CopilotInteraction records retrieved: $($allRecords.Count)"

if ($allRecords.Count -eq 0) {
    Write-Warning @"
No CopilotInteraction events were found in the audit log for the specified date range.
Possible reasons:
  1. No users have used M365 Copilot in Word, Excel, PowerPoint, or OneNote during this period.
  2. CopilotInteraction audit events may not yet be available in GCC Moderate.
  3. Audit log retention may not cover your date range (standard = 90 days; Audit Premium = 1 year).
  4. The 'CopilotInteraction' operation name may have changed - check available operations with:
       Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -RecordType CopilotInteraction -ResultSize 1
"@
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    return
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 4 - Parse, Cross-Reference, Consolidate, and Export
# ═══════════════════════════════════════════════════════════════════════════════

Write-Information "`n[Phase 4] Parsing audit records and building user-level summary..."

# Accumulate per-user, per-app data:
#   $userActivity[UPN][App] = @{ LastActive = [datetime]; Count = [int] }
$userActivity = @{}
$skippedCount = 0
$totalParsed  = 0

foreach ($record in $allRecords) {
    try {
        $audit = $record.AuditData | ConvertFrom-Json
    }
    catch {
        Write-Warning "  Skipping record with unparseable AuditData: $($record.Identity)"
        continue
    }

    $app = Resolve-AppName -AuditObj $audit
    if (-not $app -or $app -notin $TargetApps) {
        $skippedCount++
        continue
    }

    $upn = $audit.UserId
    if (-not $upn) { $upn = $record.UserIds }

    # Skip licensed users - we only care about unlicensed
    if ($licensedUsers.Contains($upn)) { continue }

    # Parse event timestamp
    $eventTime = $null
    if ($audit.CreationTime) {
        $eventTime = [datetime]::Parse($audit.CreationTime)
    } elseif ($record.CreationDate) {
        $eventTime = $record.CreationDate
    }

    $totalParsed++

    if (-not $userActivity.ContainsKey($upn)) {
        $userActivity[$upn] = @{}
    }

    if (-not $userActivity[$upn].ContainsKey($app)) {
        $userActivity[$upn][$app] = @{ LastActive = $eventTime; Count = 1 }
    } else {
        $userActivity[$upn][$app].Count++
        if ($eventTime -and $eventTime -gt $userActivity[$upn][$app].LastActive) {
            $userActivity[$upn][$app].LastActive = $eventTime
        }
    }
}

Write-Information "  Events matching target apps (unlicensed users): $totalParsed"
Write-Information "  Events skipped (non-target app or licensed user): $skippedCount"
Write-Information "  Unique impacted users: $($userActivity.Count)"

if ($userActivity.Count -eq 0) {
    Write-Information "`nNo Copilot usage by unlicensed users was found. All activity is by licensed users."
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    return
}

# ── Build user-level summary rows ────────────────────────────────────────────
$summaryRows = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($upn in $userActivity.Keys) {
    $ua = $userActivity[$upn]

    $wordLast  = if ($ua.ContainsKey('Word'))       { $ua['Word'].LastActive.ToString('yyyy-MM-dd') }       else { '' }
    $excelLast = if ($ua.ContainsKey('Excel'))      { $ua['Excel'].LastActive.ToString('yyyy-MM-dd') }      else { '' }
    $pptLast   = if ($ua.ContainsKey('PowerPoint')) { $ua['PowerPoint'].LastActive.ToString('yyyy-MM-dd') } else { '' }
    $onLast    = if ($ua.ContainsKey('OneNote'))    { $ua['OneNote'].LastActive.ToString('yyyy-MM-dd') }    else { '' }

    $wordCount  = if ($ua.ContainsKey('Word'))       { $ua['Word'].Count }       else { 0 }
    $excelCount = if ($ua.ContainsKey('Excel'))      { $ua['Excel'].Count }      else { 0 }
    $pptCount   = if ($ua.ContainsKey('PowerPoint')) { $ua['PowerPoint'].Count } else { 0 }
    $onCount    = if ($ua.ContainsKey('OneNote'))    { $ua['OneNote'].Count }    else { 0 }

    $totalCount = $wordCount + $excelCount + $pptCount + $onCount
    $appsUsed   = @($ua.Keys | Sort-Object) -join ', '

    $summaryRows.Add([PSCustomObject]@{
        User                     = $upn
        AppsUsed                 = $appsUsed
        TotalInteractions        = $totalCount
        Word_LastActive          = $wordLast
        Word_Interactions        = $wordCount
        Excel_LastActive         = $excelLast
        Excel_Interactions       = $excelCount
        PowerPoint_LastActive    = $pptLast
        PowerPoint_Interactions  = $pptCount
        OneNote_LastActive       = $onLast
        OneNote_Interactions     = $onCount
    })
}

# Highest-impact users first
$summaryRows = @($summaryRows | Sort-Object TotalInteractions -Descending)

# ── Export CSV ───────────────────────────────────────────────────────────────
$timestamp = (Get-Date -Format 'yyyyMMdd_HHmmss')
$csvFile   = Join-Path $OutputPath "CopilotImpactedUsers_$timestamp.csv"

$summaryRows | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Information "`n  CSV exported to: $csvFile"

# ── Console summary ─────────────────────────────────────────────────────────
$usersWithWord  = @($summaryRows | Where-Object { $_.Word_LastActive })
$usersWithExcel = @($summaryRows | Where-Object { $_.Excel_LastActive })
$usersWithPpt   = @($summaryRows | Where-Object { $_.PowerPoint_LastActive })
$usersWithOn    = @($summaryRows | Where-Object { $_.OneNote_LastActive })

Write-Information "`n═══════════════════════════════════════════════════"
Write-Information "  SUMMARY - Impacted Unlicensed Copilot Users"
Write-Information "═══════════════════════════════════════════════════"
Write-Information "  Date range      : $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))"
Write-Information "  Impacted users  : $($summaryRows.Count)"
Write-Information ""
Write-Information "  Users by app (a user may appear in multiple):"
Write-Information "    Word            $($usersWithWord.Count) users"
Write-Information "    Excel           $($usersWithExcel.Count) users"
Write-Information "    PowerPoint      $($usersWithPpt.Count) users"
Write-Information "    OneNote         $($usersWithOn.Count) users"
Write-Information "═══════════════════════════════════════════════════`n"

# ── Cleanup ──────────────────────────────────────────────────────────────────
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue

Write-Information "Done. Connections closed."
