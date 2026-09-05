<#
.SYNOPSIS
    Intune Housekeeper. Read-only inventory of unassigned, mis-assigned and leftover
    Windows Intune objects, plus optional owner-scoped Entra ID assignment groups.
    Produces an Excel decision tracker.

.DESCRIPTION
    Windows scope only. Delegated Graph access through your own Entra app registration.

    This script NEVER modifies, unassigns or deletes anything in your tenant. It issues
    GET requests only. A human operator reviews the tracker and makes any change by hand
    in the portal.

    All Intune calls use the Microsoft Graph beta endpoint, which Microsoft may change
    without notice.

    Windows object types covered:
      - Applications                (mobileApps, Windows types only)
      - Configuration profiles      (deviceConfigurations, configurationPolicies, groupPolicyConfigurations)
      - Compliance policies         (deviceCompliancePolicies, Windows)
      - Security baselines          (intents)
      - Remediations                (deviceHealthScripts)
      - Platform scripts            (deviceManagementScripts)
      - Entra ID groups             (owner-scoped assignment groups only, matched by
                                     -GroupNamePrefix; unused in Intune OR 0 members)

    Not collected: Autopilot deployment profiles, enrolment status page and other
    enrolment configurations, assignment filters, feature/quality/driver update
    profiles. Assignment filters are also not evaluated, so an object assigned through
    a filter that matches nothing is reported as healthy.

    Priority (order of cleanup action, not device risk):
      High     -TEST object assigned to All Devices or All Users, or a group both
               included and excluded within the same assignment intent
      Medium   unassigned or exclusion-only, nothing protecting it (the cleanup queue)
      Low      unassigned but referenced, or -TEST scoped to a group (keep/investigate)
      Watch    unassigned app created within the last -NewAppGraceMonths (parked, not
               counted as actionable)
      (blank)  assigned and healthy

    Object flags:
      Unassigned              object has no assignments at all
      OnlyExclusions          object has only exclusion targets, no inclusion
      TestNamed               display name matches -TestNameRegex
      TestNamedBroadAssign    matches -TestNameRegex AND assigned to All Devices or All Users
      DuplicateName           display name occurs more than once within its type (informational)
      IncludeExcludeOverlap   same group in include and exclude within the SAME
                              assignment intent; exclusion wins (live mistake, High).
                              Cross-intent (exclude Required / include Uninstall) is
                              a normal app pattern and is not flagged.
      SupersededByNewer       (apps) a newer version supersedes this one
      HasDependents           (apps) another app depends on this one
      RecentlyCreated         (apps) unassigned but created within the last -NewAppGraceMonths

    Group flags:
      ZeroMembers             group has 0 direct members
      NotUsedInIntune         group ID is not referenced by any Windows Intune assignment

.PARAMETER ClientId
    Client ID of your own Entra app registration (public client / native flow enabled).

.PARAMETER TenantId
    Directory (tenant) ID.

.PARAMETER OutputFolder
    Folder the workbook is written to. Created if missing.

.PARAMETER NewAppGraceMonths
    Grace window for unassigned applications. An unassigned app created within this many
    months is flagged RecentlyCreated, set to Watch and left out of the worklist and the
    actionable counts. Set it to the age of the newest version you expect to retain, or
    to 0 to flag every unassigned app regardless of creation date.

.PARAMETER TestNameRegex
    Regex identifying a test object by display name, in your own naming convention.
    Empty by default, which skips test-object detection entirely and reports it once.
    There is no built-in convention: a wrong pattern would report no test objects and
    read as a clean result. Examples: '-TEST$' for a suffix, '^TEST[-_]' for a prefix,
    '(^|[-_ ])TEST([-_ ]|$)' for either.

.PARAMETER GroupNamePrefix
    Restricts the Entra group section to your assignment group naming convention. Empty
    by default; leaving it empty reports every owned group that matches the flags, which
    is usually noise.

.PARAMETER GroupOwnerUpns
    Owner accounts whose groups are checked. Empty by default, which skips the section.

.PARAMETER HeaderColor
    Worksheet header fill colour.

.EXAMPLE
    .\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>"

    Smallest useful run. Inventories Windows Intune objects and writes the workbook to
    your Documents folder. Test-object detection and the Entra group section are both
    skipped, because neither has a naming convention supplied.

.EXAMPLE
    .\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>" `
        -TestNameRegex '-TEST$'

    Adds test-object detection for names ending in -TEST. Use single quotes: in double
    quotes PowerShell would try to expand $' as a variable. This is the run that finds
    a test policy left on All Devices, which is the only High finding that describes
    something actually reaching devices.

.EXAMPLE
    .\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>" `
        -TestNameRegex '^TEST[-_]'

    The same check for a prefix convention: TEST-Wifi or TEST_Wifi. For either position,
    use '(^|[-_ ])TEST([-_ ]|$)'.

.EXAMPLE
    .\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>" `
        -GroupOwnerUpns "alice@contoso.com","bob@contoso.com" `
        -GroupNamePrefix "<your-prefix>-"

    Adds the Entra group section. Reports groups owned by those accounts whose names
    start with the prefix and that are either empty or referenced by no Windows Intune
    assignment. Both parameters are required for the section to run: owner scoping alone
    would pull in every Teams and Microsoft 365 group the owners happen to hold.

.EXAMPLE
    .\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>" `
        -NewAppGraceMonths 0

    Flags every unassigned application regardless of creation date. Use this when you do
    not retain previous versions for rollback, or when you want to see the full picture
    once. The default of 6 parks recently created unassigned apps as Watch instead.

.EXAMPLE
    .\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>" `
        -OutputFolder "C:\Reports\Intune" -HeaderColor '#31708F'

    Writes the workbook somewhere other than Documents and changes the header fill.
    -HeaderColor takes any HTML colour string.

.EXAMPLE
    .\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>" `
        -TestNameRegex '-TEST$' `
        -GroupOwnerUpns "alice@contoso.com" -GroupNamePrefix "<your-prefix>-" `
        -NewAppGraceMonths 3 -OutputFolder "C:\Reports\Intune"

    Everything switched on, with a shorter grace window for unassigned applications.
    This is the shape of a regular review run.

.NOTES
    Required delegated permissions on the app registration, admin-consented:
      DeviceManagementApps.Read.All            applications and their assignments
      DeviceManagementConfiguration.Read.All   profiles, compliance, baselines, scripts
      Group.Read.All                           Entra group section only
      User.Read.All                            Entra group section only

    App registration: public client / native flow, redirect URI http://localhost. For
    broker (WAM) sign-in, which a tenant enforcing Conditional Access token protection
    requires, also add ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id> under Mobile
    and desktop applications. On Windows PowerShell 5.1 also add
    https://login.microsoftonline.com/common/oauth2/nativeclient.

    Broker sign-in is enabled by default on Windows in current releases of
    Microsoft.Graph.Authentication and no longer needs to be turned on in code. If a
    fresh interactive sign-in fails with AADSTS530084, update that module and check the
    broker redirect URI above. The device must also be joined or registered and
    compliant, or token protection fails regardless.

    -Scopes is deliberately not passed to Connect-MgGraph. With a custom -ClientId, MSAL
    treats requested scopes as a new authorization and triggers a consent prompt; the
    token must carry what is already consented on the app registration. The script
    verifies the granted scopes instead.

    Required modules:
      Microsoft.Graph.Authentication
      ImportExcel                     (does not require Excel to be installed)

    Windows only: broker sign-in and the header styling both depend on Windows.
    Windows PowerShell 5.1 and PowerShell 7 on Windows are both supported.

    ASCII-only file. No non-ASCII characters anywhere (no em-dashes, no smart quotes).
    Read-only: no PATCH, POST or DELETE calls are made against Graph.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true,
               HelpMessage = 'Application (client) ID of your own Entra app registration. Entra admin center > App registrations > your app > Overview.')]
    [string]   $ClientId,

    [Parameter(Mandatory = $true,
               HelpMessage = 'Directory (tenant) ID. Same Overview page as the client ID, or Entra admin center > Overview.')]
    [string]   $TenantId,

    # Completes existing directories as you type, so the folder is right before the run
    # starts rather than after sign-in.
    [ArgumentCompleter({
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        $w = [string]$wordToComplete
        $w = $w.Trim("'").Trim('"')
        if ([string]::IsNullOrWhiteSpace($w)) { $search = Join-Path (Get-Location).Path '*' }
        else                                  { $search = $w + '*' }
        foreach ($d in (Get-ChildItem -Path $search -Directory -ErrorAction SilentlyContinue)) {
            $full = [string]$d.FullName
            $text = $full
            if ($full -match '\s') { $text = "'" + $full + "'" }
            [System.Management.Automation.CompletionResult]::new($text, $d.Name, 'ProviderContainer', $full)
        }
    })]
    [string]   $OutputFolder       = "$env:USERPROFILE\Documents",

    [ValidateRange(0, 120)]
    [int]      $NewAppGraceMonths  = 6,

    # There is no default naming convention, so the completer offers the common shapes
    # already quoted correctly. A pattern like -TEST$ must be single-quoted: in double
    # quotes PowerShell tries to expand the $ sequence.
    [ArgumentCompleter({
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        $suggestions = @(
            @{ Pattern = '-TEST$';                  Tip = 'Suffix: Wifi-TEST' }
            @{ Pattern = '^TEST[-_]';               Tip = 'Prefix: TEST-Wifi or TEST_Wifi' }
            @{ Pattern = '(^|[-_ ])TEST([-_ ]|$)';  Tip = 'Either position, on a word boundary' }
            @{ Pattern = '\(test\)$';               Tip = 'Bracketed suffix: Wifi (test)' }
        )
        $w = ([string]$wordToComplete).Trim("'").Trim('"')
        foreach ($s in $suggestions) {
            if ($w -and -not $s.Pattern.StartsWith($w, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $text = "'" + $s.Pattern + "'"
            [System.Management.Automation.CompletionResult]::new($text, $s.Pattern, 'ParameterValue', $s.Tip)
        }
    })]
    [string]   $TestNameRegex      = '',

    [string]   $GroupNamePrefix    = '',
    [string]   $HeaderColor        = '#404040',
    [string[]] $GroupOwnerUpns     = @()
)

$ErrorActionPreference = 'Stop'
$script:AllReferencedGroupIds = [System.Collections.Generic.HashSet[string]]::new()

# Fail early on a malformed pattern rather than part way through the inventory. An empty
# -TestNameRegex is valid and means the operator has not supplied a naming convention,
# so test-object detection is skipped.
if ($TestNameRegex) {
    try { $null = [regex]::new($TestNameRegex) }
    catch { throw ("-TestNameRegex is not a valid regular expression: {0}" -f $_.Exception.Message) }
}

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri)
    $results = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
        }
        catch {
            Write-Warning ("Graph GET failed for {0}: {1}" -f $next, $_.Exception.Message)
            break
        }
        if ($resp.PSObject.Properties.Name -contains 'value') {
            foreach ($v in $resp.value) { $results.Add($v) }
            $next = $resp.'@odata.nextLink'
        }
        else {
            $results.Add($resp)
            $next = $null
        }
    }
    return $results
}

function Add-ReferencedGroup {
    param($GroupId)
    if ($GroupId) { [void]$script:AllReferencedGroupIds.Add([string]$GroupId) }
}

function Get-AssignmentInfo {
    param($Assignments)
    $info = [pscustomobject]@{
        Count         = 0
        AllDevices    = $false
        AllUsers      = $false
        IncludeGroups = [System.Collections.Generic.List[string]]::new()
        ExcludeGroups = [System.Collections.Generic.List[string]]::new()
        HasInclusion  = $false
        Overlap       = $false
    }
    if (-not $Assignments) { return $info }
    $list = @($Assignments)
    $info.Count = $list.Count

    # Per-intent include/exclude sets. App assignments carry an intent
    # (required/available/uninstall); other object types do not, so all their
    # assignments share the '' bucket and behave as a single set.
    $incByIntent = @{}
    $excByIntent = @{}

    foreach ($a in $list) {
        $t = [string]$a.target.'@odata.type'
        $gid = $null
        if ($a.target.PSObject.Properties.Name -contains 'groupId') { $gid = [string]$a.target.groupId }
        $intent = ''
        if ($a.PSObject.Properties.Name -contains 'intent' -and $a.intent) { $intent = [string]$a.intent }

        if     ($t -like '*allDevicesAssignmentTarget')       { $info.AllDevices = $true; $info.HasInclusion = $true }
        elseif ($t -like '*allLicensedUsersAssignmentTarget') { $info.AllUsers   = $true; $info.HasInclusion = $true }
        elseif ($t -like '*exclusionGroupAssignmentTarget')   {
            if ($gid) {
                $info.ExcludeGroups.Add($gid)
                if (-not $excByIntent.ContainsKey($intent)) { $excByIntent[$intent] = New-Object System.Collections.Generic.List[string] }
                $excByIntent[$intent].Add($gid)
            }
        }
        elseif ($t -like '*groupAssignmentTarget')            {
            if ($gid) {
                $info.IncludeGroups.Add($gid)
                if (-not $incByIntent.ContainsKey($intent)) { $incByIntent[$intent] = New-Object System.Collections.Generic.List[string] }
                $incByIntent[$intent].Add($gid)
            }
            $info.HasInclusion = $true
        }
    }

    # Overlap only counts within the SAME intent. Cross-intent reuse of a group
    # (e.g. excluded from Required, included in Uninstall) is a normal, deliberate
    # app lifecycle pattern and must not be flagged.
    foreach ($k in $incByIntent.Keys) {
        if ($info.Overlap) { break }
        if ($excByIntent.ContainsKey($k)) {
            foreach ($g in $incByIntent[$k]) {
                if ($excByIntent[$k].Contains($g)) { $info.Overlap = $true; break }
            }
        }
    }
    return $info
}

function Format-Assignment {
    param($ai)
    if ($ai.Count -eq 0) { return 'Unassigned' }
    $parts = New-Object System.Collections.Generic.List[string]
    if ($ai.AllDevices) { $parts.Add('All Devices') }
    if ($ai.AllUsers)   { $parts.Add('All Users') }
    if ($ai.IncludeGroups.Count -gt 0) { $parts.Add(("{0} include group(s)" -f $ai.IncludeGroups.Count)) }
    if ($ai.ExcludeGroups.Count -gt 0) { $parts.Add(("{0} exclude group(s)" -f $ai.ExcludeGroups.Count)) }
    if ($parts.Count -eq 0) { return 'Unassigned' }
    return ($parts -join ', ')
}

function Test-Stale {
    param($DateString, [int]$Months)
    if (-not $DateString) { return $false }
    try { return ([datetime]$DateString -lt (Get-Date).AddMonths(-$Months)) }
    catch { return $false }
}

function New-InventoryRow {
    param(
        [string]$DisplayName,
        [string]$ObjectType,
        [string]$Id,
        $Created,
        $LastModified,
        $AssignmentInfo,
        [string[]]$ExtraFlags = @(),
        [hashtable]$NameCounts
    )
    $flags = New-Object System.Collections.Generic.List[string]

    if ($AssignmentInfo.Count -eq 0)           { $flags.Add('Unassigned') }
    elseif (-not $AssignmentInfo.HasInclusion) { $flags.Add('OnlyExclusions') }

    # Test-object detection is opt-in. An empty pattern must never be passed to -match:
    # it matches every string, which would flag the entire estate as test objects.
    if ($TestNameRegex -and $DisplayName -match $TestNameRegex) {
        $flags.Add('TestNamed')
        if ($AssignmentInfo.AllDevices -or $AssignmentInfo.AllUsers) { $flags.Add('TestNamedBroadAssign') }
    }

    # DuplicateName is informational only. It annotates but does not drive risk.
    $isDuplicate = ($NameCounts -and $NameCounts.ContainsKey($DisplayName) -and $NameCounts[$DisplayName] -gt 1)
    if ($isDuplicate) { $flags.Add('DuplicateName') }

    # A group in both include and exclude within the SAME assignment intent: the
    # exclusion wins, so the group is silently not targeted. Cross-intent reuse
    # (excluded from Required, included in Uninstall) is deliberate and not flagged.
    if ($AssignmentInfo.Overlap) { $flags.Add('IncludeExcludeOverlap') }

    # Apps only. An unassigned application created within the operator-chosen grace
    # window is parked rather than queued: it is either still being worked on, or a
    # retained previous version kept for rollback. Tag it so the worklist and the
    # actionable counts exclude it. -NewAppGraceMonths 0 disables the window.
    $isApp = ($ObjectType -match '^App')
    $recentApp = ($isApp -and $AssignmentInfo.Count -eq 0 -and -not (Test-Stale $Created $NewAppGraceMonths))
    if ($recentApp) { $flags.Add('RecentlyCreated') }

    foreach ($e in $ExtraFlags) { if ($e) { $flags.Add([string]$e) } }

    $flagText = ($flags -join '; ')
    $referenced = ($flagText -match 'SupersededByNewer|HasDependents|ReferencedByRelationship')

    # Priority drives the actionable cleanup list. Informational-only flags
    # (DuplicateName) never raise it above blank on their own.
    $risk = ''
    if ($flagText -match 'TestNamedBroadAssign|IncludeExcludeOverlap') {
        $risk = 'High'
    }
    elseif (($AssignmentInfo.Count -eq 0) -or ($flagText -match 'OnlyExclusions')) {
        if ($recentApp)      { $risk = 'Watch' }
        elseif ($referenced) { $risk = 'Low' }
        else                 { $risk = 'Medium' }
    }
    elseif ($flagText -match 'TestNamed') {
        $risk = 'Low'
    }

    return [pscustomobject][ordered]@{
        DisplayName  = $DisplayName
        ObjectType   = $ObjectType
        Platform     = 'Windows'
        Id           = $Id
        Created      = $Created
        LastModified = $LastModified
        Assignment   = (Format-Assignment $AssignmentInfo)
        Priority     = $risk
        Flags        = $flagText
    }
}

function Get-NameCounts {
    param($Objects, [string]$NameProperty)
    $counts = @{}
    foreach ($o in $Objects) {
        $n = $o.$NameProperty
        if (-not $n) { $n = $o.displayName }
        if ($n) {
            $key = [string]$n
            if ($counts.ContainsKey($key)) { $counts[$key]++ } else { $counts[$key] = 1 }
        }
    }
    return $counts
}

# Object types that belong to another platform. Matched before any Windows test, so a
# type that happens to contain a Windows-ish word (macOSOfficeSuiteApp) is never
# mistaken for a Windows object.
$script:OtherPlatformPattern = 'macos|ios|android|aosp|windowsphone'

# Windows application types, matched exactly. An allow-list rather than a substring
# pattern: substring matching silently swallows any type Microsoft adds later, and the
# operator never learns it exists. Anything not on this list and not another platform is
# reported as unrecognised and left out.
$script:WindowsAppTypes = @(
    '#microsoft.graph.win32LobApp'
    '#microsoft.graph.win32CatalogApp'
    '#microsoft.graph.winGetApp'
    '#microsoft.graph.officeSuiteApp'
    '#microsoft.graph.windowsAppX'
    '#microsoft.graph.windowsUniversalAppX'
    '#microsoft.graph.windowsMobileMSI'
    '#microsoft.graph.windowsMicrosoftEdgeApp'
    '#microsoft.graph.windowsStoreApp'
    '#microsoft.graph.microsoftStoreForBusinessApp'
)

$script:UnrecognisedTypes = @{}

function Add-UnrecognisedType {
    param([string]$Area, [string]$Value)
    $key = ('{0}|{1}' -f $Area, $Value)
    if ($script:UnrecognisedTypes.ContainsKey($key)) { $script:UnrecognisedTypes[$key]++ }
    else { $script:UnrecognisedTypes[$key] = 1 }
}

function Get-PlatformClass {
    # Returns 'windows', 'other', or 'unrecognised'. Fail closed: only a positive Windows
    # match is collected, and an unrecognised value is skipped rather than guessed at.
    param(
        [string]$Value,
        [string]$WindowsPattern,
        [string[]]$WindowsTypes
    )
    $v = [string]$Value
    if (-not $v) { return 'unrecognised' }
    if ($WindowsTypes) {
        foreach ($t in $WindowsTypes) { if ($v -eq $t) { return 'windows' } }
    }
    if ($v -match $script:OtherPlatformPattern) { return 'other' }
    if ($WindowsPattern -and $v -match $WindowsPattern) { return 'windows' }
    return 'unrecognised'
}

function Select-WindowsObject {
    # Filters a raw Graph collection down to Windows objects, recording anything it does
    # not recognise. -ClassifyProperty is '@odata.type' everywhere except the settings
    # catalog, where the platform lives in 'platforms'.
    param(
        $Objects,
        [Parameter(Mandatory)][string]$Area,
        [string]$ClassifyProperty = '@odata.type',
        [string]$WindowsPattern,
        [string[]]$WindowsTypes
    )
    $kept = New-Object System.Collections.Generic.List[object]
    foreach ($o in $Objects) {
        $val = [string]$o.$ClassifyProperty
        $class = Get-PlatformClass -Value $val -WindowsPattern $WindowsPattern -WindowsTypes $WindowsTypes
        if     ($class -eq 'windows')      { $kept.Add($o) }
        elseif ($class -eq 'unrecognised') { Add-UnrecognisedType -Area $Area -Value $(if ($val) { $val } else { '(empty)' }) }
    }
    return , $kept
}

function Write-UnrecognisedTypeWarning {
    if ($script:UnrecognisedTypes.Count -eq 0) { return }
    Write-Warning 'Object types not recognised as Windows or as another platform were left out of the report:'
    foreach ($k in ($script:UnrecognisedTypes.Keys | Sort-Object)) {
        $parts = $k -split '\|', 2
        Write-Warning ("  {0}: {1} ({2} object(s))" -f $parts[0], $parts[1], $script:UnrecognisedTypes[$k])
    }
    Write-Warning 'If any of these should be in scope, open an issue with the type name so it can be added.'
}

function Add-ExpandableInventory {
    # Appends rows into the shared Collector for endpoints that support $expand=assignments.
    param(
        [ValidateNotNull()][System.Collections.Generic.List[object]]$Collector,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$TypeLabel,
        [string]$ClassifyProperty = '@odata.type',
        [string]$WindowsPattern,
        [string]$NameProperty = 'displayName',
        [switch]$AppendODataType
    )
    $raw = Invoke-GraphPaged -Uri $Uri
    if ($WindowsPattern) {
        $raw = Select-WindowsObject -Objects $raw -Area $TypeLabel `
                   -ClassifyProperty $ClassifyProperty -WindowsPattern $WindowsPattern
    }

    $counts = Get-NameCounts -Objects $raw -NameProperty $NameProperty

    foreach ($o in $raw) {
        $dn = $o.$NameProperty
        if (-not $dn) { $dn = $o.displayName }
        if (-not $dn) { $dn = '(no name)' }
        $dn = [string]$dn

        $ai = Get-AssignmentInfo $o.assignments
        foreach ($g in $ai.IncludeGroups) { Add-ReferencedGroup $g }
        foreach ($g in $ai.ExcludeGroups) { Add-ReferencedGroup $g }

        $label = $TypeLabel
        if ($AppendODataType) { $label = $TypeLabel + ' - ' + ([string]$o.'@odata.type' -replace '#microsoft.graph.', '') }

        $Collector.Add( (New-InventoryRow -DisplayName $dn -ObjectType $label -Id ([string]$o.id) `
                        -Created $o.createdDateTime -LastModified $o.lastModifiedDateTime `
                        -AssignmentInfo $ai -NameCounts $counts) )
    }
}

function Ensure-Rows {
    # An empty sheet exported with no rows loses its header row, so a single
    # placeholder is written instead. -Worklist keeps the placeholder column set
    # matching the Worklist sheet rather than the inventory sheets.
    param($Rows, [switch]$Worklist)
    if (-not $Rows -or (Get-RowCount $Rows) -eq 0) {
        if ($Worklist) {
            return , ([pscustomobject][ordered]@{
                Priority='(none found)'; DisplayName=''; ObjectType=''; Id=''; Assignment='';
                LastModified=''; Reason=''; SuggestedAction=''; Decision=''; Owner='';
                DateActioned=''; Notes=''
            })
        }
        return , ([pscustomobject][ordered]@{
            DisplayName='(none found)'; ObjectType=''; Platform=''; Id=''; Created='';
            LastModified=''; Assignment=''; Priority=''; Flags=''
        })
    }
    return , $Rows
}

function Count-Flagged {
    param($Rows)
    if ($null -eq $Rows) { return [int]0 }
    $n = 0
    foreach ($r in $Rows) { if ($r.Priority -and ([string]$r.Priority) -ne '' -and ([string]$r.Priority) -ne 'Watch') { $n++ } }
    return [int]$n
}

function Get-RowCount {
    param($Rows)
    if ($null -eq $Rows) { return [int]0 }
    if ($Rows -is [System.Collections.ICollection]) { return [int]$Rows.Count }
    $n = 0
    foreach ($r in $Rows) { $n++ }
    return [int]$n
}

# ---------------------------------------------------------------------------
# Prerequisites and connection
# ---------------------------------------------------------------------------

foreach ($m in @('Microsoft.Graph.Authentication', 'ImportExcel')) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        throw ("Required module '{0}' not found. Install with: Install-Module {0} -Scope CurrentUser" -f $m)
    }
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# Conditional Access token protection (bound tokens) is only satisfied when sign-in
# goes through the Windows broker (WAM). Current releases of
# Microsoft.Graph.Authentication enable broker sign-in by default on Windows and the
# old Set-MgGraphOption -EnableLoginByWAM switch no longer has any effect, so nothing
# is set here. If a fresh interactive sign-in fails with AADSTS530084, update
# Microsoft.Graph.Authentication and add the broker redirect URI
# ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id> to the app registration.

Write-Host 'Connecting to Microsoft Graph (delegated)...'
# Do NOT pass -Scopes here. With a custom -ClientId, MSAL treats requested scopes
# as a new authorization and triggers a consent prompt; the token must instead
# carry the permissions already consented on the app registration.
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -NoWelcome -ErrorAction Stop
$ctx = Get-MgContext
if (-not $ctx) { throw 'Failed to establish a Graph context.' }
Write-Host ("Connected as {0}" -f $ctx.Account)

# Report the scopes actually carried by the token. A cached session created before
# a consent change can persist with fewer scopes than the app registration now
# grants, which silently disables the Entra group section. Disconnect-MgGraph and
# re-run to refresh the cache if a scope is reported missing.
# Group.Read.All and User.Read.All are only needed when the Entra group section will
# actually run, so they are not reported as missing on an Intune-only run.
$grantedScopes = @()
if ($ctx.Scopes) { $grantedScopes = @($ctx.Scopes) }
$neededScopes = [System.Collections.Generic.List[string]]::new()
$neededScopes.Add('DeviceManagementApps.Read.All')
$neededScopes.Add('DeviceManagementConfiguration.Read.All')
if ($GroupOwnerUpns.Count -gt 0) {
    $neededScopes.Add('Group.Read.All')
    $neededScopes.Add('User.Read.All')
}
foreach ($needed in $neededScopes) {
    if ($grantedScopes -notcontains $needed) {
        Write-Warning ("Token does not carry '{0}'. If this is unexpected, run Disconnect-MgGraph and re-run the script to refresh the cached session." -f $needed)
    }
}

try {
    # -----------------------------------------------------------------------
    # Inventory - Windows Intune objects
    # -----------------------------------------------------------------------

    if ($TestNameRegex) {
        Write-Host ("Test-object detection using pattern: {0}" -f $TestNameRegex)
    }
    else {
        Write-Host 'Test-object detection skipped: no -TestNameRegex supplied. Objects left over from testing will not be flagged. Pass your own naming convention to enable it.'
    }

    Write-Host 'Collecting applications...'
    $appsRaw = Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$expand=assignments"
    $appsWin = Select-WindowsObject -Objects $appsRaw -Area 'Applications' -WindowsTypes $script:WindowsAppTypes
    $appNameCounts = Get-NameCounts -Objects $appsWin -NameProperty 'displayName'

    $rowsApps = New-Object System.Collections.Generic.List[object]
    foreach ($app in $appsWin) {
        $ai = Get-AssignmentInfo $app.assignments
        foreach ($g in $ai.IncludeGroups) { Add-ReferencedGroup $g }
        foreach ($g in $ai.ExcludeGroups) { Add-ReferencedGroup $g }

        # Where a publishing workflow creates Intune supersedence relationships, a
        # retained previous version can be identified directly: supersedingAppCount > 0
        # means a newer app supersedes this one, so it is intended, not a removal
        # candidate. dependentAppCount > 0 means other apps depend on this one. Both are
        # read off the app object, which is faster and more reliable than a per-app
        # relationships call. Where no supersedence exists, -NewAppGraceMonths is the
        # fallback.
        $extra = @()
        $supBy = 0; $dep = 0
        if ($app.PSObject.Properties.Name -contains 'supersedingAppCount') { $supBy = [int]$app.supersedingAppCount }
        if ($app.PSObject.Properties.Name -contains 'dependentAppCount')   { $dep   = [int]$app.dependentAppCount }
        if ($supBy -gt 0) { $extra += 'SupersededByNewer' }
        if ($dep   -gt 0) { $extra += 'HasDependents' }
        $type = 'App - ' + ([string]$app.'@odata.type' -replace '#microsoft.graph.', '')
        $rowsApps.Add( (New-InventoryRow -DisplayName ([string]$app.displayName) -ObjectType $type -Id ([string]$app.id) `
                        -Created $app.createdDateTime -LastModified $app.lastModifiedDateTime `
                        -AssignmentInfo $ai -ExtraFlags $extra -NameCounts $appNameCounts) )
    }

    $rowsConfig = New-Object System.Collections.Generic.List[object]

    Write-Host 'Collecting configuration profiles (templates)...'
    Add-ExpandableInventory -Collector $rowsConfig `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$expand=assignments" `
        -TypeLabel 'Config-Template' -AppendODataType `
        -WindowsPattern 'windows|sharedPC|editionUpgrade'

    Write-Host 'Collecting configuration profiles (settings catalog)...'
    Add-ExpandableInventory -Collector $rowsConfig `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=assignments" `
        -TypeLabel 'Config-SettingsCatalog' -NameProperty 'name' `
        -ClassifyProperty 'platforms' -WindowsPattern 'windows'

    Write-Host 'Collecting configuration profiles (ADMX)...'
    Add-ExpandableInventory -Collector $rowsConfig `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?`$expand=assignments" `
        -TypeLabel 'Config-ADMX'

    Write-Host 'Collecting compliance policies...'
    $rowsCompliance = New-Object System.Collections.Generic.List[object]
    Add-ExpandableInventory -Collector $rowsCompliance `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$expand=assignments" `
        -TypeLabel 'Compliance' `
        -WindowsPattern 'windows'

    Write-Host 'Collecting remediations (deviceHealthScripts)...'
    $rowsRemediation = New-Object System.Collections.Generic.List[object]
    Add-ExpandableInventory -Collector $rowsRemediation `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?`$expand=assignments" `
        -TypeLabel 'Remediation'

    Write-Host 'Collecting platform scripts (deviceManagementScripts)...'
    $rowsScripts = New-Object System.Collections.Generic.List[object]
    Add-ExpandableInventory -Collector $rowsScripts `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?`$expand=assignments" `
        -TypeLabel 'PlatformScript'

    Write-Host 'Collecting security baselines (intents)...'
    $intents = Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/deviceManagement/intents"
    $intentNameCounts = Get-NameCounts -Objects $intents -NameProperty 'displayName'
    $rowsBaseline = New-Object System.Collections.Generic.List[object]
    foreach ($intent in $intents) {
        $asgn = Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/deviceManagement/intents/$($intent.id)/assignments"
        $ai = Get-AssignmentInfo $asgn
        foreach ($g in $ai.IncludeGroups) { Add-ReferencedGroup $g }
        foreach ($g in $ai.ExcludeGroups) { Add-ReferencedGroup $g }
        $rowsBaseline.Add( (New-InventoryRow -DisplayName ([string]$intent.displayName) -ObjectType 'Security Baseline' -Id ([string]$intent.id) `
                            -Created $intent.createdDateTime -LastModified $intent.lastModifiedDateTime `
                            -AssignmentInfo $ai -NameCounts $intentNameCounts) )
    }

    # -----------------------------------------------------------------------
    # Inventory - owner-scoped Entra ID groups (unused in Intune OR 0 members)
    # The referenced-group set is now fully populated from every object above.
    # -----------------------------------------------------------------------

    Write-Host 'Collecting owner-scoped Entra ID groups...'
    if ($GroupOwnerUpns.Count -eq 0) {
        Write-Host '  Skipped: no -GroupOwnerUpns supplied. Pass the owner accounts whose assignment groups you want checked.'
    }
    $ownedGroups = @{}
    foreach ($upn in $GroupOwnerUpns) {
        $u = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$upn'&`$select=id,userPrincipalName"
        if (@($u).Count -eq 0) { Write-Warning ("Owner account not found: {0}" -f $upn); continue }
        $uid = [string]$u[0].id
        $groups = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$uid/ownedObjects/microsoft.graph.group?`$select=id,displayName,groupTypes,createdDateTime"
        foreach ($grp in $groups) {
            $gid = [string]$grp.id
            if (-not $ownedGroups.ContainsKey($gid)) {
                $ownedGroups[$gid] = [pscustomobject]@{ Group = $grp; Owners = @($upn) }
            }
            else {
                $ownedGroups[$gid].Owners += $upn
            }
        }
    }

    # Probe group-read permission once. Reading a group's members requires
    # Group.Read.All (or GroupMember.Read.All) on the app registration. If that
    # scope is missing the API returns 403 Forbidden and group display names come
    # back blank, so there is nothing actionable to report - skip the section.
    $groupReadOk = $true
    if ($ownedGroups.Count -gt 0) {
        $probeId = @($ownedGroups.Keys)[0]
        try {
            $null = Invoke-MgGraphRequest -Method GET -OutputType PSObject -ErrorAction Stop `
                -Uri "https://graph.microsoft.com/v1.0/groups/$probeId/members?`$top=1&`$select=id"
        }
        catch {
            if ($_.Exception.Message -match 'Forbidden|403') { $groupReadOk = $false }
        }
    }
    if (-not $groupReadOk) {
        Write-Warning 'Entra group analysis skipped: the app registration is missing Group.Read.All (delegated). Add and admin-consent that scope, then re-run to populate the EntraGroups sheet.'
    }

    $rowsGroups = New-Object System.Collections.Generic.List[object]
    foreach ($kv in $ownedGroups.GetEnumerator()) {
        if (-not $groupReadOk) { break }
        $grp = $kv.Value.Group
        $gid = [string]$grp.id
        $owners = (($kv.Value.Owners | Select-Object -Unique) -join '; ')

        # Only Intune assignment groups are in scope. Owners also hold Teams,
        # M365, and other groups that are not used for assignments, where
        # 'not used in Intune' is meaningless. Filter by naming convention before
        # any member lookup so those are never reported.
        if ($GroupNamePrefix -and ([string]$grp.displayName) -notlike ($GroupNamePrefix + '*')) { continue }

        $memberState = 'unknown'
        try {
            $mUri = "https://graph.microsoft.com/v1.0/groups/$gid/members?`$top=1&`$select=id"
            $mResp = Invoke-MgGraphRequest -Method GET -Uri $mUri -OutputType PSObject -ErrorAction Stop
            if (@($mResp.value).Count -eq 0) { $memberState = 'empty' } else { $memberState = 'has-members' }
        }
        catch {
            Write-Warning ("Member read failed for group '{0}': {1}" -f $grp.displayName, $_.Exception.Message)
        }

        $usedInIntune = $script:AllReferencedGroupIds.Contains($gid)

        $flags = @()
        if ($memberState -eq 'empty') { $flags += 'ZeroMembers' }
        if (-not $usedInIntune)       { $flags += 'NotUsedInIntune' }
        if ($flags.Count -eq 0)       { continue }

        $usedText = if ($usedInIntune) { 'Yes' } else { 'No' }
        $asg = "Members: $memberState; Used in Intune: $usedText"

        $rowsGroups.Add( [pscustomobject][ordered]@{
            DisplayName  = [string]$grp.displayName
            ObjectType   = 'Entra Group'
            Platform     = 'Entra'
            Id           = $gid
            Created      = $grp.createdDateTime
            LastModified = ''
            Assignment   = $asg
            Priority     = 'Medium'
            Flags        = ($flags -join '; ')
            GroupOwners  = $owners
        } )
    }

    # -----------------------------------------------------------------------
    # Build the summary and the operator worklist
    # -----------------------------------------------------------------------

    # The workbook has to say how it was produced. With test detection opt-in, a report
    # containing no test objects is otherwise indistinguishable from a report where the
    # check never ran. No tenant or account identifiers are recorded, so the sheet is
    # safe to share or screenshot.
    $runInfo = New-Object System.Collections.Generic.List[object]
    $runInfo.Add([pscustomobject][ordered]@{ Setting='Generated';            Value=(Get-Date -Format 'yyyy-MM-dd HH:mm') })
    $runInfo.Add([pscustomobject][ordered]@{ Setting='Graph endpoint';       Value='beta' })
    $runInfo.Add([pscustomobject][ordered]@{ Setting='Platform scope';       Value='Windows' })
    $runInfo.Add([pscustomobject][ordered]@{ Setting='TestNameRegex';        Value=$(if ($TestNameRegex) { $TestNameRegex } else { '(not set - test objects not flagged)' }) })
    $runInfo.Add([pscustomobject][ordered]@{ Setting='NewAppGraceMonths';    Value=[string]$NewAppGraceMonths })
    $runInfo.Add([pscustomobject][ordered]@{ Setting='GroupNamePrefix';      Value=$(if ($GroupNamePrefix) { $GroupNamePrefix } else { '(not set)' }) })
    $runInfo.Add([pscustomobject][ordered]@{ Setting='Group owners checked'; Value=[string]$GroupOwnerUpns.Count })

    $summary = New-Object System.Collections.Generic.List[object]
    $summary.Add([pscustomobject][ordered]@{ Category='Applications';         Total=(Get-RowCount $rowsApps);        Actionable=(Count-Flagged $rowsApps) })
    $summary.Add([pscustomobject][ordered]@{ Category='ConfigProfiles';       Total=(Get-RowCount $rowsConfig);      Actionable=(Count-Flagged $rowsConfig) })
    $summary.Add([pscustomobject][ordered]@{ Category='CompliancePolicies';   Total=(Get-RowCount $rowsCompliance);  Actionable=(Count-Flagged $rowsCompliance) })
    $summary.Add([pscustomobject][ordered]@{ Category='SecurityBaselines';    Total=(Get-RowCount $rowsBaseline);    Actionable=(Count-Flagged $rowsBaseline) })
    $summary.Add([pscustomobject][ordered]@{ Category='Remediations';         Total=(Get-RowCount $rowsRemediation); Actionable=(Count-Flagged $rowsRemediation) })
    $summary.Add([pscustomobject][ordered]@{ Category='PlatformScripts';      Total=(Get-RowCount $rowsScripts);     Actionable=(Count-Flagged $rowsScripts) })
    $summary.Add([pscustomobject][ordered]@{ Category='EntraGroups(flagged)'; Total=(Get-RowCount $rowsGroups);      Actionable=(Get-RowCount $rowsGroups) })

    $allReal = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rowsApps)        { $allReal.Add($r) }
    foreach ($r in $rowsConfig)      { $allReal.Add($r) }
    foreach ($r in $rowsCompliance)  { $allReal.Add($r) }
    foreach ($r in $rowsBaseline)    { $allReal.Add($r) }
    foreach ($r in $rowsRemediation) { $allReal.Add($r) }
    foreach ($r in $rowsScripts)     { $allReal.Add($r) }
    foreach ($r in $rowsGroups)      { $allReal.Add($r) }

    # Worklist: actionable items (High/Medium/Low), pre-sorted by priority, with a
    # plain-language Reason and a deterministic SuggestedAction for the operator.
    # Rules are evaluated top-down; first match wins.
    $bucketHigh = New-Object System.Collections.Generic.List[object]
    $bucketMed  = New-Object System.Collections.Generic.List[object]
    $bucketLow  = New-Object System.Collections.Generic.List[object]
    foreach ($r in $allReal) {
        $pri = [string]$r.Priority
        if (-not $pri -or $pri -eq '' -or $pri -eq 'Watch') { continue }

        $f = [string]$r.Flags
        $reason = 'Review manually'
        $action = 'Investigate'
        if ($f -match 'IncludeExcludeOverlap') {
            $reason = 'Same group both included and excluded in the same assignment intent; the exclusion wins and that group is silently skipped'
            $action = 'Fix assignment'
        }
        elseif ($f -match 'SupersededByNewer') {
            $reason = 'Unassigned, but a newer version supersedes it (retained rollback version)'
            $action = 'Keep'
        }
        elseif ($f -match 'HasDependents') {
            $reason = 'Unassigned, but another app depends on it'
            $action = 'Keep'
        }
        elseif ($pri -eq 'High') {
            $reason = 'TEST-named object assigned to All Devices / All Users'
            $action = 'Rename'
        }
        elseif ($f -match 'OnlyExclusions') {
            $reason = 'Only exclusion assignments; effectively not deployed'
            $action = 'Investigate'
        }
        elseif ($r.ObjectType -eq 'Entra Group') {
            $parts = @()
            if ($f -match 'ZeroMembers')     { $parts += 'no members' }
            if ($f -match 'NotUsedInIntune') { $parts += 'no Intune assignment referencing it' }
            $reason = 'Owned assignment group with ' + ($parts -join ' and ')
            $action = 'Remove'
        }
        elseif ($pri -eq 'Medium') {
            $reason = 'Unassigned; nothing references it'
            $action = 'Remove'
        }
        elseif ($f -match 'TestNamed') {
            $reason = 'TEST-named object scoped to a group; confirm the test is finished'
            $action = 'Investigate'
        }
        if ($f -match 'DuplicateName') { $reason = $reason + '; duplicate display name exists' }

        $item = [pscustomobject][ordered]@{
            Priority        = $pri
            DisplayName     = $r.DisplayName
            ObjectType      = $r.ObjectType
            Id              = $r.Id
            Assignment      = $r.Assignment
            LastModified    = $r.LastModified
            Reason          = $reason
            SuggestedAction = $action
            Decision        = ''
            Owner           = ''
            DateActioned    = ''
            Notes           = ''
        }
        if     ($pri -eq 'High')   { $bucketHigh.Add($item) }
        elseif ($pri -eq 'Medium') { $bucketMed.Add($item) }
        else                       { $bucketLow.Add($item) }
    }
    $worklist = New-Object System.Collections.Generic.List[object]
    foreach ($i in $bucketHigh) { $worklist.Add($i) }
    # Within Medium, oldest LastModified first: longest-untouched content is the most
    # confidently abandoned, so the operator starts with the safest removals.
    # Note (verified in tenant): assignment changes do NOT update lastModifiedDateTime,
    # so this orders by content abandonment only - it says nothing about when an
    # object was unassigned. Empty/unparseable dates sort last.
    $sortedMed = $bucketMed | Sort-Object -Property @{ Expression = {
        $d = $null
        if ($_.LastModified) { try { $d = [datetime]$_.LastModified } catch { } }
        if ($null -eq $d) { [datetime]::MaxValue } else { $d }
    } }
    foreach ($i in $sortedMed)  { $worklist.Add($i) }
    foreach ($i in $bucketLow)  { $worklist.Add($i) }

    # -----------------------------------------------------------------------
    # Export to Excel decision tracker
    # -----------------------------------------------------------------------

    if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmm'
    $xlsx = Join-Path $OutputFolder ("Intune-Housekeeper_{0}.xlsx" -f $stamp)
    if (Test-Path $xlsx) { Remove-Item $xlsx -Force }

    $common = @{ AutoSize = $true; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true; PassThru = $true }

    Write-Host 'Writing Excel tracker...'
    $pkg = $summary                       | Export-Excel -Path $xlsx -WorksheetName 'Summary'                 @common
    $pkg = $runInfo                       | Export-Excel -ExcelPackage $pkg -WorksheetName 'RunInfo'            @common
    $pkg = (Ensure-Rows $worklist -Worklist) | Export-Excel -ExcelPackage $pkg -WorksheetName 'Worklist'           @common
    $pkg = (Ensure-Rows $rowsApps)        | Export-Excel -ExcelPackage $pkg -WorksheetName 'Applications'       @common
    $pkg = (Ensure-Rows $rowsConfig)      | Export-Excel -ExcelPackage $pkg -WorksheetName 'ConfigProfiles'     @common
    $pkg = (Ensure-Rows $rowsCompliance)  | Export-Excel -ExcelPackage $pkg -WorksheetName 'CompliancePolicies' @common
    $pkg = (Ensure-Rows $rowsBaseline)    | Export-Excel -ExcelPackage $pkg -WorksheetName 'SecurityBaselines'  @common
    $pkg = (Ensure-Rows $rowsRemediation) | Export-Excel -ExcelPackage $pkg -WorksheetName 'Remediations'       @common
    $pkg = (Ensure-Rows $rowsScripts)     | Export-Excel -ExcelPackage $pkg -WorksheetName 'PlatformScripts'    @common
    $pkg = (Ensure-Rows $rowsGroups)      | Export-Excel -ExcelPackage $pkg -WorksheetName 'EntraGroups'        @common

    # Header styling. -HeaderColor accepts any HTML colour string.
    try {
        $brand = [System.Drawing.ColorTranslator]::FromHtml($HeaderColor)
        foreach ($ws in $pkg.Workbook.Worksheets) {
            if ($ws.Dimension) {
                $lastCol = $ws.Dimension.End.Column
                $hdr = $ws.Cells[1, 1, 1, $lastCol]
                $hdr.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $hdr.Style.Fill.BackgroundColor.SetColor($brand)
                $hdr.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                $hdr.Style.Font.Bold = $true
            }
        }
    }
    catch {
        Write-Warning ("Header styling skipped: {0}" -f $_.Exception.Message)
    }

    Close-ExcelPackage $pkg

    Write-Host ''
    Write-UnrecognisedTypeWarning
    Write-Host ("Done. Decision tracker written to: {0}" -f $xlsx)
    Write-Host ("Worklist items (High/Medium/Low): {0}" -f (Get-RowCount $worklist))
}
catch {
    Write-Host ''
    Write-Host '==================== ERROR ===================='
    Write-Host ("Type    : {0}" -f $_.Exception.GetType().FullName)
    Write-Host ("Message : {0}" -f $_.Exception.Message)
    if ($_.InvocationInfo) {
        Write-Host ("Line #  : {0}" -f $_.InvocationInfo.ScriptLineNumber)
        Write-Host ("Command : {0}" -f ([string]$_.InvocationInfo.Line).Trim())
    }
    Write-Host ("Stack   :")
    Write-Host ($_.ScriptStackTrace)
    Write-Host '==============================================='
}
finally {
    if (Get-MgContext) { Disconnect-MgGraph | Out-Null }
}
