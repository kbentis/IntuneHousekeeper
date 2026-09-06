#Requires -Version 7.0

# Intune Housekeeper. Read-only inventory of Windows Intune objects.
# The exported command is Export-IntuneHousekeeperReport; everything else here is
# internal and deliberately not exported.

$ErrorActionPreference = 'Stop'

# Module-scope run state. In a script these were per-execution; in a module they persist
# for the life of the session, so every one of them is reset at the start of each run.
# Without that, a second call in the same session double-counts and inherits the first
# run's referenced groups, unrecognised types and failure flags.
$script:AllReferencedGroupIds = [System.Collections.Generic.HashSet[string]]::new()
$script:UnrecognisedTypes     = @{}
$script:GraphReadIncomplete   = $false
$script:ConnectionOwned       = $false
$script:ObjectsExamined       = 0
$script:TestNameMatches       = 0
$script:GraphAuthVersion      = $null
$script:ConfigPathUsed        = $null

function Reset-RunState {
    $script:AllReferencedGroupIds = [System.Collections.Generic.HashSet[string]]::new()
    $script:UnrecognisedTypes     = @{}
    $script:GraphReadIncomplete   = $false
    $script:ConnectionOwned       = $false
    $script:ObjectsExamined       = 0
    $script:TestNameMatches       = 0
    $script:GraphAuthVersion      = $null
    $script:ConfigPathUsed        = $null
}


# ---------------------------------------------------------------------------
# Settings file
# ---------------------------------------------------------------------------
# Identifiers and preferences only. There is nothing else to store: the tool signs in
# through a public client flow, which has no secret. If a future change ever needs one,
# it does not belong in this file, because a settings file that holds a credential is a
# credential store with none of the protections one needs.

$script:ConfigSettingNames = @(
    'ClientId'
    'TenantId'
    'OutputFolder'
    'NewAppGraceMonths'
    'RetainedVersionMonths'
    'TestNameRegex'
    'GroupNamePrefix'
    'GroupOwnerUpns'
    'HeaderColor'
)

function Get-DefaultConfigPath {
    $base = if ($env:APPDATA) { $env:APPDATA } else { Join-Path $HOME '.config' }
    return (Join-Path (Join-Path $base 'IntuneHousekeeper') 'settings.json')
}

function Read-ConfigFile {
    # Returns a hashtable of stored settings, empty when there is no file. A key that is
    # present wins even when its value is 0 or an empty string, because both are
    # meaningful settings: 0 disables a window, '' disables test detection. Only an
    # absent key falls through to the parameter default.
    param([string]$Path)
    $result = @{}
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $result }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if (-not $raw.Trim()) { return $result }
        $json = $raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw ("Settings file '{0}' could not be read: {1}" -f $Path, $_.Exception.Message)
    }
    foreach ($name in $script:ConfigSettingNames) {
        if (($json.PSObject.Properties.Name -contains $name) -and ($null -ne $json.$name)) {
            $result[$name] = $json.$name
        }
    }
    return $result
}

function Get-IntuneHousekeeperConfig {
    <#
    .SYNOPSIS
        Reads the saved Intune Housekeeper settings.

    .DESCRIPTION
        Returns the settings currently stored, along with the path they were read from.
        Every setting is optional and only the ones that have been set are present, so an
        unset value falls through to the command's own default.

    .PARAMETER ConfigPath
        Settings file to read. Defaults to
        %APPDATA%\IntuneHousekeeper\settings.json.

    .EXAMPLE
        Get-IntuneHousekeeperConfig

        Shows what is stored, and where.
    #>
    [CmdletBinding()]
    param([string]$ConfigPath)

    if (-not $ConfigPath) { $ConfigPath = Get-DefaultConfigPath }
    $stored = Read-ConfigFile -Path $ConfigPath

    $out = [ordered]@{ Path = $ConfigPath; Exists = (Test-Path -LiteralPath $ConfigPath) }
    foreach ($name in $script:ConfigSettingNames) {
        $out[$name] = $(if ($stored.ContainsKey($name)) { $stored[$name] } else { $null })
    }
    return [pscustomobject]$out
}

function Set-IntuneHousekeeperConfig {
    <#
    .SYNOPSIS
        Saves Intune Housekeeper settings so they do not have to be typed each run.

    .DESCRIPTION
        Writes only the settings you pass, leaving anything already stored untouched.
        Values are validated on save, so a malformed regex or an out-of-range month count
        fails here rather than part way through a run three weeks later.

        Nothing secret is stored. The tool signs in through a public client flow, so the
        file holds identifiers and preferences only.

    .PARAMETER RemoveSetting
        Names of settings to delete from the file, returning them to their defaults.

    .PARAMETER ConfigPath
        Settings file to write. Defaults to
        %APPDATA%\IntuneHousekeeper\settings.json. The folder is created if needed.

    .PARAMETER PassThru
        Return the resulting settings.

    .EXAMPLE
        Set-IntuneHousekeeperConfig -ClientId "<app id>" -TenantId "<tenant id>"

        Stores the identifiers, so later runs need no parameters at all.

    .EXAMPLE
        Set-IntuneHousekeeperConfig -TestNameRegex '(^|[-_ (\[])TEST([-_ )\]]|$)' -OutputFolder 'C:\Reports\Intune'

        Adds a naming convention and an output folder to whatever is already stored.

    .EXAMPLE
        Set-IntuneHousekeeperConfig -RemoveSetting GroupOwnerUpns, GroupNamePrefix

        Stops the Entra group section running by default.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]   $ClientId,
        [string]   $TenantId,
        [string]   $OutputFolder,
        [ValidateRange(0, 120)]
        [int]      $NewAppGraceMonths,
        [ValidateRange(0, 120)]
        [int]      $RetainedVersionMonths,
        [string]   $TestNameRegex,
        [string]   $GroupNamePrefix,
        [string[]] $GroupOwnerUpns,
        [string]   $HeaderColor,
        [string[]] $RemoveSetting,
        [string]   $ConfigPath,
        [switch]   $PassThru
    )

    if (-not $ConfigPath) { $ConfigPath = Get-DefaultConfigPath }

    if ($PSBoundParameters.ContainsKey('TestNameRegex') -and $TestNameRegex) {
        try { $null = [regex]::new($TestNameRegex) }
        catch { throw ("-TestNameRegex is not a valid regular expression: {0}" -f $_.Exception.Message) }
    }
    foreach ($bad in ($RemoveSetting | Where-Object { $_ -and ($script:ConfigSettingNames -notcontains $_) })) {
        throw ("'{0}' is not a setting. Valid names: {1}" -f $bad, ($script:ConfigSettingNames -join ', '))
    }

    $stored = Read-ConfigFile -Path $ConfigPath
    foreach ($name in $script:ConfigSettingNames) {
        if ($PSBoundParameters.ContainsKey($name)) { $stored[$name] = $PSBoundParameters[$name] }
    }
    foreach ($name in $RemoveSetting) { if ($stored.ContainsKey($name)) { [void]$stored.Remove($name) } }

    $ordered = [ordered]@{}
    foreach ($name in $script:ConfigSettingNames) {
        if ($stored.ContainsKey($name)) { $ordered[$name] = $stored[$name] }
    }

    if ($PSCmdlet.ShouldProcess($ConfigPath, 'Write Intune Housekeeper settings')) {
        $folder = Split-Path -Parent $ConfigPath
        if ($folder -and -not (Test-Path -LiteralPath $folder)) {
            $null = New-Item -ItemType Directory -Path $folder -Force
        }
        ([pscustomobject]$ordered | ConvertTo-Json -Depth 4) |
            Set-Content -LiteralPath $ConfigPath -Encoding UTF8
        Write-Host ("Settings written to {0}" -f $ConfigPath)
    }

    if ($PassThru) { Get-IntuneHousekeeperConfig -ConfigPath $ConfigPath }
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
            $script:GraphReadIncomplete = $true
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

function Add-AssignmentGroupReference {
    # Records every group ID an assignment collection targets. Called on the RAW result
    # of each endpoint, before the Windows filter, so a group used only by a macOS, iOS,
    # Android or Linux object is still counted as referenced. Reporting stays Windows
    # only; referencing must not, or the group section recommends removing groups that
    # are deploying software on another platform.
    param($Assignments)
    if (-not $Assignments) { return }
    foreach ($a in $Assignments) {
        if ($a.target -and ($a.target.PSObject.Properties.Name -contains 'groupId')) {
            Add-ReferencedGroup $a.target.groupId
        }
    }
}

# Endpoints this tool does not report on, but whose assignments still reference groups.
# Confirmed necessary in a live tenant: a group targeted only by a macOS shell script was
# reported as referenced by nothing, with a suggested action against it. Reporting stays
# Windows only; the referenced-group set has to be tenant wide or the group section
# recommends action on groups that are deploying something.
$script:ReferenceOnlyEndpoints = @(
    'deviceManagement/deviceShellScripts'
    'deviceManagement/deviceCustomAttributeShellScripts'
    'deviceManagement/windowsAutopilotDeploymentProfiles'
    'deviceManagement/deviceEnrollmentConfigurations'
    'deviceManagement/windowsFeatureUpdateProfiles'
    'deviceManagement/windowsQualityUpdateProfiles'
    'deviceManagement/windowsDriverUpdateProfiles'
    'deviceAppManagement/mobileAppConfigurations'
    'deviceAppManagement/targetedManagedAppConfigurations'
    'deviceAppManagement/iosManagedAppProtections'
    'deviceAppManagement/androidManagedAppProtections'
    'deviceAppManagement/windowsManagedAppProtections'
)

function Add-ReferenceOnlyGroups {
    # Only worth the calls when the Entra group section is actually going to run. A
    # failure here marks the collection incomplete, which skips that section: a partial
    # reference set makes 'nothing references this group' unsafe to assert.
    Write-Host 'Collecting group references from object types outside the report...'
    foreach ($e in $script:ReferenceOnlyEndpoints) {
        $objects = Invoke-GraphPaged -Uri ("https://graph.microsoft.com/beta/{0}?`$expand=assignments" -f $e)
        foreach ($o in $objects) { Add-AssignmentGroupReference $o.assignments }
    }
}

function Get-AppVersion {
    # Last dotted-numeric token in a display name, as [version]. Null when there is none.
    param([string]$Name)
    if (-not $Name) { return $null }
    $m = [regex]::Matches([string]$Name, '\d+(\.\d+)+')
    if ($m.Count -eq 0) { return $null }
    try { return [version]$m[$m.Count - 1].Value } catch { return $null }
}

function Test-SupersededByName {
    # True when an unassigned application looks like a retained previous version: its
    # name carries a version, and some other app shares its base name with BOTH a higher
    # version AND a live inclusion assignment. All three conditions are required, so two
    # unassigned versions of a retired app are not demoted.
    param(
        [string]$DisplayName,
        [string]$Id,
        $Index
    )
    $thisVer = Get-AppVersion $DisplayName
    if (-not $thisVer) { return $false }
    $thisBase = Get-AppBaseName $DisplayName
    foreach ($cand in $Index) {
        if ([string]$cand.Id -eq [string]$Id) { continue }
        if (-not $cand.HasInclusion)          { continue }
        if ($cand.Base -ne $thisBase)         { continue }
        if ($cand.Version -and $cand.Version -gt $thisVer) { return $true }
    }
    return $false
}

function Get-AppBaseName {
    # Display name with dotted-numeric tokens removed, so 'App 1.2.3 (x64)' and
    # 'App 1.3.0 (x64)' share a base. Architecture and edition markers survive, because
    # they contain no dotted number, which keeps x64 and x86 packages distinct.
    param([string]$Name)
    $b = [string]$Name
    $b = $b -replace '\d+(\.\d+)+', ' '
    $b = $b -replace '\s+', ' '
    return $b.Trim()
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
    $script:ObjectsExamined++
    if ($TestNameRegex -and $DisplayName -match $TestNameRegex) {
        $script:TestNameMatches++
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
    # A name-matched retained version counts as referenced only while it is still within
    # the rollback window. Once expired it is ordinary unassigned clutter again.
    $referenced = ($flagText -match 'SupersededByNewer|HasDependents|ReferencedByRelationship')
    if (($flagText -match 'SupersededByName') -and ($flagText -notmatch 'RetainedVersionExpired')) {
        $referenced = $true
    }

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
$script:OtherPlatformPattern = 'macos|ios|android|aosp|linux|windowsphone'

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

    # Before filtering: every assignment on every platform contributes to the referenced
    # group set. Reporting is Windows only; referencing must be tenant wide.
    foreach ($o in $raw) { Add-AssignmentGroupReference $o.assignments }

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

function Get-ExportRows {
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

function Get-ActionableCount {
    param($Rows)
    if ($null -eq $Rows) { return [int]0 }
    $n = 0
    # Actionable means the operator is expected to do something: High and Medium only.
    # Low is 'keep or confirm' and Watch is parked; counting either inflates the queue.
    foreach ($r in $Rows) { if ((([string]$r.Priority) -eq 'High') -or (([string]$r.Priority) -eq 'Medium')) { $n++ } }
    return [int]$n
}

function Get-PriorityCounts {
    param($Rows)
    $c = [ordered]@{ High = 0; Medium = 0; Low = 0; Watch = 0; Healthy = 0 }
    foreach ($r in $Rows) {
        $p = [string]$r.Priority
        if     ($p -eq 'High')   { $c.High++ }
        elseif ($p -eq 'Medium') { $c.Medium++ }
        elseif ($p -eq 'Low')    { $c.Low++ }
        elseif ($p -eq 'Watch')  { $c.Watch++ }
        else                     { $c.Healthy++ }
    }
    return $c
}

function Get-RowCount {
    param($Rows)
    if ($null -eq $Rows) { return [int]0 }
    if ($Rows -is [System.Collections.ICollection]) { return [int]$Rows.Count }
    $n = 0
    foreach ($r in $Rows) { $n++ }
    return [int]$n
}


function Export-IntuneHousekeeperReport {
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
          SupersededByName        (apps) unassigned, but a newer version of the same
                                  application is assigned. Matched on display name and
                                  version, for workflows that do not create supersedence
          RetainedVersionExpired  (apps) SupersededByName, but created more than
                                  -RetainedVersionMonths ago. The rollback window has
                                  passed, so it returns to the cleanup queue
          HasDependents           (apps) another app depends on this one
          RecentlyCreated         (apps) unassigned but created within the last -NewAppGraceMonths

        Group flags:
          ZeroMembers             group has 0 direct members
          NotUsedInIntune         group ID is not referenced by any assignment found. All
                                  platforms count, and assignments are also read from object
                                  types outside the report (macOS shell and custom attribute
                                  scripts, Autopilot, enrolment configurations, update rings,
                                  app configuration and app protection policies) purely to
                                  establish references. Still verify in the portal: any
                                  object type Microsoft adds is invisible until added here

    .PARAMETER ConfigPath
        Settings file to read, written by Set-IntuneHousekeeperConfig. Defaults to
        %APPDATA%\IntuneHousekeeper\settings.json. Explicit parameters win over stored
        values, and a stored value wins over the parameter default.

    .PARAMETER ClientId
        Client ID of your own Entra app registration (public client / native flow enabled).
        Required, from this parameter or from the settings file.

    .PARAMETER TenantId
        Directory (tenant) ID. Required, from this parameter or from the settings file.

    .PARAMETER OutputFolder
        Folder the workbook is written to. Created if missing.

    .PARAMETER NewAppGraceMonths
        Grace window for unassigned applications. An unassigned app created within this many
        months is flagged RecentlyCreated, set to Watch and left out of the worklist and the
        actionable counts. Set it to the age of the newest version you expect to retain, or
        to 0 to flag every unassigned app regardless of creation date.

    .PARAMETER RetainedVersionMonths
        How long a retained previous version stays out of the cleanup queue. An unassigned
        app with a newer assigned version of the same name is kept at Low while it is newer
        than this, and returns to Medium beyond it, flagged RetainedVersionExpired. Set it
        to how long a rollback is realistically useful in your environment, or to 0 to
        disable the leniency and queue every retained copy. Applies only to name-matched
        copies: a real Intune supersedence relationship is always kept, because the newer
        app's configuration depends on it.

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
        Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>"

        Smallest useful run. Inventories Windows Intune objects and writes the workbook to
        your Documents folder. Test-object detection and the Entra group section are both
        skipped, because neither has a naming convention supplied.

    .EXAMPLE
        Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>" `
            -TestNameRegex '-TEST$'

        Adds test-object detection for names ending in -TEST. Use single quotes: in double
        quotes PowerShell would try to expand $' as a variable. This is the run that finds
        a test policy left on All Devices, which is the only High finding that describes
        something actually reaching devices.

    .EXAMPLE
        Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>" `
            -TestNameRegex '^TEST[-_]'

        The same check for a prefix convention: TEST-Wifi or TEST_Wifi. For either position,
        use '(^|[-_ ])TEST([-_ ]|$)'.

    .EXAMPLE
        Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>" `
            -GroupOwnerUpns "alice@contoso.com","bob@contoso.com" `
            -GroupNamePrefix "<your-prefix>-"

        Adds the Entra group section. Reports groups owned by those accounts whose names
        start with the prefix and that are either empty or referenced by no Windows Intune
        assignment. Both parameters are required for the section to run: owner scoping alone
        would pull in every Teams and Microsoft 365 group the owners happen to hold.

    .EXAMPLE
        Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>" `
            -NewAppGraceMonths 0

        Flags every unassigned application regardless of creation date. Use this when you do
        not retain previous versions for rollback, or when you want to see the full picture
        once. The default of 6 parks recently created unassigned apps as Watch instead.

    .EXAMPLE
        Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>" `
            -OutputFolder "C:\Reports\Intune" -HeaderColor '#31708F'

        Writes the workbook somewhere other than Documents and changes the header fill.
        -HeaderColor takes any HTML colour string.

    .EXAMPLE
        Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>" `
            -TestNameRegex '-TEST$' `
            -GroupOwnerUpns "alice@contoso.com" -GroupNamePrefix "<your-prefix>-" `
            -NewAppGraceMonths 3 -OutputFolder "C:\Reports\Intune"

        Everything switched on, with a shorter grace window for unassigned applications.
        This is the shape of a regular review run.

    .EXAMPLE
            Set-IntuneHousekeeperConfig -ClientId "<app id>" -TenantId "<tenant id>" `
                -TestNameRegex '(^|[-_ (\[])TEST([-_ )\]]|$)'
            Export-IntuneHousekeeperReport

        Save once, then run with no parameters at all. Anything passed explicitly on a
        later run overrides the stored value for that run only.

    .NOTES
        Required delegated permissions on the app registration, admin-consented:
          DeviceManagementApps.Read.All            applications and their assignments
          DeviceManagementConfiguration.Read.All   profiles, compliance, baselines
          DeviceManagementScripts.Read.All         remediations, platform scripts, and
                                                   script assignments read for the group
                                                   check
          Group.Read.All                           Entra group section only
          User.Read.All                            Entra group section only

        App registration: public client / native flow. Under Authentication, Add Redirect
        URI, Mobile and desktop applications, add both
        ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id> and http://localhost. The
        first is required for broker (WAM) sign-in, which a tenant enforcing Conditional
        Access token protection needs.

        Broker sign-in is enabled by default on Windows in current releases of
        Microsoft.Graph.Authentication and no longer needs to be turned on in code. If a
        fresh interactive sign-in fails with AADSTS530084, update that module and check the
        broker redirect URI above. The device must also be joined or registered and
        compliant, or token protection fails regardless.

        An existing Graph session for the same tenant and client ID is reused rather than
        replaced, and only a session this script opened is disconnected at the end. Running
        the script repeatedly therefore costs one sign-in, not one per run, which matters
        under Conditional Access token protection where every sign-in is a broker prompt.

        -Scopes is deliberately not passed to Connect-MgGraph. With a custom -ClientId, MSAL
        treats requested scopes as a new authorization and triggers a consent prompt; the
        token must carry what is already consented on the app registration. The script
        verifies the granted scopes instead.

        Required modules:
          Microsoft.Graph.Authentication
          ImportExcel                     (does not require Excel to be installed)

        Windows and PowerShell 7. Windows PowerShell 5.1 is not supported: .NET Framework
        allows one Microsoft.Identity.Client per process with no isolation, so an admin
        workstation carrying several Microsoft.Graph module versions fails at sign-in with
        'Could not load type ... Microsoft.Identity.Client'. PowerShell 7 loads the SDK
        dependencies in an isolated context and does not have this problem.

        A read-only app registration is recommended. The tool only issues GET, but a
        registration consented for ReadWrite holds a token that could change your tenant.

        ASCII-only file. No non-ASCII characters anywhere (no em-dashes, no smart quotes).
        Read-only: no PATCH, POST or DELETE calls are made against Graph.
    #>
    [CmdletBinding()]
    param(
        # Not mandatory, because mandatory binding happens before this function runs and
        # would prompt for a value the settings file already holds. Both are checked
        # after the settings merge, with a message naming Set-IntuneHousekeeperConfig.
        [string]   $ClientId,
        [string]   $TenantId,

        # Settings file to read. Explicit parameters always win over stored values.
        [string]   $ConfigPath,

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

        # How long a name-matched previous version is still worth keeping for rollback.
        # Beyond this it is flagged RetainedVersionExpired and returns to the cleanup queue:
        # a rollback copy nobody has needed in a year is not a rollback copy any more.
        # 0 disables the leniency entirely, so every name-matched copy stays in the queue.
        [ValidateRange(0, 120)]
        [int]      $RetainedVersionMonths = 12,

        # There is no default naming convention, so the completer offers the common shapes
        # already quoted correctly. A pattern like -TEST$ must be single-quoted: in double
        # quotes PowerShell tries to expand the $ sequence.
        [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            $suggestions = @(
                @{ Pattern = '-TEST$';                  Tip = 'Suffix: Wifi-TEST' }
                @{ Pattern = '^TEST[-_]';               Tip = 'Prefix: TEST-Wifi or TEST_Wifi' }
                @{ Pattern = '(^|[-_ (\[])TEST([-_ )\]]|$)'; Tip = 'Any position, word boundary. Does not match Latest or Attestation' }
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

    Reset-RunState

    # Settings precedence: an explicit parameter beats the settings file, which beats the
    # parameter default. $PSBoundParameters is the only way to tell '-NewAppGraceMonths 6'
    # from the default of 6, and getting that wrong would let a saved setting silently
    # override what the operator just typed.
    if (-not $ConfigPath) { $ConfigPath = Get-DefaultConfigPath }
    $script:ConfigPathUsed = $ConfigPath
    $stored = Read-ConfigFile -Path $ConfigPath
    foreach ($name in $script:ConfigSettingNames) {
        if ($PSBoundParameters.ContainsKey($name)) { continue }
        if (-not $stored.ContainsKey($name))       { continue }
        try {
            # Assigning through the parameter variable re-applies its type and any
            # ValidateRange attribute, so a bad stored value is caught here.
            Set-Variable -Name $name -Value $stored[$name] -Scope 0 -ErrorAction Stop
        }
        catch {
            throw ("Setting '{0}' in '{1}' is not valid: {2}" -f $name, $ConfigPath, $_.Exception.Message)
        }
    }

    if (-not $ClientId -or -not $TenantId) {
        throw "ClientId and TenantId are required. Pass -ClientId and -TenantId, or save them once with: Set-IntuneHousekeeperConfig -ClientId '<app id>' -TenantId '<tenant id>'. Both are on the Overview page of your app registration in the Entra admin center."
    }

    # Fail early on a malformed pattern rather than part way through the inventory. An
    # empty -TestNameRegex is valid and means the operator has not supplied a naming
    # convention, so test-object detection is skipped.
    if ($TestNameRegex) {
        try { $null = [regex]::new($TestNameRegex) }
        catch { throw ("-TestNameRegex is not a valid regular expression: {0}" -f $_.Exception.Message) }
    }

    # ---------------------------------------------------------------------------
    # Prerequisites and connection
    # ---------------------------------------------------------------------------

    # The manifest declares these in RequiredModules, so they are normally already loaded.
    # The check stays for anyone running the module from a copied folder rather than an
    # install, where the manifest's guarantee does not apply.
    foreach ($m in @('Microsoft.Graph.Authentication', 'ImportExcel')) {
        if (-not (Get-Module -Name $m) -and -not (Get-Module -ListAvailable -Name $m)) {
            throw ("Required module '{0}' not found. Install with: Install-Module {0} -Scope CurrentUser" -f $m)
        }
        Import-Module $m -ErrorAction Stop
    }

    # Multiple side-by-side versions of the Graph modules are common on an admin
    # workstation. PowerShell loads the first path in PSModulePath that holds the module,
    # not the newest version, so report what actually loaded rather than what is installed.
    $script:GraphAuthVersion = (Get-Module Microsoft.Graph.Authentication).Version
    $installedAuth = @(Get-Module Microsoft.Graph.Authentication -ListAvailable |
                       Select-Object -ExpandProperty Version | Sort-Object -Unique)
    Write-Host ("Microsoft.Graph.Authentication {0} loaded." -f $script:GraphAuthVersion)
    if ($installedAuth.Count -gt 1) {
        Write-Warning ("{0} versions of Microsoft.Graph.Authentication are installed ({1}). PowerShell loads by PSModulePath order, not by version. Mixed Graph module versions are the usual cause of sign-in failing with a 'Could not load type ... Microsoft.Identity.Client' error." -f $installedAuth.Count, ($installedAuth -join ', '))
    }

    try {
        # Conditional Access token protection (bound tokens) is only satisfied when sign-in
        # goes through the Windows broker (WAM). Current releases of
        # Microsoft.Graph.Authentication enable broker sign-in by default on Windows and the
        # old Set-MgGraphOption -EnableLoginByWAM switch no longer has any effect, so nothing
        # is set here. If a fresh interactive sign-in fails with AADSTS530084, update
        # Microsoft.Graph.Authentication and add the broker redirect URI
        # ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id> to the app registration.

        # Reuse a session that already matches. Every Connect-MgGraph under Conditional
        # Access token protection means another broker prompt, and running the script twice
        # in quick succession made the second sign-in fail with ApplicationCanceled while
        # the first was still tearing down. Only a session this script opened is closed
        # again, so a session the operator established stays theirs.
        $existing = Get-MgContext
        if ($existing -and
            ([string]$existing.TenantId -eq [string]$TenantId) -and
            ([string]$existing.ClientId -eq [string]$ClientId)) {
            $ctx = $existing
            Write-Host ("Reusing the existing Graph session for {0}" -f $ctx.Account)
        }
        else {
            if ($existing) {
                Write-Host 'An existing Graph session is for a different tenant or app registration. Reconnecting.'
                Disconnect-MgGraph | Out-Null
            }
            Write-Host 'Connecting to Microsoft Graph (delegated)...'
            # Do NOT pass -Scopes here. With a custom -ClientId, MSAL treats requested scopes
            # as a new authorization and triggers a consent prompt; the token must instead
            # carry the permissions already consented on the app registration.
            #
            # One retry, for one specific failure. Observed repeatedly in a tenant
            # enforcing token protection: the first sign-in after a previous run
            # disconnected fails with 'ApplicationCanceled / Current Request already
            # cancelled', and an immediate second attempt succeeds, because the broker is
            # still tearing down the old request. Only that signature is retried, and
            # only once, so a sign-in the operator genuinely cancelled is not forced back
            # on them repeatedly.
            try {
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -NoWelcome -ErrorAction Stop
            }
            catch {
                if ([string]$_.Exception.Message -notmatch 'ApplicationCanceled|Current Request already cancelled') { throw }
                Write-Warning 'The broker cancelled the sign-in request, which usually means it was still closing a previous session. Retrying once.'
                Start-Sleep -Seconds 3
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -NoWelcome -ErrorAction Stop
            }
            $script:ConnectionOwned = $true
            $ctx = Get-MgContext
            if (-not $ctx) { throw 'Failed to establish a Graph context.' }
            Write-Host ("Connected as {0}" -f $ctx.Account)
        }

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
        # Remediations and platform scripts sit behind their own permission, and it also
        # covers the macOS shell and custom attribute scripts read for group references.
        $neededScopes.Add('DeviceManagementScripts.Read.All')
        if ($GroupOwnerUpns.Count -gt 0) {
            $neededScopes.Add('Group.Read.All')
            $neededScopes.Add('User.Read.All')
        }
        foreach ($needed in $neededScopes) {
            # A ReadWrite grant satisfies the matching Read requirement. The tool only ever
            # issues GET, but plenty of app registrations are consented ReadWrite for other
            # tooling, and warning about a permission the token exceeds is just noise.
            $alt = $needed -replace '\.Read\.', '.ReadWrite.'
            if (($grantedScopes -notcontains $needed) -and ($grantedScopes -notcontains $alt)) {
                Write-Warning ("Token does not carry '{0}' or '{1}'. If this is unexpected, run Disconnect-MgGraph and re-run the script to refresh the cached session." -f $needed, $alt)
            }
        }
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
        foreach ($a in $appsRaw) { Add-AssignmentGroupReference $a.assignments }
        $appsWin = Select-WindowsObject -Objects $appsRaw -Area 'Applications' -WindowsTypes $script:WindowsAppTypes
        $appNameCounts = Get-NameCounts -Objects $appsWin -NameProperty 'displayName'

        # Index every Windows app by base name and version so a retained previous version can
        # be recognised without a supersedence relationship. Not every publishing workflow
        # creates supersedence, and where it does not there is no link to follow: an
        # unassigned 'App 1.2' sitting beside an assigned 'App 1.3' is a rollback copy, not
        # abandoned clutter. Built from data already collected, so no extra Graph calls.
        $appIndex = New-Object System.Collections.Generic.List[object]
        foreach ($app in $appsWin) {
            $dn = [string]$app.displayName
            $iai = Get-AssignmentInfo $app.assignments
            $appIndex.Add([pscustomobject]@{
                Id           = [string]$app.id
                Base         = (Get-AppBaseName $dn)
                Version      = (Get-AppVersion $dn)
                HasInclusion = [bool]$iai.HasInclusion
            })
        }

        $rowsApps = New-Object System.Collections.Generic.List[object]
        foreach ($app in $appsWin) {
            $ai = Get-AssignmentInfo $app.assignments

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

            # Name-and-version fallback for retained versions. Deliberately narrow: the app
            # must be unassigned, its name must carry a parseable version, and a sibling with
            # the same base name must have BOTH a higher version AND a live inclusion
            # assignment. Two unassigned versions of a retired app stay in the cleanup queue,
            # which is the correct outcome.
            if ($ai.Count -eq 0 -and $supBy -eq 0) {
                if (Test-SupersededByName -DisplayName ([string]$app.displayName) -Id ([string]$app.id) -Index $appIndex) {
                    $extra += 'SupersededByName'
                    # A rollback copy has a shelf life. Past -RetainedVersionMonths the
                    # rollback is not realistic any more and the package returns to the
                    # queue, still labelled so the operator knows what it is.
                    if (Test-Stale $app.createdDateTime $RetainedVersionMonths) {
                        $extra += 'RetainedVersionExpired'
                    }
                }
            }
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

        if ($GroupOwnerUpns.Count -gt 0) { Add-ReferenceOnlyGroups }

        Write-Host 'Collecting owner-scoped Entra ID groups...'
        if ($script:GraphReadIncomplete -and $GroupOwnerUpns.Count -gt 0) {
            Write-Warning 'Entra group analysis skipped: at least one Graph read failed, so the set of referenced groups is incomplete. A group could be reported as unreferenced only because the assignment naming it was never read. Re-run once the failure above is resolved.'
            $GroupOwnerUpns = @()
        }
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
        $runInfo.Add([pscustomobject][ordered]@{ Setting='NewAppGraceMonths';     Value=[string]$NewAppGraceMonths })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='RetainedVersionMonths'; Value=$(if ($RetainedVersionMonths -eq 0) { '0 - retained copies not excused' } else { [string]$RetainedVersionMonths }) })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='GroupNamePrefix';      Value=$(if ($GroupNamePrefix) { $GroupNamePrefix } else { '(not set)' }) })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='Group owners checked'; Value=[string]$GroupOwnerUpns.Count })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='Settings file';         Value=$(if ($script:ConfigPathUsed -and (Test-Path -LiteralPath $script:ConfigPathUsed)) { $script:ConfigPathUsed } else { '(none - parameters only)' }) })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='PowerShell';            Value=[string]$PSVersionTable.PSVersion })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='Graph auth module';     Value=[string]$script:GraphAuthVersion })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='Collection complete';   Value=$(if ($script:GraphReadIncomplete) { 'No - at least one Graph read failed' } else { 'Yes' }) })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='Actionable means';      Value='High + Medium. Low and Watch are excluded.' })
        $runInfo.Add([pscustomobject][ordered]@{ Setting='Test names matched';    Value=$(if ($TestNameRegex) { ('{0} of {1} object names' -f $script:TestNameMatches, $script:ObjectsExamined) } else { 'detection not enabled' }) })

        # Per-priority columns: with retained versions parked at Watch and Low, a bare
        # Total/Actionable pair hides where most of the estate went.
        $summary = New-Object System.Collections.Generic.List[object]
        function Add-SummaryRow {
            param([string]$Category, $Rows)
            $c = Get-PriorityCounts $Rows
            $summary.Add([pscustomobject][ordered]@{
                Category   = $Category
                Total      = (Get-RowCount $Rows)
                High       = $c.High
                Medium     = $c.Medium
                Low        = $c.Low
                Watch      = $c.Watch
                Healthy    = $c.Healthy
                Actionable = ($c.High + $c.Medium)
            })
        }
        Add-SummaryRow 'Applications'         $rowsApps
        Add-SummaryRow 'ConfigProfiles'       $rowsConfig
        Add-SummaryRow 'CompliancePolicies'   $rowsCompliance
        Add-SummaryRow 'SecurityBaselines'    $rowsBaseline
        Add-SummaryRow 'Remediations'         $rowsRemediation
        Add-SummaryRow 'PlatformScripts'      $rowsScripts
        Add-SummaryRow 'EntraGroups(flagged)' $rowsGroups

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
            elseif ($f -match 'RetainedVersionExpired') {
                $reason = 'Retained previous version of an application that is still assigned, but created outside the rollback window; a rollback this old is no longer realistic'
                $action = 'Remove'
            }
            elseif ($f -match 'SupersededByName') {
                $reason = 'Unassigned, but a newer version of the same application is assigned (retained rollback copy, matched by name and version because no supersedence relationship exists)'
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
                $reason = 'Owned assignment group with ' + ($parts -join ' and ') + '. Verify in the portal first: the group could still be used by an object type this tool does not read'
                $action = 'Investigate'
            }
            elseif ($pri -eq 'Medium') {
                $reason = 'Unassigned; nothing references it'
                $action = 'Remove'
            }
            elseif ($f -match 'TestNamed') {
                $reason = 'TEST-named object scoped to a group; confirm the test is finished'
                $action = 'Investigate'
            }
            # A test-named object that is also unassigned matches the Medium branch first, so
            # the reason would lose the fact that it is a test object. Keep the label.
            if (($f -match 'TestNamed') -and ($reason -notmatch 'test')) { $reason = $reason + '; test-named' }
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
        $pkg = (Get-ExportRows $worklist -Worklist) | Export-Excel -ExcelPackage $pkg -WorksheetName 'Worklist'           @common
        $pkg = (Get-ExportRows $rowsApps)        | Export-Excel -ExcelPackage $pkg -WorksheetName 'Applications'       @common
        $pkg = (Get-ExportRows $rowsConfig)      | Export-Excel -ExcelPackage $pkg -WorksheetName 'ConfigProfiles'     @common
        $pkg = (Get-ExportRows $rowsCompliance)  | Export-Excel -ExcelPackage $pkg -WorksheetName 'CompliancePolicies' @common
        $pkg = (Get-ExportRows $rowsBaseline)    | Export-Excel -ExcelPackage $pkg -WorksheetName 'SecurityBaselines'  @common
        $pkg = (Get-ExportRows $rowsRemediation) | Export-Excel -ExcelPackage $pkg -WorksheetName 'Remediations'       @common
        $pkg = (Get-ExportRows $rowsScripts)     | Export-Excel -ExcelPackage $pkg -WorksheetName 'PlatformScripts'    @common
        $pkg = (Get-ExportRows $rowsGroups)      | Export-Excel -ExcelPackage $pkg -WorksheetName 'EntraGroups'        @common

        # Header styling. -HeaderColor accepts any HTML colour string.
        try {
            $brand = [System.Drawing.ColorTranslator]::FromHtml($HeaderColor)
            foreach ($ws in $pkg.Workbook.Worksheets) {
                if ($ws.Dimension) {
                    $lastCol = $ws.Dimension.End.Column
                    $lastRow = $ws.Dimension.End.Row
                    $hdr = $ws.Cells[1, 1, 1, $lastCol]
                    $hdr.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $hdr.Style.Fill.BackgroundColor.SetColor($brand)
                    $hdr.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $hdr.Style.Font.Bold = $true

                    # Counts are whole objects, not measurements. Without this they render as
                    # 698.00, because the export converts numeric-looking values to doubles.
                    if ($ws.Name -eq 'Summary' -and $lastRow -gt 1 -and $lastCol -gt 1) {
                        $ws.Cells[2, 2, $lastRow, $lastCol].Style.Numberformat.Format = '0'
                    }
                    # RunInfo values are labels, some of which happen to look like numbers.
                    if ($ws.Name -eq 'RunInfo' -and $lastRow -gt 1) {
                        $ws.Cells[2, 2, $lastRow, 2].Style.Numberformat.Format = '@'
                    }
                }
            }
        }
        catch {
            Write-Warning ("Header styling skipped: {0}" -f $_.Exception.Message)
        }

        Close-ExcelPackage $pkg

        Write-Host ''
        Write-UnrecognisedTypeWarning

    if ($script:GraphReadIncomplete) {
        Write-Warning 'This report is INCOMPLETE. At least one Graph read failed (see the warnings above), so one or more categories are missing objects and their totals understate the estate. A 403 usually means the app registration is missing a permission listed in the README. RunInfo records this on the Collection complete row.'
    }

        if ($TestNameRegex) {
            Write-Host ("Test-object detection: {0} of {1} object names matched '{2}'." -f $script:TestNameMatches, $script:ObjectsExamined, $TestNameRegex)
            if ($script:TestNameMatches -eq 0) {
                Write-Warning ("No object name matched '{0}'. That is either a clean estate or the wrong pattern for your naming convention. Anchor on a word boundary rather than a bare substring: '(^|[-_ (\[])TEST([-_ )\]]|$)' matches Wifi-TEST, Wifi_TEST and Wifi (test), while a bare 'test' also matches Latest and Attestation." -f $TestNameRegex)
            }
        }

        Write-Host ("Done. Decision tracker written to: {0}" -f $xlsx)
        Write-Host ("Worklist items (High/Medium/Low): {0}" -f (Get-RowCount $worklist))
    }
    catch {
        $errMessage = [string]$_.Exception.Message

        # Translate the failures that are environmental rather than tenant problems, so the
        # operator is not left reading a .NET type-load error and guessing.
        $likelyCause = ''
        if ($errMessage -match 'Could not load type|Microsoft\.Identity\.Client|FileLoadException|Could not load file or assembly') {
            $likelyCause = 'Assembly conflict in this PowerShell session. Several side-by-side versions of the Microsoft.Graph modules share one Microsoft.Identity.Client, and Windows PowerShell 5.1 cannot isolate them. Run this in PowerShell 7, which loads the SDK dependencies in an isolated context.'
        }
        elseif ($errMessage -match 'AADSTS530084') {
            $likelyCause = 'Conditional Access token protection rejected the sign-in. The app registration needs the broker redirect URI ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id>, and the device must be joined or registered and compliant.'
        }
        elseif ($errMessage -match 'ApplicationCanceled|user_canceled|access_denied|Current Request already cancelled') {
            $likelyCause = 'The sign-in prompt was closed, cancelled, or timed out. Re-run and complete the sign-in. If no prompt appeared, check for a broker window behind the console.'
        }
        elseif ($errMessage -match 'AADSTS65001|consent') {
            $likelyCause = 'The app registration has not been admin-consented for the delegated permissions listed in the README.'
        }
        elseif ($errMessage -match 'Forbidden|403') {
            $likelyCause = 'The signed-in account or the app registration lacks a required read permission. Check the scopes reported above.'
        }

        Write-Host ''
        Write-Host '==================== ERROR ===================='
        Write-Host ("Type    : {0}" -f $_.Exception.GetType().FullName)
        Write-Host ("Message : {0}" -f $errMessage)
        if ($likelyCause) { Write-Host ("Likely  : {0}" -f $likelyCause) }
        if ($_.InvocationInfo) {
            Write-Host ("Line #  : {0}" -f $_.InvocationInfo.ScriptLineNumber)
            Write-Host ("Command : {0}" -f ([string]$_.InvocationInfo.Line).Trim())
        }
        Write-Host ("Stack   :")
        Write-Host ($_.ScriptStackTrace)
        Write-Host '==============================================='
    }
    finally {
        if ($script:ConnectionOwned -and (Get-MgContext)) { Disconnect-MgGraph | Out-Null }
    }

}

Export-ModuleMember -Function @(
    'Export-IntuneHousekeeperReport'
    'Get-IntuneHousekeeperConfig'
    'Set-IntuneHousekeeperConfig'
)
