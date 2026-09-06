# Intune Housekeeper

A read-only PowerShell module that inventories your Windows Intune estate through
Microsoft Graph and produces an Excel worklist of objects worth cleaning up:
unassigned policies, leftover test objects, broken assignments, and empty groups.

**It never modifies, unassigns, or deletes anything in your tenant.** It issues GET
requests only. Every change is made by a human, in the portal, after review. Intune has
no recycle bin, and this tool is built around that fact.

You do not have to take that on trust. Every Graph call in the script carries an
explicit method, and they are visible in one grep:

```powershell
Select-String -Path .\IntuneHousekeeper\IntuneHousekeeper.psm1 -Pattern '-Method' -SimpleMatch
```

Every hit is `-Method GET`, and CI fails the build on anything else.

---

## Why

The Intune portal has no list view of assignment status. Finding out what is unassigned
means opening every object by hand, so nobody does it, and years of test policies and
superseded applications quietly accumulate.

This gives you one sorted worklist instead.

---

## What it covers

Windows objects: applications, configuration profiles (templates, Settings Catalog,
ADMX), compliance policies, security baselines, remediations, and platform scripts.
Optionally, Entra ID assignment groups that are empty or referenced by nothing.

Not reported on: Autopilot deployment profiles, enrolment configurations, assignment
filters, feature, quality and driver update profiles, app configuration and app
protection policies, and macOS shell and custom attribute scripts.

Their **assignments are still read** when the Entra group section runs, purely to
establish which groups are referenced. Reporting is Windows only; referencing has to be
tenant wide, or a group used solely by a macOS shell script looks like it is used by
nothing.

Object types are classified as Windows, another platform, or unrecognised. Applications
are matched against an explicit list of Windows types, profiles and compliance policies
against a Windows pattern. Anything unrecognised is left out of the report and named in
a warning at the end of the run, so a type the tool has not seen before is visible
rather than silently included or silently dropped.

macOS, iOS, and Android are not covered. Pull requests welcome.

All Intune calls use the Graph **beta** endpoint, because several of these object types
and their assignment expansions are only available there. Microsoft can change beta
without notice.

---

## Requirements

- Windows. Broker sign-in and the workbook header styling both depend on it.
- **PowerShell 7.** Windows PowerShell 5.1 is not supported; the manifest requires 7.0. See below for why.
- Dependencies: `Microsoft.Graph.Authentication` and `ImportExcel`, both declared in the
  manifest and installed automatically. ImportExcel writes the workbook itself, so Excel
  does not need to be installed. See the note on the Excel dependency below before
  deploying this in a commercial environment.
- An Entra app registration, described below.

### Why PowerShell 7 and not 5.1

Windows PowerShell 5.1 runs on .NET Framework, which permits one version of a given
assembly per process and provides no isolation between modules. An admin workstation
that has accumulated several `Microsoft.Graph.*` module versions ends up with one
`Microsoft.Identity.Client` serving all of them, and sign-in fails before the tool runs:

```
Connect-MgGraph : InteractiveBrowserCredential authentication failed:
Could not load type 'Microsoft.Identity.Client.AuthScheme.TokenType'
from assembly 'Microsoft.Identity.Client, Version=4.67.2.0'
```

Fixing that means uninstalling Graph modules the admin uses for other work, which is not
a reasonable prerequisite for a read-only report. PowerShell 7 loads the SDK's
dependencies in an isolated context and does not have the problem. This was reproduced
on a normal enterprise workstation: identical command, 5.1 failed, 7 worked.

### App registration

Register your own app rather than consenting broad scopes to the shared Microsoft Graph
PowerShell app.

1. **App registrations > New registration.** Single tenant. No redirect URI is needed at
   this stage.
2. **Authentication > Add Redirect URI > Mobile and desktop applications.** Add both:
   - `ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id>`, required for broker (WAM)
     sign-in, which any tenant enforcing Conditional Access token protection needs
   - `http://localhost`
3. **API permissions.** Add these as **delegated** Microsoft Graph permissions and grant
   admin consent:

  | Permission | Used for |
  |---|---|
  | `DeviceManagementApps.Read.All` | Applications and their assignments |
  | `DeviceManagementConfiguration.Read.All` | Profiles, compliance policies, baselines |
  | `DeviceManagementScripts.Read.All` | Remediations, platform scripts, and script assignments read for the group check |
  | `Group.Read.All` | Entra group details and membership (optional section only) |
  | `User.Read.All` | Resolving owner accounts and their groups (optional section only) |

The last two are only needed if you use `-GroupOwnerUpns`. On a run without it, the
module neither expects nor warns about them.

**Use read-only scopes.** A `ReadWrite` grant satisfies the matching `Read` requirement
and the script accepts it without complaint, but a registration consented for
`ReadWrite` holds a token capable of changing your tenant. The tool only ever issues
GET, and CI enforces that on every commit, so the registration is the last place that
capability can live. A dedicated read-only app registration makes the safety property
structural rather than behavioural.

### A note on the Excel dependency

The output is a workbook rather than CSV because the Worklist is a decision tracker: it
has columns you type into, and it has to survive being shared, filtered, and returned
to. Ten CSV files would not.

`ImportExcel` builds the file using EPPlus, which is bundled with it as a DLL. EPPlus
licensing is worth knowing about if you work somewhere that cares:

- EPPlus 4.5.3.3 was the last release under LGPL.
- From version 5, EPPlus moved to the Polyform Noncommercial 1.0.0 licence, which is
  still free in some cases but requires a commercial licence for use in a commercial
  business.

That distinction sits in a dependency of this tool, not in the tool itself, so check
which version your installed copy ships rather than taking anyone's word for it:

```powershell
$m = Get-Module ImportExcel -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
(Get-Item (Join-Path $m.ModuleBase 'EPPlus.dll')).VersionInfo.FileVersion
```

A `4.5.3.x` result is the LGPL line. Anything `5.x` or later is the Polyform licence,
and if your employer has a policy about that, this is where it applies. Intune
Housekeeper itself is MIT and makes no claim about its dependencies' terms.

---

## Install

```powershell
Install-Module IntuneHousekeeper -Scope CurrentUser
```

Or clone the repository and import the module folder directly:

```powershell
Import-Module .\IntuneHousekeeper\IntuneHousekeeper.psd1
```

If that fails with "is not digitally signed", the files came from a browser download or a
zip and carry Windows' mark of the web, which `RemoteSigned` blocks. Clear it:

```powershell
Get-ChildItem . -Recurse | Unblock-File
```

Files obtained with `git clone` are never marked, and neither are files installed with
`Install-Module`. If `Get-ExecutionPolicy -List` shows `AllSigned`, that is a different
matter: nothing unsigned will load regardless of origin, and this module is not signed.

## Usage

```powershell
Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>"
```

With test-object detection and the optional Entra group check:

```powershell
Export-IntuneHousekeeperReport -ClientId "<app id>" -TenantId "<tenant id>" `
    -TestNameRegex '(^|[-_ (\[])TEST([-_ )\]]|$)' `
    -GroupOwnerUpns "alice@contoso.com","bob@contoso.com" `
    -GroupNamePrefix "<your-prefix>-"
```

### Save your settings instead of retyping them

```powershell
Set-IntuneHousekeeperConfig -ClientId "<app id>" -TenantId "<tenant id>" `
    -TestNameRegex '(^|[-_ (\[])TEST([-_ )\]]|$)' -OutputFolder 'C:\Reports\Intune'

Export-IntuneHousekeeperReport
```

Settings are written to `%APPDATA%\IntuneHousekeeper\settings.json`, or wherever
`-ConfigPath` points, and only the values you pass are stored. `Get-IntuneHousekeeperConfig`
shows what is saved and where; `Set-IntuneHousekeeperConfig -RemoveSetting <name>` clears
one.

Precedence is **explicit parameter, then settings file, then default**, so a value passed
on the command line always wins for that run without changing what is saved. A stored `0`
or `""` is a real setting, not an absent one: `"NewAppGraceMonths": 0` means no grace
window, while the key being missing means use the default.

`ClientId` and `TenantId` are the only required values, from either source. Without them
the command stops with a message telling you how to save them, rather than prompting.

Nothing secret is stored. The tool signs in through a public client flow, which has no
secret, so the file holds identifiers and preferences only. It still identifies your
tenant, so keep it out of repositories and screenshots.

The module exports three commands. `Get-Help Export-IntuneHousekeeperReport -Examples`
prints eight worked examples.

If you connect to Graph yourself first, the command reuses that session and leaves it
open. Otherwise it signs in and disconnects when it finishes.

### Finding your way around the parameters

Tab completion covers the two that are awkward to type. `-OutputFolder` completes
existing directories and quotes paths containing spaces. `-TestNameRegex` offers the
common naming conventions, already single-quoted, which matters because a pattern like
`-TEST$` breaks if you put it in double quotes.

`Get-Help Export-IntuneHousekeeperReport -Examples` prints seven worked examples, and
`-?` prints the full parameter list. If you run it with no parameters and PowerShell
prompts for `ClientId`, type `!?` at that prompt to be told where in the Entra portal
to find the value.

### Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-ClientId` | required | App registration client ID. From this parameter or the settings file |
| `-TenantId` | required | Tenant ID. From this parameter or the settings file |
| `-ConfigPath` | `%APPDATA%\IntuneHousekeeper\settings.json` | Settings file to read |
| `-OutputFolder` | `~\Documents` | Where the workbook is written |
| `-NewAppGraceMonths` | `6` | Unassigned apps created within this many months are parked as `Watch`. `0` flags every unassigned app. Range 0-120 |
| `-RetainedVersionMonths` | `12` | How long a name-matched previous version stays out of the cleanup queue. `0` queues every retained copy. Range 0-120 |
| `-TestNameRegex` | none | Regex matching your own test-object naming convention. Empty skips test-object detection |
| `-GroupNamePrefix` | none | Restricts the group check to your assignment group naming convention |
| `-GroupOwnerUpns` | none | Owner accounts whose groups are checked. Empty skips the section |
| `-HeaderColor` | `#404040` | Worksheet header fill. Any HTML colour string |

---

## The output

One workbook, `Intune-Housekeeper_<timestamp>.xlsx`. Two sheets matter:

- **Summary** - totals and actionable counts per category
- **Worklist** - every actionable object, pre-sorted, with a plain-language reason and
  a suggested action. Record your decisions here.

**RunInfo** records the settings the run used, so a workbook can be read months later
without guessing which checks were switched on. It holds no tenant or account
identifiers and is safe to share.

The remaining sheets are read-only inventory per object type.

"Actionable" means **High plus Medium**: the rows where you are expected to do
something. `Low` is "keep or confirm", `Watch` is parked, and neither is counted. Low
rows still appear on the Worklist, below the others, as context.

The Summary breaks each category down by priority, so you can see where the estate went
rather than only how much of it needs attention.

One caveat on the Summary: the `EntraGroups(flagged)` row counts only the groups that
were flagged, not every group examined. Every other row counts the whole category.

### Priority

`Priority` means **order of cleanup action, not risk to devices**.

| Priority | Meaning |
|---|---|
| **High** | Reality differs from intent: a test-named object on All Devices/All Users, or a group both included and excluded within one assignment intent. Correct it, do not delete it. |
| **Medium** | Unassigned or exclusion-only, with nothing referencing it. The cleanup queue. |
| **Low** | Unassigned but referenced by another app, or a test scoped to a group. Usually keep. |
| **Watch** | An unassigned app created within `-NewAppGraceMonths`. Parked, not counted. |
| *(blank)* | Assigned and healthy. |

Only High can actually affect a device.

Flagged Entra groups are always reported as Medium. The scale above describes Intune
objects; a group is either worth a look or it is not.

### Flags

| Flag | Meaning |
|---|---|
| `Unassigned` | No assignments at all |
| `OnlyExclusions` | Exclusion targets only, no inclusion. Deployed to nothing |
| `TestNamed` | Display name matches `-TestNameRegex` |
| `TestNamedBroadAssign` | Test-named and assigned to All Devices or All Users |
| `DuplicateName` | Display name occurs more than once within its type. Informational only, never raises priority |
| `IncludeExcludeOverlap` | Same group included and excluded within one assignment intent. The exclusion wins, so that group is silently skipped |
| `SupersededByNewer` | (apps) A newer version supersedes this one |
| `SupersededByName` | (apps) Unassigned, but a newer version of the same application is assigned. For workflows that do not create supersedence |
| `RetainedVersionExpired` | (apps) `SupersededByName`, but older than `-RetainedVersionMonths`. Back in the cleanup queue |
| `HasDependents` | (apps) Another app depends on this one |
| `RecentlyCreated` | (apps) Unassigned, created within `-NewAppGraceMonths` |
| `ZeroMembers` | (groups) No direct members |
| `NotUsedInIntune` | (groups) Not referenced by any Windows Intune assignment |

---

## Test objects are opt-in

There is no built-in test-naming convention. `-TestNameRegex` is empty by default, and
without it the `TestNamed` and `TestNamedBroadAssign` flags never fire.

This is deliberate. A shipped default such as `-TEST$` reports nothing on a tenant that
names things `TEST-Wifi` or `Wifi (test)`, and a report showing no test objects reads as
a clean result rather than as a check that never ran. Supply your own pattern:

```powershell
-TestNameRegex '-TEST$'                          # hyphen suffix only
-TestNameRegex '[-_ ]test$'                      # suffix, any separator
-TestNameRegex '(^|[-_ (\[])TEST([-_ )\]]|$)'    # any position, word boundary
```

A test-named object assigned to All Devices or All Users is the only finding in this
tool that describes something actually reaching devices, so it is worth setting
correctly.

**Anchor the pattern.** A bare `test` also matches `Latest`, `Attestation` and
`Protest`. In one real tenant, 75 object names contained the substring and 53 of those
were Latest or Attestation objects. Flagging a Device Health Attestation policy on All
Devices as a leftover test object would put a `High` row in front of you that is
completely wrong. The word-boundary form above matches `Wifi-TEST`, `Wifi_TEST`,
`Wifi TEST` and `Wifi (test)` while leaving all of those alone.

**Check the match count.** The run reports how many object names matched, and warns when
none did, because a report with no test findings otherwise reads as a clean estate
rather than as a wrong pattern. The same count is recorded on the RunInfo sheet. If you
expect leftovers and see zero, the separator is the usual culprit: matching is
case-insensitive, so `-TEST$` finding nothing while `[-_ ]test$` finds plenty means your
convention uses a space or an underscore rather than a hyphen.

---

## Unassigned applications and the grace window

Unassigned applications are flagged like anything else, but two kinds of them are
deliberate rather than abandoned: a package created recently and not assigned yet, and a
previous version retained so a bad release can be rolled back.

Where your publishing workflow creates Intune **supersedence** relationships, a retained
version is identified properly, flagged `SupersededByNewer`, and dropped to Low.

Where it does not, two fallbacks apply.

**Creation date.** Any unassigned app created within `-NewAppGraceMonths` is parked as
`Watch` and left out of the worklist and the actionable counts. You choose the window:
roughly the age of the oldest version you expect to keep, or `0` to flag every
unassigned app regardless of age.

**Name and version.** An unassigned app is flagged `SupersededByName` and dropped to Low
when its display name carries a version, and another app with the same base name has
both a higher version and a live inclusion assignment. `App 1.2` sitting unassigned
beside an assigned `App 1.3` is a rollback copy, not clutter.

The rule is deliberately narrow. Two unassigned versions of a retired app stay in the
cleanup queue, which is correct. Architecture and edition markers survive base-name
matching, so an x64 package is never matched against x86. An app whose name carries no
version is never demoted.

**Rollback copies expire.** Keeping the previous version is worth it because a bad
update can be rolled back. That argument weakens with age: nobody is rolling back to a
build from two years ago. Past `-RetainedVersionMonths`, a name-matched copy is flagged
`RetainedVersionExpired` and returns to `Medium` with a suggested action of Remove,
still labelled so you can see what it is rather than wondering why an old version is in
the queue.

The default of 12 months is a starting point, not a recommendation. Set it to how long a
rollback is realistically useful where you work, `0` to queue every retained copy, or
`120` to keep them indefinitely.

A real Intune **supersedence** relationship never expires this way. That link is
configuration the newer app depends on, so a superseded version stays `Keep` regardless
of age.

---

## Known limitations

- `NotUsedInIntune` means "not referenced by any assignment found". Every platform
  counts, and twelve object types outside the report are read for their assignments as
  well. Any type Microsoft adds later is invisible until it is added to that list, which
  is why group rows are `Investigate` and never `Remove`.
- Assignments pointing at **empty or deleted groups** are not detected. Those look
  healthy while deploying to nothing. This is the most useful thing to add next.
- **Assignment filters are not evaluated.** An object assigned through a filter that
  matches no device is reported as healthy.
- The time an object was **unassigned** is not recoverable from metadata. Assignment
  changes do not update `lastModifiedDateTime` (verified). The audit log would be
  required.
- A **wrongly scoped object with a conventional name** cannot be detected. Naming is the
  only signal of intent.
- Detection of test objects and of assignment groups both depend on **naming
  conventions** you supply. Anything that deviates is invisible, and if you supply
  nothing, that check does not run at all.
- **Unrecognised object types are excluded**, not guessed at. They are named in a
  warning at the end of the run. A future version may handle them differently.

---

## Conditional Access

If your tenant enforces **token protection**, a fresh Graph sign-in that is not brokered
fails with `AADSTS530084`.

Current releases of `Microsoft.Graph.Authentication` enable broker (WAM) sign-in by
default on Windows, and the older `Set-MgGraphOption -EnableLoginByWAM` switch no longer
has any effect, so the script sets nothing. If you hit `AADSTS530084`, update that
module and check that the app registration carries the
`ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id>` redirect URI. The device must also
be joined or registered and compliant, or token protection fails regardless.

A related trap: a cached token keeps working after a policy or consent change, so only a
*fresh* sign-in trips this. Run `Disconnect-MgGraph` and re-run to test properly.

## Adding a permission later

Consenting a permission on the app registration does not put it in your token. A cached
refresh token minted before the change is reused silently, so calls keep failing with 403
while the portal shows the permission granted. `Disconnect-MgGraph` does not reliably help
either: it can report "no application to sign out from" while the cached token lives on,
and it sometimes warns that it could not clear the MSAL cache.

Force one fresh authorization by asking for the scopes explicitly:

```powershell
Disconnect-MgGraph -ErrorAction SilentlyContinue

Connect-MgGraph -ClientId "<app id>" -TenantId "<tenant id>" -NoWelcome -Scopes `
    'DeviceManagementApps.Read.All','DeviceManagementConfiguration.Read.All',
    'DeviceManagementScripts.Read.All','Group.Read.All','User.Read.All'

(Get-MgContext).Scopes
```

Then run the report in the same session; it reuses that connection. The module itself
never passes `-Scopes`, because with a custom client ID MSAL treats it as a new
authorization and prompts for consent. Here that prompt is exactly what you want, once.

If the scope is still missing afterwards, close every PowerShell window and delete
`%LOCALAPPDATA%\.IdentityService\msal.cache`.

---

## Contributing

Forks and flavours are the point. See [`docs/DESIGN-NOTES.md`](docs/DESIGN-NOTES.md) for
why things work the way they do, including several PowerShell 5.1 landmines worth not
rediscovering, and [`AGENTS.md`](AGENTS.md) if you are working with an AI assistant on
this repo.

Two conventions worth keeping:

1. **ASCII only.** No smart quotes, no em-dashes. They survive your editor and break
   somewhere else.
2. **Read-only.** If you add write operations, make them opt-in and loud.

---

## Licence

MIT. See [`LICENSE`](LICENSE).
