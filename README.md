# Intune Housekeeper

A read-only PowerShell script that inventories your Windows Intune estate through
Microsoft Graph and produces an Excel worklist of objects worth cleaning up:
unassigned policies, leftover test objects, broken assignments, and empty groups.

**It never modifies, unassigns, or deletes anything in your tenant.** It issues GET
requests only. Every change is made by a human, in the portal, after review. Intune has
no recycle bin, and this tool is built around that fact.

You do not have to take that on trust. Every Graph call in the script carries an
explicit method, and they are visible in one grep:

```powershell
Select-String -Path .\Get-IntuneHousekeeper.ps1 -Pattern '-Method' -SimpleMatch
```

Every hit is `-Method GET`.

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

Not collected: Autopilot deployment profiles, enrolment status page and other enrolment
configurations, assignment filters, and feature, quality, or driver update profiles.

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
- Windows PowerShell 5.1 or PowerShell 7 on Windows. Both are supported.
- Modules: `Microsoft.Graph.Authentication`, `ImportExcel`. ImportExcel writes the
  workbook itself, so Excel does not need to be installed.
  ```powershell
  Install-Module Microsoft.Graph.Authentication, ImportExcel -Scope CurrentUser
  ```
  See the note on the Excel dependency below before deploying this in a commercial
  environment.
- An Entra app registration, described below.

### App registration

Register your own app rather than consenting broad scopes to the shared Microsoft Graph
PowerShell app.

1. **App registrations > New registration.** Single tenant. Under **Redirect URI**,
   choose **Public client/native** and enter `http://localhost`.
2. **Authentication > Add a platform > Mobile and desktop applications.** Add
   `ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id>`. This is required for broker
   (WAM) sign-in, which any tenant enforcing Conditional Access token protection needs.
3. If you run **Windows PowerShell 5.1**, also tick
   `https://login.microsoftonline.com/common/oauth2/nativeclient`.
4. **API permissions.** Add these as **delegated** Microsoft Graph permissions and grant
   admin consent:

  | Permission | Used for |
  |---|---|
  | `DeviceManagementApps.Read.All` | Applications and their assignments |
  | `DeviceManagementConfiguration.Read.All` | Profiles, compliance, baselines, scripts |
  | `Group.Read.All` | Entra group details and membership (optional section only) |
  | `User.Read.All` | Resolving owner accounts and their groups (optional section only) |

The last two are only needed if you use `-GroupOwnerUpns`. On a run without it, the
script neither expects nor warns about them.

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

## Usage

```powershell
.\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>"
```

With test-object detection and the optional Entra group check:

```powershell
.\Get-IntuneHousekeeper.ps1 -ClientId "<app id>" -TenantId "<tenant id>" `
    -TestNameRegex '-TEST$' `
    -GroupOwnerUpns "alice@contoso.com","bob@contoso.com" `
    -GroupNamePrefix "<your-prefix>-"
```

### Finding your way around the parameters

Tab completion covers the two that are awkward to type. `-OutputFolder` completes
existing directories and quotes paths containing spaces. `-TestNameRegex` offers the
common naming conventions, already single-quoted, which matters because a pattern like
`-TEST$` breaks if you put it in double quotes.

`Get-Help .\Get-IntuneHousekeeper.ps1 -Examples` prints seven worked examples, and
`-?` prints the full parameter list. If you run it with no parameters and PowerShell
prompts for `ClientId`, type `!?` at that prompt to be told where in the Entra portal
to find the value.

### Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-ClientId` | required | App registration client ID |
| `-TenantId` | required | Tenant ID |
| `-OutputFolder` | `~\Documents` | Where the workbook is written |
| `-NewAppGraceMonths` | `6` | Unassigned apps created within this many months are parked as `Watch`. `0` flags every unassigned app. Range 0-120 |
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

"Actionable" means High, Medium, or Low. Blank and `Watch` rows are excluded, so the
actionable column adds up to the number of rows on the Worklist.

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
-TestNameRegex '-TEST$'                    # suffix
-TestNameRegex '^TEST[-_]'                 # prefix
-TestNameRegex '(^|[-_ ])TEST([-_ ]|$)'    # either
```

The run states on screen which pattern it used, or that detection was skipped, and the
same goes on the RunInfo sheet. A test-named object assigned to All Devices or All Users
is the only finding in this tool that describes something actually reaching devices, so
it is worth setting.

---

## Unassigned applications and the grace window

Unassigned applications are flagged like anything else, but two kinds of them are
deliberate rather than abandoned: a package created recently and not assigned yet, and a
previous version retained so a bad release can be rolled back.

Where your publishing workflow creates Intune **supersedence** relationships, a retained
version is identified properly, flagged `SupersededByNewer`, and dropped to Low.

Where it does not, there is no link to follow, so creation date is the fallback. Any
unassigned app created within `-NewAppGraceMonths` is parked as `Watch` and left out of
the worklist and the actionable counts. You choose the window: set it to roughly the age
of the oldest version you expect to keep, or to `0` to flag every unassigned app
regardless of age.

Trade-off at the default of six months: an app that has had no new version in longer
than that will have its retained copy age into `Medium` and appear for review.

---

## Known limitations

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
