# Design notes

Why this tool works the way it does. If you are forking it, read this before changing
behaviour: most of the non-obvious choices here are the result of something going wrong
in a real tenant.

---

## The core principle: read-only, always

The script never modifies, unassigns, or deletes anything. It issues GET requests only.

This is not caution for its own sake. Intune has no recycle bin. A deleted policy,
application, or script is gone permanently, and recovery depends entirely on whatever
backup process you have in place. A tool that could delete faster than a human could
notice would be a liability, not a help.

If you fork this and add write operations, that is your call, but understand that you
are removing the property that makes it safe to run against production on a whim.

---

## Priority, not risk

The ranking column is called `Priority` and it means **order of cleanup action**, not
danger to devices.

The first version called it `RiskLevel`, which was wrong. An unassigned policy carries
no risk at all: it applies to nothing. Calling it "medium risk" made the report sound
alarming while describing inert clutter.

The scale as it stands:

| Priority | Meaning |
|---|---|
| High | Reality differs from intent. A test-named object on All Devices/All Users, or a group both included and excluded in one assignment intent. Correct these, never delete them. |
| Medium | Unassigned or exclusion-only, nothing protecting it. The actual cleanup queue. |
| Low | Unassigned but referenced (dependency/supersedence), or a test-named object scoped to a group. Usually keep. |
| Watch | An unassigned application created within the grace window. Parked, not actionable. |
| (blank) | Assigned and healthy. |

Read top-down, the scale is: things that are actively wrong first, then inert objects
ranked by how confidently they can be called abandoned. Only High can affect a device.

Actionable, on the Summary and in the counts, means High plus Medium only. Low is
"keep or confirm" and Watch is parked. An early version counted Low as actionable, which
put rows whose recommended action was literally `Keep` into the cleanup queue.

Flagged Entra groups are written at Medium unconditionally. The Intune scale does not
map onto a group, and a group that is empty or referenced by nothing is either worth a
look or it is not.

---

## Test-object detection ships switched off

`-TestNameRegex` has no default. Without it, `TestNamed` and `TestNamedBroadAssign`
never fire and the run says so on screen and on the RunInfo sheet.

An earlier version defaulted to `-TEST$`, which was the author's own convention. That is
worse than no default. A tenant that names things `TEST-Wifi` gets a report with no test
findings, which does not read as "this check did not run", it reads as "you have no
leftover test objects". Silent under-reporting on the one category that can affect a
device is the worst failure this tool could have, so the check now runs only when the
operator states their convention.

Two implementation notes for anyone editing this:

- `'anything' -match ''` is `$true` in PowerShell. An empty pattern reaching `-match`
  would flag every object in the estate as a test object. The call is guarded on
  `$TestNameRegex` being non-empty, and that guard must stay.
- The pattern is compiled with `[regex]::new()` at the top of the script so a malformed
  expression fails immediately with a clear message instead of part way through the
  inventory.

The parameter used to be called `-TestSuffixRegex`. The name presumed a suffix; plenty
of shops use a prefix or a bracketed tag.

---

## Platform filtering is an allow-list, not a substring guess

Every object is classified as `windows`, `other`, or `unrecognised`, in that order of
testing. Only `windows` is collected. `unrecognised` is counted per type and printed as
one warning at the end of the run.

Applications are matched against an explicit list of Windows `@odata.type` values.
Profiles and compliance policies are matched by pattern, because the Windows type list
there is long and grows.

The original filter was a substring match, `win32|windows|winget|officeSuiteApp|
microsoftStoreForBusiness`, with a `-notmatch` on `macos|ios|android`. It was correct by
a narrow margin: `macOSOfficeSuiteApp` matched `officeSuiteApp` and was saved only by
the exclusion running second. Any type Microsoft adds is either silently included or
silently dropped depending on which substrings it happens to contain, and the operator
never finds out. Testing the other-platform pattern first removes that dependency on
ordering luck.

Fail closed applies to the unrecognised bucket: those objects stay out of the report,
because guessing a platform in a report whose output is a removal queue is exactly the
guess that could get something deleted. Naming them in a warning keeps them visible.

`webApp` currently lands in the unrecognised bucket. A web link carries no platform, so
whether it belongs in a Windows report is a judgement call rather than a lookup, and it
is left for a later version rather than decided silently here.

---

## Why `lastModifiedDateTime` is not a staleness signal

An earlier version flagged everything not modified in six months. It produced roughly
four times more rows than the real cleanup list and buried the signal. It was removed.

Three reasons the timestamp is misleading:

1. **It measures content edits, not use.** A compliance policy assigned to ten thousand
   devices and untouched for a year is doing its job perfectly.
2. **Assignment changes do not update it.** Verified in a live tenant: changing a
   policy's assignment from a group to All Devices left `lastModifiedDateTime`
   unchanged. So you cannot tell when something was unassigned from object metadata.
   The audit log (`auditEvents`) is the only reliable source for that.
3. **Publishing tools re-touch objects.** Tooling that republishes application metadata
   bumps the timestamp on apps you never edited.

`LastModified` is still shown as a column, and it is used to sort the Medium bucket
(oldest content first, as a confidence heuristic). It is never used to flag anything.

---

## Include/exclude overlap must be intent-aware

If the same group appears as both an include and an exclude target on one object, the
exclusion wins and that group is silently skipped. That is a real configuration mistake
and it is flagged High.

**But application assignments carry an intent** (`required`, `available`,
`uninstall`), and reusing a group across different intents is a deliberate, common
pattern. Excluding a group from Required while including it in Uninstall is how you
actively remove an app from a set of devices.

The first implementation compared a flat include list against a flat exclude list and
produced a false positive on exactly that pattern. The check now buckets targets by
intent and only flags a collision **within the same intent**. Objects that have no
intent (everything that is not an app) fall into a single bucket and behave as before.

This was caught by a human reviewing a flagged item against the portal. It is the best
argument in the whole project for keeping a person in the loop.

**How rare is this?** The Intune portal now prevents it: pick a group as an include and
it is greyed out in the exclude list for the same assignment. So the check cannot fire on
anything created the normal way. What it can still catch is an assignment created through
Graph directly, which is how automation creates them and where nobody is watching, and
legacy configuration predating the portal validation.

That makes it defensive rather than routine, and it is kept for that reason: the cost is
one comparison per object, and the failure it describes is silent. An assignment that
targets nothing looks healthy in every view Intune offers.

The same applies to `OnlyExclusions`. The portal refuses to save an assignment that has
exclusions and no includes: "At least one group must be included when excluded groups
were selected." So that state, too, comes from Graph or from history rather than from
someone clicking through a wizard today.

Both flags are therefore worth least in a tenant built recently through the portal, and
worth most in one with years of accumulated configuration and tooling that assigns
objects programmatically. That is the estate this tool is for, so both stay.

---

## Unassigned applications and the grace window

Unassigned applications are the noisiest category, because two entirely legitimate
cases look identical to abandoned clutter: a package created recently and not yet
assigned, and a previous version retained so a bad release can be rolled back.

The clean way to detect a retained version is Intune **supersedence**, and the script
reads `supersedingAppCount` and `dependentAppCount` off the app object for exactly this.
But not every publishing workflow creates supersedence relationships, and where none
exists there is no link to follow.

The fallback is creation date. An unassigned application created within
`-NewAppGraceMonths` is flagged `RecentlyCreated`, set to `Watch`, and excluded from
both the actionable count and the worklist.

The window is the operator's decision, not a built-in assumption about anyone's release
cadence. Set it to roughly the age of the oldest version you expect to retain. Setting
it to `0` disables the grace window entirely and flags every unassigned app.

The parameter was originally called `-StaleMonths`, which was misleading: it has nothing
to do with staleness and is not applied to any object type other than applications.

### The name-and-version fallback

Creation date alone was not enough. Measured against a real tenant of 698 Windows apps
with a publishing tool that does not create supersedence: 217 retained versions fell
inside the six-month window and were parked correctly, but 66 aged past it into
`Medium`, and 57 of those had a newer version of the same app that was still assigned.
The tool was recommending removal of live rollback copies, at scale.

So an unassigned app is now flagged `SupersededByName` and dropped to Low when all of
the following hold:

- it has no assignments at all
- its display name contains a parseable dotted version
- another app shares its base name, has a higher version, and has a live inclusion
  assignment

Base name is the display name with dotted-numeric tokens removed. Architecture and
edition markers contain no dotted number and therefore survive, which keeps x64 and x86
packages from matching each other.

Every part of that is chosen to fail in the safe direction. Requiring the newer sibling
to be assigned means two unassigned versions of a retired application both stay in the
cleanup queue, which is right. Requiring a parseable version means an unversioned name
is never demoted. Over-detection tells you to keep something you wanted gone, which
wastes a minute; under-detection tells you to delete something you needed, which does
not.

### Retained versions expire

The first version of this rule demoted a name-matched copy forever, which was wrong for
the reason the rule exists at all. The justification for keeping N-1 is that a bad update
can be rolled back. That justification decays: nobody rolls back to a build from two
years ago, so an ancient N-1 is not a rollback copy, it is clutter that happens to have a
sibling.

`-RetainedVersionMonths` is therefore separate from `-NewAppGraceMonths`. They answer
different questions. The grace window asks "was this created too recently to judge". The
retention window asks "is this rollback still worth holding". Conflating them into one
number would force an operator to pick a value that is wrong for one of the two.

Past the retention window the object is flagged `RetainedVersionExpired`, stops counting
as referenced, and returns to `Medium` with an explicit reason. It keeps the
`SupersededByName` flag as well, so the operator can see it is a known previous version
rather than a mystery.

A genuine Intune supersedence relationship is exempt. That link is configuration the
newer application depends on, so age does not make removal safe.

**Known trade-off:** an application that has received no new version in longer than the
window, and no assigned newer sibling, will still age into `Medium` and appear for
review. That is intentional. A package chain that stale is worth a look.

---

## Why the output is a workbook and not CSV

CSV was considered specifically to drop the `ImportExcel` dependency, and rejected. The
Worklist is not a data dump, it is a form: `Decision`, `Owner`, `DateActioned` and
`Notes` are typed in by a person, and the sheet has to be filterable, shareable and
returnable. Splitting ten sheets into ten files loses the filter, the frozen header, and
any chance that a decision written into it survives the next run.

The dependency itself is light. `ImportExcel` needs no Excel installation, no admin
rights, and has no dependencies of its own.

One thing to keep an eye on when pinning a version: `ImportExcel` bundles EPPlus.
EPPlus 4.5.3.3 was the last LGPL release, and from version 5 it moved to Polyform
Noncommercial, which requires a commercial licence for commercial use. Whatever this
project depends on has to stay usable by the enterprise admins it is written for, so if
a future `ImportExcel` release ships EPPlus 5 or later, that is a decision to make
deliberately rather than to inherit from a version bump.

Where an environment genuinely cannot install the module, the answer is an explicit
CSV output option, not making CSV the default and degrading the tracker for everyone.

---

## Why PowerShell 7 is the floor

The original target was Windows PowerShell 5.1 and PowerShell 7 equally. Testing killed
that.

5.1 runs on .NET Framework, which permits one version of an assembly per process with no
isolation between modules. On an ordinary admin workstation carrying several
`Microsoft.Graph.*` versions side by side, they share a single
`Microsoft.Identity.Client`, and `Connect-MgGraph` fails with `Could not load type
'Microsoft.Identity.Client.AuthScheme.TokenType'` before the tool does anything.
Pinning the module version with `-RequiredVersion` did not help; the assembly is already
resolved for the process. PowerShell 7 loads the SDK's dependencies in an isolated
context, and the same command on the same machine worked first time.

The only 5.1 fix is uninstalling Graph modules the admin uses for other work, which is
not a reasonable prerequisite for a read-only report. `#Requires -Version 7.0` states
it, and the error message translation names PowerShell 7 explicitly when the assembly
signature appears.

The three 5.1 landmines below are kept anyway. They cost nothing, and the CI job still
parses under both hosts, so the option to reverse this stays open.

---

## Referencing is tenant wide, reporting is Windows only

A group referenced only by a macOS profile is still a group in use. The first version
collected referenced group IDs from objects that had already passed the Windows filter,
so those groups came back flagged with a suggested action of Remove. For a
tenant with a separate macOS or mobile estate, that is the tool telling you to delete
groups that are actively deploying software.

Group IDs are now harvested from the raw result of every endpoint, before filtering.
Reporting stays Windows only; referencing does not.

That was necessary but not sufficient. Several object types are not collected for the
report at all, and a group used only by one of those still looked unreferenced. Measured
in a live tenant: of six flagged groups, one was targeted by a macOS shell script, an
endpoint the tool never read. The report was pointing at a group that was deploying
scripts to Macs.

So a second list exists, `$script:ReferenceOnlyEndpoints`: twelve endpoints that are
read for their assignments and never reported on. macOS shell and custom attribute
scripts, Autopilot, enrolment configurations, the three update profile types, app
configuration and app protection policies. They are only queried when the group section
is going to run, so an Intune-only run does not pay for them.

That still is not a proof. Any object type Microsoft adds is invisible until someone
adds it to the list, so group rows remain `Investigate` rather than `Remove`, and the
reason text says to check the portal.

The same reasoning drives the incomplete-read rule. If any Graph GET fails, the
referenced-group set is partial, so a group could look unreferenced purely because the
assignment naming it was never read. The group section is skipped entirely in that case
rather than reported with a caveat.

---

## PowerShell 5.1 landmines

Three bugs cost real debugging time. If you are editing this, avoid repeating them.

**`@()` over a `List[object]` can throw.** Wrapping a generic list in the array
subexpression operator produced `ArgumentException: Argument types do not match` on
PS 5.1, and it moved as the code changed, which made it look like several different
bugs. All aggregation now iterates the list directly or reads `.Count` from it. Do not
reintroduce `@($someList)` on the row collections.

**`[Parameter(Mandatory)]` rejects an empty collection.** A function that fills a
caller-supplied list cannot mark that parameter mandatory, because the list is empty on
the first call and the binder refuses it before the function runs. Use
`[ValidateNotNull()]` instead.

**Wildcard `switch` falls through.** A `switch -Wildcard` on assignment target types
matched both `*exclusionGroupAssignmentTarget` and `*groupAssignmentTarget` for the same
target, quietly counting exclusions as inclusions. Rewritten as explicit `if/elseif`
with the more specific pattern tested first.

**ASCII only.** The script contains no non-ASCII characters: plain hyphens, straight
quotes. Smart quotes and dashes survive an editor but break when a file round-trips
through sync clients, other editors, or a deployment pipeline. Keep it ASCII.

---

## Authentication and Conditional Access

Two things that will bite you in a hardened tenant:

**Token protection requires brokered sign-in.** If Conditional Access enforces token
protection, a Graph sign-in that does not go through the Windows broker (WAM) fails with
`AADSTS530084`.

An earlier version called `Set-MgGraphOption -EnableLoginByWAM $true` before connecting.
That call is now obsolete: current releases of `Microsoft.Graph.Authentication` enable
broker sign-in by default on Windows and cannot have it disabled except through
`-DisableLoginByWAM` with a custom client ID, and Microsoft documents the old switch as
having no effect. The call has been removed rather than left in a `try/catch`, because
a swallowed failure printed a warning telling the user to update the module, which is
the opposite of the real situation.

What still matters, and what the earlier version got wrong by omission, is the app
registration. Brokered sign-in against your own client ID needs the redirect URI
`ms-appx-web://Microsoft.AAD.BrokerPlugin/<client id>` under Mobile and desktop
applications. The device must also be registered or joined and compliant, or token
protection fails regardless.

The failure is confusing because a cached token keeps working. Only a *fresh* sign-in
trips the policy, so it can look like one specific scope is broken when it is really
the whole authentication path.

**Intune permissions do not imply one another.** `DeviceManagementScripts.Read.All` is
required for remediations and platform scripts, and is not covered by
`DeviceManagementConfiguration.Read.All`. Found the hard way on a clean test tenant
following the README: both endpoints returned 403 while everything else worked, and the
report came back with those categories empty, which looks exactly like a tenant that has
none. The permission also covers the macOS shell and custom attribute scripts read for
group references, so its absence quietly weakens the group check too.

That is the argument for the scope check running before collection and for the
end-of-run warning when any read failed. A partial report that looks complete is the
worst output this tool can produce.

**Do not pass `-Scopes` with a custom `-ClientId`.** MSAL treats requested scopes on a
custom app registration as a new authorization and triggers a consent prompt. The token
must instead carry whatever is already consented on the app. The script therefore
connects without `-Scopes` and then *verifies* the granted scopes, warning per missing
permission rather than trying to force them. If a scope is missing unexpectedly, the
cause is usually a cached session predating a consent change: `Disconnect-MgGraph` and
re-run.

The scope check is conditional. `Group.Read.All` and `User.Read.All` are only required
when `-GroupOwnerUpns` is supplied, so an Intune-only run does not warn about them.
Warning about permissions a run does not need is the same failure mode as warning
spam: it trains the operator to ignore warnings.

The flag is called `NoAssignmentFound`, not `NotUsedInIntune`. Everything else this tool
reports is a fact read straight off an object: this policy has no assignments, this app
has a newer version, this name matches your pattern. A claim that a group is unused is
different in kind. It is a negative about the whole tenant, and establishing it means
reading every object type that could reference a group, which today is twelve endpoints
beyond the ones reported on, arrived at only because a live tenant caught the list
missing macOS shell scripts. Any object type Microsoft adds is a silent false positive
until someone adds it here.

So the flag states what the tool did rather than what the operator might conclude from
it. The suggested action is `Investigate`, never `Remove`, and the reason text says to
check the portal first.

---

## The settings file

Identifiers and preferences live in `%APPDATA%\IntuneHousekeeper\settings.json`, written
by `Set-IntuneHousekeeperConfig`. Four decisions in it are worth explaining.

**Nothing secret is stored, and nothing secret ever should be.** Sign-in uses a public
client flow, which has no secret, so there is nothing to put there. If someone later adds
certificate or client-secret support, that credential does not belong in this file: a
settings file holding a credential is a credential store with none of the protections one
needs, and it will end up in a screenshot or a repository.

**`-ClientId` and `-TenantId` are not `Mandatory`.** Mandatory binding happens before the
function body runs, so the binder would prompt for exactly the values the settings file
already holds, which defeats the point. They are ordinary optional parameters, checked
after the merge, and the error names `Set-IntuneHousekeeperConfig`. The cost is the
`HelpMessage` prompt text, which only ever appeared at the mandatory prompt. An error
message everyone sees is worth more than prompt guidance most people never trigger.

**Precedence is explicit parameter, then stored value, then default, decided with
`$PSBoundParameters`.** Nothing else can distinguish `-NewAppGraceMonths 6` from the
default of 6, because a parameter with a default always has a value. Get this wrong and a
saved setting silently overrides what the operator just typed, in a tool whose output is
a deletion queue.

**A stored `0` or empty string is a setting, not an absence.** `"NewAppGraceMonths": 0`
means no grace window and `"TestNameRegex": ""` means detection off; only a missing key
falls through to the default. The file is written sparsely for that reason, holding only
the keys that have been set.

A stored value is assigned through the parameter variable, which re-applies its type and
any `ValidateRange`, so a hand-edited file with an impossible value fails with a message
naming the file and the setting rather than misbehaving later.

---

## Scoping the Entra group check

The group section only reports groups that are **owned by a listed account** and
**match a naming prefix**. Both filters matter.

Owner scoping alone is not enough: people own Teams groups, Microsoft 365 groups, and
distribution lists, and "not used in Intune" is a meaningless statement about a Teams
group. Without the name filter the report fills with noise.

Both are parameters with no defaults (`-GroupOwnerUpns`, `-GroupNamePrefix`). If you
leave them empty the section is skipped entirely, which is the correct fail-closed
behaviour: better to report nothing than to report nonsense.

The corollary is a real limitation: an assignment group that deviates from your naming
convention will never be reported.

---

## What it deliberately does not do

- **Detect assignments targeting empty or deleted groups.** These look healthy while
  deploying to nothing, which is arguably worse than being unassigned. It needs
  resolution of every referenced group ID and is the most valuable thing to add next.
- **Evaluate assignment filters.** An object assigned through a filter that matches
  nothing is reported as healthy. Same class of blind spot as the one above.
- **Detect a wrongly scoped object with a conventional name.** There is no metadata
  signal for intent beyond a naming convention.
- **Remember decisions between cycles.** Deliberate. A "keep" decision six months ago
  is not evidence for today, and carrying it forward would let objects accumulate
  permanent exemptions without review.
- **Use AI to suggest dispositions.** The `SuggestedAction` column comes from a fixed
  rule tree in the script, so it is auditable and reproducible. An LLM was prototyped
  for this and removed: the logic turned out to be deterministic rules that needed no
  model. AI is better spent on the genuinely fuzzy judgement ("is this project
  abandoned?") than on rule-following.
