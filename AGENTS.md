# Working on this repo with an AI assistant

Context for Claude Code, Copilot, or any assistant pointed at this repository. Read
`docs/DESIGN-NOTES.md` first: it explains why the non-obvious decisions were made, and
most of them came from something breaking in a real tenant.

## What this is

Intune Housekeeper: a read-only PowerShell module that inventories Windows Intune
objects via Microsoft Graph and writes an Excel worklist. Three exported commands:
`Export-IntuneHousekeeperReport`, `Get-IntuneHousekeeperConfig` and
`Set-IntuneHousekeeperConfig`. Everything else in the .psm1 is internal. There is no
build step and no test suite, and no dependencies beyond the two modules named in the
manifest.

Module state lives at script scope in the .psm1 and therefore persists for the whole
session. `Reset-RunState` is called at the top of the exported function for that reason.
Any new counter or cache must be reset there too, or a second run in one session
inherits the first run's data.

## Hard constraints

These are not preferences. Breaking any of them breaks the tool.

1. **Read-only.** GET requests only. No PATCH, POST, or DELETE against Graph. The
   safety property of this tool is that running it cannot damage anything. If you are
   asked to add write operations, make them opt-in, explicit, and loud.

2. **ASCII only.** The script file must contain no bytes above 127. Plain hyphens,
   straight quotes. Non-ASCII characters survive an editor and then break when the file
   passes through sync clients or deployment pipelines. Verify before committing:
   ```bash
   python3 -c "print(sum(1 for b in open('IntuneHousekeeper/IntuneHousekeeper.psm1','rb').read() if b>127))"
   ```
   That must print `0`.

3. **PowerShell 7 is the supported floor**, set by `#Requires -Version 7.0`. Windows
   PowerShell 5.1 fails at `Connect-MgGraph` on any machine with several
   `Microsoft.Graph.*` versions installed, because .NET Framework cannot isolate the
   shared `Microsoft.Identity.Client`. Do not remove the `#Requires` line without
   retesting that.

   Keep the code free of 7-only syntax anyway, and keep the 5.1 parse job in CI. The
   three traps below cost nothing to avoid and preserve the option to reverse this:
   - Do not wrap a `List[object]` in `@()`. It throws `ArgumentException: Argument
     types do not match` on 5.1. Iterate the list directly or read `.Count`.
   - Do not mark a collection parameter `[Parameter(Mandatory)]` if the caller passes
     an empty list to be filled. The binder rejects empty collections. Use
     `[ValidateNotNull()]`.
   - Do not use `switch -Wildcard` for assignment target types. Patterns overlap and
     fall through. Use explicit `if/elseif`, most specific first.

4. **No organisation-specific values.** No company names, no vendor or product names in
   rationale comments, no real UPNs, no internal naming conventions as defaults, no
   numbers from a real tenant. Anything site-specific is a parameter that defaults to
   empty or neutral. Where a behaviour exists to accommodate a particular workflow,
   describe the workflow generically and let the operator set the parameter.
   The one fixed identifier in the code is Microsoft's own: the well-known client ID of
   Microsoft Graph Command Line Tools, used by `-UseGraphPowerShellApp`. It is the same
   in every tenant, so it is not a site-specific value and must stay in the code.

5. **Fail closed.** When something is ambiguous, flag it for human review or skip it.
   Never guess in a direction that could cause an object to be deleted. If a permission
   is missing, skip that section with one clear message rather than emitting garbage or
   spamming warnings. Do not warn about a permission a given run does not need.

## Design invariants

- `Priority` means order of cleanup action, not risk to devices. Only `High` describes
  something that can actually affect a device.
- Informational flags (`DuplicateName`) must never raise `Priority` on their own.
- `lastModifiedDateTime` must never be used to flag an object. It measures content
  edits only; assignment changes do not update it. It may be displayed and used as a
  sort key.
- The include/exclude overlap check must stay **intent-aware**. Reusing a group across
  different app assignment intents is deliberate and must not be flagged.
- The `Worklist` sheet is the only place decisions are recorded. Category sheets are
  read-only inventory.
- Actionable means High plus Medium. Never count `Low` or `Watch`: their recommended
  action is to keep or to wait.
- Referenced-group collection happens on the RAW result of each endpoint, before the
  Windows filter, and additionally from `$script:ReferenceOnlyEndpoints`. Adding a new
  object type to Intune means adding it there, or groups it targets will be reported as
  unreferenced. A group used only by a macOS or mobile assignment must never be
  reported as unreferenced. Group rows are `Investigate`, never `Remove`, because
  several object types are not read at all.
- If any Graph read fails, the Entra group section is skipped. An incomplete
  referenced-group set makes 'nothing references this' unsafe to assert.
- Retained-version detection must fail toward keeping. Requiring a parseable version and
  an assigned newer sibling is deliberate; loosening either risks recommending deletion
  of a live rollback copy.
- The settings file holds identifiers and preferences only. No client secret, no
  certificate, no token. A settings file that holds a credential is a credential store
  without any of the protections one needs, and the public client flow means there is
  nothing to put there anyway.
- Settings precedence is explicit parameter, then settings file, then default, decided
  with `$PSBoundParameters`. A parameter with a default always has a value, so nothing
  else can distinguish `-NewAppGraceMonths 6` from the default of 6. Getting this wrong
  lets a saved value silently override what the operator just typed.
- A stored `0` or empty string is a setting, not an absence. Only a missing key falls
  through to the default.
- `-ClientId` and `-TenantId` must not be `Mandatory`: binding happens before the
  function body runs, so the prompt would fire for values the settings file already has.
- `-NewAppGraceMonths` is the operator's call, including `0`. Do not reintroduce a
  hardcoded assumption about how long retained application versions live.
- No naming convention ships with a default. `-TestNameRegex` and `-GroupNamePrefix` are
  empty, and the check they drive is skipped and reported as skipped. A default
  convention produces a silent false negative that reads as a clean result.
- `-GroupOwnerUpns` without `-GroupNamePrefix` skips the Entra group section, decided
  right after the settings merge and before sign-in, whether the owners came from a
  parameter or the settings file. A skipped run behaves exactly like one without owners:
  no group scopes requested or checked, no reference-only endpoints read, one
  `Write-Warning` naming the missing parameter and nothing more about the section. No
  fallback: no default prefix, and no "security groups only" mode, which would move the
  risk onto Conditional Access and licensing groups. A skip is held as a reason and
  never applied by emptying `$GroupOwnerUpns`, so RunInfo and later messages report why.
- `$TestNameRegex` must never reach `-match` while empty: an empty pattern matches every
  string and would flag the whole estate.
- Object types that are neither recognised as Windows nor as another platform are left
  out of the report and named in one warning. Do not widen a filter by guessing at
  substrings; add the type to the allow-list or leave it reported as unrecognised.

## Sign-in paths

There are two ways to sign in, and they are deliberately not symmetrical.
`docs/DESIGN-NOTES.md`, "Signing in without an app registration", has the reasoning.

- An own app registration (a custom `-ClientId`) is the recommended path. The built-in
  path (`-UseGraphPowerShellApp`, Microsoft Graph Command Line Tools) is for trying the
  tool. Never make the built-in path the fallback for a missing `-ClientId`: an omitted
  parameter must not silently change which app the operator consents to.
- One setting, not two. The switch is stored and resolved as `ClientId` set to the
  well-known ID, and the mode is derived from the effective client ID. Do not add a
  separate settings key: precedence is resolved per key, and two mutually exclusive keys
  cannot be resolved that way.
- Resolve the switch after the settings merge. Before it, the merge loop sees `ClientId`
  as unbound and overwrites it with the stored value. Passing both
  `-UseGraphPowerShellApp` and `-ClientId` explicitly is an error, checked by hand
  rather than with parameter sets, so the parameter block stays flat.
- `-Scopes` is never passed with a custom client ID. On the built-in path it is always
  passed, and contains only the scopes this run needs. The consent handoff block printed
  on failure always lists all six. These rules look inconsistent and are not: do not
  align them.
- `-TenantId` is required on both paths. Do not make it optional on the built-in path:
  with no tenant, Windows offers the account the device is signed in with, which on a
  work machine is the production tenant.
- A matching session the operator opened is reused and never closed. Only a session this
  run opened is disconnected at the end.
- RunInfo records the sign-in mode by name only. No client ID or tenant ID is ever
  written to the workbook.
- The sign-in mode line is `Write-Host`, not `Write-Warning`. Warnings in this tool mean
  something to fix, and a deliberate quick-start run is not that.
- A 403 does not mean a missing scope. Delegated sign-in acts as the signed-in account,
  so role and scope both apply on either path. Do not write a message that assumes every
  403 is a permission missing from the registration.

## Style

- Comment the *why*, not the *what*, especially where behaviour looks wrong but is
  deliberate.
- Keep the parameter block flat and documented in both the comment-based help and the
  README table.
- Prefer adding a parameter over hardcoding a site-specific assumption.
- Dependencies must remain usable in a commercial environment. `ImportExcel` bundles
  EPPlus, which changed from LGPL to Polyform Noncommercial at version 5. Check the
  bundled DLL version before raising the required `ImportExcel` version.
- Update `docs/DESIGN-NOTES.md` when you make a decision a future reader would question.

## Verifying a change

There is no test harness. Before committing:

1. ASCII check (above) returns `0`.
2. Bracket balance is even.
3. The file parses, the manifest is valid, and the module imports:
   ```powershell
   $errors = $null
   [System.Management.Automation.Language.Parser]::ParseFile(
       (Resolve-Path ./IntuneHousekeeper/IntuneHousekeeper.psm1), [ref]$null, [ref]$errors)
   $errors
   Test-ModuleManifest ./IntuneHousekeeper/IntuneHousekeeper.psd1
   Import-Module ./IntuneHousekeeper/IntuneHousekeeper.psd1 -Force
   Get-Command -Module IntuneHousekeeper
   ```
   All of this runs in CI on every push.
4. Ideally, run it against a real tenant. It is read-only, so this is safe.
5. Changes to sign-in cannot be verified by CI or by stubs alone. Run both paths against
   a real tenant before release, including one run where the sign-in is cancelled. Any
   test that matches on a sign-in error message must use text copied from a real run: a
   hand-written fixture is how the first cancel check passed its test and missed the
   real message.
