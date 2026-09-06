# Working on this repo with an AI assistant

Context for Claude Code, Copilot, or any assistant pointed at this repository. Read
`docs/DESIGN-NOTES.md` first: it explains why the non-obvious decisions were made, and
most of them came from something breaking in a real tenant.

## What this is

Intune Housekeeper: a read-only PowerShell module that inventories Windows Intune
objects via Microsoft Graph and writes an Excel worklist. One exported command,
`Export-IntuneHousekeeperReport`; everything else in the .psm1 is internal. There is no
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
- `-NewAppGraceMonths` is the operator's call, including `0`. Do not reintroduce a
  hardcoded assumption about how long retained application versions live.
- No naming convention ships with a default. `-TestNameRegex` and `-GroupNamePrefix` are
  empty, and the check they drive is skipped and reported as skipped. A default
  convention produces a silent false negative that reads as a clean result.
- `$TestNameRegex` must never reach `-match` while empty: an empty pattern matches every
  string and would flag the whole estate.
- Object types that are neither recognised as Windows nor as another platform are left
  out of the report and named in one warning. Do not widen a filter by guessing at
  substrings; add the type to the allow-list or leave it reported as unrecognised.

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
