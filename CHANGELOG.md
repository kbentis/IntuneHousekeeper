# Changelog

Notable changes to Intune Housekeeper. Versions follow semantic versioning.

## Unreleased

Pre-release development, validated against a production tenant of roughly 900 Windows
objects. Changes worth recording because each came from something the tool got wrong:

- Retained previous versions of applications are recognised by display name and version
  where no Intune supersedence relationship exists, and expire after
  `-RetainedVersionMonths` so an ancient rollback copy returns to the cleanup queue.
- Group references are collected tenant wide, including from twelve object types the
  report does not cover, so a group used only by a macOS shell script is no longer
  reported as referenced by nothing.
- The Entra group section is skipped entirely if any Graph read failed, because a
  partial reference set makes "nothing references this" unsafe to assert.
- Actionable means High plus Medium. Low and Watch are excluded from the counts.
- Test-object detection has no default naming convention and reports how many names it
  matched, so a wrong pattern is visible rather than reading as a clean estate.
- PowerShell 7 is required. Windows PowerShell 5.1 fails at sign-in on any machine with
  several Microsoft.Graph module versions installed.
- An existing Graph session is reused, and only a session the script opened is closed.
- `DeviceManagementScripts.Read.All` documented and checked. Without it, remediations and
  platform scripts return 403 and come back empty while the rest of the report looks fine.
- A run where any Graph read failed now warns clearly that the report is incomplete,
  rather than only recording it on the RunInfo sheet.
- Settings file support. `Set-IntuneHousekeeperConfig` saves identifiers and
  preferences, `Get-IntuneHousekeeperConfig` shows them, and the report command reads
  them. Explicit parameters still win.
- Packaged as a module. `Install-Module IntuneHousekeeper`, then
  `Export-IntuneHousekeeperReport`. Dependencies are declared in the manifest rather
  than checked at runtime.

## 1.0.0

First public release. Entry to be written at publish time.
