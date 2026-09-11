# Changelog

Notable changes to Intune Housekeeper. Versions follow semantic versioning.

## 1.0.0

First public release.

Intune Housekeeper inventories a Windows Intune estate through Microsoft Graph and writes
an Excel decision tracker: a pre-sorted worklist of objects worth cleaning up, each with a
plain-language reason and a suggested action. It issues GET requests only and never
changes a tenant.

- Covers applications, configuration profiles (templates, Settings Catalog, ADMX),
  compliance policies, security baselines, remediations, and platform scripts.
- Flags unassigned and exclusion-only objects, test-named objects assigned broadly, and
  groups both included and excluded within one assignment intent.
- Recognises retained previous versions of applications, through Intune supersedence where
  it exists and by display name and version where it does not, and expires them after a
  window you choose.
- Optional Entra assignment group check for owned groups that are empty or that no
  assignment references. Group references are established tenant wide, including from
  object types the report does not cover.
- Settings can be saved once with `Set-IntuneHousekeeperConfig` instead of being retyped.
- Priority is colour-shaded in the workbook so a long sheet can be scanned at a glance.
- Requires PowerShell 7 on Windows.
