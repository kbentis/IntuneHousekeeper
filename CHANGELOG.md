# Changelog

Notable changes to Intune Housekeeper. Versions follow semantic versioning.

## Unreleased

Pre-release development, validated against a production tenant of roughly 1000 Windows
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

## 1.0.0

First public release. Entry to be written at publish time.
