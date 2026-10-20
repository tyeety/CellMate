# Changelog

User-visible changes are recorded here. Versions follow `MAJOR.MINOR.PATCH`.
Dates use `YYYY-MM-DD`. Keep future work under **Unreleased** until shipped;
move completed entries into a dated release section when publishing that version.

## [Unreleased]

No changes recorded yet.

## [0.2.0] - 2026-10-20

### Added

- Standalone graphical Windows installer with both 32-bit and 64-bit add-ins embedded.
- Automatic Excel architecture detection and installation for the current user without administrator access.
- Waiting for Excel to close before installation or removal, without closing user workbooks automatically.
- Registration in Windows Settings > Apps, with an uninstaller that retains personal shortcuts.
- Versioned installation paths, legacy preference migration, installation verification, and rollback on installation failure.
- Installer checks covering payload selection, migration, unrelated settings, repair, rollback, and removal.
- Public source-viewing repository, complete usage guide, license, and third-party notices.

### Changed

- Shared purple, light-background design for Setup and add-in settings, based on ShatootSoft's visual language.
- Updated CellMate icon and Windows application icon.

### Licensing

- Free personal and commercial use of official binaries under the CellMate Binary Use and Source Viewing License 1.0.
- Original source is viewable, with modification and reuse reserved to the copyright holder except mandatory platform or legal rights.

## [0.1.1] - 2026-10-20

### Changed

- Renamed RowMate to CellMate, including commands, installation paths, and settings paths.
- Switched the complete interface and messages to English with left-to-right layout.
- Preserved existing shortcut settings during migration from RowMate.

### Removed

- Duplicate Row ribbon button; duplication remains available through its configurable shortcut.

## [0.1.0] - 2026-10-20

### Added

- Initial local RowMate version, subsequently renamed CellMate.
- Duplicate current or selected rows with formulas, formatting, heights, and hidden states preserved.
- Support for contiguous and disjoint row selections, deduplicating overlapping row selections.
- Immediate undo for the latest row duplication.
- Paste Values and AutoFit Columns commands.
- Customizable shortcuts with duplicate-assignment validation, disabling, and saved preferences.
- Excel ribbon, add-in icon, and per-user installation script.

### Fixed

- Preserved existing Excel Options values when adding the startup registration.

## Release procedure (authorized maintainers)

1. Record completed changes under `Unreleased` using applicable categories: Added, Changed, Fixed, Removed, Security, or Licensing.
2. Choose a version and update `CellMate.csproj`, `Setup/CellMate.Setup.csproj`, `Setup/app.manifest`, and `Theme.Version` in `Theme.cs` consistently.
3. Move the entries to a dated version section and reset `Unreleased`. Do not label planned work as shipped.
4. Run `Build.ps1` and the installer checks documented in the README. Review the actual installer and Excel interface before claiming visual acceptance.
5. Commit, tag `vMAJOR.MINOR.PATCH`, and publish a GitHub Release containing the official installer package and its matching license/notices.
6. Record relevant verification limits in the release notes. Never attach a binary built from a different source revision.

[Unreleased]: https://github.com/tyeety/CellMate/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/tyeety/CellMate/releases/tag/v0.2.0
