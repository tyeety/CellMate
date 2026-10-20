# CellMate

**Small tools. Faster work.** An Excel desktop add-in by [ShatootSoft](https://shatootsoft.com), with configurable keyboard shortcuts and an English interface.

[Download the latest installer](https://github.com/tyeety/CellMate/releases/latest) · [Changelog](CHANGELOG.md) · [License](LICENSE) · [Report an issue](https://github.com/tyeety/CellMate/issues)

> **Free to use. Source available for viewing.** Official binaries are free for personal and commercial use. Original source is not licensed for copying, modification, building, or redistribution without written permission, except mandatory legal and hosting-platform rights. This is **not an open-source project**. Read [LICENSE](LICENSE) for the controlling terms and the GitHub viewing/forking exception.

## Requirements

- Windows 10 or Windows 11 with .NET Framework 4.8 or later.
- Microsoft Excel desktop for Windows, version 2016 or later, including Microsoft 365 desktop Excel.
- Either 32-bit or 64-bit Excel; Setup selects the matching add-in automatically.
- Installation is for the current Windows user and does not require administrator access.

Excel for the web, macOS Excel, and native ARM64 Excel are not supported. Excel itself must already be installed. Internet access is needed to download the installer, not for the add-in's commands.

## Install

1. Open [Releases](https://github.com/tyeety/CellMate/releases/latest) and download `CellMate-Setup-0.2.0.zip` (or the package for a newer release).
2. Extract the archive and run `CellMateSetup.exe`.
3. Click **Install CellMate**. If Excel is open, save your work and close all Excel windows. Setup waits and continues automatically; it never closes your workbooks for you.
4. When **Installed successfully** appears, click **Open Excel**.
5. Open or create a workbook. Find the **CellMate** ribbon tab.

The installer carries both add-in architectures, registers automatic loading, and preserves existing CellMate shortcuts. The current installer is unsigned; organizations that require signed executables or approved XLL add-ins should use their normal IT approval process. Do not disable Windows or Office protections to bypass an organizational policy.

The installer does not change global Office security settings or save your workbooks.

## Shortcuts at a glance

| Default shortcut | Command | Also on the ribbon? |
|---|---|---|
| `Ctrl+D` | Duplicate current or selected rows | No; shortcut only |
| `Ctrl+Shift+V` | Paste copied cell values | Paste Values |
| `Ctrl+Alt+O` | AutoFit selected columns | AutoFit Columns |
| `Ctrl+Alt+K` | Open shortcut settings | Settings |

Shortcuts apply across the open workbooks in the Excel instance where CellMate is loaded. Finish editing a cell before using a command. An enabled shortcut overrides Excel's normal action and may conflict with another add-in.

## Duplicate rows

### One row

Click any cell in the row, then press `Ctrl+D`. CellMate inserts a copy of the **entire row** directly underneath. You do not need to click the row header first.

For example, if A2 contains `Item A` and B2 contains `=A2`, duplicating row 2 creates row 3 with `Item A` in A3 and `=A3` in B3. Existing rows below move down.

### Several adjacent rows

Select cells spanning the rows you want, or select their row headers, then press `Ctrl+D`. The selected row block is copied below itself, in the same order.

Selecting A2:C4 duplicates whole rows 2-4 into new rows 5-7, including cells outside columns A-C.

### Separate row blocks

Hold `Ctrl` while selecting separate cells or row blocks, then use the duplicate shortcut. Each block is inserted directly below its own original. Selecting multiple areas on the same row duplicates that row only once. Adjacent selected rows are treated as one block.

CellMate preserves values, formulas using Excel's normal relative-reference adjustment, formatting, row heights, and hidden states. The inserted rows become selected. Manually hidden rows in the selected range are included; a sheet with an active filter is blocked to avoid accidental hidden-row duplication.

### Undo

Immediately press `Ctrl+Z` to undo the latest duplication. CellMate provides **one duplication undo operation**. Redo and a multilevel add-in undo history are not implemented. Add-in commands can clear Excel's existing undo history, so do not rely on earlier Excel undo entries remaining available.

### Limits

Duplication is blocked when:

- More than 10,000 rows are selected, or there is insufficient room at the bottom of the sheet.
- The worksheet is protected or multiple worksheets are grouped.
- A filter is actively hiding rows.
- A selected row contains merged cells.
- The selected rows or insertion boundary are inside or immediately beside a structured **Excel Table**.

A formatted range is supported. A structured Excel Table is a different feature and is not supported for duplication yet. CellMate displays a message before changing the sheet when these conditions are detected.

## Paste Values

1. Copy source cells **inside Excel** using `Ctrl+C`.
2. Select the destination cell or range.
3. Press `Ctrl+Shift+V` or choose **CellMate > Paste Values**.

Only values are pasted. Source formulas become their current results, and destination formatting is retained. Cut operations and arbitrary text or images copied from other applications are not supported by this command. Grouped or protected worksheets are blocked.

## AutoFit Columns

Select cells in the columns you want to resize. Press `Ctrl+Alt+O` or choose **CellMate > AutoFit Columns**. Excel adjusts the width of each selected column to fit its contents. Grouped or protected worksheets are blocked.

## Customize shortcuts

1. Choose **CellMate > Settings**, or press `Ctrl+Alt+K`.
2. Click the box next to a command.
3. Press the new combination. Include `Ctrl` or `Alt` with a letter A-Z, a number, or F1-F12. `Shift` may also be included. Some reserved combinations are rejected.
4. Click **Save changes**. The change takes effect immediately and persists between Excel sessions.

Press `Delete` or `Backspace` in a shortcut box to disable that binding. **Restore defaults** fills the default combinations; click **Save changes** to apply them. **Cancel** discards unsaved changes. A combination cannot be assigned to two CellMate commands.

**Ctrl+D normally means Fill Down in Excel.** While the CellMate binding is enabled it duplicates whole rows instead. Disable or change that binding to restore Excel's normal Ctrl+D behavior. Disabling a binding resets it to Excel's default; CellMate cannot recover a previous custom binding from a different add-in.

## Update

Download the installer package from a newer GitHub release and follow the same installation steps. Close Excel when asked. The installer reuses CellMate's startup registration, preserves shortcuts, verifies the selected binary, and attempts to roll back installation changes if a failure occurs. Running the same installer again can also reinstall/repair its version.

Earlier RowMate preferences are migrated if no CellMate settings already exist. Existing CellMate preferences take priority. Unrelated Excel startup entries and options are preserved.

## Uninstall

1. Save your work and close Excel.
2. Open **Windows Settings > Apps > Installed apps** (or **Apps & features** on Windows 10).
3. Find **CellMate**, choose **Uninstall**, and follow the removal window.

Alternatively, run `%LOCALAPPDATA%\CellMate\CellMateSetup.exe --uninstall`.

Your shortcut settings remain in `%APPDATA%\CellMate\settings.xml` so reinstalling can reuse them. Your workbooks are not deleted or modified. To reset personal settings completely, close Excel and rename or remove this settings file yourself.

## Local files and privacy

| Location | Purpose |
|---|---|
| `%LOCALAPPDATA%\CellMate\0.2.0\CellMate.xll` | Installed add-in; newer releases use their own version directory |
| `%LOCALAPPDATA%\CellMate\CellMateSetup.exe` | Installed uninstaller |
| `%LOCALAPPDATA%\CellMate\LICENSE` | CellMate license |
| `%LOCALAPPDATA%\CellMate\THIRD_PARTY_NOTICES.md` | Third-party terms |
| `%APPDATA%\CellMate\settings.xml` | Personal shortcut mappings |

Startup registration is stored under the current user's Office 16.0 Excel Options key. Windows Apps registration is also per user. CellMate's commands process cells locally; they do not upload workbook contents or send telemetry. Any independent behavior of Excel, Windows, or another add-in is outside CellMate's control.

## Troubleshooting

| Symptom | What to check |
|---|---|
| Setup says Excel was not found | Install the supported Windows desktop edition. The web and macOS editions do not load this XLL. |
| Setup keeps waiting for Excel | Save and close every Excel instance. If no windows remain, inspect Excel processes through your usual Windows tools; do not end a process containing unsaved work. |
| CellMate tab is missing | Restart Excel normally and open a workbook. Check **File > Options > Add-ins**, including disabled items, or reinstall using Setup. Follow any organization policy that blocks XLL add-ins. |
| Ctrl+D performs Fill Down | Open CellMate Settings and check that Duplicate selected rows is enabled. Another add-in may have reassigned the key. |
| Shortcut does nothing while typing | Finish or cancel editing the cell, then use the command. Check for duplicate mappings in other tools. |
| Duplication is blocked | Check the limits above: filters, Excel Tables, merges, grouped/protected sheets, row count, and space. |
| Paste Values asks for copied cells | Copy cells in Excel with Ctrl+C first; do not cut them or use a non-Excel clipboard source. |
| Settings cannot be read | Close Excel, back up and rename the settings XML, then reopen Excel to use defaults. The app also falls back to defaults when loading invalid preferences. |
| Installation fails | Read the message and resolve the cause before retrying. If rollback itself failed, retain the message and report it; do not assume installation was fully restored. |

Report the CellMate version, Excel version and bitness, Windows version, the message shown, and reproduction steps through [Issues](https://github.com/tyeety/CellMate/issues). Use a small synthetic workbook if an example is needed; do not post confidential data or credentials.

## Verification status for 0.2.0

The add-in's original row operations were exercised in real Excel, and installer checks cover architecture payloads, migration, preserving unrelated options, repair, injected-failure rollback, and removal. These checks do not establish compatibility with every Office configuration. Final visual acceptance of the 0.2.0 installer/settings was interrupted and is not claimed.

## Source and maintainer instructions

The source is published for inspection. The instructions below are for the copyright holder and separately authorized maintainers; they do not grant source-copying, build, modification, or redistribution rights. See [LICENSE](LICENSE).

### Build

Use Windows with a .NET SDK capable of building `net48` (this project is verified with SDK 10.0). NuGet restores the .NET Framework reference assemblies.

```powershell
.\Build.ps1
```

The script builds both packed XLL architectures, builds Setup with the payloads embedded, and packages the release into `dist/CellMate-Setup-0.2.0.zip`. Build outputs and personal settings are excluded from Git.

`Install.ps1` is a convenience launcher for the built graphical installer; `Install.ps1 -Uninstall` opens its removal window. End users should download the official binary release instead of building source.

### Installer checks

Run in Windows PowerShell 5.1 after building:

```powershell
[Reflection.Assembly]::LoadFrom((Resolve-Path '.\dist\CellMateSetup.exe').Path) | Out-Null
[CellMate.Setup.InstallerTests]::Run()
```

The checks use disposable temporary folders and isolated registry keys, not actual Excel installation settings. The installer also supports `--self-test`. `--preview` simulates the graphical flow without installing; it must not be used as installation evidence.

### Project map

| File | Responsibility |
|---|---|
| `Commands.cs` | Excel commands, row duplication, immediate undo |
| `Shortcuts.cs` | Shortcut validation, defaults, preference storage |
| `AddIn.cs` | Add-in lifecycle and shortcut binding |
| `Ribbon.cs` | Ribbon callbacks; no Duplicate Row button |
| `SettingsForm.cs` | English shortcut settings window |
| `Theme.cs`, `Brand.cs`, `Assets/CellMate.ico` | Shared ShatootSoft-inspired design and icons |
| `Setup/Installer.cs` | Detection, registration, migration, rollback, removal |
| `Setup/Program.cs` | Setup interface and command-line entry point |
| `Setup/InstallerTests.cs` | Isolated installer regression checks |
| `Build.ps1` | Repeatable build and release packaging |
| `CHANGELOG.md` | Version history and future release procedure |

Follow the [release procedure](CHANGELOG.md#release-procedure-authorized-maintainers) for subsequent versions. Bug reports and feature requests are welcome; source contributions require prior authorization.

## License

Copyright (c) 2026 Arash Ghasemi Rad (ShatootSoft).

Official binaries: free use, including commercial use, under the [CellMate Binary Use and Source Viewing License 1.0](LICENSE). Original source: viewing only, subject to the platform/legal exceptions in that license. Excel-DNA and other third-party components retain their own terms; see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
