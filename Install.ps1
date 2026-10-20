param([switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$setup = Join-Path $PSScriptRoot 'dist\CellMateSetup.exe'
if (!(Test-Path -LiteralPath $setup)) { throw 'Build the installer first with .\Build.ps1, or download the official release.' }
$arguments = if ($Uninstall) { '--uninstall' } else { '' }
if ($arguments) { $process = Start-Process -FilePath $setup -ArgumentList $arguments -Wait -PassThru }
else { $process = Start-Process -FilePath $setup -Wait -PassThru }
if ($process.ExitCode -ne 0) { throw "CellMate Setup exited with code $($process.ExitCode)." }
