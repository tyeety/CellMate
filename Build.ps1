$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    dotnet build CellMate.csproj -c Release
    if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
    dotnet build Setup\CellMate.Setup.csproj -c Release
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed.' }
    Copy-Item -LiteralPath LICENSE,THIRD_PARTY_NOTICES.md,README.md -Destination dist -Force
    [xml]$project = Get-Content -LiteralPath CellMate.csproj
    $version = [string]$project.Project.PropertyGroup.Version
    $package = Join-Path $PSScriptRoot "dist\CellMate-Setup-$version.zip"
    $packageFiles = @('dist\CellMateSetup.exe','dist\CellMateSetup.exe.config','dist\LICENSE','dist\THIRD_PARTY_NOTICES.md','dist\README.md')
    Compress-Archive -LiteralPath $packageFiles -DestinationPath $package -Force
    Write-Host "Ready: $package"
} finally { Pop-Location }
