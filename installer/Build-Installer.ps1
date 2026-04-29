[CmdletBinding()]
param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Debug",

    [ValidateSet("x64")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

$installerDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $installerDir
$projectPath = Join-Path $projectDir "SnapSlate.csproj"
$distDir = Join-Path $installerDir "dist"
$distAppDir = Join-Path $distDir "app"
$zipPath = Join-Path $installerDir "SnapSlate-Installer-x64.zip"
$buildOutputDir = Join-Path $projectDir "bin\$Platform\$Configuration\net10.0-windows10.0.26100.0\win-$Platform"

function Remove-DirectorySafely {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return
    }

    $resolvedPath = (Resolve-Path $Path).Path
    if ($resolvedPath -notlike "$installerDir*") {
        throw "Refusing to remove unexpected path: $resolvedPath"
    }

    Remove-Item -LiteralPath $resolvedPath -Recurse -Force
}

Write-Host "Building SnapSlate ($Configuration, $Platform)..."
dotnet build $projectPath `
    -c $Configuration `
    -p:Platform=$Platform `
    -p:WindowsPackageType=None `
    -p:WindowsAppSDKSelfContained=true `
    -p:PublishReadyToRun=false

if (-not (Test-Path $buildOutputDir)) {
    throw "The build output folder was not found: $buildOutputDir"
}

Remove-DirectorySafely -Path $distDir
New-Item -ItemType Directory -Force -Path $distAppDir | Out-Null

Copy-Item -Path (Join-Path $buildOutputDir "*") -Destination $distAppDir -Recurse -Force
Copy-Item -LiteralPath (Join-Path $installerDir "README.md") -Destination (Join-Path $distDir "README.md") -Force

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $distDir "*") -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Installer staging ready:"
Write-Host " - Folder: $distDir"
Write-Host " - Zip:    $zipPath"
