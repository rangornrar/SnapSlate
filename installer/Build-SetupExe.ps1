[CmdletBinding()]
param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",

    [ValidateSet("x64")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildInstallerScript = Join-Path $scriptDir "Build-Installer.ps1"
$innoScript = Join-Path $scriptDir "SnapSlate.iss"
$releaseDir = Join-Path $scriptDir "release"

if (-not (Test-Path $buildInstallerScript)) {
    throw "Missing build script: $buildInstallerScript"
}

if (-not (Test-Path $innoScript)) {
    throw "Missing Inno Setup script: $innoScript"
}

& $buildInstallerScript -Configuration Debug -Platform $Platform

$iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue
if (-not $iscc) {
    $fallbackIscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if (Test-Path $fallbackIscc) {
        $iscc = $fallbackIscc
    }
}

if (-not $iscc) {
    $userFallbackIscc = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
    if (Test-Path $userFallbackIscc) {
        $iscc = $userFallbackIscc
    }
}

if (-not $iscc) {
    throw "ISCC.exe was not found. Install Inno Setup 6 first."
}

if (Test-Path $releaseDir) {
    $resolvedReleaseDir = (Resolve-Path $releaseDir).Path
    if ($resolvedReleaseDir -notlike "$scriptDir*") {
        throw "Refusing to clean unexpected directory: $resolvedReleaseDir"
    }

    Remove-Item -LiteralPath $resolvedReleaseDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null

& $iscc $innoScript

$setupPath = Join-Path $releaseDir "SnapSlate-Setup.exe"
if (-not (Test-Path $setupPath)) {
    throw "Setup.exe was not produced at $setupPath"
}

Write-Host ""
Write-Host "Setup.exe ready:"
Write-Host " - $setupPath"
