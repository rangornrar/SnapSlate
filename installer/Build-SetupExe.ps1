[CmdletBinding()]
param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",

    [ValidateSet("x64")]
    [string]$Platform = "x64",
    [string]$SigningCertificateThumbprint = $env:SNAPSLATE_SIGNING_CERT_THUMBPRINT,
    [uri]$TimestampUrl = "https://timestamp.digicert.com"
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

& $buildInstallerScript -Configuration $Configuration -Platform $Platform -SigningCertificateThumbprint $SigningCertificateThumbprint -TimestampUrl $TimestampUrl

$iscc = $env:SNAPSLATE_ISCC
if (-not $iscc) { $iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue }
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
    if (-not $resolvedReleaseDir.StartsWith($scriptDir + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean unexpected directory: $resolvedReleaseDir"
    }

    Remove-Item -LiteralPath $resolvedReleaseDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null

[xml]$project = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $scriptDir) 'SnapSlate.csproj')
$version = $project.SelectSingleNode("/Project/PropertyGroup/Version").InnerText
$arguments = @("/DMyAppVersion=$version")
if ($SigningCertificateThumbprint) {
    if ($SigningCertificateThumbprint -notmatch '^[A-Fa-f0-9]{40}$' -or $TimestampUrl.Scheme -ne 'https') { throw 'Invalid signing configuration.' }
    $signTool = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe" | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
    if (-not $signTool) { throw 'SignTool missing.' }
    $signCommand = '"' + $signTool + '" sign /s My /sha1 ' + $SigningCertificateThumbprint + ' /fd SHA256 /tr ' + $TimestampUrl.AbsoluteUri + ' /td SHA256 $f'
    $arguments += @('/DSignRelease=1',"/SreleaseSign=$signCommand")
}
$arguments += $innoScript
& $iscc @arguments
if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed with exit code $LASTEXITCODE." }

$setupPath = Join-Path $releaseDir "SnapSlate-Setup.exe"
if (-not (Test-Path $setupPath)) {
    throw "Setup.exe was not produced at $setupPath"
}

Write-Host ""
Write-Host "Setup.exe ready:"
Write-Host " - $setupPath"

& (Join-Path $scriptDir "Test-Distribution.ps1")
