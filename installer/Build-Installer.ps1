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

$installerDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $installerDir
$projectPath = Join-Path $projectDir "SnapSlate.csproj"
$distDir = Join-Path $installerDir "dist"
$distAppDir = Join-Path $distDir "app"
$zipPath = Join-Path $installerDir "SnapSlate-Installer-x64.zip"
$buildOutputDir = Join-Path $installerDir "publish"

function Remove-DirectorySafely {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return
    }

    $resolvedPath = (Resolve-Path $Path).Path
    if (-not $resolvedPath.StartsWith($installerDir + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove unexpected path: $resolvedPath"
    }

    Remove-Item -LiteralPath $resolvedPath -Recurse -Force
}

Remove-DirectorySafely -Path $buildOutputDir
Write-Host "Publishing SnapSlate ($Configuration, $Platform)..."
dotnet publish $projectPath `
    -c $Configuration `
    -p:Platform=$Platform `
    -p:WindowsPackageType=None `
    -p:WindowsAppSDKSelfContained=true `
    -p:PublishReadyToRun=false `
    -p:RestoreLockedMode=true `
    -r "win-$Platform" `
    --self-contained true `
    -o $buildOutputDir

if ($LASTEXITCODE -ne 0) { throw "Publish failed with exit code $LASTEXITCODE." }

if (-not (Test-Path $buildOutputDir)) {
    throw "The build output folder was not found: $buildOutputDir"
}

if ($SigningCertificateThumbprint) {
    foreach ($ownedFile in @('SnapSlate.exe','SnapSlate.dll')) {
        & (Join-Path $installerDir 'Sign-Binary.ps1') -Path (Join-Path $buildOutputDir $ownedFile) -CertificateThumbprint $SigningCertificateThumbprint -TimestampUrl $TimestampUrl
    }
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
