[CmdletBinding()]
param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",

    [ValidateSet("x64")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

$installerDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $installerDir
$projectPath = Join-Path $projectDir "SnapSlate.csproj"
$appPackagesDir = Join-Path $projectDir "AppPackages"
$privateDir = Join-Path $installerDir "private"
$distDir = Join-Path $installerDir "dist"
$zipPath = Join-Path $installerDir "SnapSlate-Installer-x64.zip"

$certificateSubject = "CN=AppPublisher"
$certificateBaseName = "SnapSlate-Installer"
$certificatePasswordFile = Join-Path $privateDir "certificate-password.txt"
$certificatePath = Join-Path $privateDir "$certificateBaseName.cer"
$pfxPath = Join-Path $privateDir "$certificateBaseName.pfx"

function New-RandomPassword {
    param([int]$Length = 32)

    $characters = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%*-_+".ToCharArray()
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object byte[] $Length
    $random.GetBytes($bytes)

    $buffer = New-Object System.Text.StringBuilder
    foreach ($value in $bytes) {
        [void]$buffer.Append($characters[$value % $characters.Length])
    }

    return $buffer.ToString()
}

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

New-Item -ItemType Directory -Force -Path $privateDir | Out-Null

if (-not (Test-Path $certificatePasswordFile)) {
    Set-Content -Path $certificatePasswordFile -Value (New-RandomPassword) -Encoding ascii
}

$certificatePassword = (Get-Content -Path $certificatePasswordFile -Raw).Trim()
$securePassword = ConvertTo-SecureString $certificatePassword -AsPlainText -Force
$shouldCreateCertificate = (-not (Test-Path $pfxPath)) -or (-not (Test-Path $certificatePath))

if (-not $shouldCreateCertificate) {
    try {
        Get-PfxData -FilePath $pfxPath -Password $securePassword | Out-Null
    }
    catch {
        Remove-Item -LiteralPath $pfxPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $certificatePath -Force -ErrorAction SilentlyContinue
        $shouldCreateCertificate = $true
    }
}

if ($shouldCreateCertificate) {
    Write-Host "Creating signing certificate..."

    $certificate = New-SelfSignedCertificate `
        -Subject $certificateSubject `
        -Type CodeSigningCert `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -NotAfter (Get-Date).AddYears(5)

    Export-PfxCertificate `
        -Cert "Cert:\CurrentUser\My\$($certificate.Thumbprint)" `
        -FilePath $pfxPath `
        -Password $securePassword | Out-Null

    Export-Certificate `
        -Cert "Cert:\CurrentUser\My\$($certificate.Thumbprint)" `
        -FilePath $certificatePath | Out-Null
}

$certificateInfo = Get-PfxCertificate -FilePath $certificatePath

Write-Host "Building signed MSIX package..."
dotnet publish $projectPath `
    -c $Configuration `
    -p:Platform=$Platform `
    -p:GenerateAppxPackageOnBuild=true `
    -p:PackageCertificateKeyFile=$pfxPath `
    -p:PackageCertificatePassword=$certificatePassword `
    -p:PackageCertificateThumbprint=$($certificateInfo.Thumbprint)

$packageDirectory = Get-ChildItem -Path $appPackagesDir -Directory |
    Where-Object { $_.Name -like "*_${Platform}_*" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $packageDirectory) {
    throw "No package directory was produced in $appPackagesDir."
}

$msixFile = Get-ChildItem -Path $packageDirectory.FullName -Filter "*.msix" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

$cerFile = Get-ChildItem -Path $packageDirectory.FullName -Filter "*.cer" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $msixFile -or $null -eq $cerFile) {
    throw "The package output is incomplete. Expected both an .msix and a .cer file."
}

Remove-DirectorySafely -Path $distDir
New-Item -ItemType Directory -Force -Path $distDir | Out-Null

Copy-Item -LiteralPath $msixFile.FullName -Destination (Join-Path $distDir "SnapSlate.msix")
Copy-Item -LiteralPath $cerFile.FullName -Destination (Join-Path $distDir "SnapSlate.cer")

$dependenciesDirectory = Join-Path $packageDirectory.FullName "Dependencies"
if (Test-Path $dependenciesDirectory) {
    Copy-Item -LiteralPath $dependenciesDirectory -Destination (Join-Path $distDir "Dependencies") -Recurse
}

Copy-Item -LiteralPath (Join-Path $installerDir "Install-SnapSlate.ps1") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $installerDir "Install-SnapSlate.cmd") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $installerDir "Uninstall-SnapSlate.ps1") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $installerDir "README.md") -Destination $distDir

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $distDir "*") -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Installer ready:"
Write-Host " - Folder: $distDir"
Write-Host " - Zip:    $zipPath"
