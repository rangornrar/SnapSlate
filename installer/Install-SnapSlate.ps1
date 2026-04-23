[CmdletBinding()]
param(
    [switch]$SkipLaunch
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

$packageName = "E3C7A4ED-2E83-45C8-852A-A0469B8C2291"
$applicationId = "App"
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDir = Split-Path -Parent $scriptPath
$packagePath = Join-Path $scriptDir "SnapSlate.msix"
$certificatePath = Join-Path $scriptDir "SnapSlate.cer"
$dependenciesPath = Join-Path $scriptDir "Dependencies"

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ElevatedInstall {
    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$scriptPath`""
    )

    if ($SkipLaunch) {
        $arguments += "-SkipLaunch"
    }

    Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $arguments | Out-Null
}

function Import-CertificateIfMissing {
    param(
        [string]$StorePath,
        [string]$Thumbprint,
        [string]$SourcePath
    )

    $existing = Get-ChildItem -Path $StorePath | Where-Object { $_.Thumbprint -eq $Thumbprint } | Select-Object -First 1
    if ($null -eq $existing) {
        Import-Certificate -FilePath $SourcePath -CertStoreLocation $StorePath | Out-Null
    }
}

function Get-DependencyArchitectures {
    switch ([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()) {
        "X86" { return @("x86") }
        "X64" { return @("x64", "x86") }
        "Arm64" { return @("arm64", "x64", "x86") }
        default { return @("x64", "x86") }
    }
}

if (-not (Test-Path $packagePath)) {
    throw "Missing package file: $packagePath"
}

if (-not (Test-Path $certificatePath)) {
    throw "Missing certificate file: $certificatePath"
}

if (-not (Test-IsAdministrator)) {
    Write-Host "Administrator rights are required to trust the SnapSlate certificate."
    Start-ElevatedInstall
    return
}

$certificate = Get-PfxCertificate -FilePath $certificatePath
Import-CertificateIfMissing -StorePath "Cert:\LocalMachine\Root" -Thumbprint $certificate.Thumbprint -SourcePath $certificatePath
Import-CertificateIfMissing -StorePath "Cert:\LocalMachine\TrustedPeople" -Thumbprint $certificate.Thumbprint -SourcePath $certificatePath

$dependencyPackages = @()
if (Test-Path $dependenciesPath) {
    foreach ($architecture in Get-DependencyArchitectures) {
        $dependencyFolder = Join-Path $dependenciesPath $architecture
        if (Test-Path $dependencyFolder) {
            $dependencyPackages += Get-ChildItem -Path $dependencyFolder -Filter "*.msix" |
                Sort-Object FullName |
                Select-Object -ExpandProperty FullName
        }
    }
}

$installArguments = @{
    Path = $packagePath
    ForceApplicationShutdown = $true
    ForceUpdateFromAnyVersion = $true
}

if ($dependencyPackages.Count -gt 0) {
    $installArguments.DependencyPath = $dependencyPackages
}

Write-Host "Installing SnapSlate..."
Add-AppxPackage @installArguments

Write-Host "SnapSlate is installed."

if (-not $SkipLaunch) {
    $package = Get-AppxPackage -Name $packageName | Select-Object -First 1
    if ($null -ne $package) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "shell:AppsFolder\$($package.PackageFamilyName)!$applicationId" | Out-Null
    }
}
