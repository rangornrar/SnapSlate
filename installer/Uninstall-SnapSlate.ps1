[CmdletBinding()]
param(
    [switch]$KeepCertificate
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

$packageName = "E3C7A4ED-2E83-45C8-852A-A0469B8C2291"
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDir = Split-Path -Parent $scriptPath
$certificatePath = Join-Path $scriptDir "SnapSlate.cer"

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ElevatedUninstall {
    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$scriptPath`""
    )

    if ($KeepCertificate) {
        $arguments += "-KeepCertificate"
    }

    Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $arguments | Out-Null
}

function Remove-CertificateIfPresent {
    param(
        [string]$StorePath,
        [string]$Thumbprint
    )

    Get-ChildItem -Path $StorePath |
        Where-Object { $_.Thumbprint -eq $Thumbprint } |
        Remove-Item -Force
}

if (-not (Test-IsAdministrator)) {
    Write-Host "Administrator rights are required to remove the SnapSlate certificate."
    Start-ElevatedUninstall
    return
}

$package = Get-AppxPackage -Name $packageName | Select-Object -First 1
if ($null -ne $package) {
    Write-Host "Removing SnapSlate..."
    Remove-AppxPackage -Package $package.PackageFullName
}

if (-not $KeepCertificate -and (Test-Path $certificatePath)) {
    $certificate = Get-PfxCertificate -FilePath $certificatePath
    Remove-CertificateIfPresent -StorePath "Cert:\LocalMachine\TrustedPeople" -Thumbprint $certificate.Thumbprint
    Remove-CertificateIfPresent -StorePath "Cert:\LocalMachine\Root" -Thumbprint $certificate.Thumbprint
}

Write-Host "SnapSlate has been removed."
