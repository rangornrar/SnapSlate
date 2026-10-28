[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{40}$')][string]$CertificateThumbprint,
    [Parameter(Mandatory)][uri]$TimestampUrl,
    [string]$SignToolPath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0
if ($TimestampUrl.Scheme -ne 'https') { throw 'Use an HTTPS timestamp service.' }
$certificate = Get-Item -LiteralPath "Cert:\CurrentUser\My\$CertificateThumbprint"
if (-not $certificate.HasPrivateKey -or $certificate.NotAfter -le (Get-Date)) { throw 'A valid certificate with a private key is required.' }
if (-not $SignToolPath) {
    $SignToolPath = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe" | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $SignToolPath) { throw 'SignTool is missing from the Windows SDK.' }
& $SignToolPath sign /s My /sha1 $CertificateThumbprint /fd SHA256 /tr $TimestampUrl.AbsoluteUri /td SHA256 $Path
if ($LASTEXITCODE -ne 0) { throw 'SignTool failed. No unsigned fallback is allowed when signing is requested.' }
$signature = Get-AuthenticodeSignature -LiteralPath $Path
if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Signature verification failed.' }
