[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$ManifestPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0
$expected = [xml](Get-Content -LiteralPath $ManifestPath -Raw)
$archive = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Path).Path)
try {
    $entries = [Collections.Generic.Dictionary[string, IO.Compression.ZipArchiveEntry]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($fileEntry in $archive.Entries) {
        $name = $fileEntry.FullName.Replace('\', '/')
        if ($entries.ContainsKey($name)) { throw 'Duplicate Windows package path.' }
        $entries.Add($name, $fileEntry)
    }
    $entry = $entries['AppxManifest.xml']
    if (-not $entry) { throw 'MSIX manifest missing.' }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { $actual = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
    foreach ($attribute in 'Name', 'Publisher', 'Version') {
        if ($actual.Package.Identity.GetAttribute($attribute) -cne $expected.Package.Identity.GetAttribute($attribute)) {
            throw "MSIX identity mismatch: $attribute"
        }
    }
    if ($actual.Package.Identity.ProcessorArchitecture -ne 'x64' -or ([version]$actual.Package.Identity.Version).Revision -ne 0) {
        throw 'Expected an x64 package with Store revision zero.'
    }
    if ($actual.Package.Properties.PublisherDisplayName -cne $expected.Package.Properties.PublisherDisplayName) {
        throw 'MSIX publisher display name mismatch.'
    }
    $capabilities = @($actual.Package.Capabilities.ChildNodes | Where-Object NodeType -eq Element)
    if ($capabilities.Count -ne 1 -or $capabilities[0].GetAttribute('Name') -ne 'runFullTrust') {
        throw 'Unexpected Store capabilities; only runFullTrust is required.'
    }
    $families = @($actual.Package.Dependencies.TargetDeviceFamily)
    if ($families.Count -ne 1 -or $families[0].Name -ne 'Windows.Desktop') { throw 'Expected a Windows Desktop package.' }
    foreach ($file in @('SnapSlate.exe', 'SnapSlate.dll', 'SnapSlate.runtimeconfig.json', 'coreclr.dll', 'hostfxr.dll', 'Microsoft.UI.Xaml.dll', 'resources.pri', 'Assets/AppIcon.ico')) {
        if (-not $entries.ContainsKey($file)) { throw "Missing self-contained component: $file" }
    }
    if ($entries.ContainsKey('AppxSignature.p7x')) { throw 'Store submission must not contain an unexpected local signature.' }
    $runtimeReader = [IO.StreamReader]::new($entries['SnapSlate.runtimeconfig.json'].Open())
    try { $runtime = $runtimeReader.ReadToEnd() | ConvertFrom-Json } finally { $runtimeReader.Dispose() }
    if (-not $runtime.runtimeOptions.PSObject.Properties['includedFrameworks']) { throw 'Expected a self-contained .NET runtime.' }
    Write-Host "Validated MSIX: $($actual.Package.Identity.Name), $($actual.Package.Identity.Version), unsigned x64."
} finally { $archive.Dispose() }
