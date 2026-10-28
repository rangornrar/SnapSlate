[CmdletBinding()]
param(
    [switch]$Preview,
    [string]$IdentityName = $env:SNAPSLATE_STORE_IDENTITY_NAME,
    [string]$Publisher = $env:SNAPSLATE_STORE_PUBLISHER,
    [string]$PublisherDisplayName = $env:SNAPSLATE_STORE_PUBLISHER_DISPLAY_NAME,
    [string]$PackageVersion,
    [ValidateSet('x64')][string]$Platform = 'x64'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0
$projectRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $projectRoot 'SnapSlate.csproj'
[xml]$project = Get-Content -LiteralPath $projectPath
$applicationVersion = [version]$project.SelectSingleNode('/Project/PropertyGroup/Version').InnerText

if ($Preview) {
    if ($IdentityName -or $Publisher -or $PublisherDisplayName) {
        throw 'Preview cannot be combined with a Store identity.'
    }
} else {
    if ($IdentityName -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]{2,49}$' -or $IdentityName -in 'E3C7A4ED-2E83-45C8-852A-A0469B8C2291', 'SnapSlate.StorePreview') {
        throw 'Supply Package/Identity/Name exactly as assigned by Partner Center.'
    }
    if ($Publisher -notmatch '^CN=[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}$') {
        throw 'Supply Package/Identity/Publisher (CN=GUID) from Partner Center.'
    }
    if ([string]::IsNullOrWhiteSpace($PublisherDisplayName) -or $PublisherDisplayName.Length -gt 256 -or $PublisherDisplayName -eq 'AppPublisher') {
        throw 'Supply Package/Properties/PublisherDisplayName from Partner Center.'
    }
}

# Store reserves the fourth component. Preserve the day and release number in build.
if (-not $PackageVersion) {
    if ($applicationVersion.Build -lt 1 -or $applicationVersion.Build -gt 31 -or $applicationVersion.Revision -lt 0 -or $applicationVersion.Revision -gt 999) {
        throw 'The application no longer uses year.month.day.release; specify -PackageVersion explicitly.'
    }
    $PackageVersion = '{0}.{1}.{2}.0' -f $applicationVersion.Major, $applicationVersion.Minor, ($applicationVersion.Build * 1000 + $applicationVersion.Revision)
}
if ($PackageVersion -notmatch '^[0-9]+\.[0-9]+\.[0-9]+\.0$') { throw 'Store version must have four numeric components and end in .0.' }
$storeVersion = [version]$PackageVersion
if ($storeVersion.Major -eq 0 -or $storeVersion.Major -gt 65535 -or $storeVersion.Minor -gt 65535 -or $storeVersion.Build -gt 65535) {
    throw 'Store version components must fit in 0..65535, with a nonzero major version.'
}

$mode = if ($Preview) { 'PREVIEW-NOT-FOR-SUBMISSION' } else { 'submission' }
$output = Join-Path $projectRoot ('AppPackages\Store\{0}-{1}-{2}-{3}' -f $mode, $PackageVersion, $Platform, [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output | Out-Null
$manifestPath = Join-Path $output 'Package.appxmanifest'
$manifest = [xml]::new()
$manifest.PreserveWhitespace = $true
$manifest.Load((Join-Path $projectRoot 'Package.appxmanifest'))
$manifest.Package.Identity.Version = $PackageVersion
if ($Preview) {
    $manifest.Package.Identity.Name = 'SnapSlate.StorePreview'
    $manifest.Package.Properties.DisplayName = 'SnapSlate (Store preview)'
} else {
    $manifest.Package.Identity.Name = $IdentityName
    $manifest.Package.Identity.Publisher = $Publisher
    $manifest.Package.Properties.PublisherDisplayName = $PublisherDisplayName
}
$manifest.Save($manifestPath)

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Install Visual Studio Build Tools with MSVC tools for Store symbol packaging.' }
$symbolTool = & $vswhere -latest -products '*' -find 'VC\Tools\MSVC\**\bin\Hostx64\x64\mspdbcmf.exe' | Select-Object -First 1
if (-not $symbolTool) { throw 'mspdbcmf.exe is missing; install the Visual Studio MSVC build tools.' }

$buildProperties = @(
    "-p:Platform=$Platform",
    "-p:RuntimeIdentifier=win-$Platform",
    '-p:WindowsPackageType=MSIX',
    '-p:SelfContained=true',
    '-p:WindowsAppSDKSelfContained=true',
    '-p:PublishReadyToRun=false',
    '-p:AppxPackageSigningEnabled=false',
    '-p:AppxAutoIncrementPackageRevision=false',
    '-p:AppxBundle=Never',
    '-p:UapAppxPackageBuildMode=StoreUpload',
    '-p:AppxSymbolPackageEnabled=true',
    # Both discovery properties are consumed again by the upload-symbol target.
    "-p:PdbCmfx64ExeFullPath=$symbolTool",
    "-p:PdbCmfx86ExeFullPath=$symbolTool",
    '-p:GenerateAppInstallerFile=false',
    "-p:StorePackageManifest=$manifestPath",
    "-p:AppxPackageDir=$output\"
)

# Clear PRI files from any preceding unpackaged build using the packaged properties.
& dotnet clean $projectPath -c Release @buildProperties
if ($LASTEXITCODE -ne 0) { throw 'Store clean failed.' }
& dotnet build $projectPath -c Release @buildProperties -p:RestoreLockedMode=true -p:GenerateAppxPackageOnBuild=true
if ($LASTEXITCODE -ne 0) { throw 'Store package build failed.' }

$packages = @(Get-ChildItem -LiteralPath $output -Recurse -File -Filter '*.msix')
$uploads = @(Get-ChildItem -LiteralPath $output -File | Where-Object { $_.Extension -in '.msixupload', '.appxupload' })
if ($packages.Count -ne 1 -or $uploads.Count -ne 1) { throw 'Expected one x64 MSIX and one Store upload archive.' }
& (Join-Path $PSScriptRoot 'Test-StorePackage.ps1') -Path $packages[0].FullName -ManifestPath $manifestPath

# The SDK creates a separate Store payload, which can differ from the _Test MSIX.
$upload = [IO.Compression.ZipFile]::OpenRead($uploads[0].FullName)
try {
    $uploadPackages = @($upload.Entries | Where-Object { $_.FullName.EndsWith('.msix', [StringComparison]::OrdinalIgnoreCase) })
    $uploadSymbols = @($upload.Entries | Where-Object { $_.FullName.EndsWith('.appxsym', [StringComparison]::OrdinalIgnoreCase) })
    if ($uploadPackages.Count -ne 1 -or $uploadSymbols.Count -ne 1) { throw 'Store upload must contain one MSIX and its symbols.' }
    $validatedDirectory = Join-Path $output 'validated-upload'
    New-Item -ItemType Directory -Path $validatedDirectory | Out-Null
    $submissionPackage = Join-Path $validatedDirectory ([IO.Path]::GetFileName($uploadPackages[0].FullName))
    [IO.Compression.ZipFileExtensions]::ExtractToFile($uploadPackages[0], $submissionPackage, $false)
} finally { $upload.Dispose() }
& (Join-Path $PSScriptRoot 'Test-StorePackage.ps1') -Path $submissionPackage -ManifestPath $manifestPath

$revision = & git -C $projectRoot rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Git revision unavailable.' }
$dirty = [bool](& git -C $projectRoot status --porcelain)
$metadata = [ordered]@{
    submissionIdentityConfigured = -not [bool]$Preview
    mode = $mode
    applicationVersion = $applicationVersion.ToString()
    packageVersion = $PackageVersion
    architecture = $Platform
    identityName = [string]$manifest.Package.Identity.Name
    publisher = [string]$manifest.Package.Identity.Publisher
    publisherDisplayName = [string]$manifest.Package.Properties.PublisherDisplayName
    commit = $revision
    workingTreeModified = $dirty
    signed = $false
    certification = 'Not yet certified by Microsoft Store'
    sha256 = @($submissionPackage, $uploads[0].FullName) | ForEach-Object {
        [ordered]@{ file = [IO.Path]::GetFileName($_); sha256 = (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToLowerInvariant() }
    }
}
$metadata | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $output 'store-build.json') -Encoding utf8
if ($Preview) {
    'Technical preview only. Do not upload this placeholder identity to Partner Center.' |
        Set-Content -LiteralPath (Join-Path $output 'PREVIEW-NOT-FOR-SUBMISSION.txt') -Encoding utf8
}
Write-Host "Store package output: $output"
