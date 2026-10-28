[CmdletBinding()]
param([switch]$InstallSmoke)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0
$projectRoot = Split-Path -Parent $PSScriptRoot
$release = Join-Path $PSScriptRoot 'release'
$setup = Join-Path $release 'SnapSlate-Setup.exe'
$portable = Join-Path $PSScriptRoot 'SnapSlate-Installer-x64.zip'
[xml]$project = Get-Content -LiteralPath (Join-Path $projectRoot 'SnapSlate.csproj')
$version = $project.SelectSingleNode("/Project/PropertyGroup/Version").InnerText
[xml]$manifest = Get-Content -LiteralPath (Join-Path $projectRoot 'Package.appxmanifest')
if ($manifest.Package.Identity.Version -ne $version) { throw 'Manifest version mismatch.' }
foreach ($file in @($setup,$portable)) { if (-not (Test-Path -LiteralPath $file) -or (Get-Item -LiteralPath $file).Length -eq 0) { throw "Missing distribution: $file" } }
if ((Get-Item -LiteralPath $setup).VersionInfo.ProductVersion -ne $version) { throw 'Setup version mismatch.' }
$application = Join-Path $PSScriptRoot 'dist\app\SnapSlate.exe'
if ((Get-Item -LiteralPath $application).VersionInfo.FileVersion -ne $version) { throw 'Application version mismatch.' }
$zip = [IO.Compression.ZipFile]::OpenRead($portable)
try {
    foreach ($name in @('SnapSlate.exe','SnapSlate.dll','SnapSlate.runtimeconfig.json','hostfxr.dll','coreclr.dll')) {
        $entries = @($zip.Entries | Where-Object { ($_.FullName -replace '\\','/') -eq "app/$name" })
        if ($entries.Count -ne 1) { throw "Missing portable component: $name" }
        $stream = $entries[0].Open()
        try { $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream)) } finally { $stream.Dispose() }
        if ($hash -ne (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "dist\app\$name") -Algorithm SHA256).Hash) { throw "Portable component mismatch: $name" }
    }
} finally { $zip.Dispose() }
$sums = foreach ($file in @($setup,$portable)) { (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + [IO.Path]::GetFileName($file) }
[IO.File]::WriteAllLines((Join-Path $release 'SHA256SUMS.txt'),$sums,[Text.UTF8Encoding]::new($false))
$revision = git -C $projectRoot rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Git revision unavailable.' }
@{ version=$version; commit=$revision; sdk=(dotnet --version); signed=((Get-AuthenticodeSignature -LiteralPath $setup).Status -eq 'Valid'); sha256=$sums } |
    ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $release 'build-provenance.json') -Encoding utf8
if (-not $InstallSmoke) { return }
if ($env:GITHUB_ACTIONS -ne 'true') { throw 'Installation smoke tests run only on disposable GitHub runners.' }
$destination = Join-Path $env:RUNNER_TEMP ('SnapSlate-Smoke-' + [Guid]::NewGuid().ToString('N'))
$process = Start-Process -FilePath $setup -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/SP-',"/DIR=`"$destination`"",'/TASKS=') -Wait -PassThru -WindowStyle Hidden
if ($process.ExitCode -ne 0) { throw "Setup smoke failed: $($process.ExitCode)." }
$installedExe = Join-Path $destination 'SnapSlate.exe'
if ((Get-FileHash -LiteralPath $installedExe).Hash -ne (Get-FileHash -LiteralPath $application).Hash) { throw 'Installed binary mismatch.' }
$app = $null
try {
    $app = Start-Process -FilePath $installedExe -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(25)
    do { Start-Sleep -Milliseconds 250; $app.Refresh() } until ($app.HasExited -or $app.MainWindowHandle -ne [IntPtr]::Zero -or (Get-Date) -ge $deadline)
    if ($app.HasExited -or $app.MainWindowHandle -eq [IntPtr]::Zero) { throw 'Installed application failed to create its window.' }
} finally {
    if ($app -and -not $app.HasExited) { Stop-Process -Id $app.Id -Force }
    $uninstaller = Join-Path $destination 'unins000.exe'
    if (Test-Path -LiteralPath $uninstaller) {
        $uninstall = Start-Process -FilePath $uninstaller -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART') -Wait -PassThru -WindowStyle Hidden
        if ($uninstall.ExitCode -ne 0) { throw "Uninstall smoke failed: $($uninstall.ExitCode)." }
    }
}
if (Test-Path -LiteralPath $installedExe) { throw 'Uninstall left the application binary behind.' }
