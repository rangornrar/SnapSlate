[CmdletBinding()]
param([Parameter(Mandatory)][string]$Destination)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0
if ($env:GITHUB_ACTIONS -ne 'true') { throw 'This installer is restricted to disposable GitHub Actions runners.' }
$toolchain = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'toolchain.json') -Raw | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path $Destination | Out-Null
$download = Join-Path $Destination 'inno-setup.exe'
Invoke-WebRequest -Uri $toolchain.innoSetupUrl -OutFile $download
if ((Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash -ne $toolchain.innoSetupSha256) { throw 'Inno Setup checksum mismatch.' }
$installDirectory = Join-Path $Destination 'compiler'
$process = Start-Process -FilePath $download -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/SP-',"/DIR=`"$installDirectory`"") -Wait -PassThru -WindowStyle Hidden
if ($process.ExitCode -ne 0) { throw "Inno Setup installation failed: $($process.ExitCode)." }
if (-not (Test-Path -LiteralPath (Join-Path $installDirectory 'ISCC.exe'))) { throw 'Compiler missing after installation.' }
"SNAPSLATE_ISCC=$(Join-Path $installDirectory 'ISCC.exe')" | Out-File -LiteralPath $env:GITHUB_ENV -Append -Encoding utf8
