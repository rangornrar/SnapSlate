# SnapSlate Installer

This folder contains everything needed to generate and install a signed SnapSlate package.

## Build a single Setup.exe

From `C:\Users\user\Documents\Freelance\programmation\screenshot\SnapSlate`:

```powershell
.\installer\Build-SetupExe.ps1
```

This creates:

- `installer\release\SnapSlate-Setup.exe`

## Build a fresh installer

From `C:\Users\user\Documents\Freelance\programmation\screenshot\SnapSlate`:

```powershell
.\installer\Build-Installer.ps1
```

The script will:

- create or reuse a local signing certificate in `installer\private`
- publish a signed MSIX package
- stage a ready-to-share installer in `installer\dist`
- create `installer\SnapSlate-Installer-x64.zip`

## Install on a machine

Open `installer\dist\Install-SnapSlate.cmd` as administrator.

For a simpler user-facing install, use:

- `installer\release\SnapSlate-Setup.exe`

The installer will:

- trust the bundled signing certificate
- install the MSIX package and its Windows App SDK dependencies
- launch SnapSlate when the install is complete

## Uninstall

Run:

```powershell
.\installer\dist\Uninstall-SnapSlate.ps1
```
