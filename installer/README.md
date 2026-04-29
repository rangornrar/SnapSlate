# SnapSlate Installer

This folder contains the scripts used to build the SnapSlate desktop installer.

## Build the installer

From `C:\Users\user\Documents\Freelance\programmation\screenshot\SnapSlate`:

```powershell
.\installer\Build-SetupExe.ps1
```

This will:

- build SnapSlate in the self-contained Debug configuration
- stage the app files in `installer\dist\app`
- create `installer\release\SnapSlate-Setup.exe`
- create `installer\SnapSlate-Installer-x64.zip`

## Install on a machine

Run:

- `installer\release\SnapSlate-Setup.exe`

The installer will:

- copy SnapSlate files into the chosen install folder
- create a Start Menu entry
- optionally create a desktop icon
- launch SnapSlate when installation completes

## Uninstall

Use the Windows uninstall entry for SnapSlate from Apps & features.
