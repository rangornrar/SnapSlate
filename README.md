# SnapSlate

[![Build](https://github.com/rangornrar/SnapSlate/actions/workflows/build.yml/badge.svg)](https://github.com/rangornrar/SnapSlate/actions/workflows/build.yml)

Application Windows en C# / WinUI 3 pour transformer des captures d’écran en guides illustrés.

![Interface de création d’une procédure](site/shots/procedure.png)

## Utilisation

1. Télécharger `SnapSlate-Setup.exe` depuis les [versions publiées](https://github.com/rangornrar/SnapSlate/releases/latest).
2. Installer puis ouvrir SnapSlate.
3. Réaliser une capture avec `Win + Shift + S` : lorsque SnapSlate est ouverte, la capture est importée dans un nouvel onglet.
4. Ajouter textes, flèches, gommettes, cadres, zones de focus ou de masquage.
5. Enregistrer le projet `.snapslate` ou exporter les documents.

L’application doit être ouverte pour importer les captures du presse-papiers. Elle ne remplace pas le raccourci système lorsqu’elle est fermée.

## Formats et fonctionnalités

- Recadrage, palettes, réglage d’opacité et légendes de gommettes.
- Plusieurs captures/étapes, annulation et rétablissement.
- Projets JSON `.snapslate` incluant l’état des documents.
- Export d’images PNG et de guides PDF, DOCX, Markdown et HTML.
- Interface français / anglais.

Les exports HTML et Markdown référencent le sous-dossier `images/` : conserver le dossier exporté complet lors du partage.

## Développement sous Windows

Prérequis : Windows 10/11, SDK .NET 10 et outils Windows/WinUI avec le SDK Windows ciblé par `SnapSlate.csproj`.

Depuis la racine du dépôt :

```powershell
dotnet restore .\SnapSlate.csproj -p:Configuration=Release -p:Platform=x64 -r win-x64 -p:PublishReadyToRun=false
dotnet build .\SnapSlate.csproj -c Release -p:Platform=x64 -r win-x64 -p:PublishReadyToRun=false --no-restore
dotnet test .\tests\SnapSlate.Tests.csproj -c Release
```

Pour préparer les fichiers exécutables sans packaging MSIX :

```powershell
.\installer\Build-Installer.ps1 -Configuration Release
```

Les fichiers sont placés dans `installer/dist/app/`. Lancer `SnapSlate.exe` depuis ce dossier sous Windows. Les scripts historiques `Install-SnapSlate.ps1` concernent un autre parcours MSIX ; ils ne sont pas le point d’entrée de ce staging exécutable.

Pour produire le setup avec Inno Setup 6 installé :

```powershell
.\installer\Build-SetupExe.ps1
```

Résultat : `installer/release/SnapSlate-Setup.exe`. La compilation CI ne produit pas une nouvelle version distribuée et ne valide pas l’installation interactive.

## Architecture et qualité

- [Architecture et limites actuelles](docs/ARCHITECTURE.md)
- [Décisions de conception](docs/decisions/README.md)
- [Périmètre des tests et vérifications manuelles](docs/QUALITY.md)
- [Contribuer](CONTRIBUTING.md) · [Sécurité](SECURITY.md)

Le code source et les binaires ne disposent pas actuellement d’un fichier de licence général dans ce dépôt. Les dépendances ont leurs propres conditions d’utilisation.

## Sauvegarde et masquage

Le projet éditable `.snapslate` conserve les sources originales. Les masques sont appliqués de manière opaque aux images exportées ; ils ne suppriment pas les pixels du projet. Les fichiers `.bak` contiennent la version précédente du projet. Exporter un PNG ne remplace pas la sauvegarde du projet.
