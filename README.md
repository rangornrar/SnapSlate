# SnapSlate

SnapSlate est une application Windows de capture et d'annotation construite avec WinUI 3 et le Windows App SDK.

L'objectif du projet est simple: recuperer les captures faites avec `Win + Shift + S`, les ouvrir dans un nouvel onglet, puis permettre un travail d'annotation rapide sans passer par plusieurs outils.

## Fonctionnalites

- Import automatique des captures `Win + Shift + S` tant que l'application est ouverte
- Un nouvel onglet par capture
- Fermeture d'onglet avec proposition d'enregistrement si le contenu n'a pas encore ete exporte
- Crop interactif avec application et reinitialisation
- Ajout de texte libre
- Gommettes numerotees ou alphabetiques
- Option pour reinitialiser la numerotation des gommettes au changement de couleur
- Fleches droites et fleches courbes
- Encadrement de zones avec rectangles et ovales
- 10 palettes de couleurs en degrade
- Export PNG

## Limitation actuelle

Aujourd'hui, `SnapSlate` ecoute le presse-papiers Windows pendant qu'il est deja lance.

Concretement:

- `Win + Shift + S` cree bien un nouvel onglet si `SnapSlate` est ouverte
- `SnapSlate` ne remplace pas encore le raccourci systeme au niveau Windows quand l'application est fermee

Cette limite est volontairement documentee ici pour eviter toute ambiguite.

## Installation

### Option la plus simple

Genere ou telecharge le setup puis lance:

- `installer/release/SnapSlate-Setup.exe`

Le setup embarque:

- le package MSIX de l'application
- le certificat de signature necessaire a l'installation locale
- les dependances Windows App SDK utiles a l'installation

### Installation depuis les scripts du repo

Si tu veux rester sur le package MSIX et les scripts fournis:

1. Lance `installer/Build-Installer.ps1`
2. Ouvre `installer/dist/Install-SnapSlate.cmd` en administrateur

## Developpement local

### Prerequis

- Windows 10 ou Windows 11
- .NET 10 SDK
- Workload WinUI / Windows App SDK

### Build

```powershell
dotnet build .\SnapSlate.csproj -p:Platform=x64
```

### Lancement en local

Le projet est packe en MSIX. Le lancement fiable en local passe donc par l'enregistrement du manifeste puis par l'app Windows enregistree:

```powershell
Add-AppxPackage -Register .\bin\x64\Debug\net10.0-windows10.0.26100.0\win-x64\AppxManifest.xml -ForceApplicationShutdown
Start-Process explorer.exe "shell:AppsFolder\E3C7A4ED-2E83-45C8-852A-A0469B8C2291_1z32rh13vfry6!App"
```

## Build d'un vrai Setup.exe

```powershell
.\installer\Build-SetupExe.ps1
```

Sortie produite:

- `installer/release/SnapSlate-Setup.exe`

## Structure du projet

- `MainWindow.xaml` et `MainWindow.xaml.cs`: shell principal, onglets, outils d'annotation, import, export, crop
- `ScreenshotDocument.cs`: etat de chaque onglet
- `AnnotationModel.cs`: modele des annotations
- `GradientPaletteDefinition.cs`: definition des palettes degradees
- `installer/`: scripts de packaging, MSIX, Setup.exe et documentation d'installation
- `.github/workflows/build.yml`: CI Windows pour verifier que le projet compile

## Publication GitHub

Le repo contient tout le necessaire pour une publication propre:

- README principal
- changelog
- guide de contribution
- guide securite
- templates d'issues
- template de pull request
- workflow GitHub Actions
- fichier de metadonnees GitHub dans `docs/github/GITHUB_METADATA.md`

## Version actuelle

- Version application: `1.0.0.0`

## Roadmap courte

- Hook systeme complet du raccourci de capture
- Workflow de sauvegarde plus riche
- Export additionnel
- Amelioration de la gestion des historiques d'annotation
