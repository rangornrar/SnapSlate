# Procedure de deploiement Git - v2026.04.29.001

Cette procedure decrit le cycle complet pour livrer la version `2026.04.29.001` de SnapSlate.

## 1. Creer la branche de travail

Depuis `main`, cree la branche de livraison:

```powershell
git checkout main
git pull origin main
git checkout -b v20260429_001
```

## 2. Mettre a jour la version

Les fichiers a synchroniser pour cette release sont:

- `Package.appxmanifest` pour la version Windows du package
- `site/app.js` et `site/index.html` pour la version publiee et les liens de telechargement
- `docs/github/GITHUB_METADATA.md` pour les textes GitHub
- `docs/github/RELEASE_v2026.04.29_001.md` pour la note de release
- `CHANGELOG.md` pour l'historique

## 3. Construire et verifier

```powershell
dotnet build .\SnapSlate.csproj -c Debug -p:Platform=x64
.\installer\Build-SetupExe.ps1
```

Le setup attendu est:

- `installer\release\SnapSlate-Setup.exe`

## 4. Commit et push

```powershell
git status
git add .
git commit -m "Release v2026.04.29.001"
git push -u origin v20260429_001
```

## 5. Ouvrir la pull request

Creer une PR de `v20260429_001` vers `main`, puis attendre:

- la compilation CI
- les verifications visuelles du site
- la verification du setup

## 6. Merger dans main

Une fois la PR validee:

```powershell
git checkout main
git pull origin main
git merge --no-ff v20260429_001
git push origin main
```

## 7. Taguer la release

Utilise un tag GitHub cohérent avec le site:

- `2026.04.29.001`

Puis publie la release GitHub avec:

- titre: `SnapSlate v2026.04.29.001`
- asset: `SnapSlate-Setup.exe`

## 8. Mettre a jour le site

Les liens publics du site doivent pointer vers:

- `https://github.com/rangornrar/SnapSlate/releases/download/2026.04.29.001/SnapSlate-Setup.exe`

Le badge de version du site doit afficher:

- `V2026.04.29.001`

