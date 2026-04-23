# Contributing

Merci de vouloir contribuer a `SnapSlate`.

## Avant de commencer

- Ouvre une issue si le changement est important ou structurel
- Reste coherent avec le fonctionnement actuel de l'application
- Ne masque pas les limitations existantes: documente-les clairement

## Setup local

```powershell
dotnet restore .\SnapSlate.csproj
dotnet build .\SnapSlate.csproj -p:Platform=x64
```

## Regles de contribution

- Garde l'interface en francais si tu modifies des libelles deja existants
- Evite les regressions dans le flux principal: capture, annotation, export
- Si tu modifies les onglets, teste la fermeture avec contenu non sauvegarde
- Si tu modifies le flux de capture, teste `Win + Shift + S` avec l'application deja ouverte
- Si tu modifies le packaging, verifie aussi les scripts dans `installer/`

## Verification minimale avant PR

- Build local reussi
- Verification manuelle de l'ouverture de la fenetre
- Verification d'au moins un import, une annotation et un export PNG
- Documentation mise a jour si le comportement utilisateur change

## Pull requests

- Garde les PR ciblees
- Ajoute une description claire du probleme et de la solution
- Liste les tests manuels effectues
