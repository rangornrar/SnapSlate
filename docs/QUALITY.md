# Qualité et validation

## Contrôles automatisés

```powershell
dotnet test .\tests\SnapSlate.Tests.csproj -c Release
```

Sept tests couvrent : aller-retour JSON des coordonnées, champs géométriques inconnus imbriqués, types JSON invalides, échappement des contenus HTML, langue anglaise, notes multilignes et inclusion/exclusion des références d’images Markdown.

Le projet de tests utilise les fichiers de production `ProjectJson`, `GuideTextExporter` et `GuideExportModels`, sans mocks de leur logique. Il n’exige ni capture d’écran ni démarrage de WinUI. La CI Windows compile séparément l’application entière puis lance les tests.

## Vérification manuelle avant une version distribuée

- Installer et lancer sous les versions Windows ciblées, puis désinstaller.
- Importer deux captures et vérifier onglets, modification et fermeture.
- Sauvegarder, rouvrir et comparer images, annotations et recadrages.
- Vérifier annulation/rétablissement et prévenir la perte des modifications.
- Exporter PNG, PDF, DOCX, HTML et Markdown ; ouvrir chaque résultat et vérifier texte, images et légendes.
- Vérifier les masquages sur l’image exportée et les limites des projets éditables.
- Mesurer le comportement avec plusieurs grandes captures.

Cette liste décrit les contrôles à effectuer ; elle n’affirme pas qu’ils ont tous été exécutés. Les tests actuels ne fournissent pas de taux de couverture global ni de garantie de rendu visuel.
