# Qualité et validation

## Contrôles automatisés

```powershell
dotnet test .\tests\SnapSlate.Tests.csproj -c Release -p:Platform=x64 --filter "TestCategory!=Performance"
```

Les sept tests initiaux couvrent : aller-retour JSON des coordonnées, champs géométriques inconnus imbriqués, types JSON invalides, échappement des contenus HTML, langue anglaise, notes multilignes et inclusion/exclusion des références d’images Markdown.

Le lot 1 porte la suite à 42 tests : sauvegarde atomique (y compris panne et verrouillage), état modifié après PNG/Undo/écriture concurrente, validation des projets et géométries, caviardage pixel par pixel pour quatre opacités, conservation d’un export si la source est invalide et validation HTTPS de Confluence.

Le projet de tests utilise les fichiers de production `ProjectJson`, `GuideTextExporter` et `GuideExportModels`, sans mocks de leur logique. Il n’exige ni capture d’écran ni démarrage de WinUI. La CI Windows compile séparément l’application entière puis lance les tests.

Le lot 2 porte la suite à 61 tests : métadonnées globales et migration v1/v2, indépendance des brouillons, recadrage provisoire ignoré à l’export, conservation du recadrage appliqué, bornage des coordonnées, transparence et identité des PNG produits pour l’aperçu et le disque. Voir [validation du lot 2](LOT2_VALIDATION.md).

Le lot 3 porte la suite à 83 tests : coordination du presse-papiers, déplacements, empreintes, parité FR/EN, confidentialité des sources, titres et langue DOCX/PDF, contrat HTTP Notion et Confluence Cloud v2 et erreurs 401/403/429 sans secrets. Un test de profilage séparé nécessite une image et un chemin de rapport explicites. Les tests HTTP utilisent un transport factice et ne publient aucune page réelle. Voir [validation du lot 3](LOT3_VALIDATION.md).

La CI produit aussi le setup et l’archive portable, vérifie versions et contenu, puis effectue installation silencieuse, lancement et désinstallation sur Windows Server 2022. Ce test ne remplace pas une matrice Windows 10/11 ni le parcours interactif complet de l’installateur.

## Vérification manuelle avant une version distribuée

- Installer et lancer sous les versions Windows ciblées, puis désinstaller.
- Importer deux captures et vérifier onglets, modification et fermeture.
- Sauvegarder, rouvrir et comparer images, annotations et recadrages.
- Vérifier annulation/rétablissement et prévenir la perte des modifications.
- Exporter PNG, PDF, DOCX, HTML et Markdown ; ouvrir chaque résultat et vérifier texte, images et légendes.
- Vérifier les masquages sur l’image exportée et les limites des projets éditables.
- Mesurer le comportement avec plusieurs grandes captures.

Cette liste décrit les contrôles à effectuer ; elle n’affirme pas qu’ils ont tous été exécutés. Les tests actuels ne fournissent pas de taux de couverture global ni de garantie de rendu visuel.
