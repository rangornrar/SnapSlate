# SnapSlate v2026.10.27.002

Cette version réalise le lot 2 de l’audit : cohérence de l’éditeur et des exports.

- Aperçu et images exportées issus du même moteur : texte, graisse, dégradés, opacité et démo cohérents.
- Recadrage exporté seulement après confirmation ; sélection provisoire bornée et historisée, sans effacer le recadrage précédent.
- Titre, modèle, public, auteur et version du guide indépendants de l’étape sélectionnée.
- Projets v2 et reprise automatique des projets v1 depuis les métadonnées de leur première étape.
- Onze outils accessibles sur les petites captures, dans une barre indépendante et défilante.
- Exports sur instantané, dossiers distincts, progression et annulation avec suppression du package partiel.
- Réglages modifiés dans un brouillon puis enregistrés ou abandonnés avec Annuler/Échap.

## Télécharger

`SnapSlate-Setup.exe` : installateur Windows x64 autonome, .NET et Windows App SDK inclus.
`SnapSlate-Installer-x64.zip` : distribution portable.
`SHA256SUMS.txt` : empreintes des deux fichiers.

## Validation

61 tests automatisés réussis, dont exports concurrents, annulation, migration, fidélité pixel par pixel, génération PDF/DOCX et validation Open XML Office 2019. Builds Debug/Release x64 et CI GitHub vérifiés avant publication. Contrôles Windows : lancement, outils sur capture 100 × 100, abandon des métadonnées par Échap et export complet.

Les projets enregistrés en v2 nécessitent cette version ou une version ultérieure. Conserver le fichier `.bak` pour revenir à une ancienne version. L’annulation attend la fin de l’image ou du format en cours. Installation interactive, autres architectures et scénarios Windows/accessibilité complets restent à vérifier ; détails dans `docs/LOT2_VALIDATION.md`.

L’installateur n’est pas signé Authenticode. Les projets éditables et leurs `.bak` conservent les sources originales ; partager les exports pour diffuser les masquages.
