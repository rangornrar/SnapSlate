# Préparation MSIX — validation du 28 octobre 2026

## Vérifié localement

- Package x64 autonome et archive `.msixupload` générés en Release, sans erreur ni avertissement. Symboles `.appxsym` inclus.
- Application `2026.10.28.2` ; package Store `2026.10.28002.0`, quatrième composante réservée à zéro.
- Identité de développement distincte `SnapSlate.StorePreview`, marquage explicite **PREVIEW-NOT-FOR-SUBMISSION** ; aucune identité Partner Center n'est inventée.
- Contrôle du manifeste du MSIX de test et du MSIX réellement inclus dans l'archive Store : identité, éditeur, version, x64, Windows Desktop, `runFullTrust` seul, composants .NET/WinUI, ressources et absence de signature locale.
- Refus avant compilation d'une identité manquante en mode soumission, du mélange aperçu/identité et d'une version dont la quatrième composante n'est pas zéro.
- 83 tests automatisés réussis.
- Enregistrement du layout de développement avec `Add-AppxPackage -Register`, sans importer de certificat ni modifier la confiance du système.
- Lancement via l'identité MSIX, éditeur affiché, ouverture d'un projet avec deux captures et une annotation sur une image 3840 × 2160.
- Export PNG 3840 × 2160 depuis l'application MSIX ; comparaison RGB identique pixel par pixel à l'export de contrôle de la distribution classique.
- Fermeture de l'application puis désenregistrement du seul package de test créé pour cette vérification. L'installation de développement antérieure et les projets existants restent intacts.

Les images du manifeste provenaient encore du modèle WinUI. Elles ont été remplacées par des déclinaisons déterministes de l'icône existante `Assets/AppIcon.ico`, également intégrée à l'exécutable. Le chargement de l'icône de fenêtre utilise le répertoire de l'application et le contrôle XAML une image PNG, indépendamment du répertoire courant d'activation.

## À terminer avant publication Store

- Création et vérification du compte développeur, réservation du nom et association aux trois valeurs Partner Center.
- Construction du candidat avec l'identité réellement attribuée, puis vérification des fichiers obtenus.
- Windows App Certification Kit sur un poste de test avec session active et droits requis. La session actuelle n'est pas élevée ; aucune certification WACK n'est revendiquée.
- Validation sur les versions Windows et matériels ciblés, dont l'OCR avec les langues installées.
- Fiche, coordonnées de support, prix/marchés, questionnaire d'âge et URL publique de confidentialité dans Partner Center.
- Soumission et certification par Microsoft. Le package de contrôle ne doit pas être envoyé au Store.

L'inscription, la vérification d'identité et l'acceptation des conditions du compte sont réalisées par le titulaire. Aucun mot de passe, document d'identité ni certificat privé n'est nécessaire dans le dépôt.
