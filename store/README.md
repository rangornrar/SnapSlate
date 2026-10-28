# Microsoft Store — MSIX Windows x64

## État

La chaîne produit un package autonome (.NET et Windows App SDK inclus), son archive `.msixupload` et les symboles de diagnostic. Le Microsoft Store signe le package après certification : aucun certificat personnel ni clé privée n'est nécessaire à cette étape.

Le compte développeur et le nom SnapSlate ne sont pas encore réservés. Sans l'identité Partner Center, seul le mode **PREVIEW-NOT-FOR-SUBMISSION** est utilisable. Ce package de contrôle n'est pas une soumission Store et ne doit pas y être envoyé.

## À faire dans Partner Center

1. Créer un compte depuis [Microsoft Store Developer](https://storedeveloper.microsoft.com/) avec le nouveau parcours gratuit ; choisir le type adapté à son activité. Pour un projet personnel non professionnel, choisir **Individual**.
2. Effectuer soi-même la vérification d'identité et accepter les conditions du compte.
3. Créer une application **MSIX**, puis réserver « SnapSlate » si le nom est disponible.
4. Dans **Product management / Product identity**, relever exactement les valeurs sensibles à la casse :
   - `Package/Identity/Name`
   - `Package/Identity/Publisher`
   - `Package/Properties/PublisherDisplayName`

Ces trois identifiants ne sont pas des secrets. Ne pas transmettre de mot de passe, document d'identité ou clé privée dans le dépôt.

## Construire

Prérequis : Windows, SDK .NET de `global.json`, outils Windows/WinUI et outils MSVC de Visual Studio (notamment `mspdbcmf.exe` pour les symboles). Seule l'architecture x64 est validée par cette chaîne.

Depuis la racine du dépôt, aperçu technique :

```powershell
.\store\Build-StorePackage.ps1 -Preview -Platform x64
```

Après réservation dans Partner Center :

```powershell
.\store\Build-StorePackage.ps1 -Platform x64 `
  -IdentityName 'VALEUR_PACKAGE_IDENTITY_NAME' `
  -Publisher 'CN=GUID_FOURNI_PAR_PARTNER_CENTER' `
  -PublisherDisplayName 'NOM_AFFICHE_FOURNI_PAR_PARTNER_CENTER'
```

Les variables `SNAPSLATE_STORE_IDENTITY_NAME`, `SNAPSLATE_STORE_PUBLISHER` et `SNAPSLATE_STORE_PUBLISHER_DISPLAY_NAME` peuvent aussi fournir ces valeurs. L'absence d'identité ou une identité de développement bloque le mode soumission avant compilation.

Le manifeste Store est généré dans un nouveau dossier `AppPackages/Store/` ; le script ne remplace pas l'identité source ni les versions de la distribution GitHub. Il nettoie les ressources intermédiaires avec les propriétés du build MSIX avant compilation pour éviter de réutiliser un PRI du build autonome.

Le contrôle du package vérifie l'identité, la version, l'architecture, la présence des runtimes et ressources, l'absence de signature locale et les capacités demandées. `store-build.json` contient les SHA-256, la révision Git et l'état du mode aperçu. Ces contrôles ne remplacent ni le Windows App Certification Kit ni la certification Microsoft.

## Version du package

La quatrième composante est réservée au Store et reste à zéro. La version applicative `année.mois.jour.release` devient `année.mois.(jour × 1000 + release).0` : `2026.10.28.2` donne `2026.10.28002.0`.

La release du jour doit être comprise entre 0 et 999. `-PackageVersion` permet une version explicite à quatre composantes finissant par `.0`. Lors d'une mise à jour, vérifier qu'elle est supérieure à la dernière version acceptée dans Partner Center.

## GitHub Actions

Le workflow **Microsoft Store MSIX** construit un aperçu sur les PR concernées. Après intégration à `main`, **Run workflow** permet de générer le candidat Store : désactiver `preview` et renseigner les trois valeurs Partner Center. Les entrées sont passées par l'environnement puis échappées dans le XML ; elles ne sont pas interpolées dans le script du workflow.

L'artefact contient l'archive `.msixupload`, le manifeste et la provenance. Il ne déclenche aucune soumission ni publication sur le Store.

## Fiche et certification

- Textes prêts à adapter : [français](listing/fr-FR.md), [anglais](listing/en-US.md).
- Logo de fiche : [icon-300.png](listing/icon-300.png), dérivé de l'icône existante de SnapSlate.
- Captures natives de l'application : [canevas](../docs/ux-2026-10-28/after-default.png), [fenêtre maximisée](../docs/ux-2026-10-28/after-maximized.png). Le contenu de démonstration et la langue doivent correspondre à la fiche choisie.
- Politique de confidentialité de l'application : [PRIVACY.md](PRIVACY.md). Son URL publique doit être accessible avant soumission ; la politique du site web est distincte.
- Dans les notes de certification, justifier `runFullTrust` avec le texte de la fiche. Aucune permission `systemAIModels` n'est requise : l'OCR utilise `Windows.Media.Ocr`.
- Compléter prix/disponibilité, propriétés, questionnaire d'âge et coordonnées de support dans Partner Center.
- Valider le candidat avec le [Windows App Certification Kit](https://learn.microsoft.com/en-us/windows/uwp/debug-test-perf/windows-app-certification-kit) sur un poste de test. La certification Microsoft reste nécessaire avant toute disponibilité publique.

Références : [identité et exigences MSIX](https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msix/app-package-requirements), [build WinUI pour le Store](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/ci-for-winui3), [créer une soumission](https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msix/create-app-submission).
