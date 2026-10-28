# Distribution Windows x64

Depuis la racine du dépôt, avec le SDK de `global.json` et Inno Setup de `eng/toolchain.json` :

```powershell
.\installer\Build-SetupExe.ps1 -Configuration Release -Platform x64
```

Le script publie dans une sortie propre, autonome (.NET et Windows App SDK inclus), restaure les dépendances verrouillées, prépare `dist/app`, l’archive portable `SnapSlate-Installer-x64.zip`, puis `release/SnapSlate-Setup.exe`. Il échoue si une commande échoue. La version du setup vient du projet, sans numéro indépendant dans le script Inno.

Un `dotnet clean` avec les propriétés de déploiement autonome précède le publish : un build MSIX antérieur peut sinon laisser un PRI incomplet dans les fichiers intermédiaires et empêcher le chargement des thèmes WinUI au lancement.

`Test-Distribution.ps1` vérifie versions, composants et égalité des binaires ZIP/staging ; il produit `release/SHA256SUMS.txt` et `release/build-provenance.json`. La provenance contient la révision Git de construction, le SDK, les empreintes et l’état de signature. Il s’agit d’un relevé de build, pas d’une attestation cryptographique.

En CI seulement, `eng/Install-CiBuildTools.ps1` télécharge le compilateur épinglé et vérifie son SHA-256. `Test-Distribution.ps1 -InstallSmoke` installe dans un répertoire temporaire neuf du runner, vérifie le binaire installé, lance l’application et attend sa fenêtre, puis arrête cette instance et désinstalle. Cette option est refusée hors GitHub Actions ; elle ne doit pas servir sur un poste utilisateur.

## Installer ou utiliser la version portable

Exécuter le setup, puis lancer SnapSlate depuis le menu Démarrer. Désinstaller depuis les applications Windows. Le setup nécessite les droits administrateur. Pour la version portable, extraire l’archive complète puis lancer `app/SnapSlate.exe` ; conserver tous ses fichiers voisins.

Les anciens scripts d’installation MSIX ont été retirés. Aucun certificat n’est importé dans les magasins de confiance.

## Signature Authenticode

**La release 2026.10.28.001 n’est pas signée : aucun certificat de production n’est disponible.**

Avec un certificat de signature de code valide et sa clé privée déjà disponible dans `CurrentUser/My`, utiliser :

```powershell
.\installer\Build-SetupExe.ps1 -SigningCertificateThumbprint '<empreinte SHA-1 du certificat>'
```

Ou définir `SNAPSLATE_SIGNING_CERT_THUMBPRINT`. SignTool du SDK Windows signe l’exécutable et l’assembly SnapSlate en SHA-256 avec horodatage HTTPS ; Inno signe le désinstalleur et le setup. Un échec de signature interrompt le build, sans repli silencieux. `Sign-Binary.ps1` vérifie validité et empreinte après signature. Aucun mot de passe de certificat ni clé privée n’est stocké dans les scripts. Le chemin signé est préparé mais n’a pas été exécuté sans certificat ; le résultat devra être validé lors de son activation.

## Publier une release

Créer une PR, attendre les contrôles de build/tests/distribution, récupérer l’artefact `distribution-x64` du run validé puis comparer ses SHA-256. Après fusion, créer la release et joindre les deux distributions, `SHA256SUMS.txt` et `build-provenance.json`. Conserver la provenance qui indique la révision construite, même si le tag cible le commit de fusion. Le build est contrôlé ; aucune reproductibilité binaire bit à bit n’est revendiquée.
