# SnapSlate v2026.10.27.001

Cette version corrige le lot 1 de l’audit : intégrité du projet et confidentialité des captures exportées.

- La fermeture de la fenêtre demande d’enregistrer, d’ignorer ou d’annuler quand le projet est modifié.
- Exporter un PNG ne marque plus le projet éditable comme enregistré. Annuler/rétablir compare l’état à la dernière sauvegarde réelle.
- Les masques sont opaques dans les PNG et les images intégrées aux guides, quelle que soit leur ancienne opacité. Les coins du rectangle sont couverts et les masques sont appliqués après les autres annotations.
- Les projets invalides, versions inconnues, images trop grandes ou géométries invalides sont refusés avant de remplacer la session.
- La sauvegarde écrit un fichier temporaire puis remplace le fichier final ; la version précédente est conservée dans `.bak`.
- Les erreurs d’import/export conservent le projet ouvert et produisent un message récupérable.
- Confluence exige HTTPS ; les identifiants dans l’URL et les redirections automatiques sont refusés.

## Installation

Télécharger `SnapSlate-Setup.exe` et lancer l’installateur Windows x64. L’application est publiée en Release autonome avec .NET et le Windows App SDK embarqués. `SHA256SUMS.txt` permet de vérifier les fichiers distribués.

## Validation et limites

41 tests automatisés passent. Les builds Debug et Release x64 sont vérifiés. Le contrôle manuel complet des dialogues Windows a été interrompu par la touche Échap ; les tests d’installation interactive, ARM64/x86 et Windows 10 minimum restent à réaliser.

Les fichiers `.snapslate` et leurs sauvegardes `.bak` conservent les images originales : ils ne sont pas expurgés. Pour partager une image masquée, utiliser le PNG ou les images des guides exportés. Cette version ne corrige pas encore tous les écarts de fidélité du rendu, le rail d’outils sur petites images ou les publications distantes sans pièces jointes.

L’installateur ne dispose pas de signature Authenticode.
