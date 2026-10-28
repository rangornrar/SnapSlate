# SnapSlate v2026.10.28.001

Cette version réalise le lot 3 de l’audit : performances, historique, publications, accessibilité et chaîne de distribution.

- Presse-papiers traité sur événement et import volontaire répété avec `Ctrl+Shift+V` ; chaque insertion est annulable.
- Contrôle de l’état modifié sans sérialiser les images en Base64 ; export PNG en tâche de fond.
- Publications Notion et Confluence Cloud explicitement limitées au texte, sans générer inutilement les formats illustrés ; erreurs HTTP plus utiles.
- Liste d’objets et déplacements de 1 px accessibles au clavier, noms d’étapes et champs associés à leurs libellés.
- Ressources français/anglais centralisées, langue explicite des documents, titres DOCX sémantiques et descriptions des images.
- Guides partagés sans chemin local complet des fichiers sources.
- SDK et dépendances verrouillés, versions unifiées, setup et archive produits et contrôlés par la CI ; retrait des anciens scripts MSIX.

## Télécharger

- `SnapSlate-Setup.exe` : installateur Windows x64 autonome, .NET et Windows App SDK inclus.
- `SnapSlate-Installer-x64.zip` : distribution portable à extraire entièrement ; lancer `app/SnapSlate.exe`.
- `SHA256SUMS.txt` : empreintes des deux distributions.
- `build-provenance.json` : révision construite, SDK, état de signature et empreintes.

## Validation et limites

83 tests automatisés, builds x64, parcours d’export FR/EN, validation Open XML et contrats HTTP. La CI vérifie installation silencieuse, lancement et désinstallation sur un runner Windows jetable. Contrôles Windows : sélection d’objet, déplacement au clavier, annulation et aperçu de langue abandonné. Détails et mesure ciblée 1/10/50 étapes dans `docs/LOT3_VALIDATION.md`.

**Cette release n’est pas signée Authenticode : aucun certificat de production n’est disponible.** La chaîne de signature est préparée pour un futur certificat et devra être testée lors de son activation.

Notion et Confluence publient le texte et les légendes sans captures. Narrator, les thèmes contrastés, les autres architectures et la matrice Windows 10/11 ne sont pas entièrement validés. Les projets éditables et `.bak` conservent les sources originales ; partager les exports pour diffuser les masquages.
