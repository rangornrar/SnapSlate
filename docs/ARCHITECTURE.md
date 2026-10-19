# Architecture de SnapSlate

## Besoin et contraintes

Transformer des captures Windows en une suite d’étapes annotées, sauvegardable et exportable. L’application est locale, ciblée Windows, avec une interface WinUI 3 et des dialogues système. Le presse-papiers est surveillé uniquement pendant son exécution.

## Responsabilités actuelles

| Composant | Responsabilité |
| --- | --- |
| `MainWindow.xaml` et fichiers `MainWindow.*.cs` | Interface, interactions, coordination des projets et publication |
| `ScreenshotDocument`, `AnnotationModel` | État des captures et annotations, notifications de changement |
| `ProjectPersistence` | Adaptation entre état complet du projet et sérialisation |
| `ProjectJson` | JSON camelCase, enums et convertisseurs Point/Rect |
| `GuideImageRenderer` | Rendu des images annotées |
| `GuideTextExporter` | Construction HTML et Markdown sans dépendance à l’interface |
| `GuideExporter` | Production PDF / DOCX avec QuestPDF et Open XML |
| `installer/` | Staging exécutable et setup Inno Setup |

Le flux principal est : presse-papiers ou import → document → annotations → rendu d’image → manifeste d’export → fichiers de sortie. La sauvegarde de projet suit un autre flux : état du projet → `ProjectPersistence` → `ProjectJson` → fichier `.snapslate`.

## Arbitrages et dette technique

L’interface conserve une part importante de coordination et de logique dans une classe partielle `MainWindow`. La séparation en fichiers facilite la lecture, mais ne constitue pas à elle seule une séparation architecturale. Les services de JSON et d’export textuel ont été isolés pour tester les formats sans lancer WinUI.

Le projet de tests compile directement ces mêmes fichiers de production. Ce choix évite une dépendance au démarrage MSIX/WinUI ; il ne teste pas le raccordement complet de l’interface. Une bibliothèque dédiée deviendra préférable si d’autres consommateurs réutilisent ces services.

Le format JSON conserve les images et l’état en mémoire. Le coût mémoire dépend de la taille et du nombre de captures ainsi que des snapshots d’historique. Aucun budget mémoire maximal n’est garanti par cette documentation.

## Données et points de contrôle

- Les captures et projets peuvent contenir des informations confidentielles : leur partage reste une action explicite de l’utilisateur.
- Les textes injectés dans les exports HTML sont encodés ; des tests couvrent balises, guillemets et retours à la ligne.
- Les champs JSON inconnus contenant un objet ou un tableau sont ignorés correctement par les convertisseurs géométriques.
- Le masquage visuel sur un export ne garantit pas l’effacement du contenu source dans un projet éditable. Pour partager une capture masquée, vérifier le fichier exporté plutôt que transmettre le projet d’origine.

## Limites de validation

La CI vérifie la compilation Release x64 et les tests des services isolés. Le presse-papiers, le rendu WinUI, l’installation, les exports PDF/DOCX et la qualité visuelle nécessitent encore des essais Windows. Voir [QUALITY.md](QUALITY.md).
