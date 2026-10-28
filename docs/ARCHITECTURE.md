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
| `ClipboardImportCoordinator` | Coalescence des événements, imports manuels répétés et abandon des lectures obsolètes |
| `ProjectContentFingerprint`, `ProjectSaveTracker` | État modifié sans sérialisation Base64 ; empreintes d’images immuables mises en cache |
| `GuidePackageBuilder`, `ExportSource` | Instantanés d’export, progression, annulation et source sans chemin privé |
| `TextPublicationService`, `PublicationTextBuilder` | Transport HTTP testable et contenu textuel sans construire les fichiers illustrés |
| `LocalizedText`, `Strings/*/Resources.resw` | Ressources communes à l’interface et aux exporteurs, culture explicite |
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

La CI compile Release x64, teste les services de production (dont parcours import/annotation/sauvegarde/réouverture/export et génération PDF/DOCX), puis construit et installe la distribution sur un runner Windows jetable. Elle vérifie le lancement et la désinstallation avant de conserver les artefacts. Le raccordement complet des interactions WinUI, Narrator, les thèmes contrastés et les publications sur des comptes réels demandent encore des essais ciblés. Voir [QUALITY.md](QUALITY.md) et [LOT3_VALIDATION.md](LOT3_VALIDATION.md).

Les tableaux d’images sont immuables après import : les snapshots d’historique les partagent, leur SHA-256 est calculé une fois par tableau et mis en cache faible. Toute future modification de pixels doit affecter un nouveau tableau. Le JSON complet est réservé aux sauvegardes. La séparation des services est progressive ; la classe partielle MainWindow reste le coordinateur de l’interface.
