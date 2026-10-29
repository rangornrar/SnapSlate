# Copier la capture annotée — 29 octobre 2026

Le bouton **Copier l’image**, à côté d’**Exporter PNG**, copie la capture active avec ses annotations. La même commande est disponible dans Édition et avec **Ctrl + Shift + C**. Elle ne remplace pas le comportement de Ctrl + C dans les champs de texte.

Le rendu PNG est calculé sur un instantané en arrière-plan avec `GuideImageRenderer.RenderPng`, comme l’export. Il conserve la résolution de la source et applique uniquement le recadrage confirmé. Les masques sont opaques et le contour de sélection n’est pas rendu. Aucun chemin de fichier, note de guide ou projet éditable n’est ajouté au presse-papiers.

Le contenu est publié comme bitmap Windows et conservé avec `Clipboard.Flush` avant libération du flux mémoire. Un format personnalisé marque les images produites par SnapSlate pour éviter leur réimportation automatique, y compris entre plusieurs sessions de l’application. La présence du format suffit : aucune désérialisation du marqueur n’est nécessaire après sa mise en cache par Windows. Le collage manuel reste possible, ainsi que l’import des captures produites par les autres applications. Les commandes sont désactivées pendant la copie ; les échecs de rendu ou de presse-papiers donnent un message localisé.

## Vérifications réalisées

- Build Debug x64 sans erreur ni avertissement.
- 84 tests automatisés réussis, dont un nouveau contrôle de l’instantané : recadrage et masque opaque conservés lorsque le document d’origine change.
- Parcours natif avec le bouton et avec Ctrl + Shift + C ; état occupé, résultat et nom accessible exposés par l’interface.
- Une capture avec trois annotations reste sélectionnée après copie, sans import automatique de la propre image de SnapSlate.
- Collage manuel avec Ctrl + Shift + V : nouvelle capture contenant le rendu fusionné, sans objets d’annotation superposés.
- Source 3840 × 2160, recadrage confirmé 2400 × 1350, rectangle sélectionné, texte et masque demandé à opacité 10 %. Comparaison entre export PNG et copie/collage/export : dimensions et tous les pixels RGBA identiques. Le masque reste intégralement opaque `(18, 18, 18, 255)` ; aucun contour de sélection n’est copié.
- Fermeture de la session source après copie, relancement, puis collage manuel : image annotée et masque toujours disponibles.
- Les ressources FR/EN gardent les mêmes clés et paramètres de format, vérifiés par les tests.

Les formats spécifiques de toutes les applications de destination ne sont pas testés. Le protocole utilisé est le bitmap standard du presse-papiers Windows. Les essais Narrator, thèmes de contraste et autres DPI restent à effectuer ; aucune validation complète d’accessibilité n’est revendiquée.

Références : [copie d’images Microsoft](https://github.com/microsoft/Windows-universal-samples/blob/main/Samples/Clipboard/cs/CopyImage.xaml.cs), [SetContentWithOptions](https://learn.microsoft.com/en-us/uwp/api/windows.applicationmodel.datatransfer.clipboard.setcontentwithoptions), [conservation après fermeture](https://learn.microsoft.com/en-us/uwp/api/windows.applicationmodel.datatransfer.clipboard.flush).
