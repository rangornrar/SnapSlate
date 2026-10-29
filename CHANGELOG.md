# Changelog

Tous les changements notables de ce projet seront documentes ici.

Le format est volontairement simple et oriente publication GitHub.

## [2026.10.29.002] - 2026-10-29

### Ergonomie

- Accueil avec collage, import d’image et ouverture de projet ; exemple accessible à la demande.
- Réglages par catégorie, sélection lisible et boutons Annuler / Enregistrer fixes.
- Informations du guide séparées des préférences d’export ; aperçu des réglages explicite.
- Étapes redimensionnables, titres sur deux lignes, notes simplifiées et position dans le guide.
- Export avec choix capture active / guide complet, format et destination ; accès PNG direct conservé dans le menu.
- Flèches et formes regroupées, variante mémorisée et raccourcis hors saisie.
- Cinq couleurs directes, palettes traduites et nuances / opacité dans « Autres couleurs… ».

## [2026.10.29.001] - 2026-10-29

### Ajouté

- Bouton **Copier l’image** à côté de l’export PNG, commande Édition et raccourci **Ctrl + Shift + C**.
- Copie de la capture active avec annotations, masques opaques et recadrage confirmé, produite par le même moteur que l’export PNG.
- Image conservée dans le presse-papiers après fermeture de l’application, sans fichier temporaire à enregistrer.
- Import automatique protégé contre la réimportation de la propre copie de SnapSlate ; le collage manuel reste disponible.
- Libellés, aide clavier et messages de résultat en français et en anglais.

### Vérification

- 84 tests automatisés ; parcours natifs bouton/raccourci, copie/collage et fermeture/réouverture. Comparaison RGBA identique à l’export PNG, masque opaque et recadrage 2400 × 1350 conservés. Détails : `docs/CLIPBOARD_COPY.md`.

## [2026.10.28.002] - 2026-10-28

### Interface d’annotation

- Canevas dégagé au démarrage, étapes et propriétés masquées par défaut ; outils fixes à gauche, hors de la capture.
- Commandes principales regroupées, réglages contextuels et ouverture des propriétés du texte et des gommettes à la demande.
- Ajustement automatique de l’image, pourcentage de zoom exact et exposé aux outils d’assistance.
- Nouvelle annotation sélectionnée, saisie directe des textes et contour de sélection visible sur les fonds clairs et sombres.
- Annulation des formes dessinées corrigée pour restaurer l’état précédant leur création.

### Vérification

- 83 tests automatisés ; parcours natifs d’annotation, sauvegarde/réouverture et export 4K. Captures avant/après et limites : `docs/UX_WORKSPACE.md`.
- Distribution Windows x64 autonome ; signature Authenticode toujours indisponible faute de certificat de production.

## [2026.10.28.001] - 2026-10-28

### Corrigé

- Presse-papiers sur événement, imports historisés et import volontaire répété (`Ctrl+Shift+V`), origine neutre.
- Empreintes d’état sans JSON/Base64 des images, hash d’images immuables réutilisé ; export PNG en tâche de fond.
- Publications Notion et Confluence Cloud v2 clairement textuelles, sans générer les formats illustrés ; titre Notion transmis et erreurs HTTP utiles sans réponse confidentielle.
- Objets sélectionnables par liste, boutons de déplacement de 1 px, noms d’étapes/actions et champs associés à leurs libellés ; ressources de contraste système.
- Ressources FR/EN centralisées, langue explicite des guides et dates, titres DOCX sémantiques et descriptions d’images.
- Sources exportées limitées au nom de fichier ; retrait des scripts MSIX qui ajoutaient un certificat à la confiance système.
- SDK et dépendances verrouillés, version du setup issue du projet, Inno épinglé par SHA-256 en CI, contrôles installation/lancement/désinstallation et provenance des distributions.
- Nettoyage avant publication autonome pour régénérer le PRI WinUI après un build MSIX et éviter un échec de lancement sur les ressources de thème.
- Signature Authenticode optionnelle préparée ; cette release reste non signée, aucun certificat de production n’étant disponible.

### Vérification

- 83 tests automatisés ; profilage explicite pour 1/10/50 étapes. Périmètre et limites : `docs/LOT3_VALIDATION.md`.

## [2026.10.27.002] - 2026-10-27

### Corrigé

- Aperçu et exports d’images produits par un moteur commun : texte, graisse, couleurs, opacité, démo et masquage.
- Recadrage exporté uniquement après confirmation ; sélection provisoire bornée à la source et historisée sans supprimer le recadrage appliqué.
- Titre, modèle, public, auteur et version portés par le guide ; titres et notes propres à chaque étape.
- Format de projet v2 ; migration v1 à partir de la première étape, indépendamment de la sélection.
- Onze outils accessibles sur les petites images grâce à une barre indépendante de la capture et défilante.
- Chargements obsolètes ignorés ; exports sur instantané, progression et annulation entre étapes/formats, nettoyage des packages partiels.
- Images PDF/DOCX décodées en mémoire pour éviter les verrous natifs empêchant le nettoyage.
- Réglages conservés dans un brouillon ; Annuler, Échap et fermeture du panneau abandonnent les modifications. Enregistrer persiste les préférences en arrière-plan.

### Vérification

- 61 tests automatisés, dont migration des métadonnées, état modifié global, recadrage confirmé et identité aperçu/PNG.
- Contrôles Windows et limites détaillés dans `docs/LOT2_VALIDATION.md`.

## [2026.10.27.001] - 2026-10-27

### Corrige

- Confirmation de sauvegarde du projet à la fermeture de la fenêtre.
- Séparation entre export PNG et sauvegarde du projet, y compris après annulation/rétablissement.
- Masques opaques à l’export, indépendants de l’opacité, sans coins arrondis ni transparence.
- Validation des projets, images et géométries avant remplacement de la session.
- Écriture atomique avec conservation de la version précédente dans un fichier `.bak`.
- Gestion des erreurs d’import/export sans acquitter les modifications du projet.
- HTTPS obligatoire pour Confluence, sans redirection HTTP automatique.
- Installateur produit en Release autonome ; versions et dépendances Windows fixées.

### Verification

- 42 tests automatisés de persistance, intégrité des fichiers, état de sauvegarde, masquage, publication et exports texte.
- Contrôle manuel des dialogues de fermeture interrompu par l’utilisateur ; installation interactive non validée.

## [2026.04.29.001] - 2026-04-29

### Ajoute

- Nouveau workflow Git de release et deploiement documente en Markdown
- Site public aligne sur la nouvelle version et le nouvel asset de release
- Refonte de la galerie de captures avec de vraies captures SnapSlate
- Workflow editorial public avec captures, annotations, legendes et publication HTML

### Notes

- La release GitHub doit utiliser l'asset `SnapSlate-Setup.exe`
- Le tag de release conseille est `2026.04.29.001`

## [1.0.0] - 2026-04-23

### Ajoute

- Base WinUI 3 pour l'application `SnapSlate`
- Import manuel d'images
- Import automatique des captures `Win + Shift + S` pendant que l'application est ouverte
- Creation d'un nouvel onglet pour chaque capture
- Fermeture d'onglet avec verification de sauvegarde
- Annotation par texte, gommettes, fleches droites, fleches courbes, rectangles et ovales
- Crop interactif
- Export PNG
- 10 palettes degradees
- Option de reinitialisation de la numerotation des gommettes lors du changement de couleur
- Scripts de generation d'un package installable
- Setup.exe base sur Inno Setup
- Documentation GitHub, templates et CI

### Notes

- Le raccourci `Win + Shift + S` n'est pas encore remplace systeme-wide lorsque l'application est fermee
