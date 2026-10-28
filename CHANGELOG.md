# Changelog

Tous les changements notables de ce projet seront documentes ici.

Le format est volontairement simple et oriente publication GitHub.

## [2026.10.28.001] - 2026-10-28

### Corrigé

- Presse-papiers sur événement, imports historisés et import volontaire répété (`Ctrl+Shift+V`), origine neutre.
- Empreintes d’état sans JSON/Base64 des images, hash d’images immuables réutilisé ; export PNG en tâche de fond.
- Publications Notion et Confluence Cloud v2 clairement textuelles, sans générer les formats illustrés ; titre Notion transmis et erreurs HTTP utiles sans réponse confidentielle.
- Objets sélectionnables par liste, boutons de déplacement de 1 px, noms d’étapes/actions et champs associés à leurs libellés ; ressources de contraste système.
- Ressources FR/EN centralisées, langue explicite des guides et dates, titres DOCX sémantiques et descriptions d’images.
- Sources exportées limitées au nom de fichier ; retrait des scripts MSIX qui ajoutaient un certificat à la confiance système.
- SDK et dépendances verrouillés, version du setup issue du projet, Inno épinglé par SHA-256 en CI, contrôles installation/lancement/désinstallation et provenance des distributions.
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
