# Changelog

Tous les changements notables de ce projet seront documentes ici.

Le format est volontairement simple et oriente publication GitHub.

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

- 41 tests automatisés de persistance, intégrité des fichiers, état de sauvegarde, masquage, publication et exports texte.
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
