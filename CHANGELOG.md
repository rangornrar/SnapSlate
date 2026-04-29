# Changelog

Tous les changements notables de ce projet seront documentes ici.

Le format est volontairement simple et oriente publication GitHub.

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
