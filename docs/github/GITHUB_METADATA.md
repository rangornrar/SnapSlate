# GitHub Metadata

Ce fichier contient le texte pret a copier dans GitHub pour la creation du depot, la page About et la premiere release.

## Nom du depot

`SnapSlate`

## Description courte pour GitHub

WinUI 3 screenshot editor for Windows with tabs, crop, arrows, stickers, text and gradient annotation palettes.

## Description plus orientee produit

SnapSlate est un outil Windows de capture et d'annotation qui ouvre chaque screenshot dans un nouvel onglet et permet ensuite de cropper, ecrire, entourer, flecher et annoter rapidement avec plusieurs palettes degradees.

## Topics suggeres

- `windows`
- `winui3`
- `windows-app-sdk`
- `screenshot`
- `annotation`
- `image-editor`
- `desktop-app`
- `productivity`
- `msix`

## Texte court pour le haut du README ou la section About

Turn `Win + Shift + S` captures into editable tabs, then annotate them with text, stickers, arrows, shapes and gradient palettes.

## Positionnement honnete

Point important a garder dans la description publique:

- le flux `Win + Shift + S` fonctionne tant que l'application est ouverte
- le remplacement complet du raccourci systeme quand l'app est fermee fait partie de la roadmap, pas de la version actuelle

## Premiere release GitHub

### Titre

`v1.0.0 - First public SnapSlate build`

### Corps

```md
## SnapSlate v1.0.0

Premiere version publiable de SnapSlate, une application Windows d'annotation de captures d'ecran construite avec WinUI 3.

### Inclus dans cette version

- import automatique des captures `Win + Shift + S` pendant que l'application est ouverte
- un onglet par capture
- crop interactif
- texte, gommettes, fleches droites, fleches courbes, rectangles et ovales
- 10 palettes degradees
- reinitialisation optionnelle de la numerotation des gommettes lors du changement de couleur
- export PNG
- setup Windows pour simplifier l'installation

### Fichier a attacher a la release

- `SnapSlate-Setup.exe`

### Limitation actuelle

`SnapSlate` n'intercepte pas encore le raccourci systeme si l'application est fermee. Le flux automatique passe aujourd'hui par l'ecoute du presse-papiers pendant l'execution de l'application.
```

## Assets conseilles pour le depot

- une capture de l'ecran principal de l'app
- une capture montrant plusieurs onglets ouverts
- une capture montrant les palettes degradees

## Licence

Je n'ai pas cree de licence automatiquement ici, car c'est un choix juridique.

Avant publication publique, choisis explicitement une licence:

- MIT si tu veux quelque chose de tres permissif
- Apache-2.0 si tu veux une option permissive avec clauses brevets
- GPL-3.0 si tu veux une redistribution copyleft
- ou pas de licence si tu gardes le code proprietaire
