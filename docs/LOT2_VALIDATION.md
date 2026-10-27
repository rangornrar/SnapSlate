# Lot 2 — validation du 27 octobre 2026

Version : `2026.10.27.002` ; application et installateur : `2026.10.27.2`.

## Corrections et preuves

| Audit | Correction | Vérification |
| --- | --- | --- |
| F08 | Même moteur GDI+ pour l’image d’aperçu et les images des exports ; texte, graisse, transparence et dégradés cohérents. Le fond de démonstration est désormais le fond raster commun. Le calque WinUI sert à sélectionner, déplacer et saisir le texte. | Identité des PNG mémoire/disque avec tous les types d’annotation ; opacité nulle ; crop exact pixel par pixel sur source et démo. |
| F09 | Export du recadrage appliqué uniquement ; sélection provisoire conservant le précédent, bornée à l’image, alignée sur les pixels et historisée. La sélection provisoire montre le cadre complet pour permettre sa révision. | Tests du crop en attente, du remplacement provisoire et des coordonnées hors image. |
| F10 | Métadonnées globales clonées dans l’état du projet, l’historique et l’instantané d’export. Titre d’étape indépendant. | Aller-retour v2, migration v1 depuis la première étape malgré une autre sélection, modification et retour à l’état sauvegardé. |
| F11 | Rail indépendant du viewport de l’image ; boutons défilants selon la hauteur disponible ; taille indépendante du zoom. | Contrôle Windows : les onze outils sont visibles sur une image synthétique 100 × 100. |
| F12 | Identifiant de chargement pour ignorer les décodages obsolètes ; rendu d’aperçu regroupé dans une seule tâche ; générateur de packages extrait et testé ; instantané, dossier unique, progression, annulation et nettoyage. | Exports concurrents, modification du guide et inversion des étapes pendant l’export, annulation après la première image et source corrompue. Un fichier préexistant est préservé. |
| F18 | Brouillons séparés pour les préférences et les métadonnées ; écriture atomique des préférences en arrière-plan uniquement avec Enregistrer ; abandon par Annuler, Échap ou retour à la procédure. | Contrôle Windows : auteur saisi puis Échap, réouverture avec auteur vide ; dossier d’export temporaire utilisé sans enregistrer les préférences. Revue des autres chemins de sortie. |

## Résultats

- 61 tests MSTest exécutés sur les sources de production, sans remplacement de leur logique.
- Génération réelle de Markdown, HTML, PDF, DOCX et PNG dans le test de package complet.
- Validation du DOCX par OpenXmlValidator, cible Office 2019 : zéro erreur. Image incorporée identique au PNG du package.
- Lancement de l’application Debug x64 et export complet dans l’interface : fichiers générés, message de réussite.
- Builds Debug et Release x64 exigés avant publication ; compilation et suite de tests répétées par GitHub Actions sur le commit de la PR.

Les projets v1 sont lus et migrés en mémoire. Le prochain enregistrement écrit la version 2. Les champs historiques des étapes sont conservés pour ne pas perdre leur contenu, mais les exports utilisent les métadonnées du guide. Les versions anciennes de SnapSlate refusent le format v2 ; conserver le `.bak` avant une rétrogradation.

## Limites

L’annulation est coopérative entre images et formats : une opération GDI+/PDF/DOCX commencée termine son étape avant le nettoyage. Le contrôle du bouton d’annulation par clic, les changements de sélection très rapides, les déplacements/recadrages manuels complets, l’installation et désinstallation, ARM64/x86, Windows 10 minimum, Narrateur et contraste élevé restent à vérifier. Aucun résultat de performance sur de grands projets n’est revendiqué. Les tests d’identité concernent l’image raster, pas les repères de sélection ou le champ de saisie actif.

Les publications Notion/Confluence/SharePoint et les autres lots de l’audit ne sont pas validés par cette version. L’installateur x64 reste sans signature Authenticode.
