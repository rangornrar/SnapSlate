# Validation du lot 3 — 28 octobre 2026

Version distribuée : `2026.10.28.001` (version des binaires `2026.10.28.1`), Windows x64. Périmètre : F13–F17 et F19–F23 de l’audit initial, en complément des lots 1 et 2.

## Changements et preuves

| Constat | Résultat |
| --- | --- |
| F13 | Plus de timer de lecture/décodage du presse-papiers. Révisions d’événements coalescées et lectures obsolètes ignorées. Comparaisons de contenu sans JSON/Base64, empreintes d’images immuables partagées ; export PNG sur instantané en tâche de fond. |
| F14 | Insertion dans l’historique avant ajout, copie identique acceptée à chaque nouvel événement et import répété explicite via Fichier ou Ctrl+Shift+V ; origine « Presse-papiers ». |
| F15 | Publications annoncées comme texte seul, sans construire ni décoder les captures. Titre et Markdown Notion ; résolution d’espace et création de page Confluence Cloud v2. Messages distincts pour 401/403/404/429, sans corps de réponse dans les erreurs. |
| F16 | Suite des services de production et parcours import → annotation → sauvegarde atomique → réouverture → PNG/HTML/Markdown/PDF/DOCX. CI Windows avec construction du setup et contrôle installation/lancement/désinstallation avant conservation des artefacts. |
| F17 | SDK et dépendances verrouillés ; publish Release autonome, version de setup dérivée du projet, compilateur Inno téléchargé et vérifié par SHA-256 en CI, empreintes et provenance. Signature optionnelle préparée, release non signée faute de certificat. |
| F19 | Liste des objets, commandes de déplacement de 1 px, conservation du focus lors d’un déplacement, noms d’étapes/actions et associations des champs aux libellés, ressources de contraste système. |
| F20 | Ressources FR/EN communes, culture explicite des exports ; titres DOCX avec niveaux hiérarchiques et descriptions des images, langue DOCX/PDF/HTML. |
| F21 | Suppression des trois scripts historiques MSIX, aucun ajout de certificat dans Root/TrustedPeople. |
| F22 | Services de coordination, empreintes, texte de publication et localisation extraits et testables. README, architecture, qualité et chemin de distribution alignés. MainWindow conserve la coordination de l’interface. |
| F23 | Noms de fichiers sans chemin local complet dans les guides et publications ; tests Windows/UNC/Linux et exports des deux langues. |

## Tests automatisés

83 tests de non-régression passent en Release x64. Ils réutilisent les sources de production. Le parcours d’export vérifie masquage opaque, absence du chemin privé, textes et langues FR/EN, styles DOCX, descriptions des images, validation Open XML Office 2019 et métadonnées PDF. Les contrats HTTP sont testés avec un transport factice ; aucun compte Notion ou Confluence réel n’a été modifié.

```powershell
dotnet test .\tests\SnapSlate.Tests.csproj -c Release -p:Platform=x64 --filter "TestCategory!=Performance"
```

Build Debug x64 sans avertissement ni erreur. Le workflow Release exécute les tests puis construit et vérifie la distribution. Le script d’installation de test est limité aux runners jetables GitHub Actions ; il n’installe rien sur le poste utilisateur.

Le contrôle de lancement a détecté un PRI WinUI réutilisé après le build MSIX, dépourvu des ressources de thème nécessaires à la distribution autonome. Le packaging effectue maintenant un nettoyage avec les mêmes propriétés que le publish. Une publication Release propre démarre localement ; la CI exerce volontairement build MSIX → packaging autonome → lancement installé pour prévenir ce défaut.

## Mesure ciblée du contrôle d’état modifié

Même machine, même PNG de 1920 × 1032 (103 748 octets), tableaux distincts pour chaque étape ; cinq boucles `HasChanges` + `RefreshIndicators`, après mise en cache initiale. Mesures uniques indicatives, pas un benchmark statistique ni une mesure de la mémoire totale de WinUI.

| Étapes | Avant : durée / allocations | Après : durée / allocations |
| --- | --- | --- |
| 1 | 2,8733 ms / 4 190 328 octets | 0,5847 ms / 10 856 octets |
| 10 | 20,6743 ms / 41 922 904 octets | 0,2649 ms / 83 400 octets |
| 50 | 99,4483 ms / 209 508 656 octets | 1,2218 ms / 402 360 octets |

Le hash initial de chaque image reste proportionnel à sa taille ; les tableaux sont ensuite immuables et partagés par l’historique. Ce protocole répète la même image et ne mesure pas 50 captures 4K différentes ni la latence complète d’interaction. Reproduction avec `ProjectPerformanceTests` : définir `SNAPSLATE_PERF_IMAGE` et `SNAPSLATE_PERF_REPORT`, puis filtrer `TestCategory=Performance`.

## Contrôles Windows et limites

Sur le build Debug du lot 3 : lancement, fenêtre agrandie, sélection d’un rectangle par la liste, Tab vers le bouton de déplacement, activation par Espace puis annulation Ctrl+Z (message « Dernière action annulée »). L’arbre UIA expose le nom des étapes, actions et champs. Aperçu des réglages en anglais, puis abandon par Échap et retour au français vérifiés.

Narrator, toutes les variantes de contraste élevé, les autres architectures, la matrice Windows 10/11 et l’installation interactive complète restent à tester. Les ressources contrastées et la structure DOCX ne constituent pas une certification d’accessibilité. L’observation web de la lightbox est un travail distinct ; les modifications locales du site sont conservées.

Aucun certificat de production n’étant disponible, la signature n’a pas été exécutée. La release indique explicitement cette limite ; la chaîne signée devra être vérifiée lorsqu’elle sera activée. Aucun engagement de reproductibilité binaire bit à bit n’est pris.

Les projets et sauvegardes `.bak` gardent les sources originales. Les titres, notes et légendes restent du contenu utilisateur : une information privée saisie dans ces champs n’est pas supprimée automatiquement à l’export.
