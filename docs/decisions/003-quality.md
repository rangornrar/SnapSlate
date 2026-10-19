# 003 — Services isolés et compilation cohérente

**Statut :** adopté pour la fiabilisation du dépôt.

Le dernier échec historique examiné signalait une erreur ReadyToRun après une restauration sans les mêmes propriétés que la compilation Release. La CI utilise désormais la même configuration, la même plateforme et le même runtime pour les deux étapes, avec ReadyToRun désactivé comme dans le script de staging. Une publication optimisée peut rester un parcours distinct.

Les services `ProjectJson` et `GuideTextExporter` sont indépendants des préférences de l’application et du démarrage WinUI. Les tests MSTest compilent les fichiers de production en lien, sans copie de leur implémentation. Ils vérifient les formats et leurs cas limites ; la compilation de l’application vérifie séparément leur intégration au projet.

Compromis : ce premier niveau ne remplace ni les tests de bout en bout ni les essais d’installation. Si les services sont réutilisés ailleurs, les déplacer dans une bibliothèque sera préférable au partage de fichiers sources.

Référence : [ReadyToRun — documentation Microsoft](https://learn.microsoft.com/en-us/dotnet/core/deploying/ready-to-run).
