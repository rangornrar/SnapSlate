# Décisions de conception

Ces notes explicitent les choix observables dans le code et les compromis retenus lors de leur documentation. Elles ne prétendent pas reconstituer toutes les intentions historiques du projet.

| Décision | Motif | Compromis / condition de révision |
| --- | --- | --- |
| [001 — Application Windows](001-windows.md) | Intégration locale avec captures et bureau | Revoir si un besoin multiplateforme devient prioritaire |
| [002 — Projet JSON](002-json.md) | Sauvegarde lisible et état éditable | Revoir si volume des captures ou migrations de schéma le justifient |
| [003 — Services testables et CI](003-quality.md) | Vérifier formats et compilation sans coupler les tests au démarrage de l’interface | Extraire une bibliothèque si les consommateurs se multiplient |
