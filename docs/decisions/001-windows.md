# 001 — Application Windows avec WinUI 3

**Statut :** architecture existante documentée.

Le besoin comprend l’import des captures Windows et l’édition locale. WinUI 3 et le Windows App SDK permettent une intégration native avec le presse-papiers, la fenêtre et les sélecteurs de fichiers.

Une application web ou un framework multiplateforme seraient des alternatives crédibles pour élargir la distribution. Ils impliqueraient une autre intégration des captures et une migration de l’interface. Le projet conserve son socle Windows tant qu’aucun besoin multiplateforme concret ne justifie ce coût.

Conséquences : compilation Windows en CI, dépendances au SDK Windows et validation manuelle du parcours d’installation. Cette décision ne démontre aucune supériorité générale de WinUI sur les alternatives.
