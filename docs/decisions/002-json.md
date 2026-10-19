# 002 — Persistance du projet en JSON

**Statut :** format existant conservé ; convertisseurs géométriques isolés et testés.

Le format `.snapslate` représente l’état du projet via `SnapSlateProjectState`. `System.Text.Json` sérialise l’état, avec des convertisseurs dédiés à `Point` et `Rect`.

L’alternative serait un conteneur séparant manifeste et images. Elle pourrait réduire le coût de certaines opérations sur de gros projets, au prix d’une gestion supplémentaire du packaging et des migrations. Le JSON actuel reste adapté tant que des mesures ne montrent pas une limite bloquante.

Les propriétés géométriques inconnues sont ignorées, y compris les valeurs imbriquées. Les types de valeurs invalides provoquent une erreur de désérialisation. La compatibilité de toutes les versions futures n’est pas garantie : la version du schéma devra faire l’objet de migrations explicites si le modèle change.
