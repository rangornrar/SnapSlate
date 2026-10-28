# Espace d’annotation — 28 octobre 2026

Priorité confirmée : annoter une capture, avec le canevas maximal et les étapes masquées par défaut.

## Constat et décision

L’écran précédent ouvrait les étapes (256 px) et l’inspecteur (312 px), cumulait deux lignes de commandes et des cadres espacés, et plaçait le rail sur le viewport. La capture de démonstration était coupée à droite à 100 %.

Le nouvel écran réserve l’espace à l’image. Les étapes et les propriétés sont fermées au démarrage ; le rail fixe reste hors de la capture. Une ligne réunit menus, import, sauvegarde, annuler/rétablir et export PNG. Les réglages de l’outil ou de la sélection apparaissent dans une ligne contextuelle ; les propriétés détaillées s’ouvrent pour le texte et les gommettes. Le zoom et les commandes de panneaux restent en bas.

Compromis : les détails des étapes demandent un clic. La capture active reste accessible dans la liste du bas. Le choix privilégie l’annotation, conformément à la demande ; un usage majoritairement consacré à la rédaction de guides justifierait de revoir ce défaut.

L’ajustement automatique montre toute l’image sans l’agrandir au-delà de 100 %. Le zoom manuel désactive cet ajustement jusqu’à « Ajuster » ou au changement de capture. Les positions historiques du rail et l’ancien réglage de réduction des étapes restent lisibles dans les préférences, mais ne commandent plus cette disposition.

## Captures natives

Captures PNG du même écran Windows, sans recomposition :

- [Avant, fenêtre maximisée](ux-2026-10-28/before.png)
- [Après, fenêtre initiale 1440 × 900](ux-2026-10-28/after-default.png)
- [Après, fenêtre maximisée](ux-2026-10-28/after-maximized.png)
- [Capture 4K complète, avec rectangle sélectionné](ux-2026-10-28/after-4k.png)

## Validation

- Build Debug x64 : aucune erreur ni avertissement.
- 83 tests automatisés réussis ; le profil de performance optionnel reste ignoré.
- Le pourcentage courant est présent dans le nom accessible du bouton de zoom. Aperçu anglais des réglages, abandon par Échap et retour au français avec le zoom conservé.
- Création d’un texte puis saisie directe, sans recliquer le champ ; texte et sélection visibles sur le canevas.
- Ouverture/fermeture indépendante des étapes et propriétés, image entière visible aux deux tailles ci-dessus.
- Import d’une image synthétique 3840 × 2160 ; création d’un rectangle, annulation (1 → 0) et rétablissement (0 → 1).
- Enregistrement du projet, réouverture avec les deux captures et le rectangle conservés.
- Export PNG 3840 × 2160 ; la différence de pixels avec la source reste limitée au rectangle. Le contour de sélection de l’éditeur n’est pas exporté.

Les essais natifs portent sur ce poste Windows et ces tailles de fenêtre. Les écrans tactiles, les autres échelles DPI, Narrator et le mode de contraste élevé de Windows ne sont pas validés ici. Cette refonte est distribuée dans `2026.10.28.002` (binaires `2026.10.28.2`), après la release du lot 3 `2026.10.28.001`.
