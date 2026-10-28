# SnapSlate v2026.10.28.002

Une interface centrée sur l’annotation : plus de place pour la capture, moins de commandes permanentes.

- Étapes et propriétés masquées au démarrage ; outils fixes à gauche, hors de l’image.
- Import, sauvegarde, annuler/rétablir et export PNG directement accessibles.
- Réglages contextuels ; propriétés détaillées ouvertes pour le texte et les gommettes.
- Capture ajustée automatiquement au canevas, pourcentage de zoom exact et accessible.
- Nouvelle annotation sélectionnée et saisie directe des textes ; sélection visible sur fonds clairs et sombres.
- Annulation des formes dessinées corrigée.

## Télécharger

- `SnapSlate-Setup.exe` : installateur Windows x64 autonome, .NET et Windows App SDK inclus.
- `SnapSlate-Installer-x64.zip` : distribution portable à extraire entièrement ; lancer `app/SnapSlate.exe`.
- `SHA256SUMS.txt` : empreintes des deux distributions.
- `build-provenance.json` : révision construite, SDK, état de signature et empreintes.

## Validation et limites

83 tests automatisés. La CI compile la distribution, puis vérifie installation silencieuse, lancement et désinstallation sur un runner Windows jetable. Contrôles natifs : panneaux à la demande, saisie de texte, annulation/rétablissement, sauvegarde/réouverture et export PNG 3840 × 2160. Captures avant/après et périmètre dans `docs/UX_WORKSPACE.md`.

**Release non signée Authenticode : aucun certificat de production n’est disponible.**

Les essais UX portent sur le poste Windows de développement et les tailles initiale/maximisée. Les écrans tactiles, les autres échelles DPI, Narrator et le mode de contraste élevé ne sont pas validés ici. Les projets éditables et `.bak` conservent les sources originales ; partager les exports pour diffuser les masquages.
