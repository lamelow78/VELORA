Ce dossier contient le manifeste public consomme par les pages de telechargement du site.

Les binaires desktop ne sont plus commits dans le depot. Ils sont publies par le workflow GitHub dans la release `downloads-latest`, puis le manifeste est mis a jour ici avec:

- la disponibilite par plateforme
- le nom du fichier
- l'URL de telechargement
- la taille et le hash SHA-256

Les boutons du site pointent vers:

- `/telecharger/windows`
- `/telecharger/macos`
- `/telecharger/linux`

Le manifeste `website/downloads/manifest.json` indique si le binaire est deja publie.

Si tu veux changer les noms de fichiers plus tard, modifie le workflow de build et laisse le manifeste etre regenere automatiquement.
