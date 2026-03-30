Place ici les fichiers de telechargement qui seront servis directement par Netlify:

- `velora-finance-windows.exe`
- `velora-finance-macos.zip`
- `velora-finance-linux.tar.gz`

Les boutons du site pointent vers:

- `/telecharger/windows`
- `/telecharger/macos`
- `/telecharger/linux`

Le manifeste `website/downloads/manifest.json` indique si le binaire est deja publie.

Si tu veux changer les noms de fichiers plus tard, modifie `website/downloads/manifest.json` et le workflow de build.
