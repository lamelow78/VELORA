# Site web Velora Finance

Ce dossier contient une landing page statique prete pour Netlify.

## Structure

- `index.html` : page principale
- `styles.css` : design et responsive
- `app.js` : detection automatique du systeme d'exploitation
- `telecharger/` : pages de telechargement par systeme
- `download.js` : verification du manifeste et redirection automatique
- `downloads/` : dossier ou placer les installateurs

## Fichiers de telechargement attendus

Place les installateurs dans `website/downloads/` avec ces noms :

- `velora-finance-windows.exe`
- `velora-finance-macos.zip`
- `velora-finance-linux.tar.gz`

## Liens exposes sur le site

- `/telecharger/windows`
- `/telecharger/macos`
- `/telecharger/linux`

Ces pages verifient `website/downloads/manifest.json` puis lancent le telechargement si le binaire est disponible.

## Netlify

Le fichier racine `netlify.toml` publie automatiquement le dossier `website`.

Tu peux donc connecter le depot a Netlify tel quel.

## Publication automatique

Le workflow GitHub `../.github/workflows/build-downloads.yml` genere les binaires Windows, macOS et Linux puis met a jour `website/downloads/`.
