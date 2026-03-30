# Velora Finance

Velora Finance est une application desktop locale pour piloter les finances d'une entreprise:

- tableau de bord avec chiffre d'affaires, benefice, depenses et suivi des documents
- calendrier sante pour visualiser jour par jour recettes, depenses, echeances et taches
- gestion des recettes et des depenses avec montants HT et TTC
- liaison automatique des factures au chiffre d'affaires
- liaison possible des enregistrements manuels a une facture existante
- generateur de factures
- generateur de devis avec duree de validite modifiable
- todo liste avec date, heure, modification et suppression
- stockage local sur la machine dans `%LOCALAPPDATA%\Velora Finance`

## Lancer l'application

```powershell
py main.py
```

Ou bien double-cliquez sur `run_velora.bat`.

## Donnees locales

Les donnees sont enregistrees dans:

- base SQLite: `%LOCALAPPDATA%\Velora Finance\velora_finance.db`
- factures HTML: `%LOCALAPPDATA%\Velora Finance\documents\factures`
- devis HTML: `%LOCALAPPDATA%\Velora Finance\documents\devis`

## Fonctions principales

- Parametres entreprise: nom, raison sociale, SIRET, TVA, coordonnees, adresse, logo
- Dashboard avec graphiques simples
- Calendrier mensuel pour surveiller la sante financiere
- Recettes et depenses avec source manuelle ou facture
- Dates d'ajout et de creation visibles dans les listes
- Suivi des statuts de factures et devis
- Todo liste datee avec heures et statuts
- Ouverture directe des documents HTML generes
