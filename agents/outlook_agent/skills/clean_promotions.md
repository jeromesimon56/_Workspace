# Skill: clean_promotions

## Description
Recherche et supprime les emails promotionnels (newsletter/offres) dans un dossier Outlook, avec confirmation.

## Inputs
- Folder name (ex: Inbox, All)
- MaxMessages (0 = tous les messages)
- StandardCleanup (switch) : active le nettoyage standard en recherchant d'abord les emails importés localement dans `./outlook_emails`, puis supprime uniquement les messages serveur correspondant à ces imports.
- ForceDeleteDomains (array) : étend la liste des domaines recherchés dans les emails locaux lors du nettoyage standard, sinon limite la suppression aux domaines fournis.
- DryRun (switch) : affiche les messages détectés sans les supprimer
- AutoConfirm (bool) : supprime sans demander confirmation

## Output
- Supprime les messages identifiés sur la boîte (Microsoft Graph) et tente de supprimer les copies locales `.eml` lors de l'exécution.

## Script
scripts/clean_promotions.ps1
