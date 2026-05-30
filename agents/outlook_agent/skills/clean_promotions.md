# Skill: clean_promotions

## Description
Recherche et supprime les emails promotionnels (newsletter/offres) dans un dossier Outlook, avec confirmation.

## Inputs
- Folder name (ex: Inbox, All)
- MaxMessages (0 = tous les messages)
- StandardCleanup (switch) : active le nettoyage des domaines par défaut (LinkedIn, Amazon)
- ForceDeleteDomains (array) : si spécifié, seuls les emails de ces domaines sont supprimés
- DryRun (switch) : affiche les messages détectés sans les supprimer
- AutoConfirm (bool) : supprime sans demander confirmation

## Output
- Supprime les messages identifiés sur la boîte (Microsoft Graph) et tente de supprimer les copies locales `.eml` lors de l'exécution.

## Script
scripts/clean_promotions.ps1
