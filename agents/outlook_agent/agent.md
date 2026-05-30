# Outlook Agent

## Description
Agent chargé de récupérer, trier et archiver les emails Outlook via Microsoft Graph.

## Skills
- fetch_outlook_emails
- sort_outlook_emails
- extract_attachments
- clean_promotions

## Scripts
- scripts/outlook_fetch.ps1
- scripts/outlook_utils.ps1
- scripts/graph_auth.ps1
- scripts/extract_attachments.ps1
 - scripts/clean_promotions.ps1

## Config
- config/settings.json

## Usage

Lancer les actions principales depuis PowerShell :

- Récupérer les emails :
	- ` .\run_agent.ps1 -Action fetch -Folder Inbox -EmailsOutput ./outlook_emails -MaxMessages 20`
- Extraire les pièces jointes :
	- ` .\run_agent.ps1 -Action extract -Folder Inbox -AttachmentsOutput ./attachments -MaxMessages 20`
- Trier les emails (après fetch) :
	- ` .\run_agent.ps1 -Action sort -EmailsOutput ./outlook_emails -Rules ./config/settings.json`
- Tout faire :
	- ` .\run_agent.ps1 -Action all -MaxMessages 20`
