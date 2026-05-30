# Outlook Agent — README

But: récupérer, trier et extraire les emails Outlook via Microsoft Graph.

Prerequis
- PowerShell 7+ (ou Windows PowerShell avec les modules compatibles)
- Module `Microsoft.Graph` (le script `install_requirements.ps1` l'installe)

Authentification
- Le flux interactif (par défaut) : `Connect-MgGraph -Scopes "Mail.Read","Mail.ReadWrite","User.Read"`.
- Authentification non interactive (app-only) : définir les variables d'environnement suivantes si vous voulez appeler `Connect-MgGraph` en mode application :
  - `AZURE_CLIENT_ID` : ID de l'application (client)
  - `AZURE_TENANT_ID` : ID du tenant
  - `AZURE_CLIENT_SECRET` : secret client (ou utilisez un certificat)

Exemple app-only (PowerShell) :

```powershell
$env:AZURE_CLIENT_ID = "<client-id>"
$env:AZURE_TENANT_ID = "<tenant-id>"
$env:AZURE_CLIENT_SECRET = "<client-secret>"
Connect-MgGraph -ClientId $env:AZURE_CLIENT_ID -TenantId $env:AZURE_TENANT_ID -ClientSecret $env:AZURE_CLIENT_SECRET -Scopes "https://graph.microsoft.com/.default"
```

Installation rapide
- Exécuter le script d'installation (installe `Microsoft.Graph` si nécessaire) :

```powershell
.\agents\outlook_agent\install_requirements.ps1
```

Usage
- Récupérer les emails :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action fetch -Folder Inbox -EmailsOutput .\outlook_emails -MaxMessages 20
```

> Le paramètre `-Folder` accepte les noms de dossiers Outlook standard (ex. `Inbox`, `Sent Items`, `Drafts`, `Deleted Items`, `Junk Email`). Si votre boîte est localisée, ces noms standard fonctionnent aussi.
>
> `-MaxMessages` limite le nombre de messages récupérés pour des tests rapides.

- Extraire les pièces jointes :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action extract -Folder Inbox -AttachmentsOutput .\attachments
```

- Supprimer les emails promotionnels détectés (avec confirmation) :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action clean -Folder Inbox -MaxMessages 500
```

- Nettoyer tous les dossiers Outlook :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action clean -Folder All -MaxMessages 500
```

- Forcer la suppression uniquement pour certains domaines :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action clean -Folder All -MaxMessages 500 -ForceDeleteDomains 'amazon.fr','example.com' -DryRun
```

- Nettoyage standard (LinkedIn, Amazon, et domaines définis par défaut) :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action clean -Folder All -MaxMessages 500 -StandardCleanup -DryRun
```

> Quand `-StandardCleanup` est spécifié, le script recherche d'abord les emails déjà importés localement dans `./outlook_emails` pour les domaines standards, puis supprime uniquement les messages serveur qui correspondent à ces imports locaux. `-ForceDeleteDomains` étend la liste des domaines recherchés dans les fichiers locaux.

- Lister d'abord les messages détectés sans les supprimer (DryRun) :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action clean -Folder Inbox -MaxMessages 500 -DryRun
```

Note : certains expéditeurs transactionnels comme `mondialrelay` et `gerep` sont exclus tant qu'ils ne sont pas suffisamment anciens. Si ces messages datent de plus de 30 jours, ils peuvent être considérés comme nettoyables.

Pour exécuter sans invite, ajoutez `-AutoConfirm` (à utiliser avec prudence).

- Trier les emails (après `fetch`) :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action sort -EmailsOutput .\outlook_emails -Rules .\agents\outlook_agent\config\settings.json
```

- Tout exécuter :

```powershell
.\agents\outlook_agent\run_agent.ps1 -Action all
```

Notes
- Les scripts actuels utilisent `Connect-MgGraph` (authentification interactive). Si vous souhaitez des opérations automatisées côté serveur, enregistrez une application Azure AD et utilisez l'authentification app-only (voir variables d'environnement ci-dessus).
- Si `Select-MgProfile` n'est pas disponible dans votre version du module, l'authentification utilisera le profil Graph par défaut.
- Les actions manipulent des fichiers `.eml` et dossiers locaux; vérifiez les permissions du dossier de sortie.
