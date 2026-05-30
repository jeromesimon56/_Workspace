param(
    [string]$Output = "$PSScriptRoot/../../outlook_emails",
    [int]$MaxMessages = 0
)

. "$PSScriptRoot/graph_auth.ps1"

Write-Host "Extraction des pièces jointes depuis TOUS les dossiers Outlook..."

# Récupération de l'ID utilisateur
$userId = (Get-MgUser -UserId me).Id

# Fonction récursive pour parcourir tous les dossiers
function Get-AllFolders {
    param($UserId, $ParentFolderId)

    $folders = Get-MgUserMailFolder -UserId $UserId -MailFolderId $ParentFolderId -ErrorAction SilentlyContinue
    foreach ($folder in $folders) {
        $folder
        foreach ($sub in Get-AllFolders -UserId $UserId -ParentFolderId $folder.Id) {
            $sub
        }
    }
}

# Récupération du dossier racine
$rootFolders = Get-MgUserMailFolder -UserId $userId

# Tous les dossiers (racine + sous-dossiers)
$allFolders = @()
foreach ($f in $rootFolders) {
    $allFolders += $f
    $allFolders += Get-AllFolders -UserId $userId -ParentFolderId $f.Id
}

Write-Host "Dossiers trouvés :"
$allFolders.DisplayName | ForEach-Object { Write-Host " - $_" }

foreach ($folder in $allFolders) {

    Write-Host ""
    Write-Host "Traitement du dossier : $($folder.DisplayName)"

    # Récupération des messages
    $params = @{
        UserId       = $userId
        MailFolderId = $folder.Id
        All          = $true
    }
    if ($MaxMessages -gt 0) { $params.Top = $MaxMessages }

    $messages = Get-MgUserMailFolderMessage @params

    foreach ($msg in $messages) {

        # Création du dossier de sortie par message
        $safeSubject = ($msg.Subject -replace '[^a-zA-Z0-9-_ ]','_')
        if (-not $safeSubject) { $safeSubject = "SansSujet" }

        $msgFolder = Join-Path $Output $folder.DisplayName
        $msgFolder = Join-Path $msgFolder ("{0:yyyy-MM-dd}_{1}" -f $msg.ReceivedDateTime, $msg.Id)

        if (-not (Test-Path $msgFolder)) {
            New-Item -ItemType Directory -Path $msgFolder | Out-Null
        }

        # Récupération des pièces jointes
        $attachments = Get-MgUserMessageAttachment -UserId $userId -MessageId $msg.Id

        foreach ($att in $attachments) {

            # 🔥 Filtrer les pièces jointes valides
            if ($att.AdditionalProperties.'@odata.type' -ne "#microsoft.graph.fileAttachment") {
                continue
            }
            if (-not $att.ContentBytes) {
                Write-Host "Pièce jointe ignorée (pas de contenu) : $($att.Name)"
                continue
            }

            $fileName = $att.Name
            $filePath = Join-Path $msgFolder $fileName

            [IO.File]::WriteAllBytes($filePath, $att.ContentBytes)
            Write-Host "Pièce jointe sauvegardée : $filePath"
        }
    }
}

Write-Host ""
Write-Host "Extraction terminée pour tous les dossiers."
