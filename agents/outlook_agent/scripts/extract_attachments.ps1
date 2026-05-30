param(
    [string]$Folder = "Inbox",
    [string]$Output = "$PSScriptRoot/../../outlook_emails",
    [int]$MaxMessages = 0
)

. "$PSScriptRoot/graph_auth.ps1"

Write-Host "Extraction des pièces jointes depuis le dossier Outlook : $Folder"

$userId = (Get-MgUser -UserId me).Id

# Récupération du dossier Outlook
$folderObj = Get-MgUserMailFolder -UserId $userId -MailFolderId $Folder -ErrorAction SilentlyContinue
if (-not $folderObj) {
    Write-Host "Dossier introuvable : $Folder"
    exit
}

# Récupération des messages
$params = @{
    UserId       = $userId
    MailFolderId = $folderObj.Id
    All          = $true
}
if ($MaxMessages -gt 0) { $params.Top = $MaxMessages }

$messages = Get-MgUserMailFolderMessage @params

foreach ($msg in $messages) {

    # Création du chemin miroir : outlook_emails/<Folder>/<MessageId>/
    $safeSubject = ($msg.Subject -replace '[^a-zA-Z0-9-_ ]','_')
    if (-not $safeSubject) { $safeSubject = "SansSujet" }

    $msgFolder = Join-Path $Output $Folder
    $msgFolder = Join-Path $msgFolder ("{0:yyyy-MM-dd}_{1}" -f $msg.ReceivedDateTime, $msg.Id)

    if (-not (Test-Path $msgFolder)) {
        New-Item -ItemType Directory -Path $msgFolder | Out-Null
    }

    # Récupération des pièces jointes
    $attachments = Get-MgUserMessageAttachment -UserId $userId -MessageId $msg.Id

    foreach ($att in $attachments) {
        if ($att.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.fileAttachment") {

            $fileName = $att.Name
            $filePath = Join-Path $msgFolder $fileName

            [IO.File]::WriteAllBytes($filePath, $att.ContentBytes)
            Write-Host "Pièce jointe sauvegardée : $filePath"
        }
    }
}

Write-Host "Extraction terminee."
