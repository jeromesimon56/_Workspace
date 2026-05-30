param(
    [string]$Folder = "Inbox",
    [string]$Output = "./attachments",
    [int]$MaxMessages = 0
)

. "$PSScriptRoot\graph_auth.ps1"

$userId = (Get-MgUser -UserId me).Id
$knownFolderIds = @{
    'Inbox' = 'Inbox'
    'Sent Items' = 'SentItems'
    'SentItems' = 'SentItems'
    'Drafts' = 'Drafts'
    'Deleted Items' = 'DeletedItems'
    'DeletedItems' = 'DeletedItems'
    'Junk Email' = 'JunkEmail'
    'JunkEmail' = 'JunkEmail'
    'Archive' = 'Archive'
    'Outbox' = 'Outbox'
}
$folderId = if ($knownFolderIds.ContainsKey($Folder)) { $knownFolderIds[$Folder] } else { $null }

if ($folderId) {
    $folderObj = Get-MgUserMailFolder -UserId $userId -MailFolderId $folderId
} else {
    $folderObj = Get-MgUserMailFolder -UserId $userId -Top 100 | Where-Object { $_.DisplayName -ieq $Folder } | Select-Object -First 1
}
if (-not $folderObj) {
    Write-Error "Mail folder '$Folder' introuvable. Vérifiez le nom de dossier Outlook."
    exit 1
}

if (-not (Test-Path $Output)) {
    New-Item -ItemType Directory -Path $Output -Force | Out-Null
}

if ($MaxMessages -gt 0) {
    $messages = Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folderObj.Id -Top $MaxMessages
} else {
    $messages = Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folderObj.Id -All
}

foreach ($msg in $messages) {
    $id = $msg.Id
    $attachments = Get-MgUserMessageAttachment -UserId $userId -MessageId $id -All
    foreach ($att in $attachments) {
        if ($att.ODataType -eq "#microsoft.graph.fileAttachment") {
            $name = $att.Name
            $contentBytes = $att.ContentBytes
            $bytes = [System.Convert]::FromBase64String($contentBytes)
            $dir = Join-Path $Output $id
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
            $path = Join-Path $dir $name
            [System.IO.File]::WriteAllBytes($path, $bytes)
        }
    }
}
