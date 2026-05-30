param(
    [string]$Username,
    [string]$Password,
    [string]$Output = "$PSScriptRoot/attachments"
)

# Charger MailKit et MimeKit
Add-Type -Path "$PSScriptRoot/lib/MimeKit.dll"
Add-Type -Path "$PSScriptRoot/lib/MailKit.dll"

# Créer le dossier de sortie
if (-not (Test-Path $Output)) {
    New-Item -ItemType Directory -Path $Output | Out-Null
}

$client = [MailKit.Net.Imap.ImapClient]::new()

try {
    Write-Host "Connexion IMAP..."
    $client.Connect("imap-mail.outlook.com", 993, $true)

    Write-Host "Authentification..."
    $client.Authenticate($Username, $Password)

    Write-Host "Récupération des dossiers..."
    $root = $client.GetFolder($client.PersonalNamespaces[0])

    function Get-FoldersRecursively {
        param([MailKit.IMailFolder]$Folder)

        foreach ($sub in $Folder.GetSubfolders($true)) {
            $sub
            Get-FoldersRecursively -Folder $sub
        }
    }

    $folders = Get-FoldersRecursively -Folder $root

    foreach ($folder in $folders) {
        Write-Host "Dossier : $($folder.FullName)"
        $folder.Open([MailKit.FolderAccess]::ReadOnly)

        $count = $folder.Count
        if ($count -eq 0) {
            $folder.Close()
            continue
        }

        $summaries = $folder.Fetch(0, $count - 1, [MailKit.MessageSummaryItems]::UniqueId)

        foreach ($summary in $summaries) {
            $msg = $folder.GetMessage($summary.UniqueId)

            if (-not $msg.Attachments) { continue }

            $safeFolder = ($folder.FullName -replace '[\\/:*?"<>|]', '_')
            $msgFolder = Join-Path $Output $safeFolder

            if (-not (Test-Path $msgFolder)) {
                New-Item -ItemType Directory -Path $msgFolder | Out-Null
            }

            foreach ($att in $msg.Attachments) {
                $fileName = $att.FileName
                if (-not $fileName) { $fileName = "attachment.bin" }

                $filePath = Join-Path $msgFolder $fileName

                $stream = [System.IO.File]::Create($filePath)
                try {
                    if ($att -is [MimeKit.MessagePart]) {
                        $att.Message.WriteTo($stream)
                    } else {
                        $part = [MimeKit.MimePart]$att
                        $part.Content.DecodeTo($stream)
                    }
                } finally {
                    $stream.Dispose()
                }

                Write-Host "Pièce jointe sauvegardée : $filePath"
            }
        }

        $folder.Close()
    }

} finally {
    if ($client.IsConnected) {
        $client.Disconnect($true)
    }
}
