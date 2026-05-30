param(
    [string]$Username,
    [string]$Password,
    [string]$Output = "$PSScriptRoot/attachments"
)

# Charger MailKit localement
Add-Type -Path "$PSScriptRoot/lib/MimeKit.dll"
Add-Type -Path "$PSScriptRoot/lib/MailKit.dll"

$client = [MailKit.Net.Imap.ImapClient]::new()

try {
    Write-Host "🔌 Connexion IMAP..."
    $client.Connect("imap-mail.outlook.com", 993, $true)

    Write-Host "👤 Authentification..."
    $client.Authenticate($Username, $Password)

    Write-Host "Récupération des dossiers..."
    $root = $client.GetFolder($client.PersonalNamespaces[0])

    function Get-FoldersRecursively {
        param([MailKit.IMailFolder]$Folder)

        foreach ($sub in $Folder.GetSubfolders($true)) {
            $sub
            Get-FoldersRecursively $sub
        }
    }

    $folders = Get-FoldersRecursively $root

    foreach ($folder in $folders) {
        Write-Host "Dossier : $($folder.FullName)"
        $folder.Open([MailKit.FolderAccess]::ReadOnly)

        foreach ($msg in $folder.Fetch(0, $folder.Count - 1, [MailKit.MessageSummaryItems]::Full | [MailKit.MessageSummaryItems]::UniqueId)) {

            $message = $folder.GetMessage($msg.UniqueId)

            if (-not $message.Attachments) { continue }

            $safeFolder = ($folder.FullName -replace '[\\/:*?"<>|]', '_')
            $msgFolder = Join-Path $Output $safeFolder
            if (-not (Test-Path $msgFolder)) { New-Item -ItemType Directory -Path $msgFolder | Out-Null }

            foreach ($att in $message.Attachments) {
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

                Write-Host "📎 $filePath"
            }
        }

        $folder.Close()
    }

} finally {
    if ($client.IsConnected) { $client.Disconnect($true) }
}
