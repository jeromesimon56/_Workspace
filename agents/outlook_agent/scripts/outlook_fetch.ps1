param(
    [string]$Folder = "Inbox",
    [string]$Output = "./outlook_emails",
    [int]$MaxMessages = 0
)

. "$PSScriptRoot\graph_auth.ps1"

function Get-SafeFileNamePart {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $safe = $Value -replace '[^\w\s-]', ''
    return $safe.Trim()
}

function Get-MailFolderPath {
    param(
        [Parameter(Mandatory=$true)]$FolderObj,
        [Parameter(Mandatory=$true)]$UserId
    )

    $segments = @()
    $current = $FolderObj
    $visited = @{}

    while ($current) {
        if ($current.DisplayName) {
            $segments += Get-SafeFileNamePart $current.DisplayName
        }

        if (-not $current.ParentFolderId) {
            break
        }

        if ($visited.ContainsKey($current.ParentFolderId)) {
            break
        }

        $visited[$current.ParentFolderId] = $true
        $parent = Get-MgUserMailFolder -UserId $UserId -MailFolderId $current.ParentFolderId -Property 'DisplayName,ParentFolderId'
        if (-not $parent) {
            break
        }
        $current = $parent
    }

    $pathSegments = [System.Collections.ArrayList]$segments
    [void][System.Array]::Reverse($pathSegments)
    $path = $pathSegments -join '\'
    return $path
}

Write-Host "Récupération de l'utilisateur connecté..."
$user = Get-MgUser -UserId me
$userId = $user.Id
Write-Host "Utilisateur connecté : $($user.DisplayName) <$($user.Mail)> ($userId)"
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
    Write-Host "Chargement du dossier standard '$Folder' (ID=$folderId)..."
    try {
        $folderObj = Get-MgUserMailFolder -UserId $userId -MailFolderId $folderId -Verbose
        Write-Host "Dossier standard chargé : $($folderObj.DisplayName)"
    } catch {
        Write-Host "Erreur Get-MgUserMailFolder (standard) : $($_.Exception.Message)"
        throw
    }
} else {
    Write-Host "Recherche du dossier personnalisé '$Folder'..."
    try {
        $folderObj = Get-MgUserMailFolder -UserId $userId -Top 100 -Verbose | Where-Object { $_.DisplayName -ieq $Folder } | Select-Object -First 1
        Write-Host "Dossier personnalisé chargé : $($folderObj.DisplayName)"
    } catch {
        Write-Host "Erreur Get-MgUserMailFolder (custom) : $($_.Exception.Message)"
        throw
    }
}
if (-not $folderObj) {
    Write-Error "Mail folder '$Folder' introuvable. Vérifiez le nom de dossier Outlook. Si besoin, utilisez un nom exact comme 'Inbox', 'Sent Items' ou 'Drafts'."
    exit 1
}

if (-not (Test-Path $Output)) {
    New-Item -ItemType Directory -Path $Output -Force | Out-Null
}

Write-Host "Résolution du chemin local pour le dossier Outlook..."
$folderPath = Get-MailFolderPath -FolderObj $folderObj -UserId $userId
Write-Host "Chemin local résolu : $folderPath"
$Output = Join-Path $Output $folderPath
Write-Host "Chemin complet de sortie : $Output"
if (-not (Test-Path $Output)) {
    Write-Host "Création du dossier local de sortie..."
    New-Item -ItemType Directory -Path $Output -Force | Out-Null
}

Write-Host "Téléchargement du dossier Outlook '$Folder' vers le chemin local '$Output'"

if ($MaxMessages -gt 0) {
    Write-Host "Récupération de $MaxMessages message(s) depuis le dossier Outlook..."
    $messages = Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folderObj.Id -Top $MaxMessages -Property 'ReceivedDateTime,Subject'
} else {
    Write-Host "Aucun maximum spécifié : récupération de tous les messages du dossier Outlook. Cela peut prendre beaucoup de temps pour un grand dossier." 
    $messages = Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folderObj.Id -All -Property 'ReceivedDateTime,Subject'
}

$total = 0
$downloaded = 0
$skipped = 0

foreach ($msg in $messages) {
    $total++
    $id = $msg.Id
    $subjectPart = Get-SafeFileNamePart $msg.Subject
    $datePrefix = if ($msg.ReceivedDateTime) { (Get-Date $msg.ReceivedDateTime -Format 'yyyy_MM_dd') } else { (Get-Date -Format 'yyyy_MM_dd') }
    $baseName = if ($subjectPart) { "$datePrefix`_$subjectPart" } else { "$datePrefix`_$id" }
    $fileName = "$baseName.eml"
    $path = Join-Path $Output $fileName

    if (Test-Path $path) {
        Write-Host "Skipped existing email: $id -> $fileName"
        $skipped++
        continue
    }

    # If a same subject/date file already exists, append the message ID to avoid overwriting
    if (Get-ChildItem -Path $Output -Filter "$baseName*.eml" -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $path }) {
        $fileName = "$baseName`_$id.eml"
        $path = Join-Path $Output $fileName
    }

    Get-MgUserMessageContent -UserId $userId -MessageId $id -OutFile $path
    Write-Host "Downloaded: $fileName"
    $downloaded++
}

Write-Host "Messages récupérés : $($messages.Count)"
if ($total -eq 0) {
    Write-Host "Aucun message trouvé dans le dossier '$Folder'."
} else {
    Write-Host "Total vérifié: $total. Téléchargés: $downloaded. Ignorés (existants): $skipped."
}
