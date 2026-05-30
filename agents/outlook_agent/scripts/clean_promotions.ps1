param(
    [string]$Folder = "Inbox",
    [int]$MaxMessages = 0,
    [string[]]$ForceDeleteDomains = @(),
    [switch]$StandardCleanup,
    [switch]$DryRun,
    [switch]$AutoConfirm
)

. "$PSScriptRoot\graph_auth.ps1"

# Default domains used to search local .eml imports when -StandardCleanup is used
$defaultForceDeleteDomains = @(
    'linkedin.com',
    'em.linkedin.com',
    'amazon.fr',
    'amazon.com',
    'eseis-syndic.sergic.com',
    'eseis-syundic.sergic.com',
    'mags.france-abonnements.fr',
    'france-abonnements.fr',
    'leboncoin.fr',
    'info.swile.co',
    'mail.gmf.fr',
    'chronopost.fr'
)

# If StandardCleanup is used, merge with ForceDeleteDomains
if ($StandardCleanup) {
    $ForceDeleteDomains = $ForceDeleteDomains + $defaultForceDeleteDomains | Select-Object -Unique
}


function Get-SafeString {
    param([Parameter(ValueFromPipeline=$true)]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    return [string]$Value
}

function Get-SafeFileNamePart {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $safe = $Value -replace '[^\w\s-]', ''
    return $safe.Trim()
}

function Normalize-LocalEmailFileName {
    param([string]$Name)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $normalized = $base -replace '^[0-9]{4}_[0-9]{2}_[0-9]{2}[_ -]?', ''
    return $normalized
}

function Get-MailFolderChildren {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserId,
        [Parameter(Mandatory=$true)]
        [string]$ParentFolderId
    )

    $children = @()
    try {
        $childFolders = Get-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $ParentFolderId -All
    } catch {
        return @()
    }

    foreach ($child in $childFolders) {
        $children += $child
        $children += Get-MailFolderChildren -UserId $UserId -ParentFolderId $child.Id
    }

    return $children
}

function Get-AllMailFolders {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserId
    )

    $rootFolders = Get-MgUserMailFolder -UserId $UserId -Top 100
    $allFolders = @()

    foreach ($folder in $rootFolders) {
        $allFolders += $folder
        $allFolders += Get-MailFolderChildren -UserId $UserId -ParentFolderId $folder.Id
    }

    return $allFolders
}

function Get-LocalEmailFileInfo {
    param([string]$FilePath)
    
    try {
        $content = Get-Content -Path $FilePath -Raw
        $from = if ($content -match 'From: [^<]*<([^>]+)>') { $matches[1] } elseif ($content -match 'From: ([^\n]+)') { $matches[1].Trim() } else { $null }
        $subject = if ($content -match 'Subject: ([^\n]+)') { $matches[1].Trim() } else { $null }
        $date = if ($content -match 'Date: ([^\n]+)') { $matches[1].Trim() } else { $null }
        
        return @{ From = $from; Subject = $subject; Date = $date; FileName = [System.IO.Path]::GetFileName($FilePath) }
    } catch {
        return $null
    }
}

function Find-LocalEmailsInStandardDomains {
    param(
        [string]$LocalEmailFolder,
        [string[]]$Domains
    )
    
    $localEmails = @()
    if (-not (Test-Path $LocalEmailFolder)) {
        return $localEmails
    }
    
    $emlFiles = Get-ChildItem -Path $LocalEmailFolder -Filter '*.eml' -Recurse -ErrorAction SilentlyContinue
    foreach ($file in $emlFiles) {
        $info = Get-LocalEmailFileInfo -FilePath $file.FullName
        if ($info -and $info.From) {
            $fromLower = $info.From.ToLower()
            foreach ($domain in $Domains) {
                if ($fromLower -like "*@$(($domain.ToLower()).Trim())") {
                    $localEmails += $info
                    break
                }
            }
        }
    }
    
    return $localEmails
}

function Normalize-EmailText {
    param([string]$Text)
    if (-not $Text) { return '' }
    return ($Text -replace '\s+', ' ').Trim().ToLower()
}

function Get-LocalEmailMatchKey {
    param(
        [string]$From,
        [string]$Subject
    )
    return "$((Normalize-EmailText -Text $From))|$((Normalize-EmailText -Text $Subject))"
}

function Get-MessageAgeDays {
    param([datetime]$DateTime)
    if ($null -eq $DateTime) { return $null }
    return [int](((Get-Date) - $DateTime).TotalDays)
}

$ageSensitiveSenders = @{ 'mondialrelay' = 30 }

$whitelistSenders = @(
    'carine.vidal1977@gmail.com',
    'noreply-gimcovermeille@mygercop.com',
    'no-reply@amazon.fr',
    'noreply@ms.contactevery.one'
)

$whitelistDomains = @(
    'gerep.com',
    'gerep.fr',
    'info.gerep.fr',
    'horizon.fr',
    'agira.asso.fr',
    'linkedin.com',
    'natixis.com',
    'picard.fr',
    'doctolib.fr',
    'dgfip.finances.gouv.fr',
    'serviceclients.leclerc',
    'eseis-syndic.sergic.com'
)

function Is-InWhitelist {
    param([string]$From)
    $fromLower = Get-SafeString $From | ForEach-Object { $_.ToLower() }
    
    # Check exact sender match
    foreach ($wl in $whitelistSenders) {
        if ($fromLower -eq $wl.ToLower()) {
            return $true
        }
    }
    
    # Check domain match
    foreach ($domain in $whitelistDomains) {
        if ($fromLower -like "*@$($domain.ToLower())" -or $fromLower -like "*$($domain.ToLower())") {
            return $true
        }
    }
    
    return $false
}

function Is-PromoMessage {
    param(
        [string]$Subject,
        [string]$Preview,
        [string]$From,
        [array]$Headers,
        [int]$AgeDays
    )

    if ($StandardCleanup) {
        if ($localEmailMatchKeys -and $localEmailMatchKeys.Count -gt 0) {
            $matchKey = Get-LocalEmailMatchKey -From $From -Subject $Subject
            return $localEmailMatchKeys -contains $matchKey
        }
        return $false
    }

    # If ForceDeleteDomains is specified, only match emails from those domains
    if ($ForceDeleteDomains -and $ForceDeleteDomains.Count -gt 0) {
        $fromLower = Get-SafeString $From | ForEach-Object { $_.ToLower() }
        foreach ($domain in $ForceDeleteDomains) {
            if ($fromLower -like "*@$(($domain.ToLower()).Trim())") {
                return $true
            }
        }
        return $false
    }

    # Check whitelist first
    if (Is-InWhitelist -From $From) {
        return $false
    }

    $subjectLower = Get-SafeString $Subject | ForEach-Object { $_.ToLower() }
    $previewLower = Get-SafeString $Preview | ForEach-Object { $_.ToLower() }
    $fromLower = Get-SafeString $From | ForEach-Object { $_.ToLower() }

    $ageSender = $ageSensitiveSenders.GetEnumerator() | Where-Object { $fromLower -like "*$(($_.Key).ToLower())*" } | Select-Object -First 1
    if ($ageSender) {
        $minAge = $ageSender.Value
        if ($AgeDays -ge $minAge) {
            return $true
        }
        return $false
    }

    $promoKeywords = 'unsubscribe','list-unsubscribe','promo','promotion','sale','offre','newsletter','publicit','pub','discount','coupon','deal','offer','promot','free shipping','no-reply','noreply','offer','vente','promo','black friday','cyber monday','deal','amaz*'
    foreach ($k in $promoKeywords) {
        if ($subjectLower.Contains($k) -or $previewLower.Contains($k) -or $fromLower.Contains($k)) { return $true }
    }

    if ($Headers) {
        foreach ($h in $Headers) {
            $headerName = Get-SafeString $h.Name | ForEach-Object { $_.ToLower() }
            if ($headerName -eq 'list-unsubscribe') { return $true }
        }
    }

    return $false
}

# When StandardCleanup is enabled, first check local email folder
if ($StandardCleanup) {
    $localEmailFolder = Join-Path (Get-Location) 'outlook_emails'
    Write-Host "StandardCleanup: Consultation des emails localement importés dans '$localEmailFolder'..."
    $localEmails = Find-LocalEmailsInStandardDomains -LocalEmailFolder $localEmailFolder -Domains $ForceDeleteDomains

    if ($localEmails.Count -gt 0) {
        Write-Host "StandardCleanup: $($localEmails.Count) email(s) trouvé(s) localement pour suppression."
        $localEmailMatchKeys = $localEmails | ForEach-Object {
            Get-LocalEmailMatchKey -From $_.From -Subject $_.Subject
        } | Select-Object -Unique
    } else {
        Write-Host "StandardCleanup: aucun email local trouvé pour les domaines spécifiés ou dans les imports locaux."
        $localEmailMatchKeys = @()
    }
} else {
    $localEmails = @()
    $localEmailMatchKeys = @()
}

$userId = (Get-MgUser -UserId me).Id
$knownFolderIds = @{ 'Inbox'='Inbox'; 'Sent Items'='SentItems'; 'Drafts'='Drafts'; 'Deleted Items'='DeletedItems'; 'Junk Email'='JunkEmail' }
$folderObjs = @()

if ($Folder -and $Folder.ToLower() -eq 'all') {
    Write-Host "Nettoyage de tous les dossiers Outlook..."
    $folderObjs = Get-AllMailFolders -UserId $userId
    if (-not $folderObjs) {
        Write-Error "Aucun dossier Outlook trouvé pour le nettoyage."
        exit 1
    }
} else {
    $folderId = if ($knownFolderIds.ContainsKey($Folder)) { $knownFolderIds[$Folder] } else { $null }
    if ($folderId) {
        $folderObj = Get-MgUserMailFolder -UserId $userId -MailFolderId $folderId
    } else {
        $folderObj = Get-MgUserMailFolder -UserId $userId -Top 100 | Where-Object { $_.DisplayName -ieq $Folder } | Select-Object -First 1
    }
    if (-not $folderObj) {
        Write-Error "Mail folder '$Folder' introuvable."
        exit 1
    }
    $folderObjs = @($folderObj)
}

# fetch messages (limited when requested)
$msgs = @()
foreach ($folderObj in $folderObjs) {
    Write-Host "Scan du dossier '$($folderObj.DisplayName)' ($($folderObj.Id))..."
    if ($MaxMessages -gt 0) {
        $msgs += Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folderObj.Id -Top $MaxMessages
    } else {
        $msgs += Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folderObj.Id -All
    }
}

$candidates = @()
foreach ($m in $msgs) {
    $full = Get-MgUserMessage -UserId $userId -MessageId $m.Id -Property 'Subject,From,ReceivedDateTime,BodyPreview,InternetMessageHeaders'
    $subject = Get-SafeString $full.Subject
    $preview = Get-SafeString $full.BodyPreview
    $from = Get-SafeString $full.From.EmailAddress.Address
    $headers = $full.InternetMessageHeaders
    $ageDays = Get-MessageAgeDays $full.ReceivedDateTime

    if (Is-PromoMessage -Subject $subject -Preview $preview -From $from -Headers $headers -AgeDays $ageDays) {
        $reason = @()
        $subjectLower = Get-SafeString $subject | ForEach-Object { $_.ToLower() }
        $previewLower = Get-SafeString $preview | ForEach-Object { $_.ToLower() }
        $fromLower = Get-SafeString $from | ForEach-Object { $_.ToLower() }
        if ($subjectLower -match 'unsubscribe|list-unsubscribe|promo|promotion|sale|offre|newsletter|publicit|pub|discount|coupon|deal|offer|promot|free shipping|no-reply|noreply|black friday|cyber monday|vente|amazon') { $reason += 'keyword' }
        if ($fromLower -match 'no-reply|noreply|newsletter|promo|offers|deal|info@|contact@|service@') { $reason += 'sender' }
        if ($headers) {
            foreach ($h in $headers) {
                if ((Get-SafeString $h.Name).ToLower() -eq 'list-unsubscribe') { $reason += 'list-unsubscribe'; break }
            }
        }

        $ageSender = $ageSensitiveSenders.GetEnumerator() | Where-Object { $fromLower -like "*$(($_.Key).ToLower())*" } | Select-Object -First 1
        if ($ageSender -and $ageDays -ge $ageSender.Value) {
            $reason += "old-$($ageSender.Key)"
        }
        if ($ageDays -ne $null) { $reason += "ageDays:$ageDays" }
        if (-not $reason) { $reason = 'match' }

        $candidates += [PSCustomObject]@{
            Id = $m.Id
            Subject = $subject
            From = $from
            Received = $full.ReceivedDateTime
            AgeDays = $ageDays
            Reason = ($reason -join ', ')
        }
    }
}

if ($candidates.Count -eq 0) {
    Write-Host "Aucun email promotionnel détecté (selon règles heuristiques)."
    exit 0
}

Write-Host "Messages détectés comme potentiellement promotionnels : $($candidates.Count)"
$candidates | Select-Object Subject, From, Received, AgeDays, Reason | Format-Table -AutoSize

if ($DryRun) {
    Write-Host "Mode DryRun activé: aucune suppression ne sera effectuée."
    exit 0
}

if (-not $AutoConfirm) {
    $ok = Read-Host "Confirmez-vous la suppression de ces messages? Tapez OUI pour confirmer"
    if ($ok -ne 'OUI') { Write-Host 'Suppression annulée.'; exit 0 }
}

# delete messages both local files (if present) and mailbox
foreach ($c in $candidates) {
    try {
        Remove-MgUserMessage -UserId $userId -MessageId $c.Id -ErrorAction Stop -Confirm:$false
        Write-Host "Supprimé: $($c.Subject) <$($c.From)>"
    } catch {
        Write-Warning "Impossible de supprimer le message $($c.Id): $_"
    }

    $sanitizedSubject = Get-SafeFileNamePart $c.Subject
    $possibleFileMatch = "$sanitizedSubject.eml"
    $local = Get-ChildItem -Path .. -Recurse -Filter '*.eml' -ErrorAction SilentlyContinue | Where-Object {
        $normalized = Normalize-LocalEmailFileName $_.Name
        $_.Name -like "$($c.Id)*" -or $normalized -eq $sanitizedSubject
    }
    foreach ($f in $local) { Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue }
}

Write-Host "Opération terminée. $($candidates.Count) messages supprimés (tentative)."
