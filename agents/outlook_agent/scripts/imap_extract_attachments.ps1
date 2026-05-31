param(
    [string]$Username,
    [string]$Password,
    [ValidateSet('auto','login','plain','oauth2')] [string]$AuthMethod = 'auto',
    [string]$OAuth2Token = '',
    [string]$Output = ''
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $Output) {
    $Output = Join-Path $scriptDir 'attachments'
}

if (-not $Username) {
    Write-Error "Veuillez fournir -Username pour l'action IMAP."
    exit 1
}

if ($AuthMethod -eq 'oauth2') {
    if (-not $OAuth2Token) {
        Write-Error "Le mode oauth2 nécessite -OAuth2Token. Utilisez get_imap_oauth2_token.ps1 pour obtenir un jeton valide."
        exit 1
    }
} else {
    if (-not $Password) {
        Write-Error "Le mode d'authentification '$AuthMethod' nécessite un mot de passe. Utilisez -AuthMethod oauth2 avec -OAuth2Token pour XOAUTH2."
        exit 1
    }
}

function Resolve-AssemblyPath {
    param([string]$AssemblyName)
    $path = Join-Path $scriptDir "lib\$AssemblyName"
    if (Test-Path $path) { return $path }
    return $null
}

function Load-RequiredAssemblies {
    $required = @(
        'System.Threading.Tasks.Extensions.dll',
        'MimeKit.dll',
        'MailKit.dll'
    )

    $missing = @()
    foreach ($dll in $required) {
        $path = Resolve-AssemblyPath -AssemblyName $dll
        if (-not $path) {
            $missing += $dll
            continue
        }

        try {
            [Reflection.Assembly]::LoadFrom($path) | Out-Null
            Write-Host "Loaded assembly: $dll"
        } catch {
            Write-Warning "Impossible de charger l'assembly '$dll' depuis '$path': $($_.Exception.Message)"
            $missing += $dll
        }
    }

    if ($missing.Count -gt 0) {
        Write-Error "Assemblies manquants ou introuvables dans '$scriptDir\lib': $($missing -join ', ')"
        Write-Error "Lancez install_mailkit_deps.ps1 pour telecharger les dependances automatiquement."
        exit 1
    }
}

function Get-AllFolders {
    param([MailKit.IMailFolder]$Folder)

    $folders = @($Folder)
    foreach ($sub in $Folder.GetSubfolders($false)) {
        $folders += Get-AllFolders -Folder $sub
    }
    return $folders
}

function Ensure-FolderExists {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

Load-RequiredAssemblies
Ensure-FolderExists -Path $Output

$client = [MailKit.Net.Imap.ImapClient]::new()
try {
    Write-Host "Connexion IMAP..."
    $client.Connect('imap-mail.outlook.com', 993, $true)

    Write-Host "IMAP capacités : $($client.Capabilities)"
    Write-Host "Mechanismes auth supportés : $($client.AuthenticationMechanisms -join ', ')"
    Write-Host "Authentification avec AuthMethod=$AuthMethod..."

    if ($AuthMethod -eq 'oauth2') {
        if (-not $OAuth2Token) {
            Write-Error "AuthMethod oauth2 nécessite le paramètre -OAuth2Token avec un jeton valide."
            exit 1
        }
        $mechanism = New-Object MailKit.Security.SaslMechanismOAuth2($Username, $OAuth2Token)
    } else {
        $preferred = $AuthMethod
        if ($AuthMethod -eq 'auto') {
            if ($client.AuthenticationMechanisms -contains 'LOGIN') {
                $preferred = 'login'
            } elseif ($client.AuthenticationMechanisms -contains 'PLAIN') {
                $preferred = 'plain'
            }
        }

        switch ($preferred) {
            'login' {
                if (-not ($client.AuthenticationMechanisms -contains 'LOGIN')) {
                    Write-Error "Serveur IMAP ne propose pas LOGIN. Choisissez AuthMethod=plain ou vérifiez les paramètres du serveur."
                    exit 1
                }
                $mechanism = New-Object MailKit.Security.SaslMechanismLogin($Username, $Password)
            }
            'plain' {
                if (-not ($client.AuthenticationMechanisms -contains 'PLAIN')) {
                    Write-Error "Serveur IMAP ne propose pas PLAIN. Choisissez AuthMethod=login ou vérifiez les paramètres du serveur."
                    exit 1
                }
                $mechanism = New-Object MailKit.Security.SaslMechanismPlain($Username, $Password)
            }
        }
    }

    try {
        $client.Authenticate($mechanism)
    } catch {
        Write-Error "Auth failed : $($_.Exception.Message)"
        if ($_.Exception.InnerException) { Write-Error "  Inner: $($_.Exception.InnerException.Message)" }
        Write-Error "Le serveur supporte uniquement : $($client.AuthenticationMechanisms -join ', ')"
        Write-Error "Si Outlook.com rejette l'authentification basique, activez IMAP et / ou utilisez un mot de passe d'application Microsoft."
        Write-Error "Si le compte est protégé par MFA, vous devrez utiliser XOAUTH2 et un jeton d'accès valide."
        exit 1
    }

    Write-Host "Récupération des dossiers..."
    $root = $client.GetFolder($client.PersonalNamespaces[0])
    $folders = Get-AllFolders -Folder $root

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
            if (-not $msg.Attachments -or $msg.Attachments.Count -eq 0) { continue }

            $safeFolder = ($folder.FullName -replace '[\\/:*?"<>|]', '_')
            $msgFolder = Join-Path $Output $safeFolder
            Ensure-FolderExists -Path $msgFolder

            foreach ($att in $msg.Attachments) {
                $fileName = $att.FileName
                if (-not $fileName) { $fileName = 'attachment.bin' }

                $filePath = Join-Path $msgFolder $fileName
                $counter = 1
                while (Test-Path $filePath) {
                    $filePath = Join-Path $msgFolder "{0}_{1}{2}" -f [System.IO.Path]::GetFileNameWithoutExtension($fileName), $counter, [System.IO.Path]::GetExtension($fileName)
                    $counter++
                }

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
} catch {
    Write-Error "Erreur IMAP : $($_.Exception.Message)"
    if ($_.Exception.InnerException) { Write-Error "  Inner: $($_.Exception.InnerException.Message)" }
    exit 1
} finally {
    if ($client -and $client.IsConnected) {
        $client.Disconnect($true)
    }
}

