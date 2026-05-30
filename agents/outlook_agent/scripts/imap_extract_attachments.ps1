param(
    [string]$Username,
    [string]$Password,
    [string]$Output = ''
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $Output) {
    $Output = Join-Path $scriptDir 'attachments'
}

function Resolve-AssemblyPath {
    param([string]$AssemblyName)
    $path = Join-Path $scriptDir "lib\$AssemblyName"
    if (Test-Path $path) { return $path }
    return $null
}

function Load-RequiredAssemblies {
    $required = @(
        'System.Runtime.dll',
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

    Write-Host "Authentification..."
    $client.Authenticate($Username, $Password)

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

